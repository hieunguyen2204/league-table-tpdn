#!/usr/bin/env python3
"""
Build the League Table dashboard (single offline HTML file) from the TPDN Excel workbook.

Backend update flow (monthly):
    1. Replace / update the Excel file in ./data (keep the column names of the source sheet).
    2. Run:   python build_league_table.py                      # uses the newest .xlsx in ./data
              python build_league_table.py data/TPPH.xlsx -o League_Table_TPDN.html
    3. Open League_Table_TPDN.html in a browser (or send the file as-is, it needs no internet).

Metrics          : Số đợt PH = number of rows (issues); KLPH = GTPH in tỷ đồng; % phân phối = public offerings only.
Required columns : GTPH (TỶ ĐỒNG), NGÀY PHÁT HÀNH, Tổ chức tư vấn, Đại diện sở hữu trái phiếu
Optional columns : MÃ TRÁI PHIẾU, CODE, KỲ HẠN, % PHÂN PHỐI, CÁCH THỨC PHÁT HÀNH, NHÓM NGÀNH,
                   Tổ chức liên quan (cột S: advisor of private placements; several organisations separated by ';',
                   only securities companies (CTCK) are counted, each gets full credit for the issue),
                   Công ty luật tư vấn (several firms separated by ';')
Files next to this script: template.html, vendor/chart.umd.js (Chart.js 4.4.7, inlined into the output).
Requires: pandas, openpyxl.
"""
import argparse, datetime as dt, json, re, sys, unicodedata
from pathlib import Path
import pandas as pd

HERE = Path(__file__).resolve().parent

ACR = {"SSI": "SSI", "MB": "MB", "ACB": "ACB", "BIDV": "BIDV", "JB": "JB", "LPBANK": "LPBank", "VIETINBANK": "VietinBank",
       "AGRIBANK": "Agribank", "VNDIRECT": "VNDirect", "OCBS": "OCBS", "NSI": "NSI", "KAFI": "KAFI", "TCBS": "TCBS",
       "VPS": "VPS", "HSC": "HSC", "VIX": "VIX", "VCBS": "VCBS", "SHS": "SHS", "MBS": "MBS", "VCSC": "VCSC"}
PRE = re.compile(r"^(CTCP|CT TNHH|CTTNHH|CÔNG TY CỔ PHẦN|CÔNG TY TNHH|CÔNG TY)\s+CHỨNG KHOÁN\s+", re.I)


def norm(s):
    s = str(s).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def title_case(s):
    out = []
    for w in s.split(" "):
        m = re.match(r"^(\(?)(.*?)(\)?)$", w)
        pre, core, post = m.groups()
        if core.upper() in ACR:
            out.append(pre + ACR[core.upper()] + post)
        elif not re.search(r"[A-Za-zÀ-ỹ]", core):
            out.append(w)
        else:
            out.append(pre + core[:1].upper() + core[1:].lower() + post)
    return " ".join(out)


def clean_sec(v):
    """Securities-firm names: strip 'CTCP/CT TNHH Chứng khoán' and unify the spelling."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return ""
    s = PRE.sub("", re.sub(r"\s+", " ", str(v).strip()))
    return title_case(s) if s == s.upper() else s


ACR = {"SSI":"SSI","MB":"MB","ACB":"ACB","BIDV":"BIDV","JB":"JB","LPBANK":"LPBank","VIETINBANK":"VietinBank","AGRIBANK":"Agribank","VNDIRECT":"VNDirect","OCBS":"OCBS","NSI":"NSI","KAFI":"KAFI","TCBS":"TCBS","VPS":"VPS","HSC":"HSC","VIX":"VIX","VCBS":"VCBS","SHS":"SHS","MBS":"MBS","VCSC":"VCSC","UP":"UP","HD":"HD","VPBANK":"VPBank","TVSI":"TVSI","BSC":"BSC","FPTS":"FPTS","VDSC":"VDSC","SHBS":"SHBS","MSBS":"MSBS","DSC":"DSC","ABS":"ABS","EVS":"EVS","PSI":"PSI","TCSC":"TCSC","VFS":"VFS","VND":"VND","IVS":"IVS","SAIGONBANK":"Saigonbank","FPT":"FPT","APG":"APG","BOS":"BOS","AIS":"AIS","NH":"NH","DNSE":"DNSE","KIS":"KIS","KB":"KB","HSC":"HSC"}
SEC=re.compile(r"^(công ty (cổ phần|tnhh)|ctcp|ct tnhh|cttnhh)\s+chứng khoán\s+",re.I)
SEC2=re.compile(r"^công ty\s+chứng khoán\s+",re.I)
BANK=re.compile(r"^ngân hàng\s+(thương mại cổ phần|tmcp|tpcp|thương mại tnhh một thành viên|tnhh một thành viên|thương mại)?\s*",re.I)
BRAND=[("ngoại thương","VCBS","NH Vietcombank"),("công thương","VietinBank","NH VietinBank"),("đầu tư và phát triển","BIDV","NH BIDV"),("thành phố hồ chí minh","HSC",None)]
OLD=re.compile(r"\(\s*(tên cũ|việt nam)[^)]*\)",re.I)
def tcase(s):
    out=[]
    for w in s.split(" "):
        m=re.match(r"^(\(?)(.*?)(\)?)$",w);pre,core,post=m.groups()
        if core.upper() in ACR: out.append(pre+ACR[core.upper()]+post)
        elif not re.search(r"[A-Za-zÀ-ỹ]",core): out.append(w)
        else: out.append(pre+core[:1].upper()+core[1:].lower()+post)
    return " ".join(out)
def clean_org(p):
    s=re.sub(r"\s+"," ",OLD.sub("",p.replace("–","-").replace("—","-"))).strip(" ,.")
    if not s: return ""
    m=SEC.match(s) or SEC2.match(s)
    if m:
        rest=s[m.end():].strip()
        if re.match(r"^ngân hàng",rest,re.I):
            for k,sec,_ in BRAND:
                if k in rest.lower(): return sec
        if re.match(r"^thành phố hồ chí minh",rest,re.I): return "HSC"
        return tcase(rest)
    if re.match(r"^chứng khoán\s+",s,re.I): return tcase(re.sub(r"^chứng khoán\s+","",s,flags=re.I))
    m=BANK.match(s)
    if m and s.lower().startswith("ngân hàng"):
        rest=re.sub(r"\s*-\s*chi nhánh.*$","",s[m.end():].strip(),flags=re.I)
        for k,_,bn in BRAND:
            if bn and k in rest.lower(): return bn
        rest=re.sub(r"\s+việt nam$","",rest,flags=re.I)
        return "NH "+tcase(rest)
    return s
ALIAS={"Quốc Gia":"Quốc Gia (NSI)"}
def is_sec(p):
    return bool(re.search(r"chứng khoán", p, re.I))


def org_names(v, only_sec=False, tag=False):
    """Split a cell on ';', normalise every organisation; only_sec=True keeps securities companies (CTCK) only."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return []
    out = []
    for part in re.split(r"[;\n]+", str(v)):
        if only_sec and not is_sec(part):
            EXCLUDED.append(part.strip())
            continue
        n = clean_org(part); n = ALIAS.get(n, n)
        if n and tag and is_sec(part):
            n = "CTCK " + n                    # make clear these are securities companies, not their parent banks
        if n and n not in out:
            out.append(n)
    return out


def txt(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def tenor_group(t):
    if t is None:
        return ""
    return "≤ 3 năm" if t <= 3 else ("3 – 7 năm" if t <= 7 else "> 7 năm")


def num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("%", "").replace(" ", "")
    if not t or t == "-":
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
    elif "," in t:
        t = t.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", t) else t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def load_sheet(path, sheet):
    xl = pd.ExcelFile(path)
    names = [sheet] if sheet else (['Data'] if 'Data' in xl.sheet_names else xl.sheet_names)
    best = None
    for n in names:
        df = xl.parse(n, header=None)
        for i in range(min(15, len(df))):
            if any(norm(c).startswith("gtph") for c in df.iloc[i].tolist()):
                d = df.iloc[i + 1:].copy()
                names_, seen = [], {}
                for c_ in df.iloc[i].tolist():          # duplicate headers (e.g. two 'NHÓM NGÀNH'): keep the first as is
                    k_ = norm(c_); seen[k_] = seen.get(k_, 0) + 1
                    names_.append(k_ if seen[k_] == 1 else f"{k_} #{seen[k_]}")
                d.columns = names_
                if best is None or len(d) > len(best):
                    best = d
                break
    if best is None:
        sys.exit("Không tìm thấy cột 'GTPH (TỶ ĐỒNG)' trong file Excel.")
    return best


def col(df, pred):
    for c in df.columns:
        if pred(c):
            return c
    return None


SKIPPED = []
EXCLUDED = []


def build_records(df, from_date=None):
    c = {
        "vol": col(df, lambda h: h.startswith("gtph")), "date": col(df, lambda h: h == "ngay phat hanh"),
        "adv": col(df, lambda h: h == "to chuc tu van"), "tr": col(df, lambda h: h == "dai dien so huu trai phieu"),
        "bond": col(df, lambda h: h == "ma trai phieu"), "code": col(df, lambda h: h == "code"),
        "tenor": col(df, lambda h: h == "ky han"), "dist": col(df, lambda h: "phan phoi" in h),
        "method": col(df, lambda h: h == "cach thuc phat hanh"), "sector": col(df, lambda h: h == "nhom nganh"),
        "rel": col(df, lambda h: h == "to chuc lien quan"),
        "law": col(df, lambda h: "luat" in h or "phap ly" in h),
        "coll": col(df, lambda h: "quan ly" in h and ("tsbd" in h or "tai san" in h)),
    }
    g = lambda r, k: (r[c[k]] if c[k] is not None else None)
    out = []
    for _, r in df.iterrows():
        d = pd.to_datetime(g(r, "date"), errors="coerce", dayfirst=False)
        date = d.strftime("%Y-%m-%d") if pd.notna(d) else ""
        if from_date and (not date or date < from_date):
            continue
        vol = num(g(r, "vol"))
        if vol is None:
            SKIPPED.append(txt(g(r, "bond")) or txt(g(r, "code")))
            continue
        method = txt(g(r, "method")).upper()
        is_public = method.startswith("CÔNG")
        a_col, s_col = org_names(g(r, "adv"), True, True), org_names(g(r, "rel"), True, True)
        advs = (a_col or s_col) if is_public else (s_col or a_col)   # public: AJ; private: S ("Tổ chức liên quan")
        raw_dist = num(g(r, "dist"))
        if raw_dist is not None and raw_dist > 1:
            raw_dist /= 100
        if method and not method.startswith("CÔNG"):
            raw_dist = None                        # % phân phối applies to public offerings only
        tenor = num(g(r, "tenor"))
        law_raw = g(r, "law")
        law = [x.strip() for x in re.split(r"[;\n]+", str(law_raw)) if x.strip()] if law_raw is not None and not (isinstance(law_raw, float) and pd.isna(law_raw)) else []
        sector = txt(g(r, "sector")).capitalize()
        out.append({
            "bond": txt(g(r, "bond")), "code": txt(g(r, "code")),
            "vol": round(vol, 4), "date": date, "y": int(date[:4]) if date else None, "ym": date[:7] if date else "",
            "tenor": None if tenor is None else round(tenor, 2), "tg": tenor_group(tenor), "advs": advs, "coll": org_names(g(r, "coll")), "tr": (org_names(g(r, "tr"), False, True) or [""])[0], "law": law,
            "dist": None if raw_dist is None else round(raw_dist, 6), "method": method.capitalize() if method else "",
            "sector": sector,
        })
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", help="Excel file (default: newest .xlsx in ./data)")
    ap.add_argument("-o", "--output", default=str(HERE / "League_Table_TPDN.html"))
    ap.add_argument("--sheet", help="sheet name (default: auto-detect the sheet with the GTPH column)")
    ap.add_argument("--from-date", default=f"{dt.date.today().year}-01-01",
                    help="only issues from this date (YYYY-MM-DD, default 1 Jan of the current year; 'all' = no filter)")
    ap.add_argument("--title", default="League Table trái phiếu doanh nghiệp")
    a = ap.parse_args()
    src = Path(a.input) if a.input else max((HERE / "data").glob("*.xlsx"), key=lambda p: p.stat().st_mtime, default=None)
    if not src or not src.exists():
        sys.exit("Không tìm thấy file Excel. Đặt file vào thư mục data/ hoặc truyền đường dẫn.")
    fd = None if a.from_date.lower() == "all" else a.from_date
    recs = build_records(load_sheet(src, a.sheet), fd)
    meta = {"title": a.title, "source": src.name, "n": len(recs), "from": fd or "", "built": dt.datetime.now().strftime("%d/%m/%Y %H:%M")}
    tpl = (HERE / "template.html").read_text(encoding="utf-8")
    lib = HERE / "vendor" / "chart.umd.js"
    chart = ("<script>" + lib.read_text(encoding="utf-8").replace("</script", "<\\/script") + "</script>") if lib.exists() else \
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>'
    j = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = tpl.replace("__CHARTJS__", chart).replace("__DATA__", j(recs)).replace("__META__", j(meta))
    Path(a.output).write_text(html, encoding="utf-8")
    n_adv = len({x for r in recs for x in r["advs"]}); n_tr = len({r["tr"] for r in recs if r["tr"]})
    print(f"OK: {a.output}  | từ {fd or 'đầu'} | {len(recs)} đợt PH | {sum(r['vol'] for r in recs):,.1f} tỷ đồng | "
          f"{n_adv} tổ chức tư vấn | {n_tr} đại diện sở hữu | {sum(1 for r in recs if r['law'])} đợt có công ty luật")
    miss_a = sum(1 for r in recs if not r["advs"]); miss_t = sum(1 for r in recs if not r["tr"])
    if miss_a or miss_t:
        print(f"Lưu ý: {miss_a} đợt thiếu tổ chức tư vấn, {miss_t} đợt thiếu đại diện sở hữu (không được xếp hạng).")
    if EXCLUDED:
        import collections
        top = ", ".join(f"{k} ({v})" for k, v in collections.Counter(EXCLUDED).most_common(6))
        print(f"Đã loại {len(EXCLUDED)} tổ chức không phải CTCK khỏi danh sách tư vấn (ngân hàng, tổ chức nước ngoài...), ví dụ: {top}")
    if SKIPPED:
        print(f"Bỏ qua {len(SKIPPED)} dòng có GTPH không phải số (ví dụ trái phiếu quốc tế tính bằng USD): {', '.join(SKIPPED[:5])}")


if __name__ == "__main__":
    main()

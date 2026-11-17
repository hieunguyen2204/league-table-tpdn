# League Table TPDN

Dashboard league table trái phiếu doanh nghiệp (tổ chức tư vấn CTCK, đại diện sở hữu, công ty luật, tổ chức quản lý TSBĐ). Một file HTML chạy offline, dựng từ file Excel TPDN bằng Python.

## Cập nhật dữ liệu
1. Đặt file Excel mới vào `data/` (file Excel không được commit lên GitHub, xem `.gitignore`).
2. Chạy: `python build_league_table.py data/<file>.xlsx -o docs/index.html`
   - Cần `pandas`, `openpyxl`. Mặc định lấy dữ liệu từ 1/1 của năm hiện tại (`--from-date all` để lấy toàn bộ).
3. Commit và push `docs/index.html`. Workflow `.github/workflows/pages.yml` tự triển khai lên GitHub Pages.

## Triển khai lần đầu
Settings → Pages → Build and deployment → Source: **GitHub Actions**.

## Lưu ý
- `docs/index.html` chứa toàn bộ dữ liệu đã xử lý. Repo công khai thì trang và dữ liệu đều công khai; nếu chưa muốn công khai, dùng repo private (GitHub Pages cho repo private cần gói trả phí).
- Giao diện nằm trong `template.html`; Chart.js 4.4.7 nằm trong `vendor/`.

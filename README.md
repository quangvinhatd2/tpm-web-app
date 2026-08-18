# Hệ thống TPM HPCBANVE — Ghi chú vận hành

Flask + PostgreSQL (Neon) + Excel làm nguồn dữ liệu tĩnh, deploy trên Render.

---

## 1. Ba file Excel trong dự án — vai trò và nơi dùng

| File | Vai trò | Route/hàm sử dụng | Cache trong app? |
|---|---|---|---|
| `forms.xlsx` | Nội dung form đánh giá hiển thị lên web: tiêu đề, chức danh, hạng mục, tiêu chuẩn... | mọi trang xem/nhập đánh giá (`get_master_workbook()`, `get_sheet_data()`) | Có — cache trong bộ nhớ, tự kiểm tra `mtime` mỗi 10 giây (`_MTIME_CHECK_INTERVAL`) |
| `phan_giao.xlsx` | Bảng phân công: ai đánh giá / thẩm tra sheet nào | `/sync_assignments` (nút "Đồng bộ phân công" trong admin) | Không — đọc lại từ đầu mỗi lần bấm nút |
| `Template.xlsx` | Khung mẫu để sinh file báo cáo tổng hợp khiếm khuyết (K/Đ) cho tải về | `/export_summary` (nút "Xuất file báo cáo tổng") | Không — đọc lại từ đầu mỗi lần bấm nút |

⚠️ **Lưu ý tên file:** trong code, `Template.xlsx` được load bằng `load_workbook('template.xlsx')` (chữ thường). Trên Mac (filesystem không phân biệt hoa/thường) vẫn chạy được, nhưng **Render dùng Linux — phân biệt hoa/thường tuyệt đối**. Nếu file trong repo tên `Template.xlsx` (chữ T hoa), route `/export_summary` có thể lỗi "không tìm thấy file" trên production. Cần kiểm tra/sửa cho khớp tên nếu chưa test route này trên Render.

---

## 2. Vì sao sửa Excel local không tự lên web

Server chạy trên **Render**, đọc file Excel từ đĩa của chính server đó — **không phải** từ máy Mac local. Sửa file trên máy cá nhân chỉ có tác dụng khi:

1. File được `git commit` + `git push`
2. Render tự động phát hiện push mới → build lại → deploy
3. Trong quá trình build, Git checkout ghi đè toàn bộ file (kể cả Excel) vào server

Nếu bỏ qua bước 1 (quên `git add` file Excel), Render vẫn deploy nhưng dùng **bản Excel cũ nhất từng được commit** — web sẽ không đổi dù local đã sửa. Đây là lỗi đã gặp thực tế: `git status` cho thấy `phan_giao.xlsx` bị "modified" nhưng chưa từng add/commit.

---

## 3. Quy trình chuẩn — mỗi lần sửa `forms.xlsx` / `phan_giao.xlsx` / `Template.xlsx`

```bash
# Trong thư mục dự án (TPM), sau khi đã lưu (Cmd+S) file Excel trong Excel:

git add -A
git commit -m "mo ta ngan gon: vi du 'cap nhat chuc danh truc ca' hoac 'phan cong thang 9'"
git push
```

> Dùng `git add -A` (thay vì add từng file) để giảm rủi ro quên add — lệnh này tự động thêm mọi file đã thay đổi trong working directory.

Sau khi push:

1. Vào Render Dashboard → tab **Events** của service → xác nhận có dòng "Deploy started for <commit hash>"
2. Đợi deploy chuyển sang trạng thái **Live** (thường 1–8 phút tùy tải, xem log chi tiết từng bước Cloning / Installing dependencies / Building / Starting nếu muốn biết bước nào chậm)
3. Vào lại web → đăng nhập **admin** → nếu vừa sửa `phan_giao.xlsx`, bấm **"Đồng bộ phân công"** để áp dụng vào database (đây là bước bắt buộc riêng, sửa và push xong chưa tự động đồng bộ vào DB — phải bấm nút này)
4. Đọc các thông báo (flash message) hiện trên đầu trang dashboard sau khi bấm để xác nhận: số bản ghi thêm/xóa, hoặc "Không có thay đổi" nếu file không khác gì DB hiện tại

**Riêng `forms.xlsx` và `Template.xlsx`:** không cần bấm nút gì thêm sau khi deploy — `forms.xlsx` tự refresh cache trong vòng 10 giây kể từ khi Render deploy xong (do cơ chế kiểm tra mtime), `Template.xlsx` được đọc trực tiếp mỗi lần bấm "Xuất file báo cáo tổng" nên luôn là bản mới nhất ngay sau deploy.

---

## 4. Kiểm tra trước khi push (tránh lặp lỗi từng gặp)

```bash
git status
```

Trước khi push, đảm bảo file Excel vừa sửa xuất hiện trong danh sách "Changes not staged" hoặc "Changes to be committed" — nếu không thấy, có thể:
- Chưa lưu (Cmd+S) trong Excel — kiểm tra file `~$<tên file>.xlsx` có đang tồn tại không (dấu hiệu file còn đang mở trong Excel, chưa chắc đã lưu bản mới nhất)
- File đang bị `.gitignore` chặn — kiểm tra bằng `git check-ignore -v <tên file>`

---

## 5. Về tốc độ deploy trên Render

Deploy chậm (vài phút) là bình thường với gói Free/Starter, do:
- Cài lại dependencies (`pip install`) nếu không cache được
- Free tier bị giới hạn tài nguyên build
- Service "spin down" khi không có traffic → lần chạy lại sau đó chậm hơn (cold start)

Không phải lỗi — cứ theo dõi tab **Events** để biết đang ở bước nào.

---

## 6. ⚠️ Giới hạn quan trọng — KHÔNG dùng upload trực tiếp qua web để thay 3 file Excel này

Render (đặc biệt gói Free) có **filesystem tạm thời (ephemeral)**: mọi thay đổi ghi trực tiếp lên đĩa server (ví dụ qua một form "upload file" tương lai) sẽ **bị xóa mỗi khi service redeploy, restart, hoặc tự spin down do không có traffic**. Vì vậy:

- **Luôn cập nhật 3 file Excel này qua Git (mục 3), không upload thẳng lên server** trừ khi sau này có làm tính năng lưu file vào PostgreSQL (Neon) thay vì đĩa local — lúc đó mới an toàn qua các lần deploy.
- Đừng nhầm lẫn nếu sau này thấy dữ liệu "tự nhiên revert" — rất có thể do đã sửa trực tiếp trên server (SSH, hoặc tool nào đó) mà không qua Git, rồi service restart làm mất thay đổi đó.

---

## 7. Danh sách nhanh (cheat sheet)

```bash
# Sửa xong Excel, muốn đẩy lên web:
git status                      # kiểm tra file đã đổi có được nhận diện chưa
git add -A
git commit -m "..."
git push
# → theo dõi Render tab Events đến khi Live
# → nếu là phan_giao.xlsx: vào admin, bấm "Đồng bộ phân công"
```
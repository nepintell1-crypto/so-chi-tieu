# Sổ chi tiêu

App quản lý thu chi cá nhân, chạy trên điện thoại (Android, iPhone) và máy tính. Cài như một app từ trình duyệt, dùng được cả khi không có mạng.

## Chức năng
- Nhập, sửa, xoá khoản thu chi theo danh mục: `+` Thu nhập, `A` Ăn chơi, `C` Thiết yếu cá nhân, `G` Gia đình, `K` Chi khác, `T` Công ty.
- Xem theo tháng: Tổng thu, Tổng chi, Còn lại, tỷ lệ từng danh mục.
- Sổ thu chi: nhóm theo ngày, tìm kiếm không cần gõ dấu, lọc theo danh mục.
- Ngân sách tháng cho từng danh mục, cảnh báo khi đạt 80% và khi vượt.
- Báo cáo năm: biểu đồ 12 tháng, bảng chi tiết, xuất file Excel (sheet Báo cáo + T1–T12).
- Sao lưu / khôi phục toàn bộ dữ liệu bằng file `.json`.

## Dữ liệu
Dữ liệu chỉ lưu trong trình duyệt trên máy đang dùng (localStorage), không gửi đi đâu. Xoá dữ liệu trình duyệt hoặc gỡ app sẽ mất dữ liệu, nên sao lưu định kỳ.

## Cấu trúc
| File | Vai trò |
|---|---|
| `index.html` | Trang chính |
| `styles.css` | Giao diện "Premium Dark Luxury" |
| `app.js` | Toàn bộ chức năng |
| `manifest.webmanifest`, `icons/` | Thông tin để cài lên màn hình chính |
| `sw.js` | Lưu sẵn file để chạy offline. Sửa app thì tăng `VERSION` trong file này |
| `vendor/xlsx.full.min.js` | Thư viện SheetJS 0.18.5 (Apache-2.0) để tạo file Excel |

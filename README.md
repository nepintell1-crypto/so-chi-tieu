# MyLife Hub

Trợ lý cá nhân chạy trên điện thoại (Android, iPhone) và máy tính: việc cần làm, lịch tháng, mua sắm và thu chi trong một app. Nâng cấp từ app **Sổ chi tiêu**; cài như một app từ trình duyệt, dùng được cả khi không có mạng.

## Chức năng (giai đoạn 1)
- **Hôm nay**: tiền đã chi trong ngày, còn được chi theo ngân sách, cảnh báo 80%, việc đến hạn (dời sang mai 1 chạm), nhắc ghi chi, thẻ Tổng kết ngày từ 22:00.
- **Lịch**: lịch tháng, mỗi ô ngày ghi số tiền đã chi và chấm báo có việc; chạm ngày để xem việc và khoản thu chi.
- **Nút +**: ghi nhanh khoản thu chi (Số tiền → Danh mục → Lưu, có × nghìn / × triệu, gợi ý nội dung cũ, hoàn tác 5 giây), thêm việc, thêm món cần mua.
- **Việc**: thẻ Công việc / Cá nhân / Việc nhà, ưu tiên Cao / Vừa / Thấp, hạn và giờ, lặp lại (ngày, tuần, tháng, cuối tháng), chi phí dự kiến, đính kèm ảnh chụp và tài liệu (PDF, Word, Excel…).
- **Mua sắm**: gõ nhiều món cách nhau dấu phẩy, tự xếp nhóm hàng, nhớ giá lần trước. Tích "Đã mua" thì chọn: ghi riêng món đó, gộp cả lần đi chợ thành 1 khoản, hoặc không ghi.
- **Tiền**: Tổng quan, Sổ thu chi, Ngân sách, Báo cáo năm + xuất Excel theo mẫu cũ (Ngày | KH | Nội dung thu chi | Thu+/Chi- | Số tiền | Ghi chú). Khoản mục `T` (Công ty) có trạng thái Chờ hoàn / Đã hoàn.
- **Cài đặt**: thêm danh mục thu chi (mã 1 chữ cái) và nhóm hàng; mục đã có dữ liệu chỉ ẩn, không xoá. Sao lưu (nhẹ hoặc kèm tệp), khôi phục cả file sao lưu của Sổ chi tiêu cũ.

Danh mục mặc định: `+` Thu nhập, `A` Ăn chơi, `C` Thiết yếu cá nhân, `G` Gia đình, `K` Chi khác, `T` Công ty, `B` Cho con.

## Dữ liệu
Dữ liệu chỉ lưu trong trình duyệt trên máy đang dùng (IndexedDB), không gửi đi đâu. Lần đầu mở bản mới, dữ liệu của Sổ chi tiêu (localStorage `sochitieu.v1`) được chuyển sang tự động, đối chiếu số khoản và tổng tiền từng tháng; bản cũ được giữ nguyên để phòng hờ. Xoá dữ liệu trình duyệt hoặc gỡ app sẽ mất dữ liệu, nên sao lưu định kỳ.

## Cấu trúc
| File | Vai trò |
|---|---|
| `index.html` | Trang chính |
| `styles.css` | Giao diện "Premium Dark Luxury" |
| `store.js` | Lưu dữ liệu (IndexedDB), chuyển dữ liệu từ Sổ chi tiêu |
| `app.js` | Các màn hình và chức năng |
| `manifest.webmanifest`, `icons/` | Thông tin để cài lên màn hình chính |
| `sw.js` | Lưu sẵn file để chạy offline. Sửa app thì tăng `VERSION` trong file này |
| `vendor/xlsx.full.min.js` | Thư viện SheetJS 0.18.5 (Apache-2.0) để tạo file Excel |

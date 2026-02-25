# UNETI Time Mapper Extension

Chrome Extension (Manifest V3) cho `https://sinhvien.uneti.edu.vn/*` với các chức năng:

- Tự động đổi `Tiết: x - y` thành `Giờ: HH:mm - HH:mm` theo bảng giờ.
- Hỗ trợ map `normal` và `practical` + rule match 4 trường (Môn, Lớp, Phòng, GV).
- Hỗ trợ **thêm lịch thực hành bổ sung** (không cần có sẵn trên hệ thống), tự xếp đúng ô thứ/ngày trong bảng tuần.
- Hỗ trợ **import/export riêng lịch thực hành bổ sung** từ popup.
- Hiển thị thương hiệu + logo + bản quyền `TranDangKhoaTechnology`.
- Icon extension (toolbar + extension list) dùng logo tại `assets/image/logo.png`.
- Popup cấu hình JSON, Import/Export, khôi phục mặc định.

## Cài đặt

1. Mở Chrome tại `chrome://extensions`.
2. Bật `Developer mode`.
3. Chọn `Load unpacked`.
4. Trỏ tới thư mục: `d:\Downloads\Extention-LichHoc`.

## Sử dụng nhanh

1. Vào trang lịch tuần: `https://sinhvien.uneti.edu.vn/lich-theo-tuan.html`.
2. Extension tự đổi dòng `Tiết` sang `Giờ`.
3. Muốn thêm lịch thực hành bổ sung:
   - Cách 1: Bấm nút `Thêm lịch TH` trên trang lịch tuần.
   - Cách 2: Mở popup extension và điền mục `Lịch thực hành bổ sung`.
4. Chọn khung giờ thực hành cố định:
   - `06:00 - 11:30` (Sáng)
   - `13:00 - 18:30` (Chiều)
5. Import/Export lịch thực hành bổ sung:
   - Trong khối `Lịch thực hành bổ sung` bấm `Import lịch TH` hoặc `Export lịch TH`.
   - File import hỗ trợ 1 trong 2 dạng:
     - Mảng bản ghi: `[{...}, {...}]`
     - Object chứa mảng: `{ "manualPracticalSchedules": [...] }`

## Cấu hình

- Key lưu local: `unetiScheduleConfigV1`.
- Mặc định đọc từ `config/schedule-default.json` nếu chưa có cấu hình lưu.
- Có các nút:
  - `Lưu cấu hình`
  - `Đồng bộ từ JSON`
  - `Khôi phục mặc định`
  - `Import JSON`
  - `Export JSON`

## Ghi chú

- Nếu một dòng đã là giờ cụ thể (ví dụ `13h00-18h30`) thì extension giữ nguyên, không chuyển đổi.
- Lịch thực hành bổ sung chỉ hiển thị trên trang lịch tuần.

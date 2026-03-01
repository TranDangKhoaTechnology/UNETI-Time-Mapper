# UNETI Time Mapper Extension

Chrome Extension (Manifest V3) cho `https://sinhvien.uneti.edu.vn/*` với các chức năng:

- Tự động đổi `Tiết: x - y` thành `Giờ: HH:mm - HH:mm` theo bảng giờ.
- Hỗ trợ map `normal` và `practical` + rule match 4 trường (Môn, Lớp, Phòng, GV).
- Hỗ trợ **thêm lịch thực hành bổ sung** (không cần có sẵn trên hệ thống), tự xếp đúng ô thứ/ngày trong bảng tuần.
- Hỗ trợ **import/export riêng lịch thực hành bổ sung** từ popup.
- Hỗ trợ **xuất lịch học đa định dạng**: `XLSX`, `CSV`, `JSON`, `HTML`, `PNG`, `PDF`.
- Có thể xuất từ **popup** hoặc mở **panel xuất ngay trên trang lịch**.
- Hỗ trợ **xuất bảng điểm + dự đoán GPA** từ trang `Kết quả học tập` với `XLSX`, `CSV`, `JSON`.
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
6. Xuất lịch học:
   - Trên trang lịch tuần: bấm `Xuất lịch`.
   - Trong popup: dùng khối `Xuất lịch học` để xuất trực tiếp hoặc bấm `Mở panel trên trang`.
   - Hỗ trợ phạm vi `Theo tuần đang xem` hoặc `Theo tháng`.
7. Xuất điểm / GPA:
   - Mở trang `https://sinhvien.uneti.edu.vn/ket-qua-hoc-tap.html`.
   - Mở popup extension, dùng khối `Xuất điểm / GPA`, chọn `XLSX/CSV/JSON`.
   - Trên chính website có nút `Bật sửa điểm inline`: click trực tiếp vào ô điểm trên bảng gốc để sửa, `Enter` lưu, `Esc` hủy ô đang sửa.
   - Quy trình dự đoán điểm/GPA trực tiếp trên web:
     - Dự đoán nhanh: sửa trực tiếp ô `Điểm tổng kết` (hoặc các ô thành phần/điểm thi) của môn cần thử.
     - Nếu cần tính lại theo nhóm môn: giữ `Ctrl/Cmd + click` để chọn nhiều môn, bấm `Tính TB thường kỳ`.
     - GPA học kỳ + GPA tích lũy cập nhật realtime ngay sau khi bạn sửa.
     - `Điểm chữ`, `Xếp loại`, trạng thái `Đạt/Không đạt` và màu đỏ/bình thường của môn sẽ tự đồng bộ theo điểm dự đoán hiện tại.
     - Bấm `Khôi phục tất cả` để quay về dữ liệu gốc của tab hiện tại.
   - Dữ liệu dự đoán chỉ lưu tạm trong tab đang mở (reload trang sẽ về mặc định).
   - Export sau khi đã chỉnh sửa sẽ lấy đúng dữ liệu predicted hiện tại trong tab.
   - File XLSX có 4 sheet:
     - `Grades`: bảng gọn để sửa nhanh `Total10 (Edit)` + công thức.
     - `Grades_Detail`: bảng chi tiết đầy đủ điểm thành phần.
     - `GPA`: tổng hợp GPA theo kỳ và tích lũy (Current/Predicted).
     - `Lookup`: bảng quy đổi điểm hệ 10 -> hệ 4.

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

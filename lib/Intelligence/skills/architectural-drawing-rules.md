---
name: architectural-drawing-rules
description: What each architectural drawing type must show to count as complete - floor plan, ceiling plan, elevation, section, enlarged plan, detail, door and finish schedules
triggers: ban ve kien truc, hoan thien mat bang, hoan thien mat dung, hoan thien mat cat, mat bang kien truc, mat bang tran, mat dung kien truc, mat cat kien truc, mat bang phong to, chi tiet cau tao, chi tiet kien truc, thong ke cua, bang hoan thien, architectural drawing, floor plan, ceiling plan, reflected ceiling plan, elevation drawing, section drawing, enlarged plan, door schedule
agents: revit_action, revit_data, general
tools: get_current_view_elements, get_current_view_info, tag_all_rooms, tag_elements, create_dimension, create_text_note, manage_view, apply_view_template, create_schedule, get_schedule_data
standard: project
---
# Rule hoàn thiện bản vẽ KIẾN TRÚC theo loại

> **⚠️ Nguồn chuẩn** — Nội dung "phải có" dưới đây là thông lệ ngành; mã cửa, mã vật liệu, tỷ lệ, cỡ chữ thật theo BEP / tiêu chuẩn dự án (hồ sơ thiết kế kiến trúc: tham khảo TCVN 5671:2012). Không tự đặt mã khi dự án đã có quy ước — xem model trước.

Dùng sau quy trình chung `drawing-completion-rules`. Thiếu mục nào → ghi đúng mục đó vào báo cáo.

## Mặt bằng tầng (Floor plan) — 1:100 / 1:50
- Trục đủ hai phương, đầu trục ở cả hai đầu; cốt sàn (+0.000) và cốt các khu chênh (WC, ban công, sảnh).
- Kích thước ngoài nhà 3 lớp: tổng – trục – lỗ mở / mảng tường. Kích thước trong: phòng, vách ngăn, lỗ mở chính, định vị thiết bị cố định.
- Tag phòng: tên + số phòng (+ diện tích nếu dự án yêu cầu). Tag cửa đi / cửa sổ theo mã loại (ví dụ D1, W1) khớp bảng thống kê cửa.
- Ký hiệu: hướng Bắc, vết mặt cắt, ký hiệu mặt đứng, callout chi tiết; thang / ram dốc có mũi tên và chữ LÊN / XUỐNG; độ dốc sàn WC, ban công, sân thượng.
- Vật liệu hoàn thiện sàn qua legend hoặc tag sàn, không ghi chú tay lặp lại.
- Tầng điển hình: ghi rõ phạm vi áp dụng (ví dụ "MẶT BẰNG TẦNG 3–10").

## Mặt bằng trần (RCP) — cùng tỷ lệ mặt bằng
- Cao độ trần từng khu (tag trần), vật liệu trần, khe / lỗ thăm trần.
- Đèn, miệng gió, đầu báo, đầu phun (từ link MEP) hiển thị để phối hợp; lưới trần căn theo trục hoặc tâm phòng.

## Mặt đứng (Elevation) — 1:100 / 1:50
- Trục đầu, trục cuối (và trục chính); cao độ các tầng, đỉnh mái, đỉnh tường chắn mái, cốt sân.
- Ghi chú vật liệu hoàn thiện mặt ngoài (theo legend / keynote); cửa đúng mã.
- Kích thước chiều cao: tổng và từng tầng.

## Mặt cắt (Section) — 1:100 / 1:50
- Cắt qua thang và các khu chênh cốt.
- Cao độ: sàn hoàn thiện, đáy dầm / trần, đỉnh mái; kích thước chiều cao tầng, chiều cao thông thủy.
- Cấu tạo sàn / mái / tường: ghi chú lớp vật liệu hoặc dẫn chiếu chi tiết.
- Tên phòng bị cắt qua.

## Mặt bằng phóng to (WC, thang, bếp) — 1:50 / 1:20
- Định vị thiết bị vệ sinh từ trục hoặc mặt tường hoàn thiện; cốt sàn, độ dốc, phễu thu.
- Thang: số bậc, kích thước bậc (rộng × cao), bề rộng vế, chiếu nghỉ, cao độ chiếu nghỉ, tay vịn.

## Chi tiết cấu tạo (Detail) — 1:20 → 1:1
- Mỗi lớp vật liệu có ghi chú (tên + chiều dày); kích thước đầy đủ; hatch đúng vật liệu.
- Có dẫn chiếu ngược về view gốc (callout) và số chi tiết không trùng trên sheet.
- Drafting view 2D không được mâu thuẫn với model; dùng nhiều nơi thì ghi "CHI TIẾT ĐIỂN HÌNH".

## Bảng thống kê cửa / phòng / hoàn thiện
- Lấy dữ liệu từ model (`create_schedule`, `get_schedule_data`), không gõ tay; đặt trên sheet.
- Mã cửa trên mặt bằng = mã trong bảng; số lượng đọc từ `row_count`, không tự cộng.

## Kiểm nhanh bằng MCP (view đang active)
- `get_current_view_info` → tỷ lệ, detail level đúng loại bản vẽ?
- `get_current_view_elements(category="Rooms")` so với `category="RoomTags"`; tương tự `Doors`/`DoorTags`, `Windows`/`WindowTags` → thiếu tag?
- `get_current_view_elements(category="Grids")` + `category="Dimensions"` → có trục mà chưa có dim?
- Chỉ tag khi view chưa có tag loại đó (tool tag không bỏ qua phần tử đã tag — xem `drawing-completion-rules`).

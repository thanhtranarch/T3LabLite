---
name: drawing-completion-rules
description: Definition of done for a drawing sheet and the step-by-step workflow to complete sheets (views, scale, annotation, title block, revision) with T3Lab tools
triggers: hoan thien ban ve, hoan thien sheet, hoan chinh ban ve, hoan chinh sheet, hoan thien ho so ban ve, trien khai ban ve, kiem tra sheet, soat ban ve, ban ve da xong chua, finish drawing, finish sheet, complete drawing, complete sheet, drawing completion, sheet check, check sheet
agents: revit_action, revit_data, general
tools: revit_list_sheets, revit_list_views, revit_get_project_info, ai_element_filter, get_all_parameters, set_active_view, get_current_view_info, get_current_view_elements, manage_view, apply_view_template, add_view_to_sheet, place_views_on_sheets, duplicate_view, create_dimension, tag_all_rooms, tag_elements, create_text_note, manage_revision
standard: project
---
# Quy tắc hoàn thiện bản vẽ (sheet) — T3Lab

> **⚠️ Nguồn chuẩn** — Trang này quy định QUY TRÌNH và tiêu chí "xong". Mọi giá trị cụ thể (mã, tiền tố, cỡ chữ, trường khung tên…) lấy từ BEP / tiêu chuẩn của dự án. Chưa có BEP → theo quy ước đang có trong model và nói rõ là suy ra từ model. Không tự bịa mã dự án, số sheet, mã revision.

Đọc kèm: `drawing-conventions` (tỷ lệ, nét, chữ, ký hiệu), trang theo loại bản vẽ (`architectural-drawing-rules`, `structural-drawing-rules`, `mep-drawing-rules`), `annotation-standard`, `sheet-naming-standard`.

## 1. Thế nào là một sheet "XONG"
Chỉ báo hoàn thiện khi đủ cả 8 điều:
1. **Số + tên sheet** đúng chuẩn, không trùng, không còn tên tạm ("Copy", "test", "nháp").
2. **Khung tên** đủ thông tin: dự án, hạng mục, tên bản vẽ, số bản vẽ, tỷ lệ, ngày, người vẽ / kiểm / duyệt, giai đoạn, lần sửa.
3. **Từng view trên sheet**: có tiêu đề view, tỷ lệ đúng loại bản vẽ, có view template, crop gọn; viewport nằm trong khung, không chồng nhau, không đè khung tên.
4. **Định vị**: trục và cao độ hiện đủ, có đầu ký hiệu; mặt bằng có hướng Bắc; ký hiệu mặt cắt / mặt đứng / chi tiết trỏ đúng sheet.
5. **Kích thước** đủ chuỗi theo loại bản vẽ; không dim trùng, không dim đè chữ.
6. **Tag** cho mọi phòng, cửa, cấu kiện cần đọc; nội dung lấy từ dữ liệu model, không gõ tay.
7. **Ghi chú chung / legend** cần thiết đã có (ký hiệu, vật liệu, ghi chú kỹ thuật).
8. **Revision**: chỗ thay đổi so với lần phát hành trước có đám mây, revision đã gán vào sheet.

## 2. Quy trình: đọc → kế hoạch → làm → kiểm
1. **Kiểm kê (chỉ đọc)**
   - `revit_list_sheets`; `revit_get_project_info` (thông tin dự án cho khung tên).
   - `ai_element_filter(category="Viewports", fields=["id","name","view"])` → viewport nào nằm trên sheet nào (`view` = sheet chứa nó).
   - `get_all_parameters(element_id=<id sheet>)` → trường khung tên (Drawn By, Checked By, Approved By, Sheet Issue Date…).
   - Từng view: `set_active_view` → `get_current_view_info` (tỷ lệ, detail level, crop) → `get_current_view_elements(category=…)` với `Grids`, `Levels`, `Dimensions`, `RoomTags`, `DoorTags`, `WindowTags`, `TextNotes`.
2. **Báo bảng thiếu sót TRƯỚC khi sửa**: | Sheet | View | Thiếu gì | Sửa bằng tool nào | — chờ người dùng đồng ý một lần.
3. **Sửa theo đúng thứ tự** (sai thứ tự là phải làm lại):
   1. Template + tỷ lệ: `apply_view_template`, `manage_view` (`set_scale`, `set_detail_level`, `set_crop`) — chốt TRƯỚC khi dàn trang, tag, dim.
   2. Đặt view: `add_view_to_sheet` (x, y tính bằng mm trên sheet) hoặc `place_views_on_sheets`.
   3. Dim trục: `create_dimension`.
   4. Tag: `tag_all_rooms`, `tag_elements(category=…)`.
   5. Ghi chú chung: `create_text_note(sheet_number=…)`.
   6. Revision: `manage_revision` (`create`, `assign_to_sheets`).
4. **Kiểm lại bằng tool đọc** (đếm lại tag / dim / viewport) rồi mới báo.

## 3. Giới hạn tool — bắt buộc biết
- `tag_all_rooms`, `tag_all_walls`, `tag_elements` tag **mọi** phần tử trong **view đang active** và KHÔNG bỏ qua phần tử đã có tag → chạy lần hai là tag trùng. Luôn đếm tag hiện có trước (`get_current_view_elements(category="RoomTags")`…). View đã tag một phần → đề nghị người dùng dùng lệnh **Tag All** của Revit (chỉ tag phần chưa tag), hoặc xin phép chạy code chỉ đọc + tag phần còn thiếu.
- Tool tag / dim / text làm việc trên view đang active → `set_active_view` trước từng view.
- `create_dimension` chỉ dim giữa **các grid** (≥ 2). Dim tường, lỗ mở, cao độ: liệt kê chỗ thiếu cho người dùng, không giả vờ đã làm.
- Một view (trừ legend, schedule) chỉ nằm trên **một** sheet. Cần lần nữa → `duplicate_view` (`dependent` cho mặt bằng chia khu).
- Không có category tag ống / ống gió / máng cáp riêng trong tool → không đếm chính xác được tag MEP; kiểm trực quan hoặc bằng code chỉ đọc.
- Kiểm tra sâu (view chưa lên sheet, viewport chồng nhau, ký hiệu dẫn chiếu trống): nếu có tool `send_code_to_revit` (ứng dụng AI qua MCP) thì chạy code CHỈ ĐỌC, `description` ghi rõ "chỉ đọc" — người dùng được hỏi trước; nếu không có thì liệt kê mục cần người dùng tự kiểm.

## 4. Không bao giờ
- Không đổi số sheet đã phát hành khi chưa có revision và sự đồng ý của người dùng.
- Không xoá view / sheet / tag / dim trong lượt "hoàn thiện" — chỉ đề xuất.
- Không override đồ hoạ từng phần tử để "làm đẹp"; sửa view template hoặc báo người dùng. Không sửa tham số do template khoá.
- Không dùng text note thay cho dữ liệu sai trong model — sai thì sửa model.
- Không báo "xong" khi còn mục chưa kiểm được: ghi "chưa kiểm: …, lý do: …".

## 5. Báo cáo cuối
| Sheet | Trạng thái (XONG / CÒN THIẾU) | Đã làm | Còn thiếu | Id phần tử |
Một dòng tổng: X/Y sheet đạt; các việc người dùng phải tự làm (kèm lý do tool không làm được).

---
name: drawing-issue-checklist
description: Read-only checklist before issuing a drawing set (PDF/DWG) - numbering, drawing index, title blocks, revisions, cross references, content, export
triggers: phat hanh ban ve, phat hanh ho so, truoc khi phat hanh, kiem tra truoc khi xuat, kiem tra bo ban ve, nop ho so, giao nop ban ve, xuat ho so phat hanh, issue drawings, drawing issue, issue sheets, pre-issue check, transmittal
agents: revit_data, qa_check, export, general
tools: revit_list_sheets, revit_list_views, ai_element_filter, get_all_parameters, manage_revision, get_schedule_data, get_model_warnings
standard: project
---
# Kiểm tra bộ bản vẽ trước khi phát hành (T3Lab)

> **⚠️ Nguồn chuẩn** — Quy định giao nộp (mã trạng thái, cách đánh revision, tên file) lấy từ BEP / hợp đồng của dự án. Trang này chỉ là trình tự kiểm. Không tự bịa mã revision hay mã trạng thái.

Chạy toàn bộ ở chế độ **CHỈ ĐỌC**, gom thành MỘT báo cáo. Không sửa, không xuất file khi người dùng chưa xác nhận.

## 1. Danh mục & đánh số
- `revit_list_sheets`: không trùng số, không nhảy số bất thường, không còn sheet tạm (tên chứa "Copy", "test", "nháp", "WIP").
- Danh mục bản vẽ (sheet list) trên sheet bìa phải khớp danh sách sheet thật — nếu là schedule: `get_schedule_data`, so `row_count` với số sheet sẽ phát hành.
- Sheet rỗng (không viewport, không schedule): `ai_element_filter(category="Viewports", fields=["view"])` → sheet nào không xuất hiện trong cột `view` (là TÊN sheet) là sheet chưa có view (schedule đặt trên sheet không phải viewport — kiểm riêng).
- Thứ tự sheet theo trang `drawing-set-structure`.

## 2. Khung tên & revision
- Mỗi sheet phát hành: ngày phát hành, người vẽ / kiểm / duyệt, giai đoạn — `get_all_parameters(element_id=<id sheet>)`; thông tin dự án — `revit_get_project_info`.
- `manage_revision(operation="list")`:
  - Revision của đợt này đã có, đúng ngày, đúng mô tả, đúng người phát hành / nơi nhận.
  - Đã gán vào đúng các sheet thay đổi (`assign_to_sheets`) — sheet có đám mây revision mà chưa có revision là lỗi.
  - Đợt phát hành trước đã đánh dấu `set_issued` (Revit khoá đám mây của đợt đó).
- Revision chỉ tăng, không sửa lùi revision đã phát hành.

## 3. Dẫn chiếu chéo
- Ký hiệu mặt cắt / mặt đứng / callout phải hiện số chi tiết + số sheet. View được dẫn chiếu mà chưa đặt lên sheet thì ký hiệu không có số → lỗi "Chặn".
- View chưa lên sheet và không được dẫn chiếu: chỉ đề xuất dọn (skill `view-template-standard`), không xoá.
- Schedule cần phát hành phải nằm trên sheet.
- Ghi chú dạng "xem chi tiết 5/A-501" phải trỏ tới chi tiết có thật.

## 4. Nội dung
- Các sheet thay đổi trong đợt này đã qua `drawing-completion-rules`.
- Chính tả: bản vẽ tiếng Anh → skill `english-spellcheck`; tiếng Việt → thống nhất có dấu / không dấu theo dự án.
- Đơn vị thống nhất: kích thước mm, cao độ m (3 số lẻ), trừ khi dự án quy định khác.
- `get_model_warnings`: nêu các cảnh báo ảnh hưởng bản vẽ (phòng không kín / chưa đặt, phần tử trùng, tag phòng ngoài phòng).

## 5. Xuất file (sau khi người dùng xác nhận)
- Theo skill `export-standard`: `export_sheets_pdf`, `export_dwg`; khổ giấy theo khung tên; đen trắng trừ khi yêu cầu màu.
- Báo số file, thư mục, file lỗi.

## Báo cáo
| # | Hạng mục | Kết quả | Sheet / Id | Mức độ (Chặn / Cần sửa / Gợi ý) |
Kết luận một dòng: **SẴN SÀNG PHÁT HÀNH** hoặc **CHƯA** + danh sách mục "Chặn". Ghi rõ mục nào chưa kiểm được và vì sao.

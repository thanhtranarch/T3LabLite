---
name: mep-drawing-rules
description: What MEP drawings must show to count as complete - water supply, drainage, HVAC, electrical, fire protection plans, riser and schematic diagrams, coordination
triggers: ban ve mep, ban ve co dien, hoan thien mep, hoan thien ban ve mep, mat bang cap nuoc, mat bang thoat nuoc, mat bang dieu hoa, mat bang thong gio, mat bang dien, mat bang chieu sang, so do nguyen ly, so do truc dung, ban ve pccc, ban ve phong chay, mep drawing, plumbing plan, drainage plan, hvac plan, electrical plan, lighting plan, riser diagram, schematic diagram, fire protection drawing
agents: revit_action, revit_data, general
tools: ai_element_filter, get_current_view_elements, get_current_view_info, tag_elements, create_schedule, get_schedule_data, create_text_note
standard: project
---
# Rule hoàn thiện bản vẽ MEP (T3Lab)

> **⚠️ Nguồn chuẩn** — Lưu lượng, công suất, kích thước ống, khoảng cách bố trí (đầu phun, đầu báo…) do kỹ sư thiết kế quyết định theo tiêu chuẩn dự án áp dụng (cấp thoát nước, điều hoà, điện, PCCC). AI chỉ kiểm tra bản vẽ có ĐỦ và KHỚP thông tin, KHÔNG tự đặt thông số. Ký hiệu, màu, mã hệ theo legend / BEP của dự án.

Dùng sau quy trình chung `drawing-completion-rules`.

## Chung cho mọi hệ
- Nền kiến trúc / kết cấu (link) hiển thị nhạt (halftone / underlay), trục và tên phòng vẫn đọc được.
- Mỗi hệ có màu / nét / ký hiệu theo legend; legend đặt ở sheet đầu bộ môn hoặc trên từng sheet.
- Tag thiết bị: mã + thông số dự án yêu cầu (công suất, lưu lượng). Tag ống / ống gió / máng cáp: kích thước + cao độ.
- Ghi rõ quy ước cao độ ống trong ghi chú chung: BOP (đáy ống), CL (tâm ống), TOP (đỉnh ống).
- Hướng dòng chảy; vị trí và mã trục đứng (riser) — mã theo dự án.
- Phối hợp: không xung đột với dầm, cột, trần (skill `clash-coordination`); có mặt cắt phối hợp ở khu dày đặc (hành lang, trần kỹ thuật, hầm).

## Cấp nước — Thoát nước
- Đường kính danh nghĩa (DN) từng đoạn ống, vật liệu ống theo ghi chú chung.
- Thoát nước: độ dốc `i = …%` + mũi tên chiều dốc; cao độ đáy ống đầu và cuối đoạn; phễu thu, ống thông hơi, hố ga, thông tắc.
- Sơ đồ trục đứng: đủ các tầng, cao độ tầng, DN, van, đồng hồ, điểm đấu nối.

## Điều hoà — Thông gió (HVAC)
- Ống gió: kích thước W × H (mm) hoặc Ø; cao độ đáy ống gió; miệng gió có loại + lưu lượng; van gió; van chặn lửa tại vách / sàn ngăn cháy.
- Ống môi chất / nước lạnh, ống nước ngưng (có độ dốc); thiết bị (FCU, AHU, dàn nóng…) có mã khớp bảng thiết bị.

## Điện — Điện nhẹ (ELV)
- Chiếu sáng / ổ cắm: tag mạch (circuit) + tủ cấp nguồn; cao độ lắp công tắc, ổ cắm theo ghi chú.
- Máng cáp / thang cáp: kích thước W × H, cao độ; sơ đồ nguyên lý tủ (panel schedule) khớp các mạch trên mặt bằng.

## PCCC
- Đầu phun, họng nước vách tường, bình chữa cháy, đầu báo, nút ấn, chuông, đèn EXIT / sự cố: ký hiệu theo legend, đủ trên mọi tầng.
- Khoảng cách, bán kính bảo vệ, số lượng theo tiêu chuẩn PCCC dự án áp dụng — chỉ đối chiếu, không tự đặt.

## Sơ đồ nguyên lý (drafting view)
- Phải khớp model: số tầng, mã thiết bị, kích thước ống chính. Lệch giữa sơ đồ 2D và model → báo, không tự sửa.

## Kiểm nhanh bằng MCP
- Khối lượng theo hệ / tầng: `ai_element_filter(categories=["Pipes","Ducts","CableTrays","Conduits"], group_by=["category","type","level"])`.
- Thiết bị: `ai_element_filter(categories=["MechanicalEquipment","ElectricalEquipment","PlumbingFixtures","AirTerminals","Sprinklers","LightingFixtures","FireAlarmDevices"], group_by=["category","level"])` → đối chiếu bảng thiết bị (`get_schedule_data`, đọc `row_count`).
- Tool KHÔNG có category tag ống / ống gió / máng cáp riêng (chỉ `MultiCategoryTags`, `GenericModelTags`…) → không đếm chính xác tag MEP; kiểm trực quan. Không chạy `tag_elements` lên view đã có tag (tool không bỏ qua phần tử đã tag).

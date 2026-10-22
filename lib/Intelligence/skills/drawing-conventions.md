---
name: drawing-conventions
description: Construction drawing presentation wiki - paper sizes, which scale for each drawing type (site, plan, section, detail), line types and weights, text heights, common symbols, title block fields (ISO / TCVN references)
triggers: kho giay, ty le ban ve, chon ty le, ty le nao, net ve, do day net, loai net, chieu cao chu, co chu ban ve, khung ten, ky hieu ban ve, quy cach trinh bay, trinh bay ban ve, quy chuan trinh bay, drawing scale, paper size, line weight, line type, text height, title block fields, drawing symbols
agents: general, revit_action, revit_data
standard: project
---
# Wiki: Quy cách trình bày bản vẽ xây dựng

> **⚠️ Nguồn chuẩn** — Các số dưới đây là thông lệ ngành theo ISO / TCVN, dùng khi dự án chưa quy định. BEP / tiêu chuẩn công ty luôn được ưu tiên. Số hiệu tiêu chuẩn chỉ để tra cứu — kiểm tra phiên bản còn hiệu lực trước khi trích dẫn chính thức.

## Khổ giấy (dãy A — ISO 216; trình bày tờ vẽ — ISO 5457)
| Khổ | Kích thước (mm) |
|---|---|
| A0 | 841 × 1189 |
| A1 | 594 × 841 |
| A2 | 420 × 594 |
| A3 | 297 × 420 |
| A4 | 210 × 297 |
- Một bộ hồ sơ dùng thống nhất một khổ chính; trong Revit khổ giấy do family khung tên (title block) quyết định.

## Tỷ lệ (dãy khuyến nghị ISO 5455)
| Loại bản vẽ | Tỷ lệ thường dùng |
|---|---|
| Vị trí, tổng mặt bằng | 1:1000 · 1:500 · 1:200 |
| Mặt bằng, mặt đứng, mặt cắt tổng thể | 1:100 (công trình lớn 1:200) |
| Mặt bằng phóng to (WC, thang, bếp), mặt cắt tường | 1:50 · 1:20 |
| Chi tiết cấu tạo | 1:20 · 1:10 · 1:5 · 1:2 · 1:1 |
- Trong Revit, tỷ lệ view quyết định kích thước in của chữ, tag, ký hiệu → chốt tỷ lệ (`manage_view` `set_scale`) TRƯỚC khi tag, dim, dàn trang.
- View khác tỷ lệ trên cùng sheet → ghi tỷ lệ dưới từng tiêu đề view.

## Nét vẽ (ISO 128; bản vẽ xây dựng: TCVN 5570:2012)
- Dãy bề rộng nét (mm): 0,13 · 0,18 · 0,25 · 0,35 · 0,5 · 0,7 · 1,0 · 1,4 · 2,0. Tỷ lệ rất đậm : đậm : mảnh = 4 : 2 : 1.

| Loại nét | Dùng cho |
|---|---|
| Liền đậm | Đường bao bị cắt: tường, cột, sàn, dầm trong mặt bằng / mặt cắt |
| Liền mảnh / trung bình | Đường thấy phía sau mặt cắt, đường kích thước, đường dóng, đường dẫn, hatch |
| Nét đứt | Đường khuất, phần tử phía trên mặt phẳng cắt (dầm, ô thoáng), móng dưới đất |
| Chấm gạch mảnh | Trục, đường tâm, đường đối xứng |
| Chấm gạch, đậm ở hai đầu và chỗ đổi hướng | Vết mặt phẳng cắt |
- Trong Revit: độ dày nét theo Object Styles + view template; không override từng phần tử để "làm đậm".

## Chữ (ISO 3098)
- Dãy chiều cao chữ khi in (mm): 1,8 · 2,5 · 3,5 · 5 · 7 · 10 · 14 · 20.
- Thông lệ: ghi chú, chữ kích thước 2,5; tên phòng, tag 2,5–3,5; tiêu đề view 3,5–5; tên bản vẽ trong khung tên 5–7. Không dùng chữ dưới 1,8 mm.
- Cả bộ dùng 1–2 font. Chiều cao chữ là kích thước IN — annotation Revit tự nhân theo tỷ lệ view.

## Ký hiệu thường gặp (thông lệ)
- **Trục**: vòng tròn + số / chữ; số theo một phương, chữ theo phương kia, bỏ I và O (dễ nhầm với 1, 0).
- **Cao độ**: ký hiệu tam giác + giá trị bằng m, 3 số lẻ (+3.600, -1.200); ±0.000 thường là cốt hoàn thiện sàn tầng 1 — ghi rõ quy ước trong ghi chú chung.
- **Kích thước**: bằng mm, không ghi đơn vị; cao độ bằng m.
- **Mặt cắt / mặt đứng / chi tiết**: số hiệu + số sheet chứa view được dẫn chiếu.
- **Hướng Bắc** trên mọi mặt bằng; **key plan** khi công trình chia khu.
- **Độ dốc**: `i = 2%` kèm mũi tên chiều nước chảy.
- **Tên và số phòng, ký hiệu bộ phận nhà**: ISO 4157 (TCVN 6003).

## Khung tên (trường dữ liệu: ISO 7200; bản vẽ xây dựng: TCVN 5571:2012)
Trường thường có: chủ đầu tư · công trình – hạng mục – địa điểm · đơn vị thiết kế · tên bản vẽ · số / ký hiệu bản vẽ · tỷ lệ · ngày · giai đoạn thiết kế · lần sửa (revision) + ngày + nội dung · người thiết kế / vẽ / kiểm tra / chủ trì / chủ nhiệm / duyệt · số tờ.
- Trong Revit: trường dự án lấy từ Project Information (`revit_get_project_info`), trường bản vẽ lấy từ tham số sheet (`get_all_parameters`). Không gõ text đè lên khung tên.

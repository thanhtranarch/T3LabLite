---
name: structural-drawing-rules
description: What structural drawings must show to count as complete - general notes, foundation and framing plans, column, beam and slab details, Vietnamese rebar notation, bar schedule
triggers: ban ve ket cau, hoan thien ket cau, hoan thien ban ve ket cau, mat bang ket cau, mat bang mong, mat bang coc, mat bang dam san, chi tiet dam, chi tiet cot, chi tiet san, bo tri thep, thong ke thep, bang thong ke thep, ky hieu thep, ghi chu chung ket cau, structural drawing, structural plan, foundation plan, framing plan, rebar detail, bar schedule, rebar notation
agents: revit_action, revit_data, general
tools: ai_element_filter, get_current_view_elements, get_current_view_info, tag_elements, create_dimension, create_text_note, create_schedule, get_schedule_data
standard: project
---
# Rule hoàn thiện bản vẽ KẾT CẤU (T3Lab)

> **⚠️ Nguồn chuẩn** — Tính toán và yêu cầu cấu tạo (cấp bê tông, mác thép, lớp bảo vệ, neo, nối) do kỹ sư kết cấu quyết định theo hồ sơ tính toán và tiêu chuẩn dự án áp dụng. AI chỉ kiểm tra bản vẽ có ĐỦ và KHỚP thông tin, KHÔNG tự điền số liệu kỹ thuật. Số hiệu tiêu chuẩn chỉ để tra cứu — kiểm tra phiên bản còn hiệu lực.

Dùng sau quy trình chung `drawing-completion-rules`.

## Ghi chú chung kết cấu (sheet đầu bộ môn)
- Danh mục tiêu chuẩn áp dụng của dự án (ví dụ TCVN 5574:2018 — kết cấu bê tông và bê tông cốt thép).
- Vật liệu: cấp độ bền bê tông từng loại cấu kiện (B20, B25, B30… — ghi kèm mác tương đương nếu dự án dùng); thép theo TCVN 1651-1:2018 (thanh tròn trơn, ví dụ CB240-T) và TCVN 1651-2:2018 (thanh vằn, ví dụ CB300-V, CB400-V).
- Lớp bê tông bảo vệ theo loại cấu kiện, chiều dài neo / nối, vị trí nối — chép đúng từ hồ sơ thiết kế.
- Quy ước cao độ: ±0.000 tương ứng cao độ nào; đơn vị kích thước mm, cao độ m.

## Mặt bằng móng / định vị cọc — 1:100
- Trục + kích thước trục; định vị từng móng / cọc theo trục; tag mã móng (M1, M2… hoặc theo dự án).
- Cao độ đáy móng, đỉnh móng, đỉnh cọc; giằng móng có mã và tiết diện.
- Bảng thống kê móng / cọc khớp số lượng trên mặt bằng.

## Mặt bằng kết cấu dầm – sàn — 1:100
- Mã cấu kiện duy nhất: cột C…, dầm D…, sàn S…, vách V… (hoặc theo dự án); tag đọc từ tham số Mark / Type Mark, không gõ tay.
- Tag dầm có tiết diện b × h; ô sàn có chiều dày; cao độ đỉnh sàn / đỉnh dầm khi khác cao độ tầng.
- Lỗ mở sàn (định vị + kích thước), sàn hạ cốt, khe co giãn.
- Dim trục – trục; định vị cột, dầm lệch trục.

## Chi tiết cột / dầm / sàn — mặt cắt 1:20 (1:10), mặt đứng dầm 1:50 / 1:20
- Mặt cắt ngang: kích thước tiết diện, thép dọc (số lượng + Φ), thép đai (Φ + khoảng cách), số hiệu thanh, lớp bảo vệ.
- Ký hiệu thép thông dụng ở Việt Nam:
  - `2Φ20` — 2 thanh đường kính 20 mm.
  - `Φ8a200` — thép Φ8, khoảng cách 200 mm (đai, thép phân bố).
  - Số hiệu thanh ghi trong vòng tròn, trùng với dòng trong bảng thống kê thép.
- Dầm: vùng đai dày / thưa (ví dụ Φ8a100 gần gối, Φ8a200 giữa nhịp — số thật theo thiết kế), vị trí cắt và nối thép, thép gia cường chỗ dầm phụ gác lên.
- Sàn: phân biệt rõ thép lớp trên / lớp dưới; thép mũ ghi chiều dài vươn ra.

## Bảng thống kê thép (ISO 4066; bản vẽ thi công BTCT: TCVN 5572:2012)
- Cột tối thiểu: số hiệu · hình dạng và kích thước uốn · Φ (mm) · chiều dài 1 thanh · số thanh (1 cấu kiện × số cấu kiện) · tổng chiều dài · khối lượng (kg).
- Khối lượng 1 m thép ≈ 0,00617 × Φ² kg/m (Φ tính bằng mm, thép 7850 kg/m³): Φ10 ≈ 0,617 kg/m; Φ20 ≈ 2,47 kg/m.
- Số liệu lấy từ schedule / model thép (`get_schedule_data` → `row_count`, `column_totals`), không cộng tay.

## Kiểm nhanh bằng MCP
- Số cấu kiện theo loại / tầng để đối chiếu bảng thống kê: `ai_element_filter(categories=["StructuralColumns","StructuralFraming","StructuralFoundations"], group_by=["category","type","level"])`.
- Mark trống hoặc trùng: `ai_element_filter(category="StructuralFraming", parameter_name="Mark", fields=["id","type","param_value"])` → liệt kê `param_value` rỗng / lặp (tương tự cho cột, móng).
- Tag trong view đang active: `get_current_view_elements(category="StructuralFraming")` so với `category="StructuralFramingTags"`; cột: `StructuralColumns` / `StructuralColumnTags`. Chỉ `tag_elements` khi view chưa có tag loại đó.
- Thép: `ai_element_filter(category="Rebar", group_by=["type"])` để đối chiếu đường kính với bảng thống kê.

---
name: drawing-set-structure
description: Drawing set structure wiki - Vietnamese design stages and what each delivers, discipline order, sheet order inside a discipline, mandatory front sheets (index, general notes, legend)
triggers: bo ho so ban ve, cau truc ho so, cau truc bo ban ve, thu tu ban ve, thu tu sheet, sap xep sheet, danh muc ban ve, bia ho so, giai doan thiet ke, thiet ke so bo, thiet ke co so, thiet ke ky thuat, thiet ke ban ve thi cong, ban ve hoan cong, drawing set, drawing index, sheet order, design stage, as-built
agents: general, revit_data, revit_action
tools: revit_list_sheets, get_schedule_data
standard: project
---
# Wiki: Cấu trúc bộ hồ sơ bản vẽ

> **⚠️ Nguồn chuẩn** — Tiền tố bộ môn, cách đánh số và thứ tự sheet thật theo BEP / quy định của dự án; trang này là thông lệ để tham khảo. Giai đoạn thiết kế theo Luật Xây dựng và văn bản hướng dẫn — kiểm tra văn bản hiện hành trước khi trích dẫn.

## Giai đoạn thiết kế (Việt Nam)
| Giai đoạn | Thuộc hồ sơ | Bản vẽ thể hiện (tham khảo) |
|---|---|---|
| Thiết kế sơ bộ | Báo cáo nghiên cứu tiền khả thi | Phương án: mặt bằng, mặt đứng, phối cảnh, tỷ lệ nhỏ |
| Thiết kế cơ sở (TKCS) | Báo cáo nghiên cứu khả thi | Mặt bằng, mặt đứng, mặt cắt chính; giải pháp kết cấu và hệ thống kỹ thuật chính |
| Thiết kế kỹ thuật (TKKT) | Dự án thiết kế 3 bước | Đủ kích thước, vật liệu, thông số kỹ thuật chủ yếu để lập dự toán |
| Thiết kế bản vẽ thi công (TKBVTC) | Thi công | Đủ chi tiết để thi công: kích thước, cấu tạo, thống kê vật liệu |
| Hoàn công | Nghiệm thu, bàn giao | Đúng thực tế đã thi công; ghi rõ chỗ thay đổi so với TKBVTC |
- Mức chi tiết model theo giai đoạn: skill `lod-standard`. Đặt tên, trạng thái theo ISO 19650: skill `iso19650-naming`.

## Thứ tự bộ môn (thông lệ — tiền tố là ví dụ)
G (chung) → C (tổng mặt bằng, hạ tầng) → A (kiến trúc) → S (kết cấu) → M (điều hoà, thông gió) → P (cấp thoát nước) → E (điện, điện nhẹ) → FP (PCCC).

## Thứ tự sheet trong một bộ môn (thông lệ)
1. Bìa · danh mục bản vẽ · ghi chú chung · bảng ký hiệu và viết tắt
2. Tổng mặt bằng / mặt bằng định vị
3. Mặt bằng các tầng, từ dưới lên: hầm → tầng 1 → … → mái
4. Mặt bằng trần, mặt bằng hoàn thiện sàn
5. Mặt đứng → mặt cắt
6. Mặt bằng phóng to: thang, WC, khu đặc biệt
7. Chi tiết cấu tạo
8. Bảng thống kê (cửa, hoàn thiện, vật liệu, thép)
Đánh số theo nhóm (ví dụ 0xx chung, 1xx mặt bằng, 2xx mặt đứng / mặt cắt, 3xx chi tiết) — skill `sheet-naming-standard`.

## Sheet đầu bộ — bắt buộc
- **Danh mục bản vẽ**: số, tên, revision hiện hành, ngày — nên là Sheet List schedule để tự cập nhật, không gõ tay.
- **Ghi chú chung**: tiêu chuẩn áp dụng, đơn vị (mm, m), quy ước cao độ ±0.000, vật liệu chính.
- **Ký hiệu và viết tắt** dùng trong bộ (legend).

## Kiểm bằng MCP
- `revit_list_sheets` → sắp theo số, so với thứ tự trên; báo sheet sai nhóm, thiếu sheet đầu bộ.
- Danh mục là schedule → `get_schedule_data`, so `row_count` với số sheet thật.

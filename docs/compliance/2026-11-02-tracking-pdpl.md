# Rà soát pháp lý: usage tracking của T3Lab Lite (`lib/tracking/`)

**Ngày audit:** 02/11/2026
**Công cụ đối chiếu:** checklist của [vn-law-check](https://github.com/mhfed/vn-law-check) (cập nhật 02/11/2026) — mục 00, 01, 03, 07, 08.
**Phạm vi:** chỉ phần thống kê sử dụng gửi về `https://t3lab.space/api/revit/tracking`
(`lib/tracking/`, 3 call site: `startup.py`, `lib/_cpython_bootstrap.py`, `lib/core/server.py`).
Server t3lab.space không nằm trong repo này → các mục phía server chỉ nêu khuyến nghị.

> *Đây là rà soát sơ bộ do AI thực hiện, không phải tư vấn pháp lý. Số điều lấy từ checklist
> vn-law-check; mục ghi "CHƯA XÁC MINH" là chưa đối chiếu toàn văn. Xác minh lại tại
> thuvienphapluat.vn / vbpl.vn và tham vấn luật sư trước khi quyết định.*

## Phân loại (triage)

| Tín hiệu | Bằng chứng (trước khi sửa) |
|---|---|
| Theo dõi hành vi: mỗi lần mở Revit, bấm nút ribbon, gọi MCP tool | `lib/tracking/hooks.py` — `track_session_start`, `track_ribbon_click`, `track_mcp_call` |
| Bật mặc định, chỉ có opt-out trong file JSON | `service.py:102` — `if enabled is None: return True` |
| ID máy/người dùng = SHA-256 (không salt) của `hostname_username` | `service.py:119-129` |
| Gửi tên model đầy đủ | `service.py:151` — `doc_title = revit.doc.Title` → `payload["doc_name"]` |
| Gửi văn bản tự do (goal/prompt/description/query…) tới 250 ký tự | `service.py:172` |
| Server lưu IP + vị trí gần đúng (quốc gia, thành phố) | `README.md` → Network Traffic |
| Người dùng ở nhiều nước ("worldwide usage metrics") | docstring `service.py` |

→ **Mức 3 (thu thập dữ liệu hành vi quy mô, theo dõi liên tục).** Áp dụng: checklist 01 (đầy đủ),
03 (mục C), 08. Không áp dụng: TMĐT, bảo vệ người tiêu dùng, quảng cáo, giao dịch điện tử
(không bán hàng, không quảng cáo). Checklist 07 (AI): xem ghi chú cuối.

## Tổng quan

| Lĩnh vực | 🔴 | 🟡 | 🟢 |
|---|---|---|---|
| Bảo vệ dữ liệu cá nhân | 4 (3 đã sửa trong code, 1 cần chủ sở hữu xử lý) | 5 | 2 |
| An ninh mạng | 0 | 1 | 1 |

## Phát hiện

| # | Mức | Mục | Phát hiện | Căn cứ | Trạng thái |
|---|---|---|---|---|---|
| T01 | 🔴 | 01-B Đồng ý | Tracking bật mặc định, không hỏi người dùng: **im lặng bị coi là đồng ý**. Không lưu được thời điểm/nội dung đồng ý. | Luật 91/2025/QH15 Điều 9 (khoản 4 điểm d: im lặng không phải đồng ý); NĐ 356/2025/NĐ-CP (lưu bằng chứng đồng ý, số điều CHƯA XÁC MINH) | ✅ **Đã sửa** — `lib/tracking/consent.py`: hỏi 1 lần bằng TaskDialog ở Idling đầu tiên; chỉ gửi khi bấm *Allow*; Enter = *Don't allow*; đóng hộp thoại = không lưu gì, hỏi lại lần sau; lưu `value` + `version` + `at` (UTC). Bump `CONSENT_VERSION` khi đổi dữ liệu gửi → hỏi lại. |
| T02 | 🔴 | 01-A Minh bạch | Tài liệu ghi "Anonymous", "no sensitive PII collected", nhưng hash `hostname_username` không salt dò ngược được bằng từ điển, và đi kèm IP + thành phố → vẫn là **dữ liệu cá nhân** (bút danh hoá, không phải ẩn danh). | Luật 91/2025/QH15 Điều 4 (quyền được biết) | ✅ **Đã sửa** — ID là chuỗi ngẫu nhiên `uuid4` tạo trên máy (`tracking_install_id`); bỏ các câu "anonymous / no PII" trong code; README nói rõ là dữ liệu cá nhân bút danh. |
| T03 | 🔴 | 01-F Vượt mục đích | `doc_name` = `Document.Title`. Bản local của model workshared mặc định tên `<Central>_<WindowsUser>` → gửi **tên người dùng Windows**; tên dự án/khách hàng cũng là thông tin của bên thứ ba. Không cần cho mục đích "đếm công cụ được dùng". | Luật 91/2025/QH15 Điều 9 (đồng ý theo từng mục đích), Điều 7 khoản 5 | ✅ **Đã sửa** — chỉ gửi `doc_<sha256(install_id|title)[:12]>`; tên model không rời máy. |
| T04 | 🔴 | 01-B Dữ liệu cũ | Dữ liệu đã thu từ khi tracking bắt đầu được thu **không có đồng ý hợp lệ**, có tên model + hash cũ + IP/thành phố. | Luật 91/2025/QH15 Điều 9; NĐ 330/2026/NĐ-CP | ⏳ **Chủ sở hữu xử lý trên server** — xoá hoặc gộp (aggregate) các dòng cũ: bỏ `doc_name`, `user_hash`, IP; chỉ giữ số đếm theo tool/ngày/phiên bản Revit. |
| T05 | 🟡 | 01-F Tối thiểu hoá | Tham số MCP gửi gần như nguyên văn: văn bản tự do (prompt, mô tả của AI), giá trị tham số (có thể là tên người), thông báo lỗi chứa `C:\Users\<tên>\…`. | Luật 91/2025/QH15 Điều 9 | ✅ **Đã sửa** — `redact_arguments`: free text chỉ gửi độ dài; khoá chứa user/owner/author/email/phone/address/contact → `(not sent)`; mọi chuỗi còn lại (kể cả lồng nhau, `error_message`) bị lọc email, SĐT VN, thư mục user, cắt 300 ký tự. |
| T06 | 🟡 | 01-A Chính sách | Không có thông báo xử lý dữ liệu: loại dữ liệu, mục đích, thời gian lưu, kênh liên hệ, quyền, thời hạn phản hồi. | Luật 91/2025/QH15 Điều 4 khoản 1; NĐ 356/2025/NĐ-CP Điều 5 | ✅ **Đã thêm** `docs/PRIVACY.md` (link trong hộp thoại). ⏳ Chủ sở hữu điền các mục **[owner to confirm]**: tên pháp lý, nơi đặt server, thời gian lưu, IP có lưu không. |
| T07 | 🟡 | 01-B Rút đồng ý | Rút đồng ý chỉ làm được bằng cách sửa file JSON. | Luật 91/2025/QH15 Điều 10 | ⚠️ **Một phần** — có `tracking_enabled: false` / `tracking_consent.value: false` (hình thức điện tử). Đề xuất: thêm ô "Send usage statistics" trong MCP Control → Advanced (cần đi qua các gate UI/XAML). |
| T08 | 🟡 | 01-D Xuyên biên giới | Nếu `t3lab.space` đặt server ngoài Việt Nam (ví dụ Vercel/AWS nước ngoài) → là chuyển dữ liệu xuyên biên giới. | Luật 91/2025/QH15 Điều 20 (hồ sơ trong 60 ngày từ lần chuyển đầu); NĐ 356/2025 Điều 18 (Mẫu 09). Miễn trừ Điều 39 cho hộ KD/DN siêu nhỏ — áp dụng cho cá nhân: CHƯA XÁC MINH | ⏳ **Chủ sở hữu xác nhận** nơi host + diện miễn trừ; nếu không được miễn thì lập hồ sơ. |
| T09 | 🟡 | 01-A Dữ liệu nhạy cảm | Dữ liệu theo dõi hành vi trên dịch vụ trực tuyến được NĐ 356/2025 xếp vào nhóm nhạy cảm (số điều CHƯA XÁC MINH). Click ribbon/MCP gửi lên mạng có thể bị coi là loại này. | NĐ 356/2025/NĐ-CP | ⏳ Hỏi luật sư. Nếu đúng: không được hưởng miễn trừ Điều 39 và cần DPIA (Điều 21). |
| T10 | 🟡 | 03-C / 01-F Phía server | Server lưu IP đầy đủ + thành phố; chưa có thời hạn lưu, quy trình báo sự cố 72 giờ. | Luật 91/2025/QH15 Điều 23 khoản 1; Luật 116/2025/QH15 Điều 41 | ⏳ **Chủ sở hữu** — cắt IP (/24) hoặc chỉ lưu quốc gia; đặt thời hạn lưu (ví dụ 12 tháng); ghi quy trình báo sự cố. |
| T11 | 🟡 | 03-C Bảo mật | `tracking_url` cho phép đổi endpoint, kể cả `http://`. | Luật 116/2025/QH15 | Rủi ro thấp (chỉ người dùng tự sửa được). Giữ nguyên; có thể chặn `http://` sau. |
| T12 | 🟢 | 01-F | Không gửi code, secret, đường dẫn đầy đủ (đã có từ 2.2.0, giờ có test). | | ✅ |
| T13 | 🟢 | 01-F | Không mua bán, không chia sẻ cho bên thứ ba, không dùng cho quảng cáo. | Luật 91/2025/QH15 Điều 7 khoản 6, Điều 28 | ✅ (ghi trong PRIVACY.md) |
| T14 | 🟢 | 03-C | Gửi qua HTTPS, chạy nền, lỗi mạng không ảnh hưởng Revit. | | ✅ |

## Mức phạt tham khảo (chỉ cho mục 🔴)

Mức tối đa theo quy định, không phải mức chắc chắn sẽ bị phạt (checklist 08):

- Xử lý dữ liệu không có sự đồng ý: 30–50 triệu đồng; coi im lặng là đồng ý: 50–70 triệu đồng
  (NĐ 330/2026/NĐ-CP, hiệu lực 19/08/2026, số điều CHƯA XÁC MINH).
- Trần theo Luật 91/2025/QH15 Điều 8: hành vi khác tới 3 tỷ đồng (tổ chức); cá nhân 1/2 mức tổ chức;
  vi phạm chuyển dữ liệu xuyên biên giới tới 5% doanh thu năm liền trước.

## Ngoài Việt Nam

Người dùng EU: GDPR cũng yêu cầu đồng ý trước khi theo dõi hành vi (theo dõi hành vi người ở EU
đưa dịch vụ vào phạm vi GDPR). Cơ chế đồng ý mới (T01) đáp ứng hướng này; vẫn cần thông báo
đầy đủ (T06) và thời hạn lưu (T10).

## Ghi chú: Luật Trí tuệ nhân tạo (checklist 07)

Không thuộc phạm vi tracking. T3Lab Assistant hiển thị rõ là trợ lý AI và dùng nhà cung cấp/API
key do người dùng tự cấu hình (Luật 134/2025/QH15 Điều 11 khoản 1). Nếu sau này rà soát
riêng Assistant: kiểm tra nhãn nội dung AI tạo (Điều 11 khoản 3, NĐ 142/2026 Điều 18) và
thời hạn chuyển tiếp 12 tháng (Điều 35, tới 01/03/2027 cho hệ thống đã chạy trước 01/03/2026).

## Ảnh hưởng tới dashboard (do các sửa đổi)

- Số dòng sẽ giảm mạnh: chỉ máy đã bấm *Allow* mới gửi.
- `user_hash` đổi sang ID ngẫu nhiên → mỗi máy xuất hiện như người dùng mới một lần.
- `doc_name` là mã `doc_…` → vẫn đếm được số model khác nhau trên một máy, không còn tên.
- `purpose` của MCP không còn chứa văn bản AI/người dùng gõ, chỉ còn mô tả theo tool.

## Kiểm tra

```
python3 lite_guard/test_tracking_privacy.py   # 16 test: consent, ID, tên model, lọc dữ liệu
python3 lite_guard/test_mcp_server.py         # test cũ (không gửi code/secret/đường dẫn)
python3 lite_guard/guard.py                   # marker: cổng consent + doc_code không bị mất khi copy từ dev
```

## Giới hạn & khuyến nghị

- Đây là gợi ý rà soát ban đầu, không phải kết luận pháp lý. Luật dữ liệu cá nhân VN đang thay đổi
  nhanh; số điều đánh dấu CHƯA XÁC MINH cần tra lại văn bản gốc.
- Các mục ⏳ (T04, T06, T08, T09, T10) nằm ngoài repo — cần chủ sở hữu và, với T04/T08/T09, luật sư.
- Chưa kiểm tra được hộp thoại thật trong Revit (môi trường test không có Revit) — cần mở thử trên
  Revit 2022 và 2025+ trước khi phát hành.

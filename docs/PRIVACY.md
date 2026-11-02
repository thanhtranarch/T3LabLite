# T3Lab Lite — Usage statistics privacy notice

Thông báo xử lý dữ liệu cá nhân — thống kê sử dụng (Vietnamese summary at the end)

This notice covers **only the usage statistics** T3Lab Lite sends to
`t3lab.space`. Other network traffic (auto-update, the AI provider you set up,
cloud links) is listed in [README → Network Traffic](../README.md#network-traffic).

## 1. Who handles the data

- Controller: T3Lab — Trần Tiến Thanh (individual developer)
- Contact for every request below: **trantienthanh909@gmail.com**

## 2. Nothing is sent without your consent

On the first Revit start after installing or updating, T3Lab Lite asks
*"Help improve T3Lab Lite with usage statistics?"*.

| Your answer | What happens |
|---|---|
| **Allow usage statistics** | Records below are sent. Your answer, its time and the notice version are saved in `%APPDATA%\T3LabAI\mcp_paths.json` → `tracking_consent`. |
| **Don't allow** | Nothing is sent. You are not asked again for this version. |
| Close the dialog | Nothing is sent and nothing is saved. You are asked again on the next Revit start. |

When the data sent changes, the notice version goes up and you are asked again.

## 3. What is sent, and why

Purpose: to see which tools people use, which fail, and on which Revit versions,
so T3Lab can fix and improve them. **Not** used for advertising, not sold, not
shared with third parties for their own use.

| Data | Example | Notes |
|---|---|---|
| Event | `session_start`, ribbon button name, MCP tool name | |
| Panel / tool type | `Project`, `ribbon`, `mcp` | |
| Result | `success` / `error`, run time in ms | |
| Error message | `Cannot write C:\Users\(user)\…` | e-mails, phone numbers and user folder names removed; max 300 characters |
| MCP tool arguments | `view_name: Level 1`, `output_folder: Exports` | see "never sent" below |
| Revit version | `2025` | |
| Install ID | `3f9a0c…` (16 hex) | random, made on this computer; **not** your name or computer name. Delete `tracking_install_id` in `mcp_paths.json` to get a new one |
| Model code | `doc_1a2b3c4d5e6f` | a salted code of the model title; the title itself is **not** sent |
| Time | UTC timestamp | |
| Country | `VN` | worked out from the IP by the hosting provider (Vercel). The IP address and city are **not stored** |

**Never sent:** model geometry or data, code sent to Revit, full file paths,
prompts and other free text (goal, description, query, message, comment …),
fields that name a person (user, owner, author, e-mail, phone …), API keys,
passwords, tokens.

## 4. Where it is stored and for how long

- Server: `t3lab.space` runs on Vercel (region `iad1`, Washington D.C., USA);
  records are stored in Supabase (region `us-east-1`, USA). This is a transfer
  of personal data outside Vietnam (Luật 91/2025/QH15, Điều 20).
- The server drops the IP address and city before storing, removes e-mails and
  user folder names from every text field, and stores nothing from T3Lab Lite
  versions that never asked for consent.
- Kept for **12 months**, then deleted automatically every night.

## 5. Your rights

You can ask to know, see, correct or delete your data, restrict or object to
its processing, withdraw consent, and complain (Luật 91/2025/QH15, Điều 4).

- **Withdraw consent / stop sending:** set `"tracking_enabled": false` in
  `%APPDATA%\T3LabAI\mcp_paths.json` (always wins), or set
  `tracking_consent` → `"value": false`. Takes effect from the next record.
- **See or delete your data:** e-mail the contact above with your install ID
  (`tracking_install_id` in `mcp_paths.json`). Deadlines under Nghị định
  356/2025/NĐ-CP, Điều 5: acknowledge in 2 working days; give or correct data
  in 10 days; withdraw / restrict / object in 15 days; delete in 20 days.

## 6. Security

Records are sent over HTTPS. Only the T3Lab owner can read them: the
dashboard needs an admin password, and the database allows inserts only
(no public read). If the data leaks, T3Lab notifies the authority within 72
hours (Luật 91/2025/QH15, Điều 23) and posts a notice in this repository.

---

## Tóm tắt tiếng Việt

- T3Lab Lite **chỉ gửi thống kê sử dụng khi bạn bấm "Allow"**. Không trả lời
  = không gửi.
- Gửi: tên công cụ, kết quả, phiên bản Revit, một ID ngẫu nhiên của máy, mã
  hoá của tên model. Server chỉ lưu quốc gia, không lưu IP và thành phố.
- Lưu trên Vercel + Supabase tại Mỹ, giữ 12 tháng rồi tự xoá.
- Không gửi: model, code, đường dẫn đầy đủ, prompt/văn bản bạn gõ, email,
  số điện thoại, API key.
- Rút lại đồng ý: đặt `"tracking_enabled": false` trong
  `%APPDATA%\T3LabAI\mcp_paths.json`.
- Yêu cầu xem/xoá dữ liệu: gửi email kèm `tracking_install_id` tới
  trantienthanh909@gmail.com.

Căn cứ: Luật Bảo vệ dữ liệu cá nhân số 91/2025/QH15 (hiệu lực 01/01/2026),
Nghị định 356/2025/NĐ-CP. Đây không phải tư vấn pháp lý.

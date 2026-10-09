# 09 · Ribbon Icon Standard — T3 Space Line (style t3lab.space)

> Phạm vi: **toàn bộ icon trên ribbon — mọi `*.tab`** (một tab `T3Lab_Dev.tab`; 52 bundle + 4 bị khoá).
> **Khoá** đúng 4 bundle — 3 icon cloud link (logo hãng khác) + mascot T3LabAssistant: không ai được đụng vào (xem §7).
> Ngày: 2026-10-09. Thay chuẩn "đồng nhất với Revit" (2026-09-11 → 2026-10-09) — lịch sử ở §9.
> Phương án, lý do chọn hướng và concept từng icon: `dev/plan/ribbon-icon-t3space.md`.

**Trạng thái thực thi**

| Bước | Việc | Trạng thái |
|------|------|-----------|
| 1 | Đo style t3lab.space, chọn hướng (chủ extension giao toàn quyền quyết định) | ✅ xong |
| 2 | Gate + token cho style mới, chạy song song chuẩn cũ | ✅ xong |
| 3 | Vẽ lại **52/52** icon, mỗi panel một commit | ✅ xong |
| 4 | Chuyển hẳn: bỏ bảng màu và luật cũ, chuẩn này thay bản cũ | ✅ xong |
| — | **QA trong Revit thật** (§8) | ⬜ **chưa làm — cần anh mở Revit** |
| — | Đồng bộ icon lên t3lab.space | ⬜ sau khi QA Revit đạt |

Lệnh hằng ngày:

```
python3 dev/build_icons.py           # sinh dark + render PNG
python3 dev/build_icons.py --check   # không ghi gì, báo file nào lệch
python3 dev/audit_icons.py --quiet   # gate (gồm cả khoá 4 icon)
python3 dev/test_icons_t3space.py    # test từng luật của gate
```

Tài liệu này nói về **icon trên ribbon** (`icon.svg` / `icon.png` trong mỗi bundle).
Nó **không** thay thế mục "Icon" trong `T3LAB_UI_STANDARD.md` — mục đó nói về glyph
`Segoe MDL2 Assets` **bên trong cửa sổ WPF**.

---

## 1 · Vì sao

Chủ extension yêu cầu icon ribbon theo **style t3lab.space** — trang web của T3Lab.
Style đó đo từ source của repo `t3lab-space`:

| Đặc điểm | Bằng chứng | Thành luật |
|----------|-----------|-----------|
| Swiss International Style, đen trên trắng | `app/globals.css`: `--color-text-primary: #000`, `--color-accent: #000` | nét đen, không nền xám |
| Một màu nhấn | `#EA680C` (blockquote, hover link, thẻ bước) | đúng một chi tiết cam / icon |
| Icon nét | `RevitPanelIcon.tsx`: viewBox 24, `strokeWidth 1.5`, `fill none`, đầu/góc tròn | outline, nét 2 trên lưới 32 (= 1.5 / 24) |
| Xám phân cấp | `#666666` secondary | token `ink-2` |
| Khối 3D | icon `modeling-datum`: lục giác + chữ Y | glyph family / khối |

Icon T3Lab giờ **cố ý khác** icon Revit bên cạnh: nhận diện T3Lab trên ribbon thay vì
hoà vào Revit. Cái giá là nét 2 px nặng hơn nét 1 px của Revit — xem §6.

---

## 2 · Ngôn ngữ hình — luật

### 2.1 Lưới và hình học

| Luật | Giá trị | Gate |
|------|---------|------|
| Canvas | `viewBox="0 0 32 32"`, xuất PNG 64×64 (2×) | P0 |
| Dấu hiệu theo chuẩn | thẻ `<svg>` có `data-icon-style="t3space"` | P0 |
| Thuộc tính `<svg>` | `fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"` | P1 |
| Nét | **một** độ dày: 2 unit. Không đặt `stroke-width` khác ở phần tử con | P1 |
| Toạ độ | **số nguyên** ở mọi `rect` / `circle` / `line` / `path d` — nét 2 unit phủ `k±1`, rơi đúng pixel ở 64 và 32 px | P1 |
| Góc đường | 0° · 90° · 45° · dốc 2:1 (26.57° / 63.43°) và cung tròn `A`. Không bezier | P1 · P2 |
| Bo góc | `rx` chỉ `0` hoặc `2` (khung ≥ 12 unit dùng 2) | P1 |
| Tô | chỉ chi tiết `signal` được tô đặc. Chấm = đoạn dài 0 với đầu tròn (`M8 8h0`) hoặc cung bán kính 1 | P1 |
| Vùng vẽ | tâm nét trong `4–28` | review |
| Khoảng hở | hai nét **cùng token**: tier A ≥ 2 unit trống, tier B ≥ 3. Khác token được chạm nhau | review |
| Số shape | tier A ≤ 6, tier B ≤ 4 (một `<path>` nhiều đoạn tính 1) | P2 |
| Cấm | `opacity`, gradient, `filter`, `stroke-dasharray`, `<text>`, `<image>`, `<ellipse>` | P1 |

### 2.2 Bảng màu — 3 token (`dev/icons/tokens.json`)

| Token | Light | Dark | Dùng cho | Tương phản (`#EFEFEF` · `#4F4F4F`) |
|-------|-------|------|----------|-----------------------------------|
| `ink` | `#000000` | `#F2F2F2` | hình chính | 18.3 : 1 · 7.3 : 1 |
| `ink-2` | `#666666` | `#A3A3A3` | nội dung phụ: dòng chữ, đường gióng, ô con, lưới | 5.0 : 1 · 3.3 : 1 |
| `signal` | `#EA680C` | `#FF8A3D` | **một** chi tiết nhấn | 2.8 : 1 · 3.5 : 1 |

- Màu ngoài bảng = **P0**.
- `signal` dưới 3 : 1 trên nền sáng, nên **cam không bao giờ tự mang nghĩa**: icon phải
  đọc được ở thang xám (QA §8).
- Dark theme = **đổi màu nét theo bảng**, một lượt (`iconlib.to_dark`). Không silhouette,
  không vẽ tay. `icon.dark.svg` sinh tự động.
- Revit 2023 không có dark (`resolve_icon_file()` của pyRevit chỉ dùng `icon.dark.png`
  khi `HOST_APP.is_newer_than(2024)`) — bản light nét đen tự đứng được.

### 2.3 Chi tiết cam (`signal`)

| Luật | Giá trị | Gate |
|------|---------|------|
| Số lượng | đúng **1** phần tử dùng `signal` (stroke hoặc fill) | P1 |
| Diện tích | 2.5–8.5 % canvas, đo trên PNG đã build. Bộ 52 icon: 2.7–7.5 % | P2 |
| Tổng mực | 8–38 % canvas. Bộ 52 icon: 8.1–37.2 % | P2 |
| Ý nghĩa | thứ tool **làm ra / tác động lên**, hoặc huy hiệu hành động — không bao giờ là cả hình chính | review |

### 2.4 Huy hiệu hành động

Huy hiệu nằm **góc dưới-phải** (tâm nét trong ô `20–28`); hình chính **chừa trống** góc
đó — khung bị cắt hở, không vẽ đè.

| Động từ | Huy hiệu | Mẫu |
|---------|----------|-----|
| Tạo mới | `+` | SheetGen · FamiGen · CastUnit · MakePattern |
| Xuất ra | mũi tên ↘ ra khỏi hình | BatchOut · BVBSExport |
| Nhập vào | mũi tên ↖ vào hình | PDF import |
| Chuyển / chép | mũi tên → nối hai hình | FamiTransfer · CADToElements · PointCloud · Text to Element |
| Đồng bộ | mũi tên hai đầu ↔ | DatumSync · CropSync |
| Kiểm tra | dấu ✓ | IFC-SG · RebarCheck |
| Sức khoẻ | đường nhịp tim | ModelAuditor |

### 2.5 Một vật — một cách vẽ

| Đối tượng | Glyph | Mẫu |
|-----------|-------|-----|
| Sheet | khung đứng `rx 2` + 1–2 dòng `ink-2`, hở góc khi có huy hiệu | BatchOut · SheetGen |
| View | khung view + view title (vòng số hiệu + gạch chân) | ManaViews |
| Family / khối | lục giác + chữ Y, cạnh dốc 2:1 | FamiGen · ManaFami · FamiTransfer · PointCloud |
| Level · Grid | đường + bóng tròn | DatumSync |
| Thép | polyline góc tròn; mặt cắt thanh = chấm | BVBSExport · RebarCheck · RebarWizard |
| Bảng | khung + header + vạch cột `ink-2` | ManaSched |
| Workset / layer | tấm dốc 2:1 chồng nhau | ManaWorkset |
| File | chữ nhật gấp góc 45° | ManaDWG · PDF import |
| Chữ | vạch ngang `ink-2` | Feedback · Reference |

---

## 3 · Tier và đường ống render

Không đổi so với chuẩn cũ — đây là "vật lý" của pyRevit, không phải gu thẩm mỹ.

`pyrevit/coreutils/ribbon.py`: `ICON_LARGE = 32`, `ICON_MEDIUM = 24`. Nút top-level
của panel lấy `ICON_LARGE`; nút trong **stack / pulldown** lấy `ICON_MEDIUM`.
`create_bitmap()` giải mã PNG ở `icon_size * 2`, rồi WPF thu ½ lúc vẽ:

| Tier | Hiển thị | Đường đi của PNG 64 px | Nét 2 unit hiện ra |
|------|----------|------------------------|--------------------|
| **A** — top-level | 32 px | 64 → 64 → ½ | 2 px, đúng pixel nếu toạ độ nguyên |
| **B** — trong stack / pulldown | 24 px | 64 → 48 → ½ | 1.5 px, mềm |

- Render `icon.svg` ra PNG 64×64 bằng resvg (`dev/icons/render.js`) — deterministic:
  cùng SVG cho cùng byte trên Windows và Linux (đã so 2026-10-09).
- Không vượt 96 px (`check_icon_size()` của pyRevit cảnh báo). Nền trong suốt RGBA.
- `icon.svg` là nguồn **duy nhất**; `icon.dark.svg`, `icon.png`, `icon.dark.png` sinh
  bằng `python3 dev/build_icons.py`. Sửa tay sẽ bị build sau ghi đè, gate báo P0.

---

## 4 · Hạ tầng

| File | Vai trò |
|------|---------|
| `dev/icons/tokens.json` | bảng màu + ngưỡng — **nguồn duy nhất** |
| `dev/icons/iconlib.py` | quét bundle, phân tier, `style_of`, `to_dark`, `png_rgba` (giải mã PNG không cần Pillow) |
| `dev/icons/render.js` | resvg → PNG |
| `dev/build_icons.py` | sinh dark + PNG |
| `dev/audit_icons.py` | gate: `check_geometry` (§2.1–2.3) · `check_outputs` · `check_coverage` · `check_frozen` (§7) |
| `dev/test_icons_t3space.py` | mỗi luật một test |
| `dev/icons/frozen_icons.json` | sha256 của 4 icon bị khoá |

---

## 5 · Concept từng icon

Bảng 52 icon (hình chính + chi tiết cam) ở `dev/plan/ribbon-icon-t3space.md` §5. Ảnh
toàn bộ trước/sau, light/dark, đúng kích thước pyRevit vẽ:
`dev/plan/icon-t3space/preview.png`. Mỗi `icon.svg` có một dòng comment mô tả concept.

---

## 6 · Rủi ro

| Rủi ro | Xử lý |
|--------|-------|
| Nét 2 px **nặng hơn** icon Revit nét 1 px ở panel bên cạnh | Có chủ đích (nhận diện T3Lab). Bỏ nền nên tổng mực 8–37 % (chuẩn cũ 32–60 % tính cả nền). Nếu QA thấy quá nặng: hạ `stroke-width` ở `<svg>` + `stroke_widths` trong token — một chỗ cho cả bộ |
| Cam 2.8 : 1 trên nền sáng | Cam không tự mang nghĩa; QA thang xám |
| Tier B mềm nét ở 24 px | Khoảng hở tier B ≥ 3 unit, ≤ 4 shape |
| 4 icon bị khoá khác bộ | Chấp nhận — luật khoá (§7). Logo hãng và mascot khác bộ là đúng |
| Web và ribbon lệch nhau | Khi đồng bộ web: sinh từ chính `icon.svg`, không vẽ tay |
| Sao chép icon Lucide / Autodesk | Chỉ lấy *quy ước* (nét, lưới, đầu tròn); mọi đường nét vẽ mới trên lưới 32 |

---

## 7 · Miễn trừ = KHOÁ — đúng 4 bundle

> **Luật khoá — 2026-10-09, yêu cầu rõ ràng của chủ extension:** không đụng vào icon
> của 3 nút cloud link và icon T3Lab Assistant. Không sửa, vẽ lại, đổi style, thêm
> biến thể (`icon.svg`, `icon.dark.*`) hay xoá file icon của 4 bundle dưới đây — kể cả
> khi thiết kế lại **cả bộ** icon (như đợt T3 Space Line này).
>
> Thực thi bằng máy, không chỉ bằng chữ: `dev/icons/frozen_icons.json` giữ sha256 của
> từng file `icon*` trong 4 bundle; `dev/audit_icons.py` báo **P0** khi một file đổi,
> mất, hoặc có thêm file icon mới. `build_icons.py` bỏ qua 4 bundle này qua
> `iconlib.EXEMPT_BUNDLES`. Chỉ chủ extension được mở khoá — khi đó cập nhật hash trong
> cùng commit với icon mới.

| Bundle | Vì sao |
|--------|--------|
| `Support.panel/CloudLinks.stack/Autodesk Forma.urlbutton` | logo hãng khác (cloud link) |
| `Support.panel/CloudLinks.stack/Autodesk Health.urlbutton` | logo hãng khác (cloud link) |
| `Support.panel/CloudLinks.stack/Bluebeam Status.urlbutton` | logo hãng khác (cloud link) |
| `Support.panel/T3LabAssistant.pushbutton` | mascot sản phẩm |

Khoá theo **tên** thư mục bundle (duy nhất trong cả extension), không theo đường dẫn
panel: nút đổi tab / panel / stack thì khoá vẫn đi theo. Vẽ lại logo hãng là **vừa mất
nhận diện vừa đụng vào nhãn hiệu của họ** — người dùng tìm nút Forma bằng chính logo
Forma. T3LabAssistant là bề mặt trò chuyện, không phải một công cụ Revit — cùng lý do
`T3LabAssistant.xaml` bị khoá UI trong `CLAUDE.md`.

`dev/audit_icons.py` in danh sách 4 bundle này **mỗi lần chạy** (mục `KHOA`), kèm lý do,
và kiểm hash của chúng ở mọi lần chạy, kể cả `--quiet`.

---

## 8 · Checklist QA (cần Revit thật)

Reload pyRevit (Revit 2025+: **khởi động lại Revit**, xem `CLAUDE.md` rule 7) rồi kiểm
trên ribbon. Gate xanh **không** thay được mấy dòng này: audit đọc SVG, nó không biết
pyRevit thu ảnh xuống trông ra sao trên máy thật.

```
[ ] Revit 2023 (light) — cả 6 panel: nét sắc, không viền trắng quanh icon
[ ] Revit 2026 (light) — giống 2023
[ ] Revit 2026 (dark)  — nét sáng #F2F2F2, cam #FF8A3D nổi rõ trên nền tối
[ ] 35 icon tier B (stack / pulldown) ở 24 px: còn đọc ra được là cái gì
[ ] Đặt cạnh panel Revit gốc: nặng hơn nhưng không "đen kịt" — chấp nhận được?
[ ] Chụp ribbon, chuyển thang xám: vẫn phân biệt được từng tool khi mất màu cam
[ ] 125% display scaling: không vỡ nét
[ ] Hover / pressed (nền xanh nhạt): nét đen + cam vẫn rõ
[ ] Support panel: 7 tool T3Lab đã đổi; 3 logo cloud link + T3LabAssistant KHÔNG đổi (bị khoá, §7)
```

Dòng nào hỏng thì báo tên tool, sửa `icon.svg` rồi chạy lại `python3 dev/build_icons.py`.
Không sửa PNG hay `icon.dark.svg` bằng tay.

---

## 9 · Lịch sử

| Thời kỳ | Chuẩn | Ghi chú |
|---------|-------|---------|
| → 2026-09 | "Direction B" | khối đặc navy `#182A3E` + cam `#D07818`, lưới 64, có `opacity` |
| 2026-09-11 → 2026-10-09 | "Đồng nhất với Revit" | line-art đo từ `UIFrameworkRes.dll` (Revit 2026): nét 1 unit `#000000`, nền `#F3F3F3`, accent amber `#E07B00`, dark kiểu silhouette, góc vuông |
| 2026-10-09 → | **T3 Space Line** | chuẩn này |

Toàn văn chuẩn "đồng nhất với Revit" (số đo Revit, bù độ nhạt của nét khi WPF thu ½,
silhouette dark) nằm trong lịch sử git của file này.

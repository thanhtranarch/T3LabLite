---
name: construction-glossary
description: Vietnamese-English construction and drawing glossary mapped to Revit terms and to the category names T3Lab tools accept
triggers: thuat ngu xay dung, tu dien xay dung, tu vung xay dung, tieng anh xay dung, dich ban ve, dich thuat ngu, thuat ngu ban ve, thuat ngu revit, construction glossary, construction terms, revit terms, vietnamese terms
agents: general, revit_data, revit_action
---
# Wiki: Thuật ngữ xây dựng Việt – Anh – Revit

Dùng để hiểu đúng yêu cầu tiếng Việt và gọi đúng tool. Cột "Revit / tool" là tên category mà `ai_element_filter`, `tag_elements`, `get_current_view_elements` chấp nhận (hoặc loại view / tool tương ứng).

## Hồ sơ & bản vẽ
| Tiếng Việt | English | Revit / tool |
|---|---|---|
| Bản vẽ, tờ bản vẽ | Drawing, sheet | `Sheets` |
| Khung tên | Title block | Title block family |
| Danh mục bản vẽ | Drawing index, sheet list | Sheet List schedule |
| Mặt bằng | Floor plan | view `FloorPlan` |
| Mặt bằng trần | Reflected ceiling plan (RCP) | view `CeilingPlan` |
| Mặt bằng kết cấu | Structural / framing plan | view `EngineeringPlan` |
| Mặt đứng | Elevation | view `Elevation` |
| Mặt cắt | Section | view `Section` |
| Trích chi tiết, chi tiết | Callout, detail | view `Detail` |
| Bản vẽ 2D | Drafting view | view `DraftingView` |
| Chú giải ký hiệu | Legend | view `Legend` |
| Bảng thống kê | Schedule | `get_schedule_data` |
| Ghi chú | Text note | `TextNotes` |
| Ghi chú chung | General notes | text / legend |
| Chỉ dẫn kỹ thuật | Specifications | ngoài Revit |
| Lần sửa đổi | Revision | `manage_revision` |
| Đám mây sửa đổi | Revision cloud | `RevisionClouds` |
| Phát hành | Issue | — |
| Bản vẽ hoàn công | As-built drawings | — |
| Bản vẽ chế tạo, triển khai | Shop drawings | — |
| Tỷ lệ | Scale | `manage_view` (`set_scale`) |

## Giai đoạn thiết kế
| Tiếng Việt | English | Revit / tool |
|---|---|---|
| Thiết kế sơ bộ | Preliminary design | |
| Thiết kế cơ sở (TKCS) | Basic design | |
| Thiết kế kỹ thuật (TKKT) | Technical design | |
| Thiết kế bản vẽ thi công (TKBVTC) | Construction drawing design | |

## Định vị & ghi kích thước
| Tiếng Việt | English | Revit / tool |
|---|---|---|
| Trục, lưới trục | Grid | `Grids` |
| Cao độ, tầng | Level | `Levels` |
| Cốt ±0.000 | Datum level ±0.000 | |
| Cốt hoàn thiện / cốt kết cấu | Finished floor level (FFL) / structural slab level (SSL) | |
| Cao độ điểm | Spot elevation | |
| Kích thước | Dimension | `Dimensions` |
| Đường dóng / đường kích thước | Extension line / dimension line | |
| Hướng Bắc | North arrow | |
| Mặt phẳng tham chiếu | Reference plane | `ReferencePlanes` |
| Vùng phạm vi | Scope box | `ScopeBoxes` |

## Kiến trúc
| Tiếng Việt | English | Revit / tool |
|---|---|---|
| Tường | Wall | `Walls` |
| Sàn | Floor | `Floors` |
| Trần | Ceiling | `Ceilings` |
| Mái | Roof | `Roofs` |
| Cửa đi / cửa sổ | Door / window | `Doors` / `Windows` |
| Vách kính | Curtain wall | `CurtainPanels`, `CurtainWallMullions` |
| Cầu thang / lan can / ram dốc | Stair / railing / ramp | `Stairs` / `Railings` / `Ramps` |
| Phòng / diện tích | Room / area | `Rooms` / `Areas` |
| Tag phòng / cửa / cửa sổ | Room / door / window tag | `RoomTags` / `DoorTags` / `WindowTags` |
| Thiết bị vệ sinh | Plumbing fixture | `PlumbingFixtures` |
| Nội thất | Furniture | `Furniture` |

## Kết cấu
| Tiếng Việt | English | Revit / tool |
|---|---|---|
| Móng, cọc, đài cọc | Foundation, pile, pile cap | `StructuralFoundations` |
| Giằng móng | Tie beam, grade beam | `StructuralFraming` |
| Cột (kết cấu / kiến trúc) | Structural / architectural column | `StructuralColumns` / `Columns` |
| Dầm | Beam | `StructuralFraming` |
| Sàn BTCT | RC slab | `Floors` |
| Vách, lõi | Shear wall, core | `Walls` |
| Cốt thép, thép đai | Reinforcement, stirrup | `Rebar` |
| Lớp bê tông bảo vệ | Concrete cover | |
| Cấp độ bền bê tông | Concrete strength class | |
| Bảng thống kê thép | Bar schedule | schedule |

## MEP
| Tiếng Việt | English | Revit / tool |
|---|---|---|
| Ống cấp nước / thoát nước | Water supply / drainage pipe | `Pipes` |
| Phụ kiện ống / van | Pipe fitting / accessory | `PipeFittings` / `PipeAccessories` |
| Ống gió | Duct | `Ducts` |
| Miệng gió | Air terminal (diffuser, grille) | `AirTerminals` |
| Thiết bị cơ | Mechanical equipment | `MechanicalEquipment` |
| Máng cáp / ống luồn dây | Cable tray / conduit | `CableTrays` / `Conduits` |
| Đèn | Lighting fixture | `LightingFixtures` |
| Tủ điện | Electrical panel / equipment | `ElectricalEquipment` |
| Đầu phun chữa cháy | Sprinkler | `Sprinklers` |
| Đầu báo cháy | Fire alarm device | `FireAlarmDevices` |
| Trục đứng | Riser | |
| Sơ đồ nguyên lý | Schematic diagram | view `DraftingView` |

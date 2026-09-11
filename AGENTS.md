# AGENTS.md — Agent 工作室必读（本工作空间所有 Agent 的入口文档）

> ⚠️ **强制约定**：任何 Agent（含子 Agent、专家、自动化任务）在本工作空间接任务前，**必须先阅读本文件与 `docs/PROJECT.md`**（完整项目文档）。先读文档，再动代码。

---

## 项目一句话

**Odoo Excel Importer** — Chrome 扩展（Manifest V3），把**供应商数据**（供应商 PDF 或 PDF 转换版 Excel，**v3.8 上传时自动识别文件类型并路由解析**）批量写入 Odoo 采购订单（`purchase.order`）明细行（数量/包装数量/整箱批发价/备注）；并支持「商品库更新」：按 UPC 更新 Odoo 产品变体（`product.product`）的 `name`（中文简称）/ `brand`（品牌）。

## 必读文件（按顺序）

1. **本文件**（AGENTS.md）— 关键事实速查
2. **`docs/PROJECT.md`** — 权威完整文档：架构分层、两条业务流程、字段映射、UI、工程规范、已知待办
3. 需要时再查：`.workbuddy/memory/MEMORY.md`（历史决策记忆）、`GIT_WORKFLOW.md`（Git 规范）

## 关键事实速查

| 项 | 内容 |
|---|---|
| 代码位置 | **只有一个业务文件**：根目录 `content.js`（~3400 行，ES5 风格 IIFE，无构建步骤） |
| 第三方库 | `lib/xlsx.full.min.js`（SheetJS）。`lib/pdf.min.js`+`lib/pdf.worker.min.js`（pdfjs-dist 3.11.174 UMD，勿升 4.x）——**v3.2 曾移除加载，v3.8（2026-09-09）恢复**：content_scripts 加载 pdf.min.js（全局 `pdfjsLib`），pdf.worker.min.js 在 web_accessible_resources 供 worker 拉取 |
| 打包副本 | `odoowritting_extension/` 子目录 = Chrome「加载已解压」用产物；**改完根目录必须同步过去**（含 manifest 版本号） |
| 字段映射 | 集中在 content.js 顶部 `ODOO_FIELDS`（Odoo 字段名，**2026-08-14 用户确认**：包装数量=`product_packaging_qty`、整箱批发价=`box_wholesale_price`、单价=`price_unit`、备注=`remark`、订单关联=`name`；**v3.0 补读小计 `price_subtotal` 仅作比对**）与 `EXCEL_COLS`（Excel 表头名，**v3.0：数量列改「包装数量」**，原「abw交货箱数」作废），**改字段只改这两个常量**；商品库更新（v1.11.0）用 `PRODUCT_ODOO_FIELDS`（`default_code`/`brand`/`name`，用户确认均 Char）与 `PRODUCT_EXCEL_COLS`（UPC/品牌/中文简称） |
| 匹配键 | 转换版 Excel/PDF ↔ 采购单 Excel（v2.0，2026-08-20 用户需求重构）：**两级匹配** `matchPdfToExcel` —— 第一级 PDF `catalog`（CATALOG NO./Catalog#）↔ Excel「SKU」列（命中逻辑不变）；第二级对 SKU 匹配不上的行走 UPC：**整体统计两侧 UPC 出现次数，任一侧重复 → 包装拆分**（PDF 侧 description `(xN)` → `UPC+N`，如 `8809732911873+42`；Excel 侧包装列 `1 box of N pieces` → `UPC+N`），两侧唯一 → 直接 UPC 匹配；拆分失败归缺货。pair 带 `matchKey`（Catalog/UPC/拆分键）供数量核对聚合。**Excel↔Odoo（v1.9.0 + v3.7）：UPC + 包装双重匹配，两级键** —— 主键 = Excel「内部参考号」↔ `order_line/product_id/default_code`（default_code 空用 name `[数字]` 兜底）；**v3.7（2026-09-04，只读库实测：部分产品 default_code 改过号、票据 UPC 只在 `product.supplierinfo.product_code` 里）二级兜底键 = 按 PO 供应商（`partner_id`）过滤的 `supplierinfo.product_code`**（`searchPoByRef` 补读 partner_id → `getOrderLines` 查模板级+变体级 supplierinfo → 注册 `supMap/supCnt`，`buildPreviewRows` 主键未命中再按次键匹配）；Excel「订单行/包装」件数 ↔ `order_line/product_packaging_id`（qty）；包装对不上或同 UPC 多行无包装 → 报未找到不写回（防错配） |
| 两条流程 | ✅ **双 Tab（v3.8 起）**：🗂 **供应商数据导入**（步骤1 上传供应商数据——**.pdf → PDF 分支** `parsePdf`/`extractPdfTable`（v3.8 恢复，依赖 pdfjsLib）；**.xlsx/.xls → Excel 分支** `parseConvertedPdfExcel`；`detectSupplierKind` 按扩展名自动识别路由，两分支输出同构 `{format,coupon,rows}`，**可多个混合（v1.12.0）**——`mergeConvFiles` 合并去重：同 UPC+包装 Qty/Subtotal 相加、Coupon 任一非0 → 再传采购订单 Excel → `applyPdfByMode` → Modal 写回）+ 🏷 **商品库更新**（v1.11.0：上传商品库 Excel → `loadProductByUpc` 按 UPC=default_code 查 `product.product` → 预览（旧→新对照 + 勾选）→ 写回 `name`/`brand`；**不做新旧比对，匹配到即更新、空值也照写**，未匹配/UPC 重复行注明原因不写；详见 docs/PROJECT.md §6.4）。修正流共用 `parseOrderExcel` / `loadOrderLineMap` / `buildPreviewRows` / `buildPdfPreviewModal` / `executePdfUpdate` |
| 注入条件 | URL hash 含 `model=purchase.order` **或** `model=product.product`（v1.11.0 放宽；产品页默认商品库 Tab）的页面注入按钮 |
| 价格规则 | ⚠️ **v3.0（2026-08-26 用户需求）**：modal 价格**原值比对基准 = Odoo 订单行现值**（单价↔`price_unit`、0.9箱规价↔`box_wholesale_price`、0.9总价↔`price_subtotal`），**采购单 Excel 价格列不再作比对基准**；实现 = `applyPdf*` 价格字段 oldValue 占位 + `expr`，`buildPreviewRows` 命中 Odoo 行后回填重算 changed/reason（缺货行跳过）。**箱规价字段已移除**（Odoo 无对应字段）；整箱批发价计算 = PDF UNIT PRICE×0.9（`calcBoxPrice` 整数分运算）。**0.9总价动态列（v3.6，2026-09-04 用户需求）**：只读计算列，套装=0.9箱规价×包装数量、单件=单价×包装数量（单件入口新增此列）；modal 编辑 boxQty/价格 → `recalcRowTotal` 实时重算 + `refreshTotal` 刷新总和/组小计；仍比对 Odoo `price_subtotal`（不写回）；单件入口同插「订单关联 小计」行（详见 docs/PROJECT.md §6.2/#21）。**v3.6.1 补丁**：Odoo 匹配不上订单行的行价格也照常显示计算值（`fieldInitValue` newValue 优先），只是无 Odoo 原值可比对 |
| 缺货写回 | ⚠️ **v1.12.4**（取代 v1.8.0）：**所有缺货行**（①匹配不上 + ②包装数量空/0——**v3.0 由 abw交货箱数改为包装数量列判断**，Excel 入口统一）写回时除备注「缺货」外，**单价（price_unit）/ 整箱批发价（box_wholesale_price）置 0 写回**；单件入口缺货②追加「整箱批发价」字段（key=boxWholesale，v1.12.4 新增）。旧规则「Excel 入口缺货行价格不写回」已作废。**v3.9（2026-09-10）**：预览 Modal 里用户**主动把「包装数量」改成 0/空** → 同样走缺货逻辑（`syncRowOutstock`：价格置 0 并锁定、备注「缺货」、行标红、数量核对列显示缺货、自动勾选；改回非 0 按快照/calcValue 还原）；`executePdfUpdate` 写入时再判一次兜底 |
| 多选上传 | **v3.9（2026-09-10）**：供应商数据上传区支持原生多选（`makeDropZone(..., multi=true)`：input.multiple + 拖入多文件），逐个类型校验后**串行**解析（等上一个完成再下一个）；采购订单 Excel / 商品库上传区仍单选 |
| SKU 更新 | ⚠️ **v3.4（2026-08-27 用户需求，v3.3 修订）**：两级匹配中 **UPC 匹配成功**（SKU 匹配失败，`matchLevel=upc`）→ 认为 SKU 已变更，用**转换 Excel 的 SKU**（pdf.catalog）更新 `product.packaging`：**采购单 SKU 列按入口区分**（`parseOrderExcel(mode)`/`colIdxCatalog`）——单件入口用「订单行/包装/SKU」列（fuzzy 跳过含 Box 列）、套装入口用「订单行/包装/Box SKU」列；单件**不识别规格**直接定位 qty=1 记录写 `single_sku`，套装保留 XX 比对（description xN → "1 box of XX pieces"）写 `box_sku`；SKU 匹配成功不更新；Modal SKU 列旧→新对照、变更默认勾选、失败文案拆 5 种、找不到规格跳过并提示；写回 `product.packaging.write` 独立于订单行（详见 docs/PROJECT.md §6.5） |
| 套装扣减 | ⚠️ 已作废（v1.5.0 起改为求和比对 + modal 提示，不自动扣减） |
| ❌ 未实现 | 无重大未实现项；待办见 docs/PROJECT.md §9（Excel 公式列缓存值待真实样本验证、v1.7.0 待 Chrome 冒烟） |
| Git | Conventional Commits（husky+commitlint 强制），scope 建议 `excel`/`ui`/`import`/`config`/`pdf` |

## 强制约定

1. **先读 `docs/PROJECT.md` 再动手**，特别是 §9「已知问题与待办」——避免重复踩坑或重复实现
2. **改代码只改根目录 `content.js`**，改完同步 `odoowritting_extension/` 副本 + manifest 版本号
3. **代码风格保持 ES5**（var/function/IIFE），不用箭头函数/模板串；注释中文
4. **字段名/列名不硬编码**，改 `ODOO_FIELDS` / `EXCEL_COLS`
5. **用户输入展示必须 `escHtml()` 转义**（防 XSS）
6. 无自动化测试，验证 = Chrome 加载扩展手动冒烟（见 docs/PROJECT.md §8）
7. 提交信息过 commitlint：`type(scope): subject`

## 当前最高优先级待办

- [x] **Excel 导入入口改造**（v1.7.0 已实现）：吃「PDF 转换版 Excel」，比对逻辑与 PDF 流完全一致（格式 A/B + Coupon + 两步上传，详见 docs/PROJECT.md §6.3）
- [x] **Excel 入口缺货行不写回价格**（v1.8.0 已实现）：备注「缺货」的行跳过单价/整箱批发价写回，PDF 流不变（详见 docs/PROJECT.md §6.2）
- [x] **Odoo 匹配键改 UPC+包装**（v1.9.0 已实现）：UPC = `product.default_code`，包装 = `product_packaging_id.qty`，双重匹配防错配（详见 docs/PROJECT.md §6.2）
- [x] **PDF↔Excel 两级匹配重构**（v2.0 / manifest v1.10.0 已实现）：SKU 匹配不变 + UPC 包装拆分兜底（`extractDescPack` 提取 description `(xN)`、`extractPackQty` 提取包装列 pieces，拼 `UPC+N` 精确配对；唯一 UPC 直接匹配；拆分失败归缺货；数量核对按 matchKey 聚合），33 断言单测通过（详见 docs/PROJECT.md §6.2）
- [x] **商品库更新 Tab**（v1.11.0 已实现）：UPC 定位产品变体（`default_code`），写回 `name`（中文简称）/`brand`（品牌，均 Char）；**不做新旧比对、空值也照写**（用户 2026-08-21 二次确认）；状态 ok/notfound/dup，未匹配注明原因，预览勾选写回（详见 docs/PROJECT.md §6.4）
- [x] **数量列改「包装数量」+ 价格原值改比对 Odoo**（v3.0 / manifest v1.13.0 已实现）：数量比对与写回用采购单 Excel「包装数量」列（abw交货箱数列作废）；价格原值（单价→price_unit、0.9箱规价→box_wholesale_price、0.9总价→price_subtotal）从 Odoo 订单行取，`buildPreviewRows` 命中后回填重算 changed/reason；箱规价字段移除；缺货判断 = 包装数量空/0（详见 docs/PROJECT.md §6.2/#17）
- [x] **modal 按订单关联分组小计**（v3.1 / manifest v1.14.0 已实现）：套装入口表格内每组尾部插「{订单关联号} 小计」行，只对 0.9总价列做组内合计（全部行），随编辑实时刷新；单件入口（无 total09）不插；底部全表合计行保留（详见 docs/PROJECT.md §6.2/#18）
- [x] **modal 去掉整箱批发价列**（v3.1.1 / manifest v1.14.1 已实现）：fieldOrder/FIELD_LABELS/SUM_KEYS 移除 boxWholesale 列；**缺货行整箱批发价置 0 写回不受影响**（boxWholesale 字段仍在 fields 里，无输入框时 fieldDisplayVal 兜底写 0）
- [x] **移除 PDF 修正入口**（v3.2 / manifest v1.15.0 已实现）：Tab 剩 Excel 导入 + 商品库更新；删除 pdfState/renderPdfZone/processPdfFile/processPdfExcelFile/selectPdfMode/clearPdfData/resetPdfState/jumpPdfStep/previewPdfChanges/extractPdfTable/parsePdf；manifest 移除 pdf.min.js/pdf.worker.min.js 加载（lib 文件保留）；buildPdfPreviewModal 的 mode 兜底改 excelState.mode。⚠️ **v3.8 已恢复 PDF 解析能力**（改走统一「供应商数据导入」Tab，见下条）
- [x] **SKU 更新**（v3.3 + v3.4 / manifest v1.17.0 已实现）：两级匹配中 **UPC 匹配成功**（SKU 匹配失败）→ 用**转换 Excel 的 SKU**（pdf.catalog）更新 `product.packaging`——**SKU 列按入口区分**（单件=「订单行/包装/SKU」、套装=「订单行/包装/Box SKU」）；单件不识别规格定位 qty=1 记录写 `single_sku`、套装按套装数量（description xN）比对 "1 box of XX pieces" 写 `box_sku`；Modal SKU 列旧→新对照、变更默认勾选、失败文案拆 5 种、找不到规格跳过并提示；`matchPdfToExcel` pair 带 matchLevel（sku/upc）（详见 docs/PROJECT.md §6.5）
- [x] **恢复 PDF 解析 + 上传界面改「供应商数据导入」**（v3.8 / manifest v1.21.0，2026-09-09 已实现）：extractPdfTable/extractCoupon/parsePdf 从 git ee9b5bc^（v3.2 移除提交的父版本）恢复并适配 v3.3-v3.7 演进代码；原「Excel 导入」Tab/界面改名「供应商数据导入」，步骤1 上传供应商数据支持 **.pdf（PDF 分支）与 .xlsx/.xls（Excel 分支）多文件混合**，`detectSupplierKind` 自动识别路由，两分支同构输出统一合并；manifest 恢复加载 lib/pdf.min.js（content_scripts）+ lib/pdf.worker.min.js（web_accessible_resources）。语法检查通过，副本已同步，待 Chrome 冒烟验证（详见 docs/PROJECT.md §6.3/#25）
- [x] **Modal 内包装数量改 0 → 走缺货逻辑 + 供应商数据支持多选上传**（v3.9 / v3.9.1 / manifest v1.22.1，2026-09-10 已实现）：`syncRowOutstock`（价格置 0 锁定 / 备注缺货 / 标红 / 数量核对列与统计栏实时刷新 / 改回非 0 按快照还原）+ `isRowOutstock`（源数据缺货 或 boxQty 空/0）+ 单格渲染抽 `buildFieldCell`/`mountFieldInput`；`makeDropZone` 加 `multi` 参数（原生多选 + 拖入多文件 + 串行解析）。**v3.9.1 修复**：抽出的 `buildFieldCell` 内 `row._fieldRows[key] = fieldRow` 丢了行尾分号，被 ASI 与下一行 IIFE 拼成 `fieldRow(...)` 调用 → 预览报「fieldRow is not a function」，补分号修复。语法检查通过，副本已同步，待 Chrome 冒烟验证（详见 docs/PROJECT.md §6.2/#26 与 §9 #27）
- [ ] **Chrome 冒烟验证 v3.9**：① modal 把包装数量改成 0 → 行标红 / 数量核对显示「缺货」/ 单价与 0.9箱规价置 0 锁定 / 备注「缺货」/ 自动勾选；确认写入后 Odoo `product_packaging_qty`=0、`price_unit`=0、`box_wholesale_price`=0（单件入口也写箱批发价 0）、`remark`=缺货；② 改回非 0 → 价格/备注/行色/勾选还原、不写备注；③ 多选一次上传多个供应商 PDF/Excel → 逐个提示解析结果、合并行数正确、仍可单独删除
- [ ] **Chrome 冒烟验证 v3.8**：文本型 PDF 直传解析（行数/格式 A/B/Coupon）、PDF 与 Excel 混合多文件合并去重、单件/套装两入口路由 + 匹配/预览/写回回归、商品库 Tab 不受影响
- [ ] **Chrome 冒烟验证 v1.10.0**：两级匹配（同 UPC 多包装拆分命中/拆分失败缺货）、UPC+包装 Odoo 匹配（同 UPC 多包装精确命中/包装对不上报未找到）、双 Tab 切换、Excel 流、PDF 流回归
- [ ] **Chrome 冒烟验证 v1.11.0**：product.product 页注入 + 默认商品库 Tab、`极牛产品名称-KVIVA(1).xlsx` 141 行解析、UPC 匹配、预览勾选、写回 name/brand
- [ ] **Chrome 冒烟验证 v3.3**：UPC 匹配行 SKU 列旧→新对照展示、SKU 变更默认勾选、写回 single_sku（单件）/box_sku（套装）、找不到包装规格时橙色提示且订单行写回不受影响
- [ ] `searchPoByName`/`searchPoByRef` 去重合并（低优先，保留无害）

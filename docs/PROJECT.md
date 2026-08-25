# Odoo Excel Importer — 项目文档

> 本文件是项目的**权威文档**，与代码同步维护。任何 Agent 接任务前必须先读本文件。
> 最后更新：2026-08-21（manifest v1.11.0，新增「商品库更新」Tab）

---

## 1. 项目概述

| 项 | 内容 |
|---|---|
| 名称 | Odoo Excel Importer（`odoowritting-extension`） |
| 类型 | Chrome 扩展（Manifest V3），content script 注入式工具 |
| 业务目的 | 把采购 Excel / 供应商 PDF 数据批量写入 Odoo 采购订单（RFQ）明细行 |
| 技术栈 | 纯原生 JS（**ES5 风格：`var` + IIFE + function 声明**），无框架、无构建步骤 |
| 第三方库 | `lib/xlsx.full.min.js`（SheetJS）、`lib/pdf.min.js` + `lib/pdf.worker.min.js`（pdfjs-dist **3.11.174 UMD 版**，4.x 起无 UMD 故锁定 3.x） |
| 运行环境 | Odoo 实例的 `purchase.order` 页面 **或** `product.product`（产品变体）页面（URL hash 含对应 `model=`） |
| 版本 | manifest v1.11.0 |

**核心价值**：人工核对采购数据 → 自动写回 Odoo 的「数量 / 包装数量 / 整箱批发价 / 备注」字段，避免逐行手工录入。

---

## 2. 文件结构

```
odoowritting_extension/
├── content.js                 # ★ 主脚本（1866 行），全部业务逻辑与 UI
├── manifest.json              # MV3 配置（v1.2.0）
├── lib/
│   ├── xlsx.full.min.js       # Excel 解析（SheetJS）
│   ├── pdf.min.js             # PDF 解析（pdfjs-dist 3.11.174）
│   └── pdf.worker.min.js      # PDF worker（web_accessible_resources）
├── odoowritting_extension/    # ⚠️ 打包副本（Chrome「加载已解压的扩展程序」用）
│   ├── content.js             #    与根目录相同
│   ├── manifest.json          #    旧版 v1.1.0
│   ├── lib/                   #    库副本
│   └── icons/
├── icons/                     # 扩展图标（16/48/128）
├── docs/
│   └── PROJECT.md             # 本文件
├── AGENTS.md                  # Agent 入口（必读约定）
├── package.json               # 仅工程工具：husky + commitlint，无构建/测试脚本
├── commitlint.config.js       # 约定式提交规则（10 种 type）
├── GIT_WORKFLOW.md            # 团队 Git 工作流手册（Trunk-Based + Conventional Commits）
├── .github/                   # Issue/PR 模板、commitlint.yml CI、分支保护指南
└── .husky/commit-msg          # 提交信息自动校验
```

> ⚠️ **重要**：根目录是**源**，`odoowritting_extension/` 是**打包副本**。改代码改根目录的 `content.js`，改完需同步到子目录（并同步 manifest 版本号）。

---

## 3. 核心常量（改字段映射只改这里）

位于 content.js 顶部：

### 3.1 ODOO_FIELDS — Odoo 字段名映射（用户 2026-08-14 确认）

| 键 | 值（Odoo 字段） | 含义 |
|---|---|---|
| `qty` | `product_qty` | 订单行数量（当前未接入写回管道） |
| `boxQty` | `product_packaging_qty` | 包装数量 ✅ 用户确认 |
| `boxWholesalePrice` | `box_wholesale_price` | 整箱批发价 ✅ 用户确认 |
| `unitPrice` | `price_unit` | 单价 ✅ 用户确认 |
| `remark` | `remark` | 备注 ✅ 用户确认 |
| `poRef` | `name` | 采购订单「订单关联」列对应字段 ✅ 用户确认 |

> 模型：`purchase.order.line`。全部写回字段均经此映射，业务代码不硬编码字段名。

### 3.2 EXCEL_COLS — Excel 表头列名（按名动态匹配）

> 表头名为用户 2026-08-17/18 确认真实表头（catalog 列实测为「SKU_x」，非旧文档的「包装/SKU」）。

| 键 | 表头名 | 用途 |
|---|---|---|
| `upc` | 订单行/产品/内部参考号 | 匹配键1（UPC） |
| `catalog` | SKU_x | 匹配键2（Catalog）✅ 实测确认 |
| `shop` | 订单行/店铺 | 多条区分 |
| `qty` | 订单行/数量 | 总件数 |
| `boxQty` | abw交货箱数 | 交货箱数 → product_packaging_qty |
| `pack` | 订单行/包装 | 取 "1 box of XX pieces" 的 XX（套装单件数量降级用 + v1.9 Odoo 包装匹配键 + **v2.0 拆分匹配键**） |
| `boxPrice` | 箱规价 | 套装比对列1（不写回） |
| `boxPrice09` | 0.9箱规价 | 套装比对列 → box_wholesale_price |
| `total09` | 0.9总价 | 套装比对列（不写回） |
| `unitPrice` | 单价 | 单价 → price_unit |
| `remark` | 备注 | 缺货等 |
| `orderRef` | 订单关联 | 查 PO（=name） |
| `partnerRef` | 参考号 | 查 PO（=name，v1.7.1 新增）：**先用参考号查，查不到再用订单关联**（`loadOrderLineMap`）；命中订单行按两个键都注册进 lineMap |

### 3.3 其他

- `PREFIX = "__odoi_"`：所有 DOM id / class / localStorage key 前缀
- `LS_POS`：按钮位置 localStorage key；`LS_LOG`：最近日志 key
- `PDF_MODES`：`single`（单件入口📦）/ `set`（套装入口🎁）

---

## 4. 架构分层（content.js 内部结构）

单个 IIFE 内按 `═══` 注释分块，共 12 块：

| 块 | 内容 | 关键函数 |
|---|---|---|
| 常量 | 字段映射、列名、前缀 | — |
| 工具 | DOM 创建 `el()`、HTML 转义 `escHtml()`、`sleep`、价格计算 | `calcBoxPrice`（×0.9，整数分运算防浮点误差）、`parsePieces`（从 "1 box of 20 pieces" 取 20） |
| Odoo API | JSON-RPC 封装 | `rpcCall`、`searchPoByRef`、`getOrderLines`（含 `product_packaging_qty`）、`updateOrderLine`；`searchPoByName` 已无用可删 |
| Excel 解析 | 按表头列名解析 | `parseOrderExcel`（PDF/Excel 两流共用） |
| PDF 解析 | 文本块坐标聚类成表 | `parsePdf`（pdfjs 入口）、`extractPdfTable`（核心算法） |
| 业务逻辑 | 匹配修正规则 | `applyPdfByMode`（入口分发，PDF/Excel 两流共用）、`applyPdfSet`（套装）、`applyPdfSingle`（单件）、`parseConvertedPdfExcel`（Excel 流：PDF 转换版 Excel 解析，v1.7.0） |
| 日志 | localStorage 存取 | `saveLog`/`getLog`/`clearLog`/`downloadLog` |
| 状态 | 应用状态对象 | `appState`（含 `activeTab`）、`pdfState`、`excelState`、`dragCtx` |
| UI | 按钮/卡片/三 Tab/上传区/日志区 | `createDraggableButton`、`renderCardContent`、**`renderTabSwitch`（三 Tab）**、`renderExcelZone`、`renderPdfZone`、**`renderProductZone`（商品库，v1.11.0）**、`renderModePicker`（两流共用） |
| 预览 Modal | 双流共用 1 套 | `buildPdfPreviewModal(previewRows, fileName, source)`（source='pdf'/'excel'）、`loadOrderLineMap`、`buildPreviewRows` |
| 执行写回 | 双流共用批量写入 | `executePdfUpdate`（两流共用，字段经 ODOO_FIELDS 映射） |
| 初始化 | 注入条件与路由 | `inject()`、`isPurchaseOrderPage()`、`hashchange` 监听 |

---

## 5. Odoo API 通信层

- **协议**：JSON-RPC 2.0，POST 到 `{origin}/web/dataset/call_kw/{model}/{method}`
- **模型**：`purchase.order`（查询）、`purchase.order.line`（读/写）

| 函数 | 用途 | 关键点 |
|---|---|---|
| `rpcCall(endpoint, params)` | 通用调用 | body: `{jsonrpc:"2.0", method:"call", params, id}` |
| `searchPoByName(name)` | Excel 导入流查 PO | domain `[["name","=",name]]`，fields 需含 `order_line` |
| `searchPoByRef(ref)` | PDF 流查 PO | domain 用 `ODOO_FIELDS.poRef`（=name），fields 需含 `name`（否则查询为空） |
| `getOrderLines(ids)` | 批量取订单行 | fields: `id,name,price_unit,box_wholesale_price,remark,product_packaging_qty,product_id,product_packaging_id`；**v1.9 匹配键 = `product.default_code`（UPC）+ `product.packaging.qty`（包装件数）**，default_code 为空时用 `line.name` 正则 `\[(\d+)\]` 兜底（旧数据兼容）；内部追加查 `product.product`（default_code）与 `product.packaging`（name/qty） |
| `updateOrderLine(lineId, payload)` | 写回单行 | `write` 方法 |

> 注意：`searchPoByName` 与 `searchPoByRef` 逻辑几乎重复（都查 name），系历史遗留，重构时可合并。

---

## 6. 两条业务主流程

### 6.1 Excel 导入流 — ⚠️ 已移除（2026-08-14），待与用户重新对齐修改逻辑后重写

> 用户要求先删除 Excel 导入分支的全部处理逻辑，后续会给出新的修改规则。
> **当前状态**：Excel Tab 显示「Excel 导入逻辑待重新设计」占位；按钮拖入文件仅提示。
>
> 移除的函数（git 历史 `58d1f97` 可恢复）：`parseExcel`（固定列索引解析）、`groupByOrderRef`、`determineRemark`、`buildPreviewRow`、`processFile`（主流程）、`executeUpdate`（写回）、`renderUploadZone`/`pickFile`（上传区）、`buildPreviewModal`/`renderPreviewTable`/`renderPreviewRow`/`updateConfirmBtn`（预览 Modal）、`renderTabBar`/`addTab`/`removeTab`/`findTab`（标签条）、`minimizeModal`/`restoreModal`/`closeModal`。
>
> 保留的备用件：`searchPoByName`（按 name 查 PO）、`sleep`、`parsePieces`、`formatPrice`、`removeModals`、`td`、`updateLoadingOverlay`、`showResultToast`。
>
> 旧逻辑要点（重写时参考）：Excel 列 `row[0]`=UPC、`row[5]`=订单关联、`row[10]`=单价、`row[11]`=整箱批发价、`row[12]`=采购结果；写回 `price_unit`/`box_wholesale_price`/`remark`；备注按采购结果含「可能缺货/缺货」判定。

### 6.2 PDF 修正流（manifest v1.5.0 生效）

> ⚠️ **2026-08-18 规则再改**（覆盖 8-17 规则）：本次只改**套装**逻辑与 modal 展示，全局规则如下：
>
> **全局（单件/套装共用）**：
> - 按表头有无 `HS CODE` 识别格式：有 HS → PDF 表 `CATALOG NO.` 列；无 HS → PDF 表 `Catalog#` 列（解析器统一提取为 `catalog` 字段）
> - **匹配（v2.0 重构，2026-08-20 用户需求）**：两级匹配 `matchPdfToExcel`——
>   - 第一级 SKU：PDF Catalog ↔ Excel「SKU」列（`EXCEL_COLS.catalog`），能匹配上的直接配对（v1.7.2 逻辑不变）
>   - 第二级 UPC（SKU 匹配不上的行，含 Excel SKU 有值但 catalog 无对应的行）：先**整体统计两侧该 UPC 出现次数**，任一侧出现多次 → **包装拆分匹配**：PDF 侧 key = `UPC+description 的 "(xN)"`（`extractDescPack`，如 "(x42)"→42，忽略 `[195ml x 42]`），Excel 侧 key = `UPC+包装列 pieces`（`extractPackQty`，"1 box of 42 pieces"→42）；提取不到包装数字 → key 保持原 UPC 不拼接。两侧都唯一 → 直接按 UPC 匹配（原兜底）
>   - 一个 PDF 行可配对多个 Excel 行（SKU 阶段 + UPC 阶段均可，同 SKU/UPC 拆多个订单关联的 v1.7.5 合并核对场景）；每行 Excel 最多配对一次；pair 带 `matchKey`（SKU 行=Catalog，UPC 行=UPC 或拆分键）供数量核对聚合
>   - 拆分后仍匹配不上的 Excel 行 → 归入 outstock（缺货，用户确认 2026-08-20）；PDF 侧多余行静默忽略
> - ×0.9 开关：`Coupon=0` → factor=1（不打折）；`Coupon≠0`（负数）→ factor=0.9
> - 数量核对（v2.0）：**按 pair.matchKey 分组**求和 Excel「abw交货箱数」 vs **PDF Qty 总和**（同一转换版行被多个订单关联命中时 PDF Qty 只计一次）。SKU 匹配行按 Catalog 聚合（同 SKU 拆多个订单关联/多个 UPC 合并比对，不再各自报错）；UPC 唯一行按 UPC 聚合；拆分匹配行按 `UPC+包装` 聚合（不同包装独立核对，不合并求和）。不一致 modal 提示（不自动扣减）
> - **Odoo 行匹配（v1.9 改，2026-08-20 用户需求）**：写回匹配键从「UPC」改为「**UPC + 订单行/包装**」——UPC = Excel「内部参考号」↔ Odoo `order_line/product_id/default_code`（default_code 为空时用 name 正则 `\[(\d+)\]` 兜底）；包装 = Excel「订单行/包装」件数（`extractPackQty` 提取）↔ Odoo `order_line/product_packaging_id`（取其 qty）。**Excel 有包装件数 → 精确匹配（UPC+包装），失败且该 UPC 在 PO 中唯一 → 按 UPC 兜底，多行 → 报「未找到匹配的订单行」不写回（防错配其他包装行）；Excel 无包装件数 → 仅 UPC 唯一时命中，多行同样报未找到**
> - 缺货备注：① abw交货箱数为空/0；② Excel 有而 PDF 无 → 都备注「缺货」写 remark
> - **缺货行写回（v1.8.0 改，2026-08-19 用户需求）**：**Excel 入口**缺货行（备注「缺货」）的**单价（price_unit）/ 整箱批发价（box_wholesale_price）不写回 Odoo**（预览仍显示计算值供核对，写回时跳过）；PDF 入口行为不变（缺货行价格照常写回）。实现：`applyPdfByMode` 增加 `source` 参数（'pdf'/'excel'），`applyPdfSingle`/`applyPdfSet` 中缺货行（boxQty 空/0）对应价格字段 `odooField` 置 null
>
> **单件入口（不变）**：单价 = `UNIT PRICE × factor` → 比对「单价」列 → 写回 `price_unit`
>
> **套装入口（v1.5 重写，不区分 HS CODE）**：
> | 字段 | 计算 | 比对 Excel 列 | 写回 Odoo |
> |---|---|---|---|
> | 箱规价 | `UNIT PRICE IN HKD` 原值 | 箱规价 | ❌ |
> | 单价 | `UNIT PRICE × factor ÷ 套装单件数量` | 单价 | ✅ `price_unit` |
> | 0.9箱规价 | `UNIT PRICE × factor` | 0.9箱规价 | ✅ `box_wholesale_price` |
> | 0.9总价 | `Subtotal × factor` | 0.9总价 | ❌ |
> - **套装单件数量**：从 PDF `PRODUCT DESCRIPTION` 提取 `x(\d+)`（"x30"→30）；提取不到**降级**用 Excel「订单行/包装」pieces（"1 box of XX pieces" 的 XX）
> - 无 coupon 照算：factor=1 代入全部公式
> - 单件/套装都写 `abw交货箱数 product_packaging_qty`（可编辑，数量不一致时提示）
>
> **modal 展示（v1.5 增强）**：
> - 标题后显示去重后的「N 个订单关联号」
> - 不一致字段为**可编辑输入框**，单元格内两行小字：PDF 来源计算值（如 `PDF: 50×0.9÷30 = 1.5`）+ Excel 原值
> - 悬浮不一致字段显示**原因气泡**（原因由 content.js 内置 `REASONS` 字典生成）
> - 表格底部**总和行**：所有数值列（abw交货箱数/单价/箱规价/0.9箱规价/0.9总价）全表合计，随编辑实时刷新

```
步骤0 选修正入口（单件📦 / 套装🎁）
步骤1 上传 PDF → parsePdf → extractPdfTable（格式识别 + 跨页 + Coupon）
步骤2 上传采购订单 Excel → parseOrderExcel（按列名）
步骤3 匹配修正 applyPdfByMode → 预览 buildPdfPreviewModal（可编辑） → executePdfUpdate
```

#### PDF 表格提取算法（`extractPdfTable`）

1. 找表头：含 `UPC|EAN|Barcode` 的文本块，取最小 y 为表头行
2. 表头行内按 x0 排序，相邻标题左边界中点 = 列区间边界
3. 关键列按表头名正则识别：`/UPC|EAN|Barcode/i`、`/Qty/i`、`/Subtotal/i`
4. 数据行锚点：UPC 列内 8~14 位纯数字块，按 y 排序
5. 每行：y 容差 12 内文本块 → 按 x0 归属列 → 同列多块按 y 拼文本
6. 输出 `{format, coupon, rows: [{upc, catalog, qty, unitPrice, subtotal, description}]}`（列数不写死，动态识别）

#### PDF 两种格式识别（2026-08-17 用户提供两个样本确认）

供应商 PDF 有 **两种格式**，需先识别格式再提取字段：

| 特征 | 格式 A「套装/Box」单 | 格式 B「单件/Single」单 |
|---|---|---|
| 样本 | `供应商最终提供的采购信息.pdf` | `Castlers_Single_Order Invoice_35119284.pdf` |
| 页数 | 2 页 | 1 页 |
| 表头 | `UPC/EAN \| Qty \| Catalog# \| Brand \| Product Description \| Unit Price in HKD \| Subtotal`（7 列） | `UPC/EAN NUMBER \| BRAND \| CATALOG NO. \| HS CODE \| PRODUCT DESCRIPTION \| UNIT PRICE IN HKD \| QTY \| SUBTOTAL`（8 列） |
| **HS CODE 列** | ❌ 无 | ✅ 有（最可靠区分点） |
| Qty 列位置 | 第 2 列 | 第 7 列（Unit Price 之后） |
| 汇总区 | 第 2 页：`Subtotal` / `Coupon` / … / `Order Total` / `Paid by Store Credit` / **`Order Total after Store Credit`** | 第 1 页末尾：`Line Total` / `Coupon` / **`Total Amount`** |
| 客户邮箱 | purchase-box@Castlers.com | purchase-single@castlers.com |

**区分方案（判定优先级从高到低）**：
1. 表头是否含 **`HS CODE`** → 含 = 格式 B（单件）；不含 = 格式 A（套装）
2. 兜底：全文含 `Order Total after Store Credit` → 格式 A；含 `Total Amount`/`Line Total` → 格式 B

**两种格式都要提取的字段**（→ 匹配 Excel，v1.5 用途）：

| 目标 | 格式 A 表头名 | 格式 B 表头名 | 用途 |
|---|---|---|---|
| 匹配键 | UPC/EAN | UPC/EAN NUMBER | 「订单行/产品/内部参考号」 |
| 匹配键2 | Catalog# | CATALOG NO. | 「SKU_x」 |
| 数量 | Qty | QTY | 数量核对（abw交货箱数比对目标，Qty 原值） |
| 单价 | Unit Price in HKD | UNIT PRICE IN HKD | 套装箱规价/单价/0.9箱规价；单件单价 |
| 小计 | Subtotal | SUBTOTAL | 套装 0.9总价比对 |
| 描述 | Product Description | PRODUCT DESCRIPTION | 提取套装单件数量 `x(\d+)` |
| Coupon | 汇总区 | 汇总区 | ×0.9 开关（≠0 负数 → factor 0.9） |

> ✅ 已确认（2026-08-17 实测）：格式 A/B 与入口无绑定关系，单件/套装都可能出现任一格式；格式仅影响匹配键列名与数量核对目标。当前 v1.5 数量核对已统一为 PDF Qty 原值（见上方 ⚠️ 块）。

#### 套装入口规则（`applyPdfSet`，v1.5 已实现）

> 见上方 ⚠️ 块「套装入口（v1.5 重写，不区分 HS CODE）」：4 项比对（箱规价/单价/0.9箱规价/0.9总价），套装单件数量来自 PDF `PRODUCT DESCRIPTION` 的 `x(\d+)`（降级 Excel 包装列）。
> 旧规则（8-17 分格式公式、随机扣减、`Subtotal÷(数量÷pieces)` 箱规价）已全部作废。

#### 单件入口（`applyPdfSingle`，v1.4.0 已实现，v1.5 公式不变）

- 匹配：SKU 优先 / UPC 兜底（同全局规则，v1.7.2）
- 数量核对：abw交货箱数求和 vs PDF Qty 原值（同全局规则），不一致 modal 提示
- 单价：`UNIT PRICE × factor` → 比对「单价」列 → 写回 `price_unit`（不分格式）
- 缺货备注：`remark`
- 字段带 `reason`/`pdfSource` 供 modal 对照展示（v1.5 补充，公式未变）

#### PDF 流写回（`executePdfUpdate`）

- 写回字段：`product_packaging_qty`（abw交货箱数）/ `box_wholesale_price`（0.9箱规价）/ `price_unit`（单价）/ `remark`（备注），经 ODOO_FIELDS 映射
- 只处理勾选行；预览行标红场景：数量不一致（qtyMismatch）、缺货（outstock）、匹配失败（error）
- 每条变更字段带 `odooField`，为 null 的字段（箱规价/0.9总价）仅比对不写回

### 6.3 Excel 导入流（manifest v1.7.0 改造，2026-08-19）

> **用户需求（2026-08-19）**：Excel 入口上传的是「**PDF 转换版 Excel**」（供应商订单确认 PDF 转成的 Excel，如 `Castlers Box 08.27_Order Confirmation_35548880.xlsx`），**比对逻辑与 PDF 流一模一样**：同样区分有无 HS CODE（格式 A/B）、同样提取 Coupon 决定 ×0.9、同样走单件/套装公式 → 预览修改 → 写回。用户确认：**仍需再传采购订单 Excel**（与 PDF 流一致的两步上传）。
>
> ⚠️ v1.6.0 的旧逻辑（上传采购订单 Excel → `applyExcelChanges` 比对 Odoo 现有值）**已移除**。

```
步骤0 选入口（单件📦 / 套装🎁）——与 PDF 区共用 renderModePicker
步骤1 上传转换版 Excel → parseConvertedPdfExcel（新增：输出与 extractPdfTable 同构）
步骤2 上传采购订单 Excel → parseOrderExcel → applyPdfByMode（与 PDF 流同一分发）
步骤3 预览 buildPdfPreviewModal（pdf 布局）→ executePdfUpdate 写回
```

**`parseConvertedPdfExcel(data)`（新增，content.js）**：
- 表头行 = 首个含 `UPC|EAN` 关键字的行（转换版 Excel 前面有公司/客户信息行）
- 格式识别：表头含 `HS CODE` → 格式 B，否则 A（与 PDF 流一致）
- 列定位：按表头名正则（UPC/EAN、Catalog、Qty、Unit Price、Subtotal、Product Description），**容忍空列与用户附加列**（样本中 UNIT PRICE 与 QTY 之间有空列、SUBTOTAL 后用户加了 ×0.9 列，均不影响）
- 数据行：UPC 列为 8~14 位纯数字的行；重复表头行跳过；汇总区（Line Total/Coupon/Total Amount）自然跳过
- Coupon：含「Coupon」单元格所在行右侧首个数值（样本 = -73702.326 → factor 0.9）；无 Coupon 行 → 0（factor 1）
- 返回 `{ format, coupon, rows:[{upc,catalog,qty,unitPrice,subtotal,description}] }` —— 与 `extractPdfTable` 输出同构，直接喂给 `applyPdfByMode`

**写回字段 / 比对规则 / 缺货规则 / Modal 交互**：与 PDF 流（§6.2）完全一致，不再单独维护。**例外（v1.8.0）**：Excel 流缺货行（备注「缺货」）的单价/整箱批发价不写回 Odoo，详见 §6.2 ⚠️ 块。Modal 的 `mode` 改为优先取预览行自带标记（`previewRows[0].mode`），`fieldOrder` 的 excel 分支已移除。

**样本实测（2026-08-19，node 单测）**：`Castlers Box 08.27_Order Confirmation_35548880.xlsx` → 格式 B、Coupon=-73702.326、64 行、UPC 全部合法、description 含 xN（套装单件数量）、Subtotal=UnitPrice×Qty 验算通过。

### 6.4 商品库更新流（manifest v1.11.0，2026-08-21 用户需求）

> **用户需求（2026-08-21）**：入口在采购单顶部导航「产品 → 产品变体」页（`model=product.product`）；上传商品库 Excel，按 UPC 定位产品变体，用 Excel「品牌」「中文简称」更新 Odoo 对应字段。

- **注入**：`isAnyPage()` = `model=purchase.order` **或** `model=product.product`（v1.11.0 放宽）；产品页默认 Tab = 商品库更新，采购页默认 Excel；hashchange 按页面类型重置默认 Tab
- **Excel 列（`PRODUCT_EXCEL_COLS`，表头第 1 行按名定位 + 正则容错）**：`UPC`（12/13 位数字，按字符串 trim 处理）| `品牌` | `中文简称`（样本实测含 `\n`，解析时转空格）
- **Odoo 字段（`PRODUCT_ODOO_FIELDS`，用户确认均为 Char）**：`default_code`（定位键）| `brand`（品牌）| `name`（中文简称）
- **流程**：上传商品库 Excel（`makeDropZone` 复用）→ `parseProductExcel` → `loadProductByUpc`（search_read `product.product`，domain `default_code in [...]`，UPC 去重后 500/批分块）→ `buildProductPreviewRows` → `buildProductPreviewModal`（UPC + 中文简称/品牌「旧→新」对照 + 勾选）→ `executeProductUpdate`（批量 write `{name, brand}`）
- **行状态**：`ok`（✅ 待更新，UPC 匹配到唯一产品，**默认全勾选**）/ `notfound`（⚠️ 未匹配，UPC 无对应产品，注明原因）/ `dup`（🔁 UPC 匹配到多个产品，注明原因需人工）；**不做新旧比对**（用户 2026-08-21 确认）
- **写回**：Excel 为权威源，`{name, brand}` 按 Excel 值**直接写入（空值也照写，无空值保护）**；仅勾选行写回；日志/toast 与双流共用
- **样本实测（2026-08-21，openpyxl）**：`极牛产品名称-KVIVA(1).xlsx` → Sheet1 142 行（1 表头 + 141 数据）、3 列、无合并单元格、UPC 12/13 位混存（如 8809640734526 / 880933516775）、中文简称含换行

---

## 7. UI 结构

| 组件 | 说明 |
|---|---|
| 可拖动按钮 📥 | 右下角 52px 圆形紫色按钮，位置存 localStorage；拖文件到按钮上提示在卡片内选择入口 |
| 悬浮卡片 | 点按钮展开（340×500px），含 Header、三 Tab、上传区、日志区 |
| 三 Tab | 「📊 Excel 导入」/「📄 PDF 修正」/「🏷 商品库更新」（v1.11.0 新增；`appState.activeTab`，产品变体页默认商品库 Tab，采购页默认 Excel） |
| 入口选择器 | 步骤 0：单件📦 / 套装🎁 两张卡片（PDF/Excel 两流共用 `renderModePicker`），选中后锁入口，可「切换入口」重置 |
| 步骤指示器 | PDF 区：① 上传 PDF → ② 上传 Excel → ③ 预览确认；Excel 区（v1.7.0）：① 上传转换版 Excel → ② 上传采购订单 Excel → ③ 预览确认（当前步紫色、已完成绿色、可点击回退） |
| 常驻错误框 | `pdfState.error` / `excelState.error`：解析失败时在卡片内红色常驻显示原因，下次成功自动清除 |
| 预览 Modal | 92vw 宽居中弹窗，双流共用（v1.7.0 起统一 pdf 布局：数量核对/PDF Qty 列两流都有）；`mode` 优先取预览行自带标记；商品库更新 Modal（v1.11.0）：UPC + 中文简称/品牌 旧→新 对照 + 勾选写回 |
| 日志区 | 最近一次执行结果（成功/失败/跳过计数），可下载 JSON、清除 |
| Toast / Loading | 右下角提示；顶部紫色 loading 胶囊（`updateLoadingOverlay`） |

**样式约定**：主色 `#7c3aed`（紫）；行状态色 ROW_COLORS（绿=价格、橙=备注、黄=混合、灰=无操作）；PDF 预览行红=警示、绿=正常。全部内联 style，无独立 CSS 文件。

---

## 8. 工程规范与约定

### 代码风格（必须遵守）

- **ES5 风格**：`var`、`function` 声明、IIFE 包裹、`"use strict"`；不用箭头函数/解构/模板串（保持与现有代码一致）
- 函数块用 `// ═══` 分隔注释；中文注释；注释精简
- DOM id/class 一律带 `__odoi_` 前缀
- 用户输入展示必须经 `escHtml()` 转义（防 XSS）；写回前价格经 `parseFloat` 兜底
- 价格运算用整数分（`calcBoxPrice`），禁止直接浮点乘除
- 调试日志前缀：`[Odoo Excel Importer]`、`[Odoo PDF]`（测试/诊断插桩一律不留在交付代码里）

### 修改字段 / 列名

只改 content.js 顶部的 `ODOO_FIELDS` / `EXCEL_COLS` 常量，不要在业务代码里硬编码。

### 开发验证流程

1. **无构建步骤**：改完 `content.js` 直接生效
2. Chrome → `chrome://extensions` → 开发者模式 → 「加载已解压的扩展程序」→ 选 `odoowritting_extension/` 子目录
3. 改完根目录代码 → **同步到 `odoowritting_extension/`**（content.js + manifest 版本号）
4. 扩展页点刷新 → 打开 Odoo 采购订单页验证
5. 关键验证点：PDF 表格解析行数、UPC 匹配、扣减逻辑、写回后 Odoo 字段值

### Git 规范

- Conventional Commits（commitlint + husky 强制），10 种 type，本项目 scope 建议：`excel` / `ui` / `import` / `config` / `pdf`
- 分支 `<type>/<desc>`；Trunk-Based，main 不可直推；详见 `GIT_WORKFLOW.md`

---

## 9. 已知问题与待办

| # | 问题 | 位置 | 状态 |
|---|---|---|---|
| 1 | ~~Excel 导入分支处理逻辑~~ | — | ✅ v1.6.0 重写上线：Excel 入口（单件/套装 + 比对 Odoo + 可编辑 Modal），恢复双 Tab |
| 2 | ~~单件入口逻辑未实现~~ | `applyPdfSingle` | ✅ 已实现（v1.4.0，UPC/Catalog 匹配 + 单价×factor + 数量求和提示） |
| 3 | ~~Excel 的 Catalog 列名待确认~~ | `EXCEL_COLS.catalog` | ✅ 已确认 = 「SKU_x」（2026-08-17 实测） |
| 4 | Excel 公式列（0.9箱规价/0.9总价/单价）缓存值能否被 SheetJS 读到 | `parseOrderExcel` | ⚠️ 需真实 Excel 样本验证 |
| 5 | `searchPoByName` 已无用（无调用方） | Odoo API 层 | ℹ️ 可删除，保留无害 |
| 6 | `ODOO_FIELDS.qty`（product_qty）已不写回 | content.js | ℹ️ 保留备用 |
| 7 | ~~子目录打包副本版本落后~~ | `odoowritting_extension/` | ✅ 已同步（v1.6.0） |
| 8 | ~~Odoo 字段名待实测~~ | `ODOO_FIELDS` | ✅ 已确认：`product_packaging_qty` / `box_wholesale_price` / `price_unit` / `remark` |
| 9 | ~~套装分格式公式/随机扣减~~ | `applyPdfSet` | ✅ v1.5.0 重写：统一公式 + 4 项比对 + 套装单件数量（PRODUCT DESCRIPTION "x30"） |
| 10 | modal 增强（订单关联统计/PDF 对照/hover 原因/列总和） | `buildPdfPreviewModal` 等 | ⚠️ 已实现 v1.5.0，待 Chrome 冒烟验证 |
| 11 | ~~Excel 流比对 Odoo 现有值~~ | — | ❌ v1.7.0 移除：Excel 入口改为吃「PDF 转换版 Excel」，与 PDF 流同逻辑 |
| 12 | Excel 流转换版解析 + 两步上传 | `parseConvertedPdfExcel` 等 | ⚠️ v1.7.0 已实现，样本单测通过，待 Chrome 冒烟验证 |
| 13 | Odoo 匹配键改 UPC+包装（default_code + packaging qty） | `getOrderLines`/`loadOrderLineMap`/`buildPreviewRows` | ⚠️ v1.9.0 已实现，逻辑单测通过，待 Chrome 冒烟验证 |
| 14 | PDF↔Excel 匹配重构为两级匹配（SKU 优先 + UPC 包装拆分兜底） | `matchPdfToExcel`/`extractDescPack`/`markQtyMismatch` | ⚠️ v2.0（manifest v1.10.0）已实现，33 断言单测通过，待 Chrome 冒烟验证 |
| 15 | 商品库更新 Tab（UPC 定位产品变体，写回 name/brand） | `productState`/`renderProductZone`/`loadProductByUpc`/`executeProductUpdate` | ⚠️ v1.11.0 已实现（语法 + 状态判定单测通过），待 Chrome 冒烟验证（product.product 页注入、UPC 匹配、写回 name/brand） |

---

## 10. 快速上手（新 Agent 3 分钟版）

1. **只改一个文件**：`content.js`（IIFE 单文件）；字段映射在顶部两个常量对象
2. **两条在线流程（v1.7.0）**：📊 Excel 导入（上传「PDF 转换版 Excel」→ 上传采购订单 Excel → 与 PDF 同逻辑匹配修正 → Modal 写回）+ 📄 PDF 修正（上传 PDF → 上传 Excel → 匹配修正 → Modal 写回）；两流在 `applyPdfByMode` 汇合，共用 `parseOrderExcel` / `loadOrderLineMap` / `buildPreviewRows` / Modal / `executePdfUpdate`
3. **两级匹配（v2.0）**：PDF/转换版 Excel ↔ 采购单 Excel 先按 Catalog↔SKU 匹配，匹配不上的行走 UPC 匹配（同 UPC 多行用包装拆分：PDF 侧 description "(xN)"、Excel 侧包装列 pieces，拼成 `UPC+N` 精确配对；唯一 UPC 直接匹配）；数量核对按 matchKey（Catalog/UPC/拆分键）分组。**UPC+包装是 Odoo 匹配键（v1.9）**：Excel「内部参考号」+「订单行/包装」件数 ↔ Odoo `product.default_code` + `product_packaging_id.qty`；查 PO（=name）优先用「参考号」列、查不到再用「订单关联」列（v1.7.1）
4. **无构建、无测试**：改完同步到 `odoowritting_extension/` 子目录后到 Chrome 手动验证
5. Git 提交必须过 commitlint（husky），格式 `type(scope): subject`

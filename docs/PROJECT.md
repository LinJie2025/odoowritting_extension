# Odoo Excel Importer — 项目文档

> 本文件是项目的**权威文档**，与代码同步维护。任何 Agent 接任务前必须先读本文件。
> 最后更新：2026-09-10（v3.9.1 / manifest v1.22.1：Modal 内把「包装数量」改成 0/空 → 按缺货写回；供应商数据上传支持原生多选；修复 buildFieldCell 缺分号被 ASI 解析成函数调用导致「预览失败」）

---

## 1. 项目概述

| 项 | 内容 |
|---|---|
| 名称 | Odoo Excel Importer（`odoowritting-extension`） |
| 类型 | Chrome 扩展（Manifest V3），content script 注入式工具 |
| 业务目的 | 把供应商数据（**PDF 或 PDF 转换版 Excel**，v3.8 自动识别）批量写入 Odoo 采购订单（RFQ）明细行 |
| 技术栈 | 纯原生 JS（**ES5 风格：`var` + IIFE + function 声明**，无框架、无构建步骤） |
| 第三方库 | `lib/xlsx.full.min.js`（SheetJS）。`lib/pdf.min.js` + `lib/pdf.worker.min.js`（pdfjs-dist 3.11.174 UMD，勿升 4.x）——**v3.2 曾移除加载，v3.8（2026-09-09）恢复**（manifest content_scripts 加载 pdf.min.js；pdf.worker.min.js 在 web_accessible_resources 供 worker 拉取） |
| 运行环境 | Odoo 实例的 `purchase.order` 页面 **或** `product.product`（产品变体）页面（URL hash 含对应 `model=`） |
| 版本 | manifest v1.22.1 |

**核心价值**：人工核对采购数据 → 自动写回 Odoo 的「数量 / 包装数量 / 整箱批发价 / 备注」字段，避免逐行手工录入。

---

## 2. 文件结构

```
odoowritting_extension/
├── content.js                 # ★ 主脚本（约 3400 行），全部业务逻辑与 UI
├── manifest.json              # MV3 配置（v1.22.1）
├── lib/
│   ├── xlsx.full.min.js       # Excel 解析（SheetJS）
│   ├── pdf.min.js             # PDF 解析（pdfjs-dist 3.11.174；v3.2 曾移除加载，v3.8 恢复）
│   └── pdf.worker.min.js      # （同上，pdfjs worker，web_accessible_resources 提供）
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
| `catalog` | SKU_x | 匹配键2（Catalog）✅ 实测确认（旧表头，v3.4 起仅作兜底） |
| `catalogSingle` | 订单行/包装/SKU | **单件入口** SKU 列（v3.4，用户需求：单件 SKU 匹配用此列） |
| `boxCatalog` | 订单行/包装/Box SKU | **套装入口** Box SKU 列（v3.4：套装 SKU 匹配用此列） |
| `shop` | 订单行/店铺 | 多条区分 |
| `qty` | 订单行/数量 | 总件数 |
| `boxQty` | 包装数量 | 包装数量（v3.0 起；原「abw交货箱数」作废）→ product_packaging_qty，数量核对求和列 |
| `pack` | 订单行/包装 | 取 "1 box of XX pieces" 的 XX（套装单件数量降级用 + v1.9 Odoo 包装匹配键 + **v2.0 拆分匹配键**） |
| `boxPrice` | 箱规价 | v3.0 起不再作比对基准（保留列定义兼容解析） |
| `boxPrice09` | 0.9箱规价 | 套装比对列 → box_wholesale_price |
| `total09` | 0.9总价 | 套装比对列（不写回） |
| `unitPrice` | 单价 | 单价 → price_unit |
| `remark` | 备注 | 缺货等 |
| `orderRef` | 订单关联 | 查 PO（=name） |
| `partnerRef` | 参考号 | 查 PO（=name，v1.7.1 新增）：**先用参考号查，查不到再用订单关联**（`loadOrderLineMap`）；命中订单行按两个键都注册进 lineMap。**v1.12.2**：模糊匹配排除「内部参考号」（UPC 列）——无「参考号」列时 partnerRef 为空，直接走「订单关联」匹配（用户 2026-08-25 确认） |

### 3.3 其他

- `PREFIX = "__odoi_"`：所有 DOM id / class / localStorage key 前缀
- `LS_POS`：按钮位置 localStorage key；`LS_LOG`：最近日志 key
- `PDF_MODES`：`single`（单件入口📦）/ `set`（套装入口🎁）
- **`PACKAGING_ODOO_FIELDS`（v3.3）**：`product.packaging` 字段映射——`name`（包装规格，如 "piece"/"1 box of 42 pieces"）/ `qty`（包装件数，box 规格 = XX）/ `single_sku`（单件 SKU，piece 规格，标签「SKU」）/ `box_sku`（Box SKU，标签「Box SKU」），用户 2026-08-27 确认

---

## 4. 架构分层（content.js 内部结构）

单个 IIFE 内按 `═══` 注释分块，共 12 块：

| 块 | 内容 | 关键函数 |
|---|---|---|
| 常量 | 字段映射、列名、前缀 | — |
| 工具 | DOM 创建 `el()`、HTML 转义 `escHtml()`、`sleep`、价格计算 | `calcBoxPrice`（×0.9，**v3.5 起精确到 2 位小数四舍五入**）、`parsePieces`（从 "1 box of 20 pieces" 取 20） |
| Odoo API | JSON-RPC 封装 | `rpcCall`、`searchPoByRef`、`getOrderLines`（含 `product_packaging_qty`）、`updateOrderLine`；`searchPoByName` 已无用可删 |
| Excel 解析 | 按表头列名解析 | `parseOrderExcel`（PDF/Excel 双来源共用） |
| PDF 解析 | ⚠️ **v3.8（2026-09-09）恢复**：文本块坐标聚类成表 | `parsePdf`（pdfjs 入口，`pdfjsLib` 全局来自 lib/pdf.min.js）、`extractPdfTable`（核心算法）、`extractCoupon`（汇总区 Coupon） |
| 业务逻辑 | 匹配修正规则 | `applyPdfByMode`（入口分发，单件/套装共用）、`applyPdfSet`（套装）、`applyPdfSingle`（单件）、`parseConvertedPdfExcel`（**Excel 分支**：供应商 Excel/PDF 转换版解析）、`detectSupplierKind`（**v3.8：按文件扩展名判 .pdf → PDF 分支 / .xlsx/.xls → Excel 分支**） |
| 日志 | localStorage 存取 | `saveLog`/`getLog`/`clearLog`/`downloadLog` |
| 状态 | 应用状态对象 | `appState`（含 `activeTab`：'excel'=供应商数据导入 / 'product'）、`excelState`（供应商导入区，`convFiles` 带 `kind`）、`productState`、`dragCtx` |
| UI | 按钮/卡片/双 Tab/上传区/日志区 | `createDraggableButton`、`renderCardContent`、`renderTabSwitch`（🗂 供应商数据导入 / 🏷 商品库更新）、`renderExcelZone`（供应商数据导入区）、`renderProductZone`（商品库）、`renderModePicker`（单件/套装共用） |
| 预览 Modal | 双流共用 1 套 | `buildPdfPreviewModal(previewRows, fileName, source)`（source 仅兼容签名）、`loadOrderLineMap`、`buildPreviewRows` |
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
| `searchPoByRef(ref)` | PDF 流查 PO | domain 用 `ODOO_FIELDS.poRef`（=name），fields 需含 `name`（否则查询为空）；**v3.7 补读 `partner_id`**（supplierinfo 二级匹配按供应商过滤用） |
| `getOrderLines(ids, partnerId)` | 批量取订单行 | fields: `id,name,price_unit,box_wholesale_price,price_subtotal,remark,product_packaging_qty,product_id,product_packaging_id`（**v3.0 补读 price_subtotal**，套装小计比对原值）；**v1.9 主匹配键 = `product.default_code`（UPC）+ `product.packaging.qty`（包装件数）**，default_code 为空时用 `line.name` 正则 `\[(\d+)\]` 兜底（旧数据兼容）；内部追加查 `product.product`（default_code/product_tmpl_id）与 `product.packaging`（name/qty）；**v3.7（2026-09-04）二级匹配键 = `product.supplierinfo.product_code`**——传入 PO partner_id 后查该供应商名下 supplierinfo（模板级 + 变体级两段 search_read），行级返回 `supplierCodes`（去重、剔除与 default_code 重复项），供 `loadOrderLineMap` 注册 supMap/supCnt 次键索引 |
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

### 6.2 修正匹配流（供应商数据导入共用；manifest v1.5.0 起，v3.8 恢复 PDF 解析并归一为供应商数据导入）

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
> - 数量核对（v2.0）：**按 pair.matchKey 分组**求和 Excel「包装数量」（v3.0 起；原「abw交货箱数」列作废） vs **PDF Qty 总和**（同一转换版行被多个订单关联命中时 PDF Qty 只计一次）。SKU 匹配行按 Catalog 聚合（同 SKU 拆多个订单关联/多个 UPC 合并比对，不再各自报错）；UPC 唯一行按 UPC 聚合；拆分匹配行按 `UPC+包装` 聚合（不同包装独立核对，不合并求和）。不一致 modal 提示（不自动扣减）
> - **Odoo 行匹配（v1.9 改，2026-08-20 用户需求；v3.7 加二级键，2026-09-04）**：写回匹配键从「UPC」改为「**UPC + 订单行/包装**」——UPC = Excel「内部参考号」↔ Odoo `order_line/product_id/default_code`（default_code 为空时用 name 正则 `\[(\d+)\]` 兜底）；包装 = Excel「订单行/包装」件数（`extractPackQty` 提取）↔ Odoo `order_line/product_packaging_id`（取其 qty）。**v3.7（2026-09-04，只读库实测 28 行错位、活跃单 13 行）**：部分产品 `default_code` 被改过号，供应商票据 UPC 只存在于 `product.supplierinfo.product_code`（与 line.name `[码]` 同值）→ 主键匹配不上时**按 PO 供应商（partner_id）过滤的 supplierinfo.product_code 二级兜底**（`searchPoByRef` 补 partner_id → `getOrderLines(ids, partnerId)` 查模板级+变体级 supplierinfo 得行级 `supplierCodes` → `loadOrderLineMap` 注册 `supMap/supCnt` → `buildPreviewRows` 主键未命中再试次键，包装精确 + UPC 唯一兜底 + 多行防错配同主键）。**Excel 有包装件数 → 精确匹配（UPC+包装），失败且该 UPC 在 PO 中唯一 → 按 UPC 兜底，多行 → 报「未找到匹配的订单行」不写回（防错配其他包装行）；Excel 无包装件数 → 仅 UPC 唯一时命中，多行同样报未找到**
> - 缺货备注：① 包装数量为空/0；② Excel 有而 PDF 无 → 都备注「缺货」写 remark
> - **缺货行写回（v1.12.4 改，2026-08-26 用户需求，取代 v1.8.0）**：**所有缺货行**（缺货①匹配不上 + 缺货② boxQty 空/0，Excel 与 PDF 两入口统一）除备注「缺货」外，**单价（price_unit）与整箱批发价（box_wholesale_price）都置 0 写回**。实现：缺货② 的 unitPrice/boxPrice09（套装）字段 newValue 置 0、odooField 保留（不再置 null）；缺货① 的 fields 追加 `unitPrice=0` + `boxWholesale=0` 两字段；单件入口缺货② 追加 `boxWholesale` 字段（label 整箱批发价，v1.12.4 新增）。⚠️ 旧规则（v1.8.0：Excel 入口缺货行价格**不写回**）已作废。注意：缺货①行能否写回取决于 Odoo 行是否匹配到（匹配不到仍 error 不可写）
>
> **单件入口（不变）**：单价 = `UNIT PRICE × factor` → 比对「单价」列 → 写回 `price_unit`
>
> **套装入口（v1.5 重写，不区分 HS CODE；v3.0 改为 3 项比对，价格原值比对 Odoo）**：
> | 字段 | 计算（PDF） | 比对原值 | 写回 Odoo |
> |---|---|---|---|
> | 单价 | `UNIT PRICE × factor ÷ 套装单件数量` | Odoo `price_unit` | ✅ `price_unit` |
> | 0.9箱规价 | `UNIT PRICE × factor` | Odoo `box_wholesale_price` | ✅ `box_wholesale_price` |
> | 0.9总价 | `0.9箱规价 × 包装数量`（v3.6 动态列；旧 `Subtotal×factor` 作废） | Odoo `price_subtotal`（只读比对） | ❌ |
> - **箱规价字段已移除（v3.0，2026-08-26 用户确认）**：Odoo 无对应字段，modal 不再展示
> - **套装单件数量**：从 PDF `PRODUCT DESCRIPTION` 提取 `x(\d+)`（"x30"→30）；提取不到**降级**用 Excel「订单行/包装」pieces（"1 box of XX pieces" 的 XX）
> - 无 coupon 照算：factor=1 代入全部公式
> - 单件/套装都写 `包装数量 product_packaging_qty`（可编辑，数量不一致时提示）
>
> **价格比对基准（v3.0，2026-08-26 用户需求）**：modal 里单价/0.9箱规价/0.9总价的「原值」**不再读采购单 Excel 的价格列**，改为 **Odoo 订单行现值**（单价→`price_unit`、0.9箱规价→`box_wholesale_price`、0.9总价→`price_subtotal`）。实现：`applyPdf*` 生成字段时 oldValue 置占位，`buildPreviewRows` 命中 Odoo 行后回填并重算 changed/reason（缺货行跳过回填）；采购单 Excel 价格列（单价/箱规价/0.9箱规价/0.9总价）不再作比对基准
>
> **modal 展示（v1.5 增强，v3.0 原值改 Odoo）**：
> - 标题后显示去重后的「N 个订单关联号」
> - 不一致字段为**可编辑输入框**，单元格内两行小字：PDF 来源计算值（如 `PDF: 50×0.9÷30 = 1.5`）+ Odoo 原值
> - 悬浮不一致字段显示**原因气泡**（原因由 content.js 内置 `REASONS` 字典生成，v3.0 文案改「Odoo」）
> - 表格底部**总和行**：所有数值列（包装数量/单价/0.9箱规价/0.9总价）全表合计，随编辑实时刷新
> - **0.9总价动态列（v3.6，2026-09-04 用户需求）**：0.9总价列改为**只读动态计算列**——套装 = 0.9箱规价 × 包装数量、单件 = 单价 × 包装数量（单件入口新增此列，2026-09-04 用户确认）；modal 里修改「包装数量」/价格输入框时该列实时重算（`recalcRowTotal`，输入事件 → 重算 total09 → `refreshTotal`）；仍保留与 Odoo `price_subtotal` 的原值比对（不一致红框 + Odoo 原值小字），`odooField=null` 永不写回；底部总和与组小计（v3.1，单件入口本次起同样插组小计行）按动态值实时刷新
> - **组小计（v3.1，2026-08-27 用户需求；v3.6 起单件入口同样插）**：按订单关联分组，每组尾部插一行「{订单关联号} 小计」，只累计 **0.9总价列**（全部行、不区分勾选），随编辑实时刷新；底部全表合计行保留

```
步骤0 选修正入口（单件📦 / 套装🎁）
步骤1 上传供应商数据（PDF → parsePdf → extractPdfTable；转换版 Excel → parseConvertedPdfExcel；v3.8 自动识别路由，可多个合并）
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
- 数量核对：包装数量求和 vs PDF Qty 原值（同全局规则），不一致 modal 提示
- 单价：`UNIT PRICE × factor` → 比对原值 Odoo `price_unit`（v3.0 起）→ 写回 `price_unit`（不分格式）
- 缺货备注：`remark`
- 字段带 `reason`/`pdfSource` 供 modal 对照展示（v1.5 补充，公式未变）

#### PDF 流写回（`executePdfUpdate`）

- 写回字段：`product_packaging_qty`（包装数量）/ `box_wholesale_price`（0.9箱规价）/ `price_unit`（单价）/ `remark`（备注），经 ODOO_FIELDS 映射
- 只处理勾选行；预览行标红场景：数量不一致（qtyMismatch）、缺货（outstock/包装数量为 0）、匹配失败（error）
- 每条变更字段带 `odooField`，为 null 的字段（0.9总价）仅比对不写回
- **v3.9 兜底**：写回时按输入框现值再判一次缺货（`isRowOutstock`），命中则强制 `remark=缺货`、`price_unit=0`、`box_wholesale_price=0`（原包装数量字段存在时一并写 0），避免 modal 状态未同步

#### Modal 内「包装数量改为 0/空」→ 按缺货处理（v3.9，2026-09-10 用户需求）

> 用户需求：预览 Modal 里**主动把某行「包装数量」改成 0**（或清空）时，该行也要走缺货逻辑（与源数据缺货一致）。

- **触发**：`renderPdfPreviewRow` 给 boxQty 输入框挂 `input` 监听 → `syncRowOutstock(row)`（状态未变则直接返回）
- **判定** `isRowOutstock(row)`：`kind==='outstock'`（Excel 有 PDF 无）**或** boxQty 值 = 空/0（取值优先用输入框实时值，渲染早期退回字段值）
- **进入缺货态**：
  - 价格字段（单价；套装另加 0.9箱规价）`newValue=0 / changed=true / reason=缺货置 0`，输入框置 "0" 并**锁定为只读**（红底，防误改）；进入前值快照到 `row._osSnapshot`，`calcValue` 保留计算值
  - 单件入口动态补 `boxWholesale`（整箱批发价，`box_wholesale_price`）字段 = 0 写回（该列已隐藏，仅补字段）
  - 备注：无 remark 字段的行**动态补挂输入框**（`mountFieldInput`），值「缺货」；行标红、复选框自动勾选
  - 数量核对列原地刷成「缺货」（`paintQtyCell`）、悬浮气泡改为对应原因；统计栏行数实时刷新（`paintPdfStats`）
- **改回非 0**：价格按快照（无快照用 `calcValue`）回滚 + 解锁 + 按 Odoo 原值重判红框/「Odoo 原值」小字；动态补挂的备注字段整体移除（不写回备注）、单件箱批发价字段移除；行色/勾选/数量核对列还原
- **静态口径同步**：源数据即「包装数量为空/0」的匹配行，现在数量核对列同样显示「缺货」（此前显示「0 ≠ N」）、行标红、统计栏计入缺货行
- 重构：`renderPdfPreviewRow` 的单格渲染抽成 `buildFieldCell`（渲染与动态补挂共用），数量核对列抽成 `paintQtyCell`/`qtyCellTipHtml`

### 6.3 供应商数据导入流（v3.8 / manifest v1.21.0；前身 = v1.7.0「Excel 导入流」）

> **v3.8（2026-09-09 用户需求）**：从 git 历史恢复 **PDF 解析**（v3.2 曾整体移除），并把上传界面统一为「**供应商数据导入**」——
> **步骤1 上传供应商数据，脚本自动判别文件类型**：`.pdf` → 走 **PDF 分支**（pdfjs `parsePdf` → `extractPdfTable`）；`.xlsx/.xls` → 走 **Excel 分支**（`parseConvertedPdfExcel`，即「PDF 转换版 Excel」）。
> 两分支输出同构 `{format, coupon, rows}`，**可混合多个文件统一合并**（合并/匹配/预览/写回零改动）。
>
> **历史（v1.7.0，2026-08-19）**：Excel 入口上传的是「PDF 转换版 Excel」（如 `Castlers Box 08.27_Order Confirmation_35548880.xlsx`），比对逻辑与 PDF 流一模一样。
> ⚠️ v1.6.0 旧逻辑（上传采购订单 Excel → `applyExcelChanges` 比对 Odoo 现有值）**已移除**。

```
步骤0 选入口（单件📦 / 套装🎁）——renderModePicker
步骤1 上传供应商数据（PDF 或 Excel，可多个，v1.12.0 多文件）→ detectSupplierKind 自动识别：
      .pdf → parsePdf（pdfjs，v3.8 恢复）｜.xlsx/.xls → parseConvertedPdfExcel → 点「下一步」mergeConvFiles 合并
步骤2 上传采购订单 Excel → parseOrderExcel → applyPdfByMode（同一分发）
步骤3 预览 buildPdfPreviewModal（数量核对/供应商 Qty 布局）→ executePdfUpdate 写回
```

**多供应商数据文件（v1.12.0，2026-08-25；v3.8 扩展支持 PDF/Excel 混合；v3.9 支持原生多选）**：
- 上传区支持**反复添加**多个文件：已添加文件列表（类型图标/名称/行数/删除单个）+「继续添加」+「下一步」合并
- **原生多选（v3.9）**：`makeDropZone(..., multi=true)` → 文件选择框带 `multiple`、拖拽取 `dataTransfer.files` 全部文件；逐个**类型校验**（不支持的文件汇总提示并跳过）后**串行**解析（`onFile` 返回 Promise 则等上一个完成，避免多个 PDF 同时解析卡住页面）。采购订单 Excel / 商品库两个上传区仍为单选（多余文件提示后取第一个）
- `excelState.convFiles[]` 每文件 `{name, kind, rows, format, coupon}`（**kind = 'pdf' | 'excel'**，v3.8 新增）
- **合并规则**（`mergeConvFiles`，用户确认）：同 UPC+包装去重（key 与拆分匹配键一致 = `UPC+description的(xN)`，提取不到 = 原 UPC）；合并行 **Qty 相加、Subtotal 相加**，UNIT PRICE/catalog/description 取第一个
- **Coupon**（`mergeCoupon`）：任一文件非 0 → 取该值（触发 ×0.9）；全 0 → 0
- 回步骤①保留已添加列表可编辑；「重置」全清；单文件解析失败仅提示不影响已添加文件
- 合并后 rows 直接喂 `applyPdfByMode`，匹配/预览/写回/数量核对零改动

**解析双分支（输出同构）**：
- **PDF 分支（v3.8 恢复）**：`parsePdf(arrayBuffer)`（pdfjs 拉取文本块，跨页 y 偏移）→ `extractPdfTable(items)`（表头定位/格式识别/行列聚类/列边界校准）→ `extractCoupon(items)`。**依赖**：manifest content_scripts 加载 `lib/pdf.min.js`（全局 `pdfjsLib`），`lib/pdf.worker.min.js` 列入 web_accessible_resources（workerSrc=chrome.runtime.getURL）
- **Excel 分支**：`parseConvertedPdfExcel(data)`——表头行 = 首个含 `UPC|EAN` 关键字的行；格式识别：表头含 `HS CODE` → B，否则 A；列定位按表头名正则（UPC/EAN、Catalog、Qty/QTY/Quantity/QTYs、Unit Price、Subtotal、Product Description），容忍空列与用户附加列；数据行 = UPC 列为 8~14 位纯数字的行；Coupon 从含「Coupon」行右侧首个数值提取
- 返回 `{ format, coupon, rows:[{upc,catalog,qty,unitPrice,subtotal,description}] }` —— 两分支同构，直接喂 `applyPdfByMode`

**写回字段 / 比对规则 / 缺货规则 / Modal 交互**：与修正流（§6.2）完全一致，不再单独维护。**缺货行**（备注「缺货」）单价/整箱批发价置 0 写回（v1.12.4 起两入口统一），详见 §6.2 ⚠️ 块。Modal 的 `mode` 优先取预览行自带标记（`previewRows[0].mode`），兜底 `excelState.mode`。

**样本实测（2026-08-19，node 单测）**：`Castlers Box 08.27_Order Confirmation_35548880.xlsx` → 格式 B、Coupon=-73702.326、64 行、UPC 全部合法、description 含 xN（套装单件数量）、Subtotal=UnitPrice×Qty 验算通过。

### 6.4 商品库更新流（manifest v1.11.0，2026-08-21 用户需求）

> **用户需求（2026-08-21）**：入口在采购单顶部导航「产品 → 产品变体」页（`model=product.product`）；上传商品库 Excel，按 UPC 定位产品变体，用 Excel「品牌」「中文简称」更新 Odoo 对应字段。

- **注入**：`isAnyPage()` = `model=purchase.order` **或** `model=product.product`（v1.11.0 放宽）；产品页默认 Tab = 商品库更新，采购页默认 Excel；hashchange 按页面类型重置默认 Tab
- **Excel 列（`PRODUCT_EXCEL_COLS`，表头第 1 行按名定位 + 正则容错）**：`UPC`（12/13 位数字，按字符串 trim 处理）| `品牌` | `中文简称`（样本实测含 `\n`，解析时转空格）
- **Odoo 字段（`PRODUCT_ODOO_FIELDS`，用户确认均为 Char）**：`default_code`（定位键）| `brand`（品牌）| `name`（中文简称）
- **流程**：上传商品库 Excel（`makeDropZone` 复用）→ `parseProductExcel` → `loadProductByUpc`（search_read `product.product`，domain `default_code in [...]`，UPC 去重后 500/批分块）→ `buildProductPreviewRows` → `buildProductPreviewModal`（UPC + 中文简称/品牌「旧→新」对照 + 勾选）→ `executeProductUpdate`（批量 write `{name, brand}`）
- **行状态**：`ok`（✅ 将更新，**与 Odoo 原值比对有差异**）/ `same`（ℹ️ 值相同，**与 Odoo 原值比对一致，仍照写**）/ `notfound`（⚠️ 未匹配，UPC 无对应产品，注明原因）/ `dup`（🔁 UPC 匹配到多个产品，注明原因需人工）；**读取 Odoo 原 name/brand 与 Excel 比对展示**（用户 2026-08-21 要求），**写回不做比对**（匹配到唯一产品即写，空值也照写）
- **写回**：Excel 为权威源，`{name, brand}` 按 Excel 值**直接写入（空值也照写，无空值保护）**；仅勾选行写回；日志/toast 与双流共用；日志含 old/new 旧→新对照
- **⚠️ Odoo 多语言翻译（2026-08-21 实测修复）**：`product.product.name` 是**多语言翻译字段**（translate），zh_CN 用户界面显示的是**简体中文翻译值**而非 source。扩展读/写 name **必须带 `context.lang="zh_CN"`**：读 → 预览旧值=界面实际显示值；写 → 覆盖中文翻译（source 已=Excel 值无需动）。不带 lang 会导致「数据库 source=Excel 但中文界面仍显示旧翻译」的假象。`brand` 无 translate 不受影响。**修复后需重跑一次全量更新覆盖 79 条已写入产品的中文翻译**
- **样本实测（2026-08-21，openpyxl）**：`极牛产品名称-KVIVA(1).xlsx` → Sheet1 142 行（1 表头 + 141 数据）、3 列、无合并单元格、UPC 12/13 位混存（如 8809640734526 / 880933516775）、中文简称含换行

### 6.5 SKU 更新流（manifest v1.17.0，2026-08-27 用户需求）

> **用户需求（2026-08-27）**：两级匹配（SKU → UPC）中 **UPC 匹配成功**（SKU 匹配失败）的行 = 该商品 SKU 已变更，用**转换 Excel 的 SKU**（pdf.catalog）更新 Odoo 产品变体的商品包装规格 SKU 字段。**SKU 匹配成功的行不更新**（SKU 未变）。
>
> **v3.4（2026-08-27 用户澄清）**：① 单件规格**不需要识别**——入口（单件/套装）决定写 `single_sku` 还是 `box_sku`；② **采购单 Excel 的 SKU 列按入口区分**：单件入口 = 「订单行/包装/SKU」列、套装入口 = 「订单行/包装/Box SKU」列（两个采购单 Excel 同时含这两列且 Box SKU 在前，旧 fuzzy("SKU") 会误命中 Box SKU 列 → 单件 SKU 匹配失效）；③ **套装入口保留 XX 比对**（套装多规格，如 20/60/100 pieces）。

- **Odoo 字段（`PACKAGING_ODOO_FIELDS`，用户确认）**：`product.packaging.name`（包装规格，如 "1 piece" / "1 box of 42 pieces"）、`single_sku`（单件 SKU，piece 规格）、`box_sku`（Box SKU，box 规格）
- **SKU 列按入口区分（v3.4，`parseOrderExcel(data, mode)` + `colIdxCatalog`）**：单件 → `EXCEL_COLS.catalogSingle`（订单行/包装/SKU，fuzzy 兜底**跳过含 "Box" 的列**）；套装 → `EXCEL_COLS.boxCatalog`（订单行/包装/Box SKU）；旧表头「SKU_x」兼容兜底
- **按入口写回**：
  - 单件入口：`matchSkuPackaging` **不识别规格**，直接定位单件记录（**优先 qty=1 且 name 含 piece/each/single/unit → 其次任意 qty=1 → 再次仅 1 条记录兜底**），写回 `single_sku`（防误命中 box 规格 "1 box of XX pieces" 或库内 '1 piece' 但 qty≠1 的脏数据）
  - 套装入口：套装数量 = `parseSetPieces`（description xN 优先 → 包装列 pieces 兜底，与数量公式同源），**比对 "1 box of XX pieces" 的 XX**（`parsePieces(name) === setPieces`），写回 `box_sku`
- **匹配标记**：`matchPdfToExcel` pair 增加 `matchLevel`（`'sku'`/`'upc'`）；UPC 匹配（含拆分匹配）成功且转换 Excel catalog 有值 → changes 带 `skuUpdate {mode, newSku: pdf.catalog, setPieces}`
- **查询**：`loadPackagingByUpc` 按 UPC=default_code 查 product.product → 其全部 product.packaging（id/name/qty/single_sku/box_sku），500/批分块；并入预览缓存 `previewLineCache.packMap`（与 lineData 同生命周期）
- **预览 Modal**：SKU 列（宽 150px）展示「Odoo 原 SKU → 转换 Excel 新 SKU」两行对照（灰 → 紫），hover 提示写入的包装规格；SKU 有实际变更（old ≠ new）的行**默认勾选**；找不到包装规格 → SKU 橙色虚线下划线 + hover 原因（**跳过 SKU 更新，订单行写回不受影响**）
- **失败文案（v3.4 拆分）**：`UPC 未匹配到产品（default_code 无对应）` / `该产品无包装规格记录（product.packaging）` / `未找到单件(piece)包装规格` / `未找到 N 件的 box 包装规格` / `套装数量提取失败，无法定位 box 规格`
- **写回（`executePdfUpdate`）**：勾选行若 skuUpdate.ok 且 old ≠ new，额外 `product.packaging.write({single_sku | box_sku})`，**独立于订单行写回**（订单行 payload 为空但 SKU 变更的行也写）；任一侧失败计入 failed；日志 res.sku 记录 old → new + 规格名

---

## 7. UI 结构

| 组件 | 说明 |
|---|---|
| 可拖动按钮 📥 | 右下角 52px 圆形紫色按钮，位置存 localStorage；拖文件到按钮上提示在卡片内选择入口 |
| 悬浮卡片 | 点按钮展开（340×500px），含 Header、双 Tab、上传区、日志区 |
| 双 Tab | 「🗂 供应商数据导入」/「🏷 商品库更新」（v3.8 由「📊 Excel 导入」改名；`appState.activeTab`，产品变体页默认商品库 Tab，采购页默认供应商导入） |
| 入口选择器 | 步骤 0：单件📦 / 套装🎁 两张卡片（供应商数据导入区用 `renderModePicker`），选中后锁入口，可「切换入口」重置 |
| 步骤指示器 | 供应商数据导入区：① 上传供应商数据（Excel/PDF 自动识别）→ ② 上传采购订单 Excel → ③ 预览确认（当前步紫色、已完成绿色、可点击回退） |
| 常驻错误框 | `excelState.error`：解析失败时在卡片内红色常驻显示原因，下次成功自动清除 |
| 预览 Modal | 94vw 宽居中弹窗，数量核对/供应商 Qty 布局（PDF 与 Excel 来源统一展示）；`mode` 优先取预览行自带标记，兜底 `excelState.mode`；商品库更新 Modal（v1.11.0）：UPC + 中文简称/品牌 旧→新 对照 + 勾选写回。**预览 Odoo 查询缓存（v1.12.3）**：`buildPdfPreviewRows` 缓存 lineData，同一 changes（引用相同）重复预览不重查 Odoo（Modal 关闭重开/反复点预览秒开），数据重新生成（重新上传/合并）自动失效 |
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
| 16 | Excel 导入支持多个转换版文件（列表可删 + 合并去重） | `excelState.convFiles`/`mergeConvFiles`/`mergeCoupon`/`finishConvFiles`/`removeConvFile` | ⚠️ v1.12.0 已实现（17 断言单测通过），待 Chrome 冒烟验证（多文件添加/删除、同 UPC 同包装 Qty 相加、Coupon 任一非0、双流回归） |
| 17 | 数量列改「包装数量」+ 价格原值改比对 Odoo（2026-08-26 用户需求） | `EXCEL_COLS.boxQty`/`applyPdfSingle`/`applyPdfSet`/`buildPreviewRows`/`getOrderLines`/`fieldOrder` | ⚠️ v3.0（manifest v1.13.0）已实现：数量比对与写回用采购单 Excel「包装数量」列（abw交货箱数列作废）；价格原值（单价→price_unit、0.9箱规价→box_wholesale_price、0.9总价→price_subtotal）从 Odoo 订单行取，buildPreviewRows 命中后回填重算 changed/reason；箱规价字段移除；缺货判断 = 包装数量空/0。语法通过，待 Chrome 冒烟验证 |
| 18 | modal 按订单关联分组小计（2026-08-27 用户需求） | `renderPdfPreviewTable`（tbody 分组插行）/`renderPdfGroupSubtotalRow`（组小计行，只累计 total09）/`refreshTotal`（组小计实时刷新） | ⚠️ v3.1（manifest v1.14.0）已实现：套装入口每组尾部插「{订单关联号} 小计」行，只对 0.9总价列组内合计（全部行）；单件入口不插；全表合计行保留。语法通过，待 Chrome 冒烟验证 |
| 19 | 移除 PDF 修正入口（2026-08-27 用户需求；**v3.8 部分回退，见 #25**） | `renderTabSwitch`/`renderCardContent`/`renderExcelZone`/`excelState`（pdfState、renderPdfZone、processPdfFile、extractPdfTable、parsePdf 等已删） | ✅ v3.2（manifest v1.15.0）已实现：Tab 剩 Excel 导入 + 商品库更新；manifest 移除 pdf.min.js/pdf.worker.min.js 加载（lib 文件保留）；共享的 applyPdfByMode/匹配/预览/写回链路由 Excel 流独占。语法通过，副本已同步 |
| 20 | SKU 更新（2026-08-27 用户需求） | `PACKAGING_ODOO_FIELDS`/`loadPackagingByUpc`/`matchSkuPackaging`/`matchPdfToExcel`（matchLevel）/`applyPdfSingle`/`applyPdfSet`（skuUpdate）/`parseOrderExcel`（colIdxCatalog 按入口定列）/`buildPreviewRows`/`renderPdfPreviewRow`（SKU 列对照）/`executePdfUpdate` | ⚠️ v3.3 + v3.4（manifest v1.17.0）已实现：UPC 匹配成功行用转换 Excel SKU 更新 product.packaging——单件入口写 single_sku（不识别规格，qty=1 定位；SKU 列=「订单行/包装/SKU」）、套装入口按套装数量比对 XX 写 box_sku（SKU 列=「订单行/包装/Box SKU」）；SKU 列旧→新对照，变更默认勾选；失败文案拆 5 种；找不到规格跳过并提示。语法通过，副本已同步，待 Chrome 冒烟验证 |
| 21 | modal 0.9总价 改动态计算列（2026-09-04 用户需求） | `applyPdfSingle`/`applyPdfSet`（total09 字段 = 价格×包装数量）/`fieldOrder`（单件入口也含 total09）/`renderPdfPreviewRow`（total09 readonly + 初始值取 newValue）/`recalcRowTotal`（新增：编辑 boxQty/价格 → 重算本行 total09）/`refreshTotal`（总和+组小计随动态值刷新） | ⚠️ v3.6（manifest v1.19.0）+ v3.6.1 补丁（manifest v1.19.1）已实现：0.9总价列两入口都展示（套装=0.9箱规价×包装数量、单件=单价×包装数量，2026-09-04 用户确认），readonly 动态重算；保留 Odoo price_subtotal 原值比对（不一致红框）；单件入口本次起同样插「订单关联 小计」行；total09 odooField=null 不写回。**v3.6.1 补丁（2026-09-04 用户需求）**：Odoo 匹配不上订单行的价格数据（oldValue 无回填）也照常显示计算值——新增 `fieldInitValue`（newValue 优先，无计算值退回 fieldDisplayVal），输入框初值与 colMaxLen 列宽统一用该口径（含 0.9总价/单价/0.9箱规价），无法比对原值的行不显示 Odoo 原值小字。语法通过，副本已同步，待 Chrome 冒烟验证 |
| 22 | Odoo 行匹配加 supplierinfo.product_code 二级键（2026-09-04 用户需求，只读库验证后实施） | `searchPoByRef`（补 partner_id）/`getOrderLines(ids, partnerId)`（查模板级+变体级 supplierinfo → 行级 supplierCodes）/`loadOrderLineMap`（supMap/supCnt 次键注册，defaultCode 空但有 supplierCodes 的行不再丢弃）/`buildPreviewRows`（主键未命中 → 次键，包装精确+唯一兜底+多行防错配） | ⚠️ v3.7（manifest v1.20.0）已实现。验证依据（只读库 p169_erp，2026-09-04）：`8800366240575` 不在 default_code、在 `supplierinfo.product_code`（partner 1354 ABW，tmpl 8608/var 608，default_code=8800256119219）；全库 4096 行中 28 行 line.name 票据码 ≠ default_code，活跃单（purchase 11 + sent 2）13 行，样例 10/10 name `[码]` == supplierinfo.product_code。主键 default_code 优先、不中再按 PO 供应商的 supplierinfo.product_code（含包装精确/唯一兜底/多行报未找到，防错配）。语法通过，副本已同步，待 Chrome 冒烟验证（重点：P09770 等 ABW 单 medicube 防晒霜/TXA 精华行不再报未找到） |
| 23 | 缺货①行「Odoo 原值」删除线小字残留非 0（2026-09-04 用户反馈） | `applyPdfSingle`/`applyPdfSet`（outstock① 分支 unitPrice 字段） | ✅ v3.7.1（manifest v1.20.1）已修复：kind=outstock 行 unitPrice 的 oldValue 原存采购单 Excel 单价列残留值（v3.0 已废 Excel 价格列作比对、kind=outstock 又故意不回填 Odoo 现值），渲染时被标成「Odoo 原值: 非0」小字误导（Odoo 已写 0 仍显示）。修复 = oldValue 置 null（缺货行价格不比对）。语法通过，副本已同步，待 Chrome 冒烟验证 |
| 24 | 0.9总价比对实时化（2026-09-04 用户需求） | `renderPdfPreviewRow`（登记 _total09Box；total09「Odoo 原值」小字改动态、不走静态创建）/`recalcRowTotal`（重算值后触发实时比对）/新增 `syncTotal09Compare`（值 vs Odoo price_subtotal 原值动态红灰态 + 小字增删 + title） | ⚠️ v3.7.2（manifest v1.20.2）已实现：modal 里改包装数量/价格 → 0.9总价 重算后**立即与 Odoo price_subtotal 实时比对**——不一致：红框红底 + 动态删除线「Odoo 原值」小字 + title 说明；改到一致：恢复只读灰样式并移除小字；无 Odoo 原值可比对的行（匹配不上/缺货①）恒灰只展示计算值。语法通过，副本已同步，待 Chrome 冒烟验证 |
| 25 | 恢复 PDF 解析 + 上传界面改「供应商数据导入」（2026-09-09 用户需求，git 提交 ee9b5bc^ 恢复） | `content.js`（extractPdfTable/extractCoupon/parsePdf 恢复；detectSupplierKind 自动识别 .pdf→parsePdf / .xlsx/.xls→parseConvertedPdfExcel；convFiles 条目加 kind；UI 文案/步骤1 上传区改「供应商数据」并支持 PDF/Excel 混合多文件）/`manifest.json`（重新加载 lib/pdf.min.js + web_accessible_resources 加 lib/pdf.worker.min.js） | ⚠️ v3.8（manifest v1.21.0）已实现：PDF 解析链从 v3.2 移除提交的父版本恢复并适配 v3.3-v3.7 演进后的代码；流程归一为 步骤1 供应商数据（自动识别路由）→ 步骤2 采购订单 Excel → 预览写回。语法检查通过，副本已同步，待 Chrome 冒烟验证（重点：文本型 PDF 直传解析行数/格式/Coupon、Excel/PDF 混合多文件合并、路由后匹配写回回归） |
| 26 | Modal 内包装数量改 0 → 按缺货写回 + 供应商数据支持多选上传（2026-09-10 用户需求） | `isRowOutstock`（新增：源数据缺货 或 boxQty 空/0）/`syncRowOutstock`（新增：切换缺货态——快照 + 价格置 0 锁定 + 备注缺货 + 标红勾选 + 还原）/`setRowRemarkOutstock`/`setRowBoxWholesaleZero`/`buildFieldCell`+`mountFieldInput`（单格渲染抽取 + 动态补备注输入框）/`paintQtyCell`+`qtyCellTipHtml`+`paintPdfStats`（数量核对列/统计栏实时刷新）/`executePdfUpdate`（写入时兜底置 0）/`makeDropZone`（multi 参数：原生多选 + 拖入多文件 + 串行解析）/`applyPdfSingle`/`applyPdfSet`（outstock 分支记 `calcValue`） | ⚠️ v3.9（manifest v1.22.0）已实现：① modal 里把「包装数量」改成 0/空 → 该行按缺货处理（备注「缺货」+ 单价/0.9箱规价置 0 锁定，单件补 box_wholesale_price=0 写回；改回非 0 按快照/calcValue 还原）；源数据包装数量为 0 的匹配行同步显示「缺货」。② 供应商数据上传区支持一次多选/多文件拖入，逐个校验后串行解析合并。语法检查通过，副本已同步，待 Chrome 冒烟验证（重点：改 0 → 预览标红/数量核对显示缺货/确认写入后 Odoo 三字段为 0；改回非 0 还原；多选多个 PDF/Excel 混合合并） |
| 27 | 预览报错「预览失败: fieldRow is not a function」（2026-09-10 用户反馈） | `buildFieldCell`（v3.9 从 `renderPdfPreviewRow` 抽出的单格渲染函数） | ✅ v3.9.1（manifest v1.22.1）已修复：`row._fieldRows[key] = fieldRow` 抽取时丢了行尾分号，下一行是 `(function (inp, fld) {...})(input, f)`（IIFE 以 `(` 开头），ASI 不插分号 → 被解析成 `fieldRow(function(){...})(input, f)` 调用 → TypeError。修复 = 补回分号（并加注释说明，全文件已扫描确认无同类隐患）。已用最小 DOM 桩跑通 modal 渲染 + 包装数量改 0/改回全流程（单件/套装两入口），副本已同步，待 Chrome 冒烟 |

---

## 10. 快速上手（新 Agent 3 分钟版）

1. **只改一个文件**：`content.js`（IIFE 单文件）；字段映射在顶部两个常量对象
2. **一条在线修正流程（v3.8 起：供应商数据导入）**：🗂 供应商数据导入 → 步骤1 上传供应商数据（**PDF 或 PDF 转换版 Excel，自动识别路由**：.pdf → `parsePdf`/`extractPdfTable`；.xlsx/.xls → `parseConvertedPdfExcel`；**可多个混合（v1.12.0 多文件合并）**）→ 步骤2 上传采购订单 Excel → `applyPdfByMode` 匹配修正 → Modal 写回；共用 `parseOrderExcel` / `loadOrderLineMap` / `buildPreviewRows` / Modal / `executePdfUpdate`
3. **两级匹配（v2.0）**：供应商数据 ↔ 采购单 Excel 先按 Catalog↔SKU 匹配，匹配不上的行走 UPC 匹配（同 UPC 多行用包装拆分：供应商侧 description "(xN)"、Excel 侧包装列 pieces，拼成 `UPC+N` 精确配对；唯一 UPC 直接匹配）；数量核对按 matchKey（Catalog/UPC/拆分键）分组。**UPC+包装是 Odoo 匹配键（v1.9）**：Excel「内部参考号」+「订单行/包装」件数 ↔ Odoo `product.default_code` + `product_packaging_id.qty`；查 PO（=name）优先用「参考号」列、查不到再用「订单关联」列（v1.7.1）
4. **无构建、无测试**：改完同步到 `odoowritting_extension/` 子目录后到 Chrome 手动验证
5. **v3.0 比对基准**：数量 = 采购单 Excel「包装数量」列（求和 vs 供应商 Qty，写回 product_packaging_qty）；价格原值 = Odoo 订单行（price_unit / box_wholesale_price / price_subtotal），采购单 Excel 价格列不再作比对基准；箱规价字段已移除
6. **SKU 更新（v3.3）**：两级匹配中 UPC 匹配成功（SKU 匹配失败）→ 用供应商数据 SKU 更新 product.packaging——单件入口写 piece 规格的 `single_sku`；套装入口按套装数量（description xN）比对 "1 box of XX pieces" 写 `box_sku`；Modal SKU 列旧→新对照，变更默认勾选，找不到规格跳过并提示
7. Git 提交必须过 commitlint（husky），格式 `type(scope): subject`

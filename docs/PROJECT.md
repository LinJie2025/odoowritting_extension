# Odoo Excel Importer — 项目文档

> 本文件是项目的**权威文档**，与代码同步维护。任何 Agent 接任务前必须先读本文件。
> 最后更新：2026-08-18（manifest v1.5.0）

---

## 1. 项目概述

| 项 | 内容 |
|---|---|
| 名称 | Odoo Excel Importer（`odoowritting-extension`） |
| 类型 | Chrome 扩展（Manifest V3），content script 注入式工具 |
| 业务目的 | 把采购 Excel / 供应商 PDF 数据批量写入 Odoo 采购订单（RFQ）明细行 |
| 技术栈 | 纯原生 JS（**ES5 风格：`var` + IIFE + function 声明**），无框架、无构建步骤 |
| 第三方库 | `lib/xlsx.full.min.js`（SheetJS）、`lib/pdf.min.js` + `lib/pdf.worker.min.js`（pdfjs-dist **3.11.174 UMD 版**，4.x 起无 UMD 故锁定 3.x） |
| 运行环境 | 任意 Odoo 实例的 `purchase.order` 页面（URL hash 含 `model=purchase.order`） |
| 版本 | manifest v1.5.0 |

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
| `pack` | 订单行/包装 | 取 "1 box of XX pieces" 的 XX（降级用） |
| `boxPrice` | 箱规价 | 套装比对列1（不写回） |
| `boxPrice09` | 0.9箱规价 | 套装比对列 → box_wholesale_price |
| `total09` | 0.9总价 | 套装比对列（不写回） |
| `unitPrice` | 单价 | 单价 → price_unit |
| `remark` | 备注 | 缺货等 |
| `orderRef` | 订单关联 | 查 PO（=name） |

### 3.3 其他

- `PREFIX = "__odoi_"`：所有 DOM id / class / localStorage key 前缀
- `LS_POS`：按钮位置 localStorage key；`LS_LOG`：最近日志 key
- `PDF_MODES`：`single`（单件入口📦）/ `set`（套装入口🎁）

---

## 4. 架构分层（content.js 内部结构）

单个 IIFE 内按 `═══` 注释分块，共 11 块：

| 块 | 内容 | 关键函数 |
|---|---|---|
| 常量 | 字段映射、列名、前缀 | — |
| 工具 | DOM 创建 `el()`、HTML 转义 `escHtml()`、`sleep`、价格计算 | `calcBoxPrice`（×0.9，整数分运算防浮点误差）、`parsePieces`（从 "1 box of 20 pieces" 取 20） |
| Odoo API | JSON-RPC 封装 | `rpcCall`、`searchPoByName`（Excel 流重写时备用）、`searchPoByRef`、`getOrderLines`、`updateOrderLine` |
| Excel 解析 | ~~两种解析器~~ | ⚠️ `parseExcel` 已移除（2026-08-14）；`parseOrderExcel`（按表头列名）仅 PDF 流使用 |
| PDF 解析 | 文本块坐标聚类成表 | `parsePdf`（pdfjs 入口）、`extractPdfTable`（核心算法） |
| 业务逻辑 | 匹配修正规则 | `applyPdfByMode`（入口分发）、`applyPdfSet`（套装，v1.5 统一公式）、`applyPdfSingle`（单件） |
| 日志 | localStorage 存取 | `saveLog`/`getLog`/`clearLog`/`downloadLog` |
| 状态 | 应用状态对象 | `appState`、`pdfState`、`dragCtx` |
| UI | 按钮/卡片/Tab/上传区/日志区 | `createDraggableButton`、`renderCardContent`、`renderTabSwitch`、`renderPdfZone`、`renderPdfModePicker` |
| 预览 Modal | PDF 流 1 套 | `buildPdfPreviewModal`（PDF 流）；⚠️ Excel 流预览 Modal/标签条已移除待重写 |
| 执行写回 | PDF 流批量写入 | `executePdfUpdate`（PDF 流）；⚠️ Excel 流 `executeUpdate` 已移除待重写 |
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
| `getOrderLines(ids)` | 批量取订单行 | fields: `id,name,price_unit,box_wholesale_price,remark`；**从 `line.name` 正则 `\[(\d+)\]` 提取 UPC 作为匹配键** |
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
> - 按表头有无 `HS CODE` 识别格式（保留，仅用于匹配键列名）：有 HS → `UPC/EAN NUMBER`+`CATALOG NO.`；无 HS → `UPC/EAN`+`Catalog#`，任一命中
> - ×0.9 开关：`Coupon=0` → factor=1（不打折）；`Coupon≠0`（负数）→ factor=0.9
> - 数量核对：Excel「abw交货箱数」多条求和 vs **PDF Qty 原值**（不再分格式 ÷pieces，用户 8-18 确认「abw交货箱数就是Qty列的数」）；不一致 modal 提示（不自动扣减）
> - 缺货备注：① abw交货箱数为空/0；② Excel 有而 PDF 无 → 都备注「缺货」写 remark
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

- 匹配：UPC/Catalog 任一命中（同全局规则）
- 数量核对：abw交货箱数求和 vs PDF Qty 原值（同全局规则），不一致 modal 提示
- 单价：`UNIT PRICE × factor` → 比对「单价」列 → 写回 `price_unit`（不分格式）
- 缺货备注：`remark`
- 字段带 `reason`/`pdfSource` 供 modal 对照展示（v1.5 补充，公式未变）

#### PDF 流写回（`executePdfUpdate`）

- 写回字段：`product_packaging_qty`（abw交货箱数）/ `box_wholesale_price`（0.9箱规价）/ `price_unit`（单价）/ `remark`（备注），经 ODOO_FIELDS 映射
- 只处理勾选行；预览行标红场景：数量不一致（qtyMismatch）、缺货（outstock）、匹配失败（error）
- 每条变更字段带 `odooField`，为 null 的字段（箱规价/0.9总价）仅比对不写回

---

## 7. UI 结构

| 组件 | 说明 |
|---|---|
| 可拖动按钮 📥 | 右下角 52px 圆形紫色按钮，位置存 localStorage；拖文件到按钮上可直接导入 |
| 悬浮卡片 | 点按钮展开（340×500px），含 Header、双 Tab、上传区、预览窗口标签条、日志区 |
| 双 Tab | 「📊 Excel 导入」/「📄 PDF 修正」（`appState.activeTab`） |
| 入口选择器 | PDF 区步骤 0：单件📦 / 套装🎁 两张卡片，选中后锁入口，可「切换入口」重置 |
| 步骤指示器 | 选中入口后显示：① 上传 PDF → ② 上传 Excel → ③ 预览确认（当前步紫色、已完成绿色） |
| 常驻错误框 | `pdfState.error`：PDF/Excel 解析失败时在卡片内红色常驻显示原因（如「未识别到表头——请确认 PDF 为文本型而非扫描件」），下次成功自动清除 |
| 预览 Modal | 92vw 宽居中弹窗，Excel 流可最小化（点遮罩最小化、标签条恢复）；PDF 流仅关闭 |
| 标签条 | 每个 Excel 文件一个 tab（`appState.tabs`），最小化的 Modal 可从标签恢复 |
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
| 1 | ~~Excel 导入分支处理逻辑~~ | — | ✅ 已彻底移除该 Tab（v1.4.0），只留 PDF 修正 |
| 2 | ~~单件入口逻辑未实现~~ | `applyPdfSingle` | ✅ 已实现（v1.4.0，UPC/Catalog 匹配 + 单价×factor + 数量求和提示） |
| 3 | ~~Excel 的 Catalog 列名待确认~~ | `EXCEL_COLS.catalog` | ✅ 已确认 = 「SKU_x」（2026-08-17 实测） |
| 4 | Excel 公式列（0.9箱规价/0.9总价/单价）缓存值能否被 SheetJS 读到 | `parseOrderExcel` | ⚠️ 需真实 Excel 样本验证 |
| 5 | `searchPoByName` 已无用（Excel 流移除后无调用方） | Odoo API 层 | ℹ️ 可删除，保留无害 |
| 6 | `ODOO_FIELDS.qty`（product_qty）已不写回 | content.js | ℹ️ 保留备用 |
| 7 | ~~子目录打包副本版本落后~~ | `odoowritting_extension/` | ✅ 已同步（v1.5.0） |
| 8 | ~~Odoo 字段名待实测~~ | `ODOO_FIELDS` | ✅ 已确认：`product_packaging_qty` / `box_wholesale_price` / `price_unit` / `remark` |
| 9 | ~~套装分格式公式/随机扣减~~ | `applyPdfSet` | ✅ v1.5.0 重写：统一公式 + 4 项比对 + 套装单件数量（PRODUCT DESCRIPTION "x30"） |
| 10 | modal 增强（订单关联统计/PDF 对照/hover 原因/列总和） | `buildPdfPreviewModal` 等 | ⚠️ 已实现 v1.5.0，待 Chrome 冒烟验证 |

---

## 10. 快速上手（新 Agent 3 分钟版）

1. **只改一个文件**：`content.js`（~1436 行，IIFE 单文件）；字段映射在顶部两个常量对象
2. **当前只有 PDF 修正流在线**：`parseOrderExcel`（按列名）+`searchPoByRef`+`executePdfUpdate`；**Excel 导入 Tab 是占位**（逻辑已移除待重写，见 §9 #1）
3. **UPC 是核心匹配键**：Excel 列「订单行/产品/内部参考号」↔ Odoo 行 name 中 `[数字]` ↔ PDF 表 UPC 列
4. **无构建、无测试**：改完同步到 `odoowritting_extension/` 子目录后到 Chrome 手动验证
5. **未实现的功能**：单件入口（§9 #2）＋ Excel 导入流（§9 #1），接相关任务先找用户确认规则
6. Git 提交必须过 commitlint（husky），格式 `type(scope): subject`

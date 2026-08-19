# AGENTS.md — Agent 工作室必读（本工作空间所有 Agent 的入口文档）

> ⚠️ **强制约定**：任何 Agent（含子 Agent、专家、自动化任务）在本工作空间接任务前，**必须先阅读本文件与 `docs/PROJECT.md`**（完整项目文档）。先读文档，再动代码。

---

## 项目一句话

**Odoo Excel Importer** — Chrome 扩展（Manifest V3），把采购 Excel / 供应商 PDF 数据批量写入 Odoo 采购订单（`purchase.order`）明细行（数量/包装数量/整箱批发价/备注）。

## 必读文件（按顺序）

1. **本文件**（AGENTS.md）— 关键事实速查
2. **`docs/PROJECT.md`** — 权威完整文档：架构分层、两条业务流程、字段映射、UI、工程规范、已知待办
3. 需要时再查：`.workbuddy/memory/MEMORY.md`（历史决策记忆）、`GIT_WORKFLOW.md`（Git 规范）

## 关键事实速查

| 项 | 内容 |
|---|---|
| 代码位置 | **只有一个业务文件**：根目录 `content.js`（~2200 行，ES5 风格 IIFE，无构建步骤） |
| 第三方库 | `lib/xlsx.full.min.js`、`lib/pdf.min.js`+`lib/pdf.worker.min.js`（pdfjs-dist 3.11.174 UMD，勿升级 4.x） |
| 打包副本 | `odoowritting_extension/` 子目录 = Chrome「加载已解压」用产物；**改完根目录必须同步过去** |
| 字段映射 | 集中在 content.js 顶部 `ODOO_FIELDS`（Odoo 字段名，**2026-08-14 用户确认**：包装数量=`product_packaging_qty`、整箱批发价=`box_wholesale_price`、单价=`price_unit`、备注=`remark`、订单关联=`name`）与 `EXCEL_COLS`（Excel 表头名），**改字段只改这两个常量** |
| 匹配键 | 主匹配（v1.7.2）：PDF `catalog`（CATALOG NO./Catalog#）↔ Excel「SKU」列；Excel 行 SKU 缺失 → PDF `UPC` ↔ Excel「内部参考号」列兜底；Odoo 行 name 中 `\[(\d+)\]` ↔ UPC |
| 两条流程 | ✅ **双 Tab（v1.7.0）**：📊 **Excel 导入**（上传「PDF 转换版 Excel」→ `parseConvertedPdfExcel` 解析 → 再传采购订单 Excel → 与 PDF 同逻辑 `applyPdfByMode` → Modal 写回）+ 📄 **PDF 修正**（上传 PDF+Excel → `applyPdfByMode` 匹配修正 → Modal 写回）；两流在 `applyPdfByMode` 汇合，共用 `parseOrderExcel` / `loadOrderLineMap` / `buildPreviewRows` / `buildPdfPreviewModal` / `executePdfUpdate` |
| 注入条件 | 仅 URL hash 含 `model=purchase.order` 的页面注入按钮 |
| 价格规则 | 整箱批发价 = PDF Subtotal × 0.9，整数分运算（`calcBoxPrice`） |
| 缺货写回 | ⚠️ **v1.8.0**：**Excel 入口**缺货行（备注「缺货」）的**单价/整箱批发价不写回 Odoo**（预览仍显示计算值）；PDF 入口不变。实现：`applyPdfByMode` 传 `source`（'pdf'/'excel'），缺货行价格字段 `odooField` 置 null |
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
- [ ] **Chrome 冒烟验证 v1.8.0**：双 Tab 切换、Excel 流（转换版解析/两步上传/预览/缺货行写回）、PDF 流回归
- [ ] `searchPoByName`/`searchPoByRef` 去重合并（低优先，保留无害）

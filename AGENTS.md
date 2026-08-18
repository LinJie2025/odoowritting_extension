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
| 代码位置 | **只有一个业务文件**：根目录 `content.js`（1866 行，ES5 风格 IIFE，无构建步骤） |
| 第三方库 | `lib/xlsx.full.min.js`、`lib/pdf.min.js`+`lib/pdf.worker.min.js`（pdfjs-dist 3.11.174 UMD，勿升级 4.x） |
| 打包副本 | `odoowritting_extension/` 子目录 = Chrome「加载已解压」用产物；**改完根目录必须同步过去** |
| 字段映射 | 集中在 content.js 顶部 `ODOO_FIELDS`（Odoo 字段名，**2026-08-14 用户确认**：包装数量=`product_packaging_qty`、整箱批发价=`box_wholesale_price`、单价=`price_unit`、备注=`remark`、订单关联=`name`）与 `EXCEL_COLS`（Excel 表头名），**改字段只改这两个常量** |
| 匹配键 | UPC：Excel「订单行/产品/内部参考号」↔ Odoo 行 name 中 `\[(\d+)\]` ↔ PDF 表 UPC 列 |
| 两条流程 | ⚠️ **Excel 导入流已整体移除**（2026-08-14，等用户给新规则后重写，Tab 现为占位）；**PDF 修正流在线**：`parseOrderExcel`（按列名）+ `searchPoByRef` + `executePdfUpdate` |
| 注入条件 | 仅 URL hash 含 `model=purchase.order` 的页面注入按钮 |
| 价格规则 | 整箱批发价 = PDF Subtotal × 0.9，整数分运算（`calcBoxPrice`） |
| 套装扣减 | 多条包装数量求和 > PDF Qty 时随机扣减；Excel 有 PDF 无 → 备注「缺货」 |
| ❌ 未实现 | **PDF 单件入口**：`applyPdfSingleToExcel` 是 throw 占位，**规则已确认（2026-08-14）**：UPC 匹配 + 包装数量求和/随机扣减 + 第2页 Subtotal vs Order Total after Store Credit 决定 Unit Price in HKD 是否 ×0.9（详见 docs/PROJECT.md §6.2） |
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

- [ ] **与用户对齐 Excel 导入修改逻辑并重写该分支**（已整体移除：解析/匹配/预览/写回均待重写；旧实现可参考 git 历史 `58d1f97`，旧逻辑要点见 docs/PROJECT.md §6.1）
- [ ] **实现 PDF 单件入口**（`applyPdfSingleToExcel`）— 规则已确认（2026-08-14，docs/PROJECT.md §6.2）；实现前先与用户对齐 §6.2 文末歧义点（Excel「单价」表头名、单件是否标缺货、解析器需扩展提取 Unit Price in HKD + 第2页汇总值）
- [ ] `searchPoByName`/`searchPoByRef` 去重合并（Excel 流重写时一并处理）

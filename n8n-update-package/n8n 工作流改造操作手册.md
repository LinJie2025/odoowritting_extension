# n8n 工作流改造操作手册（tp7bvgKQSutq1O4c）

> 改造内容：① 新增主题过滤 ② 差异对比（有差异才写回、无差异跳过）③ 输出改动报告
> 生成时间：2026-09-02

## 改动总览

| # | 节点 | 动作 | 说明 |
|---|---|---|---|
| 1 | Email Trigger (IMAP) | 不改 | 保持现状 |
| 2 | **主题过滤**（新） | **新增节点** | `subject` 包含 "Abw Castlers Inventory Report" 才放行（兜底服务器 SUBJECT 过滤失效问题） |
| 3 | 拆出附件 / 解析 Excel / 整理行 | 不改 | — |
| 4 | Odoo 查询 | **替换代码** | 补读现值字段：`brand`/`origin_country`（product.product）、`single_sku`/`box_sku`（product.packaging） |
| 5 | 计算变更集 | **替换代码** | 逐字段对比现值：有差异才进写回队列；无差异标记 skipped；生成 report（每行：状态 + 字段级 old→new） |
| 6 | Odoo 写回 | **替换代码** | 结果透传 report，写回明细带 oldValue |
| 7 | 组装结果 / JSONL 落地 | 已删（线上确认） | 无需操作 |

新链：`Email Trigger → 主题过滤 → 拆出附件 → 解析 Excel → 整理行 → Odoo 查询 → 计算变更集 → Odoo 写回`

---

## 方案 A：在线手动修改（推荐，保留原工作流与 IMAP 凭据）

### 步骤 1：新增「主题过滤」节点

1. 打开工作流，点击画布空白处 → 搜索 `Filter` → 添加 **Filter** 节点，命名为 `主题过滤`
2. 条件配置：
   - **Conditions** 改为 `String` / `contains`（包含）
   - 左值：`={{ $json.subject }}`
   - 右值：`Abw Castlers Inventory Report`
   - 大小写：不区分
3. 接线：把 **Email Trigger (IMAP)** 的输出连到 **主题过滤**，**主题过滤** 的输出连到 **拆出附件**

（等效节点参数 JSON 见 `主题过滤-节点参数.json`，如粘贴失败按上述填写即可）

### 步骤 2：替换「Odoo 查询」代码

打开 Odoo 查询节点 → Code 编辑器 → 全选删除 → 粘贴 `Odoo 查询.jsCode.txt` 全部内容 → 保存

### 步骤 3：替换「计算变更集」代码

同上，粘贴 `计算变更集.jsCode.txt`

### 步骤 4：替换「Odoo 写回」代码

同上，粘贴 `Odoo 写回.jsCode.txt`

### 步骤 5：保存并重新激活

1. 点右上角 **Save**
2. **Deactivate → Activate**（IMAP 长连接必须重启才加载新配置）

---

## 方案 B：导入完整 JSON（更快，但会新建工作流）

1. 文件：`n8n-odoo-sku-update.workflow.json`（本包同目录）
2. n8n 编辑器：右上角菜单 → **Import from JSON/File** → 选择该文件
3. 会生成**新工作流**（名字相同），需要：
   - 删除旧工作流 `tp7bvgKQSutq1O4c`
   - 在新工作流 Email Trigger 节点上**重新选择 IMAP 凭据**（导入不携带凭据绑定）
4. 激活新工作流

> ⚠️ 方案 B 会丢 IMAP 凭据绑定、需要重建，仅在方案 A 操作不便时使用。

---

## 验证方法（改完后自行验证）

1. **主题过滤**：发一封主题不含 "Abw Castlers Inventory Report" 的邮件 → 执行记录里「主题过滤」应拦截（不进入拆出附件）；发主题匹配的 → 正常通过
2. **差异对比**：发一份库存 Excel 跑一轮，看最终节点「Odoo 写回」输出：
   - `report[]`：每行 `upc / mode / status(updated|skipped|notfound|dup) / reason / changes[]`（changes 里是字段级 `field / old / new`）
   - `stats`：`updated` / `skipped` / `notfound` / `dup` / `noPack` / `noSku` 计数
   - `results`：实际写回成功/失败明细
3. **无差异跳过**：同一份 Excel 再跑一次 → `report` 全部为 `skipped`，`skuWrites`/`productWrites` 为空，Odoo 数据不重复写入
4. **对照**：抽查一条 `updated` 记录的 UPC，在 Odoo 后台核对 `brand`/`origin_country`/`single_sku`/`box_sku` 是否等于 report 里的 `new` 值

---

## 差异对比规则（计算变更集）

- 产品字段：Excel `Brand`/`Country` **有值才对比**；与 Odoo 现值不同 → 写回，相同或 Excel 为空 → 跳过
- SKU 字段：Excel `Sku` 有值才对比；单件找 `qty=1` 包装对比 `single_sku`，套装按包装数量找包装对比 `box_sku`；找不到包装规格 → 该行 SKU 不更新（reason 注明）
- 状态定义：`updated`=有差异待写回；`skipped`=与 Odoo 一致；`notfound`=Odoo 无此 UPC；`dup`=同 UPC 多产品

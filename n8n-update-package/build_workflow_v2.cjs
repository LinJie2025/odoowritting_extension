// 构建新版 n8n 工作流：加 Filter + 差异对比 + 跳过报告
const fs = require('fs');
const crypto = require('crypto');
const SRC = 'E:/Download/邮件附件→Odoo 商品SKU_品牌_原产地更新 (1).json';
const OUT = 'D:/workingfile/ABWaddingET/odoowritting_extension/n8n-odoo-sku-update.workflow.json';

const wf = JSON.parse(fs.readFileSync(SRC, 'utf8'));

// ---------- 1. Odoo 查询：补读现值字段 ----------
const odooQuery = `const base = 'http://47.119.112.65:8069';
const db = 'castlers-260728';
const login = 'castlers';
const password = 'ca8689cfaa99bf47ceda3aff3749782e0fddb15c';
const env = { base_location: base, lang: 'en_US', timezone: 'Asia/Shanghai', user_agent: 'n8n-sku-updater/1.0' };
const helpers = this.helpers;

async function rpc(service, method, args) {
  const res = await helpers.httpRequest({
    url: base + '/jsonrpc', method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ jsonrpc: '2.0', method: 'call', params: { service: service, method: method, args: args } }),
    ignoreResponseCodeErrors: true
  });
  let body = res;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) {} }
  if (body && body.error) {
    const d = body.error.data || {};
    throw new Error(d.message || body.error.message || JSON.stringify(body.error));
  }
  return body ? body.result : null;
}

const uid = await rpc('common', 'authenticate', [db, login, password, env]);

async function callKw(model, method, args, kwargs) {
  return rpc('object', 'execute_kw', [db, uid, password, model, method, args, kwargs || {}]);
}

const items = $input.all();
const upcSet = {};
for (const it of items) { const u = it.json.upc; if (u) upcSet[u] = true; }
const upcs = Object.keys(upcSet);
const prodMap = {};
for (let s = 0; s < upcs.length; s += 500) {
  const slice = upcs.slice(s, s + 500);
  const rows = await callKw('product.product', 'search_read', [], { domain: [['default_code', 'in', slice]], fields: ['id', 'default_code', 'name', 'brand', 'origin_country'] });
  for (const r of rows) {
    const c = String(r.default_code || '').trim();
    if (!c) continue;
    if (!prodMap[c]) prodMap[c] = [];
    prodMap[c].push({ id: r.id, name: r.name || '', brand: r.brand || '', origin_country: r.origin_country || '' });
  }
}
const ids = [];
for (const k in prodMap) { for (const p of prodMap[k]) ids.push(p.id); }
const packsByProduct = {};
for (let s = 0; s < ids.length; s += 500) {
  const slice = ids.slice(s, s + 500);
  const rows = await callKw('product.packaging', 'search_read', [], { domain: [['product_id', 'in', slice]], fields: ['id', 'name', 'qty', 'product_id', 'single_sku', 'box_sku'] });
  for (const r of rows) {
    const pid = (r.product_id && r.product_id.length) ? r.product_id[0] : null;
    if (!pid) continue;
    if (!packsByProduct[pid]) packsByProduct[pid] = [];
    packsByProduct[pid].push({ id: r.id, name: r.name || '', qty: r.qty, single_sku: r.single_sku || '', box_sku: r.box_sku || '' });
  }
}
const out = [];
for (const it of items) {
  const j = it.json;
  const list = prodMap[j.upc] || null;
  const productIds = list ? list.map(function (p) { return p.id; }) : [];
  let prod = null;
  let packs = [];
  if (productIds.length === 1) {
    prod = list[0];
    packs = packsByProduct[productIds[0]] || [];
  }
  out.push({ json: { fileName: j.fileName, upc: j.upc, mode: j.mode, packN: j.packN, sku: j.sku, brand: j.brand, country: j.country, productIds: productIds, prod: prod, packs: packs } });
}
return out;`;

// ---------- 2. 计算变更集：真对比，有差异才写 ----------
const diffCalc = `const items = $input.all();
const skuWrites = [];
const productWrites = [];
const report = [];
const stats = { total: 0, notfound: 0, dup: 0, noSku: 0, noPack: 0, updated: 0, skipped: 0, productUpdated: 0, productSkipped: 0, skuUpdated: 0, skuSkipped: 0 };
const fileNames = {};

function norm(v) { return (v === undefined || v === null) ? '' : String(v).trim(); }
function same(a, b) { return norm(a) === norm(b); }

for (const it of items) {
  const r = it.json;
  stats.total++;
  if (r.fileName) fileNames[r.fileName] = true;
  const entry = { upc: r.upc, mode: r.mode, status: '', reason: '', changes: [] };

  if (!r.productIds || r.productIds.length === 0) {
    entry.status = 'notfound'; entry.reason = 'Odoo 未找到该 UPC 的产品'; stats.notfound++; report.push(entry); continue;
  }
  if (r.productIds.length > 1) {
    entry.status = 'dup'; entry.reason = 'Odoo 存在多个同 UPC 产品，无法确定目标'; stats.dup++; report.push(entry); continue;
  }
  const pid = r.productIds[0];
  const prod = r.prod || {};
  const extraTips = [];

  // 产品字段对比（brand / origin_country）：Excel 有值才对比，不同才写
  const vals = {};
  if (r.brand) {
    if (!same(prod.brand, r.brand)) { vals.brand = r.brand; entry.changes.push({ field: 'brand', old: prod.brand || '', new: r.brand }); }
  }
  if (r.country) {
    if (!same(prod.origin_country, r.country)) { vals.origin_country = r.country; entry.changes.push({ field: 'origin_country', old: prod.origin_country || '', new: r.country }); }
  }
  if (Object.keys(vals).length) { productWrites.push({ id: pid, upc: r.upc, vals: vals }); stats.productUpdated++; }
  else { stats.productSkipped++; }

  // SKU 对比（single_sku / box_sku）
  if (r.sku) {
    let p = null, field = '';
    if (r.mode === 'single') {
      field = 'single_sku';
      p = (r.packs || []).find(function (pk) { return Number(pk.qty) === 1; });
    } else {
      field = 'box_sku';
      const N = r.packN;
      p = (r.packs || []).find(function (pk) { const mm = (pk.name || '').match(/1 box of (\\d+) pieces/i); return mm && parseInt(mm[1], 10) === N; })
        || (r.packs || []).find(function (pk) { return Number(pk.qty) === N; });
    }
    if (!p) {
      extraTips.push('找不到匹配包装规格，SKU 不更新'); stats.noPack++;
    } else if (!same(p[field], r.sku)) {
      skuWrites.push({ id: p.id, upc: r.upc, sku: r.sku, field: field, spec: p.name, oldValue: p[field] || '' });
      entry.changes.push({ field: field, old: p[field] || '', new: r.sku });
      stats.skuUpdated++;
    } else {
      extraTips.push('SKU 与 Odoo 一致'); stats.skuSkipped++;
    }
  } else {
    extraTips.push('Excel 无 SKU，不处理'); stats.noSku++;
  }

  if (entry.changes.length) { entry.status = 'updated'; entry.reason = '有差异，待写回'; stats.updated++; }
  else { entry.status = 'skipped'; entry.reason = '与 Odoo 现有值一致，跳过'; stats.skipped++; }
  if (extraTips.length) entry.reason += '；' + extraTips.join('；');
  report.push(entry);
}
return [{ json: { skuWrites: skuWrites, productWrites: productWrites, report: report, stats: stats, fileNames: Object.keys(fileNames) } }];`;

// ---------- 3. Odoo 写回：末尾透传 report ----------
const odooWrite = `const base = 'http://47.119.112.65:8069';
const db = 'castlers-260728';
const login = 'castlers';
const password = 'ca8689cfaa99bf47ceda3aff3749782e0fddb15c';
const env = { base_location: base, lang: 'en_US', timezone: 'Asia/Shanghai', user_agent: 'n8n-sku-updater/1.0' };
const helpers = this.helpers;

async function rpc(service, method, args) {
  const res = await helpers.httpRequest({
    url: base + '/jsonrpc', method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ jsonrpc: '2.0', method: 'call', params: { service: service, method: method, args: args } }),
    ignoreResponseCodeErrors: true
  });
  let body = res;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) {} }
  if (body && body.error) {
    const d = body.error.data || {};
    throw new Error(d.message || body.error.message || JSON.stringify(body.error));
  }
  return body ? body.result : null;
}

const uid = await rpc('common', 'authenticate', [db, login, password, env]);

async function callKw(model, method, args, kwargs) {
  return rpc('object', 'execute_kw', [db, uid, password, model, method, args, kwargs || {}]);
}

const j = $input.first().json;
const skuWrites = j.skuWrites || [];
const productWrites = j.productWrites || [];
const results = { sku: [], product: [], errors: [] };
const groups = {};
for (const w of productWrites) {
  const k = JSON.stringify(w.vals);
  if (!groups[k]) groups[k] = [];
  groups[k].push(w.id);
}
for (const k in groups) {
  const vals = JSON.parse(k);
  try {
    await callKw('product.product', 'write', [groups[k], vals], {});
    results.product.push({ ids: groups[k], vals: vals, status: 'ok' });
  } catch (e) {
    results.errors.push({ kind: 'product', ids: groups[k], error: String(e.message || e) });
  }
}
for (const w of skuWrites) {
  const payload = {};
  payload[w.field] = w.sku;
  try {
    await callKw('product.packaging', 'write', [[w.id], payload], {});
    results.sku.push({ id: w.id, field: w.field, sku: w.sku, oldValue: w.oldValue || '', status: 'ok' });
  } catch (e) {
    results.errors.push({ kind: 'packaging', id: w.id, field: w.field, error: String(e.message || e) });
  }
}
return [{ json: { stats: j.stats, fileNames: j.fileNames, report: j.report || [], results: results } }];`;

// ---------- 4. 应用节点修改 ----------
const nodeByName = {};
for (const n of wf.nodes) nodeByName[n.name] = n;
nodeByName['Odoo 查询'].parameters.jsCode = odooQuery;
nodeByName['计算变更集'].parameters.jsCode = diffCalc;
nodeByName['Odoo 写回'].parameters.jsCode = odooWrite;

// ---------- 5. 新增 Filter 节点 ----------
const filterNode = {
  parameters: {
    conditions: {
      options: { caseSensitive: false, leftValue: '', typeValidation: 'loose' },
      combinator: 'and',
      conditions: [
        {
          id: crypto.randomUUID(),
          leftValue: '={{ $json.subject }}',
          rightValue: 'Abw Castlers Inventory Report',
          operator: { type: 'string', operation: 'contains' }
        }
      ]
    },
    options: {}
  },
  id: crypto.randomUUID(),
  name: '主题过滤',
  type: 'n8n-nodes-base.filter',
  typeVersion: 2.2,
  position: [-200, 320]
};
wf.nodes.push(filterNode);

// ---------- 6. 更新连接：Trigger -> 主题过滤 -> 拆出附件 ----------
const triggerConn = wf.connections['Email Trigger (IMAP)'];
if (triggerConn && triggerConn.main && triggerConn.main[0] && triggerConn.main[0][0]) {
  triggerConn.main[0][0].node = '主题过滤';
}
wf.connections['主题过滤'] = {
  main: [[{ node: '拆出附件', type: 'main', index: 0 }]]
};

// ---------- 7. 输出 ----------
fs.writeFileSync(OUT, JSON.stringify(wf, null, 2), 'utf8');
console.log('nodes:', wf.nodes.map(function (n) { return n.name; }).join(' | '));
console.log('OK ->', OUT);

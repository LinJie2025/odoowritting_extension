/**
 * Odoo Excel Importer v3.0
 *
 * 可拖动按钮 → 悬浮卡片(上传/标签/日志) → Modal(可最小化)
 */
(function () {
  "use strict"

  // ═══════════════════════════════════════════
  //  常量
  // ═══════════════════════════════════════════
  var PREFIX = "__odoi_"
  var LS_POS = PREFIX + "btn_pos"
  var LS_LOG = PREFIX + "log"

  // Odoo purchase.order.line 字段名映射（用户 2026-08-14 确认，如与实际字段不符，改这里即可）
  var ODOO_FIELDS = {
    qty: "product_qty",                          // 订单行/数量（当前未接入写回管道）
    boxQty: "product_packaging_qty",              // 订单行/包装数量（用户确认）
    boxWholesalePrice: "box_wholesale_price",    // 订单行/整箱批发价（用户确认）
    unitPrice: "price_unit",                     // 订单行/单价（用户确认）
    remark: "remark",                            // 订单行/备注（用户确认）
    poRef: "name"                                // 采购订单「订单关联」列对应的字段（用户确认是 name）
  }

  // Excel 表头列名（按名动态定位，不依赖列顺序）
  // 用户 2026-08-17 确认真实表头名
  var EXCEL_COLS = {
    upc: "订单行/产品/内部参考号",   // 匹配键1（UPC）
    catalog: "SKU_x",                // 匹配键2（Catalog，真实表头名「SKU_x」）✅ 2026-08-17 实测确认
    shop: "订单行/店铺",
    qty: "订单行/数量",              // 总件数（无 HS CODE 时算箱数/箱规价用）
    boxQty: "包装数量",               // 包装数量（v3.0 用户确认：数量比对与写回改用此列，取代 abw交货箱数）→ product_packaging_qty
    pack: "订单行/包装",             // 取 "1 box of XX pieces" 的 XX（套装单件数量降级用 + v1.9 Odoo 包装匹配键）
    boxPrice: "箱规价",              // 箱规价（v3.0 起不再作比对基准，保留列定义兼容解析）
    boxPrice09: "0.9箱规价",         // 0.9箱规价 → box_wholesale_price
    total09: "0.9总价",              // 0.9总价（已取消比对，保留列定义）
    unitPrice: "单价",               // 单价 → price_unit
    remark: "备注",
    orderRef: "订单关联",
    partnerRef: "参考号"             // 参考号（2026-08-19 新增：查 PO 优先用它，查不到再用订单关联；两者都匹配 Odoo name 字段）
  }

  // 商品库更新（v1.11.0 新增）：商品库 Excel 表头列名（表头第 1 行，按名定位）
  var PRODUCT_EXCEL_COLS = {
    upc: "UPC",           // 定位键（匹配 product.default_code）
    brand: "品牌",         // → product.brand（Char，用户 2026-08-21 确认）
    shortName: "中文简称"  // → product.name（Char，用户 2026-08-21 确认）
  }

  // 商品库更新：Odoo product.product（产品变体）字段映射（用户 2026-08-21 确认）
  var PRODUCT_ODOO_FIELDS = {
    upc: "default_code",  // 定位键
    brand: "brand",       // 品牌（Char，直接写字符串）
    shortName: "name"     // 商品名称（中文）（Char）
  }

  // ═══════════════════════════════════════════
  //  工具函数
  // ═══════════════════════════════════════════
  function el(tag, attrs, children) {
    var e = document.createElement(tag)
    if (attrs) {
      for (var k in attrs) {
        if (k === "style" && typeof attrs[k] === "string") {
          e.setAttribute("style", attrs[k])
        } else if (k.startsWith("on")) {
          e.addEventListener(k.slice(2).toLowerCase(), attrs[k])
        } else if (k !== "disabled") {
          e.setAttribute(k, attrs[k])
        }
      }
      if (attrs.disabled !== undefined) e.disabled = !!attrs.disabled
    }
    if (children) {
      if (typeof children === "string") {
        e.textContent = children
      } else if (Array.isArray(children)) {
        for (var i = 0; i < children.length; i++) {
          if (typeof children[i] === "string") e.appendChild(document.createTextNode(children[i]))
          else if (children[i] instanceof Node) e.appendChild(children[i])
        }
      } else if (children instanceof Node) {
        e.appendChild(children)
      }
    }
    return e
  }

  function escHtml(s) {
    var d = document.createElement("div")
    d.textContent = s || ""
    return d.innerHTML
  }

  function sleep(ms) {
    return new Promise(function (r) { setTimeout(r, ms) })
  }

  function parseFloatNum(val) {
    if (val === null || val === undefined || val === "") return null
    var n = Number(val)
    return isNaN(n) ? null : n
  }

  function formatPrice(val) {
    if (val === null || val === undefined || val === 0) return "0"
    return String(val)
  }

  // 价格 ×0.9（保留 3 位小数，匹配 Excel 公式精度，如 4946.91×0.9=4452.219）
  function calcBoxPrice(subtotal) {
    var n = parseFloat(subtotal)
    if (isNaN(n)) return null
    return Math.round(n * 0.9 * 1000) / 1000
  }

  // 从 "1 box of 20 pieces" 解析出 20
  function parsePieces(pack) {
    var m = String(pack || "").match(/of\s+(\d+)\s*pieces?/i)
    return m ? parseInt(m[1], 10) : null
  }

  // 从「订单行/包装」列提取包装件数（v1.9 用于 Odoo 包装匹配 product_packaging_id 的 qty）
  // "1 box of 20 pieces" → "20"；"20" → "20"；"1×20"/"20PCS" → "20"；空或无数字 → ""
  function extractPackQty(pack) {
    var s = String(pack || "").trim()
    if (!s) return ""
    var m = s.match(/of\s+(\d+)\s*pieces?/i)
    if (m) return m[1]
    var nums = s.match(/\d+/g)
    return nums ? nums[nums.length - 1] : ""
  }

  // 从 PDF/转换版 Excel 的 description 提取包装件数（v2.0 拆分匹配用）：
  // "Bean Cleansing Oil Jumbo (x42) (Bulk Box) [195ml x 42]" → "42"（只取 (xN) 括号格式，忽略 [195ml x 42]）；提取不到 → ""
  function extractDescPack(desc) {
    var m = String(desc || "").match(/\(\s*x\s*(\d+)\s*\)/i)
    return m ? m[1] : ""
  }

  // 套装单件数量：优先从 PDF PRODUCT DESCRIPTION 提取 "x30"（x 后数字），提取不到降级用 Excel 包装列 pieces
  function parseSetPieces(desc, pack) {
    var m = String(desc || "").match(/x\s*(\d+)/i)
    if (m) return parseInt(m[1], 10)
    return parsePieces(pack)
  }

  // 不一致原因字典（modal 悬浮展示，2026-08-18 用户需求）
  // v3.0（2026-08-26）：数量列改「包装数量」；价格原值比对基准改 Odoo 订单行（Excel 列不再作基准）
  var REASONS = {
    qtySum: function (sum, target) {
      return "数量不一致：包装数量合计 " + fmtNum(sum) + " ≠ 供应商 Qty " + fmtNum(target)
    },
    unitPrice: function (computed, expr, old) {
      return "单价不一致：计算 " + expr + " = " + fmtNum(computed) + " ≠ Odoo " + fmtNum(old)
    },
    boxPrice09: function (computed, expr, old) {
      return "0.9箱规价不一致：计算 " + expr + " = " + fmtNum(computed) + " ≠ Odoo " + fmtNum(old)
    },
    total09: function (computed, expr, old) {
      return "0.9总价不一致：计算 " + expr + " = " + fmtNum(computed) + " ≠ Odoo " + fmtNum(old)
    },
    boxQtyEmpty: "包装数量为空或 0，标记缺货",
    outstock: "Excel 有该 UPC/Catalog，但 PDF 中不存在",
    outstockZero: "缺货：单价与整箱批发价置为 0"
  }

  // 计算表达式展示（factor=1 时省略 ×1）
  function factorExpr(val, factor) {
    return factor === 1 ? String(val) : val + "×" + factor
  }

  function loadPos() {
    try {
      var raw = localStorage.getItem(LS_POS)
      if (raw) return JSON.parse(raw)
    } catch (e) {}
    return { x: window.innerWidth - 80, y: window.innerHeight - 100 }
  }

  function savePos(x, y) {
    try { localStorage.setItem(LS_POS, JSON.stringify({ x: x, y: y })) } catch (e) {}
  }

  function isPurchaseOrderPage() {
    return window.location.hash.includes("model=purchase.order")
  }

  // 商品库更新页面（v1.11.0）：产品变体列表/表单（产品 → 产品变体）
  function isProductPage() {
    return window.location.hash.includes("model=product.product")
  }

  // 扩展可注入页面（采购订单 或 产品变体）
  function isAnyPage() {
    return isPurchaseOrderPage() || isProductPage()
  }

  // ═══════════════════════════════════════════
  //  Odoo API
  // ═══════════════════════════════════════════
  async function rpcCall(endpoint, params) {
    var base = window.location.origin
    var resp = await fetch(base + endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jsonrpc: "2.0", method: "call", params: params, id: Date.now() })
    })
    var data = await resp.json()
    if (data.error) throw new Error(data.error.data?.message || data.error.message)
    return data.result
  }

  async function searchPoByName(name) {
    var result = await rpcCall("/web/dataset/call_kw/purchase.order/search_read", {
      model: "purchase.order", method: "search_read",
      args: [], kwargs: { domain: [["name", "=", name]], fields: ["id", "order_line"] }
    })
    if (!result || result.length === 0) return null
    return { id: result[0].id, orderLineIds: result[0].order_line }
  }

  // 按采购订单「订单关联」字段查询（用户确认字段=name）
  async function searchPoByRef(ref) {
    var result = await rpcCall("/web/dataset/call_kw/purchase.order/search_read", {
      model: "purchase.order", method: "search_read",
      args: [], kwargs: { domain: [[ODOO_FIELDS.poRef, "=", ref]], fields: ["id", "name", "order_line"] }
    })
    if (!result || result.length === 0) return null
    return { id: result[0].id, orderLineIds: result[0].order_line }
  }

  // 批量取订单行（v1.9 匹配键改造：UPC = product.default_code，包装 = product.packaging 的 qty）
  // 一个 UPC 可能对应多个包装不同的品，需 UPC + 包装双重匹配（用户 2026-08-20 需求）
  // v3.0：补读 price_subtotal（套装小计，价格比对原值用）
  async function getOrderLines(ids) {
    var result = await rpcCall("/web/dataset/call_kw/purchase.order.line/search_read", {
      model: "purchase.order.line", method: "search_read",
      args: [], kwargs: {
        domain: [["id", "in", ids]],
        fields: ["id", "name", "price_unit", "box_wholesale_price", "price_subtotal", "remark", "product_packaging_qty", "product_id", "product_packaging_id"]
      }
    })
    var productIds = [], packIds = []
    for (var i = 0; i < result.length; i++) {
      if (result[i].product_id && result[i].product_id.length) productIds.push(result[i].product_id[0])
      if (result[i].product_packaging_id && result[i].product_packaging_id.length) packIds.push(result[i].product_packaging_id[0])
    }
    var codeMap = {}, packMap = {}
    if (productIds.length) {
      var prods = await rpcCall("/web/dataset/call_kw/product.product/search_read", {
        model: "product.product", method: "search_read",
        args: [], kwargs: { domain: [["id", "in", productIds]], fields: ["id", "default_code"] }
      })
      for (var p = 0; p < prods.length; p++) codeMap[prods[p].id] = prods[p].default_code || ""
    }
    if (packIds.length) {
      var packs = await rpcCall("/web/dataset/call_kw/product.packaging/search_read", {
        model: "product.packaging", method: "search_read",
        args: [], kwargs: { domain: [["id", "in", packIds]], fields: ["id", "name", "qty"] }
      })
      for (var q = 0; q < packs.length; q++) packMap[packs[q].id] = { qty: packs[q].qty, name: packs[q].name || "" }
    }
    return result.map(function (r) {
      // UPC = product.default_code（用户指定字段）；为空时用 name 中 [数字] 兜底（旧数据兼容）
      var code = (r.product_id && r.product_id.length) ? (codeMap[r.product_id[0]] || "") : ""
      if (!code) {
        var m = r.name.match(/\[(\d+)\]/)
        code = m ? m[1] : ""
      }
      var pk = (r.product_packaging_id && r.product_packaging_id.length) ? packMap[r.product_packaging_id[0]] : null
      return {
        id: r.id, name: r.name, defaultCode: code,
        packQty: (pk && pk.qty !== null && pk.qty !== undefined && pk.qty !== "") ? String(pk.qty) : "",
        packName: pk ? pk.name : "",
        price_unit: r.price_unit, box_wholesale_price: r.box_wholesale_price,
        price_subtotal: r.price_subtotal,
        packaging_qty: r.product_packaging_qty,
        remark: r.remark || ""
      }
    })
  }

  async function updateOrderLine(lineId, payload) {
    await rpcCall("/web/dataset/call_kw/purchase.order.line/write", {
      model: "purchase.order.line", method: "write", args: [[lineId], payload], kwargs: {}
    })
  }

  // ═══════════════════════════════════════════
  //  Excel 解析（已移除：Excel 导入处理逻辑待与用户重新对齐后重写）
  // ═══════════════════════════════════════════

  // ═══════════════════════════════════════════
  //  PDF 解析 + 采购订单 Excel 解析 + 匹配修改
  // ═══════════════════════════════════════════

  // 从文本块数组提取表格（列数动态，关键列按表头名识别；跨页累加；识别有无 HS CODE 两种格式）
  function extractPdfTable(items) {
    // 1. 表头行：含 UPC/EAN 关键字的文本块，最小 y（跨页时每页表头 x 坐标一致，取最小 y 即可）
    var headerCandidates = items.filter(function (i) { return /UPC|EAN/i.test(i.text) })
    if (!headerCandidates.length) {
      throw new Error("未识别到表格表头（UPC/EAN）——请确认 PDF 为文本型而非扫描件")
    }
    var headerY = Math.min.apply(null, headerCandidates.map(function (i) { return i.y }))

    // 2. 表头行所有文本块（y 容差 4），按 x0 排序
    var headerRow = items.filter(function (i) { return Math.abs(i.y - headerY) <= 4 })
      .sort(function (a, b) { return a.x0 - b.x0 })
    if (!headerRow.length) {
      throw new Error("表头行文本块提取失败")
    }

    // 3. 格式识别：表头是否含 HS CODE
    var hasHsCode = false
    for (var hh = 0; hh < headerRow.length; hh++) {
      if (/HS\s*CODE/i.test(headerRow[hh].text)) { hasHsCode = true; break }
    }

    // 4. 列区间边界 = 相邻标题左边界中点
    var bounds = []
    for (var i = 0; i < headerRow.length - 1; i++) {
      bounds.push((headerRow[i].x0 + headerRow[i + 1].x0) / 2)
    }

    // 5. 关键列索引（按表头名正则识别，两格式通用）
    function findCol(re) {
      for (var i = 0; i < headerRow.length; i++) if (re.test(headerRow[i].text)) return i
      return -1
    }
    var upcCol = findCol(/UPC|EAN/i)
    var catCol = findCol(/Catalog/i)
    var qtyCol = findCol(/Qty/i)
    var priceCol = findCol(/Unit\s*Price/i)
    var subCol = findCol(/Subtotal/i)
    var descCol = findCol(/Product\s*Description/i)   // 套装单件数量来源（"x30"）
    if (upcCol < 0) {
      throw new Error("表头中未找到 UPC/EAN 列")
    }

    // 6. x0 归属列
    function colIndex(x0) {
      for (var i = 0; i < headerRow.length; i++) {
        var lo = i > 0 ? bounds[i - 1] : -Infinity
        var hi = i < bounds.length ? bounds[i] : Infinity
        if (x0 >= lo && x0 < hi) return i
      }
      return headerRow.length - 1
    }

    // 7. 数据行锚点：UPC 列内的 8~14 位纯数字块（跨页累加所有页）
    var digitRe = /^\d{8,14}$/
    var anchors = items.filter(function (i) {
      return digitRe.test(i.text) && i.y > headerY + 10 && colIndex(i.x0) === upcCol
    }).sort(function (a, b) { return a.y - b.y })
    if (!anchors.length) {
      throw new Error("表头已识别，但未找到 UPC 数据行（UPC 需为 8~14 位纯数字）")
    }

    // 8. 行聚类（B：行高自适应，2026-08-18）
    //    行高 = 相邻锚点 y 间距（跨页/异常间距 >120 时沿用上一行高）；
    //    归属范围 = [a.y - max(8, h*0.4), a.y + h*0.8] —— 下行覆盖 desc 折行，上行防串入上一行 desc 第二行
    var numishRe = /^[\d\s.,+-]+$/
    var heights = []
    for (var ha = 0; ha < anchors.length; ha++) {
      var hh = ha + 1 < anchors.length ? anchors[ha + 1].y - anchors[ha].y : -1
      if (hh < 8 || hh > 120) hh = heights.length ? heights[heights.length - 1] : 30
      heights.push(hh)
    }
    var dataRows = []   // 通过假行过滤的 { cells }（粗分：标题中点边界）
    for (var ai = 0; ai < anchors.length; ai++) {
      var a = anchors[ai]
      var h = heights[ai]
      var rowItems = items.filter(function (i) {
        return i.y > headerY && i.y >= a.y - Math.max(8, h * 0.4) && i.y <= a.y + h * 0.8
      })
      var cells = {}
      for (var ri = 0; ri < rowItems.length; ri++) {
        var it = rowItems[ri]
        var ci = colIndex(it.x0)
        if (!cells[ci]) cells[ci] = []
        cells[ci].push(it)
      }
      var row = {}
      for (var ci2 in cells) {
        cells[ci2].sort(function (x, y) { return x.y - y.y })
        row[ci2] = cells[ci2].map(function (x) { return x.text }).join(" ")
      }
      // 假行过滤：页脚/客户信息文本落入 UPC 列（如 'Phone: 18922477200'）时行内无数据列字段，跳过
      if (!(row[catCol] || row[qtyCol] || row[priceCol] || row[subCol])) continue
      dataRows.push({ cells: cells })
    }

    // 9. 列边界校准（A：内容聚类，2026-08-18）
    //    表头位置随对齐方式漂移（居中/右对齐），内容位置才是真实列边界。
    //    列边界 = 相邻两列「内容范围」的间隙中点；空列/内容重叠时用标题中点兜底。
    var colBlocks = []
    for (var cb = 0; cb < headerRow.length; cb++) colBlocks.push([])
    for (var dr = 0; dr < dataRows.length; dr++) {
      var drc = dataRows[dr].cells
      for (var dc in drc) {
        var dcNum = parseInt(dc, 10)
        for (var dbi = 0; dbi < drc[dc].length; dbi++) colBlocks[dcNum].push(drc[dc][dbi])
      }
    }
    // desc 列粗分为空（标题与内容错位时 desc 被粗分进相邻列）→ 从相邻列摘除「含字母」块补入 desc 列
    if (descCol >= 0 && colBlocks[descCol].length === 0) {
      for (var si = -1; si <= 1; si += 2) {
        var sCol = descCol + si
        if (sCol < 0 || sCol >= colBlocks.length) continue
        var kept = []
        for (var sb = 0; sb < colBlocks[sCol].length; sb++) {
          var sblk = colBlocks[sCol][sb]
          if (/[A-Za-z]/.test(sblk.text) && !numishRe.test(sblk.text)) colBlocks[descCol].push(sblk)
          else kept.push(sblk)
        }
        colBlocks[sCol] = kept
      }
    }
    var colRange = []
    for (var cr = 0; cr < colBlocks.length; cr++) {
      var minX0 = Infinity, maxX1 = -Infinity
      for (var cbi = 0; cbi < colBlocks[cr].length; cbi++) {
        var blk = colBlocks[cr][cbi]
        if (blk.x0 < minX0) minX0 = blk.x0
        if (blk.x1 > maxX1) maxX1 = blk.x1
      }
      colRange.push({ minX0: minX0, maxX1: maxX1, empty: colBlocks[cr].length === 0 })
    }
    var newBounds = []
    for (var nb = 0; nb < bounds.length; nb++) {
      var lo = colRange[nb].empty ? bounds[nb] : colRange[nb].maxX1
      var hi = colRange[nb + 1].empty ? bounds[nb] : colRange[nb + 1].minX0
      newBounds.push(lo < hi ? (lo + hi) / 2 : bounds[nb])
    }
    function colIndex2(x0) {
      for (var i = 0; i < headerRow.length; i++) {
        var lo = i > 0 ? newBounds[i - 1] : -Infinity
        var hi = i < newBounds.length ? newBounds[i] : Infinity
        if (x0 >= lo && x0 < hi) return i
      }
      return headerRow.length - 1
    }

    // 10. 按新列边界重新归属
    var rows = []
    for (var rr = 0; rr < dataRows.length; rr++) {
      var src = dataRows[rr].cells
      var cells2 = {}
      for (var sc in src) {
        for (var sbi = 0; sbi < src[sc].length; sbi++) {
          var blk2 = src[sc][sbi]
          var ci2b = colIndex2(blk2.x0)
          if (!cells2[ci2b]) cells2[ci2b] = []
          cells2[ci2b].push(blk2)
        }
      }
      var row2 = {}
      for (var ci4 in cells2) {
        cells2[ci4].sort(function (x, y) { return x.y - y.y })
        row2[ci4] = cells2[ci4].map(function (x) { return x.text }).join(" ")
      }
      // desc 兜底（保险）：重新归属后 desc 列仍空时，从相邻列逐块取含字母文本合并
      var descText = (row2[descCol] || "").trim()
      if (!descText && descCol >= 0) {
        var parts = []
        for (var ci5 in row2) {
          var cnum = parseInt(ci5, 10)
          if (cnum === descCol || Math.abs(cnum - descCol) > 1) continue
          var blocks = cells2[ci5] || []
          for (var bi2 = 0; bi2 < blocks.length; bi2++) {
            var bt = blocks[bi2].text
            if (/[A-Za-z]/.test(bt) && !numishRe.test(bt)) parts.push(bt)
          }
        }
        descText = parts.join(" ")
      }
      row2.__desc = descText
      rows.push(row2)
    }

    // 11. 输出：格式 + Coupon + 关键列
    return {
      format: hasHsCode ? "B" : "A",
      coupon: extractCoupon(items),
      rows: rows.map(function (r) {
        return {
          upc: (r[upcCol] || "").trim(),
          catalog: (r[catCol] || "").trim(),
          qty: (r[qtyCol] || "").trim(),
          unitPrice: (r[priceCol] || "").trim(),
          subtotal: (r[subCol] || "").trim(),
          description: r.__desc || ""
        }
      })
    }
  }

  // 从汇总区提取 Coupon 值（决定是否 ×0.9）
  function extractCoupon(items) {
    var blocks = items.filter(function (i) { return /^Coupon$/i.test(i.text) })
    if (!blocks.length) return 0
    var c = blocks[0]
    var cands = items.filter(function (i) {
      return Math.abs(i.y - c.y) <= 5 && i.x0 > c.x0 && /^-?\d/.test(i.text)
    }).sort(function (a, b) { return a.x0 - b.x0 })
    if (!cands.length) return 0
    var v = parseFloat(cands[0].text.replace(/[^\d.-]/g, ""))
    return isNaN(v) ? 0 : v
  }

  // 用 pdfjs-dist 解析 PDF → 文本块 → 表格
  async function parsePdf(arrayBuffer) {
    if (typeof pdfjsLib === "undefined") throw new Error("PDF 解析库未加载")
    try {
      if (typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.getURL) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL("lib/pdf.worker.min.js")
      }
    } catch (e) {}
    var task = pdfjsLib.getDocument({ data: arrayBuffer, isEvalSupported: false })
    var pdf = await task.promise
    var items = []
    try {
      for (var p = 1; p <= pdf.numPages; p++) {
        var page = await pdf.getPage(p)
        var viewport = page.getViewport({ scale: 1 })
        var content = await page.getTextContent()
        for (var i = 0; i < content.items.length; i++) {
          var it = content.items[i]
          var str = (it.str || "").trim()
          if (!str) continue
          var vp = viewport.convertToViewportPoint(it.transform[4], it.transform[5])
          // x1 = 文本右边缘（内容聚类列边界用；pdfjs width 为文本空间宽度，加 transform[4] 即右边缘 x）
          var vp1 = viewport.convertToViewportPoint(it.transform[4] + (it.width || 0), it.transform[5])
          // 页偏移：多页时各页 y 从 0 起，加 (page-1)*1000 使全局 y 单调递增，避免跨页坐标重叠
          items.push({ text: str, x0: vp[0], x1: vp1[0], y: vp[1] + (p - 1) * 1000, page: p })
        }
      }
    } finally {
      pdf.destroy()
    }
    return extractPdfTable(items)
  }

  // 按列名解析采购订单 Excel（列顺序动态）
  function parseOrderExcel(data) {
    var wb = XLSX.read(data, { type: "array" })
    var ws = wb.Sheets[wb.SheetNames[0]]
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true })
    if (rows.length < 2) throw new Error("Excel 为空或格式不正确")

    var header = rows[0].map(function (h) { return String(h || "").trim() })
    function colIdx(name) { return header.indexOf(name) }
    // 按关键词包含匹配定位（兜底，列名可能带前缀/变体）
    function colIdxFuzzy(keywords) {
      for (var j = 0; j < header.length; j++) {
        var h = header[j]
        if (!h) continue
        for (var k = 0; k < keywords.length; k++) {
          if (h.indexOf(keywords[k]) !== -1) return j
        }
      }
      return -1
    }
    function cell(row, i) {
      if (i < 0) return ""
      var v = row[i]
      return (v === undefined || v === null) ? "" : v
    }
    var idx = {
      upc: colIdx(EXCEL_COLS.upc),
      catalog: colIdx(EXCEL_COLS.catalog),
      shop: colIdx(EXCEL_COLS.shop),
      qty: colIdx(EXCEL_COLS.qty),
      boxQty: colIdx(EXCEL_COLS.boxQty),
      pack: colIdx(EXCEL_COLS.pack),
      boxPrice: colIdx(EXCEL_COLS.boxPrice),
      boxPrice09: colIdx(EXCEL_COLS.boxPrice09),
      total09: colIdx(EXCEL_COLS.total09),
      unitPrice: colIdx(EXCEL_COLS.unitPrice),
      remark: colIdx(EXCEL_COLS.remark),
      orderRef: colIdx(EXCEL_COLS.orderRef),
      partnerRef: colIdx(EXCEL_COLS.partnerRef)
    }
    if (idx.upc < 0) idx.upc = colIdxFuzzy(["内部参考号", "UPC", "EAN"])
    // v3.0：包装数量列名变体兜底（用户确认表头「包装数量」；防带前缀/英文变体）
    if (idx.boxQty < 0) idx.boxQty = colIdxFuzzy(["包装数量", "Packing Qty", "Packing", "包装数"])
    // v1.12.2：参考号模糊匹配必须排除「内部参考号」（UPC 列）——"内部参考号"含子串"参考号"，无「参考号」列时会误命中 UPC 列
    // 无「参考号」列 → partnerRef 为空 → loadOrderLineMap 直接用「订单关联」查 PO（用户 2026-08-25 确认）
    if (idx.partnerRef < 0) {
      for (var prj = 0; prj < header.length; prj++) {
        var prh = header[prj]
        if (!prh || prh.indexOf("内部参考号") !== -1) continue
        if (prh.indexOf("参考号") !== -1 || /reference/i.test(prh)) { idx.partnerRef = prj; break }
      }
    }
    // 「包装/SKU」优先于「Box SKU」——两者都含 "SKU"，若按含 SKU 兜底会先命中 Box SKU 列（2026-08-18 修复）
    if (idx.catalog < 0) idx.catalog = colIdxFuzzy(["包装/SKU", "SKU_x", "SKU", "Catalog", "货号"])
    if (idx.upc < 0) throw new Error("未找到 UPC 列（「" + EXCEL_COLS.upc + "」）")

    var result = []
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i]
      if (!row || !row.length) continue
      var upc = String(cell(row, idx.upc)).trim()
      if (!upc) continue
      result.push({
        rowIndex: i,
        upc: upc,
        catalog: String(cell(row, idx.catalog)).trim(),
        shop: String(cell(row, idx.shop)).trim(),
        qty: parseFloatNum(cell(row, idx.qty)),
        boxQty: parseFloatNum(cell(row, idx.boxQty)),
        pack: String(cell(row, idx.pack)).trim(),
        boxPrice: parseFloatNum(cell(row, idx.boxPrice)),
        boxPrice09: parseFloatNum(cell(row, idx.boxPrice09)),
        total09: parseFloatNum(cell(row, idx.total09)),
        unitPrice: parseFloatNum(cell(row, idx.unitPrice)),
        remark: String(cell(row, idx.remark)).trim(),
        orderRef: String(cell(row, idx.orderRef)).trim(),
        partnerRef: String(cell(row, idx.partnerRef)).trim()
      })
    }
    return result
  }

  // 解析「PDF 转换版 Excel」（2026-08-19 Excel 入口改造：输出与 extractPdfTable 同构，供 applyPdfByMode 复用）
  // 表头行 = 首个含 UPC/EAN 关键字的行；格式按表头有无 HS CODE 分 A/B；列按表头名正则定位（容忍空列/用户附加列）；
  // 数据行 = UPC 列为 8~14 位纯数字的行（汇总区/说明文字自然跳过）；Coupon 从汇总区单独提取（决定 ×0.9）
  function parseConvertedPdfExcel(data) {
    var wb = XLSX.read(data, { type: "array" })
    var ws = wb.Sheets[wb.SheetNames[0]]
    var grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" })
    function s(v) { return (v === undefined || v === null) ? "" : String(v).trim() }

    // 1. 定位表头行
    var headerIdx = -1
    for (var i = 0; i < grid.length; i++) {
      var row = grid[i] || []
      for (var j = 0; j < row.length; j++) {
        if (/UPC|EAN/i.test(s(row[j]))) { headerIdx = i; break }
      }
      if (headerIdx >= 0) break
    }
    if (headerIdx < 0) throw new Error("未识别到表格表头（UPC/EAN）——请确认上传的是 PDF 转换版 Excel")

    // 2. 格式识别（有无 HS CODE）+ 关键列定位
    var header = (grid[headerIdx] || []).map(function (h) { return s(h) })
    var hasHsCode = false
    for (var h = 0; h < header.length; h++) {
      if (/HS\s*CODE/i.test(header[h])) { hasHsCode = true; break }
    }
    function findCol(re) {
      for (var j = 0; j < header.length; j++) if (re.test(header[j])) return j
      return -1
    }
    var upcCol = findCol(/UPC|EAN/i)
    var catCol = findCol(/Catalog/i)
    var qtyCol = findCol(/Qty/i)
    var priceCol = findCol(/Unit\s*Price/i)
    var subCol = findCol(/Subtotal/i)
    var descCol = findCol(/Product\s*Description/i)

    // 3. 数据行（重复表头行跳过：跨页 PDF 转 Excel 可能每页重复表头）
    var digitRe = /^\d{8,14}$/
    var rows = []
    for (var r = headerIdx + 1; r < grid.length; r++) {
      var g = grid[r] || []
      var upc = s(g[upcCol])
      if (/UPC|EAN/i.test(upc)) continue
      if (!digitRe.test(upc)) continue
      rows.push({
        upc: upc,
        catalog: catCol >= 0 ? s(g[catCol]) : "",
        qty: qtyCol >= 0 ? s(g[qtyCol]) : "",
        unitPrice: priceCol >= 0 ? s(g[priceCol]) : "",
        subtotal: subCol >= 0 ? s(g[subCol]) : "",
        description: descCol >= 0 ? s(g[descCol]) : ""
      })
    }
    if (!rows.length) throw new Error("表头已识别，但未找到数据行（UPC 需为 8~14 位纯数字）")

    // 4. Coupon：含「Coupon」单元格所在行，取其右侧首个数值（无则 0 → 不 ×0.9）
    var coupon = 0
    var couponFound = false
    for (var cr = 0; cr < grid.length && !couponFound; cr++) {
      var grow = grid[cr] || []
      for (var gc = 0; gc < grow.length && !couponFound; gc++) {
        if (!/^Coupon$/i.test(s(grow[gc]))) continue
        for (var nc = gc + 1; nc < grow.length; nc++) {
          var nv = parseFloat(s(grow[nc]).replace(/[^\d.-]/g, ""))
          if (!isNaN(nv)) { coupon = nv; couponFound = true; break }
        }
        couponFound = true   // Coupon 行无数值也视为已处理（coupon 保持 0）
      }
    }

    return { format: hasHsCode ? "B" : "A", coupon: coupon, rows: rows }
  }

  // ── 多转换版 Excel 合并（v1.12.0，2026-08-25 用户需求）──
  // 同 UPC+包装去重（key 与 matchPdfToExcel 拆分键一致：UPC+description 的 "(xN)"，提取不到 = 原 UPC）；
  // 合并行 Qty 相加、Subtotal 相加（价格×箱数才与 0.9总价比对一致），UNIT PRICE/catalog/description 取第一个
  function mergeConvFiles(files) {
    var map = {}, keys = []
    for (var f = 0; f < files.length; f++) {
      var rows = files[f].rows
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i]
        var pn = extractDescPack(r.description)
        var key = pn ? r.upc + "+" + pn : r.upc
        var t = map[key]
        if (t) {
          var q1 = parseInt(t.qty, 10), q2 = parseInt(r.qty, 10)
          if (!isNaN(q1) && !isNaN(q2)) t.qty = String(q1 + q2)
          var s1 = parseFloat(t.subtotal), s2 = parseFloat(r.subtotal)
          if (!isNaN(s1) && !isNaN(s2)) t.subtotal = String(s1 + s2)
        } else {
          map[key] = { upc: r.upc, catalog: r.catalog, qty: r.qty, unitPrice: r.unitPrice, subtotal: r.subtotal, description: r.description }
          keys.push(key)
        }
      }
    }
    return keys.map(function (k) { return map[k] })
  }

  // 多文件 Coupon 合并：任一文件非 0 → 取该值（触发 ×0.9）；全部 0 → 0
  function mergeCoupon(files) {
    for (var i = 0; i < files.length; i++) {
      if (files[i].coupon) return files[i].coupon
    }
    return 0
  }

  // ── 通用匹配（v2.0 重构，2026-08-20）：两级匹配 ──
  // 第一级 SKU：PDF Catalog（格式B「CATALOG NO.」/格式A「Catalog#」，解析器已统一提取为 catalog 字段）↔ Excel「SKU」列，能匹配上的直接配对（v1.7.2 逻辑不变）
  // 第二级 UPC（SKU 匹配不上的行）：先整体统计两侧该 UPC 出现次数，任一侧出现多次（同 UPC 多包装/多行）→ 包装拆分匹配：
  //   PDF 侧 key = UPC + description 的 "(xN)"；Excel 侧 key = UPC + 包装列 "1 box of N pieces" 的 N；
  //   提取不到包装数字 → key 保持原 UPC 不拼接；两侧唯一 → 直接按 UPC 匹配（原兜底）
  // 匹配成功后比对处理与 SKU 匹配完全一致；pair.matchKey 供数量核对聚合（SKU 行=Catalog，UPC 行=UPC/拆分键）
  function matchPdfToExcel(pdfRows, excelRows) {
    var pairs = []
    var matchedIdx = {}

    // ── 第一级：SKU 匹配（Excel 行 SKU 有值才参与；每个 Excel 行最多配对一次）──
    var bySku = {}
    for (var i = 0; i < excelRows.length; i++) {
      var e = excelRows[i]
      if (!e.catalog) continue
      if (!bySku[e.catalog]) bySku[e.catalog] = []
      bySku[e.catalog].push(e)
    }
    for (var j = 0; j < pdfRows.length; j++) {
      var p = pdfRows[j]
      var skuHits = bySku[p.catalog] || []
      for (var a = 0; a < skuHits.length; a++) {
        var ex = skuHits[a]
        if (matchedIdx[ex.rowIndex]) continue
        matchedIdx[ex.rowIndex] = true
        pairs.push({ excel: ex, pdf: p, matchKey: p.catalog })
      }
    }

    // ── 第二级：UPC 匹配（含包装拆分）──
    // 整体统计 UPC 出现次数（含已 SKU 匹配的行）：任一侧 >1 → 该 UPC 走包装拆分
    var upcCntPdf = {}, upcCntExcel = {}
    for (var c = 0; c < pdfRows.length; c++) {
      var u = pdfRows[c].upc
      if (u) upcCntPdf[u] = (upcCntPdf[u] || 0) + 1
    }
    for (var d = 0; d < excelRows.length; d++) {
      var ue = excelRows[d].upc
      if (ue) upcCntExcel[ue] = (upcCntExcel[ue] || 0) + 1
    }
    function needSplit(upc) {
      return ((upcCntPdf[upc] || 0) > 1 || (upcCntExcel[upc] || 0) > 1)
    }
    // Excel 侧：剩余未匹配行，按 UPC 或「UPC+包装」建索引
    var byUpcKey = {}
    for (var m = 0; m < excelRows.length; m++) {
      var er = excelRows[m]
      if (matchedIdx[er.rowIndex]) continue
      var ek = er.upc
      if (needSplit(ek)) {
        var pkN = extractPackQty(er.pack)
        if (pkN) ek = ek + "+" + pkN
      }
      if (!byUpcKey[ek]) byUpcKey[ek] = []
      byUpcKey[ek].push(er)
    }
    // PDF 侧：所有未配对完的 Excel 行均可继续被同 key 的 PDF 行配对（一个 PDF 行可对应多个订单关联行，v1.7.5 合并核对场景）
    for (var n = 0; n < pdfRows.length; n++) {
      var pr = pdfRows[n]
      var pk = pr.upc
      if (needSplit(pk)) {
        var pdN = extractDescPack(pr.description)
        if (pdN) pk = pk + "+" + pdN
      }
      var upcHits = byUpcKey[pk] || []
      for (var b = 0; b < upcHits.length; b++) {
        var ex2 = upcHits[b]
        if (matchedIdx[ex2.rowIndex]) continue
        matchedIdx[ex2.rowIndex] = true
        pairs.push({ excel: ex2, pdf: pr, matchKey: pk })
      }
    }

    // 未匹配的 Excel 行 → 缺货（v2.0：拆分后仍匹配不上的同样归入缺货）
    var outstock = []
    for (var o = 0; o < excelRows.length; o++) {
      if (!matchedIdx[excelRows[o].rowIndex]) outstock.push(excelRows[o])
    }
    return { pairs: pairs, outstock: outstock }
  }

  // 数值相等判断（容忍浮点误差，null 视为相等）
  function numEq(a, b) {
    if (a === null || a === undefined || a === "") a = null
    if (b === null || b === undefined || b === "") b = null
    if (a === null && b === null) return true
    if (a === null || b === null) return false
    return Math.abs(parseFloat(a) - parseFloat(b)) < 0.001
  }

  // ×0.9 因子：factor=1 原值；factor=0.9 乘 0.9（整数分）
  function applyFactor(val, factor) {
    var n = parseFloatNum(val)
    if (n === null) return null
    if (factor === 1) return n
    return calcBoxPrice(n)
  }

  // 数量核对：按 matchKey 分组求和包装数量（v3.0 起由 abw交货箱数列改为包装数量列），标记是否一致
  // v2.0：聚合键改用 pair.matchKey——SKU 匹配行 = Catalog；UPC 唯一匹配行 = UPC；拆分匹配行 = UPC+包装（不同包装独立核对）
  // v1.7.5：同一 SKU 拆多个订单关联/多个 UPC（如 A 关联 1 件 + B 关联 18 件 = 19 件）时合并求和 vs PDF Qty 总和
  // v1.5（2026-08-18）：目标 = PDF Qty 原值，不区分格式、不再 ÷pieces
  // v3.0（2026-08-26）：求和列改为「包装数量」（用户确认数量比对与写回改用此列）
  function markQtyMismatch(pairs) {
    var byKey = {}
    for (var i = 0; i < pairs.length; i++) {
      // SKU 有值按 SKU 聚合（同 SKU 跨 UPC/跨订单关联合并）；UPC 匹配按 UPC 聚合；拆分匹配按 UPC+包装聚合
      var key = pairs[i].matchKey || pairs[i].pdf.catalog || pairs[i].pdf.upc || ("row-" + i)
      if (!byKey[key]) byKey[key] = []
      byKey[key].push(pairs[i])
    }
    for (var k in byKey) {
      var list = byKey[k]
      var target = 0, targetOk = true
      var sum = 0
      var seenPdf = []   // 同一转换版行被多个订单关联命中时，PDF Qty 只累计一次
      var detailMap = {} // 悬浮明细：按订单关联聚合（同 matchKey 多条相加时展示「哪几个订单关联各多少」）
      for (var s = 0; s < list.length; s++) {
        var ex = list[s].excel
        var q = ex.boxQty || 0
        sum += q
        if (q > 0) {
          var ref = ex.orderRef || ex.partnerRef || ""
          if (!detailMap[ref]) detailMap[ref] = { ref: ref, sku: list[s].matchKey || list[s].pdf.catalog || list[s].pdf.upc || "", qty: 0 }
          detailMap[ref].qty += q
        }
        var pd = list[s].pdf
        if (seenPdf.indexOf(pd) >= 0) continue
        seenPdf.push(pd)
        var rawT = parseInt(pd.qty, 10)
        if (isNaN(rawT)) { targetOk = false; continue }
        target += rawT
      }
      var detail = []
      for (var dk in detailMap) detail.push(detailMap[dk])
      for (var t = 0; t < list.length; t++) {
        list[t].qtyTarget = targetOk ? target : null
        list[t].qtySum = sum
        list[t].qtyMismatch = targetOk && !numEq(sum, target)
        list[t].qtyDetail = detail
      }
    }
  }

  // ── 单件入口：数量核对（提示）+ 单价（UNIT PRICE ×0.9 vs 单价列）──
  // v1.5：公式不变，字段补充 reason / pdfSource 供 modal 对照与悬浮展示
  // v1.8：Excel 入口（source='excel'）缺货行不写回单价（2026-08-19 用户需求）
  function applyPdfSingle(pdfRows, excelRows, factor, source) {
    var m = matchPdfToExcel(pdfRows, excelRows)
    markQtyMismatch(m.pairs)
    var changes = []
    for (var i = 0; i < m.pairs.length; i++) {
      var pair = m.pairs[i]
      var ex = pair.excel, p = pair.pdf
      var fields = []
      var outstock = (ex.boxQty == null || ex.boxQty === 0)   // 缺货行：包装数量为空/0（v3.0 由包装数量列驱动）
      // PDF 原始字段（悬浮提示第一行，2026-08-18 用户需求）
      var pdfRaw = ["UPC=" + p.upc, "QTY=" + p.qty, "UNIT PRICE=" + p.unitPrice, "Subtotal=" + p.subtotal].join(" | ")
      // 包装数量（数量核对：可编辑，默认原值；数量不一致由 qtyMismatch 提示）
      fields.push({ key: "boxQty", label: "包装数量", odooField: ODOO_FIELDS.boxQty, oldValue: ex.boxQty, newValue: ex.boxQty, changed: false, pdfRaw: pdfRaw })
      // 包装数量为空或 0 → 备注缺货
      if (outstock) {
        fields.push({ key: "remark", label: "备注", odooField: ODOO_FIELDS.remark, oldValue: ex.remark || "", newValue: "缺货", changed: true, reason: REASONS.boxQtyEmpty, pdfRaw: pdfRaw })
      }
      // 单价：PDF UNIT PRICE × factor；原值/changed 由 Odoo price_unit 回填（v3.0，不再比对 Excel 单价列）
      var unit = parseFloatNum(p.unitPrice)
      var newUnit = applyFactor(unit, factor)
      var expr = unit !== null ? factorExpr(unit, factor) : null
      var unitField = {
        key: "unitPrice", label: "单价", odooField: ODOO_FIELDS.unitPrice,
        oldValue: null, newValue: newUnit, changed: false, expr: expr,
        pdfSource: newUnit !== null ? "PDF: " + expr + " = " + newUnit : "",
        reason: null,
        pdfRaw: pdfRaw
      }
      // v1.12.4（2026-08-26 用户需求）：缺货行（两入口统一）单价/整箱批发价置 0 写回，取代 v1.8.0 的 Excel 入口不写回
      if (outstock) {
        unitField.newValue = 0
        unitField.changed = true
        unitField.pdfSource = "缺货置 0"
        unitField.reason = REASONS.outstockZero
        // 单件入口无 0.9箱规价字段，追加整箱批发价置 0（Odoo 行有该字段则写 0）
        fields.push({
          key: "boxWholesale", label: "整箱批发价", odooField: ODOO_FIELDS.boxWholesalePrice,
          oldValue: null, newValue: 0, changed: true,
          pdfSource: "缺货置 0", reason: REASONS.outstockZero, pdfRaw: pdfRaw
        })
      }
      fields.push(unitField)
      changes.push({
        kind: "match", upc: ex.upc, catalog: ex.catalog, shop: ex.shop, orderRef: ex.orderRef, partnerRef: ex.partnerRef,
        packQty: extractPackQty(ex.pack),
        qtyTarget: pair.qtyTarget, qtySum: pair.qtySum, qtyMismatch: pair.qtyMismatch, qtyDetail: pair.qtyDetail,
        fields: fields
      })
    }
    // Excel 有但 PDF 无 → 缺货：备注 + 单价/整箱批发价置 0（v1.12.4 用户需求）
    for (var q = 0; q < m.outstock.length; q++) {
      var er = m.outstock[q]
      changes.push({
        kind: "outstock", upc: er.upc, catalog: er.catalog, shop: er.shop, orderRef: er.orderRef, partnerRef: er.partnerRef,
        packQty: extractPackQty(er.pack),
        qtyTarget: null, qtySum: null, qtyMismatch: false,
        fields: [
          { key: "remark", label: "备注", odooField: ODOO_FIELDS.remark, oldValue: "", newValue: "缺货", changed: true, reason: REASONS.outstock },
          { key: "unitPrice", label: "单价", odooField: ODOO_FIELDS.unitPrice, oldValue: er.unitPrice, newValue: 0, changed: true, pdfSource: "缺货置 0", reason: REASONS.outstockZero },
          { key: "boxWholesale", label: "整箱批发价", odooField: ODOO_FIELDS.boxWholesalePrice, oldValue: null, newValue: 0, changed: true, pdfSource: "缺货置 0", reason: REASONS.outstockZero }
        ]
      })
    }
    return changes
  }

  // ── 套装入口：统一公式（不区分 HS CODE）+ 3 项比对（v3.0 去掉箱规价，价格原值比对 Odoo）──
  // 单价=UNIT PRICE×factor÷套装单件数量；0.9箱规价=UNIT PRICE×factor；0.9总价=Subtotal×factor
  // 套装单件数量：PDF PRODUCT DESCRIPTION 的 "x30" → 降级 Excel 包装列 pieces
  // v3.0（2026-08-26）：箱规价字段移除；价格原值（oldValue/changed）由 Odoo 订单行回填（price_unit/box_wholesale_price/price_subtotal）
  // v1.12.4：缺货行（两入口统一）单价/0.9箱规价置 0 写回
  function applyPdfSet(pdfRows, excelRows, factor, source) {
    var m = matchPdfToExcel(pdfRows, excelRows)
    markQtyMismatch(m.pairs)
    var changes = []
    for (var i = 0; i < m.pairs.length; i++) {
      var pair = m.pairs[i]
      var ex = pair.excel, p = pair.pdf
      var fields = []
      var outstock = (ex.boxQty == null || ex.boxQty === 0)   // 缺货行：包装数量为空/0（v3.0 由包装数量列驱动）
      // 套装单件数量：PDF PRODUCT DESCRIPTION "x30" → 降级 Excel 包装列 pieces
      var pieces = parseSetPieces(p.description, ex.pack)
      var unit = parseFloatNum(p.unitPrice)
      var sub = parseFloatNum(p.subtotal)
      // PDF 原始字段（悬浮提示第一行，2026-08-18 用户需求）
      var pdfRaw = ["UPC=" + p.upc, "QTY=" + p.qty, "UNIT PRICE=" + p.unitPrice, "Subtotal=" + p.subtotal].join(" | ")
      if (pieces && pieces > 0) pdfRaw += " | 套装单件数量=" + pieces

      // 包装数量（数量核对：可编辑，默认原值；数量不一致由 qtyMismatch 提示）
      fields.push({ key: "boxQty", label: "包装数量", odooField: ODOO_FIELDS.boxQty, oldValue: ex.boxQty, newValue: ex.boxQty, changed: false, pdfRaw: pdfRaw })
      // 包装数量为空或 0 → 备注缺货
      if (outstock) {
        fields.push({ key: "remark", label: "备注", odooField: ODOO_FIELDS.remark, oldValue: ex.remark || "", newValue: "缺货", changed: true, reason: REASONS.boxQtyEmpty, pdfRaw: pdfRaw })
      }
      // v3.0：箱规价字段已移除（用户确认：Odoo 无对应字段，不再展示）；价格原值改由 Odoo 回填

      // 单价 = UNIT PRICE × factor ÷ 套装单件数量 → price_unit；原值由 Odoo price_unit 回填（v3.0）
      var newUnit = null, unitExpr = null
      if (unit !== null) {
        unitExpr = factorExpr(unit, factor) + (pieces && pieces > 0 ? "÷" + pieces : "")
        newUnit = unit * factor
        if (pieces && pieces > 0) newUnit = Math.round((newUnit / pieces) * 10000) / 10000
        else newUnit = null
      }
      var unitField = {
        key: "unitPrice", label: "单价", odooField: ODOO_FIELDS.unitPrice,
        oldValue: null, newValue: newUnit, changed: false, expr: unitExpr,
        pdfSource: newUnit !== null ? "PDF: " + unitExpr + " = " + newUnit : "",
        reason: null,
        pdfRaw: pdfRaw
      }
      // v1.12.4（2026-08-26 用户需求）：缺货行（两入口统一）单价置 0 写回，取代 v1.8.0 的 Excel 入口不写回
      if (outstock) {
        unitField.newValue = 0
        unitField.changed = true
        unitField.pdfSource = "缺货置 0"
        unitField.reason = REASONS.outstockZero
      }
      fields.push(unitField)

      // 0.9箱规价 = UNIT PRICE × factor → box_wholesale_price；原值由 Odoo box_wholesale_price 回填（v3.0）
      var newBoxPrice09 = applyFactor(unit, factor)
      var box09Field = {
        key: "boxPrice09", label: "0.9箱规价", odooField: ODOO_FIELDS.boxWholesalePrice,
        oldValue: null, newValue: newBoxPrice09, changed: false, expr: factorExpr(unit, factor),
        pdfSource: newBoxPrice09 !== null ? "PDF: " + factorExpr(unit, factor) + " = " + newBoxPrice09 : "",
        reason: null,
        pdfRaw: pdfRaw
      }
      // v1.12.4：缺货行 0.9箱规价置 0 写回（不再不写回）
      if (outstock) {
        box09Field.newValue = 0
        box09Field.changed = true
        box09Field.pdfSource = "缺货置 0"
        box09Field.reason = REASONS.outstockZero
      }
      fields.push(box09Field)

      // 0.9总价 = Subtotal × factor（不写回，仅比对）；原值由 Odoo price_subtotal 回填（v3.0）
      var newTotal09 = applyFactor(sub, factor)
      fields.push({
        key: "total09", label: "0.9总价", odooField: null,
        oldValue: null, newValue: newTotal09, changed: false, expr: factorExpr(sub, factor),
        pdfSource: newTotal09 !== null ? "PDF: " + factorExpr(sub, factor) + " = " + newTotal09 : "",
        reason: null,
        pdfRaw: pdfRaw
      })

      changes.push({
        kind: "match", upc: ex.upc, catalog: ex.catalog, shop: ex.shop, orderRef: ex.orderRef, partnerRef: ex.partnerRef,
        packQty: extractPackQty(ex.pack),
        qtyTarget: pair.qtyTarget, qtySum: pair.qtySum, qtyMismatch: pair.qtyMismatch, qtyDetail: pair.qtyDetail,
        fields: fields
      })
    }
    // Excel 有但 PDF 无 → 缺货：备注 + 单价/整箱批发价置 0（v1.12.4 用户需求）
    for (var q = 0; q < m.outstock.length; q++) {
      var er = m.outstock[q]
      changes.push({
        kind: "outstock", upc: er.upc, catalog: er.catalog, shop: er.shop, orderRef: er.orderRef, partnerRef: er.partnerRef,
        packQty: extractPackQty(er.pack),
        qtyTarget: null, qtySum: null, qtyMismatch: false,
        fields: [
          { key: "remark", label: "备注", odooField: ODOO_FIELDS.remark, oldValue: "", newValue: "缺货", changed: true, reason: REASONS.outstock },
          { key: "unitPrice", label: "单价", odooField: ODOO_FIELDS.unitPrice, oldValue: er.unitPrice, newValue: 0, changed: true, pdfSource: "缺货置 0", reason: REASONS.outstockZero },
          { key: "boxWholesale", label: "整箱批发价", odooField: ODOO_FIELDS.boxWholesalePrice, oldValue: null, newValue: 0, changed: true, pdfSource: "缺货置 0", reason: REASONS.outstockZero }
        ]
      })
    }
    return changes
  }

  // ── 入口分发：按 PDF 修正入口路由到对应处理逻辑 ──
  // source: 'pdf' | 'excel'（v1.8：Excel 流缺货行不写回单价/整箱批发价）
  function applyPdfByMode(pdfRows, excelRows, mode, coupon, source) {
    var factor = (coupon === 0 || coupon === null || coupon === undefined) ? 1 : 0.9
    var changes = (mode === "single") ? applyPdfSingle(pdfRows, excelRows, factor, source) : applyPdfSet(pdfRows, excelRows, factor, source)
    for (var i = 0; i < changes.length; i++) changes[i].mode = mode
    return changes
  }

  // ── Excel 入口比对逻辑（2026-08-19 改造）：与 PDF 流完全一致 ──
  // 转换版 Excel 经 parseConvertedPdfExcel 解析后与 PDF 表格同构，直接走 applyPdfByMode（单件/套装同一套公式）。
  // 旧 applyExcelChanges（Excel 值 vs Odoo 现有值）已移除。

  // ═══════════════════════════════════════════
  //  业务逻辑（Excel 导入流匹配/预览逻辑已移除，待重新对齐后重写）
  // ═══════════════════════════════════════════

  // ═══════════════════════════════════════════
  //  日志
  // ═══════════════════════════════════════════
  function saveLog(logEntry) {
    try { localStorage.setItem(LS_LOG, JSON.stringify(logEntry)) } catch (e) {}
  }

  function getLog() {
    try { var r = localStorage.getItem(LS_LOG); return r ? JSON.parse(r) : null } catch (e) { return null }
  }

  function clearLog() { localStorage.removeItem(LS_LOG) }

  function downloadLog() {
    var entry = getLog()
    if (!entry) return
    var json = JSON.stringify(entry, null, 2)
    var blob = new Blob([json], { type: "application/json" })
    var url = URL.createObjectURL(blob)
    var a = document.createElement("a")
    a.href = url; a.download = "odoo-import-" + entry.timestamp.replace(/[:.]/g, "-") + ".json"
    document.body.appendChild(a); a.click(); document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  // ═══════════════════════════════════════════
  //  App 状态
  // ═══════════════════════════════════════════
  var appState = {
    cardOpen: false,
    activeTab: isProductPage() ? "product" : "excel"   // 'excel' | 'pdf' | 'product'（商品库更新，v1.11.0；产品页默认商品库 Tab）
  }

  // PDF 修正入口模式（单件 / 套装，处理逻辑区分）
  var PDF_MODES = {
    single: { id: "single", label: "单件入口", icon: "📦", desc: "按单件 UPC 匹配修正" },
    set:    { id: "set",    label: "套装入口", icon: "🎁", desc: "按套装组合匹配修正" }
  }

  // PDF 区流程状态
  var pdfState = {
    mode: null,          // 修正入口: null | 'single' | 'set'
    pdfRows: null,       // PDF 解析出的表格数据行 [{upc,catalog,qty,unitPrice,subtotal}]
    pdfFileName: null,
    pdfFormat: null,     // PDF 格式: 'A'(无HS CODE) | 'B'(有HS CODE)
    coupon: 0,           // 汇总区 Coupon 值（决定是否 ×0.9）
    excelRows: null,     // 采购订单 Excel 解析结果
    excelFileName: null,
    changes: null,       // 匹配修改结果（每条带 mode 标记）
    error: null          // 当前步骤错误信息（常驻显示，下次成功时清除）
  }

  // Excel 导入区流程状态（2026-08-19 改造：上传「PDF 转换版 Excel」→ 再传采购订单 Excel → 与 PDF 流同逻辑比对写回）
  var excelState = {
    mode: null,          // 入口: null | 'single' | 'set'
    convFiles: [],       // 已添加的转换版文件 [{name, rows, format, coupon}]（v1.12.0 多文件）
    convRows: null,      // 合并去重后的表格数据行（点「下一步」时由 mergeConvFiles 生成）
    convFileName: null,  // 展示用：全部文件名（", " 连接）
    convFormat: null,    // 展示用：格式（多文件混合时取第一个文件的）
    coupon: 0,           // 合并后 Coupon（任一文件非 0 → 取该值触发 ×0.9；全 0 → 0）
    excelRows: null,     // 采购订单 Excel 解析结果
    excelFileName: null,
    changes: null,       // applyPdfByMode 结果（与 PDF 流同一逻辑）
    error: null
  }

  // 商品库更新区流程状态（v1.11.0 新增：上传商品库 Excel → 按 UPC 查 product.product → 预览 → 写回 name/brand）
  var productState = {
    rows: null,          // 解析出的 Excel 行 [{rowIndex, upc, brand, shortName}]
    fileName: null,
    previewRows: null,   // 预览行（含 Odoo 匹配结果与状态）
    error: null          // 当前步骤错误信息
  }

  // ═══════════════════════════════════════════
  //  颜色常量（Excel 预览行配色已移除，待流程重写后按需恢复；PDF 预览行配色内联）
  // ═══════════════════════════════════════════

  var CARD_STYLE = {
    bg: "#fff", radius: "14px", shadow: "0 8px 40px rgba(0,0,0,0.18)",
    width: "340px", maxHeight: "480px", zIndex: "100001"
  }

  // ═══════════════════════════════════════════
  //  全局拖拽状态 (避免重复绑定 document 事件)
  // ═══════════════════════════════════════════
  var dragCtx = null // { el, startX, startY, startLeft, startTop }

  document.addEventListener("mousemove", function (e) {
    if (!dragCtx) return
    var dx = e.clientX - dragCtx.startX, dy = e.clientY - dragCtx.startY
    if (Math.abs(dx) < 3 && Math.abs(dy) < 3) return
    if (!dragCtx.moved) { dragCtx.moved = true; dragCtx.el.style.transition = "none" }
    var w = dragCtx.el.offsetWidth, h = dragCtx.el.offsetHeight
    var nl = Math.max(0, Math.min(window.innerWidth - w - 10, dragCtx.startLeft + dx))
    var nt = Math.max(0, Math.min(window.innerHeight - h - 10, dragCtx.startTop + dy))
    dragCtx.el.style.left = nl + "px"
    dragCtx.el.style.top = nt + "px"
  })

  document.addEventListener("mouseup", function () {
    if (!dragCtx) return
    dragCtx.el.style.transition = dragCtx.defaultTransition || ""
    dragCtx.el.style.cursor = dragCtx.origCursor || ""
    if (dragCtx.onDragEnd) dragCtx.onDragEnd(dragCtx.moved)
    dragCtx = null
  })

  // ═══════════════════════════════════════════
  //  1. 可拖动按钮
  // ═══════════════════════════════════════════
  function createDraggableButton() {
    var pos = loadPos()
    var btn = el("div", {
      id: PREFIX + "btn",
      style: [
        "position:fixed;z-index:100000",
        "left:" + pos.x + "px;top:" + pos.y + "px",
        "width:52px;height:52px;border-radius:28px",
        "background:#7c3aed;color:#fff",
        "display:flex;align-items:center;justify-content:center",
        "cursor:grab;box-shadow:0 4px 20px rgba(124,58,237,0.4)",
        "transition:transform 0.3s,border-radius 0.3s,width 0.3s,height 0.3s,left 0.3s,top 0.3s",
        "user-select:none;font-size:22px"
      ].join(";")
    }, "📥")

    var dragging = false, startX, startY, startLeft, startTop, moved = false
    var dragOver = false
    var defaultTransition = "transform 0.3s,border-radius 0.3s,width 0.3s,height 0.3s,left 0.3s,top 0.3s"

    btn.addEventListener("mousedown", function (e) {
      if (e.button !== 0) return
      dragCtx = {
        el: btn, startX: e.clientX, startY: e.clientY,
        startLeft: parseInt(btn.style.left), startTop: parseInt(btn.style.top),
        moved: false, origCursor: btn.style.cursor,
        defaultTransition: defaultTransition,
        onDragEnd: function (moved) {
          if (moved) savePos(parseInt(btn.style.left), parseInt(btn.style.top))
          else toggleCard(btn) // 点击没拖动 → 切换卡片
        }
      }
      btn.style.cursor = "grabbing"
      e.preventDefault()
    })

    // 拖文件到按钮上
    btn.addEventListener("dragover", function (e) {
      e.preventDefault(); e.stopPropagation()
      if (!dragOver) {
        dragOver = true
        btn.style.transform = "scale(1.2)"
        btn.style.boxShadow = "0 6px 30px rgba(124,58,237,0.7)"
        btn.style.cursor = "copy"
        btn.textContent = "📂"
      }
    })

    btn.addEventListener("dragleave", function (e) {
      e.preventDefault(); e.stopPropagation()
      dragOver = false
      btn.style.transform = "scale(1)"
      btn.style.boxShadow = "0 4px 20px rgba(124,58,237,0.4)"
      btn.style.cursor = "grab"
      btn.textContent = "📥"
    })

    btn.addEventListener("drop", function (e) {
      e.preventDefault(); e.stopPropagation()
      dragOver = false
      btn.style.transform = "scale(1)"
      btn.style.boxShadow = "0 4px 20px rgba(124,58,237,0.4)"
      btn.style.cursor = "grab"
      btn.textContent = "📥"
      var file = e.dataTransfer.files[0]
      if (file) showToast("请在卡片内选择入口上传文件（📊 Excel 导入 / 📄 PDF 修正 / 🏷 商品库更新）", "info")
    })

    return btn
  }

  // ═══════════════════════════════════════════
  //  2. 悬浮卡片
  // ═══════════════════════════════════════════
  var cardEl = null

  function toggleCard(btn) {
    if (appState.cardOpen) closeCard(btn)
    else openCard(btn)
  }

  function openCard(btn) {
    appState.cardOpen = true
    var rect = btn.getBoundingClientRect()
    var cardW = 340, cardH = 500

    // 确保卡片在视口内
    var left = rect.left
    var top = rect.top
    if (left + cardW > window.innerWidth - 10) left = window.innerWidth - cardW - 10
    if (top + cardH > window.innerHeight - 10) top = window.innerHeight - cardH - 10
    if (left < 10) left = 10
    if (top < 10) top = 10

    // 按钮缩放动画
    btn.style.transform = "scale(0)"
    btn.style.opacity = "0"

    cardEl = el("div", {
      id: PREFIX + "card",
      style: [
        "position:fixed;z-index:" + CARD_STYLE.zIndex,
        "left:" + left + "px;top:" + top + "px",
        "width:" + CARD_STYLE.width + ";max-height:" + CARD_STYLE.maxHeight,
        "background:" + CARD_STYLE.bg + ";border-radius:" + CARD_STYLE.radius,
        "box-shadow:" + CARD_STYLE.shadow,
        "display:flex;flex-direction:column;overflow:hidden",
        "transform:scale(0);transition:transform 0.35s cubic-bezier(0.34,1.56,0.64,1),opacity 0.35s ease,left 0.1s,top 0.1s",
        "cursor:default"
      ].join(";")
    })

    // 保存按钮位置（用于关闭时缩回方向）
    cardEl.__btnRect = rect

    // 卡片从按钮中心展开: transform-origin = 按钮中心相对卡片左上角的偏移
    var originX = rect.left + rect.width / 2 - left
    var originY = rect.top + rect.height / 2 - top

    cardEl.appendChild(renderCardContent())
    document.body.appendChild(cardEl)

    // 设置 transform-origin 后触发动画
    cardEl.style.transformOrigin = originX + "px " + originY + "px"
    requestAnimationFrame(function () { cardEl.style.transform = "scale(1)" })

    // 卡片拖动 (通过 header 拖)
    makeCardDraggable(cardEl)
  }

  function closeCard(btn) {
    appState.cardOpen = false
    if (cardEl) {
      // 重新计算 transform-origin 指向按钮当前位置
      var btnRect = btn.getBoundingClientRect()
      var cardLeft = parseInt(cardEl.style.left)
      var cardTop = parseInt(cardEl.style.top)
      var ox = btnRect.left + btnRect.width / 2 - cardLeft
      var oy = btnRect.top + btnRect.height / 2 - cardTop
      cardEl.style.transformOrigin = ox + "px " + oy + "px"
      cardEl.style.transform = "scale(0)"
      cardEl.style.opacity = "0"
      setTimeout(function () { if (cardEl) { cardEl.remove(); cardEl = null } }, 350)
    }
    // 按钮延迟出现，等卡片缩回动画进行一半
    setTimeout(function () {
      btn.style.transform = "scale(1)"
      btn.style.opacity = "1"
    }, 150)
  }

  function makeCardDraggable(card) {
    var header = card.querySelector("div[style*='border-bottom:1px solid #f3f4f6']")
    if (!header) return

    header.style.cursor = "grab"
    header.addEventListener("mousedown", function (e) {
      if (e.target.tagName === "BUTTON") return
      if (e.button !== 0) return
      dragCtx = {
        el: card, startX: e.clientX, startY: e.clientY,
        startLeft: parseInt(card.style.left), startTop: parseInt(card.style.top),
        moved: false, origCursor: "grab",
        defaultTransition: "transform 0.35s cubic-bezier(0.34,1.56,0.64,1),opacity 0.35s ease,left 0.1s,top 0.1s",
        onDragEnd: function () {}
      }
      header.style.cursor = "grabbing"
      e.preventDefault()
    })
  }

  function refreshCard() {
    if (!cardEl || !appState.cardOpen) return
    cardEl.innerHTML = ""
    cardEl.appendChild(renderCardContent())
    makeCardDraggable(cardEl)
  }

  function renderCardContent() {
    var frag = document.createDocumentFragment()
    var tab = appState.activeTab
    var isExcelTab = tab === "excel"
    var isProductTab = tab === "product"
    // 标题：商品库更新 / Excel 导入 / PDF 修正（含当前入口徽章，商品库无入口概念）
    var titleText = isProductTab ? "商品库更新" : (isExcelTab ? "Excel 导入" : "PDF 修正")
    if (!isProductTab && (isExcelTab ? excelState.mode : pdfState.mode)) {
      var m = PDF_MODES[isExcelTab ? excelState.mode : pdfState.mode]
      titleText += " · " + m.label
    }

    // Header
    frag.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;padding:14px 16px;border-bottom:1px solid #f3f4f6;user-select:none"
    }, [
      el("div", { style: "display:flex;align-items:center;gap:8px" }, [
        el("span", { style: "font-size:11px;color:#6b7280;background:#f3f4f6;padding:2px 6px;border-radius:4px" }, isProductTab ? "商品库" : "采购结果"),
        el("span", { style: "font-size:15px;font-weight:600;color:#111827" }, titleText)
      ]),
      el("button", {
        onclick: function () { closeCard(document.getElementById(PREFIX + "btn")) },
        style: "width:26px;height:26px;border:1px solid #e5e7eb;border-radius:6px;background:#fff;color:#6b7280;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center"
      }, "✕")
    ]))

    // Tab 栏：📊 Excel 导入 / 📄 PDF 修正 / 🏷 商品库更新（v1.11.0）
    frag.appendChild(renderTabSwitch())

    // 内容区（按 Tab 路由）+ 日志区（v1.12.1：包可滚动容器 flex:1，文件列表超出卡片高度时内容区出现滚动条，Header/Tab 固定）
    var bodyWrap = el("div", { style: "flex:1;overflow-y:auto;min-height:0;padding-bottom:10px" })
    if (isExcelTab) bodyWrap.appendChild(renderExcelZone())
    else if (isProductTab) bodyWrap.appendChild(renderProductZone())
    else bodyWrap.appendChild(renderPdfZone())

    // 日志区 (有日志才显示)
    var log = getLog()
    if (log) {
      bodyWrap.appendChild(el("div", { style: "padding:0 16px;margin-top:2px" }, [
        el("div", { style: "font-size:11px;color:#9ca3af;margin-bottom:6px;text-transform:uppercase;letter-spacing:0.5px" }, "最近日志"),
        renderLogSection(log)
      ]))
    }
    frag.appendChild(bodyWrap)

    return frag
  }

  // ═══════════════════════════════════════════
  //  3. 上传区（Excel 导入 Tab 已移除，只保留 PDF 修正入口）
  // ═══════════════════════════════════════════

  // ═══════════════════════════════════════════
  //  3. Tab 栏（📊 Excel 导入 / 📄 PDF 修正，2026-08-18 恢复双 Tab）
  // ═══════════════════════════════════════════
  function renderTabSwitch() {
    var bar = el("div", { style: "display:flex;padding:10px 16px 0;gap:8px" })
    var tabs = [
      { id: "excel", label: "📊 Excel 导入" },
      { id: "pdf", label: "📄 PDF 修正" },
      { id: "product", label: "🏷 商品库更新" }
    ]
    for (var i = 0; i < tabs.length; i++) {
      (function (t) {
        var active = appState.activeTab === t.id
        bar.appendChild(el("div", {
          onclick: function () { appState.activeTab = t.id; refreshCard() },
          style: "flex:1;text-align:center;padding:8px 6px;border-radius:8px;font-size:12px;cursor:pointer;font-weight:600;color:" +
            (active ? "#7c3aed" : "#6b7280") + ";background:" + (active ? "#f5f3ff" : "#f9fafb") +
            ";border:1px solid " + (active ? "#ddd6fe" : "#e5e7eb") + ";transition:background 0.15s,border-color 0.15s"
        }, t.label))
      })(tabs[i])
    }
    return bar
  }

  // ═══════════════════════════════════════════
  //  3b. PDF 区 UI
  // ═══════════════════════════════════════════


  function renderPdfZone() {
    var wrap = el("div", { style: "margin:12px 16px;display:flex;flex-direction:column;gap:10px" })

    // 当前入口徽章（可切换）
    if (pdfState.mode) {
      var m = PDF_MODES[pdfState.mode]
      wrap.appendChild(el("div", {
        style: "display:flex;justify-content:space-between;align-items:center;font-size:12px;color:#7c3aed;background:#f5f3ff;border:1px solid #ddd6fe;border-radius:8px;padding:8px 10px"
      }, [
        el("span", {}, m.icon + " 当前入口：" + m.label),
        el("span", {
          onclick: function () { resetPdfState() },
          style: "cursor:pointer;text-decoration:underline;color:#6b7280"
        }, "切换入口")
      ]))
    }

    // 常驻错误提示（解析失败原因，成功后自动清除）
    if (pdfState.error) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#dc2626;background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:8px 10px;line-height:1.5"
      }, "⚠️ " + pdfState.error))
    }

    // 步骤指示器：① 上传 PDF → ② 上传 Excel → ③ 预览确认（已完成步骤可点击回退，防误上传）
    if (pdfState.mode) {
      var step = pdfState.pdfRows ? (pdfState.excelRows ? 3 : 2) : 1
      var steps = [
        { n: 1, label: "上传 PDF", icon: "📄" },
        { n: 2, label: "上传 Excel", icon: "📊" },
        { n: 3, label: "预览确认", icon: "👁" }
      ]
      var stepBar = el("div", { style: "display:flex;align-items:center;gap:6px;font-size:11px" })
      for (var si = 0; si < steps.length; si++) {
        (function (s) {
          var done = s.n < step, active = s.n === step
          var clickable = done   // 仅已完成步骤可点击回退
          var chipStyle = "flex:1;text-align:center;padding:6px 4px;border-radius:6px;border:1px solid " +
            (active ? "#7c3aed" : (done ? "#a7f3d0" : "#e5e7eb")) + ";background:" +
            (active ? "#f5f3ff" : (done ? "#f0fdf4" : "#f9fafb")) + ";color:" +
            (active ? "#7c3aed" : (done ? "#059669" : "#9ca3af")) + ";font-weight:" + (active ? "600" : "400") +
            (clickable ? ";cursor:pointer;transition:background 0.15s,border-color 0.15s" : ";cursor:default")
          var chipAttrs = { style: chipStyle, title: clickable ? "点击返回此步骤重新上传" : "" }
          if (clickable) {
            chipAttrs.onclick = function () { jumpPdfStep(s.n) }
            chipAttrs.onmouseenter = function () { chip.style.background = "#ecfdf5"; chip.style.borderColor = "#34d399" }
            chipAttrs.onmouseleave = function () { chip.style.background = "#f0fdf4"; chip.style.borderColor = "#a7f3d0" }
          }
          var chip = el("div", chipAttrs, (done ? "✅" : "") + s.icon + " " + s.label)
          stepBar.appendChild(chip)
        })(steps[si])
        if (si < steps.length - 1) {
          stepBar.appendChild(el("span", { style: "color:#d1d5db" }, "→"))
        }
      }
      wrap.appendChild(stepBar)
    }

    // 状态提示
    if (pdfState.pdfRows) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#059669;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:8px 10px"
      }, "✅ PDF 已解析：" + (pdfState.pdfFileName || "") + "（" + pdfState.pdfRows.length + " 行）"))
    }
    if (pdfState.excelRows) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#059669;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:8px 10px"
      }, "✅ Excel 已解析：" + (pdfState.excelFileName || "") + "（" + pdfState.excelRows.length + " 行）"))
    }

    if (!pdfState.mode) {
      // 步骤0：选择修正入口（单件 / 套装）
      wrap.appendChild(renderModePicker(selectPdfMode))
    } else if (!pdfState.pdfRows) {
      // 步骤1：上传 PDF
      wrap.appendChild(makeDropZone("点击上传或拖拽 PDF", "支持 .pdf（文本型，非扫描件）", "📄", "pdf", function (file) { processPdfFile(file) }))
    } else if (!pdfState.excelRows) {
      // 步骤2：上传 Excel（匹配规则与 Excel 入口一致：SKU 优先、UPC 兜底，v1.7.2）
      wrap.appendChild(makeDropZone("请上传采购订单 Excel", "需含 SKU 列（优先匹配）与内部参考号列（SKU 缺失兜底）", "📊", "excel", function (file) { processPdfExcelFile(file) }))
    } else {
      // 步骤3：预览 + 重置
      var btnRow = el("div", { style: "display:flex;gap:8px" }, [
        el("button", {
          onclick: function () { previewPdfChanges() },
          style: "flex:1;padding:10px;border:none;border-radius:8px;background:#7c3aed;color:#fff;font-size:13px;font-weight:600;cursor:pointer"
        }, "👁 预览修改"),
        el("button", {
          onclick: function () { resetPdfState() },
          style: "padding:10px 14px;border:1px solid #d1d5db;border-radius:8px;background:#fff;color:#6b7280;font-size:13px;cursor:pointer"
        }, "重置")
      ])
      wrap.appendChild(btnRow)
    }
    return wrap
  }

  // 入口选择器：单件 / 套装（PDF/Excel 两流共用，onSelect 接收 modeId）
  function renderModePicker(onSelect) {
    var wrap = el("div", { style: "display:flex;flex-direction:column;gap:8px" }, [
      el("div", { style: "font-size:12px;color:#374151;font-weight:600" }, "请选择修正入口")
    ])
    var row = el("div", { style: "display:flex;gap:10px" })
    var ids = ["single", "set"]
    for (var i = 0; i < ids.length; i++) {
      (function (mid) {
        var m = PDF_MODES[mid]
        var btn = el("div", {
          onclick: function () { onSelect(mid) },
          style: "flex:1;padding:14px 10px;border:2px solid #e5e7eb;border-radius:10px;text-align:center;cursor:pointer;transition:border-color 0.2s,background 0.2s"
        }, [
          el("div", { style: "font-size:26px;margin-bottom:6px" }, m.icon),
          el("div", { style: "font-size:13px;font-weight:600;color:#111827" }, m.label),
          el("div", { style: "font-size:11px;color:#9ca3af;margin-top:4px" }, m.desc)
        ])
        btn.addEventListener("mouseenter", function () { btn.style.borderColor = "#7c3aed"; btn.style.background = "#f5f3ff" })
        btn.addEventListener("mouseleave", function () { btn.style.borderColor = "#e5e7eb"; btn.style.background = "" })
        row.appendChild(btn)
      })(ids[i])
    }
    wrap.appendChild(row)
    return wrap
  }

  function selectPdfMode(modeId) {
    if (!PDF_MODES[modeId]) return
    clearPdfData()
    pdfState.mode = modeId
    refreshCard()
  }

  function clearPdfData() {
    pdfState.mode = null
    pdfState.pdfRows = null
    pdfState.pdfFileName = null
    pdfState.pdfFormat = null
    pdfState.coupon = 0
    pdfState.excelRows = null
    pdfState.excelFileName = null
    pdfState.changes = null
    pdfState.error = null
  }

  function fileExt(name) {
    var m = /\.([a-z0-9]+)$/i.exec(name || "")
    return m ? m[1].toLowerCase() : ""
  }

  function fileTypeOk(name, accept) {
    var ext = fileExt(name)
    if (accept === "pdf") return ext === "pdf"
    return ext === "xlsx" || ext === "xls"
  }

  function makeDropZone(title, hint, icon, accept, onFile) {
    var zone = el("div", {
      style: "padding:20px;border:2px dashed #d1d5db;border-radius:10px;text-align:center;cursor:pointer;transition:border-color 0.2s,background 0.2s"
    })
    zone.addEventListener("click", function () {
      var input = document.createElement("input")
      input.type = "file"
      input.accept = accept === "pdf" ? ".pdf" : ".xlsx,.xls"
      input.addEventListener("change", function () {
        var f = input.files ? input.files[0] : null
        if (f) {
          if (!fileTypeOk(f.name, accept)) {
            showToast(accept === "pdf" ? "请选择 PDF 文件" : "请选择 Excel 文件（.xlsx / .xls）", "error")
          } else {
            onFile(f)
          }
        }
        input.remove()
      })
      input.click()
    })
    zone.addEventListener("dragover", function (e) { e.preventDefault(); zone.style.borderColor = "#7c3aed"; zone.style.background = "#f5f3ff" })
    zone.addEventListener("dragleave", function () { zone.style.borderColor = "#d1d5db"; zone.style.background = "" })
    zone.addEventListener("drop", function (e) {
      e.preventDefault()
      zone.style.borderColor = "#d1d5db"; zone.style.background = ""
      var f = e.dataTransfer.files[0]
      if (!f) return
      if (!fileTypeOk(f.name, accept)) {
        showToast("文件类型不匹配：" + (accept === "pdf" ? "此步骤请上传 PDF" : "此步骤请上传 Excel（.xlsx / .xls）"), "error")
        return
      }
      onFile(f)
    })
    zone.appendChild(el("div", { style: "font-size:28px;margin-bottom:6px" }, icon))
    zone.appendChild(el("div", { style: "font-size:13px;color:#374151;font-weight:500" }, title))
    zone.appendChild(el("div", { style: "font-size:11px;color:#9ca3af;margin-top:4px" }, hint))
    return zone
  }

  function resetPdfState() {
    clearPdfData()
    refreshCard()
  }

  // ═══════════════════════════════════════════
  //  3c. Excel 导入区 UI（2026-08-19 改造：与 PDF 流完全一致——
  //      上传「PDF 转换版 Excel」→ 上传采购订单 Excel → 预览写回）
  // ═══════════════════════════════════════════
  function renderExcelZone() {
    var wrap = el("div", { style: "margin:12px 16px;display:flex;flex-direction:column;gap:10px" })

    // 当前入口徽章（可切换）
    if (excelState.mode) {
      var m = PDF_MODES[excelState.mode]
      wrap.appendChild(el("div", {
        style: "display:flex;justify-content:space-between;align-items:center;font-size:12px;color:#7c3aed;background:#f5f3ff;border:1px solid #ddd6fe;border-radius:8px;padding:8px 10px"
      }, [
        el("span", {}, m.icon + " 当前入口：" + m.label),
        el("span", {
          onclick: function () { resetExcelState() },
          style: "cursor:pointer;text-decoration:underline;color:#6b7280"
        }, "切换入口")
      ]))
    }

    // 常驻错误提示
    if (excelState.error) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#dc2626;background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:8px 10px;line-height:1.5"
      }, "⚠️ " + excelState.error))
    }

    // 步骤指示器：① 上传转换版 Excel → ② 上传采购订单 Excel → ③ 预览确认（已完成步骤可点击回退）
    if (excelState.mode) {
      var step = excelState.convRows ? (excelState.excelRows ? 3 : 2) : 1
      var steps = [
        { n: 1, label: "上传转换版 Excel", icon: "📑" },
        { n: 2, label: "上传采购订单 Excel", icon: "📊" },
        { n: 3, label: "预览确认", icon: "👁" }
      ]
      var stepBar = el("div", { style: "display:flex;align-items:center;gap:6px;font-size:11px" })
      for (var si = 0; si < steps.length; si++) {
        (function (s) {
          var done = s.n < step, active = s.n === step
          var clickable = done
          var chipStyle = "flex:1;text-align:center;padding:6px 4px;border-radius:6px;border:1px solid " +
            (active ? "#7c3aed" : (done ? "#a7f3d0" : "#e5e7eb")) + ";background:" +
            (active ? "#f5f3ff" : (done ? "#f0fdf4" : "#f9fafb")) + ";color:" +
            (active ? "#7c3aed" : (done ? "#059669" : "#9ca3af")) + ";font-weight:" + (active ? "600" : "400") +
            (clickable ? ";cursor:pointer;transition:background 0.15s,border-color 0.15s" : ";cursor:default")
          var chipAttrs = { style: chipStyle, title: clickable ? "点击返回此步骤重新上传" : "" }
          if (clickable) {
            chipAttrs.onclick = function () { jumpExcelStep(s.n) }
            chipAttrs.onmouseenter = function () { chip.style.background = "#ecfdf5"; chip.style.borderColor = "#34d399" }
            chipAttrs.onmouseleave = function () { chip.style.background = "#f0fdf4"; chip.style.borderColor = "#a7f3d0" }
          }
          var chip = el("div", chipAttrs, (done ? "✅" : "") + s.icon + " " + s.label)
          stepBar.appendChild(chip)
        })(steps[si])
        if (si < steps.length - 1) {
          stepBar.appendChild(el("span", { style: "color:#d1d5db" }, "→"))
        }
      }
      wrap.appendChild(stepBar)
    }

    // 状态提示
    if (excelState.convRows) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#059669;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:8px 10px"
      }, "✅ 已合并 " + excelState.convFiles.length + " 个转换版文件（" + excelState.convRows.length + " 行，格式" + (excelState.convFormat || "?") + (excelState.coupon ? "，Coupon=" + excelState.coupon : "") + "）"))
    }
    if (excelState.excelRows) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#059669;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:8px 10px"
      }, "✅ 采购订单 Excel 已解析：" + (excelState.excelFileName || "") + "（" + excelState.excelRows.length + " 行）"))
    }

    if (!excelState.mode) {
      // 步骤0：选择入口（单件 / 套装，与 PDF 区共用选择器）
      wrap.appendChild(renderModePicker(selectExcelMode))
    } else if (!excelState.convRows) {
      // 步骤1：上传转换版 Excel（可多个，v1.12.0）
      if (excelState.convFiles.length) {
        // 已添加文件列表（可删除单个）
        var listWrap = el("div", { style: "display:flex;flex-direction:column;gap:6px" })
        for (var fi = 0; fi < excelState.convFiles.length; fi++) {
          (function (i, f) {
            listWrap.appendChild(el("div", {
              style: "display:flex;justify-content:space-between;align-items:center;font-size:12px;background:#f5f3ff;border:1px solid #ddd6fe;border-radius:8px;padding:6px 10px"
            }, [
              el("span", { style: "flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" }, "📑 " + f.name + "（" + f.rows.length + " 行）"),
              el("span", {
                onclick: function () { removeConvFile(i) },
                title: "删除该文件",
                style: "cursor:pointer;color:#dc2626;font-weight:700;margin-left:8px;padding:0 4px;user-select:none"
              }, "✕")
            ]))
          })(fi, excelState.convFiles[fi])
        }
        wrap.appendChild(listWrap)
        // 继续添加 + 下一步（合并）
        wrap.appendChild(makeDropZone("继续添加转换版 Excel（可选）", "PDF 转成的 Excel（.xlsx / .xls），含 UPC/EAN 表头", "📑", "excel", function (file) { processExcelFile(file) }))
        wrap.appendChild(el("button", {
          onclick: function () { finishConvFiles() },
          style: "padding:10px;border:none;border-radius:8px;background:#7c3aed;color:#fff;font-size:13px;font-weight:600;cursor:pointer"
        }, "下一步：上传采购订单 Excel"))
      } else {
        wrap.appendChild(makeDropZone("点击上传或拖拽转换版 Excel（可多个）", "PDF 转成的 Excel（.xlsx / .xls），含 UPC/EAN 表头，可反复添加", "📑", "excel", function (file) { processExcelFile(file) }))
      }
    } else if (!excelState.excelRows) {
      // 步骤2：上传采购订单 Excel（匹配规则与 PDF 流一致：SKU 优先、UPC 兜底，v1.7.2）
      wrap.appendChild(makeDropZone("请上传采购订单 Excel", "需含 SKU 列（优先匹配）与内部参考号列（SKU 缺失兜底）", "📊", "excel", function (file) { processExcelOrderFile(file) }))
    } else {
      // 步骤3：预览 + 重置
      var btnRow = el("div", { style: "display:flex;gap:8px" }, [
        el("button", {
          onclick: function () { previewExcelChanges() },
          style: "flex:1;padding:10px;border:none;border-radius:8px;background:#7c3aed;color:#fff;font-size:13px;font-weight:600;cursor:pointer"
        }, "👁 预览修改"),
        el("button", {
          onclick: function () { resetExcelState() },
          style: "padding:10px 14px;border:1px solid #d1d5db;border-radius:8px;background:#fff;color:#6b7280;font-size:13px;cursor:pointer"
        }, "重置")
      ])
      wrap.appendChild(btnRow)
    }
    return wrap
  }

  // ═══════════════════════════════════════════
  //  3d. 商品库更新区（v1.11.0 新增：上传商品库 Excel → 按 UPC 查 product.product → 预览 → 写回 name/brand）
  // ═══════════════════════════════════════════
  function renderProductZone() {
    var wrap = el("div", { style: "margin:12px 16px;display:flex;flex-direction:column;gap:10px" })

    // 常驻错误提示
    if (productState.error) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#dc2626;background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:8px 10px;line-height:1.5"
      }, "⚠️ " + productState.error))
    }

    // 解析成功提示（含匹配统计：有差异 / 值相同 / 未匹配 / 重复；缺 UPC 跳过数单独提示）
    if (productState.previewRows) {
      var ok = productState.previewRows.filter(function (r) { return r.status === "ok" }).length
      var same = productState.previewRows.filter(function (r) { return r.status === "same" }).length
      var notfound = productState.previewRows.filter(function (r) { return r.status === "notfound" }).length
      var dup = productState.previewRows.filter(function (r) { return r.status === "dup" }).length
      var skipTip = (productState.rows && productState.rows.skippedCount) ? "，" + productState.rows.skippedCount + " 行缺 UPC 已跳过" : ""
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#059669;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:8px 10px;line-height:1.5"
      }, "✅ 已解析 " + (productState.fileName || "") + "（" + productState.previewRows.length + " 行）：" + ok + " 行有差异将更新" + (same ? "，" + same + " 行值相同" : "") + (notfound ? "，" + notfound + " 行未匹配到产品" : "") + (dup ? "，" + dup + " 行 UPC 重复" : "") + skipTip))
    } else if (productState.rows) {
      wrap.appendChild(el("div", {
        style: "font-size:12px;color:#059669;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:8px 10px"
      }, "✅ 商品库已解析：" + (productState.fileName || "") + "（" + productState.rows.length + " 行），正在查询 Odoo 产品..."))
    }

    if (!productState.rows) {
      // 步骤1：上传商品库 Excel
      wrap.appendChild(makeDropZone("点击上传或拖拽商品库 Excel", "需含 UPC / 品牌 / 中文简称 三列（表头第 1 行）", "🏷", "excel", function (file) { processProductFile(file) }))
    } else if (productState.previewRows) {
      // 步骤2：预览 + 重置
      var btnRow = el("div", { style: "display:flex;gap:8px" }, [
        el("button", {
          onclick: function () { previewProductChanges() },
          style: "flex:1;padding:10px;border:none;border-radius:8px;background:#7c3aed;color:#fff;font-size:13px;font-weight:600;cursor:pointer"
        }, "👁 预览修改"),
        el("button", {
          onclick: function () { resetProductState() },
          style: "padding:10px 14px;border:1px solid #d1d5db;border-radius:8px;background:#fff;color:#6b7280;font-size:13px;cursor:pointer"
        }, "重置")
      ])
      wrap.appendChild(btnRow)
    } else {
      // 查询中
      wrap.appendChild(el("div", { style: "text-align:center;padding:16px;font-size:12px;color:#6b7280" }, "⏳ 正在按 UPC 查询 Odoo 产品..."))
    }
    return wrap
  }

  // 商品库 Excel 解析：表头第 1 行，按列名定位 UPC/品牌/中文简称（正则容错列名变体）
  function parseProductExcel(data) {
    var wb = XLSX.read(data, { type: "array" })
    var ws = wb.Sheets[wb.SheetNames[0]]
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true })
    if (!rows.length) throw new Error("Excel 为空或格式不正确")
    var header = rows[0].map(function (h) { return String(h || "").trim() })
    function colIdxFuzzy(keywords) {
      for (var j = 0; j < header.length; j++) {
        var h = header[j]
        if (!h) continue
        for (var k = 0; k < keywords.length; k++) {
          if (h.indexOf(keywords[k]) !== -1) return j
        }
      }
      return -1
    }
    function cell(row, i) {
      if (i < 0) return ""
      var v = row[i]
      return (v === undefined || v === null) ? "" : v
    }
    var idx = {
      upc: colIdxFuzzy(["UPC", "EAN", "条码", "Barcode"]),
      brand: colIdxFuzzy(["品牌", "Brand"]),
      shortName: colIdxFuzzy(["中文简称", "简称", "中文名", "名称"])
    }
    if (idx.upc < 0) throw new Error("未找到 UPC 列（表头需含「UPC」）")
    if (idx.brand < 0) throw new Error("未找到品牌列（表头需含「品牌」）")
    if (idx.shortName < 0) throw new Error("未找到中文简称列（表头需含「中文简称」）")

    var result = []
    result.skippedCount = 0   // 缺 UPC 被跳过的行数（无法定位产品，主动告知用户）
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i]
      if (!row || !row.length) continue
      var upc = String(cell(row, idx.upc)).trim()
      if (!upc) { result.skippedCount++; continue }
      result.push({
        rowIndex: i,
        upc: upc,
        brand: String(cell(row, idx.brand)).trim(),
        // 中文简称样本实测含换行符，写入前转空格（name 单行显示）
        shortName: String(cell(row, idx.shortName)).replace(/\n+/g, " ").trim()
      })
    }
    if (!result.length) throw new Error("Excel 中没有有效数据（UPC 列为空）")
    return result
  }

  // 按 UPC（=default_code）批量查 product.product；一个 UPC 可能对应多个产品 → 返回数组
  // context.lang=zh_CN：Odoo name 字段是多语言翻译字段，带 lang 读中文翻译值（=用户界面实际显示的值），比对/预览才准（2026-08-21 修复）
  async function loadProductByUpc(upcList) {
    var result = await rpcCall("/web/dataset/call_kw/product.product/search_read", {
      model: "product.product", method: "search_read",
      args: [], kwargs: {
        domain: [["default_code", "in", upcList]],
        fields: ["id", "name", "default_code", PRODUCT_ODOO_FIELDS.brand],
        context: { lang: "zh_CN" }
      }
    })
    var map = {}
    for (var i = 0; i < result.length; i++) {
      var r = result[i]
      var code = String(r.default_code || "").trim()
      if (!code) continue
      if (!map[code]) map[code] = []
      map[code].push({ id: r.id, name: r.name || "", brand: r[PRODUCT_ODOO_FIELDS.brand] || "", defaultCode: code })
    }
    return map
  }

  // 合并 Excel 行与 Odoo 匹配结果 → 预览行（状态：ok 将更新 / same 值相同 / notfound 未匹配 / dup UPC 重复）
  // 读取 Odoo 原 name/brand 与 Excel 比对展示（用户 2026-08-21 要求）；写回不做比对：匹配到唯一产品 → 直接更新（空值也照写）
  function buildProductPreviewRows(rows, map) {
    return rows.map(function (r) {
      var list = map[r.upc] || null
      if (!list || !list.length) {
        return { upc: r.upc, brand: r.brand, shortName: r.shortName, status: "notfound", error: "UPC 未匹配到产品（default_code 无对应）", checked: false, hasAction: false }
      }
      if (list.length > 1) {
        return { upc: r.upc, brand: r.brand, shortName: r.shortName, status: "dup", error: "UPC 匹配到 " + list.length + " 个产品，无法确定目标，需人工处理", checked: false, hasAction: false }
      }
      var p = list[0]
      // 与 Odoo 原值比对（仅展示用途，写回不受影响）
      var nameSame = (r.shortName === (p.name || ""))
      var brandSame = (r.brand === (p.brand || ""))
      var isSame = nameSame && brandSame
      return {
        upc: r.upc,
        oldName: p.name || "", newName: r.shortName,
        oldBrand: p.brand || "", newBrand: r.brand,
        nameSame: nameSame, brandSame: brandSame,
        odooProductId: p.id,
        status: isSame ? "same" : "ok",
        error: "",
        checked: true,   // 匹配到即默认勾选（值相同也照写，用户确认「匹配到就更新」）
        hasAction: true
      }
    })
  }

  async function processProductFile(file) {
    try {
      productState.error = null
      var data = await file.arrayBuffer()
      var rows = parseProductExcel(data)
      productState.rows = rows
      productState.fileName = file.name
      productState.previewRows = null
      refreshCard()
      updateLoadingOverlay("正在按 UPC 查询 Odoo 产品...")
      // 去重 UPC 后分块查询（Odoo domain in 一次不宜过大）
      var upcSet = {}
      for (var i = 0; i < rows.length; i++) upcSet[rows[i].upc] = true
      var upcs = Object.keys(upcSet)
      var map = {}
      for (var s = 0; s < upcs.length; s += 500) {
        var slice = upcs.slice(s, s + 500)
        var part = await loadProductByUpc(slice)
        for (var k in part) map[k] = part[k]
      }
      productState.previewRows = buildProductPreviewRows(rows, map)
      updateLoadingOverlay(null)
      refreshCard()
      var ok = productState.previewRows.filter(function (r) { return r.status === "ok" }).length
      var same = productState.previewRows.filter(function (r) { return r.status === "same" }).length
      var notfound = productState.previewRows.filter(function (r) { return r.status === "notfound" }).length
      var dup = productState.previewRows.filter(function (r) { return r.status === "dup" }).length
      var skipTip = rows.skippedCount ? "，" + rows.skippedCount + " 行缺 UPC 已跳过" : ""
      showToast("查询完成：" + ok + " 行有差异" + (same ? "，" + same + " 行值相同" : "") + (notfound ? "，" + notfound + " 行未匹配" : "") + (dup ? "，" + dup + " 行 UPC 重复" : "") + skipTip, (notfound + dup) > 0 ? "info" : "success")
    } catch (err) {
      productState.error = err.message || String(err)
      updateLoadingOverlay(null)
      refreshCard()
      showToast("解析失败: " + productState.error, "error")
      console.error("[Odoo Product]", err)
    }
  }

  function previewProductChanges() {
    if (!productState.previewRows) return
    buildProductPreviewModal(productState.previewRows, productState.fileName)
  }

  function resetProductState() {
    productState.rows = null
    productState.fileName = null
    productState.previewRows = null
    productState.error = null
    refreshCard()
  }

  // 预览 Modal：UPC + 中文简称/品牌 旧值→新值 对照 + 勾选（复用现有 Modal 基础设施）
  function buildProductPreviewModal(previewRows, fileName) {
    removeModals()
    var overlay = el("div", {
      id: PREFIX + "modal_overlay",
      style: "position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:100010;display:flex;align-items:center;justify-content:center"
    })
    overlay.addEventListener("click", function (e) { if (e.target === overlay) removeModals() })

    var modal = el("div", {
      style: "background:#fff;border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,0.2);width:92vw;max-width:1100px;max-height:86vh;display:flex;flex-direction:column;overflow:hidden"
    })
    overlay.appendChild(modal)

    // Header
    modal.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;gap:12px;padding:14px 20px;border-bottom:1px solid #e5e7eb"
    }, [
      el("h2", { style: "margin:0;font-size:16px;color:#111827;flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" }, "🏷 商品库更新 · " + (fileName || "预览")),
      el("button", {
        onclick: function () { removeModals() },
        style: "width:28px;height:28px;border:1px solid #e5e7eb;border-radius:6px;background:#fff;color:#6b7280;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center"
      }, "✕")
    ]))

    // 统计栏（含「值相同」行——已与 Odoo 原值比对一致，仍会照写）
    var ok = previewRows.filter(function (r) { return r.status === "ok" }).length
    var same = previewRows.filter(function (r) { return r.status === "same" }).length
    var notfound = previewRows.filter(function (r) { return r.status === "notfound" }).length
    var dup = previewRows.filter(function (r) { return r.status === "dup" }).length
    modal.appendChild(el("div", {
      style: "padding:10px 20px;background:#f0f9ff;font-size:12px;color:#1e40af"
    }, "共 " + previewRows.length + " 行  ·  " + ok + " 行有差异将更新" + (same > 0 ? "  ·  " + same + " 行与 Odoo 值相同" : "") + (notfound > 0 ? "  ·  " + notfound + " 行未匹配到产品" : "") + (dup > 0 ? "  ·  " + dup + " 行 UPC 重复" : "") + "  ·  仅勾选行会写入 Odoo"))

    // 表格
    var tableWrap = el("div", { style: "flex:1;overflow:auto;max-height:58vh" })
    tableWrap.appendChild(renderProductPreviewTable(previewRows))
    modal.appendChild(tableWrap)

    // Footer
    modal.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-top:1px solid #e5e7eb"
    }, [
      el("span", { style: "font-size:12px;color:#6b7280" }, "未匹配到产品或 UPC 重复的行不可勾选"),
      el("button", {
        id: PREFIX + "product_confirm_btn",
        onclick: function () { executeProductUpdate(previewRows, fileName) },
        style: "padding:8px 20px;border:none;border-radius:8px;background:#7c3aed;color:#fff;font-size:13px;font-weight:600;cursor:pointer"
      }, "✅ 确认写入 Odoo")
    ]))

    document.body.appendChild(overlay)
    currentModalOverlay = overlay
  }

  function renderProductPreviewTable(previewRows) {
    var table = el("table", { style: "width:100%;border-collapse:collapse;font-size:12px" })
    var hoverStyle = document.createElement("style")
    hoverStyle.textContent = "#" + PREFIX + "modal_overlay tbody tr:hover td{background:#f5f3ff!important}"
    table.appendChild(hoverStyle)
    var thead = document.createElement("thead")
    var tr = document.createElement("tr")
    tr.style.cssText = "background:#f3f4f6"
    var headers = [
      { w: "36px", html: '<input type="checkbox" id="' + PREFIX + 'product_select_all" checked style="cursor:pointer">' },
      { w: "130px", text: "UPC" },
      { text: "中文简称（旧 → 新）" },
      { text: "品牌（旧 → 新）" },
      { w: "130px", text: "状态" }
    ]
    for (var i = 0; i < headers.length; i++) {
      var th = document.createElement("th")
      th.style.cssText = "padding:8px 8px;text-align:center;font-weight:600;color:#374151;border-bottom:2px solid #e5e7eb;font-size:11px;white-space:nowrap;vertical-align:middle;position:sticky;top:0;z-index:3;background:#f3f4f6"
      if (headers[i].w) th.style.width = headers[i].w
      if (headers[i].html) th.innerHTML = headers[i].html
      else th.textContent = headers[i].text
      tr.appendChild(th)
    }
    thead.appendChild(tr)
    table.appendChild(thead)

    // 全选/取消：DOM 与数据行 r.checked 双向同步（2026-08-21 修复：此前只改 DOM 导致执行时误写全部匹配行）
    var selectAll = table.querySelector("#" + PREFIX + "product_select_all")
    if (selectAll) {
      selectAll.addEventListener("change", function () {
        var checked = selectAll.checked
        var boxes = table.querySelectorAll("input[type=checkbox]:not(#" + PREFIX + "product_select_all)")
        for (var i = 0; i < boxes.length; i++) {
          var enable = !boxes[i].disabled
          boxes[i].checked = enable && checked
          if (boxes[i]._row) boxes[i]._row.checked = boxes[i].checked   // 同步数据层，防执行时误写
        }
        refreshProductConfirmCount()
      })
    }

    var statusMeta = {
      ok:        { text: "✅ 将更新", color: "#059669", bg: "#f0fdf4", border: "#bbf7d0" },
      same:      { text: "ℹ️ 值相同", color: "#6b7280", bg: "#f9fafb", border: "#e5e7eb" },
      notfound:  { text: "⚠️ 未匹配", color: "#dc2626", bg: "#fef2f2", border: "#fca5a5" },
      dup:       { text: "🔁 UPC 重复", color: "#d97706", bg: "#fffbeb", border: "#fde68a" }
    }
    var tbody = document.createElement("tbody")
    for (var i = 0; i < previewRows.length; i++) {
      (function (r) {
        var meta = statusMeta[r.status] || statusMeta.notfound
        var trRow = document.createElement("tr")
        trRow.style.cssText = "border-bottom:1px solid #f3f4f6"
        // 勾选
        var tdCb = document.createElement("td")
        tdCb.style.cssText = "padding:6px 10px;text-align:center;vertical-align:middle"
        var cb = document.createElement("input")
        cb.type = "checkbox"
        cb.checked = !!r.checked
        cb.disabled = !r.hasAction
        cb._row = r                                   // 全选时通过它同步数据层
        cb.addEventListener("change", function () {
          r.checked = cb.checked
          refreshProductConfirmCount()
        })
        tdCb.appendChild(cb)
        trRow.appendChild(tdCb)
        // UPC
        trRow.appendChild(td(r.upc, "#111827", "600"))
        // 中文简称 / 品牌 旧→新
        trRow.appendChild(buildDiffTd(r.oldName, r.newName))
        trRow.appendChild(buildDiffTd(r.oldBrand, r.newBrand))
        // 状态徽章
        var tdStatus = document.createElement("td")
        tdStatus.style.cssText = "padding:6px 10px;text-align:center;vertical-align:middle"
        tdStatus.appendChild(el("span", {
          style: "display:inline-block;padding:2px 10px;border-radius:10px;font-size:11px;white-space:nowrap;color:" + meta.color + ";background:" + meta.bg + ";border:1px solid " + meta.border
        }, meta.text))
        trRow.appendChild(tdStatus)
        tbody.appendChild(trRow)
      })(previewRows[i])
    }
    table.appendChild(tbody)
    refreshProductConfirmCount()
    return table
  }

  // 确认按钮实时显示勾选行数；0 行时禁用（防误点导致批量误写，2026-08-21 修复）
  function refreshProductConfirmCount() {
    var btn = document.getElementById(PREFIX + "product_confirm_btn")
    if (!btn) return
    var boxes = document.querySelectorAll("#" + PREFIX + "modal_overlay tbody input[type=checkbox]")
    var n = 0
    for (var i = 0; i < boxes.length; i++) {
      if (boxes[i].checked && !boxes[i].disabled) n++
    }
    btn.textContent = n > 0 ? "✅ 确认写入 Odoo（" + n + " 行）" : "✅ 确认写入 Odoo（0 行）"
    btn.disabled = n === 0
    btn.style.opacity = n === 0 ? "0.5" : "1"
    btn.style.cursor = n === 0 ? "not-allowed" : "pointer"
  }

  // 旧值 → 新值 单元格（两行小字；textContent 自动转义防 XSS）
  // Excel 值为空时按空写入（用户确认「匹配到就按 Excel 值更新，空不空无所谓」），新值为空直接显示空白、不加提示
  function buildDiffTd(oldVal, newVal) {
    var t = document.createElement("td")
    t.style.cssText = "padding:6px 10px;text-align:center;vertical-align:middle"
    var oldDiv = document.createElement("div")
    oldDiv.style.cssText = "font-size:11px;color:#9ca3af;word-break:break-all;max-width:240px"
    oldDiv.textContent = (oldVal === undefined || oldVal === null || oldVal === "") ? "—" : oldVal
    t.appendChild(oldDiv)
    var newDiv = document.createElement("div")
    newDiv.style.cssText = "font-size:12px;color:#7c3aed;font-weight:600;word-break:break-all;max-width:240px;margin-top:2px"
    newDiv.textContent = (newVal === undefined || newVal === null || newVal === "") ? "" : newVal
    t.appendChild(newDiv)
    return t
  }

  // 批量写回 product.product（name / brand，Char 直接写；不做比对，空值也写入；勾选行才写）
  async function executeProductUpdate(previewRows, fileName) {
    // 以 DOM 勾选状态为准同步数据层（防全选等操作导致数据与界面不同步，2026-08-21 修复）
    var boxes = document.querySelectorAll("#" + PREFIX + "modal_overlay tbody input[type=checkbox]")
    for (var i = 0; i < previewRows.length && i < boxes.length; i++) {
      previewRows[i].checked = boxes[i].checked
    }
    var selected = previewRows.filter(function (r) { return r.checked && r.hasAction })
    if (!selected.length) return
    removeModals()

    var logEntry = {
      timestamp: new Date().toISOString(),
      fileName: fileName,
      totalRows: previewRows.length,
      selectedRows: selected.length,
      results: []
    }

    updateLoadingOverlay("正在写入 Odoo 产品...")
    for (var i = 0; i < selected.length; i++) {
      var row = selected[i]
      // 不做比对：按 Excel 值直接写入 name/brand（空值也写入，用户 2026-08-21 确认）
      var payload = {}
      var changedDesc = []
      payload[PRODUCT_ODOO_FIELDS.shortName] = row.newName
      changedDesc.push("中文简称=" + row.newName)
      payload[PRODUCT_ODOO_FIELDS.brand] = row.newBrand
      changedDesc.push("品牌=" + row.newBrand)
      try {
        // context.lang=zh_CN：name 是 Odoo 多语言翻译字段，带 lang 写简体中文翻译（用户 zh_CN 界面显示的才是新值）；
        // source 值此前已写入且=Excel 值，无需重复；brand 无 translate 不受影响（2026-08-21 修复）
        await rpcCall("/web/dataset/call_kw/product.product/write", {
          model: "product.product", method: "write", args: [[row.odooProductId], payload],
          kwargs: { context: { lang: "zh_CN" } }
        })
        logEntry.results.push({ upc: row.upc, productId: row.odooProductId, status: "success", changes: changedDesc, old: { name: row.oldName, brand: row.oldBrand }, new: { name: row.newName, brand: row.newBrand } })
      } catch (err) {
        console.error("[Odoo Product] 写回失败 id=" + row.odooProductId, err)
        logEntry.results.push({ upc: row.upc, productId: row.odooProductId, status: "failed", error: err.message || "写入失败", changes: changedDesc, old: { name: row.oldName, brand: row.oldBrand }, new: { name: row.newName, brand: row.newBrand } })
      }
    }
    updateLoadingOverlay(null)

    var skipped = previewRows.filter(function (r) { return !r.checked || !r.hasAction })
    for (var s = 0; s < skipped.length; s++) {
      logEntry.results.push({ upc: skipped[s].upc, status: "skipped", reason: skipped[s].error || "用户跳过或无需操作" })
    }

    saveLog(logEntry)
    refreshCard()
    showResultToast(logEntry)
    // 写回成功后提示刷新 Odoo 页面（列表视图有缓存，需 F5 或重进菜单才显示新值）
    setTimeout(function () {
      showToast("💡 写入已完成，如 Odoo 页面未显示新值，请刷新页面（F5）或重新进入该菜单", "info")
    }, 4200)
  }

  function selectExcelMode(modeId) {
    if (!PDF_MODES[modeId]) return
    clearExcelData()
    excelState.mode = modeId
    refreshCard()
  }

  function clearExcelData() {
    excelState.mode = null
    excelState.convFiles = []
    excelState.convRows = null
    excelState.convFileName = null
    excelState.convFormat = null
    excelState.coupon = 0
    excelState.excelRows = null
    excelState.excelFileName = null
    excelState.changes = null
    excelState.error = null
  }

  function resetExcelState() {
    clearExcelData()
    refreshCard()
  }

  // 步骤条跳转（Excel 流）：1=上传转换版 Excel, 2=上传采购订单 Excel, 3=预览确认
  // v1.12.0：回步骤1 保留已添加文件列表（可增删后重新点「下一步」合并），清掉合并结果与下游
  function jumpExcelStep(target) {
    if (target === 1) {
      excelState.convRows = null
      excelState.convFileName = null
      excelState.convFormat = null
      excelState.coupon = 0
      excelState.excelRows = null
      excelState.excelFileName = null
      excelState.changes = null
      excelState.error = null
    } else if (target === 2) {
      if (!excelState.convRows) return
      excelState.excelRows = null
      excelState.excelFileName = null
      excelState.changes = null
    } else if (target === 3) {
      if (!excelState.convRows || !excelState.excelRows) return
    }
    refreshCard()
  }

  // 步骤条跳转：回到已完成步骤重新执行（防误上传，如误传 PDF 后点「上传 PDF」重传）
  // target: 1=上传PDF, 2=上传Excel, 3=预览确认
  function jumpPdfStep(target) {
    if (target === 1) {
      // 回到上传 PDF：清空 PDF 及其下游（Excel / 匹配结果 / 格式 / Coupon）
      pdfState.pdfRows = null
      pdfState.pdfFileName = null
      pdfState.pdfFormat = null
      pdfState.coupon = 0
      pdfState.excelRows = null
      pdfState.excelFileName = null
      pdfState.changes = null
      pdfState.error = null
    } else if (target === 2) {
      // 回到上传 Excel：清空 Excel 及匹配结果，保留已解析的 PDF
      if (!pdfState.pdfRows) return
      pdfState.excelRows = null
      pdfState.excelFileName = null
      pdfState.changes = null
    } else if (target === 3) {
      // 预览确认：需 PDF 与 Excel 均已上传（无需清数据，UI 自动切到该步骤）
      if (!pdfState.pdfRows || !pdfState.excelRows) return
    }
    refreshCard()
  }

  // ═══════════════════════════════════════════
  //  4. 标签条（已移除：Excel 预览窗口 Tab 管理待重新设计后恢复）
  // ═══════════════════════════════════════════

  // ═══════════════════════════════════════════
  //  5. 日志区
  // ═══════════════════════════════════════════
  function renderLogSection(log) {
    var success = log.results.filter(function (r) { return r.status === "success" }).length
    var failed = log.results.filter(function (r) { return r.status === "failed" }).length
    var skipped = log.results.filter(function (r) { return r.status === "skipped" }).length

    var section = el("div", { style: "padding:10px 12px;background:#f9fafb;border-radius:8px;margin-bottom:12px" }, [
      el("div", { style: "display:flex;gap:12px;margin-bottom:8px;font-size:12px" }, [
        el("span", { style: "color:#059669" }, "✓ " + success + " 成功"),
        el("span", { style: "color:#d97706" }, failed > 0 ? "✗ " + failed + " 失败" : ""),
        el("span", { style: "color:#9ca3af" }, "— " + skipped + " 跳过")
      ]),
      el("div", { style: "font-size:11px;color:#6b7280" }, [
        el("span", {}, log.fileName || "未知文件"),
        el("span", { style: "margin-left:8px" }, new Date(log.timestamp).toLocaleString("zh-CN"))
      ]),
      el("div", { style: "margin-top:8px;display:flex;gap:8px" }, [
        el("button", {
          onclick: downloadLog,
          style: "padding:5px 12px;font-size:11px;border:1px solid #d1d5db;border-radius:6px;background:#fff;color:#374151;cursor:pointer"
        }, "📥 下载日志"),
        el("button", {
          onclick: function () { clearLog(); refreshCard() },
          style: "padding:5px 12px;font-size:11px;border:1px solid #fecaca;border-radius:6px;background:#fff;color:#dc2626;cursor:pointer"
        }, "清除")
      ])
    ])

    return section
  }

  // ═══════════════════════════════════════════
  //  6. Modal 基础设施（Excel 预览 Modal 已移除待重写；PDF 预览 Modal 见 7c）
  // ═══════════════════════════════════════════
  var currentModalOverlay = null

  function td(text, color, weight, extraStyle) {
    var t = document.createElement("td")
    var s = "padding:6px 10px;color:" + color + ";text-align:center;vertical-align:middle"
    if (weight) s += ";font-weight:" + weight
    if (extraStyle) s += ";font-family:" + extraStyle
    t.style.cssText = s
    t.textContent = text
    return t
  }

  // ═══════════════════════════════════════════
  //  7. Modal 关闭（最小化/恢复/标签恢复已移除，待 Excel 流程重写后按需恢复）
  // ═══════════════════════════════════════════
  function removeModals() {
    if (currentModalOverlay) {
      currentModalOverlay.remove()
      currentModalOverlay = null
    }
    var existing = document.getElementById(PREFIX + "modal_overlay")
    if (existing) existing.remove()
  }

  // ═══════════════════════════════════════════
  //  7b. PDF 区流程：选入口 → 上传 → 解析 → 匹配(按入口分发) → 预览 → 写回
  // ═══════════════════════════════════════════
  async function processPdfFile(file) {
    try {
      showToast("解析 PDF 中...", "info")
      pdfState.error = null
      var data = await file.arrayBuffer()
      var result = await parsePdf(data)
      var pdfRows = result.rows || []
      if (!pdfRows.length) {
        pdfState.error = "PDF 中未找到表格数据（0 行）"
        refreshCard()
        showToast(pdfState.error, "error")
        return
      }
      pdfState.pdfRows = pdfRows
      pdfState.pdfFileName = file.name
      pdfState.pdfFormat = result.format
      pdfState.coupon = result.coupon || 0
      pdfState.excelRows = null
      pdfState.excelFileName = null
      pdfState.changes = null
      refreshCard()
      showToast("PDF 解析成功：" + pdfRows.length + " 行（格式" + result.format + "，Coupon=" + pdfState.coupon + "），请上传 Excel", "success")
    } catch (err) {
      pdfState.error = err.message || String(err)
      refreshCard()
      showToast("PDF 解析失败: " + pdfState.error, "error")
      console.error("[Odoo PDF]", err)
    }
  }

  async function processPdfExcelFile(file) {
    try {
      pdfState.error = null
      var data = await file.arrayBuffer()
      var excelRows = parseOrderExcel(data)
      if (!excelRows.length) {
        pdfState.error = "Excel 中没有有效数据"
        refreshCard()
        showToast(pdfState.error, "error")
        return
      }
      pdfState.excelRows = excelRows
      pdfState.excelFileName = file.name
      pdfState.changes = applyPdfByMode(pdfState.pdfRows, excelRows, pdfState.mode, pdfState.coupon, "pdf")
      refreshCard()
      showToast("匹配完成，可点击预览", "success")
    } catch (err) {
      pdfState.error = err.message || String(err)
      refreshCard()
      showToast("Excel 解析失败: " + pdfState.error, "error")
      console.error("[Odoo PDF]", err)
    }
  }

  // ── Excel 导入流（2026-08-19 改造：与 PDF 流同逻辑；v1.12.0 支持多个转换版文件）──
  // 步骤1：上传「PDF 转换版 Excel」（可多个）→ 解析后追加到 convFiles（不立即合并，点「下一步」再合并）
  async function processExcelFile(file) {
    try {
      showToast("解析转换版 Excel 中...", "info")
      excelState.error = null
      var data = await file.arrayBuffer()
      var result = parseConvertedPdfExcel(data)
      if (!result.rows.length) {
        excelState.error = "转换版 Excel 中未找到表格数据（0 行）"
        refreshCard()
        showToast(excelState.error, "error")
        return
      }
      excelState.convFiles.push({ name: file.name, rows: result.rows, format: result.format, coupon: result.coupon || 0 })
      excelState.excelRows = null
      excelState.excelFileName = null
      excelState.changes = null
      refreshCard()
      showToast("已添加 " + file.name + "（" + result.rows.length + " 行，格式" + result.format + "），可继续添加或点「下一步」", "success")
    } catch (err) {
      // 解析失败：仅提示，不影响已添加的文件
      showToast("转换版 Excel 解析失败: " + (err.message || String(err)), "error")
      console.error("[Odoo Excel]", err)
    }
  }

  // 步骤1「下一步」：合并全部转换版文件 → convRows（含 Coupon 合并、文件名汇总）
  function finishConvFiles() {
    if (!excelState.convFiles.length) return
    excelState.convRows = mergeConvFiles(excelState.convFiles)
    excelState.coupon = mergeCoupon(excelState.convFiles)
    excelState.convFileName = excelState.convFiles.map(function (f) { return f.name }).join(", ")
    excelState.convFormat = excelState.convFiles[0].format
    excelState.excelRows = null
    excelState.excelFileName = null
    excelState.changes = null
    excelState.error = null
    refreshCard()
    showToast("已合并 " + excelState.convFiles.length + " 个转换版文件（" + excelState.convRows.length + " 行），请上传采购订单 Excel", "success")
  }

  // 删除已添加的转换版文件（合并结果作废，需重新点「下一步」）
  function removeConvFile(idx) {
    if (idx < 0 || idx >= excelState.convFiles.length) return
    excelState.convFiles.splice(idx, 1)
    excelState.convRows = null
    excelState.convFileName = null
    excelState.convFormat = null
    excelState.coupon = 0
    excelState.excelRows = null
    excelState.excelFileName = null
    excelState.changes = null
    refreshCard()
  }

  // 步骤2：上传采购订单 Excel → 与 PDF 流同一套匹配修正逻辑（applyPdfByMode）
  async function processExcelOrderFile(file) {
    try {
      excelState.error = null
      var data = await file.arrayBuffer()
      var excelRows = parseOrderExcel(data)
      if (!excelRows.length) {
        excelState.error = "采购订单 Excel 中没有有效数据"
        refreshCard()
        showToast(excelState.error, "error")
        return
      }
      excelState.excelRows = excelRows
      excelState.excelFileName = file.name
      excelState.changes = applyPdfByMode(excelState.convRows, excelRows, excelState.mode, excelState.coupon, "excel")
      refreshCard()
      showToast("匹配完成，可点击预览", "success")
    } catch (err) {
      excelState.error = err.message || String(err)
      refreshCard()
      showToast("采购订单 Excel 解析失败: " + excelState.error, "error")
      console.error("[Odoo Excel]", err)
    }
  }

  // 预览：查 Odoo 订单行 → Modal（与 PDF 流同一预览/写回通道）
  async function previewExcelChanges() {
    if (!excelState.changes) return
    try {
      showToast("查询 Odoo 订单行中...", "info")
      var previewRows = await buildPdfPreviewRows(excelState.changes)
      updateLoadingOverlay(null)
      var modeLabel = excelState.mode ? "[" + PDF_MODES[excelState.mode].label + "] " : ""
      buildPdfPreviewModal(previewRows, modeLabel + excelState.convFileName + " + " + excelState.excelFileName)
    } catch (err) {
      updateLoadingOverlay(null)
      showToast("预览失败: " + (err.message || err), "error")
      console.error("[Odoo Excel]", err)
    }
  }

  // 查 PO → 订单行，构建 lineMap（v1.9：key = ref|defaultCode 兜底 + ref|defaultCode|包装件数 精确）
  // 2026-08-19：每行优先用「参考号」查 PO（两者都查 Odoo name 字段），查不到再用「订单关联」；命中行按两个键都注册
  // 返回 { map: 行查找表, cnt: ref|defaultCode 出现次数 }——包装匹配失败时按 cnt 判断是否唯一可兜底（避免错配同 UPC 其他包装行）
  async function loadOrderLineMap(refPairs) {
    var lineMap = {}, cnt = {}
    for (var r = 0; r < refPairs.length; r++) {
      var orderRef = refPairs[r].orderRef, partnerRef = refPairs[r].partnerRef
      try {
        var po = null
        if (partnerRef) po = await searchPoByRef(partnerRef)   // 先匹配参考号
        if (!po && orderRef) po = await searchPoByRef(orderRef) // 查不到再匹配订单关联
        if (po && po.orderLineIds.length) {
          var lines = await getOrderLines(po.orderLineIds)
          var refs = []
          if (orderRef) refs.push(orderRef)
          if (partnerRef && partnerRef !== orderRef) refs.push(partnerRef)
          for (var j = 0; j < lines.length; j++) {
            if (!lines[j].defaultCode) continue
            for (var x = 0; x < refs.length; x++) {
              var k0 = refs[x] + "|" + lines[j].defaultCode
              cnt[k0] = (cnt[k0] || 0) + 1
              lineMap[k0] = lines[j]
              if (lines[j].packQty !== "") lineMap[k0 + "|" + lines[j].packQty] = lines[j]
            }
          }
        }
      } catch (e) {
        console.error("[Odoo Excel Importer] 查询失败 " + (partnerRef || orderRef), e)
      }
    }
    return { map: lineMap, cnt: cnt }
  }

  // 预览 Odoo 查询缓存（v1.12.3，2026-08-25 用户需求）：同一 changes 数组（引用相同）重复预览时
  // 复用已查的 lineData，不再重复请求 Odoo（Modal 关闭重开 / 反复点预览都不重查，直到数据重新生成）
  var previewLineCache = { changes: null, lineData: null }

  // 构建 PDF 区预览行（含 Odoo 订单行匹配；参考号/订单关联去重成 refPairs）
  async function buildPdfPreviewRows(changes) {
    var seen = {}, refPairs = []
    for (var i = 0; i < changes.length; i++) {
      var c = changes[i]
      if (!c.orderRef && !c.partnerRef) continue
      var k = (c.partnerRef || "") + "|" + (c.orderRef || "")
      if (seen[k]) continue
      seen[k] = true
      refPairs.push({ orderRef: c.orderRef || "", partnerRef: c.partnerRef || "" })
    }
    var lineData
    if (previewLineCache.changes === changes && previewLineCache.lineData) {
      lineData = previewLineCache.lineData   // 缓存命中：不查 Odoo
    } else {
      lineData = await loadOrderLineMap(refPairs)
      previewLineCache.changes = changes
      previewLineCache.lineData = lineData
    }
    return buildPreviewRows(changes, lineData)
  }

  // changes + lineData({map,cnt}) → 预览行（PDF/Excel 两流共用；lineData 已查好时直接走此函数）
  // v1.9 匹配规则：UPC = Excel「内部参考号」↔ Odoo product.default_code；包装 = Excel「订单行/包装」件数 ↔ Odoo product_packaging.qty
  // - Excel 有包装件数 → 精确匹配 UPC+包装；失败且该 UPC 在 PO 中唯一 → 按 UPC 兜底；多行 → 报未找到（不降级到其他包装行，避免写错行）
  // - Excel 无包装件数 → 仅当该 UPC 在 PO 中唯一时按 UPC 命中；多行 → 报未找到
  function buildPreviewRows(changes, lineData) {
    var lineMap = lineData.map, cnt = lineData.cnt || {}
    var previewRows = []
    for (var k = 0; k < changes.length; k++) {
      var c = changes[k]
      var line = null
      var refs = []
      if (c.orderRef) refs.push(c.orderRef)
      if (c.partnerRef && c.partnerRef !== c.orderRef) refs.push(c.partnerRef)
      for (var rr = 0; rr < refs.length && !line; rr++) {
        var ref = refs[rr]
        var k0 = ref + "|" + c.upc
        if (c.packQty) {
          line = lineMap[ref + "|" + c.upc + "|" + c.packQty] || null
          if (line) break
        }
        // 精确包装匹配不到（或 Excel 无包装件数）：仅当该 UPC 在 PO 中唯一时按 UPC 兜底
        if (cnt[k0] === 1) { line = lineMap[k0] || null; break }
        if (cnt[k0] > 1) break
      }
      // v3.0（2026-08-26 用户需求）：价格原值从 Odoo 订单行取（单价→price_unit、0.9箱规价→box_wholesale_price、0.9总价→price_subtotal）
      // 缺货行（kind=outstock）价格已定死（置 0），不参与回填
      if (line && c.kind !== "outstock" && c.fields) {
        var odooOldMap = { unitPrice: "price_unit", boxPrice09: "box_wholesale_price", total09: "price_subtotal" }
        for (var fi = 0; fi < c.fields.length; fi++) {
          var cf = c.fields[fi]
          var odooKey = odooOldMap[cf.key]
          if (!odooKey) continue
          var oldVal = line[odooKey]
          if (oldVal === null || oldVal === undefined || oldVal === "" || cf.newValue === null || cf.newValue === undefined) continue
          cf.oldValue = oldVal
          cf.changed = !numEq(oldVal, cf.newValue)
          cf.reason = cf.changed ? (cf.key === "unitPrice" ? REASONS.unitPrice(cf.newValue, cf.expr, oldVal)
            : cf.key === "boxPrice09" ? REASONS.boxPrice09(cf.newValue, cf.expr, oldVal)
            : REASONS.total09(cf.newValue, cf.expr, oldVal)) : null
        }
      }
      var hasChanged = false
      for (var f = 0; f < c.fields.length; f++) { if (c.fields[f].changed) hasChanged = true }
      // v1.7.7：数据一致（白底无标记）的行也可勾选写入；仅「未找到匹配的订单行」不可勾选
      var lineErr = line ? null : "未找到匹配的订单行（UPC " + (c.upc || "—") + (c.packQty ? " / 包装 " + c.packQty + " 件" : "") + "）"
      previewRows.push({
        idx: k,
        upc: c.upc, catalog: c.catalog, shop: c.shop, orderRef: c.orderRef || c.partnerRef || "",
        mode: c.mode || null,
        odooLineId: line ? line.id : null,
        odooLineName: line ? line.name : null,
        kind: c.kind,
        qtyTarget: c.qtyTarget, qtySum: c.qtySum, qtyMismatch: c.qtyMismatch, qtyDetail: c.qtyDetail,
        fields: c.fields,
        hasAction: !lineErr,
        checked: hasChanged || c.kind === "outstock",
        error: lineErr
      })
    }
    // 按订单关联号排序
    previewRows.sort(function (a, b) {
      var ra = String(a.orderRef || ""), rb = String(b.orderRef || "")
      return ra < rb ? -1 : (ra > rb ? 1 : 0)
    })
    // 排序后重编 idx
    for (var s = 0; s < previewRows.length; s++) previewRows[s].idx = s
    return previewRows
  }

  function fmtNum(v) {
    return (v === null || v === undefined || v === "") ? "—" : String(v)
  }

  async function previewPdfChanges() {
    if (!pdfState.changes) return
    try {
      showToast("查询 Odoo 订单行中...", "info")
      var previewRows = await buildPdfPreviewRows(pdfState.changes)
      updateLoadingOverlay(null)
      var modeLabel = pdfState.mode ? "[" + PDF_MODES[pdfState.mode].label + "] " : ""
      buildPdfPreviewModal(previewRows, modeLabel + pdfState.pdfFileName + " + " + pdfState.excelFileName)
    } catch (err) {
      updateLoadingOverlay(null)
      showToast("预览失败: " + (err.message || err), "error")
      console.error("[Odoo PDF]", err)
    }
  }

  // 字段展示顺序（按入口；2026-08-19 起 Excel 流与 PDF 流同一布局，source 仅保留兼容签名）
  // v3.0：去掉箱规价（用户确认：Odoo 无对应字段，不再展示）
  // v3.1.1：去掉整箱批发价列（用户确认 2026-08-27；缺货行 boxWholesale 字段仍保留在 fields 里，置 0 写回不受影响）
  function fieldOrder(mode, source) {
    var order = ["boxQty", "unitPrice"]
    if (mode === "set") order = order.concat(["boxPrice09", "total09"])
    return order.concat(["remark"])
  }

  var FIELD_LABELS = {
    boxQty: "包装数量", boxPrice09: "0.9箱规价",
    unitPrice: "单价", total09: "0.9总价", remark: "备注"
  }

  function findField(fields, key) {
    for (var i = 0; i < fields.length; i++) if (fields[i].key === key) return fields[i]
    return null
  }

  function fieldDisplayVal(f) {
    var v = f.changed ? f.newValue : f.oldValue
    return (v === null || v === undefined || v === "") ? "" : String(v)
  }

  // source: 'pdf' | 'excel'（2026-08-18 双流共用：Excel 流无 PDF 比对，标题/字段列布局按 source 区分）
  function buildPdfPreviewModal(previewRows, fileName, source) {
    removeModals()
    source = source || "pdf"
    // mode 优先取预览行自带标记（applyPdfByMode 已按入口盖章），兜底再读两流状态（2026-08-19：Excel 流也走 pdf 布局）
    var mode = (previewRows.length && previewRows[0].mode) || (source === "excel" ? excelState.mode : pdfState.mode)
    var titleIcon = source === "excel" ? "📊" : "📄"
    // 去重订单关联号列表（2026-08-18 用户要求：显示具体关联号，如 P00393|P00395|...）
    var refSet = {}
    for (var ri = 0; ri < previewRows.length; ri++) {
      if (previewRows[ri].orderRef) refSet[previewRows[ri].orderRef] = true
    }
    var refList = Object.keys(refSet)
    var overlay = el("div", {
      id: PREFIX + "modal_overlay",
      style: "position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:100010;display:flex;align-items:center;justify-content:center"
    })
    overlay.addEventListener("click", function (e) { if (e.target === overlay) removeModals() })

    var modal = el("div", {
      style: "background:#fff;border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,0.2);width:94vw;max-width:1400px;max-height:88vh;display:flex;flex-direction:column;overflow:hidden"
    })
    overlay.appendChild(modal)

    // Header
    var h2Title = el("h2", { style: "margin:0;font-size:16px;color:#111827;flex:1;min-width:0" }, titleIcon + " " + (fileName || (source === "excel" ? "Excel 导入预览" : "PDF 修正预览")))
    if (refList.length > 0) {
      h2Title.appendChild(el("span", { style: "color:#7c3aed;font-size:13px;margin-left:10px;font-weight:600;word-break:break-all" }, refList.join(" | ")))
    }
    modal.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;gap:12px;padding:14px 20px;border-bottom:1px solid #e5e7eb"
    }, [
      h2Title,
      el("button", {
        onclick: function () { removeModals() },
        style: "width:28px;height:28px;border:1px solid #e5e7eb;border-radius:6px;background:#fff;color:#6b7280;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center"
      }, "✕")
    ]))

    // 统计栏
    var actionable = previewRows.filter(function (r) { return r.hasAction && !r.error }).length
    var errors = previewRows.filter(function (r) { return r.error }).length
    var outstock = previewRows.filter(function (r) { return r.kind === "outstock" }).length
    var qtyMis = previewRows.filter(function (r) { return r.qtyMismatch }).length
    modal.appendChild(el("div", {
      style: "padding:10px 20px;background:#f0f9ff;font-size:12px;color:#1e40af"
    }, "共 " + previewRows.length + " 行  ·  " + actionable + " 行可写入" + (qtyMis > 0 ? "  ·  " + qtyMis + " 行数量不一致" : "") + (outstock > 0 ? "  ·  " + outstock + " 行缺货" : "") + (errors > 0 ? "  ·  " + errors + " 行匹配失败" : "") + "  · 不一致字段可编辑，悬浮可查看原因，底部为各列总和"))

    // 表格
    var tableWrap = el("div", { style: "flex:1;overflow:auto;max-height:60vh" })
    tableWrap.appendChild(renderPdfPreviewTable(previewRows, mode, source))
    modal.appendChild(tableWrap)

    // Footer
    modal.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-top:1px solid #e5e7eb"
    }, [
      el("span", { style: "font-size:12px;color:#6b7280" }, "勾选要写入 Odoo 的行，可修改输入框后写入"),
      el("button", {
        id: PREFIX + "pdf_confirm_btn",
        onclick: function () { executePdfUpdate(previewRows, fileName) },
        style: "padding:8px 20px;border:none;border-radius:8px;background:#7c3aed;color:#fff;font-size:13px;font-weight:600;cursor:pointer"
      }, "✅ 确认写入 Odoo")
    ]))

    document.body.appendChild(overlay)
    currentModalOverlay = overlay
  }

  function renderPdfPreviewTable(previewRows, mode, source) {
    var order = fieldOrder(mode, source)
    var isExcel = source === "excel"
    var table = el("table", { style: "width:100%;border-collapse:collapse;font-size:12px" })
    // hover 行高亮（2026-08-18 美化）
    var hoverStyle = document.createElement("style")
    hoverStyle.textContent = "#" + PREFIX + "modal_overlay tbody tr:hover td{background:#f5f3ff!important}"
    table.appendChild(hoverStyle)
    var thead = document.createElement("thead")
    var tr = document.createElement("tr")
    tr.style.cssText = "background:#f3f4f6"
    var headers = [
      { w: "34px", html: '<input type="checkbox" id="' + PREFIX + 'pdf_select_all" checked style="cursor:pointer">' },
      { w: "80px", text: "店铺" },
      { w: "120px", text: "UPC" },
      { w: "90px", text: "SKU" },
      { w: "90px", text: "订单关联" },
      { text: "产品名" }
    ]
    // PDF 流额外两列：数量核对 + 供应商 Qty（Excel 流无 PDF 源，不展示；2026-08-19 改列名）
    if (!isExcel) {
      headers.push({ w: "110px", text: "数量核对" }, { w: "90px", text: "供应商 Qty" })
    }
    for (var i = 0; i < order.length; i++) {
      headers.push({ w: "110px", text: FIELD_LABELS[order[i]] })
    }
    for (var i = 0; i < headers.length; i++) {
      var th = document.createElement("th")
      // position:sticky 固定表头：滚动表格时列名行始终可见（2026-08-19 用户需求）
      th.style.cssText = "padding:8px 8px;text-align:center;font-weight:600;color:#374151;border-bottom:2px solid #e5e7eb;font-size:11px;white-space:nowrap;vertical-align:middle;position:sticky;top:0;z-index:3;background:#f3f4f6"
      if (headers[i].w) th.style.width = headers[i].w
      if (headers[i].html) th.innerHTML = headers[i].html
      else th.textContent = headers[i].text
      tr.appendChild(th)
    }
    thead.appendChild(tr)
    table.appendChild(thead)

    // 输入框宽度自适应：以该列最长数值为基准（最少 7 位数字），2026-08-18 用户要求
    var colMaxLen = {}
    for (var ci = 0; ci < order.length; ci++) colMaxLen[order[ci]] = 7
    for (var rj = 0; rj < previewRows.length; rj++) {
      var flds = previewRows[rj].fields
      for (var fj = 0; fj < flds.length; fj++) {
        var ff = flds[fj]
        var ln = String(fieldDisplayVal(ff)).length
        if (ln > colMaxLen[ff.key]) colMaxLen[ff.key] = ln
      }
    }

    var tbody = document.createElement("tbody")
    // v3.1（2026-08-27 用户需求）：按订单关联分组，每组尾部插一行「组小计」（只累计 0.9总价列）
    // previewRows 已按 orderRef 排序（buildPreviewRows），同组行连续，遍历时组切换即插小计行
    // 单件入口无 total09 字段（order 不含 total09）→ 不插组小计行
    var groupCtx = []
    var hasTotal09 = order.indexOf("total09") >= 0
    for (var i = 0; i < previewRows.length; i++) {
      var pr = previewRows[i]
      var curRef = String(pr.orderRef || "")
      var lastGroup = groupCtx.length ? groupCtx[groupCtx.length - 1] : null
      if (hasTotal09 && lastGroup && curRef !== lastGroup.ref) {
        tbody.appendChild(renderPdfGroupSubtotalRow(lastGroup, order, source))
      }
      if (!lastGroup || curRef !== lastGroup.ref) {
        groupCtx.push({ ref: curRef, rowIdx: [], cell: null })
        lastGroup = groupCtx[groupCtx.length - 1]
      }
      lastGroup.rowIdx.push(i)
      tbody.appendChild(renderPdfPreviewRow(pr, order, colMaxLen, i, source))
    }
    if (hasTotal09 && groupCtx.length) tbody.appendChild(renderPdfGroupSubtotalRow(groupCtx[groupCtx.length - 1], order, source))
    table.appendChild(tbody)

    // 列总和行（2026-08-18 用户需求：所有数值列全表合计，随输入框编辑实时刷新）
    table.appendChild(renderPdfTotalRow(previewRows, order, source, groupCtx))

    setTimeout(function () {
      var sa = document.getElementById(PREFIX + "pdf_select_all")
      if (sa) {
        sa.addEventListener("change", function () {
          var cbs = document.querySelectorAll("." + PREFIX + "pdf_row_cb:not([disabled])")
          for (var i = 0; i < cbs.length; i++) cbs[i].checked = sa.checked
          for (var i = 0; i < previewRows.length; i++) {
            if (previewRows[i].hasAction && !previewRows[i].error) previewRows[i].checked = sa.checked
          }
        })
      }
    }, 10)
    return table
  }

  // 组小计行（v3.1，2026-08-27 用户需求）：按订单关联分组，只对 0.9总价（total09）列做组内合计
  // 列结构与普通行对齐（勾选/店铺/UPC/SKU/订单关联/产品名 + [数量核对/供应商Qty] + 字段列），其余列占位
  function renderPdfGroupSubtotalRow(group, order, source) {
    var isExcel = source === "excel"
    var tr = document.createElement("tr")
    tr.style.cssText = "background:#f5f3ff;border-bottom:1px solid #e5e7eb"
    tr.appendChild(el("td", { style: "padding:4px 8px" }))
    for (var i = 0; i < 3; i++) tr.appendChild(el("td", { style: "padding:4px 8px" }))
    var tdRef = el("td", {
      style: "padding:4px 8px;font-size:11px;color:#6d28d9;font-weight:600;text-align:center;white-space:nowrap"
    }, (group.ref || "—") + " 小计")
    tr.appendChild(tdRef)
    tr.appendChild(el("td", { style: "padding:4px 8px" }))
    if (!isExcel) {
      tr.appendChild(el("td", { style: "padding:4px 8px" }))
      tr.appendChild(el("td", { style: "padding:4px 8px" }))
    }
    for (var o = 0; o < order.length; o++) {
      var key = order[o]
      if (key === "total09") {
        group.cell = el("td", {
          style: "padding:4px 8px;font-size:12px;color:#6d28d9;font-weight:600;text-align:center;white-space:nowrap"
        }, "0")
        tr.appendChild(group.cell)
      } else {
        tr.appendChild(el("td", { style: "padding:4px 8px;font-size:11px;color:#c4b5fd;text-align:center" }, "—"))
      }
    }
    return tr
  }

  function renderPdfPreviewRow(row, order, colMaxLen, rowIndex, source) {
    var tr = document.createElement("tr")
    var isExcel = source === "excel"
    var alertRow = row.error || row.qtyMismatch || row.kind === "outstock"
    // 交替行色美化；警示行红底优先
    var bg = alertRow ? "#fef2f2" : ((rowIndex || 0) % 2 === 0 ? "#ffffff" : "#f9fafb")
    var border = alertRow ? "#fca5a5" : "#e5e7eb"
    tr.style.cssText = "border-bottom:1px solid " + border + ";background:" + bg

    // 勾选框
    var tdCb = document.createElement("td")
    tdCb.style.cssText = "padding:6px 8px;text-align:center"
    if (row.hasAction && !row.error) {
      var cb = document.createElement("input")
      cb.type = "checkbox"; cb.className = PREFIX + "pdf_row_cb"
      cb.checked = row.checked; cb.style.cursor = "pointer"
      cb.addEventListener("change", function () { row.checked = cb.checked })
      tdCb.appendChild(cb)
    } else {
      tdCb.innerHTML = '<span style="color:#d1d5db">—</span>'
    }
    tr.appendChild(tdCb)

    tr.appendChild(td(row.shop || "—", "#374151", "600"))
    tr.appendChild(td(row.upc, "#6b7280", null, "monospace;font-size:11px"))
    tr.appendChild(td(row.catalog || "—", "#6b7280", null, "font-size:11px"))
    tr.appendChild(td(row.orderRef || "—", "#374151", null, "font-size:11px"))
    tr.appendChild(td(row.odooLineName || "—", row.error ? "#ef4444" : "#374151"))

    // 数量核对 + PDF Qty 两列（仅 PDF 流；Excel 流无 PDF 源，2026-08-18）
    if (!isExcel) {
      // 数量核对：PDF Qty vs 包装数量（多条求和）；一致 ✓ 绿色，不一致红色 + hover 原因
      var qtyTd = document.createElement("td")
      qtyTd.style.cssText = "padding:6px 8px;font-size:11px;text-align:center;vertical-align:middle;white-space:nowrap"
      if (row.kind === "outstock") {
        qtyTd.innerHTML = '<span style="color:#dc2626;font-weight:600;text-decoration:underline dotted;cursor:help">缺货</span>'
        var osTip = null
        qtyTd.addEventListener("mouseenter", function () {
          if (osTip) return
          osTip = el("div", {
            style: "position:fixed;z-index:100050;background:#1f2937;color:#f9fafb;font-size:11px;padding:6px 10px;border-radius:6px;max-width:340px;line-height:1.5;pointer-events:none"
          }, REASONS.outstock)
          document.body.appendChild(osTip)
          var r = qtyTd.getBoundingClientRect()
          osTip.style.left = Math.max(8, Math.min(r.left, window.innerWidth - osTip.offsetWidth - 8)) + "px"
          osTip.style.top = Math.max(8, r.top - osTip.offsetHeight - 8) + "px"
        })
        qtyTd.addEventListener("mouseleave", function () {
          if (osTip) { osTip.remove(); osTip = null }
        })
      } else if (row.qtyMismatch || (row.qtyDetail && row.qtyDetail.length > 1)) {
        // 数量不一致 或 同 SKU 多条订单关联相加 → 悬浮显示明细（2026-08-19 用户需求）
        var isMis = row.qtyMismatch
        qtyTd.innerHTML = isMis
          ? '<span style="color:#dc2626;font-weight:600;text-decoration:underline dotted;cursor:help">' + fmtNum(row.qtySum) + ' ≠ ' + fmtNum(row.qtyTarget) + '</span>'
          : '<span style="color:#16a34a;font-weight:600;text-decoration:underline dotted;cursor:help">✓ ' + fmtNum(row.qtyTarget) + '</span>'
        var qTip = null
        qtyTd.addEventListener("mouseenter", function () {
          if (qTip) return
          qTip = el("div", {
            style: "position:fixed;z-index:100050;background:#1f2937;color:#f9fafb;font-size:11px;padding:8px 10px;border-radius:6px;max-width:340px;line-height:1.6;pointer-events:none"
          })
          var qhtml = '<div style="font-weight:600;margin-bottom:4px;color:#e5e7eb">供应商 Qty = ' + escHtml(fmtNum(row.qtyTarget)) + '</div>'
          var dd = row.qtyDetail || []
          if (dd.length > 1) {
            qhtml += '<div style="margin-bottom:2px">数量由以下订单关联相加：</div>'
            for (var di = 0; di < dd.length; di++) {
              qhtml += '<div>· ' + escHtml(dd[di].ref || "—") + '（SKU ' + escHtml(dd[di].sku || "—") + '）：' + fmtNum(dd[di].qty) + ' 件</div>'
            }
          }
          qhtml += '<div>' + escHtml(isMis ? REASONS.qtySum(row.qtySum, row.qtyTarget) : ("合计 " + fmtNum(row.qtySum) + " = 供应商 Qty " + fmtNum(row.qtyTarget))) + '</div>'
          qTip.innerHTML = qhtml
          document.body.appendChild(qTip)
          var r = qtyTd.getBoundingClientRect()
          qTip.style.left = Math.max(8, Math.min(r.left, window.innerWidth - qTip.offsetWidth - 8)) + "px"
          qTip.style.top = Math.max(8, r.top - qTip.offsetHeight - 8) + "px"
        })
        qtyTd.addEventListener("mouseleave", function () {
          if (qTip) { qTip.remove(); qTip = null }
        })
      } else {
        qtyTd.innerHTML = '<span style="color:#16a34a;font-weight:600">✓ ' + fmtNum(row.qtyTarget) + '</span>'
      }
      tr.appendChild(qtyTd)

      // PDF Qty 列：PDF 原始 qty 值（= qtyTarget，outstock 无值显示 —，2026-08-18 新增）
      var pdfQtyTd = document.createElement("td")
      pdfQtyTd.style.cssText = "padding:6px 8px;font-size:11px;text-align:center;vertical-align:middle;white-space:nowrap;color:#374151"
      pdfQtyTd.textContent = fmtNum(row.qtyTarget)
      tr.appendChild(pdfQtyTd)
    }

    // 字段输入框 + PDF 来源对照小字 + 不一致原因悬浮（2026-08-18 增强）
    row._inputs = {}
    for (var oi = 0; oi < order.length; oi++) {
      var key = order[oi]
      var f = findField(row.fields, key)
      var ftd = document.createElement("td")
      ftd.style.cssText = "padding:4px 6px;text-align:center;vertical-align:middle"
      if (!f) {
        ftd.innerHTML = '<span style="color:#d1d5db">—</span>'
      } else {
        var changed = f.changed && !row.error
        var input = document.createElement("input")
        input.type = "text"
        input.value = fieldDisplayVal(f)
        // 宽度：该列所有行计算值中最多位数者为基底，全列统一；+1ch+14px 余量保证数字完整显示
        var w = (colMaxLen && colMaxLen[key]) || 7
        input.style.cssText = "width:calc(" + (w + 1) + "ch + 14px);padding:5px 6px;border:1px solid " + (changed ? "#fca5a5" : "#d1d5db") +
          ";border-radius:6px;font-size:12px;color:#111827;background:" + (changed ? "#fff5f5" : "#fff") +
          ";box-sizing:border-box;text-align:center"
        // 输入框与 Odoo 原值平级（同一行居中，2026-08-18：去掉 PDF 来源小字；v3.0：原值改 Odoo 值）
        var fieldRow = el("div", { style: "display:flex;align-items:center;justify-content:center;gap:8px;white-space:nowrap" })
        fieldRow.appendChild(input)
        if (changed && f.oldValue !== null && f.oldValue !== undefined && f.oldValue !== "") {
          fieldRow.appendChild(el("span", { style: "font-size:10px;color:#9ca3af;text-decoration:line-through" }, "Odoo 原值: " + fmtNum(f.oldValue)))
        }
        ftd.appendChild(fieldRow)
        row._inputs[key] = input;
        // 不一致原因悬浮气泡 + 编辑实时刷新总和（IIFE 捕获 input/f，避免闭包共享）
        (function (inp, fld) {
          if (fld.reason && changed) {
            var tip = null
            inp.addEventListener("mouseenter", function () {
              if (tip || !fld.reason) return
              tip = el("div", {
                style: "position:fixed;z-index:100050;background:#1f2937;color:#f9fafb;font-size:11px;padding:8px 10px;border-radius:6px;max-width:440px;line-height:1.5;pointer-events:none"
              })
              // 两行悬浮：第一行 PDF 原始字段与数据，第二行计算过程（2026-08-18 用户需求）
              tip.innerHTML = fld.pdfRaw
                ? '<div style="font-weight:600;margin-bottom:4px;color:#e5e7eb;white-space:pre-wrap">' + escHtml(fld.pdfRaw) + '</div>' +
                  '<div style="white-space:pre-wrap">' + escHtml(fld.reason) + '</div>'
                : escHtml(fld.reason)
              document.body.appendChild(tip)
              var r = inp.getBoundingClientRect()
              tip.style.left = Math.max(8, Math.min(r.left, window.innerWidth - tip.offsetWidth - 8)) + "px"
              tip.style.top = Math.max(8, r.top - tip.offsetHeight - 8) + "px"
            })
            inp.addEventListener("mouseleave", function () {
              if (tip) { tip.remove(); tip = null }
            })
          }
          inp.addEventListener("input", function () { refreshTotal() })
        })(input, f)
      }
      tr.appendChild(ftd)
    }
    return tr
  }

  // ── 列总和行（2026-08-18：所有数值列全表合计，随输入框编辑实时刷新）──
  var SUM_KEYS = { boxQty: 1, unitPrice: 1, boxPrice09: 1, total09: 1 }
  var pdfTotalCtx = null

  function refreshTotal() {
    var ctx = pdfTotalCtx
    if (!ctx) return
    var sums = {}
    for (var i = 0; i < ctx.rows.length; i++) {
      var row = ctx.rows[i]
      if (!row._inputs) continue
      for (var k in ctx.cells) {
        if (k === "pdfQty") continue   // 供应商 Qty 非输入框，下方单独累加
        var inp = row._inputs[k]
        if (!inp) continue
        var n = parseFloat(inp.value)
        if (!isNaN(n)) sums[k] = (sums[k] || 0) + n
      }
    }
    // 供应商 Qty 列总和：按 SKU 组去重后累加（同 SKU 多行共享同一 qtyTarget，只算一次）
    if (ctx.cells.pdfQty) {
      var pq = 0
      var seenGroup = {}
      for (var pi = 0; pi < ctx.rows.length; pi++) {
        var rt = ctx.rows[pi].qtyTarget
        var rn = parseFloat(rt)
        if (isNaN(rn)) continue
        var gk = (ctx.rows[pi].catalog || "") + "|" + (ctx.rows[pi].upc || "")
        if (seenGroup[gk]) continue
        seenGroup[gk] = true
        pq += rn
      }
      sums.pdfQty = pq
    }
    for (var k2 in ctx.cells) {
      ctx.cells[k2].textContent = sums[k2] === undefined ? "0" : String(Math.round(sums[k2] * 1000) / 1000)
    }
    // 组小计（v3.1，2026-08-27 用户需求）：按订单关联分组，只累计 0.9总价（total09），随编辑实时刷新
    if (ctx.groups) {
      for (var g = 0; g < ctx.groups.length; g++) {
        var grp = ctx.groups[g]
        if (!grp.cell) continue
        var gsum = 0
        for (var gi = 0; gi < grp.rowIdx.length; gi++) {
          var gr = ctx.rows[grp.rowIdx[gi]]
          if (!gr._inputs || !gr._inputs.total09) continue
          var gn = parseFloat(gr._inputs.total09.value)
          if (!isNaN(gn)) gsum += gn
        }
        grp.cell.textContent = String(Math.round(gsum * 1000) / 1000)
      }
    }
  }

  function renderPdfTotalRow(previewRows, order, source, groupCtx) {
    var isExcel = source === "excel"
    var tfoot = document.createElement("tfoot")
    var tr = document.createElement("tr")
    tr.style.cssText = "background:#f3f4f6"
    // 前 6 列（勾选/店铺/UPC/SKU/订单关联/产品名）合并为标签（v1.7.8：加 SKU 列，colSpan 5→6）
    var tdLabel = document.createElement("td")
    tdLabel.colSpan = 6
    tdLabel.style.cssText = "padding:6px 8px;font-size:11px;color:#374151;font-weight:600;text-align:center;white-space:nowrap"
    tdLabel.textContent = "总和（全表合计）"
    tr.appendChild(tdLabel)
    var pdfQtyCell = null
    if (!isExcel) {
      var tdQty = document.createElement("td")
      tdQty.style.cssText = "padding:6px 8px;font-size:11px;color:#9ca3af;text-align:center"
      tr.appendChild(tdQty)
      // 供应商 Qty 列（qtyTarget 求和，2026-08-19 用户需求）
      var tdPdfQty = document.createElement("td")
      tdPdfQty.style.cssText = "padding:6px 8px;font-size:12px;color:#111827;font-weight:600;text-align:center;white-space:nowrap"
      tdPdfQty.textContent = "0"
      tr.appendChild(tdPdfQty)
      pdfQtyCell = tdPdfQty
    }
    var cells = {}
    for (var i = 0; i < order.length; i++) {
      var key = order[i]
      var td = document.createElement("td")
      td.style.cssText = "padding:6px 8px;font-size:12px;color:#111827;font-weight:600;text-align:center;white-space:nowrap"
      if (SUM_KEYS[key]) {
        td.textContent = "0"
        cells[key] = td
      } else {
        td.textContent = "—"
      }
      tr.appendChild(td)
    }
    if (pdfQtyCell) cells.pdfQty = pdfQtyCell
    tfoot.appendChild(tr)
    pdfTotalCtx = { rows: previewRows, cells: cells, groups: groupCtx || [] }
    refreshTotal()
    return tfoot
  }

  async function executePdfUpdate(previewRows, fileName) {
    var selected = previewRows.filter(function (r) { return r.checked && r.hasAction && !r.error })
    if (!selected.length) return
    removeModals()

    var logEntry = {
      timestamp: new Date().toISOString(),
      fileName: fileName,
      totalRows: previewRows.length,
      selectedRows: selected.length,
      results: []
    }

    for (var i = 0; i < selected.length; i++) {
      var row = selected[i]
      var payload = {}
      var changedDesc = []
      for (var j = 0; j < row.fields.length; j++) {
        var f = row.fields[j]
        if (!f.odooField) continue
        var input = row._inputs && row._inputs[f.key]
        var val = input ? input.value : fieldDisplayVal(f)
        if (f.key === "remark") {
          payload[f.odooField] = String(val || "")
          changedDesc.push(f.label + "=" + val)
        } else {
          var n = parseFloat(val)
          if (!isNaN(n)) {
            payload[f.odooField] = n
            changedDesc.push(f.label + "=" + n)
          }
        }
      }
      if (Object.keys(payload).length === 0) continue
      try {
        await updateOrderLine(row.odooLineId, payload)
        logEntry.results.push({ upc: row.upc, shop: row.shop, orderRef: row.orderRef, status: "success", changes: changedDesc })
      } catch (err) {
        console.error("[Odoo PDF] 写回失败 lineId=" + row.odooLineId, err)
        logEntry.results.push({ upc: row.upc, shop: row.shop, orderRef: row.orderRef, status: "failed", error: err.message || "写入失败", changes: changedDesc })
      }
    }

    var skipped = previewRows.filter(function (r) { return !r.checked || !r.hasAction || r.error })
    for (var i = 0; i < skipped.length; i++) {
      logEntry.results.push({ upc: skipped[i].upc, shop: skipped[i].shop, orderRef: skipped[i].orderRef, status: "skipped", reason: skipped[i].error || "用户跳过或无需操作" })
    }

    saveLog(logEntry)
    refreshCard()
    showResultToast(logEntry)
  }

  // ═══════════════════════════════════════════
  //  8. 文件处理（Excel 导入主流程已移除，待与用户对齐修改逻辑后重写）
  // ═══════════════════════════════════════════

  function updateLoadingOverlay(text) {
    var id = PREFIX + "loading"
    var existing = document.getElementById(id)
    if (text) {
      if (!existing) {
        existing = el("div", {
          id: id,
          style: "position:fixed;top:16px;left:50%;transform:translateX(-50%);background:#7c3aed;color:#fff;padding:8px 20px;border-radius:20px;font-size:13px;z-index:100020;box-shadow:0 4px 12px rgba(124,58,237,0.4)"
        })
        document.body.appendChild(existing)
      }
      existing.textContent = "⏳ " + text
    } else {
      if (existing) existing.remove()
    }
  }

  // ═══════════════════════════════════════════
  //  9. 执行更新（Excel 流写回已移除，待重新设计；PDF 流写回见 7b executePdfUpdate）
  // ═══════════════════════════════════════════

  function showResultToast(logEntry) {
    var success = logEntry.results.filter(function (r) { return r.status === "success" }).length
    var failed = logEntry.results.filter(function (r) { return r.status === "failed" }).length
    var failedItem = logEntry.results.find(function (r) { return r.status === "failed" })
    var msg = "完成: " + success + " 行已更新"
    if (failed > 0) {
      msg += "，" + failed + " 行失败"
      if (failedItem && failedItem.error) msg += "（" + failedItem.error + "）"
    }
    showToast(msg, failed > 0 ? "error" : "success")
  }

  // ═══════════════════════════════════════════
  //  10. Toast
  // ═══════════════════════════════════════════
  function showToast(msg, type) {
    type = type || "info"
    var id = PREFIX + "toast"
    var existing = document.getElementById(id)
    if (existing) existing.remove()
    var colors = { info: "#1890ff", success: "#52c41a", error: "#ff4d4f" }
    var toast = el("div", {
      id: id,
      style: "position:fixed;bottom:80px;right:20px;z-index:100030;padding:12px 20px;border-radius:8px;color:#fff;font-size:14px;font-family:sans-serif;background:" + (colors[type] || colors.info) + ";box-shadow:0 4px 12px rgba(0,0,0,0.25);transition:opacity 0.3s;opacity:1;max-width:400px"
    }, msg)
    document.body.appendChild(toast)
    setTimeout(function () { toast.style.opacity = "0"; setTimeout(function () { toast.remove() }, 300) }, 4000)
  }

  // ═══════════════════════════════════════════
  //  11. 初始化
  // ═══════════════════════════════════════════
  function inject() {
    if (document.getElementById(PREFIX + "btn")) return
    if (!document.body) { setTimeout(inject, 200); return }
    document.body.appendChild(createDraggableButton())
  }

  if (isAnyPage()) inject()

  window.addEventListener("hashchange", function () {
    var existing = document.getElementById(PREFIX + "btn")
    if (isAnyPage()) {
      if (!existing) inject()
      // 页面类型切换时重置默认 Tab（采购页 → excel，产品页 → product），卡片开着则刷新
      var defaultTab = isProductPage() ? "product" : "excel"
      if (appState.activeTab !== defaultTab && (isPurchaseOrderPage() || isProductPage())) {
        appState.activeTab = defaultTab
        refreshCard()
      }
    } else {
      if (existing) existing.remove()
      if (cardEl) { cardEl.remove(); cardEl = null }
      removeModals()
      appState.cardOpen = false
    }
  })
})()

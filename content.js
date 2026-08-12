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

  async function getOrderLines(ids) {
    var result = await rpcCall("/web/dataset/call_kw/purchase.order.line/search_read", {
      model: "purchase.order.line", method: "search_read",
      args: [], kwargs: {
        domain: [["id", "in", ids]],
        fields: ["id", "name", "price_unit", "box_wholesale_price", "remark"]
      }
    })
    return result.map(function (r) {
      var m = r.name.match(/\[(\d+)\]/)
      return {
        id: r.id, name: r.name, upc: m ? m[1] : "",
        price_unit: r.price_unit, box_wholesale_price: r.box_wholesale_price,
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
  //  Excel 解析
  // ═══════════════════════════════════════════
  function parseExcel(data) {
    var wb = XLSX.read(data, { type: "array" })
    var ws = wb.Sheets[wb.SheetNames[0]]
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1 })
    if (rows.length < 2) throw new Error("Excel 为空或格式不正确")

    var results = []
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i]
      if (!row || row.length === 0) continue
      var upc = String(row[0] || "").trim()
      if (!upc) continue
      results.push({
        upc: upc,
        productName: String(row[1] || "").trim(),
        orderRef: String(row[5] || "").trim(),
        unitPrice: parseFloatNum(row[10]),
        boxWholesalePrice: parseFloatNum(row[11]),
        purchaseResult: String(row[12] || "").trim()
      })
    }
    return results
  }

  // ═══════════════════════════════════════════
  //  业务逻辑
  // ═══════════════════════════════════════════
  function determineRemark(purchaseResult) {
    if (purchaseResult.indexOf("可能缺货") !== -1) return "可能缺货"
    if (purchaseResult.indexOf("缺货") !== -1) return "缺货"
    return null
  }

  function groupByOrderRef(rows) {
    var map = new Map()
    for (var i = 0; i < rows.length; i++) {
      var key = rows[i].orderRef
      if (!map.has(key)) map.set(key, [])
      map.get(key).push(rows[i])
    }
    return map
  }

  function buildPreviewRow(excelRow, odooLine) {
    var remark = determineRemark(excelRow.purchaseResult)
    var hasPrice = excelRow.unitPrice !== null || excelRow.boxWholesalePrice !== null
    var hasRemark = remark !== null
    var hasAction = hasPrice || hasRemark

    var changes = []
    if (hasRemark) changes.push({ type: "remark", label: "备注", oldValue: (odooLine ? (odooLine.remark || "(空)") : "?"), newValue: remark })
    if (excelRow.unitPrice !== null) changes.push({ type: "price_unit", label: "单价", oldValue: odooLine ? formatPrice(odooLine.price_unit) : "?", newValue: formatPrice(excelRow.unitPrice) })
    if (excelRow.boxWholesalePrice !== null) changes.push({ type: "box_wholesale_price", label: "整箱批发价", oldValue: odooLine ? formatPrice(odooLine.box_wholesale_price) : "?", newValue: formatPrice(excelRow.boxWholesalePrice) })

    var rowType = "none"
    if (hasPrice && hasRemark) rowType = "mixed"
    else if (hasPrice) rowType = "price"
    else if (hasRemark) rowType = "remark"

    return {
      upc: excelRow.upc, productName: excelRow.productName, orderRef: excelRow.orderRef,
      purchaseResult: excelRow.purchaseResult,
      odooLineId: odooLine ? odooLine.id : null, odooLineName: odooLine ? odooLine.name : null,
      rowType: rowType, hasAction: hasAction, changes: changes,
      checked: hasAction, error: odooLine ? null : "未找到匹配的订单行"
    }
  }

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
    previewRows: [],
    tabs: [],           // { id, label, fileName, previewRows }
    activeModal: null    // 'preview' | 'result'
  }

  // ═══════════════════════════════════════════
  //  颜色常量
  // ═══════════════════════════════════════════
  var ROW_COLORS = {
    price:  { bg: "#f0fdf4", border: "#4ade80", tag: "#16a34a", label: "价格更新" },
    remark: { bg: "#fff7ed", border: "#fb923c", tag: "#ea580c", label: "备注" },
    mixed:  { bg: "#fefce8", border: "#facc15", tag: "#ca8a04", label: "价格+备注" },
    none:   { bg: "#f9fafb", border: "#d1d5db", tag: "#9ca3af", label: "无需操作" }
  }

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
      if (file) processFile(file)
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

    // Header
    frag.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;padding:14px 16px;border-bottom:1px solid #f3f4f6;user-select:none"
    }, [
      el("div", { style: "display:flex;align-items:center;gap:8px" }, [
        el("span", { style: "font-size:11px;color:#6b7280;background:#f3f4f6;padding:2px 6px;border-radius:4px" }, "采购结果"),
        el("span", { style: "font-size:15px;font-weight:600;color:#111827" }, "Excel 导入")
      ]),
      el("button", {
        onclick: function () { closeCard(document.getElementById(PREFIX + "btn")) },
        style: "width:26px;height:26px;border:1px solid #e5e7eb;border-radius:6px;background:#fff;color:#6b7280;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center"
      }, "✕")
    ]))

    // 上传区 (始终显示)
    frag.appendChild(renderUploadZone())

    // 标签条 (有 tabs 才显示)
    if (appState.tabs.length > 0) {
      frag.appendChild(el("div", { style: "padding:0 16px;margin-top:2px" }, [
        el("div", { style: "font-size:11px;color:#9ca3af;margin-bottom:6px;text-transform:uppercase;letter-spacing:0.5px" }, "预览窗口"),
        renderTabBar()
      ]))
    }

    // 日志区 (有日志才显示)
    var log = getLog()
    if (log) {
      frag.appendChild(el("div", { style: "padding:0 16px;margin-top:2px" }, [
        el("div", { style: "font-size:11px;color:#9ca3af;margin-bottom:6px;text-transform:uppercase;letter-spacing:0.5px" }, "最近日志"),
        renderLogSection(log)
      ]))
    }

    return frag
  }

  // ═══════════════════════════════════════════
  //  3. 上传区
  // ═══════════════════════════════════════════
  function renderUploadZone() {
    var zone = el("div", {
      style: "margin:12px 16px;padding:20px;border:2px dashed #d1d5db;border-radius:10px;text-align:center;cursor:pointer;transition:border-color 0.2s,background 0.2s"
    })

    zone.addEventListener("click", function () { pickFile() })
    zone.addEventListener("dragover", function (e) { e.preventDefault(); zone.style.borderColor = "#7c3aed"; zone.style.background = "#f5f3ff" })
    zone.addEventListener("dragleave", function () { zone.style.borderColor = "#d1d5db"; zone.style.background = "" })
    zone.addEventListener("drop", function (e) {
      e.preventDefault()
      zone.style.borderColor = "#d1d5db"; zone.style.background = ""
      var file = e.dataTransfer.files[0]
      if (file) processFile(file)
    })

    zone.appendChild(el("div", { style: "font-size:28px;margin-bottom:6px" }, "📁"))
    zone.appendChild(el("div", { style: "font-size:13px;color:#374151;font-weight:500" }, "点击上传或拖拽 Excel"))
    zone.appendChild(el("div", { style: "font-size:11px;color:#9ca3af;margin-top:4px" }, "支持 .xlsx / .xls"))

    return zone
  }

  function pickFile() {
    var input = document.createElement("input")
    input.type = "file"; input.accept = ".xlsx,.xls"
    input.addEventListener("change", function () {
      var file = input.files ? input.files[0] : null
      if (file) processFile(file)
      input.remove()
    })
    input.click()
  }

  // ═══════════════════════════════════════════
  //  4. 标签条
  // ═══════════════════════════════════════════
  function renderTabBar() {
    var bar = el("div", { style: "display:flex;flex-wrap:wrap;gap:6px" })
    for (var i = 0; i < appState.tabs.length; i++) {
      (function (idx) {
        var tab = appState.tabs[idx]
        var tabEl = el("div", {
          style: "display:flex;align-items:center;gap:6px;padding:6px 10px;border-radius:8px;background:#f3f4f6;font-size:12px;color:#374151;cursor:pointer;max-width:100%;transition:background 0.15s",
          onclick: function () { restoreModal(tab.id) },
          onmouseenter: function () { tabEl.style.background = "#e5e7eb" },
          onmouseleave: function () { tabEl.style.background = "#f3f4f6" }
        }, [
          el("span", {}, "📊"),
          el("span", {
            style: "overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
          }, tab.label || tab.fileName || "预览")
        ])
        bar.appendChild(tabEl)
      })(i)
    }
    return bar
  }

  function addTab(fileName, previewRows) {
    var id = "tab_" + Date.now()
    appState.tabs.push({ id: id, label: fileName, fileName: fileName, previewRows: previewRows })
    refreshCard()
    return id
  }

  function removeTab(id) {
    appState.tabs = appState.tabs.filter(function (t) { return t.id !== id })
    refreshCard()
  }

  function findTab(id) {
    return appState.tabs.find(function (t) { return t.id === id }) || null
  }

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
  //  6. Preview Modal (居中 + 可最小化)
  // ═══════════════════════════════════════════
  var currentModalOverlay = null
  var currentTabId = null

  function buildPreviewModal(previewRows, fileName, tabId) {
    removeModals()
    currentTabId = tabId

    var overlay = el("div", {
      id: PREFIX + "modal_overlay",
      style: "position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:100010;display:flex;align-items:center;justify-content:center"
    })
    overlay.addEventListener("click", function (e) { if (e.target === overlay) minimizeModal() })

    var modal = el("div", {
      style: "background:#fff;border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,0.2);width:92vw;max-width:1100px;max-height:85vh;display:flex;flex-direction:column;overflow:hidden"
    })
    overlay.appendChild(modal)

    // Header
    modal.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-bottom:1px solid #e5e7eb"
    }, [
      el("h2", { style: "margin:0;font-size:16px;color:#111827" }, "📋 " + (fileName || "导入预览")),
      el("div", { style: "display:flex;gap:6px" }, [
        el("button", {
          onclick: minimizeModal,
          title: "最小化",
          style: "width:28px;height:28px;border:1px solid #e5e7eb;border-radius:6px;background:#fff;color:#6b7280;font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center;line-height:1"
        }, "—"),
        el("button", {
          onclick: function () { closeModal(tabId) },
          title: "关闭",
          style: "width:28px;height:28px;border:1px solid #e5e7eb;border-radius:6px;background:#fff;color:#6b7280;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center"
        }, "✕")
      ])
    ]))

    // 统计栏
    var actionable = previewRows.filter(function (r) { return r.hasAction && !r.error }).length
    var errors = previewRows.filter(function (r) { return r.error }).length
    var none = previewRows.length - actionable - errors
    modal.appendChild(el("div", {
      style: "padding:10px 20px;background:#f0f9ff;font-size:12px;color:#1e40af"
    }, "共 " + previewRows.length + " 行  ·  " + actionable + " 行可更新  ·  " + none + " 行无需操作" + (errors > 0 ? "  ·  " + errors + " 行匹配失败" : "")))

    // 表格
    var tableWrap = el("div", { style: "flex:1;overflow-y:auto;max-height:50vh" })
    tableWrap.appendChild(renderPreviewTable(previewRows))
    modal.appendChild(tableWrap)

    // Footer
    var checked = previewRows.filter(function (r) { return r.checked && r.hasAction && !r.error }).length
    modal.appendChild(el("div", {
      style: "display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-top:1px solid #e5e7eb"
    }, [
      el("span", { style: "font-size:12px;color:#6b7280" }, "勾选 " + checked + " / " + actionable + " 行"),
      el("div", { style: "display:flex;gap:8px" }, [
        el("button", {
          onclick: minimizeModal,
          style: "padding:8px 16px;border:1px solid #d1d5db;border-radius:8px;background:#fff;color:#6b7280;font-size:13px;cursor:pointer"
        }, "最小化"),
        el("button", {
          id: PREFIX + "confirm_btn",
          disabled: checked === 0,
          onclick: function () { executeUpdate(previewRows, fileName, tabId) },
          style: "padding:8px 20px;border:none;border-radius:8px;background:" + (checked === 0 ? "#d1d5db" : "#7c3aed") + ";color:#fff;font-size:13px;font-weight:600;cursor:" + (checked === 0 ? "default" : "pointer")
        }, "✅ 确认更新 (" + checked + " 行)")
      ])
    ]))

    document.body.appendChild(overlay)
    currentModalOverlay = overlay
  }

  function renderPreviewTable(previewRows) {
    var table = el("table", {
      style: "width:100%;border-collapse:collapse;font-size:12px"
    })

    // thead
    var thead = document.createElement("thead")
    var tr = document.createElement("tr")
    tr.style.cssText = "background:#f9fafb;position:sticky;top:0;z-index:2"
    var headers = [
      { w: "38px", html: '<input type="checkbox" id="' + PREFIX + 'select_all" checked style="cursor:pointer">' },
      { w: "90px", text: "询价单" },
      { w: "130px", text: "UPC" },
      { text: "产品名" },
      { w: "70px", text: "操作类型" },
      { w: "180px", text: "变更详情" }
    ]
    for (var i = 0; i < headers.length; i++) {
      var th = document.createElement("th")
      th.style.cssText = "padding:8px 10px;text-align:left;font-weight:600;color:#374151;border-bottom:2px solid #e5e7eb;font-size:11px;white-space:nowrap"
      if (headers[i].w) th.style.width = headers[i].w
      if (headers[i].html) th.innerHTML = headers[i].html
      else th.textContent = headers[i].text
      tr.appendChild(th)
    }
    thead.appendChild(tr)
    table.appendChild(thead)

    // tbody
    var tbody = document.createElement("tbody")
    for (var i = 0; i < previewRows.length; i++) {
      tbody.appendChild(renderPreviewRow(previewRows[i], previewRows))
    }
    table.appendChild(tbody)

    // 全选事件
    setTimeout(function () {
      var sa = document.getElementById(PREFIX + "select_all")
      if (sa) {
        sa.addEventListener("change", function () {
          var cbs = document.querySelectorAll("." + PREFIX + "row_cb:not([disabled])")
          for (var i = 0; i < cbs.length; i++) {
            cbs[i].checked = sa.checked
          }
          for (var i = 0; i < previewRows.length; i++) {
            if (previewRows[i].hasAction && !previewRows[i].error) {
              previewRows[i].checked = sa.checked
            }
          }
          updateConfirmBtn(previewRows)
        })
      }
    }, 10)

    table.__previewRows = previewRows
    return table
  }

  function renderPreviewRow(row, allRows) {
    var tr = document.createElement("tr")
    var c = ROW_COLORS[row.rowType] || ROW_COLORS.none
    tr.style.cssText = "border-bottom:1px solid " + c.border + ";background:" + c.bg

    // 复选框
    var tdCb = document.createElement("td")
    tdCb.style.cssText = "padding:6px 10px;text-align:center"
    if (row.hasAction && !row.error) {
      var cb = document.createElement("input")
      cb.type = "checkbox"; cb.className = PREFIX + "row_cb"
      cb.checked = row.checked; cb.style.cursor = "pointer"
      cb.addEventListener("change", function () {
        row.checked = cb.checked
        updateConfirmBtn(allRows)
      })
      tdCb.appendChild(cb)
    } else {
      tdCb.innerHTML = '<span style="color:#d1d5db">—</span>'
    }
    tr.appendChild(tdCb)

    tr.appendChild(td(row.orderRef, "#374151", "600"))
    tr.appendChild(td(row.upc, "#6b7280", null, "monospace;font-size:11px"))
    tr.appendChild(td(row.odooLineName || row.productName || "—", row.error ? "#ef4444" : "#374151"))

    // 操作标签
    var typeTd = document.createElement("td")
    typeTd.style.cssText = "padding:6px 10px"
    if (row.error) {
      typeTd.innerHTML = '<span style="display:inline-block;padding:2px 6px;border-radius:3px;font-size:10px;background:#fef2f2;color:#dc2626">匹配失败</span>'
    } else {
      typeTd.innerHTML = '<span style="display:inline-block;padding:2px 6px;border-radius:3px;font-size:10px;background:' + c.bg + ';color:' + c.tag + ';border:1px solid ' + c.border + '">' + c.label + '</span>'
    }
    tr.appendChild(typeTd)

    // 变更详情
    var detailTd = document.createElement("td")
    detailTd.style.cssText = "padding:6px 10px;font-size:11px"
    if (row.error) {
      detailTd.innerHTML = '<span style="color:#ef4444">' + escHtml(row.error) + '</span>'
    } else if (row.changes.length === 0) {
      detailTd.innerHTML = '<span style="color:#9ca3af">—</span>'
    } else {
      var parts = []
      for (var j = 0; j < row.changes.length; j++) {
        var ch = row.changes[j]
        parts.push(
          '<div style="margin-bottom:1px"><span style="color:#6b7280">' + escHtml(ch.label) + ':</span> ' +
          '<span style="color:#9ca3af;text-decoration:line-through">' + escHtml(ch.oldValue) + '</span> ' +
          '<span style="color:#059669;font-weight:600">→ ' + escHtml(ch.newValue) + '</span></div>'
        )
      }
      detailTd.innerHTML = parts.join("")
    }
    tr.appendChild(detailTd)

    return tr
  }

  function td(text, color, weight, extraStyle) {
    var t = document.createElement("td")
    var s = "padding:6px 10px;color:" + color
    if (weight) s += ";font-weight:" + weight
    if (extraStyle) s += ";font-family:" + extraStyle
    t.style.cssText = s
    t.textContent = text
    return t
  }

  function updateConfirmBtn(rows) {
    var btn = document.getElementById(PREFIX + "confirm_btn")
    if (!btn) return
    var checked = rows.filter(function (r) { return r.checked && r.hasAction && !r.error }).length
    btn.disabled = checked === 0
    btn.style.background = checked === 0 ? "#d1d5db" : "#7c3aed"
    btn.style.cursor = checked === 0 ? "default" : "pointer"
    btn.textContent = "✅ 确认更新 (" + checked + " 行)"
  }

  // ═══════════════════════════════════════════
  //  7. 最小化 / 恢复
  // ═══════════════════════════════════════════
  function minimizeModal() {
    if (currentModalOverlay) {
      currentModalOverlay.remove()
      currentModalOverlay = null
    }
    appState.activeModal = null
    if (!appState.cardOpen) {
      var btn = document.getElementById(PREFIX + "btn")
      if (btn) openCard(btn)
    }
  }

  function restoreModal(tabId) {
    var tab = appState.tabs.find(function (t) { return t.id === tabId })
    if (!tab) return
    buildPreviewModal(tab.previewRows, tab.fileName, tabId)
  }

  function closeModal(tabId) {
    if (currentModalOverlay) {
      currentModalOverlay.remove()
      currentModalOverlay = null
    }
    removeTab(tabId)
    appState.activeModal = null
  }

  function removeModals() {
    if (currentModalOverlay) {
      currentModalOverlay.remove()
      currentModalOverlay = null
    }
    var existing = document.getElementById(PREFIX + "modal_overlay")
    if (existing) existing.remove()
  }

  // ═══════════════════════════════════════════
  //  8. 文件处理 → 查询 → 展示 Modal
  // ═══════════════════════════════════════════
  async function processFile(file) {
    // 展开卡片
    if (!appState.cardOpen) {
      var btn = document.getElementById(PREFIX + "btn")
      if (btn) openCard(btn)
    }

    try {
      var data = await file.arrayBuffer()
      var excelRows = parseExcel(data)
      if (excelRows.length === 0) {
        showToast("Excel 中没有有效数据", "error")
        return
      }

      var groups = groupByOrderRef(excelRows)

      // 先创建空标签
      var tabId = addTab(file.name, [])

      // 构建初始预览行（无 odoo 数据）
      var previewRows = excelRows.map(function (r) {
        return buildPreviewRow(r, null)
      })
      var tab = findTab(tabId)
      if (tab) tab.previewRows = previewRows
      refreshCard()

      // 显示 Modal（loading 态）
      buildPreviewModal(previewRows, file.name, tabId)

      // 逐个 PO 查询
      var groupEntries = Array.from(groups.entries())
      var updatedPreviews = []

      for (var gi = 0; gi < groupEntries.length; gi++) {
        var orderRef = groupEntries[gi][0]
        var groupRows = groupEntries[gi][1]

        updateLoadingOverlay("查询 " + orderRef + "（" + groupRows.length + " 行）...")
        await sleep(150)

        var po = await searchPoByName(orderRef)
        var lineMap = new Map()

        if (po && po.orderLineIds.length > 0) {
          var odooLines = await getOrderLines(po.orderLineIds)
          for (var i = 0; i < odooLines.length; i++) {
            if (odooLines[i].upc) lineMap.set(odooLines[i].upc, odooLines[i])
          }
        }

        for (var i = 0; i < groupRows.length; i++) {
          var excelRow = groupRows[i]
          var odooLine = po ? (lineMap.get(excelRow.upc) || null) : null
          var pr = buildPreviewRow(excelRow, odooLine)
          if (!po && !pr.error) pr.error = "未找到询价单 " + orderRef
          updatedPreviews.push(pr)
        }

        // 更新 Modal 表格
        var overlay = document.getElementById(PREFIX + "modal_overlay")
        if (overlay) {
          var tableWrap = overlay.querySelector("div[style*='overflow-y:auto']")
          if (tableWrap) {
            tableWrap.innerHTML = ""
            tableWrap.appendChild(renderPreviewTable(updatedPreviews))
          }
          // 更新统计
          var statsDiv = overlay.querySelector("div[style*='background:#f0f9ff']")
          if (statsDiv) {
            var act = updatedPreviews.filter(function (r) { return r.hasAction && !r.error }).length
            var errs = updatedPreviews.filter(function (r) { return r.error }).length
            var none2 = updatedPreviews.length - act - errs
            statsDiv.textContent = "共 " + updatedPreviews.length + " 行  ·  " + act + " 行可更新  ·  " + none2 + " 行无需操作" + (errs > 0 ? "  ·  " + errs + " 行匹配失败" : "")
          }
          updateConfirmBtn(updatedPreviews)
        }
      }

      updateLoadingOverlay(null)

      // 更新 tab
      var tab2 = findTab(tabId)
      if (tab2) tab2.previewRows = updatedPreviews
      refreshCard()
    } catch (err) {
      showToast("导入失败: " + (err.message || err), "error")
      console.error("[Odoo Excel Importer]", err)
    }
  }

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
  //  9. 执行更新
  // ═══════════════════════════════════════════
  async function executeUpdate(previewRows, fileName, tabId) {
    var selected = previewRows.filter(function (r) { return r.checked && r.hasAction && !r.error })
    if (selected.length === 0) return

    removeModals()
    removeTab(tabId)

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
      for (var j = 0; j < row.changes.length; j++) {
        var ch = row.changes[j]
        if (ch.type === "remark") payload.remark = ch.newValue === "(空)" ? "" : ch.newValue
        else if (ch.type === "price_unit") payload.price_unit = parseFloat(ch.newValue) || 0
        else if (ch.type === "box_wholesale_price") payload.box_wholesale_price = parseFloat(ch.newValue) || 0
      }
      try {
        await updateOrderLine(row.odooLineId, payload)
        logEntry.results.push({ upc: row.upc, orderRef: row.orderRef, productName: row.productName, status: "success", changes: row.changes })
      } catch (err) {
        logEntry.results.push({ upc: row.upc, orderRef: row.orderRef, productName: row.productName, status: "failed", error: err.message || "写入失败", changes: row.changes })
      }
    }

    // 跳过的行
    var skipped = previewRows.filter(function (r) { return !r.checked || !r.hasAction || r.error })
    for (var i = 0; i < skipped.length; i++) {
      logEntry.results.push({ upc: skipped[i].upc, orderRef: skipped[i].orderRef, productName: skipped[i].productName, status: "skipped", reason: skipped[i].error || "用户跳过或无需操作" })
    }

    saveLog(logEntry)
    refreshCard()
    showResultToast(logEntry)
  }

  function showResultToast(logEntry) {
    var success = logEntry.results.filter(function (r) { return r.status === "success" }).length
    var failed = logEntry.results.filter(function (r) { return r.status === "failed" }).length
    var msg = "完成: " + success + " 行已更新"
    if (failed > 0) msg += "，" + failed + " 行失败"
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

  if (isPurchaseOrderPage()) inject()

  window.addEventListener("hashchange", function () {
    var existing = document.getElementById(PREFIX + "btn")
    if (isPurchaseOrderPage()) {
      if (!existing) inject()
    } else {
      if (existing) existing.remove()
      if (cardEl) { cardEl.remove(); cardEl = null }
      removeModals()
      appState.cardOpen = false
    }
  })
})()

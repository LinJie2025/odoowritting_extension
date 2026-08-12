/**
 * Odoo Excel Importer - Content Script v2.0
 *
 * 上传 Excel → 预览 Modal(颜色区分+逐行勾选) → 确认写入 → 日志下载
 */

(function () {
  "use strict"

  // ═══════════════════════════════════════════
  //  1. 页面检测
  // ═══════════════════════════════════════════
  function isPurchaseOrderPage() {
    return window.location.hash.includes("model=purchase.order")
  }

  // ═══════════════════════════════════════════
  //  2. Odoo JSONRPC
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
  //  3. Excel 解析
  // ═══════════════════════════════════════════
  function parseExcel(data) {
    var wb = XLSX.read(data, { type: "array" })
    var ws = wb.Sheets[wb.SheetNames[0]]
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1 })
    if (rows.length < 2) throw new Error("Excel 文件为空或格式不正确")

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

  function parseFloatNum(val) {
    if (val === null || val === undefined || val === "") return null
    var n = Number(val)
    return isNaN(n) ? null : n
  }

  // ═══════════════════════════════════════════
  //  4. 业务逻辑
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

  /**
   * 生成单行预览数据
   */
  function buildPreviewRow(excelRow, odooLine) {
    var remark = determineRemark(excelRow.purchaseResult)
    var hasPrice = excelRow.unitPrice !== null || excelRow.boxWholesalePrice !== null
    var hasRemark = remark !== null
    var hasAction = hasPrice || hasRemark

    // 构建变更列表
    var changes = []
    if (hasRemark) {
      changes.push({
        type: "remark",
        label: "备注",
        oldValue: (odooLine ? (odooLine.remark || "(空)") : "?"),
        newValue: remark
      })
    }
    if (excelRow.unitPrice !== null) {
      changes.push({
        type: "price_unit",
        label: "单价",
        oldValue: odooLine ? formatPrice(odooLine.price_unit) : "?",
        newValue: formatPrice(excelRow.unitPrice)
      })
    }
    if (excelRow.boxWholesalePrice !== null) {
      changes.push({
        type: "box_wholesale_price",
        label: "整箱批发价",
        oldValue: odooLine ? formatPrice(odooLine.box_wholesale_price) : "?",
        newValue: formatPrice(excelRow.boxWholesalePrice)
      })
    }

    // 行颜色：绿色=价格更新，橙色=备注，灰色=无操作
    var rowType = "none"
    if (hasPrice && hasRemark) rowType = "mixed"
    else if (hasPrice) rowType = "price"
    else if (hasRemark) rowType = "remark"

    return {
      upc: excelRow.upc,
      productName: excelRow.productName,
      orderRef: excelRow.orderRef,
      purchaseResult: excelRow.purchaseResult,
      odooLineId: odooLine ? odooLine.id : null,
      odooLineName: odooLine ? odooLine.name : null,
      rowType: rowType,
      hasAction: hasAction,
      changes: changes,
      checked: hasAction,  // 默认勾选有操作的行
      error: odooLine ? null : "未找到匹配的订单行"
    }
  }

  function formatPrice(val) {
    if (val === null || val === undefined || val === 0) return "0"
    return String(val)
  }

  // ═══════════════════════════════════════════
  //  5. Modal UI 构建
  // ═══════════════════════════════════════════
  var ROW_COLORS = {
    price:   { bg: "#f0fdf4", border: "#4ade80", tag: "#16a34a", label: "价格更新" },
    remark:  { bg: "#fff7ed", border: "#fb923c", tag: "#ea580c", label: "备注" },
    mixed:   { bg: "#fefce8", border: "#facc15", tag: "#ca8a04", label: "价格+备注" },
    none:    { bg: "#f9fafb", border: "#d1d5db", tag: "#9ca3af", label: "无需操作" }
  }

  function buildPreviewModal(previewRows, fileName) {
    removeExistingModals()

    var overlay = createOverlay("__odoo_preview_modal")
    var modal = createModalContainer()
    overlay.appendChild(modal)

    // Header
    var header = el("div", { style: "padding:20px 24px 16px;border-bottom:1px solid #e5e7eb" }, [
      el("div", { style: "display:flex;justify-content:space-between;align-items:center" }, [
        el("h2", { style: "margin:0;font-size:18px;color:#111827" }, "📋 Excel 导入预览"),
        el("div", { style: "display:flex;gap:8px" }, [
          el("span", { style: "font-size:12px;color:#6b7280;line-height:28px" }, fileName || ""),
          createCloseButton(overlay)
        ])
      ])
    ])
    modal.appendChild(header)

    // 统计栏
    var statsDiv = el("div", { id: "__odoo_stats", style: "padding:12px 24px;background:#f0f9ff;font-size:13px;color:#1e40af" })
    modal.appendChild(statsDiv)

    // 表格容器
    var tableWrap = el("div", {
      style: "max-height:50vh;overflow-y:auto;border-top:1px solid #e5e7eb"
    })
    var table = el("table", {
      style: "width:100%;border-collapse:collapse;font-size:13px"
    })
    tableWrap.appendChild(table)
    modal.appendChild(tableWrap)

    // Footer
    var footer = el("div", {
      style: "padding:16px 24px;border-top:1px solid #e5e7eb;display:flex;justify-content:space-between;align-items:center"
    }, [
      el("div", { style: "display:flex;gap:8px" }, [
        el("button", {
          id: "__odoo_confirm_btn",
          style: "padding:10px 24px;background:#7c3aed;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;opacity:0.5",
          disabled: "true"
        }, "⏳ 查询中..."),
        el("button", {
          onclick: function () { overlay.remove() },
          style: "padding:10px 16px;background:#fff;color:#6b7280;border:1px solid #d1d5db;border-radius:8px;font-size:14px;cursor:pointer"
        }, "取消")
      ]),
      el("span", { id: "__odoo_footer_info", style: "font-size:12px;color:#9ca3af" })
    ])
    modal.appendChild(footer)

    document.body.appendChild(overlay)

    // 渲染表头和行
    renderTableHeader(table)
    renderRows(table, previewRows)
    updateStats(previewRows)
    updateConfirmButton(previewRows)

    return { overlay: overlay, table: table }
  }

  function removeExistingModals() {
    var ids = ["__odoo_preview_modal", "__odoo_result_modal"]
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i])
      if (el) el.remove()
    }
  }

  function createOverlay(id) {
    var overlay = el("div", {
      id: id,
      style: "position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:100000;display:flex;align-items:center;justify-content:center"
    })
    overlay.addEventListener("click", function (e) {
      if (e.target === overlay) overlay.remove()
    })
    return overlay
  }

  function createModalContainer() {
    return el("div", {
      style: "background:#fff;border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,0.2);width:92vw;max-width:1100px;max-height:90vh;display:flex;flex-direction:column;overflow:hidden"
    })
  }

  function createCloseButton(overlay) {
    return el("button", {
      onclick: function () { overlay.remove() },
      style: "width:28px;height:28px;border:1px solid #e5e7eb;border-radius:6px;background:#fff;color:#6b7280;font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center"
    }, "✕")
  }

  function el(tag, attrs, children) {
    var e = document.createElement(tag)
    if (attrs) {
      for (var k in attrs) {
        if (k === "style" && typeof attrs[k] === "string") {
          e.setAttribute("style", attrs[k])
        } else if (k === "disabled") {
          // skip, handle separately
        } else if (k.startsWith("on")) {
          e.addEventListener(k.slice(2).toLowerCase(), attrs[k])
        } else {
          e.setAttribute(k, attrs[k])
        }
      }
    }
    if (attrs && attrs.disabled) e.disabled = true
    if (children) {
      if (typeof children === "string") {
        e.textContent = children
      } else if (Array.isArray(children)) {
        for (var i = 0; i < children.length; i++) {
          if (typeof children[i] === "string") {
            e.appendChild(document.createTextNode(children[i]))
          } else if (children[i]) {
            e.appendChild(children[i])
          }
        }
      } else if (children instanceof Node) {
        e.appendChild(children)
      }
    }
    return e
  }

  // ═══════════════════════════════════════════
  //  6. 表格渲染
  // ═══════════════════════════════════════════
  function renderTableHeader(table) {
    var thead = document.createElement("thead")
    var tr = document.createElement("tr")
    tr.style.cssText = "background:#f9fafb;position:sticky;top:0;z-index:2"

    var headers = [
      { w: "40px", html: '<input type="checkbox" id="__odoo_select_all" checked style="cursor:pointer">' },
      { w: "100px", text: "询价单" },
      { w: "140px", text: "UPC" },
      { w: "auto", text: "产品名" },
      { w: "80px", text: "加购结果" },
      { w: "80px", text: "操作类型" },
      { w: "180px", text: "变更详情" }
    ]

    for (var i = 0; i < headers.length; i++) {
      var th = document.createElement("th")
      th.style.cssText = "padding:10px 12px;text-align:left;font-weight:600;color:#374151;border-bottom:2px solid #e5e7eb;font-size:12px;white-space:nowrap"
      if (headers[i].w !== "auto") th.style.width = headers[i].w
      if (headers[i].html) {
        th.innerHTML = headers[i].html
      } else {
        th.textContent = headers[i].text
      }
      tr.appendChild(th)
    }
    thead.appendChild(tr)
    table.appendChild(thead)

    // 全选事件
    setTimeout(function () {
      var selectAll = document.getElementById("__odoo_select_all")
      if (selectAll) {
        selectAll.addEventListener("change", function () {
          var cbs = document.querySelectorAll(".__odoo_row_cb:not([disabled])")
          for (var i = 0; i < cbs.length; i++) { cbs[i].checked = selectAll.checked }
          updateConfirmButton(getAllPreviewRows())
        })
      }
    }, 10)
  }

  function renderRows(table, previewRows) {
    var tbody = document.getElementById("__odoo_preview_tbody")
    if (!tbody) {
      tbody = document.createElement("tbody")
      tbody.id = "__odoo_preview_tbody"
      table.appendChild(tbody)
    }

    // 先渲染已有行
    for (var i = 0; i < previewRows.length; i++) {
      appendRow(tbody, previewRows[i])
    }

    // 存储引用
    table.__previewRows = previewRows
  }

  function appendRow(tbody, row) {
    var tr = document.createElement("tr")
    var c = ROW_COLORS[row.rowType] || ROW_COLORS.none
    tr.style.cssText = "border-bottom:1px solid " + c.border + ";background:" + c.bg + ";transition:background 0.2s"
    tr.setAttribute("data-upc", row.upc)
    tr.setAttribute("data-orderref", row.orderRef)

    // 复选框
    var tdCb = document.createElement("td")
    tdCb.style.cssText = "padding:8px 12px;text-align:center"
    if (row.hasAction && !row.error) {
      var cb = document.createElement("input")
      cb.type = "checkbox"
      cb.className = "__odoo_row_cb"
      cb.checked = row.checked
      cb.style.cursor = "pointer"
      cb.addEventListener("change", function () {
        row.checked = cb.checked
        updateConfirmButton(getAllPreviewRows())
      })
      tdCb.appendChild(cb)
    } else {
      tdCb.innerHTML = '<span style="color:#d1d5db">—</span>'
    }
    tr.appendChild(tdCb)

    // 询价单
    tr.appendChild(cell(row.orderRef, { color: "#374151", weight: "600" }))
    // UPC
    tr.appendChild(cell(row.upc, { color: "#6b7280", mono: true }))
    // 产品名
    tr.appendChild(cell(row.odooLineName || row.productName || "—", { color: row.error ? "#ef4444" : "#374151" }))
    // 加购结果
    var resultText = row.purchaseResult
    var resultColor = "#6b7280"
    if (resultText.indexOf("缺货") !== -1) resultColor = "#ef4444"
    else if (resultText.indexOf("成功") !== -1) resultColor = "#16a34a"
    tr.appendChild(cell(resultText || "—", { color: resultColor, size: "11px" }))

    // 操作类型标签
    var tdType = document.createElement("td")
    tdType.style.cssText = "padding:8px 12px"
    if (row.error) {
      tdType.innerHTML = '<span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;background:#fef2f2;color:#dc2626">匹配失败</span>'
    } else {
      tdType.innerHTML = '<span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;background:' + c.bg + ';color:' + c.tag + ';border:1px solid ' + c.border + '">' + c.label + '</span>'
    }
    tr.appendChild(tdType)

    // 变更详情
    var tdDetail = document.createElement("td")
    tdDetail.style.cssText = "padding:8px 12px;font-size:11px"
    if (row.error) {
      tdDetail.innerHTML = '<span style="color:#ef4444">' + escHtml(row.error) + '</span>'
    } else if (row.changes.length === 0) {
      tdDetail.innerHTML = '<span style="color:#9ca3af">—</span>'
    } else {
      var parts = []
      for (var i = 0; i < row.changes.length; i++) {
        var ch = row.changes[i]
        parts.push(
          '<div style="margin-bottom:2px">' +
          '<span style="color:#6b7280">' + escHtml(ch.label) + ':</span> ' +
          '<span style="color:#9ca3af;text-decoration:line-through">' + escHtml(ch.oldValue) + '</span> ' +
          '<span style="color:#059669;font-weight:600">→ ' + escHtml(ch.newValue) + '</span>' +
          '</div>'
        )
      }
      tdDetail.innerHTML = parts.join("")
    }
    tr.appendChild(tdDetail)

    tbody.appendChild(tr)
  }

  function cell(text, opts) {
    var td = document.createElement("td")
    var style = "padding:8px 12px"
    if (opts.color) style += ";color:" + opts.color
    if (opts.weight) style += ";font-weight:" + opts.weight
    if (opts.size) style += ";font-size:" + opts.size
    if (opts.mono) style += ";font-family:monospace;font-size:12px"
    td.style.cssText = style
    td.textContent = text
    return td
  }

  function escHtml(s) {
    var d = document.createElement("div")
    d.textContent = s
    return d.innerHTML
  }

  function getAllPreviewRows() {
    var table = document.querySelector("#__odoo_preview_tbody")
    if (!table || !table.parentElement) return []
    return table.parentElement.__previewRows || []
  }

  function updateStats(rows) {
    var stats = document.getElementById("__odoo_stats")
    if (!stats) return
    var total = rows.length
    var actionable = rows.filter(function (r) { return r.hasAction && !r.error }).length
    var errors = rows.filter(function (r) { return r.error }).length
    var none = total - actionable - errors
    var parts = ["共 " + total + " 行"]
    if (actionable > 0) parts.push(actionable + " 行可更新")
    if (none > 0) parts.push(none + " 行无需操作")
    if (errors > 0) parts.push(errors + " 行匹配失败")
    stats.textContent = parts.join("  ·  ")
  }

  function updateConfirmButton(rows) {
    var btn = document.getElementById("__odoo_confirm_btn")
    var info = document.getElementById("__odoo_footer_info")
    if (!btn) return

    var checked = rows.filter(function (r) { return r.checked && r.hasAction && !r.error }).length
    btn.disabled = checked === 0
    btn.style.opacity = checked === 0 ? "0.5" : "1"
    btn.textContent = "✅ 确认更新 (" + checked + " 行)"
    if (info) {
      info.textContent = "共 " + rows.length + " 行，勾选 " + checked + " 行"
    }
  }

  // ═══════════════════════════════════════════
  //  7. 结果 Modal
  // ═══════════════════════════════════════════
  function showResultModal(logEntry) {
    removeExistingModals()

    var overlay = createOverlay("__odoo_result_modal")
    var modal = createModalContainer()
    modal.style.maxWidth = "480px"
    overlay.appendChild(modal)

    var success = logEntry.results.filter(function (r) { return r.status === "success" }).length
    var failed = logEntry.results.filter(function (r) { return r.status === "failed" }).length
    var skipped = logEntry.results.filter(function (r) { return r.status === "skipped" }).length
    var icon = failed > 0 ? "⚠️" : "✅"
    var titleColor = failed > 0 ? "#d97706" : "#059669"

    // Header
    modal.appendChild(el("div", {
      style: "padding:28px 24px 16px;text-align:center"
    }, [
      el("div", { style: "font-size:48px;margin-bottom:8px" }, icon),
      el("h2", { style: "margin:0 0 16px;font-size:20px;color:" + titleColor }, "导入完成"),
      el("div", {
        style: "display:flex;justify-content:center;gap:24px;font-size:14px"
      }, [
        el("div", { style: "text-align:center" }, [el("div", { style: "font-size:28px;font-weight:700;color:#059669" }, String(success)), el("div", { style: "color:#6b7280;margin-top:4px" }, "成功")]),
        el("div", { style: "text-align:center" }, [el("div", { style: "font-size:28px;font-weight:700;color:#9ca3af" }, String(skipped)), el("div", { style: "color:#6b7280;margin-top:4px" }, "跳过")]),
        el("div", { style: "text-align:center" }, [el("div", { style: "font-size:28px;font-weight:700;color:#ef4444" }, String(failed)), el("div", { style: "color:#6b7280;margin-top:4px" }, "失败")])
      ])
    ]))

    // 失败明细
    if (failed > 0) {
      var failRows = logEntry.results.filter(function (r) { return r.status === "failed" })
      var detailDiv = el("div", {
        style: "margin:0 24px 16px;padding:12px;background:#fef2f2;border-radius:8px;max-height:150px;overflow-y:auto"
      })
      for (var i = 0; i < failRows.length; i++) {
        detailDiv.appendChild(el("div", {
          style: "font-size:12px;color:#dc2626;margin-bottom:4px;font-family:monospace"
        }, failRows[i].upc + ": " + (failRows[i].error || "写入失败")))
      }
      modal.appendChild(detailDiv)
    }

    // Buttons
    modal.appendChild(el("div", {
      style: "padding:16px 24px 24px;display:flex;justify-content:center;gap:12px;border-top:1px solid #f3f4f6"
    }, [
      el("button", {
        onclick: function () { downloadLog(logEntry) },
        style: "padding:10px 20px;background:#fff;color:#374151;border:1px solid #d1d5db;border-radius:8px;font-size:14px;cursor:pointer"
      }, "📥 下载日志"),
      el("button", {
        onclick: function () { overlay.remove() },
        style: "padding:10px 24px;background:#7c3aed;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer"
      }, "关闭")
    ]))

    document.body.appendChild(overlay)
  }

  // ═══════════════════════════════════════════
  //  8. 日志 (localStorage + JSON 下载)
  // ═══════════════════════════════════════════
  var LOG_KEY = "__odoo_import_log"

  function saveLog(logEntry) {
    try {
      localStorage.setItem(LOG_KEY, JSON.stringify(logEntry))
    } catch (e) {
      console.warn("[Odoo Excel Importer] 日志保存失败:", e)
    }
  }

  function getLog() {
    try {
      var raw = localStorage.getItem(LOG_KEY)
      return raw ? JSON.parse(raw) : null
    } catch (e) {
      return null
    }
  }

  function downloadLog(logEntry) {
    var json = JSON.stringify(logEntry, null, 2)
    var blob = new Blob([json], { type: "application/json" })
    var url = URL.createObjectURL(blob)
    var a = document.createElement("a")
    a.href = url
    a.download = "odoo-import-log-" + logEntry.timestamp.replace(/[:.]/g, "-") + ".json"
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  // ═══════════════════════════════════════════
  //  9. 浮动按钮
  // ═══════════════════════════════════════════
  function createFloatingButton() {
    var btn = document.createElement("button")
    btn.id = "__odoo_import_btn"
    btn.textContent = "📥 导入Excel"
    btn.style.cssText = "position:fixed;bottom:20px;right:20px;z-index:99999;padding:12px 24px;border-radius:28px;border:none;background-color:#7c3aed;color:#fff;font-size:15px;font-weight:600;font-family:sans-serif;cursor:pointer;box-shadow:0 4px 16px rgba(124,58,237,0.4);transition:transform 0.15s,box-shadow 0.15s;outline:none"
    btn.addEventListener("mouseenter", function () {
      btn.style.transform = "scale(1.05)"
      btn.style.boxShadow = "0 6px 20px rgba(124,58,237,0.5)"
    })
    btn.addEventListener("mouseleave", function () {
      btn.style.transform = "scale(1)"
      btn.style.boxShadow = "0 4px 16px rgba(124,58,237,0.4)"
    })
    btn.addEventListener("click", handleImport)
    return btn
  }

  // ═══════════════════════════════════════════
  //  10. 主导入流程
  // ═══════════════════════════════════════════
  async function handleImport() {
    var input = document.createElement("input")
    input.type = "file"
    input.accept = ".xlsx,.xls"
    input.style.display = "none"
    document.body.appendChild(input)

    input.addEventListener("change", async function () {
      var file = input.files ? input.files[0] : null
      input.remove()
      if (!file) return

      var btn = document.getElementById("__odoo_import_btn")
      if (btn) { btn.textContent = "⏳ 处理中..."; btn.disabled = true }

      try {
        // 1. 解析 Excel
        var data = await file.arrayBuffer()
        var excelRows = parseExcel(data)
        if (excelRows.length === 0) {
          showToast("Excel 中没有有效数据", "error")
          resetButton(btn)
          return
        }

        var groups = groupByOrderRef(excelRows)

        // 2. 先创建 Modal (空表格+loading态)
        var previewRows = excelRows.map(function (r) {
          return buildPreviewRow(r, null)
        })
        var modalState = buildPreviewModal(previewRows, file.name)

        // 3. 逐个 PO 查询，渐进更新表格
        var groupEntries = Array.from(groups.entries())
        var updatedPreviews = []

        for (var gi = 0; gi < groupEntries.length; gi++) {
          var orderRef = groupEntries[gi][0]
          var groupRows = groupEntries[gi][1]

          updateLoadingText(modalState.overlay, "正在查询 " + orderRef + "（" + groupRows.length + " 行）...")
          await sleep(200) // 给 UI 渲染时间

          var po = await searchPoByName(orderRef)
          var odooLines = []
          var lineMap = new Map()

          if (po && po.orderLineIds.length > 0) {
            odooLines = await getOrderLines(po.orderLineIds)
            for (var i = 0; i < odooLines.length; i++) {
              if (odooLines[i].upc) lineMap.set(odooLines[i].upc, odooLines[i])
            }
          }

          // 构建该组的预览行
          for (var i = 0; i < groupRows.length; i++) {
            var excelRow = groupRows[i]
            var odooLine = po ? (lineMap.get(excelRow.upc) || null) : null

            if (!po) {
              excelRow.__poError = "未找到询价单 " + orderRef
            }
            var pr = buildPreviewRow(excelRow, odooLine)
            if (excelRow.__poError && !pr.error) pr.error = excelRow.__poError
            updatedPreviews.push(pr)
          }

          // 重新渲染表格
          var tbody = document.getElementById("__odoo_preview_tbody")
          if (tbody) tbody.innerHTML = ""
          var table = modalState.table
          table.__previewRows = updatedPreviews
          if (tbody) {
            for (var i = 0; i < updatedPreviews.length; i++) {
              appendRow(tbody, updatedPreviews[i])
            }
          }

          updateStats(updatedPreviews)
          updateConfirmButton(updatedPreviews)
        }

        // 查询完成
        updateLoadingText(modalState.overlay, null)

        // 设置确认按钮
        var confirmBtn = document.getElementById("__odoo_confirm_btn")
        if (confirmBtn) {
          confirmBtn.textContent = "✅ 确认更新 (" + updatedPreviews.filter(function (r) { return r.checked && r.hasAction && !r.error }).length + " 行)"
          confirmBtn.style.opacity = "1"
          confirmBtn.disabled = false
          confirmBtn.onclick = function () { executeUpdate(updatedPreviews, file.name, modalState.overlay) }
        }
      } catch (err) {
        showToast("导入失败: " + (err.message || err), "error")
        console.error("[Odoo Excel Importer]", err)
        resetButton(btn)
      }
    })

    input.click()
  }

  function updateLoadingText(overlay, text) {
    var indicator = overlay.querySelector(".__odoo_loading")
    if (text) {
      if (!indicator) {
        indicator = document.createElement("div")
        indicator.className = "__odoo_loading"
        indicator.style.cssText = "position:absolute;top:12px;left:50%;transform:translateX(-50%);background:#7c3aed;color:#fff;padding:8px 20px;border-radius:20px;font-size:13px;z-index:10;box-shadow:0 4px 12px rgba(124,58,237,0.4)"
        overlay.appendChild(indicator)
      }
      indicator.textContent = "⏳ " + text
      indicator.style.display = ""
    } else {
      if (indicator) indicator.style.display = "none"
    }
  }

  async function executeUpdate(previewRows, fileName, previewOverlay) {
    var selectedRows = previewRows.filter(function (r) { return r.checked && r.hasAction && !r.error })

    if (selectedRows.length === 0) {
      showToast("没有勾选任何可更新的行", "error")
      return
    }

    // 关闭预览 Modal
    previewOverlay.remove()

    var btn = document.getElementById("__odoo_import_btn")
    if (btn) { btn.textContent = "⏳ 写入中..."; btn.disabled = true }

    var logEntry = {
      timestamp: new Date().toISOString(),
      fileName: fileName,
      totalRows: previewRows.length,
      selectedRows: selectedRows.length,
      results: []
    }

    for (var i = 0; i < selectedRows.length; i++) {
      var row = selectedRows[i]
      var payload = {}

      for (var j = 0; j < row.changes.length; j++) {
        var ch = row.changes[j]
        if (ch.type === "remark") {
          payload.remark = ch.newValue === "(空)" ? "" : ch.newValue
        } else if (ch.type === "price_unit") {
          payload.price_unit = parseFloat(ch.newValue) || 0
        } else if (ch.type === "box_wholesale_price") {
          payload.box_wholesale_price = parseFloat(ch.newValue) || 0
        }
      }

      try {
        await updateOrderLine(row.odooLineId, payload)
        logEntry.results.push({
          upc: row.upc,
          orderRef: row.orderRef,
          productName: row.productName,
          status: "success",
          changes: row.changes
        })
      } catch (err) {
        logEntry.results.push({
          upc: row.upc,
          orderRef: row.orderRef,
          productName: row.productName,
          status: "failed",
          error: err.message || "写入失败",
          changes: row.changes
        })
      }
    }

    // 跳过的行
    var skipped = previewRows.filter(function (r) { return !r.checked || !r.hasAction || r.error })
    for (var i = 0; i < skipped.length; i++) {
      logEntry.results.push({
        upc: skipped[i].upc,
        orderRef: skipped[i].orderRef,
        productName: skipped[i].productName,
        status: "skipped",
        reason: skipped[i].error || "用户跳过或无需操作"
      })
    }

    // 保存日志
    saveLog(logEntry)

    // 显示结果
    showResultModal(logEntry)
    resetButton(btn)
  }

  function showToast(msg, type) {
    type = type || "info"
    var existing = document.getElementById("__odoo_import_toast")
    if (existing) existing.remove()
    var colors = { info: "#1890ff", success: "#52c41a", error: "#ff4d4f" }
    var toast = document.createElement("div")
    toast.id = "__odoo_import_toast"
    toast.textContent = msg
    toast.style.cssText = "position:fixed;bottom:80px;right:20px;z-index:100001;padding:12px 20px;border-radius:8px;color:#fff;font-size:14px;font-family:sans-serif;background-color:" + (colors[type] || colors.info) + ";box-shadow:0 4px 12px rgba(0,0,0,0.25);transition:opacity 0.3s;opacity:1;max-width:400px;word-break:break-word"
    document.body.appendChild(toast)
    setTimeout(function () { toast.style.opacity = "0"; setTimeout(function () { toast.remove() }, 300) }, 4000)
  }

  function resetButton(btn) {
    if (btn) {
      btn.textContent = "📥 导入Excel"
      btn.disabled = false
    }
  }

  function sleep(ms) {
    return new Promise(function (r) { setTimeout(r, ms) })
  }

  // ═══════════════════════════════════════════
  //  11. 注入
  // ═══════════════════════════════════════════
  function injectButton() {
    if (document.getElementById("__odoo_import_btn")) return
    if (!document.body) { setTimeout(injectButton, 200); return }
    document.body.appendChild(createFloatingButton())
  }

  if (isPurchaseOrderPage()) injectButton()

  window.addEventListener("hashchange", function () {
    var existing = document.getElementById("__odoo_import_btn")
    if (isPurchaseOrderPage()) {
      if (!existing) injectButton()
    } else {
      if (existing) existing.remove()
      removeExistingModals()
    }
  })
})()

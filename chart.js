/* ============================================================================
 *  chart.js — Motor de gráficos de XlsView (ECharts).
 *
 *  - Toma el rango seleccionado en la hoja y construye un gráfico.
 *  - Diálogo dedicado con: tipo de gráfico, título, opciones, tema, exportar.
 *  - Live-update: refresca al cambiar los datos del rango en la hoja.
 *  - Múltiples gráficos guardados por pestaña; persisten al guardar/reabrir
 *    (se serializan dentro del snapshot del libro en __charts).
 *
 *  API expuesta:  window.XlsxChart.init({ getApi, getActiveTab, host, onDirty })
 *    getApi()        -> univerAPI de Univer
 *    getActiveTab()  -> pestaña activa (para leer/guardar sus gráficos)
 *    host            -> objeto con postToHost(msg) y flag inHostApp (para export)
 *    onDirty()       -> marca la pestaña como modificada
 * ==========================================================================*/
(function () {
  "use strict";

  let deps = null;
  let ec = null;                 // instancia ECharts activa en el diálogo
  let liveTimer = null;          // temporizador de live-update
  let current = null;            // gráfico en edición { id, cfg }
  const $ = (id) => document.getElementById(id);

  // -------------------------------------------------------------------------
  //  Configuración por defecto de un gráfico
  // -------------------------------------------------------------------------
  function defaultConfig() {
    return {
      id: "chart-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
      type: "bar",              // bar|line|area|pie|scatter|radar|funnel|gauge|
                                // stacked-bar|combo|heatmap|treemap
      range: null,             // { sheetId, sheetName, startRow, startColumn, endRow, endColumn }
      title: "",
      firstRowIsHeader: true,
      firstColIsCategory: true,
      showLegend: true,
      showLabels: false,
      smooth: false,
      horizontal: false,
      dualAxis: false,
      palette: "vivid",
      name: "Gráfico",
    };
  }

  const PALETTES = {
    vivid: ["#3d7eff", "#12915a", "#e2a03f", "#c0392b", "#8e44ad", "#16a085", "#d35400", "#2c3e50"],
    pastel: ["#93c5fd", "#a7f3d0", "#fde68a", "#fca5a5", "#c4b5fd", "#99f6e4", "#fdba74", "#cbd5e1"],
    mono: ["#1f4e78", "#2e75b6", "#5b9bd5", "#9dc3e6", "#bdd7ee", "#deebf7"],
    warm: ["#c0392b", "#e67e22", "#f39c12", "#f1c40f", "#d35400", "#e74c3c"],
    cool: ["#16a085", "#2980b9", "#8e44ad", "#27ae60", "#2c3e50", "#3498db"],
  };

  // Detecta el tema (claro/oscuro) del documento — la app es oscura.
  function isDark() { return true; }

  // -------------------------------------------------------------------------
  //  Extracción de datos desde la hoja
  // -------------------------------------------------------------------------
  // Lee el rango activo de la hoja y devuelve { range, values } o null.
  function readActiveRange() {
    try {
      const wb = deps.getApi().getActiveWorkbook();
      const ws = wb.getActiveSheet();
      const rng = ws.getActiveRange ? ws.getActiveRange() : null;
      if (!rng) return null;
      const values = rng.getValues();
      const startRow = rng.getRow ? rng.getRow() : 0;
      const startColumn = rng.getColumn ? rng.getColumn() : 0;
      const h = rng.getHeight ? rng.getHeight() : (values ? values.length : 0);
      const w = rng.getWidth ? rng.getWidth() : (values && values[0] ? values[0].length : 0);
      const sheetId = ws.getSheetId ? ws.getSheetId() : null;
      const sheetName = ws.getSheetName ? ws.getSheetName() : "";
      return {
        range: { sheetId, sheetName, startRow, startColumn, endRow: startRow + h - 1, endColumn: startColumn + w - 1 },
        values,
      };
    } catch (e) { return null; }
  }

  // Relee los valores de un rango guardado (para live-update).
  function readRangeValues(range) {
    try {
      const wb = deps.getApi().getActiveWorkbook();
      let ws = null;
      if (range.sheetId && wb.getSheetBySheetId) ws = wb.getSheetBySheetId(range.sheetId);
      if (!ws && range.sheetName && wb.getSheetByName) ws = wb.getSheetByName(range.sheetName);
      if (!ws) ws = wb.getActiveSheet();
      const nRows = range.endRow - range.startRow + 1;
      const nCols = range.endColumn - range.startColumn + 1;
      const rng = ws.getRange(range.startRow, range.startColumn, nRows, nCols);
      return rng.getValues();
    } catch (e) { return null; }
  }

  // Convierte la matriz de valores + config en { categories, series }.
  function shapeData(values, cfg) {
    if (!values || !values.length) return { categories: [], series: [] };
    let rows = values.map((r) => r.slice());
    let headers = null;
    if (cfg.firstRowIsHeader && rows.length > 1) {
      headers = rows[0].map((h) => (h == null ? "" : String(h)));
      rows = rows.slice(1);
    }
    let categories = [];
    let dataCols = rows.length ? rows[0].length : 0;
    let startCol = 0;
    if (cfg.firstColIsCategory && dataCols > 1) {
      categories = rows.map((r) => (r[0] == null ? "" : String(r[0])));
      startCol = 1;
    } else {
      categories = rows.map((_, i) => String(i + 1));
    }
    const series = [];
    for (let c = startCol; c < dataCols; c++) {
      const name = headers ? (headers[c] || ("Serie " + (c - startCol + 1))) : ("Serie " + (c - startCol + 1));
      const data = rows.map((r) => {
        const v = r[c];
        const n = typeof v === "number" ? v : parseFloat(v);
        return isNaN(n) ? null : n;
      });
      series.push({ name, data });
    }
    return { categories, series };
  }

  // -------------------------------------------------------------------------
  //  Construcción de la opción de ECharts según el tipo
  // -------------------------------------------------------------------------
  function buildOption(cfg, shaped) {
    const dark = isDark();
    const textColor = dark ? "#e6e6e6" : "#222";
    const axisColor = dark ? "#555" : "#ccc";
    const splitColor = dark ? "rgba(255,255,255,0.06)" : "rgba(0,0,0,0.06)";
    const color = PALETTES[cfg.palette] || PALETTES.vivid;

    const base = {
      color,
      backgroundColor: "transparent",
      textStyle: { color: textColor, fontFamily: "Segoe UI, sans-serif" },
      title: cfg.title ? { text: cfg.title, left: "center", textStyle: { color: textColor, fontSize: 16 } } : undefined,
      tooltip: { trigger: "item" },
      legend: cfg.showLegend ? { bottom: 0, textStyle: { color: textColor } } : undefined,
      animationDuration: 500,
    };

    const cats = shaped.categories;
    const series = shaped.series;
    const labelOpt = cfg.showLabels ? { show: true, color: textColor } : { show: false };

    const catAxis = { type: "category", data: cats, axisLine: { lineStyle: { color: axisColor } }, axisLabel: { color: textColor } };
    const valAxis = { type: "value", axisLine: { lineStyle: { color: axisColor } }, axisLabel: { color: textColor }, splitLine: { lineStyle: { color: splitColor } } };

    switch (cfg.type) {
      case "bar":
      case "stacked-bar": {
        const stack = cfg.type === "stacked-bar" ? "total" : undefined;
        const opt = Object.assign({}, base, {
          tooltip: { trigger: "axis" },
          xAxis: cfg.horizontal ? valAxis : catAxis,
          yAxis: cfg.horizontal ? catAxis : valAxis,
          series: series.map((s) => ({ name: s.name, type: "bar", stack, data: s.data, label: labelOpt })),
        });
        return opt;
      }
      case "line":
      case "area": {
        return Object.assign({}, base, {
          tooltip: { trigger: "axis" },
          xAxis: catAxis, yAxis: valAxis,
          series: series.map((s) => ({
            name: s.name, type: "line", smooth: cfg.smooth, data: s.data, label: labelOpt,
            areaStyle: cfg.type === "area" ? {} : undefined,
          })),
        });
      }
      case "combo": {
        // primera serie como columnas, resto como líneas; doble eje opcional
        const yAxes = cfg.dualAxis
          ? [valAxis, Object.assign({}, valAxis, { position: "right" })]
          : valAxis;
        return Object.assign({}, base, {
          tooltip: { trigger: "axis" },
          xAxis: catAxis, yAxis: yAxes,
          series: series.map((s, i) => ({
            name: s.name, type: i === 0 ? "bar" : "line", smooth: cfg.smooth,
            yAxisIndex: cfg.dualAxis && i > 0 ? 1 : 0, data: s.data, label: labelOpt,
          })),
        });
      }
      case "pie": {
        const s0 = series[0] || { data: [] };
        return Object.assign({}, base, {
          tooltip: { trigger: "item", formatter: "{b}: {c} ({d}%)" },
          series: [{
            type: "pie", radius: ["40%", "70%"], data: cats.map((c, i) => ({ name: c, value: s0.data[i] })),
            label: { color: textColor, show: cfg.showLabels || true },
          }],
        });
      }
      case "scatter": {
        // pares (x=serie1, y=serie2) si hay 2 series; si no, índice vs valor
        let pts;
        if (series.length >= 2) pts = series[0].data.map((x, i) => [x, series[1].data[i]]);
        else pts = (series[0] ? series[0].data : []).map((y, i) => [i + 1, y]);
        return Object.assign({}, base, {
          xAxis: valAxis, yAxis: valAxis,
          series: [{ type: "scatter", data: pts, symbolSize: 12 }],
        });
      }
      case "radar": {
        const max = Math.max(1, ...series.flatMap((s) => s.data.filter((v) => v != null)));
        return Object.assign({}, base, {
          radar: { indicator: cats.map((c) => ({ name: c, max })), axisName: { color: textColor } },
          series: [{ type: "radar", data: series.map((s) => ({ name: s.name, value: s.data })) }],
        });
      }
      case "funnel": {
        const s0 = series[0] || { data: [] };
        return Object.assign({}, base, {
          series: [{ type: "funnel", data: cats.map((c, i) => ({ name: c, value: s0.data[i] })), label: { color: textColor } }],
        });
      }
      case "gauge": {
        const val = (series[0] && series[0].data.find((v) => v != null)) || 0;
        return Object.assign({}, base, {
          series: [{ type: "gauge", data: [{ value: val, name: cfg.title || (series[0] && series[0].name) || "" }],
            detail: { color: textColor }, axisLabel: { color: textColor } }],
        });
      }
      case "heatmap": {
        const data = [];
        let maxV = 0;
        series.forEach((s, sc) => s.data.forEach((v, ci) => { if (v != null) { data.push([ci, sc, v]); maxV = Math.max(maxV, v); } }));
        return Object.assign({}, base, {
          tooltip: { position: "top" },
          grid: { top: 30, bottom: 50 },
          xAxis: { type: "category", data: cats, axisLabel: { color: textColor } },
          yAxis: { type: "category", data: series.map((s) => s.name), axisLabel: { color: textColor } },
          visualMap: { min: 0, max: maxV || 1, calculable: true, orient: "horizontal", left: "center", bottom: 0, textStyle: { color: textColor } },
          series: [{ type: "heatmap", data, label: { show: cfg.showLabels, color: "#000" } }],
        });
      }
      case "treemap": {
        const s0 = series[0] || { data: [] };
        return Object.assign({}, base, {
          series: [{ type: "treemap", data: cats.map((c, i) => ({ name: c, value: s0.data[i] || 0 })),
            label: { color: "#fff" }, breadcrumb: { show: false } }],
        });
      }
      default:
        return base;
    }
  }

  // -------------------------------------------------------------------------
  //  Renderizado en el diálogo
  // -------------------------------------------------------------------------
  function renderPreview(cfg) {
    const host = $("chartCanvas");
    if (!host) return;
    if (!ec) ec = echarts.init(host, null, { renderer: "canvas" });
    const src = current._values || (cfg.range ? readRangeValues(cfg.range) : null);
    const shaped = shapeData(src, cfg);
    if (!shaped.series.length) {
      ec.clear();
      ec.setOption({ title: { text: "Sin datos numéricos en el rango", left: "center", top: "center", textStyle: { color: "#888" } } });
      return;
    }
    ec.setOption(buildOption(cfg, shaped), true);
    ec.resize();
  }

  // -------------------------------------------------------------------------
  //  Live-update: revisa el rango cada 1.2s y re-renderiza si cambió.
  // -------------------------------------------------------------------------
  function startLive(cfg) {
    stopLive();
    if (!cfg.range) return;
    let lastJson = JSON.stringify(current._values || null);
    liveTimer = setInterval(() => {
      if (!$("chartModal").classList.contains("open")) return;
      const v = readRangeValues(cfg.range);
      const j = JSON.stringify(v);
      if (j !== lastJson) {
        lastJson = j;
        current._values = v;
        renderPreview(cfg);
      }
    }, 1200);
  }
  function stopLive() { if (liveTimer) { clearInterval(liveTimer); liveTimer = null; } }

  // -------------------------------------------------------------------------
  //  Persistencia de gráficos por pestaña
  // -------------------------------------------------------------------------
  function tabCharts() {
    const t = deps.getActiveTab();
    if (!t) return [];
    if (!t.charts) t.charts = [];
    return t.charts;
  }
  function saveChart(cfg) {
    const list = tabCharts();
    const idx = list.findIndex((c) => c.id === cfg.id);
    const clean = Object.assign({}, cfg); delete clean._values;
    if (idx >= 0) list[idx] = clean; else list.push(clean);
    if (deps.onDirty) deps.onDirty();
    renderChartList();
  }
  function deleteChart(id) {
    const list = tabCharts();
    const idx = list.findIndex((c) => c.id === id);
    if (idx >= 0) { list.splice(idx, 1); if (deps.onDirty) deps.onDirty(); renderChartList(); }
  }

  // -------------------------------------------------------------------------
  //  UI del diálogo
  // -------------------------------------------------------------------------
  function openModal(fromSelection) {
    const t = deps.getActiveTab();
    if (!t) return;
    // Si se abre desde una selección nueva, crear un gráfico a partir de ella.
    if (fromSelection) {
      const picked = readActiveRange();
      if (!picked) { deps.toast("Selecciona primero un rango con datos", "err"); return; }
      current = { id: defaultConfig().id, cfg: defaultConfig(), _values: picked.values };
      current.cfg.range = picked.range;
      current.cfg.name = "Gráfico " + (tabCharts().length + 1);
    } else if (!current) {
      current = { id: defaultConfig().id, cfg: defaultConfig() };
    }
    cfgToForm(current.cfg);
    renderChartList();
    $("chartModal").classList.add("open");
    setTimeout(() => { renderPreview(current.cfg); startLive(current.cfg); }, 60);
  }
  function closeModal() {
    stopLive();
    $("chartModal").classList.remove("open");
  }

  function cfgToForm(cfg) {
    $("chType").value = cfg.type;
    $("chTitle").value = cfg.title || "";
    $("chName").value = cfg.name || "";
    $("chPalette").value = cfg.palette || "vivid";
    $("chHeader").checked = !!cfg.firstRowIsHeader;
    $("chCategory").checked = !!cfg.firstColIsCategory;
    $("chLegend").checked = !!cfg.showLegend;
    $("chLabels").checked = !!cfg.showLabels;
    $("chSmooth").checked = !!cfg.smooth;
    $("chHorizontal").checked = !!cfg.horizontal;
    $("chDual").checked = !!cfg.dualAxis;
    $("chRangeInfo").textContent = cfg.range
      ? cfg.range.sheetName + "!" + a1(cfg.range.startRow, cfg.range.startColumn) + ":" + a1(cfg.range.endRow, cfg.range.endColumn)
      : "(sin rango)";
  }
  function formToCfg() {
    const cfg = current.cfg;
    cfg.type = $("chType").value;
    cfg.title = $("chTitle").value;
    cfg.name = $("chName").value || "Gráfico";
    cfg.palette = $("chPalette").value;
    cfg.firstRowIsHeader = $("chHeader").checked;
    cfg.firstColIsCategory = $("chCategory").checked;
    cfg.showLegend = $("chLegend").checked;
    cfg.showLabels = $("chLabels").checked;
    cfg.smooth = $("chSmooth").checked;
    cfg.horizontal = $("chHorizontal").checked;
    cfg.dualAxis = $("chDual").checked;
    return cfg;
  }

  function a1(row, col) {
    let s = "", c = col;
    do { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; } while (c >= 0);
    return s + (row + 1);
  }

  function renderChartList() {
    const ul = $("chList");
    if (!ul) return;
    const list = tabCharts();
    ul.innerHTML = "";
    if (!list.length) {
      const li = document.createElement("li");
      li.className = "empty";
      li.textContent = "Sin gráficos guardados.";
      ul.appendChild(li);
      return;
    }
    list.forEach((c) => {
      const li = document.createElement("li");
      li.className = current && c.id === current.cfg.id ? "active" : "";
      li.innerHTML = '<span class="cl-name"></span><span class="cl-del" title="Eliminar">✕</span>';
      li.querySelector(".cl-name").textContent = "📊 " + (c.name || "Gráfico");
      li.querySelector(".cl-name").addEventListener("click", () => loadChart(c.id));
      li.querySelector(".cl-del").addEventListener("click", (e) => { e.stopPropagation(); deleteChart(c.id); });
      ul.appendChild(li);
    });
  }

  function loadChart(id) {
    const c = tabCharts().find((x) => x.id === id);
    if (!c) return;
    current = { id: c.id, cfg: JSON.parse(JSON.stringify(c)) };
    current._values = c.range ? readRangeValues(c.range) : null;
    cfgToForm(current.cfg);
    renderChartList();
    renderPreview(current.cfg);
    startLive(current.cfg);
  }

  // -------------------------------------------------------------------------
  //  Exportar imagen
  // -------------------------------------------------------------------------
  function exportImage(kind) {
    if (!ec) return;
    let url, ext;
    if (kind === "svg") {
      // ECharts SVG requiere renderer svg; re-render temporal
      const tmp = document.createElement("div");
      tmp.style.cssText = "position:absolute;left:-9999px;width:900px;height:560px";
      document.body.appendChild(tmp);
      const svgEc = echarts.init(tmp, null, { renderer: "svg" });
      const shaped = shapeData(current._values || readRangeValues(current.cfg.range), current.cfg);
      svgEc.setOption(buildOption(current.cfg, shaped));
      url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgEc.renderToSVGString());
      svgEc.dispose(); tmp.remove();
      ext = "svg";
    } else {
      url = ec.getDataURL({ type: "png", pixelRatio: 2, backgroundColor: "#1e1e1e" });
      ext = "png";
    }
    const name = (current.cfg.name || "grafico").replace(/[\\\/:*?"<>|]/g, "_") + "." + ext;
    // Descargar (en la app o el navegador)
    const a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    deps.toast("Imagen exportada: " + name, "ok");
  }

  // -------------------------------------------------------------------------
  //  Wiring
  // -------------------------------------------------------------------------
  function wire() {
    $("chClose").addEventListener("click", closeModal);
    $("chCancel").addEventListener("click", closeModal);
    $("chNew").addEventListener("click", () => openModal(true));

    // Cambios en cualquier control -> re-render en vivo
    ["chType", "chTitle", "chName", "chPalette", "chHeader", "chCategory",
     "chLegend", "chLabels", "chSmooth", "chHorizontal", "chDual"].forEach((id) => {
      const el = $(id);
      if (!el) return;
      const ev = el.type === "checkbox" || el.tagName === "SELECT" ? "change" : "input";
      el.addEventListener(ev, () => { formToCfg(); renderPreview(current.cfg); });
    });

    $("chSave").addEventListener("click", () => { formToCfg(); saveChart(current.cfg); deps.toast("Gráfico guardado ✓", "ok"); });
    $("chExportPng").addEventListener("click", () => exportImage("png"));
    $("chExportSvg").addEventListener("click", () => exportImage("svg"));

    window.addEventListener("resize", () => { if (ec && $("chartModal").classList.contains("open")) ec.resize(); });
  }

  // -------------------------------------------------------------------------
  //  API pública
  // -------------------------------------------------------------------------
  window.XlsxChart = {
    init(d) { deps = d; wire(); },
    open() { openModal(true); },        // abrir con la selección actual
    openManager() { openModal(false); }, // abrir sin crear (gestor)
    // Al guardar el libro, inyectar los gráficos de la pestaña en el snapshot.
    attachToSnapshot(snapshot, tab) {
      if (tab && tab.charts && tab.charts.length) snapshot.__charts = tab.charts;
    },
    // Al abrir, recuperar los gráficos del snapshot a la pestaña.
    restoreFromData(uniData, tab) {
      if (uniData && uniData.__charts) tab.charts = uniData.__charts;
    },
  };
})();

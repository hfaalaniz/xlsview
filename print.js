/* ============================================================================
 *  print.js — Impresión / vista previa fiel de una hoja de cálculo.
 *
 *  Univer core no exporta a PDF ni imprime, así que renderizamos nosotros un
 *  HTML paginado a partir del snapshot de Univer + la configuración de página
 *  (tamaño de papel, orientación, márgenes, escala/ajuste, área de impresión,
 *  filas/columnas repetidas, encabezado/pie, cuadrícula, centrado).
 *
 *  El resultado se abre en una ventana de vista previa con botón "Imprimir",
 *  que usa window.print() de WebView2 (el usuario elige impresora o PDF).
 *
 *  API:  window.XlsxPrint.preview(snapshot, sheetId, page, meta)
 *        window.XlsxPrint.buildHtml(snapshot, sheetId, page, meta) -> string
 * ==========================================================================*/
(function () {
  "use strict";

  const DPI = 96;                 // px por pulgada en pantalla/impresión CSS
  const SSF = (window.XLSX && window.XLSX.SSF) || null;

  // ---- helpers de estilo -----------------------------------------------------
  const H_ALIGN = { 1: "left", 2: "center", 3: "right", 4: "justify" };
  const V_ALIGN = { 1: "top", 2: "middle", 3: "bottom" };
  // grosor de borde Univer -> CSS
  const BORDER_W = { 1: "1px", 2: "2px", 3: "3px", 7: "3px", 8: "2px", 13: "3px" };
  const BORDER_STYLE_CSS = { 7: "double", 10: "dashed", 4: "dashed", 3: "dotted" };

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function rgbOf(o) { return o && o.rgb ? o.rgb : null; }

  // Valor mostrado de una celda (aplica formato numérico si lo hay).
  function displayValue(cell) {
    if (!cell) return "";
    let v = cell.v;
    if (v == null || v === "") return "";
    const pattern = cell.s && cell.s.n && cell.s.n.pattern;
    // t: 1=string, 2=number, 3=boolean (según nuestra convención de conversión)
    if ((cell.t === 2 || typeof v === "number") && pattern && SSF) {
      try { return SSF.format(pattern, Number(v)); } catch (e) { return String(v); }
    }
    if (cell.t === 3) return v ? "TRUE" : "FALSE";
    return String(v);
  }

  // CSS inline del contenido de una celda a partir de su estilo Univer.
  //  scale: factor de escala aplicado también a la fuente (fit-to-page).
  function cellCss(s, scale) {
    if (!s) return "";
    scale = scale || 1;
    const css = [];
    if (s.ff) css.push("font-family:'" + s.ff.replace(/'/g, "") + "',Calibri,Arial,sans-serif");
    const fs = (s.fs || 11) * scale;
    css.push("font-size:" + (Math.round(fs * 10) / 10) + "pt");
    if (s.bl) css.push("font-weight:bold");
    if (s.it) css.push("font-style:italic");
    let deco = "";
    if (s.ul && s.ul.s) deco += " underline";
    if (s.st && s.st.s) deco += " line-through";
    if (deco) css.push("text-decoration:" + deco.trim());
    if (s.cl && rgbOf(s.cl)) css.push("color:" + rgbOf(s.cl));
    if (s.bg && rgbOf(s.bg)) css.push("background-color:" + rgbOf(s.bg));
    if (s.ht && H_ALIGN[s.ht]) css.push("text-align:" + H_ALIGN[s.ht]);
    if (s.vt && V_ALIGN[s.vt]) css.push("vertical-align:" + V_ALIGN[s.vt]);
    if (s.tb === 3) css.push("white-space:normal;word-break:break-word");
    else css.push("white-space:nowrap");
    return css.join(";");
  }

  // Bordes de la celda -> CSS (lados). Devuelve string de estilos border-*.
  function borderCss(s) {
    if (!s || !s.bd) return "";
    const css = [];
    const map = { l: "left", r: "right", t: "top", b: "bottom" };
    for (const k in map) {
      const b = s.bd[k];
      if (!b) continue;
      const w = BORDER_W[b.s] || "1px";
      const st = BORDER_STYLE_CSS[b.s] || "solid";
      const col = (b.cl && b.cl.rgb) || "#000";
      css.push("border-" + map[k] + ":" + w + " " + st + " " + col);
    }
    return css.join(";");
  }

  // Diagonal como fondo (line gradient) si existe.
  function diagCss(s) {
    if (!s || !s.bd) return "";
    const parts = [];
    if (s.bd.tl_br) parts.push("to bottom right");
    if (s.bd.bl_tr) parts.push("to top right");
    if (!parts.length) return "";
    const col = (s.bd.tl_br && s.bd.tl_br.cl && s.bd.tl_br.cl.rgb) ||
                (s.bd.bl_tr && s.bd.bl_tr.cl && s.bd.bl_tr.cl.rgb) || "#000";
    const grads = parts.map((p) =>
      "linear-gradient(" + p + ", transparent calc(50% - 0.6px), " + col + " calc(50% - 0.6px), " + col + " calc(50% + 0.6px), transparent calc(50% + 0.6px))");
    return "background-image:" + grads.join(",");
  }

  // ---- geometría de la hoja --------------------------------------------------
  function colWidthPx(sheet, c) {
    const cd = sheet.columnData && sheet.columnData[c];
    if (cd && cd.w) return cd.w;
    return sheet.defaultColumnWidth || 88;
  }
  function rowHeightPx(sheet, r) {
    const rd = sheet.rowData && sheet.rowData[r];
    if (rd && rd.h) return rd.h;
    return sheet.defaultRowHeight || 22;
  }
  function isColHidden(sheet, c) {
    const cd = sheet.columnData && sheet.columnData[c];
    return !!(cd && cd.hd);
  }
  function isRowHidden(sheet, r) {
    const rd = sheet.rowData && sheet.rowData[r];
    return !!(rd && rd.hd);
  }

  // Rango usado de la hoja (última fila/col con datos o estilo).
  function usedRange(sheet) {
    let maxR = 0, maxC = 0;
    const cd = sheet.cellData || {};
    for (const rk in cd) {
      const r = +rk;
      if (r > maxR) maxR = r;
      for (const ck in cd[rk]) { const c = +ck; if (c > maxC) maxC = c; }
    }
    // Considerar también columnas/filas con tamaño definido.
    for (const ck in (sheet.columnData || {})) { const c = +ck; if (c > maxC) maxC = c; }
    return { maxR, maxC };
  }

  // Parsea "A1:G50" -> {s:{r,c}, e:{r,c}} (0-based). Ignora el nombre de hoja.
  function parseRange(ref) {
    if (!ref) return null;
    const m = ref.split("!").pop().replace(/\$/g, "");
    const mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m);
    if (!mm) return null;
    const col = (s) => { let c = 0; for (let i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64); return c - 1; };
    return { s: { r: +mm[2] - 1, c: col(mm[1]) }, e: { r: +mm[4] - 1, c: col(mm[3]) } };
  }

  // Mapa de celdas cubiertas por merges: "r,c" -> { anchor, rowspan, colspan } | "skip".
  function buildMergeMap(sheet) {
    const map = {};
    (sheet.mergeData || []).forEach((m) => {
      const r0 = m.startRow, c0 = m.startColumn, r1 = m.endRow, c1 = m.endColumn;
      map[r0 + "," + c0] = { anchor: true, rowspan: r1 - r0 + 1, colspan: c1 - c0 + 1 };
      for (let r = r0; r <= r1; r++)
        for (let c = c0; c <= c1; c++)
          if (!(r === r0 && c === c0)) map[r + "," + c] = "skip";
    });
    return map;
  }

  // ---- paginación ------------------------------------------------------------
  //  Reparte columnas y filas en páginas según el ancho/alto imprimible.
  function paginate(sheet, page, range) {
    const scale = (page.setup.scale || 100) / 100;
    const paper = paperSize(page);
    // área imprimible en px (papel - márgenes), a 96 dpi
    const availW = (paper.w - page.margins.left - page.margins.right) * DPI;
    const availH = (paper.h - page.margins.top - page.margins.bottom - 0.4) * DPI; // 0.4in encabezado/pie

    // columnas visibles del rango
    const cols = [];
    for (let c = range.s.c; c <= range.e.c; c++) if (!isColHidden(sheet, c)) cols.push(c);
    const rows = [];
    for (let r = range.s.r; r <= range.e.r; r++) if (!isRowHidden(sheet, r)) rows.push(r);

    // Repetir columnas/filas (títulos): se anteponen en cada página.
    const repCols = page.repeatCols ? rangeList(page.repeatCols[0], page.repeatCols[1]).filter((c) => !isColHidden(sheet, c)) : [];
    const repRows = page.repeatRows ? rangeList(page.repeatRows[0], page.repeatRows[1]).filter((r) => !isRowHidden(sheet, r)) : [];
    const repColsSet = new Set(repCols), repRowsSet = new Set(repRows);
    const bodyCols = cols.filter((c) => !repColsSet.has(c));
    const bodyRows = rows.filter((r) => !repRowsSet.has(r));

    // "Ajustar a página": si fitToPage, calculamos una escala que quepa a lo ancho.
    let sc = scale;
    if (page.setup.fitToPage && page.setup.fitToWidth) {
      const totalW = [...repCols, ...bodyCols].reduce((a, c) => a + colWidthPx(sheet, c), 0);
      const maxW = (availW / page.setup.fitToWidth) * 0.985;   // 1.5% de margen de seguridad
      if (totalW > maxW) sc = Math.min(sc, maxW / totalW);
    }

    // repartir columnas en franjas horizontales
    const colPages = [];
    let cur = [], curW = repCols.reduce((a, c) => a + colWidthPx(sheet, c) * sc, 0);
    for (const c of bodyCols) {
      const w = colWidthPx(sheet, c) * sc;
      if (cur.length && curW + w > availW) { colPages.push(cur); cur = []; curW = repCols.reduce((a, cc) => a + colWidthPx(sheet, cc) * sc, 0); }
      cur.push(c); curW += w;
    }
    if (cur.length || !bodyCols.length) colPages.push(cur);

    // repartir filas en franjas verticales
    const rowPages = [];
    let curR = [], curH = repRows.reduce((a, r) => a + rowHeightPx(sheet, r) * sc, 0);
    for (const r of bodyRows) {
      const h = rowHeightPx(sheet, r) * sc;
      if (curR.length && curH + h > availH) { rowPages.push(curR); curR = []; curH = repRows.reduce((a, rr) => a + rowHeightPx(sheet, rr) * sc, 0); }
      curR.push(r); curH += h;
    }
    if (curR.length || !bodyRows.length) rowPages.push(curR);

    // producto: cada página = (franja de columnas) x (franja de filas)
    const pages = [];
    for (const rp of rowPages) {
      for (const cp of colPages) {
        pages.push({
          cols: [...repCols, ...cp],
          rows: [...repRows, ...rp],
        });
      }
    }
    return { pages, scale: sc, availW, availH, paper };
  }

  function rangeList(a, b) { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; }

  function paperSize(page) {
    const P = { "1": { w: 8.5, h: 11 }, "5": { w: 8.5, h: 14 }, "8": { w: 11.69, h: 16.54 },
                "9": { w: 8.27, h: 11.69 }, "11": { w: 5.83, h: 8.27 } };
    let p = P[page.setup.paperSize] || P["9"];
    if (page.setup.orientation === "landscape") p = { w: p.h, h: p.w };
    return p;
  }

  // ---- render de una página --------------------------------------------------
  function renderPage(sheet, pg, scale, mergeMap, page) {
    const cols = pg.cols, rows = pg.rows;
    // Posición de cada fila/columna DENTRO de esta página (para recortar merges).
    const colPos = new Map(); cols.forEach((c, i) => colPos.set(c, i));
    const rowPos = new Map(); rows.forEach((r, i) => rowPos.set(r, i));
    const colHdrOffset = page.options.headings ? 1 : 0;

    // Ancho total de la tabla = suma de anchos de columna (+ encabezado de fila).
    // Es imprescindible fijarlo para que table-layout:fixed respete el colgroup.
    let totalW = page.options.headings ? 34 : 0;
    for (const c of cols) totalW += Math.round(colWidthPx(sheet, c) * scale);

    let html = '<table class="grid" cellspacing="0" cellpadding="0" style="width:' + totalW + 'px"><colgroup>';
    if (page.options.headings) html += '<col style="width:34px">';
    for (const c of cols) html += '<col style="width:' + Math.round(colWidthPx(sheet, c) * scale) + 'px">';
    html += "</colgroup>";

    if (page.options.headings) {
      html += '<tr class="hdr" style="height:18px"><td class="corner"></td>';
      for (const c of cols) html += '<td class="chdr">' + colName(c) + "</td>";
      html += "</tr>";
    }

    for (const r of rows) {
      html += '<tr style="height:' + Math.round(rowHeightPx(sheet, r) * scale) + 'px">';
      if (page.options.headings) html += '<td class="rhdr">' + (r + 1) + "</td>";
      for (const c of cols) {
        const key = r + "," + c;
        const mm = mergeMap[key];
        if (mm === "skip") continue;
        const cell = sheet.cellData && sheet.cellData[r] && sheet.cellData[r][c];
        const s = cell && cell.s;
        // Fuente base escalada para celdas sin estilo propio de fuente.
        const baseFs = !s || !s.fs ? "font-size:" + (Math.round(11 * scale * 10) / 10) + "pt;" : "";
        const styleStr = baseFs + [cellCss(s, scale), borderCss(s), diagCss(s)].filter(Boolean).join(";");

        // Recortar el merge al rango de columnas/filas PRESENTES en esta página,
        // para que colspan/rowspan no desborden la tabla (rompen table-layout).
        let span = "";
        if (mm && mm.anchor) {
          const lastC = c + mm.colspan - 1, lastR = r + mm.rowspan - 1;
          let cs = 0; for (const cc of cols) if (cc >= c && cc <= lastC) cs++;
          let rs = 0; for (const rr of rows) if (rr >= r && rr <= lastR) rs++;
          if (cs > 1) span += ' colspan="' + cs + '"';
          if (rs > 1) span += ' rowspan="' + rs + '"';
        }
        const gl = page.options.gridLines ? " gl" : "";
        html += '<td class="cell' + gl + '"' + span + ' style="' + styleStr + '">' +
                esc(displayValue(cell)) + "</td>";
      }
      html += "</tr>";
    }
    html += "</table>";
    return html;
  }

  function colName(c) {
    let s = ""; c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }

  // Sustituye tokens de encabezado/pie: &P página, &N total, &D fecha, &F archivo, &A hoja.
  function hf(text, ctx) {
    return esc(text || "")
      .replace(/&P/g, ctx.page).replace(/&N/g, ctx.total)
      .replace(/&D/g, ctx.date).replace(/&F/g, ctx.file).replace(/&A/g, ctx.sheet);
  }

  // ---- construcción del documento completo ----------------------------------
  function buildHtml(snapshot, sheetId, page, meta) {
    meta = meta || {};
    const sheet = snapshot.sheets[sheetId];
    if (!sheet) return "<p>Hoja no encontrada.</p>";

    // Rango a imprimir: área de impresión si existe, si no el rango usado.
    let range = page.printArea ? parseRange(page.printArea) : null;
    if (!range) { const u = usedRange(sheet); range = { s: { r: 0, c: 0 }, e: { r: u.maxR, c: u.maxC } }; }

    const mergeMap = buildMergeMap(sheet);
    const pag = paginate(sheet, page, range);
    const paper = pag.paper;
    const date = new Date().toLocaleDateString();

    let pagesHtml = "";
    pag.pages.forEach((pg, i) => {
      const ctx = { page: i + 1, total: pag.pages.length, date: date,
                    file: meta.fileName || "", sheet: sheet.name || "" };
      const hdr = page.header ? '<div class="pg-hdr">' + hf(page.header, ctx) + "</div>" : "";
      const ftr = page.footer ? '<div class="pg-ftr">' + hf(page.footer, ctx) + "</div>" : "";
      const alignH = page.options.horizontalCentered ? "center" : "flex-start";
      const alignV = page.options.verticalCentered ? "center" : "flex-start";
      pagesHtml +=
        '<div class="page" style="width:' + paper.w + "in;height:" + paper.h + "in;" +
        "padding-top:" + page.margins.top + "in;padding-bottom:" + page.margins.bottom + "in;" +
        "padding-left:" + page.margins.left + "in;padding-right:" + page.margins.right + 'in">' +
        hdr +
        '<div class="pg-body" style="justify-content:' + alignH + ";align-items:" + alignV + '">' +
        renderPage(sheet, pg, pag.scale, mergeMap, page) +
        "</div>" + ftr +
        "</div>";
    });

    const css = pageCss(paper);
    return { html: pagesHtml, css: css, pageCount: pag.pages.length };
  }

  function pageCss(paper) {
    return `
      * { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      html, body { margin: 0; padding: 0; background: #525659; }
      .page { background: #fff; margin: 12px auto; box-shadow: 0 2px 12px rgba(0,0,0,.4);
              display: flex; flex-direction: column; overflow: hidden; }
      .pg-hdr, .pg-ftr { flex: 0 0 auto; font: 10px Arial; color: #444; text-align: center; padding: 2px 0; }
      .pg-ftr { margin-top: auto; }
      .pg-body { flex: 1 1 auto; display: flex; overflow: hidden; }
      table.grid { border-collapse: collapse; table-layout: fixed; }
      table.grid td { overflow: hidden; padding: 0 3px; line-height: 1.2;
                      vertical-align: bottom; text-overflow: clip;
                      font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
      td.cell.gl { border: 0.5px solid #d0d0d0; }
      td.chdr, td.rhdr, td.corner { background: #f0f0f0; border: 0.5px solid #bbb;
                text-align: center; font: 10px Arial; color: #666; }
      td.rhdr { min-width: 30px; }
      @media print {
        html, body { background: #fff; }
        .page { margin: 0; box-shadow: none; page-break-after: always; }
        .page:last-child { page-break-after: auto; }
      }
      @page { size: ${paper.w}in ${paper.h}in; margin: 0; }
    `;
  }

  // ---- ventana de vista previa ----------------------------------------------
  function preview(snapshot, sheetId, page, meta) {
    const built = buildHtml(snapshot, sheetId, page, meta);
    const win = document.getElementById("printPreview");
    const frame = document.getElementById("printFrame");
    if (!win || !frame) return built;

    const doc = frame.contentDocument;
    doc.open();
    doc.write("<!doctype html><html><head><meta charset='utf-8'><style>" +
      built.css + "</style></head><body>" + built.html + "</body></html>");
    doc.close();

    const pc = document.getElementById("printPageCount");
    if (pc) pc.textContent = built.pageCount + (built.pageCount === 1 ? " página" : " páginas");
    win.classList.add("open");
    return built;
  }

  function doPrint() {
    const frame = document.getElementById("printFrame");
    if (!frame) return;
    try {
      frame.contentWindow.focus();
      frame.contentWindow.print();
    } catch (e) { /* no-op */ }
  }

  function close() {
    const win = document.getElementById("printPreview");
    if (win) win.classList.remove("open");
  }

  window.XlsxPrint = { buildHtml: buildHtml, preview: preview, print: doPrint, close: close };
})();

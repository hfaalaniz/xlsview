/* ============================================================================
 *  xlsx-styles.js — Lector de estilos ricos de archivos .xlsx / .xlsm.
 *
 *  SheetJS (edición community) NO lee fuentes, bordes ni alineación desde el
 *  archivo. Aquí parseamos directamente el XML interno del .xlsx (que es un ZIP)
 *  usando XLSX.read(buf, { bookFiles:true }) para recuperar TODO el estilo del
 *  archivo original y replicarlo en Univer:
 *
 *    xl/styles.xml       -> fonts, fills, borders, numFmts, cellXfs
 *    xl/theme/theme1.xml -> paleta de colores de tema (para color theme="n")
 *    xl/worksheets/N.xml -> el atributo s="idx" de cada celda -> índice cellXfs
 *
 *  API:  window.XlsxStyles.build(workbookWithBookFiles)
 *        -> { sheetsByName: { "Hoja1": { "R,C": estiloUniver } }, xfCount }
 *
 *  El "estiloUniver" usa las claves que entiende Univer:
 *    ff fontFamily · fs fontSize · bl bold · it italic · ul underline · st strike
 *    cl color · bg background · ht hAlign · vt vAlign · tb wrap · n numFmt · bd bordes
 * ==========================================================================*/
(function () {
  "use strict";

  // ---- colores indexados legados de Excel (los más comunes) ----------------
  const INDEXED = {
    0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF",
    5: "FFFF00", 6: "FF00FF", 7: "00FFFF", 8: "000000", 9: "FFFFFF",
    10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF",
    15: "00FFFF", 16: "800000", 17: "008000", 18: "000080", 19: "808000",
    20: "800080", 21: "008080", 22: "C0C0C0", 23: "808080", 64: "000000",
  };

  // Orden de resolución de colores de tema (dk1/lt1 van intercambiados en OOXML).
  // Índices theme: 0=lt1(bg1) 1=dk1(text1) 2=lt2(bg2) 3=dk2(text2) 4..9=accent1..6
  const THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
                       "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

  const parser = new DOMParser();
  const decode = (v) => {
    if (v == null) return "";
    if (typeof v === "string") return v;
    // Uint8Array / array de bytes
    try { return new TextDecoder("utf-8").decode(new Uint8Array(v)); }
    catch (e) { return String(v); }
  };

  function getFile(wb, name) {
    if (!wb || !wb.files) return null;
    const f = wb.files[name];
    return f ? decode(f.content) : null;
  }

  function parseXml(text) {
    if (!text) return null;
    return parser.parseFromString(text, "application/xml");
  }

  // -- Normaliza un hex a "#RRGGBB" (descarta alfa AARRGGBB) ------------------
  function hex6(v) {
    if (!v) return undefined;
    v = String(v).replace(/^#/, "").toUpperCase();
    if (v.length === 8) v = v.slice(2);
    if (v.length === 6) return "#" + v;
    return undefined;
  }

  // Aplica tinte (tint de -1..1) a un hex, como hace Excel con colores de tema.
  function applyTint(hex, tint) {
    if (!hex || !tint) return hex;
    let r = parseInt(hex.slice(1, 3), 16);
    let g = parseInt(hex.slice(3, 5), 16);
    let b = parseInt(hex.slice(5, 7), 16);
    const f = (c) => {
      if (tint < 0) return Math.round(c * (1 + tint));
      return Math.round(c * (1 - tint) + 255 * tint);
    };
    r = Math.max(0, Math.min(255, f(r)));
    g = Math.max(0, Math.min(255, f(g)));
    b = Math.max(0, Math.min(255, f(b)));
    const h = (n) => n.toString(16).padStart(2, "0").toUpperCase();
    return "#" + h(r) + h(g) + h(b);
  }

  // -- Extrae la paleta de tema desde theme1.xml -----------------------------
  function parseTheme(doc) {
    const palette = [];
    if (!doc) return palette;
    const scheme = doc.getElementsByTagName("a:clrScheme")[0]
                || doc.getElementsByTagName("clrScheme")[0];
    if (!scheme) return palette;
    const map = {};
    for (const key of THEME_ORDER) {
      const node = scheme.getElementsByTagName("a:" + key)[0]
                || scheme.getElementsByTagName(key)[0];
      if (!node) { map[key] = undefined; continue; }
      const srgb = node.getElementsByTagName("a:srgbClr")[0]
                || node.getElementsByTagName("srgbClr")[0];
      const sys = node.getElementsByTagName("a:sysClr")[0]
                || node.getElementsByTagName("sysClr")[0];
      let val;
      if (srgb) val = srgb.getAttribute("val");
      else if (sys) val = sys.getAttribute("lastClr") || (sys.getAttribute("val") === "window" ? "FFFFFF" : "000000");
      map[key] = val ? "#" + val.toUpperCase() : undefined;
    }
    // Devolver por índice de tema (0..).
    return THEME_ORDER.map((k) => map[k]);
  }

  // -- Resuelve un nodo <color> (rgb / theme / indexed) a "#RRGGBB" -----------
  function resolveColor(node, theme) {
    if (!node) return undefined;
    const rgb = node.getAttribute("rgb");
    if (rgb) return hex6(rgb);
    const idx = node.getAttribute("indexed");
    if (idx != null && INDEXED[+idx]) return "#" + INDEXED[+idx];
    const th = node.getAttribute("theme");
    if (th != null) {
      let base = theme[+th];
      if (base) {
        const tint = parseFloat(node.getAttribute("tint") || "0");
        return applyTint(base, tint);
      }
    }
    return undefined;
  }

  function firstChildTag(parent, tag) {
    if (!parent) return null;
    return parent.getElementsByTagName(tag)[0] || null;
  }

  // -- Parsea styles.xml -> arreglos de fonts/fills/borders/numFmts + cellXfs -
  function parseStyles(doc, theme) {
    const out = { fonts: [], fills: [], borders: [], numFmts: {}, xfs: [] };
    if (!doc) return out;

    // numFmts personalizados
    const nfNodes = doc.getElementsByTagName("numFmt");
    for (let i = 0; i < nfNodes.length; i++) {
      const id = nfNodes[i].getAttribute("numFmtId");
      const code = nfNodes[i].getAttribute("formatCode");
      if (id != null) out.numFmts[id] = code;
    }

    // fonts
    const fontsRoot = doc.getElementsByTagName("fonts")[0];
    if (fontsRoot) {
      const fs = fontsRoot.getElementsByTagName("font");
      for (let i = 0; i < fs.length; i++) {
        const f = fs[i];
        const font = {};
        const name = firstChildTag(f, "name");
        if (name) font.name = name.getAttribute("val");
        const sz = firstChildTag(f, "sz");
        if (sz) font.size = parseFloat(sz.getAttribute("val"));
        if (f.getElementsByTagName("b").length) font.bold = true;
        if (f.getElementsByTagName("i").length) font.italic = true;
        if (f.getElementsByTagName("u").length) font.underline = true;
        if (f.getElementsByTagName("strike").length) font.strike = true;
        const col = firstChildTag(f, "color");
        if (col) font.color = resolveColor(col, theme);
        out.fonts.push(font);
      }
    }

    // fills
    const fillsRoot = doc.getElementsByTagName("fills")[0];
    if (fillsRoot) {
      const fl = fillsRoot.getElementsByTagName("fill");
      for (let i = 0; i < fl.length; i++) {
        const pat = firstChildTag(fl[i], "patternFill");
        const fill = {};
        if (pat) {
          const type = pat.getAttribute("patternType");
          if (type && type !== "none") {
            const fg = firstChildTag(pat, "fgColor");
            fill.color = resolveColor(fg, theme);
          }
        }
        out.fills.push(fill);
      }
    }

    // borders (incluye diagonales: diagonalUp = ↗ bl_tr, diagonalDown = ↘ tl_br)
    const bordersRoot = doc.getElementsByTagName("borders")[0];
    if (bordersRoot) {
      const bd = bordersRoot.getElementsByTagName("border");
      for (let i = 0; i < bd.length; i++) {
        const bEl = bd[i];
        const border = {};
        for (const side of ["left", "right", "top", "bottom"]) {
          const sNode = firstChildTag(bEl, side);
          if (!sNode) continue;
          const style = sNode.getAttribute("style");
          if (!style || style === "none") continue;
          const colNode = firstChildTag(sNode, "color");
          border[side] = { style, color: resolveColor(colNode, theme) || "#000000" };
        }
        // Diagonal: <diagonal style="..."> con flags de dirección en <border>.
        const diag = firstChildTag(bEl, "diagonal");
        const dStyle = diag && diag.getAttribute("style");
        if (dStyle && dStyle !== "none") {
          let up = bEl.getAttribute("diagonalUp") === "1";
          let down = bEl.getAttribute("diagonalDown") === "1";
          // Algunos escritores (p. ej. exceljs) omiten los flags de dirección;
          // si falta, asumimos ↘ (tl_br), la diagonal más habitual.
          if (!up && !down) down = true;
          const dColor = resolveColor(firstChildTag(diag, "color"), theme) || "#000000";
          if (down) border.tl_br = { style: dStyle, color: dColor };  // ↘
          if (up) border.bl_tr = { style: dStyle, color: dColor };    // ↗
        }
        out.borders.push(border);
      }
    }

    // cellXfs (los estilos aplicados a celdas)
    const cellXfsRoot = doc.getElementsByTagName("cellXfs")[0];
    if (cellXfsRoot) {
      const xfs = cellXfsRoot.getElementsByTagName("xf");
      for (let i = 0; i < xfs.length; i++) {
        const xf = xfs[i];
        const rec = {
          numFmtId: xf.getAttribute("numFmtId"),
          fontId: parseInt(xf.getAttribute("fontId") || "0", 10),
          fillId: parseInt(xf.getAttribute("fillId") || "0", 10),
          borderId: parseInt(xf.getAttribute("borderId") || "0", 10),
          align: null,
        };
        const al = firstChildTag(xf, "alignment");
        if (al) {
          rec.align = {
            h: al.getAttribute("horizontal") || null,
            v: al.getAttribute("vertical") || null,
            wrap: al.getAttribute("wrapText") === "1" || al.getAttribute("wrapText") === "true",
          };
        }
        // Protección de celda. Solo la registramos si el xf la declara
        // (applyProtection) para no marcar TODAS las celdas como bloqueadas.
        const prot = firstChildTag(xf, "protection");
        if (prot && xf.getAttribute("applyProtection") === "1") {
          rec.protection = {
            locked: prot.getAttribute("locked") !== "0",  // por defecto locked
            hidden: prot.getAttribute("hidden") === "1",
          };
        }
        out.xfs.push(rec);
      }
    }

    return out;
  }

  // Formatos numéricos incorporados de Excel (ids < 164), los más usados.
  const BUILTIN_NUMFMT = {
    "1": "0", "2": "0.00", "3": "#,##0", "4": "#,##0.00",
    "9": "0%", "10": "0.00%", "11": "0.00E+00", "12": "# ?/?",
    "13": "# ??/??", "14": "m/d/yyyy", "15": "d-mmm-yy", "16": "d-mmm",
    "17": "mmm-yy", "18": "h:mm AM/PM", "19": "h:mm:ss AM/PM",
    "20": "h:mm", "21": "h:mm:ss", "22": "m/d/yyyy h:mm",
    "37": "#,##0 ;(#,##0)", "38": "#,##0 ;[Red](#,##0)",
    "39": "#,##0.00;(#,##0.00)", "40": "#,##0.00;[Red](#,##0.00)",
    "44": '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    "45": "mm:ss", "46": "[h]:mm:ss", "47": "mmss.0", "48": "##0.0E+0", "49": "@",
  };

  // Mapeo de alineación OOXML -> claves de Univer.
  const H_ALIGN = { left: 1, center: 2, right: 3, justify: 4 };
  const V_ALIGN = { top: 1, middle: 2, center: 2, bottom: 3 };
  // Estilo de borde OOXML -> BorderStyleTypes de Univer (mapeo completo).
  const BORDER_STYLE = {
    thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6,
    double: 7, medium: 8, mediumDashed: 9, mediumDashDot: 10,
    mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
  };
  // Inverso: BorderStyleTypes de Univer -> estilo OOXML (para escribir).
  const BORDER_STYLE_REV = {
    1: "thin", 2: "hair", 3: "dotted", 4: "dashed", 5: "dashDot", 6: "dashDotDot",
    7: "double", 8: "medium", 9: "mediumDashed", 10: "mediumDashDot",
    11: "mediumDashDotDot", 12: "slantDashDot", 13: "thick",
  };

  // -- Construye el estilo Univer para un índice de cellXfs ------------------
  function xfToUniver(xf, styles) {
    if (!xf) return null;
    const s = {};
    let has = false;

    // Formato numérico
    if (xf.numFmtId && xf.numFmtId !== "0") {
      const code = styles.numFmts[xf.numFmtId] || BUILTIN_NUMFMT[xf.numFmtId];
      if (code) { s.n = { pattern: code }; has = true; }
    }

    // Fuente
    const font = styles.fonts[xf.fontId];
    if (font) {
      if (font.name) { s.ff = font.name; has = true; }
      if (font.size) { s.fs = font.size; has = true; }
      if (font.bold) { s.bl = 1; has = true; }
      if (font.italic) { s.it = 1; has = true; }
      if (font.underline) { s.ul = { s: 1 }; has = true; }
      if (font.strike) { s.st = { s: 1 }; has = true; }
      if (font.color) { s.cl = { rgb: font.color }; has = true; }
    }

    // Relleno
    const fill = styles.fills[xf.fillId];
    if (fill && fill.color) { s.bg = { rgb: fill.color }; has = true; }

    // Alineación
    if (xf.align) {
      if (xf.align.h && H_ALIGN[xf.align.h]) { s.ht = H_ALIGN[xf.align.h]; has = true; }
      if (xf.align.v && V_ALIGN[xf.align.v]) { s.vt = V_ALIGN[xf.align.v]; has = true; }
      if (xf.align.wrap) { s.tb = 3; has = true; }
    }

    // Bordes (lados + diagonales ↘ tl_br, ↗ bl_tr)
    const border = styles.borders[xf.borderId];
    if (border) {
      const bd = {};
      const sideKey = { left: "l", right: "r", top: "t", bottom: "b", tl_br: "tl_br", bl_tr: "bl_tr" };
      for (const side of ["left", "right", "top", "bottom", "tl_br", "bl_tr"]) {
        const b = border[side];
        if (!b) continue;
        bd[sideKey[side]] = { s: BORDER_STYLE[b.style] || 1, cl: { rgb: b.color || "#000000" } };
      }
      if (Object.keys(bd).length) { s.bd = bd; has = true; }
    }

    // Protección de celda. OOXML: por defecto locked=1 (bloqueada). Registramos
    // el estado explícito para (a) imponer el bloqueo con hoja protegida y
    // (b) re-escribir el xf fielmente. __unlocked marca las editables (locked=0).
    if (xf.protection) {
      if (xf.protection.locked === false) s.__unlocked = true;  // editable
      else s.__locked = true;                                   // bloqueada explícita
      if (xf.protection.hidden) s.__hidden = true;              // fórmula oculta
      has = true;
    }

    return has ? s : null;
  }

  // -- Mapea nombres de hoja a su archivo xl/worksheets/N.xml -----------------
  function sheetFileMap(wb) {
    // workbook.xml lista <sheet name r:id> ; workbook.xml.rels mapea r:id -> target.
    const map = {};
    const wbXml = parseXml(getFile(wb, "xl/workbook.xml"));
    const relsXml = parseXml(getFile(wb, "xl/_rels/workbook.xml.rels"));
    if (!wbXml || !relsXml) return map;

    const rels = {};
    const relNodes = relsXml.getElementsByTagName("Relationship");
    for (let i = 0; i < relNodes.length; i++) {
      rels[relNodes[i].getAttribute("Id")] = relNodes[i].getAttribute("Target");
    }
    const sheetNodes = wbXml.getElementsByTagName("sheet");
    for (let i = 0; i < sheetNodes.length; i++) {
      const name = sheetNodes[i].getAttribute("name");
      const rid = sheetNodes[i].getAttribute("r:id")
               || sheetNodes[i].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      let target = rels[rid];
      if (!target) continue;
      target = target.replace(/^\/?xl\//, "").replace(/^\//, "");
      map[name] = "xl/" + target;
    }
    return map;
  }

  // -- Decodifica "A1" -> {c,r} (0-based) ------------------------------------
  function decodeRef(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!m) return null;
    let col = 0;
    for (let i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
    return { c: col - 1, r: parseInt(m[2], 10) - 1 };
  }

  // -- Construye el mapa de estilos por hoja/celda ---------------------------
  function build(wb) {
    const result = { sheetsByName: {}, xfCount: 0, ok: false };
    if (!wb || !wb.files) return result;

    const themeDoc = parseXml(getFile(wb, "xl/theme/theme1.xml"));
    const theme = parseTheme(themeDoc);
    const stylesDoc = parseXml(getFile(wb, "xl/styles.xml"));
    const styles = parseStyles(stylesDoc, theme);
    result.xfCount = styles.xfs.length;

    // Precalcular el estilo Univer de cada índice de cellXfs.
    const xfUniver = styles.xfs.map((xf) => xfToUniver(xf, styles));

    result.panesByName = {};
    result.protectionByName = {};   // { hoja: { protected, opts } }
    result.pageByName = {};         // { hoja: { setup, margins, options, printArea, repeatRows, repeatCols } }

    // Áreas de impresión y filas/columnas repetidas viven en definedName del wb.
    const printAreas = {}, repeatByIdx = {};
    try {
      const wbDoc2 = parseXml(getFile(wb, "xl/workbook.xml"));
      if (wbDoc2) {
        const dn = wbDoc2.getElementsByTagName("definedName");
        for (let i = 0; i < dn.length; i++) {
          const nm = dn[i].getAttribute("name");
          const sheetIdx = dn[i].getAttribute("localSheetId");
          const val = (dn[i].textContent || "").trim();
          if (nm === "_xlnm.Print_Area" && sheetIdx != null) printAreas[sheetIdx] = val;
          if (nm === "_xlnm.Print_Titles" && sheetIdx != null) repeatByIdx[sheetIdx] = val;
        }
      }
    } catch (e) { /* sin definedNames */ }

    const files = sheetFileMap(wb);
    let sheetIndex = -1;
    for (const sheetName in files) {
      sheetIndex++;
      const sheetXml = getFile(wb, files[sheetName]);
      if (!sheetXml) continue;
      const doc = parseXml(sheetXml);
      if (!doc) continue;

      const cellStyles = {};
      const cNodes = doc.getElementsByTagName("c");
      for (let i = 0; i < cNodes.length; i++) {
        const c = cNodes[i];
        const sIdx = c.getAttribute("s");
        if (sIdx == null) continue;                 // celda sin estilo aplicado
        const style = xfUniver[+sIdx];
        if (!style) continue;                        // el estilo por defecto (0) no aporta nada
        const ref = c.getAttribute("r");
        const pos = ref ? decodeRef(ref) : null;
        if (!pos) continue;
        cellStyles[pos.r + "," + pos.c] = style;
      }
      result.sheetsByName[sheetName] = cellStyles;

      // Freeze panes: <pane xSplit="n" ySplit="m" state="frozen"/>
      const paneNode = doc.getElementsByTagName("pane")[0];
      if (paneNode && /frozen/i.test(paneNode.getAttribute("state") || "")) {
        const xs = parseInt(paneNode.getAttribute("xSplit") || "0", 10) || 0;
        const ys = parseInt(paneNode.getAttribute("ySplit") || "0", 10) || 0;
        if (xs > 0 || ys > 0) result.panesByName[sheetName] = { xSplit: xs, ySplit: ys };
      }

      // Protección de hoja: <sheetProtection sheet="1" .../>
      const spNode = doc.getElementsByTagName("sheetProtection")[0];
      if (spNode && spNode.getAttribute("sheet") === "1") {
        // Conservamos los atributos originales para reescribirlos fielmente.
        const opts = {};
        for (let a = 0; a < spNode.attributes.length; a++) {
          const at = spNode.attributes[a];
          opts[at.name] = at.value;
        }
        result.protectionByName[sheetName] = { protected: true, opts: opts };
      }

      // Configuración de página: pageSetup, pageMargins, printOptions.
      const page = { setup: {}, margins: {}, options: {} };
      const psNode = doc.getElementsByTagName("pageSetup")[0];
      if (psNode) {
        page.setup.orientation = psNode.getAttribute("orientation") || null;   // portrait/landscape
        page.setup.paperSize = psNode.getAttribute("paperSize") || null;       // 9=A4, 1=Letter…
        page.setup.scale = psNode.getAttribute("scale") || null;
        page.setup.fitToWidth = psNode.getAttribute("fitToWidth") || null;
        page.setup.fitToHeight = psNode.getAttribute("fitToHeight") || null;
      }
      const pmNode = doc.getElementsByTagName("pageMargins")[0];
      if (pmNode) {
        for (const m of ["left", "right", "top", "bottom", "header", "footer"]) {
          const v = pmNode.getAttribute(m);
          if (v != null) page.margins[m] = parseFloat(v);
        }
      }
      // sheetPr > pageSetUpPr fitToPage
      const setupPr = doc.getElementsByTagName("pageSetUpPr")[0];
      if (setupPr && setupPr.getAttribute("fitToPage") === "1") page.setup.fitToPage = true;
      const poNode = doc.getElementsByTagName("printOptions")[0];
      if (poNode) {
        page.options.gridLines = poNode.getAttribute("gridLines") === "1";
        page.options.headings = poNode.getAttribute("headings") === "1";
        page.options.horizontalCentered = poNode.getAttribute("horizontalCentered") === "1";
        page.options.verticalCentered = poNode.getAttribute("verticalCentered") === "1";
      }
      // Área de impresión y filas/columnas repetidas (de definedName).
      if (printAreas[sheetIndex]) page.printArea = printAreas[sheetIndex];
      if (repeatByIdx[sheetIndex]) {
        const parsed = parsePrintTitles(repeatByIdx[sheetIndex]);
        if (parsed.rows) page.repeatRows = parsed.rows;
        if (parsed.cols) page.repeatCols = parsed.cols;
      }
      result.pageByName[sheetName] = page;
    }

    result.ok = true;
    return result;
  }

  // Parsea "Hoja!$1:$3,Hoja!$A:$B" -> { rows:[0,2], cols:[0,1] } (0-based).
  function parsePrintTitles(val) {
    const out = {};
    (val || "").split(",").forEach((part) => {
      const m = part.split("!").pop();
      const rowM = /^\$?(\d+):\$?(\d+)$/.exec(m);
      const colM = /^\$?([A-Z]+):\$?([A-Z]+)$/.exec(m);
      if (rowM) out.rows = [parseInt(rowM[1], 10) - 1, parseInt(rowM[2], 10) - 1];
      else if (colM) out.cols = [colToIdx(colM[1]), colToIdx(colM[2])];
    });
    return out;
  }
  function colToIdx(s) {
    let c = 0;
    for (let i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64);
    return c - 1;
  }

  // ===========================================================================
  //  ESCRITURA de estilos (writer)
  //
  //  SheetJS community descarta fuentes/bordes/alineación al guardar. Aquí, tras
  //  que SheetJS escriba el .xlsx (datos + fórmulas + numFmt), reescribimos el
  //  ZIP interno: generamos un styles.xml completo desde los estilos de Univer y
  //  asignamos a cada celda su índice de estilo (s=) en xl/worksheets/N.xml.
  // ===========================================================================

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  const H_ALIGN_REV = { 1: "left", 2: "center", 3: "right", 4: "justify" };
  const V_ALIGN_REV = { 1: "top", 2: "center", 3: "bottom" };
  // estilo de borde Univer -> OOXML (mapeo completo, definido arriba).
  const BORDER_REV = BORDER_STYLE_REV;

  // Normaliza "#RRGGBB" -> "FFRRGGBB" (ARGB opaco) para OOXML.
  function toArgb(rgb) {
    if (!rgb) return null;
    let v = String(rgb).replace(/^#/, "").toUpperCase();
    if (v.length === 6) return "FF" + v;
    if (v.length === 8) return v;
    return null;
  }

  // Construye las tablas de estilos únicos a partir de todas las celdas Univer.
  //  protection: { "sheet-N": { enabled, opts, unlocked:{"R,C":true} } }
  function collectStyles(snapshot, protection) {
    protection = protection || {};
    // Tablas con deduplicación por clave JSON.
    const fonts = [{}];   // índice 0 = fuente por defecto
    const fills = [{ none: true }, { gray125: true }];  // 0 none, 1 gray125 (obligatorio)
    const borders = [{}]; // índice 0 = sin bordes
    const numFmts = {};   // code -> id (custom, id >= 164)
    let nextNumFmtId = 164;
    const xfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }]; // xf 0 por defecto
    const xfKey = new Map([[JSON.stringify(xfs[0]), 0]]);

    const idxOf = (arr, keyObj) => {
      const k = JSON.stringify(keyObj);
      for (let i = 0; i < arr.length; i++) if (JSON.stringify(arr[i]) === k) return i;
      arr.push(keyObj);
      return arr.length - 1;
    };

    // Mapa: sheetOrderName -> { "R,C": xfIndex }
    const perSheet = {};

    const order = snapshot.sheetOrder || Object.keys(snapshot.sheets);
    order.forEach((sid) => {
      const sh = snapshot.sheets[sid];
      if (!sh) return;
      const name = sh.name || sid;
      const cellStyleIdx = {};
      const cellData = sh.cellData || {};
      // Protección de esta hoja (por id de hoja Univer).
      const prot = protection[sid];
      const sheetProtected = !!(prot && prot.enabled);
      const unlocked = (prot && prot.unlocked) || {};
      Object.keys(cellData).forEach((rk) => {
        const row = cellData[rk];
        Object.keys(row).forEach((ck) => {
          const uc = row[ck];
          if (!uc || !uc.s) return;
          const s = uc.s;

          // Fuente
          let fontId = 0;
          if (s.ff || s.fs || s.bl || s.it || s.ul || s.st || s.cl) {
            fontId = idxOf(fonts, {
              name: s.ff || null, sz: s.fs || null,
              b: !!s.bl, i: !!s.it, u: !!(s.ul && s.ul.s), strike: !!(s.st && s.st.s),
              color: s.cl && s.cl.rgb ? toArgb(s.cl.rgb) : null,
            });
          }
          // Relleno
          let fillId = 0;
          if (s.bg && s.bg.rgb) {
            fillId = idxOf(fills, { solid: toArgb(s.bg.rgb) });
          }
          // Borde (lados + diagonales tl_br ↘ / bl_tr ↗)
          let borderId = 0;
          if (s.bd && (s.bd.l || s.bd.r || s.bd.t || s.bd.b || s.bd.tl_br || s.bd.bl_tr)) {
            const b = {};
            [["l", "left"], ["r", "right"], ["t", "top"], ["b", "bottom"],
             ["tl_br", "tl_br"], ["bl_tr", "bl_tr"]].forEach(([k, side]) => {
              if (s.bd[k]) b[side] = { style: BORDER_REV[s.bd[k].s] || "thin", color: (s.bd[k].cl && s.bd[k].cl.rgb) ? toArgb(s.bd[k].cl.rgb) : "FF000000" };
            });
            borderId = idxOf(borders, b);
          }
          // Formato numérico
          let numFmtId = 0;
          if (s.n && s.n.pattern) {
            if (numFmts[s.n.pattern] == null) { numFmts[s.n.pattern] = nextNumFmtId++; }
            numFmtId = numFmts[s.n.pattern];
          }
          // Alineación
          let align = null;
          if (s.ht || s.vt || s.tb === 3) {
            align = {
              h: H_ALIGN_REV[s.ht] || null,
              v: V_ALIGN_REV[s.vt] || null,
              wrap: s.tb === 3,
            };
          }

          // Protección: solo relevante si la hoja está protegida. Marcamos
          // locked=false para las celdas explícitamente desbloqueadas.
          let prot2 = null;
          if (sheetProtected && unlocked[rk + "," + ck]) prot2 = { locked: false };

          const xf = { numFmtId, fontId, fillId, borderId, align, prot: prot2 };
          const key = JSON.stringify(xf);
          let xi = xfKey.get(key);
          if (xi == null) { xi = xfs.length; xfs.push(xf); xfKey.set(key, xi); }
          cellStyleIdx[rk + "," + ck] = xi;
        });
      });
      perSheet[name] = cellStyleIdx;
    });

    return { fonts, fills, borders, numFmts, xfs, perSheet };
  }

  // Serializa styles.xml desde las tablas recogidas.
  function buildStylesXml(t) {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

    // numFmts personalizados
    const nfCodes = Object.keys(t.numFmts);
    if (nfCodes.length) {
      xml += '<numFmts count="' + nfCodes.length + '">';
      nfCodes.forEach((code) => {
        xml += '<numFmt numFmtId="' + t.numFmts[code] + '" formatCode="' + esc(code) + '"/>';
      });
      xml += '</numFmts>';
    }

    // fonts
    xml += '<fonts count="' + t.fonts.length + '">';
    t.fonts.forEach((f) => {
      xml += "<font>";
      if (f.b) xml += "<b/>";
      if (f.i) xml += "<i/>";
      if (f.u) xml += "<u/>";
      if (f.strike) xml += "<strike/>";
      xml += '<sz val="' + (f.sz || 11) + '"/>';
      if (f.color) xml += '<color rgb="' + f.color + '"/>';
      else xml += '<color theme="1"/>';
      xml += '<name val="' + esc(f.name || "Calibri") + '"/>';
      xml += "</font>";
    });
    xml += "</fonts>";

    // fills
    xml += '<fills count="' + t.fills.length + '">';
    t.fills.forEach((f) => {
      if (f.none) xml += '<fill><patternFill patternType="none"/></fill>';
      else if (f.gray125) xml += '<fill><patternFill patternType="gray125"/></fill>';
      else if (f.solid) xml += '<fill><patternFill patternType="solid"><fgColor rgb="' + f.solid + '"/><bgColor indexed="64"/></patternFill></fill>';
      else xml += '<fill><patternFill patternType="none"/></fill>';
    });
    xml += "</fills>";

    // borders (con diagonales: tl_br ↘ = diagonalDown, bl_tr ↗ = diagonalUp)
    xml += '<borders count="' + t.borders.length + '">';
    t.borders.forEach((b) => {
      const down = b.tl_br, up = b.bl_tr;
      let attrs = "";
      if (down) attrs += ' diagonalDown="1"';
      if (up) attrs += ' diagonalUp="1"';
      xml += "<border" + attrs + ">";
      ["left", "right", "top", "bottom"].forEach((side) => {
        const s = b[side];
        if (s) xml += "<" + side + ' style="' + s.style + '"><color rgb="' + s.color + '"/></' + side + ">";
        else xml += "<" + side + "/>";
      });
      // Un solo <diagonal>; si hay ambas direcciones comparten estilo/color.
      const diag = down || up;
      if (diag) xml += '<diagonal style="' + diag.style + '"><color rgb="' + diag.color + '"/></diagonal>';
      else xml += "<diagonal/>";
      xml += "</border>";
    });
    xml += "</borders>";

    // cellStyleXfs (uno base)
    xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';

    // cellXfs
    xml += '<cellXfs count="' + t.xfs.length + '">';
    t.xfs.forEach((xf) => {
      const applyNum = xf.numFmtId ? ' applyNumberFormat="1"' : "";
      const applyFont = xf.fontId ? ' applyFont="1"' : "";
      const applyFill = xf.fillId ? ' applyFill="1"' : "";
      const applyBorder = xf.borderId ? ' applyBorder="1"' : "";
      const applyAlign = xf.align ? ' applyAlignment="1"' : "";
      const applyProt = xf.prot ? ' applyProtection="1"' : "";
      xml += '<xf numFmtId="' + xf.numFmtId + '" fontId="' + xf.fontId + '" fillId="' + xf.fillId +
             '" borderId="' + xf.borderId + '" xfId="0"' + applyNum + applyFont + applyFill + applyBorder + applyAlign + applyProt + ">";
      if (xf.align) {
        let a = "<alignment";
        if (xf.align.h) a += ' horizontal="' + xf.align.h + '"';
        if (xf.align.v) a += ' vertical="' + xf.align.v + '"';
        if (xf.align.wrap) a += ' wrapText="1"';
        a += "/>";
        xml += a;
      }
      if (xf.prot) {
        xml += '<protection locked="' + (xf.prot.locked === false ? "0" : "1") + '"/>';
      }
      xml += "</xf>";
    });
    xml += "</cellXfs>";

    xml += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
    xml += "</styleSheet>";
    return xml;
  }

  // Codifica {r,c} (0-based) -> "A1"
  function encodeRef(r, c) {
    let s = "";
    c = c + 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s + (r + 1);
  }

  // Inserta/actualiza el <pane> de freeze dentro del <sheetView> de la hoja.
  function injectPane(sheetXml, freeze) {
    if (!freeze || (!freeze.xSplit && !freeze.ySplit)) return sheetXml;
    const xs = freeze.xSplit || 0, ys = freeze.ySplit || 0;
    const topLeft = encodeRef(ys, xs);
    const activePane = xs > 0 && ys > 0 ? "bottomRight" : (xs > 0 ? "topRight" : "bottomLeft");
    let pane = '<pane';
    if (xs) pane += ' xSplit="' + xs + '"';
    if (ys) pane += ' ySplit="' + ys + '"';
    pane += ' topLeftCell="' + topLeft + '" activePane="' + activePane + '" state="frozen"/>';

    // Quitar cualquier <pane .../> previo dentro del sheetView.
    let xml = sheetXml.replace(/<pane\b[^>]*\/>/, "");
    // El <pane> debe ir DENTRO del <sheetView>. Distinguimos el sheetView
    // autocerrado (<sheetView .../>) del que tiene contenido (<sheetView ...>…).
    if (/<sheetView\b[^>]*\/>/.test(xml)) {
      // Autocerrado: expandir para alojar el pane.
      xml = xml.replace(/(<sheetView\b[^>]*?)\s*\/>/, "$1>" + pane + "</sheetView>");
    } else if (/<sheetView\b[^>]*>/.test(xml)) {
      // Con contenido: insertar el pane justo tras la etiqueta de apertura.
      xml = xml.replace(/(<sheetView\b[^>]*>)/, "$1" + pane);
    }
    return xml;
  }

  // Inserta/actualiza la configuración de página (printOptions, pageMargins,
  // pageSetup) en el XML de la hoja. El orden OOXML es: … sheetData, [mergeCells],
  // [printOptions], [pageMargins], [pageSetup] …
  function injectPageSetup(sheetXml, page) {
    if (!page) return sheetXml;
    let xml = sheetXml;

    // Quitar los existentes para no duplicar.
    xml = xml.replace(/<printOptions\b[^>]*\/>/g, "")
             .replace(/<pageMargins\b[^>]*\/>/g, "")
             .replace(/<pageSetup\b[^>]*\/>/g, "");

    const o = page.options || {};
    const m = page.margins || {};
    const s = page.setup || {};

    let block = "";
    // printOptions (solo si hay algo activado)
    if (o.gridLines || o.headings || o.horizontalCentered || o.verticalCentered) {
      block += "<printOptions";
      if (o.horizontalCentered) block += ' horizontalCentered="1"';
      if (o.verticalCentered) block += ' verticalCentered="1"';
      if (o.headings) block += ' headings="1"';
      if (o.gridLines) block += ' gridLines="1"';
      block += "/>";
    }
    // pageMargins (siempre)
    const mv = (k, d) => (m[k] != null ? m[k] : d);
    block += '<pageMargins left="' + mv("left", 0.7) + '" right="' + mv("right", 0.7) +
             '" top="' + mv("top", 0.75) + '" bottom="' + mv("bottom", 0.75) +
             '" header="' + mv("header", 0.3) + '" footer="' + mv("footer", 0.3) + '"/>';
    // pageSetup
    block += '<pageSetup paperSize="' + (s.paperSize || "9") + '"';
    if (s.fitToPage) {
      block += ' fitToWidth="' + (s.fitToWidth || 1) + '" fitToHeight="' + (s.fitToHeight || 0) + '"';
    } else {
      block += ' scale="' + (s.scale || 100) + '"';
    }
    block += ' orientation="' + (s.orientation || "portrait") + '"/>';

    // pageSetup/pageMargins/printOptions van al FINAL del worksheet (tras
    // sheetProtection, mergeCells, etc.). Insertar antes de </worksheet> respeta
    // el orden del esquema de forma robusta. Colocamos antes de otros elementos
    // finales conocidos (drawing, legacyDrawing) si los hubiera.
    const finalTags = /(<drawing\b|<legacyDrawing\b|<\/worksheet>)/;
    if (finalTags.test(xml)) {
      xml = xml.replace(finalTags, block + "$1");
    } else {
      xml = xml.replace(/(<\/sheetData>)/, "$1" + block);
    }

    // Marcar fitToPage en sheetPr > pageSetUpPr si aplica.
    if (s.fitToPage) {
      if (/<pageSetUpPr\b[^>]*\/>/.test(xml)) {
        xml = xml.replace(/<pageSetUpPr\b[^>]*\/>/, '<pageSetUpPr fitToPage="1"/>');
      } else if (/<sheetPr\b[^>]*>/.test(xml)) {
        xml = xml.replace(/(<sheetPr\b[^>]*>)/, '$1<pageSetUpPr fitToPage="1"/>');
      } else if (/<sheetPr\b[^>]*\/>/.test(xml)) {
        xml = xml.replace(/(<sheetPr\b[^>]*?)\s*\/>/, '$1><pageSetUpPr fitToPage="1"/></sheetPr>');
      }
    }
    return xml;
  }

  // Inserta/actualiza <sheetProtection .../> tras el bloque de cols o datos.
  // Debe ir después de <sheetData> según el esquema OOXML.
  function injectSheetProtection(sheetXml, prot) {
    if (!prot || !prot.enabled) return sheetXml;
    // Reconstruir los atributos originales (o unos por defecto sensatos).
    const opts = prot.opts || {};
    let attrs = ' sheet="1"';
    // Conservar los flags originales si estaban; si no, valores por defecto de
    // Excel (bloquea edición pero permite seleccionar celdas).
    const keep = ["algorithmName", "hashValue", "saltValue", "spinCount", "password",
                  "objects", "scenarios", "formatCells", "formatColumns", "formatRows",
                  "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns",
                  "deleteRows", "selectLockedCells", "sort", "autoFilter",
                  "pivotTables", "selectUnlockedCells"];
    for (const k of keep) if (opts[k] != null) attrs += ' ' + k + '="' + opts[k] + '"';
    const tag = "<sheetProtection" + attrs + "/>";

    let xml = sheetXml.replace(/<sheetProtection\b[^>]*\/>/, "");   // quitar previo
    // Insertar justo tras el cierre de </sheetData>.
    if (/<\/sheetData>/.test(xml)) {
      xml = xml.replace(/(<\/sheetData>)/, "$1" + tag);
    }
    return xml;
  }

  // Reescribe el atributo s= de cada celda en el XML de una hoja (+ freeze +
  // protección + configuración de página).
  function rewriteSheetXml(sheetXml, cellStyleIdx, freeze, prot, page) {
    let xml = sheetXml.replace(/<c\b([^>]*?)(\/?)>/g, (full, attrs, selfClose) => {
      const rm = /\br="([A-Z]+\d+)"/.exec(attrs);
      if (!rm) return full;
      const pos = decodeRef(rm[1]);
      if (!pos) return full;
      const xi = cellStyleIdx[pos.r + "," + pos.c];
      if (xi == null || xi === 0) return full;   // sin estilo o estilo por defecto
      let newAttrs = attrs;
      if (/\bs="\d+"/.test(newAttrs)) newAttrs = newAttrs.replace(/\bs="\d+"/, 's="' + xi + '"');
      else newAttrs = ' s="' + xi + '"' + newAttrs;
      return "<c" + newAttrs + selfClose + ">";
    });
    xml = injectPane(xml, freeze);
    xml = injectSheetProtection(xml, prot);
    xml = injectPageSetup(xml, page);
    return xml;
  }

  // Escribe/actualiza los definedName de Área de impresión y Títulos repetidos
  // en workbook.xml, según la config de página de cada hoja (por índice).
  function updateDefinedNames(zip, CFB, readPath, snapshot, pageByName, nameToPath) {
    let wbXml = readPath("xl/workbook.xml");
    if (!wbXml) return;

    // Orden de hojas por índice (el localSheetId de definedName es 0-based).
    const order = snapshot.sheetOrder || Object.keys(snapshot.sheets || {});
    const names = order.map((sid) => (snapshot.sheets[sid] && snapshot.sheets[sid].name) || sid);

    // Construir los definedName nuevos.
    let defs = "";
    names.forEach((nm, idx) => {
      const p = pageByName[nm];
      if (!p) return;
      const q = "'" + String(nm).replace(/'/g, "''") + "'";
      if (p.printArea) {
        const ref = String(p.printArea).split("!").pop().replace(/\$/g, "");
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
        if (m) {
          defs += '<definedName name="_xlnm.Print_Area" localSheetId="' + idx + '">' +
                  q + "!$" + m[1] + "$" + m[2] + ":$" + m[3] + "$" + m[4] + "</definedName>";
        }
      }
      const titles = [];
      if (p.repeatCols) titles.push(q + "!$" + colName2(p.repeatCols[0]) + ":$" + colName2(p.repeatCols[1]));
      if (p.repeatRows) titles.push(q + "!$" + (p.repeatRows[0] + 1) + ":$" + (p.repeatRows[1] + 1));
      if (titles.length) {
        defs += '<definedName name="_xlnm.Print_Titles" localSheetId="' + idx + '">' +
                titles.join(",") + "</definedName>";
      }
    });

    // Quitar los definedName de impresión previos.
    wbXml = wbXml.replace(/<definedName name="_xlnm\.Print_(Area|Titles)"[\s\S]*?<\/definedName>/g, "");
    // Quitar un <definedNames> vacío que pudiera quedar.
    wbXml = wbXml.replace(/<definedNames>\s*<\/definedNames>/g, "");

    if (defs) {
      // Usamos función de reemplazo para que los "$" de las referencias ($A$1)
      // no se interpreten como grupos de captura de String.replace.
      if (/<definedNames>/.test(wbXml)) {
        wbXml = wbXml.replace(/<definedNames>/, () => "<definedNames>" + defs);
      } else {
        // Insertar el bloque tras </sheets> (posición válida en el esquema).
        wbXml = wbXml.replace(/<\/sheets>/, () => "</sheets><definedNames>" + defs + "</definedNames>");
      }
    }
    CFB.utils.cfb_add(zip, "/xl/workbook.xml", str2buf(wbXml));
  }

  function colName2(c) {
    let s = ""; c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }

  // API pública de escritura: recibe el buffer xlsx (Uint8Array/ArrayBuffer), el
  // snapshot de Univer y opts { protection, page }; devuelve un nuevo Uint8Array
  // con los estilos, freeze panes, protección y config de página aplicados.
  function apply(xlsxBuf, snapshot, opts) {
    opts = opts || {};
    const protection = opts.protection || {};
    const pageByName = opts.page || {};   // config de página por nombre de hoja
    const CFB = (window.XLSX && window.XLSX.CFB) || null;
    if (!CFB) return xlsxBuf;   // sin CFB no podemos reescribir; devolver tal cual

    const buf = xlsxBuf instanceof Uint8Array ? xlsxBuf : new Uint8Array(xlsxBuf);
    let zip;
    try { zip = CFB.read(buf, { type: "buffer" }); }
    catch (e) { return xlsxBuf; }

    const tables = collectStyles(snapshot, protection);
    const stylesXml = buildStylesXml(tables);

    // Mapear nombre de hoja -> ruta sheetN.xml, leyendo workbook.xml(.rels).
    const findPath = (p) => zip.FullPaths.find((fp) => fp.replace(/^Root Entry\//, "") === p || fp.endsWith("/" + p));
    const readPath = (p) => {
      const fp = findPath(p);
      if (!fp) return null;
      const entry = zip.FileIndex[zip.FullPaths.indexOf(fp)];
      const c = entry && entry.content;
      if (c == null) return null;
      return typeof c === "string" ? c : decode(c);
    };

    // wb rels para nombre->target
    const wbDoc = parseXml(readPath("xl/workbook.xml"));
    const relsDoc = parseXml(readPath("xl/_rels/workbook.xml.rels"));
    const nameToPath = {};
    if (wbDoc && relsDoc) {
      const rels = {};
      const rl = relsDoc.getElementsByTagName("Relationship");
      for (let i = 0; i < rl.length; i++) rels[rl[i].getAttribute("Id")] = rl[i].getAttribute("Target");
      const sn = wbDoc.getElementsByTagName("sheet");
      for (let i = 0; i < sn.length; i++) {
        const nm = sn[i].getAttribute("name");
        const rid = sn[i].getAttribute("r:id") ||
          sn[i].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        let target = rels[rid];
        if (!target) continue;
        target = target.replace(/^\/?xl\//, "").replace(/^\//, "");
        nameToPath[nm] = "xl/" + target;
      }
    }

    // Mapas nombre de hoja -> freeze / protección (desde snapshot y opts).
    const freezeByName = {};
    const protByName = {};
    const order = snapshot.sheetOrder || Object.keys(snapshot.sheets || {});
    order.forEach((sid) => {
      const sh = snapshot.sheets[sid];
      const nm = (sh && sh.name) || sid;
      if (sh && sh.freeze && (sh.freeze.xSplit || sh.freeze.ySplit)) freezeByName[nm] = sh.freeze;
      if (protection[sid] && protection[sid].enabled) protByName[nm] = protection[sid];
    });

    // Reescribir cada hoja con estilos (+ freeze panes + protección + página).
    // Iteramos sobre TODAS las hojas (no solo las que tienen estilos), porque la
    // configuración de página aplica aunque la hoja no lleve estilos ricos.
    Object.keys(nameToPath).forEach((name) => {
      const path = nameToPath[name];
      const xml = readPath(path);
      if (!xml) return;
      const rewritten = rewriteSheetXml(
        xml, tables.perSheet[name] || {}, freezeByName[name], protByName[name], pageByName[name]);
      CFB.utils.cfb_add(zip, "/" + path, str2buf(rewritten));
    });

    // Área de impresión / títulos repetidos -> definedName en workbook.xml.
    updateDefinedNames(zip, CFB, readPath, snapshot, pageByName, nameToPath);

    // Inyectar el nuevo styles.xml.
    CFB.utils.cfb_add(zip, "/xl/styles.xml", str2buf(stylesXml));

    try {
      const out = CFB.write(zip, { type: "buffer", fileType: "zip", compression: true });
      return out instanceof Uint8Array ? out : new Uint8Array(out);
    } catch (e) {
      return xlsxBuf;
    }
  }

  function str2buf(s) {
    return new TextEncoder().encode(s);
  }

  window.XlsxStyles = { build: build, apply: apply };
})();

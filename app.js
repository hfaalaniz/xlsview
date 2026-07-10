/* ============================================================================
 *  XlsView — lógica del visor/editor de hojas de cálculo.
 *
 *  - Lee .xlsx/.xlsm/.xls/.csv con SheetJS y los muestra en Univer (edición
 *    completa: fórmulas, formato, varias hojas).
 *  - Guarda de vuelta al disco a través del host (WebView2 / servidor local),
 *    o descarga en el navegador si no hay host.
 *
 *  Puente con el host (idéntico patrón a pdfview):
 *    - Apertura inicial:  ?file=/xls/<token>
 *    - Aperturas posteriores: se recogen por sondeo de /pending
 *    - Guardar:           POST /save/<token>  (cuerpo = binario del archivo)
 *    - Guardar como:      POST /saveas?name=...  -> el host abre "Guardar como"
 *    - Control ventana:   postMessage("close"|"fullscreen"|"minimize"|"drag…")
 * ==========================================================================*/
(function () {
  "use strict";

  // ---- Globals de Univer expuestos por los bundles UMD ----
  //  En Univer 0.5.x, LocaleType vive en UniverCore y NO existe mergeLocales:
  //  el objeto de locale (ya combinado por su propio bundle) se pasa tal cual.
  const { createUniver } = window.UniverPresets;
  const { LocaleType } = window.UniverCore;
  const { UniverSheetsCorePreset } = window.UniverPresetSheetsCore;
  const { defaultTheme } = window.UniverDesign;   // requerido por ThemeService
  const localeEnUS = window.UniverPresetSheetsCoreEnUS;

  const inHostApp = !!(window.chrome && window.chrome.webview);

  // ---------------------------------------------------------------------------
  //  Estado
  // ---------------------------------------------------------------------------
  let univerAPI = null;          // API facade de Univer
  const tabs = [];               // { id, name, token, unitId, dirty }
  let activeTab = null;
  let suppressDirty = false;      // ignora eventos de edición durante la carga

  // Tamaños de papel (pulgadas) por código OOXML paperSize.
  const PAPER = {
    "1": { name: "Carta", w: 8.5, h: 11 },
    "5": { name: "Legal", w: 8.5, h: 14 },
    "8": { name: "A3", w: 11.69, h: 16.54 },
    "9": { name: "A4", w: 8.27, h: 11.69 },
    "11": { name: "A5", w: 5.83, h: 8.27 },
  };

  // Config de página por defecto (estilo Excel) para una hoja.
  function defaultPage() {
    return {
      setup: { orientation: "portrait", paperSize: "9", scale: 100,
               fitToPage: false, fitToWidth: 1, fitToHeight: 0 },
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      options: { gridLines: false, headings: false, horizontalCentered: false, verticalCentered: false },
      printArea: null, repeatRows: null, repeatCols: null,
      header: "", footer: "",
    };
  }

  // Combina la config leída del archivo con los valores por defecto.
  function normalizePageConfig(byName, sheetNames) {
    const out = {};
    (sheetNames || []).forEach((sn) => {
      const d = defaultPage();
      const src = byName && byName[sn];
      if (src) {
        if (src.setup) {
          if (src.setup.orientation) d.setup.orientation = src.setup.orientation;
          if (src.setup.paperSize) d.setup.paperSize = src.setup.paperSize;
          if (src.setup.scale) d.setup.scale = parseInt(src.setup.scale, 10) || 100;
          if (src.setup.fitToPage) d.setup.fitToPage = true;
          if (src.setup.fitToWidth != null) d.setup.fitToWidth = parseInt(src.setup.fitToWidth, 10);
          if (src.setup.fitToHeight != null) d.setup.fitToHeight = parseInt(src.setup.fitToHeight, 10);
        }
        if (src.margins) Object.assign(d.margins, src.margins);
        if (src.options) Object.assign(d.options, src.options);
        if (src.printArea) d.printArea = src.printArea;
        if (src.repeatRows) d.repeatRows = src.repeatRows;
        if (src.repeatCols) d.repeatCols = src.repeatCols;
      }
      out[sn] = d;
    });
    return out;
  }

  // ---------------------------------------------------------------------------
  //  Utilidades DOM
  // ---------------------------------------------------------------------------
  const $ = (id) => document.getElementById(id);

  function toast(msg, kind) {
    const t = $("toast");
    t.textContent = msg;
    t.className = "show" + (kind ? " " + kind : "");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { t.className = ""; }, 2600);
  }

  function showLoading(text) {
    $("loadingText").textContent = text || "Cargando…";
    $("loading").classList.add("show");
  }
  function hideLoading() { $("loading").classList.remove("show"); }

  function setWelcome(visible) {
    $("welcome").classList.toggle("hidden", !visible);
    $("univer-host").classList.toggle("hidden", visible);
  }

  // ---------------------------------------------------------------------------
  //  Puente con el host
  // ---------------------------------------------------------------------------
  function postToHost(message) {
    if (inHostApp) { window.chrome.webview.postMessage(message); return true; }
    return false;
  }

  // ---------------------------------------------------------------------------
  //  Inicialización de Univer
  // ---------------------------------------------------------------------------
  function initUniver() {
    const { univerAPI: api } = createUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: localeEnUS },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({ container: "univer-host" }),
      ],
    });
    univerAPI = api;

    // Marcar la pestaña activa como "modificada" ante cualquier edición.
    // SheetValueChanged es la señal directa de cambio de datos; añadimos
    // CommandExecuted (filtrado) como red de seguridad para estilos/formato.
    try {
      const E = univerAPI.Event;
      if (E && E.SheetValueChanged) {
        univerAPI.addEvent(E.SheetValueChanged, () => {
          if (activeTab && !suppressDirty) markDirty(activeTab, true);
        });
      }
      if (E && E.CommandExecuted) {
        univerAPI.addEvent(E.CommandExecuted, (p) => {
          if (!activeTab || suppressDirty) return;
          const id = p && p.command && p.command.id ? p.command.id : "";
          if (/insert|remove|move|clear|paste|numfmt|style|merge|font|border|background/i.test(id)) {
            markDirty(activeTab, true);
          }
        });
      }

      // Protección de hoja: si la hoja activa está protegida y la celda está
      // bloqueada (locked), cancelamos el inicio de la edición.
      if (E && E.BeforeSheetEditStart) {
        univerAPI.addEvent(E.BeforeSheetEditStart, (p) => {
          if (!isCellLocked(p && p.worksheet, p && p.row, p && p.column)) return;
          if (p) p.cancel = true;   // impide entrar en modo edición
          toast("Celda bloqueada (hoja protegida)", "err");
        });
      }
    } catch (e) { /* la API de eventos puede variar entre versiones */ }
    // Nota: registerCustomFunctions() se llama al montar el primer workbook
    // (mountTab), no aquí: Univer necesita un libro activo para conservarlas.
  }

  let customFnsRegistered = false;

  // ---------------------------------------------------------------------------
  //  Funciones propias de negocio (registerFunction de la Facade de Univer).
  //  Se comportan como funciones nativas: se usan en celdas y se recalculan.
  //  Añade las tuyas aquí siguiendo el mismo patrón [fn, "NOMBRE"].
  // ---------------------------------------------------------------------------
  function registerCustomFunctions() {
    if (customFnsRegistered) return;
    if (!univerAPI || !univerAPI.registerFunction) return;
    // Firma de la Facade (Univer 0.5.x): registerFunction({ name, func, description }).
    // 'func' recibe los argumentos ya evaluados y devuelve el resultado.
    //
    // La Facade de Univer 0.5.x espera UN objeto { calculate }, donde calculate
    // es un array de tripletes [func, "NOMBRE", "descripción abstracta"].
    // (NO pasamos 'description' como objeto: cuando falta, Univer genera el
    // descriptor a partir del triplete. Pasarlo como string dispara un fallo del
    // LocaleService.)  'func' recibe los argumentos ya evaluados (primitivos).
    // Utilidades locales para las funciones.
    const num = (x, def) => { const n = Number(x); return isNaN(n) ? (def == null ? 0 : def) : n; };
    const opt = (x, def) => (x == null || x === "") ? def : Number(x);

    const calculate = [
      // ===================== COMERCIAL / CONTABLE =====================
      [(monto, tasa) => num(monto) * (1 + opt(tasa, 0.21)),
        "IVA", "Aplica IVA a un monto (21% por defecto): IVA(monto,[tasa])"],
      [(total, tasa) => num(total) / (1 + opt(tasa, 0.21)),
        "NETO", "Quita el IVA de un total (21% por defecto): NETO(total,[tasa])"],
      [(base, tasa) => num(base) * opt(tasa, 0.21),
        "IIBB", "Ingresos brutos sobre una base: IIBB(base,[tasa])"],
      [(costo, margen) => { const m = opt(margen, 0.4); return num(costo) / (1 - m); },
        "PRECIOVENTA", "Precio de venta según margen: PRECIOVENTA(costo,[margen%])"],
      [(costo, pv) => num(pv) === 0 ? 0 : (num(pv) - num(costo)) / num(pv),
        "MARGEN", "Margen sobre venta: MARGEN(costo, precioVenta)"],
      [(costo, pv) => num(costo) === 0 ? 0 : (num(pv) - num(costo)) / num(costo),
        "MARKUP", "Recargo sobre costo: MARKUP(costo, precioVenta)"],
      [(precio, pct) => num(precio) * (1 - opt(pct, 0)),
        "DESCUENTO", "Precio con descuento: DESCUENTO(precio, pct)"],
      [(viejo, nuevo) => num(viejo) === 0 ? 0 : (num(nuevo) - num(viejo)) / num(viejo),
        "VARIACIONPCT", "Variación porcentual: VARIACIONPCT(anterior, actual)"],
      [(parte, total) => num(total) === 0 ? 0 : num(parte) / num(total),
        "PORCENTAJEDE", "Qué porcentaje es parte de total: PORCENTAJEDE(parte, total)"],

      // ===================== FINANZAS =====================
      [(capital, tasa, periodos) => num(capital) * opt(tasa, 0) * num(periodos),
        "INTERESSIMPLE", "Interés simple: INTERESSIMPLE(capital, tasa, periodos)"],
      [(capital, tasa, periodos) => num(capital) * Math.pow(1 + opt(tasa, 0), num(periodos)) - num(capital),
        "INTERESCOMP", "Interés compuesto ganado: INTERESCOMP(capital, tasa, periodos)"],
      [(vp, tasa, periodos) => num(vp) * Math.pow(1 + opt(tasa, 0), num(periodos)),
        "VALORFUTURO", "Valor futuro: VALORFUTURO(valorPresente, tasa, periodos)"],
      [(monto, tasa, cuotas) => {
        const i = opt(tasa, 0), n = num(cuotas);
        if (i === 0) return num(monto) / n;
        return num(monto) * i / (1 - Math.pow(1 + i, -n));
      }, "CUOTAFIJA", "Cuota fija (sistema francés): CUOTAFIJA(monto, tasaPeriodo, nCuotas)"],
      [(nominalAnual, m) => Math.pow(1 + opt(nominalAnual, 0) / opt(m, 12), opt(m, 12)) - 1,
        "TASAEFECTIVA", "Tasa efectiva anual: TASAEFECTIVA(nominalAnual,[capitalizaciones=12])"],

      // ===================== INGENIERÍA ELÉCTRICA =====================
      [(tension, corriente, cosphi) => Math.sqrt(3) * num(tension) * num(corriente) * opt(cosphi, 0.85),
        "POTENCIA3F", "Potencia trifásica (W): POTENCIA3F(V, I, [cosφ])"],
      [(tension, corriente, cosphi) => num(tension) * num(corriente) * opt(cosphi, 1),
        "POTENCIA1F", "Potencia monofásica (W): POTENCIA1F(V, I, [cosφ])"],
      [(longitud, corriente, seccion, kappa) =>
        (2 * num(longitud) * num(corriente)) / (opt(kappa, 56) * num(seccion, 1)),
        "CAIDATENSION", "Caída de tensión monofásica (V): CAIDATENSION(L, I, S, [κ])"],
      [(longitud, corriente, seccion, cosphi, kappa) =>
        (Math.sqrt(3) * num(longitud) * num(corriente) * opt(cosphi, 0.85)) / (opt(kappa, 56) * num(seccion, 1)),
        "CAIDA3F", "Caída de tensión trifásica (V): CAIDA3F(L, I, S, [cosφ], [κ])"],
      [(tension, resistencia) => num(resistencia) === 0 ? 0 : num(tension) / num(resistencia),
        "OHM", "Ley de Ohm, corriente (A): OHM(tensión, resistencia)"],
      [(potencia, horas, precioKwh) => (num(potencia) / 1000) * num(horas) * opt(precioKwh, 0),
        "CONSUMOKWH", "Consumo eléctrico ($): CONSUMOKWH(potenciaW, horas, [precioKwh])"],
      [(corriente, tension, potenciaAparente) =>
        num(potenciaAparente) === 0 ? 0 : (num(tension) * num(corriente)) / num(potenciaAparente),
        "FACTORPOT", "Factor de potencia: FACTORPOT(I, V, potenciaAparente)"],

      // ===================== MATEMÁTICA / GEOMETRÍA =====================
      [(a, b) => Math.sqrt(num(a) * num(a) + num(b) * num(b)),
        "HIPOTENUSA", "Hipotenusa (Pitágoras): HIPOTENUSA(catetoA, catetoB)"],
      [(radio) => Math.PI * num(radio) * num(radio),
        "AREACIRCULO", "Área de un círculo: AREACIRCULO(radio)"],
      [(base, altura) => num(base) * num(altura) / 2,
        "AREATRIANGULO", "Área de un triángulo: AREATRIANGULO(base, altura)"],
      [(valor, multiplo) => { const mlt = opt(multiplo, 1); return mlt === 0 ? num(valor) : Math.round(num(valor) / mlt) * mlt; },
        "REDONDEARM", "Redondea al múltiplo más cercano: REDONDEARM(valor, múltiplo)"],
      [(peso, altura) => { const h = num(altura); return h === 0 ? 0 : num(peso) / (h * h); },
        "IMC", "Índice de masa corporal: IMC(pesoKg, alturaM)"],
      [(a, b, c) => num(b) === 0 ? 0 : num(a) * num(c) / num(b),
        "REGLA3", "Regla de tres simple: REGLA3(a, b, c) = a·c/b"],

      // ===================== CONVERSIÓN =====================
      [(c) => num(c) * 9 / 5 + 32, "CELSIUS_F", "Celsius→Fahrenheit: CELSIUS_F(°C)"],
      [(f) => (num(f) - 32) * 5 / 9, "FAHRENHEIT_C", "Fahrenheit→Celsius: FAHRENHEIT_C(°F)"],
      [(cv) => num(cv) * 735.49875, "CV_W", "Caballos de vapor→Watts: CV_W(cv)"],
      [(km, litros) => num(litros) === 0 ? 0 : num(km) / num(litros),
        "CONSUMOKML", "Rendimiento km/l: CONSUMOKML(km, litros)"],

      // ===================== TEXTO / UTILIDAD =====================
      [(texto) => String(texto == null ? "" : texto).split(/\s+/).filter(Boolean)
        .map(p => p.charAt(0).toUpperCase()).join(""),
        "INICIALES", "Iniciales de un nombre: INICIALES(\"Juan Pérez\") → JP"],
      [(texto) => String(texto == null ? "" : texto).replace(/[^0-9]/g, ""),
        "SOLONUMEROS", "Deja solo los dígitos: SOLONUMEROS(\"AB-12-3\") → 123"],
      [(texto) => {
        const w = String(texto == null ? "" : texto).toLowerCase().split(/\s+/);
        return w.map(x => x ? x.charAt(0).toUpperCase() + x.slice(1) : x).join(" ");
      }, "TITULAR", "Pone en mayúscula cada palabra: TITULAR(\"hola mundo\")"],
      [(texto) => String(texto == null ? "" : texto).split(/\s+/).filter(Boolean).length,
        "CONTARPALABRAS", "Cuenta palabras: CONTARPALABRAS(texto)"],

      // ===================== FECHA =====================
      [(nacimiento) => {
        // nacimiento como número de serie de Excel (días desde 1899-12-30)
        const serial = num(nacimiento);
        if (!serial) return 0;
        const ms = (serial - 25569) * 86400000;   // a epoch Unix
        const d = new Date(ms);
        const hoy = new Date();
        let edad = hoy.getUTCFullYear() - d.getUTCFullYear();
        const mm = hoy.getUTCMonth() - d.getUTCMonth();
        if (mm < 0 || (mm === 0 && hoy.getUTCDate() < d.getUTCDate())) edad--;
        return edad;
      }, "EDAD", "Edad en años desde una fecha: EDAD(fechaNacimiento)"],
      [(fecha) => {
        const serial = num(fecha);
        if (!serial) return 0;
        const d = new Date((serial - 25569) * 86400000);
        return Math.floor(d.getUTCMonth() / 3) + 1;
      }, "TRIMESTRE", "Trimestre (1-4) de una fecha: TRIMESTRE(fecha)"],
    ];
    try { univerAPI.registerFunction({ calculate }); customFnsRegistered = true; }
    catch (e) { /* si la firma cambia entre versiones, no rompemos el arranque */ }
  }

  // ¿La celda (row,col) de esa worksheet está bloqueada por protección de hoja?
  function isCellLocked(worksheet, row, column) {
    if (!activeTab || !activeTab.protection) return false;
    let sheetId = null;
    try { sheetId = worksheet && worksheet.getSheetId ? worksheet.getSheetId() : null; } catch (e) {}
    if (!sheetId) return false;
    const prot = activeTab.protection[sheetId];
    if (!prot || !prot.enabled) return false;
    // OOXML: con la hoja protegida, toda celda está bloqueada salvo las que se
    // marcaron locked="0" (editables). Así una celda está bloqueada si NO está
    // en el conjunto de desbloqueadas.
    return !(prot.unlocked && prot.unlocked[row + "," + column]);
  }

  // ===========================================================================
  //  Conversión  SheetJS  <->  Univer
  // ===========================================================================

  // -- Alto nivel: workbook de SheetJS -> datos de workbook de Univer --------
  //  richStyles: mapa opcional { nombreHoja: { "R,C": estiloUniver } } con los
  //  estilos completos leídos del archivo original (xlsx-styles.js).
  function sheetjsToUniver(wb, name, richStyles, richPanes, richProtection) {
    const sheets = {};
    const sheetOrder = [];
    let firstId = null;
    // Protección recolectada por id de hoja Univer (no cabe en el snapshot):
    //   { "sheet-0": { enabled, opts, locked: {"R,C":true} } }
    const protection = {};

    wb.SheetNames.forEach((sn, idx) => {
      const ws = wb.Sheets[sn];
      const sheetId = "sheet-" + idx;
      if (idx === 0) firstId = sheetId;
      sheetOrder.push(sheetId);

      const sheetStyles = (richStyles && richStyles[sn]) || null;
      const pane = (richPanes && richPanes[sn]) || null;
      const prot = (richProtection && richProtection[sn]) || null;

      const ref = ws["!ref"] || "A1";
      const range = XLSX.utils.decode_range(ref);
      const rowCount = Math.max(range.e.r + 1, 20);
      const colCount = Math.max(range.e.c + 1, 12);

      const cellData = {};
      const unlockedCells = {};   // "R,C" -> true  (celdas editables: locked="0")
      for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = ws[addr];
          const rich = sheetStyles ? sheetStyles[R + "," + C] : null;
          // Puede haber estilo sin valor (celda vacía con formato/borde).
          if (!cell && !rich) continue;
          // Registrar celdas explícitamente DESbloqueadas (para hoja protegida).
          if (rich && rich.__unlocked) unlockedCells[R + "," + C] = true;
          const uCell = convertCellToUniver(cell || {});
          const merged = mergeCellStyle(uCell, stripProtFlags(rich));
          if (merged) {
            (cellData[R] || (cellData[R] = {}))[C] = merged;
          }
        }
      }

      // Anchos de columna / altos de fila (si SheetJS los expuso)
      const columnData = {};
      (ws["!cols"] || []).forEach((c, i) => {
        if (c && (c.wpx || c.wch)) {
          // wpx viene en píxeles reales; wch en caracteres (~7px c/u + margen).
          columnData[i] = { w: c.wpx ? Math.round(c.wpx) : Math.round(c.wch * 7 + 5) };
        }
        if (c && c.hidden) (columnData[i] || (columnData[i] = {})).hd = 1;
      });
      const rowData = {};
      (ws["!rows"] || []).forEach((r, i) => {
        // SheetJS reporta hpx == hpt (puntos), sin convertir a píxeles. Excel usa
        // 1pt = 4/3 px, así que convertimos desde hpt para el alto real en pantalla.
        const h = r && (r.hpt != null ? r.hpt * 4 / 3 : r.hpx);
        if (h) rowData[i] = { h: Math.round(h) };
        if (r && r.hidden) (rowData[i] || (rowData[i] = {})).hd = 1;
      });

      // Celdas combinadas
      const mergeData = (ws["!merges"] || []).map((m) => ({
        startRow: m.s.r, startColumn: m.s.c, endRow: m.e.r, endColumn: m.e.c,
      }));

      const sheetDef = {
        id: sheetId,
        name: sn,
        rowCount,
        columnCount: colCount,
        cellData,
        columnData,
        rowData,
        mergeData,
        defaultColumnWidth: 88,
        defaultRowHeight: 22,
      };

      // Inmovilizar filas/columnas (freeze panes) del archivo original.
      if (pane && (pane.xSplit > 0 || pane.ySplit > 0)) {
        sheetDef.freeze = {
          xSplit: pane.xSplit || 0,          // nº de columnas congeladas
          ySplit: pane.ySplit || 0,          // nº de filas congeladas
          startRow: pane.ySplit || 0,        // primera fila desplazable
          startColumn: pane.xSplit || 0,     // primera columna desplazable
        };
      }

      // Protección de la hoja: guardamos aparte. Con la hoja protegida, TODA
      // celda está bloqueada salvo las explícitamente desbloqueadas (locked="0").
      if (prot && prot.protected) {
        protection[sheetId] = {
          enabled: true,
          opts: prot.opts || {},
          unlocked: unlockedCells,
          name: sn,
        };
      }

      sheets[sheetId] = sheetDef;
    });

    return {
      id: "wb-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
      name: name || "Workbook",
      appVersion: "xlsview",
      locale: LocaleType.EN_US,
      sheetOrder,
      sheets,
      __protection: protection,   // consumido por la app (no por Univer)
    };
  }

  // -- Celda SheetJS -> celda Univer -----------------------------------------
  function convertCellToUniver(cell) {
    const u = {};
    let has = false;
    const hasFormula = !!cell.f;

    // Valor y tipo. En celdas con fórmula solo tomamos el valor si SheetJS
    // guardó un resultado cacheado válido (no un error): así Univer recalcula
    // la fórmula en vez de heredar un "#ERROR" cosmético.
    if (cell.t === "n") { u.v = cell.v; u.t = 2; has = true; }        // number
    else if (cell.t === "b") { u.v = cell.v ? 1 : 0; u.t = 3; has = true; } // boolean
    else if (cell.t === "d") { u.v = cell.w || String(cell.v); u.t = 1; has = true; }
    else if (cell.t === "e") { if (!hasFormula) { u.v = cell.w || "#ERROR"; u.t = 1; has = true; } }
    else if (cell.v !== undefined && cell.v !== null) { u.v = String(cell.v); u.t = 1; has = true; } // string

    // Fórmula
    if (hasFormula) {
      u.f = "=" + cell.f;   // Univer espera la fórmula con "="
      has = true;
    }

    // Estilo básico de SheetJS (relleno + formato numérico). El estilo rico
    // completo se fusiona luego en mergeCellStyle().
    const s = convertStyleToUniver(cell);
    if (s) { u.s = s; has = true; }

    return has ? u : null;
  }

  // Devuelve una copia del estilo rico sin los flags internos de protección
  // (__locked/__hidden), que no son propiedades de estilo válidas para Univer.
  function stripProtFlags(rich) {
    if (!rich || (!rich.__locked && !rich.__hidden && !rich.__unlocked)) return rich;
    const clean = {};
    for (const k in rich) {
      if (k === "__locked" || k === "__hidden" || k === "__unlocked") continue;
      clean[k] = rich[k];
    }
    return Object.keys(clean).length ? clean : null;
  }

  // Fusiona el estilo rico (del parser XML, con fuentes/bordes/alineación) sobre
  // el estilo básico de SheetJS. El rico tiene prioridad porque proviene directo
  // del archivo. Devuelve la celda (o null si queda vacía sin contenido).
  function mergeCellStyle(uCell, rich) {
    if (!rich) return uCell && Object.keys(uCell).length ? uCell : null;
    const cell = uCell || {};
    const base = cell.s || {};
    cell.s = Object.assign({}, base, rich);   // rich pisa las claves de base
    // Marcar que hay contenido aunque la celda no tuviera valor (celda con solo
    // formato/borde): Univer necesita el registro para pintar el estilo.
    if (Object.keys(cell).length === 1 && cell.s) {
      // solo estilo, sin valor ni fórmula: lo dejamos igual (Univer lo acepta).
    }
    return Object.keys(cell).length ? cell : null;
  }

  function argb(hex) {
    if (!hex) return undefined;
    hex = String(hex).replace(/^#/, "");
    if (hex.length === 8) hex = hex.slice(2);       // quitar alfa AARRGGBB
    if (hex.length === 6) return "#" + hex.toUpperCase();
    return undefined;
  }

  function convertStyleToUniver(cell) {
    const st = cell.s;
    const numFmt = cell.z;
    const style = {};
    let has = false;

    if (numFmt && numFmt !== "General") {
      style.n = { pattern: numFmt };
      has = true;
    }

    if (st) {
      // Fuente
      if (st.font) {
        const f = st.font;
        if (f.name) { style.ff = f.name; has = true; }
        if (f.sz)   { style.fs = f.sz; has = true; }
        if (f.bold) { style.bl = 1; has = true; }
        if (f.italic) { style.it = 1; has = true; }
        if (f.underline) { style.ul = { s: 1 }; has = true; }
        if (f.strike) { style.st = { s: 1 }; has = true; }
        if (f.color && f.color.rgb) { style.cl = { rgb: argb(f.color.rgb) }; has = true; }
      }
      // Relleno
      if (st.fill && st.fill.fgColor && st.fill.fgColor.rgb) {
        style.bg = { rgb: argb(st.fill.fgColor.rgb) };
        has = true;
      }
      // Alineación:  Univer horizontal 1=left 2=center 3=right ; vertical 1=top 2=middle 3=bottom
      if (st.alignment) {
        const a = st.alignment;
        const hMap = { left: 1, center: 2, right: 3 };
        const vMap = { top: 1, center: 2, middle: 2, bottom: 3 };
        if (a.horizontal && hMap[a.horizontal]) { style.ht = hMap[a.horizontal]; has = true; }
        if (a.vertical && vMap[a.vertical]) { style.vt = vMap[a.vertical]; has = true; }
        if (a.wrapText) { style.tb = 3; has = true; } // 3 = wrap
      }
    }

    return has ? style : null;
  }

  // -- Univer -> workbook SheetJS (para guardar) ------------------------------
  function univerToSheetjs(snapshot) {
    const wb = XLSX.utils.book_new();
    const order = snapshot.sheetOrder || Object.keys(snapshot.sheets);

    order.forEach((sid) => {
      const sh = snapshot.sheets[sid];
      if (!sh) return;
      const ws = {};
      let maxR = 0, maxC = 0;
      const cellData = sh.cellData || {};

      Object.keys(cellData).forEach((rk) => {
        const R = +rk;
        const row = cellData[rk];
        Object.keys(row).forEach((ck) => {
          const C = +ck;
          const uc = row[ck];
          if (!uc) return;
          const cell = univerCellToSheetjs(uc);
          if (cell === null) return;
          ws[XLSX.utils.encode_cell({ r: R, c: C })] = cell;
          if (R > maxR) maxR = R;
          if (C > maxC) maxC = C;
        });
      });

      ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });

      // Anchos de columna
      const cols = [];
      const cd = sh.columnData || {};
      Object.keys(cd).forEach((ck) => {
        const c = +ck;
        if (cd[ck] && cd[ck].w) cols[c] = { wpx: cd[ck].w };
        if (cd[ck] && cd[ck].hd) cols[c] = Object.assign(cols[c] || {}, { hidden: true });
      });
      if (cols.length) ws["!cols"] = cols;

      // Altos de fila. Univer guarda el alto en PÍXELES; OOXML lo espera en
      // PUNTOS (1pt = 4/3 px). SheetJS escribe hpt tal cual, así que convertimos
      // px -> pt para no inflar la altura en cada ciclo guardar/reabrir.
      //
      // ALTURA EFECTIVA: Univer muestra la fila con esta fórmula (misma que su
      // getRowHeight): si ia (isAutoHeight) es null/true y hay ah (auto-height),
      // usa ah; si no, usa h. El pincel "copiar formato" recalcula ah pero deja
      // h intacto; guardar solo h perdía la altura que se veía en pantalla.
      // Aquí NO recalculamos nada: leemos la altura que Univer YA está mostrando.
      const rows = [];
      const rd = sh.rowData || {};
      Object.keys(rd).forEach((rk) => {
        const r = +rk;
        const row = rd[rk];
        if (!row) return;
        const useAuto = (row.ia == null || row.ia === 1) && typeof row.ah === "number";
        const px = useAuto ? row.ah : row.h;
        if (px) rows[r] = { hpt: Math.round((px * 3 / 4) * 100) / 100 };
        if (row.hd) rows[r] = Object.assign(rows[r] || {}, { hidden: true });
      });
      if (rows.length) ws["!rows"] = rows;

      // Combinadas
      if (sh.mergeData && sh.mergeData.length) {
        ws["!merges"] = sh.mergeData.map((m) => ({
          s: { r: m.startRow, c: m.startColumn },
          e: { r: m.endRow, c: m.endColumn },
        }));
      }

      let sheetName = (sh.name || sid).substring(0, 31).replace(/[\\\/\?\*\[\]:]/g, "_");
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    if (!wb.SheetNames.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[""]]), "Sheet1");
    }
    return wb;
  }

  function univerCellToSheetjs(uc) {
    const cell = {};
    let has = false;

    // Fórmula
    if (uc.f) { cell.f = String(uc.f).replace(/^=/, ""); has = true; }

    // Valor
    if (uc.v !== undefined && uc.v !== null && uc.v !== "") {
      if (uc.t === 2 || typeof uc.v === "number") { cell.t = "n"; cell.v = Number(uc.v); }
      else if (uc.t === 3) { cell.t = "b"; cell.v = !!uc.v; }
      else { cell.t = "s"; cell.v = String(uc.v); }
      has = true;
    } else if (cell.f) {
      cell.t = "n"; cell.v = 0;  // placeholder para celda solo-fórmula
    }

    // Formato numérico -> z
    if (uc.s && uc.s.n && uc.s.n.pattern) { cell.z = uc.s.n.pattern; has = true; }

    return has ? cell : null;
  }

  // ===========================================================================
  //  Libro nuevo en blanco
  // ===========================================================================
  let blankCounter = 0;

  // Construye el snapshot de un libro vacío (una hoja "Hoja1").
  function blankWorkbookData(name) {
    const sheetId = "sheet-0";
    return {
      id: "wb-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
      name: name || "Libro",
      appVersion: "xlsview",
      locale: LocaleType.EN_US,
      sheetOrder: [sheetId],
      sheets: {
        [sheetId]: {
          id: sheetId,
          name: "Hoja1",
          rowCount: 100,
          columnCount: 26,
          cellData: {},
          columnData: {},
          rowData: {},
          mergeData: [],
          defaultColumnWidth: 88,
          defaultRowHeight: 22,
        },
      },
      __protection: {},
    };
  }

  // Crea un libro nuevo en blanco en una pestaña propia y la activa.
  function newBlankWorkbook() {
    blankCounter += 1;
    const name = "Libro " + blankCounter + ".xlsx";
    const uniData = blankWorkbookData(name);
    const tab = {
      id: "tab-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
      name,
      token: null,                 // sin archivo en disco: al guardar irá a "Guardar como"
      unitId: null,
      snapshot: uniData,
      dirty: false,
      protection: {},
      pageConfig: normalizePageConfig(null, ["Hoja1"]),
      isNew: true,                 // marca de libro sin guardar aún
    };
    tabs.push(tab);
    renderTabs();
    activateTab(tab);
    setWelcome(false);
    updateToolbar();
    return tab;
  }

  // ===========================================================================
  //  Carga / apertura de archivos
  // ===========================================================================
  async function openFromToken(token, displayName) {
    // Ya abierto: activar su pestaña.
    const existing = tabs.find((t) => t.token === token);
    if (existing) { activateTab(existing); return; }
    try {
      showLoading("Abriendo " + (displayName || "") + "…");
      const res = await fetch("/xls/" + token, { cache: "no-store" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const buf = await res.arrayBuffer();
      const name = displayName || decodeURIComponent(token.split("/").pop() || "archivo");
      await openArrayBuffer(buf, name, token);
    } catch (e) {
      toast("No se pudo abrir el archivo: " + e.message, "err");
    } finally {
      hideLoading();
      signalRendered();
    }
  }

  async function openArrayBuffer(buf, name, token) {
    showLoading("Procesando " + name + "…");
    // bookFiles:true conserva el XML interno para leer estilos ricos (fuentes,
    // bordes, alineación) que la edición community de SheetJS no expone.
    const wb = XLSX.read(buf, {
      type: "array",
      cellStyles: true,
      cellNF: true,
      cellDates: true,
      cellFormula: true,
      bookFiles: true,
    });

    // Leer los estilos completos del archivo original (solo xlsx/xlsm; el CSV y
    // el xls antiguo no tienen styles.xml y se ignoran silenciosamente).
    let richStyles = null, richPanes = null, richProtection = null, richPage = null;
    try {
      if (window.XlsxStyles && wb.files) {
        const built = window.XlsxStyles.build(wb);
        if (built.ok) {
          richStyles = built.sheetsByName;
          richPanes = built.panesByName;
          richProtection = built.protectionByName;
          richPage = built.pageByName;
        }
      }
    } catch (e) { /* si falla el parser, seguimos con el estilo básico */ }

    const uniData = sheetjsToUniver(wb, name, richStyles, richPanes, richProtection);
    // Config de página por nombre de hoja (para impresión). Normalizamos a un
    // objeto por defecto si el archivo no la trae.
    const pageConfig = normalizePageConfig(richPage, wb.SheetNames);

    // NO montamos todos los workbooks a la vez (Univer los superpondría en el
    // mismo contenedor). Cada pestaña guarda su SNAPSHOT; solo hay un workbook
    // montado en Univer a la vez. Al activar una pestaña se monta el suyo.
    const tab = {
      id: "tab-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
      name,
      token: token || null,
      unitId: null,               // id del workbook montado (null si no montado)
      snapshot: uniData,          // datos para (re)montar en Univer
      dirty: false,
      // Protección por id de hoja Univer: { "sheet-0": { enabled, opts, locked, name } }
      protection: uniData.__protection || {},
      // Config de página por nombre de hoja (impresión).
      pageConfig: pageConfig,
    };
    tabs.push(tab);
    renderTabs();
    activateTab(tab);             // desmonta el anterior y monta este
    setWelcome(false);
    updateToolbar();
  }

  // Monta el workbook de una pestaña en Univer desde su snapshot.
  function mountTab(t) {
    if (t.unitId) return;         // ya montado
    suppressDirty = true;
    const wbApi = univerAPI.createWorkbook(t.snapshot);
    t.unitId = (wbApi && wbApi.getId) ? wbApi.getId() : t.snapshot.id;
    // Las funciones propias deben registrarse con un workbook ya existente
    // (si se registran en el arranque, sin libro, se pierden). Idempotente.
    registerCustomFunctions();
    // Reactivar el seguimiento de cambios tras asentar el render.
    setTimeout(() => { suppressDirty = false; }, 500);
  }

  // Desmonta el workbook de una pestaña, guardando antes su estado editado.
  function unmountTab(t) {
    if (!t.unitId) return;
    try {
      // Preservar las ediciones: recuperar el snapshot vigente antes de disponer.
      const wb = univerAPI.getActiveWorkbook && univerAPI.getActiveWorkbook();
      if (wb && wb.save) {
        const snap = wb.save();
        if (snap) {
          // Conservar la protección (no cabe en el snapshot de Univer).
          snap.__protection = t.snapshot.__protection;
          t.snapshot = snap;
        }
      }
    } catch (e) { /* si falla, conservamos el snapshot previo */ }
    try {
      if (univerAPI.disposeUnit) univerAPI.disposeUnit(t.unitId);
    } catch (e) { /* no-op */ }
    t.unitId = null;
  }

  // ===========================================================================
  //  Guardado
  // ===========================================================================
  function currentSnapshot(tab) {
    // Obtener el snapshot del workbook activo desde Univer.
    const wb = univerAPI.getActiveWorkbook
      ? univerAPI.getActiveWorkbook()
      : null;
    let snap = null;
    if (wb && wb.save) snap = wb.save();
    else if (wb && wb.getSnapshot) snap = wb.getSnapshot();
    if (!snap) throw new Error("No se pudo obtener el contenido de la hoja.");
    resolveSnapshotStyles(snap);
    // NO se recalculan ni comparan alturas: el snapshot de Univer YA contiene lo
    // que el usuario modificó (datos, estilos, alturas, anchos). Se guarda tal
    // cual — si la altura es X, se persiste X y se recupera X. Cualquier lógica
    // de "altura efectiva/render" quedó eliminada a pedido explícito del usuario.
    return snap;
  }

  // Univer guarda los estilos editados en una TABLA (snap.styles) y deja en la
  // celda un string-id (cell.s = "eQAnLq") en vez del objeto inline. Aquí
  // normalizamos: reemplazamos cada string-id por su objeto de estilo, para que
  // la serialización y la impresión (que esperan estilo inline) los conserven.
  function resolveSnapshotStyles(snap) {
    const table = snap.styles;
    if (!table) return snap;
    for (const sid in (snap.sheets || {})) {
      const sheet = snap.sheets[sid];
      const cd = sheet && sheet.cellData;
      if (!cd) continue;
      for (const r in cd) {
        const row = cd[r];
        for (const c in row) {
          const cell = row[c];
          if (cell && typeof cell.s === "string") {
            const resolved = table[cell.s];
            cell.s = resolved ? resolved : undefined;   // objeto inline (o sin estilo)
          }
        }
      }
    }
    return snap;
  }

  function bookTypeFor(name) {
    const ext = (name.split(".").pop() || "xlsx").toLowerCase();
    if (ext === "csv") return "csv";
    if (ext === "xls") return "xls";
    if (ext === "xlsm") return "xlsm";
    return "xlsx";
  }

  function serialize(tab) {
    const snap = currentSnapshot(tab);
    const wb = univerToSheetjs(snap);
    const bookType = bookTypeFor(tab.name);
    let out = XLSX.write(wb, { bookType, type: "array", cellStyles: true });

    // SheetJS descarta fuentes/bordes/alineación al escribir; reinyectamos los
    // estilos completos de Univer en el ZIP (solo formatos OOXML: xlsx/xlsm).
    // Pasamos también la protección (por id de hoja) para preservarla al guardar.
    if ((bookType === "xlsx" || bookType === "xlsm") && window.XlsxStyles && window.XlsxStyles.apply) {
      try {
        const styled = window.XlsxStyles.apply(out, snap, {
          protection: tab.protection || {},
          page: tab.pageConfig || {},   // config de página por nombre de hoja
        });
        if (styled && styled.length) out = styled;
      } catch (e) { /* si falla, conservamos el archivo sin estilos ricos */ }
    }
    return { data: out, bookType };
  }

  async function saveActive() {
    if (!activeTab) return;
    try {
      showLoading("Guardando…");
      const { data } = serialize(activeTab);

      if (inHostApp && activeTab.token) {
        // Guardar en disco a través del host.
        const res = await fetch("/save/" + activeTab.token, {
          method: "POST",
          headers: { "Content-Type": "application/octet-stream" },
          body: data,
        });
        if (!res.ok) throw new Error("HTTP " + res.status);
        markDirty(activeTab, false);
        toast("Guardado ✓", "ok");
      } else if (inHostApp) {
        // En la app pero sin ruta conocida (archivo soltado por arrastre):
        // delegar en "Guardar como" para elegir dónde escribir en disco.
        hideLoading();
        await saveActiveAs();
        return;
      } else {
        // Navegador: descargar.
        downloadBlob(data, activeTab.name);
        markDirty(activeTab, false);
        toast("Archivo descargado", "ok");
      }
    } catch (e) {
      toast("Error al guardar: " + e.message, "err");
    } finally {
      hideLoading();
    }
  }

  async function saveActiveAs() {
    if (!activeTab) return;
    try {
      showLoading("Guardando como…");
      const { data } = serialize(activeTab);
      if (inHostApp) {
        const res = await fetch("/saveas?name=" + encodeURIComponent(activeTab.name), {
          method: "POST",
          headers: { "Content-Type": "application/octet-stream" },
          body: data,
        });
        if (res.ok) {
          const j = await res.json().catch(() => ({}));
          if (j && j.saved) {
            if (j.name) { activeTab.name = j.name; renderTabs(); updateToolbar(); }
            if (j.token) activeTab.token = j.token;
            markDirty(activeTab, false);
            toast("Guardado como " + (j.name || "") + " ✓", "ok");
          } else {
            toast("Guardar como: cancelado");
          }
        } else {
          throw new Error("HTTP " + res.status);
        }
      } else {
        downloadBlob(data, activeTab.name);
        toast("Archivo descargado", "ok");
      }
    } catch (e) {
      toast("Error: " + e.message, "err");
    } finally {
      hideLoading();
    }
  }

  function downloadBlob(data, name) {
    const blob = new Blob([data], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  // ===========================================================================
  //  Pestañas
  // ===========================================================================
  function renderTabs() {
    const bar = $("tabbar");
    bar.innerHTML = "";
    tabs.forEach((t) => {
      const el = document.createElement("div");
      el.className = "tab" + (t === activeTab ? " active" : "");
      el.title = t.name;
      el.innerHTML =
        (t.dirty ? '<span class="dot">•</span>' : "") +
        '<span class="label"></span>' +
        '<span class="close" title="Cerrar">✕</span>';
      el.querySelector(".label").textContent = t.name;
      el.addEventListener("mousedown", (e) => {
        if (e.button === 1) { e.preventDefault(); closeTab(t); return; }  // clic central
      });
      el.addEventListener("click", (e) => {
        if (e.target.classList.contains("close")) { closeTab(t); return; }
        activateTab(t);
      });
      bar.appendChild(el);
    });
  }

  function activateTab(t) {
    if (activeTab === t && t.unitId) {
      // Ya activa y montada: solo refrescar UI.
      renderTabs(); setWelcome(false); updateToolbar();
      return;
    }
    // Desmontar el workbook de la pestaña anterior (para no superponer renders).
    if (activeTab && activeTab !== t) unmountTab(activeTab);
    activeTab = t;
    mountTab(t);                  // montar el de esta pestaña
    renderTabs();
    setWelcome(false);
    updateToolbar();
  }

  function closeTab(t) {
    if (t.dirty) {
      const ok = confirm('"' + t.name + '" tiene cambios sin guardar.\n¿Cerrar de todos modos?');
      if (!ok) return;
    }
    const idx = tabs.indexOf(t);
    if (idx < 0) return;
    // Desmontar/eliminar el workbook de Univer (libera memoria y su render).
    if (t.unitId) {
      try { if (univerAPI.disposeUnit) univerAPI.disposeUnit(t.unitId); } catch (e) {}
      t.unitId = null;
    }

    const wasActive = activeTab === t;
    tabs.splice(idx, 1);
    if (wasActive) {
      activeTab = null;                          // evita que activateTab intente desmontar 't'
      const next = tabs[idx] || tabs[idx - 1] || null;
      if (next) activateTab(next);
      else { setWelcome(true); }
    }
    renderTabs();
    updateToolbar();
  }

  function markDirty(t, dirty) {
    if (t.dirty === dirty) return;
    t.dirty = dirty;
    renderTabs();
    updateToolbar();
  }

  // ===========================================================================
  //  Toolbar / estado de UI
  // ===========================================================================
  function updateToolbar() {
    const has = !!activeTab;
    $("btnSave").disabled = !has;
    $("btnSaveAs").disabled = !has;
    $("btnZoomIn").disabled = !has;
    $("btnZoomOut").disabled = !has;
    $("btnPrint").disabled = !has;
    $("btnPageSetup").disabled = !has;
    $("btnMacro").disabled = !has;
    $("curName").textContent = has
      ? activeTab.name + (activeTab.dirty ? "  •" : "")
      : "— sin archivo —";
    document.title = has ? activeTab.name + " — XlsView" : "XlsView";
  }

  // ===========================================================================
  //  Impresión / configuración de página
  // ===========================================================================

  // Nombre de la hoja activa en Univer (para indexar pageConfig por nombre).
  function activeSheetInfo() {
    try {
      const wb = univerAPI.getActiveWorkbook && univerAPI.getActiveWorkbook();
      const snap = wb && wb.save ? wb.save() : null;
      if (!snap) return null;
      // Determinar la hoja activa: Univer marca sheetOrder[0] o hay un current.
      let sheetId = null;
      try {
        const ws = wb.getActiveSheet && wb.getActiveSheet();
        if (ws && ws.getSheetId) sheetId = ws.getSheetId();
      } catch (e) {}
      if (!sheetId) sheetId = (snap.sheetOrder && snap.sheetOrder[0]) || Object.keys(snap.sheets)[0];
      const sheet = snap.sheets[sheetId];
      return { snap, sheetId, sheetName: sheet && sheet.name };
    } catch (e) { return null; }
  }

  // Devuelve la config de página de la hoja activa (crea una por defecto si falta).
  function currentPage() {
    const info = activeSheetInfo();
    if (!info || !activeTab) return null;
    const name = info.sheetName || info.sheetId;
    if (!activeTab.pageConfig) activeTab.pageConfig = {};
    if (!activeTab.pageConfig[name]) activeTab.pageConfig[name] = defaultPage();
    return { page: activeTab.pageConfig[name], info: info, name: name };
  }

  // --- Modal de configuración de página ---
  function openPageSetup() {
    const cp = currentPage();
    if (!cp) return;
    loadPageForm(cp.page);
    $("pageModal").classList.add("open");
  }

  function loadPageForm(p) {
    $("pgOrientation").value = p.setup.orientation || "portrait";
    $("pgPaper").value = p.setup.paperSize || "9";
    $("pgMTop").value = p.margins.top;
    $("pgMBottom").value = p.margins.bottom;
    $("pgMLeft").value = p.margins.left;
    $("pgMRight").value = p.margins.right;
    const fit = !!p.setup.fitToPage;
    $("pgModeFit").checked = fit;
    $("pgModeScale").checked = !fit;
    $("pgScale").value = p.setup.scale || 100;
    $("pgFitW").value = p.setup.fitToWidth != null ? p.setup.fitToWidth : 1;
    $("pgFitH").value = p.setup.fitToHeight != null ? p.setup.fitToHeight : 0;
    $("pgPrintArea").value = cleanRef(p.printArea) || "";
    $("pgRepeatRows").value = p.repeatRows ? (p.repeatRows[0] + 1) + ":" + (p.repeatRows[1] + 1) : "";
    $("pgRepeatCols").value = p.repeatCols ? colName(p.repeatCols[0]) + ":" + colName(p.repeatCols[1]) : "";
    $("pgHeader").value = p.header || "";
    $("pgFooter").value = p.footer || "";
    $("pgGridLines").checked = !!p.options.gridLines;
    $("pgHeadings").checked = !!p.options.headings;
    $("pgHCenter").checked = !!p.options.horizontalCentered;
    $("pgVCenter").checked = !!p.options.verticalCentered;
  }

  function readPageForm(p) {
    p.setup.orientation = $("pgOrientation").value;
    p.setup.paperSize = $("pgPaper").value;
    p.margins.top = parseFloat($("pgMTop").value) || 0;
    p.margins.bottom = parseFloat($("pgMBottom").value) || 0;
    p.margins.left = parseFloat($("pgMLeft").value) || 0;
    p.margins.right = parseFloat($("pgMRight").value) || 0;
    p.setup.fitToPage = $("pgModeFit").checked;
    p.setup.scale = parseInt($("pgScale").value, 10) || 100;
    p.setup.fitToWidth = parseInt($("pgFitW").value, 10) || 1;
    p.setup.fitToHeight = parseInt($("pgFitH").value, 10) || 0;
    p.printArea = $("pgPrintArea").value.trim() || null;
    p.repeatRows = parseRowSpec($("pgRepeatRows").value);
    p.repeatCols = parseColSpec($("pgRepeatCols").value);
    p.header = $("pgHeader").value;
    p.footer = $("pgFooter").value;
    p.options.gridLines = $("pgGridLines").checked;
    p.options.headings = $("pgHeadings").checked;
    p.options.horizontalCentered = $("pgHCenter").checked;
    p.options.verticalCentered = $("pgVCenter").checked;
    return p;
  }

  function cleanRef(ref) { return ref ? String(ref).split("!").pop().replace(/\$/g, "") : ""; }
  function colName(c) { let s = ""; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function colIdx(s) { let c = 0; for (let i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64); return c - 1; }
  function parseRowSpec(v) { const m = /^(\d+):(\d+)$/.exec((v || "").trim()); return m ? [(+m[1]) - 1, (+m[2]) - 1] : null; }
  function parseColSpec(v) { const m = /^([A-Za-z]+):([A-Za-z]+)$/.exec((v || "").trim()); return m ? [colIdx(m[1].toUpperCase()), colIdx(m[2].toUpperCase())] : null; }

  function applyPageSetup() {
    const cp = currentPage();
    if (!cp) return;
    readPageForm(cp.page);
    markDirty(activeTab, true);   // cambió la configuración de página
    $("pageModal").classList.remove("open");
    toast("Configuración de página aplicada", "ok");
  }

  function openPreview() {
    const cp = currentPage();
    if (!cp) return;
    // Tomar el snapshot MÁS reciente (con ediciones) para imprimir.
    const info = activeSheetInfo();
    if (!info) { toast("No hay hoja para imprimir", "err"); return; }
    window.XlsxPrint.preview(info.snap, info.sheetId, cp.page, { fileName: activeTab.name });
  }

  function wirePrintUI() {
    $("btnPrint").addEventListener("click", openPreview);
    $("btnPageSetup").addEventListener("click", openPageSetup);
    $("pgClose").addEventListener("click", () => $("pageModal").classList.remove("open"));
    $("pgCancel").addEventListener("click", () => $("pageModal").classList.remove("open"));
    $("pgApply").addEventListener("click", applyPageSetup);
    $("pgPreview").addEventListener("click", () => {
      const cp = currentPage(); if (cp) readPageForm(cp.page);
      $("pageModal").classList.remove("open");
      openPreview();
    });
    $("ppPrint").addEventListener("click", () => window.XlsxPrint.print());
    $("ppClose").addEventListener("click", () => window.XlsxPrint.close());
    $("ppSetup").addEventListener("click", () => { window.XlsxPrint.close(); openPageSetup(); });
  }

  function zoom(delta) {
    try {
      const wb = univerAPI.getActiveWorkbook && univerAPI.getActiveWorkbook();
      const ws = wb && wb.getActiveSheet && wb.getActiveSheet();
      if (ws && ws.getZoomRatio && ws.setZoomRatio) {
        let z = ws.getZoomRatio() || 1;
        z = Math.min(4, Math.max(0.3, z + delta));
        ws.setZoomRatio(z);
      }
    } catch (e) { /* no-op */ }
  }

  // ===========================================================================
  //  Arrastre de ventana (asa de la toolbar) — igual que pdfview
  // ===========================================================================
  function wireDrag() {
    if (!inHostApp) return;
    const handle = $("dragHandle");
    let dragging = false;
    handle.addEventListener("mousedown", (e) => {
      if (e.button !== 0) return;
      dragging = true;
      postToHost("dragstart:" + Math.round(e.screenX) + "," + Math.round(e.screenY));
      e.preventDefault();
    });
    window.addEventListener("mousemove", (e) => {
      if (!dragging) return;
      postToHost("dragmove:" + Math.round(e.screenX) + "," + Math.round(e.screenY));
    });
    window.addEventListener("mouseup", () => {
      if (!dragging) return;
      dragging = false;
      postToHost("dragend");
    });
    // Doble clic en el asa = pantalla completa
    handle.addEventListener("dblclick", () => postToHost("fullscreen"));
  }

  // ===========================================================================
  //  Sondeo de /pending (archivos abiertos tras arrancar la app)
  // ===========================================================================
  function startPendingPoll() {
    if (!inHostApp) return;
    async function poll() {
      try {
        const res = await fetch("/pending", { cache: "no-store" });
        if (res.ok) {
          const list = await res.json();
          for (const path of list) {
            // path viene como "/xls/<token>"
            const token = path.replace(/^\/xls\//, "");
            await openFromToken(token);
          }
        }
      } catch (e) { /* servidor puede no exponer /pending */ }
      setTimeout(poll, 800);
    }
    setTimeout(poll, 800);
  }

  // Avisar al host que ya pintamos (cierra el splash nativo).
  let renderedSent = false;
  function signalRendered() {
    if (renderedSent) return;
    renderedSent = true;
    postToHost("rendered");
  }

  // ===========================================================================
  //  Wiring de botones y teclado
  // ===========================================================================
  //  Macros — scripts JavaScript del usuario que automatizan tareas sobre la
  //  hoja activa mediante la Facade de Univer. Se guardan en localStorage.
  // ===========================================================================
  const MACRO_STORE = "xlsview.macros";
  let macros = [];            // { name, code }
  let macroSel = -1;          // índice seleccionado

  const MACRO_EXAMPLES = {
    numerar:
      "// Numera la primera columna a partir de la fila seleccionada.\n" +
      "// Ajusta 'desde' y cuántas filas quieres numerar.\n" +
      "const desde = 8;      // fila 9 (0-based)\n" +
      "const cuantas = 12;\n" +
      "for (let i = 0; i < cuantas; i++) {\n" +
      "  cell(desde + i, 0).setValue(i + 1);\n" +
      "}\n" +
      "toast('Numeradas ' + cuantas + ' filas');",
    totales:
      "// Suma cada columna numérica de un rango y escribe una fila de totales.\n" +
      "const filaIni = 8, filaFin = 20;        // filas de datos (0-based)\n" +
      "const colIni = 3, colFin = 6;           // columnas D..G\n" +
      "const filaTotal = filaFin + 1;\n" +
      "for (let c = colIni; c <= colFin; c++) {\n" +
      "  const letra = String.fromCharCode(65 + c);\n" +
      "  cell(filaTotal, c).setValue('=SUM(' + letra + (filaIni+1) + ':' + letra + (filaFin+1) + ')');\n" +
      "}\n" +
      "cell(filaTotal, colIni - 1).setValue('TOTAL');\n" +
      "toast('Fila de totales agregada');",
    limpiar:
      "// Quita el formato (mantiene los valores) del rango seleccionado.\n" +
      "const sel = sheet.getActiveRange ? sheet.getActiveRange() : null;\n" +
      "if (sel && sel.setFontWeight) {\n" +
      "  sel.setFontWeight('normal').setFontStyle('normal');\n" +
      "  if (sel.setBackground) sel.setBackground(null);\n" +
      "  toast('Formato limpiado');\n" +
      "} else {\n" +
      "  toast('Selecciona un rango primero');\n" +
      "}",
    resaltar:
      "// Pinta de rojo el texto de las celdas con números negativos\n" +
      "// dentro del rango de datos.\n" +
      "const f0 = 8, f1 = 30, c0 = 3, c1 = 7;\n" +
      "for (let r = f0; r <= f1; r++) {\n" +
      "  for (let c = c0; c <= c1; c++) {\n" +
      "    const v = cell(r, c).getValue();\n" +
      "    if (typeof v === 'number' && v < 0) cell(r, c).setFontColor('#c0392b');\n" +
      "  }\n" +
      "}\n" +
      "toast('Negativos resaltados');",
  };

  // Macros de ejemplo que se SIEMBRAN la primera vez (aparecen ya guardadas en
  // el botón ⚡ Macro). Coinciden con la hoja "Macros" del libro de ejemplos.
  const MACRO_SEED = [
    { name: "Numerar filas", code: MACRO_EXAMPLES.numerar },
    { name: "Fila de totales", code: MACRO_EXAMPLES.totales },
    { name: "Resaltar negativos", code: MACRO_EXAMPLES.resaltar },
    { name: "Limpiar formato", code: MACRO_EXAMPLES.limpiar },
  ];

  function loadMacros() {
    try { macros = JSON.parse(localStorage.getItem(MACRO_STORE) || "[]"); }
    catch (e) { macros = []; }
    if (!Array.isArray(macros)) macros = [];
    // Primera vez (nunca se guardó nada): sembrar las macros de ejemplo.
    if (localStorage.getItem(MACRO_STORE) == null) {
      macros = MACRO_SEED.map((m) => ({ name: m.name, code: m.code }));
      saveMacros();
    }
  }
  function saveMacros() {
    try { localStorage.setItem(MACRO_STORE, JSON.stringify(macros)); } catch (e) {}
  }

  function renderMacroList() {
    const ul = $("mcList");
    ul.innerHTML = "";
    if (!macros.length) {
      const li = document.createElement("li");
      li.className = "empty";
      li.textContent = "Sin macros. Crea una con ＋ o parte de un ejemplo.";
      ul.appendChild(li);
      return;
    }
    macros.forEach((m, i) => {
      const li = document.createElement("li");
      li.className = i === macroSel ? "active" : "";
      li.textContent = "⚡ " + (m.name || "(sin nombre)");
      li.title = m.name;
      li.addEventListener("click", () => selectMacro(i));
      ul.appendChild(li);
    });
  }

  function selectMacro(i) {
    macroSel = i;
    const m = macros[i];
    if (m) { $("mcName").value = m.name || ""; $("mcCode").value = m.code || ""; }
    renderMacroList();
  }

  function newMacro() {
    macroSel = -1;
    $("mcName").value = "";
    $("mcCode").value = "";
    renderMacroList();
    $("mcName").focus();
  }

  function saveCurrentMacro() {
    const name = $("mcName").value.trim();
    const code = $("mcCode").value;
    if (!name) { toast("Ponle un nombre a la macro", "err"); return; }
    const existing = macros.findIndex((m) => m.name === name);
    if (macroSel >= 0 && macros[macroSel]) {
      macros[macroSel] = { name, code };
    } else if (existing >= 0) {
      macros[existing] = { name, code };
      macroSel = existing;
    } else {
      macros.push({ name, code });
      macroSel = macros.length - 1;
    }
    saveMacros();
    renderMacroList();
    toast("Macro guardada ✓", "ok");
  }

  function deleteCurrentMacro() {
    if (macroSel < 0 || !macros[macroSel]) { toast("No hay macro seleccionada"); return; }
    if (!confirm('¿Eliminar la macro "' + macros[macroSel].name + '"?')) return;
    macros.splice(macroSel, 1);
    saveMacros();
    newMacro();
    toast("Macro eliminada", "ok");
  }

  // Ejecuta el código de la macro con una API acotada sobre la hoja activa.
  function runMacroCode(code) {
    if (!activeTab || !univerAPI) { toast("Abre una hoja primero", "err"); return; }
    const workbook = univerAPI.getActiveWorkbook && univerAPI.getActiveWorkbook();
    const sheet = workbook && workbook.getActiveSheet && workbook.getActiveSheet();
    if (!sheet) { toast("No hay hoja activa", "err"); return; }

    // Helpers: cell(f,c) y range(f,c,nf,nc) devuelven FRange de Univer.
    const cell = (f, c) => sheet.getRange(f, c);
    const range = (f, c, nf, nc) => sheet.getRange(f, c, nf || 1, nc || 1);
    // Dimensiones de datos (para bucles)
    let rows = 0, cols = 0;
    try {
      const snap = currentSnapshot(activeTab);
      const sh = snap.sheets[(snap.sheetOrder || [])[0]];
      rows = sh ? sh.rowCount : 0; cols = sh ? sh.columnCount : 0;
    } catch (e) {}

    suppressDirty = false;
    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(
        "sheet", "cell", "range", "rows", "cols", "api", "workbook", "toast",
        '"use strict";\n' + code
      );
      fn(sheet, cell, range, rows, cols, univerAPI, workbook, (m) => toast(String(m), "ok"));
      if (activeTab) markDirty(activeTab, true);
      toast("Macro ejecutada ✓", "ok");
    } catch (e) {
      toast("Error en la macro: " + e.message, "err");
      // También lo dejamos en consola para depurar.
      try { console.error("[macro]", e); } catch (_) {}
    }
  }

  function openMacroModal() {
    if (!activeTab) return;
    loadMacros();
    renderMacroList();
    if (macros.length && macroSel < 0) selectMacro(0);
    else if (!macros.length) newMacro();
    $("macroModal").classList.add("open");
  }
  function closeMacroModal() { $("macroModal").classList.remove("open"); }

  function wireMacros() {
    $("btnMacro").addEventListener("click", openMacroModal);
    $("mcClose").addEventListener("click", closeMacroModal);
    $("mcCancel").addEventListener("click", closeMacroModal);
    $("mcNew").addEventListener("click", newMacro);
    $("mcSave").addEventListener("click", saveCurrentMacro);
    $("mcDelete").addEventListener("click", deleteCurrentMacro);
    $("mcRun").addEventListener("click", () => runMacroCode($("mcCode").value));
    document.querySelectorAll(".macro-ex").forEach((btn) => {
      btn.addEventListener("click", () => {
        const ex = MACRO_EXAMPLES[btn.dataset.ex];
        if (ex == null) return;
        macroSel = -1;
        if (!$("mcName").value.trim()) $("mcName").value = btn.textContent.trim();
        $("mcCode").value = ex;
        renderMacroList();
      });
    });
    // Tab inserta dos espacios en el editor en vez de saltar de foco.
    $("mcCode").addEventListener("keydown", (e) => {
      if (e.key === "Tab") {
        e.preventDefault();
        const ta = e.target, s = ta.selectionStart, en = ta.selectionEnd;
        ta.value = ta.value.slice(0, s) + "  " + ta.value.slice(en);
        ta.selectionStart = ta.selectionEnd = s + 2;
      }
    });
  }

  // ===========================================================================
  function wireUI() {
    $("btnNew").addEventListener("click", () => newBlankWorkbook());
    $("btnOpen").addEventListener("click", pickFile);
    $("btnOpenWelcome").addEventListener("click", pickFile);
    $("btnSave").addEventListener("click", saveActive);
    $("btnSaveAs").addEventListener("click", saveActiveAs);
    $("btnZoomIn").addEventListener("click", () => zoom(0.1));
    $("btnZoomOut").addEventListener("click", () => zoom(-0.1));
    wireMacros();
    $("btnMin").addEventListener("click", () => { if (!postToHost("minimize")) toast("Minimizar solo en la app"); });
    $("btnFull").addEventListener("click", () => postToHost("fullscreen"));
    $("btnClose").addEventListener("click", () => { if (!postToHost("close")) window.close(); });

    // Selector de archivo (navegador o "Abrir" dentro de la app)
    $("filePicker").addEventListener("change", async (e) => {
      const file = e.target.files && e.target.files[0];
      e.target.value = "";
      if (!file) return;
      const buf = await file.arrayBuffer();
      await openArrayBuffer(buf, file.name, null);
      hideLoading();
    });

    // Arrastrar y soltar
    const dz = document.body;
    ["dragenter", "dragover"].forEach((ev) =>
      dz.addEventListener(ev, (e) => { e.preventDefault(); $("welcome").classList.add("dragover"); }));
    ["dragleave", "drop"].forEach((ev) =>
      dz.addEventListener(ev, (e) => { e.preventDefault(); $("welcome").classList.remove("dragover"); }));
    dz.addEventListener("drop", async (e) => {
      const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (!file) return;
      const buf = await file.arrayBuffer();
      await openArrayBuffer(buf, file.name, null);
      hideLoading();
    });

    // Teclado
    window.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
        e.preventDefault();
        if (e.shiftKey) saveActiveAs(); else saveActive();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "n") {
        e.preventDefault(); newBlankWorkbook();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") {
        e.preventDefault(); pickFile();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "w") {
        e.preventDefault(); if (activeTab) closeTab(activeTab);
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "p") {
        e.preventDefault(); if (activeTab) openPreview();
      } else if (e.key === "Escape") {
        // Cerrar primero los overlays de impresión si están abiertos.
        if ($("printPreview").classList.contains("open")) { window.XlsxPrint.close(); return; }
        if ($("pageModal").classList.contains("open")) { $("pageModal").classList.remove("open"); return; }
      }
    });

    // Advertir cambios sin guardar al cerrar
    window.addEventListener("beforeunload", (e) => {
      if (tabs.some((t) => t.dirty)) { e.preventDefault(); e.returnValue = ""; }
    });
  }

  // Abrir un archivo. En la app de escritorio usamos el diálogo nativo del host
  // (que da la RUTA real, así el guardado sobrescribe en disco). En el navegador
  // usamos el selector HTML (que solo entrega el contenido, sin ruta -> descarga).
  async function pickFile() {
    if (inHostApp) {
      try {
        showLoading("Abriendo…");
        const res = await fetch("/openfile", { method: "POST" });
        const j = await res.json().catch(() => ({}));
        if (j && j.opened && j.token) {
          const token = j.token.replace(/^\/xls\//, "");
          await openFromToken(token, j.name);
        }
      } catch (e) {
        toast("No se pudo abrir: " + e.message, "err");
      } finally {
        hideLoading();
      }
      return;
    }
    $("filePicker").click();
  }

  // ===========================================================================
  //  Arranque
  // ===========================================================================
  function boot() {
    initUniver();
    wireUI();
    wireDrag();
    wirePrintUI();
    loadMacros();          // siembra las macros de ejemplo la primera vez
    setWelcome(true);
    updateToolbar();

    // Hook de diagnóstico (para depuración con DevTools).
    window.__xlsview = {
      get univerAPI() { return univerAPI; },
      get activeTab() { return activeTab; },
      currentSnapshot: () => activeTab ? currentSnapshot(activeTab) : null,
      serialize: () => activeTab ? serialize(activeTab) : null,
      univerToSheetjs, currentPage,
    };

    // Apertura inicial por ?file=/xls/<token>
    const params = new URLSearchParams(location.search);
    const file = params.get("file");
    if (file) {
      const token = file.replace(/^\/xls\//, "");
      openFromToken(token).then(signalRendered);
    } else {
      // Sin archivo: arrancar con un libro nuevo en blanco (no pantalla vacía).
      newBlankWorkbook();
      setTimeout(signalRendered, 300);
    }

    startPendingPoll();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();

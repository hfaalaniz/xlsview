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
    t.textContent = "";
    const span = document.createElement("span");
    span.className = "toast-msg";
    span.textContent = msg;
    t.appendChild(span);

    // Los errores llevan un botón "Copiar" para reportar el mensaje si hace falta.
    if (kind === "err") {
      const btn = document.createElement("button");
      btn.className = "toast-copy";
      btn.type = "button";
      btn.title = "Copiar el mensaje de error";
      btn.innerHTML = "📋 Copiar";
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        copyErrorText(String(msg), btn);
      });
      t.appendChild(btn);
    }

    t.className = "show" + (kind ? " " + kind : "");
    clearTimeout(toast._t);
    // Los toasts de error permanecen 4 s visibles; el resto, 2.6 s.
    const ms = kind === "err" ? 4000 : 2600;
    toast._t = setTimeout(() => { t.className = ""; }, ms);
  }

  // Copia el texto del error al portapapeles y confirma en el propio botón.
  // Usa la API async si está disponible; si no, un fallback con execCommand.
  function copyErrorText(text, btn) {
    const done = () => {
      const prev = btn.innerHTML;
      btn.innerHTML = "✓ Copiado";
      // Mantener el toast visible unos segundos más tras copiar.
      clearTimeout(toast._t);
      toast._t = setTimeout(() => { $("toast").className = ""; }, 3000);
      setTimeout(() => { btn.innerHTML = prev; }, 1500);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done).catch(() => fallbackCopy(text, done));
    } else {
      fallbackCopy(text, done);
    }
  }

  function fallbackCopy(text, onDone) {
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
      if (onDone) onDone();
    } catch (e) { /* si no se puede copiar, no hacemos nada */ }
  }

  // ---------------------------------------------------------------------------
  //  Diálogo personalizado (reemplaza alert/confirm/prompt nativos).
  //  Todas las funciones devuelven una Promise.
  // ---------------------------------------------------------------------------
  const dlg = (function () {
    let resolver = null;

    function close(result) {
      $("dlgOverlay").classList.remove("open");
      const r = resolver; resolver = null;
      if (r) r(result);
    }

    // opts: { title, message, icon, html, input, buttons:[{label,value,variant,default}] }
    function show(opts) {
      return new Promise((resolve) => {
        // Si ya hay un diálogo abierto, lo cerramos como cancelado.
        if (resolver) { const r = resolver; resolver = null; r(undefined); }
        resolver = resolve;

        $("dlgIcon").textContent = opts.icon || "❓";
        $("dlgTitle").textContent = opts.title || "XlsView";
        const msg = $("dlgMsg");
        if (opts.html) msg.innerHTML = opts.html; else msg.textContent = opts.message || "";

        const inp = $("dlgInput");
        if (opts.input != null) {
          inp.style.display = "";
          inp.value = String(opts.input);
        } else {
          inp.style.display = "none";
          inp.value = "";
        }

        const box = $("dlgButtons");
        box.innerHTML = "";
        const buttons = opts.buttons || [{ label: "Aceptar", value: true, variant: "primary", default: true }];
        buttons.forEach((b) => {
          const el = document.createElement("button");
          el.textContent = b.label;
          if (b.variant) el.className = b.variant;
          el.addEventListener("click", () => {
            const val = opts.input != null && b.value !== false ? inp.value : b.value;
            close(val);
          });
          box.appendChild(el);
          if (b.default) setTimeout(() => el.focus(), 30);
        });

        $("dlgOverlay").classList.add("open");
        if (opts.input != null) setTimeout(() => { inp.focus(); inp.select(); }, 30);
      });
    }

    // Cierre por teclado: Esc = cancelar, Enter = botón por defecto.
    document.addEventListener("keydown", (e) => {
      if (!$("dlgOverlay").classList.contains("open")) return;
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(undefined); }
      else if (e.key === "Enter") {
        const def = $("dlgButtons").querySelector("button.primary") || $("dlgButtons").querySelector("button");
        if (def) { e.preventDefault(); def.click(); }
      }
    }, true);
    // Clic fuera del cuadro = cancelar.
    $("dlgOverlay").addEventListener("mousedown", (e) => {
      if (e.target === $("dlgOverlay")) close(undefined);
    });

    return {
      show,
      alert(message, opts) {
        return show(Object.assign({ icon: "ℹ️", title: "XlsView", message,
          buttons: [{ label: "Aceptar", value: true, variant: "primary", default: true }] }, opts || {}));
      },
      confirm(message, opts) {
        return show(Object.assign({ icon: "❓", title: "Confirmar", message,
          buttons: [
            { label: (opts && opts.cancelLabel) || "Cancelar", value: false },
            { label: (opts && opts.okLabel) || "Aceptar", value: true, variant: (opts && opts.danger) ? "danger" : "primary", default: true },
          ] }, opts || {}));
      },
      prompt(message, defaultValue, opts) {
        return show(Object.assign({ icon: "✏️", title: "XlsView", message, input: defaultValue || "",
          buttons: [
            { label: "Cancelar", value: false },
            { label: "Aceptar", variant: "primary", default: true },
          ] }, opts || {}));
      },
    };
  })();

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
  // Fusión profunda de objetos de locale (Univer usa objetos anidados).
  // No hay mergeLocales en los bundles UMD 0.5.x, así que lo hacemos a mano.
  function deepMergeLocale(target, source) {
    if (!source || typeof source !== "object") return target;
    for (const k of Object.keys(source)) {
      const sv = source[k];
      if (sv && typeof sv === "object" && !Array.isArray(sv)) {
        if (!target[k] || typeof target[k] !== "object") target[k] = {};
        deepMergeLocale(target[k], sv);
      } else {
        target[k] = sv;
      }
    }
    return target;
  }

  function initUniver() {
    // ------------------------------------------------------------------
    //  Plugins OSS adicionales (todos open source, sin licencia).
    //  Cada entrada: [globalUMD, nombrePlugin]. Se registran en orden de
    //  dependencia, tras el preset core que trae sus dependencias base.
    //  Si algún bundle no cargó, se omite sin romper la app.
    // ------------------------------------------------------------------
    const pluginSpecs = [
      // Validación de datos (listas, casillas, número, fecha…)
      ["UniverDataValidation", "UniverDataValidationPlugin"],
      ["UniverSheetsDataValidation", "UniverSheetsDataValidationPlugin"],
      ["UniverSheetsDataValidationUi", "UniverSheetsDataValidationUIPlugin"],
      // NOTA: el grupo "drawing" (imágenes) se omite adrede: el preset core 0.5.5
      // ya trae drawing-manager.service; re-registrarlo rompe la inicialización.
      // Comentarios / hilos (thread-comment)
      ["UniverThreadComment", "UniverThreadCommentPlugin"],
      ["UniverThreadCommentUi", "UniverThreadCommentUIPlugin"],
      ["UniverSheetsThreadComment", "UniverSheetsThreadCommentPlugin"],
      ["UniverSheetsThreadCommentUi", "UniverSheetsThreadCommentUIPlugin"],
      // Filtros
      ["UniverSheetsFilter", "UniverSheetsFilterPlugin"],
      ["UniverSheetsFilterUi", "UniverSheetsFilterUIPlugin"],
      // Ordenar
      ["UniverSheetsSort", "UniverSheetsSortPlugin"],
      ["UniverSheetsSortUi", "UniverSheetsSortUIPlugin"],
      // Buscar y reemplazar
      ["UniverFindReplace", "UniverFindReplacePlugin"],
      ["UniverSheetsFindReplace", "UniverSheetsFindReplacePlugin"],
      // Formato condicional
      ["UniverSheetsConditionalFormatting", "UniverSheetsConditionalFormattingPlugin"],
      ["UniverSheetsConditionalFormattingUi", "UniverSheetsConditionalFormattingUIPlugin"],
      // Hipervínculos
      ["UniverSheetsHyperLink", "UniverSheetsHyperLinkPlugin"],
      ["UniverSheetsHyperLinkUi", "UniverSheetsHyperLinkUIPlugin"],
    ];

    const ossPlugins = [];
    for (const [g, name] of pluginSpecs) {
      try {
        const mod = window[g];
        const plugin = mod && mod[name];
        if (plugin) ossPlugins.push(plugin);
        else console.warn("[XlsView] Plugin OSS no disponible:", g, name);
      } catch (e) { /* omitir el que falle */ }
    }

    // Combinar los textos (locales) de los presets adicionales sobre el core.
    // Cada global es un objeto de locale ya listo; se hace merge profundo.
    const mergedLocale = deepMergeLocale({}, localeEnUS);
    [
      window.UniverSheetsDataValidationUiEnUS,   // ya lo cargabas (univer-dv-locale)
      window.UniverThreadCommentUiEnUS,
      window.UniverSheetsFilterUiEnUS,
      window.UniverSheetsSortEnUS,
      window.UniverFindReplaceEnUS,
      window.UniverSheetsConditionalFormattingUiEnUS,
      window.UniverSheetsHyperLinkUiEnUS,
    ].forEach((loc) => { if (loc) deepMergeLocale(mergedLocale, loc); });

    // Traducción al español (propia). 0.5.5 no publica es-ES y su enum
    // LocaleType tampoco lo incluye, así que en vez de registrar un idioma
    // nuevo mergeamos el es-ES PARCIAL sobre el en-US ya combinado y seguimos
    // usando LocaleType.EN_US como clave. Las claves traducidas quedan en
    // español; las que falten conservan el inglés (fallback natural).
    if (window.UniverXlsViewEsES) deepMergeLocale(mergedLocale, window.UniverXlsViewEsES);

    const cfg = {
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: mergedLocale },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({ container: "univer-host" }),
      ],
    };
    if (ossPlugins.length) cfg.plugins = ossPlugins;
    const { univerAPI: api } = createUniver(cfg);
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
  //  WEBSERVICE — caché síncrona de descargas (el motor no admite async).
  //   estado: "loading" | "ok" | "error"
  // ---------------------------------------------------------------------------
  const wsCache = new Map();   // url -> { state, value }
  let wsRecalcPending = false;

  function webserviceSync(url) {
    if (!url) return "#N/A";
    if (!inHostApp) return "#N/A";        // sin host no hay acceso a red
    const hit = wsCache.get(url);
    if (hit) {
      if (hit.state === "ok") return hit.value;
      if (hit.state === "error") return "#VALUE!";
      return "Cargando…";                 // "loading": aún en curso
    }
    // Primera vez: marcar en curso y disparar la descarga en segundo plano.
    wsCache.set(url, { state: "loading", value: "" });
    fetch("/fetch?url=" + encodeURIComponent(url), { cache: "no-store" })
      .then((r) => r.ok ? r.json() : { ok: false })
      .then((j) => {
        if (j && j.ok) wsCache.set(url, { state: "ok", value: String(j.body) });
        else wsCache.set(url, { state: "error", value: "" });
        scheduleWebserviceRecalc();
      })
      .catch(() => { wsCache.set(url, { state: "error", value: "" }); scheduleWebserviceRecalc(); });
    return "Cargando…";
  }

  // Cuando llega una respuesta, refrescamos las celdas con WEBSERVICE. Univer
  // cachea el resultado de una función "pura", así que no la re-evalúa solo
  // porque cambió nuestro caché: hay que reescribir la fórmula de cada celda que
  // use WEBSERVICE para forzar su re-evaluación (ya con el valor en caché).
  function scheduleWebserviceRecalc() {
    if (wsRecalcPending) return;
    wsRecalcPending = true;
    setTimeout(() => {
      wsRecalcPending = false;
      try {
        const wb = univerAPI.getActiveWorkbook && univerAPI.getActiveWorkbook();
        if (!wb) return;
        const snap = wb.save ? wb.save() : null;
        if (!snap) return;
        const order = snap.sheetOrder || Object.keys(snap.sheets || {});
        suppressDirty = true;
        order.forEach((sid) => {
          const sh = snap.sheets[sid];
          if (!sh || !sh.cellData) return;
          const uni = univerAPI.getUniverSheet ? univerAPI.getUniverSheet(sid) : null;
          const sheetApi = wb.getSheetBySheetId ? wb.getSheetBySheetId(sid) : null;
          Object.keys(sh.cellData).forEach((rk) => {
            const row = sh.cellData[rk];
            Object.keys(row).forEach((ck) => {
              const cell = row[ck];
              const f = cell && cell.f;
              if (f && /WEBSERVICE\s*\(/i.test(f) && sheetApi) {
                // reescribir la misma fórmula -> Univer la re-evalúa
                try { sheetApi.getRange(+rk, +ck).setValue(f); } catch (e) {}
              }
            });
          });
        });
        setTimeout(() => { suppressDirty = false; }, 200);
      } catch (e) { /* si falla, editar la celda manualmente recalcula */ }
    }, 80);
  }

  // ---------------------------------------------------------------------------
  //  Banco de funciones propias (registerFunction de la Facade de Univer).
  //  Se comportan como funciones nativas: se usan en celdas y se recalculan.
  //  Añade las tuyas al array 'calculate' como [fn, "NOMBRE", "descripción"].
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

      // ===================== WEB / XML =====================
      // FILTERXML: aplica un XPath a un texto XML y devuelve el primer nodo
      // coincidente (local, sin red). Univer trae el nombre pero no el cálculo.
      [(xml, xpath) => {
        try {
          const doc = new DOMParser().parseFromString(String(xml), "text/xml");
          if (doc.getElementsByTagName("parsererror").length) return "#VALUE!";
          const res = doc.evaluate(String(xpath), doc, null,
            XPathResult.FIRST_ORDERED_NODE_TYPE, null);
          const node = res.singleNodeValue;
          if (!node) return "#N/A";
          return node.textContent != null ? node.textContent : String(node.nodeValue || "");
        } catch (e) { return "#VALUE!"; }
      }, "FILTERXML", "Extrae datos de un XML con XPath: FILTERXML(xml, xpath)"],

      // WEBSERVICE: trae el texto de una URL a través del host C# (ruta /fetch,
      // que sí tiene acceso a internet). El motor de Univer NO soporta funciones
      // asíncronas (una Promise se interpreta como matriz y da #SPILL!), así que
      // usamos un patrón de CACHÉ SÍNCRONA: la función devuelve el valor cacheado;
      // si aún no está, dispara la descarga en segundo plano y devuelve "cargando…";
      // al llegar la respuesta, forzamos un recálculo para que la celda se refresque.
      [(url) => webserviceSync(String(url == null ? "" : url)), "WEBSERVICE",
        "Devuelve el texto de una URL: WEBSERVICE(url)"],
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

  // Lee los gráficos de XlsView guardados dentro del ZIP (xl/xlsview-charts.json).
  function chartsFromWorkbook(wb) {
    try {
      const f = wb && wb.files && wb.files["xl/xlsview-charts.json"];
      if (!f) return null;
      const txt = typeof f.content === "string"
        ? f.content
        : new TextDecoder("utf-8").decode(new Uint8Array(f.content));
      const arr = JSON.parse(txt);
      return Array.isArray(arr) ? arr : null;
    } catch (e) { return null; }
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
      if (!res.ok) {
        let detail = "HTTP " + res.status;
        try { const j = await res.json(); if (j && j.error) detail += " — " + j.error; } catch (e) {}
        throw new Error(detail);
      }
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
    let richStyles = null, richPanes = null, richProtection = null, richPage = null, richValidations = null;
    try {
      if (window.XlsxStyles && wb.files) {
        const built = window.XlsxStyles.build(wb);
        if (built.ok) {
          richStyles = built.sheetsByName;
          richPanes = built.panesByName;
          richProtection = built.protectionByName;
          richPage = built.pageByName;
          richValidations = built.validationsByName;
        }
      }
    } catch (e) { /* si falla el parser, seguimos con el estilo básico */ }

    const uniData = sheetjsToUniver(wb, name, richStyles, richPanes, richProtection);

    // Validaciones de datos leídas del OOXML -> recurso del plugin de Univer.
    // sheetjsToUniver asigna sheet-0, sheet-1… en el orden de wb.SheetNames.
    try {
      if (richValidations && Object.keys(richValidations).length) {
        const bySheetId = {};
        wb.SheetNames.forEach((sn, idx) => {
          if (richValidations[sn]) bySheetId["sheet-" + idx] = richValidations[sn];
        });
        if (Object.keys(bySheetId).length) {
          uniData.resources = uniData.resources || [];
          uniData.resources.push({
            name: "SHEET_DATA_VALIDATION_PLUGIN",
            data: JSON.stringify(bySheetId),
          });
        }
      }
    } catch (e) { /* si falla, se abre sin validaciones */ }
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
      // Gráficos guardados (leídos del archivo si los hubiera).
      charts: chartsFromWorkbook(wb) || [],
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
          charts: tab.charts || [],     // gráficos de XlsView (json en el ZIP)
        });
        if (styled && styled.length) out = styled;
      } catch (e) { /* si falla, conservamos el archivo sin estilos ricos */ }
    }
    return { data: out, bookType };
  }

  // Devuelve true si se guardó, false si se canceló o falló.
  async function saveActive() {
    if (!activeTab) return false;
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
        return true;
      } else if (inHostApp) {
        // En la app pero sin ruta conocida (archivo nuevo o soltado por arrastre):
        // delegar en "Guardar como" para elegir dónde escribir en disco.
        hideLoading();
        return await saveActiveAs();
      } else {
        // Navegador: descargar.
        downloadBlob(data, activeTab.name);
        markDirty(activeTab, false);
        toast("Archivo descargado", "ok");
        return true;
      }
    } catch (e) {
      toast("Error al guardar: " + e.message, "err");
      return false;
    } finally {
      hideLoading();
    }
  }

  // Devuelve true si se guardó, false si se canceló o falló.
  async function saveActiveAs() {
    if (!activeTab) return false;
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
            if (j.path) pushRecent(j.path, j.name);
            markDirty(activeTab, false);
            toast("Guardado como " + (j.name || "") + " ✓", "ok");
            return true;
          } else {
            toast("Guardar como: cancelado");
            return false;
          }
        } else {
          throw new Error("HTTP " + res.status);
        }
      } else {
        downloadBlob(data, activeTab.name);
        toast("Archivo descargado", "ok");
        return true;
      }
    } catch (e) {
      toast("Error: " + e.message, "err");
      return false;
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

  async function closeTab(t) {
    if (t.dirty) {
      const choice = await dlg.show({
        icon: "⚠️", title: "Cambios sin guardar",
        message: '"' + t.name + '" tiene cambios sin guardar.\n¿Qué quieres hacer?',
        buttons: [
          { label: "Cancelar", value: "cancel" },
          { label: "No guardar", value: "discard", variant: "danger" },
          { label: "Guardar", value: "save", variant: "primary", default: true },
        ],
      });
      if (choice === "cancel" || choice === undefined) return;
      if (choice === "save") {
        // Activar la pestaña para guardarla, y abortar si el guardado falla/cancela.
        if (activeTab !== t) activateTab(t);
        const saved = await saveActive();
        if (saved === false) return;   // cancelado en "Guardar como"
      }
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
    $("btnChart").disabled = !has;
    $("btnValidate").disabled = !has;
    $("btnForm").disabled = !has;
    if (window.XlsxMenus) window.XlsxMenus.refresh();
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
            registerRecentFromToken(token);
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

  async function deleteCurrentMacro() {
    if (macroSel < 0 || !macros[macroSel]) { toast("No hay macro seleccionada"); return; }
    const ok = await dlg.confirm('¿Eliminar la macro "' + macros[macroSel].name + '"?',
      { icon: "🗑️", title: "Eliminar macro", okLabel: "Eliminar", danger: true });
    if (!ok) return;
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

  // ===========================================================================
  //  Gráficos (ECharts, en chart.js)
  // ===========================================================================
  function initCharts() {
    if (!window.XlsxChart) return;
    window.XlsxChart.init({
      getApi: () => univerAPI,
      getActiveTab: () => activeTab,
      toast: toast,
      onDirty: () => { if (activeTab) markDirty(activeTab, true); },
    });
    $("btnChart").addEventListener("click", () => {
      if (!activeTab) return;
      window.XlsxChart.open();      // crea un gráfico desde la selección actual
    });
  }

  // ===========================================================================
  //  Formulario de captura / data entry (form.js)
  // ===========================================================================
  function initForm() {
    if (!window.XlsxForm) return;
    window.XlsxForm.init({
      getApi: () => univerAPI,
      getActiveTab: () => activeTab,
      toast: toast,
      confirm: (msg, opts) => dlg.confirm(msg, opts),
      onDirty: () => { if (activeTab) markDirty(activeTab, true); },
    });
    $("btnForm").addEventListener("click", () => {
      if (!activeTab) return;
      window.XlsxForm.open();
    });
  }

  // ===========================================================================
  //  Archivos recientes (persistidos en localStorage)
  // ===========================================================================
  const RECENT_STORE = "xlsview.recent";
  const RECENT_MAX = 12;

  function getRecents() {
    try { return JSON.parse(localStorage.getItem(RECENT_STORE) || "[]"); }
    catch (e) { return []; }
  }

  // Registra (o promueve al frente) un archivo abierto/guardado con ruta real.
  function pushRecent(path, name) {
    if (!path) return;
    let list = getRecents().filter((r) => r.path !== path);
    list.unshift({ path: path, name: name || path.split(/[\\/]/).pop(), at: Date.now() });
    if (list.length > RECENT_MAX) list = list.slice(0, RECENT_MAX);
    try { localStorage.setItem(RECENT_STORE, JSON.stringify(list)); } catch (e) {}
    if (window.XlsxMenus) window.XlsxMenus.refresh();
  }

  function clearRecents() {
    try { localStorage.removeItem(RECENT_STORE); } catch (e) {}
    if (window.XlsxMenus) window.XlsxMenus.refresh();
  }

  // Resuelve la ruta real de un token ya servido (host) y lo registra reciente.
  async function registerRecentFromToken(token) {
    if (!inHostApp) return;
    try {
      const res = await fetch("/pathfor?token=" + encodeURIComponent(token), { cache: "no-store" });
      const j = await res.json().catch(() => ({}));
      if (j && j.path) pushRecent(j.path, j.path.split(/[\\/]/).pop());
    } catch (e) { /* sin host o sin ruta: no se registra */ }
  }

  // Reabre un archivo reciente por su ruta real (pide al host que lo registre).
  async function openRecent(path) {
    if (!inHostApp) { toast("Recientes solo en la app de escritorio", "err"); return; }
    try {
      showLoading("Abriendo…");
      const res = await fetch("/openpath?path=" + encodeURIComponent(path), { method: "POST" });
      const j = await res.json().catch(() => ({}));
      if (j && j.opened && j.token) {
        const token = j.token.replace(/^\/xls\//, "");
        await openFromToken(token, j.name);
        pushRecent(j.path || path, j.name);
      } else {
        toast("El archivo ya no existe; se quita de recientes", "err");
        const list = getRecents().filter((r) => r.path !== path);
        try { localStorage.setItem(RECENT_STORE, JSON.stringify(list)); } catch (e) {}
        if (window.XlsxMenus) window.XlsxMenus.refresh();
      }
    } catch (e) {
      toast("No se pudo abrir: " + e.message, "err");
    } finally {
      hideLoading();
    }
  }

  // ===========================================================================
  //  Barra de menús clásicos desplegables (menus.js)
  // ===========================================================================
  function initMenus() {
    if (!window.XlsxMenus) return;
    window.XlsxMenus.init({
      getApi: () => univerAPI,
      getActiveTab: () => activeTab,
      toast: toast,
      onDirty: () => { if (activeTab) markDirty(activeTab, true); },
      getRecents: getRecents,
      openRecent: openRecent,
      clearRecents: clearRecents,
    });
  }

  // ===========================================================================
  //  Validación de datos (plugin de Univer)
  // ===========================================================================
  let valRange = null;   // rango activo capturado al abrir el modal

  function a1n(row, col) {
    let s = "", c = col;
    do { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; } while (c >= 0);
    return s + (row + 1);
  }

  function openValidateModal() {
    if (!activeTab || !univerAPI.newDataValidation) {
      toast("La validación no está disponible", "err"); return;
    }
    try {
      const ws = univerAPI.getActiveWorkbook().getActiveSheet();
      const rng = ws.getActiveRange && ws.getActiveRange();
      if (!rng) { toast("Selecciona primero un rango", "err"); return; }
      const r = rng.getRow ? rng.getRow() : 0;
      const c = rng.getColumn ? rng.getColumn() : 0;
      const h = rng.getHeight ? rng.getHeight() : 1;
      const w = rng.getWidth ? rng.getWidth() : 1;
      valRange = { row: r, col: c, h, w, sheetName: ws.getSheetName ? ws.getSheetName() : "" };
      $("vaRange").textContent = valRange.sheetName + "!" + a1n(r, c) +
        (h > 1 || w > 1 ? ":" + a1n(r + h - 1, c + w - 1) : "");
      updateValFields();
      $("valModal").classList.add("open");
    } catch (e) { toast("No se pudo leer la selección: " + e.message, "err"); }
  }
  function closeValidateModal() { $("valModal").classList.remove("open"); }

  // Muestra/oculta los campos según el tipo elegido.
  function updateValFields() {
    const type = $("vaType").value;
    document.querySelector(".va-list").style.display = type === "list" ? "" : "none";
    document.querySelector(".va-range").style.display = type === "range" ? "" : "none";
    document.querySelector(".va-cmp").style.display = (type === "number" || type === "date") ? "" : "none";
    // valor 2 solo para between/notBetween
    const op = $("vaOp").value;
    document.querySelector(".va-val2").style.display = (op === "between" || op === "notBetween") ? "" : "none";
    $("vaLbl1").textContent = (op === "between" || op === "notBetween") ? "Desde" : "Valor";
    $("vaVal1").type = type === "date" ? "date" : "text";
    $("vaVal2").type = type === "date" ? "date" : "text";
    $("vaDropdown").parentElement.style.opacity = (type === "list" || type === "range") ? "1" : "0.5";
  }

  function applyValidation() {
    if (!valRange) return;
    try {
      const ws = univerAPI.getActiveWorkbook().getActiveSheet();
      const rng = ws.getRange(valRange.row, valRange.col, valRange.h, valRange.w);
      const type = $("vaType").value;
      const b = univerAPI.newDataValidation();
      let builder;

      if (type === "list") {
        const raw = $("vaListValues").value;
        const vals = raw.split(/[\n,]+/).map((s) => s.trim()).filter(Boolean);
        if (!vals.length) { toast("Escribe al menos un valor", "err"); return; }
        builder = b.requireValueInList(vals, true, $("vaDropdown").checked);
      } else if (type === "range") {
        const src = $("vaSourceRange").value.trim();
        if (!src) { toast("Indica el rango de origen", "err"); return; }
        // El builder espera una fórmula tipo =$F$1:$F$10 o Hoja!F1:F10
        builder = b.requireValueInRange(src.startsWith("=") ? src : "=" + src, true, $("vaDropdown").checked);
      } else if (type === "checkbox") {
        builder = b.requireCheckbox();
      } else if (type === "number") {
        builder = numOrDateBuilder(b, "number");
        if (!builder) return;
      } else if (type === "date") {
        builder = numOrDateBuilder(b, "date");
        if (!builder) return;
      }
      if (!builder) return;
      if (builder.setAllowInvalid) builder.setAllowInvalid(!$("vaReject").checked);
      const rule = builder.build();
      rng.setDataValidation(rule);
      markDirty(activeTab, true);
      closeValidateModal();
      toast("Validación aplicada ✓", "ok");
    } catch (e) { toast("Error al aplicar: " + e.message, "err"); }
  }

  // Construye la regla de número o fecha según el operador.
  function numOrDateBuilder(b, kind) {
    const op = $("vaOp").value;
    const v1s = $("vaVal1").value, v2s = $("vaVal2").value;
    if (kind === "number") {
      const v1 = parseFloat(v1s), v2 = parseFloat(v2s);
      if (isNaN(v1)) { toast("Valor no válido", "err"); return null; }
      switch (op) {
        case "between": return b.requireNumberBetween(v1, isNaN(v2) ? v1 : v2);
        case "notBetween": return b.requireNumberNotBetween(v1, isNaN(v2) ? v1 : v2);
        case "equal": return b.requireNumberEqualTo(v1);
        case "notEqual": return b.requireNumberNotEqualTo(v1);
        case "gt": return b.requireNumberGreaterThan(v1);
        case "gte": return b.requireNumberGreaterThanOrEqualTo(v1);
        case "lt": return b.requireNumberLessThan(v1);
        case "lte": return b.requireNumberLessThanOrEqualTo(v1);
      }
    } else {
      // fechas como texto ISO YYYY-MM-DD
      if (!v1s) { toast("Indica una fecha", "err"); return null; }
      switch (op) {
        case "between": return b.requireDateBetween(v1s, v2s || v1s);
        case "notBetween": return b.requireDateNotBetween(v1s, v2s || v1s);
        case "equal": return b.requireDateEqualTo(v1s);
        case "gt": return b.requireDateAfter(v1s);
        case "gte": return b.requireDateOnOrAfter(v1s);
        case "lt": return b.requireDateBefore(v1s);
        case "lte": return b.requireDateOnOrBefore(v1s);
        default: return b.requireDateBetween(v1s, v2s || v1s);
      }
    }
    return null;
  }

  function clearValidation() {
    if (!valRange) return;
    try {
      const ws = univerAPI.getActiveWorkbook().getActiveSheet();
      const rng = ws.getRange(valRange.row, valRange.col, valRange.h, valRange.w);
      if (rng.setDataValidation) rng.setDataValidation(null);
      markDirty(activeTab, true);
      closeValidateModal();
      toast("Validación quitada", "ok");
    } catch (e) { toast("Error: " + e.message, "err"); }
  }

  function wireValidation() {
    if (!$("btnValidate")) return;
    $("btnValidate").addEventListener("click", openValidateModal);
    $("vaClose").addEventListener("click", closeValidateModal);
    $("vaCancel").addEventListener("click", closeValidateModal);
    $("vaApply").addEventListener("click", applyValidation);
    $("vaClear").addEventListener("click", clearValidation);
    $("vaType").addEventListener("change", updateValFields);
    $("vaOp").addEventListener("change", updateValFields);
  }

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
        if ($("chartModal").classList.contains("open")) { $("chartModal").classList.remove("open"); return; }
        if ($("valModal").classList.contains("open")) { $("valModal").classList.remove("open"); return; }
        if ($("formModal").classList.contains("open")) { $("formModal").classList.remove("open"); return; }
        if ($("macroModal").classList.contains("open")) { $("macroModal").classList.remove("open"); return; }
      }
    });

    // El host pregunta antes de cerrar la ventana: comprobamos cambios sin
    // guardar y respondemos "force-close" si procede.
    if (inHostApp && window.chrome.webview) {
      window.chrome.webview.addEventListener("message", (ev) => {
        const m = typeof ev.data === "string" ? ev.data : "";
        if (m === "query-close") handleCloseRequest();
      });
    }
    // En navegador: aviso nativo mínimo (WebView2 no lo muestra, por eso el flujo
    // propio de arriba lo cubre en la app).
    window.addEventListener("beforeunload", (e) => {
      if (!inHostApp && tabs.some((t) => t.dirty)) { e.preventDefault(); e.returnValue = ""; }
    });
  }

  // ===========================================================================
  //  Cierre con detección de cambios sin guardar
  // ===========================================================================
  let closing = false;

  async function handleCloseRequest() {
    if (closing) return;
    const dirtyTabs = tabs.filter((t) => t.dirty);
    if (!dirtyTabs.length) { postToHost("force-close"); return; }

    closing = true;
    try {
      const listHtml = "Hay cambios sin guardar en:" +
        "<ul>" + dirtyTabs.map((t) => "<li>" + escapeHtml(t.name) + "</li>").join("") + "</ul>";
      const choice = await dlg.show({
        icon: "⚠️", title: "Cerrar XlsView",
        html: listHtml,
        buttons: [
          { label: "Cancelar", value: "cancel" },
          { label: "No guardar", value: "discard", variant: "danger" },
          { label: dirtyTabs.length > 1 ? "Guardar todo" : "Guardar", value: "save", variant: "primary", default: true },
        ],
      });
      if (choice === "cancel" || choice === undefined) { closing = false; return; }
      if (choice === "save") {
        // Guardar cada pestaña con cambios; si alguna se cancela, abortar el cierre.
        for (const t of dirtyTabs) {
          if (activeTab !== t) activateTab(t);
          const ok = await saveActive();
          if (ok === false) { closing = false; toast("Cierre cancelado", ""); return; }
        }
      }
      // "discard" o guardado completo -> cerrar de verdad.
      postToHost("force-close");
    } catch (e) {
      closing = false;
    }
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
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
          pushRecent(j.path, j.name);
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
    initCharts();          // motor de gráficos (ECharts)
    wireValidation();      // validación de datos (plugin de Univer)
    initForm();            // formulario de captura (data entry)
    initMenus();           // barra de menús clásicos desplegables (menus.js)
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
      openFromToken(token).then(() => {
        registerRecentFromToken(token);
        signalRendered();
      });
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

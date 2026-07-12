/* ===========================================================================
 *  menus.js  —  Barra de menús clásicos desplegables de XlsView (propio)
 *
 *  Reorganiza las acciones de la toolbar en menús categorizados (Archivo,
 *  Insertar, Fórmulas, Datos, Herramientas, Ver) con iconos y submenús.
 *
 *  - NO es un plugin de Univer. Es UI propia.
 *  - Las acciones heredadas se disparan haciendo .click() sobre los botones
 *    ocultos en #legacyButtons (conservan todo el wiring de app.js).
 *  - Las acciones nuevas (fórmulas, presets Univer) reciben un contexto `ctx`
 *    que app.js inyecta en XlsxMenus.init({ getApi, getActiveTab, toast }).
 *
 *  API pública: window.XlsxMenus.init(ctx) y .refresh() (habilitar/deshabilitar).
 * =========================================================================== */
(function () {
  "use strict";

  let ctx = null;              // { getApi, getActiveTab, toast }
  let openMenu = null;         // elemento .menu abierto actualmente
  const $ = (id) => document.getElementById(id);

  // -------------------------------------------------------------------------
  //  Catálogo de fórmulas por categoría. Cada entrada: [NOMBRE, plantilla].
  //  La plantilla usa ⟨⟩ como marcadores; al insertar se coloca el cursor
  //  entre el primer paréntesis. Nombres en la nomenclatura de Univer (EN),
  //  que es como calculan (ver memoria: NO se localizan al español).
  // -------------------------------------------------------------------------
  const FORMULAS = {
    "Matemáticas": [
      "SUM", "SUMIF", "SUMIFS", "PRODUCT", "SUMPRODUCT", "ROUND", "ROUNDUP",
      "ROUNDDOWN", "INT", "MOD", "ABS", "POWER", "SQRT", "EXP", "LN", "LOG",
      "LOG10", "CEILING", "FLOOR", "TRUNC", "SIGN", "GCD", "LCM", "RAND",
      "RANDBETWEEN", "PI", "SUBTOTAL", "AGGREGATE",
    ],
    "Estadísticas": [
      "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "COUNT", "COUNTA", "COUNTBLANK",
      "COUNTIF", "COUNTIFS", "MAX", "MIN", "MAXIFS", "MINIFS", "MEDIAN",
      "MODE", "STDEV", "STDEVP", "VAR", "VARP", "LARGE", "SMALL", "RANK",
      "PERCENTILE", "QUARTILE", "CORREL", "TREND", "FORECAST",
    ],
    "Texto": [
      "CONCATENATE", "CONCAT", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN",
      "FIND", "SEARCH", "REPLACE", "SUBSTITUTE", "UPPER", "LOWER", "PROPER",
      "TRIM", "TEXT", "VALUE", "REPT", "EXACT", "CLEAN", "CHAR", "CODE",
      "UNICHAR", "NUMBERVALUE",
    ],
    "Fecha y hora": [
      "TODAY", "NOW", "DATE", "TIME", "YEAR", "MONTH", "DAY", "HOUR",
      "MINUTE", "SECOND", "WEEKDAY", "WEEKNUM", "DATEDIF", "EDATE",
      "EOMONTH", "NETWORKDAYS", "WORKDAY", "DATEVALUE", "TIMEVALUE", "DAYS",
    ],
    "Lógicas": [
      "IF", "IFS", "IFERROR", "IFNA", "AND", "OR", "NOT", "XOR", "TRUE",
      "FALSE", "SWITCH", "ISBLANK", "ISERROR", "ISNA", "ISNUMBER",
      "ISTEXT", "ISLOGICAL", "ISEVEN", "ISODD",
    ],
    "Búsqueda y referencia": [
      "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", "INDEX", "MATCH", "XMATCH",
      "OFFSET", "INDIRECT", "CHOOSE", "ROW", "ROWS", "COLUMN", "COLUMNS",
      "TRANSPOSE", "ADDRESS", "HYPERLINK",
    ],
    "Financieras": [
      "PMT", "PV", "FV", "NPV", "IRR", "XIRR", "RATE", "NPER", "IPMT",
      "PPMT", "SLN", "DB", "DDB", "EFFECT", "NOMINAL",
    ],
    "Ingeniería": [
      "CONVERT", "DEC2BIN", "DEC2HEX", "BIN2DEC", "HEX2DEC", "BITAND",
      "BITOR", "BITXOR", "COMPLEX", "IMSUM", "IMPRODUCT", "DELTA",
      "GESTEP", "ERF", "BESSELJ",
    ],
    "Web": [
      "ENCODEURL", "FILTERXML", "WEBSERVICE",
    ],
    "Propias de XlsView": [
      "IVA", "NETO", "MANOOBRA", "POTENCIA3F", "CAIDATENSION",
    ],
  };

  // -------------------------------------------------------------------------
  //  Command IDs de los presets OSS de Univer (verificados vía executeCommand).
  // -------------------------------------------------------------------------
  function api() { return ctx && ctx.getApi && ctx.getApi(); }
  function hasTab() { return !!(ctx && ctx.getActiveTab && ctx.getActiveTab()); }
  function toast(m, k) { if (ctx && ctx.toast) ctx.toast(m, k); }

  function exec(cmd, params) {
    const a = api();
    if (!a) { toast("Editor no listo", "err"); return; }
    try { return a.executeCommand(cmd, params); }
    catch (e) { toast("No se pudo ejecutar: " + (e.message || cmd), "err"); }
  }

  function activeSheet() {
    const a = api();
    const wb = a && a.getActiveWorkbook && a.getActiveWorkbook();
    return wb && wb.getActiveSheet && wb.getActiveSheet();
  }

  // Funciones que toman un rango como primer argumento (agregados). Para el
  // resto se inserta =NOMBRE() vacío para que el usuario complete argumentos.
  const RANGE_FNS = new Set([
    "SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTBLANK", "MAX", "MIN", "MEDIAN",
    "MODE", "STDEV", "STDEVP", "VAR", "VARP", "PRODUCT", "SUBTOTAL", "AGGREGATE",
    "CONCAT", "TEXTJOIN", "TRANSPOSE",
  ]);

  // Inserta una fórmula. Si hay un rango de varias celdas seleccionado y la
  // función admite rango, escribe =NOMBRE(rango) en la celda contigua y calcula.
  // Si es una sola celda (o la función no toma rango), escribe =NOMBRE().
  function insertFormula(name) {
    const wb = api() && api().getActiveWorkbook && api().getActiveWorkbook();
    const ws = wb && wb.getActiveSheet && wb.getActiveSheet();
    if (!ws) { toast("Abre o crea una hoja primero", "err"); return; }
    const rng = ws.getActiveRange && ws.getActiveRange();
    if (!rng) { toast("Selecciona una celda", "err"); return; }

    try {
      const g = rng.getRange ? rng.getRange() : null;
      const multi = g && (g.endRow > g.startRow || g.endColumn > g.startColumn);

      if (multi && RANGE_FNS.has(name)) {
        // Elegir celda destino: debajo si es columna, a la derecha si es fila,
        // debajo de la esquina inferior-izquierda en cualquier otro caso.
        const a1 = rng.getA1Notation();
        let destRow, destCol;
        if (g.endColumn === g.startColumn) {          // columna vertical
          destRow = g.endRow + 1; destCol = g.startColumn;
        } else if (g.endRow === g.startRow) {         // fila horizontal
          destRow = g.startRow; destCol = g.endColumn + 1;
        } else {                                      // bloque
          destRow = g.endRow + 1; destCol = g.startColumn;
        }
        const dest = ws.getRange(destRow, destCol);
        dest.setValue("=" + name + "(" + a1 + ")");
        wb.setActiveRange(dest);
        toast("Insertado =" + name + "(" + a1 + ")", "ok");
      } else {
        rng.setValue("=" + name + "()");
        toast("Insertado =" + name + "() — completa los argumentos", "ok");
      }
      if (ctx && ctx.onDirty) ctx.onDirty();
    } catch (e) {
      toast("No se pudo insertar la fórmula", "err");
    }
  }

  // Presets Univer ---------------------------------------------------------
  function toggleFilter() { exec("sheet.command.smart-toggle-filter"); }
  function sortAsc()      { exec("sheet.command.sort-range-asc"); }
  function sortDesc()     { exec("sheet.command.sort-range-desc"); }
  function customSort()   { exec("sheet.command.sort-range-custom"); }
  function openFind()     { exec("ui.operation.open-find-dialog"); }
  function openReplace()  { exec("ui.operation.open-replace-dialog"); }
  function openComment()  { exec("sheets.operation.show-comment-modal"); }
  function insertLink()   { exec("sheet.operation.open-hyper-link-edit-panel"); }
  function cfDuplicate()  { exec("sheet.command.add-duplicate-values-conditional-rule"); }
  function cfColorScale() { exec("sheet.command.add-color-scale-conditional-rule"); }
  function cfDataBar()    { exec("sheet.command.add-data-bar-conditional-rule"); }
  function cfClear()      { exec("sheet.command.clear-range-conditional-rule"); }

  // Dispara un botón heredado por id (conserva el wiring de app.js).
  function legacy(id) {
    return function () {
      const b = $(id);
      if (b && !b.disabled) b.click();
    };
  }

  // -------------------------------------------------------------------------
  //  Definición declarativa de los menús.
  //  item: { icon, label, key?, run?, sub?, needsTab?, sep?, head? }
  //  sub: array de items (submenú). head: título de sección (no clicable).
  // -------------------------------------------------------------------------
  // Submenú de archivos recientes (dinámico desde ctx.getRecents()).
  function recentItems() {
    const list = (ctx && ctx.getRecents && ctx.getRecents()) || [];
    if (!list.length) {
      return [{ label: "(sin archivos recientes)", disabled: true }];
    }
    const items = list.map((r) => ({
      icon: "📄",
      label: r.name || r.path.split(/[\\/]/).pop(),
      hint: r.path,
      run: () => ctx.openRecent(r.path),
    }));
    items.push({ sep: true });
    items.push({ icon: "🧹", label: "Borrar lista de recientes", run: () => ctx.clearRecents() });
    return items;
  }

  function buildDef() {
    const formulaSubs = Object.keys(FORMULAS).map((cat) => ({
      icon: "ƒ", label: cat, needsTab: true,
      sub: FORMULAS[cat].map((fn) => ({
        label: fn, run: () => insertFormula(fn), needsTab: true,
      })),
    }));

    return [
      {
        title: "📄 Archivo",
        items: [
          { icon: "📄", label: "Nuevo", key: "Ctrl+N", run: legacy("btnNew") },
          { icon: "📂", label: "Abrir…", key: "Ctrl+O", run: legacy("btnOpen") },
          { icon: "🕘", label: "Archivos recientes", sub: recentItems() },
          { sep: true },
          { icon: "💾", label: "Guardar", key: "Ctrl+S", run: legacy("btnSave"), needsTab: true },
          { icon: "📥", label: "Guardar como…", run: legacy("btnSaveAs"), needsTab: true },
          { sep: true },
          { icon: "🖨️", label: "Imprimir…", key: "Ctrl+P", run: legacy("btnPrint"), needsTab: true },
          { icon: "⚙️", label: "Configurar página…", run: legacy("btnPageSetup"), needsTab: true },
        ],
      },
      {
        title: "➕ Insertar",
        items: [
          { icon: "📊", label: "Gráfico…", run: legacy("btnChart"), needsTab: true },
          { icon: "📝", label: "Formulario de captura…", run: legacy("btnForm"), needsTab: true },
          { sep: true },
          { icon: "🔗", label: "Hipervínculo…", run: insertLink, needsTab: true },
          { icon: "💬", label: "Comentario", run: openComment, needsTab: true },
        ],
      },
      {
        title: "ƒ Fórmulas",
        items: [
          { head: "Insertar función por categoría" },
          ...formulaSubs,
        ],
      },
      {
        title: "🗂️ Datos",
        items: [
          { icon: "🔽", label: "Filtro (activar/desactivar)", run: toggleFilter, needsTab: true },
          { sep: true },
          { icon: "🔼", label: "Ordenar ascendente", run: sortAsc, needsTab: true },
          { icon: "🔽", label: "Ordenar descendente", run: sortDesc, needsTab: true },
          { icon: "↕️", label: "Orden personalizado…", run: customSort, needsTab: true },
          { sep: true },
          { icon: "✓", label: "Validación de datos…", run: legacy("btnValidate"), needsTab: true },
          { sep: true },
          {
            icon: "🎨", label: "Formato condicional", needsTab: true,
            sub: [
              { icon: "🔁", label: "Resaltar duplicados", run: cfDuplicate, needsTab: true },
              { icon: "🌈", label: "Escala de color", run: cfColorScale, needsTab: true },
              { icon: "📶", label: "Barras de datos", run: cfDataBar, needsTab: true },
              { sep: true },
              { icon: "🧹", label: "Borrar reglas del rango", run: cfClear, needsTab: true },
            ],
          },
        ],
      },
      {
        title: "🛠️ Herramientas",
        items: [
          { icon: "🔍", label: "Buscar…", key: "Ctrl+F", run: openFind, needsTab: true },
          { icon: "🔁", label: "Reemplazar…", key: "Ctrl+H", run: openReplace, needsTab: true },
          { sep: true },
          { icon: "⚡", label: "Macros…", run: legacy("btnMacro"), needsTab: true },
        ],
      },
      {
        title: "👁️ Ver",
        items: [
          { icon: "➕", label: "Acercar", key: "Ctrl++", run: () => $("btnZoomIn").click(), needsTab: true },
          { icon: "➖", label: "Alejar", key: "Ctrl+−", run: () => $("btnZoomOut").click(), needsTab: true },
        ],
      },
    ];
  }

  // -------------------------------------------------------------------------
  //  Motor de render/dropdown
  // -------------------------------------------------------------------------
  function closeAll() {
    document.querySelectorAll(".menu-pop").forEach((p) => p.remove());
    if (openMenu) openMenu.classList.remove("open");
    openMenu = null;
  }

  // Renderiza un popup con una lista de items. Devuelve el elemento .menu-pop.
  function renderPopup(items, anchorRect, parentPop) {
    const pop = document.createElement("div");
    pop.className = "menu-pop";
    let subTimer = null;

    items.forEach((it) => {
      if (it.sep) {
        const s = document.createElement("div");
        s.className = "menu-sep";
        pop.appendChild(s);
        return;
      }
      if (it.head) {
        const h = document.createElement("div");
        h.className = "menu-head";
        h.textContent = it.head;
        pop.appendChild(h);
        return;
      }
      const el = document.createElement("div");
      el.className = "menu-item";
      const disabled = it.disabled || (it.needsTab && !hasTab());
      if (disabled) el.classList.add("disabled");
      if (it.hint) el.title = it.hint;

      const ico = document.createElement("span");
      ico.className = "mi-ico";
      ico.textContent = it.icon || "";
      const lab = document.createElement("span");
      lab.className = "mi-label";
      lab.textContent = it.label;
      el.appendChild(ico);
      el.appendChild(lab);

      if (it.sub) {
        el.classList.add("has-sub");
        const arr = document.createElement("span");
        arr.className = "mi-arrow";
        arr.textContent = "▶";
        el.appendChild(arr);
      } else if (it.key) {
        const k = document.createElement("span");
        k.className = "mi-key";
        k.textContent = it.key;
        el.appendChild(k);
      }

      // Cerrar submenús hermanos al pasar el mouse.
      el.addEventListener("mouseenter", () => {
        pop.querySelectorAll(":scope > .menu-item.hl").forEach((x) => x.classList.remove("hl"));
        // quitar cualquier subpop abierto de este nivel
        if (pop._subPop) { pop._subPop.remove(); pop._subPop = null; }
        if (it.sub && !disabled) {
          el.classList.add("hl");
          clearTimeout(subTimer);
          subTimer = setTimeout(() => {
            const r = el.getBoundingClientRect();
            const sp = renderPopup(it.sub, r, pop);
            pop._subPop = sp;
          }, 60);
        }
      });

      if (!it.sub) {
        el.addEventListener("click", (e) => {
          e.stopPropagation();
          if (disabled) return;
          closeAll();
          try { it.run && it.run(); } catch (err) { toast("Error: " + err.message, "err"); }
        });
      }

      pop.appendChild(el);
    });

    document.body.appendChild(pop);

    // Posicionamiento: submenú al lado derecho del item; menú raíz debajo del título.
    const pw = pop.offsetWidth, ph = pop.offsetHeight;
    const vw = window.innerWidth, vh = window.innerHeight;
    let left, top;
    if (parentPop) {
      left = anchorRect.right - 2;
      top = anchorRect.top - 5;
      if (left + pw > vw - 4) left = anchorRect.left - pw + 2; // voltear a la izquierda
    } else {
      left = anchorRect.left;
      top = anchorRect.bottom + 3;
      if (left + pw > vw - 4) left = vw - pw - 4;
    }
    if (top + ph > vh - 4) top = Math.max(4, vh - ph - 4);
    pop.style.left = Math.max(4, left) + "px";
    pop.style.top = top + "px";
    return pop;
  }

  function openTopMenu(menuEl, def) {
    if (openMenu === menuEl) { closeAll(); return; }
    closeAll();
    openMenu = menuEl;
    menuEl.classList.add("open");
    const rect = menuEl.querySelector(".menu-title").getBoundingClientRect();
    renderPopup(def.items, rect, null);
  }

  function buildMenus() {
    const bar = $("menubar");
    if (!bar) return;
    bar.innerHTML = "";
    const defs = buildDef();

    defs.forEach((def) => {
      const menu = document.createElement("div");
      menu.className = "menu";
      const title = document.createElement("button");
      title.className = "menu-title";
      title.innerHTML = def.title.replace(/^(\S+)\s(.+)$/,
        '<span class="mi-ico">$1</span><span>$2</span>') + '<span class="caret">▼</span>';
      menu.appendChild(title);

      title.addEventListener("click", (e) => {
        e.stopPropagation();
        openTopMenu(menu, def);
      });
      // Hover-switch: si ya hay un menú abierto, pasar el mouse cambia a este.
      title.addEventListener("mouseenter", () => {
        if (openMenu && openMenu !== menu) openTopMenu(menu, def);
      });

      bar.appendChild(menu);
    });
  }

  // Listeners globales: registrados UNA sola vez (buildMenus puede re-ejecutarse
  // en refresh(), así que no se enganchan aquí para no acumularlos).
  let globalWired = false;
  function wireGlobal() {
    if (globalWired) return;
    globalWired = true;
    document.addEventListener("click", (e) => {
      if (openMenu && !openMenu.contains(e.target) &&
          !e.target.closest(".menu-pop")) closeAll();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && openMenu) closeAll();
    });
    window.addEventListener("resize", closeAll);
  }

  // -------------------------------------------------------------------------
  //  API pública
  // -------------------------------------------------------------------------
  window.XlsxMenus = {
    init(context) {
      ctx = context || {};
      wireGlobal();
      buildMenus();
    },
    // Reconstruir la barra (p.ej. tras cambiar la lista de recientes). El
    // estado habilitado/deshabilitado por hasTab() se evalúa al abrir el popup.
    refresh() { buildMenus(); },
  };
})();

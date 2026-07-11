/* ============================================================================
 *  form.js — Formulario de captura (data entry) de XlsView.
 *
 *  Como el "Formulario" clásico de Excel: detecta la fila de encabezados de una
 *  tabla, genera un campo por columna y permite dar de alta, editar y navegar
 *  registros. Escribe celdas normales en la hoja (se guardan como cualquier dato).
 *
 *  Propio, sin dependencias. API:  window.XlsxForm.init({ getApi, getActiveTab, toast, onDirty })
 * ==========================================================================*/
(function () {
  "use strict";

  let deps = null;
  const $ = (id) => document.getElementById(id);

  // Estado del formulario abierto.
  let form = null;   // { sheetName, headerRow, firstCol, lastCol, headers[], dataStart, current }

  function api() { return deps.getApi(); }
  function activeSheet() { return api().getActiveWorkbook().getActiveSheet(); }

  function a1(row, col) {
    let s = "", c = col;
    do { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; } while (c >= 0);
    return s + (row + 1);
  }

  // -------------------------------------------------------------------------
  //  Detección de la tabla a partir de la selección
  // -------------------------------------------------------------------------
  //  - Si el usuario seleccionó un rango con >1 fila: 1ª fila = encabezados.
  //  - Si seleccionó una sola celda/fila: usamos esa fila como encabezados y
  //    detectamos el ancho por las celdas con contenido contiguas.
  function detectTable() {
    const ws = activeSheet();
    const rng = ws.getActiveRange && ws.getActiveRange();
    if (!rng) return null;
    const r = rng.getRow ? rng.getRow() : 0;
    const c = rng.getColumn ? rng.getColumn() : 0;
    const h = rng.getHeight ? rng.getHeight() : 1;
    const w = rng.getWidth ? rng.getWidth() : 1;

    let headerRow = r, firstCol = c, lastCol = c + w - 1;

    if (w <= 1) {
      // Selección de una sola columna/celda: expandir a las columnas contiguas
      // con encabezado en esa fila.
      firstCol = c; lastCol = c;
      // hacia la derecha
      for (let cc = c; cc < c + 64; cc++) {
        const v = ws.getRange(headerRow, cc).getValue();
        if (v == null || v === "") break;
        lastCol = cc;
      }
      // hacia la izquierda
      for (let cc = c - 1; cc >= 0; cc--) {
        const v = ws.getRange(headerRow, cc).getValue();
        if (v == null || v === "") break;
        firstCol = cc;
      }
    }

    const headers = [];
    for (let cc = firstCol; cc <= lastCol; cc++) {
      const v = ws.getRange(headerRow, cc).getValue();
      headers.push({ col: cc, name: (v == null || v === "") ? a1(headerRow, cc) : String(v) });
    }
    if (!headers.length) return null;

    // Contar cuántas filas de datos hay (contiguas debajo del encabezado).
    const dataStart = headerRow + 1;
    let dataEnd = headerRow;   // última fila con algún dato
    for (let rr = dataStart; rr < dataStart + 5000; rr++) {
      let any = false;
      for (let cc = firstCol; cc <= lastCol; cc++) {
        const v = ws.getRange(rr, cc).getValue();
        if (v != null && v !== "") { any = true; break; }
      }
      if (!any) break;
      dataEnd = rr;
    }

    return { sheetName: ws.getSheetName ? ws.getSheetName() : "", headerRow, firstCol, lastCol,
             headers, dataStart, dataEnd, current: dataStart };
  }

  // -------------------------------------------------------------------------
  //  Render del formulario
  // -------------------------------------------------------------------------
  function renderFields() {
    const box = $("frmFields");
    box.innerHTML = "";
    const ws = activeSheet();
    const isNew = form.current > form.dataEnd;
    form.headers.forEach((hd) => {
      const wrap = document.createElement("label");
      wrap.className = "frm-field";
      const lab = document.createElement("span");
      lab.className = "frm-label";
      lab.textContent = hd.name;
      const inp = document.createElement("input");
      inp.type = "text";
      inp.dataset.col = hd.col;
      inp.value = isNew ? "" : cellText(ws, form.current, hd.col);
      wrap.appendChild(lab);
      wrap.appendChild(inp);
      box.appendChild(wrap);
    });
    updateFormStatus();
    // foco en el primer campo
    const first = box.querySelector("input");
    if (first) setTimeout(() => first.focus(), 30);
  }

  function cellText(ws, row, col) {
    const v = ws.getRange(row, col).getValue();
    return (v == null) ? "" : String(v);
  }

  function updateFormStatus() {
    const total = Math.max(0, form.dataEnd - form.dataStart + 1);
    const isNew = form.current > form.dataEnd;
    const idx = isNew ? total + 1 : (form.current - form.dataStart + 1);
    $("frmStatus").textContent = isNew
      ? "Nuevo registro (fila " + (form.current + 1) + ")"
      : "Registro " + idx + " de " + total;
    $("frmRangeInfo").textContent = form.sheetName + "!" +
      a1(form.headerRow, form.firstCol) + ":" + a1(form.headerRow, form.lastCol) + " (encabezados)";
    $("frmPrev").disabled = isNew ? total === 0 : form.current <= form.dataStart;
    $("frmDelete").disabled = isNew || total === 0;
  }

  // -------------------------------------------------------------------------
  //  Acciones
  // -------------------------------------------------------------------------
  // Vuelca los campos del formulario a la fila indicada.
  function writeRow(row) {
    const ws = activeSheet();
    const inputs = $("frmFields").querySelectorAll("input");
    inputs.forEach((inp) => {
      const col = +inp.dataset.col;
      const raw = inp.value;
      // Convertir a número si el texto es numérico (mantiene fórmulas si empieza con =).
      let val = raw;
      if (raw !== "" && !/^=/.test(raw) && !isNaN(Number(raw)) && /\d/.test(raw)) val = Number(raw);
      ws.getRange(row, col).setValue(val === "" ? null : val);
    });
    if (deps.onDirty) deps.onDirty();
  }

  function saveCurrent() {
    const wasNew = form.current > form.dataEnd;
    writeRow(form.current);
    if (wasNew) { form.dataEnd = form.current; }
    deps.toast("Registro guardado ✓", "ok");
    updateFormStatus();
  }

  function newRecord() {
    form.current = form.dataEnd + 1;
    renderFields();
  }

  function nav(delta) {
    const isNew = form.current > form.dataEnd;
    let target;
    if (isNew && delta < 0) target = form.dataEnd;        // desde "nuevo" hacia atrás
    else target = form.current + delta;
    if (target < form.dataStart) target = form.dataStart;
    if (target > form.dataEnd) { newRecord(); return; }   // pasar de la última = nuevo
    form.current = target;
    renderFields();
  }

  async function deleteRecord() {
    if (form.current > form.dataEnd) return;
    const ok = await deps.confirm('¿Eliminar el registro de la fila ' + (form.current + 1) + '?',
      { icon: "🗑️", title: "Eliminar registro", okLabel: "Eliminar", danger: true });
    if (!ok) return;
    const ws = activeSheet();
    // Desplazar hacia arriba las filas inferiores (borrado simple por celdas).
    for (let rr = form.current; rr < form.dataEnd; rr++) {
      for (let cc = form.firstCol; cc <= form.lastCol; cc++) {
        const below = ws.getRange(rr + 1, cc).getValue();
        ws.getRange(rr, cc).setValue(below == null ? null : below);
      }
    }
    // Limpiar la última fila (ahora duplicada).
    for (let cc = form.firstCol; cc <= form.lastCol; cc++) ws.getRange(form.dataEnd, cc).setValue(null);
    form.dataEnd = Math.max(form.headerRow, form.dataEnd - 1);
    if (form.current > form.dataEnd) form.current = Math.max(form.dataStart, form.dataEnd);
    if (deps.onDirty) deps.onDirty();
    renderFields();
    deps.toast("Registro eliminado", "ok");
  }

  // Buscar el siguiente registro que contenga el texto (en cualquier columna).
  function findNext() {
    const q = ($("frmFind").value || "").toLowerCase().trim();
    if (!q) return;
    const ws = activeSheet();
    for (let rr = form.current + 1; rr <= form.dataEnd; rr++) {
      if (rowMatches(ws, rr, q)) { form.current = rr; renderFields(); return; }
    }
    // envolver desde el inicio
    for (let rr = form.dataStart; rr <= form.current; rr++) {
      if (rowMatches(ws, rr, q)) { form.current = rr; renderFields(); return; }
    }
    deps.toast("Sin coincidencias para “" + q + "”");
  }
  function rowMatches(ws, row, q) {
    for (let cc = form.firstCol; cc <= form.lastCol; cc++) {
      const v = ws.getRange(row, cc).getValue();
      if (v != null && String(v).toLowerCase().indexOf(q) >= 0) return true;
    }
    return false;
  }

  // -------------------------------------------------------------------------
  //  Apertura / cierre
  // -------------------------------------------------------------------------
  function open() {
    if (!deps.getActiveTab()) return;
    const t = detectTable();
    if (!t) { deps.toast("Selecciona una tabla con encabezados", "err"); return; }
    if (!t.headers.length) { deps.toast("No se detectaron encabezados", "err"); return; }
    form = t;
    // Empezar en el primer registro si hay datos; si no, en "nuevo".
    form.current = (form.dataEnd >= form.dataStart) ? form.dataStart : form.dataStart;
    renderFields();
    $("formModal").classList.add("open");
  }
  function close() { $("formModal").classList.remove("open"); }

  // -------------------------------------------------------------------------
  //  Wiring
  // -------------------------------------------------------------------------
  function wire() {
    $("frmClose").addEventListener("click", close);
    $("frmCancel").addEventListener("click", close);
    $("frmSave").addEventListener("click", saveCurrent);
    $("frmNew").addEventListener("click", newRecord);
    $("frmPrev").addEventListener("click", () => nav(-1));
    $("frmNext").addEventListener("click", () => nav(1));
    $("frmDelete").addEventListener("click", deleteRecord);
    $("frmFindBtn").addEventListener("click", findNext);
    $("frmFind").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); findNext(); } });
    // Enter en un campo = guardar; flechas para navegar registros.
    $("frmFields").addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); saveCurrent(); }
    });
  }

  window.XlsxForm = {
    init(d) { deps = d; wire(); },
    open,
  };
})();

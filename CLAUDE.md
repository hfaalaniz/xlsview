# XlsView — notas de desarrollo (contexto para retomar)

Visor/editor de hojas de cálculo `.xlsx/.xlsm/.xls/.csv`. App de escritorio
**C# / WinForms / WebView2** (.NET) que sirve una UI web local. Motor de edición:
**Univer 0.5.5** cargado como **bundles UMD locales** (sin bundler) desde `lib/`.
Lectura/escritura de archivos: **SheetJS** (`lib/xlsx.full.min.js`).

## Arquitectura (resumen)

- `index.html` — UI (toolbar oscura propia, pestañas, modales) + carga de todos
  los `<script>`/`<link>` de Univer y libs.
- `app.js` (~2150 líneas) — lógica: SheetJS ⇆ Univer, pestañas, guardar/abrir,
  `initUniver()` (registro de plugins + merge de locales).
- `xlsx-styles.js`, `print.js`, `chart.js`, `form.js` — funciones **propias**
  hechas a mano (gráficos ECharts, impresión, macros, formularios, config de
  página). **NO son plugins de Univer y no se tocan.**
- `app/` — proyecto C# (`Program.cs`, `Server.cs`, `ViewerForm.cs`, `Installer.cs`).
  - `app/build.cmd` → compila a `app/bin/Release/net8.0-windows/XlsViewApp.exe`.
  - `deploy.cmd` → publica a `%LOCALAPPDATA%\Programs\XlsView` y asocia extensiones.
- Assets web: el exe NO los copia. `Program.cs::ResolveWebRoot()` sube hasta 8
  niveles desde el exe buscando `index.html` → usa la **raíz del proyecto**.
  Por eso editar `index.html`/`app.js`/`lib/` en la raíz basta; recompilar C#
  solo hace falta si cambias `.cs`.

## Lo que se hizo (2026-07-11): agregar presets OSS que faltaban

Objetivo del usuario: que xlsview tuviera "todo lo disponible" de Univer OSS,
tomando como referencia el ejemplo que corrimos desde el monorepo `univer/`.

### Referencia usada de `univer/` (monorepo, en `c:\Users\Fabian\univer`)
- Ese monorepo está en Univer **0.25.1** (ESM/Vite) — NO compatible directo con
  este proyecto (0.5.5, UMD). Sirvió solo como **referencia** de qué presets
  existen y cómo se componen (`examples/src/preset-sheets-core/main.ts` importa
  todos los presets de sheets).
- Los ejemplos del monorepo se arrancan con `pnpm dev` (sirve en :3002).
- También se creó un proyecto standalone limpio en
  `c:\Users\Fabian\univer-sheet-clean` (Vite + `@univerjs/preset-sheets-core`
  0.25.1, workbook vacío, es-ES) — solo demostrativo, no relacionado con xlsview.

### Restricción encontrada (clave)
- En **0.5.5** los bundles UMD de los *presets* (`@univerjs/preset-sheets-*`)
  NO traen un bundle autocontenido con la función `UniverSheetsXxxPreset`
  (su `lib/umd/index.js` es en realidad el primer submódulo).
- ⇒ Vía viable = la que ya usaba data-validation: cargar cada **submódulo UMD**
  como `<script>` y registrar sus **Plugins** manualmente en `cfg.plugins`.
- Todos los globals base que necesitan los submódulos (`UniverCore`,
  `UniverSheets`, `UniverSheetsUi`, `UniverEngineFormula`, `UniverSheetsFormulaUi`,
  `UniverUi`, `UniverDocs`, `UniverDesign`, `UniverRpc`, `UniverSheetsDataValidation`)
  YA están presentes en `univer-presets.js` + `univer-preset-sheets-core.js`.
- **No hay `mergeLocales`** en UMD 0.5.5 → se implementó `deepMergeLocale()` en `app.js`.

### Presets agregados (6) — todos verificados funcionando, 0 errores
| Preset | Submódulos UMD (orden de carga) | Plugins registrados |
|---|---|---|
| Filtros | sheets-filter, sheets-filter-ui | UniverSheetsFilterPlugin, UniverSheetsFilterUIPlugin |
| Ordenar | sheets-sort, sheets-sort-ui | UniverSheetsSortPlugin, UniverSheetsSortUIPlugin |
| Buscar/reemplazar | find-replace, sheets-find-replace | UniverFindReplacePlugin, UniverSheetsFindReplacePlugin |
| Hipervínculos | sheets-hyper-link, sheets-hyper-link-ui | UniverSheetsHyperLinkPlugin, UniverSheetsHyperLinkUIPlugin |
| Formato condicional | sheets-conditional-formatting, -ui | UniverSheetsConditionalFormattingPlugin, ...UIPlugin |
| Comentarios (hilos) | thread-comment, thread-comment-ui, sheets-thread-comment, sheets-thread-comment-ui | UniverThreadCommentPlugin, UniverThreadCommentUIPlugin, UniverSheetsThreadCommentPlugin, UniverSheetsThreadCommentUIPlugin |

Los `.js`, `.css` y locales `en-US` se copiaron a `lib/` desde npm @0.5.5
(via `npm pack @univerjs/<pkg>@0.5.5`, extrayendo `package/lib/umd/index.js`,
`package/lib/index.css` y `package/lib/umd/locales/en-US.js`).

Locales combinados en `app.js` (globals): `UniverSheetsDataValidationUiEnUS`,
`UniverThreadCommentUiEnUS`, `UniverSheetsFilterUiEnUS`, `UniverSheetsSortEnUS`,
`UniverFindReplaceEnUS`, `UniverSheetsConditionalFormattingUiEnUS`,
`UniverSheetsHyperLinkUiEnUS`.

### Descartados a propósito
- **Drawing (imágenes)** ❌ — el preset core 0.5.5 YA incluye
  `univer.drawing-manager.service`; re-registrarlo rompe la init con
  `[redi] Identifier "univer.drawing-manager.service" already exists`.
  (Detectado en prueba headless; archivos drawing eliminados de `lib/`.)
- **Notas (note)** y **Tablas (table)** ❌ — NO existían como paquete en 0.5.5
  (son posteriores). Para agregarlos habría que traerlos de una versión donde
  existan (riesgo de incompatibilidad con el core 0.5.5).

### Otro cambio
- Se restauró `lib/univer-preset-sheets-core.css` (estaba borrado localmente;
  recuperado con `git checkout HEAD -- lib/univer-preset-sheets-core.css`).
  Sin él faltaban estilos base de Univer.

## Cómo compilar y ejecutar
```bat
app\build.cmd
app\bin\Release\net8.0-windows\XlsViewApp.exe "ejemplos\Ejemplos XlsView.xlsx"
```
(Herramientas presentes en la máquina: .NET SDK 10.x, WebView2, Node, Chrome.)

## Cómo verificar sin abrir la ventana (prueba headless)
Se sirvió la carpeta por HTTP y se cargó `index.html` en Chrome
`--headless=new` vía CDP, capturando `Runtime.consoleAPICalled` /
`exceptionThrown` para detectar errores de carga de plugins. Comprobaciones:
`window.__xlsview.univerAPI` listo, `createWorkbook({})` OK, y los 12 globals de
submódulo presentes. (Hook de diagnóstico existente: `window.__xlsview`.)

## Estado / próximos pasos posibles
- ✅ 6 presets OSS activos y verificados. App compila y arranca.
- Pendiente si se retoma:
  1. `deploy.cmd` para publicar la versión estable.
  2. Intentar añadir `note` y `table` desde una versión de Univer donde existan.
  3. Traducir la UI de Univer a español (hoy en-US; en 0.5.5 los presets no
     traían locale es-ES).
- Cambios git sin commitear: `M app.js`, `M index.html`, y los nuevos `lib/univer-*`.

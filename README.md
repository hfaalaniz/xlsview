# XlsView — Visor / Editor de hojas de cálculo

Visor y **editor completo** de hojas de cálculo (`.xlsx`, `.xlsm`, `.xls`, `.csv`)
con edición tipo Excel: fórmulas, formato de celdas, varias hojas, filtros…

Mismo concepto que **PdfView**: un ejecutable propio con motor Edge/WebView2 que
abre una ventana sin bordes (maximizada, con la barra de tareas visible), sirve
la interfaz web desde un servidor local y asocia las extensiones en Windows.

## Arquitectura

```
xlsview/
├─ index.html            UI (toolbar, pestañas, editor, modales de impresión)
├─ app.js                Lógica: SheetJS <-> Univer, pestañas, guardar/abrir
├─ xlsx-styles.js        Lee/escribe estilos ricos + config de página
├─ print.js              Impresión: render HTML paginado + vista previa
├─ lib/                  Librerías locales (sin CDN)
│   ├─ xlsx.full.min.js              SheetJS: lee/escribe xlsx/xlsm/xls/csv
│   ├─ react·react-dom·rxjs          dependencias de Univer (UMD)
│   ├─ univer-presets.js             Univer core + fórmulas + render
│   ├─ univer-preset-sheets-core.js  UI de hoja de cálculo
│   ├─ univer-preset-sheets-core.css estilos de Univer
│   └─ univer-locale-en-US.js        textos de la interfaz
├─ app/                  App de escritorio (C# / WebView2 / .NET 8)
│   ├─ Program.cs        arranque, instancia única, argumentos
│   ├─ ViewerForm.cs     ventana WebView2 (sin bordes, arrastre, F11…)
│   ├─ Server.cs         servidor HTTP local (servir, /save, /saveas, /pending)
│   ├─ Installer.cs      asociación de extensiones + PATH (a nivel de usuario)
│   ├─ SplashForm.cs     pantalla de carga
│   └─ build.cmd         compila el ejecutable
└─ Abrir Editor.cmd      lanzador rápido
```

**Motor de edición:** [Univer](https://univer.ai) (open source) — experiencia
tipo Excel con fórmulas, estilos, formato numérico y varias hojas.
**Lectura/escritura de archivos:** [SheetJS](https://sheetjs.com).

## Compilar

Requiere el **SDK de .NET 8+** y el runtime **WebView2** (presente en Windows
10/11 modernos).

```bat
app\build.cmd
```

El ejecutable queda en:

```
app\bin\Release\net8.0-windows\XlsViewApp.exe
```

## Uso

```bat
XlsViewApp.exe "hoja.xlsx"     Abrir una hoja
XlsViewApp.exe                  Abrir el editor vacío
XlsViewApp.exe --install        Asociar .xlsx .xlsm .xls .csv y añadir al PATH
XlsViewApp.exe --uninstall      Revertir
XlsViewApp.exe --stop           Cerrar la ventana abierta
```

Tras `--install` puedes abrir hojas por doble clic o desde la terminal:

```bat
xlsview hoja.xlsx
```

### Atajos en la ventana

| Tecla            | Acción                          |
|------------------|---------------------------------|
| `Ctrl + S`       | Guardar                         |
| `Ctrl + Shift+S` | Guardar como…                   |
| `Ctrl + O`       | Abrir archivo                   |
| `Ctrl + W`       | Cerrar pestaña                  |
| `F11`            | Pantalla completa               |
| `Ctrl + M`       | Minimizar                       |
| `Esc`            | Cerrar (fuera de edición)       |

Arrastra la barra superior para mover la ventana; suéltala sobre otro monitor
para reencajarla. También puedes **arrastrar y soltar** un archivo en la ventana.

## Estilos (fidelidad de formato)

XlsView lee y **replica los estilos del archivo original**. Como la edición
community de SheetJS no expone fuentes, bordes ni alineación, `xlsx-styles.js`
parsea directamente el XML interno del `.xlsx` para recuperarlos:

| Estilo                                   | Leer | Guardar (xlsx/xlsm) |
|------------------------------------------|:----:|:-------------------:|
| Fuente (nombre, tamaño)                  |  ✓   |         ✓           |
| Negrita / cursiva / subrayado / tachado  |  ✓   |         ✓           |
| Color de texto (rgb, tema, indexado)     |  ✓   |         ✓           |
| Color de relleno                         |  ✓   |         ✓           |
| Bordes por lado (13 estilos, color)      |  ✓   |         ✓           |
| Bordes diagonales (↘ ↗ y X)              |  ✓   |         ✓           |
| Alineación H/V y ajuste de texto         |  ✓   |         ✓           |
| Formato numérico (incl. moneda `$`, %)   |  ✓   |         ✓           |
| Celdas combinadas · anchos · altos       |  ✓   |         ✓           |
| Inmovilizar filas/columnas (freeze panes)|  ✓   |         ✓           |
| Protección de hoja / celdas bloqueadas   |  ✓   |         ✓           |
| Fórmulas · varias hojas                  |  ✓   |         ✓           |

Al **guardar**, tras que SheetJS escriba los datos, reinyectamos un `styles.xml`
completo en el ZIP del `.xlsx` con los estilos de Univer, verificado como válido
con ExcelJS.

**Protección:** si la hoja original está protegida (`sheetProtection`), las
celdas bloqueadas (`locked`) pasan a ser de solo lectura en el editor — al
intentar editarlas avisa «Celda bloqueada» — mientras que las celdas
desbloqueadas (`locked="0"`) siguen siendo editables, igual que en Excel. La
protección y el estado de cada celda se conservan al guardar.

> El formato antiguo `.xls` y el `.csv` no llevan estilos; se guardan solo con
> datos y fórmulas. Para máxima fidelidad de formato, usa `.xlsx`.
> La protección de hoja no usa contraseña (Univer core no la implementa): es un
> bloqueo de edición en el editor, no un cifrado.

## Guardado

- **Guardar** sobrescribe el archivo original en disco a través del host.
- **Guardar como…** abre el diálogo nativo de Windows y elige formato
  (`xlsx`, `xlsm`, `xls`, `csv`).
- **Cambios sin guardar al cerrar:** al cerrar una pestaña o toda la ventana
  (aspa, `Esc`, `Alt+F4`) con cambios pendientes, XlsView intercepta el cierre y
  muestra un diálogo propio con **Guardar / No guardar / Cancelar**. Si hay
  varias pestañas modificadas, las lista todas y ofrece «Guardar todo». Sin
  cambios, cierra directo sin preguntar.

> Los avisos y confirmaciones dentro del editor usan **diálogos propios** con el
> estilo de la app (no los cuadros nativos del navegador). Los `MessageBox`
> nativos solo aparecen en fallos de arranque/instalación, cuando la interfaz
> web aún no está disponible.

## Impresión y configuración de página

Univer no trae impresión, así que XlsView genera su propio **render HTML fiel y
paginado** de la hoja (celdas, estilos, bordes, merges, formato numérico) y usa
la impresión de WebView2 — el usuario elige impresora o **«Guardar como PDF»**.

- **🖨️ Imprimir** (o `Ctrl+P`): abre la **vista previa** paginada con el número
  de páginas; botón «Imprimir» para enviar a la impresora/PDF.
- **⚙️ Página**: diálogo de configuración con:

| Control                    | Detalle                                             |
|----------------------------|-----------------------------------------------------|
| Orientación                | Vertical / Horizontal                               |
| Tamaño de papel            | A4, Carta, Legal, A3, A5                             |
| Márgenes                   | Superior / inferior / izquierdo / derecho (pulgadas)|
| Escala                     | Ajustar al N % · o «ajustar a N páginas de ancho»   |
| Área de impresión          | Rango a imprimir (ej. `A1:G50`)                      |
| Repetir filas / columnas   | Títulos que se repiten en cada página (ej. `1:7`)   |
| Encabezado / pie           | Con tokens `&A` hoja · `&P` página · `&N` total · `&D` fecha · `&F` archivo |
| Opciones                   | Líneas de cuadrícula · encabezados fila/col · centrar H/V |

La configuración se **lee del archivo original** (`pageSetup`, `pageMargins`,
`printOptions`, área de impresión y títulos) y se **conserva al guardar**,
verificado como válido con ExcelJS.

## Gráficos

Motor propio con **ECharts** (open source, sin licencias ni marca de agua — los
gráficos nativos de Univer son de pago). Botón **📊 Gráfico**: selecciona un
rango con datos y ábrelo.

- **Tipos:** columnas, barras, líneas, área, torta/dona, columnas apiladas,
  combinado (columna+línea, doble eje), dispersión (XY), radar, embudo, medidor,
  mapa de calor y treemap.
- **Datos:** detecta encabezados (1ª fila) y categorías (1ª columna); múltiples
  series.
- **Personalización:** título, 5 paletas, leyenda, etiquetas de datos, líneas
  suavizadas, barras horizontales, doble eje. Tema oscuro coherente con la app.
- **Live-update:** el gráfico se refresca al editar los datos del rango.
- **Exportar:** PNG (2×) o SVG vectorial.
- **Múltiples gráficos** por hoja: se guardan en el `.xlsx` (como
  `xl/xlsview-charts.json` dentro del ZIP) y se recuperan al reabrir.

## Validación de datos

Botón **✓ Validar**: selecciona un rango y define qué se puede escribir en él
(plugin oficial de Univer, open source). Tipos:

- **Lista de valores** (escribes las opciones) o **lista desde un rango** →
  desplegable en la celda
- **Casilla de verificación**
- **Número** con condición (entre, igual, mayor que…)
- **Fecha** con condición

Opciones: mostrar/ocultar el desplegable, y **rechazar** entradas inválidas o
solo advertir. Se puede **quitar** la validación de un rango.

La validación se guarda como **`<dataValidation>` OOXML estándar** dentro del
`.xlsx` —así **Excel también la lee**— y XlsView la recupera al reabrir.

## Formulario de captura (data entry)

Botón **📝 Formulario**: da de alta y edita registros en una tabla mediante un
diálogo, como el «Formulario» clásico de Excel.

- **Detecta la tabla** a partir de la selección: la fila seleccionada son los
  encabezados y genera un campo por columna. Si seleccionas una sola celda,
  expande a las columnas contiguas con encabezado.
- **Navega** los registros existentes (◀ Anterior / Siguiente ▶), muestra
  «Registro N de M».
- **＋ Nuevo** agrega una fila al final; **💾 Guardar registro** vuelca los
  campos a la hoja (los números se guardan como número, las fórmulas `=…` como
  fórmula).
- **🗑️ Eliminar** quita el registro y desplaza los de abajo.
- **Buscar** salta al siguiente registro que contenga un texto.

Escribe celdas normales, así que los datos se guardan como cualquier contenido.
`Enter` en un campo guarda el registro.

## Fórmulas, funciones propias y macros

El motor de Univer trae **525 funciones** (las 14 categorías de Excel) con
anidamiento, `IF/IFS/SWITCH`, `LAMBDA/LET`, `VLOOKUP/XLOOKUP/INDEX/MATCH`,
referencias entre hojas y literales de matriz. Único límite: los arrays con
**derrame (spill)** (`SEQUENCE`, `FILTER`, `SORT`, `UNIQUE`) no se materializan
en varias celdas con el preset actual.

**Banco de funciones propias** — 38 funciones definidas en `app.js`
(`registerCustomFunctions()`), registradas con `univerAPI.registerFunction` y
usables como nativas. Se anidan con las de Univer (`=ROUND(IVA(1000),0)`).

*Comercial / contable*
| Función | Qué hace |
|---------|----------|
| `=IVA(monto; [tasa])` | Aplica IVA (21% por defecto) |
| `=NETO(total; [tasa])` | Quita el IVA de un total |
| `=IIBB(base; [tasa])` | Ingresos brutos sobre una base |
| `=PRECIOVENTA(costo; [margen%])` | Precio de venta según margen |
| `=MARGEN(costo; precioVenta)` | Margen sobre la venta |
| `=MARKUP(costo; precioVenta)` | Recargo sobre el costo |
| `=DESCUENTO(precio; pct)` | Precio con descuento |
| `=VARIACIONPCT(anterior; actual)` | Variación porcentual |
| `=PORCENTAJEDE(parte; total)` | Qué % es la parte del total |

*Finanzas*
| Función | Qué hace |
|---------|----------|
| `=INTERESSIMPLE(capital; tasa; periodos)` | Interés simple |
| `=INTERESCOMP(capital; tasa; periodos)` | Interés compuesto ganado |
| `=VALORFUTURO(vp; tasa; periodos)` | Valor futuro |
| `=CUOTAFIJA(monto; tasa; nCuotas)` | Cuota fija (sistema francés) |
| `=TASAEFECTIVA(nominal; [capit.=12])` | Tasa efectiva anual (TEA) |

*Ingeniería eléctrica*
| Función | Qué hace |
|---------|----------|
| `=POTENCIA3F(V; I; [cosφ])` | Potencia trifásica (W) |
| `=POTENCIA1F(V; I; [cosφ])` | Potencia monofásica (W) |
| `=CAIDATENSION(L; I; S; [κ])` | Caída de tensión monofásica (V) |
| `=CAIDA3F(L; I; S; [cosφ]; [κ])` | Caída de tensión trifásica (V) |
| `=OHM(tensión; resistencia)` | Corriente por ley de Ohm (A) |
| `=CONSUMOKWH(W; horas; [$kWh])` | Costo del consumo eléctrico |
| `=FACTORPOT(I; V; potAparente)` | Factor de potencia |

*Matemática / geometría*
| Función | Qué hace |
|---------|----------|
| `=HIPOTENUSA(a; b)` | Hipotenusa (Pitágoras) |
| `=AREACIRCULO(radio)` · `=AREATRIANGULO(base; alt)` | Áreas |
| `=REDONDEARM(valor; múltiplo)` | Redondea al múltiplo más cercano |
| `=IMC(pesoKg; alturaM)` | Índice de masa corporal |
| `=REGLA3(a; b; c)` | Regla de tres simple (a·c/b) |

*Conversión*
| Función | Qué hace |
|---------|----------|
| `=CELSIUS_F(°C)` · `=FAHRENHEIT_C(°F)` | Temperatura |
| `=CV_W(cv)` | Caballos de vapor → watts |
| `=CONSUMOKML(km; litros)` | Rendimiento km/litro |

*Texto / fecha*
| Función | Qué hace |
|---------|----------|
| `=INICIALES("Juan Pérez")` | Iniciales → JP |
| `=SOLONUMEROS(texto)` | Deja solo los dígitos |
| `=TITULAR(texto)` | Mayúscula inicial en cada palabra |
| `=CONTARPALABRAS(texto)` | Cuenta palabras |
| `=EDAD(fecha)` · `=TRIMESTRE(fecha)` | Edad en años · trimestre 1-4 |

*Web / XML* (completan el hueco del preset de Univer)
| Función | Qué hace |
|---------|----------|
| `=FILTERXML(xml; xpath)` | Extrae datos de un XML con una expresión XPath (local, sin red) |
| `=WEBSERVICE(url)` | Trae el texto de una URL. La petición la hace el host C# (acceso a internet); muestra «Cargando…» y refresca al llegar la respuesta |

> **Ingeniería, estadística, financieras…** ya vienen **nativas** en Univer (525
> funciones): `BIN2DEC`, `BITAND`, `COMPLEX`, `IMSUM`, `BESSELJ`, `ERF`,
> `CONVERT`, `DELTA`, `GESTEP`, etc. Se invocan por su **nombre en inglés**.
>
> **Nota:** también existe `=MANOOBRA(horas; valorHora; [cargas%])`. Para añadir
> las tuyas, agrega un triplete `[func, "NOMBRE", "descripción"]` al array
> `calculate` en `registerCustomFunctions()` de `app.js` y ejecuta `deploy.cmd`.
> El libro **`ejemplos/Ejemplos XlsView.xlsx`** (hoja «Banco de funciones»)
> muestra las 38 en acción.

**Macros** — botón **⚡ Macro** en la barra. Abre un editor donde escribes
JavaScript que automatiza la hoja activa. Tienes disponibles: `sheet`,
`cell(fila, col)`, `range(f, c, nf, nc)`, `rows`, `cols`, `api`, `workbook` y
`toast(mensaje)`. Las macros se guardan por navegador y hay 4 ejemplos listos
(numerar filas, fila de totales, limpiar formato, resaltar negativos). No es
VBA: es scripting sobre la API de Univer.

## Modo de depuración (opcional)

Define `XLSVIEW_DEBUG_PORT` para exponer la ventana a Chrome DevTools:

```bat
set XLSVIEW_DEBUG_PORT=9333
XlsViewApp.exe "hoja.xlsx"
```

---
Desarrollado por Fabian Alaniz · SheetJS · Univer · WebView2 · .NET

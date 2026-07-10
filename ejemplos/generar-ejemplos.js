// Genera "Ejemplos XlsView.xlsx": libro didáctico con las capacidades reales
// del motor de fórmulas de Univer + funciones propias + macros de XlsView.
//
// Uso:  npm i exceljs   &&   node generar-ejemplos.js
const ExcelJS = require('exceljs');

const wb = new ExcelJS.Workbook();
wb.creator = 'XlsView';
wb.created = new Date();

// ---------- Paleta / estilos reutilizables ----------
const AZUL = 'FF1F4E78', AZULCLARO = 'FFD9E1F2', AMARILLO = 'FFFFF2CC',
      VERDE = 'FFE2EFDA', GRIS = 'FFF2F2F2', BORDE = 'FFBFBFBF';

function fill(color) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: color } }; }
function thinBorder() {
  const s = { style: 'thin', color: { argb: BORDE } };
  return { top: s, left: s, bottom: s, right: s };
}
function styleHeaderRow(ws, row, cols) {
  for (let c = 1; c <= cols; c++) {
    const cell = ws.getCell(row, c);
    cell.fill = fill(AZUL);
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.border = thinBorder();
  }
}
function title(ws, text, colspan) {
  ws.mergeCells(1, 1, 1, colspan);
  const c = ws.getCell(1, 1);
  c.value = text;
  c.font = { name: 'Calibri', size: 16, bold: true, color: { argb: AZUL } };
  c.alignment = { vertical: 'middle' };
  ws.getRow(1).height = 26;
}
function subtitle(ws, row, text, colspan) {
  ws.mergeCells(row, 1, row, colspan);
  const c = ws.getCell(row, 1);
  c.value = text;
  c.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF808080' } };
}
// Sección con banda de color
function sectionBand(ws, row, text, colspan) {
  ws.mergeCells(row, 1, row, colspan);
  const c = ws.getCell(row, 1);
  c.value = text;
  c.fill = fill(AZULCLARO);
  c.font = { name: 'Calibri', size: 11, bold: true, color: { argb: AZUL } };
  c.alignment = { vertical: 'middle' };
  ws.getRow(row).height = 20;
}

// =====================================================================
//  HOJA 1 — Índice / portada
// =====================================================================
{
  const ws = wb.addWorksheet('Índice', { properties: { tabColor: { argb: AZUL } } });
  ws.getColumn(1).width = 4;
  ws.getColumn(2).width = 26;
  ws.getColumn(3).width = 70;
  title(ws, '📊 XlsView — Ejemplos de fórmulas, funciones y macros', 3);
  subtitle(ws, 2, 'Motor: Univer 0.5.5 · 525 funciones · construido y verificado en la app real', 3);

  const items = [
    ['Fórmulas', 'Operadores, precedencia, referencias, rangos, texto y comparaciones'],
    ['Anidamientos', 'Funciones dentro de funciones, IF/IFS anidados, LAMBDA y LET'],
    ['Funciones', 'Catálogo por categoría: lógicas, texto, fecha, búsqueda, matemática/estadística'],
    ['Funciones propias', 'Las 5 originales: IVA, NETO, MANOOBRA, POTENCIA3F, CAIDATENSION'],
    ['Banco de funciones', '38 funciones propias por rubro: comercial, finanzas, eléctrica, mate, conversión, texto, fecha'],
    ['Macros', 'Scripts JavaScript para el botón ⚡ Macro, con su código listo para copiar'],
    ['Limitaciones', 'Qué NO funciona todavía (arrays con derrame) y cómo sortearlo'],
  ];
  let r = 4;
  ws.getCell(r, 2).value = 'Hoja';
  ws.getCell(r, 3).value = 'Contenido';
  styleHeaderRow(ws, r, 3); ws.getCell(r,1).fill = fill(AZUL); ws.getCell(r,1).border = thinBorder();
  r++;
  items.forEach(([hoja, desc], i) => {
    ws.getCell(r, 2).value = hoja;
    ws.getCell(r, 2).font = { bold: true, color: { argb: AZUL } };
    ws.getCell(r, 3).value = desc;
    for (let c = 1; c <= 3; c++) {
      ws.getCell(r, c).border = thinBorder();
      if (i % 2) ws.getCell(r, c).fill = fill(GRIS);
    }
    ws.getRow(r).height = 22;
    r++;
  });
  r += 1;
  ws.getCell(r, 2).value = 'Consejo:';
  ws.getCell(r, 2).font = { bold: true };
  ws.mergeCells(r, 3, r, 3);
  ws.getCell(r, 3).value = 'Haz clic en cualquier celda con fórmula (columna con fondo amarillo) para ver la fórmula en la barra.';
  ws.getCell(r, 3).font = { italic: true, color: { argb: 'FF808080' } };
}

// Helper para hojas de ejemplos con 3 columnas: descripción | fórmula (texto) | resultado (fórmula real)
function exampleSheet(name, tabColor, headerTitle, subtitleText, rowsDef) {
  const ws = wb.addWorksheet(name, { properties: { tabColor: { argb: tabColor } } });
  ws.getColumn(1).width = 4;
  ws.getColumn(2).width = 40;   // descripción
  ws.getColumn(3).width = 42;   // fórmula como texto
  ws.getColumn(4).width = 22;   // resultado (fórmula viva)
  ws.views = [{ state: 'frozen', ySplit: 4 }];
  title(ws, headerTitle, 4);
  subtitle(ws, 2, subtitleText, 4);

  let r = 4;
  const header = () => {
    ws.getCell(r, 2).value = 'Qué hace';
    ws.getCell(r, 3).value = 'Fórmula';
    ws.getCell(r, 4).value = 'Resultado';
    styleHeaderRow(ws, r, 4);
    ws.getCell(r, 1).fill = fill(AZUL); ws.getCell(r, 1).border = thinBorder();
    r++;
  };
  header();
  let banded = 0;
  for (const def of rowsDef) {
    if (def.section) {
      sectionBand(ws, r, def.section, 4);
      r++; banded = 0;
      continue;
    }
    const [desc, formulaText, formula] = def;
    ws.getCell(r, 2).value = desc;
    // fórmula como texto (con apóstrofo para que no se evalúe)
    ws.getCell(r, 3).value = formulaText;
    ws.getCell(r, 3).font = { name: 'Consolas', size: 10, color: { argb: 'FF203040' } };
    ws.getCell(r, 3).fill = fill(GRIS);
    // resultado: fórmula viva (ExcelJS guarda la fórmula; Univer recalcula al abrir)
    if (formula != null) {
      ws.getCell(r, 4).value = { formula: formula };
      ws.getCell(r, 4).fill = fill(AMARILLO);
      ws.getCell(r, 4).font = { name: 'Calibri', size: 11, bold: true };
    }
    for (let c = 1; c <= 4; c++) ws.getCell(r, c).border = thinBorder();
    ws.getRow(r).height = 20;
    r++; banded++;
  }
  return ws;
}

// =====================================================================
//  HOJA 2 — Fórmulas básicas
// =====================================================================
exampleSheet('Fórmulas', 'FF2E75B6',
  '➕ Fórmulas básicas',
  'Operadores aritméticos, comparación, texto, referencias y rangos. La columna amarilla contiene la fórmula viva.',
  [
    { section: 'Aritmética y precedencia' },
    ['Suma, resta, multiplicación', '=2+3*4-1', '2+3*4-1'],
    ['Paréntesis cambian la precedencia', '=(2+3)*4', '(2+3)*4'],
    ['División y decimales', '=10/4', '10/4'],
    ['Potencia', '=2^10', '2^10'],
    ['Porcentaje (sufijo %): 512', '=1024*0.5', '1024*0.5'],
    ['Negativo y unario', '=-5+8', '-5+8'],
    { section: 'Comparación (devuelven VERDADERO/FALSO)' },
    ['Mayor que', '=10>3', '10>3'],
    ['Distinto de', '=5<>5', '5<>5'],
    ['Comparación usada en aritmética', '=(5>3)*100', '(5>3)*100'],
    { section: 'Texto' },
    ['Concatenar con &', '="Hola "&"mundo"', '"Hola "&"mundo"'],
    ['Concatenar número', '="Total: "&(2+3)', '"Total: "&(2+3)'],
    { section: 'Referencias y rangos' },
    ['Referencia a otra celda (G14 = 42)', '=G14', 'G14'],
    ['Suma de un rango', '=SUM(G14:G16)', 'SUM(G14:G16)'],
    ['Referencia entre hojas', "='Funciones propias'!D6", "'Funciones propias'!D6"],
    ['Literal de matriz (suma)', '=SUM({1,2,3;4,5,6})', 'SUM({1,2,3;4,5,6})'],
  ]);
// datos de apoyo para las referencias, en zona libre (G14:G16, no pisa la col D).
{
  const ws = wb.getWorksheet('Fórmulas');
  ws.getColumn(7).width = 8;
  ws.getCell('F14').value = 'Apoyo →';
  ws.getCell('F14').font = { italic: true, color: { argb: 'FF808080' } };
  ws.getCell('G14').value = 42;
  ws.getCell('G15').value = 8;
  ws.getCell('G16').value = 10;
}

// =====================================================================
//  HOJA 3 — Anidamientos
// =====================================================================
exampleSheet('Anidamientos', 'FF548235',
  '🔗 Anidamientos de funciones',
  'Funciones dentro de funciones. El motor evalúa de dentro hacia afuera, sin límite práctico de profundidad.',
  [
    { section: 'IF anidado (if/then/else)' },
    ['IF simple', '=IF(10>5,"sí","no")', 'IF(10>5,"sí","no")'],
    ['IF dentro de IF (2 niveles)', '=IF(G14>50,"alto",IF(G14>10,"medio","bajo"))', 'IF(G14>50,"alto",IF(G14>10,"medio","bajo"))'],
    ['IFS (varias condiciones)', '=IFS(G14>100,"A",G14>40,"B",TRUE,"C")', 'IFS(G14>100,"A",G14>40,"B",TRUE,"C")'],
    ['SWITCH', '=SWITCH(3,1,"uno",2,"dos",3,"tres","otro")', 'SWITCH(3,1,"uno",2,"dos",3,"tres","otro")'],
    ['IFERROR envuelve un error', '=IFERROR(10/0,"sin dividir")', 'IFERROR(10/0,"sin dividir")'],
    { section: 'Anidar funciones de cálculo' },
    ['ROUND(AVERAGE(...))', '=ROUND(AVERAGE(G14:G16),1)', 'ROUND(AVERAGE(G14:G16),1)'],
    ['SUM dentro de un IF', '=IF(SUM(G14:G16)>50,"supera","no")', 'IF(SUM(G14:G16)>50,"supera","no")'],
    ['MAX de dos SUM', '=MAX(SUM(G14:G15),SUM(G15:G16))', 'MAX(SUM(G14:G15),SUM(G15:G16))'],
    ['Texto + número redondeado', '="Prom: "&ROUND(AVERAGE(G14:G16),2)', '"Prom: "&ROUND(AVERAGE(G14:G16),2)'],
    { section: 'LAMBDA y LET (funciones modernas)' },
    ['LAMBDA en línea (x→x²) con 7', '=LAMBDA(x,x*x)(7)', 'LAMBDA(x,x*x)(7)'],
    ['LET (variables locales)', '=LET(a,5,b,3,a*b+a)', 'LET(a,5,b,3,a*b+a)'],
    ['MAP+LAMBDA consumido por SUM', '=SUM(MAP(G14:G16,LAMBDA(v,v*2)))', 'SUM(MAP(G14:G16,LAMBDA(v,v*2)))'],
    ['REDUCE (acumular)', '=REDUCE(0,G14:G16,LAMBDA(acc,v,acc+v))', 'REDUCE(0,G14:G16,LAMBDA(acc,v,acc+v))'],
  ]);
{
  const ws = wb.getWorksheet('Anidamientos');
  ws.getColumn(7).width = 8;
  ws.getCell('F14').value = 'Apoyo →';
  ws.getCell('F14').font = { italic: true, color: { argb: 'FF808080' } };
  ws.getCell('G14').value = 42; ws.getCell('G15').value = 8; ws.getCell('G16').value = 10;
}

// =====================================================================
//  HOJA 4 — Funciones por categoría
// =====================================================================
exampleSheet('Funciones', 'FFBF8F00',
  '🧮 Catálogo de funciones por categoría',
  'Una muestra representativa de las 525 funciones. Datos de apoyo en la zona F14:H18.',
  [
    { section: 'Lógicas' },
    ['AND / OR / NOT', '=AND(5>1,3<9)', 'AND(5>1,3<9)'],
    ['XOR', '=XOR(TRUE,FALSE)', 'XOR(TRUE,FALSE)'],
    { section: 'Texto' },
    ['LEFT / RIGHT / MID', '=MID("Univer",2,3)', 'MID("Univer",2,3)'],
    ['LEN (longitud)', '=LEN("hoja de cálculo")', 'LEN("hoja de cálculo")'],
    ['UPPER / LOWER / PROPER', '=PROPER("juan pérez")', 'PROPER("juan pérez")'],
    ['SUBSTITUTE (reemplazar)', '=SUBSTITUTE("a-b-c","-","/")', 'SUBSTITUTE("a-b-c","-","/")'],
    ['TEXT (formato)', '=TEXT(1234.5,"#,##0.00")', 'TEXT(1234.5,"#,##0.00")'],
    ['TEXTJOIN', '=TEXTJOIN("-",TRUE,"a","b","c")', 'TEXTJOIN("-",TRUE,"a","b","c")'],
    ['REGEXEXTRACT (regex)', '=REGEXEXTRACT("Orden 1234","[0-9]+")', 'REGEXEXTRACT("Orden 1234","[0-9]+")'],
    { section: 'Fecha y hora' },
    ['HOY / AHORA', '=TODAY()', 'TODAY()'],
    ['DÍAS entre fechas', '=DAYS(DATE(2026,12,31),DATE(2026,1,1))', 'DAYS(DATE(2026,12,31),DATE(2026,1,1))'],
    ['DATEDIF (meses)', '=DATEDIF(DATE(2025,1,1),DATE(2026,3,1),"m")', 'DATEDIF(DATE(2025,1,1),DATE(2026,3,1),"m")'],
    ['EOMONTH (fin de mes)', '=EOMONTH(DATE(2026,2,15),0)', 'EOMONTH(DATE(2026,2,15),0)'],
    { section: 'Búsqueda y referencia' },
    ['VLOOKUP (tabla F14:H18)', '=VLOOKUP("B",F14:H18,3,FALSE)', 'VLOOKUP("B",F14:H18,3,FALSE)'],
    ['INDEX + MATCH', '=INDEX(H14:H18,MATCH("C",F14:F18,0))', 'INDEX(H14:H18,MATCH("C",F14:F18,0))'],
    ['XLOOKUP', '=XLOOKUP("D",F14:F18,H14:H18)', 'XLOOKUP("D",F14:F18,H14:H18)'],
    ['CHOOSE', '=CHOOSE(2,"rojo","verde","azul")', 'CHOOSE(2,"rojo","verde","azul")'],
    { section: 'Matemática y estadística' },
    ['SUMIF (con criterio)', '=SUMIF(H14:H18,">20")', 'SUMIF(H14:H18,">20")'],
    ['COUNTIF', '=COUNTIF(F14:F18,"?")', 'COUNTIF(F14:F18,"?")'],
    ['AVERAGE / MEDIAN', '=MEDIAN(H14:H18)', 'MEDIAN(H14:H18)'],
    ['SUMPRODUCT', '=SUMPRODUCT(H14:H18,H14:H18)', 'SUMPRODUCT(H14:H18,H14:H18)'],
    ['ROUND / MOD / POWER', '=MOD(17,5)', 'MOD(17,5)'],
  ]);
// tabla de apoyo F14:H18 para VLOOKUP/INDEX
{
  const ws = wb.getWorksheet('Funciones');
  ws.getColumn(6).width = 8; ws.getColumn(7).width = 14; ws.getColumn(8).width = 10;
  const tabla = [['Cód','Nombre','Valor'],['A','Alfa',10],['B','Beta',25],['C','Gamma',30],['D','Delta',18]];
  tabla.forEach((row, i) => {
    row.forEach((v, j) => {
      const cell = ws.getCell(13 + i, 6 + j);
      cell.value = v;
      cell.border = thinBorder();
      if (i === 0) { cell.fill = fill(AZUL); cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; }
      else if (i % 2) cell.fill = fill(GRIS);
    });
  });
  ws.getCell('E12').value = 'Tabla de apoyo (VLOOKUP / INDEX / XLOOKUP):';
  ws.getCell('E12').font = { italic: true, bold: true, color: { argb: 'FF808080' } };
}

// =====================================================================
//  HOJA 5 — Funciones propias de XlsView
// =====================================================================
exampleSheet('Funciones propias', 'FFC00000',
  '⚙️ Funciones propias de XlsView',
  'Registradas por la app con registerFunction. Se usan y se anidan igual que las nativas. (Fuera de XlsView darán #NOMBRE?.)',
  [
    { section: 'Comercial / contable' },
    ['IVA — aplica IVA (21% por defecto)', '=IVA(1000)', 'IVA(1000)'],
    ['IVA — con tasa (10.5%)', '=IVA(1000,0.105)', 'IVA(1000,0.105)'],
    ['NETO — quita el IVA de un total', '=NETO(1210)', 'NETO(1210)'],
    ['MANOOBRA — horas·valor·(1+cargas)', '=MANOOBRA(10,65000,0.35)', 'MANOOBRA(10,65000,0.35)'],
    { section: 'Eléctricas' },
    ['POTENCIA3F — √3·V·I·cosφ (W)', '=POTENCIA3F(380,20,0.85)', 'POTENCIA3F(380,20,0.85)'],
    ['CAIDATENSION — 2·L·I/(κ·S) (V)', '=CAIDATENSION(50,20,4)', 'CAIDATENSION(50,20,4)'],
    { section: 'Anidadas con funciones nativas' },
    ['Redondear el IVA', '=ROUND(IVA(1000),0)', 'ROUND(IVA(1000),0)'],
    ['Total con IVA formateado', '="$ "&TEXT(IVA(2500),"#,##0.00")', '"$ "&TEXT(IVA(2500),"#,##0.00")'],
    ['Potencia en kW (dividir /1000)', '=ROUND(POTENCIA3F(380,20,0.85)/1000,2)', 'ROUND(POTENCIA3F(380,20,0.85)/1000,2)'],
    ['¿Caída aceptable? (<3%·230V)', '=IF(CAIDATENSION(50,20,4)<230*0.03,"OK","Revisar")', 'IF(CAIDATENSION(50,20,4)<230*0.03,"OK","Revisar")'],
  ]);
// dato D5 referenciado desde la hoja Fórmulas
{
  const ws = wb.getWorksheet('Funciones propias');
  // D5 es la celda de resultado de la primera fila (IVA(1000)); ya tiene fórmula.
}

// =====================================================================
//  HOJA 5b — Banco de funciones (todas las propias por categoría)
// =====================================================================
exampleSheet('Banco de funciones', 'FF9E480E',
  '🏦 Banco de funciones propias de XlsView',
  '38 funciones registradas por la app, agrupadas por rubro. Se usan y anidan igual que las nativas.',
  [
    { section: 'Comercial / contable' },
    ['IVA — aplica IVA (21% def.)', '=IVA(1000)', 'IVA(1000)'],
    ['NETO — quita el IVA', '=NETO(1210)', 'NETO(1210)'],
    ['IIBB — ingresos brutos', '=IIBB(1000,0.03)', 'IIBB(1000,0.03)'],
    ['PRECIOVENTA — según margen', '=PRECIOVENTA(600,0.4)', 'PRECIOVENTA(600,0.4)'],
    ['MARGEN — sobre venta', '=MARGEN(600,1000)', 'MARGEN(600,1000)'],
    ['MARKUP — recargo sobre costo', '=MARKUP(600,1000)', 'MARKUP(600,1000)'],
    ['DESCUENTO — precio con dto.', '=DESCUENTO(1000,0.15)', 'DESCUENTO(1000,0.15)'],
    ['VARIACIONPCT — % de cambio', '=VARIACIONPCT(100,150)', 'VARIACIONPCT(100,150)'],
    ['PORCENTAJEDE — parte/total', '=PORCENTAJEDE(30,120)', 'PORCENTAJEDE(30,120)'],
    { section: 'Finanzas' },
    ['INTERESSIMPLE', '=INTERESSIMPLE(1000,0.05,3)', 'INTERESSIMPLE(1000,0.05,3)'],
    ['INTERESCOMP — interés ganado', '=INTERESCOMP(1000,0.05,3)', 'INTERESCOMP(1000,0.05,3)'],
    ['VALORFUTURO', '=VALORFUTURO(1000,0.05,3)', 'VALORFUTURO(1000,0.05,3)'],
    ['CUOTAFIJA — sistema francés', '=CUOTAFIJA(12000,0.03,12)', 'CUOTAFIJA(12000,0.03,12)'],
    ['TASAEFECTIVA — TEA', '=TASAEFECTIVA(0.24,12)', 'TASAEFECTIVA(0.24,12)'],
    { section: 'Ingeniería eléctrica' },
    ['POTENCIA3F — trifásica (W)', '=POTENCIA3F(380,20,0.85)', 'POTENCIA3F(380,20,0.85)'],
    ['POTENCIA1F — monofásica (W)', '=POTENCIA1F(220,10,0.9)', 'POTENCIA1F(220,10,0.9)'],
    ['CAIDATENSION — monofásica (V)', '=CAIDATENSION(50,20,4)', 'CAIDATENSION(50,20,4)'],
    ['CAIDA3F — trifásica (V)', '=CAIDA3F(50,20,4,0.85)', 'CAIDA3F(50,20,4,0.85)'],
    ['OHM — corriente (A)', '=OHM(220,10)', 'OHM(220,10)'],
    ['CONSUMOKWH — costo ($)', '=CONSUMOKWH(2000,5,80)', 'CONSUMOKWH(2000,5,80)'],
    ['FACTORPOT — factor de potencia', '=FACTORPOT(10,220,2500)', 'FACTORPOT(10,220,2500)'],
    { section: 'Matemática / geometría' },
    ['HIPOTENUSA — Pitágoras', '=HIPOTENUSA(3,4)', 'HIPOTENUSA(3,4)'],
    ['AREACIRCULO', '=AREACIRCULO(2)', 'AREACIRCULO(2)'],
    ['AREATRIANGULO', '=AREATRIANGULO(10,6)', 'AREATRIANGULO(10,6)'],
    ['REDONDEARM — al múltiplo', '=REDONDEARM(47,5)', 'REDONDEARM(47,5)'],
    ['IMC — masa corporal', '=IMC(80,1.75)', 'IMC(80,1.75)'],
    ['REGLA3 — regla de tres', '=REGLA3(3,6,10)', 'REGLA3(3,6,10)'],
    { section: 'Conversión' },
    ['CELSIUS_F — °C→°F', '=CELSIUS_F(100)', 'CELSIUS_F(100)'],
    ['FAHRENHEIT_C — °F→°C', '=FAHRENHEIT_C(212)', 'FAHRENHEIT_C(212)'],
    ['CV_W — caballos→watts', '=CV_W(1)', 'CV_W(1)'],
    ['CONSUMOKML — km/litro', '=CONSUMOKML(500,40)', 'CONSUMOKML(500,40)'],
    { section: 'Texto / utilidad' },
    ['INICIALES', '=INICIALES("Juan Carlos Pérez")', 'INICIALES("Juan Carlos Pérez")'],
    ['SOLONUMEROS — deja dígitos', '=SOLONUMEROS("AB-12-3x4")', 'SOLONUMEROS("AB-12-3x4")'],
    ['TITULAR — Mayúscula Inicial', '=TITULAR("hola mundo cruel")', 'TITULAR("hola mundo cruel")'],
    ['CONTARPALABRAS', '=CONTARPALABRAS("uno dos tres")', 'CONTARPALABRAS("uno dos tres")'],
    { section: 'Fecha' },
    ['EDAD — años desde una fecha', '=EDAD(DATE(1990,6,15))', 'EDAD(DATE(1990,6,15))'],
    ['TRIMESTRE — 1..4 de una fecha', '=TRIMESTRE(DATE(2026,8,1))', 'TRIMESTRE(DATE(2026,8,1))'],
  ]);

// =====================================================================
//  HOJA 6 — Macros
// =====================================================================
{
  const ws = wb.addWorksheet('Macros', { properties: { tabColor: { argb: 'FF7030A0' } } });
  ws.getColumn(1).width = 4;
  ws.getColumn(2).width = 30;
  ws.getColumn(3).width = 90;
  ws.views = [{ state: 'frozen', ySplit: 5 }];
  title(ws, '⚡ Macros — automatización con JavaScript', 3);
  subtitle(ws, 2, 'Pega el código en el botón ⚡ Macro de XlsView y pulsa Ejecutar. Estas ya vienen precargadas en la app.', 3);
  ws.mergeCells(3, 1, 3, 3);
  ws.getCell(3, 1).value = 'API disponible:  sheet · cell(fila,col) · range(f,c,nf,nc) · rows · cols · api · workbook · toast(msg)   —   índices 0-based.';
  ws.getCell(3, 1).font = { name: 'Consolas', size: 9.5, color: { argb: 'FF7030A0' } };

  let r = 5;
  ws.getCell(r, 2).value = 'Macro';
  ws.getCell(r, 3).value = 'Código (cópialo al botón ⚡ Macro)';
  styleHeaderRow(ws, r, 3); ws.getCell(r,1).fill=fill(AZUL); ws.getCell(r,1).border=thinBorder();
  r++;

  const macros = [
    ['Numerar filas', "// Numera la primera columna desde la fila 6.\nvar desde = 5, cuantas = 10;\nfor (var i = 0; i < cuantas; i++) cell(desde + i, 0).setValue(i + 1);\ntoast('Numeradas ' + cuantas + ' filas');"],
    ['Fila de totales', "// Suma columnas D..G del rango de datos y escribe la fila de totales.\nvar fIni = 5, fFin = 15, cIni = 3, cFin = 6, fTot = fFin + 1;\nfor (var c = cIni; c <= cFin; c++) {\n  var L = String.fromCharCode(65 + c);\n  cell(fTot, c).setValue('=SUM(' + L + (fIni+1) + ':' + L + (fFin+1) + ')');\n}\ncell(fTot, cIni - 1).setValue('TOTAL');\ntoast('Totales agregados');"],
    ['Resaltar negativos', "// Pinta de rojo los números negativos del rango.\nvar f0=5, f1=30, c0=3, c1=7;\nfor (var r=f0; r<=f1; r++) for (var c=c0; c<=c1; c++) {\n  var v = cell(r,c).getValue();\n  if (typeof v==='number' && v<0) cell(r,c).setFontColor('#c0392b');\n}\ntoast('Negativos resaltados');"],
    ['Rellenar serie', "// Rellena B6:B25 con una progresión (inicio 100, paso 25).\nvar inicio=100, paso=25;\nfor (var i=0; i<20; i++) cell(5+i, 1).setValue(inicio + i*paso);\ntoast('Serie generada');"],
    ['Encabezado con estilo', "// Escribe y da formato a una fila de encabezado (fila 6, A..E).\nvar tit = ['Ítem','Descripción','Cantidad','Precio','Total'];\nfor (var c=0; c<tit.length; c++) {\n  var cel = cell(5, c);\n  cel.setValue(tit[c]);\n  cel.setFontWeight('bold').setFontColor('#ffffff').setBackground('#1F4E78');\n}\ntoast('Encabezado creado');"],
    ['Suma rápida de la selección', "// Suma el rango seleccionado y muestra el total.\nvar sel = sheet.getActiveRange ? sheet.getActiveRange() : null;\nif (!sel) { toast('Selecciona un rango'); }\nelse {\n  var vals = sel.getValues(), t = 0;\n  for (var i=0;i<vals.length;i++) for (var j=0;j<vals[i].length;j++)\n    if (typeof vals[i][j]==='number') t += vals[i][j];\n  toast('Suma = ' + t);\n}"],
  ];
  macros.forEach(([name, code], i) => {
    ws.getCell(r, 2).value = name;
    ws.getCell(r, 2).font = { bold: true, color: { argb: 'FF7030A0' } };
    ws.getCell(r, 2).alignment = { vertical: 'top' };
    ws.getCell(r, 3).value = code;
    ws.getCell(r, 3).font = { name: 'Consolas', size: 9.5 };
    ws.getCell(r, 3).alignment = { vertical: 'top', wrapText: true };
    ws.getCell(r, 3).fill = fill(GRIS);
    for (let c = 1; c <= 3; c++) ws.getCell(r, c).border = thinBorder();
    const lines = code.split('\n').length;
    ws.getRow(r).height = Math.max(20, lines * 14);
    r++;
  });
}

// =====================================================================
//  HOJA 7 — Limitaciones
// =====================================================================
{
  const ws = wb.addWorksheet('Limitaciones', { properties: { tabColor: { argb: 'FFED7D31' } } });
  ws.getColumn(1).width = 4;
  ws.getColumn(2).width = 34;
  ws.getColumn(3).width = 34;
  ws.getColumn(4).width = 44;
  ws.views = [{ state: 'frozen', ySplit: 4 }];
  title(ws, '⚠️ Limitaciones y cómo sortearlas', 4);
  subtitle(ws, 2, 'El motor calcula arrays internamente, pero el preset actual no los DERRAMA (spill) en varias celdas.', 4);

  let r = 4;
  ws.getCell(r,2).value='Función'; ws.getCell(r,3).value='Resultado en XlsView'; ws.getCell(r,4).value='Alternativa que sí funciona';
  styleHeaderRow(ws, r, 4); ws.getCell(r,1).fill=fill(AZUL); ws.getCell(r,1).border=thinBorder();
  r++;

  const lim = [
    ['SEQUENCE(n)', '#REF! (ni envuelta)', 'Generar la serie con una macro, o escribirla a mano'],
    ['FILTER(rango,cond)', '#SPILL!', 'Usar SUMIF/COUNTIF, o SUMPRODUCT con la condición'],
    ['SORT(rango)', '#SPILL!', 'Ordenar los datos manualmente en la hoja (menú de datos)'],
    ['UNIQUE(rango)', '#SPILL!', 'Quitar duplicados a mano, o COUNTIF para marcar repetidos'],
    ['MAP(...) sola', '#SPILL!', 'Envolver en SUM/MAX: SÍ funciona =SUM(MAP(...))'],
  ];
  lim.forEach(([f, res, alt], i) => {
    ws.getCell(r,2).value = f; ws.getCell(r,2).font={name:'Consolas',size:10,bold:true};
    ws.getCell(r,3).value = res; ws.getCell(r,3).font={color:{argb:'FFC00000'},bold:true};
    ws.getCell(r,4).value = alt; ws.getCell(r,4).font={name:'Consolas',size:9.5,color:{argb:'FF375623'}};
    for (let c=1;c<=4;c++){ ws.getCell(r,c).border=thinBorder(); if(i%2) ws.getCell(r,c).fill=fill(GRIS); }
    ws.getRow(r).height = 22;
    r++;
  });
  r += 1;
  ws.mergeCells(r,2,r,4);
  ws.getCell(r,2).value = '✔ MAP y REDUCE SÍ funcionan si su array se consume dentro de una función:';
  ws.getCell(r,2).font = { bold: true, color: { argb: 'FF375623' } };
  r++;
  ws.getCell(r,2).value = 'SUM(MAP({2,4,6},LAMBDA(v,v*v))) = 4+16+36';
  ws.getCell(r,2).font = { name: 'Consolas', size: 10 };
  ws.getCell(r,3).value = { formula: 'SUM(MAP({2,4,6},LAMBDA(v,v*v)))' };
  ws.getCell(r,3).fill = fill(AMARILLO); ws.getCell(r,3).font = { bold: true };
  ws.getCell(r,3).border = thinBorder();
  r++;
  ws.getCell(r,2).value = 'REDUCE(0,{1,2,3,4,5},LAMBDA(a,v,a+v))';
  ws.getCell(r,2).font = { name: 'Consolas', size: 10 };
  ws.getCell(r,3).value = { formula: 'REDUCE(0,{1,2,3,4,5},LAMBDA(a,v,a+v))' };
  ws.getCell(r,3).fill = fill(AMARILLO); ws.getCell(r,3).font = { bold: true };
  ws.getCell(r,3).border = thinBorder();
  r++;
  ws.getCell(r,2).value = 'SUM({1,2,3,4,5}) — matriz literal';
  ws.getCell(r,2).font = { name: 'Consolas', size: 10 };
  ws.getCell(r,3).value = { formula: 'SUM({1,2,3,4,5})' };
  ws.getCell(r,3).fill = fill(AMARILLO); ws.getCell(r,3).font = { bold: true };
  ws.getCell(r,3).border = thinBorder();
}

const OUT = require('path').join(__dirname, 'Ejemplos XlsView.xlsx');
wb.xlsx.writeFile(OUT).then(() => {
  console.log('OK escrito:', OUT);
  console.log('Hojas:', wb.worksheets.map(w => w.name).join(' · '));
});

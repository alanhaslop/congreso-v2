// ═══════════════════════════════════════════════════════════════
//  AGF Mesenchymal — Ventas Congreso · Google Apps Script
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1K2rELW54yEqQ8pKXHyPeyDqVLygeCQyQJMeH8w4n5m4';
const ADMIN_EMAIL    = 'alan.haslop@dermacells.com.ar';
const REPLY_TO       = 'alan.haslop@dermacells.com.ar';
const SHEET_VENTAS    = 'Ventas';
const SHEET_RESUMEN   = 'Resumen';
const SHEET_STOCK     = 'Stock';
const SHEET_DASHBOARD = 'Dashboard';
const PDF_FOLDER_ID  = '1qGxephO8pfFAz41B28f39BZM40YdM2GS';

// Precios base (deben coincidir con index.html)
const P_CERRADA   = 750;
const P_COMBINADA = 900;


// ── ENTRY POINT ─────────────────────────────────────────────────
function doPost(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    if (!e || !e.postData) throw new Error('Sin payload');
    const p = JSON.parse(e.postData.contents);
    if (!p.ventaNum) throw new Error('Payload inválido: falta ventaNum');

    // PDF: si viene base64 del cliente lo usamos directamente (calidad exacta del recibo visual)
    // Si no, fallback a la conversión HTML→GDoc (legacy / offline sync sin PDF)
    var pdfBlob, pdfUrl;
    if (p.pdfBase64) {
      var nombre    = 'Recibo_Venta_' + p.ventaNum + '_' + (p.cliente && p.cliente.apellido || 'cliente') + '.pdf';
      var pdfBytes  = Utilities.base64Decode(p.pdfBase64);
      var rawBlob   = Utilities.newBlob(pdfBytes, 'application/pdf', nombre);
      var folder    = DriveApp.getFolderById(PDF_FOLDER_ID);
      var savedFile = folder.createFile(rawBlob);
      savedFile.setName(nombre);
      pdfUrl  = savedFile.getUrl();
      pdfBlob = savedFile.getBlob();
      pdfBlob.setName(nombre);
    } else {
      var result = generarPdfRecibo(p);
      pdfBlob = result.blob;
      pdfUrl  = result.url;
    }

    guardarVenta(p, pdfUrl);
    actualizarResumen(p);
    obtenerOCrearHoja(SHEET_STOCK, crearHojaStock);         // crea la hoja si no existe
    obtenerOCrearHoja(SHEET_DASHBOARD, crearHojaDashboard); // crea la hoja si no existe

    if (p.cliente && p.cliente.mail) enviarEmailCliente(p, pdfBlob);
    if (ADMIN_EMAIL)                 enviarEmailAdmin(p, pdfBlob);

    out.setContent(JSON.stringify({ ok: true, ventaNum: p.ventaNum }));
  } catch (err) {
    console.error('doPost error:', err.message, err.stack);
    out.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return out;
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'AGF Ventas API activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── CALCULAR UNIDADES POR PRODUCTO ───────────────────────────────
// Cerrada: 5 unidades del producto indicado
// Combinada: 1 unidad por cada slot
function calcUnidades(p) {
  const u = { Dermal: 0, Capillary: 0, Pink: 0, Biomask: 0 };
  (p.cajas || []).forEach(function(cj) {
    if (cj.tipo === 'cerrada') {
      if (u[cj.detalle] !== undefined) u[cj.detalle] += 5;
    } else {
      cj.detalle.split(',').forEach(function(prod) {
        const key = prod.trim();
        if (u[key] !== undefined) u[key]++;
      });
    }
  });
  return u;
}


// ── GUARDAR VENTA EN HOJA ────────────────────────────────────────
// Columnas: A–C datos venta | D Vendedor | E–W datos venta | X PDF link | Y–AB unidades por producto
function guardarVenta(p, pdfUrl) {
  const sheet = obtenerOCrearHoja(SHEET_VENTAS, crearHeadersVentas);
  const c = p.cliente     || {};
  const f = p.facturacion || {};
  const u = calcUnidades(p);

  const cajasDetalle = (p.cajas || []).map(function(cj) {
    let txt = 'Caja ' + cj.caja + ' (' + cj.tipo + '): ' + cj.detalle;
    if (cj.descCaja > 0) txt += ' [-' + cj.descCaja + '%]';
    txt += ' = u$' + cj.precio;
    return txt;
  }).join('\n');

  sheet.appendRow([
    new Date(),                                           // A  Timestamp
    p.ventaNum,                                           // B  Venta #
    p.dispositivo    || '',                               // C  Dispositivo
    p.vendedor       || '',                               // D  Vendedor
    p.fecha          || '',                               // E  Fecha local
    c.nombre         || '',                               // F  Nombre
    c.apellido       || '',                               // G  Apellido
    c.cuit           || '',                               // H  CUIT/CUIL
    c.mail           || '',                               // I  Email
    c.tel            || '',                               // J  Teléfono
    c.localidad      || '',                               // K  Localidad
    p.condFiscal     || '',                               // L  Cond. Fiscal
    f.mismosContacto ? '' : (f.razonSocial     || ''),   // M  Razón Social
    f.mismosContacto ? '' : (f.cuitFacturacion || ''),   // N  CUIT Fact.
    p.metodoCobro    || '',                               // O  Método cobro
    p.moneda         || 'USD',                            // P  Moneda
    p.tipoCambio     || '',                               // Q  Tipo cambio
    (p.cajas || []).length,                               // R  Cant. cajas
    cajasDetalle,                                         // S  Detalle cajas
    p.descuentoGlobal || 0,                               // T  Desc. global %
    p.subtotalUSD    || p.totalUSD || 0,                  // U  Subtotal U$D
    p.totalUSD       || 0,                                // V  Total U$D
    p.totalARS       || '',                               // W  Total ARS
    '',                                                   // X  Recibo PDF (fórmula abajo)
    u.Dermal,                                             // Y  Dermal (unidades)
    u.Capillary,                                          // Z  Capillary (unidades)
    u.Pink,                                               // AA Pink (unidades)
    u.Biomask                                             // AB Biomask (unidades)
  ]);

  // Hipervínculo al PDF en columna X (24)
  if (pdfUrl) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 24).setFormula('=HYPERLINK("' + pdfUrl + '","📄 Ver recibo")');
  }
}

function crearHeadersVentas(sheet) {
  const headers = [
    'Timestamp','Venta #','Dispositivo','Vendedor','Fecha local',
    'Nombre','Apellido','CUIT/CUIL','Email','Teléfono','Localidad',
    'Cond. Fiscal','Razón Social Fact.','CUIT Facturación',
    'Método cobro','Moneda','Tipo de cambio',
    'Cant. cajas','Detalle cajas',
    'Desc. global %','Subtotal U$D','Total U$D','Total ARS',
    'Recibo PDF',
    'Dermal (u)','Capillary (u)','Pink (u)','Biomask (u)'
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1D9E75').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1,  160);
  sheet.setColumnWidth(4,  120); // Vendedor
  sheet.setColumnWidth(5,  140); // Fecha local
  sheet.setColumnWidth(19, 280); // Detalle cajas
  sheet.setColumnWidths(6, 2, 110); // Nombre, Apellido
  sheet.setColumnWidth(24, 100); // PDF
  // Destacar columnas de stock en celeste
  sheet.getRange(1, 25, 1, 4).setBackground('#1B3A52');
}


// ── HOJA DE RESUMEN (por día + método + vendedor) ────────────────
function actualizarResumen(p) {
  const sheet = obtenerOCrearHoja(SHEET_RESUMEN, crearHeadersResumen);
  const hoy   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const datos = sheet.getDataRange().getValues();
  const u     = calcUnidades(p);

  const COL = { fecha:0, metodo:1, vendedor:2, cajas:3, usd:4, ars:5, ventas:6,
                dermal:7, capillary:8, pink:9, biomask:10 };
  const NCOLS = 11;

  const vendedorKey = (p.vendedor || '').toString().trim();

  let filaExistente = -1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][COL.fecha]    === hoy
     && datos[i][COL.metodo]   === p.metodoCobro
     && datos[i][COL.vendedor] === vendedorKey) {
      filaExistente = i + 1;
      break;
    }
  }

  const cantCajas = (p.cajas || []).length;
  const totalUSD  = p.totalUSD  || 0;
  const totalARS  = p.totalARS  || 0;

  if (filaExistente > 0) {
    const fila = sheet.getRange(filaExistente, 1, 1, NCOLS).getValues()[0];
    sheet.getRange(filaExistente, COL.cajas    + 1).setValue(fila[COL.cajas]    + cantCajas);
    sheet.getRange(filaExistente, COL.usd      + 1).setValue(fila[COL.usd]      + totalUSD);
    sheet.getRange(filaExistente, COL.ars      + 1).setValue(fila[COL.ars]      + totalARS);
    sheet.getRange(filaExistente, COL.ventas   + 1).setValue(fila[COL.ventas]   + 1);
    sheet.getRange(filaExistente, COL.dermal   + 1).setValue(fila[COL.dermal]   + u.Dermal);
    sheet.getRange(filaExistente, COL.capillary+ 1).setValue(fila[COL.capillary]+ u.Capillary);
    sheet.getRange(filaExistente, COL.pink     + 1).setValue(fila[COL.pink]     + u.Pink);
    sheet.getRange(filaExistente, COL.biomask  + 1).setValue(fila[COL.biomask]  + u.Biomask);
  } else {
    sheet.appendRow([hoy, p.metodoCobro, vendedorKey, cantCajas, totalUSD, totalARS, 1,
                     u.Dermal, u.Capillary, u.Pink, u.Biomask]);
  }
}

function crearHeadersResumen(sheet) {
  sheet.appendRow(['Fecha','Método cobro','Vendedor','Cajas vendidas','Total U$D','Total ARS','# Ventas',
                   'Dermal (u)','Capillary (u)','Pink (u)','Biomask (u)']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 11)
    .setBackground('#1B3A52').setFontColor('#ffffff').setFontWeight('bold');
}


// ── HOJA DE STOCK ────────────────────────────────────────────────
function crearHojaStock(sheet) {
  const PRODS = ['Dermal', 'Capillary', 'Pink', 'Biomask'];
  // Columnas en hoja Ventas para cada producto
  const COL_VENTAS = { Dermal: 'Y', Capillary: 'Z', Pink: 'AA', Biomask: 'AB' };

  // ── Título ──
  sheet.getRange('A1').setValue('CONTROL DE STOCK — BAAS 2026');
  sheet.getRange('A1:F1').merge()
    .setBackground('#1B3A52').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  // ── Tabla de stock ──
  sheet.getRange('A3:F3').setValues([[
    'Producto', 'Stock Inicial', 'Vendidas Total', 'Stock Disponible', 'Vendidas Hoy', 'Stock Fin del Día'
  ]]);
  sheet.getRange('A3:F3')
    .setBackground('#1D9E75').setFontColor('#ffffff').setFontWeight('bold');

  PRODS.forEach(function(prod, i) {
    var row  = 4 + i;
    var col  = COL_VENTAS[prod];
    sheet.getRange('A' + row).setValue(prod).setFontWeight('bold');
    sheet.getRange('B' + row).setValue(0);   // ← ingresá el stock inicial acá
    // Vendidas Total (todas las ventas)
    sheet.getRange('C' + row).setFormula(
      '=IFERROR(SUM(Ventas!' + col + '2:' + col + '),0)'
    );
    // Stock disponible = inicial - vendidas total
    sheet.getRange('D' + row).setFormula('=B' + row + '-C' + row);
    // Vendidas hoy
    sheet.getRange('E' + row).setFormula(
      '=IFERROR(SUMIFS(Ventas!' + col + ':' + col + ',Ventas!A:A,">="&TODAY(),Ventas!A:A,"<"&TODAY()+1),0)'
    );
    // Stock fin del día = disponible - vendidas hoy
    sheet.getRange('F' + row).setFormula('=D' + row + '-E' + row);
  });

  // Destacar celdas editables de Stock Inicial en amarillo
  sheet.getRange('B4:B7')
    .setBackground('#FFF9C4')
    .setNote('✏️ Ingresá el stock inicial antes del evento');

  // Bordes y colores alternos en tabla
  sheet.getRange('A3:F7').setBorder(true, true, true, true, true, true);
  sheet.getRange('A4:F4').setBackground('#f9f9f7');
  sheet.getRange('A5:F5').setBackground('#ffffff');
  sheet.getRange('A6:F6').setBackground('#f9f9f7');
  sheet.getRange('A7:F7').setBackground('#ffffff');

  // ── Separator ──
  sheet.getRange('A9').setValue('INFORME DIARIO POR PRODUCTO')
    .setFontWeight('bold').setFontSize(11).setFontColor('#1B3A52');

  sheet.getRange('A10:I10').setValues([[
    'Fecha', 'Dermal', 'Capillary', 'Pink', 'Biomask', 'Cajas', 'Total U$D', 'Total ARS', '# Ventas'
  ]]);
  sheet.getRange('A10:I10')
    .setBackground('#1D9E75').setFontColor('#ffffff').setFontWeight('bold');

  // QUERY sobre hoja Resumen (ya tiene datos agregados por día, mucho más simple)
  // Resumen: A=Fecha, B=Método, C=Vendedor, D=Cajas, E=U$D, F=ARS, G=#Ventas, H=Dermal, I=Capillary, J=Pink, K=Biomask
  sheet.getRange('A11').setFormula(
    '=IFERROR(QUERY(Resumen!A2:K,'
    + '"SELECT A, SUM(H), SUM(I), SUM(J), SUM(K), SUM(D), SUM(E), SUM(F), SUM(G) '
    + 'WHERE A IS NOT NULL '
    + 'GROUP BY A '
    + 'ORDER BY A DESC '
    + 'LABEL A \'\', SUM(H) \'\', SUM(I) \'\', SUM(J) \'\', SUM(K) \'\', SUM(D) \'\', SUM(E) \'\', SUM(F) \'\', SUM(G) \'\'"'
    + ',0),"Sin ventas aún")'
  );

  // ── Anchos de columna ──
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 130);
  sheet.setFrozenRows(0);
}


// ── HOJA DE DASHBOARD (ranking de vendedores por día + total) ────
// Lógica dinámica: detecta automáticamente los 3 días más recientes con ventas
// (cualquier fila de Ventas que tenga vendedor cargado).
//
// Referencias a hoja Ventas: A=Timestamp(Date), B=Venta #, D=Vendedor,
// R=Cant.cajas, V=Total U$D, W=Total ARS
function crearHojaDashboard(sheet) {
  // ── Título principal ──
  sheet.getRange('A1').setValue('RANKING DE VENDEDORES');
  sheet.getRange('A1:E1').merge()
    .setBackground('#1B3A52').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  // ── Panel "Días detectados" (columnas G-H, a modo de referencia) ──
  // Fórmula para la k-ésima fecha más reciente entre ventas con vendedor.
  // FILTER devuelve los timestamps de Ventas!A donde Ventas!D no está vacía;
  // INT los convierte a fecha pura; UNIQUE elimina duplicados; LARGE toma el k-ésimo.
  function fechaRecienteK(k) {
    return '=IFERROR(LARGE(UNIQUE(INT(FILTER(Ventas!A2:A, Ventas!D2:D<>""))), ' + k + '), "")';
  }

  sheet.getRange('G1').setValue('DÍAS DETECTADOS (AUTO)')
    .setFontWeight('bold').setFontSize(10).setFontColor('#1B3A52');
  sheet.getRange('G1:H1').merge().setBackground('#E1F5EE');
  sheet.getRange('G2').setValue('1º (+ reciente):');
  sheet.getRange('G3').setValue('2º:');
  sheet.getRange('G4').setValue('3º:');
  sheet.getRange('H2').setFormula(fechaRecienteK(1));
  sheet.getRange('H3').setFormula(fechaRecienteK(2));
  sheet.getRange('H4').setFormula(fechaRecienteK(3));
  sheet.getRange('H2:H4').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('G2:G4').setFontColor('#666').setFontSize(10);

  // ── Fórmulas de ranking ──
  // Título dinámico: "Ranking del dd/mm/yyyy" o "(día sin ventas aún)"
  function tituloDia(cellFecha) {
    return '=IF(' + cellFecha + '="","(día sin ventas aún)","Ranking del "&TEXT(' + cellFecha + ',"dd/mm/yyyy"))';
  }

  // QUERY dinámica por día: filtra ventas cuyo timestamp cae en el día indicado por la celda
  function formulaRankingDia(cellFecha) {
    return '=IF(' + cellFecha + '="","Sin ventas aún",IFERROR(QUERY(Ventas!A:W,'
      + '"SELECT D, COUNT(B), SUM(V), SUM(W), SUM(R) '
      + 'WHERE A >= date \'"&TEXT(' + cellFecha + ',"yyyy-mm-dd")&"\' '
      + 'AND A < date \'"&TEXT(' + cellFecha + '+1,"yyyy-mm-dd")&"\' '
      + 'AND D IS NOT NULL AND D <> \'\' '
      + 'GROUP BY D ORDER BY COUNT(B) DESC '
      + 'LABEL D \'Vendedor\', COUNT(B) \'# Ventas\', SUM(V) \'Total U$D\', SUM(W) \'Total ARS\', SUM(R) \'Cajas\'"'
      + ',1),"Sin ventas aún"))';
  }

  // QUERY para el total histórico (sin filtro de fecha)
  const formulaRankingTotal =
    '=IFERROR(QUERY(Ventas!A:W,'
    + '"SELECT D, COUNT(B), SUM(V), SUM(W), SUM(R) '
    + 'WHERE D IS NOT NULL AND D <> \'\' '
    + 'GROUP BY D ORDER BY COUNT(B) DESC '
    + 'LABEL D \'Vendedor\', COUNT(B) \'# Ventas\', SUM(V) \'Total U$D\', SUM(W) \'Total ARS\', SUM(R) \'Cajas\'"'
    + ',1),"Sin ventas aún")';

  // ── Render de las 3 tablas por día ──
  const CELDAS_FECHA = ['H2', 'H3', 'H4'];
  let row = 3;
  CELDAS_FECHA.forEach(function(cell) {
    sheet.getRange(row, 1).setFormula(tituloDia(cell))
      .setFontWeight('bold').setFontSize(11).setFontColor('#1B3A52');
    sheet.getRange(row, 1, 1, 5).setBackground('#E1F5EE');
    row++;
    sheet.getRange(row, 1).setFormula(formulaRankingDia(cell));
    row += 16; // espacio reservado para la tabla + separación
  });

  // ── Tabla total ──
  sheet.getRange(row, 1).setValue('RANKING TOTAL (TODAS LAS VENTAS)')
    .setFontWeight('bold').setFontSize(11).setFontColor('#ffffff');
  sheet.getRange(row, 1, 1, 5).setBackground('#1D9E75');
  row++;
  sheet.getRange(row, 1).setFormula(formulaRankingTotal);

  // ── Anchos de columna ──
  sheet.setColumnWidth(1, 160); // Vendedor
  sheet.setColumnWidth(2, 100); // # Ventas
  sheet.setColumnWidth(3, 120); // Total U$D
  sheet.setColumnWidth(4, 130); // Total ARS
  sheet.setColumnWidth(5, 100); // Cajas
  sheet.setColumnWidth(7, 140); // label "días detectados"
  sheet.setColumnWidth(8, 110); // fechas
  sheet.setFrozenRows(1);
}



// ── PDF RECIBO ───────────────────────────────────────────────────
/**
 * Convierte el HTML del recibo a PDF usando Drive API:
 * 1. Sube el HTML como Google Doc (Drive lo convierte automáticamente)
 * 2. Exporta el Google Doc como PDF
 * 3. Guarda el PDF en la carpeta PDF_FOLDER_ID
 * 4. Borra el Google Doc temporal
 * Devuelve { blob, url }
 */
function generarPdfRecibo(p) {
  const html   = buildReciboHTML(p);
  const nombre = 'Recibo_AGF_Venta' + p.ventaNum + '.pdf';
  const token  = ScriptApp.getOAuthToken();

  // ── 1. Subir HTML como Google Doc ──
  const boundary = 'agf_pdf_boundary';
  const body = '--' + boundary + '\r\n'
    + 'Content-Type: application/json; charset=UTF-8\r\n\r\n'
    + JSON.stringify({
        name: '_temp_recibo_' + p.ventaNum,
        mimeType: 'application/vnd.google-apps.document'
      })
    + '\r\n--' + boundary + '\r\n'
    + 'Content-Type: text/html; charset=UTF-8\r\n\r\n'
    + html
    + '\r\n--' + boundary + '--';

  const uploadResp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'POST',
      contentType: 'multipart/related; boundary=' + boundary,
      headers: { Authorization: 'Bearer ' + token },
      payload: body,
      muteHttpExceptions: true
    }
  );
  const docId = JSON.parse(uploadResp.getContentText()).id;
  if (!docId) throw new Error('Error creando Google Doc temporal: ' + uploadResp.getContentText());

  try {
    // ── 2. Exportar como PDF ──
    const exportResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + docId + '/export?mimeType=application/pdf',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    const pdfBlob = exportResp.getBlob();
    pdfBlob.setName(nombre);

    // ── 3. Guardar PDF en carpeta definitiva ──
    const folder    = DriveApp.getFolderById(PDF_FOLDER_ID);
    const savedFile = folder.createFile(pdfBlob);
    savedFile.setName(nombre);

    const emailBlob = savedFile.getBlob();
    emailBlob.setName(nombre);
    return { blob: emailBlob, url: savedFile.getUrl() };

  } finally {
    // ── 4. Borrar Google Doc temporal siempre ──
    try { DriveApp.getFileById(docId).setTrashed(true); } catch(e) {}
  }
}

// ── FORMATEAR FECHA PARA DOCUMENTOS ─────────────────────────────
function formatFechaDoc(fechaStr) {
  var MESES = ['enero','febrero','marzo','abril','mayo','junio',
               'julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var m = (fechaStr || '').match(/^(\d+)\/(\d+)\/(\d+)/);
  if (!m) return fechaStr || '';
  return m[1] + ' de ' + MESES[parseInt(m[2], 10) - 1] + ' de ' + m[3];
}

/**
 * HTML del recibo — diseño clínico v8. Layout 100% tablas (Google Docs ignora flexbox/grid).
 * Todos los ternarios pre-computados antes del return para evitar SyntaxError en Apps Script.
 */
function buildReciboHTML(p) {
  var c    = p.cliente     || {};
  var f    = p.facturacion || {};
  var DARK = '#1B3A52';

  var SH  = 'font-size:7px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:#999;border-bottom:1px solid #eaeae8;padding-bottom:4px;margin-bottom:10px;margin-top:18px';
  var SH0 = 'font-size:7px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:#999;border-bottom:1px solid #eaeae8;padding-bottom:4px;margin-bottom:10px;margin-top:0';
  var DL  = 'padding:2px 0;color:#aaa;font-size:10px;width:105px';
  var DV  = 'padding:2px 0;color:#1a1a1a;font-size:11px;font-weight:500';

  var fechaDoc = formatFechaDoc(p.fecha);

  // ── Filas de cliente ──
  var cuitRow  = c.cuit      ? '<tr><td style="' + DL + '">CUIT/CUIL</td><td style="' + DV + '">' + c.cuit      + '</td></tr>' : '';
  var emailRow = c.mail      ? '<tr><td style="' + DL + '">Email</td><td style="' + DV + '">' + c.mail      + '</td></tr>' : '';
  var telRow   = c.tel       ? '<tr><td style="' + DL + '">Tel\u00e9fono</td><td style="' + DV + '">' + c.tel       + '</td></tr>' : '';
  var locRow   = c.localidad ? '<tr><td style="' + DL + '">Localidad</td><td style="' + DV + '">' + c.localidad + '</td></tr>' : '';

  // ── Filas de detalle (precio de lista, sin descuentos inline) ──
  var filasDetalle = (p.cajas || []).map(function(l) {
    var tipo    = l.tipo === 'cerrada' ? 'Cerrada' : 'Combinada';
    var listaPx = l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA;
    return '<tr style="border-bottom:1px solid #f2f2f0">'
      + '<td style="padding:6px 0;font-size:10px;color:#555;width:100px">Caja ' + l.caja + ' \u2014 ' + tipo + '</td>'
      + '<td style="padding:6px 8px;font-size:10px;color:#777">' + l.detalle + '</td>'
      + '<td style="padding:6px 0;text-align:right;font-size:10px;color:#888;white-space:nowrap">u$ ' + listaPx + '</td>'
      + '</tr>';
  }).join('');

  // ── Cálculo de descuentos ──
  var precioLista = (p.cajas || []).reduce(function(acc, l) {
    return acc + (l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA);
  }, 0);
  var totalDescuentos = precioLista - (p.totalUSD || 0);
  var hayDescuentos   = totalDescuentos > 0;

  // Sub-líneas por fuente de descuento (grisadas)
  var subDescRows = '';
  if (hayDescuentos) {
    (p.cajas || []).forEach(function(l) {
      if (l.descCaja > 0) {
        var listaPx = l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA;
        var monto   = Math.round(listaPx * l.descCaja / 100);
        var tipo    = l.tipo === 'cerrada' ? 'Cerrada' : 'Combinada';
        subDescRows += '<tr>'
          + '<td style="padding:2px 0 2px 12px;font-size:9px;color:#bbb">Caja ' + l.caja + ' \u2014 ' + tipo + '</td>'
          + '<td style="padding:2px 0;text-align:right;font-size:9px;color:#bbb;white-space:nowrap">\u2212' + l.descCaja + '%</td>'
          + '<td style="padding:2px 0 2px 10px;text-align:right;font-size:9px;color:#bbb;white-space:nowrap">\u2212u$ ' + monto + '</td>'
          + '</tr>';
      }
    });
    if (p.descuentoGlobal > 0) {
      var subtotalDescCaja = (p.cajas || []).reduce(function(acc, l) {
        return acc + Math.round((l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA) * (1 - (l.descCaja || 0) / 100));
      }, 0);
      var descGlobalMonto = subtotalDescCaja - (p.totalUSD || 0);
      subDescRows += '<tr>'
        + '<td style="padding:2px 0 2px 12px;font-size:9px;color:#bbb">Desc. global</td>'
        + '<td style="padding:2px 0;text-align:right;font-size:9px;color:#bbb;white-space:nowrap">\u2212' + p.descuentoGlobal + '%</td>'
        + '<td style="padding:2px 0 2px 10px;text-align:right;font-size:9px;color:#bbb;white-space:nowrap">\u2212u$ ' + descGlobalMonto + '</td>'
        + '</tr>';
    }
  }

  // Fila resumen de descuentos (precio lista + línea principal + sub-líneas)
  var resumenRows = '';
  if (hayDescuentos) {
    var pctRaw = precioLista > 0 ? totalDescuentos / precioLista * 100 : 0;
    var pctFmt = (pctRaw % 1 === 0 ? pctRaw.toFixed(0) : pctRaw.toFixed(1)).replace('.', ',') + '%';
    resumenRows = '<tr><td colspan="3" style="padding-top:10px">'
      + '<table cellpadding="0" cellspacing="0" width="100%">'
      + '<tr>'
      +   '<td style="padding:3px 0;font-size:10px;color:#aaa">Precio de lista</td>'
      +   '<td style="padding:3px 0;font-size:10px;color:#aaa;text-align:right;white-space:nowrap"></td>'
      +   '<td style="padding:3px 0 3px 10px;text-align:right;font-size:10px;color:#aaa;white-space:nowrap">u$ ' + precioLista + '</td>'
      + '</tr>'
      + '<tr>'
      +   '<td style="padding:3px 0;font-size:10px;color:#1a1a1a">Descuentos aplicados</td>'
      +   '<td style="padding:3px 0;text-align:right;font-size:10px;color:#1a1a1a;white-space:nowrap">\u2212' + pctFmt + '</td>'
      +   '<td style="padding:3px 0 3px 10px;text-align:right;font-size:10px;color:#1a1a1a;white-space:nowrap">\u2212u$ ' + totalDescuentos + '</td>'
      + '</tr>'
      + subDescRows
      + '</table>'
      + '</td></tr>';
  }

  // ── Total ARS ──
  var totalARSStr = p.totalARS
    ? '<div style="font-size:10px;color:#aaa;margin-top:3px">AR$ ' + Number(p.totalARS).toLocaleString('es-AR') + '</div>'
    : '';

  // ── Cobro ──
  var tcRow = p.tipoCambio
    ? '<tr><td style="' + DL + '">Tipo de cambio</td><td style="' + DV + '">AR$ ' + Number(p.tipoCambio).toLocaleString('es-AR') + ' / U$D</td></tr>'
    : '';

  // ── Facturación (solo si aplica; mismosContacto usa datos del cliente) ──
  var facturacionBlock = '';
  if (p.condFiscal) {
    var facRazon, facCuit;
    if (f.mismosContacto) {
      facRazon = ((c.nombre || '') + ' ' + (c.apellido || '')).trim();
      facCuit  = c.cuit || '';
    } else {
      facRazon = f.razonSocial     || '';
      facCuit  = f.cuitFacturacion || '';
    }
    var facRazonRow = facRazon ? '<tr><td style="' + DL + '">Raz\u00f3n social</td><td style="' + DV + '">' + facRazon + '</td></tr>' : '';
    var facCuitRow  = facCuit  ? '<tr><td style="' + DL + '">CUIT</td><td style="' + DV + '">' + facCuit + '</td></tr>'               : '';
    var condRow     = '<tr><td style="' + DL + '">Cond. fiscal</td><td style="' + DV + '">' + p.condFiscal + '</td></tr>';
    facturacionBlock = '<div style="' + SH + '">Facturaci\u00f3n</div>'
      + '<table cellpadding="0" cellspacing="0">'
      + facRazonRow + facCuitRow + condRow
      + '</table>';
  }

  // ════ CONSTRUCCIÓN DEL HTML ════
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">'
    + '<style>body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#1a1a1a;background:#ececea;padding:24px}table{border-collapse:collapse}td,th{vertical-align:top}</style>'
    + '</head><body>'
    + '<table align="center" cellpadding="0" cellspacing="0" width="640" bgcolor="#ffffff">'

    // ── HEADER ──
    + '<tr><td style="padding:22px 36px 18px;border-bottom:1px solid #eaeae8">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td valign="top">'
    +       '<div style="font-family:Georgia,serif;font-size:19px;font-weight:700;color:' + DARK + ';letter-spacing:-.2px">AGF Mesenchymal</div>'
    +       '<div style="font-size:9px;color:#aaa;font-style:italic;margin-top:2px">Argentina</div>'
    +       '<div style="font-size:7px;color:#ccc;letter-spacing:.18em;text-transform:uppercase;margin-top:5px">BAAS 2026</div>'
    +     '</td>'
    +     '<td valign="top" align="right">'
    +       '<div style="font-size:7px;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:4px">Comprobante de compra</div>'
    +       '<div style="font-family:Georgia,serif;font-size:20px;font-weight:400;color:' + DARK + ';line-height:1">N.\u00ba ' + p.ventaNum + '</div>'
    +       '<div style="font-size:9px;color:#bbb;margin-top:4px">' + fechaDoc + '</div>'
    +     '</td>'
    +   '</tr></table>'
    + '</td></tr>'

    // ── CONTENIDO ──
    + '<tr><td style="padding:20px 36px 28px">'

    //   DATOS DEL CLIENTE
    + '<div style="' + SH0 + '">Datos del cliente</div>'
    + '<table cellpadding="0" cellspacing="0" style="margin-bottom:0">'
    + '<tr><td style="' + DL + '">Nombre</td><td style="' + DV + '">' + (c.nombre || '') + ' ' + (c.apellido || '') + '</td></tr>'
    + cuitRow + emailRow + telRow + locRow
    + '</table>'

    //   DETALLE
    + '<div style="' + SH + '">Detalle</div>'
    + '<table cellpadding="0" cellspacing="0" width="100%">'
    + '<thead><tr style="border-bottom:1px solid #eaeae8">'
    + '<th style="padding:4px 0;text-align:left;font-size:7px;font-weight:600;letter-spacing:.15em;text-transform:uppercase;color:#bbb;width:100px">Caja</th>'
    + '<th style="padding:4px 8px;text-align:left;font-size:7px;font-weight:600;letter-spacing:.15em;text-transform:uppercase;color:#bbb">Contenido</th>'
    + '<th style="padding:4px 0;text-align:right;font-size:7px;font-weight:600;letter-spacing:.15em;text-transform:uppercase;color:#bbb;white-space:nowrap">Lista</th>'
    + '</tr></thead>'
    + '<tbody>' + filasDetalle + resumenRows + '</tbody>'
    + '</table>'

    //   TOTAL
    + '<table cellpadding="0" cellspacing="0" width="100%" style="margin-top:14px;padding-top:12px;border-top:1px solid #eaeae8"><tr>'
    + '<td style="font-size:7px;letter-spacing:.2em;text-transform:uppercase;color:#bbb;padding-top:6px">Total</td>'
    + '<td align="right">'
    +   '<div style="font-family:Georgia,serif;font-size:24px;font-weight:700;color:' + DARK + ';line-height:1">u$ ' + p.totalUSD + '</div>'
    +   totalARSStr
    + '</td>'
    + '</tr></table>'

    //   CONDICIONES DE COBRO
    + '<div style="' + SH + '">Condiciones de cobro</div>'
    + '<table cellpadding="0" cellspacing="0">'
    + '<tr><td style="' + DL + '">M\u00e9todo</td><td style="' + DV + '">' + (p.metodoCobro || '') + '</td></tr>'
    + '<tr><td style="' + DL + '">Moneda</td><td style="' + DV + '">' + (p.moneda === 'USD' ? 'D\u00f3lares americanos' : 'Pesos argentinos') + '</td></tr>'
    + tcRow
    + '</table>'

    //   FACTURACIÓN
    + facturacionBlock

    + '</td></tr>'

    // ── PIE ──
    + '<tr><td style="padding:10px 36px;border-top:1px solid #eaeae8">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td style="font-size:8px;color:#ccc">Documento v\u00e1lido como comprobante de compra</td>'
    +     '<td align="right" style="font-size:8px;color:#ccc">AGF Mesenchymal Argentina &middot; BAAS 2026</td>'
    +   '</tr></table>'
    + '</td></tr>'

    + '</table>'
    + '</body></html>';
}


// ── EMAIL CLIENTE ────────────────────────────────────────────────
function enviarEmailCliente(p, pdfBlob) {
  const c = p.cliente;
  GmailApp.sendEmail(
    c.mail,
    'Gracias por elegirnos \uD83D\uDC99 Comprobante N.\u00ba ' + p.ventaNum + ' \u2014 AGF Mesenchymal Argentina',
    buildTextoPlano(p),
    { htmlBody: buildEmailHTML(p, false), name: 'AGF Mesenchymal Argentina',
      replyTo: REPLY_TO, attachments: [pdfBlob] }
  );
}

function enviarEmailAdmin(p, pdfBlob) {
  const c = p.cliente || {};
  GmailApp.sendEmail(
    ADMIN_EMAIL,
    '[Venta #' + p.ventaNum + '] ' + c.nombre + ' ' + c.apellido + ' — u$' + p.totalUSD + ' — ' + p.metodoCobro,
    buildTextoPlano(p),
    { htmlBody: buildEmailHTML(p, true), name: 'AGF Ventas Congreso', attachments: [pdfBlob] }
  );
}


// ── TEXTO PLANO ──────────────────────────────────────────────────
function buildTextoPlano(p) {
  var c        = p.cliente || {};
  var saludo   = c.nombre ? c.nombre + ',' : 'Estimado cliente,';
  var contacto = p.vendedor
    ? 'Ante cualquier inquietud, no dudes en contactarte con ' + p.vendedor + ' o con Alan J. Haslop.'
    : 'Ante cualquier inquietud, no dudes en contactarte con Alan J. Haslop.';
  return [
    saludo + ' gracias por confiar en nosotros.',
    '',
    'Adjunto el comprobante de tu compra. Tus productos ya estan reservados en el stand B28 al final de la expo. Cuando puedas, te esperamos para que retires tu pedido.',
    '',
    contacto,
    '',
    '--',
    'Alan J. Haslop',
    'Director Ejecutivo',
    'AGF Mesenchymal Argentina',
    '',
    '---',
    'Venta #' + p.ventaNum + ' | ' + (p.fecha || ''),
    'Total: u$' + p.totalUSD + (p.totalARS ? ' / AR$' + Number(p.totalARS).toLocaleString('es-AR') : ''),
    'Metodo: ' + (p.metodoCobro || '') + ' | Moneda: ' + (p.moneda || 'USD')
  ].join('\n');
}


// ── EMAIL HTML (cuerpo del mensaje) ─────────────────────────────
function buildEmailHTML(p, isAdmin) {
  var c    = p.cliente || {};
  var DARK = '#1B3A52';

  var saludo   = c.nombre ? c.nombre + ',' : 'Estimado cliente,';
  var contacto = p.vendedor
    ? 'Ante cualquier inquietud, no dudes en contactarte con <strong>' + p.vendedor + '</strong> o conmigo.'
    : 'Ante cualquier inquietud, no dudes en contactarte conmigo.';

  var adminBanner = isAdmin
    ? '<tr><td style="padding:0 36px 16px">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +   '<td style="background:#fff3cd;border:1px solid #ffc107;border-radius:4px;padding:10px 14px;font-size:12px;color:#856404;font-family:Arial,Helvetica,sans-serif">'
    +   '<strong>Copia administrador</strong> &mdash; Venta #' + p.ventaNum + ' &middot; ' + (p.dispositivo || '?') + ' &middot; ' + (p.fecha || '')
    +   '</td></tr></table>'
    +   '</td></tr>'
    : '';

  return '<!DOCTYPE html><html lang="es"><head>'
    + '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '</head>'
    + '<body style="margin:0;padding:0;background:#f5f5f3">'
    + '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    + '<td align="center" style="padding:32px 16px">'
    + '<table cellpadding="0" cellspacing="0" width="520">'
    + '<tr><td style="background:#ffffff;border-radius:2px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.08)">'
    + '<table cellpadding="0" cellspacing="0" width="100%">'

    // Línea de acento
    + '<tr><td style="height:3px;background:' + DARK + '"></td></tr>'

    // Header
    + '<tr><td style="padding:24px 36px 0">'
    +   '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:15px;font-weight:700;color:' + DARK + '">AGF Mesenchymal</div>'
    +   '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:10px;color:#aaa;font-style:italic;margin-top:1px">Argentina</div>'
    + '</td></tr>'

    // Banner admin (vacío si no es admin)
    + adminBanner

    // Cuerpo
    + '<tr><td style="padding:22px 36px 28px;font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.75;color:#2a2a2a">'
    +   '<p style="margin:0 0 12px">' + saludo + ' gracias por confiar en nosotros.</p>'
    +   '<p style="margin:0 0 12px">Adjunto el comprobante de tu compra. Tus productos ya est\u00e1n reservados en el stand B28 al final de la expo. Cuando puedas, te esperamos para que retires tu pedido.</p>'
    +   '<p style="margin:0 0 28px">' + contacto + '</p>'

    // Firma
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +   '<td style="border-top:1px solid #e8e8e0;padding-top:18px">'
    +     '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:13px;font-weight:700;color:' + DARK + '">Alan J. Haslop</div>'
    +     '<div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#999;margin-top:3px">Director Ejecutivo</div>'
    +     '<div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#999">AGF Mesenchymal <em>Argentina</em></div>'
    +   '</td></tr></table>'
    + '</td></tr>'

    // Footer
    + '<tr><td style="background:#f7f7f5;border-top:1px solid #eeece6;padding:10px 36px">'
    +   '<span style="font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#bbb">AGF Mesenchymal &middot; BAAS 2026</span>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr>'
    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';
}


// ── UTILIDADES ───────────────────────────────────────────────────
function obtenerOCrearHoja(nombre, fnHeaders) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    fnHeaders(sheet);
  }
  return sheet;
}


// ── TEST MANUAL ──────────────────────────────────────────────────
function testManual() {
  const basePayload = {
    ventaNum: 99,
    fecha: '22/4/2026, 10:00:00',
    dispositivo: 'TEST',
    cliente: { nombre:'Ana', apellido:'García', cuit:'20-12345678-0', mail:'alan.haslop@dermacells.com.ar', tel:'', localidad:'CABA' },
    facturacion: { razonSocial:'Empresa SA', cuitFacturacion:'30-99999999-0', mismosContacto:false },
    condFiscal: 'Resp. Inscripto',
    cajas: [
      { caja:1, tipo:'cerrada',   detalle:'Dermal',                                   descCaja:0,  precio:750 },
      { caja:2, tipo:'combinada', detalle:'Dermal, Capillary, Pink, Biomask, Dermal', descCaja:10, precio:810 },
      { caja:3, tipo:'cerrada',   detalle:'Capillary',                                descCaja:0,  precio:750 }
    ],
    totalUSD: 2310,
    totalARS: null,
    moneda: 'USD',
    tipoCambio: null,
    metodoCobro: 'Transferencia',
    descuentoGlobal: 0
  };

  // ── Unidades esperadas: Dermal=11 (5+2+0), Capillary=6 (0+1+5), Pink=1, Biomask=1 ──
  const u = calcUnidades(basePayload);
  Logger.log('Unidades calculadas: ' + JSON.stringify(u));

  // ── PDF de prueba (una sola vez, vendedor = Gabriel) ──
  const pdfPayload = Object.assign({}, basePayload, { vendedor:'Gabriel' });
  const { blob, url } = generarPdfRecibo(pdfPayload);
  Logger.log('PDF generado: ' + blob.getName() + ' (' + blob.getBytes().length + ' bytes)');
  Logger.log('URL en Drive: ' + url);

  // ── 4 ventas en 3 días distintos con 4 vendedores distintos ──
  // Poblará el Dashboard con las 3 tablas diarias + el ranking total.
  // NOTA: el Resumen agrupa por "hoy" (no usa timestamp falseado), así que todas
  // las filas del Resumen quedan con la fecha de hoy. El Dashboard sí lee el
  // timestamp real de la col A de Ventas y por eso distribuye bien los días.
  const ventasTest = [
    { vendedor:'Gabriel',     diasAtras:0, ventaNum:991 },
    { vendedor:'Federico',    diasAtras:0, ventaNum:992 },
    { vendedor:'Cecilia',     diasAtras:1, ventaNum:993 },
    { vendedor:'Calcopietro', diasAtras:2, ventaNum:994 }
  ];

  const sheet = obtenerOCrearHoja(SHEET_VENTAS, crearHeadersVentas);

  ventasTest.forEach(function(t) {
    const p = Object.assign({}, basePayload, { vendedor:t.vendedor, ventaNum:t.ventaNum });
    guardarVenta(p, url);
    // Sobreescribir col A (timestamp) con fecha falseada para poder testear Dashboard
    const lastRow = sheet.getLastRow();
    const fakeDate = new Date();
    fakeDate.setDate(fakeDate.getDate() - t.diasAtras);
    sheet.getRange(lastRow, 1).setValue(fakeDate);
    actualizarResumen(p);
    Logger.log('Venta test guardada: #' + t.ventaNum + ' · ' + t.vendedor + ' · ' + fakeDate.toDateString());
  });

  obtenerOCrearHoja(SHEET_STOCK, crearHojaStock);
  obtenerOCrearHoja(SHEET_DASHBOARD, crearHojaDashboard);

  // Un solo mail (no spamear con 4)
  enviarEmailCliente(pdfPayload, blob);
  Logger.log('Email cliente enviado a: ' + pdfPayload.cliente.mail);
  enviarEmailAdmin(pdfPayload, blob);
  Logger.log('Email admin enviado a: ' + ADMIN_EMAIL);

  Logger.log('testManual completado OK — revisá la hoja Dashboard');
}

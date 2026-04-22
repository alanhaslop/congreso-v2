// ═══════════════════════════════════════════════════════════════
//  AGF Messenchymal — Ventas Congreso · Google Apps Script
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

/**
 * HTML del recibo — diseño v7, layout 100% tablas (Google Docs ignora flexbox/grid).
 * Sin firmas. Todos los ternarios condicionales pre-computados antes del return
 * para evitar SyntaxError "Unexpected token ?" / ":" en Apps Script.
 */
function buildReciboHTML(p) {
  var c    = p.cliente     || {};
  var f    = p.facturacion || {};
  var DARK  = '#1B3A52';
  var GREEN = '#9FD4C0';
  var u    = calcUnidades(p);

  // ── Estilos reutilizables ──
  var ST   = 'font-size:8px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:' + DARK + ';border-bottom:1px solid #d8d4cc;padding-bottom:4px;margin-bottom:10px;margin-top:16px';
  var ST0  = 'font-size:8px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:' + DARK + ';border-bottom:1px solid #d8d4cc;padding-bottom:4px;margin-bottom:10px;margin-top:0';
  var DL   = 'padding:2px 0;color:#bbb;font-size:10px;width:90px';
  var DV   = 'padding:2px 0;color:#1a1a1a;font-weight:500;font-size:11px';
  var CL   = 'padding:2px 0;color:#bbb;font-size:10px;width:120px';
  var CV   = 'padding:2px 0;color:#1a1a1a;font-weight:500;font-size:11px';

  // ── Todas las celdas condicionales pre-computadas (evita "+ ?" o "+ :" en return) ──
  var cuitRow      = c.cuit      ? '<tr><td style="' + DL + '">CUIT/CUIL</td><td style="' + DV + '">' + c.cuit      + '</td></tr>' : '';
  var emailRow     = c.mail      ? '<tr><td style="' + DL + '">Email</td><td style="' + DV + '">' + c.mail      + '</td></tr>' : '';
  var telRow       = c.tel       ? '<tr><td style="' + DL + '">Teléfono</td><td style="' + DV + '">' + c.tel       + '</td></tr>' : '';
  var localRow     = c.localidad ? '<tr><td style="' + DL + '">Localidad</td><td style="' + DV + '">' + c.localidad + '</td></tr>' : '';
  var condClientRow = p.condFiscal ? '<tr><td style="' + DL + '">Cond. fiscal</td><td style="' + DV + '">' + p.condFiscal + '</td></tr>' : '';

  // ── Filas de cajas ──
  var filas = (p.cajas || []).map(function(l) {
    var tipo       = l.tipo === 'cerrada' ? 'Cerrada' : 'Combinada';
    var tagBg      = l.tipo === 'cerrada' ? '#EEF4F8' : '#f0faf6';
    var descBadge  = l.descCaja > 0 ? ' <span style="color:#e67e22;font-size:10px;font-style:italic">&#8212; desc. ' + l.descCaja + '%</span>' : '';
    return '<tr style="border-bottom:1px solid #f5f2ec">'
      + '<td style="padding:6px 8px;width:80px"><table cellpadding="0" cellspacing="0"><tr>'
      + '<td bgcolor="' + tagBg + '" style="padding:2px 7px;font-size:9px;font-weight:500;color:' + DARK + '">' + tipo + '</td>'
      + '</tr></table></td>'
      + '<td style="padding:6px 8px;color:#777;font-size:10.5px">' + l.detalle + descBadge + '</td>'
      + '<td style="padding:6px 8px;text-align:right;font-weight:600;color:' + DARK + ';font-size:11px;white-space:nowrap">u$' + l.precio + '</td>'
      + '</tr>';
  }).join('');

  // ── Subtotal + descuento global ──
  var subtotalUSD = (p.cajas || []).reduce(function(acc, l) {
    return acc + Math.round((l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA) * (1 - (l.descCaja || 0) / 100));
  }, 0);
  var ahorroUSD = (p.descuentoGlobal > 0 && p.descuentoGlobal < 100) ? subtotalUSD - p.totalUSD : 0;
  var subtotalRows = p.descuentoGlobal > 0
    ? '<tr><td colspan="2" style="padding:4px 8px;font-size:10px;color:#999">Subtotal</td><td style="padding:4px 8px;text-align:right;font-size:10px;color:#999">u$' + subtotalUSD + '</td></tr>'
      + '<tr><td colspan="2" style="padding:4px 8px;font-size:10px;color:#e67e22">Desc. general ' + p.descuentoGlobal + '%</td><td style="padding:4px 8px;text-align:right;font-size:10px;color:#e67e22;font-weight:600">' + (ahorroUSD > 0 ? '&#8722; u$' + ahorroUSD : '') + '</td></tr>'
    : '';

  // ── Filas de cobro ──
  var metodoRow   = p.metodoCobro  ? '<tr><td style="' + CL + '">Método de pago</td><td style="' + CV + '">' + p.metodoCobro + '</td></tr>' : '';
  var tcCobroRow  = p.tipoCambio   ? '<tr><td style="' + CL + '">Tipo de cambio</td><td style="' + CV + '">AR$' + Number(p.tipoCambio).toLocaleString('es-AR') + ' / U$D</td></tr>' : '';
  var descCobroRow = p.descuentoGlobal > 0 ? '<tr><td style="' + CL + '">Descuento general</td><td style="padding:2px 0;color:#e67e22;font-weight:600;font-size:11px">' + p.descuentoGlobal + '% &#8212; u$' + ahorroUSD + ' de ahorro</td></tr>' : '';
  var totalARSStr = p.totalARS     ? '<div style="font-size:11px;color:#aaa;margin-top:3px">AR$ ' + Number(p.totalARS).toLocaleString('es-AR') + '</div>' : '';

  // ── Bloque de facturación ──
  var instrMap = {
    'Resp. Inscripto': { bg:'#eff6ff', border:'#93c5fd', color:'#1e3a8a', icono:'&#9650;', texto:'Se emitirá Factura A. Verificar CUIT y razón social con administración antes de procesar.' },
    'Monotributista':  { bg:'#f0fdf4', border:'#86efac', color:'#14532d', icono:'&#9679;', texto:'Se emitirá Factura B. Datos registrados para procesamiento posterior por administración.' },
    'Cons. Final':     { bg:'#f0fdf4', border:'#86efac', color:'#14532d', icono:'&#9679;', texto:'Se emitirá Factura B (Consumidor Final). Sin CUIT específico requerido.' }
  };
  var instrData = p.condFiscal && instrMap[p.condFiscal] ? instrMap[p.condFiscal] : null;
  var facRazon  = (!f.mismosContacto && f.razonSocial)     ? f.razonSocial     : '';
  var facCuit   = (!f.mismosContacto && f.cuitFacturacion) ? f.cuitFacturacion : '';
  var facRazonTd = facRazon ? '<td style="padding-right:20px"><div style="font-size:8px;opacity:.7">Razón social</div><div style="font-size:10.5px;font-weight:600">' + facRazon + '</div></td>' : '';
  var facCuitTd  = facCuit  ? '<td><div style="font-size:8px;opacity:.7">CUIT</div><div style="font-size:10.5px;font-weight:600">' + facCuit  + '</div></td>' : '';
  var datosFactTr = (facRazon || facCuit) ? '<tr><td colspan="2" style="padding-top:7px;border-top:1px solid rgba(0,0,0,0.1)"><table cellpadding="0" cellspacing="0"><tr>' + facRazonTd + facCuitTd + '</tr></table></td></tr>' : '';
  var facturacionBlock = instrData
    ? '<table cellpadding="0" cellspacing="0" width="100%" style="margin-top:16px;border:1px solid ' + instrData.border + '"><tr>'
      + '<td bgcolor="' + instrData.bg + '" style="padding:10px 14px">'
      + '<table cellpadding="0" cellspacing="0" width="100%">'
      + '<tr><td style="font-size:8px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;color:' + instrData.color + ';padding-bottom:4px">' + instrData.icono + ' Instrucciones de facturación</td></tr>'
      + '<tr><td style="font-size:10.5px;color:' + instrData.color + ';line-height:1.5">' + instrData.texto + '</td></tr>'
      + datosFactTr
      + '</table></td></tr></table>'
    : '';

  // ════ CONSTRUCCIÓN DEL HTML ════
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">'
    + '<style>body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#1a1a1a;background:#e8e6e0;padding:24px}table{border-collapse:collapse}td,th{vertical-align:top}</style>'
    + '</head><body>'
    + '<table align="center" cellpadding="0" cellspacing="0" width="680" bgcolor="#ffffff">'

    // ── HEADER ──
    + '<tr><td style="padding:20px 36px 16px;border-bottom:1px solid #e0ddd6">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td valign="middle"><div style="font-size:22px;font-weight:700;color:' + DARK + ';letter-spacing:-.5px">AGF Messenchymal</div>'
    +       '<div style="font-size:10px;color:#aaa;margin-top:3px;font-style:italic">dermacells.com.ar &middot; Argentina</div></td>'
    +     '<td valign="middle" align="right"><div style="font-size:28px;font-weight:300;color:' + DARK + ';line-height:1">Venta #' + p.ventaNum + '</div>'
    +       '<div style="font-size:10px;color:#aaa;margin-top:4px">' + (p.fecha || '') + '</div></td>'
    +   '</tr></table>'
    + '</td></tr>'

    // ── BANDA ──
    + '<tr><td bgcolor="' + DARK + '" style="padding:7px 36px">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td style="font-size:9px;font-weight:500;letter-spacing:.2em;text-transform:uppercase;color:' + GREEN + '">Comprobante de compra</td>'
    +     '<td align="right" style="font-size:9px;font-weight:500;letter-spacing:.2em;text-transform:uppercase;color:' + GREEN + '">Dermacells S.A.</td>'
    +   '</tr></table>'
    + '</td></tr>'

    // ── CONTENIDO ──
    + '<tr><td style="padding:18px 36px 24px">'

    //   DATOS DEL CLIENTE
    +   '<div style="' + ST0 + '">Datos del cliente</div>'
    +   '<table cellpadding="0" cellspacing="0" style="margin-bottom:16px">'
    +     '<tr><td style="' + DL + '">Nombre</td><td style="' + DV + '">' + (c.nombre || '') + ' ' + (c.apellido || '') + '</td></tr>'
    +     cuitRow + emailRow + telRow + localRow + condClientRow
    +   '</table>'

    //   PRODUCTOS ADQUIRIDOS
    +   '<div style="' + ST + '">Productos adquiridos</div>'
    +   '<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:16px;border:1px solid #f0ede6"><tr>'
    +     '<td align="center" style="padding:8px;border-right:1px solid #f0ede6">'
    +       '<div style="font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb;margin-bottom:4px">Dermal</div>'
    +       '<div style="font-size:26px;font-weight:600;color:' + DARK + ';line-height:1">' + u.Dermal + '</div>'
    +       '<div style="font-size:7px;color:#bbb;letter-spacing:.1em;text-transform:uppercase;margin-top:2px">unidades</div>'
    +     '</td>'
    +     '<td align="center" style="padding:8px;border-right:1px solid #f0ede6">'
    +       '<div style="font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb;margin-bottom:4px">Capillary</div>'
    +       '<div style="font-size:26px;font-weight:600;color:' + DARK + ';line-height:1">' + u.Capillary + '</div>'
    +       '<div style="font-size:7px;color:#bbb;letter-spacing:.1em;text-transform:uppercase;margin-top:2px">unidades</div>'
    +     '</td>'
    +     '<td align="center" style="padding:8px;border-right:1px solid #f0ede6">'
    +       '<div style="font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb;margin-bottom:4px">Pink</div>'
    +       '<div style="font-size:26px;font-weight:600;color:' + DARK + ';line-height:1">' + u.Pink + '</div>'
    +       '<div style="font-size:7px;color:#bbb;letter-spacing:.1em;text-transform:uppercase;margin-top:2px">unidades</div>'
    +     '</td>'
    +     '<td align="center" style="padding:8px">'
    +       '<div style="font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb;margin-bottom:4px">Biomask</div>'
    +       '<div style="font-size:26px;font-weight:600;color:' + DARK + ';line-height:1">' + u.Biomask + '</div>'
    +       '<div style="font-size:7px;color:#bbb;letter-spacing:.1em;text-transform:uppercase;margin-top:2px">unidades</div>'
    +     '</td>'
    +   '</tr></table>'

    //   DETALLE DE CAJAS
    +   '<div style="' + ST + '">Detalle de cajas</div>'
    +   '<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:0">'
    +   '<thead><tr style="border-bottom:1px solid #e0ddd6">'
    +     '<th style="padding:5px 8px;text-align:left;font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb;width:80px">Tipo</th>'
    +     '<th style="padding:5px 8px;text-align:left;font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb">Contenido</th>'
    +     '<th style="padding:5px 8px;text-align:right;font-size:8px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bbb;white-space:nowrap">Precio</th>'
    +   '</tr></thead>'
    +   '<tbody>' + filas + subtotalRows + '</tbody>'
    +   '</table>'

    //   COBRO / TOTAL
    +   '<table cellpadding="0" cellspacing="0" width="100%" style="margin-top:16px;padding-top:12px;border-top:2px solid ' + DARK + '"><tr>'
    +     '<td valign="top">'
    +       '<table cellpadding="0" cellspacing="0">' + metodoRow + tcCobroRow + descCobroRow + '</table>'
    +     '</td>'
    +     '<td valign="top" align="right" style="white-space:nowrap">'
    +       '<div style="font-size:8px;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:3px">Total</div>'
    +       '<div style="font-size:36px;font-weight:600;color:' + DARK + ';line-height:1">u$' + p.totalUSD + '</div>'
    +       totalARSStr
    +     '</td>'
    +   '</tr></table>'

    //   FACTURACIÓN
    +   facturacionBlock

    + '</td></tr>'

    // ── PIE ──
    + '<tr bgcolor="#f7f6f2"><td style="padding:8px 36px;border-top:1px solid #e8e5de">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td style="font-size:9px;color:#ccc">Documento válido como comprobante de compra</td>'
    +     '<td align="right" style="font-size:9px;color:#ccc">AGF Messenchymal &middot; BAAS 2026</td>'
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
    'AGF Messenchymal — Comprobante de compra #' + p.ventaNum,
    buildTextoPlano(p),
    { htmlBody: buildEmailHTML(p, false), name: 'AGF Messenchymal Argentina',
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
  const c = p.cliente || {};
  const u = calcUnidades(p);
  const lineas = [
    'AGF Messenchymal — Venta #' + p.ventaNum + '  |  Dispositivo ' + (p.dispositivo || '?'),
    'Fecha: ' + p.fecha,
    'Cliente: ' + c.nombre + ' ' + c.apellido + '  |  CUIT: ' + c.cuit,
    '',
    'Detalle:',
    ...(p.cajas || []).map(function(cj) {
      return '  Caja ' + cj.caja + ' (' + cj.tipo + '): ' + cj.detalle
        + (cj.descCaja > 0 ? ' [-' + cj.descCaja + '%]' : '') + ' = u$' + cj.precio;
    }),
    '',
    p.descuentoGlobal > 0
      ? '  Descuento global: ' + p.descuentoGlobal + '%' : null,
    'Total: u$' + p.totalUSD + (p.totalARS ? '  /  AR$' + Number(p.totalARS).toLocaleString('es-AR') : ''),
    'Método: ' + p.metodoCobro + '  |  Moneda: ' + p.moneda + (p.tipoCambio ? '  |  TC: $' + p.tipoCambio : ''),
    p.condFiscal ? 'Cond. fiscal: ' + p.condFiscal : null,
    '',
    'Unidades: Dermal ' + u.Dermal + ' | Capillary ' + u.Capillary + ' | Pink ' + u.Pink + ' | Biomask ' + u.Biomask
  ].filter(function(l) { return l !== null; }).join('\n');
  return lineas;
}


// ── EMAIL HTML (cuerpo del mensaje) ─────────────────────────────
function buildEmailHTML(p, isAdmin) {
  const c     = p.cliente     || {};
  const f     = p.facturacion || {};
  const GREEN = '#1D9E75';
  const DARK  = '#1B3A52';
  const u     = calcUnidades(p);

  const filasHTML = (p.cajas || []).map(function(cj, i) {
    const bg = i % 2 === 0 ? '#f9f9f7' : '#ffffff';
    const descBadge = cj.descCaja > 0
      ? ' <span style="color:#e67e22;font-size:11px">[-' + cj.descCaja + '%]</span>' : '';
    return '<tr style="background:' + bg + '">'
      + '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-size:13px">Caja ' + cj.caja + ' — ' + (cj.tipo === 'cerrada' ? 'Cerrada' : 'Combinada') + '</td>'
      + '<td style="padding:8px 12px;border-bottom:1px solid #eee;color:#666;font-size:12px">' + cj.detalle + descBadge + '</td>'
      + '<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600;font-size:13px">u$' + cj.precio + '</td>'
      + '</tr>';
  }).join('');

  const subtotalUSD = (p.cajas || []).reduce(function(acc, l) {
    return acc + Math.round((l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA) * (1 - (l.descCaja || 0) / 100));
  }, 0);
  const descRowHTML = p.descuentoGlobal > 0
    ? '<tr><td colspan="2" style="padding:5px 12px;font-size:11px;color:#999;font-style:italic">Descuento global: ' + p.descuentoGlobal + '%</td>'
      + '<td style="padding:5px 12px;text-align:right;color:#c0392b;font-size:12px">−u$' + (subtotalUSD - p.totalUSD) + '</td></tr>'
    : '';

  const totalDisplay = 'u$' + p.totalUSD
    + (p.totalARS ? '&nbsp;<span style="font-size:14px;color:#666">/ AR$' + Number(p.totalARS).toLocaleString('es-AR') + '</span>' : '');

  const tcRowHTML = p.tipoCambio
    ? '<tr><td style="padding:3px 0;color:#999">Tipo de cambio</td><td style="padding:3px 0">$' + Number(p.tipoCambio).toLocaleString('es-AR') + ' AR$/U$D</td></tr>' : '';
  const condFiscalHTML = p.condFiscal
    ? (f.mismosContacto
        ? '<tr><td style="padding:3px 0;color:#999">Cond. fiscal</td><td style="padding:3px 0">' + p.condFiscal + '</td></tr>'
        : '<tr><td style="padding:3px 0;color:#999;vertical-align:top">Facturación</td><td style="padding:3px 0">' + (f.razonSocial || '') + '<br>CUIT: ' + (f.cuitFacturacion || '') + ' — ' + p.condFiscal + '</td></tr>')
    : '';
  const adminBannerHTML = isAdmin
    ? '<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:10px 16px;margin-bottom:18px;font-size:13px;color:#856404">'
      + '<strong>Copia administrador</strong> — Venta registrada ' + p.fecha + ' · Dispositivo&nbsp;' + (p.dispositivo || '?') + '</div>' : '';
  const unidadesHTML = '<div style="background:#f0f4f8;border-radius:6px;padding:10px 14px;margin-top:16px;font-size:12px">'
    + '<strong style="color:' + DARK + '">Unidades vendidas:</strong>&nbsp;&nbsp;'
    + 'Dermal <b>' + u.Dermal + '</b> &nbsp;|&nbsp; Capillary <b>' + u.Capillary + '</b> &nbsp;|&nbsp; Pink <b>' + u.Pink + '</b> &nbsp;|&nbsp; Biomask <b>' + u.Biomask + '</b>'
    + '</div>';

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;padding:0;background:#f5f5f3;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a">'
    + '<div style="max-width:560px;margin:24px auto;background:#ffffff;border-radius:10px;overflow:hidden;box-shadow:0 2px 14px rgba(0,0,0,.1)">'
    + '<div style="background:' + DARK + ';padding:20px 28px;display:flex;justify-content:space-between;align-items:center">'
    + '<div><div style="font-size:19px;font-weight:700;color:' + GREEN + '">AGF Messenchymal</div>'
    + '<div style="font-size:11px;color:#9FD4C0;margin-top:2px">dermacells.com.ar · Argentina</div></div>'
    + '<div style="text-align:right"><div style="font-size:10px;color:#9ab;letter-spacing:.1em;text-transform:uppercase">Comprobante de compra</div>'
    + '<div style="font-size:14px;font-weight:700;color:' + GREEN + ';margin-top:4px">Venta #' + p.ventaNum + '</div>'
    + '<div style="font-size:10px;color:#9ab;margin-top:2px">BAAS 2026</div></div></div>'
    + '<div style="padding:24px 28px">' + adminBannerHTML
    + '<div style="font-size:10px;font-weight:700;letter-spacing:.15em;color:' + DARK + ';text-transform:uppercase;border-bottom:1px solid #e8e8e0;padding-bottom:4px;margin-bottom:10px">Datos del cliente</div>'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px">'
    + '<tr><td style="padding:3px 0;color:#999;width:32%">Nombre</td><td style="padding:3px 0">' + (c.nombre || '') + ' ' + (c.apellido || '') + '</td></tr>'
    + '<tr><td style="padding:3px 0;color:#999">CUIT / CUIL</td><td style="padding:3px 0">' + (c.cuit || '') + '</td></tr>'
    + (c.mail      ? '<tr><td style="padding:3px 0;color:#999">Email</td><td style="padding:3px 0">' + c.mail + '</td></tr>' : '')
    + (c.tel       ? '<tr><td style="padding:3px 0;color:#999">Teléfono</td><td style="padding:3px 0">' + c.tel  + '</td></tr>' : '')
    + (c.localidad ? '<tr><td style="padding:3px 0;color:#999">Localidad</td><td style="padding:3px 0">' + c.localidad + '</td></tr>' : '')
    + condFiscalHTML + '</table>'
    + '<div style="font-size:10px;font-weight:700;letter-spacing:.15em;color:' + DARK + ';text-transform:uppercase;border-bottom:1px solid #e8e8e0;padding-bottom:4px;margin-bottom:0">Detalle de cajas</div>'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:20px">'
    + '<thead><tr style="border-bottom:1px solid #e8e8e0">'
    + '<th style="padding:8px 12px;text-align:left;font-weight:500;color:#999;font-size:11px">Descripción</th>'
    + '<th style="padding:8px 12px;text-align:left;font-weight:500;color:#999;font-size:11px">Contenido</th>'
    + '<th style="padding:8px 12px;text-align:right;font-weight:500;color:#999;font-size:11px">Importe</th>'
    + '</tr></thead><tbody>' + filasHTML + descRowHTML + '</tbody></table>'
    + '<div style="font-size:10px;font-weight:700;letter-spacing:.15em;color:' + DARK + ';text-transform:uppercase;border-bottom:1px solid #e8e8e0;padding-bottom:4px;margin-bottom:10px">Condiciones de cobro</div>'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px">'
    + '<tr><td style="padding:3px 0;color:#999;width:32%">Método</td><td style="padding:3px 0">' + (p.metodoCobro || '') + '</td></tr>'
    + '<tr><td style="padding:3px 0;color:#999">Moneda</td><td style="padding:3px 0">' + (p.moneda === 'USD' ? 'Dólares americanos' : 'Pesos argentinos') + '</td></tr>'
    + tcRowHTML + '</table>'
    + '<div style="border-top:2px solid ' + DARK + ';padding-top:14px;display:flex;justify-content:space-between;align-items:center">'
    + '<span style="font-size:12px;color:#999;text-transform:uppercase;letter-spacing:.08em">Total</span>'
    + '<span style="font-size:22px;font-weight:700;color:' + DARK + '">' + totalDisplay + '</span></div>'
    + unidadesHTML
    + '</div>'
    + '<div style="background:#f7f7f5;border-top:1px solid #e8e8e0;padding:12px 28px;display:flex;justify-content:space-between;align-items:center">'
    + '<span style="font-size:11px;color:#bbb">Documento válido como comprobante de compra</span>'
    + '<span style="font-size:11px;color:#bbb">AGF Messenchymal · 2026</span></div>'
    + '</div></body></html>';
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
  const payload = {
    ventaNum: 99,
    fecha: '19/4/2026, 10:00:00',
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

  const u = calcUnidades(payload);
  Logger.log('Unidades calculadas: ' + JSON.stringify(u));
  // Esperado: Dermal=11 (5+2+0), Capillary=6 (0+1+5), Pink=1, Biomask=1

  const { blob, url } = generarPdfRecibo(payload);
  Logger.log('PDF generado: ' + blob.getName() + ' (' + blob.getBytes().length + ' bytes)');
  Logger.log('URL en Drive: ' + url);

  guardarVenta(payload, url);
  actualizarResumen(payload);
  obtenerOCrearHoja(SHEET_STOCK, crearHojaStock);

  // Enviar emails con el PDF adjunto (igual que doPost)
  enviarEmailCliente(payload, blob);
  Logger.log('Email cliente enviado a: ' + payload.cliente.mail);
  enviarEmailAdmin(payload, blob);
  Logger.log('Email admin enviado a: ' + ADMIN_EMAIL);

  Logger.log('testManual completado OK');
}

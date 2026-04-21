// ═══════════════════════════════════════════════════════════════
//  AGF Messenchymal — Ventas Congreso · Google Apps Script
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1K2rELW54yEqQ8pKXHyPeyDqVLygeCQyQJMeH8w4n5m4';
const ADMIN_EMAIL    = 'alan.haslop@dermacells.com.ar';
const REPLY_TO       = 'alan.haslop@dermacells.com.ar';
const SHEET_VENTAS   = 'Ventas';
const SHEET_RESUMEN  = 'Resumen';
const SHEET_STOCK    = 'Stock';
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

    // Generar PDF primero para tener la URL antes de guardar en Sheets
    const { blob: pdfBlob, url: pdfUrl } = generarPdfRecibo(p);

    guardarVenta(p, pdfUrl);
    actualizarResumen(p);
    obtenerOCrearHoja(SHEET_STOCK, crearHojaStock); // crea la hoja si no existe

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
// Columnas: A–V datos venta | W PDF link | X–AA unidades por producto
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
    p.fecha          || '',                               // D  Fecha local
    c.nombre         || '',                               // E  Nombre
    c.apellido       || '',                               // F  Apellido
    c.cuit           || '',                               // G  CUIT/CUIL
    c.mail           || '',                               // H  Email
    c.tel            || '',                               // I  Teléfono
    c.localidad      || '',                               // J  Localidad
    p.condFiscal     || '',                               // K  Cond. Fiscal
    f.mismosContacto ? '' : (f.razonSocial     || ''),   // L  Razón Social
    f.mismosContacto ? '' : (f.cuitFacturacion || ''),   // M  CUIT Fact.
    p.metodoCobro    || '',                               // N  Método cobro
    p.moneda         || 'USD',                            // O  Moneda
    p.tipoCambio     || '',                               // P  Tipo cambio
    (p.cajas || []).length,                               // Q  Cant. cajas
    cajasDetalle,                                         // R  Detalle cajas
    p.descuentoGlobal || 0,                               // S  Desc. global %
    p.subtotalUSD    || p.totalUSD || 0,                  // T  Subtotal U$D
    p.totalUSD       || 0,                                // U  Total U$D
    p.totalARS       || '',                               // V  Total ARS
    '',                                                   // W  Recibo PDF (fórmula abajo)
    u.Dermal,                                             // X  Dermal (unidades)
    u.Capillary,                                          // Y  Capillary (unidades)
    u.Pink,                                               // Z  Pink (unidades)
    u.Biomask                                             // AA Biomask (unidades)
  ]);

  // Hipervínculo al PDF en columna W
  if (pdfUrl) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 23).setFormula('=HYPERLINK("' + pdfUrl + '","📄 Ver recibo")');
  }
}

function crearHeadersVentas(sheet) {
  const headers = [
    'Timestamp','Venta #','Dispositivo','Fecha local',
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
  sheet.setColumnWidth(4,  140);
  sheet.setColumnWidth(18, 280);
  sheet.setColumnWidths(5, 2, 110);
  sheet.setColumnWidth(23, 100); // PDF
  // Destacar columnas de stock en celeste
  sheet.getRange(1, 24, 1, 4).setBackground('#1B3A52');
}


// ── HOJA DE RESUMEN (por día + método) ──────────────────────────
function actualizarResumen(p) {
  const sheet = obtenerOCrearHoja(SHEET_RESUMEN, crearHeadersResumen);
  const hoy   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const datos = sheet.getDataRange().getValues();
  const u     = calcUnidades(p);

  const COL = { fecha:0, metodo:1, cajas:2, usd:3, ars:4, ventas:5,
                dermal:6, capillary:7, pink:8, biomask:9 };

  let filaExistente = -1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][COL.fecha] === hoy && datos[i][COL.metodo] === p.metodoCobro) {
      filaExistente = i + 1;
      break;
    }
  }

  const cantCajas = (p.cajas || []).length;
  const totalUSD  = p.totalUSD  || 0;
  const totalARS  = p.totalARS  || 0;

  if (filaExistente > 0) {
    const fila = sheet.getRange(filaExistente, 1, 1, 10).getValues()[0];
    sheet.getRange(filaExistente, COL.cajas    + 1).setValue(fila[COL.cajas]    + cantCajas);
    sheet.getRange(filaExistente, COL.usd      + 1).setValue(fila[COL.usd]      + totalUSD);
    sheet.getRange(filaExistente, COL.ars      + 1).setValue(fila[COL.ars]      + totalARS);
    sheet.getRange(filaExistente, COL.ventas   + 1).setValue(fila[COL.ventas]   + 1);
    sheet.getRange(filaExistente, COL.dermal   + 1).setValue(fila[COL.dermal]   + u.Dermal);
    sheet.getRange(filaExistente, COL.capillary+ 1).setValue(fila[COL.capillary]+ u.Capillary);
    sheet.getRange(filaExistente, COL.pink     + 1).setValue(fila[COL.pink]     + u.Pink);
    sheet.getRange(filaExistente, COL.biomask  + 1).setValue(fila[COL.biomask]  + u.Biomask);
  } else {
    sheet.appendRow([hoy, p.metodoCobro, cantCajas, totalUSD, totalARS, 1,
                     u.Dermal, u.Capillary, u.Pink, u.Biomask]);
  }
}

function crearHeadersResumen(sheet) {
  sheet.appendRow(['Fecha','Método cobro','Cajas vendidas','Total U$D','Total ARS','# Ventas',
                   'Dermal (u)','Capillary (u)','Pink (u)','Biomask (u)']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 10)
    .setBackground('#1B3A52').setFontColor('#ffffff').setFontWeight('bold');
}


// ── HOJA DE STOCK ────────────────────────────────────────────────
function crearHojaStock(sheet) {
  const PRODS = ['Dermal', 'Capillary', 'Pink', 'Biomask'];
  // Columnas en hoja Ventas para cada producto
  const COL_VENTAS = { Dermal: 'X', Capillary: 'Y', Pink: 'Z', Biomask: 'AA' };

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
  // Resumen: A=Fecha, B=Método, C=Cajas, D=U$D, E=ARS, F=#Ventas, G=Dermal, H=Capillary, I=Pink, J=Biomask
  sheet.getRange('A11').setFormula(
    '=IFERROR(QUERY(Resumen!A2:J,'
    + '"SELECT A, SUM(G), SUM(H), SUM(I), SUM(J), SUM(C), SUM(D), SUM(E), SUM(F) '
    + 'WHERE A IS NOT NULL '
    + 'GROUP BY A '
    + 'ORDER BY A DESC '
    + 'LABEL A \'\', SUM(G) \'\', SUM(H) \'\', SUM(I) \'\', SUM(J) \'\', SUM(C) \'\', SUM(D) \'\', SUM(E) \'\', SUM(F) \'\'"'
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
 * HTML del recibo — mismo diseño que imprimirRecibo() en index.html, sin firmas.
 * NOTA: Usa layout 100% basado en tablas para compatibilidad con conversión Google Docs → PDF.
 * CSS flexbox, gap, display:inline-block son ignorados por Google Docs.
 */
function buildReciboHTML(p) {
  const c     = p.cliente     || {};
  const f     = p.facturacion || {};
  const GREEN = '#9FD4C0';
  const DARK  = '#1B3A52';

  // ── Filas de cajas con tachado si hay descuento por caja ──
  const filas = (p.cajas || []).map(function(l) {
    const tipo  = l.tipo === 'cerrada' ? 'Cerrada' : 'Combinada';
    const tagBg = l.tipo === 'cerrada' ? '#EEF4F8' : '#f0faf6';
    const baseP = l.tipo === 'cerrada' ? P_CERRADA  : P_COMBINADA;
    const tachado = l.descCaja > 0
      ? '<span style="text-decoration:line-through;color:#ccc;font-size:9px;margin-right:4px">u$' + baseP + '</span>'
      : '';
    return '<tr style="border-bottom:1px solid #f5f2ec">'
      + '<td style="padding:6px 8px;width:80px">'
      +   '<table cellpadding="0" cellspacing="0"><tr>'
      +     '<td bgcolor="' + tagBg + '" style="padding:2px 7px;font-size:9px;font-weight:500;color:' + DARK + '">' + tipo + '</td>'
      +   '</tr></table>'
      + '</td>'
      + '<td style="padding:6px 8px;color:#777;font-size:10.5px">' + l.detalle + '</td>'
      + '<td style="padding:6px 8px;text-align:right;font-weight:600;color:' + DARK + ';font-size:11px;white-space:nowrap">' + tachado + 'u$' + l.precio + '</td>'
      + '</tr>';
  }).join('');

  // ── Filas de subtotal + descuento global ──
  const subtotalUSD = (p.cajas || []).reduce(function(acc, l) {
    return acc + Math.round((l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA) * (1 - (l.descCaja || 0) / 100));
  }, 0);
  const ahorroUSD = (p.descuentoGlobal > 0 && p.descuentoGlobal < 100)
    ? subtotalUSD - p.totalUSD : 0;

  const descGlobalRows = p.descuentoGlobal > 0
    ? '<tr><td colspan="2" style="padding:4px 8px;font-size:10px;color:#999">Subtotal</td>'
      + '<td style="padding:4px 8px;text-align:right;font-size:10px;color:#999">u$' + subtotalUSD + '</td></tr>'
      + '<tr><td colspan="2" style="padding:4px 8px;font-size:10px;color:#e67e22">Desc. general ' + p.descuentoGlobal + '%</td>'
      + '<td style="padding:4px 8px;text-align:right;font-size:10px;color:#e67e22;font-weight:600">'
      + (ahorroUSD > 0 ? '&#8722; u$' + ahorroUSD : '') + '</td></tr>'
    : '';

  // ── Detalles de pago (filas de tabla en lugar de divs flex) ──
  const metodoRow = p.metodoCobro
    ? '<tr><td style="padding:2px 0;color:#bbb;font-size:10px;width:130px">Método de cobro</td>'
      + '<td style="padding:2px 0;color:#1a1a1a;font-weight:500;font-size:11px">' + p.metodoCobro + '</td></tr>'
    : '';
  const tcRow = p.tipoCambio
    ? '<tr><td style="padding:2px 0;color:#bbb;font-size:10px">Tipo de cambio</td>'
      + '<td style="padding:2px 0;color:#1a1a1a;font-weight:500;font-size:11px">AR$' + Number(p.tipoCambio).toLocaleString('es-AR') + ' / U$D</td></tr>'
    : '';
  const condRow = p.condFiscal
    ? '<tr><td style="padding:2px 0;color:#bbb;font-size:10px">Cond. fiscal</td>'
      + '<td style="padding:2px 0;color:#1a1a1a;font-weight:500;font-size:11px">' + p.condFiscal + (!f.mismosContacto && f.razonSocial ? ' &#8212; ' + f.razonSocial : '') + '</td></tr>'
    : '';
  const totalARSStr = p.totalARS
    ? '<div style="font-size:11px;color:#aaa;margin-top:3px">AR$ ' + Number(p.totalARS).toLocaleString('es-AR') + '</div>'
    : '';

  // ── Bloque de facturación ──
  const instrMap = {
    'Resp. Inscripto': { bg:'#eff6ff', border:'#93c5fd', color:'#1e3a8a', icono:'&#9650;',
      texto:'Se emitirá Factura A. Verificar CUIT y razón social con administración antes de procesar.' },
    'Monotributista':  { bg:'#f0fdf4', border:'#86efac', color:'#14532d', icono:'&#9679;',
      texto:'Se emitirá Factura B. Datos registrados para procesamiento posterior por administración.' },
    'Cons. Final':     { bg:'#f0fdf4', border:'#86efac', color:'#14532d', icono:'&#9679;',
      texto:'Se emitirá Factura B (Consumidor Final). Sin CUIT específico requerido.' }
  };
  const instrData  = p.condFiscal && instrMap[p.condFiscal] ? instrMap[p.condFiscal] : null;
  const facRazon   = (!f.mismosContacto && f.razonSocial)     ? f.razonSocial     : '';
  const facCuit    = (!f.mismosContacto && f.cuitFacturacion) ? f.cuitFacturacion : '';

  const datosFactRow = (facRazon || facCuit)
    ? '<tr><td colspan="2" style="padding-top:7px;border-top:1px solid rgba(0,0,0,0.1)">'
      + '<table cellpadding="0" cellspacing="0"><tr>'
      + (facRazon
        ? '<td style="padding-right:20px"><div style="font-size:8px;opacity:.7">Razón social</div>'
          + '<div style="font-size:10.5px;font-weight:600">' + facRazon + '</div></td>'
        : '')
      + (facCuit
        ? '<td><div style="font-size:8px;opacity:.7">CUIT</div>'
          + '<div style="font-size:10.5px;font-weight:600">' + facCuit + '</div></td>'
        : '')
      + '</tr></table></td></tr>'
    : '';

  const facturacionBlock = instrData
    ? '<table cellpadding="0" cellspacing="0" width="100%" style="margin-top:16px;border:1px solid ' + instrData.border + '">'
      + '<tr><td bgcolor="' + instrData.bg + '" style="padding:10px 14px">'
      + '<table cellpadding="0" cellspacing="0" width="100%">'
      + '<tr><td style="font-size:8px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;color:' + instrData.color + ';padding-bottom:4px">'
      + instrData.icono + ' Instrucciones de facturación</td></tr>'
      + '<tr><td style="font-size:10.5px;color:' + instrData.color + ';line-height:1.5">' + instrData.texto + '</td></tr>'
      + datosFactRow
      + '</table>'
      + '</td></tr></table>'
    : '';

  // ── Pre-computados para evitar "? / :" al inicio de línea con "+" (rompe parser Apps Script) ──
  const localidadCell = c.localidad
    ? '<td><div style="font-size:8px;color:#bbb;margin-bottom:2px">Localidad</div><div style="font-size:12px">' + c.localidad + '</div></td>'
    : '';
  const emailTelTable = (c.mail || c.tel)
    ? '<table cellpadding="0" cellspacing="0" style="margin-bottom:14px"><tr>'
      + (c.mail ? '<td style="padding-right:24px"><div style="font-size:8px;color:#bbb;margin-bottom:2px">Email</div><div style="font-size:11px">' + c.mail + '</div></td>' : '')
      + (c.tel  ? '<td><div style="font-size:8px;color:#bbb;margin-bottom:2px">Teléfono</div><div style="font-size:11px">' + c.tel  + '</div></td>' : '')
      + '</tr></table>'
    : '';

  // ════ CONSTRUCCIÓN DEL HTML ════
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">'
    + '<style>body{background:#FAF7F2;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#1a1a1a;padding:24px}table{border-collapse:collapse}td,th{vertical-align:top}</style>'
    + '</head><body>'

    // ── Contenedor principal ──
    + '<table align="center" cellpadding="0" cellspacing="0" width="680" bgcolor="#ffffff">'

    // ── HEADER ──
    + '<tr><td style="padding:20px 36px 16px;border-bottom:1px solid #e0ddd6">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td valign="middle">'
    +       '<div style="font-size:22px;font-weight:700;color:' + DARK + ';letter-spacing:-.5px">AGF Messenchymal</div>'
    +       '<div style="font-size:10px;color:#aaa;margin-top:3px;font-style:italic">dermacells.com.ar &middot; Argentina</div>'
    +     '</td>'
    +     '<td valign="middle" align="right">'
    +       '<div style="font-size:28px;font-weight:300;color:' + DARK + ';line-height:1">Venta #' + p.ventaNum + '</div>'
    +       '<div style="font-size:10px;color:#aaa;margin-top:4px">' + (p.fecha || '') + '</div>'
    +     '</td>'
    +   '</tr></table>'
    + '</td></tr>'

    // ── BANDA OSCURA ──
    + '<tr><td bgcolor="' + DARK + '" style="padding:7px 36px">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td style="font-size:9px;font-weight:500;letter-spacing:.2em;text-transform:uppercase;color:' + GREEN + '">Comprobante de compra</td>'
    +     '<td align="right" style="font-size:9px;font-weight:500;letter-spacing:.2em;text-transform:uppercase;color:' + GREEN + '">Dermacells S.A.</td>'
    +   '</tr></table>'
    + '</td></tr>'

    // ── CUERPO ──
    + '<tr><td style="padding:18px 36px 24px">'

    //   CLIENTE — título sección
    +   '<div style="font-size:8px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:8px;padding-bottom:4px;border-bottom:1px solid #e8e8e0">Datos del cliente</div>'

    //   CLIENTE — fila nombre / cuit / localidad
    +   '<table cellpadding="0" cellspacing="0" style="margin-bottom:14px"><tr>'
    +     '<td style="padding-right:24px;min-width:150px">'
    +       '<div style="font-size:8px;color:#bbb;margin-bottom:2px">Nombre</div>'
    +       '<div style="font-size:12px;font-weight:600">' + (c.nombre || '') + ' ' + (c.apellido || '') + '</div>'
    +     '</td>'
    +     '<td style="padding-right:24px;min-width:130px">'
    +       '<div style="font-size:8px;color:#bbb;margin-bottom:2px">CUIT / CUIL</div>'
    +       '<div style="font-size:12px;font-weight:600">' + (c.cuit || '') + '</div>'
    +     '</td>'
    +     localidadCell
    +   '</tr></table>'

    //   CLIENTE — fila email / tel
    +   emailTelTable

    //   DETALLE — título sección
    +   '<div style="font-size:8px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:6px;padding-bottom:4px;border-bottom:1px solid #e8e8e0">Detalle</div>'

    //   DETALLE — tabla de cajas
    +   '<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:16px">'
    +   '<thead><tr style="border-bottom:2px solid ' + DARK + '">'
    +   '<th style="padding:5px 8px;text-align:left;font-size:9px;font-weight:600;color:#bbb;letter-spacing:.1em;text-transform:uppercase;width:80px">Tipo</th>'
    +   '<th style="padding:5px 8px;text-align:left;font-size:9px;font-weight:600;color:#bbb;letter-spacing:.1em;text-transform:uppercase">Contenido</th>'
    +   '<th style="padding:5px 8px;text-align:right;font-size:9px;font-weight:600;color:#bbb;letter-spacing:.1em;text-transform:uppercase;white-space:nowrap">Precio</th>'
    +   '</tr></thead><tbody>' + filas + descGlobalRows + '</tbody></table>'

    //   TOTAL + PAGO — divididos izquierda/derecha con tabla
    +   '<table cellpadding="0" cellspacing="0" width="100%" style="border-top:1px solid #e8e8e0;padding-top:14px"><tr>'

    //     Izquierda: detalles de pago
    +     '<td valign="top">'
    +       '<div style="font-size:8px;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:6px">Detalles de pago</div>'
    +       '<table cellpadding="0" cellspacing="0">' + metodoRow + tcRow + condRow + '</table>'
    +     '</td>'

    //     Derecha: total
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
    + '<tr bgcolor="#f7f7f5"><td style="padding:10px 36px;border-top:1px solid #e8e8e0">'
    +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
    +     '<td style="font-size:9px;color:#bbb">Documento válido como comprobante de compra</td>'
    +     '<td align="right" style="font-size:9px;color:#bbb">AGF Messenchymal &middot; BAAS 2026</td>'
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

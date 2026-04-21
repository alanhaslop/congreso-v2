// ═══════════════════════════════════════════════════════════════
//  AGF Messenchymal — Ventas Congreso · Google Apps Script
//  Pegá este código en script.google.com y configurá las dos
//  constantes de abajo antes de hacer "Implementar > Web app".
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID  = '1K2rELW54yEqQ8pKXHyPeyDqVLygeCQyQJMeH8w4n5m4';
const ADMIN_EMAIL     = 'alan.haslop@dermacells.com.ar';
const REPLY_TO        = 'alan.haslop@dermacells.com.ar';
const SHEET_VENTAS    = 'Ventas';
const SHEET_RESUMEN   = 'Resumen';
const PDF_FOLDER_ID   = '1qGxephO8pfFAz41B28f39BZM40YdM2GS'; // Carpeta Drive donde se guardan los recibos

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

    // Generar PDF, guardarlo en Drive y obtener URL + blob para email
    const { blob: pdfBlob, url: pdfUrl } = generarPdfRecibo(p);

    guardarVenta(p, pdfUrl);    // pasa la URL para el hipervínculo en Sheets
    actualizarResumen(p);

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


// ── GUARDAR VENTA EN HOJA ────────────────────────────────────────
function guardarVenta(p, pdfUrl) {
  const sheet = obtenerOCrearHoja(SHEET_VENTAS, crearHeadersVentas);
  const c = p.cliente     || {};
  const f = p.facturacion || {};

  const cajasDetalle = (p.cajas || []).map(cj => {
    let txt = `Caja ${cj.caja} (${cj.tipo}): ${cj.detalle}`;
    if (cj.descCaja > 0) txt += ` [-${cj.descCaja}%]`;
    txt += ` = u$${cj.precio}`;
    return txt;
  }).join('\n');

  sheet.appendRow([
    new Date(),
    p.ventaNum,
    p.dispositivo    || '',
    p.fecha          || '',
    c.nombre         || '',
    c.apellido       || '',
    c.cuit           || '',
    c.mail           || '',
    c.tel            || '',
    c.localidad      || '',
    p.condFiscal     || '',
    f.mismosContacto ? '' : (f.razonSocial     || ''),
    f.mismosContacto ? '' : (f.cuitFacturacion || ''),
    p.metodoCobro    || '',
    p.moneda         || 'USD',
    p.tipoCambio     || '',
    (p.cajas || []).length,
    cajasDetalle,
    p.descuentoGlobal || 0,
    p.subtotalUSD    || p.totalUSD || 0,
    p.totalUSD       || 0,
    p.totalARS       || '',
    ''  // col W: placeholder para el hipervínculo al PDF
  ]);

  // Agregar hipervínculo al PDF en la última fila, columna W (23)
  if (pdfUrl) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 23).setFormula(`=HYPERLINK("${pdfUrl}","📄 Ver recibo")`);
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
    'Recibo PDF'  // col W
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setBackground('#1D9E75').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1,  160);
  sheet.setColumnWidth(4,  140);
  sheet.setColumnWidth(18, 280);
  sheet.setColumnWidths(5, 2, 110);
}


// ── HOJA DE RESUMEN ──────────────────────────────────────────────
function actualizarResumen(p) {
  const sheet = obtenerOCrearHoja(SHEET_RESUMEN, crearHeadersResumen);
  const hoy   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const datos = sheet.getDataRange().getValues();

  const colFecha = 0, colMetodo = 1, colCajas = 2, colUSD = 3, colARS = 4, colVentas = 5;
  let filaExistente = -1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colFecha] === hoy && datos[i][colMetodo] === p.metodoCobro) {
      filaExistente = i + 1;
      break;
    }
  }

  const cantCajas = (p.cajas || []).length;
  const totalUSD  = p.totalUSD || 0;
  const totalARS  = p.totalARS || 0;

  if (filaExistente > 0) {
    const fila = sheet.getRange(filaExistente, 1, 1, 6).getValues()[0];
    sheet.getRange(filaExistente, colCajas  + 1).setValue(fila[colCajas]  + cantCajas);
    sheet.getRange(filaExistente, colUSD    + 1).setValue(fila[colUSD]    + totalUSD);
    sheet.getRange(filaExistente, colARS    + 1).setValue(fila[colARS]    + totalARS);
    sheet.getRange(filaExistente, colVentas + 1).setValue(fila[colVentas] + 1);
  } else {
    sheet.appendRow([hoy, p.metodoCobro, cantCajas, totalUSD, totalARS, 1]);
  }
}

function crearHeadersResumen(sheet) {
  sheet.appendRow(['Fecha','Método cobro','Cajas vendidas','Total U$D','Total ARS','# Ventas']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 6).setBackground('#1B3A52').setFontColor('#ffffff').setFontWeight('bold');
}


// ── PDF RECIBO ───────────────────────────────────────────────────
/**
 * Genera el PDF del recibo, lo guarda en PDF_FOLDER_ID y devuelve
 * { blob, url } donde blob se usa para adjuntar al email y url para
 * el hipervínculo en la hoja de Sheets.
 */
function generarPdfRecibo(p) {
  const html     = buildReciboHTML(p);
  const nombre   = 'Recibo_AGF_Venta' + p.ventaNum + '.pdf';

  // 1. Crear HTML temporal para conversión
  const htmlBlob  = Utilities.newBlob(html, 'text/html', 'recibo_temp.html');
  const tempFile  = DriveApp.createFile(htmlBlob);

  try {
    // 2. Convertir a PDF
    const pdfBlob = tempFile.getAs('application/pdf');
    pdfBlob.setName(nombre);

    // 3. Guardar PDF permanentemente en la carpeta indicada
    const folder    = DriveApp.getFolderById(PDF_FOLDER_ID);
    const savedFile = folder.createFile(pdfBlob);
    savedFile.setName(nombre);

    // 4. Obtener blob fresco del archivo guardado para adjuntar al email
    const emailBlob = savedFile.getBlob();
    emailBlob.setName(nombre);

    return { blob: emailBlob, url: savedFile.getUrl() };

  } finally {
    tempFile.setTrashed(true); // borrar solo el HTML temporal
  }
}

/**
 * HTML del recibo — mismo diseño que imprimirRecibo() en index.html,
 * sin la sección de firmas, sin imágenes de logos (texto branding).
 */
function buildReciboHTML(p) {
  const c     = p.cliente     || {};
  const f     = p.facturacion || {};
  const GREEN = '#9FD4C0';
  const DARK  = '#1B3A52';

  // ── Filas de cajas con tachado si hay descuento por caja ──
  const filas = (p.cajas || []).map(function(l) {
    const tipo     = l.tipo === 'cerrada' ? 'Cerrada' : 'Combinada';
    const tagBg    = l.tipo === 'cerrada' ? '#EEF4F8' : '#f0faf6';
    const baseP    = l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA;
    const tachado  = l.descCaja > 0
      ? '<span style="text-decoration:line-through;color:#ccc;font-size:9px;margin-right:4px">u$' + baseP + '</span>'
      : '';
    return '<tr style="border-bottom:1px solid #f5f2ec">'
      + '<td style="padding:6px 8px"><span style="display:inline-block;background:' + tagBg + ';color:' + DARK + ';font-size:9px;padding:2px 7px;border-radius:2px;font-weight:500">' + tipo + '</span></td>'
      + '<td style="padding:6px 8px;color:#777;font-size:10.5px">' + l.detalle + '</td>'
      + '<td style="padding:6px 8px;text-align:right;font-weight:600;color:' + DARK + ';font-size:11px">' + tachado + 'u$' + l.precio + '</td>'
      + '</tr>';
  }).join('');

  // ── Filas de subtotal + descuento global ──
  const subtotalUSD = (p.cajas || []).reduce(function(acc, l) {
    const base  = l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA;
    const after = Math.round(base * (1 - (l.descCaja || 0) / 100));
    return acc + after;
  }, 0);
  const ahorroUSD = (p.descuentoGlobal > 0 && p.descuentoGlobal < 100)
    ? subtotalUSD - p.totalUSD : 0;

  const descGlobalRows = p.descuentoGlobal > 0
    ? '<tr><td colspan="2" style="padding:4px 8px;font-size:10px;color:#999">Subtotal</td>'
      + '<td style="padding:4px 8px;text-align:right;font-size:10px;color:#999">u$' + subtotalUSD + '</td></tr>'
      + '<tr><td colspan="2" style="padding:4px 8px;font-size:10px;color:#e67e22">Desc. general ' + p.descuentoGlobal + '%</td>'
      + '<td style="padding:4px 8px;text-align:right;font-size:10px;color:#e67e22;font-weight:600">'
      + (ahorroUSD > 0 ? '− u$' + ahorroUSD : '') + '</td></tr>'
    : '';

  // ── Método de pago / TC ──
  const metodoRow = p.metodoCobro
    ? '<div style="display:flex;gap:12px;padding:2px 0">'
      + '<span style="color:#bbb;min-width:120px;font-size:10px">Método de cobro</span>'
      + '<span style="color:#1a1a1a;font-weight:500;font-size:11px">' + p.metodoCobro + '</span></div>'
    : '';
  const tcRow = p.tipoCambio
    ? '<div style="display:flex;gap:12px;padding:2px 0">'
      + '<span style="color:#bbb;min-width:120px;font-size:10px">Tipo de cambio</span>'
      + '<span style="color:#1a1a1a;font-weight:500;font-size:11px">AR$' + Number(p.tipoCambio).toLocaleString('es-AR') + ' / U$D</span></div>'
    : '';

  // ── Condición fiscal (solo si hay) ──
  const condRow = p.condFiscal
    ? '<div style="display:flex;gap:12px;padding:2px 0">'
      + '<span style="color:#bbb;min-width:120px;font-size:10px">Cond. fiscal</span>'
      + '<span style="color:#1a1a1a;font-weight:500;font-size:11px">' + p.condFiscal
      + (!f.mismosContacto && f.razonSocial ? ' — ' + f.razonSocial : '')
      + '</span></div>'
    : '';

  // ── Total ARS ──
  const totalARSStr = p.totalARS
    ? '<div style="font-size:11px;color:#aaa;margin-top:3px">AR$ ' + Number(p.totalARS).toLocaleString('es-AR') + '</div>'
    : '';

  // ── Bloque de facturación (igual que en index.html PATCH A) ──
  const instrMap = {
    'Resp. Inscripto': { bg:'#eff6ff', border:'#93c5fd', color:'#1e3a8a', icono:'▲',
      texto:'Se emitirá Factura A. Verificar CUIT y razón social con administración antes de procesar.' },
    'Monotributista':  { bg:'#f0fdf4', border:'#86efac', color:'#14532d', icono:'●',
      texto:'Se emitirá Factura B. Datos registrados para procesamiento posterior por administración.' },
    'Cons. Final':     { bg:'#f0fdf4', border:'#86efac', color:'#14532d', icono:'●',
      texto:'Se emitirá Factura B (Consumidor Final). Sin CUIT específico requerido.' }
  };
  const instrData = p.condFiscal && instrMap[p.condFiscal] ? instrMap[p.condFiscal] : null;

  // Solo mostrar razón social / CUIT de facturación si son distintos al cliente
  const facRazon = (!f.mismosContacto && f.razonSocial)     ? f.razonSocial     : '';
  const facCuit  = (!f.mismosContacto && f.cuitFacturacion) ? f.cuitFacturacion : '';
  const datosFactRow = (facRazon || facCuit)
    ? '<div style="margin-top:7px;padding-top:7px;border-top:1px solid rgba(0,0,0,0.1);display:flex;gap:20px">'
      + (facRazon ? '<div><div style="font-size:8px;opacity:.7">Razón social</div><div style="font-size:10.5px;font-weight:600">' + facRazon + '</div></div>' : '')
      + (facCuit  ? '<div><div style="font-size:8px;opacity:.7">CUIT</div><div style="font-size:10.5px;font-weight:600">' + facCuit + '</div></div>' : '')
      + '</div>'
    : '';
  const facturacionBlock = instrData
    ? '<div style="margin-top:16px;padding:10px 14px;background:' + instrData.bg + ';border:1px solid ' + instrData.border + ';border-radius:6px">'
      + '<div style="font-size:8px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;color:' + instrData.color + ';margin-bottom:4px">' + instrData.icono + ' Instrucciones de facturación</div>'
      + '<div style="font-size:10.5px;color:' + instrData.color + ';line-height:1.5">' + instrData.texto + '</div>'
      + datosFactRow
      + '</div>'
    : '';

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">'
    + '<style>'
    + '* { box-sizing: border-box; margin: 0; padding: 0; }'
    + 'body { background: #FAF7F2; font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #1a1a1a; padding: 24px; }'
    + 'table { border-collapse: collapse; width: 100%; }'
    + '</style></head><body>'
    + '<div style="max-width:720px;margin:0 auto;background:#fff">'

    // ── HEADER ──
    + '<div style="padding:20px 36px 16px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #e0ddd6">'
    + '<div>'
    + '<div style="font-size:22px;font-weight:700;color:' + DARK + ';letter-spacing:-.5px">AGF Messenchymal</div>'
    + '<div style="font-size:10px;color:#aaa;margin-top:3px;font-style:italic">dermacells.com.ar · Argentina</div>'
    + '</div>'
    + '<div style="text-align:right">'
    + '<div style="font-size:28px;font-weight:300;color:' + DARK + ';line-height:1">Venta #' + p.ventaNum + '</div>'
    + '<div style="font-size:10px;color:#aaa;margin-top:4px">' + (p.fecha || '') + '</div>'
    + '</div>'
    + '</div>'

    // ── BANDA ──
    + '<div style="background:' + DARK + ';padding:7px 36px;display:flex;justify-content:space-between">'
    + '<span style="font-size:9px;font-weight:500;letter-spacing:.2em;text-transform:uppercase;color:' + GREEN + '">Comprobante de compra</span>'
    + '<span style="font-size:9px;font-weight:500;letter-spacing:.2em;text-transform:uppercase;color:' + GREEN + '">Dermacells S.A.</span>'
    + '</div>'

    + '<div style="padding:18px 36px 24px">'

    // ── DATOS DEL CLIENTE ──
    + '<div style="font-size:8px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:8px;padding-bottom:4px;border-bottom:1px solid #e8e8e0">Datos del cliente</div>'
    + '<div style="display:flex;gap:20px;margin-bottom:14px;flex-wrap:wrap">'
    + '<div style="min-width:140px"><div style="font-size:8px;color:#bbb;margin-bottom:2px">Nombre</div><div style="font-size:12px;font-weight:600">' + (c.nombre || '') + ' ' + (c.apellido || '') + '</div></div>'
    + '<div style="min-width:120px"><div style="font-size:8px;color:#bbb;margin-bottom:2px">CUIT / CUIL</div><div style="font-size:12px;font-weight:600">' + (c.cuit || '') + '</div></div>'
    + (c.localidad ? '<div><div style="font-size:8px;color:#bbb;margin-bottom:2px">Localidad</div><div style="font-size:12px">' + c.localidad + '</div></div>' : '')
    + '</div>'
    + ((c.mail || c.tel)
      ? '<div style="display:flex;gap:20px;margin-bottom:14px;flex-wrap:wrap">'
        + (c.mail ? '<div><div style="font-size:8px;color:#bbb;margin-bottom:2px">Email</div><div style="font-size:11px">' + c.mail + '</div></div>' : '')
        + (c.tel  ? '<div><div style="font-size:8px;color:#bbb;margin-bottom:2px">Teléfono</div><div style="font-size:11px">' + c.tel + '</div></div>' : '')
        + '</div>'
      : '')

    // ── TABLA DE CAJAS ──
    + '<div style="font-size:8px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:6px;padding-bottom:4px;border-bottom:1px solid #e8e8e0">Detalle</div>'
    + '<table style="margin-bottom:16px">'
    + '<thead><tr style="border-bottom:2px solid ' + DARK + '">'
    + '<th style="padding:5px 8px;text-align:left;font-size:9px;font-weight:600;color:#bbb;letter-spacing:.1em;text-transform:uppercase">Tipo</th>'
    + '<th style="padding:5px 8px;text-align:left;font-size:9px;font-weight:600;color:#bbb;letter-spacing:.1em;text-transform:uppercase">Contenido</th>'
    + '<th style="padding:5px 8px;text-align:right;font-size:9px;font-weight:600;color:#bbb;letter-spacing:.1em;text-transform:uppercase">Precio</th>'
    + '</tr></thead>'
    + '<tbody>' + filas + descGlobalRows + '</tbody>'
    + '</table>'

    // ── TOTAL + DETALLES ──
    + '<div style="display:flex;justify-content:space-between;align-items:flex-start;border-top:1px solid #e8e8e0;padding-top:14px">'
    + '<div>'
    + '<div style="font-size:8px;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:6px">Detalles de pago</div>'
    + metodoRow + tcRow + condRow
    + '</div>'
    + '<div style="text-align:right">'
    + '<div style="font-size:8px;letter-spacing:.2em;text-transform:uppercase;color:#bbb;margin-bottom:3px">Total</div>'
    + '<div style="font-size:36px;font-weight:600;color:' + DARK + ';line-height:1">u$' + p.totalUSD + '</div>'
    + totalARSStr
    + '</div>'
    + '</div>'

    + facturacionBlock

    + '</div>' // /padding

    // ── PIE ──
    + '<div style="background:#f7f7f5;border-top:1px solid #e8e8e0;padding:10px 36px;display:flex;justify-content:space-between">'
    + '<span style="font-size:9px;color:#bbb">Documento válido como comprobante de compra</span>'
    + '<span style="font-size:9px;color:#bbb">AGF Messenchymal · BAAS 2026</span>'
    + '</div>'

    + '</div></body></html>';
}


// ── EMAIL CLIENTE ────────────────────────────────────────────────
function enviarEmailCliente(p, pdfBlob) {
  const c = p.cliente;
  GmailApp.sendEmail(
    c.mail,
    'AGF Messenchymal — Comprobante de compra #' + p.ventaNum,
    buildTextoPlano(p),
    {
      htmlBody:    buildEmailHTML(p, false),
      name:        'AGF Messenchymal Argentina',
      replyTo:     REPLY_TO,
      attachments: [pdfBlob]
    }
  );
}

function enviarEmailAdmin(p, pdfBlob) {
  const c = p.cliente || {};
  GmailApp.sendEmail(
    ADMIN_EMAIL,
    '[Venta #' + p.ventaNum + '] ' + c.nombre + ' ' + c.apellido + ' — u$' + p.totalUSD + ' — ' + p.metodoCobro,
    buildTextoPlano(p),
    {
      htmlBody:    buildEmailHTML(p, true),
      name:        'AGF Ventas Congreso',
      attachments: [pdfBlob]
    }
  );
}


// ── TEXTO PLANO (fallback) ───────────────────────────────────────
function buildTextoPlano(p) {
  const c = p.cliente || {};
  const lineas = [
    'AGF Messenchymal — Venta #' + p.ventaNum + '  |  Dispositivo ' + (p.dispositivo || '?'),
    'Fecha: ' + p.fecha,
    'Cliente: ' + c.nombre + ' ' + c.apellido + '  |  CUIT: ' + c.cuit,
    '',
    'Detalle:',
    ...(p.cajas || []).map(cj =>
      '  Caja ' + cj.caja + ' (' + cj.tipo + '): ' + cj.detalle
      + (cj.descCaja > 0 ? ' [-' + cj.descCaja + '%]' : '') + ' = u$' + cj.precio
    ),
    '',
    p.descuentoGlobal > 0
      ? '  Descuento global: ' + p.descuentoGlobal + '%  (-u$' + (subtotalCalc(p) - p.totalUSD) + ')' : null,
    'Total: u$' + p.totalUSD + (p.totalARS ? '  /  AR$' + Number(p.totalARS).toLocaleString('es-AR') : ''),
    'Método: ' + p.metodoCobro + '  |  Moneda: ' + p.moneda + (p.tipoCambio ? '  |  TC: $' + p.tipoCambio : ''),
    p.condFiscal ? 'Cond. fiscal: ' + p.condFiscal : null
  ].filter(function(l) { return l !== null; }).join('\n');
  return lineas;
}

function subtotalCalc(p) {
  return (p.cajas || []).reduce(function(acc, l) {
    const base = l.tipo === 'cerrada' ? P_CERRADA : P_COMBINADA;
    return acc + Math.round(base * (1 - (l.descCaja || 0) / 100));
  }, 0);
}


// ── EMAIL HTML (cuerpo del mensaje, sin adjunto) ─────────────────
function buildEmailHTML(p, isAdmin) {
  const c     = p.cliente     || {};
  const f     = p.facturacion || {};
  const GREEN = '#1D9E75';
  const DARK  = '#1B3A52';

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

  const descRowHTML = p.descuentoGlobal > 0
    ? '<tr><td colspan="2" style="padding:5px 12px;font-size:11px;color:#999;font-style:italic">Descuento global: ' + p.descuentoGlobal + '%</td>'
      + '<td style="padding:5px 12px;text-align:right;color:#c0392b;font-size:12px">−u$' + (subtotalCalc(p) - p.totalUSD) + '</td></tr>'
    : '';

  const totalDisplay = 'u$' + p.totalUSD
    + (p.totalARS ? '&nbsp;<span style="font-size:14px;color:#666">/ AR$' + Number(p.totalARS).toLocaleString('es-AR') + '</span>' : '');

  const tcRowHTML = p.tipoCambio
    ? '<tr><td style="padding:3px 0;color:#999">Tipo de cambio</td><td style="padding:3px 0">$' + Number(p.tipoCambio).toLocaleString('es-AR') + ' AR$/U$D</td></tr>'
    : '';

  const condFiscalHTML = p.condFiscal
    ? (f.mismosContacto
        ? '<tr><td style="padding:3px 0;color:#999">Cond. fiscal</td><td style="padding:3px 0">' + p.condFiscal + '</td></tr>'
        : '<tr><td style="padding:3px 0;color:#999;vertical-align:top">Facturación</td>'
          + '<td style="padding:3px 0">' + (f.razonSocial || '') + '<br>CUIT: ' + (f.cuitFacturacion || '') + ' — ' + p.condFiscal + '</td></tr>')
    : '';

  const adminBannerHTML = isAdmin
    ? '<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:10px 16px;margin-bottom:18px;font-size:13px;color:#856404">'
      + '<strong>Copia administrador</strong> — Venta registrada ' + p.fecha + ' · Dispositivo&nbsp;' + (p.dispositivo || '?')
      + '</div>'
    : '';

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;padding:0;background:#f5f5f3;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a">'
    + '<div style="max-width:560px;margin:24px auto;background:#ffffff;border-radius:10px;overflow:hidden;box-shadow:0 2px 14px rgba(0,0,0,.1)">'

    + '<div style="background:' + DARK + ';padding:20px 28px;display:flex;justify-content:space-between;align-items:center">'
    + '<div><div style="font-size:19px;font-weight:700;color:' + GREEN + '">AGF Messenchymal</div>'
    + '<div style="font-size:11px;color:#9FD4C0;margin-top:2px">dermacells.com.ar · Argentina</div></div>'
    + '<div style="text-align:right">'
    + '<div style="font-size:10px;color:#9ab;letter-spacing:.1em;text-transform:uppercase">Comprobante de compra</div>'
    + '<div style="font-size:14px;font-weight:700;color:' + GREEN + ';margin-top:4px">Venta #' + p.ventaNum + '</div>'
    + '<div style="font-size:10px;color:#9ab;margin-top:2px">BAAS 2026</div>'
    + '</div></div>'

    + '<div style="padding:24px 28px">'
    + adminBannerHTML

    + '<div style="font-size:10px;font-weight:700;letter-spacing:.15em;color:' + DARK + ';text-transform:uppercase;border-bottom:1px solid #e8e8e0;padding-bottom:4px;margin-bottom:10px">Datos del cliente</div>'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px">'
    + '<tr><td style="padding:3px 0;color:#999;width:32%">Nombre</td><td style="padding:3px 0">' + (c.nombre || '') + ' ' + (c.apellido || '') + '</td></tr>'
    + '<tr><td style="padding:3px 0;color:#999">CUIT / CUIL</td><td style="padding:3px 0">' + (c.cuit || '') + '</td></tr>'
    + (c.mail      ? '<tr><td style="padding:3px 0;color:#999">Email</td><td style="padding:3px 0">' + c.mail + '</td></tr>' : '')
    + (c.tel       ? '<tr><td style="padding:3px 0;color:#999">Teléfono</td><td style="padding:3px 0">' + c.tel + '</td></tr>' : '')
    + (c.localidad ? '<tr><td style="padding:3px 0;color:#999">Localidad</td><td style="padding:3px 0">' + c.localidad + '</td></tr>' : '')
    + condFiscalHTML
    + '</table>'

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
    + tcRowHTML
    + '</table>'

    + '<div style="border-top:2px solid ' + DARK + ';padding-top:14px;display:flex;justify-content:space-between;align-items:center">'
    + '<span style="font-size:12px;color:#999;text-transform:uppercase;letter-spacing:.08em">Total</span>'
    + '<span style="font-size:22px;font-weight:700;color:' + DARK + '">' + totalDisplay + '</span>'
    + '</div>'
    + '</div>'

    + '<div style="background:#f7f7f5;border-top:1px solid #e8e8e0;padding:12px 28px;display:flex;justify-content:space-between;align-items:center">'
    + '<span style="font-size:11px;color:#bbb">Documento válido como comprobante de compra</span>'
    + '<span style="font-size:11px;color:#bbb">AGF Messenchymal · 2026</span>'
    + '</div>'

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
    cliente: { nombre:'Ana', apellido:'García', cuit:'20-12345678-0', mail:'', tel:'', localidad:'CABA' },
    facturacion: { razonSocial:'Ana García', cuitFacturacion:'20-12345678-0', mismosContacto:true },
    condFiscal: 'Resp. Inscripto',
    cajas: [
      { caja:1, tipo:'cerrada',   detalle:'Dermal',           descCaja:0,  precio:750 },
      { caja:2, tipo:'combinada', detalle:'Dermal, Capillary, Pink, Biomask, Dermal', descCaja:10, precio:810 }
    ],
    totalUSD: 1560,
    totalARS: null,
    moneda: 'USD',
    tipoCambio: null,
    metodoCobro: 'Transferencia',
    descuentoGlobal: 0
  };

  guardarVenta(payload);
  actualizarResumen(payload);

  // Probar generación y guardado de PDF en Drive
  const { blob, url } = generarPdfRecibo(payload);
  Logger.log('PDF generado OK: ' + blob.getName() + ' (' + blob.getBytes().length + ' bytes)');
  Logger.log('URL en Drive: ' + url);
}

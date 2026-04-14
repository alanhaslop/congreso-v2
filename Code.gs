// ═══════════════════════════════════════════════════════════════
//  AGF Messenchymal — Ventas Congreso · Google Apps Script
//  Pegá este código en script.google.com y configurá las dos
//  constantes de abajo antes de hacer "Implementar > Web app".
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '';           // ← ID de la hoja (en la URL: /d/XXXX/edit)
const ADMIN_EMAIL    = '';           // ← email que recibe copia de cada venta (puede quedar vacío)
const REPLY_TO       = 'contacto@dermacells.com.ar';
const SHEET_VENTAS   = 'Ventas';
const SHEET_RESUMEN  = 'Resumen';


// ── ENTRY POINT ─────────────────────────────────────────────────
function doPost(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    if (!e || !e.postData) throw new Error('Sin payload');
    const p = JSON.parse(e.postData.contents);
    if (!p.ventaNum)   throw new Error('Payload inválido: falta ventaNum');

    guardarVenta(p);
    actualizarResumen(p);
    if (p.cliente && p.cliente.mail)  enviarEmailCliente(p);
    if (ADMIN_EMAIL)                  enviarEmailAdmin(p);

    out.setContent(JSON.stringify({ ok: true, ventaNum: p.ventaNum }));
  } catch (err) {
    console.error('doPost error:', err.message, err.stack);
    out.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return out;
}

// Endpoint GET para verificar que el deploy funciona
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'AGF Ventas API activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── GUARDAR VENTA EN HOJA ────────────────────────────────────────
function guardarVenta(p) {
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
    new Date(),                                            // A: Timestamp
    p.ventaNum,                                            // B: Venta #
    p.dispositivo    || '',                                // C: Dispositivo
    p.fecha          || '',                                // D: Fecha (local)
    c.nombre         || '',                                // E: Nombre
    c.apellido       || '',                                // F: Apellido
    c.cuit           || '',                                // G: CUIT/CUIL
    c.mail           || '',                                // H: Email
    c.tel            || '',                                // I: Teléfono
    c.localidad      || '',                                // J: Localidad
    p.condFiscal     || '',                                // K: Cond. Fiscal
    f.mismosContacto ? '' : (f.razonSocial     || ''),    // L: Razón Social fact.
    f.mismosContacto ? '' : (f.cuitFacturacion || ''),    // M: CUIT Facturación
    p.metodoCobro    || '',                                // N: Método cobro
    p.moneda         || 'USD',                             // O: Moneda
    p.tipoCambio     || '',                                // P: Tipo de cambio
    (p.cajas || []).length,                                // Q: Cant. cajas
    cajasDetalle,                                          // R: Detalle cajas
    p.descuentoGlobal || 0,                               // S: Desc. global %
    p.subtotalUSD    || p.totalUSD || 0,                  // T: Subtotal U$D
    p.totalUSD       || 0,                                 // U: Total U$D
    p.totalARS       || ''                                 // V: Total ARS
  ]);
}

function crearHeadersVentas(sheet) {
  const headers = [
    'Timestamp','Venta #','Dispositivo','Fecha local',
    'Nombre','Apellido','CUIT/CUIL','Email','Teléfono','Localidad',
    'Cond. Fiscal','Razón Social Fact.','CUIT Facturación',
    'Método cobro','Moneda','Tipo de cambio',
    'Cant. cajas','Detalle cajas',
    'Desc. global %','Subtotal U$D','Total U$D','Total ARS'
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);

  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setBackground('#1D9E75').setFontColor('#ffffff').setFontWeight('bold');

  sheet.setColumnWidth(1,  160); // Timestamp
  sheet.setColumnWidth(4,  140); // Fecha local
  sheet.setColumnWidth(18, 280); // Detalle cajas
  sheet.setColumnWidths(5, 2, 110); // Nombre / Apellido
}


// ── HOJA DE RESUMEN (acumulado del día) ─────────────────────────
function actualizarResumen(p) {
  const sheet = obtenerOCrearHoja(SHEET_RESUMEN, crearHeadersResumen);

  const hoy      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const datos    = sheet.getDataRange().getValues();
  const colFecha = 0, colMetodo = 1, colCajas = 2, colUSD = 3, colARS = 4, colVentas = 5;

  // Buscar fila existente para hoy + método
  let filaExistente = -1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colFecha] === hoy && datos[i][colMetodo] === p.metodoCobro) {
      filaExistente = i + 1; // 1-indexed
      break;
    }
  }

  const cantCajas = (p.cajas || []).length;
  const totalUSD  = p.totalUSD  || 0;
  const totalARS  = p.totalARS  || 0;

  if (filaExistente > 0) {
    // Acumular en fila existente
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
  const hdr = sheet.getRange(1, 1, 1, 6);
  hdr.setBackground('#1B3A52').setFontColor('#ffffff').setFontWeight('bold');
}


// ── EMAIL CLIENTE ────────────────────────────────────────────────
function enviarEmailCliente(p) {
  const c = p.cliente;
  GmailApp.sendEmail(
    c.mail,
    `AGF Messenchymal — Comprobante de compra #${p.ventaNum}`,
    buildTextoPlano(p),
    { htmlBody: buildEmailHTML(p, false), name: 'AGF Messenchymal Argentina', replyTo: REPLY_TO }
  );
}

function enviarEmailAdmin(p) {
  const c = p.cliente || {};
  GmailApp.sendEmail(
    ADMIN_EMAIL,
    `[Venta #${p.ventaNum}] ${c.nombre} ${c.apellido} — u$${p.totalUSD} — ${p.metodoCobro}`,
    buildTextoPlano(p),
    { htmlBody: buildEmailHTML(p, true), name: 'AGF Ventas Congreso' }
  );
}

function buildTextoPlano(p) {
  const c = p.cliente || {};
  const lineas = [
    `AGF Messenchymal — Venta #${p.ventaNum}  |  Dispositivo ${p.dispositivo || '?'}`,
    `Fecha: ${p.fecha}`,
    `Cliente: ${c.nombre} ${c.apellido}  |  CUIT: ${c.cuit}`,
    '',
    'Detalle:',
    ...(p.cajas || []).map(cj =>
      `  Caja ${cj.caja} (${cj.tipo}): ${cj.detalle}${cj.descCaja > 0 ? ` [-${cj.descCaja}%]` : ''} = u$${cj.precio}`
    ),
    '',
    p.descuentoGlobal > 0 ? `  Descuento global: ${p.descuentoGlobal}%  (-u$${(p.subtotalUSD || p.totalUSD) - p.totalUSD})` : null,
    `Total: u$${p.totalUSD}${p.totalARS ? '  /  AR$' + Number(p.totalARS).toLocaleString('es-AR') : ''}`,
    `Método: ${p.metodoCobro}  |  Moneda: ${p.moneda}${p.tipoCambio ? '  |  TC: $' + p.tipoCambio : ''}`,
    p.condFiscal ? `Cond. fiscal: ${p.condFiscal}` : null,
  ].filter(l => l !== null).join('\n');
  return lineas;
}


// ── EMAIL HTML ───────────────────────────────────────────────────
function buildEmailHTML(p, isAdmin) {
  const c     = p.cliente     || {};
  const f     = p.facturacion || {};
  const GREEN = '#1D9E75';
  const DARK  = '#1B3A52';

  // Filas de cajas
  const filasHTML = (p.cajas || []).map((cj, i) => {
    const bg = i % 2 === 0 ? '#f9f9f7' : '#ffffff';
    const descBadge = cj.descCaja > 0
      ? ` <span style="color:#e67e22;font-size:11px">[-${cj.descCaja}%]</span>` : '';
    return `
      <tr style="background:${bg}">
        <td style="padding:8px 12px;border-bottom:1px solid #eee;font-size:13px">
          Caja ${cj.caja} — ${cj.tipo === 'cerrada' ? 'Cerrada' : 'Combinada'}
        </td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#666;font-size:12px">
          ${cj.detalle}${descBadge}
        </td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600;font-size:13px">
          u$${cj.precio}
        </td>
      </tr>`;
  }).join('');

  const descRowHTML = p.descuentoGlobal > 0
    ? `<tr>
        <td colspan="2" style="padding:5px 12px;font-size:11px;color:#999;font-style:italic">
          Descuento global: ${p.descuentoGlobal}%
        </td>
        <td style="padding:5px 12px;text-align:right;color:#c0392b;font-size:12px">
          −u$${(p.subtotalUSD || p.totalUSD) - p.totalUSD}
        </td>
      </tr>` : '';

  const totalDisplay = `u$${p.totalUSD}${p.totalARS
    ? `&nbsp;<span style="font-size:14px;color:#666">/ AR$${Number(p.totalARS).toLocaleString('es-AR')}</span>`
    : ''}`;

  const tcRowHTML = p.tipoCambio
    ? `<tr>
        <td style="padding:3px 0;color:#999">Tipo de cambio</td>
        <td style="padding:3px 0">$${Number(p.tipoCambio).toLocaleString('es-AR')} AR$/U$D</td>
       </tr>` : '';

  const condFiscalHTML = p.condFiscal
    ? (f.mismosContacto
        ? `<tr><td style="padding:3px 0;color:#999">Cond. fiscal</td><td style="padding:3px 0">${p.condFiscal}</td></tr>`
        : `<tr>
            <td style="padding:3px 0;color:#999;vertical-align:top">Facturación</td>
            <td style="padding:3px 0">${f.razonSocial || ''}<br>
              CUIT: ${f.cuitFacturacion || ''} — ${p.condFiscal}
            </td>
           </tr>`)
    : '';

  const adminBannerHTML = isAdmin
    ? `<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:10px 16px;
                   margin-bottom:18px;font-size:13px;color:#856404">
         <strong>Copia administrador</strong> — Venta registrada ${p.fecha} · Dispositivo&nbsp;${p.dispositivo || '?'}
       </div>` : '';

  return `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#f5f5f3;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a">
<div style="max-width:560px;margin:24px auto;background:#ffffff;border-radius:10px;overflow:hidden;box-shadow:0 2px 14px rgba(0,0,0,.1)">

  <!-- ── Header ── -->
  <div style="background:${DARK};padding:20px 28px;display:flex;justify-content:space-between;align-items:center">
    <div>
      <div style="font-size:19px;font-weight:700;color:${GREEN}">AGF Messenchymal</div>
      <div style="font-size:11px;color:#9FD4C0;margin-top:2px">dermacells.com.ar · Argentina</div>
    </div>
    <div style="text-align:right">
      <div style="font-size:10px;color:#9ab;letter-spacing:.1em;text-transform:uppercase">Comprobante de compra</div>
      <div style="font-size:14px;font-weight:700;color:${GREEN};margin-top:4px">Venta #${p.ventaNum}</div>
      <div style="font-size:10px;color:#9ab;margin-top:2px">BAAS 2026</div>
    </div>
  </div>

  <div style="padding:24px 28px">
    ${adminBannerHTML}

    <!-- ── Cliente ── -->
    <div style="font-size:10px;font-weight:700;letter-spacing:.15em;color:${DARK};text-transform:uppercase;
                border-bottom:1px solid #e8e8e0;padding-bottom:4px;margin-bottom:10px">Datos del cliente</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px">
      <tr><td style="padding:3px 0;color:#999;width:32%">Nombre</td>
          <td style="padding:3px 0">${c.nombre || ''} ${c.apellido || ''}</td></tr>
      <tr><td style="padding:3px 0;color:#999">CUIT / CUIL</td>
          <td style="padding:3px 0">${c.cuit || ''}</td></tr>
      ${c.mail      ? `<tr><td style="padding:3px 0;color:#999">Email</td><td style="padding:3px 0">${c.mail}</td></tr>` : ''}
      ${c.tel       ? `<tr><td style="padding:3px 0;color:#999">Teléfono</td><td style="padding:3px 0">${c.tel}</td></tr>` : ''}
      ${c.localidad ? `<tr><td style="padding:3px 0;color:#999">Localidad</td><td style="padding:3px 0">${c.localidad}</td></tr>` : ''}
      ${condFiscalHTML}
    </table>

    <!-- ── Cajas ── -->
    <div style="font-size:10px;font-weight:700;letter-spacing:.15em;color:${DARK};text-transform:uppercase;
                border-bottom:1px solid #e8e8e0;padding-bottom:4px;margin-bottom:0">Detalle de cajas</div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
      <thead>
        <tr style="border-bottom:1px solid #e8e8e0">
          <th style="padding:8px 12px;text-align:left;font-weight:500;color:#999;font-size:11px">Descripción</th>
          <th style="padding:8px 12px;text-align:left;font-weight:500;color:#999;font-size:11px">Contenido</th>
          <th style="padding:8px 12px;text-align:right;font-weight:500;color:#999;font-size:11px">Importe</th>
        </tr>
      </thead>
      <tbody>
        ${filasHTML}
        ${descRowHTML}
      </tbody>
    </table>

    <!-- ── Cobro ── -->
    <div style="font-size:10px;font-weight:700;letter-spacing:.15em;color:${DARK};text-transform:uppercase;
                border-bottom:1px solid #e8e8e0;padding-bottom:4px;margin-bottom:10px">Condiciones de cobro</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px">
      <tr><td style="padding:3px 0;color:#999;width:32%">Método</td>
          <td style="padding:3px 0">${p.metodoCobro || ''}</td></tr>
      <tr><td style="padding:3px 0;color:#999">Moneda</td>
          <td style="padding:3px 0">${p.moneda === 'USD' ? 'Dólares americanos' : 'Pesos argentinos'}</td></tr>
      ${tcRowHTML}
    </table>

    <!-- ── Total ── -->
    <div style="border-top:2px solid ${DARK};padding-top:14px;display:flex;justify-content:space-between;align-items:center">
      <span style="font-size:12px;color:#999;text-transform:uppercase;letter-spacing:.08em">Total</span>
      <span style="font-size:22px;font-weight:700;color:${DARK}">${totalDisplay}</span>
    </div>
  </div>

  <!-- ── Footer ── -->
  <div style="background:#f7f7f5;border-top:1px solid #e8e8e0;padding:12px 28px;
              display:flex;justify-content:space-between;align-items:center">
    <span style="font-size:11px;color:#bbb">Documento válido como comprobante de compra</span>
    <span style="font-size:11px;color:#bbb">AGF Messenchymal · 2026</span>
  </div>

</div>
</body>
</html>`;
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

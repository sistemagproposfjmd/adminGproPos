// ============================================================================
// MÓDULO DE COBRANZA DIGITAL Y SOPORTES WEB (GPro)
// ============================================================================
const HOJA_SOPORTE_WEB = 'SoporteWeb';

function asegurarHojaSoporteWeb_(ss) {
  var hoja = ss.getSheetByName(HOJA_SOPORTE_WEB);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_SOPORTE_WEB);
    hoja.appendRow([
      'IdSoporte',          // Col 1
      'FechaRegistro',      // Col 2
      'IdSpreadsheet',      // Col 3
      'NombreEmpresa',      // Col 4
      'RifCliente',         // Col 5
      'NombreCliente',      // Col 6
      'CorreoCliente',      // Col 7
      'FacturasAfectadas',  // Col 8 (JSON o lista separada por comas)
      'Moneda',             // Col 9
      'MontoUSD',           // Col 10
      'MontoBs',            // Col 11
      'Tasa',               // Col 12
      'FormaPago',          // Col 13 (PagoMovil, Transferencia, Zelle, Deposito)
      'BancoOrigen',        // Col 14
      'BancoDestino',       // Col 15
      'NroCuentaDestino',   // Col 16
      'Referencia',         // Col 17
      'FechaPago',          // Col 18
      'Status',             // Col 19: PENDIENTE, CONVERTIDO_ANTICIPO, RECHAZADO
      'IdAnticipoGenerado', // Col 20: ID en MovBancoCaja / CobrosPagos
      'Observaciones'       // Col 21
    ]);
  }
  return hoja;
}
function obtenerDatosParaEnvioCobro(rifOCliente, idSpreadsheetEmpresa) {
  try {
    var ss = idSpreadsheetEmpresa ? SpreadsheetApp.openById(idSpreadsheetEmpresa) : SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Buscar Correo en hoja Clientes
    var shClientes = ss.getSheetByName('Clientes');
    var correoCliente = '';
    var nombreCliente = '';
    var telefonoCliente = '';

    if (shClientes) {
      var dCli = shClientes.getDataRange().getValues();
      var hCli = dCli[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
      var cRif = hCli.indexOf('rifclientes');
      var cNom = hCli.indexOf('nombreclientes');
      var cCor = hCli.indexOf('correocliente');
      var cTel = hCli.indexOf('telefonocliente');

      var busqueda = String(rifOCliente || '').trim().toLowerCase();
      for (var i = 1; i < dCli.length; i++) {
        var rVal = cRif !== -1 ? String(dCli[i][cRif] || '').trim().toLowerCase() : '';
        var nVal = cNom !== -1 ? String(dCli[i][cNom] || '').trim().toLowerCase() : '';
        if (rVal === busqueda || nVal === busqueda) {
          correoCliente = cCor !== -1 ? String(dCli[i][cCor] || '') : '';
          nombreCliente = cNom !== -1 ? String(dCli[i][cNom] || '') : '';
          telefonoCliente = cTel !== -1 ? String(dCli[i][cTel] || '') : '';
          break;
        }
      }
    }

    // 2. Facturas pendientes desde CobrosPagos
    var facturas = getFacturasPorTransaccion('Venta', idSpreadsheetEmpresa);
    var busq = String(rifOCliente || '').trim().toLowerCase();
    var facturasCliente = facturas.filter(function(f) {
      var rF = String(f.Rif || '').trim().toLowerCase();
      var nF = String(f.ProveedorCliente || '').trim().toLowerCase();
      var saldo = Number(f.SaldoFacturaUSD || 0);
      var noAnulada = String(f.cobradaPagada || '').toLowerCase() !== 'anulada';
      return (rF === busq || nF === busq) && saldo > 0.01 && noAnulada;
    });

    return {
      success: true,
      nombre: nombreCliente || rifOCliente,
      rif: rifOCliente,
      correo: correoCliente,
      telefono: telefonoCliente,
      facturas: facturasCliente
    };
  } catch (err) {
    Logger.log('Error obtenerDatosParaEnvioCobro: ' + err);
    return { success: false, message: err.message };
  }
}
function enviarNotificacionCobroCliente(payload) {
  try {
    var rif = payload.rif;
    var correoDestino = payload.correo;
    var nombre = payload.nombre;
    var facturasSeleccionadas = payload.facturas || [];
    var idSpreadsheet = payload.idSpreadsheet || '';
    var empresaNombre = payload.empresaNombre || 'GPro POS';

    if (!correoDestino || correoDestino.indexOf('@') === -1) {
      throw new Error('El cliente no posee un correo electrónico válido.');
    }
    if (!facturasSeleccionadas.length) {
      throw new Error('Debe seleccionar al menos una factura para notificar.');
    }

    // Cuentas Bancarias activas para colocar en el correo
    var ss = idSpreadsheet ? SpreadsheetApp.openById(idSpreadsheet) : SpreadsheetApp.getActiveSpreadsheet();
    var cuentasBancarias = [];
    var shBancos = ss.getSheetByName('datosBancarios');
    if (shBancos) {
      var dB = shBancos.getDataRange().getValues();
      for (var b = 1; b < dB.length; b++) {
        var nBanco = String(dB[b][11] || dB[b][2] || '').trim(); // NombreBancoCuenta
        var nCta = String(dB[b][9] || '').trim();              // NroCuenta
        var tCta = String(dB[b][13] || '').trim();             // tipoCuenta
        var monB = String(dB[b][12] || 'Bolívares').trim();    // MonedaBanco
        if (nBanco && nCta) {
          cuentasBancarias.push({ banco: nBanco, cuenta: nCta, tipo: tCta, moneda: monB });
        }
      }
    }

    // Generar Token de cobro seguro (expira en 15 días)
    var payloadToken = {
      tipo: 'soporte_pago_cliente',
      idSpreadsheet: idSpreadsheet,
      empresaNombre: empresaNombre,
      rif: rif,
      nombre: nombre,
      facturas: facturasSeleccionadas.map(function(f) {
        return { nro: f.NroFactura, saldoUSD: f.SaldoFacturaUSD, saldoBs: f.SaldoFactura, fecha: f.FechaFactura };
      }),
      exp: Date.now() + (15 * 24 * 60 * 60 * 1000) // 15 días
    };

    var jsonToken = JSON.stringify(payloadToken);
    var tokenB64 = Utilities.base64EncodeWebSafe(jsonToken);
    var firmaB64 = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(jsonToken, SESSION_SECRET));
    var tokenFinal = (tokenB64 + '.' + firmaB64).replace(/=+$/g, '');

    // URL WebApp limpia
    var rawUrl = ScriptApp.getService().getUrl();
    var baseUrl = (rawUrl && rawUrl.indexOf('http') === 0) ? rawUrl.split('?')[0] : URL_WEBAPP_FIJA;
    var urlReporte = baseUrl + '?action=reportarPago&token=' + encodeURIComponent(tokenFinal);

    // Armar tabla HTML de facturas
    var totalUSD = 0;
    var filasFacturasHtml = facturasSeleccionadas.map(function(f) {
      totalUSD += Number(f.SaldoFacturaUSD || 0);
      return `<tr>
        <td style="padding:10px;border-bottom:1px solid #e2e8f0;font-weight:600;">#${f.NroFactura}</td>
        <td style="padding:10px;border-bottom:1px solid #e2e8f0;color:#64748b;">${f.FechaFactura || '-'}</td>
        <td style="padding:10px;border-bottom:1px solid #e2e8f0;text-align:right;color:#0f172a;font-weight:700;">$ ${Number(f.SaldoFacturaUSD || 0).toFixed(2)}</td>
      </tr>`;
    }).join('');

    // Armar bloque de cuentas bancarias
    var cuentasHtml = cuentasBancarias.slice(0, 4).map(function(c) {
      return `<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:8px;">
        <strong style="color:#0f172a;display:block;">${c.banco} (${c.moneda})</strong>
        <span style="font-family:monospace;color:#475569;font-size:13px;">${c.cuenta}</span>
        <span style="display:block;font-size:11px;color:#94a3b8;">${c.tipo}</span>
      </div>`;
    }).join('');

    // Plantilla de Correo Responsive
    var htmlCorreo = `
    <!DOCTYPE html>
    <html>
    <head><meta charset="utf-8"></head>
    <body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;background-color:#f1f5f9;margin:0;padding:24px;color:#334155;">
      <div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 10px 25px rgba(0,0,0,0.08);border:1px solid #e2e8f0;">
        <div style="background:linear-gradient(135deg,#0f172a,#1e293b);padding:30px 24px;text-align:center;border-bottom:4px solid #f97316;">
          <h1 style="color:#ffffff;margin:0;font-size:22px;letter-spacing:-0.5px;">${empresaNombre}</h1>
          <p style="color:#94a3b8;margin:6px 0 0;font-size:14px;">Estado de Cuenta / Notificación de Cobro</p>
        </div>

        <div style="padding:28px 24px;">
          <p style="font-size:15px;line-height:1.6;margin-top:0;">Estimado(a) <strong>${nombre}</strong> (RIF: ${rif}),</p>
          <p style="font-size:14px;color:#64748b;line-height:1.6;">Le hacemos llegar el detalle de sus facturas pendientes por conciliar a la fecha:</p>

          <table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:13.5px;">
            <thead>
              <tr style="background:#f8fafc;color:#475569;text-align:left;">
                <th style="padding:10px;border-bottom:2px solid #cbd5e1;">N° Factura</th>
                <th style="padding:10px;border-bottom:2px solid #cbd5e1;">Emisión</th>
                <th style="padding:10px;border-bottom:2px solid #cbd5e1;text-align:right;">Saldo</th>
              </tr>
            </thead>
            <tbody>
              ${filasFacturasHtml}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="2" style="padding:12px 10px;font-weight:700;font-size:15px;color:#0f172a;">Total por Cancelar:</td>
                <td style="padding:12px 10px;text-align:right;font-weight:800;font-size:17px;color:#ea580c;">$ ${totalUSD.toFixed(2)}</td>
              </tr>
            </tfoot>
          </table>

          <div style="text-align:center;margin:32px 0;">
            <a href="${urlReporte}" target="_blank" style="display:inline-block;background:linear-gradient(135deg,#f97316,#ea580c);color:#ffffff;text-decoration:none;padding:15px 36px;border-radius:12px;font-weight:800;font-size:15px;box-shadow:0 6px 20px rgba(249,115,22,0.35);">
              💳 Reportar Soporte de Pago Aquí
            </a>
            <p style="font-size:11.5px;color:#94a3b8;margin-top:10px;">(Haga clic para ingresar su número de referencia, monto y fecha)</p>
          </div>

          ${cuentasHtml ? `
          <div style="margin-top:24px;border-top:1px dashed #cbd5e1;padding-top:20px;">
            <h4 style="margin:0 0 12px;color:#0f172a;font-size:14px;">Cuentas Bancarias Disponibles:</h4>
            ${cuentasHtml}
          </div>` : ''}

          <p style="font-size:12px;color:#94a3b8;margin-top:28px;text-align:center;line-height:1.5;">
            Si ya realizó este pago, por favor haga clic en el botón superior para ingresar su referencia y actualizar su saldo automáticamente.
          </p>
        </div>

        <div style="background:#f8fafc;padding:16px 24px;text-align:center;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;">
          ${empresaNombre} &copy; 2026 — Sistema Automatizado de Cobranzas GPro
        </div>
      </div>
    </body>
    </html>`;

    MailApp.sendEmail({
      to: correoDestino,
      subject: `📄 Recordatorio de Pago / Facturas Pendientes - ${empresaNombre}`,
      htmlBody: htmlCorreo,
      name: empresaNombre
    });

    return { success: true, message: 'Correo de cobranza enviado exitosamente a ' + correoDestino };
  } catch (err) {
    Logger.log('Error enviarNotificacionCobroCliente: ' + err);
    return { success: false, message: err.message };
  }
}

function validarTokenCobroCliente_(token) {
  if (!token || typeof token !== 'string') return null;
  
  var partes = token.trim().split('.');
  if (partes.length !== 2) return null;

  try {
    var payloadB64 = partes[0];
    var firmaRecibida = partes[1];

    // Decodificar payload en formato WebSafe
    var jsonPayload = Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString();
    
    // Generar firma esperada sobre el JSON original
    var firmaEsperada = Utilities.base64EncodeWebSafe(
      Utilities.computeHmacSha256Signature(jsonPayload, SESSION_SECRET)
    ).replace(/=+$/g, '');

    // Comparación segura en formato WebSafe exacto
    if (firmaEsperada !== firmaRecibida) {
      Logger.log('[TOKEN] Firma inválida. Esperada: ' + firmaEsperada + ' | Recibida: ' + firmaRecibida);
      return null;
    }

    var payload = JSON.parse(jsonPayload);
    if (payload.tipo !== 'soporte_pago_cliente') return null;
    if (payload.exp && payload.exp < Date.now()) {
      Logger.log('[TOKEN] Token expirado');
      return null;
    }

    return payload;
  } catch (e) {
    Logger.log('[TOKEN] Error en validarTokenCobroCliente_: ' + e.message);
    return null;
  }
}

function registrarSoporteWebCliente(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);

    if (!payload || !payload.token) throw new Error('Token de validación ausente.');
    var sesionToken = validarTokenCobroCliente_(payload.token);
    if (!sesionToken) throw new Error('El enlace de reporte es inválido o ha expirado. Solicite uno nuevo.');

    var idTarget = sesionToken.idSpreadsheet;
    var ss = idTarget ? SpreadsheetApp.openById(idTarget) : SpreadsheetApp.getActiveSpreadsheet();
    var shSoporte = asegurarHojaSoporteWeb_(ss);

    var idSoporte = 'SOP-' + Date.now().toString().slice(-7);
    var fechaRegistro = new Date();

    var facturasTexto = Array.isArray(payload.facturasSeleccionadas)
      ? payload.facturasSeleccionadas.join(', ')
      : String(payload.facturasSeleccionadas || '');

    var fila = [
      idSoporte,                               // Col 1: IdSoporte
      fechaRegistro,                           // Col 2: FechaRegistro
      idTarget,                                // Col 3: IdSpreadsheet
      sesionToken.empresaNombre || '',         // Col 4: NombreEmpresa
      sesionToken.rif || '',                   // Col 5: RifCliente
      sesionToken.nombre || '',                // Col 6: NombreCliente
      payload.correoContacto || '',            // Col 7: CorreoCliente
      facturasTexto,                           // Col 8: FacturasAfectadas
      payload.moneda || 'USD',                 // Col 9: Moneda
      Number(payload.montoUSD || 0),           // Col 10: MontoUSD
      Number(payload.montoBs || 0),            // Col 11: MontoBs
      Number(payload.tasa || 1),               // Col 12: Tasa
      payload.formaPago || 'Transferencia',    // Col 13: FormaPago
      payload.bancoOrigen || '',               // Col 14: BancoOrigen
      payload.bancoDestino || '',              // Col 15: BancoDestino
      payload.nroCuentaDestino || '',          // Col 16: NroCuentaDestino
      payload.referencia || '',                // Col 17: Referencia
      payload.fechaPago || '',                 // Col 18: FechaPago
      'PENDIENTE',                             // Col 19: Status
      '',                                      // Col 20: IdAnticipoGenerado
      payload.observaciones || ''              // Col 21: Observaciones
    ];

    shSoporte.appendRow(fila);

    // Opcional: Notificar al correo del administrador
    try {
      MailApp.sendEmail({
        to: 'sistemagpropos@gmail.com',
        subject: `🔔 Nuevo Soporte de Cobro Reportado: ${sesionToken.nombre} ($ ${Number(payload.montoUSD || 0).toFixed(2)})`,
        htmlBody: `
          <h3>Se ha recibido un nuevo soporte de pago web</h3>
          <ul>
            <li><strong>Cliente:</strong> ${sesionToken.nombre} (${sesionToken.rif})</li>
            <li><strong>Monto:</strong> $ ${Number(payload.montoUSD || 0).toFixed(2)} / Bs. ${Number(payload.montoBs || 0).toFixed(2)}</li>
            <li><strong>Método:</strong> ${payload.formaPago} (${payload.bancoOrigen} ➔ ${payload.bancoDestino})</li>
            <li><strong>Referencia:</strong> ${payload.referencia}</li>
            <li><strong>Facturas:</strong> ${facturasTexto}</li>
          </ul>
          <p>Ingrese al módulo <em>Facturas Lista</em> para verificar y convertir en Anticipo/Cobro.</p>
        `
      });
    } catch(eMail) {}

    return {
      success: true,
      idSoporte: idSoporte,
      mensaje: '¡Soporte registrado exitosamente! Su pago será verificado por el departamento de administración.'
    };
  } catch (err) {
    Logger.log('Error registrarSoporteWebCliente: ' + err);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}
function obtenerSoportesWebPendientes(idSpreadsheetEmpresa) {
  try {
    var ss = idSpreadsheetEmpresa ? SpreadsheetApp.openById(idSpreadsheetEmpresa) : SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_SOPORTE_WEB);
    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var soportes = [];

    for (var i = 1; i < datos.length; i++) {
      var row = datos[i];
      var st = String(row[18] || '').trim().toUpperCase(); // Status
      if (st === 'PENDIENTE') {
        soportes.push({
          fila: i + 1,
          idSoporte: row[0],
          fechaRegistro: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : row[1],
          rif: row[4],
          cliente: row[5],
          correo: row[6],
          facturas: row[7],
          moneda: row[8],
          montoUSD: Number(row[9] || 0),
          montoBs: Number(row[10] || 0),
          tasa: Number(row[11] || 1),
          formaPago: row[12],
          bancoOrigen: row[13],
          bancoDestino: row[14],
          cuentaDestino: row[15],
          referencia: row[16],
          fechaPago: row[17],
          observaciones: row[20]
        });
      }
    }
    return soportes;
  } catch (e) {
    Logger.log('Error obtenerSoportesWebPendientes: ' + e);
    return [];
  }
}
function convertirSoporteEnAnticipo(idSoporte, cuentaBancariaSeleccionada, idSpreadsheetEmpresa) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = idSpreadsheetEmpresa ? SpreadsheetApp.openById(idSpreadsheetEmpresa) : SpreadsheetApp.getActiveSpreadsheet();
    var shSoporte = ss.getSheetByName(HOJA_SOPORTE_WEB);
    if (!shSoporte) throw new Error('Hoja SoporteWeb no encontrada');

    var dataSop = shSoporte.getDataRange().getValues();
    var filaTarget = -1;
    var soporte = null;

    for (var i = 1; i < dataSop.length; i++) {
      if (String(dataSop[i][0]).trim() === String(idSoporte).trim()) {
        filaTarget = i + 1;
        soporte = dataSop[i];
        break;
      }
    }

    if (!soporte) throw new Error('Soporte no encontrado.');
    if (String(soporte[18]).toUpperCase() !== 'PENDIENTE') {
      throw new Error('Este soporte ya fue procesado anteriormente.');
    }

    var montoUSD = Number(soporte[9]) || 0;
    var montoBs = Number(soporte[10]) || 0;
    var tasa = Number(soporte[11]) || 1;
    var cliente = soporte[5];
    var rif = soporte[4];
    var referencia = soporte[16];
    var fechaPago = soporte[17] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Desglosar cuenta bancaria
    var partesCta = String(cuentaBancariaSeleccionada || '').split('|');
    var ctaTipo = partesCta[0] || 'Banco';
    var ctaNombre = partesCta[1] || soporte[14] || 'Banco';
    var ctaCod = partesCta[2] || soporte[15] || '';
    var ctaCodContable = partesCta[3] || '';
    var ctaNomContable = partesCta[4] || ctaNombre;

    // 1. Asentar en MovBancoCaja (las 20 columnas)
    var shMB = ss.getSheetByName('MovBancoCaja');
    var idMov = 'MOV-' + Date.now().toString().slice(-7);
    if (shMB) {
      var partesF = String(fechaPago).split('-');
      var fDate = (partesF.length === 3) 
        ? new Date(parseInt(partesF[0], 10), parseInt(partesF[1], 10) - 1, parseInt(partesF[2], 10))
        : new Date();

      var filaMB = new Array(20).fill('');
      filaMB[0] = idMov;                           // IdMovimiento
      filaMB[1] = fDate;                           // FechaMov
      filaMB[2] = ctaTipo;                         // TipoMov
      filaMB[3] = ctaCod;                          // NroCuenta
      filaMB[4] = ctaNombre;                       // NombreBanco
      filaMB[5] = ctaCod;                          // CodigoBancos
      filaMB[6] = ctaNombre;                       // Cuenta
      filaMB[7] = referencia;                      // ReferenciaMov
      filaMB[8] = 'Anticipo Web Cliente: ' + cliente + ' (Ref: ' + referencia + ')'; // descMov
      filaMB[9] = 'Entrada';                       // TipoTransaccion
      filaMB[10] = montoUSD;                       // MontoMovUSD
      filaMB[11] = soporte[8] || 'Dolares';        // MonedaMov
      filaMB[12] = tasa;                           // TasaMov
      filaMB[13] = 0;                              // IGTF
      filaMB[14] = montoBs;                        // MontoMovBs
      filaMB[15] = ctaCodContable;                 // CodigoCuentaContable
      filaMB[16] = ctaNomContable;                 // CuentaContable
      filaMB[17] = '';                             // CuentaGasto
      filaMB[18] = cliente;                        // NombreCuentaMov
      filaMB[19] = new Date();                     // fechaRegistroMov

      shMB.appendRow(filaMB);
    }

    // 2. Asentar en CobrosPagos (las 23 columnas como Anticipo disponible)
    var shCP = ss.getSheetByName('CobrosPagos');
    var idAnticipoCP = (typeof obtenerSiguienteCorrelativoCobrosPagos === 'function')
      ? obtenerSiguienteCorrelativoCobrosPagos('A', true, idSpreadsheetEmpresa)
      : ('ANT-' + Date.now().toString().slice(-6));

    if (shCP) {
      var filaCP = new Array(23).fill('');
      filaCP[0] = cliente;                         // ProveedorCliente
      filaCP[1] = rif;                             // Rif
      filaCP[2] = idAnticipoCP;                    // NroFactura (ID del Anticipo)
      filaCP[3] = new Date();                      // FechaFactura
      filaCP[4] = 0;                               // TotalFactura
      filaCP[5] = montoBs;                         // SaldoFactura
      filaCP[6] = montoUSD;                        // SaldoFacturaUSD
      filaCP[7] = 'anticipoCobro';                 // cobradaPagada
      filaCP[8] = 'Anticipo';                      // Transaccion
      filaCP[9] = idAnticipoCP;                    // id
      filaCP[10] = montoUSD;                       // Abono
      filaCP[11] = montoBs;                        // AbonoBs
      filaCP[12] = 'Dolares';                      // Moneda
      filaCP[13] = tasa;                           // Tasa
      filaCP[14] = 'Soporte Web validado Ref: ' + referencia; // Descripcion
      filaCP[15] = new Date();                     // fechaAbono
      filaCP[16] = 'ANTICIPO';                     // TipoMovimiento
      filaCP[17] = ctaNombre;                      // CuentaDestino
      filaCP[18] = ctaCod;                         // NroCuentaDestino
      filaCP[19] = ctaNombre;                      // BancoDestino
      filaCP[20] = new Date();                     // fechaVencimiento

      shCP.appendRow(filaCP);
    }

    // 3. Marcar soporte como convertido
    shSoporte.getRange(filaTarget, 19).setValue('CONVERTIDO_ANTICIPO');
    shSoporte.getRange(filaTarget, 20).setValue(idAnticipoCP);

    return {
      success: true,
      mensaje: 'Soporte convertido exitosamente en Anticipo #' + idAnticipoCP + '. Ya se encuentra disponible para cruzar con facturas.',
      idAnticipo: idAnticipoCP
    };
  } catch (err) {
    Logger.log('Error convertirSoporteEnAnticipo: ' + err);
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================================
/** FUNCIONES COBRANZAS - PAGOS */
// ==========================================================
var CFG_CXC = {
  HOJA_COBROS_PAGOS: 'CobrosPagos',
  HOJA_MOV_BANCO:    'MovBancoCaja',
  TOLERANCIA:        0.01,
  TIPO_MOV_ENTRADA:  'Entrada',   // ajusta a tu nomenclatura si difiere
  TIPO_MOV_SALIDA:   'Salida',
  /* Si tu CUENTA_POR_PAGAR en getConstantsContables sigue leyendo P20
     (que es COSTOVENTAS), define aquí la cuenta correcta y el módulo
     la usará en su lugar. Ejemplo:
     OVERRIDE_CUENTA_POR_PAGAR: { COD: '210101001', NOM: 'Cuentas por Pagar Proveedores', CTA: 'Pasivo', SUBCTA: 'CUENTAS POR PAGAR' }, */
  OVERRIDE_CUENTA_POR_PAGAR: null
};

var CXC_COL_CP = {
  PROVEEDOR_CLIENTE: 0,      // Columna A
  RIF: 1,                    // Columna B
  NRO_FACTURA: 2,            // Columna C
  FECHA_FACTURA: 3,          // Columna D
  TOTAL_FACTURA: 4,          // Columna E
  SALDO_FACTURA: 5,          // Columna F
  SALDO_FACTURA_USD: 6,      // Columna G
  COBRADA_PAGADA: 7,         // Columna H
  TRANSACCION: 8,            // Columna I
  ID: 9,                     // Columna J
  ABONO: 10,                 // Columna K
  ABONO_BS: 11,              // Columna L  ← NUEVO
  MONEDA: 12,                // Columna M
  TASA: 13,                  // Columna N
  DESCRIPCION: 14,           // Columna O
  FECHA_ABONO: 15,           // Columna P
  TIPO_MOVIMIENTO: 16,       // Columna Q
  CUENTA_DESTINO: 17,        // Columna R
  NRO_CUENTA_DESTINO: 18,    // Columna S
  BANCO_DESTINO: 19,         // Columna T
  FECHA_VENCIMIENTO: 20,     // Columna U
  NRO_COMPROBANTES_RET_IVA: 21,   // Columna V
  NRO_COMPROBANTES_RET_ISLR: 22   // Columna W
};

var CXC_COL_MB = {
  ID_MOVIMIENTO: 0,          // Columna A
  FECHA_MOV: 1,              // Columna B
  TIPO_MOV: 2,               // Columna C
  NRO_CUENTA: 3,             // Columna D
  NOMBRE_BANCO: 4,           // Columna E
  CODIGO_BANCOS: 5,          // Columna F
  CUENTA: 6,                 // Columna G
  REFERENCIA_MOV: 7,         // Columna H
  DESC_MOV: 8,               // Columna I
  TIPO_TRANSACCION: 9,       // Columna J
  MONTO_MOV_USD: 10,         // Columna K
  MONEDA_MOV: 11,            // Columna L
  TASA_MOV: 12,              // Columna M
  IGTF: 13,                  // Columna N
  MONTO_MOV_BS: 14,          // Columna O
  CODIGO_CTA_CONTABLE: 15,   // Columna P
  CTA_CONTABLE: 16,          // Columna Q
  CTA_GASTO: 17,             // Columna R
  NOMBRE_CTA_MOV: 18,        // Columna S
  FECHA_REGISTRO: 19         // Columna T
};

// =========================================================================
// MÓDULO BACKEND: COBROS, PAGOS Y LIQUIDACIONES (MULTIEMPRESA)
// Archivo: CobrosPagos.gs
// =========================================================================
function _resolverSSCP_(idSpreadsheet) {
  if (idSpreadsheet && typeof _SS_ === 'function') {
    return _SS_(idSpreadsheet);
  }
  if (idSpreadsheet) {
    try { return SpreadsheetApp.openById(idSpreadsheet); } catch(e) {}
  }
  if (typeof Contexto !== 'undefined' && Contexto.getEmpresaSS) {
    var idActiva = Contexto.getEmpresaActivaId();
    if (idActiva) return Contexto.getEmpresaSS(idActiva);
  }
  if (typeof SPREADSHEET_ID !== 'undefined') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}
function getFacturasPorTransaccion(transaccion, idSpreadsheet) {
  try {
    var ss = _resolverSSCP_(idSpreadsheet);
    if (!ss) return [];
    var sheet = ss.getSheetByName('CobrosPagos');
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var esVenta = String(transaccion || '').toLowerCase().indexOf('vent') === 0 
               || String(transaccion || '').toLowerCase().indexOf('cobr') === 0;

    var facturas = [];
    var vistas = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nroFac = String(row[2] || '').trim();
      var transFila = String(row[8] || '').toLowerCase();

      var match = esVenta
        ? (transFila.indexOf('vent') === 0 || transFila.indexOf('cobr') === 0)
        : (transFila.indexOf('compr') === 0 || transFila.indexOf('pag') === 0);

      if (!match || !nroFac) continue;

      var tot = Number(row[4]) || 0;
      var abono = Number(row[10]) || 0;
      var tipoMov = String(row[16] || '').trim();

      // Fila principal
      if (tot > 0 && abono === 0 && !tipoMov) {
        var key = nroFac + '_' + transFila;
        if (vistas[key]) continue;
        vistas[key] = true;

        facturas.push({
          ProveedorCliente: row[0] || '',
          Rif: row[1] || '',
          NroFactura: nroFac,
          FechaFactura: row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          TotalFactura: tot,
          SaldoFactura: Number(row[5]) || 0,
          SaldoFacturaUSD: Number(row[6]) || 0,
          cobradaPagada: row[7] || '',
          Transaccion: row[8] || '',
          Tasa: Number(row[13]) || 1,
          Id: row[9] || ''
        });
      }
    }
    return facturas;
  } catch(e) {
    Logger.log("Error en getFacturasPorTransaccion: " + e);
    return [];
  }
}
function procesarLiquidacionCompleta(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var idSpreadsheet = p.idSpreadsheet || p.empresaId;
    var ss = _resolverSSCP_(idSpreadsheet);
    if (!ss) throw new Error("No se pudo conectar a la empresa activa.");

    var sheetCP = ss.getSheetByName("CobrosPagos");
    var sheetMBC = ss.getSheetByName("MovBancoCaja");
    if (!sheetCP) throw new Error("No se encontró la hoja CobrosPagos.");

    var ahora = new Date();
    var fechaStr = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaHoraStr = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    var facturas = p.facturas || [];
    var detallesPago = p.detallesPago || [];

    // 1. Descontar saldo de cada factura en CobrosPagos
    var dataCP = sheetCP.getDataRange().getValues();
    for (var f = 0; f < facturas.length; f++) {
      var facObj = facturas[f];
      var nroFacBuscado = String(facObj.nro || facObj.nroFactura || '').trim();

      for (var r = 1; r < dataCP.length; r++) {
        var nroEnFila = String(dataCP[r][2] || '').trim();
        var totEnFila = Number(dataCP[r][4]) || 0;
        var abonoEnFila = Number(dataCP[r][10]) || 0;

        // Fila madre
        if (nroEnFila === nroFacBuscado && totEnFila > 0 && abonoEnFila === 0) {
          var saldoActualUSD = Number(dataCP[r][6]) || 0;
          var montoAbonoUSD = parseFloat(facObj.montoAplicadoUSD || facObj.abonoUSD) || 0;
          var nuevoSaldoUSD = Math.max(0, saldoActualUSD - montoAbonoUSD);
          var tasaFila = Number(dataCP[r][13]) || 1;
          var nuevoSaldoBs = nuevoSaldoUSD * tasaFila;

          sheetCP.getRange(r + 1, 6).setValue(nuevoSaldoBs);  // SaldoFactura (Bs)
          sheetCP.getRange(r + 1, 7).setValue(nuevoSaldoUSD); // SaldoFacturaUSD
          if (nuevoSaldoUSD <= 0.01) {
            sheetCP.getRange(r + 1, 8).setValue("Cobrada");   // cobradaPagada
          }
          break;
        }
      }
    }

    // 2. Registrar renglón de abono en CobrosPagos
    var lastRowCP = sheetCP.getLastRow();
    var nuevoIdCP = "ABO-" + ("000" + lastRowCP).slice(-4);
    var primerPago = detallesPago[0] || {};

    sheetCP.appendRow([
      p.clienteNombre || p.proveedorNombre || "",  // 1: ProveedorCliente
      p.rif || "",                                 // 2: Rif
      facturas.map(function(x){ return x.nro || x.nroFactura; }).join(", "), // 3: NroFactura
      fechaStr,                                    // 4: FechaFactura
      0,                                           // 5: TotalFactura
      0,                                           // 6: SaldoFactura
      0,                                           // 7: SaldoFacturaUSD
      "Cobrada",                                   // 8: cobradaPagada
      p.transaccion || "Cobro",                    // 9: Transaccion
      nuevoIdCP,                                   // 10: id
      parseFloat(p.totalAplicadoUSD) || 0,         // 11: Abono
      (parseFloat(p.totalAplicadoUSD) || 0) * (parseFloat(p.tasa) || 1), // 12: AbonoBs
      p.moneda || "USD",                           // 13: Moneda
      parseFloat(p.tasa) || 1,                     // 14: Tasa
      "Liquidación Facturas: " + facturas.map(function(x){ return x.nro; }).join(", "), // 15: Descripcion
      fechaStr,                                    // 16: fechaAbono
      primerPago.formaPago || "TRANSFERENCIA",     // 17: TipoMovimiento
      primerPago.cuenta || "",                     // 18: CuentaDestino
      primerPago.nroCuenta || "",                  // 19: NroCuentaDestino
      primerPago.banco || "",                      // 20: BancoDestino
      "", "", ""
    ]);

    // 3. Registrar en MovBancoCaja si aplica
    if (sheetMBC && detallesPago.length > 0) {
      for (var d = 0; d < detallesPago.length; d++) {
        var dp = detallesPago[d];
        var lastMBC = sheetMBC.getLastRow();
        var idMov = "MOV-" + ("0000" + lastMBC).slice(-5);
        var mUSD = parseFloat(dp.montoUSD) || 0;
        var tMov = parseFloat(p.tasa) || 1;

        sheetMBC.appendRow([
          idMov,
          ahora,
          dp.tipo || "Banco",
          dp.nroCuenta || "",
          dp.banco || "",
          "",
          dp.cuenta || "Principal",
          dp.referencia || nuevoIdCP,
          "Cobranza Facturas: " + facturas.map(function(x){ return x.nro; }).join(", "),
          "Transacción Positiva",
          mUSD,
          "USD",
          tMov,
          0,
          mUSD * tMov,
          "", "Cobranzas", "", dp.banco || "",
          fechaHoraStr
        ]);
      }
    }

    return {
      success: true,
      mensaje: "Liquidación " + nuevoIdCP + " procesada con éxito en la empresa activa."
    };

  } catch(err) {
    Logger.log("Error en procesarLiquidacionCompleta: " + err);
    return { success: false, mensaje: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
function generarCorrelativoAsiento(skipLock) {
  var lock = null;
  if (!skipLock) {
    try {
      lock = LockService.getScriptLock();
      lock.waitLock(15000);
    } catch (e) {
      // Si ya está tomado por el caller (que tiene el lock superior),
      // seguimos sin lock porque es seguro.
      Logger.log('[ASIENTO] Lock ya tomado por caller, siguiendo sin lock: ' + e.message);
      lock = null;
    }
  }
  try {
    var sheet = getSheetByName('Asientos');
    if (!sheet) return "00-001";

    var ultimaFila = sheet.getLastRow();
    if (ultimaFila < 2) return "00-001";

    var valores = sheet.getRange(2, 1, ultimaFila - 1, 1).getValues().flat().filter(String);
    if (valores.length === 0) return "00-001";

    var ultimoCorrelativo = String(valores[valores.length - 1]);
    var partes = ultimoCorrelativo.split("-");
    var numero = parseInt(partes[1], 10);
    if (isNaN(numero) || !isFinite(numero)) numero = 0;

    return "00-" + String(numero + 1).padStart(3, '0');
  } finally {
    if (lock) lock.releaseLock();
  }
}
function registrarAnticipo(data) {
  try {
    if (!data) return { success: false, message: 'Datos vacíos' };
    if (!data.MontoMovUSD || data.MontoMovUSD <= 0) return { success: false, message: 'Monto inválido' };
    if (!data.ReferenciaMov) return { success: false, message: 'Referencia obligatoria' };

    // 1. Registrar en MovBancoCaja (afecta saldo)
    var resultMov = agregarEditarMovBancoCaja(data);
    if (!resultMov.success) return resultMov;
    var idAnticipo = resultMov.id;

    // 2. Registrar en CobrosPagos (22 columnas, A-V)
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var hojaCP = ss.getSheetByName('CobrosPagos');
      if (hojaCP) {
        var fechaAbono = data.FechaMov;
        if (typeof fechaAbono === 'string') {
          var p = fechaAbono.split('-');
          if (p.length === 3) fechaAbono = new Date(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10));
        }

        var esCobro = String(data.CuentaContable || '').toLowerCase().indexOf('anticipo de cliente') !== -1;
        var fila = [];
        fila[CP_COLS.PROVEEDOR_CLIENTE - 1] = data.NombreCuentaMov || '';
        fila[CP_COLS.RIF - 1] = '';
        fila[CP_COLS.NRO_FACTURA - 1] = '';
        fila[CP_COLS.FECHA_FACTURA - 1] = fechaAbono;
        fila[CP_COLS.TOTAL_FACTURA - 1] = data.MontoMovUSD || 0;
        fila[CP_COLS.SALDO_FACTURA - 1] = data.MontoMovBs || 0;
        fila[CP_COLS.SALDO_FACTURA_USD - 1] = data.MontoMovUSD || 0;
        fila[CP_COLS.COBRADA_PAGADA - 1] = 'No';
        fila[CP_COLS.TRANSACCION - 1] = 'Anticipo';
        fila[CP_COLS.ID - 1] = idAnticipo;
        fila[CP_COLS.ABONO - 1] = data.MontoMovBs || 0;
        fila[CP_COLS.MONEDA - 1] = data.MonedaMov || '';
        fila[CP_COLS.TASA - 1] = data.TasaMov || 1;
        fila[CP_COLS.DESCRIPCION - 1] = data.descMov || '';
        fila[CP_COLS.FECHA_ABONO - 1] = fechaAbono;
        fila[CP_COLS.TIPO_MOVIMIENTO - 1] = esCobro ? 'Cobro' : 'Pago';
        fila[CP_COLS.CUENTA_DESTINO - 1] = data.Cuenta || '';
        fila[CP_COLS.NRO_CUENTA_DESTINO - 1] = data.NroCuenta || '';
        fila[CP_COLS.BANCO_DESTINO - 1] = data.NombreBanco || '';
        fila[CP_COLS.FECHA_VENCIMIENTO - 1] = '';
        fila[CP_COLS.NRO_COMP_RET_IVA - 1] = '';
        fila[CP_COLS.NRO_COMP_RET_ISLR - 1] = '';
        hojaCP.appendRow(fila);
      } else {
        Logger.log('Hoja CobrosPagos no encontrada, anticipo solo en MovBancoCaja');
      }
    } catch (eCP) {
      Logger.log('Error en CobrosPagos: ' + eCP);
    }

    return { success: true, message: 'Anticipo registrado correctamente', id: idAnticipo };
  } catch (e) {
    Logger.log('Error registrarAnticipo: ' + e);
    return { success: false, message: 'Error: ' + (e.message || String(e)) };
  }
}
function registrarLiquidacionCxCp(payload, esCobro) {
  if (!payload || !payload.facturas || !Array.isArray(payload.facturas)) {
    throw new Error('Payload inválido para la operación.');
  }

  var info = payload.infoPago || {};
  var tasa = parseNumeroLatino(info.tasa);
  if (!isFinite(tasa) || tasa <= 0) throw new Error('Tasa de cambio inválida: ' + info.tasa);

  var fechaOp = info.fecha || formatDate(new Date());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shCP = ss.getSheetByName(CFG_CXC.HOJA_COBROS_PAGOS);
  var shMB = ss.getSheetByName(CFG_CXC.HOJA_MOV_BANCO);
  var shAsientos = ss.getSheetByName('Asientos');

  if (!shCP) throw new Error("No se encontró la hoja CobrosPagos.");
  if (!shAsientos) throw new Error("No se encontró la hoja Asientos.");

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var prefijo = esCobro ? 'C' : 'P';
    var nroCorrelativo = obtenerSiguienteCorrelativoCobrosPagos(prefijo, true);

    var dataCP = shCP.getDataRange().getValues();
    var hdCP = dataCP[0];
    var cCP = {
      nombre:   hdCP.indexOf('ProveedorCliente'),
      rif:      hdCP.indexOf('Rif'),
      factura:  hdCP.indexOf('NroFactura'),
      fechaFac: hdCP.indexOf('FechaFactura'),
      total:    hdCP.indexOf('TotalFactura'),
      saldoBs:  hdCP.indexOf('SaldoFactura'),
      saldoUsd: hdCP.indexOf('SaldoFacturaUSD'),
      status:   hdCP.indexOf('cobradaPagada'),
      abono:    hdCP.indexOf('Abono'),
      abonoBs:  hdCP.indexOf('AbonoBs'),
      id:       hdCP.indexOf('id'),
      fVenc:    hdCP.indexOf('fechaVencimiento')
    };

    var anticiposCruzados = info.anticiposCruzados || [];
    var retencionesAplicadas = info.retencionesAplicadas || [];
    var totalCruzadoUSD = 0;
    anticiposCruzados.forEach(function(a) { totalCruzadoUSD += Number(a.montoCruzadoUSD) || 0; });

    var updatesEmision = [];
    var filasTx = [];
    var excedenteAnticipoTotal = 0;
    var facturasPayload = payload.facturas;

    var detallesPago = (info.detallesPago || []).filter(function(d) { return Number(d.monto) > 0; });
    var ctaDestinoFila = { nombre: '', codigo: '', banco: '', codContable: '', nomContable: '' };
    if (detallesPago.length > 0) {
      var ctaParts = String(detallesPago[0].cuenta || '').split('|');
      ctaDestinoFila = {
        nombre: String(ctaParts[1] || '').trim(),
        codigo: String(ctaParts[2] || '').trim(),
        banco: String(ctaParts[1] || '').trim(),
        codContable: String(ctaParts[3] || '').trim(),
        nomContable: String(ctaParts[4] || '').trim()
      };
    }

    facturasPayload.forEach(function(f) {
      var facturaNum = String(f.NroFacturaCliente || f.NroFacturaProveedor || f.Id || '').trim();
      var montoAnticipo = Number(f.montoAnticipoAplicadoUSD) || 0;
      var montoRetencion = Number(f.montoRetencionAplicadaUSD) || 0;
      var montoBanco = Number(f.montoAbonoBancoUSD) || 0;
      var totalAplicado = Number(f.montoTotalAplicadoUSD) || (montoAnticipo + montoRetencion + montoBanco);
      var abonadoBs = Math.round((totalAplicado * tasa) * 100) / 100;

      var facturaEncontrada = false;
      for (var r = 1; r < dataCP.length; r++) {
        var facFila = String(dataCP[r][cCP.factura]).trim();
        var totFila = Number(dataCP[r][cCP.total]) || 0;

        if (facFila === facturaNum && totFila > 0) {
          facturaEncontrada = true;
          var sBs = Math.round((Number(dataCP[r][cCP.saldoBs]) || 0) * 100) / 100;
          var sUsd = Math.round((Number(dataCP[r][cCP.saldoUsd]) || 0) * 100) / 100;

          if (abonadoBs > sBs + 0.01) {
            throw new Error('El abono (Bs ' + abonadoBs.toFixed(2) + ') excede el saldo de la factura ' + facturaNum + ' (Bs ' + sBs.toFixed(2) + ').');
          }
          if (totalAplicado > sUsd + 0.01) {
            throw new Error('El abono ($' + totalAplicado.toFixed(2) + ') excede el saldo USD de la factura ' + facturaNum + ' ($' + sUsd.toFixed(2) + ').');
          }

          var excedente = Math.max(0, totalAplicado - sUsd);
          if (excedente > 0.01) {
            excedenteAnticipoTotal += excedente;
          }

          var restBs = Math.max(0, Math.round((sBs - abonadoBs) * 100) / 100);
          var restUsd = Math.max(0, Math.round((sUsd - totalAplicado) * 100) / 100);
          var st = (restBs <= 0.01) ? (esCobro ? 'Cobrada' : 'Pagada') : (esCobro ? 'Abono - Por Cobrar' : 'Abono - Por Pagar');

          updatesEmision.push({ fila: r + 1, restBs: restBs, restUsd: restUsd, status: st });

          var filaNueva = new Array(23).fill('');
          filaNueva[CXC_COL_CP.PROVEEDOR_CLIENTE] = dataCP[r][cCP.nombre];
          filaNueva[CXC_COL_CP.RIF]               = dataCP[r][cCP.rif];
          filaNueva[CXC_COL_CP.NRO_FACTURA]       = facturaNum;
          filaNueva[CXC_COL_CP.FECHA_FACTURA]     = dataCP[r][cCP.fechaFac];
          filaNueva[CXC_COL_CP.TOTAL_FACTURA]     = 0;
          filaNueva[CXC_COL_CP.SALDO_FACTURA]     = restBs;
          filaNueva[CXC_COL_CP.SALDO_FACTURA_USD] = restUsd;
          filaNueva[CXC_COL_CP.COBRADA_PAGADA]    = st;
          filaNueva[CXC_COL_CP.TRANSACCION]       = esCobro ? 'Cobro' : 'Pago';
          filaNueva[CXC_COL_CP.ID]                = nroCorrelativo;
          filaNueva[CXC_COL_CP.ABONO]             = totalAplicado;
          filaNueva[CXC_COL_CP.ABONO_BS]          = abonadoBs;
          filaNueva[CXC_COL_CP.MONEDA]            = 'Dolares';
          filaNueva[CXC_COL_CP.TASA]              = tasa;
          filaNueva[CXC_COL_CP.DESCRIPCION]       = (esCobro ? 'Cobro' : 'Pago') + ' Factura ' + facturaNum 
                                                    + (montoAnticipo > 0 ? ' [Anticipo: $' + montoAnticipo.toFixed(2) + ']' : '')
                                                    + (montoRetencion > 0 ? ' [Retención: $' + montoRetencion.toFixed(2) + ']' : '')
                                                    + (montoBanco > 0 ? ' [Banco: $' + montoBanco.toFixed(2) + ']' : '');
          filaNueva[CXC_COL_CP.FECHA_ABONO]       = new Date(fechaOp + 'T00:00:00');
          filaNueva[CXC_COL_CP.TIPO_MOVIMIENTO]   = esCobro ? 'CXC' : 'CXP';
          filaNueva[CXC_COL_CP.CUENTA_DESTINO]    = ctaDestinoFila.nombre;
          filaNueva[CXC_COL_CP.NRO_CUENTA_DESTINO] = ctaDestinoFila.codigo;
          filaNueva[CXC_COL_CP.BANCO_DESTINO]     = ctaDestinoFila.banco;
          filaNueva[CXC_COL_CP.FECHA_VENCIMIENTO] = dataCP[r][cCP.fVenc];

          if (f.retencionInfo) {
            var comp = String(f.retencionInfo.comprobante || '').trim();
            if (comp) {
              if (String(f.retencionInfo.tipoRetencion).toUpperCase() === 'IVA') {
                filaNueva[CXC_COL_CP.NRO_COMPROBANTES_RET_IVA] = comp;
              } else {
                filaNueva[CXC_COL_CP.NRO_COMPROBANTES_RET_ISLR] = comp;
              }
            }
          }

          filasTx.push(filaNueva);
          break;
        }
      }
      if (!facturaEncontrada) {
        throw new Error('Factura no encontrada en CobrosPagos: ' + facturaNum);
      }
    });

    anticiposCruzados.forEach(function(ant) {
      var idAnt = String(ant.idAnticipo || ant.referencia).trim().toLowerCase();
      for (var rA = 1; rA < dataCP.length; rA++) {
        var idFila = String(dataCP[rA][cCP.id]).trim().toLowerCase();
        var facFila = String(dataCP[rA][cCP.factura]).trim().toLowerCase();

        if (idFila === idAnt || facFila === idAnt) {
          var sRestUSD = Math.max(0, Math.round((Number(ant.saldoRestanteUSD)) * 100) / 100);
          var sRestBs = Math.max(0, Math.round((sRestUSD * tasa) * 100) / 100);
          var stAnt = (sRestUSD <= 0.01) ? 'Anticipo' : 'Abono - Parcial';

          updatesEmision.push({ fila: rA + 1, restBs: sRestBs, restUsd: sRestUSD, status: stAnt });
          break;
        }
      }
    });

    retencionesAplicadas.forEach(function(ret) {
      var idRet = String(ret.idRetencion || ret.comprobante).trim().toLowerCase();
      for (var rR = 1; rR < dataCP.length; rR++) {
        var idFila = String(dataCP[rR][cCP.id]).trim().toLowerCase();
        var compFila = String(dataCP[rR][CXC_COL_CP.NRO_COMPROBANTES_RET_IVA] || '').trim().toLowerCase();
        var compFila2 = String(dataCP[rR][CXC_COL_CP.NRO_COMPROBANTES_RET_ISLR] || '').trim().toLowerCase();

        if (idFila === idRet || compFila === idRet || compFila2 === idRet) {
          var sRestUSD = Math.max(0, Math.round(((Number(dataCP[rR][cCP.saldoUsd]) || 0) - (Number(ret.montoAplicadoUSD) || 0)) * 100) / 100);
          var sRestBs = Math.max(0, Math.round(((Number(dataCP[rR][cCP.saldoBs]) || 0) - (Number(ret.montoAplicadoBs) || 0)) * 100) / 100);
          var stRet = (sRestUSD <= 0.01) ? 'Anticipo' : 'Abono - Parcial';

          updatesEmision.push({ fila: rR + 1, restBs: sRestBs, restUsd: sRestUSD, status: stRet });
          break;
        }
      }
    });

    updatesEmision.forEach(function(u) {
      shCP.getRange(u.fila, cCP.saldoBs + 1).setValue(u.restBs);
      shCP.getRange(u.fila, cCP.saldoUsd + 1).setValue(u.restUsd);
      shCP.getRange(u.fila, cCP.status + 1).setValue(u.status);
    });
    if (filasTx.length > 0) {
      shCP.getRange(shCP.getLastRow() + 1, 1, filasTx.length, 23).setValues(filasTx);
    }

    try {
      var hojaFuente = ss.getSheetByName(esCobro ? 'Ventas' : 'Compras');
      if (hojaFuente) {
        var dataFuente = hojaFuente.getDataRange().getValues();
        var hFuente = dataFuente[0].map(function(h) {
          return String(h || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-#.]/g, '');
        });
        var colFactF = esCobro ? hFuente.indexOf('nrofacturacliente') : hFuente.indexOf('nrofacturaproveedor');
        var colSaldoBsF = esCobro ? hFuente.indexOf('saldocliente') : hFuente.indexOf('saldoproveedor');
        var colSaldoUsdF = esCobro ? hFuente.indexOf('saldoclienteusd') : hFuente.indexOf('saldoproveedordivisa');
        var colStatusF = hFuente.indexOf('cobradapagada');
        var colAbonoF = esCobro ? hFuente.indexOf('abonocliente') : hFuente.indexOf('abono');
        var colFechaAbonoF = esCobro ? hFuente.indexOf('fechaabonocliente') : hFuente.indexOf('fechaabono');
        var colStatusGlobalF = hFuente.indexOf('status');

        if (colFactF !== -1 && colSaldoBsF !== -1) {
          for (var p = 0; p < facturasPayload.length; p++) {
            var fp = facturasPayload[p];
            var nroF = String(fp.NroFacturaCliente || fp.NroFacturaProveedor || '').trim().toLowerCase();
            for (var q = 1; q < dataFuente.length; q++) {
              if (String(dataFuente[q][colFactF] || '').trim().toLowerCase() === nroF) {
                var nuevoBs = Math.max(0, Math.round(((Number(dataFuente[q][colSaldoBsF]) || 0) - (fp.montoTotalAplicadoBs || 0)) * 100) / 100);
                hojaFuente.getRange(q + 1, colSaldoBsF + 1).setValue(nuevoBs);
                if (colSaldoUsdF !== -1) {
                  var nuevoUsd = Math.max(0, Math.round(((Number(dataFuente[q][colSaldoUsdF]) || 0) - (fp.montoTotalAplicadoUSD || 0)) * 100) / 100);
                  hojaFuente.getRange(q + 1, colSaldoUsdF + 1).setValue(nuevoUsd);
                }
                if (colStatusF !== -1) {
                  hojaFuente.getRange(q + 1, colStatusF + 1).setValue(nuevoBs <= 0.01 ? (esCobro ? 'Cobrada' : 'Pagada') : (esCobro ? 'Abono - Por Cobrar' : 'Abono - Por Pagar'));
                }
                if (colStatusGlobalF !== -1) {
                  hojaFuente.getRange(q + 1, colStatusGlobalF + 1).setValue(nuevoBs <= 0.01 ? (esCobro ? 'Cobrada' : 'Pagada') : 'Activa');
                }
                if (colAbonoF !== -1) {
                  var abonoActual = Math.round(((Number(dataFuente[q][colAbonoF]) || 0) + (fp.montoTotalAplicadoBs || 0)) * 100) / 100;
                  hojaFuente.getRange(q + 1, colAbonoF + 1).setValue(abonoActual);
                }
                if (colFechaAbonoF !== -1) {
                  hojaFuente.getRange(q + 1, colFechaAbonoF + 1).setValue(new Date(fechaOp + 'T00:00:00'));
                }
                break;
              }
            }
          }
        }
      }
    } catch (eFuente) {
      Logger.log('[CXC] Error actualizando Ventas/Compras: ' + eFuente.message);
    }

    if (excedenteAnticipoTotal > 0.01) {
      try {
        var correlativoNuevoAnt = obtenerSiguienteCorrelativoCobrosPagos('A', true);
        var filaExcedente = new Array(23).fill('');
        filaExcedente[CXC_COL_CP.PROVEEDOR_CLIENTE] = facturasPayload[0].NombreCliente || facturasPayload[0].NombreProveedor || '';
        filaExcedente[CXC_COL_CP.RIF]               = facturasPayload[0].RifCliente || facturasPayload[0].RifProveedor || '';
        filaExcedente[CXC_COL_CP.NRO_FACTURA]       = correlativoNuevoAnt;
        filaExcedente[CXC_COL_CP.FECHA_FACTURA]     = new Date(fechaOp + 'T00:00:00');
        filaExcedente[CXC_COL_CP.TOTAL_FACTURA]     = 0;
        filaExcedente[CXC_COL_CP.SALDO_FACTURA]     = Math.round((excedenteAnticipoTotal * tasa) * 100) / 100;
        filaExcedente[CXC_COL_CP.SALDO_FACTURA_USD] = Math.round(excedenteAnticipoTotal * 100) / 100;
        filaExcedente[CXC_COL_CP.COBRADA_PAGADA]    = esCobro ? 'anticipoCobro' : 'anticipoPago';
        filaExcedente[CXC_COL_CP.TRANSACCION]       = 'Anticipo';
        filaExcedente[CXC_COL_CP.ID]                = correlativoNuevoAnt;
        filaExcedente[CXC_COL_CP.ABONO]             = 0;
        filaExcedente[CXC_COL_CP.ABONO_BS]          = 0;
        filaExcedente[CXC_COL_CP.MONEDA]            = 'Dolares';
        filaExcedente[CXC_COL_CP.TASA]              = tasa;
        filaExcedente[CXC_COL_CP.DESCRIPCION]       = 'Excedente de anticipo generado por sobrepago en factura ' + nroCorrelativo;
        filaExcedente[CXC_COL_CP.FECHA_ABONO]       = new Date(fechaOp + 'T00:00:00');
        filaExcedente[CXC_COL_CP.TIPO_MOVIMIENTO]   = 'ANTICIPO';
        filaExcedente[CXC_COL_CP.FECHA_VENCIMIENTO] = new Date(fechaOp + 'T00:00:00');

        shCP.appendRow(filaExcedente);
      } catch (eExc) {
        Logger.log('[CXC] Error creando anticipo por excedente: ' + eExc.message);
      }
    }

    if (detallesPago.length > 0 && shMB) {
      var nombreTerceroMB = (facturasPayload[0] && (facturasPayload[0].NombreCliente || facturasPayload[0].NombreProveedor)) || '';
      detallesPago.forEach(function(d) {
        var ctaParts = String(d.cuenta || '').split('|');
        var ctaTipo = String(ctaParts[0] || 'Banco').trim();
        var ctaNombre = String(ctaParts[1] || 'Caja/Banco').trim();
        var ctaCod = String(ctaParts[2] || '').trim();
        var ctaCodContable = String(ctaParts[3] || '').trim();
        var ctaNomContable = String(ctaParts[4] || '').trim();
        
        var mUSD = (d.moneda === 'USD') ? Math.round(Number(d.monto) * 100) / 100 : Math.round((Number(d.monto) / tasa) * 100) / 100;
        var mBs = (d.moneda === 'USD') ? Math.round((Number(d.monto) * tasa) * 100) / 100 : Math.round(Number(d.monto) * 100) / 100;

        var filaMB = new Array(20).fill('');
        filaMB[CXC_COL_MB.ID_MOVIMIENTO]    = "MOV-" + nroCorrelativo;
        filaMB[CXC_COL_MB.FECHA_MOV]        = new Date(fechaOp + 'T00:00:00');
        filaMB[CXC_COL_MB.TIPO_MOV]         = ctaTipo;
        filaMB[CXC_COL_MB.NRO_CUENTA]       = ctaCod;
        filaMB[CXC_COL_MB.NOMBRE_BANCO]     = ctaNombre;
        filaMB[CXC_COL_MB.CODIGO_BANCOS]    = ctaCod;
        filaMB[CXC_COL_MB.CUENTA]           = ctaNombre;
        filaMB[CXC_COL_MB.REFERENCIA_MOV]   = d.referencia || 'S/R';
        filaMB[CXC_COL_MB.DESC_MOV]         = (esCobro ? 'Cobro' : 'Pago') + ' Facturas: ' + facturasPayload.map(function(x){return x.NroFacturaCliente||x.NroFacturaProveedor;}).join(', ');
        filaMB[CXC_COL_MB.TIPO_TRANSACCION] = esCobro ? 'Entrada' : 'Salida';
        filaMB[CXC_COL_MB.MONTO_MOV_USD]    = mUSD;
        filaMB[CXC_COL_MB.MONEDA_MOV]       = (d.moneda === 'USD') ? 'Dolares' : 'Bolivares';
        filaMB[CXC_COL_MB.TASA_MOV]         = tasa;
        filaMB[CXC_COL_MB.IGTF]             = 0;
        filaMB[CXC_COL_MB.MONTO_MOV_BS]     = mBs;
        filaMB[CXC_COL_MB.CODIGO_CTA_CONTABLE] = ctaCodContable;
        filaMB[CXC_COL_MB.CTA_CONTABLE]        = ctaNomContable;
        filaMB[CXC_COL_MB.NOMBRE_CTA_MOV]   = nombreTerceroMB;
        filaMB[CXC_COL_MB.FECHA_REGISTRO]   = new Date();

        shMB.appendRow(filaMB);
      });
    }

    var nroAsientoContable = null;
    try {
      var totalFacturasBs = 0;
      facturasPayload.forEach(function(f) {
        var m = Number(f.montoTotalAplicadoBs) || Math.round((Number(f.montoTotalAplicadoUSD) * tasa) * 100) / 100;
        totalFacturasBs += m;
      });
      totalFacturasBs = Math.round(totalFacturasBs * 100) / 100;

      if (totalFacturasBs > 0) {
        nroAsientoContable = generarCorrelativoAsiento(true);

        var ctaBanco;
        if (ctaDestinoFila.codContable) {
          ctaBanco = { 
            codContable: ctaDestinoFila.codContable, 
            nomContable: ctaDestinoFila.nomContable || ctaDestinoFila.nombre || 'Caja/Banco' 
          };
        } else {
          var ctaContable = obtenerDatosContablesCuenta(detallesPago.length > 0 ? detallesPago[0].cuenta : '');
          if (!ctaContable.cod || ctaContable.cod === '1.1.01.01') {
            throw new Error('No se pudo determinar el código contable de la cuenta "' + ctaDestinoFila.nombre + '". Verifique datosBancarios.');
          }
          ctaBanco = { codContable: ctaContable.cod, nomContable: ctaContable.nom };
        }

        var contra = resolverContrapartidaCxCp(esCobro);
        var glosa = (esCobro ? 'Cobro' : 'Pago') + ' consolidado ' + nroCorrelativo;
        var refFacturas = facturasPayload.map(function(x){return x.NroFacturaCliente||x.NroFacturaProveedor;}).join(', ');
        var tercero = (facturasPayload[0] && (facturasPayload[0].NombreCliente || facturasPayload[0].NombreProveedor)) || '';

        var lineasAsiento = construirAsientoLiquidacion(
          esCobro, nroAsientoContable, fechaOp, tercero, glosa, refFacturas, totalFacturasBs, ctaBanco, contra
        );

        guardarAsientoMasivo(lineasAsiento);
      }
    } catch (eAsiento) {
      Logger.log('[CXC] Error generando asiento: ' + eAsiento.message);
    }

    try {
      if (detallesPago.length > 0) {
        actualizarSaldosBancariosOptimizado();
      }
    } catch (eSaldos) {
      Logger.log('[CXC] Error actualizando saldos: ' + eSaldos.message);
    }

    return { 
      ok: true, 
      asiento: nroCorrelativo, 
      nroAsientoContable: nroAsientoContable,
      id: nroCorrelativo,
      excedenteAnticipo: excedenteAnticipoTotal > 0.01 ? excedenteAnticipoTotal : 0
    };

  } finally {
    lock.releaseLock();
  }
}
function obtenerNroControlReal(ss, nroFactura, esVenta) {
    try {
        var hojaFuente = ss.getSheetByName(esVenta ? 'Ventas' : 'Compras');
        if (!hojaFuente) return '00-00000000';
        
        var datos = hojaFuente.getDataRange().getValues();
        if (datos.length < 2) return '00-00000000';
        
        var headers = datos[0].map(function(h) {
            return String(h).trim().toLowerCase();
        });
        
        var cFact = esVenta
            ? headers.indexOf('nrofacturacliente')
            : headers.indexOf('nrofacturaproveedor');
        
        var cCtrl = esVenta
            ? headers.indexOf('nrocontrol')
            : headers.indexOf('nrocontrolproveedor');
        
        if (cFact === -1 || cCtrl === -1) {
            Logger.log('[RET] Columnas no encontradas. Fact=' + cFact + ' Ctrl=' + cCtrl);
            return '00-00000000';
        }
        
        var nroBuscado = String(nroFactura).trim().toLowerCase();
        for (var i = 1; i < datos.length; i++) {
            if (String(datos[i][cFact]).trim().toLowerCase() === nroBuscado) {
                var control = String(datos[i][cCtrl] || '').trim();
                return control || '00-00000000';
            }
        }
    } catch (e) {
        Logger.log('[RET] Error buscando nroControl: ' + e.message);
    }
    return '00-00000000';
}
function registrarPago(datosPago) {
  // Redirige al flujo moderno (evita contaminar CobrosPagos con 7 columnas)
  if (!datosPago || !datosPago.infoPago) {
    throw new Error('registrarPago legacy: payload inválido.');
  }
  // Adaptar formato antiguo → moderno
  var payload = {
    infoPago: datosPago.infoPago || {},
    facturas: datosPago.facturas || []
  };
  return registrarLiquidacionCxCp(payload, false);
}
function inicioDeDia(fecha) {
  if (!fecha || !(fecha instanceof Date)) {
    fecha = new Date();
  }
  var d = new Date(fecha);
  d.setHours(0, 0, 0, 0);
  return d;
}
function sumarDias(fecha, dias) {
  if (!fecha || !(fecha instanceof Date)) {
    fecha = new Date();
  }
  var d = new Date(fecha);
  d.setDate(d.getDate() + dias);
  return d;
}
function formatearFecha(fecha) {
  if (!fecha || !(fecha instanceof Date)) {
    return '';
  }
  var dia = String(fecha.getDate()).padStart(2, '0');
  var mes = String(fecha.getMonth() + 1).padStart(2, '0');
  var año = fecha.getFullYear();
  return dia + '/' + mes + '/' + año;
}
function formatDate(fecha) {
  if (!fecha || !(fecha instanceof Date)) {
    fecha = new Date();
  }
  var dia = String(fecha.getDate()).padStart(2, '0');
  var mes = String(fecha.getMonth() + 1).padStart(2, '0');
  var año = fecha.getFullYear();
  return año + '-' + mes + '-' + dia;
}
function limpiarTexto(texto) {
  if (!texto) {
    return '';
  }
  return String(texto)
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}
function clasificarTipoDocumento(status) {
  if (!status) {
    return 'OTRO';
  }
  var statusLower = String(status).toLowerCase();
  if (statusLower.indexOf('cobr') !== -1 || statusLower.indexOf('venta') !== -1) {
    return 'CXC';
  }
  if (statusLower.indexOf('pag') !== -1 || statusLower.indexOf('compr') !== -1) {
    return 'CXP';
  }
  if (statusLower.indexOf('anticip') !== -1) {
    return 'ANTICIPO';
  }
  if (statusLower.indexOf('ret') !== -1) {
    return 'RETENCION';
  }
  return 'OTRO';
}
function registrarCobranza(payload) {return registrarLiquidacionCxCp(payload, true);}
function registrarPagoCompra(payload) {return registrarLiquidacionCxCp(payload, false);}
function registrarAnticipo(anticipoData) {
  try {
    const hojaMovimientos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MovBancoCaja");
    if (!hojaMovimientos) throw new Error("La hoja 'MovBancoCaja' no fue encontrada.");

    const esCobro = anticipoData.TipoMov === 'Cobro' || anticipoData.tipo === 'Cobro'; // Soporta ambos formatos
    const tipoTransaccion = esCobro ? 'Transf Positiva(+)' : 'Transf Negativa(-)';
    
    // ✅ CORRECCIÓN 1: Evitar "undefined. undefined" usando valores por defecto
    const clienteProv = anticipoData.NombreCuentaMov || anticipoData.clienteProveedor || 'N/D';
    const concepto = anticipoData.descMov || anticipoData.concepto || 'Anticipo';
    const descripcionMovimiento = `Anticipo ${esCobro ? 'de Cliente' : 'a Proveedor'}: ${clienteProv}. ${concepto}`;
    
    // ✅ CORRECCIÓN 2: Usar los campos reales que envía el frontend
    const codigoCuenta = anticipoData.CodigoBancos || anticipoData.Cuenta || '';
    const nombreCuenta = anticipoData.NombreBanco || '';
    const nroCuentaReal = anticipoData.NroCuenta || ''; 

    const idMov = "MOV-" + Utilities.getUuid().substring(0, 8);
    
    const tasa = parseFloat(anticipoData.TasaMov || anticipoData.tasa) || 1;
    const montoOriginal = parseFloat(anticipoData.MontoMovUSD || anticipoData.monto) || 0;
    const montoEquivalente = parseFloat(anticipoData.MontoMovBs || anticipoData.montoEquivalente) || (montoOriginal * tasa);
    const moneda = anticipoData.MonedaMov || anticipoData.moneda || 'Dolar';
    
    const montoUSD = moneda.toLowerCase().includes('dolar') ? montoOriginal : (montoEquivalente / tasa);

    // ✅ CORRECCIÓN 3: Validación segura de fecha para evitar 1969
    let fechaRegistro = new Date();
    if (anticipoData.FechaMov || anticipoData.fecha) {
      const fechaTemp = new Date(anticipoData.FechaMov || anticipoData.fecha);
      if (!isNaN(fechaTemp.getTime())) fechaRegistro = fechaTemp;
    }

    const nuevaFila = [
      idMov,                                    // 1. IdMovimiento
      fechaRegistro,                            // 2. FechaMov (Objeto Date válido)
      anticipoData.TipoMov || anticipoData.origenDinero || "Banco", // 3. TipoMov
      nroCuentaReal,                            // 4. NroCuenta (El número real, no el string "COD - NOMBRE")
      nombreCuenta,                             // 5. NombreBanco
      codigoCuenta,                             // 6. CodigoBancos
      codigoCuenta,                             // 7. Cuenta
      anticipoData.ReferenciaMov || anticipoData.referencia || '', // 8. ReferenciaMov
      descripcionMovimiento,                    // 9. descMov
      tipoTransaccion,                          // 10. TipoTransaccion
      montoUSD,                                 // 11. MontoMovUSD
      moneda,                                   // 12. MonedaMov
      tasa,                                     // 13. TasaMov
      parseFloat(anticipoData.IGTF) || 0,       // 14. IGTF
      montoEquivalente,                         // 15. MontoMovBs
      '',                                       // 16. CodigoCuentaContable
      '',                                       // 17. CuentaContable
      '',                                       // 18. CuentaGasto
      '',                                       // 19. NombreCuentaMov
      new Date()                                // 20. fechaRegistroMov
    ];

    hojaMovimientos.appendRow(nuevaFila);

    return { success: true, message: "Anticipo registrado exitosamente.", id: idMov };
  } catch (error) {
    Logger.log('Error en registrarAnticipo: ' + error.toString());
    return { success: false, message: 'Error en el servidor: ' + error.message };
  }
}
function obtenerDatosInforme(){try{const ss=SpreadsheetApp.getActiveSpreadsheet();const hojaMovBancos=ss.getSheetByName("MovBancoCaja");if(!hojaMovBancos){return{saldoInicial:0,totalIngresos:0,totalEgresos:0,saldoFinal:0,mapaIngresos:{},mapaEgresos:{comisionesComTrf:{monto:0,transacciones:[]},otros:{}}}}const data=hojaMovBancos.getDataRange().getValues();const headers=data[0].map(h=>String(h).toLowerCase().trim());const colFecha=headers.findIndex(h=>h.includes("fecha"));const colTipo=headers.findIndex(h=>h.includes("tipo")&&h.includes("mov"));const colMonto=headers.findIndex(h=>h.includes("monto")&&h.includes("bs"));const colDesc=headers.findIndex(h=>h.includes("desc"));const colRef=headers.findIndex(h=>h.includes("referencia"));let saldoInicial=0,totalIngresos=0,totalEgresos=0;const mapaIngresos={},mapaEgresos={comisionesComTrf:{monto:0,transacciones:[],descripcion:"Comisiones Transferencias"},otros:{}};for(let i=1;i<data.length;i++){const row=data[i];const fecha=row[colFecha]?new Date(row[colFecha]).toLocaleDateString('es-VE'):'';const tipo=String(row[colTipo]||'').toLowerCase();const monto=parseNumeroLatino(row[colMonto])||0;const desc=String(row[colDesc]||'Movimiento');const ref=String(row[colRef]||'');if(tipo.includes('entrada')||tipo.includes('ingreso')){totalIngresos+=monto;if(!mapaIngresos[fecha])mapaIngresos[fecha]={};if(!mapaIngresos[fecha][desc])mapaIngresos[fecha][desc]={monto:0,items:[]};mapaIngresos[fecha][desc].monto+=monto;mapaIngresos[fecha][desc].items.push({monto:monto,ref:ref})}else if(tipo.includes('salida')||tipo.includes('egreso')){totalEgresos+=monto;if(desc.toLowerCase().includes('comision')||desc.toLowerCase().includes('transferencia')){mapaEgresos.comisionesComTrf.monto+=monto;mapaEgresos.comisionesComTrf.transacciones.push({fecha:fecha,ref:ref,desc:desc,monto:monto})}else{if(!mapaEgresos.otros[desc])mapaEgresos.otros[desc]={monto:0,items:[]};mapaEgresos.otros[desc].monto+=monto;mapaEgresos.otros[desc].items.push({fecha:fecha,ref:ref,monto:monto})}}}const saldoFinal=saldoInicial+totalIngresos-totalEgresos;return{saldoInicial:saldoInicial,totalIngresos:totalIngresos,totalEgresos:totalEgresos,saldoFinal:saldoFinal,mapaIngresos:mapaIngresos,mapaEgresos:mapaEgresos}}catch(e){Logger.log('Error obtenerDatosInforme: '+e);return{saldoInicial:0,totalIngresos:0,totalEgresos:0,saldoFinal:0,mapaIngresos:{},mapaEgresos:{comisionesComTrf:{monto:0,transacciones:[]},otros:{}}}}}

function obtenerRelacionPagos(){try{const ss=SpreadsheetApp.getActiveSpreadsheet();const hojaMovBancos=ss.getSheetByName("MovBancoCaja");if(!hojaMovBancos)return[];const data=hojaMovBancos.getDataRange().getValues();const headers=data[0].map(h=>String(h).toLowerCase().trim());const colFecha=headers.findIndex(h=>h.includes("fecha"));const colRef=headers.findIndex(h=>h.includes("referencia"));const colDesc=headers.findIndex(h=>h.includes("desc"));const colMonto=headers.findIndex(h=>h.includes("monto")&&h.includes("bs"));const relacion=[];let balance=0;for(let i=1;i<data.length;i++){const row=data[i];const fecha=row[colFecha]?new Date(row[colFecha]).toLocaleDateString('es-VE'):'';const referencia=String(row[colRef]||'');const descripcion=String(row[colDesc]||'');const monto=parseNumeroLatino(row[colMonto])||0;balance+=monto;relacion.push({fecha:fecha,referencia:referencia,descripcion:descripcion,monto:monto,balance:balance})}return relacion}catch(e){Logger.log('Error obtenerRelacionPagos: '+e);return[]}}

function resolverContrapartidaCxCp(esCobro) {
  var constants = getConstantsContables(); // TU función
  var cta;
  if (esCobro) {
    cta = constants.CUENTA_POR_COBRAR;
    if (!cta || !cta.COD) throw new Error('CONFIGURACIÓN: CUENTA_POR_COBRAR sin código (Resumen P15).');
  } else {
    cta = constants.CUENTA_POR_PAGAR;
    if (!cta || !cta.COD) throw new Error('CONFIGURACIÓN: CUENTA_POR_PAGAR sin código en getConstantsContables.');
    if (constants.COSTOVENTAS && String(cta.COD) === String(constants.COSTOVENTAS)) {
      throw new Error('CONFIGURACIÓN BLOQUEADA: CUENTA_POR_PAGAR lee Resumen P20, que contiene COSTOVENTAS (' +
                      cta.COD + '). Corrige la celda de Cuentas por Pagar en getConstantsContables.');
    }
  }
  return {
    cod: limpiarCodigoContable(cta.COD),
    nom: String(cta.NOM || ''),
    mayor: String(cta.CTA || (esCobro ? 'Activo' : 'Pasivo')),
    subcuenta: String(cta.SUBCTA || cta.SUB || 'S/D')
  };
}
function limpiarCodigoContable(str) {
  return str ? String(str).split(' - ')[0].trim() : '';
}
function construirAsientoLiquidacion(esCobro, nroAsiento, fechaOp, tercero, glosa, refFacturas, totalBs, ctaBanco, contra) {

  var fechaAsiento = new Date(fechaOp + 'T00:00:00');
  var codBancoLimpio = limpiarCodigoContable(ctaBanco.codContable);

  // Clasificación CUENTA/SUBCUENTA del banco desde TU catálogo (ahora sí matchea)
  var catalogo = getCatalogosContables();
  var detBanco = { cuenta: 'S/D', subcuenta: ctaBanco.nomContable || 'S/D' };
  for (var k = 0; k < catalogo.length; k++) {
    if (String(catalogo[k].cod) === codBancoLimpio) {
      detBanco = { cuenta: catalogo[k].mayor || 'S/D', subcuenta: catalogo[k].nom || 'S/D' };
      break;
    }
  }

  var montoBanco  = esCobro ? totalBs : (totalBs * -1);
  var montoContra = esCobro ? (totalBs * -1) : totalBs;

  return [
    { nroAsiento: nroAsiento, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
      clienteProveedor: tercero, descripcionGeneral: glosa,
      codigoCuenta: codBancoLimpio,
      cuentaContable: ctaBanco.nomContable,
      montoTransaccion: montoBanco,
      descripcionTransaccion: esCobro ? 'Cobro Facturas' : 'Pago Facturas',
      transaccionRef: refFacturas,
      mayor: detBanco.cuenta, subcuenta: detBanco.subcuenta },

    { nroAsiento: nroAsiento, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
      clienteProveedor: tercero, descripcionGeneral: glosa,
      codigoCuenta: contra.cod, cuentaContable: contra.nom,
      montoTransaccion: montoContra,
      descripcionTransaccion: esCobro ? 'Cobro Facturas' : 'Pago Facturas',
      transaccionRef: refFacturas,
      mayor: contra.mayor, subcuenta: contra.subcuenta }
  ];
}
function _cxcDetalleCuentaCatalogo(codigo, fallbackNombre) {
  var fallback = { mayor: 'S/D', subcuenta: fallbackNombre || 'S/D' };
  if (!codigo) return fallback;
  try {
    var cat = getCatalogosContables(); // ← TU función
    for (var i = 0; i < cat.length; i++) {
      if (String(cat[i].cod) === String(codigo)) {
        return { mayor: String(cat[i].mayor || 'S/D'), subcuenta: String(cat[i].nom || fallbackNombre || 'S/D') };
      }
    }
  } catch (e) { /* catálogo no disponible: usa fallback */ }
  return fallback;
}
function _cxcResolverCuenta(tipo, nombre, codigo) {
  var lista = getBancosYCajas(); // ← TU función
  if (!lista || !lista.length) {
    throw new Error('No hay cuentas cargadas en "datosBancarios" (getBancosYCajas devolvió vacío).');
  }
  var norm = function(s) { return String(s || '').trim().toLowerCase(); };
  var match = null, i, c;

  if (codigo) {
    for (i = 0; i < lista.length && !match; i++) {
      c = lista[i];
      if (norm(tipo) === norm(c.tipo) && norm(codigo) === norm(c.CodigoCuenta || c.codigoCaja || '')) match = c;
    }
  }
  if (!match && nombre) {
    for (i = 0; i < lista.length && !match; i++) {
      c = lista[i];
      if (norm(tipo) === norm(c.tipo) && norm(nombre) === norm(c.NombreBanco || c.nombreCaja || '')) match = c;
    }
  }
  if (!match) {
    throw new Error('Cuenta/Caja no encontrada en datosBancarios: ' + [tipo, nombre, codigo].filter(Boolean).join(' | '));
  }

  if (match.tipo === 'Banco') {
    return {
      NroCuenta: String(match.NroCuenta || ''), NombreBanco: String(match.NombreBanco || ''),
      CodigoBancos: String(match.CodigoBancos || ''), Cuenta: String(match.NroCuenta || ''),
      CodigoCuentaContable: String(match.codContable || ''), CuentaContable: String(match.nomContable || ''),
      NombreCuentaMov: String(match.NombreBanco || '')
    };
  }
  return {
    NroCuenta: '', NombreBanco: String(match.nombreCaja || ''),
    CodigoBancos: String(match.codigoCaja || ''), Cuenta: String(match.codigoCaja || ''),
    CodigoCuentaContable: String(match.codContable || ''), CuentaContable: String(match.nomContable || ''),
    NombreCuentaMov: String(match.nombreCaja || '')
  };
}
function _cxcBuscarFilaEmision(datos, nombre, factura) {
  var norm = function(s) { return String(s || '').trim().toLowerCase(); };
  var nN = norm(nombre), nF = norm(factura);
  var esFilaTx = function(fila) {
    var abono = _cxcNum(fila[CXC_COL_CP.ABONO]);
    var trans = norm(fila[CXC_COL_CP.TRANSACCION]);
    var total = _cxcNum(fila[CXC_COL_CP.TOTAL_FACTURA]);
    return abono > 0 || ((trans === 'cobro' || trans === 'pago') && total <= 0);
  };
  var r;
  for (r = datos.length - 1; r >= 0; r--) {
    if (norm(datos[r][CXC_COL_CP.PROVEEDOR_CLIENTE]) === nN &&
        norm(datos[r][CXC_COL_CP.NRO_FACTURA]) === nF &&
        _cxcNum(datos[r][CXC_COL_CP.TOTAL_FACTURA]) > 0 &&
        !esFilaTx(datos[r])) return r;
  }
  for (r = datos.length - 1; r >= 0; r--) {
    if (norm(datos[r][CXC_COL_CP.PROVEEDOR_CLIENTE]) === nN &&
        norm(datos[r][CXC_COL_CP.NRO_FACTURA]) === nF) return r;
  }
  return -1;
}
function _cxcNum(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;
  if (typeof valor === 'number') return isFinite(valor) ? valor : 0;
  var str = String(valor).trim().replace(/[^\d.,\-]/g, '');
  if (!str) return 0;
  var lastC = str.lastIndexOf(','), lastP = str.lastIndexOf('.');
  if (lastC !== -1 && lastP !== -1) {
    str = (lastC > lastP) ? str.replace(/\./g, '').replace(',', '.') : str.replace(/,/g, '');
  } else if (lastC !== -1) {
    var partes = str.split(',');
    str = (partes.length === 2 && partes[1].length === 3) ? partes.join('') : str.replace(/,/g, '.');
  }
  var n = parseFloat(str);
  return isNaN(n) ? 0 : n;
}
function _cxcR2(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }
function _cxcHoyISO() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function getHistoricoData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName("CobrosPagos") || 
              ss.getSheetByName("COBROS PAGOS") || 
              ss.getSheetByName("Cobros_Pagos") || 
              ss.getActiveSheet();
  
  var data = { cobros: [], pagos: [] };
  
  if (!sheet) return data;
  
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return data;
  
  var headers = rows[0].map(function(h) { 
    return String(h).toLowerCase().trim().replace(/[\r\n\t]/g, "").replace(/\s+/g, ""); 
  });
  
  var idxProvCli          = headers.indexOf("proveedorcliente");
  var idxRif              = headers.indexOf("rif");
  var idxNroFactura       = headers.indexOf("nrofactura");
  var idxFechaFactura     = headers.indexOf("fechafactura");
  var idxTotaFactura      = headers.indexOf("totalfactura"); // ✅ CORREGIDO
  var idxSaldoFactura     = headers.indexOf("saldofactura");
  var idxSaldoFacturaUSD  = headers.indexOf("saldofacturausd");
  var idxCobradaPagada    = headers.indexOf("cobradapagada");
  var idxTransaccion      = headers.indexOf("transaccion");
  var idxAbono            = headers.indexOf("abono");
  var idxMoneda           = headers.indexOf("moneda");
  var idxTasa             = headers.indexOf("tasa");
  var idxDescripcion      = headers.indexOf("descripcion");
  var idxFechaAbono       = headers.indexOf("fechaabono");
  var idxTipoMovimiento   = headers.indexOf("tipomovimiento");
  var idxCuentaDestino    = headers.indexOf("cuentadestino");
  var idxNroCuentaDestino = headers.indexOf("nrocuentadestino");
  var idxBancoDestino     = headers.indexOf("bancodestino");
  
  var parseVzlaNum = function(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    var str = String(val).trim();
    str = str.replace(/\./g, "").replace(",", ".");
    var res = parseFloat(str);
    return isNaN(res) ? 0 : res;
  };
  
  var formatFecha = function(val) {
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    return String(val || "").trim();
  };

  var consolidado = {};

  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    var provCli = idxProvCli !== -1 ? String(row[idxProvCli] || "").trim() : "";
    var nroFact = idxNroFactura !== -1 ? String(row[idxNroFactura] || "").trim() : "";
    var trans   = idxTransaccion !== -1 ? String(row[idxTransaccion] || "").trim() : ""; 
    
    if (!provCli || !nroFact) continue;
    
    var key = provCli + "|" + nroFact + "|" + trans;
    
    var abono      = idxAbono !== -1 ? parseVzlaNum(row[idxAbono]) : 0;
    var saldoUSD   = idxSaldoFacturaUSD !== -1 ? parseVzlaNum(row[idxSaldoFacturaUSD]) : 0;
    var tasa       = idxTasa !== -1 ? parseVzlaNum(row[idxTasa]) || 1 : 1;
    var totalFactUSD = idxTotaFactura !== -1 ? parseVzlaNum(row[idxTotaFactura]) : 0;
    
    var monedaStr = idxMoneda !== -1 ? String(row[idxMoneda]).toLowerCase() : "";
    if (monedaStr.indexOf("dolar") === -1 && monedaStr !== "usd" && tasa > 1) {
      totalFactUSD = totalFactUSD / tasa;
    }

    if (!consolidado[key]) {
      consolidado[key] = {
        proveedorCliente: provCli,
        rif: idxRif !== -1 ? String(row[idxRif] || "").trim() : "",
        factura: nroFact,
        fecha: idxFechaFactura !== -1 ? formatFecha(row[idxFechaFactura]) : "",
        montoTotalUSD: totalFactUSD,
        saldoUSD: saldoUSD,
        transaccion: trans,
        abonos: []
      };
    }
    
    if (saldoUSD < consolidado[key].saldoUSD) {
      consolidado[key].saldoUSD = saldoUSD;
    }
    
    if (abono > 0) {
      var abonoUSD = abono;
      if (monedaStr.indexOf("dolar") === -1 && monedaStr !== "usd" && tasa > 1) {
        abonoUSD = abono / tasa;
      }
      
      var fechaAb       = idxFechaAbono !== -1 ? formatFecha(row[idxFechaAbono]) : "";
      var tipoMov       = idxTipoMovimiento !== -1 ? String(row[idxTipoMovimiento] || "") : "";
      var cuentaDest    = idxCuentaDestino !== -1 ? String(row[idxCuentaDestino] || "") : "";
      var nroCuentaDest = idxNroCuentaDestino !== -1 ? String(row[idxNroCuentaDestino] || "") : "";
      var bancoDest     = idxBancoDestino !== -1 ? String(row[idxBancoDestino] || "") : "";
      var desc          = idxDescripcion !== -1 ? String(row[idxDescripcion] || "") : "";

      consolidado[key].abonos.push({
        fecha: fechaAb || consolidado[key].fecha,
        modalidad: tipoMov,
        banco: cuentaDest,
        nroCuenta: nroCuentaDest,
        bancoDest: bancoDest,
        ref: desc,
        monto: abonoUSD,
        moneda: "USD"
      });
    }
  }

  for (var k in consolidado) {
    var item = consolidado[k];
    
    if (item.saldoUSD <= 0.05) {
      item.estado = "PAGADO";
    } else if (item.abonos.length > 0) {
      item.estado = "ABONO PARCIAL";
    } else {
      item.estado = "PENDIENTE";
    }
    
    var objUI = {
      cliente: item.proveedorCliente,
      proveedor: item.proveedorCliente,
      rif: item.rif,
      factura: item.factura,
      fecha: item.fecha,
      montoTotalUSD: item.montoTotalUSD,
      montoCobradoUSD: item.montoTotalUSD - item.saldoUSD,
      montoPagadoUSD: item.montoTotalUSD - item.saldoUSD,
      saldoUSD: item.saldoUSD,
      estado: item.estado,
      abonos: item.abonos
    };
    
    if (item.transaccion === "Cobro") {
      data.cobros.push(objUI);
    } else {
      data.pagos.push(objUI);
    }
  }
  
  return data;
}
function obtenerDatosContablesCuenta(cuentaFrontend) {
  var valor = String(cuentaFrontend || '').trim();
  if (!valor) return { cod: '1.1.01.01', nom: 'Banco/Caja Default' };

  var partes = valor.split('|');
  
  // Formato nuevo (5 partes): tipo|nombre|codigo|codContable|nomContable
  if (partes.length >= 5) {
    var cod = String(partes[3] || '').trim();
    var nom = String(partes[4] || '').trim();
    if (!nom) nom = String(partes[1] || 'Banco/Caja').trim();
    if (!cod) {
      // Sin código contable: buscar en datosBancarios por nombre
      return buscarCuentaContablePorNombre(partes[1]);
    }
    return { cod: cod, nom: nom };
  }
  
  // Formato viejo (3 partes): tipo|nombre|codigo
  var nombreBuscar = partes.length >= 2 ? partes[1].trim() : valor;
  return buscarCuentaContablePorNombre(nombreBuscar);
}
function buscarCuentaContablePorNombre(nombreBuscar) {
  nombreBuscar = String(nombreBuscar || '').trim();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('datosBancarios');
    if (!hoja) return { cod: '1.1.01.01', nom: nombreBuscar || 'Banco/Caja' };

    var data = hoja.getDataRange().getValues();
    var headers = data[0].map(function(h) {
      return String(h || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-#.]/g, '');
    });

    var colNombreBanco = headers.indexOf('nombrebanco');
    var colNombreCaja = headers.indexOf('nombrecaja');
    var colCodContBanco = headers.indexOf('codcontable');
    var colNomContBanco = headers.indexOf('nomcontable');

    for (var i = 1; i < data.length; i++) {
      var nombreBanco = colNombreBanco !== -1 ? String(data[i][colNombreBanco] || '').trim() : '';
      var nombreCaja = colNombreCaja !== -1 ? String(data[i][colNombreCaja] || '').trim() : '';
      
      if (nombreBanco && nombreBanco.toUpperCase() === nombreBuscar.toUpperCase()) {
        return {
          cod: colCodContBanco !== -1 ? String(data[i][colCodContBanco] || '1.1.01.01').trim() : '1.1.01.01',
          nom: colNomContBanco !== -1 ? String(data[i][colNomContBanco] || nombreBanco).trim() : nombreBanco
        };
      }
      if (nombreCaja && nombreCaja.toUpperCase() === nombreBuscar.toUpperCase()) {
        return {
          cod: colCodContBanco !== -1 ? String(data[i][colCodContBanco] || '1.1.01.01').trim() : '1.1.01.01',
          nom: colNomContBanco !== -1 ? String(data[i][colNomContBanco] || nombreCaja).trim() : nombreCaja
        };
      }
    }
  } catch (e) {
    Logger.log('[CXC] Error buscando cuenta contable: ' + e.message);
  }
  return { cod: '1.1.01.01', nom: nombreBuscar || 'Banco/Caja' };
}
function obtenerDatosCobrosPagos() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var hoja = ss.getSheetByName('CobrosPagos');
        if (!hoja) return [];

        if (hoja.getLastRow() <= 1) return [];

        var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
        
        var hojaVentas = ss.getSheetByName('Ventas');
        var hojaCompras = ss.getSheetByName('Compras');
        
        var infoFacturas = {};
        
        function indexarInfo(hojaFuente, esVenta) {
            if (!hojaFuente) return;
            var data = hojaFuente.getDataRange().getValues();
            if (data.length < 2) return;
            var h = data[0].map(function(x) { return String(x).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-#.]/g, ''); });
            
            var colFact = h.indexOf(esVenta ? 'nrofacturacliente' : 'nrofacturaproveedor');
            var colIva = h.indexOf('ivacliente');
            if (colIva === -1) colIva = h.indexOf('ivaproveedor');
            var colSubTotal = h.indexOf(esVenta ? 'subtotalcliente' : 'subtotalproveedor');
            var colExento = h.indexOf(esVenta ? 'totalexentoventa' : 'totalexentocompra');

            if (colFact !== -1) {
                for (var i = 1; i < data.length; i++) {
                    var fac = String(data[i][colFact]).trim();
                    if (fac) {
                        infoFacturas[fac] = {
                            iva: colIva !== -1 ? (parseFloat(data[i][colIva]) || 0) : 0,
                            subTotal: colSubTotal !== -1 ? (parseFloat(data[i][colSubTotal]) || 0) : 0,
                            exento: colExento !== -1 ? (parseFloat(data[i][colExento]) || 0) : 0
                        };
                    }
                }
            }
        }
        
        indexarInfo(hojaVentas, true);
        indexarInfo(hojaCompras, false);

        var datosConInfo = datos.map(function(row) {
            var nroFactura = String(row[2]).trim();
            var info = infoFacturas[nroFactura] || { iva: 0, subTotal: 0, exento: 0 };
            
            // Agregamos al final del array (posiciones 23, 24, 25)
            while (row.length < 26) row.push('');
            
            row[23] = info.iva;
            row[24] = info.subTotal;
            row[25] = info.exento;
            return row;
        });

        return datosConInfo.map(function(row) {
            return row.map(function(celda) {
                if (celda instanceof Date) return celda.toISOString();
                return celda;
            });
        });

    } catch (error) {
        console.error('Error en obtenerDatosCobrosPagos:', error);
        return [];
    }
}
function normalizarFactura(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().replace(/^0+/, '').toLowerCase(); // Quita ceros a la izquierda
}
function obtenerConfiguracionTercero(nombreTercero, tipo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Busca en la hoja correspondiente según sea CXC (Clientes) o CXP (Proveedores)
  const hoja = ss.getSheetByName(tipo === 'CXC' ? "Clientes" : "Proveedores");
  
  const config = {
    pctIva: 75,   // Valor por defecto SENIAT
    pctIslr: 2    // Valor por defecto general
  };

  if (!hoja) return config;

  const data = hoja.getDataRange().getValues();
  if (data.length < 2) return config;

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const colNombre = headers.findIndex(h => h.includes("nombre") || h.includes("cliente") || h.includes("proveedor"));
  const colPctIva = headers.findIndex(h => h.includes("ret iva") || h.includes("% iva") || h.includes("retencion iva"));
  const colPctIslr = headers.findIndex(h => h.includes("ret islr") || h.includes("% islr") || h.includes("retencion islr"));

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colNombre]).trim().toLowerCase() === String(nombreTercero).trim().toLowerCase()) {
      if (colPctIva !== -1 && data[i][colPctIva] !== "") {
        config.pctIva = parseFloat(data[i][colPctIva]) || config.pctIva;
      }
      if (colPctIslr !== -1 && data[i][colPctIslr] !== "") {
        config.pctIslr = parseFloat(data[i][colPctIslr]) || config.pctIslr;
      }
      break;
    }
  }

  return config;
}
function _resolverCuenta(tipo, nombre, codigo) {
  var vacio = {
    NroCuenta: '', NombreBanco: '', CodigoBancos: '', Cuenta: '',
    CodigoCuentaContable: '', CuentaContable: '', NombreCuentaMov: nombre || ''
  };
  try {
    var lista = getBancosYCajas(); // ← TU función, sin modificar
    if (!lista || !lista.length) throw new Error('No hay cuentas bancarias/cajas cargadas en "datosBancarios".');

    var norm = function(s) { return String(s || '').trim().toLowerCase(); };
    var match = null, i, c;

    // 1) Prioridad: coincidencia por código (única incluso con bancos repetidos)
    if (codigo) {
      for (i = 0; i < lista.length && !match; i++) {
        c = lista[i];
        var codC = norm(c.CodigoCuenta || c.codigoCaja || '');
        if (norm(tipo) === norm(c.tipo) && codC === norm(codigo)) match = c;
      }
    }
    // 2) Fallback: por nombre
    if (!match && nombre) {
      for (i = 0; i < lista.length && !match; i++) {
        c = lista[i];
        var nomC = norm(c.NombreBanco || c.nombreCaja || '');
        if (norm(tipo) === norm(c.tipo) && nomC === norm(nombre)) match = c;
      }
    }

    // Contablemente seguro: si la cuenta no existe, NO se asienta.
    if (!match) {
      throw new Error('Cuenta/Caja no encontrada en "datosBancarios": ' + (tipo || '') + ' ' + (nombre || '') + (codigo ? ' (' + codigo + ')' : ''));
    }

    if (match.tipo === 'Banco') {
      return {
        NroCuenta:            String(match.NroCuenta || ''),
        NombreBanco:          String(match.NombreBanco || ''),
        CodigoBancos:         String(match.CodigoBancos || ''),
        Cuenta:               String(match.NroCuenta || ''),
        CodigoCuentaContable: String(match.codContable || ''),
        CuentaContable:       String(match.nomContable || ''),
        NombreCuentaMov:      String(match.NombreBanco || '')
      };
    }
    // Caja
    return {
      NroCuenta:            '',
      NombreBanco:          String(match.nombreCaja || ''),
      CodigoBancos:         String(match.codigoCaja || ''),
      Cuenta:               String(match.codigoCaja || ''),
      CodigoCuentaContable: String(match.codContable || ''),
      CuentaContable:       String(match.nomContable || ''),
      NombreCuentaMov:      String(match.nombreCaja || '')
    };
  } catch (e) {
    throw new Error(e.message || 'Error al resolver la cuenta bancaria.');
  }
}
function parseNumeroCliente(val) {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return isFinite(val) ? val : 0;
    var str = String(val).trim().replace(/[^\d.,\-]/g, '');
    if (!str) return 0;
    var lastC = str.lastIndexOf(',');
    var lastP = str.lastIndexOf('.');
    if (lastC !== -1 && lastP !== -1) {
      str = (lastC > lastP) ? str.replace(/\./g, '').replace(',', '.') : str.replace(/,/g, '');
    } else if (lastC !== -1) {
      var partes = str.split(',');
      str = (partes.length === 2 && partes[1].length === 3) ? partes.join('') : str.replace(/,/g, '.');
    }
    var n = parseFloat(str);
    return isNaN(n) ? 0 : n;
  }

  function filtrarYMostrarFacturas() {
    var inp = document.getElementById('searchProveedor');
    var f = (inp ? inp.value : '').trim().toLowerCase();
    var fDesde = document.getElementById('filtroFechaDesde') ? document.getElementById('filtroFechaDesde').value : '';
    var fHasta = document.getElementById('filtroFechaHasta') ? document.getElementById('filtroFechaHasta').value : '';
    var tb = document.getElementById('tablaFacturasBody');
    if (!tb) return;

    var esVenta = (state.filtroEstado === 'porCobrar' || state.filtroEstado === 'cobrada');
    var soloSaldosPendientes = (state.filtroEstado === 'porPagar' || state.filtroEstado === 'porCobrar');

    var list = state.facturas.filter(function (x) {
      var trans = safeStr(x.Transaccion).toLowerCase();
      var matchTipo = esVenta
        ? (trans.indexOf('vent') === 0 || trans.indexOf('cobr') === 0)
        : (trans.indexOf('compr') === 0 || trans.indexOf('pag') === 0);

      var nom = safeStr(x.ProveedorCliente).toLowerCase();
      var nro = safeStr(x.NroFactura).toLowerCase();
      var matchTexto = (!f || nom.indexOf(f) !== -1 || nro.indexOf(f) !== -1);

      var fecha = safeStr(x.FechaFactura).slice(0, 10);
      var matchFecha = true;
      if (fDesde && fecha && fecha < fDesde) matchFecha = false;
      if (fHasta && fecha && fecha > fHasta) matchFecha = false;

      // Usar parseNumeroCliente para tolerar comas o números ya convertidos
      var saldoBs = parseNumeroCliente(x.SaldoFactura);
      var saldoUSD = parseNumeroCliente(x.SaldoFacturaUSD);
      var tieneSaldo = (saldoBs > 0.01 || saldoUSD > 0.01);
      var noAnulada = safeStr(x.cobradaPagada).toLowerCase() !== 'anulada';

      if (soloSaldosPendientes) {
        return matchTipo && matchTexto && matchFecha && tieneSaldo && noAnulada;
      } else {
        return matchTipo && matchTexto && matchFecha && (!tieneSaldo) && noAnulada;
      }
    });

    tb.innerHTML = '';

    if (!list.length) {
      tb.innerHTML = '<tr><td colspan="9" class="text-center py-4 text-muted"><i class="bi bi-inbox me-2"></i> No hay facturas en esta categoría</td></tr>';
      return;
    }

    list.forEach(function (x) {
      var nombre = x.ProveedorCliente || 'N/A';
      var nro = x.NroFactura || 'N/A';
      var trans = x.Transaccion || (esVenta ? 'Venta' : 'Compra');
      var tasa = parseNumeroCliente(x.Tasa) || 1;
      var totalBs = parseNumeroCliente(x.TotalFactura || x.TotaFactura);
      var totalUSD = tasa > 0 ? (totalBs / tasa) : 0;
      var saldoBs = parseNumeroCliente(x.SaldoFactura);
      var saldoUSD = parseNumeroCliente(x.SaldoFacturaUSD || (tasa > 0 ? saldoBs / tasa : 0));
      var fechaStr = safeStr(x.FechaFactura).slice(0, 10);

      var nSan = sanitizeJsAttr(nro);
      var tSan = sanitizeJsAttr(trans);

      var stVal = safeStr(x.cobradaPagada || 'Pendiente');
      var an = stVal.toLowerCase() === 'anulada';
      var sp = (saldoBs > 0.01) || (saldoUSD > 0.01);

      var sb = an ? 'badge-anulada' : (!sp ? 'badge-active' : 'badge-pendiente');
      var st = an ? 'Anulada' : (!sp ? (esVenta ? 'Cobrada' : 'Pagada') : 'Pendiente');

      var tieneRet = (x.nroComprobantesRetIVA && safeStr(x.nroComprobantesRetIVA).trim() !== '') ||
                     (x.nroComprobantesRetISLR && safeStr(x.nroComprobantesRetISLR).trim() !== '');

      var tr = document.createElement('tr');
      tr.innerHTML =
        '<td class="fw-bold text-light">' + nombre + '</td>' +
        '<td><span class="badge bg-dark border border-secondary">' + nro + '</span></td>' +
        '<td class="text-muted small">' + fechaStr + '</td>' +
        '<td class="text-end text-muted">' + fmtBs(saldoBs) + '</td>' +
        '<td class="text-end text-warning fw-bold">' + fmtUSD(saldoUSD) + '</td>' +
        '<td class="text-end text-success fw-bold">' + fmtUSD(totalUSD) + '</td>' +
        '<td class="text-center"><span class="badge ' + (esVenta ? 'bg-primary' : 'bg-info') + '">' + trans + '</span></td>' +
        '<td class="text-center"><span class="badge-status ' + sb + '">' + st + '</span></td>' +
        '<td class="text-center"><div class="btn-action-group">' +
          '<button type="button" class="btn-action ver" onclick="ModuloFacturasLista.verFactura(\'' + nSan + '\')" title="Ver Detalles"><i class="bi bi-eye"></i></button>' +
          '<button type="button" class="btn-action historial" onclick="ModuloFacturasLista.verHistorialAbonos(\'' + nSan + '\')" title="Historial de Abonos y Pagos"><i class="bi bi-clock-history"></i></button>' +
          (tieneRet ?
            '<button type="button" class="btn-action retencion" onclick="ModuloFacturasLista.imprimirRetencionFactura(\'' + nSan + '\')" title="Imprimir Comprobante"><i class="bi bi-file-earmark-check"></i></button>'
          : '') +
          (!an && sp ?
            '<button type="button" class="btn-action editar" onclick="ModuloFacturasLista.abrirModalEdicion(\'' + nSan + '\')" title="Editar"><i class="bi bi-pencil-square"></i></button>' +
            '<button type="button" class="btn-action pagar" onclick="ModuloFacturasLista.abrirModalLiquidacion(\'' + nSan + '\')" title="' + (esVenta ? 'Cobrar' : 'Pagar') + '"><i class="bi bi-cash-coin"></i></button>'
          : '') +
          ((state.userRole === 'Admin' || state.userRole === 'Owner') && !an ?
            '<button type="button" class="btn-action anular" onclick="ModuloFacturasLista.anularFactura(\'' + nSan + '\', \'' + tSan + '\')" title="Anular Factura"><i class="bi bi-ban"></i></button>' +
            '<button type="button" class="btn-action eliminar" onclick="ModuloFacturasLista.borrarFactura(\'' + nSan + '\')" title="Eliminar"><i class="bi bi-trash"></i></button>'
          : '') +
        '</div></td>';
      tb.appendChild(tr);
    });
  }
  
function validarIntegridadDatos() {
  Logger.log('[VALIDAR INTEGRIDAD] Iniciando validación general...');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaCP = ss.getSheetByName('CobrosPagos');
  var hojaAsientos = ss.getSheetByName('Asientos');
  
  var reporte = {
    timestamp: new Date().toISOString(),
    errores: [],
    advertencias: [],
    info: [],
    estado: 'OK'
  };
  
  // ============================================================
  // VALIDACIÓN 1: CobrosPagos existe y tiene datos
  // ============================================================
  
  if (!hojaCP) {
    reporte.errores.push('Hoja CobrosPagos no encontrada');
    reporte.estado = 'ERROR';
    return reporte;
  }
  
  var dataCP = hojaCP.getDataRange().getValues();
  if (dataCP.length < 2) {
    reporte.advertencias.push('CobrosPagos está vacía (solo encabezado)');
  } else {
    reporte.info.push('CobrosPagos: ' + (dataCP.length - 1) + ' filas de datos');
  }
  
  // ============================================================
  // VALIDACIÓN 2: Verificar encabezados críticos
  // ============================================================
  
  var encabezadosCriticos = [
    'ProveedorCliente', 'NroFactura', 'TotalFactura', 'SaldoFactura', 
    'SaldoFacturaUSD', 'cobradaPagada', 'id', 'fechaVencimiento'
  ];
  
  var headersCP = dataCP[0];
  var encabezadosFaltantes = [];
  
  encabezadosCriticos.forEach(function(enc) {
    if (headersCP.indexOf(enc) === -1) {
      encabezadosFaltantes.push(enc);
    }
  });
  
  if (encabezadosFaltantes.length > 0) {
    reporte.errores.push('Encabezados faltantes en CobrosPagos: ' + encabezadosFaltantes.join(', '));
    reporte.estado = 'ERROR';
  } else {
    reporte.info.push('✓ Todos los encabezados críticos presentes');
  }
  
  // ============================================================
  // VALIDACIÓN 3: Detectar duplicados de facturas
  // ============================================================
  
  var colFactura = headersCP.indexOf('NroFactura');
  var colTercero = headersCP.indexOf('ProveedorCliente');
  var duplicados = {};
  var duplicadosEncontrados = [];
  
  for (var i = 1; i < dataCP.length; i++) {
    var tercero = String(dataCP[i][colTercero] || '').trim();
    var factura = String(dataCP[i][colFactura] || '').trim();
    
    if (!tercero || !factura) continue;
    
    var clave = tercero + '|' + factura;
    
    if (duplicados[clave]) {
      duplicados[clave].push(i + 1);
    } else {
      duplicados[clave] = [i + 1];
    }
  }
  
  Object.keys(duplicados).forEach(function(clave) {
    if (duplicados[clave].length > 1) {
      duplicadosEncontrados.push(clave + ' (filas: ' + duplicados[clave].join(', ') + ')');
    }
  });
  
  if (duplicadosEncontrados.length > 0) {
    reporte.advertencias.push('Facturas duplicadas encontradas: ' + duplicadosEncontrados.join('; '));
  } else {
    reporte.info.push('✓ No hay duplicados de facturas');
  }
  
  // ============================================================
  // VALIDACIÓN 4: Verificar saldos negativos
  // ============================================================
  
  var colSaldoUSD = headersCP.indexOf('SaldoFacturaUSD');
  var colSaldoBs = headersCP.indexOf('SaldoFactura');
  var saldosNegativos = [];
  
  for (var j = 1; j < dataCP.length; j++) {
    var saldoUSD = parseNumeroLatino(dataCP[j][colSaldoUSD]);
    var saldoBs = parseNumeroLatino(dataCP[j][colSaldoBs]);
    
    if (saldoUSD < -0.01 || saldoBs < -0.01) {
      var factura = String(dataCP[j][colFactura] || '').trim();
      saldosNegativos.push(factura + ' (USD: $' + saldoUSD.toFixed(2) + ', Bs: ' + saldoBs.toFixed(2) + ')');
    }
  }
  
  if (saldosNegativos.length > 0) {
    reporte.advertencias.push('Saldos negativos detectados: ' + saldosNegativos.join('; '));
  } else {
    reporte.info.push('✓ No hay saldos negativos');
  }
  
  // ============================================================
  // VALIDACIÓN 5: Verificar fechas inválidas
  // ============================================================
  
  var colFechaFactura = headersCP.indexOf('FechaFactura');
  var colFechaVencimiento = headersCP.indexOf('fechaVencimiento');
  var fechasInvalidas = [];
  
  for (var k = 1; k < dataCP.length; k++) {
    var fechaEmision = parseFechaSafe(dataCP[k][colFechaFactura]);
    var fechaVencimiento = parseFechaSafe(dataCP[k][colFechaVencimiento]);
    
    if (fechaEmision && fechaVencimiento && fechaVencimiento < fechaEmision) {
      var factura = String(dataCP[k][colFactura] || '').trim();
      fechasInvalidas.push(factura + ' (emisión: ' + formatearFecha(fechaEmision) + 
                          ', vencimiento: ' + formatearFecha(fechaVencimiento) + ')');
    }
  }
  
  if (fechasInvalidas.length > 0) {
    reporte.advertencias.push('Fechas inconsistentes: ' + fechasInvalidas.join('; '));
  } else {
    reporte.info.push('✓ Fechas consistentes');
  }
  
  // ============================================================
  // VALIDACIÓN 6: Verificar Asientos balanceados
  // ============================================================
  
  if (hojaAsientos) {
    var dataAsientos = hojaAsientos.getDataRange().getValues();
    
    if (dataAsientos.length > 1) {
      var asientosNoBalanceados = [];
      var asientoActual = null;
      var totalDebe = 0;
      var totalHaber = 0;
      
      for (var m = 1; m < dataAsientos.length; m++) {
        var nroAsiento = String(dataAsientos[m][0] || '').trim();
        var monto = parseNumeroLatino(dataAsientos[m][7]); // Columna H: MontoTransaccion
        
        if (nroAsiento !== asientoActual && asientoActual !== null) {
          // Cambió el asiento, verificar balance anterior
          if (Math.abs(totalDebe - totalHaber) > 0.01) {
            asientosNoBalanceados.push(asientoActual + ' (Debe: ' + totalDebe.toFixed(2) + 
                                      ', Haber: ' + totalHaber.toFixed(2) + ')');
          }
          totalDebe = 0;
          totalHaber = 0;
        }
        
        asientoActual = nroAsiento;
        
        if (monto > 0) {
          totalDebe += monto;
        } else {
          totalHaber += Math.abs(monto);
        }
      }
      
      // Verificar último asiento
      if (Math.abs(totalDebe - totalHaber) > 0.01) {
        asientosNoBalanceados.push(asientoActual + ' (Debe: ' + totalDebe.toFixed(2) + 
                                  ', Haber: ' + totalHaber.toFixed(2) + ')');
      }
      
      if (asientosNoBalanceados.length > 0) {
        reporte.errores.push('Asientos desbalanceados: ' + asientosNoBalanceados.join('; '));
        reporte.estado = 'ERROR';
      } else {
        reporte.info.push('✓ Todos los asientos están balanceados');
      }
    }
  }
  
  // ============================================================
  // RESUMEN FINAL
  // ============================================================
  
  if (reporte.errores.length > 0) {
    reporte.estado = 'ERROR';
  } else if (reporte.advertencias.length > 0) {
    reporte.estado = 'ADVERTENCIA';
  }
  
  Logger.log('[VALIDAR INTEGRIDAD] Estado final: ' + reporte.estado);
  Logger.log('[VALIDAR INTEGRIDAD] Errores: ' + reporte.errores.length + 
             ' | Advertencias: ' + reporte.advertencias.length + 
             ' | Info: ' + reporte.info.length);
  
  return reporte;
}
function ejecutarPagoCobro(data) {
  toggleLoader(true, `Registrando ${data.tipoMovimiento.toLowerCase()}...`);

  if (typeof google !== 'undefined' && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler(function(resp) {
        toggleLoader(false);
        // resp puede ser {ok:true, saldoBs, saldoUSD, cobradaPagada} o true
        aplicarPagoLocal(data, resp);
        showSwalToast(`${data.tipoMovimiento} registrado: $${data.montoUSD.toFixed(2)}`, 'success');
      })
      .withFailureHandler(function(e) {
        toggleLoader(false);
        showSwalToast('Error: ' + (e.message || 'No se pudo registrar'), 'error');
      })
      .registrarPagoCobro({
        nroFactura: data.nroFactura,
        monto: data.monto,
        moneda: data.moneda,
        montoUSD: data.montoUSD,
        montoBS: data.montoBS,
        fecha: data.fecha,
        descripcion: data.descripcion,
        tipoMovimiento: data.tipoMovimiento,
        tasa: data.tasa
      });
  } else {
    // Modo demo sin Apps Script: aplicamos localmente
    toggleLoader(false);
    aplicarPagoLocal(data, null);
    showSwalToast(`${data.tipoMovimiento} registrado (modo local): $${data.montoUSD.toFixed(2)}`, 'info');
  }
}
function aplicarPagoLocal(data, resp) {
  const idx = facturasData.findIndex(x => String(x.NroFactura) === String(data.nroFactura));
  if (idx === -1) return;

  const f = facturasData[idx];
  const nuevoSaldoBs  = resp && typeof resp.saldoBs  === 'number' ? resp.saldoBs  : Math.max(0, Number(f.SaldoFactura || 0) - data.montoBS);
  const nuevoSaldoUSD = resp && typeof resp.saldoUSD === 'number' ? resp.saldoUSD : Math.max(0, Number(f.SaldoFacturaUSD || 0) - data.montoUSD);

  f.SaldoFactura = nuevoSaldoBs;
  f.SaldoFacturaUSD = nuevoSaldoUSD;
  f.cobradaPagada = (nuevoSaldoBs <= 0.01 && nuevoSaldoUSD <= 0.01)
    ? (data.esVenta ? 'Cobrada' : 'Pagada')
    : 'Pendiente';
  f.fechaAbono = data.fecha;
  f.Abono = (Number(f.Abono || 0) + data.montoUSD);

  filtrarYMostrarFacturas(transaccionActual);
}
function getHistorialAbonos(nroFactura) {
  try {
    if (!nroFactura) return [];

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var shCP = ss.getSheetByName(CFG_CXC.HOJA_COBROS_PAGOS);
    if (!shCP) return [];

    var data = shCP.getDataRange().getValues();
    if (data.length < 2) return [];

    var hd = data[0].map(function(h) { return String(h).trim(); });

    var idx = {
      nroFactura:     hd.indexOf('NroFactura'),
      proveedor:      hd.indexOf('ProveedorCliente'),
      rif:            hd.indexOf('Rif'),
      total:          hd.indexOf('TotaFactura') !== -1 ? hd.indexOf('TotaFactura') : hd.indexOf('TotalFactura'),
      saldoBs:        hd.indexOf('SaldoFactura'),
      saldoUSD:       hd.indexOf('SaldoFacturaUSD'),
      status:         hd.indexOf('cobradaPagada'),
      transaccion:    hd.indexOf('Transaccion'),
      abono:          hd.indexOf('Abono'),
      moneda:         hd.indexOf('Moneda'),
      tasa:           hd.indexOf('Tasa'),
      descripcion:    hd.indexOf('Descripcion'),
      fechaAbono:     hd.indexOf('fechaAbono'),
      tipoMovimiento: hd.indexOf('TipoMovimiento'),
      cuentaDestino:  hd.indexOf('CuentaDestino'),
      nroCuentaDest:  hd.indexOf('NroCuentaDestino'),
      bancoDestino:   hd.indexOf('BancoDestino'),
      compIVA:        hd.indexOf('nroComprobantesRetIVA'),
      compISLR:       hd.indexOf('nroComprobantesRetISLR'),
      id:             hd.indexOf('id')
    };

    var nroBuscado = String(nroFactura).trim().toLowerCase();
    var out = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var facFila = String(row[idx.nroFactura] || '').trim().toLowerCase();
      if (facFila !== nroBuscado) continue;

      var totFila  = Number(row[idx.total]) || 0;
      var abono    = parseNumeroLatino(row[idx.abono]);
      var tipoMov  = String(row[idx.tipoMovimiento] || '').trim().toUpperCase();
      var trans    = String(row[idx.transaccion] || '').trim();
      var compIVA  = String(row[idx.compIVA]  || '').trim();
      var compISLR = String(row[idx.compISLR] || '').trim();

      // Saltar la fila "madre" (factura original sin movimientos)
      var esFilaMadre = (totFila > 0 && abono === 0 && !tipoMov && !compIVA && !compISLR);
      if (esFilaMadre) continue;

      // Fecha normalizada
      var fechaRaw = row[idx.fechaAbono];
      var fechaStr = '';
      if (fechaRaw instanceof Date) {
        fechaStr = Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        fechaStr = String(fechaRaw || '').slice(0, 10);
      }

      var tasa   = parseNumeroLatino(row[idx.tasa]) || 1;
      var moneda = String(row[idx.moneda] || '').toUpperCase();
      var montoUSD = abono;
      if (moneda.indexOf('BOLIVAR') !== -1 || moneda === 'BS') {
        montoUSD = abono / tasa;
      }
      var montoBs = montoUSD * tasa;

      // Tipo legible
      var tipoLegible = 'Abono';
      if (tipoMov === 'RETENCION' || trans.toUpperCase().indexOf('RETENCION') !== -1) {
        tipoLegible = 'Retención ' + (compIVA ? 'IVA' : compISLR ? 'ISLR' : '');
      } else if (tipoMov === 'ANTICIPO') {
        tipoLegible = 'Anticipo';
      } else if (tipoMov === 'CXC' || tipoMov === 'CXP') {
        tipoLegible = String(trans || 'Cobro/Pago');
      } else if (tipoMov) {
        tipoLegible = tipoMov;
      } else if (trans) {
        tipoLegible = trans;
      }

      out.push({
        Fecha:             fechaStr,
        FechaAbono:        fechaStr,
        TipoMovimiento:    tipoLegible,
        Transaccion:       trans,
        MontoUSD:          montoUSD,
        MontoBS:           montoBs,
        Moneda:            moneda || 'USD',
        Tasa:              tasa,
        Descripcion:       String(row[idx.descripcion] || ''),
        ProveedorCliente:  String(row[idx.proveedor] || ''),
        Rif:               String(row[idx.rif] || ''),
        NroFactura:        String(row[idx.nroFactura] || ''),
        StatusDespues:     String(row[idx.status] || ''),
        SaldoUSDDespues:   parseNumeroLatino(row[idx.saldoUSD]),
        SaldoBsDespues:    parseNumeroLatino(row[idx.saldoBs]),
        CuentaDestino:     String(row[idx.cuentaDestino] || ''),
        NroCuentaDestino:  String(row[idx.nroCuentaDest] || ''),
        BancoDestino:      String(row[idx.bancoDestino] || ''),
        ComprobanteIVA:    compIVA,
        ComprobanteISLR:   compISLR,
        id:                String(row[idx.id] || '')
      });
    }

    out.sort(function(a, b) { return String(b.Fecha).localeCompare(String(a.Fecha)); });
    return out;

  } catch (e) {
    Logger.log('[getHistorialAbonos] Error: ' + e.message);
    throw new Error('getHistorialAbonos: ' + e.message);
  }
}
function parseNumeroLatino(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;
  if (typeof valor === 'number') return isFinite(valor) ? valor : 0;
  var str = String(valor).trim().replace(/[^\d.,\-]/g, '');
  if (!str) return 0;
  var lastC = str.lastIndexOf(',');
  var lastP = str.lastIndexOf('.');
  if (lastC !== -1 && lastP !== -1) {
    str = (lastC > lastP) ? str.replace(/\./g, '').replace(',', '.') : str.replace(/,/g, '');
  } else if (lastC !== -1) {
    var partes = str.split(',');
    str = (partes.length === 2 && partes[1].length === 3) ? partes.join('') : str.replace(/,/g, '.');
  }
  var n = parseFloat(str);
  return isNaN(n) ? 0 : n;
}
function registrarAnticipoTercero(payload, token) {
  if (!payload) throw new Error('Payload vacío para registrarAnticipoTercero.');

  var tipo = String(payload.tipo || '').toUpperCase();
  if (tipo !== 'CXC' && tipo !== 'CXP') throw new Error('Tipo inválido: ' + payload.tipo);

  var esCobro = (tipo === 'CXC');
  var tasa = parseNumeroLatino(payload.tasa);
  if (!isFinite(tasa) || tasa <= 0) throw new Error('Tasa inválida: ' + payload.tasa);

  var monto = parseNumeroLatino(payload.monto);
  if (!isFinite(monto) || monto <= 0) throw new Error('Monto inválido: ' + payload.monto);

  var moneda = String(payload.moneda || 'USD').toUpperCase();
  if (moneda !== 'USD' && moneda !== 'BS') throw new Error('Moneda inválida: ' + payload.moneda);

  var fecha = payload.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fecha))) throw new Error('Fecha inválida: ' + fecha);

  var nombreEntidad = String(payload.nombreEntidad || '').trim();
  if (!nombreEntidad) throw new Error('Nombre del tercero requerido.');

  var modalidad = String(payload.modalidad || 'Transferencia').trim();
  var cuentaFront = String(payload.cuenta || '').trim();
  if (!cuentaFront) throw new Error('Cuenta/Caja requerida.');

  var esEfectivo = (modalidad === 'Efectivo');
  var referencia = String(payload.referencia || '').trim();
  if (!esEfectivo && !referencia) throw new Error('Referencia bancaria requerida para modalidad ' + modalidad + '.');
  if (esEfectivo && !referencia) referencia = 'EFECTIVO';

  var conceptoPersonalizado = String(payload.conceptoPersonalizado || '').trim()
    || ('Anticipo ' + (esCobro ? 'de Cliente' : 'a Proveedor') + ': ' + nombreEntidad);

  var montoUSD = Math.round(((moneda === 'USD') ? monto : (monto / tasa)) * 100) / 100;
  var montoBs = Math.round(((moneda === 'USD') ? (monto * tasa) : monto) * 100) / 100;

  var ctaParts = cuentaFront.split('|');
  var ctaTipo = String(ctaParts[0] || (esEfectivo ? 'Caja' : 'Banco')).trim();
  var ctaNombre = String(ctaParts[1] || 'Caja/Banco').trim();
  var ctaCod = String(ctaParts[2] || '').trim();
  var ctaCodContable = String(ctaParts[3] || '').trim();
  var ctaNomContable = String(ctaParts[4] || '').trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  var shMB = null;
  var filaMBEscrita = -1;
  var filasCPEscritas = 0;
  var shCPref = null;
  var primeraFilaCP = -1;
  var nroAsientoContable = null;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    shMB = ss.getSheetByName(CFG_CXC.HOJA_MOV_BANCO);
    shCPref = ss.getSheetByName(CFG_CXC.HOJA_COBROS_PAGOS);
    if (!shCPref) throw new Error('No existe la hoja CobrosPagos.');

    var correlativo = obtenerSiguienteCorrelativoCobrosPagos('A', true);

    var dataMov = {
      TipoMov: ctaTipo,
      NroCuenta: ctaCod,
      NombreBanco: ctaNombre,
      CodigoBancos: ctaCod,
      Cuenta: ctaNombre,
      ReferenciaMov: referencia,
      descMov: conceptoPersonalizado,
      TipoTransaccion: esCobro ? 'Entrada' : 'Salida',
      MontoMovUSD: montoUSD,
      MonedaMov: (moneda === 'USD') ? 'Dolares' : 'Bolivares',
      TasaMov: tasa,
      IGTF: 0,
      MontoMovBs: montoBs,
      CodigoCuentaContable: ctaCodContable,
      CuentaContable: ctaNomContable,
      NombreCuentaMov: nombreEntidad,
      FechaMov: fecha,
      detalles: [{
        montoUSD: montoUSD,
        montoBs: montoBs,
        concepto: conceptoPersonalizado,
        codigo: ctaCodContable,
        nombre: ctaNomContable,
        cuentaGasto: ''
      }]
    };

    var resMov = agregarEditarMovBancoCaja(dataMov);
    if (!resMov || !resMov.success) throw new Error('Error al registrar en MovBancoCaja: ' + ((resMov && resMov.message) || 'desconocido'));
    filaMBEscrita = shMB.getLastRow();

    var filaCP = new Array(23).fill('');
    filaCP[CXC_COL_CP.PROVEEDOR_CLIENTE] = nombreEntidad;
    filaCP[CXC_COL_CP.RIF]               = String(payload.rifEntidad || '').trim();
    filaCP[CXC_COL_CP.NRO_FACTURA]       = correlativo;
    filaCP[CXC_COL_CP.FECHA_FACTURA]     = new Date(fecha + 'T00:00:00');
    filaCP[CXC_COL_CP.TOTAL_FACTURA]     = 0;
    filaCP[CXC_COL_CP.SALDO_FACTURA]     = montoBs;
    filaCP[CXC_COL_CP.SALDO_FACTURA_USD] = montoUSD;
    filaCP[CXC_COL_CP.COBRADA_PAGADA]    = esCobro ? 'anticipoCobro' : 'anticipoPago';
    filaCP[CXC_COL_CP.TRANSACCION]       = 'Anticipo';
    filaCP[CXC_COL_CP.NRO_FACTURA]       = correlativo;
    filaCP[CXC_COL_CP.ABONO]             = (moneda === 'USD') ? monto : 0;
    filaCP[CXC_COL_CP.ABONO_BS]          = (moneda === 'BS') ? monto : 0;
    filaCP[CXC_COL_CP.MONEDA]            = (moneda === 'USD') ? 'Dolares' : 'Bolivares';
    filaCP[CXC_COL_CP.TASA]              = tasa;
    filaCP[CXC_COL_CP.DESCRIPCION]       = conceptoPersonalizado;
    filaCP[CXC_COL_CP.FECHA_ABONO]       = new Date(fecha + 'T00:00:00');
    filaCP[CXC_COL_CP.TIPO_MOVIMIENTO]   = 'ANTICIPO';
    filaCP[CXC_COL_CP.CUENTA_DESTINO]    = ctaNombre;
    filaCP[CXC_COL_CP.NRO_CUENTA_DESTINO] = ctaCod;
    filaCP[CXC_COL_CP.BANCO_DESTINO]     = ctaNombre;
    filaCP[CXC_COL_CP.FECHA_VENCIMIENTO] = new Date(fecha + 'T00:00:00');

    primeraFilaCP = shCPref.getLastRow() + 1;
    shCPref.getRange(primeraFilaCP, 1, 1, 23).setValues([filaCP]);
    filasCPEscritas = 1;

    try {
      nroAsientoContable = generarCorrelativoAsiento(true);

      var constants = getConstantsContables();
      var ctaAnticipo = esCobro ? constants.ANTICIPO_CLIENTE : constants.ANTICIPO_PROVEEDOR;
      if (!ctaAnticipo || !ctaAnticipo.COD) throw new Error('CONFIGURACIÓN: Falta cuenta ANTICIPO.');

      var metaBanco = _obtenerMetadatosCuentaCatalogo(ctaCodContable, ctaNomContable, ctaTipo, ctaNombre);
      var metaAnticipo = _obtenerMetadatosCuentaCatalogo(ctaAnticipo.COD, ctaAnticipo.NOM, ctaAnticipo.CTA, ctaAnticipo.SUBCTA || ctaAnticipo.SUB);

      var fechaAsiento = new Date(fecha + 'T00:00:00');
      var glosa = 'Anticipo ' + (esCobro ? 'recibido' : 'otorgado') + ' ' + correlativo;
      var lineasAsiento = [];

      if (esCobro) {
        lineasAsiento.push({
          nroAsiento: nroAsientoContable, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
          clienteProveedor: nombreEntidad, descripcionGeneral: glosa,
          codigoCuenta: metaBanco.cod, cuentaContable: metaBanco.nom,
          montoTransaccion: montoBs, descripcionTransaccion: 'Anticipo recibido',
          transaccionRef: correlativo, contabilidad: "1",
          mayor: metaBanco.mayor, subcuenta: metaBanco.subcuenta,
          eeff: metaBanco.eeff, tipoCuenta: metaBanco.tipoCuenta, grupoBalance: metaBanco.grupoBalance
        });
        lineasAsiento.push({
          nroAsiento: nroAsientoContable, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
          clienteProveedor: nombreEntidad, descripcionGeneral: glosa,
          codigoCuenta: metaAnticipo.cod, cuentaContable: metaAnticipo.nom,
          montoTransaccion: montoBs * -1, descripcionTransaccion: 'Anticipo de cliente',
          transaccionRef: correlativo, contabilidad: "1",
          mayor: metaAnticipo.mayor, subcuenta: metaAnticipo.subcuenta,
          eeff: metaAnticipo.eeff, tipoCuenta: metaAnticipo.tipoCuenta, grupoBalance: metaAnticipo.grupoBalance
        });
      } else {
        lineasAsiento.push({
          nroAsiento: nroAsientoContable, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
          clienteProveedor: nombreEntidad, descripcionGeneral: glosa,
          codigoCuenta: metaAnticipo.cod, cuentaContable: metaAnticipo.nom,
          montoTransaccion: montoBs, descripcionTransaccion: 'Anticipo a proveedor',
          transaccionRef: correlativo, contabilidad: "1",
          mayor: metaAnticipo.mayor, subcuenta: metaAnticipo.subcuenta,
          eeff: metaAnticipo.eeff, tipoCuenta: metaAnticipo.tipoCuenta, grupoBalance: metaAnticipo.grupoBalance
        });
        lineasAsiento.push({
          nroAsiento: nroAsientoContable, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
          clienteProveedor: nombreEntidad, descripcionGeneral: glosa,
          codigoCuenta: metaBanco.cod, cuentaContable: metaBanco.nom,
          montoTransaccion: montoBs * -1, descripcionTransaccion: 'Anticipo otorgado',
          transaccionRef: correlativo, contabilidad: "1",
          mayor: metaBanco.mayor, subcuenta: metaBanco.subcuenta,
          eeff: metaBanco.eeff, tipoCuenta: metaBanco.tipoCuenta, grupoBalance: metaBanco.grupoBalance
        });
      }

      guardarAsientoMasivo(lineasAsiento);
    } catch (eAsiento) {
      throw new Error('Se registró el anticipo pero el asiento contable falló: ' + eAsiento.message);
    }

    try {
      if (typeof actualizarSaldosBancariosOptimizado === 'function') actualizarSaldosBancariosOptimizado();
    } catch (eSaldos) {
      Logger.log('[ANTICIPO] Error actualizando saldos: ' + eSaldos.message);
    }

    return {
      ok: true, success: true,
      message: 'Anticipo registrado correctamente.',
      asiento: correlativo, id: correlativo,
      nroAsientoContable: nroAsientoContable,
      montoUSD: montoUSD, montoBs: montoBs, moneda: moneda
    };

  } catch (e) {
    if (filasCPEscritas > 0 && shCPref && primeraFilaCP > 0) {
      try { shCPref.deleteRows(primeraFilaCP, filasCPEscritas); } catch (rb) {}
    }
    if (filaMBEscrita > 0 && shMB) {
      try { shMB.deleteRow(filaMBEscrita); } catch (rb) {}
    }
    throw new Error(e.message || 'Error al registrar anticipo.');
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
function registrarPagoCobro(payload) {
  try {
    if (!payload || !payload.nroFactura) {
      throw new Error('Payload inválido: falta número de factura.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shCP = ss.getSheetByName('CobrosPagos');
    if (!shCP) throw new Error('No existe la hoja CobrosPagos');

    var data = shCP.getDataRange().getValues();
    var hd = data[0].map(function(h) { return String(h || '').trim(); });

    var cFact = hd.indexOf('NroFactura');
    var cProv = hd.indexOf('ProveedorCliente');
    var cRif = hd.indexOf('Rif');
    var cTotal = hd.indexOf('TotalFactura') !== -1 ? hd.indexOf('TotalFactura') : hd.indexOf('TotaFactura');
    var cTipoMov = hd.indexOf('TipoMovimiento');
    var cSaldoBs = hd.indexOf('SaldoFactura');
    var cSaldoUSD = hd.indexOf('SaldoFacturaUSD');
    var cStatus = hd.indexOf('cobradaPagada');

    var nroBuscado = String(payload.nroFactura).trim().toLowerCase();
    var factura = null;

    for (var i = 1; i < data.length; i++) {
      var facFila = String(data[i][cFact] || '').trim().toLowerCase();
      if (facFila !== nroBuscado) continue;

      var totFila = Number(data[i][cTotal]) || 0;
      var tipoMov = String(data[i][cTipoMov] || '').trim();
      if (totFila > 0 && !tipoMov) {
        factura = data[i];
        break;
      }
    }

    if (!factura) throw new Error('Factura no encontrada en CobrosPagos: ' + payload.nroFactura);

    var esCobro = (String(payload.tipoMovimiento || '').toLowerCase() === 'cobro');
    var tasa = Number(payload.tasa) || 1;
    var montoUSD = Number(payload.montoUSD) || 0;
    var montoBs = Number(payload.montoBS) || (montoUSD * tasa);

    if (montoUSD <= 0 && montoBs <= 0) {
      throw new Error('El monto debe ser mayor a 0.');
    }

    // Resolver datos bancarios
    var detallesPago = [];
    if (payload.cuentaSeleccionada) {
      detallesPago.push({
        cuenta: payload.cuentaSeleccionada, // formato: Tipo|Nombre|Codigo|CodContable|NomContable
        monto: (payload.moneda === 'USD') ? montoUSD : montoBs,
        moneda: payload.moneda || 'USD',
        referencia: payload.descripcion || 'Abono Factura ' + payload.nroFactura
      });
    }

    var facturasPayload = [{
      NroFacturaCliente: factura[cFact],
      NroFacturaProveedor: factura[cFact],
      NombreCliente: factura[cProv],
      NombreProveedor: factura[cProv],
      RifCliente: factura[cRif],
      RifProveedor: factura[cRif],
      montoAnticipoAplicadoUSD: 0,
      montoRetencionAplicadaUSD: 0,
      montoAbonoBancoUSD: montoUSD,
      montoTotalAplicadoUSD: montoUSD,
      montoTotalAplicadoBs: montoBs
    }];

    var infoPago = {
      fecha: payload.fecha || formatDate(new Date()),
      tasa: tasa,
      anticiposCruzados: [],
      retencionesAplicadas: [],
      detallesPago: detallesPago
    };

    var resultado = registrarLiquidacionCxCp({ infoPago: infoPago, facturas: facturasPayload }, esCobro);

    // Releer saldos actualizados
    var saldoFinalBs = 0, saldoFinalUSD = 0, statusFinal = '';
    var dataPost = shCP.getDataRange().getValues();
    for (var k = 1; k < dataPost.length; k++) {
      if (String(dataPost[k][cFact] || '').trim().toLowerCase() === nroBuscado) {
        var totPost = Number(dataPost[k][cTotal]) || 0;
        var tipoPost = String(dataPost[k][cTipoMov] || '').trim();
        if (totPost > 0 && !tipoPost) {
          saldoFinalBs = Number(dataPost[k][cSaldoBs]) || 0;
          saldoFinalUSD = Number(dataPost[k][cSaldoUSD]) || 0;
          statusFinal = String(dataPost[k][cStatus] || '');
          break;
        }
      }
    }

    return {
      ok: true,
      saldoBs: saldoFinalBs,
      saldoUSD: saldoFinalUSD,
      cobradaPagada: statusFinal,
      asiento: resultado.asiento,
      id: resultado.id
    };
  } catch (e) {
    Logger.log('[registrarPagoCobro] Error: ' + e.message);
    throw new Error(e.message);
  }
}
function borrarFacturaCobrosPagos(nroFactura) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shCP = ss.getSheetByName('CobrosPagos');
    if (!shCP) throw new Error('Hoja CobrosPagos no encontrada');

    var data = shCP.getDataRange().getValues();
    var nroBuscado = String(nroFactura).trim().toLowerCase();
    var filasAEliminar = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2] || '').trim().toLowerCase() === nroBuscado) {
        filasAEliminar.push(i + 1);
      }
    }

    // Eliminar de abajo hacia arriba para preservar índices
    for (var k = filasAEliminar.length - 1; k >= 0; k--) {
      shCP.deleteRow(filasAEliminar[k]);
    }

    return { ok: true, message: 'Factura eliminada de CobrosPagos.' };
  } catch (e) {
    Logger.log('[borrarFacturaCobrosPagos] Error: ' + e.message);
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}
function obtenerAnticiposDisponibles(nombreTercero, rifTercero) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shCP = ss.getSheetByName('CobrosPagos');
    if (!shCP) return [];

    var data = shCP.getDataRange().getValues();
    if (data.length < 2) return [];

    var hd = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var cNom = hd.indexOf('proveedorcliente');
    var cRif = hd.indexOf('rif');
    var cTrans = hd.indexOf('transaccion');
    var cTipoMov = hd.indexOf('tipomovimiento');
    var cSaldoUSD = hd.indexOf('saldofacturausd');
    var cSaldoBs = hd.indexOf('saldofactura');
    var cId = hd.indexOf('id');
    var cFact = hd.indexOf('nrofactura');

    var nNorm = String(nombreTercero || '').trim().toLowerCase();
    var rNorm = String(rifTercero || '').trim().toLowerCase();
    var anticipos = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var trans = String(row[cTrans] || '').toLowerCase();
      var tipoMov = String(row[cTipoMov] || '').toLowerCase();
      var esAnticipo = (trans.indexOf('anticip') !== -1 || tipoMov === 'anticipo');
      if (!esAnticipo) continue;

      var nomFila = String(row[cNom] || '').trim().toLowerCase();
      var rifFila = String(row[cRif] || '').trim().toLowerCase();
      var coincide = (rNorm && rifFila === rNorm) || (nNorm && nomFila === nNorm);
      if (!coincide) continue;

      var saldoUSD = parseNumeroLatino(row[cSaldoUSD]);
      var saldoBs = parseNumeroLatino(row[cSaldoBs]);

      if (saldoUSD > 0.01 || saldoBs > 0.01) {
        anticipos.push({
          idAnticipo: String(row[cId] || row[cFact] || ('ANT-' + i)),
          nroFactura: String(row[cFact] || ''),
          saldoUSD: saldoUSD,
          saldoBs: saldoBs
        });
      }
    }
    return anticipos;
  } catch (e) {
    Logger.log('[ANTICIPOS] Error: ' + e.message);
    return [];
  }
}
function getFacturasPorTransaccion(transaccion, idSpreadsheetEmpresa) {
  try {
    var ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) throw new Error('No se pudo abrir la hoja de cálculo de la empresa activa.');

    var sheet = ss.getSheetByName('CobrosPagos');
    if (!sheet) throw new Error('No se encontró la hoja "CobrosPagos" en la empresa activa.');

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var headers = data[0].map(function(h) { 
      return String(h || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-#.]/g, ''); 
    });

    var idx = {
      proveedor: headers.indexOf('proveedorcliente'),
      rif: headers.indexOf('rif'),
      factura: headers.indexOf('nrofactura'),
      fecha: headers.indexOf('fechafactura'),
      total: headers.indexOf('totalfactura') !== -1 ? headers.indexOf('totalfactura') : headers.indexOf('totafactura'),
      saldoBs: headers.indexOf('saldofactura'),
      saldoUSD: headers.indexOf('saldofacturausd'),
      status: headers.indexOf('cobradapagada'),
      transaccion: headers.indexOf('transaccion'),
      tasa: headers.indexOf('tasa'),
      id: headers.indexOf('id'),
      abono: headers.indexOf('abono'),
      tipoMov: headers.indexOf('tipomovimiento'),
      banco: headers.indexOf('bancodestino'),
      cuenta: headers.indexOf('cuentadestino'),
      nroCuenta: headers.indexOf('nrocuentadestino'),
      fVenc: headers.indexOf('fechavencimiento'),
      compIVA: headers.indexOf('nrocomprobantesretiva'),
      compISLR: headers.indexOf('nrocomprobantesretislr')
    };

    var esVenta = String(transaccion || '').toLowerCase().indexOf('vent') === 0 
               || String(transaccion || '').toLowerCase().indexOf('cobr') === 0;

    // Indexar datos fiscales desde Ventas / Compras
    var infoFiscal = {};
    var shFuente = ss.getSheetByName(esVenta ? 'Ventas' : 'Compras');
    if (shFuente && shFuente.getLastRow() > 1) {
      var dF = shFuente.getDataRange().getValues();
      var hF = dF[0].map(function(x) { 
        return String(x || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-#.]/g, ''); 
      });
      var cFactF = esVenta ? hF.indexOf('nrofacturacliente') : hF.indexOf('nrofacturaproveedor');
      var cIvaF = esVenta ? hF.indexOf('ivacliente') : hF.indexOf('ivaproveedor');
      var cSubF = esVenta ? hF.indexOf('subtotalcliente') : hF.indexOf('subtotalproveedor');
      var cExeF = esVenta ? hF.indexOf('totalexentoventa') : hF.indexOf('totalexentocompra');

      for (var f = 1; f < dF.length; f++) {
        var nFacF = String(dF[f][cFactF] || '').trim().toLowerCase();
        if (nFacF) {
          infoFiscal[nFacF] = {
            iva: cIvaF !== -1 ? parseNumeroLatino(dF[f][cIvaF]) : 0,
            baseImponible: cSubF !== -1 ? parseNumeroLatino(dF[f][cSubF]) : 0,
            exento: cExeF !== -1 ? parseNumeroLatino(dF[f][cExeF]) : 0
          };
        }
      }
    }

    var facturas = [];
    var facturasVistas = {};
    var timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'America/Caracas';

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nroFac = String(row[idx.factura] || '').trim();
      var transFila = String(row[idx.transaccion] || '').toLowerCase();

      var matchTipo = esVenta
        ? (transFila.indexOf('vent') === 0 || transFila.indexOf('cobr') === 0)
        : (transFila.indexOf('compr') === 0 || transFila.indexOf('pag') === 0);

      if (!matchTipo || !nroFac) continue;

      // PARSEO SEGURO CON parseNumeroLatino
      var totFila = parseNumeroLatino(row[idx.total]);
      var abono = parseNumeroLatino(row[idx.abono]);
      var tipoMov = String(row[idx.tipoMov] || '').trim();
      var saldoBs = parseNumeroLatino(row[idx.saldoBs]);
      var saldoUSD = parseNumeroLatino(row[idx.saldoUSD]);

      // Identificar fila madre activa
      var esFilaMadre = (totFila > 0 && abono === 0 && !tipoMov);
      if (!esFilaMadre) continue;

      var idUnico = nroFac + '_' + transFila;
      if (facturasVistas[idUnico]) continue;
      facturasVistas[idUnico] = true;

      var fechaStr = '';
      if (row[idx.fecha] instanceof Date) {
        fechaStr = Utilities.formatDate(row[idx.fecha], timeZone, 'yyyy-MM-dd');
      } else {
        fechaStr = String(row[idx.fecha] || '').slice(0, 10);
      }

      var fVencStr = '';
      if (idx.fVenc !== -1 && row[idx.fVenc] instanceof Date) {
        fVencStr = Utilities.formatDate(row[idx.fVenc], timeZone, 'yyyy-MM-dd');
      } else if (idx.fVenc !== -1) {
        fVencStr = String(row[idx.fVenc] || '').slice(0, 10);
      }

      var fiscal = infoFiscal[nroFac.toLowerCase()] || { iva: 0, baseImponible: 0, exento: 0 };
      var tasaFila = parseNumeroLatino(row[idx.tasa]) || 1;

      facturas.push({
        ProveedorCliente: String(row[idx.proveedor] || ''),
        Rif: String(row[idx.rif] || ''),
        NroFactura: nroFac,
        FechaFactura: fechaStr,
        TotalFactura: totFila,
        TotaFactura: totFila,
        SaldoFactura: saldoBs,
        SaldoFacturaUSD: saldoUSD,
        cobradaPagada: String(row[idx.status] || 'Pendiente'),
        Transaccion: String(row[idx.transaccion] || (esVenta ? 'Venta' : 'Compra')),
        Tasa: tasaFila,
        id: String(row[idx.id] || ''),
        BancoDestino: idx.banco !== -1 ? String(row[idx.banco] || '') : '',
        CuentaDestino: idx.cuenta !== -1 ? String(row[idx.cuenta] || '') : '',
        NroCuentaDestino: idx.nroCuenta !== -1 ? String(row[idx.nroCuenta] || '') : '',
        fechaVencimiento: fVencStr,
        nroComprobantesRetIVA: idx.compIVA !== -1 ? String(row[idx.compIVA] || '') : '',
        nroComprobantesRetISLR: idx.compISLR !== -1 ? String(row[idx.compISLR] || '') : '',
        MontoIVA: fiscal.iva,
        BaseImponible: fiscal.baseImponible,
        MontoExento: fiscal.exento
      });
    }

    return facturas;
  } catch (error) {
    Logger.log('Error en getFacturasPorTransaccion: ' + error.message);
    throw new Error(error.message);
  }
}
function anularFacturaUniversal(nroFactura, transaccion, idSpreadsheetEmpresa) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) throw new Error('No se pudo acceder a la empresa activa.');

    var nroBuscado = String(nroFactura).trim().toLowerCase();
    var esVenta = String(transaccion || '').toLowerCase().indexOf('vent') !== -1;
    var shCP = ss.getSheetByName('CobrosPagos');

    if (!shCP) throw new Error('Hoja CobrosPagos no encontrada en la empresa activa.');

    var dataCP = shCP.getDataRange().getValues();
    var pagosActivos = 0;
    var filaMadreCP = -1;

    for (var i = 1; i < dataCP.length; i++) {
      var facCP = String(dataCP[i][2] || '').trim().toLowerCase();
      var totCP = Number(dataCP[i][4]) || 0;
      var abonoCP = Number(dataCP[i][10]) || 0;
      var tipoMovCP = String(dataCP[i][16] || '').trim();

      if (facCP === nroBuscado) {
        if (totCP > 0 && !tipoMovCP) {
          filaMadreCP = i + 1;
        } else if (abonoCP > 0 || tipoMovCP !== '') {
          pagosActivos++;
        }
      }
    }

    if (pagosActivos > 0) {
      throw new Error(
        'No se puede anular la factura N° ' + nroFactura + ' porque tiene ' + pagosActivos +
        ' pago(s)/retención(es) registrado(s). Debe anular primero todos los pagos en el Historial de Abonos.'
      );
    }

    if (filaMadreCP === -1) {
      throw new Error('No se encontró la factura N° ' + nroFactura + ' en CobrosPagos.');
    }

    // Anular en CobrosPagos
    shCP.getRange(filaMadreCP, 6).setValue(0); // SaldoFactura
    shCP.getRange(filaMadreCP, 7).setValue(0); // SaldoFacturaUSD
    shCP.getRange(filaMadreCP, 8).setValue('Anulada'); // cobradaPagada

    // Anular en Ventas o Compras
    if (esVenta) {
      anularFacturaVenta(nroFactura, idSpreadsheetEmpresa);
    } else {
      var shCom = ss.getSheetByName('Compras');
      if (shCom) {
        var dC = shCom.getDataRange().getValues();
        var cFactC = dC[0].findIndex(function(h) { return String(h).toLowerCase().includes('factura'); });
        var cStatC = dC[0].findIndex(function(h) { return String(h).toLowerCase() === 'status' || String(h).toLowerCase() === 'cobradapagada'; });
        var cSaldoC = dC[0].findIndex(function(h) { return String(h).toLowerCase().includes('saldoproveedor'); });

        if (cFactC !== -1) {
          for (var j = 1; j < dC.length; j++) {
            if (String(dC[j][cFactC] || '').trim().toLowerCase() === nroBuscado) {
              if (cStatC !== -1) shCom.getRange(j + 1, cStatC + 1).setValue('Anulada');
              if (cSaldoC !== -1) shCom.getRange(j + 1, cSaldoC + 1).setValue(0);
              break;
            }
          }
        }
      }
    }

    return { ok: true, message: 'Factura ' + nroFactura + ' anulada exitosamente.' };
  } catch (e) {
    Logger.log('[anularFacturaUniversal] Error: ' + e.message);
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function anularPagoIndividual(idPago, nroFactura, idSpreadsheetEmpresa) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);

    var ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) throw new Error('No se pudo acceder a la empresa activa.');

    var shCP = ss.getSheetByName('CobrosPagos');
    if (!shCP) throw new Error('Hoja CobrosPagos no encontrada.');

    var dataCP = shCP.getDataRange().getValues();
    var nroFacNorm = String(nroFactura).trim().toLowerCase();
    var idPagoNorm = String(idPago).trim().toLowerCase();

    var filaPagoIdx = -1;
    var filaMadreIdx = -1;
    var montoAbonoUSD = 0;
    var montoAbonoBs = 0;
    var esVenta = false;
    var tipoMovimiento = '';

    for (var i = 1; i < dataCP.length; i++) {
      var facFila = String(dataCP[i][2] || '').trim().toLowerCase();
      var idFila = String(dataCP[i][9] || '').trim().toLowerCase();
      var compIVA = String(dataCP[i][21] || '').trim().toLowerCase();
      var compISLR = String(dataCP[i][22] || '').trim().toLowerCase();
      var totFila = Number(dataCP[i][4]) || 0;

      if (facFila === nroFacNorm && (idFila === idPagoNorm || compIVA === idPagoNorm || compISLR === idPagoNorm) && totFila === 0) {
        filaPagoIdx = i + 1;
        montoAbonoUSD = Number(dataCP[i][10]) || 0;
        montoAbonoBs = Number(dataCP[i][11]) || (montoAbonoUSD * (Number(dataCP[i][13]) || 1));
        tipoMovimiento = String(dataCP[i][16] || '');
        break;
      }
    }

    if (filaPagoIdx === -1) {
      throw new Error('No se encontró el registro del pago/retención con ID: ' + idPago);
    }

    for (var j = 1; j < dataCP.length; j++) {
      var facM = String(dataCP[j][2] || '').trim().toLowerCase();
      var totM = Number(dataCP[j][4]) || 0;
      var transM = String(dataCP[j][8] || '').toLowerCase();

      if (facM === nroFacNorm && totM > 0) {
        filaMadreIdx = j + 1;
        esVenta = (transM.indexOf('vent') !== -1 || transM.indexOf('cobr') !== -1);
        
        var saldoActualBs = Number(dataCP[j][5]) || 0;
        var saldoActualUSD = Number(dataCP[j][6]) || 0;

        var nuevoSaldoBs = Math.round((saldoActualBs + montoAbonoBs) * 100) / 100;
        var nuevoSaldoUSD = Math.round((saldoActualUSD + montoAbonoUSD) * 100) / 100;

        shCP.getRange(filaMadreIdx, 6).setValue(nuevoSaldoBs);
        shCP.getRange(filaMadreIdx, 7).setValue(nuevoSaldoUSD);
        shCP.getRange(filaMadreIdx, 8).setValue(esVenta ? 'porCobrar' : 'porPagar');

        if (tipoMovimiento === 'RETENCION') {
          if (String(dataCP[j][21]).trim().toLowerCase() === idPagoNorm) shCP.getRange(filaMadreIdx, 22).setValue('');
          if (String(dataCP[j][22]).trim().toLowerCase() === idPagoNorm) shCP.getRange(filaMadreIdx, 23).setValue('');
        }
        break;
      }
    }

    shCP.deleteRow(filaPagoIdx);

    // Sincronizar en hoja Compras o Ventas
    var shFuente = ss.getSheetByName(esVenta ? 'Ventas' : 'Compras');
    if (shFuente) {
      var dF = shFuente.getDataRange().getValues();
      var hF = dF[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
      var cFactF = esVenta ? hF.indexOf('nrofacturacliente') : hF.indexOf('nrofacturaproveedor');
      var cSaldoBsF = esVenta ? hF.indexOf('saldocliente') : hF.indexOf('saldoproveedor');
      var cSaldoUsdF = esVenta ? hF.indexOf('saldoclienteusd') : hF.indexOf('saldoproveedordivisa');
      var cAbonoF = esVenta ? hF.indexOf('abonocliente') : hF.indexOf('abono');
      var cStatusF = hF.indexOf('cobradapagada');

      if (cFactF !== -1) {
        for (var k = 1; k < dF.length; k++) {
          if (String(dF[k][cFactF] || '').trim().toLowerCase() === nroFacNorm) {
            var fIdx = k + 1;
            var sBsAnt = Number(dF[k][cSaldoBsF]) || 0;
            var sUsdAnt = Number(dF[k][cSaldoUsdF]) || 0;
            var abonoAnt = Number(dF[k][cAbonoF]) || 0;

            if (cSaldoBsF !== -1) shFuente.getRange(fIdx, cSaldoBsF + 1).setValue(Math.round((sBsAnt + montoAbonoBs) * 100) / 100);
            if (cSaldoUsdF !== -1) shFuente.getRange(fIdx, cSaldoUsdF + 1).setValue(Math.round((sUsdAnt + montoAbonoUSD) * 100) / 100);
            if (cAbonoF !== -1) shFuente.getRange(fIdx, cAbonoF + 1).setValue(Math.max(0, Math.round((abonoAnt - montoAbonoBs) * 100) / 100));
            if (cStatusF !== -1) shFuente.getRange(fIdx, cStatusF + 1).setValue(esVenta ? 'porCobrar' : 'porPagar');
            break;
          }
        }
      }
    }

    // Eliminar asiento en MovBancoCaja
    var shMB = ss.getSheetByName('MovBancoCaja');
    if (shMB) {
      var dMB = shMB.getDataRange().getValues();
      var idBuscadoMB = 'mov-' + idPagoNorm;
      for (var m = dMB.length - 1; m >= 1; m--) {
        var idMB = String(dMB[m][0] || '').trim().toLowerCase();
        if (idMB === idBuscadoMB || idMB.indexOf(idPagoNorm) !== -1) {
          shMB.deleteRow(m + 1);
          break;
        }
      }
      try { actualizarSaldosBancariosOptimizado(idSpreadsheetEmpresa); } catch (_) {}
    }

    return {
      ok: true,
      mensaje: 'Pago ' + idPago + ' anulado exitosamente. Se restituyeron $' + montoAbonoUSD.toFixed(2) + ' al saldo.'
    };
  } catch (e) {
    Logger.log('[anularPagoIndividual] Error: ' + e.message);
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}
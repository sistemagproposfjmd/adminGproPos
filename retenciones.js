/* ================================================================
   retenciones.gs — VERSIÓN UNIFICADA Y CORREGIDA
   Generación automática + manual de comprobantes de retención
   IVA e ISLR con correlativo único, asientos contables y
   reimpresión desde ComprobantesRetencion y CobrosPagos.
   ================================================================ */

var HOJA_RETENCIONES = 'ComprobantesRetencion';

/* ================================================================
   BLOQUE 1 — HOJA Y CORRELATIVO
   ================================================================ */

function _getHojaRetenciones_() {
  var ss = SpreadsheetApp.getActive();
  var hoja = ss.getSheetByName(HOJA_RETENCIONES);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_RETENCIONES);
    hoja.appendRow([
      'NroComprobante', 'Tipo', 'Fecha', 'RifTercero', 'RazonSocialTercero',
      'NroFactura', 'BaseUSD', 'BaseBs', 'Porcentaje',
      'Concepto', 'MontoRetUSD', 'Tasa', 'MontoRetBs',
      'Usuario', 'Timestamp', 'Estado'
    ]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function _normHeaderRet_(h) {
  return String(h || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[\s_\-#.]/g, '');
}

function _fmtPeriodoRet_(fecha) {
  var d = (fecha instanceof Date) ? fecha : new Date(fecha + 'T00:00:00');
  if (isNaN(d.getTime())) d = new Date();
  return String(d.getFullYear()) + String(d.getMonth() + 1).padStart(2, '0');
}

function _largoSecuencialRet_() {
  return 7;
}

/**
 * Correlativo único por tipo y periodo YYYYMM.
 * Busca en: ComprobantesRetencion (auto), Compras, CobrosPagos, Asientos.
 * REQUIERE LockService del caller cuando skipLock=false.
 */
function generarCorrelativoDesdeAbonos(tipo, fechaOperacion, skipLock) {
  tipo = String(tipo || 'IVA').toUpperCase();
  if (tipo !== 'IVA' && tipo !== 'ISLR') {
    throw new Error('Tipo de retención inválido: ' + tipo);
  }

  var lock = null;
  if (!skipLock) {
    lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (e) {
      Logger.log('[CORRELATIVO RET] Lock ya tomado por caller: ' + e.message);
      lock = null;
    }
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prefijo = _fmtPeriodoRet_(fechaOperacion || new Date());
    var largo = _largoSecuencialRet_();
    var patronComp = new RegExp('^' + prefijo + '(\\d{6,8})$');
    var maxSec = 0;

    // ---- Fuente 0: ComprobantesRetencion (flujo automático) ----
    var shCR = ss.getSheetByName(HOJA_RETENCIONES);
    if (shCR && shCR.getLastRow() > 1) {
      var dataCR = shCR.getDataRange().getValues();
      var hCR = dataCR[0].map(_normHeaderRet_);
      var cNro = hCR.indexOf('nrocomprobante');
      var cTipo = hCR.indexOf('tipo');
      if (cNro !== -1) {
        for (var cr = 1; cr < dataCR.length; cr++) {
          var tipoCR = String(dataCR[cr][cTipo] || '').toUpperCase();
          if (tipoCR && tipoCR !== tipo) continue;
          var nroCR = String(dataCR[cr][cNro] || '').trim();
          var mCR = nroCR.match(patronComp);
          if (mCR) {
            var nCR = parseInt(mCR[1], 10);
            if (nCR > maxSec) maxSec = nCR;
          }
        }
      }
    }

    // ---- Fuente 1: Compras ----
    var shCompras = ss.getSheetByName('Compras');
    if (shCompras && shCompras.getLastRow() > 1) {
      var dataCom = shCompras.getDataRange().getValues();
      var hCom = dataCom[0].map(_normHeaderRet_);
      var colNroCompCom = hCom.indexOf('nrocomprobantes');
      if (colNroCompCom !== -1) {
        for (var c = 1; c < dataCom.length; c++) {
          var compCom = String(dataCom[c][colNroCompCom] || '').trim();
          var mCom = compCom.match(patronComp);
          if (mCom) {
            var nCom = parseInt(mCom[1], 10);
            if (nCom > maxSec) maxSec = nCom;
          }
        }
      }
    }

    // ---- Fuente 2: CobrosPagos ----
    var shCP = ss.getSheetByName('CobrosPagos');
    if (shCP && shCP.getLastRow() > 1) {
      var dataCP = shCP.getDataRange().getValues();
      var hCP = dataCP[0].map(_normHeaderRet_);
      var colIvaCP = hCP.indexOf('nrocomprobantesretiva');
      var colIslrCP = hCP.indexOf('nrocomprobantesretislr');
      var colTransCP = hCP.indexOf('transaccion');
      var colTipoMovCP = hCP.indexOf('tipomovimiento');
      var colComp = (tipo === 'ISLR') ? colIslrCP : colIvaCP;
      if (colComp === -1) colComp = colTransCP;

      for (var i = 1; i < dataCP.length; i++) {
        if (colTipoMovCP !== -1) {
          var tm = String(dataCP[i][colTipoMovCP] || '').toUpperCase();
          if (tm !== 'RETENCION') continue;
        }
        var comp = String(dataCP[i][colComp] || '').trim();
        var m = comp.match(patronComp);
        if (m) {
          var n = parseInt(m[1], 10);
          if (n > maxSec) maxSec = n;
        }
      }
    }

    // ---- Fuente 3: Asientos ----
    var shAs = ss.getSheetByName('Asientos');
    if (shAs && shAs.getLastRow() > 1) {
      var dataAs = shAs.getDataRange().getValues();
      var hAs = dataAs[0].map(_normHeaderRet_);
      var colRefAs = hAs.indexOf('transaccionref');
      var colDescAs = hAs.indexOf('descripciongeneral');
      if (colRefAs === -1) colRefAs = 9;
      if (colDescAs === -1) colDescAs = 4;

      for (var j = 1; j < dataAs.length; j++) {
        var desc = String(dataAs[j][colDescAs] || '').toUpperCase();
        var ref  = String(dataAs[j][colRefAs]  || '').toUpperCase();
        var esDelTipo = (tipo === 'IVA')
          ? (desc.indexOf('RETENCIÓN IVA') !== -1 || desc.indexOf('RETENCION IVA') !== -1)
          : (desc.indexOf('RETENCIÓN ISLR') !== -1 || desc.indexOf('RETENCION ISLR') !== -1);
        if (!esDelTipo) continue;
        var mAs = ref.match(patronComp);
        if (mAs) {
          var nAs = parseInt(mAs[1], 10);
          if (nAs > maxSec) maxSec = nAs;
        }
      }
    }

    var siguiente = String(maxSec + 1).padStart(largo, '0');
    var nroFinal = prefijo + siguiente;
    Logger.log('[CORRELATIVO RET] ' + tipo + ' -> ' + nroFinal);
    return nroFinal;
  } catch (e) {
    Logger.log('[CORRELATIVO RET] ERROR: ' + e.message);
    throw new Error('No se pudo generar el correlativo: ' + e.message);
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (_) {} }
  }
}

/**
 * Delegado: usa la fuente unificada (ComprobantesRetencion + Compras + CP + Asientos).
 * IMPORTANTE: el caller debe tener el lock tomado.
 */
function generarCorrelativoRetencion(tipo, anio) {
  var fecha = new Date(anio || new Date().getFullYear(), 0, 1);
  return generarCorrelativoDesdeAbonos(tipo, fecha, true);
}

function obtenerSiguienteCorrelativoRetencion(tipo, fechaOperacion) {
  try {
    var nro = generarCorrelativoDesdeAbonos(tipo, fechaOperacion, false);
    return {
      ok: true,
      nroComprobante: nro,
      tipo: String(tipo || 'IVA').toUpperCase(),
      fecha: fechaOperacion || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
    };
  } catch (e) {
    Logger.log('[obtenerSiguienteCorrelativoRetencion] ' + e.message);
    throw new Error(e.message);
  }
}

function getProximoNroComprobante(tipo, fechaOperacion) {
  return generarCorrelativoDesdeAbonos(tipo, fechaOperacion, false);
}

function previewCorrelativoRetencionCompras(tipo, fecha) {
  try {
    // skipLock=true → NO bloquea al preview
    return generarCorrelativoDesdeAbonos(tipo || 'IVA', fecha || new Date(), true);
  } catch (e) {
    Logger.log('Error en previewCorrelativo: ' + e.message);
    return '';
  }
}

/* ================================================================
   BLOQUE 2 — GENERAR COMPROBANTE Y REGISTRAR EN FACTURA
   ================================================================ */

/**
 * Genera comprobante en ComprobantesRetencion y devuelve el Nro.
 * REQUIERE lock del caller.
 */
function generarComprobanteRetencion(datos) {
  datos.tipo = String(datos.tipo || 'IVA').toUpperCase();
  if (datos.tipo !== 'IVA' && datos.tipo !== 'ISLR') {
    throw new Error('Tipo de retención inválido: ' + datos.tipo);
  }

  var hoja = _getHojaRetenciones_();
  var anio = new Date(datos.fecha || new Date()).getFullYear();
  var nroComprobante = generarCorrelativoRetencion(datos.tipo, anio);

  hoja.appendRow([
    nroComprobante,
    datos.tipo,
    datos.fecha,
    datos.rif || '',
    datos.tercero || '',
    datos.factura || '',
    Number(datos.baseUSD || 0),
    Number(datos.baseBs || 0),
    Number(datos.porcentaje || 0),
    datos.concepto || '',
    Number(datos.montoRetUSD || 0),
    Number(datos.tasa || 1),
    Number(datos.montoRetBs || 0),
    Session.getActiveUser().getEmail(),
    new Date(),
    'Vigente'
  ]);

  return nroComprobante;
}

function registrarComprobanteEnFactura(nroFactura, tipo, nroComprobante) {
  var ss = SpreadsheetApp.getActive();
  var hoja = ss.getSheetByName('Facturas');
  if (!hoja) return;

  var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var colNro    = headers.indexOf('NroFactura') + 1;
  var colTarget = headers.indexOf(tipo === 'ISLR' ? 'nroComprobantesRetISLR' : 'nroComprobantesRetIVA') + 1;

  if (colNro < 1 || colTarget < 1) return;

  var data = hoja.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colNro - 1]) === String(nroFactura)) {
      hoja.getRange(i + 1, colTarget).setValue(nroComprobante);
      return;
    }
  }
}

/* ================================================================
   BLOQUE 3 — DISTRIBUCIÓN DE RETENCIONES POR FACTURA
   ================================================================ */

function _distribuirRetencion_(facturas, baseKey, montoTotal) {
  var elegibles = facturas.filter(function (f) {
    var bandera = (baseKey === 'baseIvaUSD') ? f.tieneRetIVA : f.tieneRetISLR;
    return !bandera && Number(f[baseKey] || 0) > 0;
  });
  if (!elegibles.length || montoTotal <= 0) return [];

  var baseTotal = elegibles.reduce(function (s, f) { return s + Number(f[baseKey] || 0); }, 0);
  if (baseTotal <= 0) return [];

  // Calcular montos con redondeo y ajustar la última con MAYOR base (no la última del array)
  var resultado = elegibles.map(function (f) {
    var proporcion = Number(f[baseKey] || 0) / baseTotal;
    return {
      factura: f.nroFactura,
      baseUSD: Number(f[baseKey] || 0),
      montoRetUSD: Math.round(montoTotal * proporcion * 100) / 100
    };
  });

  // Ajuste final: la diferencia va a la factura con mayor base
  var suma = resultado.reduce(function (s, r) { return s + r.montoRetUSD; }, 0);
  var diff = Math.round((montoTotal - suma) * 100) / 100;
  if (Math.abs(diff) >= 0.01) {
    var idxMayor = 0;
    for (var k = 1; k < resultado.length; k++) {
      if (resultado[k].baseUSD > resultado[idxMayor].baseUSD) idxMayor = k;
    }
    resultado[idxMayor].montoRetUSD = Math.round((resultado[idxMayor].montoRetUSD + diff) * 100) / 100;
  }

  return resultado;
}

/* ================================================================
   BLOQUE 4 — VALIDACIONES
   ================================================================ */

function validarComprobanteUnico(nroComprobante, tipoRetencion) {
  try {
    if (!nroComprobante) return false;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tipo = String(tipoRetencion || 'IVA').toUpperCase();
    var nro = String(nroComprobante).trim();

    // ---- ComprobantesRetencion (respetar Estado=Anulado) ----
    var shCR = ss.getSheetByName(HOJA_RETENCIONES);
    if (shCR && shCR.getLastRow() > 1) {
      var dataCR = shCR.getDataRange().getValues();
      var hCR = dataCR[0].map(_normHeaderRet_);
      var cNro  = hCR.indexOf('nrocomprobante');
      var cTipo = hCR.indexOf('tipo');
      var cEst  = hCR.indexOf('estado');
      for (var i = 1; i < dataCR.length; i++) {
        if (String(dataCR[i][cNro] || '').trim() !== nro) continue;
        var est = String(dataCR[i][cEst] || '').toUpperCase();
        if (est === 'ANULADO') continue;
        var t = String(dataCR[i][cTipo] || '').toUpperCase();
        if (!t || t === tipo) {
          Logger.log('[validarComprobanteUnico] DUPLICADO en ComprobantesRetencion fila ' + (i + 1));
          return false;
        }
      }
    }

    // ---- CobrosPagos ----
    var shCP = ss.getSheetByName('CobrosPagos');
    if (shCP && shCP.getLastRow() > 1) {
      var dataCP = shCP.getDataRange().getValues();
      var hCP = dataCP[0].map(_normHeaderRet_);
      var colIva = hCP.indexOf('nrocomprobantesretiva');
      var colIslr = hCP.indexOf('nrocomprobantesretislr');
      var col = (tipo === 'ISLR') ? colIslr : colIva;
      if (col !== -1) {
        for (var j = 1; j < dataCP.length; j++) {
          if (String(dataCP[j][col] || '').trim() === nro) {
            Logger.log('[validarComprobanteUnico] DUPLICADO en CobrosPagos fila ' + (j + 1));
            return false;
          }
        }
      }
    }

    // ---- Asientos ----
    var shAs = ss.getSheetByName('Asientos');
    if (shAs && shAs.getLastRow() > 1) {
      var dataAs = shAs.getDataRange().getValues();
      var hAs = dataAs[0].map(_normHeaderRet_);
      var colRef = hAs.indexOf('transaccionref');
      if (colRef === -1) colRef = 9;
      for (var k = 1; k < dataAs.length; k++) {
        if (String(dataAs[k][colRef] || '').trim() === nro) {
          Logger.log('[validarComprobanteUnico] DUPLICADO en Asientos fila ' + (k + 1));
          return false;
        }
      }
    }

    return true;
  } catch (error) {
    Logger.log('[validarComprobanteUnico] ERROR (falla cerrada): ' + error.message);
    return false;
  }
}

function validarNoExisteRetencion(nroComprobante, tipoRetencion, nroFactura) {
  if (!nroComprobante || typeof nroComprobante !== 'string') {
    throw new Error('N° de comprobante requerido para validar');
  }
  nroComprobante = String(nroComprobante).trim().toUpperCase();
  tipoRetencion = String(tipoRetencion || 'IVA').toUpperCase();
  nroFactura = String(nroFactura || '').trim();

  Logger.log('[VALIDAR RET] ' + nroComprobante + ' | ' + tipoRetencion + ' | ' + nroFactura);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var encontrados = [];

  // Asientos
  var hojaAsientos = ss.getSheetByName('Asientos');
  if (hojaAsientos && hojaAsientos.getLastRow() > 1) {
    var dataA = hojaAsientos.getDataRange().getValues();
    for (var i = 1; i < dataA.length; i++) {
      var desc = String(dataA[i][4] || '').toUpperCase();
      var ref  = String(dataA[i][9] || '').toUpperCase();
      if ((desc.indexOf(nroComprobante) !== -1 || ref.indexOf(nroComprobante) !== -1) &&
          (desc.indexOf(tipoRetencion) !== -1 || desc.indexOf('RET') !== -1)) {
        encontrados.push({ ubicacion: 'Asientos', fila: i + 1, detalle: desc.substring(0, 80) });
      }
    }
  }

  // CobrosPagos
  var hojaCP = ss.getSheetByName('CobrosPagos');
  if (hojaCP && hojaCP.getLastRow() > 1) {
    var dataCP = hojaCP.getDataRange().getValues();
    var hCP = dataCP[0].map(_normHeaderRet_);
    var cIva  = hCP.indexOf('nrocomprobantesretiva');
    var cIslr = hCP.indexOf('nrocomprobantesretislr');
    var cTrans = hCP.indexOf('transaccion');
    var cFact  = hCP.indexOf('nrofactura');
    for (var j = 1; j < dataCP.length; j++) {
      var trans = String(dataCP[j][cTrans] || '').toUpperCase();
      var fact  = String(dataCP[j][cFact]  || '').toUpperCase();
      if (trans.indexOf(nroComprobante) === -1) continue;
      var cTipo = (tipoRetencion === 'IVA') ? cIva : cIslr;
      if (cTipo !== -1 && String(dataCP[j][cTipo] || '').toUpperCase().indexOf(nroComprobante) !== -1) {
        encontrados.push({ ubicacion: 'CobrosPagos', fila: j + 1, detalle: 'Retención ' + tipoRetencion + ' - Factura: ' + fact });
      }
    }
  }

  if (encontrados.length > 0) {
    var detalles = encontrados.map(function (e) {
      return e.ubicacion + ' (Fila ' + e.fila + '): ' + e.detalle;
    }).join('\n');
    Logger.log('[VALIDAR RET] ⚠️ DUPLICADO:\n' + detalles);
    return { existe: true, detalles: detalles, cantidad: encontrados.length };
  }

  return { existe: false, detalles: 'Sin duplicados', cantidad: 0 };
}

/* ================================================================
   BLOQUE 5 — PROCESAR LIQUIDACIÓN COMPLETA (INTEGRADO)
   ================================================================ */

function procesarLiquidacionCompleta(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var ss = SpreadsheetApp.getActive();
    var hojaFacturas = ss.getSheetByName('Facturas');
    var hojaAbonos   = ss.getSheetByName('AbonosFacturas') || ss.getSheetByName('Abonos');
    if (!hojaFacturas) throw new Error('No existe la hoja Facturas');
    if (!hojaAbonos)   throw new Error('No existe la hoja de Abonos');

    var tasa    = Number(payload.tasa || 1) || 1;
    var fecha   = payload.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var esCobro = String(payload.tipoMovimiento || '').toLowerCase() === 'cobro';

    var comprobantesGenerados = [];

    // 1) Anticipos seleccionados
    if (Array.isArray(payload.anticiposSeleccionados) && payload.anticiposSeleccionados.length) {
      _aplicarAnticipos_(payload.anticiposSeleccionados, payload, fecha, tasa, hojaAbonos);
    }

    // 2) Anticipo directo → crear nuevo anticipo
    if (payload.esAnticipoDirecto) {
      _crearAnticipo_(payload, fecha, tasa);
    }

    // 3) Movimientos bancarios
    if (Array.isArray(payload.detallesPago) && payload.detallesPago.length) {
      _registrarMovimientosBancarios_(payload.detallesPago, payload, esCobro, fecha, tasa);
    }

    // 4) Abonos por factura (solo el monto en banco)
    if (!payload.esAnticipoDirecto && Array.isArray(payload.facturasSeleccionadas)) {
      _registrarAbonosPorFactura_(payload.facturasSeleccionadas, payload, esCobro, fecha, tasa);
    }

    // 5) RETENCIÓN IVA — automática
    if (Number(payload.montoRetIVA) > 0 && Array.isArray(payload.facturasSeleccionadas)) {
      var distIVA = _distribuirRetencion_(payload.facturasSeleccionadas, 'baseIvaUSD', Number(payload.montoRetIVA));
      distIVA.forEach(function (item) {
        _procesarUnaRetencion_({
          tipo: 'IVA',
          porcentaje: 75,
          concepto: 'Retención IVA',
          item: item,
          payload: payload,
          fecha: fecha,
          tasa: tasa,
          esCobro: esCobro,
          hojaAbonos: hojaAbonos,
          comprobantesGenerados: comprobantesGenerados
        });
      });
    }

    // 6) RETENCIÓN ISLR — automática
    if (Number(payload.montoRetISLR) > 0 && Array.isArray(payload.facturasSeleccionadas)) {
      var distISLR = _distribuirRetencion_(payload.facturasSeleccionadas, 'baseIslrUSD', Number(payload.montoRetISLR));
      var pctISLR = Number(payload.porcentajeISLR || 0);
      distISLR.forEach(function (item) {
        _procesarUnaRetencion_({
          tipo: 'ISLR',
          porcentaje: pctISLR,
          concepto: payload.conceptoISLR || 'Retención ISLR',
          item: item,
          payload: payload,
          fecha: fecha,
          tasa: tasa,
          esCobro: esCobro,
          hojaAbonos: hojaAbonos,
          comprobantesGenerados: comprobantesGenerados
        });
      });
    }

    // 7) Recalcular saldos de facturas afectadas
    if (Array.isArray(payload.facturasSeleccionadas)) {
      payload.facturasSeleccionadas.forEach(function (f) {
        _recalcularSaldoFactura_(hojaFacturas, f.nroFactura);
      });
    }

    return {
      ok: true,
      mensaje: 'Liquidación procesada. ' + comprobantesGenerados.length + ' comprobante(s) generado(s).',
      comprobantes: comprobantesGenerados
    };
  } catch (err) {
    throw new Error('procesarLiquidacionCompleta: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

function _procesarUnaRetencion_(ctx) {
  var item = ctx.item;
  var payload = ctx.payload;

  // a) Generar comprobante en ComprobantesRetencion
  var nro = generarComprobanteRetencion({
    tipo: ctx.tipo,
    fecha: ctx.fecha,
    rif: payload.rif,
    tercero: payload.tercero,
    factura: item.factura,
    baseUSD: item.baseUSD,
    baseBs: +(item.baseUSD * ctx.tasa).toFixed(2),
    porcentaje: ctx.porcentaje,
    concepto: ctx.concepto,
    montoRetUSD: item.montoRetUSD,
    tasa: ctx.tasa,
    montoRetBs: +(item.montoRetUSD * ctx.tasa).toFixed(2)
  });
  registrarComprobanteEnFactura(item.factura, ctx.tipo, nro);
  ctx.comprobantesGenerados.push(nro);

  // b) Registrar en AbonosFacturas (para que el saldo se descuente)
  _registrarAbonoRetencion_(ctx.hojaAbonos, {
    fecha: ctx.fecha,
    factura: item.factura,
    tercero: payload.tercero,
    rif: payload.rif,
    tipoMovimiento: 'Retención ' + ctx.tipo,
    montoUSD: item.montoRetUSD,
    montoBS: +(item.montoRetUSD * ctx.tasa).toFixed(2),
    descripcion: 'Retención ' + ctx.tipo + ' factura ' + item.factura +
                 (ctx.tipo === 'ISLR' && payload.conceptoISLR ? ' — ' + payload.conceptoISLR : ''),
    comprobanteIVA: ctx.tipo === 'IVA' ? nro : '',
    comprobanteISLR: ctx.tipo === 'ISLR' ? nro : ''
  });

  // c) Registrar en CobrosPagos (historial + descuento en Ventas/Compras)
  try {
    asentarRetencionEnCobrosPagos(SpreadsheetApp.getActive(), {
      tipoRetencion: ctx.tipo,
      nroComprobante: nro,
      fecha: ctx.fecha,
      tasa: ctx.tasa,
      nombreEntidad: payload.tercero,
      rifEntidad: payload.rif,
      esVenta: ctx.esCobro,
      facturas: [{
        nroFactura: item.factura,
        montoRetenidoUSD: item.montoRetUSD,
        montoRetenidoBs: +(item.montoRetUSD * ctx.tasa).toFixed(2)
      }]
    });
  } catch (eCP) {
    Logger.log('[RET ' + ctx.tipo + '] Error asentando en CobrosPagos: ' + eCP.message);
  }

  // d) Asiento contable
  try {
    var nroAsiento = generarCorrelativoAsiento(true);
    var contra = resolverContrapartidaCxCp(ctx.esCobro);
    var lineas = construirAsientoRetencion(
      ctx.esCobro, nroAsiento, ctx.fecha, payload.tercero,
      'Retención ' + ctx.tipo + ' Comprobante ' + nro, nro,
      Math.round(item.montoRetUSD * ctx.tasa * 100) / 100,
      ctx.tipo, contra
    );
    guardarAsientoMasivo(lineas);
  } catch (eAs) {
    Logger.log('[RET ' + ctx.tipo + '] Error generando asiento: ' + eAs.message);
  }
}

/* ================================================================
   BLOQUE 6 — HELPERS INTERNOS
   ================================================================ */

function _aplicarAnticipos_(anticipos, payload, fecha, tasa, hojaAbonos) {
  var hoja = SpreadsheetApp.getActive().getSheetByName('Anticipos');
  if (!hoja) return;

  var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var colId    = headers.indexOf('idAnticipo') + 1;
  var colSaldo = headers.indexOf('saldoUSD') + 1;
  if (colId < 1 || colSaldo < 1) return;

  var data = hoja.getDataRange().getValues();
  anticipos.forEach(function (a) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colId - 1]) === String(a.idAnticipo)) {
        var nuevo = Number(data[i][colSaldo - 1] || 0) - Number(a.montoUSD || 0);
        hoja.getRange(i + 1, colSaldo).setValue(Math.max(0, nuevo));

        // Registrar en AbonosFacturas para que el saldo de la factura se descuente
        if (hojaAbonos && payload.facturasSeleccionadas && payload.facturasSeleccionadas.length) {
          var fac = payload.facturasSeleccionadas[0];
          _registrarAbonoRetencion_(hojaAbonos, {
            fecha: fecha,
            factura: fac.nroFactura,
            tercero: payload.tercero,
            rif: payload.rif,
            tipoMovimiento: 'Anticipo',
            montoUSD: Number(a.montoUSD || 0),
            montoBS: +(Number(a.montoUSD || 0) * tasa).toFixed(2),
            descripcion: 'Anticipo ' + a.idAnticipo + ' aplicado a factura ' + fac.nroFactura,
            comprobanteIVA: '',
            comprobanteISLR: ''
          });
        }
        break;
      }
    }
  });
}

function _crearAnticipo_(payload, fecha, tasa) {
  var ss = SpreadsheetApp.getActive();
  var hoja = ss.getSheetByName('Anticipos');
  if (!hoja) {
    hoja = ss.insertSheet('Anticipos');
    hoja.appendRow(['idAnticipo','Fecha','Tercero','RIF','MontoUSD','MontoBS','Tasa','SaldoUSD','Estado']);
  }
  var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var id = 'ANT-' + new Date().getTime();
  var monto = Number(payload.totalAplicadoUSD || 0);

  var fila = new Array(headers.length).fill('');
  var idx = function (name) { return headers.indexOf(name); };
  if (idx('idAnticipo') !== -1) fila[idx('idAnticipo')] = id;
  if (idx('Fecha') !== -1) fila[idx('Fecha')] = fecha;
  if (idx('Tercero') !== -1) fila[idx('Tercero')] = payload.tercero;
  if (idx('RIF') !== -1) fila[idx('RIF')] = payload.rif;
  if (idx('MontoUSD') !== -1) fila[idx('MontoUSD')] = monto;
  if (idx('MontoBS') !== -1) fila[idx('MontoBS')] = +(monto * tasa).toFixed(2);
  if (idx('Tasa') !== -1) fila[idx('Tasa')] = tasa;
  if (idx('SaldoUSD') !== -1) fila[idx('SaldoUSD')] = monto;
  if (idx('Estado') !== -1) fila[idx('Estado')] = 'Disponible';

  hoja.appendRow(fila);
  return id;
}

function _registrarMovimientosBancarios_(detalles, payload, esCobro, fecha, tasa) {
  var hoja = SpreadsheetApp.getActive().getSheetByName('MovimientosBancarios');
  if (!hoja) return;
  detalles.forEach(function (d) {
    var partes = String(d.cuenta || '').split('|');
    hoja.appendRow([
      fecha, partes[0] || '', partes[1] || '', partes[2] || '',
      partes[3] || '', partes[4] || '',
      esCobro ? 'Ingreso' : 'Egreso',
      d.moneda, d.monto, d.referencia,
      payload.tercero, payload.rif,
      new Date()
    ]);
  });
}

function _registrarAbonosPorFactura_(facturas, payload, esCobro, fecha, tasa) {
  var hoja = SpreadsheetApp.getActive().getSheetByName('AbonosFacturas')
          || SpreadsheetApp.getActive().getSheetByName('Abonos');
  if (!hoja) return;

  var totalSaldo = facturas.reduce(function (s, f) { return s + Number(f.saldoUSD || 0); }, 0);
  if (totalSaldo <= 0) return;

  var montoBanco = Number(payload.totalBancoUSD || 0);

  facturas.forEach(function (f) {
    var proporcion = Number(f.saldoUSD || 0) / totalSaldo;
    var monto = +(montoBanco * proporcion).toFixed(2);
    if (monto <= 0) return;

    hoja.appendRow([
      fecha, f.nroFactura, payload.tercero, payload.rif,
      esCobro ? 'Cobro' : 'Pago', monto, +(monto * tasa).toFixed(2),
      'Abono factura ' + f.nroFactura,
      '', '', new Date(), Session.getActiveUser().getEmail()
    ]);
  });
}

function _registrarAbonoRetencion_(hoja, datos) {
  hoja.appendRow([
    datos.fecha, datos.factura, datos.tercero, datos.rif,
    datos.tipoMovimiento, datos.montoUSD, datos.montoBS,
    datos.descripcion,
    datos.comprobanteIVA || '',
    datos.comprobanteISLR || '',
    new Date(), Session.getActiveUser().getEmail()
  ]);
}

function _recalcularSaldoFactura_(hojaFacturas, nroFactura) {
  var headers = hojaFacturas.getRange(1, 1, 1, hojaFacturas.getLastColumn()).getValues()[0];
  var colNro    = headers.indexOf('NroFactura') + 1;
  var colTotal  = headers.indexOf('TotalFactura') + 1;
  if (colTotal < 1) colTotal = headers.indexOf('TotaFactura') + 1;
  var colSaldo  = headers.indexOf('SaldoFactura') + 1;
  var colSaldoU = headers.indexOf('SaldoFacturaUSD') + 1;
  var colTasa   = headers.indexOf('Tasa') + 1;

  var data = hojaFacturas.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colNro - 1]) !== String(nroFactura)) continue;

    var totalBs = Number(data[i][colTotal - 1] || 0);
    var tasa    = Number(data[i][colTasa - 1] || 1) || 1;

    var hojaAb = SpreadsheetApp.getActive().getSheetByName('AbonosFacturas')
              || SpreadsheetApp.getActive().getSheetByName('Abonos');
    var pagadoUSD = 0;
    if (hojaAb && hojaAb.getLastRow() > 1) {
      var hAb = hojaAb.getRange(1, 1, 1, hojaAb.getLastColumn()).getValues()[0];
      var cNro = hAb.indexOf('NroFactura') + 1;
      var cUSD = hAb.indexOf('MontoUSD') + 1;
      if (cNro > 0 && cUSD > 0) {
        var dAb = hojaAb.getDataRange().getValues();
        for (var j = 1; j < dAb.length; j++) {
          if (String(dAb[j][cNro - 1]) === String(nroFactura)) {
            pagadoUSD += Number(dAb[j][cUSD - 1] || 0);
          }
        }
      }
    }

    var saldoUSD = Math.max(0, (totalBs / tasa) - pagadoUSD);
    var saldoBs  = +(saldoUSD * tasa).toFixed(2);

    if (colSaldo  > 0) hojaFacturas.getRange(i + 1, colSaldo).setValue(saldoBs);
    if (colSaldoU > 0) hojaFacturas.getRange(i + 1, colSaldoU).setValue(saldoUSD);
    break;
  }
}

/* ================================================================
   BLOQUE 7 — REGISTRO MANUAL (registrarMultiplesRetenciones)
   ================================================================ */

function registrarMultiplesRetenciones(p) {
  var lock = LockService.getScriptLock();
  var filasEscritas = 0;
  var shCPref = null;
  var primeraFilaEscrita = -1;

  try {
    lock.waitLock(30000);

    if (!p) throw new Error('Payload vacío.');
    if (!p.facturas || p.facturas.length === 0) throw new Error('Debe incluir al menos una factura.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shCP = ss.getSheetByName('CobrosPagos');
    var shAsientos = ss.getSheetByName('Asientos');
    if (!shCP) throw new Error('No existe la hoja CobrosPagos.');
    if (!shAsientos) throw new Error('No existe la hoja Asientos.');
    shCPref = shCP;

    var data = shCP.getDataRange().getValues();
    var tasa = Number(p.tasa) || 1;
    var esVenta = (p.tipo === 'VENTA' || p.tipo === 'CXC');
    var tipoRet = String(p.tipoRetencion || 'IVA').toUpperCase();
    var campoCompCol = (tipoRet === 'IVA')
      ? CXC_COL_CP.NRO_COMPROBANTES_RET_IVA
      : CXC_COL_CP.NRO_COMPROBANTES_RET_ISLR;

    // Correlativo: si no viene, generarlo
    var nroComp = String(p.nroComprobante || '').trim();
    if (!nroComp) {
      nroComp = generarCorrelativoDesdeAbonos(tipoRet, p.fecha, true);
    }

    // Validar unicidad (también para ISLR)
    if (!validarComprobanteUnico(nroComp, tipoRet)) {
      throw new Error('El comprobante ' + nroComp + ' ya existe. No se puede duplicar.');
    }

    // Enriquecer NroControl
    p.facturas.forEach(function (f) {
      if (!f.nroControl || f.nroControl === '00-00000000') {
        f.nroControl = obtenerNroControlReal(ss, f.nroFactura, esVenta);
      }
    });

    // Indexar CobrosPagos
    var idxPorFactura = {};
    for (var r = 1; r < data.length; r++) {
      var facKey = String(data[r][CXC_COL_CP.NRO_FACTURA] || '').trim().toLowerCase();
      var provKey = String(data[r][CXC_COL_CP.PROVEEDOR_CLIENTE] || '').trim().toLowerCase();
      if (facKey && provKey) idxPorFactura[facKey + '|' + provKey] = r;
    }

    // Hoja fuente para actualizar saldos
    var hojaFuente = ss.getSheetByName(esVenta ? 'Ventas' : 'Compras');
    var dataFuente = null, hFuente = null;
    if (hojaFuente) {
      dataFuente = hojaFuente.getDataRange().getValues();
      hFuente = dataFuente[0].map(_normHeaderRet_);
    }
    var colFactF = esVenta ? hFuente.indexOf('nrofacturacliente') : hFuente.indexOf('nrofacturaproveedor');
    var colSaldoBsF = esVenta ? hFuente.indexOf('saldocliente') : hFuente.indexOf('saldoproveedor');
    var colSaldoUsdF = esVenta ? hFuente.indexOf('saldoclienteusd') : hFuente.indexOf('saldoproveedordivisa');
    var colStatusF = hFuente.indexOf('cobradapagada');
    var colAbonoF = esVenta ? hFuente.indexOf('abonocliente') : hFuente.indexOf('abono');

    var filasNuevas = [];
    var totalRetBs = 0;

    p.facturas.forEach(function (f) {
      var nroFacNormal = String(f.nroFactura).trim().toLowerCase();
      var provNormal = String(p.nombreEntidad).trim().toLowerCase();
      var retUSD = Number(f.montoRetenidoUSD) || 0;
      var retBs = Number(f.montoRetenidoBs) || (retUSD * tasa);
      totalRetBs += retBs;

      // Actualizar fila existente en CobrosPagos
      var rIdx = idxPorFactura[nroFacNormal + '|' + provNormal];
      if (rIdx !== undefined) {
        var saldoActualBs = Number(data[rIdx][CXC_COL_CP.SALDO_FACTURA]) || 0;
        var saldoActualUSD = Number(data[rIdx][CXC_COL_CP.SALDO_FACTURA_USD]) || 0;
        var nuevoSaldoBs = Math.max(0, saldoActualBs - retBs);
        var nuevoSaldoUSD = Math.max(0, saldoActualUSD - retUSD);

        shCP.getRange(rIdx + 1, CXC_COL_CP.SALDO_FACTURA + 1).setValue(nuevoSaldoBs);
        shCP.getRange(rIdx + 1, CXC_COL_CP.SALDO_FACTURA_USD + 1).setValue(nuevoSaldoUSD);
        shCP.getRange(rIdx + 1, campoCompCol + 1).setValue(nroComp);
        if (nuevoSaldoBs <= 0.01) {
          shCP.getRange(rIdx + 1, CXC_COL_CP.COBRADA_PAGADA + 1).setValue(esVenta ? 'Cobrada' : 'Pagada');
        }
      }

      // Actualizar Ventas / Compras
      if (dataFuente && colFactF !== -1) {
        for (var w = 1; w < dataFuente.length; w++) {
          if (String(dataFuente[w][colFactF] || '').trim().toLowerCase() === nroFacNormal) {
            var nvoBsF = Math.max(0, Math.round(((Number(dataFuente[w][colSaldoBsF]) || 0) - retBs) * 100) / 100);
            hojaFuente.getRange(w + 1, colSaldoBsF + 1).setValue(nvoBsF);
            if (colSaldoUsdF !== -1) {
              var nvoUsdF = Math.max(0, Math.round(((Number(dataFuente[w][colSaldoUsdF]) || 0) - retUSD) * 100) / 100);
              hojaFuente.getRange(w + 1, colSaldoUsdF + 1).setValue(nvoUsdF);
            }
            if (colStatusF !== -1) {
              hojaFuente.getRange(w + 1, colStatusF + 1).setValue(nvoBsF <= 0.01 ? (esVenta ? 'Cobrada' : 'Pagada') : 'Abono - Por Cobrar');
            }
            if (colAbonoF !== -1) {
              var abonoF = Math.round(((Number(dataFuente[w][colAbonoF]) || 0) + retBs) * 100) / 100;
              hojaFuente.getRange(w + 1, colAbonoF + 1).setValue(abonoF);
            }
            break;
          }
        }
      }

      // Nueva fila histórica en CobrosPagos
      var filaRet = new Array(23).fill('');
      filaRet[CXC_COL_CP.PROVEEDOR_CLIENTE] = p.nombreEntidad;
      filaRet[CXC_COL_CP.RIF]               = p.rifEntidad || '';
      filaRet[CXC_COL_CP.NRO_FACTURA]       = f.nroFactura;
      filaRet[CXC_COL_CP.FECHA_FACTURA]     = new Date(p.fecha + 'T00:00:00');
      filaRet[CXC_COL_CP.TOTAL_FACTURA]     = 0;
      filaRet[CXC_COL_CP.SALDO_FACTURA]     = 0;
      filaRet[CXC_COL_CP.SALDO_FACTURA_USD] = 0;
      filaRet[CXC_COL_CP.COBRADA_PAGADA]    = esVenta ? 'Retencion_CXC' : 'Retencion_CXP';
      filaRet[CXC_COL_CP.TRANSACCION]       = 'Retencion ' + tipoRet;
      filaRet[CXC_COL_CP.ID]                = nroComp;
      filaRet[CXC_COL_CP.ABONO]             = retUSD;
      filaRet[CXC_COL_CP.ABONO_BS]          = retBs;
      filaRet[CXC_COL_CP.MONEDA]            = 'Dolares';
      filaRet[CXC_COL_CP.TASA]              = tasa;
      filaRet[CXC_COL_CP.DESCRIPCION]       = 'Retención ' + tipoRet + ' Comprobante N° ' + nroComp + ' - Factura ' + f.nroFactura;
      filaRet[CXC_COL_CP.FECHA_ABONO]       = new Date(p.fecha + 'T00:00:00');
      filaRet[CXC_COL_CP.TIPO_MOVIMIENTO]   = 'RETENCION';
      filaRet[campoCompCol] = nroComp;

      filasNuevas.push(filaRet);
    });

    if (filasNuevas.length > 0) {
      primeraFilaEscrita = shCP.getLastRow() + 1;
      shCP.getRange(primeraFilaEscrita, 1, filasNuevas.length, 23).setValues(filasNuevas);
      filasEscritas = filasNuevas.length;
    }

    // Asiento
    var nroAsientoContable = null;
    if (totalRetBs > 0) {
      try {
        nroAsientoContable = generarCorrelativoAsiento(true);
        var contra = resolverContrapartidaCxCp(esVenta);
        var glosa = 'Retención ' + tipoRet + ' Comprobante ' + nroComp;
        var lineasAsiento = construirAsientoRetencion(
          esVenta, nroAsientoContable, p.fecha, p.nombreEntidad, glosa, nroComp,
          Math.round(totalRetBs * 100) / 100, tipoRet, contra
        );
        guardarAsientoMasivo(lineasAsiento);
      } catch (eAsiento) {
        throw new Error('Se registraron filas pero el asiento falló: ' + eAsiento.message + '. Se revertirán cambios.');
      }
    }

    // Guardar comprobante en ComprobantesRetencion para que sea reimprimible
    // (Solo si NO existe ya)
    try {
      if (validarComprobanteUnico(nroComp, tipoRet)) {
        var hojaCR = _getHojaRetenciones_();
        var totalUSD = Math.round((totalRetBs / tasa) * 100) / 100;
        hojaCR.appendRow([
          nroComp, tipoRet, p.fecha, p.rifEntidad || '', p.nombreEntidad,
          p.facturas.length === 1 ? p.facturas[0].nroFactura : '(varias)',
          totalUSD, totalRetBs,
          tipoRet === 'IVA' ? 75 : (Number(p.porcentajeISLR) || 0),
          'Retención ' + tipoRet + ' (manual)',
          totalUSD, tasa, totalRetBs,
          Session.getActiveUser().getEmail(), new Date(), 'Vigente'
        ]);
      }
    } catch (eCR) {
      Logger.log('[registrarMultiplesRetenciones] No se pudo guardar en ComprobantesRetencion: ' + eCR.message);
    }

    return {
      ok: true,
      comprobante: nroComp,
      total: p.facturas.length,
      nroAsientoContable: nroAsientoContable
    };
  } catch (e) {
    if (filasEscritas > 0 && shCPref && primeraFilaEscrita > 0) {
      try { shCPref.deleteRows(primeraFilaEscrita, filasEscritas); } catch (rb) {}
    }
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ================================================================
   BLOQUE 8 — ASIENTOS Y ASENTAMIENTO EN COBROS/PAGOS
   ================================================================ */

function construirAsientoRetencion(esCXC, nroAsiento, fechaOp, tercero, glosa, refDoc, montoBs, tipoRetencion, ctaContrapartida) {
  var constants = getConstantsContables();
  var fechaAsiento = new Date(fechaOp + 'T00:00:00');
  var lineas = [];

  var ctaRet;
  if (esCXC) {
    ctaRet = (tipoRetencion === 'IVA') ? constants.RET_IVA_COBRAR : constants.RET_ISLR_COBRAR;
  } else {
    ctaRet = (tipoRetencion === 'IVA') ? constants.RET_IVA_PAGAR : constants.RET_ISLR_PAGAR;
  }
  if (!ctaRet || !ctaRet.COD) {
    throw new Error('CONFIG: Falta cuenta de retención ' + tipoRetencion + '.');
  }

  var metaRet = _obtenerMetadatosCuentaCatalogo(ctaRet.COD, ctaRet.NOM, ctaRet.CTA, ctaRet.SUBCTA || ctaRet.SUB);
  var metaContra = _obtenerMetadatosCuentaCatalogo(ctaContrapartida.cod, ctaContrapartida.nom, ctaContrapartida.mayor, ctaContrapartida.subcuenta);

  if (esCXC) {
    lineas.push({
      nroAsiento: nroAsiento, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
      clienteProveedor: tercero, descripcionGeneral: glosa,
      codigoCuenta: metaRet.cod, cuentaContable: metaRet.nom,
      montoTransaccion: montoBs,
      descripcionTransaccion: 'Retención ' + tipoRetencion + ' por cobrar',
      transaccionRef: refDoc, contabilidad: "1",
      mayor: metaRet.mayor, subcuenta: metaRet.subcuenta,
      eeff: metaRet.eeff, tipoCuenta: metaRet.tipoCuenta, grupoBalance: metaRet.grupoBalance
    });
    lineas.push({
      nroAsiento: nroAsiento, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
      clienteProveedor: tercero, descripcionGeneral: glosa,
      codigoCuenta: metaContra.cod, cuentaContable: metaContra.nom,
      montoTransaccion: montoBs * -1,
      descripcionTransaccion: 'Retención ' + tipoRetencion + ' aplicada a CxC',
      transaccionRef: refDoc, contabilidad: "1",
      mayor: metaContra.mayor, subcuenta: metaContra.subcuenta,
      eeff: metaContra.eeff, tipoCuenta: metaContra.tipoCuenta, grupoBalance: metaContra.grupoBalance
    });
  } else {
    lineas.push({
      nroAsiento: nroAsiento, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
      clienteProveedor: tercero, descripcionGeneral: glosa,
      codigoCuenta: metaContra.cod, cuentaContable: metaContra.nom,
      montoTransaccion: montoBs,
      descripcionTransaccion: 'Retención ' + tipoRetencion + ' aplicada a CxP',
      transaccionRef: refDoc, contabilidad: "1",
      mayor: metaContra.mayor, subcuenta: metaContra.subcuenta,
      eeff: metaContra.eeff, tipoCuenta: metaContra.tipoCuenta, grupoBalance: metaContra.grupoBalance
    });
    lineas.push({
      nroAsiento: nroAsiento, fechaAsiento: fechaAsiento, tipoAsiento: 'Automático',
      clienteProveedor: tercero, descripcionGeneral: glosa,
      codigoCuenta: metaRet.cod, cuentaContable: metaRet.nom,
      montoTransaccion: montoBs * -1,
      descripcionTransaccion: 'Retención ' + tipoRetencion + ' por pagar',
      transaccionRef: refDoc, contabilidad: "1",
      mayor: metaRet.mayor, subcuenta: metaRet.subcuenta,
      eeff: metaRet.eeff, tipoCuenta: metaRet.tipoCuenta, grupoBalance: metaRet.grupoBalance
    });
  }
  return lineas;
}

function asentarRetencionEnCobrosPagos(ss, payload) {
  if (!ss || !payload) return;
  var hojaCP = ss.getSheetByName('CobrosPagos');
  if (!hojaCP) return;

  var dataCP = hojaCP.getDataRange().getValues();
  if (dataCP.length < 2) return;

  var hCP = dataCP[0].map(_normHeaderRet_);
  var cols = {
    proveedor: hCP.indexOf('proveedorcliente'),
    rif: hCP.indexOf('rif'),
    factura: hCP.indexOf('nrofactura'),
    fechaFac: hCP.indexOf('fechafactura'),
    totalFac: hCP.indexOf('totalfactura'),
    saldoFac: hCP.indexOf('saldofactura'),
    saldoFacUsd: hCP.indexOf('saldofacturausd'),
    cobrada: hCP.indexOf('cobradapagada'),
    transaccion: hCP.indexOf('transaccion'),
    id: hCP.indexOf('id'),
    abono: hCP.indexOf('abono'),
    abonoBs: hCP.indexOf('abonobs'),
    moneda: hCP.indexOf('moneda'),
    tasa: hCP.indexOf('tasa'),
    descripcion: hCP.indexOf('descripcion'),
    fechaAbono: hCP.indexOf('fechaabono'),
    tipoMov: hCP.indexOf('tipomovimiento'),
    retIVA: hCP.indexOf('nrocomprobantesretiva'),
    retISLR: hCP.indexOf('nrocomprobantesretislr')
  };

  var tipoRet = String(payload.tipoRetencion || 'IVA').toUpperCase();
  var nroComp = String(payload.nroComprobante || '').trim();
  var fecha = payload.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var tasa = Number(payload.tasa) || 1;
  if (!nroComp || !payload.facturas || !payload.facturas.length) return;

  var esVenta = payload.esVenta !== undefined ? payload.esVenta : false;
  var campoComp = (tipoRet === 'IVA') ? cols.retIVA : cols.retISLR;

  var filas = [];
  payload.facturas.forEach(function (f) {
    var montoBs = Number(f.montoRetenidoBs || 0);
    var montoUSD = Number(f.montoRetenidoUSD || 0);
    if (montoBs <= 0 && montoUSD <= 0) return;

    var fila = new Array(hCP.length).fill('');
    if (cols.proveedor !== -1) fila[cols.proveedor] = payload.nombreEntidad || '';
    if (cols.rif !== -1) fila[cols.rif] = payload.rifEntidad || '';
    if (cols.factura !== -1) fila[cols.factura] = f.nroFactura;
    if (cols.fechaFac !== -1) fila[cols.fechaFac] = new Date(fecha + 'T00:00:00');
    if (cols.totalFac !== -1) fila[cols.totalFac] = 0;
    if (cols.saldoFac !== -1) fila[cols.saldoFac] = 0;
    if (cols.saldoFacUsd !== -1) fila[cols.saldoFacUsd] = 0;
    if (cols.cobrada !== -1) fila[cols.cobrada] = esVenta ? 'Retencion_CXC' : 'Retencion_CXP';
    if (cols.transaccion !== -1) fila[cols.transaccion] = 'Retencion ' + tipoRet;
    if (cols.id !== -1) fila[cols.id] = nroComp;
    if (cols.abono !== -1) fila[cols.abono] = montoUSD;
    if (cols.abonoBs !== -1) fila[cols.abonoBs] = montoBs;
    if (cols.moneda !== -1) fila[cols.moneda] = 'Dolares';
    if (cols.tasa !== -1) fila[cols.tasa] = tasa;
    if (cols.descripcion !== -1) fila[cols.descripcion] = 'Retención ' + tipoRet + ' N° ' + nroComp + ' - Factura ' + f.nroFactura;
    if (cols.fechaAbono !== -1) fila[cols.fechaAbono] = new Date(fecha + 'T00:00:00');
    if (cols.tipoMov !== -1) fila[cols.tipoMov] = 'RETENCION';
    if (campoComp !== -1) fila[campoComp] = nroComp;

    filas.push(fila);
  });

  if (filas.length) {
    hojaCP.getRange(hojaCP.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
  }
}

/* ================================================================
   BLOQUE 9 — REIMPRESIÓN UNIFICADA
   ================================================================ */

/**
 * Busca en ComprobantesRetencion primero; si no, cae a CobrosPagos.
 * Retorna { ok, html }.
 */
function reimprimirComprobanteHistorial(nroComprobante, tipoRet, nroFactura) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tipoFinal = String(tipoRet || 'IVA').toUpperCase();

    // 1) Buscar en ComprobantesRetencion
    var shCR = ss.getSheetByName(HOJA_RETENCIONES);
    if (shCR && shCR.getLastRow() > 1) {
      var dataCR = shCR.getDataRange().getValues();
      var hCR = dataCR[0].map(_normHeaderRet_);
      var cNro = hCR.indexOf('nrocomprobante');
      var cTipo = hCR.indexOf('tipo');
      var cEstado = hCR.indexOf('estado');

      for (var i = 1; i < dataCR.length; i++) {
        if (String(dataCR[i][cNro]).trim() !== String(nroComprobante).trim()) continue;
        var t = String(dataCR[i][cTipo] || '').toUpperCase();
        if (t && t !== tipoFinal) continue;
        var est = String(dataCR[i][cEstado] || '').toUpperCase();
        if (est === 'ANULADO') {
          return { ok: false, html: '', mensaje: 'El comprobante está ANULADO.' };
        }
        return _reimprimirDesdeComprobanteRetencion_(dataCR[i], hCR);
      }
    }

    // 2) Fallback: CobrosPagos
    var res = generarReciboRetencionPorComprobante(nroComprobante, tipoFinal);
    if (res && res.html) return { ok: true, html: res.html };
    throw new Error('No se generó el HTML del comprobante.');
  } catch (e) {
    Logger.log('Error al reimprimir: ' + e.message);
    throw new Error(e.message);
  }
}

function _reimprimirDesdeComprobanteRetencion_(fila, headers) {
  var obj = {};
  headers.forEach(function (h, idx) { obj[h] = fila[idx]; });

  var datosParaTemplate = {
    tipo: obj.tipo,
    tipoRetencion: obj.tipo,
    nroComprobante: obj.nrocomprobante,
    fecha: obj.fecha,
    tasa: obj.tasa || 1,
    nombreEntidad: obj.razonsocialtercero,
    rifEntidad: obj.riftercero,
    empresa: obtenerDatosEmpresa(),
    facturas: [{
      nroFactura: obj.nrofactura,
      nroControl: '00-00000000',
      fechaFactura: obj.fecha,
      totalFactura: 0,
      baseImponible: obj.baseusd || 0,
      montoIva: 0,
      porcentaje: obj.porcentaje || 0,
      montoRetenidoBs: obj.montoretbs || 0,
      montoRetenidoUSD: obj.montoretusd || 0
    }],
    montoEnLetras: numeroALetras(obj.montoretbs || 0)
  };

  var template = HtmlService.createTemplateFromFile('ReciboRetencion');
  template.datos = datosParaTemplate;
  return { ok: true, html: template.evaluate().getContent() };
}

/* ================================================================
   BLOQUE 10 — REIMPRESIÓN DESDE CobrosPagos (fallback)
   ================================================================ */

function generarReciboRetencionPorComprobante(identificador, tipoRet, idSpreadsheetEmpresa) {
  try {
    if (!identificador) throw new Error('Debe indicar un N° de comprobante o factura.');

    var ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) throw new Error('No se pudo acceder a la empresa activa.');

    var hojaCP = ss.getSheetByName('CobrosPagos');
    var tipoFinal = String(tipoRet || 'IVA').toUpperCase();
    var idBusqueda = String(identificador).trim().toLowerCase();
    var idNorm = idBusqueda.replace(/^0+/, '');
    var filasEncontradas = [];

    // Pre-cargar datos fiscales de Compras y Ventas
    var infoFacturasMadre = {};
    ['Compras', 'Ventas'].forEach(function(nomHoja) {
      var sh = ss.getSheetByName(nomHoja);
      if (!sh || sh.getLastRow() < 2) return;
      var d = sh.getDataRange().getValues();
      var h = d[0].map(_normHeaderRet_);
      var cF = h.indexOf('nrofacturaproveedor') !== -1 ? h.indexOf('nrofacturaproveedor') : h.indexOf('nrofacturacliente');
      var cCtrl = h.indexOf('nrocontrolproveedor') !== -1 ? h.indexOf('nrocontrolproveedor') : h.indexOf('nrocontrol');
      var cBase = h.indexOf('subtotalproveedor') !== -1 ? h.indexOf('subtotalproveedor') : h.indexOf('subtotalcliente');
      var cIva = h.indexOf('ivaproveedor') !== -1 ? h.indexOf('ivaproveedor') : h.indexOf('ivacliente');
      var cTot = h.indexOf('totalproveedor') !== -1 ? h.indexOf('totalproveedor') : h.indexOf('totalcliente');
      var cExe = h.indexOf('totalexentocompra') !== -1 ? h.indexOf('totalexentocompra') : h.indexOf('totalexentoventa');

      for (var f = 1; f < d.length; f++) {
        var numF = String(d[f][cF] || '').trim().toLowerCase();
        if (numF) {
          infoFacturasMadre[numF] = {
            nroControl: cCtrl !== -1 ? String(d[f][cCtrl] || '00-00000000') : '00-00000000',
            baseImponible: cBase !== -1 ? (parseFloat(d[f][cBase]) || 0) : 0,
            montoIva: cIva !== -1 ? (parseFloat(d[f][cIva]) || 0) : 0,
            totalFactura: cTot !== -1 ? (parseFloat(d[f][cTot]) || 0) : 0,
            montoExento: cExe !== -1 ? (parseFloat(d[f][cExe]) || 0) : 0
          };
        }
      }
    });

    if (hojaCP && hojaCP.getLastRow() > 1) {
      var dataCP = hojaCP.getDataRange().getValues();
      var hCP = dataCP[0].map(_normHeaderRet_);

      var cFact = hCP.indexOf('nrofactura');
      var cCompIva = hCP.indexOf('nrocomprobantesretiva');
      var cCompIslr = hCP.indexOf('nrocomprobantesretislr');
      var cTrans = hCP.indexOf('transaccion');
      var cAbono = hCP.indexOf('abono');
      var cAbonoBs = hCP.indexOf('abonobs');
      var cTasa = hCP.indexOf('tasa');
      var cFechaFac = hCP.indexOf('fechafactura');
      var cTercero = hCP.indexOf('proveedorcliente');
      var cRif = hCP.indexOf('rif');

      var colComp = (tipoFinal === 'ISLR') ? cCompIslr : cCompIva;

      for (var i = 1; i < dataCP.length; i++) {
        var r = dataCP[i];
        var valFact = cFact !== -1 ? String(r[cFact]).trim().toLowerCase() : '';
        var valComp = colComp !== -1 ? String(r[colComp]).trim().toLowerCase() : '';
        var valTrans = cTrans !== -1 ? String(r[cTrans]).trim().toLowerCase() : '';

        var coincide = (valComp && valComp === idBusqueda) ||
                       (valFact && (valFact === idBusqueda || valFact.replace(/^0+/, '') === idNorm)) ||
                       (valTrans && valTrans === idBusqueda);

        if (!coincide) continue;

        var montoUSD = cAbono !== -1 ? (parseFloat(r[cAbono]) || 0) : 0;
        var tasaFila = cTasa !== -1 ? (parseFloat(r[cTasa]) || 1) : 1;
        var montoBs = cAbonoBs !== -1 ? (parseFloat(r[cAbonoBs]) || 0) : (montoUSD * tasaFila);

        if (montoBs <= 0 && montoUSD <= 0) continue;

        var compReal = (colComp !== -1 && r[colComp]) ? r[colComp] : (cTrans !== -1 ? r[cTrans] : identificador);
        var fechaFac = cFechaFac !== -1 ? r[cFechaFac] : new Date();
        var madre = infoFacturasMadre[valFact] || { nroControl: '00-00000000', baseImponible: 0, montoIva: 0, totalFactura: 0 };

        filasEncontradas.push({
          nroFactura: cFact !== -1 ? r[cFact] : '',
          nroControl: madre.nroControl,
          fechaFactura: fechaFac instanceof Date
            ? Utilities.formatDate(fechaFac, Session.getScriptTimeZone(), 'dd/MM/yyyy')
            : fechaFac,
          totalFactura: madre.totalFactura,
          baseImponible: madre.baseImponible,
          montoIva: madre.montoIva,
          porcentaje: (tipoFinal === 'IVA' ? 75 : 2),
          montoRetenidoBs: montoBs,
          montoRetenidoUSD: montoUSD,
          tercero: cTercero !== -1 ? r[cTercero] : '',
          rif: cRif !== -1 ? r[cRif] : '',
          comprobante: compReal,
          fecha: fechaFac,
          tasa: tasaFila
        });
      }
    }

    if (filasEncontradas.length === 0) {
      throw new Error('No se encontraron registros activos de retención ' + tipoFinal + ' para: ' + identificador);
    }

    var primerReg = filasEncontradas[0];
    var totalAcumuladoBs = 0;
    filasEncontradas.forEach(function (f) { totalAcumuladoBs += f.montoRetenidoBs; });

    var datosParaTemplate = {
      tipo: tipoFinal,
      tipoRetencion: tipoFinal,
      nroComprobante: primerReg.comprobante || identificador,
      fecha: primerReg.fecha || new Date(),
      tasa: primerReg.tasa || 1,
      nombreEntidad: primerReg.tercero,
      rifEntidad: primerReg.rif,
      empresa: obtenerDatosEmpresa(),
      facturas: filasEncontradas,
      montoEnLetras: numeroALetras(totalAcumuladoBs)
    };

    var template = HtmlService.createTemplateFromFile('ReciboRetencion');
    template.datos = datosParaTemplate;
    return { ok: true, html: template.evaluate().getContent() };

  } catch (err) {
    Logger.log('Error en generarReciboRetencionPorComprobante: ' + err.message);
    throw new Error(err.message);
  }
}

function generarHtmlComprobanteRetencion(payload) {
  try {
    var tipoRet = String(payload.tipoRetencion || 'IVA').toUpperCase();
    var totalRetBs = 0;

    var facturas = (payload.facturas || []).map(function (f) {
      var retBs = parseFloat(f.montoRetenidoBs || 0);
      totalRetBs += retBs;
      return {
        nroFactura: f.nroFactura || f.factura || 'S/F',
        nroControl: f.nroControl || '00-00000000',
        fechaFactura: f.fechaFactura || payload.fecha || new Date(),
        totalFactura: parseFloat(f.totalFactura || f.total || 0),
        baseImponible: parseFloat(f.baseImponible || f.base || 0),
        montoIva: parseFloat(f.montoIva || f.iva || 0),
        porcentaje: parseFloat(f.porcentaje || (tipoRet === 'IVA' ? 75 : 2)),
        montoRetenidoBs: retBs,
        montoRetenidoUSD: parseFloat(f.montoRetenidoUSD || 0)
      };
    });

    var datosParaTemplate = {
      tipo: tipoRet,
      tipoRetencion: tipoRet,
      nroComprobante: payload.nroComprobante || 'SIN-NUMERO',
      fecha: payload.fecha || new Date(),
      tasa: parseFloat(payload.tasa || 1),
      nombreEntidad: payload.nombreEntidad || payload.tercero || '',
      rifEntidad: payload.rifEntidad || payload.rif || '',
      empresa: obtenerDatosEmpresa(),
      facturas: facturas,
      montoEnLetras: numeroALetras(totalRetBs)
    };

    var template = HtmlService.createTemplateFromFile('ReciboRetencion');
    template.datos = datosParaTemplate;
    return { ok: true, html: template.evaluate().getContent() };
  } catch (e) {
    Logger.log('Error en generarHtmlComprobanteRetencion: ' + e.message);
    throw new Error(e.message);
  }
}

function enviarRetencionPorCorreo(datos) {
  try {
    var resHtml = generarHtmlComprobanteRetencion(datos);
    if (!resHtml || !resHtml.html) throw new Error('No se pudo compilar el comprobante.');

    var blob = Utilities.newBlob(resHtml.html, MimeType.HTML).getAs(MimeType.PDF);
    blob.setName('Comprobante_Retencion_' + datos.tipo + '_' + datos.nroComprobante + '.pdf');

    MailApp.sendEmail({
      to: datos.emailDestino,
      subject: datos.asunto || ('Comprobante de Retención ' + datos.tipo + ' N° ' + datos.nroComprobante),
      htmlBody: datos.mensajeBody || '<p>Estimado(s), adjuntamos su comprobante de retención oficial.</p>',
      attachments: [blob]
    });

    return { success: true, message: 'Comprobante enviado exitosamente por correo.' };
  } catch (e) {
    Logger.log('Error en enviarRetencionPorCorreo: ' + e.message);
    throw new Error(e.message);
  }
}

/* ================================================================
   BLOQUE 11 — UTILIDADES AUXILIARES
   ================================================================ */

function obtenerRetencionesDisponibles(nombreTercero, rifTercero) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shCP = ss.getSheetByName('CobrosPagos');
    if (!shCP || shCP.getLastRow() < 2) return [];

    var data = shCP.getDataRange().getValues();
    var hd = data[0].map(_normHeaderRet_);

    var cNombre   = hd.indexOf('proveedorcliente');
    var cRif      = hd.indexOf('rif');
    var cFactura  = hd.indexOf('nrofactura');
    var cCompIVA  = hd.indexOf('nrocomprobantesretiva');
    var cCompISLR = hd.indexOf('nrocomprobantesretislr');
    var cMonto    = hd.indexOf('abono');
    var cMontoBs  = hd.indexOf('abonobs');
    var cTasa     = hd.indexOf('tasa');
    var cSaldoUSD = hd.indexOf('saldofacturausd');
    var cTipoMov  = hd.indexOf('tipomovimiento');

    var norm = function (s) { return String(s || '').trim().toLowerCase(); };
    var nNombre = norm(nombreTercero);
    var nRif = norm(rifTercero);
    var retenciones = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var tipoMov = norm(row[cTipoMov]);
      if (tipoMov !== 'retencion') continue;

      var nombreFila = norm(row[cNombre]);
      var rifFila = norm(row[cRif]);
      if (nombreFila !== nNombre && rifFila !== nRif) continue;

      var compIVA = String(row[cCompIVA] || '').trim();
      var compISLR = String(row[cCompISLR] || '').trim();
      var comprobante = compIVA || compISLR;
      if (!comprobante) continue;

      // Determinar tipo por columna que la contiene
      var tipoRet = compIVA ? 'IVA' : 'ISLR';
      var montoUSD = Number(row[cMonto]) || 0;
      var montoBs  = cMontoBs !== -1 ? (Number(row[cMontoBs]) || 0) : 0;
      var tasa = Number(row[cTasa]) || 1;
      if (montoBs <= 0) montoBs = montoUSD * tasa;

      retenciones.push({
        id: comprobante,
        comprobante: comprobante,
        tipoRetencion: tipoRet,
        factura: String(row[cFactura] || '').trim(),
        montoUSD: montoUSD,
        montoBs: montoBs,
        tasa: tasa,
        yaAplicada: true
      });
    }

    return retenciones;
  } catch (e) {
    Logger.log('[RET] Error obteniendo retenciones disponibles: ' + e.message);
    return [];
  }
}

function fmtNum(n) {
  return (Number(n) || 0).toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatearFechaRango(fechaISO) {
  if (!fechaISO) return '';
  var p = String(fechaISO).split('-');
  if (p.length !== 3) return fechaISO;
  return p[2] + '/' + p[1] + '/' + p[0];
}

/* ================================================================
   FUNCIONES EXTERNAS REQUERIDAS (deben existir en otros .gs)
   ----------------------------------------------------------------
   - obtenerDatosEmpresa()
   - parseNumeroLatino(v)
   - numeroALetras(n)
   - obtenerNroControlReal(ss, nroFactura, esVenta)
   - generarCorrelativoAsiento(esAutomatico)
   - resolverContrapartidaCxCp(esVenta)
   - guardarAsientoMasivo(lineas)
   - _obtenerMetadatosCuentaCatalogo(cod, nom, cta, sub)
   - getConstantsContables()  → { RET_IVA_COBRAR, RET_ISLR_COBRAR,
                                   RET_IVA_PAGAR, RET_ISLR_PAGAR }
   - CXC_COL_CP (objeto de índices de columnas CobrosPagos)
   ----------------------------------------------------------------
   HOJAS REQUERIDAS EN EL SPREADSHEET:
   - Facturas          (NroFactura, TotalFactura|TotaFactura, SaldoFactura,
                        SaldoFacturaUSD, Tasa, nroComprobantesRetIVA, nroComprobantesRetISLR)
   - AbonosFacturas    (o "Abonos"): 12 columnas en este orden →
                        [Fecha, NroFactura, Tercero, RIF, TipoMovimiento,
                         MontoUSD, MontoBS, Descripcion, ComprobanteIVA,
                         ComprobanteISLR, Timestamp, Usuario]
   - ComprobantesRetencion (auto-creada con 16 columnas)
   - CobrosPagos       (ProveedorCliente, Rif, NroFactura, FechaFactura,
                        TotalFactura, SaldoFactura, SaldoFacturaUSD,
                        CobradaPagada, Transaccion, ID, Abono, AbonoBS,
                        Moneda, Tasa, Descripcion, FechaAbono,
                        TipoMovimiento, nroComprobantesRetIVA,
                        nroComprobantesRetISLR, ...)
   - Asientos          (NroAsiento, FechaAsiento, ..., Descripcion, ...,
                        TransaccionRef, ...)
   - Anticipos         (idAnticipo, Fecha, Tercero, RIF, MontoUSD,
                        MontoBS, Tasa, SaldoUSD, Estado)
   - MovimientosBancarios (Fecha, Tipo, Banco, Cuenta, CodContable,
                           NomContable, Ingreso|Egreso, Moneda, Monto,
                           Referencia, Tercero, RIF, Timestamp)
   - Ventas / Compras  (columnas de saldo y status según esVenta)
   ================================================================ */
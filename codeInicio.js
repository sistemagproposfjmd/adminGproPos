/************************************************************
  PANEL FINANCIERO · CodeInicio.gs
  100% DINÁMICO · INTEGRACIÓN DE TASAS (RESUMEN / TASAS)
  Lee: Compras, Ventas, DetalleFacturas, CobrosPagos, Resumen, Tasas
************************************************************/

const CONFIG = {

  TOKEN_SECRET_KEY: 'GPRO_SECRET_F82#kL9@zPq!', // Reemplazar con clave secreta robusta
  SESSION_HOURS: 8,
  MAX_OTP_PER_HOUR: 4,
  LOCK_TIMEOUT_MS: 20000, // 20 segundos de espera máxima de bloqueo exclusivo

  HOJA_COMPRAS:  'Compras',
  HOJA_VENTAS:   'Ventas',
  HOJA_DETALLE:  'DetalleFacturas',
  HOJA_COBROS:   'CobrosPagos',
  HOJA_MOVIM:    'MovBancoCaja',
  HOJA_CINTA:    'Cinta',
  HOJA_RESUMEN:  'Resumen',
  HOJA_TASAS:    'Tasas',
  ESTADO_PEND_COBRAR: 'porcobrar',
  ESTADO_PEND_PAGAR:  'porpagar',
  TXT_ANTICIPO:  'ANTICIPO',
  MESES_HISTORICO: 6,
  ID_HOJA: '' // Vacío si el script vive dentro de la hoja

  
};

const MES_ABBR = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const fmtUsd_ = n => '$' + fmt_(n);
const fmtBs_  = n => 'Bs. ' + fmt_(n);
const fmtInt_ = n => fmt_(n, 0);
const esUSD_ = m => /usd|u\.s\.d|dollar|\$/.test(norm_(m));
const serieATexto_ = serie => serie.map(m => m.label + ':' + Math.round(m.v)).join(',');

function obtenerHistorialTasasGrafico_(tasasObj) {
  try {
    var ssM = Contexto.getMasterSS();
    var shTasas = ssM ? ssM.getSheetByName('Tasas') : null;
    var histUSD = [], histEUR = [];

    if (shTasas) {
      var lastRow = shTasas.getLastRow();
      if (lastRow >= 3) {
        var numFilas = Math.min(lastRow - 2, 60);
        var vals = shTasas.getRange(3, 1, numFilas, 14).getValues();
        
        var head = shTasas.getRange(1, 1, 1, 14).getValues()[0].map(function(h){
          return String(h||'').trim().toLowerCase();
        });
        var iFecha = head.indexOf('fecha');
        var iUSD   = head.findIndex(function(h){ return /usd|dolar|dólar|bcv/.test(h) && !/eur|euro/.test(h); });
        var iEUR   = head.findIndex(function(h){ return /eur|euro/.test(h); });
        
        if (iFecha === -1) iFecha = 0;
        if (iUSD   === -1) iUSD   = 4;
        if (iEUR   === -1) iEUR   = 7;

        for (var i = 0; i < vals.length; i++) {
          var f = vals[i][iFecha];
          var u = parseFloat(String(vals[i][iUSD]||'').replace(',','.')) || 0;
          var e = parseFloat(String(vals[i][iEUR]||'').replace(',','.')) || 0;
          var esFechaValida = (f instanceof Date && !isNaN(f.getTime()));
          
          if (esFechaValida && u > 0) histUSD.push(u);
          if (esFechaValida && e > 0) histEUR.push(e);
        }
      }
    }

    var tasaActualUSD = (tasasObj && tasasObj.tasaUSD) || 0;
    var tasaActualEUR = (tasasObj && tasasObj.tasaEUR) || 0;

    if (histUSD.length < 2 && tasaActualUSD > 0) histUSD = [tasaActualUSD, tasaActualUSD];
    if (histEUR.length < 2 && tasaActualEUR > 0) histEUR = [tasaActualEUR, tasaActualEUR];

    histUSD = histUSD.slice(-6);
    histEUR = histEUR.slice(-6);

    var fmt = function(arr){ return arr.map(function(n){ return Math.round(n*100)/100; }).join(','); };

    return {
      barrasUSD: histUSD.length >= 2 ? fmt(histUSD) : '',
      barrasEUR: histEUR.length >= 2 ? fmt(histEUR) : ''
    };
  } catch (e) {
    Logger.log('Error obtenerHistorialTasasGrafico_: ' + e);
    return { barrasUSD:'', barrasEUR:'' };
  }
}
function listarHojasSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('❌ No hay Spreadsheet activo. Ejecuta esto desde el proyecto bound de emp1.');
    return;
  }

  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📊 INVENTARIO DE HOJAS');
  Logger.log('📄 Archivo: ' + ss.getName());
  Logger.log('🆔 Spreadsheet ID: ' + ss.getId());
  Logger.log('🔗 URL: ' + ss.getUrl());
  Logger.log('📅 Fecha: ' + new Date().toLocaleString('es-VE'));
  Logger.log('═══════════════════════════════════════════════════');

  const hojas = ss.getSheets();
  Logger.log('Total de hojas: ' + hojas.length);
  Logger.log('');

  const resumen = [];

  hojas.forEach(function(hoja, idx) {
    const nombre = hoja.getName();
    const lastRow = hoja.getLastRow();
    const lastCol = hoja.getLastColumn();
    const maxFilas = hoja.getMaxRows();
    const maxCols = hoja.getMaxColumns();

    let encabezados = [];
    let tieneDatos = false;
    let tieneFormulas = false;
    let tieneValidaciones = false;
    let protegida = false;
    let filasConDatos = 0;

    try {
      if (lastCol > 0) {
        encabezados = hoja.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      }
      filasConDatos = Math.max(0, lastRow - 1);
      tieneDatos = filasConDatos > 0;

      // Buscar fórmulas y validaciones en un rango razonable
      if (lastRow > 0 && lastCol > 0) {
        const rango = hoja.getRange(1, 1, Math.min(lastRow, 50), Math.min(lastCol, 20));
        const formulas = rango.getFormulas();
        tieneFormulas = formulas.some(function(fila) {
          return fila.some(function(celda) { return celda !== ''; });
        });
        const validaciones = rango.getDataValidations();
        tieneValidaciones = validaciones.some(function(fila) {
          return fila.some(function(celda) { return celda !== null; });
        });
      }

      // Verificar protección
      try { protegida = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET).length > 0; } catch(e) {}

    } catch (e) {
      Logger.log('⚠️ Error leyendo hoja "' + nombre + '": ' + e.message);
    }

    Logger.log('───────────────────────────────────────────────');
    Logger.log('[' + (idx + 1) + '] HOJA: "' + nombre + '"');
    Logger.log('    Dimensiones activas: ' + lastRow + ' filas × ' + lastCol + ' cols');
    Logger.log('    Dimensiones máximas: ' + maxFilas + ' filas × ' + maxCols + ' cols');
    Logger.log('    Filas con datos (sin header): ' + filasConDatos);
    Logger.log('    ¿Tiene datos?: ' + (tieneDatos ? '✅ SÍ' : '❌ NO (solo headers)')); 
    Logger.log('    ¿Tiene fórmulas?: ' + (tieneFormulas ? '✅ SÍ' : '❌ NO'));
    Logger.log('    ¿Tiene validaciones?: ' + (tieneValidaciones ? '✅ SÍ' : '❌ NO'));
    Logger.log('    ¿Protegida?: ' + (protegida ? '🔒 SÍ' : '🔓 NO'));
    Logger.log('    Encabezados (' + encabezados.length + '): ');
    Logger.log('      [' + encabezados.map(function(h, i) { return (i+1) + ':' + h; }).join(' | ') + ']');

    resumen.push({
      nombre: nombre,
      filasConDatos: filasConDatos,
      columnas: lastCol,
      tieneDatos: tieneDatos,
      tieneFormulas: tieneFormulas,
      tieneValidaciones: tieneValidaciones,
      protegida: protegida,
      encabezados: encabezados.join(' | ')
    });
  });

  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📋 RESUMEN EN FORMATO TABLA (para copiar a Excel)');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('Nombre\tFilas datos\tColumnas\tDatos\tFórmulas\tValidaciones\tProtegida\tEncabezados');
  resumen.forEach(function(r) {
    Logger.log([
      r.nombre,
      r.filasConDatos,
      r.columnas,
      r.tieneDatos ? 'SÍ' : 'NO',
      r.tieneFormulas ? 'SÍ' : 'NO',
      r.tieneValidaciones ? 'SÍ' : 'NO',
      r.protegida ? 'SÍ' : 'NO',
      r.encabezados
    ].join('\t'));
  });

  Logger.log('');
  Logger.log('✅ Inventario completo.');
  Logger.log('👉 Copia el resumen tabulado y pégalo en el chat para clasificar las hojas.');
}
function buscarReferenciasEnCodigo() {
  const TERMINOS = [
    'empresas', 'Empresas',
    'Usuarios', 'usuarios',
    'Log_Accesos', 'log_accesos', 'LogAccesos',
    'manual_ES', 'manual_EN', 'Manual_ES', 'Manual_EN',
    'Resumen', 'resumen',
    'Tasas', 'tasas',
    'Hoja 18', 'Hoja18'
  ];

  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🔍 BÚSQUEDA DE REFERENCIAS A HOJAS EN EL CÓDIGO');
  Logger.log('📅 Fecha: ' + new Date().toLocaleString('es-VE'));
  Logger.log('═══════════════════════════════════════════════════');

  // Listar todos los archivos del proyecto
  let archivos = [];
  try {
    // En proyectos bound, esto funciona para .gs
    const files = DriveApp.getFileById(ScriptApp.getScriptId());
    // Fallback: no podemos listar archivos del proyecto por API directamente
    // Usamos una lista manual: pedimos al usuario que la complete si es necesario
  } catch (e) {
    Logger.log('⚠️ No se pudo obtener el ID del proyecto: ' + e.message);
  }

  // Como Apps Script no expone una API para listar archivos del proyecto,
  // usamos esta estrategia: buscar en TODOS los archivos conocidos
  // Si algún archivo no está en la lista, se puede añadir manualmente.

  // Lista conocida (basada en lo que me has pasado). Añade aquí si falta alguno.
  const ARCHIVOS_CONOCIDOS = [
    'code.gs',
    'BuscarReferencias.gs'
    // ⚠️ AÑADE AQUÍ EL RESTO DE TUS ARCHIVOS .gs, .html, .js, .css
    // Ejemplo:
    // 'DashBoard.gs', 'BancosCajas.gs', 'Compras.gs', 'Ventas.gs',
    // 'Productos.gs', 'Contabilidad.gs', 'Pedidos.gs', 'UsuariosLista.gs',
    // 'index.html', 'DashBoard.html', 'BancosCajas.html', ...
  ];

  Logger.log('');
  Logger.log('📁 Archivos a escanear: ' + ARCHIVOS_CONOCIDOS.length);
  Logger.log('');

  const resultados = {};
  const archivosEscaneados = [];
  const archivosConError = [];

  ARCHIVOS_CONOCIDOS.forEach(function(nombreArchivo) {
    try {
      let contenido = '';
      if (nombreArchivo.endsWith('.html') || nombreArchivo.endsWith('.gs') || nombreArchivo.endsWith('.js') || nombreArchivo.endsWith('.css')) {
        // Intentar leer como HTML (funciona para .gs también en Apps Script)
        contenido = HtmlService.createHtmlOutputFromFile(nombreArchivo.replace(/\.(gs|js|css)$/, '')).getContent();
        // Si es .gs, el nombre real es sin extensión en HtmlService
      }

      // Estrategia alternativa: leer archivos .gs con getContent no funciona bien
      // Para .gs usamos otra aproximación
      try {
        contenido = HtmlService.createHtmlOutputFromFile(nombreArchivo.replace(/\.gs$/, '')).getContent();
      } catch(e) {
        // Si falla, intentar con el nombre tal cual (por si es .html)
        contenido = HtmlService.createHtmlOutputFromFile(nombreArchivo.replace(/\.html$/, '')).getContent();
      }

      archivosEscaneados.push(nombreArchivo);
    } catch (e) {
      archivosConError.push(nombreArchivo + ' (' + e.message + ')');
    }
  });

  Logger.log('✅ Archivos escaneados: ' + archivosEscaneados.length);
  Logger.log('⚠️ Archivos con error: ' + archivosConError.length);
  if (archivosConError.length > 0) {
    Logger.log('   Errores:');
    archivosConError.forEach(function(e) { Logger.log('     - ' + e); });
  }

  Logger.log('');
  Logger.log('───────────────────────────────────────────────');
  Logger.log('📊 RESULTADO');
  Logger.log('───────────────────────────────────────────────');
  Logger.log('⚠️ LIMITACIÓN CONOCIDA: Apps Script no permite listar archivos del proyecto');
  Logger.log('   automáticamente. Este script solo escanea los archivos que le indiques');
  Logger.log('   en ARCHIVOS_CONOCIDOS.');
  Logger.log('');
  Logger.log('🔧 RECOMENDACIÓN: En lugar de este script, haz una búsqueda manual:');
  Logger.log('   1. Abre el editor de Apps Script');
  Logger.log('   2. En el panel izquierdo, haz click en el ícono de lupa (🔍)');
  Logger.log('   3. Busca: "empresas"');
  Logger.log('   4. Anota cada archivo y línea donde aparezca');
  Logger.log('   5. Repite para: Usuarios, Log_Accesos, manual_ES, manual_EN, Resumen, Tasas');
  Logger.log('');
  Logger.log('   La búsqueda de Apps Script soporta regex y muestra:');
  Logger.log('   - Nombre del archivo');
  Logger.log('   - Número de línea');
  Logger.log('   - Contexto');
  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('✅ Fin.');
  Logger.log('═══════════════════════════════════════════════════');
}
function obtenerTarjetas() {
  const hoy = new Date();
  const anio = hoy.getFullYear(), mes = hoy.getMonth();
  const enMes = f => f instanceof Date && !isNaN(f.getTime()) && f.getFullYear() === anio && f.getMonth() === mes;

  const compras = leerCompras_() || [];
  const ventas  = leerVentas_() || [];
  const cobros  = leerCobros_() || [];

  // Obtener tasas usando tu propia función
  const tasasObj = obtenerTasaCambioActual();
  const tasaUSD = tasasObj.tasaUSD;
  const tasaEUR = tasasObj.tasaEUR;
  const curvasTasas = obtenerHistorialTasasGrafico_(cobros, tasaUSD, tasaEUR);

  /* ───── VENTAS Y COMPARATIVA MENSUAL REAL ───── */
  const venMesF = ventas.filter(r => enMes(getDate_(r, 'FechaFacturaCliente', 'FechaFactura', 'Fecha')));
  const venMesUsd = venMesF.reduce((s, r) => s + getNum_(r, 'TotalClienteUSD', 'TotalUSD', 'MontoUSD'), 0);

  const dAnt = new Date(anio, mes - 1, 1);
  const venMesAnt = ventas.filter(r => {
    const f = getDate_(r, 'FechaFacturaCliente', 'FechaFactura', 'Fecha');
    return f && f.getFullYear() === dAnt.getFullYear() && f.getMonth() === dAnt.getMonth();
  }).reduce((s, r) => s + getNum_(r, 'TotalClienteUSD', 'TotalUSD', 'MontoUSD'), 0);

  const deltaVentasVal = venMesAnt > 0 ? ((venMesUsd - venMesAnt) / venMesAnt * 100) : (venMesUsd > 0 ? 100 : 0);
  const deltaVentasTxt = (deltaVentasVal >= 0 ? '+' : '') + deltaVentasVal.toFixed(1) + '%';

  /* ───── COMPRAS Y COMPARATIVA MENSUAL REAL ───── */
  const comMesF = compras.filter(r => enMes(getDate_(r, 'FechaFacturaProveedor', 'FechaFactura', 'Fecha')));
  const comMesUsd = comMesF.reduce((s, r) => s + getNum_(r, 'TotalProveedorDivisa', 'TotalUSD', 'MontoUSD'), 0);

  const comMesAnt = compras.filter(r => {
    const f = getDate_(r, 'FechaFacturaProveedor', 'FechaFactura', 'Fecha');
    return f && f.getFullYear() === dAnt.getFullYear() && f.getMonth() === dAnt.getMonth();
  }).reduce((s, r) => s + getNum_(r, 'TotalProveedorDivisa', 'TotalUSD', 'MontoUSD'), 0);

  const deltaComprasVal = comMesAnt > 0 ? ((comMesUsd - comMesAnt) / comMesAnt * 100) : (comMesUsd > 0 ? 100 : 0);
  const deltaComprasTxt = (deltaComprasVal >= 0 ? '+' : '') + deltaComprasVal.toFixed(1) + '%';

  /* ───── CARTERAS REALES ───── */
  const porCobrar = ventas.filter(r => r.t('cobradaPagada') === norm_(CONFIG.ESTADO_PEND_COBRAR) || getNum_(r, 'SaldoClienteUSD', 'SaldoUSD') > 0);
  const porPagar  = compras.filter(r => r.t('cobradaPagada') === norm_(CONFIG.ESTADO_PEND_PAGAR) || getNum_(r, 'SaldoProveedorDivisa', 'SaldoUSD') > 0);

  const cxcUsd = porCobrar.reduce((s, r) => s + getNum_(r, 'SaldoClienteUSD', 'SaldoUSD'), 0);
  const cxpUsd = porPagar.reduce((s, r) => s + getNum_(r, 'SaldoProveedorDivisa', 'SaldoUSD'), 0);

  const vencCob = vencidas_(porCobrar, 'FechaVencimientoCliente', 'SaldoClienteUSD');
  const vencPag = vencidas_(porPagar, 'FechaVencimientoProveedor', 'SaldoProveedorDivisa');

  /* ───── SERIES MENSUALES ───── */
  const serieVen = serieMensual_(ventas, r => getDate_(r, 'FechaFacturaCliente', 'FechaFactura', 'Fecha'), r => getNum_(r, 'TotalClienteUSD', 'TotalUSD'));
  const serieCom = serieMensual_(compras, r => getDate_(r, 'FechaFacturaProveedor', 'FechaFactura', 'Fecha'), r => getNum_(r, 'TotalProveedorDivisa', 'TotalUSD'));

  /* ───── TOPS ───── */
  const topCli  = agruparSaldo_(porCobrar, 'NombreCliente', 'SaldoClienteUSD');
  const topProv = agruparSaldo_(porPagar, 'NombreProveedor', 'SaldoProveedorDivisa');

  /* ───── FLUJO REAL Y PROGRESO DE COBRANZA ───── */
  const abonosMes = cobros.filter(r => enMes(getDate_(r, 'fechaAbono', 'Fecha')) && getNum_(r, 'Abono', 'Monto') > 0);
  const cobradoRealUSD = abonosMes.filter(r => r.t('Transaccion') === 'cobro').reduce((s, r) => {
    const ab = getNum_(r, 'Abono', 'Monto');
    const tasa = getNum_(r, 'Tasa', 'TasaCambio') || 1;
    return s + (esUSD_(getVal_(r, 'Moneda')) ? ab : (tasa > 0 ? ab / tasa : 0));
  }, 0);

  const totalFacturasMes = venMesF.length;
  const facturasCobradasMes = venMesF.filter(r => r.t('cobradaPagada') !== norm_(CONFIG.ESTADO_PEND_COBRAR) && getNum_(r, 'SaldoClienteUSD', 'SaldoUSD') === 0).length;
  const pctCobroReal = venMesUsd > 0 ? Math.min(100, Math.round((cobradoRealUSD / venMesUsd) * 100)) : (totalFacturasMes > 0 ? Math.round((facturasCobradasMes / totalFacturasMes) * 100) : 0);

  const ultVenta = ventas.length > 0 ? ventas[ventas.length - 1] : null;
  const ultCobro = cobros.length > 0 ? cobros[cobros.length - 1] : null;

  const t = [];

  /* ═══════════════ COLUMNA 1 (VENTAS & GRÁFICO TASA USD) ═══════════════ */
  t.push({
    columna: 1, tipo: 'stat', etiqueta: 'Ventas del mes',
    valor: fmtUsd_(venMesUsd), delta: deltaVentasTxt,
    barras: serieVen.map(m => Math.round(m.v)).join(','),
    meta: totalFacturasMes + ' facturas emitidas este mes', tono: 'naranja'
  });

  // Tasa USD en formato STAT CON BARRAS DE HISTORIAL
  t.push({
    columna: 1, tipo: 'stat', etiqueta: 'Tasa Dólar Oficial (USD)',
    valor: tasaUSD > 0 ? fmtBs_(tasaUSD) : 'Sin tasa',
    delta: 'USD/BS',
    barras: curvasTasas.barrasUSD,
    meta: 'Hoja Resumen/Tasas · Clic para ampliar',
    pill: 'BCV', tono: 'naranja'
  });

  t.push({
    columna: 1, tipo: 'progreso', etiqueta: 'Efectividad de cobranza',
    valor: pctCobroReal + '%',
    barras: (totalFacturasMes > 0 ? facturasCobradasMes : 0) + '/' + (totalFacturasMes > 0 ? totalFacturasMes : 8),
    meta: fmtUsd_(cobradoRealUSD) + ' cobrados de ' + fmtUsd_(venMesUsd),
    pill: pctCobroReal >= 70 ? 'Al día' : (pctCobroReal > 0 ? 'En progreso' : 'Pendiente'), tono: 'naranja'
  });

  if (ultVenta) {
    const nroFac = getVal_(ultVenta, 'NroFactura', 'Factura', 'NumeroFactura', 'Control') || 'Última Factura';
    const nomCli = getVal_(ultVenta, 'NombreCliente', 'Cliente', 'RazonSocial') || 'Cliente';
    const totFac = getNum_(ultVenta, 'TotalClienteUSD', 'TotalUSD', 'MontoUSD');
    const fFac   = getDate_(ultVenta, 'FechaFacturaCliente', 'Fecha');
    const fStr   = fFac ? Utilities.formatDate(fFac, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'Reciente';

    t.push({
      columna: 1, tipo: 'factura', etiqueta: 'Última Venta Registrada',
      valor: String(nroFac), pill: 'Emitida',
      linea1: nomCli + '|' + fmtUsd_(totFac),
      linea2: 'Fecha emisión|' + fStr,
      delta: fmtUsd_(totFac), nota: 'Factura reciente en sistema'
    });
  }

  t.push({
    columna: 1, tipo: 'cliente',
    nota: topCli.length > 0 ? topCli[0].n : (ventas.length > 0 ? getVal_(ventas[0], 'NombreCliente', 'Cliente') : 'Sin deudores'),
    meta: 'Cuenta por cobrar',
    linea1: 'Saldo Pendiente|' + fmtUsd_(topCli.length > 0 ? topCli[0].m : 0),
    linea2: 'Cartera Total|' + fmtUsd_(cxcUsd), tono: 'naranja'
  });

  /* ═══════════════ COLUMNA 2 (COMPRAS & GRÁFICO TASA EUR) ═══════════════ */
  t.push({
    columna: 2, tipo: 'stat', etiqueta: 'Compras del mes',
    valor: fmtUsd_(comMesUsd), delta: deltaComprasTxt,
    barras: serieCom.map(m => Math.round(m.v)).join(','),
    meta: comMesF.length + ' facturas de compras registradas', tono: 'azul'
  });

  // Tasa EUR en formato STAT CON BARRAS DE HISTORIAL
  t.push({
    columna: 2, tipo: 'stat', etiqueta: 'Tasa Euro Oficial (EUR)',
    valor: tasaEUR > 0 ? fmtBs_(tasaEUR) : 'Sin tasa',
    delta: 'EUR/BS',
    barras: curvasTasas.barrasEUR,
    meta: 'Hoja Resumen/Tasas · Clic para ampliar',
    pill: 'BCV', tono: 'azul'
  });

  t.push({
    columna: 2, tipo: 'gasto', etiqueta: 'Deuda a proveedores',
    nota: topProv.length > 0 ? topProv[0].n : 'Proveedores',
    valor: fmtUsd_(cxpUsd),
    meta: porPagar.length + ' facturas pendientes por pagar', tono: 'azul'
  });

  t.push({
    columna: 2, tipo: 'cliente',
    nota: topProv.length > 0 ? topProv[0].n : (compras.length > 0 ? getVal_(compras[0], 'NombreProveedor', 'Proveedor') : 'Sin acreedores'),
    meta: 'Cuenta por pagar',
    linea1: 'Saldo a Pagar|' + fmtUsd_(topProv.length > 0 ? topProv[0].m : 0),
    linea2: 'Compromiso|' + (vencPag.n > 0 ? vencPag.n + ' vencidas' : 'Al día'), tono: 'azul'
  });

  const margenReal = venMesUsd > 0 ? Math.round(((venMesUsd - comMesUsd) / venMesUsd) * 100) : 0;
  t.push({
    columna: 2, tipo: 'progreso', etiqueta: 'Margen de Rendimiento',
    valor: margenReal + '%',
    barras: Math.min(8, Math.max(0, Math.round((margenReal / 100) * 8))) + '/8',
    meta: fmtUsd_(venMesUsd - comMesUsd) + ' margen neto de ventas',
    pill: margenReal > 25 ? 'Rentable' : 'Operativo', tono: 'azul'
  });

  /* ═══════════════ COLUMNA 3 (ALERTAS REALES & MOVIMIENTO) ═══════════════ */
  let notaAlerta = '';
  if (vencCob.n > 0) {
    notaAlerta = 'Tienes ' + vencCob.n + ' factura(s) por cobrar vencida(s) que suman ' + fmtUsd_(vencCob.usd) + '. Recomendado gestionar cobro hoy.';
  } else if (vencPag.n > 0) {
    notaAlerta = 'Tienes ' + vencPag.n + ' factura(s) de proveedores vencida(s) por ' + fmtUsd_(vencPag.usd) + '. Revisa tu flujo de pagos.';
  } else {
    notaAlerta = 'Todas tus facturas están al día. Has recaudado ' + fmtUsd_(cobradoRealUSD) + ' en abonos durante este mes.';
  }

  t.push({
    columna: 3, tipo: 'recordatorio', etiqueta: 'Notificación Financiera',
    pill: (vencCob.n > 0 || vencPag.n > 0) ? 'Atención' : 'Excelente',
    nota: '"' + notaAlerta + '"',
    meta: 'Análisis en tiempo real'
  });

  t.push({
    columna: 3, tipo: 'stat', etiqueta: 'Cuentas por cobrar',
    valor: fmtUsd_(cxcUsd),
    pill: vencCob.n > 0 ? vencCob.n + ' vencidas' : 'Al día',
    meta: porCobrar.length + ' facturas pendientes de cobro', tono: 'naranja'
  });

  if (ultCobro) {
    const nomAbono = getVal_(ultCobro, 'NombreCliente', 'Cliente', 'Beneficiario', 'Descripcion') || 'Cobro registrado';
    const montAbono = getNum_(ultCobro, 'Abono', 'Monto');
    const fAbono = getDate_(ultCobro, 'fechaAbono', 'Fecha');
    const fAbonoStr = fAbono ? Utilities.formatDate(fAbono, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'Reciente';

    t.push({
      columna: 3, tipo: 'recibo', etiqueta: 'Último Abono Recibido',
      pill: 'Confirmado', nota: nomAbono + ' · ' + fAbonoStr,
      delta: fmtUsd_(montAbono), meta: 'Acreditado en cuenta'
    });
  }

  t.push({
    columna: 3, tipo: 'stat', etiqueta: 'Recaudación mensual',
    valor: fmtUsd_(cobradoRealUSD), delta: deltaVentasTxt,
    barras: serieVen.map(m => Math.round(m.v)).join(','),
    meta: abonosMes.length + ' ingresos registrados este mes', pill: 'Caja Real', tono: 'naranja'
  });

  return t;
}
function obtenerDatos() {return { tarjetas: obtenerTarjetas(), frases: leerFrases_() };}
function probarCalculos() { Logger.log(JSON.stringify(obtenerDatos(), null, 2)); }

// ============================================================
// RESOLUTOR MULTI-EMPRESA PARA EL DASHBOARD
// ============================================================
function libro_(idOpcional) {
  var idTarget = idOpcional ||
    (typeof Contexto !== 'undefined' && Contexto.getEmpresaActivaId ? Contexto.getEmpresaActivaId() : null) ||
    PropertiesService.getUserProperties().getProperty('EMPRESA_ACTIVA_ID');

  if (!idTarget) {
    throw new Error('libro_: empresaId no resuelto. Evita fuga de datos entre empresas.');
  }
  if (typeof Contexto !== 'undefined' && Contexto.getEmpresaSS) {
    return Contexto.getEmpresaSS(idTarget);
  }
  return SpreadsheetApp.openById(idTarget);
}

var __hojaCache = {};
function leerHoja_(nombre, ssOpcional) {
  var ss = (ssOpcional && typeof ssOpcional === 'object' && ssOpcional.getSheetByName)
    ? ssOpcional
    : libro_(ssOpcional);

  if (!ss) return [];
  
  var cacheKey = ss.getId() + '::' + nombre;
  var ahora = Date.now();
  
  // Caché de 30 segundos para respuestas rápidas
  if (__hojaCache[cacheKey] && (ahora - __hojaCache[cacheKey].ts) < 30000) {
    return __hojaCache[cacheKey].data;
  }

  var sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) {
    __hojaCache[cacheKey] = { ts: ahora, data: [] };
    return [];
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var head = vals[0].map(function(h){ return norm_(h); });

  var result = vals.slice(1)
    .filter(function(row){
      // Ignorar filas completamente vacías
      return row.some(function(c){ return c !== '' && c !== null && c !== undefined; });
    })
    .map(function(row){
      var o = { raw: row };
      var normRow = row.map(function(c){ return norm_(c); });
      
      o.v = function(n){
        var k = head.indexOf(norm_(n));
        return k < 0 ? '' : row[k];
      };
      o.n = function(n){ return num_(o.v(n)); };
      o.f = function(n){ return fecha_(o.v(n)); };
      o.t = function(n){
        var k = head.indexOf(norm_(n));
        return k < 0 ? '' : normRow[k];
      };
      return o;
    });

  __hojaCache[cacheKey] = { ts: ahora, data: result };
  return result;
}

function leerCompras_(ss) { 
  var res = leerHoja_(CONFIG.HOJA_COMPRAS, ss); 
  return Array.isArray(res) ? res : []; 
}

function leerVentas_(ss)  { 
  var res = leerHoja_(CONFIG.HOJA_VENTAS, ss); 
  return Array.isArray(res) ? res : []; 
}

function leerDetalle_(ss) { 
  var res = leerHoja_(CONFIG.HOJA_DETALLE, ss); 
  return Array.isArray(res) ? res : []; 
}

function leerCobros_(ss)  { 
  var res = leerHoja_(CONFIG.HOJA_COBROS, ss); 
  return Array.isArray(res) ? res : []; 
}

function obtenerTarjetasDashboard(token, idEmpresa, modo) {
  try {
    modo = modo || 'dashboard';

    // 1) Resolver Spreadsheet (con o sin token)
    var ss = null;
    
    if (token) {
      var sesion = validarSessionToken(token);
      if (sesion) {
        var respEmp = Contexto.obtenerEmpresasDeUsuario(sesion.sub);
        var empresas = (respEmp && respEmp.success) ? (respEmp.empresas || []) : [];
        var permitida = empresas.filter(function(e) {
          return String(e.idSpreadsheet) === String(idEmpresa) || String(e.idEmpresa) === String(idEmpresa);
        })[0];
        
        var esAdmin = (sesion.role || '').toLowerCase() === 'admin' || (sesion.role || '').toLowerCase() === 'owner';
        if (permitida || esAdmin) {
          var idSS = permitida ? permitida.idSpreadsheet : idEmpresa;
          if (idSS) ss = SpreadsheetApp.openById(idSS);
        }
      }
    }
    
    // Fallback: Si no hay token o falló, usar idEmpresa directo
    if (!ss && idEmpresa) {
      try {
        ss = SpreadsheetApp.openById(idEmpresa);
      } catch(e) {
        return { success: false, error: 'No se pudo abrir el Spreadsheet: ' + e.message };
      }
    }

    if (!ss) return { success: false, error: 'No se pudo acceder a la hoja de la empresa' };

    // 2) Leer datos
    var hoy = new Date();
    var anio = hoy.getFullYear(), mes = hoy.getMonth();
    var enMes = function(f){ return f instanceof Date && !isNaN(f.getTime()) && f.getFullYear()===anio && f.getMonth()===mes; };
    var enAnio = function(f){ return f instanceof Date && !isNaN(f.getTime()) && f.getFullYear()===anio; };

    var compras = leerCompras_(ss) || [];
    var ventas  = leerVentas_(ss) || [];
    var detalle = leerDetalle_(ss) || [];
    var cobros  = leerCobros_(ss) || [];

    // 3) Tasas
    var tasasObj = obtenerTasaCambioActual();
    var tasaUSD = tasasObj.tasaUSD || 0;
    var tasaEUR = tasasObj.tasaEUR || 0;
    var curvas = obtenerHistorialTasasGrafico_(tasasObj);

    // 4) Cálculos del mes
    var venMesF = ventas.filter(function(r){ return enMes(getDate_(r,'FechaFacturaCliente','FechaFactura','Fecha')); });
    var venMesUsd = venMesF.reduce(function(s,r){ return s + getNum_(r,'TotalClienteUSD','TotalUSD','MontoUSD'); }, 0);

    var dAnt = new Date(anio, mes-1, 1);
    var venMesAnt = ventas.filter(function(r){
      var f = getDate_(r,'FechaFacturaCliente','FechaFactura','Fecha');
      return f && f.getFullYear()===dAnt.getFullYear() && f.getMonth()===dAnt.getMonth();
    }).reduce(function(s,r){ return s + getNum_(r,'TotalClienteUSD','TotalUSD','MontoUSD'); }, 0);
    var deltaVen = venMesAnt>0 ? ((venMesUsd-venMesAnt)/venMesAnt*100) : 0;

    var comMesF = compras.filter(function(r){ return enMes(getDate_(r,'FechaFacturaProveedor','FechaFactura','Fecha')); });
    var comMesUsd = comMesF.reduce(function(s,r){ return s + getNum_(r,'TotalProveedorDivisa','TotalUSD','MontoUSD'); }, 0);
    var comMesAnt = compras.filter(function(r){
      var f = getDate_(r,'FechaFacturaProveedor','FechaFactura','Fecha');
      return f && f.getFullYear()===dAnt.getFullYear() && f.getMonth()===dAnt.getMonth();
    }).reduce(function(s,r){ return s + getNum_(r,'TotalProveedorDivisa','TotalUSD','MontoUSD'); }, 0);
    var deltaCom = comMesAnt>0 ? ((comMesUsd-comMesAnt)/comMesAnt*100) : 0;

    var utilidadMes = venMesUsd - comMesUsd;
    var margenPct   = venMesUsd>0 ? (utilidadMes/venMesUsd*100) : 0;

    var porCobrar = ventas.filter(function(r){ return r.t('cobradaPagada')===norm_(CONFIG.ESTADO_PEND_COBRAR) || getNum_(r,'SaldoClienteUSD','SaldoUSD')>0; });
    var cxcUsd = porCobrar.reduce(function(s,r){ return s + getNum_(r,'SaldoClienteUSD','SaldoUSD'); },0);
    var vencCob = vencidas_(porCobrar,'FechaVencimientoCliente','SaldoClienteUSD');

    var porPagar = compras.filter(function(r){ return r.t('cobradaPagada')===norm_(CONFIG.ESTADO_PEND_PAGAR) || getNum_(r,'SaldoProveedorDivisa','SaldoUSD')>0; });
    var cxpUsd = porPagar.reduce(function(s,r){ return s + getNum_(r,'SaldoProveedorDivisa','SaldoUSD'); },0);
    var vencPag = vencidas_(porPagar,'FechaVencimientoProveedor','SaldoProveedorDivisa');

    // Series mensuales
    var serieVen = serieMensual_(ventas,  function(r){ return getDate_(r,'FechaFacturaCliente','FechaFactura','Fecha'); }, function(r){ return getNum_(r,'TotalClienteUSD','TotalUSD'); });
    var serieCom = serieMensual_(compras, function(r){ return getDate_(r,'FechaFacturaProveedor','FechaFactura','Fecha'); }, function(r){ return getNum_(r,'TotalProveedorDivisa','TotalUSD'); });
    var labels = serieVen.map(function(m){ return m.label; });
    var venVals = serieVen.map(function(m){ return Math.round(m.v||0); });
    var comVals = serieCom.map(function(m){ return Math.round(m.v||0); });
    var utilVals = venVals.map(function(v,i){ return v - (comVals[i]||0); });

    // Tops
    var topCli  = agruparSaldo_(porCobrar,'NombreCliente','SaldoClienteUSD').slice(0,5);
    var topProv = agruparSaldo_(porPagar,'NombreProveedor','SaldoProveedorDivisa').slice(0,5);

    // Top productos
    var detVenta = detalle.filter(function(r){ return r.t('Transaccion')==='venta'; });
    var cUni = {}, cMonto = {};
    detVenta.forEach(function(r){
      var nom = getVal_(r,'Descripcion','Producto','Articulo') || 'Ítem';
      var cant = getNum_(r,'Cantidad') || 1;
      var monto = getNum_(r,'PrecioTotalUSD','TotalUSD','PrecioTotal');
      cUni[nom] = (cUni[nom]||0) + cant;
      cMonto[nom] = (cMonto[nom]||0) + monto;
    });
    var topProdUni   = Object.keys(cUni).map(function(k){ return {label:k,val:cUni[k]}; }).sort(function(a,b){return b.val-a.val;}).slice(0,5);
    var topProdMonto = Object.keys(cMonto).map(function(k){ return {label:k,val:Math.round(cMonto[k])}; }).sort(function(a,b){return b.val-a.val;}).slice(0,5);

    // Flujo
    var abonosMes = cobros.filter(function(r){ return enMes(getDate_(r,'fechaAbono','Fecha')) && getNum_(r,'Abono','Monto')>0; });
    var sumaMon = function(arr, esUsd) {
      return arr.reduce(function(s,r){
        var ab = getNum_(r,'Abono','Monto'), t = getNum_(r,'Tasa','TasaCambio')||0;
        var usd = esUSD_(getVal_(r,'Moneda')) ? ab : (t>0 ? ab/t : 0);
        var bs  = esUSD_(getVal_(r,'Moneda')) ? ab*t : ab;
        return s + (esUsd ? usd : bs);
      }, 0);
    };
    var cobradosUSD  = sumaMon(abonosMes.filter(function(r){ return r.t('Transaccion')==='cobro'; }), true);
    var pagadosUSD   = sumaMon(abonosMes.filter(function(r){ return r.t('Transaccion')==='pago'; }), true);
    var anticiposUSD = sumaMon(cobros.filter(function(r){ return r.t('TipoMovimiento').indexOf(norm_(CONFIG.TXT_ANTICIPO))>=0; }), true);

    // ============================================================
    // MODO LANDING → shape de tarjetas verticales (inicio.html)
    // ============================================================
    if (modo === 'landing') {
      var tarjetasLanding = [];
      tarjetasLanding.push({
        columna:1, tipo:'stat', etiqueta:'Ventas del mes',
        valor: fmtUsd_(venMesUsd), delta: (deltaVen>=0?'+':'') + deltaVen.toFixed(1) + '%',
        barras: serieVen.map(function(m){ return Math.round(m.v); }).join(','),
        meta: venMesF.length + ' facturas emitidas este mes', tono:'naranja'
      });
      tarjetasLanding.push({
        columna:1, tipo:'stat', etiqueta:'Tasa Dólar Oficial (USD)',
        valor: tasaUSD>0 ? fmtBs_(tasaUSD) : 'Sin tasa',
        delta:'USD/BS',
        barras: curvas.barrasUSD || '',
        meta:'Hoja Resumen/Tasas · Clic para ampliar',
        pill:'BCV', tono:'naranja'
      });
      tarjetasLanding.push({
        columna:2, tipo:'stat', etiqueta:'Compras del mes',
        valor: fmtUsd_(comMesUsd), delta: (deltaCom>=0?'+':'') + deltaCom.toFixed(1) + '%',
        barras: serieCom.map(function(m){ return Math.round(m.v); }).join(','),
        meta: comMesF.length + ' facturas de compras registradas', tono:'azul'
      });
      tarjetasLanding.push({
        columna:2, tipo:'stat', etiqueta:'Tasa Euro Oficial (EUR)',
        valor: tasaEUR>0 ? fmtBs_(tasaEUR) : 'Sin tasa',
        delta:'EUR/BS',
        barras: curvas.barrasEUR || '',
        meta:'Hoja Resumen/Tasas · Clic para ampliar',
        pill:'BCV', tono:'azul'
      });
      tarjetasLanding.push({
        columna:3, tipo:'recordatorio', etiqueta:'Notificación Financiera',
        pill: (vencCob.n>0 || vencPag.n>0) ? 'Atención' : 'Excelente',
        nota: vencCob.n>0
          ? 'Tienes ' + vencCob.n + ' factura(s) vencida(s) por ' + fmtUsd_(vencCob.usd) + '.'
          : (vencPag.n>0
              ? 'Tienes ' + vencPag.n + ' factura(s) de proveedores vencida(s) por ' + fmtUsd_(vencPag.usd) + '.'
              : 'Todas tus facturas están al día.'),
        meta: 'Análisis en tiempo real'
      });
      return tarjetasLanding;
    }

    // ============================================================
    // MODO DASHBOARD → shape actual (compatibilidad)
    // ============================================================
    var setClientes = new Set(); ventas.forEach(function(r){ var n=getVal_(r,'NombreCliente','Cliente','RazonSocial').trim(); if(n)setClientes.add(n); });
    var setProveedores = new Set(); compras.forEach(function(r){ var n=getVal_(r,'NombreProveedor','Proveedor','RazonSocial').trim(); if(n)setProveedores.add(n); });
    var setCliAntes = new Set(); ventas.forEach(function(r){ var f=getDate_(r,'FechaFacturaCliente','Fecha'); if(f && f < new Date(anio,mes,1)){ var n=getVal_(r,'NombreCliente','Cliente').trim(); if(n)setCliAntes.add(n); } });
    var setProvAntes = new Set(); compras.forEach(function(r){ var f=getDate_(r,'FechaFacturaProveedor','Fecha'); if(f && f < new Date(anio,mes,1)){ var n=getVal_(r,'NombreProveedor','Proveedor').trim(); if(n)setProvAntes.add(n); } });
    var cliNuevosMes  = [].concat.apply([], [Array.from(setClientes)]).filter(function(c){ return !setCliAntes.has(c); }).length;
    var provNuevosMes = [].concat.apply([], [Array.from(setProveedores)]).filter(function(p){ return !setProvAntes.has(p); }).length;

    var tarjetas = [
      { id:'ventas', etiqueta:'Ventas del Mes', icono:'trending', tono:'naranja', valor:fmtUsd_(venMesUsd), sub: '', delta:deltaVen, meta:venMesF.length+' facturas emitidas', barras:serieATexto_(serieVen) },
      { id:'compras', etiqueta:'Compras del Mes', icono:'truck', tono:'azul', valor:fmtUsd_(comMesUsd), sub:'', delta:deltaCom, meta:comMesF.length+' facturas de proveedores', barras:serieATexto_(serieCom) },
      { id:'utilidad', etiqueta:'Utilidad del Mes', icono:'wallet', tono:'emerald', valor:fmtUsd_(utilidadMes), sub:'Margen '+margenPct.toFixed(1)+'%', delta:null, meta:'Ventas − Compras', barras:utilVals.map(function(v,i){ return labels[i]+':'+Math.abs(v); }).join(',') },
      { id:'cxc', etiqueta:'Cuentas por Cobrar', icono:'clock', tono:'naranja', valor:fmtUsd_(cxcUsd), sub:'', delta:null, meta:porCobrar.length+' facturas pendientes', pill: vencCob.n>0 ? vencCob.n+' vencidas' : '' },
      { id:'cxp', etiqueta:'Cuentas por Pagar', icono:'clock', tono:'azul', valor:fmtUsd_(cxpUsd), sub:'', delta:null, meta:porPagar.length+' facturas pendientes', pill: vencPag.n>0 ? vencPag.n+' vencidas' : '' },
      { id:'nroVentas', etiqueta:'Nro. de Ventas', icono:'file', tono:'naranja', valor:fmtInt_(venMesF.length), sub:'Este mes', delta:null, meta:venMesF.length+' en el año' },
      { id:'nroCompras', etiqueta:'Nro. de Compras', icono:'box', tono:'azul', valor:fmtInt_(comMesF.length), sub:'Este mes', delta:null, meta:comMesF.length+' en el año' },
      { id:'nroClientes', etiqueta:'Nro. de Clientes', icono:'users', tono:'naranja', valor:fmtInt_(setClientes.size), sub:'Registrados', delta:null, meta: (cliNuevosMes>0?'+'+cliNuevosMes+' nuevos este mes':'Sin nuevos este mes') },
      { id:'nroProveedores', etiqueta:'Nro. de Proveedores', icono:'truck', tono:'azul', valor:fmtInt_(setProveedores.size), sub:'Registrados', delta:null, meta: (provNuevosMes>0?'+'+provNuevosMes+' nuevos este mes':'Sin nuevos este mes') }
    ];

    var graficos = {
      ventasVsCompras: { labels:labels, ventas:venVals, compras:comVals },
      utilidadMensual: { labels:labels, data:utilVals },
      topProductosVendidos: { labels:topProdUni.map(function(x){return x.label;}), data:topProdUni.map(function(x){return x.val;}) },
      topProductosVendidosMonto: { labels:topProdMonto.map(function(x){return x.label;}), data:topProdMonto.map(function(x){return x.val;}) },
      topClientes: { labels:topCli.map(function(x){return x.n;}), data:topCli.map(function(x){return Math.round(x.m);}) },
      topProveedores: { labels:topProv.map(function(x){return x.n;}), data:topProv.map(function(x){return Math.round(x.m);}) },
      carteraCxC: { alDia: Math.max(0, cxcUsd - (vencCob.usd||0)), vencida: (vencCob.usd||0) },
      carteraCxP: { alDia: Math.max(0, cxpUsd - (vencPag.usd||0)), vencida: (vencPag.usd||0) },
      flujoCaja: { labels:['Cobrado','Pagado','Anticipo'], data:[Math.round(cobradosUSD), Math.round(pagadosUSD), Math.round(anticiposUSD)] },
      balanceAnual: {
        ventas: Math.round(ventas.filter(function(r){ return enAnio(getDate_(r,'FechaFacturaCliente','FechaFactura','Fecha')); }).reduce(function(s,r){return s+getNum_(r,'TotalClienteUSD','TotalUSD');},0)),
        compras: Math.round(compras.filter(function(r){ return enAnio(getDate_(r,'FechaFacturaProveedor','FechaFactura','Fecha')); }).reduce(function(s,r){return s+getNum_(r,'TotalProveedorDivisa','TotalUSD');},0))
      }
    };

    return { success:true, generado:new Date().toISOString(), tarjetas:tarjetas, graficos:graficos };

  } catch (err) {
    Logger.log('Error en obtenerTarjetasDashboard: ' + err.stack);
    return { success:false, error: err.toString() };
  }
}

function norm_(s) {return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v == null ? '' : v).trim().replace(/[^\d.,\-]/g, '');
  if (!s) return 0;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}
function fecha_(v) {
  if (v instanceof Date) return v;
  const s = String(v == null ? '' : v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[2] - 1, +m[1]); }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function fmt_(n, dec) {
  if (dec === undefined) dec = 2;
  const neg = n < 0; n = Math.abs(n);
  const p = n.toFixed(dec).split('.');
  p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-' : '') + p.join(',');
}
function getVal_(r, ...nombres) {
  for (let n of nombres) {
    const val = r.v(n);
    if (val !== '' && val !== null && val !== undefined) return String(val).trim();
  }
  return '';
}
function getNum_(r, ...nombres) {
  for (let n of nombres) {
    const val = r.n(n);
    if (val !== 0 && !isNaN(val)) return val;
  }
  return 0;
}
function getDate_(r, ...nombres) {
  for (let n of nombres) {
    const val = r.f(n);
    if (val instanceof Date && !isNaN(val.getTime())) return val;
  }
  return null;
}
function serieMensual_(filas, fnFecha, fnMonto) {
  const hoy = new Date();
  const meses = [];
  for (let i = CONFIG.MESES_HISTORICO - 1; i >= 0; i--) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
    meses.push({ y: d.getFullYear(), m: d.getMonth(), label: MES_ABBR[d.getMonth()], v: 0 });
  }
  filas.forEach(r => {
    const f = fnFecha(r);
    if (!f) return;
    const m = meses.find(x => x.y === f.getFullYear() && x.m === f.getMonth());
    if (m) m.v += fnMonto(r);
  });
  return meses;
}
function agruparSaldo_(filas, colNombre, colSaldo) {
  const m = {};
  filas.forEach(r => {
    const k = getVal_(r, colNombre, 'Cliente', 'Proveedor', 'RazonSocial') || '—';
    m[k] = (m[k] || 0) + getNum_(r, colSaldo, 'SaldoUSD', 'Saldo');
  });
  return Object.keys(m).map(k => ({ n: k, m: m[k] }))
    .filter(x => x.m > 0).sort((a, b) => b.m - a.m).slice(0, 5);
}
function vencidas_(filasPendientes, colFechaVenc, colSaldoUsd) {
  const hoy = new Date(); 
  hoy.setHours(0, 0, 0, 0);
  
  const venc = (filasPendientes || []).filter(r => {
    const f = getDate_(r, colFechaVenc, 'FechaVencimiento', 'FechaFacturaCliente', 'FechaFacturaProveedor', 'Fecha');
    const saldo = getNum_(r, colSaldoUsd, 'SaldoUSD', 'SaldoClienteUSD', 'SaldoProveedorDivisa');
    return f instanceof Date && !isNaN(f.getTime()) && f < hoy && saldo > 0;
  });

  return { 
    n: venc.length, 
    usd: venc.reduce((s, r) => s + getNum_(r, colSaldoUsd, 'SaldoUSD', 'SaldoClienteUSD', 'SaldoProveedorDivisa'), 0) 
  };
}
function leerFrases_() {
  const sh = libro_().getSheetByName(CONFIG.HOJA_CINTA);
  if (!sh || sh.getLastRow() < 1) return [];
  return sh.getDataRange().getDisplayValues().flat().map(s => String(s).trim()).filter(Boolean);
}



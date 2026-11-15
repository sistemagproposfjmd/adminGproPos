

/**
 * ============================================================
 * SCRIPT DE INVENTARIO DE HOJAS — Ejecutar 1 vez
 * ============================================================
 * Devuelve en el log:
 *  - Lista de hojas
 *  - Filas/columnas
 *  - Encabezados
 *  - Si tiene datos, fórmulas, validaciones
 *  - Si está protegida
 *
 * Uso: Ejecutar → listarHojasSistema → Ver → Registros
 * ============================================================
 */
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

/* ============================================================
   MANUAL DE USUARIO — Backend
   ============================================================ */
function obtenerRolUsuarioManual(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName('Usuarios');

    // Si no existe la hoja → acceso total
    if (!hoja) return { rol: 'admin', modulos: null };

    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { rol: 'operador', modulos: null };

    const headers = datos[0].map(h => String(h).trim());
    const idx = {
      username: headers.indexOf('username'),
      correo: headers.indexOf('correo'),
      role: headers.indexOf('role'),
      status: headers.indexOf('status')
    };

    const emailLower = String(email || '').toLowerCase().trim();
    let fila = null;

    // Buscar por correo primero, luego por username
    for (let i = 1; i < datos.length; i++) {
      const correo = idx.correo > -1 ? String(datos[i][idx.correo] || '').toLowerCase().trim() : '';
      const username = idx.username > -1 ? String(datos[i][idx.username] || '').toLowerCase().trim() : '';
      if (correo === emailLower || username === emailLower) {
        fila = datos[i];
        break;
      }
    }

    // Usuario no encontrado
    if (!fila) return { rol: 'operador', modulos: null };

    // Verificar status activo
    const status = idx.status > -1 ? String(fila[idx.status] || '').toLowerCase().trim() : '';
    if (status && !['activo', 'active', 'true', '1', 'si', 'sí'].includes(status)) {
      return { rol: 'inactivo', modulos: [] }; // No ve nada
    }

    // Rol
    const role = idx.role > -1 ? String(fila[idx.role] || 'operador').toLowerCase().trim() : 'operador';

    // Admin → todos los módulos
    if (role === 'admin' || role === 'administrador') {
      return { rol: 'admin', modulos: null };
    }

    // Mapeo columna del Sheet → módulo del manual
    const MAPA_PERMISOS = {
      cargarFacturaCompra: 'compras',
      proveedoresLista: 'compras',
      pagosProveedores: 'compras',
      cxcycxp: 'compras',
      comprasMasiva: 'compras',
      cargarFacturaVentas: 'ventas',
      clientesLista: 'ventas',
      pedidos: 'ventas',
      pedidosVendedor: 'ventas',
      productosLista: 'inventario',
      ajusteInventario: 'inventario',
      asientos: 'Contable',
      contabilizacion: 'Contabilidad',
      cuentaContableLista: 'contabilidad',
      impuestos: 'reportes',
      reportes: 'reportes',
      facturasLista: 'reportes',
      bancosCajas: 'dashboard',
      cobrarYpagar: 'dashboard'
    };

    const modulosPermitidos = new Set();

    // Siempre incluir módulos básicos (siempre visibles)
    modulosPermitidos.add('introduccion');
    modulosPermitidos.add('autenticacion');
    modulosPermitidos.add('carga');
    modulosPermitidos.add('practicas');

    // Evaluar cada permiso
    Object.keys(MAPA_PERMISOS).forEach(columna => {
      const i = headers.indexOf(columna);
      if (i === -1) return; // Columna no existe
      const val = String(fila[i] || '').toLowerCase().trim();
      const activo = ['true', '1', 'si', 'sí', 'x', 'activo', 'yes'].includes(val);
      if (activo) modulosPermitidos.add(MAPA_PERMISOS[columna]);
    });

    return {
      rol: role,
      modulos: Array.from(modulosPermitidos)
    };

  } catch (err) {
    console.error('Error obteniendo rol manual:', err);
    return { rol: 'admin', modulos: null };
  }
}
function obtenerManualDesdeSheet(lang) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(lang === 'en' ? 'Manual_EN' : 'Manual_ES');
    if (!hoja) return null;

    const filas = hoja.getDataRange().getValues();
    filas.shift();
    const modulos = {};

    filas.forEach(([modulo, orden, titulo, detalle, items, nota, alerta, descModulo, icono, tabla]) => {
      if (!modulo) return;
      const id = _slugManual(modulo);
      if (!modulos[id]) {
        modulos[id] = {
          id, nombre: modulo, icono: icono || '📘',
          descripcion: descModulo || '', pasos: []
        };
      }
      const paso = {
        orden: Number(orden) || 999, titulo, detalle,
        items: items ? String(items).split('|').map(s => s.trim()).filter(Boolean) : [],
        nota: nota || '', alerta: alerta || ''
      };
      if (tabla) paso.tabla = _parsearTablaManual(tabla);
      modulos[id].pasos.push(paso);
    });

    return Object.values(modulos).map(m => {
      m.pasos.sort((a, b) => a.orden - b.orden);
      m.pasos.forEach(p => delete p.orden);
      return m;
    });
  } catch (err) {
    console.error('Error manual desde Sheet:', err);
    return null;
  }
}
function _parsearTablaManual(str) {
  const partes = String(str).split('||').map(s => s.trim()).filter(Boolean);
  if (partes.length < 2) return null;
  return {
    headers: partes[0].split(';').map(s => s.trim()),
    rows: partes.slice(1).map(f => f.split(';').map(s => s.trim()))
  };
}
function _slugManual(txt) {
  return String(txt).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}
function getManualData(lang) {
  return obtenerManualDesdeSheet(lang);
}
function getConfigClienteManual() {
  const email = Session.getActiveUser().getEmail() || 'anonimo';
  return { email, rol: obtenerRolUsuarioManual(email), lang: 'es' };
}
function construirManifestManual() {
  return {
    name: 'Manual de Usuario GProPOS',
    short_name: 'GProPOS',
    description: 'Manual de usuario oficial del sistema GProPOS.',
    start_url: '?page=manual',
    scope: './',
    display: 'standalone',
    orientation: 'portrait-primary',
    background_color: '#0f1115',
    theme_color: '#ff6b35',
    lang: 'es-VE',
    categories: ['business', 'productivity', 'education'],
    icons: [
      { src: 'data:image/svg+xml;base64,' + _base64IconoManual('#ff6b35', '#ff8c5a', 'GP'), sizes: '192x192', type: 'image/svg+xml', purpose: 'any' },
      { src: 'data:image/svg+xml;base64,' + _base64IconoManual('#ff6b35', '#ff8c5a', 'GP'), sizes: '512x512', type: 'image/svg+xml', purpose: 'any' },
      { src: 'data:image/svg+xml;base64,' + _base64IconoManual('#0f1115', '#ff6b35', 'GP', true), sizes: '512x512', type: 'image/svg+xml', purpose: 'maskable' }
    ],
    shortcuts: [
      { name: 'Inventario', short_name: 'Inventario', url: '?page=manual#/inventario' },
      { name: 'Ventas', short_name: 'Ventas', url: '?page=manual#/ventas' },
      { name: 'Compras', short_name: 'Compras', url: '?page=manual#/compras' }
    ]
  };
}
function _base64IconoManual(bg, accent, texto, mask) {
  const padding = mask ? 20 : 10;
  const size = 512;
  const inner = size - padding * 2;
  const font = inner * 0.42;
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${size} ${size}">
      <defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${bg}"/><stop offset="100%" stop-color="${accent}"/>
      </linearGradient></defs>
      <rect width="${size}" height="${size}" fill="url(#g)" rx="${mask ? 0 : 100}"/>
      <text x="50%" y="50%" font-family="system-ui, sans-serif" font-size="${font}" font-weight="800"
            fill="white" text-anchor="middle" dominant-baseline="central">${texto}</text>
    </svg>`.trim();
  return Utilities.base64Encode(svg);
}
function testRolManual() {
  const emails = [
    'admin@empresa.com',
    'contador@empresa.com',
    'vendedor@empresa.com',
    Session.getActiveUser().getEmail()
  ];

  emails.forEach(email => {
    const r = obtenerRolUsuarioManual(email);
    Logger.log('📧 %s → rol: %s · módulos: %s',
      email, r.rol, JSON.stringify(r.modulos));
  });
}
function diagnosticoManual() {
  const out = [];
  
  try {
    out.push('=== INICIO DIAGNÓSTICO ===');
    
    // 1. Verificar que el archivo existe
    try {
      const raw = HtmlService.createHtmlOutputFromFile('manual').getContent();
      out.push('1. Archivo "manual" existe: SÍ');
      out.push('2. Tamaño raw: ' + raw.length + ' chars');
      out.push('3. Raw tiene <?= ?: ' + raw.includes('<?='));
      out.push('4. Raw primeros 200 chars:');
      out.push(raw.substring(0, 200));
    } catch (err) {
      out.push('❌ ERROR cargando archivo manual: ' + err.message);
      Logger.log(out.join('\n'));
      return;
    }
    
    // 2. Intentar evaluar como template
    try {
      const tpl = HtmlService.createTemplateFromFile('manual');
      tpl.lang = 'es';
      tpl.userEmail = 'test@test.com';
      tpl.rol = { rol: 'admin', modulos: null };
      tpl.manifestData = construirManifestManual();
      
      out.push('');
      out.push('5. Template creado OK');
      
      const html = tpl.evaluate().getContent();
      
      out.push('6. HTML evaluado: ' + html.length + ' chars');
      out.push('7. ¿Quedan <?= ?: ' + html.includes('<?='));
      out.push('8. ¿Quedan <?!= ?: ' + html.includes('<?!='));
      
      // Buscar el body
      const idx = html.indexOf('<body');
      if (idx > -1) {
        out.push('9. Body encontrado en posición: ' + idx);
        out.push('10. Body tag: ' + html.substring(idx, idx + 100));
      } else {
        out.push('9. ❌ NO se encontró <body>');
      }
      
      // Buscar la línea 18 (aproximada)
      const lines = html.split('\n');
      out.push('11. Total líneas: ' + lines.length);
      if (lines[17]) out.push('12. Línea 18: ' + lines[17].substring(0, 150));
      
    } catch (err) {
      out.push('');
      out.push('❌ ERROR evaluando template:');
      out.push('   Mensaje: ' + err.message);
      out.push('   Stack: ' + (err.stack || 'sin stack'));
    }
    
  } catch (err) {
    out.push('');
    out.push('❌ ERROR GENERAL: ' + err.message);
    out.push('   Stack: ' + (err.stack || 'sin stack'));
  }
  
  // Imprimir todo junto
  Logger.log(out.join('\n'));
}
function testServirManual() {
  const out = [];
  
  try {
    // Simular lo mismo que hace el doGet
    const lang = 'es';
    const user = 'sistemagpropos@gmail.com';
    
    const tpl = HtmlService.createTemplateFromFile('manual');
    tpl.lang = lang;
    tpl.userEmail = user;
    tpl.rol = obtenerRolUsuarioManual(user);
    tpl.manifestData = construirManifestManual();
    
    const html = tpl.evaluate().getContent();
    
    out.push('=== TEST SERVIR MANUAL ===');
    out.push('1. HTML length: ' + html.length);
    out.push('2. ¿Tiene "data-lang=\\"es\\"": ' + html.includes('data-lang="es"'));
    out.push('3. ¿Tiene "lang is not defined"?: ' + html.includes('lang is not defined'));
    out.push('4. ¿Tiene "<?=": ' + html.includes('<?='));
    out.push('5. ¿Tiene "manifestData" sin evaluar?: ' + html.includes('manifestData'));
    
    // Buscar el body
    const bodyIdx = html.indexOf('<body');
    out.push('6. Body: ' + html.substring(bodyIdx, bodyIdx + 120));
    
    // Buscar si hay scripts con abrirManual
    const hasAbrirManual = html.includes('abrirManual');
    out.push('7. ¿Tiene función abrirManual?: ' + hasAbrirManual);
    
    // Buscar si hay window.URL_APP
    const hasURLApp = html.includes('URL_APP');
    out.push('8. ¿Tiene URL_APP?: ' + hasURLApp);
    
  } catch (err) {
    out.push('❌ ERROR: ' + err.message);
  }
  
  Logger.log(out.join('\n'));
}
function obtenerResumenAlertasSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const alertas = [];
  let totalAlertas = 0;

  try {
    // -------------------------------------------------------------------------
    // 1 & 2. HOJA PEDIDOS (Pedidos por Facturar y Pedidos con más de 25 días)
    // -------------------------------------------------------------------------
    const shPedidos = ss.getSheetByName("PEDIDOS") || ss.getSheetByName("Pedidos");
    if (shPedidos && shPedidos.getLastRow() > 1) {
      const data = shPedidos.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());

      const colStatus = headers.indexOf("StatusPedido");
      const colFecha = headers.indexOf("FechaRegistroPedido") !== -1 
                        ? headers.indexOf("FechaRegistroPedido") 
                        : headers.indexOf("fechaFacturaClientePedido");
      const colIdPed = headers.indexOf("idPedido") !== -1 
                        ? headers.indexOf("idPedido") 
                        : headers.indexOf("IdNroPedido");

      let cantPorFacturar = 0;
      let cantMas25Dias = 0;
      let ejemplos25 = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = colStatus !== -1 ? String(row[colStatus]).trim().toLowerCase() : "";

        if (status === "pendiente" || status === "por facturar") {
          cantPorFacturar++;

          if (colFecha !== -1 && row[colFecha]) {
            const fPed = new Date(row[colFecha]);
            if (!isNaN(fPed.getTime())) {
              const diffDias = Math.floor((hoy - fPed) / (1000 * 60 * 60 * 24));
              if (diffDias > 25) {
                cantMas25Dias++;
                if (ejemplos25.length < 2) {
                  const idDoc = colIdPed !== -1 ? row[colIdPed] : i;
                  ejemplos25.push(`N° ${idDoc} (${diffDias}d)`);
                }
              }
            }
          }
        }
      }

      if (cantPorFacturar > 0) {
        alertas.push({
          id: 'pedidos_facturar',
          titulo: 'Pedidos por Facturar',
          subtitulo: `${cantPorFacturar} pedido(s) en espera de emisión`,
          cantidad: cantPorFacturar,
          icono: 'fas fa-file-invoice-dollar',
          colorClass: 'warn-orange',
          pagina: 'pPedidos'
        });
      }

      if (cantMas25Dias > 0) {
        alertas.push({
          id: 'pedidos_25dias',
          titulo: 'Pedidos > 25 Días',
          subtitulo: `${cantMas25Dias} crítico(s): ${ejemplos25.join(', ')}`,
          cantidad: cantMas25Dias,
          icono: 'fas fa-clock',
          colorClass: 'warn-red',
          pagina: 'pPedidos'
        });
      }
    }

    // -------------------------------------------------------------------------
    // 3 & 4. HOJA PRODUCTOS (Stock Cero y Stock Mínimo)
    // -------------------------------------------------------------------------
    const shProd = ss.getSheetByName("PRODUCTOS") || ss.getSheetByName("Productos");
    if (shProd && shProd.getLastRow() > 1) {
      const data = shProd.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());

      const colNombre = headers.indexOf("NombreProducto");
      const colStock = headers.indexOf("StockProducto");
      const colMin = headers.indexOf("MinimoProducto");

      let cantCero = 0;
      let cantMin = 0;
      let ejemplosCero = [];
      let ejemplosMin = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (colStock === -1) break;

        const stock = parseFloat(String(row[colStock]).replace(",", ".")) || 0;
        const minimo = colMin !== -1 ? (parseFloat(String(row[colMin]).replace(",", ".")) || 0) : 0;
        const nom = colNombre !== -1 ? String(row[colNombre]).trim() : `Prod ${i}`;

        if (stock <= 0) {
          cantCero++;
          if (ejemplosCero.length < 2) ejemplosCero.push(nom);
        } else if (minimo > 0 && stock <= minimo) {
          cantMin++;
          if (ejemplosMin.length < 2) ejemplosMin.push(`${nom} (${stock})`);
        }
      }

      if (cantCero > 0) {
        alertas.push({
          id: 'stock_cero',
          titulo: 'Productos Stock Cero',
          subtitulo: `${cantCero} agotado(s): ${ejemplosCero.join(', ')}`,
          cantidad: cantCero,
          icono: 'fas fa-box-open',
          colorClass: 'warn-red',
          pagina: 'pProductosLista'
        });
      }

      if (cantMin > 0) {
        alertas.push({
          id: 'stock_minimo',
          titulo: 'Stock Mínimo o Crítico',
          subtitulo: `${cantMin} por agotarse: ${ejemplosMin.join(', ')}`,
          cantidad: cantMin,
          icono: 'fas fa-triangle-exclamation',
          colorClass: 'warn-amber',
          pagina: 'pProductosLista'
        });
      }
    }

    // -------------------------------------------------------------------------
    // 5. FACTURAS VENCIDAS POR COBRAR
    // -------------------------------------------------------------------------
    const shFact = ss.getSheetByName("FacturasVentas") || ss.getSheetByName("Facturas") || ss.getSheetByName("Ventas");
    if (shFact && shFact.getLastRow() > 1) {
      const data = shFact.getDataRange().getValues();
      const headers = data[0].map(h => String(h).toLowerCase().trim());

      const colTipo = headers.findIndex(h => h.includes("tipo") || h.includes("transaccion"));
      const colSaldo = headers.findIndex(h => h.includes("saldo") || h.includes("pendiente"));
      const colVenc = headers.findIndex(h => h.includes("vencimiento") || h.includes("fecha venc"));
      const colStatus = headers.findIndex(h => h.includes("estado") || h.includes("estatus"));

      let cantVencidas = 0;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tipo = colTipo !== -1 ? String(row[colTipo]).toLowerCase() : "venta";
        const saldo = colSaldo !== -1 ? (parseFloat(String(row[colSaldo]).replace(",", ".")) || 0) : 0;
        const estatus = colStatus !== -1 ? String(row[colStatus]).toLowerCase() : "";

        if ((tipo.includes("venta") || tipo === "") && (saldo > 0.01 || (!estatus.includes("pagad") && !estatus.includes("anulad")))) {
          if (colVenc !== -1 && row[colVenc]) {
            const fVenc = new Date(row[colVenc]);
            if (!isNaN(fVenc.getTime()) && fVenc < hoy) {
              cantVencidas++;
            }
          }
        }
      }

      if (cantVencidas > 0) {
        alertas.push({
          id: 'facturas_vencidas',
          titulo: 'Facturas Vencidas',
          subtitulo: `${cantVencidas} documento(s) con cobro expirado`,
          cantidad: cantVencidas,
          icono: 'fas fa-hand-holding-dollar',
          colorClass: 'warn-red',
          pagina: 'pCxcCxp'
        });
      }
    }

    // -------------------------------------------------------------------------
    // 6. CLIENTES CON DATOS INCOMPLETOS
    // -------------------------------------------------------------------------
    const shCli = ss.getSheetByName("CLIENTES") || ss.getSheetByName("Clientes");
    if (shCli && shCli.getLastRow() > 1) {
      const data = shCli.getDataRange().getValues();
      const headers = data[0].map(h => String(h).toLowerCase().trim());

      const colDoc = headers.findIndex(h => h.includes("rif") || h.includes("cedula") || h.includes("dni") || h.includes("documento"));
      const colTel = headers.findIndex(h => h.includes("telf") || h.includes("telefono") || h.includes("celular"));
      const colEmail = headers.findIndex(h => h.includes("correo") || h.includes("email"));
      const colNombre = headers.findIndex(h => h.includes("nombre") || h.includes("cliente"));

      let cantIncompletos = 0;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const nom = colNombre !== -1 ? String(row[colNombre]).trim() : "";
        if (!nom) continue;

        const doc = colDoc !== -1 ? String(row[colDoc]).trim() : "";
        const tel = colTel !== -1 ? String(row[colTel]).trim() : "";
        const mail = colEmail !== -1 ? String(row[colEmail]).trim() : "";

        if (!doc || !tel || !mail || !mail.includes("@")) {
          cantIncompletos++;
        }
      }

      if (cantIncompletos > 0) {
        alertas.push({
          id: 'clientes_incompletos',
          titulo: 'Clientes Incompletos',
          subtitulo: `${cantIncompletos} cliente(s) sin RIF, Teléfono o Email`,
          cantidad: cantIncompletos,
          icono: 'fas fa-user-xmark',
          colorClass: 'warn-blue',
          pagina: 'pClientesLista'
        });
      }
    }

    totalAlertas = alertas.reduce((acc, curr) => acc + curr.cantidad, 0);

    return {
      success: true,
      totalAlertas: totalAlertas,
      alertas: alertas
    };

  } catch (error) {
    console.error("Error en obtenerResumenAlertasSistema:", error);
    return { success: false, error: error.message, totalAlertas: 0, alertas: [] };
  }
}
function calcularVencimientoFactura(fechaEmision, condicionODias) {
  try {
    if (!fechaEmision) {
      throw new Error("La fecha de emisión es obligatoria.");
    }

    // 1. Parsear la fecha de emisión de forma segura
    let fechaBase;
    if (fechaEmision instanceof Date) {
      fechaBase = new Date(fechaEmision.getTime());
    } else if (typeof fechaEmision === "string") {
      // Soporta "YYYY-MM-DD" o "DD/MM/YYYY"
      if (fechaEmision.includes("/")) {
        const parts = fechaEmision.split("/");
        fechaBase = new Date(parts[2], parts[1] - 1, parts[0]);
      } else if (fechaEmision.includes("-")) {
        const parts = fechaEmision.split("-");
        fechaBase = new Date(parts[0], parts[1] - 1, parts[2]);
      } else {
        fechaBase = new Date(fechaEmision);
      }
    }

    if (isNaN(fechaBase.getTime())) {
      throw new Error("Formato de fecha de emisión inválido.");
    }

    // 2. Determinar los días de crédito según el parámetro
    let dias = 0;
    if (typeof condicionODias === "number") {
      dias = condicionODias;
    } else if (typeof condicionODias === "string") {
      const texto = condicionODias.trim().toLowerCase();
      if (texto.includes("contado") || texto === "0") {
        dias = 0;
      } else {
        // Extrae el primer número que encuentre en el texto (ej. "Crédito 30 días" -> 30)
        const match = texto.match(/\d+/);
        dias = match ? parseInt(match[0], 10) : 0;
      }
    }

    // 3. Sumar los días a la fecha base
    const fechaVenc = new Date(fechaBase.getTime());
    fechaVenc.setDate(fechaVenc.getDate() + dias);

    // 4. Comparar contra el día de hoy (sin horas para precisión de días exactos)
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    const fVencComparar = new Date(fechaVenc.getTime());
    fVencComparar.setHours(0, 0, 0, 0);

    const diffTiempo = fVencComparar.getTime() - hoy.getTime();
    const diasRestantes = Math.round(diffTiempo / (1000 * 60 * 60 * 24));

    // Formatear salidas
    const yyyy = fechaVenc.getFullYear();
    const mm = String(fechaVenc.getMonth() + 1).padStart(2, '0');
    const dd = String(fechaVenc.getDate()).padStart(2, '0');

    const fechaISO = `${yyyy}-${mm}-${dd}`;       // Para inputs HTML type="date"
    const fechaFormat = `${dd}/${mm}/${yyyy}`;    // Para tablas o reportes

    let status = "Vigente";
    let color = "success";

    if (diasRestantes < 0) {
      status = `Vencida hace ${Math.abs(diasRestantes)} día(s)`;
      color = "danger";
    } else if (diasRestantes === 0) {
      status = "Vence Hoy";
      color = "warning";
    } else {
      status = `Vence en ${diasRestantes} día(s)`;
      color = "info";
    }

    return {
      success: true,
      fechaVencimientoISO: fechaISO,       // "2026-10-15"
      fechaVencimientoFormat: fechaFormat, // "15/10/2026"
      diasCredito: dias,
      diasRestantes: diasRestantes,
      estaVencida: diasRestantes < 0,
      statusTexto: status,
      colorBootstrap: color
    };

  } catch (error) {
    console.error("Error en calcularVencimientoFactura:", error);
    return {
      success: false,
      error: error.message
    };
  }
}
// =============================================
//   O     C    R   &   C O M P R A S   (V2.0)
// =============================================
function guardarLoteComprasOCR(lote) {
  if (!lote || !lote.length) {
    return { status: "error", message: "No se recibieron datos para guardar." };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetCompras = ss.getSheetByName("Compras");
  if (!sheetCompras) {
    return { status: "error", message: "Hoja 'Compras' no encontrada." };
  }

  // Cargar índice de duplicados existentes en la hoja
  var duplicadosExistentes = _cmCargarDuplicadosExistentes_(sheetCompras);

  var guardadas = 0;
  var asientosGenerados = [];
  var duplicadosDetectados = [];
  var erroresPorFactura = [];
  var idsGenerados = [];

  // Claves vistas dentro de este lote (para detectar duplicados internos)
  var clavesEnLote = {};

  for (var i = 0; i < lote.length; i++) {
    var item = lote[i];
    var nroFacturaLabel = (item && item.nroFacturaProveedor) ? item.nroFacturaProveedor : "(sin número)";

    try {
      // ------------------------------------------------------------
      // 1. Normalizar RIF y fecha
      // ------------------------------------------------------------
      var rifNorm = _cmNormalizarRIF_(item.rifProveedor || "");
      var nroFactura = String(item.nroFacturaProveedor || "").trim();
      var nroControl = String(item.nroControlProveedor || "").trim();

      // ------------------------------------------------------------
      // 2. Detectar duplicados
      // ------------------------------------------------------------
      var clave = rifNorm + "|" + nroFactura;
      if (rifNorm && nroFactura) {
        if (duplicadosExistentes[clave]) {
          duplicadosDetectados.push({
            nroFactura: nroFactura,
            rif: rifNorm,
            motivo: "Ya existe en la hoja Compras"
          });
          continue;
        }
        if (clavesEnLote[clave]) {
          duplicadosDetectados.push({
            nroFactura: nroFactura,
            rif: rifNorm,
            motivo: "Duplicado dentro del mismo lote"
          });
          continue;
        }
        clavesEnLote[clave] = true;
      }

      // ------------------------------------------------------------
      // 3. Calcular montos y exento
      // ------------------------------------------------------------
      var subTotalBs = parseFloat(item.subTotalProveedor) || 0;
      var ivaBs      = parseFloat(item.ivaProveedor) || 0;
      var totalBs    = parseFloat(item.totalProveedor) || 0;
      var tasa       = parseFloat(item.tasaProveedor) || 1;

      // Si la factura es exenta (IVA = 0 y subtotal > 0)
      var esExento = (item.tipoExento === "EXENTO") || (ivaBs === 0 && subTotalBs > 0);
      var totalExento = esExento ? subTotalBs : 0;

      // Coherencia: si total y subtotal no cuadran, ajustar
      if (!esExento && totalBs > 0 && subTotalBs > 0) {
        if (Math.abs(subTotalBs + ivaBs - totalBs) > 0.5 && ivaBs === 0) {
          subTotalBs = parseFloat((totalBs / 1.16).toFixed(2));
          ivaBs = parseFloat((totalBs - subTotalBs).toFixed(2));
        }
      }

      // Divisa
      var subTotalDiv = parseFloat((subTotalBs / tasa).toFixed(2));
      var ivaDiv      = parseFloat((ivaBs / tasa).toFixed(2));
      var totalDiv    = parseFloat((totalBs / tasa).toFixed(2));

      // ------------------------------------------------------------
      // 4. Fecha en formato DD/MM/YYYY (como espera el backend)
      // ------------------------------------------------------------
      var fechaDDMMYYYY = _cmNormalizarFechaDDMMYYYY_(item.fechaFacturaProveedor);

      // ------------------------------------------------------------
      // 5. Item genérico (G000001 GENERICO)
      // ------------------------------------------------------------
      var itemGenerico = {
        NroFactura: nroFactura,
        Producto: "GENERICO",
        Categoria: "",
        Cantidad: 1,
        Precio: subTotalBs,
        PorcentajeIVA: esExento ? 0 : 16,
        MontoIVA: ivaBs,
        PrecioTotal: totalBs,
        PrecioUSD: subTotalDiv,
        PorcentajeIVAUSD: esExento ? 0 : 16,
        MontoIVAUSD: ivaDiv,
        PrecioTotalUSD: totalDiv,
        CuentaContable: "120505001",
        Cuenta: "INVENTARIO DE SUMINISTROS",
        SubCuenta: "INVENTARIO",
        Transaccion: "Compra"
      };

      // ------------------------------------------------------------
      // 6. Armar facturaData con los nombres EXACTOS que espera
      //    guardarDatosFacturaCompras
      // ------------------------------------------------------------
      var facturaData = {
        // Identificación del proveedor
        NombreProveedor: String(item.nombreProveedor || "").trim(),
        RifProveedor: rifNorm,
        FechaFacturaProveedor: fechaDDMMYYYY,
        TipoProveedor: String(item.tipoProveedor || "Factura"),

        // Documento
        NroFacturaProveedor: nroFactura,
        NroControlProveedor: nroControl,

        // Moneda y tasa
        MonedaProveedor: String(item.monedaProveedor || "Bolivares"),
        TasaProveedor: tasa,

        // Montos Bs
        SubTotalProveedor: subTotalBs,
        totalExentoCompra: totalExento,
        IvaProveedor: ivaBs,
        TotalProveedor: totalBs,
        SaldoProveedor: totalBs,

        // Montos divisa
        SubTotalProveedorDivisa: subTotalDiv,
        IVAProveedorDivisa: ivaDiv,
        TotalProveedorDivisa: totalDiv,
        SaldoProveedorDivisa: totalDiv,

        // Metadata
        TransaccionProveedor: "Compra",
        tipoTransaccion: "Compra",
        cobradaPagada: "porPagar",
        Abono: 0,
        fechaAbono: "",
        descripcionCompra: String(item.descripcionCompra || "COMPRA DIVERSA"),

        // Items (un solo genérico)
        items: [itemGenerico]
      };

      // ------------------------------------------------------------
      // 7. Delegar a guardarDatosFacturaCompras
      // ------------------------------------------------------------
      var resultado = guardarDatosFacturaCompras(facturaData);

      if (resultado && resultado.status === "success") {
        guardadas++;
        if (resultado.nroAsiento) {
          asientosGenerados.push({
            nroFactura: nroFactura,
            nroAsiento: resultado.nroAsiento
          });
        }
        if (resultado.idCompra) {
          idsGenerados.push({
            nroFactura: nroFactura,
            idCompra: resultado.idCompra
          });
        }
        // Registrar en índice para evitar duplicados subsiguientes
        if (rifNorm && nroFactura) {
          duplicadosExistentes[rifNorm + "|" + nroFactura] = true;
        }
      } else {
        erroresPorFactura.push({
          nroFactura: nroFacturaLabel,
          error: (resultado && resultado.message) ? resultado.message : "Error desconocido"
        });
      }

    } catch (errItem) {
      Logger.log("❌ Error procesando factura " + nroFacturaLabel + ": " + errItem.message);
      erroresPorFactura.push({
        nroFactura: nroFacturaLabel,
        error: errItem.message
      });
    }
  }

  Logger.log(
    "✅ guardarLoteComprasOCR: " + guardadas + " guardadas, " +
    asientosGenerados.length + " asientos, " +
    duplicadosDetectados.length + " duplicados, " +
    erroresPorFactura.length + " errores."
  );

  return {
    status: "success",
    guardadas: guardadas,
    asientos: asientosGenerados,
    duplicados: duplicadosDetectados,
    errores: erroresPorFactura,
    ids: idsGenerados
  };
}
function procesarOcrLoteServidor(loteBase64) {
  var resultados = [];
  
  if (typeof Drive === "undefined" || !Drive.Files || !Drive.Files.insert) {
    return [{ status: "error", message: "Error: Debes habilitar el Servicio Avanzado de Drive API v2 en el editor de Apps Script." }];
  }

  for (var i = 0; i < loteBase64.length; i++) {
    var fileId = null;
    try {
      var parts = loteBase64[i].split(",");
      if (parts.length < 2) {
        resultados.push({ status: "error", message: "Formato de imagen inválido." });
        continue;
      }
      
      var bytes = Utilities.base64Decode(parts[1]);
      var blob = Utilities.newBlob(bytes, "image/jpeg", "ocr_temp_" + new Date().getTime() + ".jpg");

      // 1. Insertar con OCR activado
      var resource = { 
        title: blob.getName(), 
        mimeType: blob.getContentType() 
      };
      
      var ocrFile = Drive.Files.insert(resource, blob, { 
        ocr: true, 
        ocrLanguage: "es", 
        convert: true 
      });
      
      fileId = ocrFile.id;

      // 2. Extraer texto del documento generado
      var doc = DocumentApp.openById(fileId);
      var text = doc.getBody().getText();
      
      resultados.push({ status: "success", text: text });
      
    } catch (e) {
      resultados.push({ status: "error", message: e.toString() });
    } finally {
      // 3. Destrucción permanente para no llenar la papelera ni agotar almacenamiento
      if (fileId) {
        try {
          Drive.Files.remove(fileId);
        } catch (errDelete) {
          DriveApp.getFileById(fileId).setTrashed(true);
        }
      }
    }
  }
  return resultados;
}
function guardarLoteComprasOCR(lote) {
  try {
    console.log('📥 Recibiendo lote de', lote ? lote.length : 0, 'facturas');
    
    if (!lote || !lote.length) {
      return { status: "error", message: "No se recibieron datos para guardar." };
    }

    // ============================================
    // 1. GUARDAR EN HOJA "Compras"
    // ============================================
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Compras");
    
    // Verificar que exista la hoja, si no, crearla
    if (!sheet) {
      sheet = ss.insertSheet("Compras");
      var headers = [
        "NombreProveedor", "RifProveedor", "FechaFacturaProveedor", "TipoProveedor",
        "NroFacturaProveedor", "NroControlProveedor", "MonedaProveedor", "TasaProveedor",
        "SubTotalProveedor", "totalExentoCompra", "IvaProveedor", "TotalProveedor",
        "SaldoProveedor", "SubTotalProveedorDivisa", "IVAProveedorDivisa",
        "TotalProveedorDivisa", "SaldoProveedorDivisa", "TransaccionProveedor",
        "Id", "cobradaPagada", "Abono", "fechaAbono", "Contabilidad", "nroAsiento",
        "descripcionCompra", "Status", "fechaRetencion", "RetencionIva", "RetencionISLR",
        "nroComprobantes"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
    }
    
    // ============================================
    // 2. PREPARAR FILAS PARA LA HOJA "Compras"
    // ============================================
    var filasNuevas = [];
    var asientosGenerados = [];
    var nroAsientoActual = null;

    lote.forEach(function(item, index) {
      console.log('📝 Procesando factura', index + 1, '-', item.nroFacturaProveedor || item.NroFacturaProveedor);
      
      // Usar los nombres de campo correctos (tanto mayúsculas como minúsculas)
      var nombreProv = item.nombreProveedor || item.NombreProveedor || '';
      var rifProv = item.rifProveedor || item.RifProveedor || '';
      var fechaFact = item.fechaFacturaProveedor || item.FechaFacturaProveedor || '';
      var tipoProv = item.tipoProveedor || item.TipoProveedor || 'Factura';
      var nroFact = item.nroFacturaProveedor || item.NroFacturaProveedor || '';
      var nroControl = item.nroControlProveedor || item.NroControlProveedor || '';
      var moneda = item.monedaProveedor || item.MonedaProveedor || 'Dolares';
      var tasa = parseFloat(item.tasaProveedor || item.TasaProveedor || 1);
      var subTotal = parseFloat(item.subTotalProveedor || item.SubTotalProveedor || 0);
      var exento = parseFloat(item.totalExentoCompra || 0);
      var iva = parseFloat(item.ivaProveedor || item.IvaProveedor || 0);
      var total = parseFloat(item.totalProveedor || item.TotalProveedor || 0);
      var saldo = parseFloat(item.saldoProveedor || item.SaldoProveedor || total);
      var subTotalDiv = parseFloat(item.subTotalProveedorDivisa || item.SubTotalProveedorDivisa || 0);
      var ivaDiv = parseFloat(item.ivaProveedorDivisa || item.IvaProveedorDivisa || 0);
      var totalDiv = parseFloat(item.totalProveedorDivisa || item.TotalProveedorDivisa || 0);
      var saldoDiv = parseFloat(item.saldoProveedorDivisa || item.SaldoProveedorDivisa || totalDiv);
      var transaccion = item.transaccionProveedor || item.TransaccionProveedor || "Compra";
      var id = item.id || item.Id || 'FAC-COMP-' + Date.now() + '-' + (index + 1);
      var cobrada = item.cobradaPagada || "porPagar";
      var abono = parseFloat(item.abono || item.Abono || 0);
      var fechaAbono = item.fechaAbono || "";
      var contabilidad = item.contabilidad || item.Contabilidad || "SÍ";
      var nroAsiento = item.nroAsiento || "";
      var descripcion = item.descripcionCompra || item.descripcionCompra || 'COMPRA DIVERSA';

      // Crear fila para la hoja "Compras"
      var fila = [
        nombreProv,                  // NombreProveedor
        rifProv,                     // RifProveedor
        fechaFact,                   // FechaFacturaProveedor
        tipoProv,                    // TipoProveedor
        nroFact,                     // NroFacturaProveedor
        nroControl,                  // NroControlProveedor
        moneda,                      // MonedaProveedor
        tasa,                        // TasaProveedor
        subTotal,                    // SubTotalProveedor
        exento,                      // totalExentoCompra
        iva,                         // IvaProveedor
        total,                       // TotalProveedor
        saldo,                       // SaldoProveedor
        subTotalDiv,                 // SubTotalProveedorDivisa
        ivaDiv,                      // IVAProveedorDivisa
        totalDiv,                    // TotalProveedorDivisa
        saldoDiv,                    // SaldoProveedorDivisa
        transaccion,                 // TransaccionProveedor
        id,                          // Id
        cobrada,                     // cobradaPagada
        abono,                       // Abono
        fechaAbono,                  // fechaAbono
        contabilidad,                // Contabilidad
        nroAsiento,                  // nroAsiento
        descripcion,                 // descripcionCompra
        "",                          // Status
        "",                          // fechaRetencion
        "",                          // RetencionIva
        "",                          // RetencionISLR
        ""                           // nroComprobantes
      ];
      
      filasNuevas.push(fila);
      
      // ============================================
      // 3. GENERAR ASIENTO CONTABLE AUTOMÁTICO
      // ============================================
      // Preparar datos para generarAsientoCompra
      var facturaData = {
        NombreProveedor: nombreProv,
        RifProveedor: rifProv,
        FechaFacturaProveedor: fechaFact,
        TipoProveedor: tipoProv,
        NroFacturaProveedor: nroFact,
        NroControlProveedor: nroControl,
        MonedaProveedor: moneda,
        TasaProveedor: tasa,
        SubTotalProveedor: subTotal,
        IvaProveedor: iva,
        TotalProveedor: total,
        SaldoProveedor: saldo,
        SubTotalProveedorDivisa: subTotalDiv,
        IvaProveedorDivisa: ivaDiv,
        TotalProveedorDivisa: totalDiv,
        SaldoProveedorDivisa: saldoDiv,
        TransaccionProveedor: transaccion,
        tipoTransaccion: "Compra",
        cobradaPagada: cobrada,
        Abono: abono,
        fechaAbono: fechaAbono,
        Contabilidad: contabilidad,
        nroAsiento: nroAsiento,
        descripcionCompra: descripcion,
        CuentaContable: item.CuentaContable || '401-01',
        items: item.items || [{
          Producto: descripcion,
          Cantidad: 1,
          Precio: subTotal,
          CuentaContable: item.CuentaContable || '401-01',
          Cuenta: 'Gasto',
          SubCuenta: 'COMPRAS',
          Iva: iva
        }]
      };
      
      console.log('🔄 Generando asiento para factura:', nroFact);
      console.log('   Datos:', JSON.stringify(facturaData, null, 2));
      
      try {
        var resultadoAsiento = generarAsientoCompra(facturaData);
        console.log('✅ Asiento generado:', resultadoAsiento);
        
        if (resultadoAsiento && resultadoAsiento.status === 'success') {
          nroAsientoActual = resultadoAsiento.nroAsiento;
          // Actualizar el número de asiento en la fila
          filasNuevas[filasNuevas.length - 1][23] = nroAsientoActual; // Columna nroAsiento
        } else {
          console.warn('⚠️ No se pudo generar asiento para', nroFact);
        }
      } catch (error) {
        console.error('❌ Error generando asiento:', error.message);
      }
    });

    // ============================================
    // 4. GUARDAR FILAS EN LA HOJA
    // ============================================
    if (filasNuevas.length > 0) {
      var ultimaFila = sheet.getLastRow();
      var startRow = ultimaFila > 1 ? ultimaFila + 1 : 2;
      var totalFilas = filasNuevas.length;
      var totalColumnas = filasNuevas[0].length;
      
      console.log('📊 Guardando', totalFilas, 'filas desde la fila', startRow);
      
      // Asegurar que todas las filas tengan el mismo número de columnas
      var filasFinales = filasNuevas.map(function(fila) {
        while (fila.length < totalColumnas) {
          fila.push('');
        }
        return fila;
      });
      
      sheet.getRange(startRow, 1, totalFilas, totalColumnas).setValues(filasFinales);
    }

    console.log('✅ Proceso completado. Facturas:', filasNuevas.length, '| Asiento:', nroAsientoActual || 'N/A');

    return { 
      status: "success", 
      guardadas: filasNuevas.length,
      nroAsiento: nroAsientoActual || 'N/A'
    };
    
  } catch (e) {
    console.error('❌ Error en guardarLoteComprasOCR:', e.toString());
    return { status: "error", message: e.toString() };
  }
}
function enviarCorreoCorporativo(destinatarios, asunto, htmlBody, adjuntos, replyTo) {
  try {
    if (!destinatarios || destinatarios.trim() === "") {
      throw new Error("No se especificaron destinatarios");
    }

    var opciones = {
      htmlBody: htmlBody,
      name: "Sistema de Facturación"
    };

    if (adjuntos && adjuntos.length > 0) {
      opciones.attachments = adjuntos.filter(function(blob) {
        return blob !== null;
      });
    }

    if (replyTo && replyTo.trim() !== "") {
      opciones.replyTo = replyTo;
    }

    GmailApp.sendEmail(destinatarios, asunto, "Su cliente de correo no soporta HTML. Por favor contacte al emisor.", opciones);

    console.log("✅ Correo enviado a: " + destinatarios);
    return { enviado: true, destinatarios: destinatarios };

  } catch (e) {
    console.error("❌ Error en enviarCorreoCorporativo: " + e.message);
    throw new Error("No se pudo enviar el correo: " + e.message);
  }
}
function enviarCorreoFacturaVenta(nroFactura, facturaData, datosCliente, datosEmpresa) {
  try {
    console.log("📧 Iniciando envío de correo para factura: " + nroFactura);

    // Validar que tenemos lo mínimo
    if (!facturaData) {
      throw new Error("No se recibieron datos de la factura");
    }
    if (!datosCliente) {
      throw new Error("No se encontró el cliente '" + (facturaData.nombreCliente || "desconocido") + "'");
    }

    // --- Construir el objeto que necesitan las funciones de HTML ---
    var datosFactura = {
      nombreCliente: facturaData.nombreCliente || "",
      rifCliente: facturaData.rifCliente || "",
      fechaFacturaCliente: facturaData.fechaFacturaCliente || "",
      nroFacturaCliente: facturaData.nroFacturaCliente || nroFactura,
      nroControlCliente: facturaData.nroControlCliente || "",
      tipoDocumento: facturaData.tipoFacturaCliente || "Factura",
      moneda: facturaData.selectMonedaVenta || "Dolares",
      tasa: parseFloat(facturaData.monedaTasaVenta) || 1,
      subTotal: parseFloat(facturaData.subTotal) || 0,
      iva: parseFloat(facturaData.iva) || 0,
      total: parseFloat(facturaData.total) || 0,
      totalUSD: parseFloat(facturaData.totalClienteUSD) || 0,
      DireccionCliente: datosCliente.direccion || ""
    };

    // Mapear los items al formato que espera generarHtmlCorreoFacturaVenta
    if (facturaData.items && facturaData.items.length > 0) {
      datosFactura.items = facturaData.items.map(function(item) {
        return {
          producto: item.nombre || item.producto || "",
          nombre: item.nombre || item.producto || "",
          cantidad: parseFloat(item.cantidad) || 0,
          precio: parseFloat(item.precio) || 0,
          tasaIva: parseFloat(item.tasaIva) || 0,
          montoIva: parseFloat(item.montoIva) || 0,
          precioTotal: parseFloat(item.precioTotal) || 0
        };
      });
    } else {
      datosFactura.items = [];
    }

    console.log("✅ datosFactura construido. Items: " + datosFactura.items.length);
    console.log("✅ Cliente: " + datosCliente.nombre + " | Correo: " + datosCliente.correo);
    console.log("✅ Empresa: " + datosEmpresa.nombre);

    // --- Generar PDF ---
    var pdfBlob = null;
    try {
      pdfBlob = generarPdfFacturaVenta(datosFactura, datosCliente, datosEmpresa);
      console.log("✅ PDF generado correctamente");
    } catch (pdfError) {
      console.warn("⚠️ No se pudo generar el PDF: " + pdfError.message);
    }

    // --- Generar HTML del correo ---
    var htmlCorreo = generarHtmlCorreoFacturaVenta(datosFactura, datosCliente, datosEmpresa);
    console.log("✅ HTML del correo generado (" + htmlCorreo.length + " chars)");

    // --- Construir destinatarios ---
    var destinatarios = [];

    if (datosCliente.correo && datosCliente.correo.trim() !== "") {
      destinatarios.push(datosCliente.correo.trim());
    }

    if (datosEmpresa.correo && datosEmpresa.correo.trim() !== "" &&
        datosEmpresa.correo.trim() !== datosCliente.correo) {
      destinatarios.push(datosEmpresa.correo.trim());
    }

    if (destinatarios.length === 0) {
      console.warn("⚠️ No hay destinatarios válidos para la factura " + nroFactura);
      return {
        status: "warning",
        message: "Factura guardada pero sin destinatarios de correo configurados."
      };
    }

    // --- Construir asunto ---
    var tipoDoc = datosFactura.tipoDocumento || "Factura";
    var asunto = datosEmpresa.nombre + " - " + tipoDoc + " de Venta N° " + nroFactura + " - " + datosFactura.nombreCliente;

    // --- Enviar correo ---
    var adjuntos = pdfBlob ? [pdfBlob] : [];
    var resultadoEnvio = enviarCorreoCorporativo(
      destinatarios.join(","),
      asunto,
      htmlCorreo,
      adjuntos,
      datosEmpresa.correo || ""
    );

    console.log("📧 Correo enviado exitosamente a: " + destinatarios.join(", "));

    return {
      status: "success",
      message: "Correo enviado a: " + destinatarios.join(", "),
      destinatarios: destinatarios,
      pdfGenerado: pdfBlob !== null
    };

  } catch (e) {
    console.error("❌ Error en enviarCorreoFacturaVenta: " + e.message);
    return { status: "error", message: "Error al enviar correo: " + e.message };
  }
}
function obtenerDatosClientePorNombre(nombreCliente) {
  try {
    var sheet = getSheetByName("Clientes");
    if (!sheet) {
      console.warn("⚠️ Hoja 'Clientes' no encontrada");
      return null;
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var colMap = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    for (var i = 1; i < data.length; i++) {
      var nombreEnHoja = String(data[i][colMap["NombreClientes"]] || "").trim();
      if (nombreEnHoja === String(nombreCliente).trim()) {
        return {
          id: data[i][colMap["idClientes"]] || "",
          nombre: data[i][colMap["NombreClientes"]] || "",
          rif: data[i][colMap["RifClientes"]] || "",
          direccion: data[i][colMap["DireccionCliente"]] || "",
          correo: data[i][colMap["CorreoCliente"]] || "",
          telefono: data[i][colMap["TelefonoCliente"]] || "",
          tipoContribuyente: data[i][colMap["TipoContribuyenteCliente"]] || "",
          retIva: data[i][colMap["RetIVACliente"]] || 0,
          retIslr: data[i][colMap["RetISLRCliente"]] || 0,
          responsable: data[i][colMap["Responsable"]] || "",
          status: data[i][colMap["Status"]] || ""
        };
      }
    }

    console.warn("⚠️ Cliente no encontrado: " + nombreCliente);
    return null;
  } catch (e) {
    console.error("❌ Error en obtenerDatosClientePorNombre: " + e.message);
    return null;
  }
}
function obtenerDatosCabeceraFactura(nroFactura) {
  try {
    var sheet = getSheetByName("Ventas");
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colMap["NroFacturaCliente"]] || "").trim() === String(nroFactura).trim()) {
        return {
          tipoDocumento: data[i][colMap["TipoCliente"]] || "Factura",
          moneda: data[i][colMap["MonedaCliente"]] || "Dolares",
          tasa: parseFloat(data[i][colMap["TasaCliente"]]) || 1,
          subTotal: parseFloat(data[i][colMap["SubTotalCliente"]]) || 0,
          iva: parseFloat(data[i][colMap["IvaCliente"]]) || 0,
          total: parseFloat(data[i][colMap["TotalCliente"]]) || 0,
          totalUSD: parseFloat(data[i][colMap["TotalClienteUSD"]]) || 0,
          status: data[i][colMap["Status"]] || "Activa"
        };
      }
    }
    return null;
  } catch (e) {
    console.error("❌ Error en obtenerDatosCabeceraFactura: " + e.message);
    return null;
  }
}
function formatNumberCorreo(num) {
  if (num === null || num === undefined || isNaN(num)) return "0,00";
  return parseFloat(num).toLocaleString('es-VE', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}
function generarHtmlCorreoFacturaVenta(datosFactura, datosCliente, datosEmpresa) {

  var subTotal = 0;
  var totalIva = 0;
  var totalGeneral = 0;
  var filasItems = "";

  if (datosFactura.items && datosFactura.items.length > 0) {
    datosFactura.items.forEach(function(item, index) {
      var precioTotal = parseFloat(item.precioTotal) || 0;
      var montoIva = parseFloat(item.montoIva) || 0;
      var precioUnitario = parseFloat(item.precio) || 0;
      var cantidad = parseFloat(item.cantidad) || 0;

      subTotal += (precioTotal - montoIva);
      totalIva += montoIva;
      totalGeneral += precioTotal;

      var bgColor = index % 2 === 0 ? "#ffffff" : "#f8fafc";

      filasItems += "<tr style='background-color:" + bgColor + ";'>"
        + "<td style='padding:11px 14px; border-bottom:1px solid #edf2f7; text-align:center; font-size:13px; color:#718096;'>" + (index + 1) + "</td>"
        + "<td style='padding:11px 14px; border-bottom:1px solid #edf2f7; font-size:13px; color:#2d3748; font-weight:500;'>" + (item.producto || item.nombre || '') + "</td>"
        + "<td style='padding:11px 14px; border-bottom:1px solid #edf2f7; text-align:center; font-size:13px; color:#718096;'>" + cantidad + "</td>"
        + "<td style='padding:11px 14px; border-bottom:1px solid #edf2f7; text-align:right; font-size:13px; color:#718096;'>" + formatNumberCorreo(precioUnitario) + "</td>"
        + "<td style='padding:11px 14px; border-bottom:1px solid #edf2f7; text-align:right; font-size:13px; color:#718096;'>" + formatNumberCorreo(precioTotal - montoIva) + "</td>"
        + "<td style='padding:11px 14px; border-bottom:1px solid #edf2f7; text-align:right; font-size:13px; color:#718096;'>" + formatNumberCorreo(montoIva) + "</td>"
        + "<td style='padding:11px 14px; border-bottom:1px solid #edf2f7; text-align:right; font-size:13px; color:#2d3748; font-weight:600;'>" + formatNumberCorreo(precioTotal) + "</td>"
        + "</tr>";
    });
  }

  if (datosFactura.subTotal) subTotal = datosFactura.subTotal;
  if (datosFactura.iva) totalIva = datosFactura.iva;
  if (datosFactura.total) totalGeneral = datosFactura.total;

  var tasa = datosFactura.tasa || datosFactura.monedaTasaVenta || 1;
  var totalUsd = datosFactura.totalUSD || (tasa > 0 ? totalGeneral / tasa : 0);
  var tipoDoc = datosFactura.tipoDocumento || "Factura";
  var fechaFormateada = datosFactura.fechaFacturaCliente || "";
  var simboloMoneda = datosFactura.moneda === "Bolivares" ? "Bs" : "$";

  var logoHtml = "";
  if (datosEmpresa.logo && datosEmpresa.logo.trim() !== "") {
    logoHtml = "<img src='" + datosEmpresa.logo + "' alt='Logo' style='max-height:65px; max-width:180px; object-fit:contain;'>";
  } else {
    logoHtml = "<div style='font-size:20px; font-weight:700; color:#1a365d; letter-spacing:1px;'>" + (datosEmpresa.nombre || "EMPRESA") + "</div>";
  }

  var C_PRIMARIO = "#1a365d";
  var C_SECUNDARIO = "#2c5282";
  var C_ACENTO = "#3182ce";
  var C_FONDO = "#f0f4f8";
  var C_BORDE = "#e2e8f0";
  var C_TEXTO = "#2d3748";
  var C_TEXTO_SEC = "#718096";

  var html = "<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body style='margin:0; padding:0; background-color:" + C_FONDO + "; font-family:\"Segoe UI\",Roboto,Helvetica,Arial,sans-serif;'>"
    + "<table width='100%' cellpadding='0' cellspacing='0' style='background-color:" + C_FONDO + "; padding:30px 0;'>"
    + "<tr><td align='center'>"
    + "<table width='700' cellpadding='0' cellspacing='0' style='background-color:#ffffff; border-radius:10px; overflow:hidden; box-shadow:0 6px 24px rgba(0,0,0,0.07);'>"
    + "<tr><td style='background:linear-gradient(135deg," + C_PRIMARIO + "," + C_SECUNDARIO + "); padding:0; height:5px;'></td></tr>"
    + "<tr><td style='padding:32px 40px 20px 40px;'>"
    + "<table width='100%' cellpadding='0' cellspacing='0'><tr>"
    + "<td style='vertical-align:middle;'>"
    + logoHtml
    + "<div style='margin-top:8px; font-size:11px; color:" + C_TEXTO_SEC + "; line-height:1.6;'>RIF: " + (datosEmpresa.rif || "N/A") + "<br>" + (datosEmpresa.direccion || "Dirección no configurada") + "</div></td>"
    + "<td align='right' style='vertical-align:top;'>"
    + "<div style='background:linear-gradient(135deg," + C_PRIMARIO + "," + C_SECUNDARIO + "); color:#fff; padding:14px 28px; border-radius:8px; text-align:center;'>"
    + "<div style='font-size:10px; text-transform:uppercase; letter-spacing:2.5px; opacity:0.8;'>" + tipoDoc + " de Venta</div>"
    + "<div style='font-size:24px; font-weight:700; margin-top:5px;'>" + (datosFactura.nroFacturaCliente || "") + "</div></div>"
    + "<div style='margin-top:10px; font-size:11px; color:" + C_TEXTO_SEC + "; text-align:right;'>Nro. Control: <strong style='color:" + C_TEXTO + ";'>" + (datosFactura.nroControlCliente || "N/A") + "</strong></div></td>"
    + "</tr></table></td></tr>"
    + "<tr><td style='padding:0 40px;'><hr style='border:none; border-top:2px solid " + C_BORDE + "; margin:0;'></td></tr>"
    + "<tr><td style='padding:24px 40px;'>"
    + "<table width='100%' cellpadding='0' cellspacing='0'><tr>"
    + "<td width='55%' style='vertical-align:top;'>"
    + "<div style='font-size:9px; text-transform:uppercase; letter-spacing:2px; color:" + C_ACENTO + "; font-weight:700; margin-bottom:10px;'>Datos del Cliente</div>"
    + "<div style='font-size:15px; font-weight:600; color:" + C_TEXTO + "; margin-bottom:5px;'>" + (datosCliente.nombre || datosFactura.nombreCliente || "") + "</div>"
    + "<div style='font-size:12px; color:" + C_TEXTO_SEC + "; line-height:1.8;'>RIF: " + (datosCliente.rif || datosFactura.rifCliente || "") + "<br>" + (datosCliente.direccion || datosFactura.DireccionCliente || "N/A") + (datosCliente.telefono ? "<br>Tel: " + datosCliente.telefono : "") + (datosCliente.responsable ? "<br>Responsable: <strong>" + datosCliente.responsable + "</strong>" : "") + "</div></td>"
    + "<td width='45%' align='right' style='vertical-align:top;'>"
    + "<div style='font-size:9px; text-transform:uppercase; letter-spacing:2px; color:" + C_ACENTO + "; font-weight:700; margin-bottom:10px;'>Información de Emisión</div>"
    + "<div style='font-size:13px; color:#4a5568; line-height:2.2;'><strong>Fecha:</strong> " + fechaFormateada + "<br><strong>Moneda:</strong> " + (datosFactura.moneda || "Dolares") + "<br><strong>Tasa:</strong> " + formatNumberCorreo(tasa) + " Bs/$</div></td>"
    + "</tr></table></td></tr>"
    + "<tr><td style='padding:0 40px 24px 40px;'>"
    + "<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid " + C_BORDE + "; border-radius:8px; overflow:hidden;'>"
    + "<tr style='background:linear-gradient(135deg," + C_PRIMARIO + "," + C_SECUNDARIO + ");'>"
    + "<th style='padding:13px 14px; text-align:center; font-size:10px; color:#fff; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>#</th>"
    + "<th style='padding:13px 14px; text-align:left; font-size:10px; color:#fff; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>Descripción</th>"
    + "<th style='padding:13px 14px; text-align:center; font-size:10px; color:#fff; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>Cant.</th>"
    + "<th style='padding:13px 14px; text-align:right; font-size:10px; color:#fff; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>Precio Unit.</th>"
    + "<th style='padding:13px 14px; text-align:right; font-size:10px; color:#fff; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>Base Imp.</th>"
    + "<th style='padding:13px 14px; text-align:right; font-size:10px; color:#fff; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>IVA</th>"
    + "<th style='padding:13px 14px; text-align:right; font-size:10px; color:#fff; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>Total</th></tr>"
    + filasItems
    + "</table></td></tr>"
    + "<tr><td style='padding:0 40px 24px 40px;'>"
    + "<table width='350' cellpadding='0' cellspacing='0' align='right' style='border:1px solid " + C_BORDE + "; border-radius:8px; overflow:hidden;'>"
    + "<tr style='background-color:#f7fafc;'><td style='padding:11px 18px; font-size:12px; color:" + C_TEXTO_SEC + ";'>Sub Total:</td><td style='padding:11px 18px; font-size:13px; color:" + C_TEXTO + "; text-align:right; font-weight:500;'>" + formatNumberCorreo(subTotal) + " " + simboloMoneda + "</td></tr>"
    + "<tr><td style='padding:11px 18px; font-size:12px; color:" + C_TEXTO_SEC + ";'>IVA:</td><td style='padding:11px 18px; font-size:13px; color:" + C_TEXTO + "; text-align:right; font-weight:500;'>" + formatNumberCorreo(totalIva) + " " + simboloMoneda + "</td></tr>"
    + "<tr style='background:linear-gradient(135deg," + C_PRIMARIO + "," + C_SECUNDARIO + ");'><td style='padding:14px 18px; font-size:14px; color:#fff; font-weight:700;'>TOTAL:</td><td style='padding:14px 18px; font-size:17px; color:#fff; text-align:right; font-weight:700;'>" + formatNumberCorreo(totalGeneral) + " " + simboloMoneda + "</td></tr>"
    + "<tr style='background-color:#edf2f7;'><td style='padding:9px 18px; font-size:11px; color:#a0aec0;'>Equiv. USD:</td><td style='padding:9px 18px; font-size:12px; color:" + C_TEXTO_SEC + "; text-align:right;'>$ " + formatNumberCorreo(totalUsd) + "</td></tr>"
    + "</table></td></tr>"
    + "<tr><td style='padding:0 40px 24px 40px;'>"
    + "<div style='background-color:#ebf8ff; border-left:4px solid " + C_ACENTO + "; padding:14px 20px; border-radius:0 8px 8px 0;'>"
    + "<div style='font-size:12px; color:#2c5282; line-height:1.7;'><strong style='color:" + C_PRIMARIO + ";'>📄 Documento Adjunto:</strong> Se adjunta al presente correo la " + tipoDoc.toLowerCase() + " de venta N° " + (datosFactura.nroFacturaCliente || "") + " en formato PDF para su verificación y archivo.</div></div></td></tr>"
    + "<tr><td style='padding:0 40px;'><hr style='border:none; border-top:1px solid " + C_BORDE + "; margin:0;'></td></tr>"
    + "<tr><td style='padding:20px 40px 28px 40px;'>"
    + "<table width='100%' cellpadding='0' cellspacing='0'><tr>"
    + "<td style='vertical-align:top;'><div style='font-size:10px; color:#a0aec0; line-height:1.8;'>Este correo fue generado automáticamente por el sistema de facturación.<br>Por favor no responder a este correo de forma directa.</div></td>"
    + "<td align='right' style='vertical-align:top;'><div style='font-size:13px; font-weight:600; color:" + C_PRIMARIO + ";'>" + (datosEmpresa.nombre || "") + "</div><div style='font-size:10px; color:#a0aec0; margin-top:3px; line-height:1.7;'>" + (datosEmpresa.rif || "") + (datosEmpresa.telefono ? "<br>" + datosEmpresa.telefono : "") + (datosEmpresa.correo ? "<br>" + datosEmpresa.correo : "") + "</div></td>"
    + "</tr></table></td></tr>"
    + "<tr><td style='background:linear-gradient(135deg," + C_PRIMARIO + "," + C_SECUNDARIO + "); padding:0; height:5px;'></td></tr>"
    + "</table></td></tr></table></body></html>";

  return html;
}
function generarHtmlPdfFacturaVenta(datosFactura, datosCliente, datosEmpresa) {

  var subTotal = datosFactura.subTotal || 0;
  var totalIva = datosFactura.iva || 0;
  var totalGeneral = datosFactura.total || 0;
  var tasa = datosFactura.tasa || datosFactura.monedaTasaVenta || 1;
  var totalUsd = datosFactura.totalUSD || (tasa > 0 ? totalGeneral / tasa : 0);
  var tipoDoc = datosFactura.tipoDocumento || "Factura";
  var fechaFormateada = datosFactura.fechaFacturaCliente || "";
  var simboloMoneda = datosFactura.moneda === "Bolivares" ? "Bs" : "$";

  var filasItems = "";
  if (datosFactura.items && datosFactura.items.length > 0) {
    datosFactura.items.forEach(function(item, index) {
      var precioTotal = parseFloat(item.precioTotal) || 0;
      var montoIva = parseFloat(item.montoIva) || 0;
      var precioUnitario = parseFloat(item.precio) || 0;
      var cantidad = parseFloat(item.cantidad) || 0;
      var tasaIvaItem = parseFloat(item.tasaIva) || 0;
      var bgRow = index % 2 === 0 ? "#ffffff" : "#f9fafb";

      filasItems += "<tr style='background-color:" + bgRow + ";'>"
        + "<td style='padding:8px 10px; border:1px solid #d1d5db; text-align:center; font-size:11px; color:#374151;'>" + (index + 1) + "</td>"
        + "<td style='padding:8px 10px; border:1px solid #d1d5db; font-size:11px; color:#111827;'>" + (item.producto || item.nombre || "") + "</td>"
        + "<td style='padding:8px 10px; border:1px solid #d1d5db; text-align:center; font-size:11px; color:#374151;'>" + cantidad + "</td>"
        + "<td style='padding:8px 10px; border:1px solid #d1d5db; text-align:right; font-size:11px; color:#374151;'>" + formatNumberCorreo(precioUnitario) + "</td>"
        + "<td style='padding:8px 10px; border:1px solid #d1d5db; text-align:center; font-size:11px; color:#374151;'>" + tasaIvaItem + "%</td>"
        + "<td style='padding:8px 10px; border:1px solid #d1d5db; text-align:right; font-size:11px; color:#374151;'>" + formatNumberCorreo(montoIva) + "</td>"
        + "<td style='padding:8px 10px; border:1px solid #d1d5db; text-align:right; font-size:11px; color:#111827; font-weight:600;'>" + formatNumberCorreo(precioTotal) + "</td>"
        + "</tr>";
    });
  }

  var logoHtml = "";
  if (datosEmpresa.logo && datosEmpresa.logo.trim() !== "") {
    logoHtml = "<img src='" + datosEmpresa.logo + "' style='max-height:55px; max-width:160px;'>";
  } else {
    logoHtml = "<span style='font-size:18px; font-weight:700; color:#111827;'>" + (datosEmpresa.nombre || "EMPRESA") + "</span>";
  }

  var html = "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
    + "<style>"
    + "@page { size: letter; margin: 12mm 15mm; }"
    + "body { font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #111827; margin: 0; padding: 0; }"
    + "table { border-collapse: collapse; width: 100%; }"
    + ".borde-exterior { border: 2px solid #1f2937; padding: 20px; }"
    + ".titulo-doc { background-color: #1f2937; color: white; text-align: center; padding: 8px; font-size: 14px; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; margin-bottom: 15px; }"
    + ".datos-box { border: 1px solid #9ca3af; padding: 10px; margin-bottom: 12px; }"
    + ".datos-label { font-size: 9px; text-transform: uppercase; letter-spacing: 1px; color: #6b7280; font-weight: 700; margin-bottom: 6px; border-bottom: 1px solid #e5e7eb; padding-bottom: 4px; }"
    + ".total-box { border: 2px solid #1f2937; width: 280px; }"
    + ".total-fila { display: flex; justify-content: space-between; padding: 6px 12px; }"
    + ".total-fila-grande { background-color: #1f2937; color: white; padding: 10px 12px; font-size: 14px; font-weight: 700; }"
    + ".footer-line { border-top: 2px solid #1f2937; margin-top: 20px; padding-top: 10px; text-align: center; font-size: 9px; color: #6b7280; }"
    + ".firma-line { border-top: 1px solid #374151; width: 200px; margin: 40px auto 5px auto; }"
    + ".firma-text { text-align: center; font-size: 10px; color: #374151; }"
    + "</style></head><body>"

    + "<div class='borde-exterior'>"

    // Header
    + "<table style='margin-bottom:10px;'><tr>"
    + "<td style='vertical-align:top; width:55%; padding-right:15px;'>"
    + logoHtml
    + "<div style='margin-top:6px; font-size:10px; color:#374151; line-height:1.6;'>"
    + "<strong>RIF:</strong> " + (datosEmpresa.rif || "N/A") + "<br>"
    + (datosEmpresa.direccion || "")
    + (datosEmpresa.telefono ? "<br><strong>Tel:</strong> " + datosEmpresa.telefono : "")
    + (datosEmpresa.correo ? "<br><strong>Email:</strong> " + datosEmpresa.correo : "")
    + "</div></td>"
    + "<td style='vertical-align:top; text-align:right; width:45%;'>"
    + "<div class='titulo-doc'>" + tipoDoc + "</div>"
    + "<table style='border:1px solid #374151; margin-left:auto;'>"
    + "<tr><td style='padding:5px 12px; font-size:10px; color:#6b7280; border-bottom:1px solid #d1d5db;'>N° Factura:</td><td style='padding:5px 12px; font-size:13px; font-weight:700; color:#111827; border-bottom:1px solid #d1d5db;'>" + (datosFactura.nroFacturaCliente || "") + "</td></tr>"
    + "<tr><td style='padding:5px 12px; font-size:10px; color:#6b7280; border-bottom:1px solid #d1d5db;'>N° Control:</td><td style='padding:5px 12px; font-size:13px; font-weight:700; color:#111827; border-bottom:1px solid #d1d5db;'>" + (datosFactura.nroControlCliente || "N/A") + "</td></tr>"
    + "<tr><td style='padding:5px 12px; font-size:10px; color:#6b7280;'>Fecha Emisión:</td><td style='padding:5px 12px; font-size:12px; font-weight:600; color:#111827;'>" + fechaFormateada + "</td></tr>"
    + "</table></td></tr></table>"

    // Cliente
    + "<div class='datos-box'>"
    + "<div class='datos-label'>Datos del Cliente</div>"
    + "<table><tr>"
    + "<td style='width:50%; vertical-align:top;'>"
    + "<div style='margin-bottom:2px;'><strong>Nombre / Razón Social:</strong> " + (datosCliente.nombre || datosFactura.nombreCliente || "") + "</div>"
    + "<div style='margin-bottom:2px;'><strong>RIF:</strong> " + (datosCliente.rif || datosFactura.rifCliente || "") + "</div></td>"
    + "<td style='width:50%; vertical-align:top;'>"
    + "<div style='margin-bottom:2px;'><strong>Dirección:</strong> " + (datosCliente.direccion || datosFactura.DireccionCliente || "N/A") + "</div>"
    + (datosCliente.telefono ? "<div style='margin-bottom:2px;'><strong>Teléfono:</strong> " + datosCliente.telefono + "</div>" : "")
    + (datosCliente.responsable ? "<div style='margin-bottom:2px;'><strong>Responsable:</strong> " + datosCliente.responsable + "</div>" : "")
    + "</td></tr></table></div>"

    // Moneda
    + "<div style='display:flex; gap:20px; margin-bottom:12px;'>"
    + "<div style='border:1px solid #9ca3af; padding:6px 14px; font-size:10px;'><strong>Moneda:</strong> " + (datosFactura.moneda || "Dolares") + "</div>"
    + "<div style='border:1px solid #9ca3af; padding:6px 14px; font-size:10px;'><strong>Tasa de Cambio:</strong> " + formatNumberCorreo(tasa) + " Bs/$</div>"
    + "<div style='border:1px solid #9ca3af; padding:6px 14px; font-size:10px;'><strong>Tipo Contribuyente:</strong> " + (datosCliente.tipoContribuyente || "N/A") + "</div>"
    + "</div>"

    // Items
    + "<table style='margin-bottom:15px;'>"
    + "<tr style='background-color:#1f2937;'>"
    + "<th style='padding:8px 10px; border:1px solid #374151; color:white; font-size:10px; text-align:center; width:30px;'>#</th>"
    + "<th style='padding:8px 10px; border:1px solid #374151; color:white; font-size:10px; text-align:left;'>Descripción</th>"
    + "<th style='padding:8px 10px; border:1px solid #374151; color:white; font-size:10px; text-align:center; width:50px;'>Cant.</th>"
    + "<th style='padding:8px 10px; border:1px solid #374151; color:white; font-size:10px; text-align:right; width:90px;'>P. Unitario</th>"
    + "<th style='padding:8px 10px; border:1px solid #374151; color:white; font-size:10px; text-align:center; width:50px;'>% IVA</th>"
    + "<th style='padding:8px 10px; border:1px solid #374151; color:white; font-size:10px; text-align:right; width:90px;'>Monto IVA</th>"
    + "<th style='padding:8px 10px; border:1px solid #374151; color:white; font-size:10px; text-align:right; width:100px;'>Total</th></tr>"
    + filasItems + "</table>"

    // Totales
    + "<div style='display:flex; justify-content:flex-end;'><div class='total-box'>"
    + "<div class='total-fila' style='background-color:#f9fafb; border-bottom:1px solid #d1d5db;'><span style='font-size:11px; color:#374151;'>Sub Total:</span><span style='font-size:12px; font-weight:600; color:#111827;'>" + formatNumberCorreo(subTotal) + " " + simboloMoneda + "</span></div>"
    + "<div class='total-fila' style='border-bottom:1px solid #d1d5db;'><span style='font-size:11px; color:#374151;'>IVA:</span><span style='font-size:12px; font-weight:600; color:#111827;'>" + formatNumberCorreo(totalIva) + " " + simboloMoneda + "</span></div>"
    + "<div class='total-fila-grande'><span>TOTAL:</span><span>" + formatNumberCorreo(totalGeneral) + " " + simboloMoneda + "</span></div>"
    + "<div class='total-fila' style='background-color:#f3f4f6;'><span style='font-size:10px; color:#6b7280;'>Equiv. USD:</span><span style='font-size:11px; color:#374151;'>$ " + formatNumberCorreo(totalUsd) + "</span></div>"
    + "</div></div>"

    // Firmas
    + "<table style='margin-top:35px; width:100%;'><tr>"
    + "<td style='text-align:center; width:50%;'><div class='firma-line'></div><div class='firma-text'>Firma y Sello del Emisor</div></td>"
    + "<td style='text-align:center; width:50%;'><div class='firma-line'></div><div class='firma-text'>Firma del Receptor</div></td>"
    + "</tr></table>"

    // Footer
    + "<div class='footer-line'>"
    + (datosEmpresa.nombre || "") + " | RIF: " + (datosEmpresa.rif || "") + " | " + (datosEmpresa.direccion || "")
    + "<br>Documento generado por el Sistema de Facturación | Fecha de impresión: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
    + "</div>"

    + "</div></body></html>";

  return html;
}
function generarPdfFacturaVenta(datosFactura, datosCliente, datosEmpresa) {
  try {
    var htmlContent = generarHtmlPdfFacturaVenta(datosFactura, datosCliente, datosEmpresa);

    var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(780)
      .setHeight(1100);

    var blob = htmlOutput.getBlob();
    var pdfBlob = blob.getAs(MimeType.PDF);
    pdfBlob.setName("Factura_Venta_" + (datosFactura.nroFacturaCliente || "sin_nro") + ".pdf");

    console.log("✅ PDF generado: " + pdfBlob.getName() + " (" + (pdfBlob.getBytes().length / 1024).toFixed(1) + " KB)");
    return pdfBlob;

  } catch (e) {
    console.error("❌ Error generando PDF: " + e.message);
    console.error("Stack: " + e.stack);
    return null;
  }
}

// ============================================================
// HELPERS DE COMPRAS MASIVA
// ============================================================
function _cmNormalizarRIF_(rif) {
  if (!rif) return "";
  var c = String(rif).toUpperCase().replace(/[^JVGE0-9]/g, "");
  var m = c.match(/^([JVGE])(\d{8,9})$/);
  if (!m) return String(rif).toUpperCase();
  var L = m[1], d = m[2];
  if (d.length === 9) return L + "-" + d.slice(0, 8) + "-" + d.slice(8);
  return L + "-" + d.slice(0, 7) + "-" + d.slice(7);
}
function _cmNormalizarFechaDDMMYYYY_(fecha) {
  if (!fecha) return "";
  try {
    if (fecha instanceof Date) {
      return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    var s = String(fecha).trim();
    // ISO YYYY-MM-DD
    var iso = s.match(/^(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})/);
    if (iso) {
      var y = iso[1], mm = String(parseInt(iso[2], 10)).padStart(2, "0"), dd = String(parseInt(iso[3], 10)).padStart(2, "0");
      return dd + "/" + mm + "/" + y;
    }
    // DD/MM/YYYY o DD-MM-YYYY
    var lat = s.match(/^(\d{1,2})[\-\/\.](\d{1,2})[\-\/\.](\d{2,4})/);
    if (lat) {
      var d = String(parseInt(lat[1], 10)).padStart(2, "0");
      var m = String(parseInt(lat[2], 10)).padStart(2, "0");
      var yy = parseInt(lat[3], 10);
      if (yy < 100) yy += 2000;
      return d + "/" + m + "/" + yy;
    }
    return s;
  } catch (e) {
    return "";
  }
}
function _cmCargarDuplicadosExistentes_(sheet) {
  var map = {};
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;
    // Columnas: 2 = RifProveedor, 5 = NroFacturaProveedor
    var valores = sheet.getRange(2, 2, lastRow - 1, 4).getValues();
    for (var i = 0; i < valores.length; i++) {
      var rif = _cmNormalizarRIF_(valores[i][0]);
      var nroFact = String(valores[i][3] || "").trim();
      if (rif && nroFact) {
        map[rif + "|" + nroFact] = true;
      }
    }
  } catch (e) {
    Logger.log("⚠️ Error leyendo duplicados: " + e.message);
  }
  return map;
}

// =============================================
//   FUNCIONALIDAD contable
// =============================================
function obtenerDatosAbonos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAbonos = ss.getSheetByName('CobrosPagos');
    if (!hojaAbonos) { console.warn('obtenerDatosAbonos: no existe CobrosPagos'); return []; }
    
    const lastRow = hojaAbonos.getLastRow();
    if (lastRow < 2) return [];

    // Obtenemos cabeceras y datos
    const headers = hojaAbonos.getRange(1, 1, 1, hojaAbonos.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const data = hojaAbonos.getRange(2, 1, lastRow - 1, headers.length).getValues();

    const idx = {};
    headers.forEach((h, i) => { idx[h] = i; });

    const limpiarNum = (v) => {
      if (v === null || v === undefined || v === '') return 0;
      if (typeof v === 'number') return v;
      let s = String(v).trim();
      let neg = s.startsWith('(') && s.endsWith(')');
      if (neg) s = s.slice(1, -1);
      // Formato español: 1.234,56 -> 1234.56
      if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
      const n = parseFloat(s);
      return isNaN(n) ? 0 : (neg ? -n : n);
    };

    const formatearFecha = (v) => {
      if (!v) return '';
      if (v instanceof Date) {
        return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
      return String(v).trim();
    };

    const resultado = [];

    data.forEach(row => {
      // 1. Identificar si la fila es una Retención
      const transaccion = String(row[idx['Transaccion']] || '').trim();
      const transaccionUpper = transaccion.toUpperCase();
      
      if (!transaccionUpper.includes('RETENCION')) return; // Si no es retención, la saltamos

      // 2. Extraer datos directos de CobrosPagos
      const provCli = String(row[idx['ProveedorCliente']] || '').trim();
      const rif = String(row[idx['Rif']] || '').trim();
      const nroFactura = String(row[idx['NroFactura']] || '').trim();
      const fechaFactura = formatearFecha(row[idx['FechaFactura']]);
      const fechaAbono = formatearFecha(row[idx['fechaAbono']]);
      const comprobante = String(row[idx['id']] || '').trim(); // En tu hoja, "id" tiene el número de comprobante
      const montoUSD = limpiarNum(row[idx['Abono']]); // En tu hoja, el monto está en "Abono"
      const moneda = String(row[idx['Moneda']] || 'USD').trim();
      const tasa = limpiarNum(row[idx['Tasa']]) || 1;

      const montoBs = montoUSD * tasa;

      // 3. Armar el objeto
      resultado.push({
        FechaFactura: fechaFactura,
        Fecha: fechaFactura,
        fechaAbono: fechaAbono, // Útil para el filtro
        Cliente: provCli,
        ProveedorCliente: provCli,
        Rif: rif,
        Factura: nroFactura,
        NroFactura: nroFactura,
        NroControl: '', // Tu hoja no tiene NroControl directo en CobrosPagos
        Referencia: transaccion,
        Transaccion: transaccion,
        MontoBs: montoBs,
        'Monto Bs': montoBs,
        '% Retencion': transaccionUpper.includes('IVA') ? '75' : '5', // Estimado
        Porcentaje: transaccionUpper.includes('IVA') ? '75' : '5',
        'Nro Comprobante': comprobante,
        Comprobante: comprobante,
        id: comprobante,
        Moneda: moneda,
        Tasa: tasa,
        MontoUSD: montoUSD,
        Abono: montoUSD
      });
    });

    console.log('obtenerDatosAbonos: ' + resultado.length + ' retenciones encontradas');
    return resultado;

  } catch (e) {
    console.error('Error en obtenerDatosAbonos:', e.toString());
    return [];
  }
}
function obtenerDatosAsientos() {
  return getDatosAsientos();
}
function obtenerCategoriasConId() {
  try {
    const sheet = getSheetByName('Tasas');
    if (!sheet) {
      console.error("La hoja 'Tasas' no fue encontrada.");
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Leer rango O1:P (columnas 15 y 16)
    const data = sheet.getRange(1, 15, lastRow, 2).getValues();
    const headers = data[0].map(h => String(h).trim());

    const idColIndex = headers.indexOf('IdCategoria');
    const nombreColIndex = headers.indexOf('NombreCategoria');

    if (idColIndex === -1 || nombreColIndex === -1) {
      console.error("No se encontraron las columnas 'IdCategoria' o 'NombreCategoria' en O1:P.");
      return [];
    }

    const categorias = data.slice(1).map(row => {
      const id = row[idColIndex];
      const nombre = row[nombreColIndex];
      if (id && nombre && String(id).trim() !== '' && String(nombre).trim() !== '') {
        return { id: String(id).trim(), nombre: String(nombre).trim() };
      }
      return null;
    }).filter(Boolean);

    console.log('obtenerCategoriasConId: ' + categorias.length + ' categorías');
    return categorias;

  } catch (e) {
    console.error('Error en obtenerCategoriasConId: ' + e.toString());
    return [];
  }
}
function getDatosAsientos() {
  try {
    console.log("🚀 Iniciando getDatosAsientos() desde Web App...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log("📊 Spreadsheet activa:", ss.getName());
    
    const sheet = ss.getSheetByName("Asientos");
    
    if (!sheet) {
      console.error("❌ Hoja 'Asientos' no encontrada");
      console.log("📋 Hojas disponibles:", ss.getSheets().map(s => s.getName()).join(", "));
      return [];
    }
    
    console.log("✅ Hoja 'Asientos' encontrada");
    
    // Obtener todos los datos
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    console.log(`📐 Dimensiones: ${lastRow} filas, ${lastCol} columnas`);
    
    if (lastRow <= 1) {
      console.log("⚠️ La hoja solo tiene encabezados o está vacía");
      return [];
    }
    
    // Obtener encabezados de la fila 1
    const headersRange = sheet.getRange(1, 1, 1, lastCol);
    const headers = headersRange.getValues()[0];
    console.log("📋 Encabezados encontrados:", headers);
    
    // Obtener datos con getValues() - esto evalúa las fórmulas automáticamente
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    
    console.log(`📊 Obteniendo ${values.length} filas de datos (con fórmulas evaluadas)`);
    
    const resultado = [];
    
    for (let i = 0; i < values.length; i++) {
      const fila = values[i];
      const registro = {};
      
      // Verificar si la fila tiene datos mínimos
      let tieneDatos = false;
      
      // Crear objeto con todos los campos
      for (let j = 0; j < headers.length; j++) {
        let valor = fila[j];
        const headerNombre = String(headers[j] || '').trim();
        
        // Convertir fechas a string
        if (valor instanceof Date) {
          valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        
        // Manejar valores null/undefined
        if (valor === undefined || valor === null) {
          valor = '';
        }
        
        // Para números, mantenerlos como números (NO convertir a string con formato)
        // Esto es importante para el parseo correcto en el frontend
        if (typeof valor === 'number') {
          // Mantener como número, no convertir a string
          valor = valor;
        } else if (typeof valor === 'string') {
          // Si es string, mantenerlo como está (puede tener formato español)
          valor = valor.trim();
        }
        
        registro[headerNombre] = valor;
        
        // Verificar si tiene datos importantes
        if (headerNombre === 'CodigoCuenta' && valor && valor.toString().trim() !== '') {
          tieneDatos = true;
        }
      }
      
      // Solo agregar si tiene código de cuenta
      if (tieneDatos) {
        resultado.push(registro);
        
        // Log para depuración de los primeros registros
        if (resultado.length <= 3) {
          console.log(`📝 Registro ${resultado.length}:`, {
            CodigoCuenta: registro['CodigoCuenta'],
            CuentaContable: registro['CuentaContable'],
            MontoTransaccion: registro['MontoTransaccion'],
            MontoType: typeof registro['MontoTransaccion'],
            CUENTA: registro['CUENTA'],
            EEFF: registro['EEFF'],
            TipoCuenta: registro['TipoCuenta'],
            GrupoBalance: registro['GrupoBalance']
          });
        }
      }
    }
    
    console.log(`✅ Procesados ${resultado.length} registros válidos`);
    
    // Verificar que las columnas con fórmula tienen valores
    if (resultado.length > 0) {
      const muestra = resultado[0];
      console.log("🔍 Verificando columnas con fórmula:");
      console.log("  - EEFF:", muestra['EEFF'] || "(vacío)");
      console.log("  - TipoCuenta:", muestra['TipoCuenta'] || "(vacío)");
      console.log("  - GrupoBalance:", muestra['GrupoBalance'] || "(vacío)");
      
      // Contar cuántos tienen valores
      const conEEFF = resultado.filter(r => r['EEFF'] && r['EEFF'].toString().trim() !== '').length;
      const conTipoCuenta = resultado.filter(r => r['TipoCuenta'] && r['TipoCuenta'].toString().trim() !== '').length;
      const conGrupoBalance = resultado.filter(r => r['GrupoBalance'] && r['GrupoBalance'].toString().trim() !== '').length;
      
      console.log(`📊 Población de columnas con fórmula:`);
      console.log(`  - EEFF: ${conEEFF}/${resultado.length} (${Math.round(conEEFF/resultado.length*100)}%)`);
      console.log(`  - TipoCuenta: ${conTipoCuenta}/${resultado.length} (${Math.round(conTipoCuenta/resultado.length*100)}%)`);
      console.log(`  - GrupoBalance: ${conGrupoBalance}/${resultado.length} (${Math.round(conGrupoBalance/resultado.length*100)}%)`);
    }
    
    return resultado;
    
  } catch (error) {
    console.error("❌ ERROR en getDatosAsientos:", error.toString());
    console.error("Stack trace:", error.stack);
    return [];
  }
}
function getTasa() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Resumen") || ss.getSheetByName("Tasas");
    
    if (hoja) {
      // Intentar obtener tasa de varias celdas posibles
      let tasa = hoja.getRange("C1").getValue();
      if (!tasa || isNaN(parseFloat(tasa))) {
        tasa = hoja.getRange("C3").getValue();
      }
      if (!tasa || isNaN(parseFloat(tasa))) {
        tasa = hoja.getRange("B2").getValue();
      }
      
      if (tasa && !isNaN(parseFloat(tasa))) {
        console.log("💰 Tasa obtenida:", parseFloat(tasa));
        return parseFloat(tasa);
      }
    }
  } catch(e) {
    console.error("Error obteniendo tasa:", e);
  }
  
  console.log("💰 Usando tasa por defecto: 43.00");
  return 43.00;
}
function getDatosEmpresa() {
  const NOM_EMPRESA = "Translogistica Transporte y Logistica de Venezuela, C.A.";
  const RIF_EMPRESA = "J-30667902-7";
  const DIR_EMPRESA = "Av. Principal, Edificio Centro, Piso 1, Oficina A, Caracas";
  const LOGO_EMPRESA = "https://lh3.googleusercontent.com/d/1dCncG3S9eZYyFMl5LnkvWFKjV32kgyGX=w120-h120";
  
  return {
    nombre: NOM_EMPRESA,
    rif: RIF_EMPRESA,
    direccion: DIR_EMPRESA,
    logo: LOGO_EMPRESA
  };
}
function testGetDatosAsientos() {
  console.log("========== INICIANDO TEST DESDE EDITOR ==========");
  
  const resultado = getDatosAsientos();
  
  console.log("📊 RESULTADO DEL TEST:");
  console.log("- Tipo:", typeof resultado);
  console.log("- Es array:", Array.isArray(resultado));
  console.log("- Longitud:", resultado ? resultado.length : "N/A");
  
  if (resultado && resultado.length > 0) {
    console.log("- Primer registro completo:");
    console.log(JSON.stringify(resultado[0], null, 2));
    console.log("- ✅ Test exitoso! Se encontraron", resultado.length, "registros");
    
    // Mostrar resumen de tipos de cuenta
    const tipos = {};
    resultado.forEach(r => {
      const tipo = r['CUENTA'] || 'Sin tipo';
      tipos[tipo] = (tipos[tipo] || 0) + 1;
    });
    console.log("📊 Distribución por tipo de cuenta:", tipos);
    
  } else {
    console.log("- ⚠️ No se encontraron registros en la hoja");
  }
  
  console.log("=================================================");
  
  return resultado;
}
function diagnosticarFormulas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Asientos");
    
    if (!sheet) {
      console.log("❌ Hoja Asientos no encontrada");
      return;
    }
    
    console.log("=== DIAGNÓSTICO DE FÓRMULAS ===");
    
    // Verificar columna EEFF (columna N, índice 13)
    const eeffRange = sheet.getRange("N2:N11");
    const eeffValues = eeffRange.getValues();
    console.log("EEFF (col N) - primeros 10 valores:");
    eeffValues.forEach((v, i) => console.log(`  Fila ${i+2}: "${v[0]}"`));
    
    // Verificar columna TipoCuenta (columna O, índice 14)
    const tipoCuentaRange = sheet.getRange("O2:O11");
    const tipoCuentaValues = tipoCuentaRange.getValues();
    console.log("TipoCuenta (col O) - primeros 10 valores:");
    tipoCuentaValues.forEach((v, i) => console.log(`  Fila ${i+2}: "${v[0]}"`));
    
    // Verificar columna GrupoBalance (columna P, índice 15)
    const grupoBalanceRange = sheet.getRange("P2:P11");
    const grupoBalanceValues = grupoBalanceRange.getValues();
    console.log("GrupoBalance (col P) - primeros 10 valores:");
    grupoBalanceValues.forEach((v, i) => console.log(`  Fila ${i+2}: "${v[0]}"`));
    
    // Verificar que las fórmulas existen en la fila 1
    const formulaEEFF = sheet.getRange("N1").getFormula();
    console.log("\nFórmula en N1:", formulaEEFF);
    
    const formulaTipoCuenta = sheet.getRange("O1").getFormula();
    console.log("Fórmula en O1:", formulaTipoCuenta);
    
    const formulaGrupoBalance = sheet.getRange("P1").getFormula();
    console.log("Fórmula en P1:", formulaGrupoBalance);
    
    // Verificar si la hoja PlandeCuentas existe
    const planCuentas = ss.getSheetByName("PlandeCuentas");
    if (planCuentas) {
      console.log("✅ Hoja PlandeCuentas encontrada");
      console.log("  Filas con datos:", planCuentas.getLastRow());
      console.log("  Columnas:", planCuentas.getLastColumn());
    } else {
      console.log("❌ Hoja PlandeCuentas NO encontrada");
    }
    
    console.log("================================");
    
  } catch(e) {
    console.error("Error:", e);
  }
}
function testGetDatosAsientos() {
  console.log("========== TEST COMPLETO ==========");
  
  const resultado = getDatosAsientos();
  
  console.log("=== RESULTADO ===");
  console.log("¿Es array?", Array.isArray(resultado));
  console.log("Longitud:", resultado ? resultado.length : 0);
  
  if (resultado && resultado.length > 0) {
    console.log("Primer registro completo:");
    const primero = resultado[0];
    
    // Mostrar todas las propiedades del primer registro
    for (let key in primero) {
      console.log(`  ${key}: ${primero[key]} (${typeof primero[key]})`);
    }
    
    // Verificar específicamente las columnas que necesitamos
    console.log("\n=== VERIFICACIÓN DE COLUMNAS CRÍTICAS ===");
    console.log("CUENTA:", resultado[0]['CUENTA']);
    console.log("EEFF:", resultado[0]['EEFF']);
    console.log("GrupoBalance:", resultado[0]['GrupoBalance']);
    console.log("TipoCuenta:", resultado[0]['TipoCuenta']);
    console.log("CodigoCuenta:", resultado[0]['CodigoCuenta']);
    console.log("CuentaContable:", resultado[0]['CuentaContable']);
    console.log("MontoTransaccion:", resultado[0]['MontoTransaccion']);
    console.log("Tipo de Monto:", typeof resultado[0]['MontoTransaccion']);
    
    // Contar registros por tipo
    const tipos = {};
    resultado.forEach(r => {
      const tipo = r['CUENTA'] || 'SIN TIPO';
      tipos[tipo] = (tipos[tipo] || 0) + 1;
    });
    console.log("\n=== DISTRIBUCIÓN POR TIPO ===");
    console.log(tipos);
    
    // Contar cuántos tienen EEFF
    const conEEFF = resultado.filter(r => r['EEFF'] && r['EEFF'] !== '').length;
    console.log(`\nRegistros con EEFF: ${conEEFF}/${resultado.length}`);
    
    // Mostrar los primeros 5 montos
    console.log("\n=== PRIMEROS 5 MONTOS ===");
    for (let i = 0; i < Math.min(5, resultado.length); i++) {
      console.log(`Registro ${i+1}: ${resultado[i]['MontoTransaccion']} (${typeof resultado[i]['MontoTransaccion']})`);
    }
    
  } else {
    console.log("❌ No hay datos");
  }
  
  console.log("====================================");
  
  return resultado;
}
function enviarEstadoFinancieroPorCorreo(destinatario, tipoEEFF, moneda, comentarios, mesFiltro, tasaCambio) {
  try {
    // Validaciones básicas
    if (!destinatario || !destinatario.includes('@')) {
      throw new Error('Destinatario inválido');
    }
    
    const mostrarUSD = moneda === 'USD';
    const tasa = tasaCambio || 0.00;
    
    // Obtener y procesar datos
    const datos = obtenerDatosProcesados(mesFiltro);
    
    // Generar contenido HTML
    const htmlContent = generarHtmlEstadoFinanciero(datos, tipoEEFF, mostrarUSD, tasa, comentarios);
    
    // Configurar asunto
    let asunto = `Estado Financiero - ${NOM_EMPRESA}`;
    if (tipoEEFF === 'ES') asunto = `Balance General - ${NOM_EMPRESA}`;
    else if (tipoEEFF === 'ER') asunto = `Estado de Resultado - ${NOM_EMPRESA}`;
    
    if (mesFiltro && mesFiltro !== 'todos') {
      const meses = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
      asunto += ` - ${meses[parseInt(mesFiltro)]}`;
    }
    
    // Enviar correo
MailApp.sendEmail({
  to: destinatario,
  subject: asunto,
  htmlBody: htmlContent,
  name: NOM_EMPRESA
});
    
    return { success: true, message: 'Correo enviado exitosamente' };
    
  } catch (error) {
    console.error('Error al enviar correo:', error);
    return { success: false, message: error.toString() };
  }
}
function obtenerDatosProcesados(mesFiltro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Asientos');
  
  if (!hoja) throw new Error('No se encontró la hoja Asientos');
  
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  
  // Encontrar índices de columnas
  const idxFecha = headers.indexOf('FechaAsiento');
  const idxCodigo = headers.indexOf('CodigoCuenta');
  const idxNombre = headers.indexOf('CuentaContable');
  const idxMonto = headers.indexOf('MontoTransaccion');
  const idxCuenta = headers.indexOf('CUENTA');
  const idxTipo = headers.indexOf('TipoCuenta');
  const idxGrupo = headers.indexOf('GrupoBalance');
  const idxEEFF = headers.indexOf('EEFF');
  
  const cuentas = {};
  
  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    
    // Filtrar por mes si aplica
    if (mesFiltro && mesFiltro !== 'todos' && idxFecha >= 0) {
      const fecha = row[idxFecha];
      if (fecha instanceof Date) {
        if ((fecha.getMonth() + 1) !== parseInt(mesFiltro)) continue;
      } else if (typeof fecha === 'string') {
        let mes = '';
        if (fecha.includes('/')) mes = fecha.split('/')[1];
        else if (fecha.includes('-')) mes = fecha.split('-')[1];
        if (mes && parseInt(mes) !== parseInt(mesFiltro)) continue;
      }
    }
    
    const codigo = String(row[idxCodigo] || 'SIN_CODIGO');
    const monto = parseMontoGs(row[idxMonto]);
    
    if (!cuentas[codigo]) {
      cuentas[codigo] = {
        codigo: codigo,
        nombre: row[idxNombre] || 'Sin nombre',
        cuenta: String(row[idxCuenta] || '').trim(),
        tipoCuenta: String(row[idxTipo] || '').trim(),
        grupoBalance: String(row[idxGrupo] || '').trim(),
        eeff: String(row[idxEEFF] || '').trim(),
        saldo: 0
      };
    }
    
    cuentas[codigo].saldo += monto;
  }
  
  return Object.values(cuentas);
}
function parseMontoGs(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;
  if (typeof valor === 'number') return valor;
  
  let str = String(valor).trim();
  if (str === '') return 0;
  
  let esNegativo = str.startsWith('(') && str.endsWith(')');
  if (esNegativo) str = str.substring(1, str.length - 1);
  
  let tieneSigno = str.startsWith('-');
  if (tieneSigno) str = str.substring(1);
  
  str = str.replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  let num = parseFloat(str);
  if (isNaN(num)) return 0;
  
  return (esNegativo || tieneSigno) ? -num : num;
}
function generarHtmlEstadoFinanciero(datos, tipoEEFF, mostrarUSD, tasa, comentarios) {
  const monedaSymbol = mostrarUSD ? '$' : 'Bs';
  
  let html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body style="margin: 0; padding: 20px; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f8f9fa; color: #212529;">
    <div style="max-width: 900px; margin: 0 auto; background: #ffffff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow: hidden;">
      
      <!-- Header con Logo y Datos de Empresa -->
      <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; color: white;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="width: 80px; vertical-align: middle;">
              <img src="${LOGO_EMPRESA}" style="width: 70px; height: 70px; border-radius: 50%; object-fit: cover; border: 3px solid rgba(255,255,255,0.3);" alt="Logo">
            </td>
            <td style="padding-left: 20px; vertical-align: middle;">
              <h1 style="margin: 0 0 5px 0; font-size: 24px; font-weight: bold;">${NOM_EMPRESA}</h1>
              <p style="margin: 0 0 3px 0; font-size: 14px; opacity: 0.9;">RIF: ${RIF_EMPRESA}</p>
              <p style="margin: 0; font-size: 12px; opacity: 0.8;">${DIR_EMPRESA}</p>
            </td>
            <td style="text-align: right; vertical-align: middle;">
              <div style="background: rgba(255,255,255,0.2); padding: 10px 15px; border-radius: 8px; display: inline-block;">
                <p style="margin: 0; font-size: 12px; opacity: 0.9;">Tasa BCV</p>
                <p style="margin: 0; font-size: 18px; font-weight: bold;">${tasa.toFixed(2)} Bs/$</p>
              </div>
            </td>
          </tr>
        </table>
      </div>
      
      <!-- Contenido -->
      <div style="padding: 30px;">
        
        <!-- Comentarios del usuario -->
        ${comentarios ? `
        <div style="background: #e7f3ff; border-left: 4px solid #2196F3; padding: 15px; margin-bottom: 25px; border-radius: 4px;">
          <p style="margin: 0; font-size: 14px; color: #555; font-style: italic;">
            <strong>Nota:</strong> ${comentarios.replace(/\n/g, '<br>')}
          </p>
        </div>
        ` : ''}
        
        <!-- Estado de Situación -->
        ${(tipoEEFF === 'ES' || tipoEEFF === 'AMBOS') ? generarTablaSeccion(datos, 'ES', mostrarUSD, tasa, monedaSymbol, 'ESTADO DE SITUACIÓN (BALANCE GENERAL)') : ''}
        
        <!-- Estado de Resultado -->
        ${(tipoEEFF === 'ER' || tipoEEFF === 'AMBOS') ? generarTablaSeccion(datos, 'ER', mostrarUSD, tasa, monedaSymbol, 'ESTADO DE RESULTADO') : ''}
        
      </div>
      
      <!-- Footer -->
      <div style="background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #e9ecef;">
        <p style="margin: 0 0 5px 0; font-size: 12px; color: #6c757d;">
          Documento generado el ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}
        </p>
        <p style="margin: 0; font-size: 11px; color: #adb5bd;">
          ${NOM_EMPRESA} - Sistema Contable Integrado
        </p>
      </div>
      
    </div>
  </body>
  </html>
  `;
  
  return html;
}
function generarTablaSeccion(datos, tipoEEFF, mostrarUSD, tasa, monedaSymbol, titulo) {
  const datosFiltrados = datos.filter(d => d.eeff === tipoEEFF);
  
  if (datosFiltrados.length === 0) {
    return `<div style="margin-bottom: 30px;"><h2 style="color: #495057; border-bottom: 2px solid #e9ecef; padding-bottom: 10px;">${titulo}</h2><p style="color: #6c757d;">No hay datos disponibles para esta sección.</p></div>`;
  }
  
  // Estructurar datos igual que en el frontend
  const estructura = {};
  datosFiltrados.forEach(d => {
    const cuenta = d.cuenta || 'SIN CLASIFICAR';
    if (!estructura[cuenta]) estructura[cuenta] = {};
    
    const tipoCuenta = d.tipoCuenta || 'SIN TIPO';
    if (!estructura[cuenta][tipoCuenta]) estructura[cuenta][tipoCuenta] = {};
    
    const grupoBalance = d.grupoBalance || 'SIN GRUPO';
    if (!estructura[cuenta][tipoCuenta][grupoBalance]) estructura[cuenta][tipoCuenta][grupoBalance] = [];
    
    estructura[cuenta][tipoCuenta][grupoBalance].push(d);
  });
  
  let html = `
  <div style="margin-bottom: 40px;">
    <h2 style="color: #495057; border-bottom: 3px solid #667eea; padding-bottom: 10px; margin-bottom: 20px; font-size: 20px;">
      📊 ${titulo}
    </h2>
    
    <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 20px;">
  `;
  
  const ordenCuentas = ['Activo', 'Pasivo', 'Patrimonio', 'Ingresos', 'Costos', 'Gastos', 'SIN CLASIFICAR'];
  let totalGeneral = 0;
  let activoTotal = 0, pasivoTotal = 0, patrimonioTotal = 0;
  
  for (const cuenta of ordenCuentas) {
    if (!estructura[cuenta]) continue;
    
    let totalCuenta = 0;
    html += `
      <tr style="background-color: #f8f9fa;">
        <td colspan="2" style="padding: 12px 15px; font-weight: bold; font-size: 15px; color: #212529; border-top: 2px solid #dee2e6; text-transform: uppercase;">
          ${cuenta}
        </td>
      </tr>
    `;
    
    const tiposCuenta = estructura[cuenta];
    for (const tipoCuenta of Object.keys(tiposCuenta).sort()) {
      let totalTipo = 0;
      html += `
        <tr style="background-color: #ffffff;">
          <td colspan="2" style="padding: 10px 15px 10px 30px; font-weight: 600; color: #495057; border-top: 1px solid #e9ecef;">
            ${tipoCuenta}
          </td>
        </tr>
      `;
      
      const gruposBalance = tiposCuenta[tipoCuenta];
      for (const grupoBalance of Object.keys(gruposBalance).sort()) {
        let totalGrupo = 0;
        html += `
          <tr style="background-color: #ffffff;">
            <td colspan="2" style="padding: 8px 15px 8px 45px; font-weight: 600; color: #6c757d; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">
              ▼ ${grupoBalance}
            </td>
          </tr>
        `;
        
        const cuentasDetalle = gruposBalance[grupoBalance];
        cuentasDetalle.sort((a, b) => a.nombre.localeCompare(b.nombre));
        
        for (const c of cuentasDetalle) {
          totalGrupo += c.saldo;
          const valor = mostrarUSD ? c.saldo / tasa : c.saldo;
          const color = c.saldo < 0 ? '#dc3545' : '#212529';
          const valorFormateado = valor.toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + ' ' + monedaSymbol;
          
          html += `
            <tr style="background-color: #ffffff;">
              <td style="padding: 6px 15px 6px 60px; color: #495057; border-bottom: 1px solid #f1f3f5;">
                ${c.nombre}
              </td>
              <td style="padding: 6px 15px; text-align: right; color: ${color}; font-weight: 500; border-bottom: 1px solid #f1f3f5; white-space: nowrap;">
                ${valorFormateado}
              </td>
            </tr>
          `;
        }
        
        totalTipo += totalGrupo;
        const valorGrupo = mostrarUSD ? totalGrupo / tasa : totalGrupo;
        const colorGrupo = totalGrupo < 0 ? '#dc3545' : '#28a745';
        
        html += `
          <tr style="background-color: #f1f3f5;">
            <td style="padding: 8px 15px 8px 45px; text-align: right; font-weight: bold; color: #495057; font-size: 12px;">
              TOTAL ${grupoBalance}:
            </td>
            <td style="padding: 8px 15px; text-align: right; font-weight: bold; color: ${colorGrupo}; white-space: nowrap;">
              ${valorGrupo.toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${monedaSymbol}
            </td>
          </tr>
        `;
      }
      
      totalCuenta += totalTipo;
      const valorTipo = mostrarUSD ? totalTipo / tasa : totalTipo;
      const colorTipo = totalTipo < 0 ? '#dc3545' : '#212529';
      
      html += `
        <tr style="background-color: #e9ecef;">
          <td style="padding: 10px 15px 10px 30px; text-align: right; font-weight: bold; color: #212529;">
            TOTAL ${tipoCuenta}:
          </td>
          <td style="padding: 10px 15px; text-align: right; font-weight: bold; color: ${colorTipo}; font-size: 14px; white-space: nowrap;">
            ${valorTipo.toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${monedaSymbol}
          </td>
        </tr>
      `;
    }
    
    // Guardar totales para ecuación contable
    if (cuenta === 'Activo') activoTotal = totalCuenta;
    if (cuenta === 'Pasivo') pasivoTotal = totalCuenta;
    if (cuenta === 'Patrimonio') patrimonioTotal = totalCuenta;
    
    totalGeneral += totalCuenta;
    const valorCuenta = mostrarUSD ? totalCuenta / tasa : totalCuenta;
    const colorCuenta = totalCuenta < 0 ? '#dc3545' : '#495057';
    
    html += `
      <tr style="background-color: #dee2e6;">
        <td style="padding: 12px 15px; text-align: right; font-weight: bold; color: #212529; font-size: 15px; text-transform: uppercase;">
          Total ${cuenta}:
        </td>
        <td style="padding: 12px 15px; text-align: right; font-weight: bold; color: ${colorCuenta}; font-size: 15px; white-space: nowrap;">
          ${valorCuenta.toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${monedaSymbol}
        </td>
      </tr>
    `;
  }
  
  // Footer con ecuación contable si es Balance General
  if (tipoEEFF === 'ES') {
    const diferencia = Math.abs(activoTotal) - (Math.abs(pasivoTotal) + Math.abs(patrimonioTotal));
    const esCorrecto = Math.abs(diferencia) < 0.01;
    const colorDiff = esCorrecto ? '#28a745' : '#dc3545';
    const mensaje = esCorrecto ? '✓ Ecuación contable cumplida' : '⚠ Diferencia detectada';
    
    html += `
      <tr style="background-color: ${esCorrecto ? '#d4edda' : '#f8d7da'};">
        <td style="padding: 15px; text-align: right; font-weight: bold; color: #212529; font-size: 14px;">
          ECUACIÓN CONTABLE (Activo - Pasivo - Patrimonio):
        </td>
        <td style="padding: 15px; text-align: right; font-weight: bold; color: ${colorDiff}; font-size: 16px; white-space: nowrap;">
          ${diferencia.toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${monedaSymbol}
          <br><span style="font-size: 11px; color: ${colorDiff};">${mensaje}</span>
        </td>
      </tr>
    `;
  } else {
    // Total para Estado de Resultado
    const valorTotal = mostrarUSD ? totalGeneral / tasa : totalGeneral;
    const colorTotal = totalGeneral < 0 ? '#dc3545' : '#28a745';
    const resultadoTexto = totalGeneral >= 0 ? 'UTILIDAD NETA' : 'PÉRDIDA NETA';
    
    html += `
      <tr style="background-color: #212529; color: white;">
        <td style="padding: 15px; text-align: right; font-weight: bold; font-size: 16px; text-transform: uppercase;">
          ${resultadoTexto}:
        </td>
        <td style="padding: 15px; text-align: right; font-weight: bold; font-size: 18px; color: ${totalGeneral >= 0 ? '#38ef7d' : '#f45c43'}; white-space: nowrap;">
          ${Math.abs(valorTotal).toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${monedaSymbol}
        </td>
      </tr>
    `;
  }
  
  html += `</table></div>`;
  return html;
}
function getEmpresaInfo() {
  return getDatosEmpresa();
}
// =============================================
/** IMPUESTOS                     */
// =============================================
function crearXML() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("abonos");
  const datos = hoja.getDataRange().getValues();
  
  // Definir fechas directamente para pruebas (puedes cambiar estas fechas)
  const fechaDesde = new Date('2026-03-01'); // Cambia esta fecha según sea necesario
  const fechaHasta = new Date('2026-03-14'); // Cambia esta fecha según sea necesario
  
  // Validar fechas
  if (isNaN(fechaDesde) || isNaN(fechaHasta)) {
    Logger.log("Fechas inválidas. Por favor, ingrese fechas en el formato correcto.");
    return;
  }
  
  const RIF_EMPRESA = "J-12345678-9"; // Cambia esto según sea necesario
  const periodo = `${fechaDesde.getFullYear()}${(fechaDesde.getMonth() + 1).toString().padStart(2, '0')}`;
  
  // Crear el XML
  const xmlRoot = `<RelacionRetencionesISLR RifAgente="${RIF_EMPRESA}" Periodo="${periodo}">`;
  let xmlContenido = "";

  for (let i = 1; i < datos.length; i++) { // Comenzar desde 1 para omitir encabezados
    const fila = datos[i];
    const fechaOperacion = new Date(fila[5]); // Columna "Fecha" (índice 5)
    
    if (fila[6] === "RETENCION ISLR" && fechaOperacion >= fechaDesde && fechaOperacion <= fechaHasta) {
      xmlContenido += `
        <DetalleRetencion>
          <RifRetenido>${fila[2]}</RifRetenido>
          <NumeroFactura>${fila[3]}</NumeroFactura>
          <NumeroControl>${fila[4]}</NumeroControl>
          <FechaOperacion>${Utilities.formatDate(fechaOperacion, Session.getScriptTimeZone(), "dd/MM/yyyy")}</FechaOperacion>
          <CodigoConcepto>${fila[6]}</CodigoConcepto>
          <MontoOperacion>${fila[7].toString().replace(',', '.')}</MontoOperacion>
          <PorcentajeRetencion>${fila[8].toString().replace('%', '')}</PorcentajeRetencion>
        </DetalleRetencion>`;
    }
  }

  const xmlFinal = xmlRoot + xmlContenido + "</RelacionRetencionesISLR>";
  
  // Guardar el archivo XML
  const blob = Utilities.newBlob(xmlFinal, 'text/xml', `XML_relacionRetencionesISLR_${periodo}.xml`);
  const carpeta = DriveApp.getRootFolder(); // Cambia esto si necesitas otra carpeta
  carpeta.createFile(blob);
  
  Logger.log("Archivo XML creado con éxito.");
}
function sheetToObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(sheetName);
  
  if (!hoja) {
    console.error(`La hoja "${sheetName}" no fue encontrada.`);
    return []; // Devolver un array vacío si la hoja no existe
  }
  
  const datos = hoja.getDataRange().getValues();
  
  if (datos.length < 2) {
    console.log(`No hay datos suficientes en la hoja "${sheetName}".`);
    return []; // No hay datos para procesar
  }
  
  const headers = datos[0].map(h => h.toString().trim());
  const filas = datos.slice(1);
  
  const resultados = filas.map(fila => {
    let obj = {};
    headers.forEach((header, index) => {
      let valor = fila[index];
      // Si la celda es una fecha, la convertimos a un formato estándar YYYY-MM-DD
      // Esto simplifica mucho el manejo de fechas en el lado del cliente (JavaScript)
      if (valor instanceof Date) {
        obj[header] = valor.toISOString().split('T')[0];
      } else {
        obj[header] = valor;
      }
    });
    return obj;
  }).filter(obj => obj[headers[0]] !== ''); // Filtrar filas vacías basadas en la primera columna

  console.log(`Procesados ${resultados.length} registros de la hoja "${sheetName}".`);
  return resultados;
}
function obtenerCuentasContablesUnicas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Asientos');
    if (!hoja) return [];
    
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return [];
    
    const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => h.trim());
    const colIndex = headers.indexOf('CuentaContable');
    if (colIndex === -1) return [];
    
    const valores = hoja.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
    const cuentasSet = new Set(valores.flat().filter(c => c && c.toString().trim() !== ''));
    
    return Array.from(cuentasSet).sort();
  } catch (error) {
    console.error('Error en obtenerCuentasContablesUnicas:', error);
    return [];
  }
}
function filtrarAsientosPorFechaYCuenta(fechaInicio, fechaFin, cuenta) {
  try {
    const datos = obtenerDatosAsientos(); // Reutilizamos la función principal
    
    const dDesde = new Date(fechaInicio + 'T00:00:00');
    const dHasta = new Date(fechaFin + 'T23:59:59');

    const resultados = datos.filter(fila => {
      let fechaFila;
      const fechaValor = fila.FechaAsiento;

      if (!fechaValor) return false;
      
      if (typeof fechaValor === 'string' && fechaValor.includes('/')) {
        const [dia, mes, anio] = fechaValor.split('/');
        fechaFila = new Date(`${anio}-${mes}-${dia}T00:00:00`);
      } else {
        fechaFila = new Date(fechaValor + 'T00:00:00'); // Asume YYYY-MM-DD
      }

      if (isNaN(fechaFila.getTime())) return false;
      
      const enRango = fechaFila >= dDesde && fechaFila <= dHasta;
      if (!enRango) return false;
      
      if (cuenta && cuenta !== 'TODAS') {
        return fila.CuentaContable === cuenta;
      }
      
      return true;
    });
    
    console.log(`Filtro encontró ${resultados.length} resultados.`);
    return resultados;
  } catch (error) {
    console.error('Error en filtrarAsientosPorFechaYCuenta:', error);
    return [];
  }
}
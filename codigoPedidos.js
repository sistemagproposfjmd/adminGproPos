// ============================================================
// ARCHIVO: codigoPedidos.gs
// MÓDULO UNIFICADO: POS (pedidos.html) + TERMINAL (pedidosVendedor.html)
// ============================================================

var ID_HOJA_DEFAULT = '1bu0s7SQlGHuLkV7nHb_QBIRnXGaZ9IG66bj92rEhUcY'; // Tu hoja de respaldo

function _SS_(idOpcional) {
  // 1. Si la función recibe el ID directamente, usa ese
  if (idOpcional && typeof idOpcional === 'string' && idOpcional.length > 15) {
    return SpreadsheetApp.openById(idOpcional);
  }
  // 2. Si el sistema tiene Contexto de empresa activa en sesión
  if (typeof Contexto !== 'undefined' && Contexto.getEmpresaActivaId) {
    var idActiva = Contexto.getEmpresaActivaId();
    if (idActiva) return Contexto.getEmpresaSS(idActiva);
  }
  // 3. Si está en UserProperties
  var idUser = PropertiesService.getUserProperties().getProperty('EMPRESA_ACTIVA_ID');
  if (idUser) {
    return SpreadsheetApp.openById(idUser);
  }
  // 4. Respaldo por defecto
  return SpreadsheetApp.openById(ID_HOJA_DEFAULT);
}

// Helpers de formato y vista
function getPedidosVendedorHtml() {
  return HtmlService.createTemplateFromFile('pedidosVendedor').evaluate().getContent();
}
function _fmtBs_(n){ return 'Bs. '+(Number(n)||0).toLocaleString('es-VE',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function _fmtUSD_(n){ return '$ '+(Number(n)||0).toLocaleString('es-VE',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function _fmtF_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); }

function getColVal(row, headers, name) {
  var idx = headers.indexOf(name);
  return (idx > -1) ? row[idx] : "";
}
function formatearFechaSafe(fecha) {
  if (!fecha) return "";
  try {
    if (typeof fecha === 'string') return fecha;
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch(e) { return ""; }
}
function formatearFechaParaInput(fecha) {
  if (!fecha) return new Date().toISOString().split('T')[0];
  if (typeof fecha === 'string' && fecha.includes('/')) {
    var parts = fecha.split('/');
    if (parts.length === 3) return parts[2] + '-' + parts[1] + '-' + parts[0];
  }
  return new Date().toISOString().split('T')[0];
}
function formatNumberCorreo(num) {
  if (num === null || num === undefined || isNaN(num)) return "0.00";
  return Number(num).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}
function getLoginUrl() {
  return ScriptApp.getService().getUrl();
}

function obtenerSpreadsheetEmpresa_(idSpreadsheetEmpresa) {
  try {
    if (idSpreadsheetEmpresa && String(idSpreadsheetEmpresa).trim() !== '') {
      return SpreadsheetApp.openById(String(idSpreadsheetEmpresa).trim());
    }
    if (typeof _SS_ === 'function') {
      return _SS_(idSpreadsheetEmpresa);
    }
    if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function formatearUrlImagenDrive_(url) {
  if (!url || typeof url !== 'string') return '';
  var raw = url.trim();
  if (!raw) return '';

  var driveMatch = raw.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
                   raw.match(/id=([a-zA-Z0-9_-]+)/) ||
                   raw.match(/[-\w]{25,}/);
  if (driveMatch && driveMatch[1]) {
    return 'https://drive.google.com/thumbnail?id=' + driveMatch[1] + '&sz=w300';
  } else if (driveMatch && driveMatch[0] && raw.indexOf('drive.google.com') !== -1) {
    return 'https://drive.google.com/thumbnail?id=' + driveMatch[0] + '&sz=w300';
  }
  return raw;
}

// Helper seguro para formatear fechas
function formatearFechaSafe_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    try {
      return Utilities.formatDate(valor, Session.getScriptTimeZone() || 'America/Caracas', 'yyyy-MM-dd');
    } catch(e) {
      return valor.toISOString().split('T')[0];
    }
  }
  return String(valor).trim();
}

// ============================================================
// 1. CARGA DE CATÁLOGO Y PEDIDOS PARA EL POS (pedidos.html)
// ============================================================
function obtenerDatosPedidos(token, idSpreadsheetEmpresa) {
  try {
    var ss = obtenerSpreadsheetEmpresa_(idSpreadsheetEmpresa);

    var respuesta = {
      status: 'success',
      tasa: 1,
      nuevoIdPedido: 1,
      productos: [],
      clientes: [],
      categorias: []
    };

    // 1.1 Tasa de cambio actual
    try {
      var tc = (typeof obtenerTasaCambioActual === 'function') ? obtenerTasaCambioActual() : null;
      if (tc && tc.tasaUSD > 0) {
        respuesta.tasa = tc.tasaUSD;
      } else if (typeof getTasasDeCambio === 'function') {
        var tc2 = getTasasDeCambio();
        if (tc2 && tc2.tasaDolar > 0) respuesta.tasa = tc2.tasaDolar;
      }
    } catch(e) { respuesta.tasa = 1; }

    // 1.2 Catálogo de Productos
    var sheetProd = ss.getSheetByName('Productos');
    if (sheetProd) {
      var dataProd = sheetProd.getDataRange().getValues();
      if (dataProd.length > 1) {
        var hProd = dataProd[0].map(function(c) { return String(c || '').trim().toLowerCase(); });
        
        var idxId = hProd.indexOf("idproducto"); if (idxId === -1) idxId = hProd.indexOf("id"); if (idxId === -1) idxId = 0;
        var idxNom = hProd.indexOf("nombreproducto"); if (idxNom === -1) idxNom = hProd.indexOf("nombre"); if (idxNom === -1) idxNom = 1;
        var idxPre = hProd.indexOf("preciousd"); if (idxPre === -1) idxPre = hProd.indexOf("precio"); if (idxPre === -1) idxPre = 2;
        var idxStk = hProd.indexOf("existencia"); if (idxStk === -1) idxStk = hProd.indexOf("stock"); if (idxStk === -1) idxStk = 3;
        var idxCat = hProd.indexOf("categoria"); if (idxCat === -1) idxCat = hProd.indexOf("rubro");
        var idxImg = hProd.indexOf("imagenproducto"); if (idxImg === -1) idxImg = hProd.indexOf("imagen"); if (idxImg === -1) idxImg = hProd.indexOf("foto");

        var catsObj = {};
        for (var i = 1; i < dataProd.length; i++) {
          var f = dataProd[i];
          if (!f[idxId] && !f[idxNom]) continue;

          var cNom = (idxCat >= 0 && f[idxCat]) ? String(f[idxCat]).trim() : 'General';
          if (cNom && !catsObj[cNom]) { 
            catsObj[cNom] = true; 
            respuesta.categorias.push(cNom); 
          }

          var rawImg = (idxImg >= 0 && f[idxImg]) ? String(f[idxImg]).trim() : '';

          respuesta.productos.push({
            id: f[idxId],
            idProducto: f[idxId],
            nombre: f[idxNom],
            precio: parseFloat(String(f[idxPre] || '0').replace(',', '.')) || 0,
            stock: parseInt(f[idxStk], 10) || 0,
            categoria: cNom,
            imagen: formatearUrlImagenDrive_(rawImg)
          });
        }
      }
    }

    // 1.3 Lista de Clientes
    var sheetCli = ss.getSheetByName('Clientes');
    if (sheetCli) {
      var dataCli = sheetCli.getDataRange().getValues();
      if (dataCli.length > 1) {
        var hCli = dataCli[0].map(function(c) { return String(c || '').trim().toLowerCase(); });
        var idxCId = hCli.indexOf("idclientes"); if (idxCId === -1) idxCId = 0;
        var idxCNom = hCli.indexOf("nombreclientes"); if (idxCNom === -1) idxCNom = 1;
        var idxCRif = hCli.indexOf("rifclientes"); if (idxCRif === -1) idxCRif = 2;

        for (var j = 1; j < dataCli.length; j++) {
          if (!dataCli[j][idxCNom]) continue;
          respuesta.clientes.push({
            id: dataCli[j][idxCId],
            nombre: dataCli[j][idxCNom],
            rif: idxCRif >= 0 ? dataCli[j][idxCRif] : ''
          });
        }
      }
    }

    // 1.4 Próximo ID de Pedido
    var sheetPed = ss.getSheetByName('Pedidos');
    if (sheetPed) {
      var lastRow = sheetPed.getLastRow();
      respuesta.nuevoIdPedido = (lastRow > 1) ? lastRow : 1;
    }

    return respuesta;
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
// ============================================================
// 2. SINCRONIZACIÓN INTEGRAL (pedidosVendedor.html y Móvil)
// ============================================================
function syncTodo(vendedorIdentificador, idSpreadsheetEmpresa) {
  var response = {
    status: 'success',
    pedidos: [],
    productos: [],
    clientesBasicos: [],
    clientesFull: [],
    categorias: [],
    anticipos: [],
    bancos: [],
    tasaUSD: 1,
    nuevoIdPedido: 1
  };

  try {
    var ss = obtenerSpreadsheetEmpresa_(idSpreadsheetEmpresa);
    var vendIdClean = String(vendedorIdentificador || '').trim().toLowerCase();
    
    // Identificación de permisos de vendedor / administrador
    var esAdmin = (
      vendIdClean === '' || 
      vendIdClean === 'admin' || 
      vendIdClean === 'administrador' || 
      vendIdClean.indexOf('owner') !== -1 ||
      vendIdClean.indexOf('moreno') !== -1 ||
      vendIdClean.indexOf('freddy') !== -1
    );

    // 2.1 Tasa Maestra
    try {
      var tc = (typeof obtenerTasaCambioActual === 'function') ? obtenerTasaCambioActual() : null;
      if (tc && tc.tasaUSD > 0) {
        response.tasaUSD = tc.tasaUSD;
      } else if (typeof getTasasDeCambio === 'function') {
        var tc2 = getTasasDeCambio();
        if (tc2 && tc2.tasaDolar > 0) response.tasaUSD = tc2.tasaDolar;
      }
    } catch(errTc) { response.tasaUSD = 1; }

    // 2.2 Pedidos
    var sheetPed = ss.getSheetByName('Pedidos');
    if (sheetPed) {
      var dataPed = sheetPed.getDataRange().getValues();
      if (dataPed.length > 1) {
        var hPed = dataPed[0].map(function(c){ return String(c || '').trim().toLowerCase(); });
        var idxPId = hPed.indexOf("idpedido"); if (idxPId === -1) idxPId = 0;
        var idxNomCli = hPed.indexOf("nombreclientepedido");
        var idxFecha = hPed.indexOf("fechafacturaclientepedido");
        var idxTotUSD = hPed.indexOf("totalclienteusd");
        var idxStat = hPed.indexOf("statuspedido");
        var maxId = 0;

        for (var k = 1; k < dataPed.length; k++) {
          var pid = parseInt(dataPed[k][idxPId], 10);
          if (isNaN(pid)) continue;
          if (pid > maxId) maxId = pid;
          response.pedidos.push({
            idPedido: dataPed[k][idxPId],
            nombreClientePedido: idxNomCli >= 0 ? dataPed[k][idxNomCli] : '',
            fechaFacturaClientePedido: idxFecha >= 0 ? formatearFechaSafe_(dataPed[k][idxFecha]) : '',
            TotalClienteUSD: idxTotUSD >= 0 ? (parseFloat(String(dataPed[k][idxTotUSD]).replace(',', '.')) || 0) : 0,
            StatusPedido: idxStat >= 0 ? dataPed[k][idxStat] : 'Pendiente'
          });
        }
        response.nuevoIdPedido = maxId + 1;
      }
    }

    // 2.3 Productos (con hProd declarado adentro de la función)
    var sheetProd = ss.getSheetByName('Productos');
    if (sheetProd) {
      var dataProd = sheetProd.getDataRange().getValues();
      if (dataProd.length > 1) {
        var hProd = dataProd[0].map(function(c) { return String(c || '').trim().toLowerCase(); });

        var idxPId = hProd.indexOf("idproducto"); if (idxPId === -1) idxPId = hProd.indexOf("id"); if (idxPId === -1) idxPId = 0;
        var idxPNom = hProd.indexOf("nombreproducto"); if (idxPNom === -1) idxPNom = hProd.indexOf("nombre"); if (idxPNom === -1) idxPNom = 1;
        var idxPPre = hProd.indexOf("preciousd"); if (idxPPre === -1) idxPPre = hProd.indexOf("precio"); if (idxPPre === -1) idxPPre = 2;
        var idxPStk = hProd.indexOf("existencia"); if (idxPStk === -1) idxPStk = hProd.indexOf("stock"); if (idxPStk === -1) idxPStk = 3;
        var idxPCat = hProd.indexOf("categoria"); if (idxPCat === -1) idxPCat = hProd.indexOf("rubro");
        var idxPImg = hProd.indexOf("imagenproducto"); if (idxPImg === -1) idxPImg = hProd.indexOf("imagen"); if (idxPImg === -1) idxPImg = hProd.indexOf("foto");

        var setCats = {};
        for (var m = 1; m < dataProd.length; m++) {
          var fila = dataProd[m];
          var idItem = fila[idxPId];
          var nomItem = fila[idxPNom];
          if (!idItem && !nomItem) continue;

          var precioVal = parseFloat(String(fila[idxPPre] || '0').replace(',', '.')) || 0;
          var stockVal = parseInt(fila[idxPStk], 10);
          if (isNaN(stockVal)) stockVal = 0;

          var catVal = (idxPCat >= 0 && fila[idxPCat]) ? String(fila[idxPCat]).trim() : 'General';
          if (catVal && !setCats[catVal]) {
            setCats[catVal] = true;
            response.categorias.push(catVal);
          }

          var rawImg = (idxPImg >= 0 && fila[idxPImg]) ? String(fila[idxPImg]).trim() : '';

          response.productos.push({
            id: idItem,
            idProducto: idItem,
            nombre: nomItem,
            producto: nomItem,
            precio: precioVal,
            precioUSD: precioVal,
            stock: stockVal,
            existencia: stockVal,
            categoria: catVal,
            imagen: formatearUrlImagenDrive_(rawImg),
            imagenProducto: formatearUrlImagenDrive_(rawImg)
          });
        }
      }
    }

    // 2.4 Clientes (Carga completa y asignación inteligente)
    var sheetCli = ss.getSheetByName('Clientes');
    if (sheetCli) {
      var dataCli = sheetCli.getDataRange().getValues();
      if (dataCli.length > 1) {
        var hCli = dataCli[0].map(function(c){ return String(c || '').trim().toLowerCase(); });
        var idxIdCli = hCli.indexOf("idclientes"); if (idxIdCli === -1) idxIdCli = 0;
        var idxNom = hCli.indexOf("nombreclientes"); if (idxNom === -1) idxNom = 1;
        var idxRif = hCli.indexOf("rifclientes"); if (idxRif === -1) idxRif = 2;
        var idxVend = hCli.indexOf("vendedor");
        var idxTel = hCli.indexOf("telefonocliente");
        var idxStat = hCli.indexOf("status");

        for (var c = 1; c < dataCli.length; c++) {
          var filaCli = dataCli[c];
          var nomCliente = filaCli[idxNom];
          if (!nomCliente) continue;

          var vendAsignado = (idxVend >= 0 && filaCli[idxVend]) ? String(filaCli[idxVend]).trim().toLowerCase() : '';
          var statusCli = (idxStat >= 0 && filaCli[idxStat]) ? String(filaCli[idxStat]).trim().toLowerCase() : 'activo';

          // Filtro por vendedor: Administradores ven todos. Vendedores ven los suyos y los "sin asignar"
          if (!esAdmin && vendIdClean !== '') {
            var esMio = (vendAsignado === vendIdClean);
            var esSinAsignar = (vendAsignado === '' || vendAsignado === 'sin asignar');
            if (!esMio && !esSinAsignar) continue;
          }

          var itemCli = {
            id: filaCli[idxIdCli],
            idCliente: filaCli[idxIdCli],
            nombre: nomCliente,
            NombreClientes: nomCliente,
            rif: idxRif >= 0 ? filaCli[idxRif] : '',
            RifClientes: idxRif >= 0 ? filaCli[idxRif] : '',
            telefono: idxTel >= 0 ? filaCli[idxTel] : '',
            vendedor: vendAsignado,
            status: statusCli
          };

          response.clientesBasicos.push(itemCli);
          response.clientesFull.push(itemCli);
        }
      }
    }

    // Alias para compatibilidad con vistas que lean .clientes
    response.clientes = response.clientesBasicos;

    return response;
  } catch (err) {
    response.status = 'error';
    response.message = err.toString();
    return response;
  }
}

// ==========================================
// 2. GUARDAR / EDITAR PEDIDO
// ==========================================
function guardarPedido(pedidoData) {
  var response = {
    status: 'success',
    message: 'Guardado correctamente',
    idPedido: null,
    nuevoId: null,
    nombreCliente: null,
    rifCliente: null,
    fecha: null,
    items: [],
    subtotal: 0,
    totalUSD: 0,
    totalBs: 0,
    tasa: 1,
    correoEnviado: false
  };
  var lock = LockService.getScriptLock();

  try {
    if (lock.tryLock(15000)) {
      // Abre dinámicamente la empresa correcta
      var ss = _SS_(pedidoData.idSpreadsheet);
      var hojaPedidos = ss.getSheetByName("Pedidos");
      var hojaDetalles = ss.getSheetByName("DetallesPedidos");

      if (!hojaPedidos) throw new Error('No se encontró la hoja "Pedidos" en la empresa activa');
      if (!hojaDetalles) throw new Error('No se encontró la hoja "DetallesPedidos" en la empresa activa');

      var idPedido = parseInt(pedidoData.idPedido);
      var esEdicion = (idPedido && idPedido > 0);

      // GENERAR NUEVO ID EN LA EMPRESA ACTIVA
      if (!esEdicion) {
        var dataP = hojaPedidos.getDataRange().getValues();
        var max = 0;
        if (dataP.length > 1) {
          var idx = dataP[0].indexOf("idPedido");
          if (idx === -1) idx = 0;
          for (var i = 1; i < dataP.length; i++) {
            var v = parseInt(dataP[i][idx]);
            if (!isNaN(v) && v > max) max = v;
          }
        }
        idPedido = max + 1;
      }

      // TOTALES
      var subtotal = 0;
      for (var j = 0; j < pedidoData.carrito.length; j++) {
        subtotal += (parseFloat(pedidoData.carrito[j].precio) * parseInt(pedidoData.carrito[j].cantidad));
      }
      var totalUSD = subtotal;
      var tasa = parseFloat(pedidoData.tasa) || 1;
      var totalBs = totalUSD * tasa;

      // DATOS DEL CLIENTE Y DE LA EMPRESA
      var datosClienteFull = obtenerDatosClientePorNombre(pedidoData.cliente.nombre, pedidoData.idSpreadsheet);
      var emailCliente = datosClienteFull ? (datosClienteFull.correo || datosClienteFull.email || '') : '';
      var dirCliente = datosClienteFull ? (datosClienteFull.direccion || '') : '';
      var datosEmpresa = getEmpresaData(pedidoData.idSpreadsheet);

      // PRESERVAR ESTADO SI ES EDICIÓN
      var estadoFinal = "Pendiente";
      if (esEdicion) {
        var dataEstado = hojaPedidos.getDataRange().getValues();
        var idxIdEst = dataEstado[0].indexOf("idPedido");
        if (idxIdEst === -1) idxIdEst = 0;
        var idxStatusEst = dataEstado[0].indexOf("StatusPedido");
        for (var rE = 1; rE < dataEstado.length; rE++) {
          if (String(dataEstado[rE][idxIdEst]) === String(idPedido)) {
            if (idxStatusEst >= 0 && dataEstado[rE][idxStatusEst]) {
              estadoFinal = String(dataEstado[rE][idxStatusEst]).trim() || "Pendiente";
            }
            break;
          }
        }
      }

      // CONSTRUIR FILA DE PEDIDOS
      var headers = hojaPedidos.getDataRange().getValues()[0];
      var nuevaFila = [];

      var mapa = {
        "idPedido": idPedido,
        "nombreClientePedido": pedidoData.cliente.nombre,
        "rifClientePedido": pedidoData.cliente.rif || '',
        "fechaFacturaClientePedido": pedidoData.fecha,
        "TipoClientePedido": "Jurídico",
        "NroFacturaClientePedido": 'PED-' + String(idPedido).padStart(4, '0'),
        "MonedaClientePedido": 'USD',
        "TasaClientePedido": tasa,
        "TotalClientePedido": totalBs,
        "TotalClienteUSD": totalUSD,
        "SaldoClienteUSD": totalUSD,
        "TransaccionClientePedido": "Pedido",
        "IdNroPedido": 'PED-' + String(idPedido).padStart(4, '0'),
        "FechaRegistroPedido": new Date(),
        "StatusPedido": estadoFinal,
        "clienteId": pedidoData.cliente.id || '',
        "vendedor": pedidoData.vendedor || (Session.getActiveUser().getEmail().split('@')[0]),
        "emailCliente": emailCliente
      };

      for (var h = 0; h < headers.length; h++) {
        var key = String(headers[h]).trim();
        nuevaFila.push(mapa[key] !== undefined ? mapa[key] : "");
      }

      if (esEdicion) {
        var data = hojaPedidos.getDataRange().getValues();
        var idxId = data[0].indexOf("idPedido");
        if (idxId === -1) idxId = 0;

        for (var r = 1; r < data.length; r++) {
          if (String(data[r][idxId]) === String(idPedido)) {
            hojaPedidos.getRange(r + 1, 1, 1, nuevaFila.length).setValues([nuevaFila]);
            break;
          }
        }

        // Eliminar detalles previos para reescribirlos
        var dataDet = hojaDetalles.getDataRange().getValues();
        if (dataDet.length > 1) {
          var idxPedDet = dataDet[0].indexOf("ID Pedido");
          if (idxPedDet === -1) idxPedDet = 0;
          for (var d = dataDet.length - 1; d >= 1; d--) {
            if (String(dataDet[d][idxPedDet]) === String(idPedido)) {
              hojaDetalles.deleteRow(d + 1);
            }
          }
        }
      } else {
        hojaPedidos.appendRow(nuevaFila);
      }

      // INSERTAR DETALLES EN LA EMPRESA ACTIVA
      var headersDet = hojaDetalles.getDataRange().getValues()[0];
      var filasDetalle = [];
      var fechaActual = new Date();

      for (var iD = 0; iD < pedidoData.carrito.length; iD++) {
        var prod = pedidoData.carrito[iD];
        var filaD = [];
        var totalLinea = parseInt(prod.cantidad) * parseFloat(prod.precio);

        for (var col = 0; col < headersDet.length; col++) {
          var colName = String(headersDet[col]).trim();
          if (colName === "ID Pedido" || colName === "IDPedido") filaD.push(idPedido);
          else if (colName === "ID Producto" || colName === "IDProducto") filaD.push(prod.id);
          else if (colName === "Producto") filaD.push(prod.nombre);
          else if (colName === "Cantidad") filaD.push(parseInt(prod.cantidad));
          else if (colName === "Precio Unitario" || colName === "PrecioUnitario") filaD.push(parseFloat(prod.precio));
          else if (colName === "Total Producto" || colName === "TotalProducto") filaD.push(totalLinea);
          else if (colName === "Total") filaD.push(totalUSD);
          else if (colName === "Fecha" || colName === "FechaRegistro") filaD.push(fechaActual);
          else filaD.push("");
        }
        filasDetalle.push(filaD);
      }

      if (filasDetalle.length > 0) {
        hojaDetalles.getRange(hojaDetalles.getLastRow() + 1, 1, filasDetalle.length, filasDetalle[0].length).setValues(filasDetalle);
      }

      response.idPedido = idPedido;
      response.nuevoId = idPedido + 1;
      response.nombreCliente = pedidoData.cliente.nombre;
      response.rifCliente = pedidoData.cliente.rif || '';
      response.fecha = pedidoData.fecha;
      response.items = pedidoData.carrito;
      response.subtotal = subtotal;
      response.totalUSD = totalUSD;
      response.totalBs = totalBs;
      response.tasa = tasa;
      response.emailCliente = emailCliente;
      response.dirCliente = dirCliente;
      response.logo = datosEmpresa.logo || '';
      response.empresa = datosEmpresa.nombre || '';
      response.rifEmpresa = datosEmpresa.rif || '';
      response.dirEmpresa = datosEmpresa.direccion || '';
      response.telEmpresa = datosEmpresa.telefono || '';

    } else {
      response.status = 'error';
      response.message = 'Servidor ocupado. Intenta de nuevo.';
    }
  } catch (e) {
    response.status = 'error';
    response.message = e.toString();
    Logger.log("Error en guardarPedido: " + e.toString());
  } finally {
    lock.releaseLock();
  }
  return response;
}

// ==========================================
// 3. CARGAR PEDIDO PARA EDICIÓN O TICKET
// ==========================================
function cargarPedidoParaEdicion(idPedido, idSpreadsheetEmpresa) {
  try {
    if (!idPedido) return { status: 'error', message: 'ID de pedido inválido' };

    var ss = _SS_(idSpreadsheetEmpresa);
    var sheetPed = ss.getSheetByName('Pedidos');
    if (!sheetPed) return { status: 'error', message: 'Hoja Pedidos no encontrada' };

    var dataPed = sheetPed.getDataRange().getValues();
    var hPed = dataPed[0];

    var idxId = -1, idxIdCliente = -1, idxCliente = -1, idxRif = -1,
        idxTasa = -1, idxFecha = -1, idxTotal = -1, idxStatus = -1;

    for (var i = 0; i < hPed.length; i++) {
      var col = String(hPed[i]).trim().toLowerCase();
      if (col === "idpedido" || col === "id pedido" || col === "id_pedido") idxId = i;
      if (col === "idcliente" || col === "id cliente" || col === "clienteid") idxIdCliente = i;
      if (col === "nombreclientepedido" || col === "cliente" || col === "nombrecliente") idxCliente = i;
      if (col === "rifclientepedido" || col === "rif" || col === "rifcliente") idxRif = i;
      if (col === "tasapedido" || col === "tasa" || col === "tasaclientepedido") idxTasa = i;
      if (col === "fechafacturaclientepedido" || col === "fecha" || col === "fechapedido") idxFecha = i;
      if (col === "totalclientepedido" || col === "total" || col === "totalusd") idxTotal = i;
      if (col === "statuspedido" || col === "status" || col === "estado") idxStatus = i;
    }

    if (idxId === -1) return { status: 'error', message: 'No se encontró la columna idPedido' };

    var pedidoEncontrado = null;
    for (var r = 1; r < dataPed.length; r++) {
      if (String(dataPed[r][idxId]).trim() === String(idPedido).trim()) {
        var fechaFmt = idxFecha > -1 ? formatearFechaParaInput(dataPed[r][idxFecha]) : '';
        var tasa = idxTasa > -1 ? (parseFloat(dataPed[r][idxTasa]) || 1) : 1;
        var total = idxTotal > -1 ? (parseFloat(dataPed[r][idxTotal]) || 0) : 0;
        var status = idxStatus > -1 ? String(dataPed[r][idxStatus] || 'Pendiente') : 'Pendiente';
        var nombre = idxCliente > -1 ? (dataPed[r][idxCliente] || '') : '';
        var rif = idxRif > -1 ? (dataPed[r][idxRif] || '') : '';
        var idCli = idxIdCliente > -1 ? String(dataPed[r][idxIdCliente] || '') : '';

        pedidoEncontrado = {
          idPedido: dataPed[r][idxId],
          idCliente: idCli,
          clienteId: idCli,
          clienteNombre: nombre,
          nombreClientePedido: nombre,
          rif: rif,
          rifClientePedido: rif,
          tasa: tasa,
          TasaClientePedido: tasa,
          fecha: fechaFmt,
          fechaFacturaClientePedido: fechaFmt,
          TotalClienteUSD: total,
          StatusPedido: status
        };
        break;
      }
    }

    if (!pedidoEncontrado) return { status: 'error', message: 'Pedido #' + idPedido + ' no encontrado' };

    // Detalles
    var sheetDet = ss.getSheetByName('DetallesPedidos');
    var carrito = [];
    var subtotal = 0;

    if (sheetDet) {
      var dataDet = sheetDet.getDataRange().getValues();
      var hDet = dataDet[0];
      var idxIdPed = -1, idxIdProd = -1, idxProd = -1, idxCant = -1, idxPrec = -1;

      for (var j = 0; j < hDet.length; j++) {
        var colD = String(hDet[j]).trim().toLowerCase();
        if (colD === "id pedido" || colD === "idpedido" || colD === "id_pedido") idxIdPed = j;
        if (colD === "id producto" || colD === "idproducto" || colD === "id_producto") idxIdProd = j;
        if (colD === "producto" || colD === "nombre" || colD === "descripcion") idxProd = j;
        if (colD === "cantidad" || colD === "cant") idxCant = j;
        if (colD === "precio unitario" || colD === "preciounitario" || colD === "precio") idxPrec = j;
      }

      if (idxIdPed === -1) idxIdPed = 0;
      if (idxIdProd === -1) idxIdProd = 1;
      if (idxProd === -1) idxProd = 2;
      if (idxCant === -1) idxCant = 3;
      if (idxPrec === -1) idxPrec = 4;

      for (var k = 1; k < dataDet.length; k++) {
        if (String(dataDet[k][idxIdPed]).trim() === String(idPedido).trim()) {
          var cant = parseInt(dataDet[k][idxCant]) || 1;
          var prec = parseFloat(dataDet[k][idxPrec]) || 0;
          var totLin = cant * prec;
          subtotal += totLin;
          carrito.push({
            idProducto: String(dataDet[k][idxIdProd]),
            nombre: String(dataDet[k][idxProd]),
            producto: String(dataDet[k][idxProd]),
            cantidad: cant,
            precio: prec,
            total: totLin
          });
        }
      }
    }

    return {
      status: 'success',
      pedido: pedidoEncontrado,
      carrito: carrito,
      subtotal: subtotal,
      iva: subtotal * 0.16,
      total: subtotal * 1.16
    };

  } catch (e) {
    Logger.log("Error en cargarPedidoParaEdicion: " + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

// ==========================================
// 4. ELIMINAR / CAMBIAR ESTADO
// ==========================================
function eliminarPedido(idPedido, idSpreadsheetEmpresa) {
  try {
    var ss = _SS_(idSpreadsheetEmpresa);
    var sheet = ss.getSheetByName('Pedidos');
    if (!sheet) return { status: 'error', message: 'Hoja Pedidos no encontrada' };

    var data = sheet.getDataRange().getValues();
    var idx = data[0].indexOf("idPedido");
    if (idx === -1) idx = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idx]) === String(idPedido)) {
        sheet.deleteRow(i + 1);
        
        var sheetDet = ss.getSheetByName('DetallesPedidos');
        if (sheetDet) {
          var dataDet = sheetDet.getDataRange().getValues();
          var idxDet = dataDet[0].indexOf("ID Pedido");
          if (idxDet === -1) idxDet = 0;
          for (var j = dataDet.length - 1; j >= 1; j--) {
            if (String(dataDet[j][idxDet]) === String(idPedido)) sheetDet.deleteRow(j + 1);
          }
        }
        return { status: 'success', message: 'Pedido eliminado correctamente' };
      }
    }
    return { status: 'error', message: 'Pedido no encontrado' };
  } catch(e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

function cambiarEstado(idPedido, estado, idSpreadsheetEmpresa) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = _SS_(idSpreadsheetEmpresa);
    var sheet = ss.getSheetByName('Pedidos');
    if (!sheet) throw new Error("Hoja Pedidos no encontrada");
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var idxId = headers.indexOf("idpedido");
    var idxStat = headers.indexOf("statuspedido");

    if (idxId === -1 || idxStat === -1) throw new Error("Columnas requeridas no encontradas");

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxId]).trim() === String(idPedido).trim()) {
        sheet.getRange(i + 1, idxStat + 1).setValue(estado);
        return { status: 'success', message: 'Pedido #' + idPedido + ' marcado como: ' + estado };
      }
    }
    return { status: 'error', message: 'Pedido #' + idPedido + ' no encontrado' };
  } catch(e) { 
    return { status: 'error', message: e.toString() }; 
  } finally {
    lock.releaseLock();
  }
}


// ==========================================
// 6. DATOS DE LA EMPRESA (MULTI-EMPRESA)
// ==========================================
function getEmpresaData(idSpreadsheetEmpresa) {
  try {
    // Si tenemos Contexto con datos de la empresa activa
    if (typeof Contexto !== 'undefined' && Contexto.getEmpresaActivaDatos) {
      var dActiva = Contexto.getEmpresaActivaDatos();
      if (dActiva && dActiva.nombre) {
        return {
          nombre: dActiva.nombre,
          rif: dActiva.rif || '',
          direccion: dActiva.direccion || '',
          telefono: dActiva.telefono || '',
          correo: dActiva.correo || '',
          logo: dActiva.logoUrl || ''
        };
      }
    }

    var ss = _SS_(idSpreadsheetEmpresa);
    var sheet = ss.getSheetByName('Resumen') || ss.getSheetByName('Empresa');
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h){ return String(h).toLowerCase().trim(); });
      var findVal = function(colName) {
        var idx = headers.findIndex(function(h){ return h.includes(colName); });
        return idx >= 0 ? String(data[1][idx] || '') : '';
      };

      return {
        nombre: findVal('nombre') || ss.getName().replace('GPRO_DATA - ', ''),
        rif: findVal('rif') || 'J-00000000-0',
        direccion: findVal('direcc') || '',
        telefono: findVal('telefono') || '',
        correo: findVal('correo') || '',
        logo: findVal('logo') || ''
      };
    }
  } catch (e) {
    Logger.log('Error en getEmpresaData: ' + e.toString());
  }

  return {
    nombre: 'EMPRESA',
    rif: 'J-00000000-0',
    direccion: '',
    telefono: '',
    correo: '',
    logo: ''
  };
}

// ==========================================
// 7. CLIENTES Y ANTICIPOS AUXILIARES
// ==========================================
function obtenerDatosClientePorNombre(nombreCliente, idSpreadsheetEmpresa) {
  try {
    var ss = _SS_(idSpreadsheetEmpresa);
    var sheet = ss.getSheetByName("Clientes");
    if (!sheet) return null;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });
    var idxNom = headers.indexOf("NombreClientes");
    var idxRif = headers.indexOf("RifClientes");
    var idxDir = headers.indexOf("DireccionCliente");
    var idxMail = headers.indexOf("CorreoCliente");
    var idxTel = headers.indexOf("TelefonoCliente");

    var busqueda = String(nombreCliente || '').trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxNom]).trim().toLowerCase() === busqueda) {
        return {
          nombre: data[i][idxNom],
          rif: data[i][idxRif] || '',
          direccion: data[i][idxDir] || '',
          correo: data[i][idxMail] || '',
          telefono: data[i][idxTel] || ''
        };
      }
    }
    return {
      nombre: nombreCliente || "CLIENTE GENÉRICO",
      rif: "N/A",
      direccion: "Mostrador / Venta directa",
      correo: "",
      telefono: ""
    };
  } catch (e) {
    return { nombre: nombreCliente || "CLIENTE", rif: "", direccion: "", correo: "", telefono: "" };
  }
}
function _obtenerAnticiposAppInterno(ss) {
  try {
    var sh = ss.getSheetByName('CobrosPagos');
    if (!sh) return [];

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return [];

    var h = data[0].map(function(x){ return String(x).trim(); });
    var ix = function(n){ return h.indexOf(n); };

    var idxTipoMov = ix('TipoMovimiento');
    var idxCobPag = ix('cobradaPagada');
    var idxId = ix('id');
    var idxCliente = ix('ProveedorCliente');
    var idxRif = ix('Rif');
    var idxFecha = ix('fechaAbono');
    var idxAbono = ix('Abono');
    var idxMoneda = ix('Moneda');
    var idxTasa = ix('Tasa');
    var idxBanco = ix('BancoDestino');
    var idxNroCta = ix('NroCuentaDestino');
    var idxDesc = ix('Descripcion');

    var out = [];
    for (var i = data.length - 1; i >= 1 && out.length < 100; i--) {
      var mov = idxTipoMov >= 0 ? String(data[i][idxTipoMov] || '').toUpperCase() : '';
      var cob = idxCobPag >= 0 ? String(data[i][idxCobPag] || '').toLowerCase() : '';

      if (mov === 'ANTICIPO' || cob === 'anticipo') {
        var monto = idxAbono >= 0 ? (Number(data[i][idxAbono]) || 0) : 0;
        var moneda = idxMoneda >= 0 ? String(data[i][idxMoneda] || 'BS').toUpperCase() : 'BS';
        var tasa = idxTasa >= 0 ? (Number(data[i][idxTasa]) || 1) : 1;
        var montoBs = (moneda === 'BS') ? monto : monto * tasa;
        var montoUSD = (moneda === 'USD') ? monto : (monto / tasa);

        var f = idxFecha >= 0 ? data[i][idxFecha] : null;
        var fechaMs = Date.now();
        if (f instanceof Date) fechaMs = f.getTime();
        else if (typeof f === 'string' && f) {
          var parsed = Date.parse(f);
          if (!isNaN(parsed)) fechaMs = parsed;
        }

        out.push({
          id: String(idxId >= 0 ? (data[i][idxId] || '') : ''),
          cliente: String(idxCliente >= 0 ? (data[i][idxCliente] || '') : ''),
          rif: String(idxRif >= 0 ? (data[i][idxRif] || '') : ''),
          fecha: fechaMs,
          montoBs: montoBs,
          montoUSD: montoUSD,
          banco: String(idxBanco >= 0 ? (data[i][idxBanco] || '') : ''),
          referencia: String(idxNroCta >= 0 ? (data[i][idxNroCta] || '') : ''),
          descripcion: String(idxDesc >= 0 ? (data[i][idxDesc] || '') : '')
        });
      }
    }
    return out;
  } catch (e) {
    return [];
  }
}

function _getBancosYCajasInterno(ss) {
  try {
    const sheet = ss.getSheetByName('datosBancarios') || ss.getSheetByName('Bancos');
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const resultados = [];
    const totalData = sheet.getRange(2, 1, lastRow - 1, Math.min(26, sheet.getLastColumn())).getValues();
    const idx = (col) => col - 1;

    const CT_COLS = typeof CUENTA_COLS !== 'undefined' ? CUENTA_COLS : {
      ID: 1, CODIGO: 2, NUMERO: 3, ID_BANCO_REF: 4, NOM_BANCO: 5, MONEDA: 6, TIPO: 7, SALDO: 8, COD_CONTABLE: 9, NOM_CONTABLE: 10
    };
    const CJ_COLS = typeof CAJA_COLS !== 'undefined' ? CAJA_COLS : {
      ID: 11, CODIGO: 12, NOMBRE: 13, MONEDA: 14, SALDO: 15, COD_CONTABLE: 16, NOM_CONTABLE: 17
    };

    totalData.forEach((row) => {
      const codCta = String(row[idx(CT_COLS.CODIGO)] || '').trim();
      const nroCta = String(row[idx(CT_COLS.NUMERO)] || '').trim();
      const nomBanco = String(row[idx(CT_COLS.NOM_BANCO)] || '').trim();

      if (codCta !== "" || nroCta !== "" || nomBanco !== "") {
        resultados.push({
          tipo: 'Banco',
          id: String(row[idx(CT_COLS.ID)] || ''),
          codigo: codCta,
          nro: nroCta,
          nombre: nomBanco || 'Banco',
          moneda: String(row[idx(CT_COLS.MONEDA)] || 'BS').trim().toUpperCase(),
          tipoCta: String(row[idx(CT_COLS.TIPO)] || 'Corriente').trim(),
          saldo: Number(row[idx(CT_COLS.SALDO)]) || 0,
          codContable: String(row[idx(CT_COLS.COD_CONTABLE)] || ''),
          nomContable: String(row[idx(CT_COLS.NOM_CONTABLE)] || '')
        });
      }

      const codCaja = String(row[idx(CJ_COLS.CODIGO)] || '').trim();
      const nomCaja = String(row[idx(CJ_COLS.NOMBRE)] || '').trim();

      if (codCaja !== "" || nomCaja !== "") {
        resultados.push({
          tipo: 'Caja',
          id: String(row[idx(CJ_COLS.ID)] || ''),
          codigo: codCaja,
          nro: '',
          nombre: nomCaja || 'Caja Principal',
          moneda: String(row[idx(CJ_COLS.MONEDA)] || 'BS').trim().toUpperCase(),
          tipoCta: 'Efectivo',
          saldo: Number(row[idx(CJ_COLS.SALDO)]) || 0,
          codContable: String(row[idx(CJ_COLS.COD_CONTABLE)] || ''),
          nomContable: String(row[idx(CJ_COLS.NOM_CONTABLE)] || '')
        });
      }
    });

    return resultados;
  } catch (e) {
    return [];
  }
}

// ==========================================
// 8. ANTICIPOS Y DOCUMENTOS
// ==========================================
function guardarAnticipoApp(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = _SS_(p.idSpreadsheet);
    var sheetCP = ss.getSheetByName("CobrosPagos");
    var sheetMBC = ss.getSheetByName("MovBancoCaja");
    
    if (!sheetCP) throw new Error("No se encontró la hoja CobrosPagos");
    if (!sheetMBC) throw new Error("No se encontró la hoja MovBancoCaja");

    var ahora = new Date();
    var fechaStr = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaHoraStr = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    var lastRowCP = sheetCP.getLastRow();
    var nuevoIdCP = "ANT-" + ("000" + (lastRowCP)).slice(-4);

    var cta = p.cuentaDestino || {};
    var esBanco = (cta.tipo === "Banco");
    var montoBs = parseFloat(p.montoBs) || 0;
    var montoUSD = parseFloat(p.montoUSD) || 0;
    var tasa = parseFloat(p.tasa) || 1;
    var moneda = p.moneda || (montoUSD > 0 && montoBs === 0 ? "USD" : "BS");

    var filaCP = [
      p.clienteNombre || "",
      p.rif || "",
      "",
      "",
      0,
      0,
      0,
      "Cobrado",
      "Ingreso",
      nuevoIdCP,
      (moneda === "BS" ? montoBs : montoUSD),
      moneda,
      tasa,
      p.descripcion || "ANTICIPO DE CLIENTE",
      fechaStr,
      "ANTICIPO",
      cta.codigo || cta.nombre || "",
      cta.nro || "",
      cta.nombre || p.banco || "",
      "",
      "",
      ""
    ];
    sheetCP.appendRow(filaCP);

    var lastRowMBC = sheetMBC.getLastRow();
    var idMov = "MOV-" + ("0000" + lastRowMBC).slice(-5);

    var filaMBC = [
      idMov,
      ahora,
      (esBanco ? "Banco" : "Caja"),
      cta.nro || "",
      cta.nombre || p.banco || "",
      cta.codigo || "",
      cta.tipoCta || (esBanco ? "Corriente" : "Caja Principal"),
      p.referencia || nuevoIdCP,
      (p.descripcion ? p.descripcion + " - " : "") + "Anticipo: " + p.clienteNombre,
      "Transacción Positiva",
      montoUSD,
      moneda,
      tasa,
      0,
      montoBs,
      cta.codContable || "",
      cta.nomContable || "Anticipos Recibidos",
      "",
      cta.nombre || "",
      fechaHoraStr
    ];
    sheetMBC.appendRow(filaMBC);

    return {
      status: 'success',
      id: nuevoIdCP,
      idMov: idMov,
      message: 'Anticipo ' + nuevoIdCP + ' guardado con éxito'
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
function obtenerCategoriasResumen(idSpreadsheetEmpresa) {
  try {
    var ss = _SS_(idSpreadsheetEmpresa);
    var sh = ss.getSheetByName('Resumen');
    if (!sh) return {};
    
    var lastRow = sh.getLastRow();
    if (lastRow < 44) return {};
    
    // Rango dinámico desde I44 hasta J(última fila)
    var numFilas = lastRow - 44 + 1;
    var data = sh.getRange(44, 9, numFilas, 2).getValues();
    var mapa = {};

    for (var i = 0; i < data.length; i++) {
      var idCat = String(data[i][0] || '').trim();
      var nomCat = String(data[i][1] || '').trim();
      if (idCat && nomCat) {
        mapa[idCat] = nomCat;
      }
    }
    return mapa;
  } catch(e) {
    Logger.log("Error en obtenerCategoriasResumen: " + e.toString());
    return {};
  }
}
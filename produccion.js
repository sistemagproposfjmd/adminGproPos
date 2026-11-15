// =========================================================================
// MÓDULO BACKEND: Produccion.gs (MULTI-EMPRESA COMPLETO)
// =========================================================================

const HOJA_PRODUCTOS = "Productos";
const HOJA_PRODUCCIONES = "Producciones";
const HOJA_RECETAS_HISTORIAL = "RecetasHistorial";
const MAX_PROFUNDIDAD_RECETA = 5;

function _SS_Produccion_(idOpcional) {
  if (idOpcional && typeof idOpcional === 'string' && idOpcional.trim().length > 15) {
    try {
      return SpreadsheetApp.openById(idOpcional.trim());
    } catch(e) {}
  }
  if (typeof Contexto !== 'undefined' && Contexto.getEmpresaSS) {
    return Contexto.getEmpresaSS(idOpcional);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetByNameProduccion(nombreHoja, idSpreadsheetEmpresa) {
  var ss = _SS_Produccion_(idSpreadsheetEmpresa);
  if (!ss) return null;
  return ss.getSheetByName(nombreHoja);
}

/**
 * Inicializa y valida las hojas de producción en la empresa activa
 */
function initModuloProduccion(idSpreadsheetEmpresa) {
  try {
    validarYCrearHojasProduccion(idSpreadsheetEmpresa);
    var usuario = Session.getActiveUser().getEmail() || "Operador";
    return {
      success: true,
      detalles: {
        usuarioActual: usuario,
        empresaId: idSpreadsheetEmpresa || ''
      }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Valida la existencia de las hojas en la hoja de cálculo de la empresa activa
 */
function validarYCrearHojasProduccion(idSpreadsheetEmpresa) {
  var ss = _SS_Produccion_(idSpreadsheetEmpresa);
  if (!ss) return { success: false, message: "No se pudo acceder a la hoja de la empresa" };

  // 1. Validar Producciones
  var shProd = ss.getSheetByName(HOJA_PRODUCCIONES);
  if (!shProd) {
    shProd = ss.insertSheet(HOJA_PRODUCCIONES);
    shProd.appendRow([
      "idProduccion", "fecha", "idCompuesto", "nombreCompuesto",
      "cantidadProducida", "costoTotalUSD", "costoUnitarioUSD",
      "usuario", "estado", "desgloseJSON", "notas"
    ]);
    shProd.setFrozenRows(1);
  }

  // 2. Validar RecetasHistorial
  var shHist = ss.getSheetByName(HOJA_RECETAS_HISTORIAL);
  if (!shHist) {
    shHist = ss.insertSheet(HOJA_RECETAS_HISTORIAL);
    shHist.appendRow([
      "idCompuesto", "version", "fechaModificacion", "usuario", "recetaJSON", "motivo"
    ]);
    shHist.setFrozenRows(1);
  }

  return { success: true, message: "Hojas verificadas correctamente" };
}

/**
 * Retorna todos los productos de la empresa activa
 */
function obtenerProductos(idSpreadsheetEmpresa) {
  var ss = _SS_Produccion_(idSpreadsheetEmpresa);
  if (!ss) return [];
  var sh = ss.getSheetByName(HOJA_PRODUCTOS);
  if (!sh) return [];

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idxId     = headers.indexOf("IdProducto");
  var idxNombre = headers.indexOf("NombreProducto");
  var idxPrecio = headers.indexOf("PrecioProducto");
  var idxVenta  = headers.indexOf("precioVenta");
  var idxStock  = headers.indexOf("StockProducto");
  var idxMinimo = headers.indexOf("MinimoProducto");
  var idxCat    = headers.indexOf("CategoriaProducto");
  var idxTipo   = headers.indexOf("tipo");
  var idxStatus = headers.indexOf("Status");
  var idxReceta = headers.indexOf("receta");
  var idxMedida = headers.indexOf("Medida");
  var idxFactor = headers.indexOf("FactorConversion");

  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var id = String(r[idxId] || "").trim();
    if (!id) continue;

    lista.push({
      IdProducto: id,
      NombreProducto: String(r[idxNombre] || ""),
      PrecioProducto: parseFloat(r[idxPrecio]) || 0,
      precioVenta: idxVenta !== -1 ? (parseFloat(r[idxVenta]) || 0) : 0,
      StockProducto: parseFloat(r[idxStock]) || 0,
      MinimoProducto: parseFloat(r[idxMinimo]) || 0,
      CategoriaProducto: idxCat !== -1 ? String(r[idxCat] || "") : "",
      tipo: idxTipo !== -1 ? String(r[idxTipo] || "Simple") : "Simple",
      Status: idxStatus !== -1 ? String(r[idxStatus] || "Activo") : "Activo",
      receta: idxReceta !== -1 ? String(r[idxReceta] || "") : "",
      Medida: idxMedida !== -1 ? String(r[idxMedida] || "UNID") : "UNID",
      FactorConversion: idxFactor !== -1 ? (parseFloat(r[idxFactor]) || 1) : 1
    });
  }
  return lista;
}

/**
 * Retorna las categorías disponibles
 */
function obtenerCategoriasConId(idSpreadsheetEmpresa) {
  var ss = _SS_Produccion_(idSpreadsheetEmpresa);
  if (!ss) return [];
  var sh = ss.getSheetByName("Categoria") || ss.getSheetByName("Categorias") || ss.getSheetByName("Resumen");
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var categorias = [];
  for (var i = 1; i < data.length; i++) {
    var val = String(data[i][0] || "").trim();
    if (val && !categorias.some(function(c) { return c.id === val; })) {
      categorias.push({ id: val, nombre: val });
    }
  }
  return categorias;
}

/**
 * Calcula la disponibilidad de producción de un compuesto basado en el stock de sus ingredientes
 */
function calcularDisponibilidadCompuesto(idProducto, idSpreadsheetEmpresa) {
  try {
    var productos = obtenerProductos(idSpreadsheetEmpresa);
    var compuesto = productos.find(function(p) { return p.IdProducto === idProducto; });
    if (!compuesto) return { success: false, message: "Compuesto no encontrado" };

    if (!compuesto.receta) {
      return { success: false, message: "El producto no tiene una receta definida" };
    }

    var recetaArr = [];
    try {
      recetaArr = JSON.parse(compuesto.receta);
    } catch(e) {
      return { success: false, message: "Estructura de receta inválida en la BD" };
    }

    if (!Array.isArray(recetaArr) || recetaArr.length === 0) {
      return { success: false, message: "La receta no contiene ingredientes" };
    }

    var maxProducible = Infinity;
    var desglose = [];
    var costoCalculado = 0;

    for (var i = 0; i < recetaArr.length; i++) {
      var ing = recetaArr[i];
      var ingProd = productos.find(function(p) { return p.IdProducto === ing.id; });
      var stockActual = ingProd ? (ingProd.StockProducto || 0) : 0;
      var costoIng = ingProd ? (ingProd.PrecioProducto || 0) : 0;
      var req = parseFloat(ing.cantidad) || 0;

      var maxParaEste = req > 0 ? Math.floor(stockActual / req) : 0;
      if (maxParaEste < maxProducible) maxProducible = maxParaEste;

      costoCalculado += (costoIng * req);

      desglose.push({
        id: ing.id,
        nombre: ingProd ? ingProd.NombreProducto : (ing.nombre || ing.id),
        requeridoPorUnidad: req,
        unidad: ing.unidad || 'UNID',
        stockDisponible: stockActual,
        maxParaEste: maxParaEste,
        costoUnitario: costoIng
      });
    }

    if (maxProducible === Infinity) maxProducible = 0;

    return {
      success: true,
      maxProducible: Math.max(0, maxProducible),
      costoUnitarioEstimado: costoCalculado,
      desglose: desglose
    };

  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Ejecuta la producción descontando inventario y sumando stock al compuesto
 */
function producirCompuesto(idProducto, cantidad, opciones) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch(e) {
    return { success: false, message: "El sistema está ocupado. Intenta de nuevo en unos segundos." };
  }

  try {
    opciones = opciones || {};
    var modo = opciones.modo || 'normal';
    var idSpreadsheet = opciones.idSpreadsheet || '';

    var ss = _SS_Produccion_(idSpreadsheet);
    if (!ss) return { success: false, message: "No se pudo abrir la empresa activa" };

    var shProd = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!shProd) return { success: false, message: "Hoja de Productos no encontrada" };

    var disp = calcularDisponibilidadCompuesto(idProducto, idSpreadsheet);
    if (!disp.success) return disp;

    var cant = parseFloat(cantidad) || 0;
    if (cant <= 0) return { success: false, message: "Cantidad inválida" };

    var costoTotal = disp.costoUnitarioEstimado * cant;

    if (modo === 'simulacion') {
      return {
        success: true,
        costoTotal: costoTotal,
        costoUnitario: disp.costoUnitarioEstimado,
        message: "Simulación completada satisfactoriamente"
      };
    }

    if (cant > disp.maxProducible) {
      return { success: false, message: "Stock insuficiente de ingredientes para producir " + cant + " unidades" };
    }

    var data = shProd.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var idxId    = headers.indexOf("IdProducto");
    var idxStock = headers.indexOf("StockProducto");
    var idxCosto = headers.indexOf("PrecioProducto");

    // 1. Descontar ingredientes
    for (var d = 0; d < disp.desglose.length; d++) {
      var item = disp.desglose[d];
      var totalDeducir = item.requeridoPorUnidad * cant;
      for (var f = 1; f < data.length; f++) {
        if (String(data[f][idxId]).trim() === item.id) {
          var stockActual = parseFloat(data[f][idxStock]) || 0;
          var nuevoStock = Math.max(0, stockActual - totalDeducir);
          shProd.getRange(f + 1, idxStock + 1).setValue(nuevoStock);
          break;
        }
      }
    }

    // 2. Aumentar stock del producto terminado y actualizar costo
    var nombreCompuesto = idProducto;
    for (var p = 1; p < data.length; p++) {
      if (String(data[p][idxId]).trim() === idProducto) {
        var stockComp = parseFloat(data[p][idxStock]) || 0;
        shProd.getRange(p + 1, idxStock + 1).setValue(stockComp + cant);
        if (disp.costoUnitarioEstimado > 0) {
          shProd.getRange(p + 1, idxCosto + 1).setValue(disp.costoUnitarioEstimado);
        }
        nombreCompuesto = data[p][headers.indexOf("NombreProducto")] || idProducto;
        break;
      }
    }

    // 3. Registrar en Producciones
    validarYCrearHojasProduccion(idSpreadsheet);
    var shHist = ss.getSheetByName(HOJA_PRODUCCIONES);
    var idProduccion = "PROD-" + Date.now();
    var usuario = Session.getActiveUser().getEmail() || "Operador";

    shHist.appendRow([
      idProduccion,
      new Date().toISOString(),
      idProducto,
      nombreCompuesto,
      cant,
      costoTotal,
      disp.costoUnitarioEstimado,
      usuario,
      "Completada",
      JSON.stringify(disp.desglose),
      ""
    ]);

    return {
      success: true,
      message: "Se produjeron " + cant + " unidades de " + nombreCompuesto + " con éxito.",
      idProduccion: idProduccion
    };

  } catch(e) {
    return { success: false, message: "Error en producción: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Consulta el historial de producciones con soporte de paginación y filtros
 */
function obtenerHistorialProducciones(filtros) {
  try {
    filtros = filtros || {};
    var idSpreadsheet = filtros.idSpreadsheet || '';
    var ss = _SS_Produccion_(idSpreadsheet);
    if (!ss) return { success: false, message: "Empresa no disponible" };

    var sh = ss.getSheetByName(HOJA_PRODUCCIONES);
    if (!sh) return { success: true, producciones: [], total: 0 };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, producciones: [], total: 0 };

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var producciones = [];

    for (var i = data.length - 1; i >= 1; i--) {
      var r = data[i];
      if (!r[0]) continue;
      producciones.push({
        idProduccion: String(r[0]),
        fecha: r[1],
        idCompuesto: String(r[2]),
        nombreCompuesto: String(r[3]),
        cantidadProducida: parseFloat(r[4]) || 0,
        costoTotalUSD: parseFloat(r[5]) || 0,
        costoUnitarioUSD: parseFloat(r[6]) || 0,
        usuario: String(r[7] || ''),
        estado: String(r[8] || 'Completada'),
        desgloseJSON: String(r[9] || '[]')
      });
    }

    var pagina = parseInt(filtros.pagina) || 1;
    var porPagina = parseInt(filtros.porPagina) || 20;
    var inicio = (pagina - 1) * porPagina;
    var paginados = producciones.slice(inicio, inicio + porPagina);

    return {
      success: true,
      producciones: paginados,
      total: producciones.length,
      pagina: pagina,
      totalPaginas: Math.ceil(producciones.length / porPagina)
    };

  } catch(e) {
    return { success: false, message: e.message, producciones: [] };
  }
}

/**
 * Guarda o actualiza la receta de un producto compuesto y almacena una versión en el historial
 */
function guardarReceta(idProducto, recetaArr, idSpreadsheetEmpresa) {
  try {
    var ss = _SS_Produccion_(idSpreadsheetEmpresa);
    if (!ss) return { success: false, message: "No se pudo acceder a la empresa activa" };

    var sh = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!sh) return { success: false, message: "Hoja Productos no encontrada" };

    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var idxId = headers.indexOf("IdProducto");
    var idxReceta = headers.indexOf("receta");
    var idxTipo = headers.indexOf("tipo");

    if (idxReceta === -1) {
      sh.getRange(1, headers.length + 1).setValue("receta");
      idxReceta = headers.length;
    }

    var filaEncontrada = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxId]).trim() === idProducto) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) return { success: false, message: "Producto no encontrado" };

    var jsonStr = JSON.stringify(recetaArr);
    sh.getRange(filaEncontrada, idxReceta + 1).setValue(jsonStr);
    if (idxTipo !== -1) {
      sh.getRange(filaEncontrada, idxTipo + 1).setValue("Compuesto");
    }

    // Guardar versión en historial
    validarYCrearHojasProduccion(idSpreadsheetEmpresa);
    var shHist = ss.getSheetByName(HOJA_RECETAS_HISTORIAL);
    if (shHist) {
      var usuario = Session.getActiveUser().getEmail() || "Operador";
      shHist.appendRow([
        idProducto,
        Date.now(),
        new Date().toISOString(),
        usuario,
        jsonStr,
        "Actualización manual de receta"
      ]);
    }

    return { success: true, message: "Receta guardada exitosamente" };

  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Recalcula el costo unitario de todos los compuestos de la empresa
 */
function recalcularCostosCompuestos(idSpreadsheetEmpresa) {
  try {
    var productos = obtenerProductos(idSpreadsheetEmpresa);
    var compuestos = productos.filter(function(p) { return p.tipo === "Compuesto"; });
    var ss = _SS_Produccion_(idSpreadsheetEmpresa);
    var sh = ss.getSheetByName(HOJA_PRODUCTOS);
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var idxId = headers.indexOf("IdProducto");
    var idxCosto = headers.indexOf("PrecioProducto");

    var actualizados = 0;
    compuestos.forEach(function(c) {
      var disp = calcularDisponibilidadCompuesto(c.IdProducto, idSpreadsheetEmpresa);
      if (disp.success && disp.costoUnitarioEstimado > 0) {
        for (var f = 1; f < data.length; f++) {
          if (String(data[f][idxId]).trim() === c.IdProducto) {
            sh.getRange(f + 1, idxCosto + 1).setValue(disp.costoUnitarioEstimado);
            actualizados++;
            break;
          }
        }
      }
    });

    return { success: true, message: "Se recalcularon " + actualizados + " compuestos." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function obtenerRecetaProducto(codigoProducto, idSpreadsheet) {
  const ss = Contexto.getEmpresaSS(idSpreadsheet);
  const sheetRecetas = ss.getSheetByName("RecetasHistorial");
  if (!sheetRecetas) throw new Error("Hoja RecetasHistorial no encontrada.");

  const data = sheetRecetas.getDataRange().getValues();
  const insumos = [];
  let cabecera = null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === codigoProducto.trim()) {
      cabecera = {
        idReceta: data[i][0],
        codigoProducto: data[i][1],
        nombreProducto: data[i][2],
        rendimiento: Number(data[i][3]) || 1,
        unidad: data[i][4]
      };
      insumos.push({
        codigoInsumo: data[i][5],
        descripcionInsumo: data[i][6],
        cantidad: Number(data[i][7]) || 0,
        unidadMedida: data[i][8],
        costoUnitarioUSD: Number(data[i][9]) || 0
      });
    }
  }

  return { cabecera: cabecera, insumos: insumos };
}

function registrarOrdenProduccion(orden, idSpreadsheet, auth) {
  const ss = Contexto.getEmpresaSS(idSpreadsheet);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheetProd = ss.getSheetByName("Producciones");
    const sheetStock = ss.getSheetByName("Productos");
    
    // 1. Validar existencias de todos los insumos requeridos
    const dataStock = sheetStock.getDataRange().getValues();
    const stockMap = {};
    for (let i = 1; i < dataStock.length; i++) {
      stockMap[String(dataStock[i][0]).trim()] = { fila: i + 1, actual: Number(dataStock[i][4]) || 0 };
    }

    // 2. Descontar materias primas e incrementar producto terminado
    orden.insumosRequeridos.forEach(function(ins) {
      const item = stockMap[ins.codigoInsumo];
      if (item) {
        const nuevoStock = item.actual - ins.cantidadTotal;
        sheetStock.getRange(item.fila, 5).setValue(nuevoStock);
      }
    });

    // 3. Aumentar stock del producto terminado
    const terminado = stockMap[orden.codigoProductoTerminado];
    if (terminado) {
      sheetStock.getRange(terminado.fila, 5).setValue(terminado.actual + orden.cantidadProducida);
    }

    // 4. Asentar en historial de Producción
    sheetProd.appendRow([
      orden.idOrden || Utilities.getUuid(),
      new Date(),
      orden.codigoProductoTerminado,
      orden.nombreProductoTerminado,
      orden.cantidadProducida,
      orden.lote,
      orden.costoTotalUSD,
      "Finalizada",
      orden.responsable
    ]);

    return { success: true, message: "Producción procesada y stock actualizado." };
  } finally {
    lock.releaseLock();
  }
}
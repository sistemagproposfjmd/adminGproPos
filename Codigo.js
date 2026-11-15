
// =============================================
// MÓDULO LOGIN: codigo.gs
// VARIABLES GLOBALES Y CONFIGURACIÓN
// =============================================
const NOM_EMPRESA = "Tu Nombre Comercial, C.A.";
const RIF_EMPRESA = "J-12345678-9";
const DIR_EMPRESA = "Av. Principal, Edificio Centro, Piso 1, Oficina A, Caracas";
const LOGO_EMPRESA = "https://lh3.googleusercontent.com/d/1dCncG3S9eZYyFMl5LnkvWFKjV32kgyGX=w120-h120"; 

const SHEET_NAME_CLIENTES = "Clientes"; 
const SHEET_NAME_COBROS = "CobrosPagos";
const RET_SHEET_NAME = "Abonos";
const RET_ID_COL = 1;
const RET_NRO_FACTURA_COL = 3;
const RET_MONTO_USD_COL = 7;
const RET_MONTO_BS_COL = 8;
const HOJA_FACTURAS_NOMBRE = "Ventas";
const HOJA_CLIENTES_NOMBRE = "Clientes";
const HOJA_DETALLES_NOMBRE = "DetalleFacturas";

const correoADMIN = "fmorenocontabilidad@gmail.com";
const CLAVE_SECRETA = "CARACAS";

const URL_WEB_APP = "https://script.google.com/macros/s/AKfycbxrxRa14Svz-6MhEw3aCoRqDnBUoq7Kqz-9MPAaJUozYG9do4tZN93a0oDP2gERdhuYMg/exec";
const URL_WEB_APPBlog = "https://sistemagpropos.blogspot.com/";


function borrarDatos() {
  var hojaNombres = ["Compras", "Ventas", "Pedidos", "Abonos", "DetalleFacturas", "CobrosPagos", "DetallesPedidos","Asientos","MovBancoCaja"];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  hojaNombres.forEach(function(nombreHoja) {
    var hoja = ss.getSheetByName(nombreHoja);
    if (hoja) {
      // Obtener el rango desde la fila 2 hasta la última fila y columna
      var ultimaFila = hoja.getLastRow();
      var ultimaColumna = hoja.getLastColumn();
      if (ultimaFila > 1) {
        hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).clearContent(); // Borrar contenido
      }
    } else {
      Logger.log("La hoja '" + nombreHoja + "' no existe.");
    }
  });
}

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Funciones')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Datos Analisis')
          .addItem('Abrir Web', 'abrirWebAppEnNuevaPestana')
          .addItem('Abrir Maestro', 'abrirHojaAnalisis')) // <- Apunta a la nueva función
      .addSeparator()
      .addItem('Borrar Datos', 'borrarDatos')
      .addToUi();
}

function abrirWebAppEnNuevaPestana() {
  const urlWebApp = URL_WEB_APP;
  const urlWebAppBlog = URL_WEB_APPBlog;
  
  const htmlScript = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_blank">
        <style>
          * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
          }
          body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            background: radial-gradient(circle at top left, #1e1b4b, #0f172a, #020617);
            color: #f8fafc;
            padding: 16px;
          }
          .glass-card {
            background: rgba(30, 41, 59, 0.45);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.125);
            border-radius: 16px;
            padding: 20px 24px;
            width: 100%;
            text-align: center;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
            display: flex;
            flex-direction: column;
            gap: 12px;
          }
          p {
            font-size: 13px;
            color: #94a3b8;
            margin-bottom: 4px;
            line-height: 1.4;
          }
          .btn-open {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            width: 100%;
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.8), rgba(99, 102, 241, 0.8));
            border: 1px solid rgba(255, 255, 255, 0.2);
            color: #ffffff !important;
            padding: 10px 18px;
            text-decoration: none;
            border-radius: 10px;
            font-weight: 600;
            font-size: 13.5px;
            letter-spacing: 0.3px;
            box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            transition: all 0.25s ease;
            cursor: pointer;
          }
          .btn-open.secondary {
            background: linear-gradient(135deg, rgba(15, 23, 42, 0.8), rgba(30, 41, 59, 0.8));
            border: 1px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
          }
          .btn-open:hover {
            opacity: 0.95;
            box-shadow: 0 6px 20px rgba(59, 130, 246, 0.5);
            transform: translateY(-1px);
          }
        </style>
      </head>
      <body>
        <div class="glass-card">
          <p>Selecciona el sistema que deseas abrir:</p>
          
          <!-- Botón para WebApp Blog -->
          <a href="${urlWebAppBlog}" target="_blank" class="btn-open" onclick="cerrarModal()">🚀 Abrir WebApp Blog</a>

          <!-- Botón para WebApp Principal (Opcional) -->
          <a href="${urlWebApp}" target="_blank" class="btn-open secondary" onclick="cerrarModal()">🔗 Abrir WebApp Principal</a>
        </div>

        <script>
          function cerrarModal() {
            // Espera un instante para que la pestaña alcance a abrirse antes de cerrar el diálogo
            setTimeout(function() {
              google.script.host.close();
            }, 300);
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlScript)
    .setWidth(380)
    .setHeight(220);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Acceso al Sistema');
}
function abrirHojaAnalisis() {
  const idArchivo = '1fFyzPgSjBnmnHmJrKAV300B5aShq2ah9BnhlrjeXr2k';
  const urlHoja = `https://docs.google.com/spreadsheets/d/${idArchivo}/edit`;
  
  const htmlScript = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_blank">
        <style>
          * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
          }
          body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            background: radial-gradient(circle at top left, #1e1b4b, #0f172a, #020617);
            color: #f8fafc;
            padding: 16px;
          }
          .glass-card {
            background: rgba(30, 41, 59, 0.45);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.125);
            border-radius: 16px;
            padding: 20px 24px;
            width: 100%;
            text-align: center;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
            display: flex;
            flex-direction: column;
            gap: 12px;
          }
          p {
            font-size: 13px;
            color: #94a3b8;
            margin-bottom: 4px;
            line-height: 1.4;
          }
          .btn-open {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            width: 100%;
            background: linear-gradient(135deg, rgba(16, 185, 129, 0.8), rgba(5, 150, 105, 0.8));
            border: 1px solid rgba(255, 255, 255, 0.2);
            color: #ffffff !important;
            padding: 10px 18px;
            text-decoration: none;
            border-radius: 10px;
            font-weight: 600;
            font-size: 13.5px;
            letter-spacing: 0.3px;
            box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
            transition: all 0.25s ease;
            cursor: pointer;
          }
          .btn-open:hover {
            opacity: 0.95;
            box-shadow: 0 6px 20px rgba(16, 185, 129, 0.5);
            transform: translateY(-1px);
          }
        </style>
      </head>
      <body>
        <div class="glass-card">
          <p>Haz clic para abrir el archivo de análisis en una nueva pestaña:</p>
          
          <a href="${urlHoja}" target="_blank" class="btn-open" onclick="cerrarModal()">
            📊 Abrir Google Sheet
          </a>
        </div>

        <script>
          function cerrarModal() {
            setTimeout(function() {
              google.script.host.close();
            }, 300);
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlScript)
    .setWidth(380)
    .setHeight(180);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Acceso a Hoja de Análisis');
}
// =======================================================
/** FUNCIONES TASAS                         */
// =======================================================
function importBCVEuroDolares() {
  const url = "https://www.bcv.org.ve/";
  const options = { 
    'muteHttpExceptions': true, 
    'headers': { 
      'User-Agent': 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)' 
    },
    'timeout': 10000
  };
  
  const HojaTasa = getSheetByName("TASAS");
  if (!HojaTasa) { 
    console.error("CRÍTICO: No se encontró la hoja 'Resumen'."); 
    return { status: "error", message: "Hoja Resumen no encontrada" };
  }
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Error HTTP ${response.getResponseCode()}`);
    }
    
    const content = response.getContentText();
    console.log(content.substring(0, 500));
    
    const dolarRegex = /<div id="dolar"[\s\S]*?<strong>\s*([\d.,]+)\s*<\/strong>/i;
    const euroRegex = /<div id="euro"[\s\S]*?<strong>\s*([\d.,]+)\s*<\/strong>/i;
    
    const dolarMatch = content.match(dolarRegex);
    const euroMatch = content.match(euroRegex);
    
    if (dolarMatch && dolarMatch[1] && euroMatch && euroMatch[1]) {
      const dollarRateStr = dolarMatch[1].trim().replace(/\./g, '').replace(',', '.');
      const euroRateStr = euroMatch[1].trim().replace(/\./g, '').replace(',', '.');
      
      const dollarRate = parseFloat(dollarRateStr);
      const euroRate = parseFloat(euroRateStr);
      
      HojaTasa.getRange("C3:c4").clearContent();
      HojaTasa.getRange("c3").setValue(dollarRate); 
      HojaTasa.getRange("c4").setValue(euroRate);
      
      console.log(`ÉXITO. Tasas actualizadas: Dólar = ${dollarRate}, Euro = ${euroRate}`);
      return { status: "success", dollar: dollarRate, euro: euroRate };
    } else {
      const errorMsg = "Error: No se encontró el patrón esperado en la web del BCV.";
      console.error(errorMsg);
      HojaTasa.getRange("C3:c4").setValues([[errorMsg],[errorMsg]]);
      return { status: "error", message: errorMsg };
    }
  } catch (e) {
    const errorMsg = "Error de conexión al procesar la página: " + e.toString();
    console.error(errorMsg);
    HojaTasa.getRange("F14:F15").setValues([["Error de conexión"], ["Error de conexión"]]);
    return { status: "error", message: errorMsg };
  }
}
function obtenerTasaCambioActual() {
  try {
    console.log('🔍 Ejecutando obtenerTasaCambioActual() desde archivo externo');
    
    var idArchivoExterno = "1fFyzPgSjBnmnHmJrKAV300B5aShq2ah9BnhlrjeXr2k";
    var archivoExterno = SpreadsheetApp.openById(idArchivoExterno);
    var hojaTasas = archivoExterno.getSheetByName("Tasas");
    
    if (!hojaTasas) {
      console.warn('⚠️ No se encontró la hoja "Tasas" en el archivo externo.');
      return { tasaUSD: 0, tasaEUR: 0 };
    }
    
    var valorUSD = hojaTasas.getRange("C3").getValue();
    var valorEUR = hojaTasas.getRange("C4").getValue();
    
    // Limpieza segura para números de Google Sheets
    const limpiarTasa = (val) => {
      let strVal = String(val).trim();
      if (strVal.includes(',')) {
        strVal = strVal.replace(/\./g, '').replace(',', '.');
      }
      return parseFloat(strVal) || 0;
    };
    
    var tasaUSD = limpiarTasa(valorUSD);
    var tasaEUR = limpiarTasa(valorEUR);
    
    console.log('📊 Valores obtenidos de Tasas:', { tasaUSD, tasaEUR });
    
    return { tasaUSD, tasaEUR };
    
  } catch (error) {
    console.error('❌ Error en obtenerTasaCambioActual:', error);
    return { tasaUSD: 0, tasaEUR: 0 };
  }
}
function getTasasDeCambio() {
  try {
    // La hoja Tasas está en GPRO_MAESTRO
    const HOJA_TASAS = (typeof getSheetByName === 'function')
      ? getSheetByName('Tasas')
      : (typeof Contexto !== 'undefined' ? Contexto.getMasterSS().getSheetByName('Tasas') : null);

    if (!HOJA_TASAS) throw new Error("La hoja 'Tasas' no fue encontrada.");
    
    const limpiarTasa = (val) => {
      let strVal = String(val).trim();
      if (strVal.includes(',')) {
        strVal = strVal.replace(/\./g, '').replace(',', '.');
      }
      return parseFloat(strVal) || 0;
    };

    const tasaDolar = limpiarTasa(HOJA_TASAS.getRange("C3").getValue());
    const tasaEuro  = limpiarTasa(HOJA_TASAS.getRange("C4").getValue());

    if (isNaN(tasaDolar) || isNaN(tasaEuro)) {
      throw new Error("Los valores en C3 o C4 de 'Tasas' no son números válidos.");
    }

    return { tasaDolar, tasaEuro };
  } catch (e) {
    Logger.log("Error en getTasasDeCambio: " + e.toString());
    throw new Error("No se pudieron obtener las tasas de cambio: " + e.message);
  }
}
function getDataMoneda() {
  try {
    var sheet = (typeof getSheetByName === 'function')
      ? getSheetByName("Tasas")
      : (typeof Contexto !== 'undefined' ? Contexto.getMasterSS().getSheetByName("Tasas") : null);

    if (!sheet) {
      return [{ nombre: "Dolares", tasa: 1 }, { nombre: "Bolivares", tasa: 1 }];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [{ nombre: "Dolares", tasa: 1 }, { nombre: "Bolivares", tasa: 1 }];
    }

    var values = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
    var monedas = [];

    for (var i = 0; i < values.length; i++) {
      var nombre = String(values[i][0] || "").trim();
      if (!nombre) continue;

      var rawTasa = values[i][1];
      var tasaNum = 1;

      if (typeof rawTasa === 'number') {
        tasaNum = isNaN(rawTasa) || rawTasa <= 0 ? 1 : rawTasa;
      } else {
        var strTasa = String(rawTasa || "").trim();
        if (strTasa.includes(',') && !strTasa.includes('.')) {
          strTasa = strTasa.replace(',', '.');
        } else if (strTasa.includes(',') && strTasa.includes('.')) {
          strTasa = strTasa.replace(/\./g, '').replace(',', '.');
        }
        strTasa = strTasa.replace(/[^0-9.-]/g, '');
        tasaNum = parseFloat(strTasa);
        if (isNaN(tasaNum) || tasaNum <= 0) tasaNum = 1;
      }

      var lower = nombre.toLowerCase();
      if (lower.includes('bolivar') || lower.includes('bolívar') || lower === 'bs' || lower === 'ves') {
        tasaNum = 1;
      }

      monedas.push({ nombre: nombre, tasa: tasaNum });
    }

    return monedas.length > 0 ? monedas : [{ nombre: "Dolares", tasa: 1 }, { nombre: "Bolivares", tasa: 1 }];
  } catch (e) {
    Logger.log("Error en getDataMoneda: " + e.toString());
    return [{ nombre: "Dolares", tasa: 1 }, { nombre: "Bolivares", tasa: 1 }];
  }
}
function getTasaByMoneda(monedaNombre) {
  try {
    const monedas = getDataMoneda();
    const moneda = monedas.find(m => m.nombre === monedaNombre);
    return moneda ? moneda.tasa : 1;
  } catch (e) {
    Logger.log("Error en getTasaByMoneda: " + e.toString());
    return 1;
  }
}
function obtenerTasaApiCotizave() {
  const API_KEY = "ctz_live_6FaMwGbo8Fy6UOWMDGR3Z4yo4neTFJcDiUDZqv";
  const URL_API = "https://api.cotizave.com/v1/fx/rates";
  
  // 1. Revisar si la tasa está en caché temporal (15 minutos)
  const cache = CacheService.getScriptCache();
  const cached = cache.get("TASA_COTIZAVE_CACHE");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  try {
    const options = {
      method: "get",
      headers: {
        "X-API-Key": API_KEY,
        "Accept": "application/json"
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(URL_API, options);
    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      let tasaUSD = 0;
      let tasaEUR = 0;

      if (data && Array.isArray(data.rates)) {
        // Buscar tasa USD
        const itemUSD = data.rates.find(r => r.base === "USD" || r.market === "reference");
        if (itemUSD && itemUSD.mid) tasaUSD = parseFloat(itemUSD.mid);

        // Buscar tasa EUR
        const itemEUR = data.rates.find(r => r.base === "EUR" || r.market === "eur_reference");
        if (itemEUR && itemEUR.mid) tasaEUR = parseFloat(itemEUR.mid);
      }

      if (tasaUSD > 0) {
        const resultado = {
          success: true,
          tasaUSD: tasaUSD,
          tasaEUR: tasaEUR,
          fecha: new Date().toISOString()
        };

        // Guardar en caché por 15 minutos (900 segundos)
        cache.put("TASA_COTIZAVE_CACHE", JSON.stringify(resultado), 900);

        // Opcional: auto-actualizar la celda C3 de tu hoja Tasas para que el resto del sistema la tenga
        try {
          const idArchivoExterno = "1fFyzPgSjBnmnHmJrKAV300B5aShq2ah9BnhlrjeXr2k";
          const sh = SpreadsheetApp.openById(idArchivoExterno).getSheetByName("Tasas");
          if (sh) {
            sh.getRange("C3").setValue(tasaUSD);
            if (tasaEUR > 0) sh.getRange("C4").setValue(tasaEUR);
          }
        } catch(eSheet) {
          // Si no tiene permisos de escritura, continúa sin detener el proceso
        }

        return resultado;
      }
    }
  } catch (error) {
    Logger.log("Error consultando API Cotizave: " + error.toString());
  }

  // 2. Respaldo a tu función de hoja externa si la API falla
  try {
    const backup = obtenerTasaCambioActual();
    return {
      success: true,
      tasaUSD: backup.tasaUSD || 1,
      tasaEUR: backup.tasaEUR || 1,
      fuente: "Hoja de respaldo"
    };
  } catch(e) {
    return { success: false, tasaUSD: 1, tasaEUR: 1 };
  }
}
// =============================================
/** ASIENTOS                     */
// =============================================
function getConstantsContables() {
  return {
    // 1. VENTAS
    DEBITO_FISCAL: {
      COD: "213005001",
      NOM: "IVA - DEBITO FISCAL",
      CTA: "Pasivo",
      SUB: "IMPUESTOS POR PAGAR",
      SUBCTA: "IMPUESTOS POR PAGAR"
    },
    CUENTA_POR_COBRAR: {
      COD: "112005001",
      NOM: "CUENTAS POR COBRAR COMERCIALES",
      CTA: "Activo",
      SUB: "CUENTAS POR COBRAR",
      SUBCTA: "CUENTAS POR COBRAR"
    },
    INGRESOS_VENTAS: {
      COD: "410505001",
      NOM: "VENTAS / INGRESOS",
      CTA: "Ingresos",
      SUB: "INGRESOS",
      SUBCTA: "INGRESOS"
    },
    RET_IVA_COBRAR: {
      COD: "115005008",
      NOM: "IVA - RETENCIONES POR DESCONTAR",
      CTA: "Activo",
      SUB: "IMPUESTOS",
      SUBCTA: "IMPUESTOS"
    },
    RET_ISLR_COBRAR: {
      COD: "115005001",
      NOM: "ISLR RETENIDO POR CLIENTES",
      CTA: "Activo",
      SUB: "IMPUESTOS",
      SUBCTA: "IMPUESTOS"
    },

    // 2. COMPRAS
    CREDITO_FISCAL: {
      COD: "115005006",
      NOM: "IVA - CREDITO FISCAL",
      CTA: "Activo",
      SUB: "IMPUESTOS",
      SUBCTA: "IMPUESTOS"
    },
    CUENTA_POR_PAGAR: {
      COD: "210505001",
      NOM: "CUENTAS POR PAGAR PROVEEDORES NAC.",
      CTA: "Pasivo",
      SUB: "CUENTAS POR PAGAR",
      SUBCTA: "CUENTAS POR PAGAR"
    },
    RET_IVA_PAGAR: {
      COD: "213005005",
      NOM: "IVA - RETENIDO POR PAGAR",
      CTA: "Pasivo",
      SUB: "IMPUESTOS POR PAGAR",
      SUBCTA: "IMPUESTOS POR PAGAR"
    },
    RET_ISLR_PAGAR: {
      COD: "213005004",
      NOM: "ISLR RETENIDO POR PAGAR",
      CTA: "Pasivo",
      SUB: "IMPUESTOS POR PAGAR",
      SUBCTA: "IMPUESTOS POR PAGAR"
    },

    // 3. INVENTARIOS Y COSTOS
    COSTO_VENTAS: {
      COD: "510505002",
      NOM: "COSTOS DE VENTAS",
      CTA: "Costos",
      SUB: "COSTOS",
      SUBCTA: "COSTOS"
    },
    INVENTARIO: {
      COD: "120505001",
      NOM: "INVENTARIO DE SUMINISTROS",
      CTA: "Activo",
      SUB: "INVENTARIO",
      SUBCTA: "INVENTARIO"
    },

    // 4. ANTICIPOS
    ANTICIPO_PROVEEDOR: {
      COD: "116005001",
      NOM: "ANTICIPOS A PROVEEDORES",
      CTA: "Activo",
      SUB: "CUENTAS POR PAGAR",
      SUBCTA: "CUENTAS POR PAGAR"
    },
    ANTICIPO_CLIENTE: {
      COD: "216005001",
      NOM: "AVANCES RECIBIDOS DE CLIENTES - ANTICIPOS",
      CTA: "Pasivo",
      SUB: "ANTICIPO CLIENTES",
      SUBCTA: "ANTICIPO CLIENTES"
    }
  };
}
function getCuentaContableCompleta(codigoCuenta) {
  try {
    const resultado = {
      CuentaContable: "",
      NombreCuenta: "",
      Cuenta: "",
      SubCuenta: ""
    };

    if (!codigoCuenta) return resultado;

    const codigo = String(codigoCuenta).trim();
    resultado.CuentaContable = codigo;

    const sheetCuentas = getSheetByName("PlandeCuentas");
    if (sheetCuentas) {
      const lastRow = sheetCuentas.getLastRow();
      if (lastRow > 1) {
        const data = sheetCuentas.getDataRange().getValues();
        const headers = data[0].map(h => String(h).trim());

        const idxCodigo = headers.indexOf("Cuenta Contable") !== -1
          ? headers.indexOf("Cuenta Contable")
          : headers.indexOf("CuentaContable");
        const idxNombre = headers.indexOf("Cuenta") !== -1
          ? headers.indexOf("Cuenta")
          : headers.indexOf("NombreCuenta");
        const idxSub = headers.indexOf("SubCuenta");

        if (idxCodigo !== -1 && idxNombre !== -1) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxCodigo]).trim() === codigo) {
              resultado.NombreCuenta = String(data[i][idxNombre] || "");
              if (idxSub !== -1) {
                resultado.SubCuenta = String(data[i][idxSub] || "");
              }
              break;
            }
          }
        }
      }
    }

    // Derivar Cuenta (primeros 4 dígitos) y SubCuenta (primeros 6 dígitos)
    if (!resultado.Cuenta && codigo.length >= 4) {
      resultado.Cuenta = codigo.substring(0, 4);
    }
    if (!resultado.SubCuenta && codigo.length >= 6) {
      resultado.SubCuenta = codigo.substring(0, 6);
    }

    return resultado;

  } catch (e) {
    console.error("Error en getCuentaContableCompleta: " + e.toString());
    return {
      CuentaContable: String(codigoCuenta || ""),
      NombreCuenta: "",
      Cuenta: "",
      SubCuenta: ""
    };
  }
}
function getSiguienteNroAsiento() {return generarCorrelativoAsiento();}
function obtenerUltimoNroFacturaAsientos() {  
  const sheetAsientos = getSheetByName('Asientos');
  const data = sheetAsientos.getDataRange().getValues();  

  const numerosAsientos = data.slice(1).map(row => row[0]);
  const numerosValidos = numerosAsientos.filter(nro => nro !== '' && !isNaN(nro)).map(Number);  
  const ultimoNumero = Math.max(...numerosValidos);  

  const fechaActual = new Date();  
  const mes = String(fechaActual.getMonth() + 1).padStart(2, '0');  
  const anio = String(fechaActual.getFullYear()).slice(2);  

  let nuevoNumeroFactura;  
  if (ultimoNumero !== -Infinity) {  
    const ultimoNumeroFormateado = String(ultimoNumero + 1).padStart(3, '0');  
    nuevoNumeroFactura = `${ultimoNumeroFormateado}-${mes}${anio}`;  
  } else {  
    nuevoNumeroFactura = `00`;  
  }  

  return nuevoNumeroFactura;  
}
function generarAsientoVenta(facturaData) {
  try {
    const asientoSheet = getSheetByName("Asientos");
    if (!asientoSheet) throw new Error("No se encontró la hoja 'Asientos'");

    const constants = getConstantsContables();
    const nroAsiento = generarCorrelativoAsiento();

    const fecha = facturaData.fechaFacturaCliente;
    const tipoAsiento = "Automático";
    const cliente = facturaData.nombreCliente;
    const descTransaccion = facturaData.tipoTransaccion || "Venta"; 
    const transaccionId = facturaData.nroFacturaCliente;
    const contabilidad = 1;
    const tasaCambio = parseFloat(facturaData.monedaTasaVenta) || 1;
    
    let filasAsiento = [];
    let totalCostoVentasBs = 0;

    // 1. DEBE: Cuentas por Cobrar (Total Factura)
    const cCxC = constants.CUENTA_POR_COBRAR;
    filasAsiento.push([
      nroAsiento,
      fecha,
      tipoAsiento,
      cliente,
      `Asiento ${nroAsiento} - Factura ${transaccionId}`,
      cCxC.COD,
      cCxC.NOM,
      parseFloat(facturaData.total) || 0,
      descTransaccion,
      transaccionId,
      contabilidad,
      cCxC.CTA,
      cCxC.SUBCTA || cCxC.SUB,
      "ES",
      "Activo Corriente",
      "Efectos y cuentas por cobrar comerciales"
    ]);

    // 2. HABER: Débito Fiscal IVA (Monto Negativo)
    const ivaTotalBs = parseFloat(facturaData.iva) || 0;
    if (ivaTotalBs > 0) {
      const cIva = constants.DEBITO_FISCAL;
      filasAsiento.push([
        nroAsiento,
        fecha,
        tipoAsiento,
        cliente,
        `IVA Factura ${transaccionId}`,
        cIva.COD,
        cIva.NOM,
        (ivaTotalBs * -1),
        descTransaccion,
        transaccionId,
        contabilidad,
        cIva.CTA,
        cIva.SUB,
        "ES",
        "Pasivo Corriente",
        "Retenciones y tributos por pagar"
      ]);
    }

    // 3. HABER: Ingresos por Ventas (Detalle de productos)
    facturaData.items.forEach((item) => {
      const montoIngresoBs = parseFloat(item.precioTotal) || 0;
      const cIngreso = constants.INGRESOS_VENTAS;
      const nombreProducto = item.nombre || item.producto || "Producto";
      
      filasAsiento.push([
        nroAsiento,
        fecha,
        tipoAsiento,
        cliente,
        `Venta Det- Factura ${transaccionId} - ${nombreProducto}`,
        cIngreso.COD,
        cIngreso.NOM,
        (montoIngresoBs * -1),
        descTransaccion,
        transaccionId,
        contabilidad,
        cIngreso.CTA,
        cIngreso.SUB,
        "ER",
        "Ingresos Operacionales",
        "Ingresos por actividades ordinarias"
      ]);
    });

    // 4. COSTO DE VENTAS Y SALIDA DE INVENTARIOS
    facturaData.items.forEach((item) => {
      const cantidad = parseFloat(item.cantidad) || 0;
      const precioCompraUsd = parseFloat(item.precioCompraUSD) || 0;
      
      if (precioCompraUsd > 0) {
        totalCostoVentasBs += (cantidad * precioCompraUsd) * tasaCambio;
      }
    });

    if (totalCostoVentasBs > 0) {
      // DEBE: Costo de Ventas
      const cCosto = constants.COSTO_VENTAS;
      filasAsiento.push([
        nroAsiento,
        fecha,
        tipoAsiento,
        cliente,
        `Costo de Ventas - Factura ${transaccionId}`,
        cCosto.COD,
        cCosto.NOM,
        totalCostoVentasBs,
        descTransaccion,
        transaccionId,
        contabilidad,
        cCosto.CTA,
        cCosto.SUB,
        "ER",
        "Costo de Ventas",
        "Costo de ventas"
      ]);

      // HABER: Salida de Inventario
      const cInv = constants.INVENTARIO;
      filasAsiento.push([
        nroAsiento,
        fecha,
        tipoAsiento,
        cliente,
        `Salida Inventario - Factura ${transaccionId}`,
        cInv.COD,
        cInv.NOM,
        (totalCostoVentasBs * -1),
        descTransaccion,
        transaccionId,
        contabilidad,
        cInv.CTA,
        cInv.SUB,
        "ES",
        "Activo Corriente",
        "Inventarios"
      ]);
    }

    // 5. AJUSTE DE CENTAVOS Y ESCRITURA EN HOJA
    let totalDebe = 0;
    let totalHaber = 0;
    filasAsiento.forEach(fila => {
      const monto = fila[7];
      if (monto > 0) totalDebe += monto;
      else totalHaber += Math.abs(monto);
    });

    const diferencia = totalDebe - totalHaber;
    if (Math.abs(diferencia) > 0 && Math.abs(diferencia) < 0.05) {
      filasAsiento[0][7] -= diferencia;
      totalDebe -= diferencia;
    }

    if (filasAsiento.length > 0) {
      const lastRow = asientoSheet.getLastRow();
      const targetRow = lastRow > 1 ? lastRow + 1 : 2;
      asientoSheet.getRange(targetRow, 1, filasAsiento.length, 16).setValues(filasAsiento);
    }

    return {
      nroAsiento: nroAsiento,
      debe: totalDebe,
      haber: totalHaber,
      balanceado: Math.abs(totalDebe - totalHaber) < 0.01
    };
    
  } catch (e) {
    console.error("Error en generarAsientoVenta: " + e.toString());
    throw e;
  }
}
function generarAsientoCompra(facturaData) {
  try {
    const sheetAsientos = getSheetByName("Asientos");
    if (!sheetAsientos) throw new Error("Hoja 'Asientos' no encontrada.");

    const consts = getConstantsContables();
    const nroAsiento = generarCorrelativoAsiento();

    const fecha = facturaData.FechaFacturaProveedor || new Date();
    const proveedor = facturaData.NombreProveedor || "Proveedor";
    const nroFactura = facturaData.NroFacturaProveedor || "S/F";
    const descTransaccion = facturaData.tipoTransaccion || "Compra";
    const totalBs = parseFloat(facturaData.TotalProveedor) || 0;
    const ivaBs = parseFloat(facturaData.IvaProveedor) || 0;
    const exentoBs = parseFloat(facturaData.totalExentoCompra) || 0;
    const baseImponible = totalBs - ivaBs - exentoBs;

    const filasAsiento = [];

    // 1. DEBE: Inventario / Gastos (Base Imponible)
    if (baseImponible > 0) {
      const cInventario = consts.INVENTARIO;
      filasAsiento.push([
        nroAsiento, fecha, "Automático", proveedor,
        `Compra Base Fact. ${nroFactura}`, 
        cInventario.COD, cInventario.NOM, baseImponible, 
        descTransaccion, nroFactura, 1, 
        cInventario.CTA, cInventario.SUB, "ES", "Activo Corriente", "Inventario"
      ]);
    }

    // 2. DEBE: IVA Crédito Fiscal
    if (ivaBs > 0) {
      const cCredito = consts.CREDITO_FISCAL;
      filasAsiento.push([
        nroAsiento, fecha, "Automático", proveedor,
        `IVA Crédito Fiscal Fact. ${nroFactura}`, 
        cCredito.COD, cCredito.NOM, ivaBs, 
        descTransaccion, nroFactura, 1, 
        cCredito.CTA, cCredito.SUB, "ES", "Activo Corriente", "Tributos por cobrar"
      ]);
    }

    // 3. DEBE: Exento (Si aplica)
    if (exentoBs > 0) {
      const cExento = consts.INVENTARIO; // O una cuenta de gasto específica
      filasAsiento.push([
        nroAsiento, fecha, "Automático", proveedor,
        `Compra Exenta Fact. ${nroFactura}`, 
        cExento.COD, cExento.NOM, exentoBs, 
        descTransaccion, nroFactura, 1, 
        cExento.CTA, cExento.SUB, "ES", "Activo Corriente", "Inventario"
      ]);
    }

    // 4. HABER: Cuentas por Pagar (Total Factura)
    const cCxP = consts.CUENTA_POR_PAGAR;
    filasAsiento.push([
      nroAsiento, fecha, "Automático", proveedor,
      `C x P Fact. ${nroFactura}`, 
      cCxP.COD, cCxP.NOM, (totalBs * -1), // Negativo para el Haber
      descTransaccion, nroFactura, 1, 
      cCxP.CTA, cCxP.SUB, "ES", "Pasivo Corriente", "Cuentas por pagar"
    ]);

    // 5. HABER: Retención IVA (Si aplica)
    const retIva = parseFloat(facturaData.RetencionIva) || 0;
    if (retIva > 0) {
      const cRetIva = consts.RET_IVA_PAGAR;
      filasAsiento.push([
        nroAsiento, fecha, "Automático", proveedor,
        `Retención IVA Fact. ${nroFactura}`, 
        cRetIva.COD, cRetIva.NOM, (retIva * -1), 
        descTransaccion, nroFactura, 1, 
        cRetIva.CTA, cRetIva.SUB, "ES", "Pasivo Corriente", "Retenciones por pagar"
      ]);
    }

    // 5. HABER: Retención ISLR (Si aplica)
    const retIslr = parseFloat(facturaData.RetencionIslr) || 0;
    if (retIslr > 0) {
      const cRetIslr = consts.RET_ISLR_PAGAR;
      filasAsiento.push([
        nroAsiento, fecha, "Automático", proveedor,
        `Retención ISLR Fact. ${nroFactura}`, 
        cRetIslr.COD, cRetIslr.NOM, (retIslr * -1), 
        descTransaccion, nroFactura, 1, 
        cRetIslr.CTA, cRetIslr.SUB, "ES", "Pasivo Corriente", "Retenciones por pagar"
      ]);
    }

    if (filasAsiento.length > 0) {
      const lastRow = sheetAsientos.getLastRow();
      const targetRow = lastRow > 1 ? lastRow + 1 : 2;
      sheetAsientos.getRange(targetRow, 1, filasAsiento.length, 16).setValues(filasAsiento);
    }

    return nroAsiento;

  } catch (e) {
    console.error("Error en generarAsientoCompra: " + e.toString());
    throw e;
  }
}
function generarAsientoMovimientoBancario(movData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAsientos = ss.getSheetByName("Asientos");
  if (!sheetAsientos) return "Sin Asiento (Hoja no encontrada)";

  const nroAsiento = generarCorrelativoAsiento();
  
  const fecha = movData.fecha;
  const beneficiario = movData.nombreBanco;
  const ref = movData.referenciaMov;
  const desc = movData.descripcionMov;
  const esEgreso = movData.tipoTransaccion.includes("Negativa");
  let montoBs = Math.round(parseFloat(movData.montoEquivalente) * 100) / 100;

  const limpiarCodigo = (str) => str ? String(str).split(" - ")[0].trim() : "";
  
  const codContrapartida = limpiarCodigo(movData.asocCodContable);
  const codBancoCaja = limpiarCodigo(movData.bankCodContable);

  const detallesCuentas = getCatalogosContables();
  
  const buscarDetalle = (codigo) => {
    const cuenta = detallesCuentas.find(c => c.cod === codigo);
    return cuenta ? {
      cuenta: cuenta.mayor || "ACTIVO", 
      subcuenta: cuenta.nom || "MOVIMIENTO",
      eeff: (cuenta.mayor === "ACTIVO" || cuenta.mayor === "PASIVO" || cuenta.mayor === "PATRIMONIO") ? "ES" : "ER",
      tipo: cuenta.mayor,
      grupo: cuenta.mayor
    } : { cuenta: "S/D", subcuenta: "S/D", eeff: "S/D", tipo: "S/D", grupo: "S/D" };
  };

  const detContra = buscarDetalle(codContrapartida);
  const detBanco = buscarDetalle(codBancoCaja);

  let filasAsiento = [];

  // FILA 1: CONTRAPARTIDA
  filasAsiento.push([
    nroAsiento,
    fecha,
    "Automático",
    beneficiario,
    desc,
    codContrapartida,
    movData.asocNomContable,
    esEgreso ? montoBs : (montoBs * -1),
    movData.tipoTransaccion,
    ref,
    1,
    detContra.cuenta,
    detContra.subcuenta,
    detContra.eeff,
    detContra.tipo,
    detContra.grupo
  ]);

  // FILA 2: BANCO O CAJA
  filasAsiento.push([
    nroAsiento,
    fecha,
    "Automático",
    beneficiario,
    `Ref: ${ref} - Registro Bancario`,
    codBancoCaja,
    movData.bankNomContable,
    esEgreso ? (montoBs * -1) : montoBs, 
    movData.tipoTransaccion,
    ref,
    1,
    detBanco.cuenta,
    detBanco.subcuenta,
    detBanco.eeff,
    detBanco.tipo,
    detBanco.grupo
  ]);

  sheetAsientos.getRange(sheetAsientos.getLastRow() + 1, 1, filasAsiento.length, 16).setValues(filasAsiento);
  return nroAsiento;
}
function construirAsientoAnticipo(esCXC, nroAsiento, fechaOp, tercero, glosa, refDoc, totalBs, ctaPayload, modalidad) {
  var constants = getConstantsContables();
  var fechaAsiento = new Date(fechaOp + 'T00:00:00');

  // 1. Obtener datos de la Caja o Banco seleccionada
  var ctaTesoreria = obtenerDatosContablesCuenta(ctaPayload);
  var metaTesoreria = _obtenerMetadatosCuentaCatalogo(
    ctaTesoreria.cod, 
    ctaTesoreria.nom, 
    'Activo', 
    (modalidad === 'Efectivo' ? 'CAJA' : 'BANCOS')
  );

  // 2. Obtener datos de la cuenta de Anticipo desde getConstantsContables()
  var ctaAnticipoDef = esCXC ? constants.ANTICIPO_CLIENTE : constants.ANTICIPO_PROVEEDOR;
  if (!ctaAnticipoDef || !ctaAnticipoDef.COD) {
    throw new Error('CONFIGURACIÓN: Falta cuenta de anticipo en getConstantsContables().');
  }

  var metaAnticipo = _obtenerMetadatosCuentaCatalogo(
    ctaAnticipoDef.COD,
    ctaAnticipoDef.NOM,
    ctaAnticipoDef.CTA || (esCXC ? 'Pasivo' : 'Activo'),
    ctaAnticipoDef.SUBCTA || ctaAnticipoDef.SUB || 'ANTICIPOS'
  );

  var lineas = [];

  if (esCXC) {
    // -------------------------------------------------------------
    // COBRO: ANTICIPO RECIBIDO DE CLIENTE (Ingreso de Fondos)
    // -------------------------------------------------------------
    // 1. DEBE: Caja / Banco (Aumenta Activo)
    lineas.push({
      nroAsiento: nroAsiento,
      fechaAsiento: fechaAsiento,
      tipoAsiento: 'Automático',
      clienteProveedor: tercero,
      descripcionGeneral: glosa,
      codigoCuenta: metaTesoreria.cod,
      cuentaContable: metaTesoreria.nom,
      montoTransaccion: totalBs, // DEBE (+)
      descripcionTransaccion: 'Ingreso por Anticipo (' + modalidad + ')',
      transaccionRef: refDoc,
      contabilidad: "1",
      mayor: metaTesoreria.mayor,
      subcuenta: metaTesoreria.subcuenta,
      eeff: metaTesoreria.eeff,
      tipoCuenta: metaTesoreria.tipoCuenta,
      grupoBalance: metaTesoreria.grupoBalance
    });

    // 2. HABER: Anticipos Recibidos de Clientes (Aumenta Pasivo)
    lineas.push({
      nroAsiento: nroAsiento,
      fechaAsiento: fechaAsiento,
      tipoAsiento: 'Automático',
      clienteProveedor: tercero,
      descripcionGeneral: glosa,
      codigoCuenta: metaAnticipo.cod,
      cuentaContable: metaAnticipo.nom,
      montoTransaccion: totalBs * -1, // HABER (-)
      descripcionTransaccion: 'Registro Anticipo Recibido',
      transaccionRef: refDoc,
      contabilidad: "1",
      mayor: metaAnticipo.mayor,
      subcuenta: metaAnticipo.subcuenta,
      eeff: metaAnticipo.eeff,
      tipoCuenta: metaAnticipo.tipoCuenta,
      grupoBalance: metaAnticipo.grupoBalance
    });

  } else {
    // -------------------------------------------------------------
    // PAGO: ANTICIPO OTORGADO A PROVEEDOR (Egreso de Fondos)
    // -------------------------------------------------------------
    // 1. DEBE: Anticipos a Proveedores (Aumenta Activo)
    lineas.push({
      nroAsiento: nroAsiento,
      fechaAsiento: fechaAsiento,
      tipoAsiento: 'Automático',
      clienteProveedor: tercero,
      descripcionGeneral: glosa,
      codigoCuenta: metaAnticipo.cod,
      cuentaContable: metaAnticipo.nom,
      montoTransaccion: totalBs, // DEBE (+)
      descripcionTransaccion: 'Registro Anticipo Otorgado',
      transaccionRef: refDoc,
      contabilidad: "1",
      mayor: metaAnticipo.mayor,
      subcuenta: metaAnticipo.subcuenta,
      eeff: metaAnticipo.eeff,
      tipoCuenta: metaAnticipo.tipoCuenta,
      grupoBalance: metaAnticipo.grupoBalance
    });

    // 2. HABER: Caja / Banco (Disminuye Activo)
    lineas.push({
      nroAsiento: nroAsiento,
      fechaAsiento: fechaAsiento,
      tipoAsiento: 'Automático',
      clienteProveedor: tercero,
      descripcionGeneral: glosa,
      codigoCuenta: metaTesoreria.cod,
      cuentaContable: metaTesoreria.nom,
      montoTransaccion: totalBs * -1, // HABER (-)
      descripcionTransaccion: 'Egreso por Anticipo (' + modalidad + ')',
      transaccionRef: refDoc,
      contabilidad: "1",
      mayor: metaTesoreria.mayor,
      subcuenta: metaTesoreria.subcuenta,
      eeff: metaTesoreria.eeff,
      tipoCuenta: metaTesoreria.tipoCuenta,
      grupoBalance: metaTesoreria.grupoBalance
    });
  }

  return lineas;
}
function obtenerCuentasContables(){try{const sheet=getSheetByName("PlandeCuentas");if(!sheet)return[];const lastRow=sheet.getLastRow();if(lastRow<2)return[];const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>String(h).trim());const idxCod=headers.indexOf("CuentaContable")!==-1?headers.indexOf("CuentaContable"):headers.indexOf("Cuenta Contable");const idxNom=headers.indexOf("NombreCuenta")!==-1?headers.indexOf("NombreCuenta"):headers.indexOf("Cuenta");const idxCue=headers.indexOf("Cuenta");const idxSub=headers.indexOf("SubCuenta");if(idxCod===-1)return[];const data=sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();return data.map(row=>({codigo:String(row[idxCod]||"").trim(),nombre:String(row[idxNom]||"").trim(),cuenta:idxCue!==-1?String(row[idxCue]||"").trim():"",subcuenta:idxSub!==-1?String(row[idxSub]||"").trim():""})).filter(c=>c.codigo&&c.nombre);}catch(e){Logger.log("Error en obtenerCuentasContables: "+e.toString());return[];}}

function getDataListaCuentaContable() {  
  const sheetPlandeCuentas = getSheetByName('PlandeCuentas');  
  const cuentaContableData = sheetPlandeCuentas.getDataRange().getDisplayValues();  
  const cuentaContablesHeaders = cuentaContableData.shift();  
  let cuentaContableDataArray = [];  

  cuentaContableData.forEach(dataRow => {  
    const temporaryObject = {};  
    cuentaContablesHeaders.forEach((headerTitle, index) => {  
      temporaryObject[headerTitle] = dataRow[index];  
    });  
    cuentaContableDataArray.push(temporaryObject);  
  });  

  return cuentaContableDataArray;  
}
function getListaCuentasContables() {
  try {
    const sheet = getSheetByName('PlandeCuentas');
    if (!sheet) throw new Error('Hoja PlandeCuentas no encontrada');
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const range = sheet.getRange('B2:C' + lastRow).getValues();
    return range
      .filter(row => row[0] && row[1])
      .map(row => ({
        codigo: String(row[0]),
        nombre: String(row[1])
      }));
  } catch (e) {
    console.error('Error en getListaCuentasContables:', e);
    return [];
  }
}
function getCatalogosContables() {
  try {
    var sheet = getSheetByName("PlandeCuentas");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    var headers = data[0].map(function(h) { return String(h).trim(); });
    
    // Mapeo dinámico de columnas según tu estructura real
    var colAgrupacion    = headers.indexOf("Agrupacion");
    var colCuentaCont    = headers.indexOf("Cuenta Contable");
    var colCuenta        = headers.indexOf("Cuenta");
    var colEEFF          = headers.indexOf("EEFF");
    var colOrden         = headers.indexOf("Orden");
    var colTipo          = headers.indexOf("Tipo");
    var colTipoCuenta    = headers.indexOf("TipoCuenta");
    var colGrupoBalance  = headers.indexOf("GrupoBalance");
    
    // Fallbacks por si cambian los nombres
    if (colAgrupacion   === -1) colAgrupacion   = 0;
    if (colCuentaCont   === -1) colCuentaCont   = 1;
    if (colCuenta       === -1) colCuenta       = 2;
    if (colEEFF         === -1) colEEFF         = 3;
    if (colOrden        === -1) colOrden        = 4;
    if (colTipo         === -1) colTipo         = 5;
    if (colTipoCuenta   === -1) colTipoCuenta   = 6;
    if (colGrupoBalance === -1) colGrupoBalance = 7;
    
    var resultado = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var cod = String(r[colCuentaCont] || '').trim();
      if (!cod) continue;
      
      resultado.push({
        sub:          String(r[colAgrupacion]   || '').trim(),  // Agrupacion
        cod:          cod,                                       // Cuenta Contable
        nom:          String(r[colCuenta]       || '').trim(),  // Cuenta
        eeff:         String(r[colEEFF]         || '').trim(),  // EEFF
        orden:        String(r[colOrden]        || '').trim(),  // Orden
        mayor:        String(r[colTipo]         || '').trim(),  // Tipo (Activo/Pasivo/Ingresos...)
        tipoCuenta:   String(r[colTipoCuenta]   || '').trim(),  // TipoCuenta
        grupoBalance: String(r[colGrupoBalance] || '').trim(),  // GrupoBalance
        agrupacion:   String(r[colAgrupacion]   || '').trim()   // Alias
      });
    }
    return resultado;
  } catch (e) {
    Logger.log("Error en getCatalogosContables: " + e.toString());
    return [];
  }
}
function _obtenerMetadatosCuentaCatalogo(codigoABuscar, fallbackNombre, mayorDefault, subDefault) {
  var codBuscado = String(codigoABuscar || '').split(' - ')[0].trim();
  
  var meta = {
    cod:          codBuscado,
    nom:          fallbackNombre || 'S/D',
    mayor:        mayorDefault || 'Activo',
    subcuenta:    subDefault || 'S/D',
    eeff:         (mayorDefault === 'Pasivo' || mayorDefault === 'Activo' || mayorDefault === 'Patrimonio') ? 'Balance General' : 'Estado de Resultados',
    tipoCuenta:   mayorDefault || 'Activo',
    grupoBalance: mayorDefault === 'Activo' ? 'Activo Corriente' : (mayorDefault === 'Pasivo' ? 'Pasivo Corriente' : 'Operacional'),
    agrupacion:   subDefault || 'S/D'
  };

  if (!codBuscado) return meta;

  try {
    var cat = getCatalogosContables();
    for (var i = 0; i < cat.length; i++) {
      if (String(cat[i].cod) === codBuscado) {
        meta.cod          = cat[i].cod;
        meta.nom          = cat[i].nom || fallbackNombre || meta.nom;
        meta.mayor        = cat[i].mayor || meta.mayor;
        meta.subcuenta    = cat[i].nom || subDefault || meta.subcuenta;
        meta.eeff         = cat[i].eeff || meta.eeff;
        meta.tipoCuenta   = cat[i].tipoCuenta || meta.tipoCuenta;
        meta.grupoBalance = cat[i].grupoBalance || meta.grupoBalance;
        meta.agrupacion   = cat[i].agrupacion || meta.agrupacion;
        break;
      }
    }
  } catch (err) {
    console.error('Error al resolver metadatos de cuenta en PlandeCuentas:', err);
  }

  return meta;
}
function getDataPlanCuentas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('PlandeCuentas');
    if (!sheet) throw new Error('No se encontró la hoja "PlandeCuentas"');
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    var headers = data[0];
    var result = [];
    var colMap = {};
    headers.forEach((header, index) => { colMap[String(header).trim()] = index; });
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var hasData = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
      if (!hasData) continue;
      
      result.push({
        CuentaContable: row[colMap['Cuenta Contable']] || '',
        Cuenta: row[colMap['Cuenta']] || '',
        EEFF: row[colMap['EEFF']] || '',
        Orden: row[colMap['Orden']] || '',
        Tipo: row[colMap['Tipo']] || '',
        Agrupacion: row[colMap['Agrupacion']] || '',
        TipoCuenta: row[colMap['TipoCuenta']] || '',
        GrupoBalance: row[colMap['GrupoBalance']] || ''
      });
    }
    return result;
  } catch (error) {
    console.error('Error en getDataPlanCuentas:', error);
    return { error: error.message };
  }
}
function getAsientosData() {
  try {
    const sheet = getSheetByName("Asientos");
    if (!sheet) return { error: "No se encontró la hoja 'Asientos'" };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { error: "La hoja está vacía" };

    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const getCol = (name) => headers.indexOf(name.toUpperCase());

    const idx = {
      nro: getCol("NROASIENTO"),
      fecha: getCol("FECHAASIENTO"),
      cliente: getCol("CLIENTE/PROVEEDOR"),
      desc: getCol("DESCRIPCION"),
      cuenta: getCol("CUENTA"),
      sub: getCol("SUBCUENTA"),
      nombre: getCol("CUENTACONTABLE"),
      monto: getCol("MONTOTRANSACCION")
    };

    if (idx.cuenta === -1 || idx.sub === -1 || idx.monto === -1) {
      return { error: "Faltan columnas obligatorias (CUENTA, SUBCUENTA o MONTOTRANSACCION)" };
    }

    const rows = data.slice(1);
    const asientos = rows.map((r) => {
      if (!r[idx.cuenta] && !r[idx.monto]) return null;

      let m = r[idx.monto];
      if (typeof m === 'string') m = parseFloat(m.replace(/\./g, '').replace(',', '.')) || 0;

      return {
        nro: idx.nro !== -1 ? String(r[idx.nro]) : "S/N",
        fecha: (idx.fecha !== -1 && r[idx.fecha] instanceof Date) ? r[idx.fecha].getTime() : null,
        cliente: idx.cliente !== -1 ? String(r[idx.cliente]) : "",
        descripcion: idx.desc !== -1 ? String(r[idx.desc]) : "",
        cuenta: String(r[idx.cuenta]),
        subcuenta: String(r[idx.sub]),
        cuentaContable: String(r[idx.nombre] || "Sin nombre"),
        monto: parseFloat(m) || 0
      };
    }).filter(x => x !== null);

    // Obtener tasas actualizadas desde el archivo externo (Hoja "Tasas")
    let tasaUSD = 1;
    let tasaEUR = 1;
    try {
      // Aprovecha la función que creamos antes para consultar el archivo externo
      const tasasActuales = obtenerTasaCambioActual();
      tasaUSD = tasasActuales.tasaUSD || 1;
      tasaEUR = tasasActuales.tasaEUR || 1;
    } catch(e) {
      console.warn("⚠️ No se pudieron obtener las tasas externas, usando valores por defecto:", e);
    }

    return { 
      asientos, 
      tasa: tasaUSD,      // Mantiene compatibilidad si tu código usa 'tasa' para USD
      tasaUSD: tasaUSD,   // Tasa del dólar (C1)
      tasaEUR: tasaEUR    // Tasa del euro (E1)
    };

  } catch (e) {
    return { error: "Error en servidor: " + e.toString() };
  }
}
function formatDate(date) {
  if (!date) return '';
  let d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  let year = d.getFullYear();
  let month = String(d.getMonth() + 1).padStart(2, '0');
  let day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}
function getVentas() {
  try {
    const HOJA_VENTAS = getSheetByName('Ventas');
    if (!HOJA_VENTAS) throw new Error("No se encontró la hoja 'Ventas'.");
    
    const data = HOJA_VENTAS.getDataRange().getValues();
    const headers = data.shift();

    const retIvaIndex = headers.indexOf("RetencionIva");
    const retIslrIndex = headers.indexOf("RetencionISLR");

    if (retIvaIndex === -1 || retIslrIndex === -1) {
      throw new Error("Asegúrate de que las columnas 'RetencionIva' y 'RetencionISLR' existan en la hoja 'Ventas'.");
    }

    const ventasJson = data.map(row => {
      let venta = {};
      headers.forEach((header, i) => {
        if (header && typeof header === 'string' && header.trim() !== '') {
          if (header.toLowerCase().includes('fecha') && row[i] instanceof Date) {
              venta[header] = row[i].toISOString();
          } else {
              venta[header] = row[i];
          }
        }
      });
      
      venta.tieneRetencionIva = !!row[retIvaIndex];
      venta.tieneRetencionIslr = !!row[retIslrIndex];
      return venta;
    });
    
    return ventasJson.filter(v => v.SaldoCliente && parseFloat(v.SaldoCliente) > 0.01);

  } catch (e) {
    Logger.log(e);
    throw new Error("Error al obtener datos de Ventas: " + e.message);
  }
}
function getCompras() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Compras");
    if (!hoja) throw new Error("Hoja 'Compras' no encontrada");
    
    const data = hoja.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const headers = data[0];
    const colSaldoBs = headers.indexOf("SaldoProveedor");
    const colSaldoUSD = headers.indexOf("SaldoProveedorDivisa");
    const colEstado = headers.indexOf("cobradaPagada");
    const colNroFactura = headers.indexOf("NroFacturaProveedor");
    const colNombre = headers.indexOf("NombreProveedor");
    const colRif = headers.indexOf("RifProveedor");
    const colIva = headers.indexOf("IvaProveedor");
    
    const parseNum = (val) => {
      if (typeof val === 'number') return val;
      if (!val) return 0;
      const str = String(val).replace(/\./g, '').replace(',', '.');
      const num = parseFloat(str);
      return isNaN(num) ? 0 : num;
    };
    
    const compras = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const saldoBs = colSaldoBs !== -1 ? parseNum(row[colSaldoBs]) : 0;
      const saldoUSD = colSaldoUSD !== -1 ? parseNum(row[colSaldoUSD]) : 0;
      const estado = String(row[colEstado] || "").toLowerCase().trim();
      const estadosPagados = ['pagada', 'pagado', 'cobrada', 'cobrado', 'pagada completa', 'pagado completo'];
      
      if ((saldoBs > 0.01 || saldoUSD > 0.01) && !estadosPagados.includes(estado)) {
        compras.push({
          Id: String(row[headers.indexOf("Id")] || row[colNroFactura] || "").trim(),
          NombreProveedor: String(row[colNombre] || "").trim(),
          RifProveedor: String(row[colRif] || "").trim(),
          NroFacturaProveedor: String(row[colNroFactura] || "").trim(),
          SaldoProveedor: saldoBs,
          SaldoProveedorDivisa: saldoUSD,
          IvaProveedor: colIva !== -1 ? parseNum(row[colIva]) : 0,
          cobradaPagada: row[colEstado] || "porPagar"
        });
      }
    }
    return compras;
  } catch (error) {
    Logger.log("❌ Error en getCompras: " + error.toString());
    throw new Error("Error al obtener compras: " + error.message);
  }
}
function deleteVenta(nroFactura) {
  try {
    const sheet = getSheetByName("Ventas");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nroFacturaIndex = headers.indexOf("NroFacturaCliente");

    for (let i = 1; i < data.length; i++) {
      if (data[i][nroFacturaIndex] == nroFactura) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return true;
  } catch (error) {
    console.error("Error en deleteVenta:", error);
    throw error;
  }
}
function updateVenta(updatedVenta) {
  try {
    const sheet = getSheetByName("Ventas");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nroFacturaIndex = headers.indexOf("NroFacturaCliente");

    for (let i = 1; i < data.length; i++) {
      if (data[i][nroFacturaIndex] == updatedVenta.NroFacturaCliente) {
        for (const [key, value] of Object.entries(updatedVenta)) {
          const colIndex = headers.indexOf(key);
          if (colIndex !== -1) {
            sheet.getRange(i + 1, colIndex + 1).setValue(value);
          }
        }
        break;
      }
    }
    return true;
  } catch (error) {
    console.error("Error en updateVenta:", error);
    throw error;
  }
}
function guardarAsientoMasivo(filas) {
  if (!filas || filas.length === 0) return "";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Asientos");
  if (!sheet) throw new Error("Hoja 'Asientos' no encontrada.");

  var rows = filas.map(function(f) {
    return [
      f.nroAsiento,                               // Col A: NroAsiento
      f.fechaAsiento,                             // Col B: FechaAsiento
      f.tipoAsiento || 'Automático',             // Col C: TipoAsiento
      f.clienteProveedor || '',                   // Col D: cliente/Proveedor
      f.descripcionGeneral || '',                 // Col E: Descripcion
      f.codigoCuenta || '',                       // Col F: CodigoCuenta
      f.cuentaContable || '',                     // Col G: CuentaContable
      f.montoTransaccion,                         // Col H: MontoTransaccion
      f.descripcionTransaccion || '',             // Col I: descripcionTransaccion
      f.transaccionRef || '',                     // Col J: Transaccion
      f.contabilidad != null ? f.contabilidad : "1", // Col K: Contabilidad
      f.mayor || '',                              // Col L: CUENTA
      f.subcuenta || '',                          // Col M: SUBCUENTA
      f.eeff || '',                               // Col N: EEFF
      f.tipoCuenta || '',                         // Col O: TipoCuenta
      f.grupoBalance || ''                        // Col P: GrupoBalance
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 16).setValues(rows);
  return filas[0].nroAsiento;
}
function guardarDatosEnAsiento(registros) {
  const hoja = getSheetByName("asientos");
  const ultimaFila = hoja.getLastRow();

  registros.forEach(function(registro) {
    const datos = [
      registro.nroAsiento,
      registro.FechaFactura,
      "TipoAsiento",
      registro.clienteProveedor,
      "Descripcion",
      "CodigoCuenta",
      "CuentaContable",
      registro.Total,
      registro.Contabilidad
    ];

    hoja.appendRow(datos);
  });
}
function parseFechaSafe(valor) {  
  if (!valor) {
    return null;
  }
  
  // Si ya es un Date válido, retornar
  if (valor instanceof Date) {
    return isValidDate(valor) ? valor : null;
  }
  
  // Convertir a string
  var str = String(valor).trim();
  
  if (str === '') {
    return null;
  }
  
  // ============================================================
  // INTENTAR PARSEAR COMO FECHA
  // ============================================================
  
  var fecha = null;
  
  // FORMATO 1: YYYY-MM-DD (ISO)
  var regexISO = /^(\d{4})-(\d{2})-(\d{2})$/;
  var matchISO = str.match(regexISO);
  if (matchISO) {
    fecha = new Date(matchISO[1] + '-' + matchISO[2] + '-' + matchISO[3] + 'T00:00:00Z');
    if (isValidDate(fecha)) return fecha;
  }
  
  // FORMATO 2: DD/MM/YYYY
  var regexDMY1 = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
  var matchDMY1 = str.match(regexDMY1);
  if (matchDMY1) {
    var dia = parseInt(matchDMY1[1], 10);
    var mes = parseInt(matchDMY1[2], 10);
    var año = parseInt(matchDMY1[3], 10);
    fecha = new Date(año, mes - 1, dia);
    if (isValidDate(fecha)) return fecha;
  }
  
  // FORMATO 3: DD-MM-YYYY
  var regexDMY2 = /^(\d{1,2})-(\d{1,2})-(\d{4})$/;
  var matchDMY2 = str.match(regexDMY2);
  if (matchDMY2) {
    var dia = parseInt(matchDMY2[1], 10);
    var mes = parseInt(matchDMY2[2], 10);
    var año = parseInt(matchDMY2[3], 10);
    fecha = new Date(año, mes - 1, dia);
    if (isValidDate(fecha)) return fecha;
  }
  
  // FORMATO 4: MM/DD/YYYY (anglosajón)
  var regexMDY = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
  var matchMDY = str.match(regexMDY);
  if (matchMDY) {
    // Intentar como DD/MM primero
    var dia = parseInt(matchMDY[1], 10);
    var mes = parseInt(matchMDY[2], 10);
    var año = parseInt(matchMDY[3], 10);
    
    // Si el primer número es > 12, definitivamente es día
    if (dia > 12) {
      fecha = new Date(año, mes - 1, dia);
    } else if (mes > 12) {
      // Si el segundo número es > 12, definitivamente es día
      fecha = new Date(año, dia - 1, mes);
    } else {
      // Ambiguo: asumir DD/MM (formato latino)
      fecha = new Date(año, mes - 1, dia);
    }
    
    if (isValidDate(fecha)) return fecha;
  }
  
  // FORMATO 5: Intentar parseo estándar de JavaScript
  fecha = new Date(str);
  if (isValidDate(fecha)) return fecha;
  
  // ============================================================
  // SI NADA FUNCIONA, RETORNAR NULL
  // ============================================================
  
  Logger.log('[PARSE FECHA] No se pudo parsear: ' + str);
  return null;
}
function isValidDate(fecha) {
  if (!(fecha instanceof Date)) return false;
  return !isNaN(fecha.getTime());
}
function insertarLineaAsiento(sheet, d) {
  // NroAsiento, FechaAsiento, TipoAsiento, cliente/Proveedor, Descripcion, CodigoCuenta, CuentaContable, MontoTransaccion, descripcionTransaccion, Transaccion, Contabilidad, CUENTA, SUBCUENTA
  sheet.appendRow([
    d.nroAsiento,        // A
    d.fecha,             // B
    "Automático",        // C
    d.cliente,           // D
    d.glosa,             // E
    d.codCuenta,         // F
    d.nomCuenta,         // G
    d.monto,             // H (MontoTransaccion - Positivo Debe / Negativo Haber)
    d.glosaTransaccion,  // I
    d.nroDocumento,      // J (Referencia)
    "1",                 // K
    d.categoria,         // L (Ej: ACTIVO / PASIVO)
    d.subcategoria       // M (Ej: BANCOS / ANTICIPOS)
  ]);
}
function guardarAsientoContable(asiento, idSpreadsheet, auth) {
  // Validación estricta de partida doble
  let sumaDebeUSD = 0;
  let sumaHaberUSD = 0;

  asiento.lineas.forEach(function(l) {
    sumaDebeUSD += Number(l.debeUSD) || 0;
    sumaHaberUSD += Number(l.haberUSD) || 0;
  });

  const diferencia = Math.abs(sumaDebeUSD - sumaHaberUSD);
  if (diferencia > 0.01) {
    throw new Error("Asiento descuadrado. Suma Debe: " + sumaDebeUSD.toFixed(2) + " vs Suma Haber: " + sumaHaberUSD.toFixed(2));
  }

  const ss = Contexto.getEmpresaSS(idSpreadsheet);
  const sheetAsientos = ss.getSheetByName("Asientos");
  if (!sheetAsientos) throw new Error("Hoja Asientos no encontrada.");

  const nroAsiento = asiento.numero || ("ASI-" + Date.now().toString().slice(-6));
  const fecha = asiento.fecha || new Date();

  // Inserción de líneas contables
  const filas = asiento.lineas.map(function(l) {
    return [
      nroAsiento,
      fecha,
      asiento.concepto,
      l.codigoCuenta,
      l.nombreCuenta,
      l.debeUSD,
      l.haberUSD,
      l.debeUSD * (asiento.tasaBCV || 1),
      l.haberUSD * (asiento.tasaBCV || 1),
      asiento.origenModulo || "Manual",
      asiento.usuario || "Sistema"
    ];
  });

  sheetAsientos.getRange(sheetAsientos.getLastRow() + 1, 1, filas.length, 11).setValues(filas);
  return { success: true, nroAsiento: nroAsiento, totalUSD: sumaDebeUSD };
}

function calcularBalanceGeneral(idSpreadsheet) {
  const ss = Contexto.getEmpresaSS(idSpreadsheet);
  const sheetPlan = ss.getSheetByName("PlandeCuentas");
  const data = sheetPlan.getDataRange().getValues();

  let totalActivo = 0;
  let totalPasivo = 0;
  let totalPatrimonio = 0;

  for (let i = 1; i < data.length; i++) {
    const tipo = String(data[i][2]).trim();
    const nivel = Number(data[i][3]);
    const saldo = Number(data[i][4]) || 0;

    if (nivel === 1) {
      if (tipo === "Activo") totalActivo = saldo;
      if (tipo === "Pasivo") totalPasivo = saldo;
      if (tipo === "Patrimonio") totalPatrimonio = saldo;
    }
  }

  return {
    activo: totalActivo,
    pasivo: totalPasivo,
    patrimonio: totalPatrimonio,
    cuadrado: Math.abs(totalActivo - (totalPasivo + totalPatrimonio)) < 0.01
  };
}
function generarXmlSeniatISLR(periodoFiscal, idSpreadsheet) {
  const ss = Contexto.getEmpresaSS(idSpreadsheet);
  const sheetEmpresa = Contexto.getMasterSS().getSheetByName("Empresas");
  
  // Obtener RIF del Agente de Retención
  let rifAgente = "J000000000";
  const dataEmpresas = sheetEmpresa.getDataRange().getValues();
  for (let i = 1; i < dataEmpresas.length; i++) {
    if (String(dataEmpresas[i][3]).trim() === idSpreadsheet.trim()) {
      rifAgente = String(dataEmpresas[i][2]).replace(/[^A-Za-z0-9]/g, '');
      break;
    }
  }

  const sheetRet = ss.getSheetByName("CobrosPagos");
  const dataRet = sheetRet.getDataRange().getValues();

  let xml = '<?xml version="1.0" encoding="ISO-8859-1"?>\n';
  xml += '<RelacionRetencionesISLR RifAgente="' + rifAgente + '" Periodo="' + periodoFiscal + '">\n';

  for (let i = 1; i < dataRet.length; i++) {
    const tipo = String(dataRet[i][1]).trim();
    if (tipo === "RetencionISLR") {
      const rifRetenido = String(dataRet[i][3]).replace(/[^A-Za-z0-9]/g, '');
      const factura = String(dataRet[i][4]);
      const control = String(dataRet[i][5]);
      const fecha = Utilities.formatDate(new Date(dataRet[i][2]), "America/Caracas", "yyyy-MM-dd");
      const concepto = String(dataRet[i][10] || "055"); // 055=Bienes, 056=Flete, 057=Honorarios, 058=Arrendamiento
      const montoBs = Number(dataRet[i][7] || 0).toFixed(2);
      const porcentaje = Number(dataRet[i][8] || 2).toFixed(2);

      xml += '  <DetalleRetencion>\n';
      xml += '    <RifRetenido>' + rifRetenido + '</RifRetenido>\n';
      xml += '    <NumeroFactura>' + factura + '</NumeroFactura>\n';
      xml += '    <NumeroControl>' + control + '</NumeroControl>\n';
      xml += '    <FechaOperacion>' + fecha + '</FechaOperacion>\n';
      xml += '    <CodigoConcepto>' + concepto + '</CodigoConcepto>\n';
      xml += '    <MontoOperacion>' + montoBs + '</MontoOperacion>\n';
      xml += '    <PorcentajeRetencion>' + porcentaje + '</PorcentajeRetencion>\n';
      xml += '  </DetalleRetencion>\n';
    }
  }

  xml += '</RelacionRetencionesISLR>';
  return xml;
}
// ==========================================================
/** FUNCIONES BANCOS Y CAJAS  */
// ==========================================================
const BANCO_COLS = {
  START: 1,
  ID: 1,        // A IdBanco
  CODIGO: 2,    // B CodigoBancos
  NOMBRE: 3,    // C NombreBanco
  MONEDA: 4,    // D MonedaCta
  TASA: 5,      // E TasaBanco
  SALDO: 6,     // F SaldoCuenta
  COUNT: 6
};

const CUENTA_COLS = {
  START: 8,           // H
  ID: 8,              // H IdCuenta
  CODIGO: 9,          // I CodigoCuenta
  NUMERO: 10,         // J NroCuenta
  ID_BANCO_REF: 11,   // K CodigoBancosCuenta
  NOM_BANCO: 12,      // L NombreBancoCuenta
  MONEDA: 13,         // M MonedaBanco
  TIPO: 14,           // N tipoCuenta
  SALDO: 15,          // O SaldoBanco
  COD_CONTABLE: 16,   // P CodigoContable
  NOM_CONTABLE: 17,   // Q CuentaContable
  COUNT: 10
};

const CAJA_COLS = {
  START: 19,          // S
  ID: 19,             // S IdCaja
  CODIGO: 20,         // T CodigoCaja
  NOMBRE: 21,         // U NombreCaja
  MONEDA: 22,         // V MonedaCaja
  TASA: 23,           // W TasaCaja
  SALDO: 24,          // X SaldoCaja
  COD_CONTABLE: 25,   // Y CodigoContable
  NOM_CONTABLE: 26,   // Z CuentaContable
  COUNT: 8
};

// ==========================================================
// RESOLUTOR MULTI-EMPRESA
// ==========================================================
function _SS_BC_(idOpcional) {
  var idTarget = idOpcional || 
                 (typeof Contexto !== 'undefined' && Contexto.getEmpresaActivaId ? Contexto.getEmpresaActivaId() : null) ||
                 PropertiesService.getUserProperties().getProperty('EMPRESA_ACTIVA_ID') ||
                 (typeof ID_HOJA !== 'undefined' ? ID_HOJA : null);

  if (typeof Contexto !== 'undefined' && Contexto.getEmpresaSS && idTarget) {
    return Contexto.getEmpresaSS(idTarget);
  }
  if (idTarget && typeof idTarget === 'string' && idTarget.length > 15) {
    return SpreadsheetApp.openById(idTarget);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}
function getSheetBancos(idSpreadsheetEmpresa) {
  var ss = _SS_BC_(idSpreadsheetEmpresa);
  if (!ss) return null;
  return ss.getSheetByName('datosBancarios') || ss.getSheetByName('Bancos');
}
function getListaBancos(idSpreadsheetEmpresa) {
  try {
    const sheet = getSheetBancos(idSpreadsheetEmpresa);
    if (!sheet) return [];
    const datos = sheet.getDataRange().getValues();
    const bancos = [];
    const idx = (col) => col - 1;

    for (let i = 1; i < datos.length; i++) {
      const codigoBanco = datos[i][idx(BANCO_COLS.CODIGO)];
      if (codigoBanco && codigoBanco.toString().trim() !== "") {
        bancos.push([
          datos[i][idx(BANCO_COLS.ID)],
          codigoBanco,
          datos[i][idx(BANCO_COLS.NOMBRE)],
          datos[i][idx(BANCO_COLS.MONEDA)],
          datos[i][idx(BANCO_COLS.TASA)],
          datos[i][idx(BANCO_COLS.SALDO)]
        ]);
      }
    }
    return bancos;
  } catch (error) {
    console.error("Error al obtener bancos:", error);
    return [];
  }
}
function getListaCuentas(idSpreadsheetEmpresa) {
  try {
    const sheet = getSheetBancos(idSpreadsheetEmpresa);
    if (!sheet) return [];
    const datos = sheet.getDataRange().getValues();
    const cuentas = [];
    const idx = (col) => col - 1;

    for (let i = 1; i < datos.length; i++) {
      const nroCuenta = datos[i][idx(CUENTA_COLS.NUMERO)];
      if (nroCuenta && nroCuenta.toString().trim() !== "") {
        cuentas.push([
          datos[i][idx(CUENTA_COLS.ID)],
          datos[i][idx(CUENTA_COLS.ID_BANCO_REF)],
          datos[i][idx(CUENTA_COLS.NOM_BANCO)],
          datos[i][idx(CUENTA_COLS.CODIGO)],
          nroCuenta,
          datos[i][idx(CUENTA_COLS.MONEDA)],
          datos[i][idx(CUENTA_COLS.TIPO)],
          datos[i][idx(CUENTA_COLS.SALDO)],
          datos[i][idx(CUENTA_COLS.COD_CONTABLE)],
          datos[i][idx(CUENTA_COLS.NOM_CONTABLE)]
        ]);
      }
    }
    return cuentas;
  } catch (error) {
    console.error("Error al obtener cuentas:", error);
    return [];
  }
}
function getListaCajas(idSpreadsheetEmpresa) {
  try {
    const sheet = getSheetBancos(idSpreadsheetEmpresa);
    if (!sheet) return [];
    const datos = sheet.getDataRange().getValues();
    const cajas = [];
    const idx = (col) => col - 1;

    for (let i = 1; i < datos.length; i++) {
      const codigoCaja = datos[i][idx(CAJA_COLS.CODIGO)];
      if (codigoCaja && codigoCaja.toString().trim() !== "") {
        cajas.push([
          datos[i][idx(CAJA_COLS.ID)],
          codigoCaja,
          datos[i][idx(CAJA_COLS.NOMBRE)],
          datos[i][idx(CAJA_COLS.MONEDA)],
          datos[i][idx(CAJA_COLS.TASA)],
          datos[i][idx(CAJA_COLS.SALDO)],
          datos[i][idx(CAJA_COLS.COD_CONTABLE)],
          datos[i][idx(CAJA_COLS.NOM_CONTABLE)]
        ]);
      }
    }
    return cajas;
  } catch (error) {
    console.error("Error al obtener cajas:", error);
    return [];
  }
}
function saveData(data) {
  try {
    const sheet = getSheetBancos(data.idSpreadsheet);
    if (!sheet) throw new Error('No se encontró la hoja de bancos');

    const id = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyyyyddHHmmss') + 
               Math.floor(Math.random() * 1000);
    
    const targetRow = getNextAvailableRow(sheet, data.type);
    const finalRow = Math.max(targetRow, 2);
    
    if (data.type === 'banco') {
      sheet.getRange(finalRow, BANCO_COLS.START, 1, BANCO_COLS.COUNT).setValues([[
        id,
        data.codigo,
        data.nombre,
        data.moneda,
        parseFloat(data.tasa) || 0,
        parseFloat(data.saldo) || 0
      ]]);
      
    } else if (data.type === 'cuenta') {
      let codigoBanco = data.banco || '';
      let nombreBanco = '';
      if (data.banco && data.banco.includes('|')) {
        const parts = data.banco.split('|');
        codigoBanco = parts[0];
        nombreBanco = parts[1] || '';
      }
      
      sheet.getRange(finalRow, CUENTA_COLS.START, 1, CUENTA_COLS.COUNT).setValues([[
        id,
        codigoBanco,
        nombreBanco,
        data.codigo,
        data.numero,
        data.moneda,
        data.tipo,
        0,
        data.codContable || "",
        data.nomContable || ""
      ]]);
      
    } else if (data.type === 'caja') {
      sheet.getRange(finalRow, CAJA_COLS.START, 1, CAJA_COLS.COUNT).setValues([[
        id,
        data.codigo,
        data.nombre,
        data.moneda,
        parseFloat(data.tasa) || 0,
        parseFloat(data.saldo) || 0,
        data.codContable || "",
        data.nomContable || ""
      ]]);
    }
    
    return {
      success: true,
      message: `${data.type} registrado correctamente`,
      id: id,
      row: finalRow
    };
    
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}
function deleteRecord(id, type, idSpreadsheetEmpresa) {
  try {
    const sheet = getSheetBancos(idSpreadsheetEmpresa);
    if (!sheet) throw new Error('Hoja de bancos no encontrada');
    
    const CONFIG_RANGOS = {
      banco:  { idCol: BANCO_COLS.ID,  start: BANCO_COLS.START,  count: BANCO_COLS.COUNT },
      cuenta: { idCol: CUENTA_COLS.ID, start: CUENTA_COLS.START, count: CUENTA_COLS.COUNT },
      caja:   { idCol: CAJA_COLS.ID,   start: CAJA_COLS.START,   count: CAJA_COLS.COUNT }
    };
    
    const config = CONFIG_RANGOS[type.toLowerCase()];
    if (!config) return { success: false, message: "Tipo no válido: " + type };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No hay datos para eliminar" };
    
    const idColumn = sheet.getRange(2, config.idCol, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < idColumn.length; i++) {
      if (String(idColumn[i][0]) === String(id)) {
        const rowNum = i + 2;
        sheet.getRange(rowNum, config.start, 1, config.count).clearContent();
        return { success: true, message: `${type} eliminado correctamente` };
      }
    }
    return { success: false, message: "No se encontró el registro con ID: " + id };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}
function agregarEditarMovBancoCaja(data) {
  try {
    if (!data) return { success: false, message: 'Datos vacíos' };

    var ss = _SS_BC_(data.idSpreadsheet);
    if (!ss) throw new Error('No se pudo acceder a la hoja de la empresa activa');

    var hojaCabecera = ss.getSheetByName('MovBancoCaja');
    if (!hojaCabecera) throw new Error('No se encontró la hoja "MovBancoCaja"');

    // Verificar o crear la hoja detalle si no existe en la empresa activa
    var hojaDetalle = ss.getSheetByName('detalleBancosCajas');
    if (!hojaDetalle) {
      hojaDetalle = ss.insertSheet('detalleBancosCajas');
      hojaDetalle.appendRow([
        'IdMovimiento', 'FechaMov', 'TipoMov', 'NroCuenta', 'NombreBanco',
        'CodigoBancos', 'Cuenta', 'ReferenciaMov', 'descMov', 'TipoTransaccion',
        'MontoMovUSD', 'MonedaMov', 'TasaMov', 'IGTF', 'MontoMovBs',
        'CodigoCuentaContable', 'CuentaContable', 'CuentaGasto', 'NombreCuentaMov', 'fechaRegistroMov'
      ]);
    }

    var idMov = (data.IdMovimiento || '').toString().trim();
    var esEdicion = !!idMov;
    var filaExistente = -1;

    if (esEdicion) {
      var lastRow = hojaCabecera.getLastRow();
      if (lastRow >= 2) {
        var ids = hojaCabecera.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim() === idMov) { filaExistente = i + 2; break; }
        }
      }
      if (filaExistente === -1) {
        return { success: false, message: 'No se encontró el movimiento a editar (ID: ' + idMov + ')' };
      }
    } else {
      idMov = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyyyyddHHmmss')
            + Math.floor(Math.random() * 1000);
    }

    var fechaMov = data.FechaMov;
    if (typeof fechaMov === 'string') {
      var parts = fechaMov.split('-');
      if (parts.length === 3) {
        fechaMov = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      } else {
        fechaMov = new Date();
      }
    } else if (!(fechaMov instanceof Date) || isNaN(fechaMov.getTime())) {
      fechaMov = new Date();
    }

    var tasa = parseFloat(data.TasaMov) || 1;
    var lineas = Array.isArray(data.detalles) && data.detalles.length > 0 ? data.detalles : [];

    var totalUSD = 0;
    var totalBs = 0;

    if (lineas.length > 0) {
      lineas.forEach(function(lin) {
        var u = parseFloat(lin.montoUSD) || 0;
        var b = parseFloat(lin.montoBs) || (u * tasa);
        totalUSD += u;
        totalBs += b;
      });
    } else {
      totalUSD = parseFloat(data.MontoMovUSD) || 0;
      totalBs = parseFloat(data.MontoMovBs) || (totalUSD * tasa);
    }

    var codContableCab = lineas.length === 1 ? (lineas[0].codigo || '') : (lineas.length > 1 ? 'MULTILINEA' : (data.CodigoCuentaContable || ''));
    var nomContableCab = lineas.length === 1 ? (lineas[0].nombre || '') : (lineas.length > 1 ? 'Distribución Multilínea (' + lineas.length + ' cuentas)' : (data.CuentaContable || ''));

    var filaCabecera = [
      idMov, fechaMov, data.TipoMov || '', data.NroCuenta || '', data.NombreBanco || '',
      data.CodigoBancos || '', data.Cuenta || '', data.ReferenciaMov || '', data.descMov || '',
      data.TipoTransaccion || '', totalUSD, data.MonedaMov || '', tasa,
      parseFloat(data.IGTF) || 0, totalBs, codContableCab,
      nomContableCab, data.CuentaGasto || '', data.NombreCuentaMov || '',
      data.fechaRegistroMov || new Date().toLocaleString('es-VE')
    ];

    if (esEdicion) {
      hojaCabecera.getRange(filaExistente, 1, 1, filaCabecera.length).setValues([filaCabecera]);
    } else {
      hojaCabecera.appendRow(filaCabecera);
    }

    // Actualizar líneas en detalleBancosCajas
    var lastRowDet = hojaDetalle.getLastRow();
    if (lastRowDet >= 2) {
      var idsDet = hojaDetalle.getRange(2, 1, lastRowDet - 1, 1).getValues();
      for (var d = idsDet.length - 1; d >= 0; d--) {
        if (String(idsDet[d][0]).trim() === idMov) {
          hojaDetalle.deleteRow(d + 2);
        }
      }
    }

    if (lineas.length > 0) {
      var filasDetalle = lineas.map(function(lin) {
        var mUSD = parseFloat(lin.montoUSD) || 0;
        var mBs = parseFloat(lin.montoBs) || (mUSD * tasa);
        return [
          idMov, fechaMov, data.TipoMov || '', data.NroCuenta || '', data.NombreBanco || '',
          data.CodigoBancos || '', data.Cuenta || '', data.ReferenciaMov || '',
          lin.concepto || data.descMov || '', data.TipoTransaccion || '',
          mUSD, data.MonedaMov || '', tasa, parseFloat(data.IGTF) || 0, mBs,
          lin.codigo || '', lin.nombre || '', lin.cuentaGasto || '', data.NombreCuentaMov || '',
          new Date().toLocaleString('es-VE')
        ];
      });

      hojaDetalle.getRange(hojaDetalle.getLastRow() + 1, 1, filasDetalle.length, filasDetalle[0].length).setValues(filasDetalle);
    } else {
      hojaDetalle.appendRow([
        idMov, fechaMov, data.TipoMov || '', data.NroCuenta || '', data.NombreBanco || '',
        data.CodigoBancos || '', data.Cuenta || '', data.ReferenciaMov || '', data.descMov || '',
        data.TipoTransaccion || '', totalUSD, data.MonedaMov || '', tasa,
        parseFloat(data.IGTF) || 0, totalBs, data.CodigoCuentaContable || '',
        data.CuentaContable || '', data.CuentaGasto || '', data.NombreCuentaMov || '',
        new Date().toLocaleString('es-VE')
      ]);
    }

    // Actualizar saldos bancarios
    try {
      actualizarSaldosBancariosOptimizado(data.idSpreadsheet);
    } catch (e) {
      Logger.log('Error actualizando saldos: ' + e);
    }

    return {
      success: true,
      message: esEdicion ? 'Movimiento actualizado correctamente' : 'Movimiento registrado correctamente',
      id: idMov,
      totalUSD: totalUSD,
      totalBs: totalBs
    };
  } catch (e) {
    Logger.log('Error en agregarEditarMovBancoCaja: ' + e);
    return { success: false, message: 'Error: ' + (e.message || String(e)) };
  }
}
function eliminarMovBancoCaja(id, idSpreadsheetEmpresa) {
  try {
    const ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) throw new Error('No se pudo acceder a la hoja de la empresa activa');

    const sheetCab = ss.getSheetByName('MovBancoCaja');
    const sheetDet = ss.getSheetByName('detalleBancosCajas');
    if (!sheetCab) throw new Error('No se encontró hoja MovBancoCaja');

    const idBuscado = String(id).trim();

    const lastRowCab = sheetCab.getLastRow();
    if (lastRowCab >= 2) {
      const idsCab = sheetCab.getRange(2, 1, lastRowCab - 1, 1).getValues().flat();
      const rowIndex = idsCab.findIndex(c => String(c).trim() === idBuscado);
      if (rowIndex !== -1) {
        sheetCab.deleteRow(rowIndex + 2);
      }
    }

    if (sheetDet) {
      const lastRowDet = sheetDet.getLastRow();
      if (lastRowDet >= 2) {
        const idsDet = sheetDet.getRange(2, 1, lastRowDet - 1, 1).getValues().flat();
        for (let d = idsDet.length - 1; d >= 0; d--) {
          if (String(idsDet[d]).trim() === idBuscado) {
            sheetDet.deleteRow(d + 2);
          }
        }
      }
    }

    try {
      actualizarSaldosBancariosOptimizado(idSpreadsheetEmpresa);
    } catch (e) {
      Logger.log('Error saldos: ' + e);
    }

    return { success: true, message: 'Movimiento eliminado correctamente.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
function obtenerDetallesMovimiento(idMov, idSpreadsheetEmpresa) {
  try {
    var ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) return [];

    var hojaDet = ss.getSheetByName('detalleBancosCajas');
    if (!hojaDet) return [];

    var lastRow = hojaDet.getLastRow();
    if (lastRow < 2) return [];

    var idBuscado = String(idMov).trim();
    var data = hojaDet.getRange(2, 1, lastRow - 1, 20).getValues();
    var detalles = [];

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === idBuscado) {
        detalles.push({
          codigo: String(data[i][15] || ''),
          nombre: String(data[i][16] || ''),
          cuentaGasto: String(data[i][17] || ''),
          concepto: String(data[i][8] || ''),
          montoUSD: parseFloat(data[i][10]) || 0,
          montoBs: parseFloat(data[i][14]) || 0
        });
      }
    }

    return detalles;
  } catch (e) {
    Logger.log('Error en obtenerDetallesMovimiento: ' + e);
    return [];
  }
}
function actualizarSaldosBancariosOptimizado(idSpreadsheetEmpresa) {
  try {
    const ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) throw new Error("No se pudo acceder a la hoja de la empresa activa");

    const hojaDatos = ss.getSheetByName("datosBancarios") || ss.getSheetByName("Bancos");
    const hojaMovimientos = ss.getSheetByName("MovBancoCaja");
    if (!hojaDatos || !hojaMovimientos) throw new Error("Hojas requeridas no encontradas");

    const datosMovimientos = hojaMovimientos.getDataRange().getValues();
    const datosBancarios = hojaDatos.getDataRange().getValues();

    const deltas = { banco: {}, cuenta: {}, caja: {} };

    if (datosMovimientos.length >= 2) {
      const enc = datosMovimientos[0].map(h => String(h).trim());
      const idxMov = {
        codigoBancos: enc.indexOf("CodigoBancos"),
        nroCuenta: enc.indexOf("NroCuenta"),
        tipoMov: enc.indexOf("TipoMov"),
        tipoTransaccion: enc.indexOf("TipoTransaccion"),
        montoMov: enc.indexOf("MontoMovUSD") !== -1 ? enc.indexOf("MontoMovUSD") : enc.indexOf("MontoMov")
      };

      for (let i = 1; i < datosMovimientos.length; i++) {
        const fila = datosMovimientos[i];
        const tipoMov = String(fila[idxMov.tipoMov] || '').trim();
        const tipoTrans = String(fila[idxMov.tipoTransaccion] || '').trim();
        const monto = parseFloat(fila[idxMov.montoMov]) || 0;
        const factor = tipoTrans.includes("Positiva") ? 1 : (tipoTrans.includes("Negativa") ? -1 : 0);
        const ajuste = monto * factor;

        if (tipoMov === "Banco") {
          const codBanco = String(fila[idxMov.codigoBancos] || '').trim();
          if (codBanco) deltas.banco[codBanco] = (deltas.banco[codBanco] || 0) + ajuste;
          const nroCuenta = String(fila[idxMov.nroCuenta] || '').trim();
          if (nroCuenta) deltas.cuenta[nroCuenta] = (deltas.cuenta[nroCuenta] || 0) + ajuste;
        } else if (tipoMov === "Caja") {
          const codCaja = String(fila[idxMov.codigoBancos] || '').trim();
          if (codCaja) deltas.caja[codCaja] = (deltas.caja[codCaja] || 0) + ajuste;
        }
      }
    }

    let modificadas = 0;
    const idx0 = (c) => c - 1;

    for (let i = 1; i < datosBancarios.length; i++) {
      const fila = datosBancarios[i];
      const filaNum = i + 1;

      // Banco (F)
      const codBanco = String(fila[idx0(BANCO_COLS.CODIGO)] || '').trim();
      if (codBanco && deltas.banco.hasOwnProperty(codBanco)) {
        const actual = parseFloat(fila[idx0(BANCO_COLS.SALDO)]) || 0;
        hojaDatos.getRange(filaNum, BANCO_COLS.SALDO).setValue(actual + deltas.banco[codBanco]);
        modificadas++;
        deltas.banco[codBanco] = 0;
      }

      // Cuenta (O)
      const nroCuenta = String(fila[idx0(CUENTA_COLS.NUMERO)] || '').trim();
      if (nroCuenta && deltas.cuenta.hasOwnProperty(nroCuenta)) {
        const actual = parseFloat(fila[idx0(CUENTA_COLS.SALDO)]) || 0;
        hojaDatos.getRange(filaNum, CUENTA_COLS.SALDO).setValue(actual + deltas.cuenta[nroCuenta]);
        modificadas++;
        deltas.cuenta[nroCuenta] = 0;
      }

      // Caja (X)
      const codCaja = String(fila[idx0(CAJA_COLS.CODIGO)] || '').trim();
      if (codCaja && deltas.caja.hasOwnProperty(codCaja)) {
        const actual = parseFloat(fila[idx0(CAJA_COLS.SALDO)]) || 0;
        hojaDatos.getRange(filaNum, CAJA_COLS.SALDO).setValue(actual + deltas.caja[codCaja]);
        modificadas++;
        deltas.caja[codCaja] = 0;
      }
    }

    return { success: true, message: `Saldos actualizados correctamente` };
  } catch (error) {
    console.error("Error actualizarSaldosBancariosOptimizado:", error);
    return { success: false, error: error.toString() };
  }
}
function getMovimientos(limite) {
  limite = limite || 1000;
  try {
    // Resuelve la hoja de la empresa activa sin usar getActiveSpreadsheet()
    const sheet = (typeof getSheetByName === 'function')
      ? getSheetByName('MovBancoCaja')
      : (typeof Contexto !== 'undefined' ? Contexto.getSheet('MovBancoCaja') : null);

    if (!sheet) throw new Error('No se encontró la hoja MovBancoCaja en la empresa activa');
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const startRow = (limite < 1000) ? Math.max(2, lastRow - limite + 1) : 2;
    const numRows = lastRow - startRow + 1;
    const range = sheet.getRange(startRow, 1, numRows, 20);
    const values = range.getValues();
    const movimientos = [];
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      try {
        const formatearNumero = (val) => {
          if (val === "" || val === null || val === undefined) return 0;
          if (typeof val === 'number') return val;
          let num = parseFloat(String(val).replace(/\./g, '').replace(',', '.'));
          return isNaN(num) ? 0 : num;
        };
        
        let fechaISO = '';
        if (row[1] instanceof Date) {
          fechaISO = row[1].toISOString();
        } else if (row[1]) {
          let d = new Date(row[1]);
          fechaISO = isNaN(d.getTime()) ? '' : d.toISOString();
        }
        
        movimientos.push({
          IdMovimiento: String(row[0] || ''),
          FechaMov: fechaISO,
          Cuenta: String(row[6] || ''),
          NroCuenta: String(row[3] || ''),
          NombreBanco: String(row[4] || ''),
          ReferenciaMov: String(row[7] || ''),
          descMov: String(row[8] || ''),
          TipoTransaccion: String(row[9] || ''),
          MontoMovUSD: formatearNumero(row[10]),
          MontoMovBs: formatearNumero(row[14]),
          MonedaMov: String(row[11] || ''),
          tipoMov: String(row[2] || ''),
          codigoCuentaMovInput: row[6] + ' - ' + row[4],
          CodigoBancos: String(row[5] || ''),
          TasaMov: formatearNumero(row[12]),
          IGTF: formatearNumero(row[13]),
          CodigoCuentaContable: String(row[15] || ''),
          CuentaContable: String(row[16] || ''),
          CuentaGasto: String(row[17] || ''),
          NombreCuentaMov: String(row[18] || '')
        });
      } catch (e) {
        console.error('Error procesando fila de movimiento', e);
      }
    }
    console.log(`getMovimientos: ${movimientos.length} registros procesados`);
    return movimientos;
  } catch (error) {
    console.error('Error en getMovimientos:', error);
    return [];
  }
}
function getBancosYCajas(idSpreadsheetEmpresa) {
  try {
    const sheet = getSheetBancos(idSpreadsheetEmpresa);
    if (!sheet) {
      Logger.log('getBancosYCajas: Hoja de bancos no encontrada');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const numCols = Math.min(28, sheet.getLastColumn());
    const totalData = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    const idx = (col) => col - 1;

    const B_COLS = typeof BANCO_COLS !== 'undefined' ? BANCO_COLS : { CODIGO: 2, NOMBRE: 3 };
    const CT_COLS = typeof CUENTA_COLS !== 'undefined' ? CUENTA_COLS : { 
      ID: 1, CODIGO: 2, NUMERO: 3, ID_BANCO_REF: 4, NOM_BANCO: 5, MONEDA: 6, TIPO: 7, COD_CONTABLE: 8, NOM_CONTABLE: 9 
    };
    const CJ_COLS = typeof CAJA_COLS !== 'undefined' ? CAJA_COLS : { 
      ID: 10, CODIGO: 11, NOMBRE: 12, MONEDA: 13, COD_CONTABLE: 14, NOM_CONTABLE: 15 
    };

    const codigosBancosSet = {};
    totalData.forEach((row) => {
      const cod = String(row[idx(B_COLS.CODIGO)] || '').trim();
      if (cod) codigosBancosSet[cod] = true;
    });

    const resultados = [];

    totalData.forEach((row) => {
      // --- Cuentas Bancarias ---
      if (row[idx(CT_COLS.CODIGO)]) {
        const codigoCuenta = String(row[idx(CT_COLS.CODIGO)] || '').trim();
        const codigoBancosCuenta = String(row[idx(CT_COLS.ID_BANCO_REF)] || '').trim();
        const nombreBanco = String(row[idx(CT_COLS.NOM_BANCO)] || '').trim();

        let codigoBancosReal = '';
        if (codigosBancosSet[codigoCuenta]) codigoBancosReal = codigoCuenta;
        if (!codigoBancosReal && codigoBancosCuenta) {
          const soloDigitos = codigoBancosCuenta.replace(/\D/g, '');
          if (soloDigitos.length >= 4) {
            const prefijo4 = soloDigitos.substring(0, 4);
            if (codigosBancosSet[prefijo4]) codigoBancosReal = prefijo4;
          }
        }
        if (!codigoBancosReal && nombreBanco) {
          for (const k in codigosBancosSet) {
            const filaBanco = totalData.find(r => String(r[idx(B_COLS.CODIGO)] || '').trim() === k);
            if (filaBanco && String(filaBanco[idx(B_COLS.NOMBRE)] || '').trim().toUpperCase() === nombreBanco.toUpperCase()) {
              codigoBancosReal = k;
              break;
            }
          }
        }

        resultados.push({
          tipo: 'Banco',
          IdCuenta: String(row[idx(CT_COLS.ID)] || ''),
          CodigoCuenta: codigoCuenta,
          NroCuenta: String(row[idx(CT_COLS.NUMERO)] || ''),
          CodigoBancos: codigoBancosReal,
          CodigoBancosCuenta: codigoBancosCuenta,
          NombreBanco: nombreBanco,
          MonedaBanco: String(row[idx(CT_COLS.MONEDA)] || ''),
          tipoCuenta: String(row[idx(CT_COLS.TIPO)] || ''),
          codContable: String(row[idx(CT_COLS.COD_CONTABLE)] || ''),
          nomContable: String(row[idx(CT_COLS.NOM_CONTABLE)] || '')
        });
      }

      // --- Cajas ---
      if (row[idx(CJ_COLS.CODIGO)]) {
        resultados.push({
          tipo: 'Caja',
          IdCaja: String(row[idx(CJ_COLS.ID)] || ''),
          codigoCaja: String(row[idx(CJ_COLS.CODIGO)] || ''),
          nombreCaja: String(row[idx(CJ_COLS.NOMBRE)] || ''),
          MonedaCaja: String(row[idx(CJ_COLS.MONEDA)] || ''),
          codContable: String(row[idx(CJ_COLS.COD_CONTABLE)] || ''),
          nomContable: String(row[idx(CJ_COLS.NOM_CONTABLE)] || '')
        });
      }
    });

    return resultados;
  } catch (e) {
    Logger.log('Error en getBancosYCajas: ' + e);
    return [];
  }
}
function obtenerCuentasContablesParaSelect_() {
  try {
    // Resuelve de la empresa activa sin getActiveSpreadsheet()
    var hoja = (typeof getSheetByName === 'function')
      ? (getSheetByName('PlandeCuentas') || getSheetByName('PlanDeCuentas') || getSheetByName('Plan de Cuentas') || getSheetByName('PLANDECUENTAS'))
      : (typeof Contexto !== 'undefined' ? (Contexto.getSheet('PlandeCuentas') || Contexto.getSheet('PlanDeCuentas')) : null);

    if (!hoja) { Logger.log('Hoja PlandeCuentas no encontrada'); return []; }
    
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return [];

    var headers = datos[0].map(function(h) {
      return String(h).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    });
    var iCodigo = headers.findIndex(function(h) {
      return h.indexOf('cuenta contable') !== -1 || h === 'codigo';
    });
    var iNombre = headers.findIndex(function(h) {
      return h === 'cuenta' || h.indexOf('nombre') !== -1;
    });
    if (iCodigo === -1) iCodigo = 1; // B
    if (iNombre === -1) iNombre = 2; // C

    var res = [];
    for (var i = 1; i < datos.length; i++) {
      var cod = String(datos[i][iCodigo] || '').trim();
      var nom = String(datos[i][iNombre] || '').trim();
      if (cod) res.push({ codigo: cod, nombre: nom });
    }
    return res;
  } catch (e) { Logger.log('Error obtenerCuentasContablesParaSelect_: ' + e); return []; }
}
function getDatosBancarios() {
  try {
    const sheet = getSheetBancos();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const resultados = [];
    const totalData = sheet.getRange(2, 1, lastRow - 1, 26).getValues();

    totalData.forEach(row => {
      // 1. Procesar Cuentas Bancarias (Convertir índices de base 1 a base 0 de la fila)
      const idxCta = (col) => col - 1;
      if (row[idxCta(CUENTA_COLS.NUMERO)]) { 
        resultados.push({
          tipo: 'Banco',
          nombre: String(row[idxCta(CUENTA_COLS.NOM_BANCO)]),
          nroCuenta: String(row[idxCta(CUENTA_COLS.NUMERO)]),
          codigo: String(row[idxCta(CUENTA_COLS.CODIGO)]),
          codigoContable: String(row[idxCta(CUENTA_COLS.COD_CONTABLE)])
        });
      }

      // 2. Procesar Cajas
      const idxCaj = (col) => col - 1;
      if (row[idxCaj(CAJA_COLS.CODIGO)]) {
        resultados.push({
          tipo: 'Caja',
          nombre: String(row[idxCaj(CAJA_COLS.NOMBRE)]),
          nroCuenta: '',
          codigo: String(row[idxCaj(CAJA_COLS.CODIGO)]),
          codigoContable: String(row[idxCaj(CAJA_COLS.COD_CONTABLE)])
        });
      }
    });

    return resultados;
  } catch (e) {
    Logger.log('Error en getDatosBancarios: ' + e.toString());
    return [];
  }
}
function getNextAvailableRow(sheet, type) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  
  let colIndex;
  switch(type.toLowerCase()) {
    case 'banco': colIndex = BANCO_COLS.ID; break;
    case 'cuenta': colIndex = CUENTA_COLS.ID; break;
    case 'caja': colIndex = CAJA_COLS.ID; break;
    default: return lastRow + 1;
  }
  
  const values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0];
    if (!cellValue || cellValue === '') {
      return i + 2;
    }
  }
  return lastRow + 1;
}
function obtenerSaldoCuenta(nroCuenta) {
  try {
    const sheet = getSheetBancos();
    const datos = sheet.getDataRange().getValues();
    const idx = (col) => col - 1;
    
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][idx(CUENTA_COLS.NUMERO)]) === String(nroCuenta)) {
        return {
          success: true,
          nroCuenta: nroCuenta,
          saldo: datos[i][idx(CUENTA_COLS.SALDO)]
        };
      }
    }
    return { success: false, message: "Cuenta no encontrada" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
function obtenerSaldoCaja(codigoCaja) {
  try {
    const sheet = getSheetBancos();
    const datos = sheet.getDataRange().getValues();
    const idx = (col) => col - 1;
    
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][idx(CAJA_COLS.CODIGO)]) === String(codigoCaja)) {
        return {
          success: true,
          codigoCaja: codigoCaja,
          saldo: datos[i][idx(CAJA_COLS.SALDO)]
        };
      }
    }
    return { success: false, message: "Caja no encontrada" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
function findNextAvailableRow(sheet, colNum) {
  const columnValues = sheet.getRange(2, colNum, sheet.getLastRow()).getValues();
  let i = 0;
  while (i < columnValues.length && columnValues[i][0] !== "") {
    i++;
  }
  return i + 2;
}
function populateStaticDropdownsMov() {
    populateDropdown('cuentaContableMov', allData.listaCuentasContables.map(c => `${c.codigo} - ${c.nombre}`));
    populateDropdown('monedaMov', allData.monedas.map(m => m.nombre));
}
function actualizarSaldosBancarios() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDatos = ss.getSheetByName("datosBancarios");
    const hojaMovimientos = ss.getSheetByName("MovBancoCaja");
    
    if (!hojaDatos || !hojaMovimientos) {
      throw new Error("No se encontraron las hojas necesarias");
    }
    
    // Obtener datos de movimientos
    const datosMovimientos = hojaMovimientos.getDataRange().getValues();
    const encabezadosMov = datosMovimientos[0];
    
    // Mapear índices de columnas en MovBancoCaja
    const idxMov = {
      codigoBancos: encabezadosMov.indexOf("CodigoBancos"),
      nroCuenta: encabezadosMov.indexOf("NroCuenta"),
      tipoMov: encabezadosMov.indexOf("TipoMov"),
      tipoTransaccion: encabezadosMov.indexOf("TipoTransaccion"),
      montoMov: encabezadosMov.indexOf("MontoMov"),
      monedaMov: encabezadosMov.indexOf("MonedaMov")
    };
    
    // Validar que existan las columnas necesarias
    if (idxMov.nroCuenta === -1) {
      throw new Error("No se encontró la columna NroCuenta en MovBancoCaja");
    }
    
    // Obtener datos de datosBancarios
    const datosBancarios = hojaDatos.getDataRange().getValues();
    const encabezados = datosBancarios[0];
    
    // Mapear índices de columnas en datosBancarios
    const idx = {
      // Bancos (columnas A-F)
      codigoBancos: encabezados.indexOf("CodigoBancos"),
      saldoCuenta: encabezados.indexOf("SaldoCuenta"), // Saldo del Banco
      
      // Cuentas Bancarias (columnas H-Q)
      nroCuenta: encabezados.indexOf("NroCuenta"),
      codigoBancosCuenta: encabezados.indexOf("CodigoBancosCuenta"),
      saldoBanco: encabezados.indexOf("SaldoBanco"), // Saldo de la Cuenta
      
      // Cajas (columnas S-Z)
      codigoCaja: encabezados.indexOf("CodigoCaja"),
      saldoCaja: encabezados.indexOf("SaldoCaja") // Saldo de la Caja
    };
    
    // Inicializar acumuladores
    const saldosPorBanco = {};      // Key: CodigoBancos
    const saldosPorCuenta = {};     // Key: NroCuenta
    const saldosPorCaja = {};       // Key: CodigoCaja
    
    // Procesar movimientos (desde fila 1, saltando encabezado)
    for (let i = 1; i < datosMovimientos.length; i++) {
      const fila = datosMovimientos[i];
      
      const codigoBanco = fila[idxMov.codigoBancos];
      const nroCuenta = fila[idxMov.nroCuenta];
      const tipoMov = fila[idxMov.tipoMov];
      const tipoTransaccion = fila[idxMov.tipoTransaccion];
      const monto = parseFloat(fila[idxMov.montoMov]) || 0;
      
      // Determinar signo del movimiento
      let factor = 0;
      if (tipoTransaccion && tipoTransaccion.toString().includes("Positiva")) {
        factor = 1; // Ingreso (+)
      } else if (tipoTransaccion && tipoTransaccion.toString().includes("Negativa")) {
        factor = -1; // Egreso (-)
      }
      
      const montoAjustado = monto * factor;
      
      // Acumular por BANCO (si es tipo Banco)
      if (tipoMov === "Banco" && codigoBanco) {
        saldosPorBanco[codigoBanco] = (saldosPorBanco[codigoBanco] || 0) + montoAjustado;
      }
      
      // Acumular por CUENTA BANCARIA (si tiene número de cuenta)
      if (tipoMov === "Banco" && nroCuenta && nroCuenta.toString().trim() !== "") {
        saldosPorCuenta[nroCuenta] = (saldosPorCuenta[nroCuenta] || 0) + montoAjustado;
      }
      
      // Acumular por CAJA (si es tipo Caja)
      if (tipoMov === "Caja" && codigoBanco) {
        // Para cajas, usamos el CodigoBancos que en realidad es el CodigoCaja
        saldosPorCaja[codigoBanco] = (saldosPorCaja[codigoBanco] || 0) + montoAjustado;
      }
    }
    
    // Actualizar hoja datosBancarios
    const rangosAActualizar = [];
    
    for (let i = 1; i < datosBancarios.length; i++) {
      const fila = datosBancarios[i];
      let filaModificada = false;
      
      // 1. Actualizar SALDO DE BANCO (por CodigoBancos)
      const codigoBanco = fila[idx.codigoBancos];
      if (codigoBanco && saldosPorBanco.hasOwnProperty(codigoBanco)) {
        // La columna SaldoCuenta está en la sección de Bancos
        const nuevoSaldo = saldosPorBanco[codigoBanco];
        if (fila[idx.saldoCuenta] !== nuevoSaldo) {
          hojaDatos.getRange(i + 1, idx.saldoCuenta + 1).setValue(nuevoSaldo);
          filaModificada = true;
        }
      }
      
      // 2. Actualizar SALDO DE CUENTA BANCARIA (por NroCuenta)
      const nroCuenta = fila[idx.nroCuenta];
      if (nroCuenta && nroCuenta.toString().trim() !== "" && saldosPorCuenta.hasOwnProperty(nroCuenta)) {
        const nuevoSaldoCuenta = saldosPorCuenta[nroCuenta];
        if (fila[idx.saldoBanco] !== nuevoSaldoCuenta) {
          hojaDatos.getRange(i + 1, idx.saldoBanco + 1).setValue(nuevoSaldoCuenta);
          filaModificada = true;
        }
      }
      
      // 3. Actualizar SALDO DE CAJA (por CodigoCaja)
      const codigoCaja = fila[idx.codigoCaja];
      if (codigoCaja && saldosPorCaja.hasOwnProperty(codigoCaja)) {
        const nuevoSaldoCaja = saldosPorCaja[codigoCaja];
        if (fila[idx.saldoCaja] !== nuevoSaldoCaja) {
          hojaDatos.getRange(i + 1, idx.saldoCaja + 1).setValue(nuevoSaldoCaja);
          filaModificada = true;
        }
      }
    }
    
    // Registrar log para depuración
    console.log("Saldos por Banco:", saldosPorBanco);
    console.log("Saldos por Cuenta:", saldosPorCuenta);
    console.log("Saldos por Caja:", saldosPorCaja);
    
    return {
      success: true,
      message: "Saldos actualizados correctamente",
      detalles: {
        bancosActualizados: Object.keys(saldosPorBanco).length,
        cuentasActualizadas: Object.keys(saldosPorCuenta).length,
        cajasActualizadas: Object.keys(saldosPorCaja).length,
        saldosBanco: saldosPorBanco,
        saldosCuenta: saldosPorCuenta,
        saldosCaja: saldosPorCaja
      }
    };
    
  } catch (error) {
    console.error("Error en actualizarSaldosBancarios:", error);
    return {
      success: false,
      error: error.toString(),
      message: "Error al actualizar saldos: " + error.message
    };
  }
}
function procesarFecha(fechaInput) {
  if (!fechaInput) return '';
  
  if (fechaInput instanceof Date && !isNaN(fechaInput)) {
    return fechaInput;
  }
  
  const dateString = String(fechaInput);
  const partes = dateString.split('/');
  
  if (partes.length === 3) {
    const dia = parseInt(partes[0], 10);
    const mes = parseInt(partes[1], 10) - 1;
    const anio = parseInt(partes[2], 10);
    
    const fechaFormateada = new Date(anio, mes, dia);
    
    if (!isNaN(fechaFormateada.getTime()) && 
        fechaFormateada.getDate() === dia && 
        fechaFormateada.getMonth() === mes && 
        fechaFormateada.getFullYear() === anio) {
      return fechaFormateada;
    }
  }
  
  console.log('No se pudo procesar la fecha:', fechaInput, 'Formato esperado: DD/MM/YYYY');
  return '';
}
function obtenerSiguienteCorrelativoCobrosPagos(prefijo, skipLock, idSpreadsheetEmpresa) {
  if (!prefijo) prefijo = 'A';
  prefijo = String(prefijo).trim().toUpperCase();
  if (['C', 'P', 'A'].indexOf(prefijo) === -1) {
    throw new Error('Prefijo de correlativo inválido: ' + prefijo + ' (esperado C, P o A).');
  }

  var lock = null;
  if (!skipLock) {
    lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (e) {
      Logger.log('[CORRELATIVO CP] Lock ya tomado por caller: ' + e.message);
      lock = null;
    }
  }

  try {
    var ss = _SS_BC_(idSpreadsheetEmpresa);
    if (!ss) return prefijo + '0000001';

    var shCP = ss.getSheetByName(CFG_CXC.HOJA_COBROS_PAGOS || 'CobrosPagos');
    if (!shCP) return prefijo + '0000001';

    var lastRow = shCP.getLastRow();
    if (lastRow < 2) return prefijo + '0000001';

    var headers = shCP.getRange(1, 1, 1, shCP.getLastColumn()).getValues()[0];
    var colId = -1;
    for (var h = 0; h < headers.length; h++) {
      var hn = String(headers[h] || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (hn === 'id') { colId = h; break; }
    }
    if (colId === -1) colId = 9; // Fallback columna J

    var valoresId = shCP.getRange(2, colId + 1, lastRow - 1, 1).getValues();
    var maxNum = 0;
    var regex = new RegExp('^' + prefijo + '(\\d+)$', 'i');

    for (var i = 0; i < valoresId.length; i++) {
      var val = String(valoresId[i][0] || '').trim();
      if (!val) continue;

      var match = val.match(regex);
      if (match && match[1]) {
        var num = parseInt(match[1], 10);
        if (isFinite(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }

    var siguiente = maxNum + 1;
    return prefijo + String(siguiente).padStart(7, '0');

  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}
function buscarRespaldoEnComprasVentas(ss, idBusqueda, idNorm, tipoFinal) {
  var hojas = ["Compras", "Ventas"];
  var encontrados = [];
  
  for (var h = 0; h < hojas.length; h++) {
    var hoja = ss.getSheetByName(hojas[h]);
    if (!hoja) continue;
    var data = hoja.getDataRange().getValues();
    if (data.length <= 1) continue;

    var headers = data[0].map(function(head) {
      return String(head).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s_\-#.]/g, "");
    });

    var cFact = headers.findIndex(function(x) { return x.includes("factura"); });
    var cComp = headers.findIndex(function(x) { return x.includes("comprobante"); });
    var cRet = headers.findIndex(function(x) { return x === (tipoFinal === 'IVA' ? "retencioniva" : "retencionislr"); });
    var cTer = headers.findIndex(function(x) { return x.includes("proveedor") || x.includes("cliente") || x.includes("tercero"); });
    var cRif = headers.findIndex(function(x) { return x.includes("rif"); });
    var cFecha = headers.findIndex(function(x) { return x.includes("fecha"); });

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var valFact = cFact !== -1 ? String(row[cFact]).trim().toLowerCase() : "";
      var valComp = cComp !== -1 ? String(row[cComp]).trim().toLowerCase() : "";

      if ((valComp && valComp.indexOf(idBusqueda) !== -1) || (valFact && (valFact === idBusqueda || valFact.replace(/^0+/, '') === idNorm))) {
        var montoRet = cRet !== -1 ? (parseFloat(row[cRet]) || 0) : 0;
        if (montoRet > 0) {
          encontrados.push({
            nroFactura: row[cFact],
            nroControl: "00-00000000",
            fechaFactura: row[cFecha] || new Date(),
            totalFactura: 0,
            baseImponible: 0,
            montoIva: 0,
            porcentaje: (tipoFinal === 'IVA' ? 75 : 2),
            montoRetenidoBs: montoRet,
            montoRetenidoUSD: 0,
            tercero: cTer !== -1 ? row[cTer] : "",
            rif: cRif !== -1 ? row[cRif] : "",
            comprobante: cComp !== -1 && row[cComp] ? row[cComp] : idBusqueda,
            fecha: row[cFecha],
            tasa: 1
          });
        }
      }
    }
    if (encontrados.length > 0) break;
  }
  return encontrados;
}
function obtenerDatosEmpresa() {
  return {
    nombre: typeof NOM_EMPRESA !== 'undefined' ? NOM_EMPRESA : 'EMPRESA C.A.',
    rif: typeof RIF_EMPRESA !== 'undefined' ? RIF_EMPRESA : 'J-00000000-0',
    direccion: typeof DIR_EMPRESA !== 'undefined' ? DIR_EMPRESA : '',
    logo: typeof LOGO_EMPRESA !== 'undefined' ? LOGO_EMPRESA : '',
    telefono: ''
  };
}
function numeroALetras(cantidad) {
  var numero = parseFloat(cantidad) || 0;
  var enteros = Math.floor(numero);
  var centavos = Math.round((numero - enteros) * 100);
  
  function unidades(num) {
    var u = ['', 'UN', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE'];
    return u[num] || '';
  }
  function decenas(num) {
    var d = ['', 'DIEZ', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA'];
    var especiales = { 11:'ONCE', 12:'DOCE', 13:'TRECE', 14:'CATORCE', 15:'QUINCE', 16:'DIECISEIS', 17:'DIECISIETE', 18:'DIECIOCHO', 19:'DIECINUEVE', 21:'VEINTIUNO', 22:'VEINTIDOS', 23:'VEINTITRES', 24:'VEINTICUATRO', 25:'VEINTICINCO', 26:'VEINTISEIS', 27:'VEINTISIETE', 28:'VEINTIOCHO', 29:'VEINTINUEVE' };
    if (especiales[num]) return especiales[num];
    var dec = Math.floor(num / 10);
    var uni = num % 10;
    if (dec === 0) return unidades(uni);
    if (uni === 0) return d[dec];
    return d[dec] + ' Y ' + unidades(uni);
  }
  function centenas(num) {
    var c = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS'];
    if (num === 100) return 'CIEN';
    var cen = Math.floor(num / 100);
    var res = num % 100;
    return (c[cen] ? c[cen] + ' ' : '') + decenas(res);
  }
  function seccion(num, divisor, strSingular, strPlural) {
    var c = Math.floor(num / divisor);
    var r = num - (c * divisor);
    var letras = '';
    if (c > 0) {
      if (c === 1 && strSingular === 'UN MILLON') letras = 'UN MILLON ';
      else if (c > 1) letras = centenas(c) + ' ' + strPlural + ' ';
      else letras = centenas(c) + ' ' + strSingular + ' ';
    }
    return { letras: letras, resto: r };
  }
  
  if (enteros === 0) return 'CERO BOLÍVARES CON ' + String(centavos).padStart(2, '0') + '/100';
  
  var mill = seccion(enteros, 1000000, 'UN MILLON', 'MILLONES');
  var miles = seccion(mill.resto, 1000, 'MIL', 'MIL');
  var cents = centenas(miles.resto);
  
  var resultado = (mill.letras + miles.letras + cents).trim() + ' BOLÍVARES CON ' + String(centavos).padStart(2, '0') + '/100';
  return '**** ' + resultado.replace(/\s+/g, ' ') + ' ****';
}




// =============================================
/** COMPRAS                    */
// =============================================
function _resolverSSCompras_(idSpreadsheet) {
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
function guardarDatosFacturaCompras(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!payload) throw new Error("No se recibieron datos de la compra.");
    
    // ✅ FIX 1: Validar que el proveedor y la factura no lleguen vacíos
    var nombreProv = String(payload.NombreProveedor || '').trim();
    var nroFact = String(payload.NroFacturaProveedor || '').trim();
    var tasa = parseFloat(payload.TasaProveedor) || 0;

    if (!nombreProv) throw new Error("El nombre del proveedor está vacío. Seleccione uno válido.");
    if (!nroFact) throw new Error("El número de factura está vacío.");
    if (tasa <= 0) throw new Error("La tasa de cambio es inválida (debe ser mayor a 0).");

    var idSpreadsheet = payload.idSpreadsheet || payload.empresaId || payload.idEmpresa;
    var ss = _resolverSSCompras_(idSpreadsheet);
    if (!ss) throw new Error("No se pudo conectar a la base de datos de la empresa activa.");

    var sheetCompras = ss.getSheetByName("Compras");
    var sheetDetalle = ss.getSheetByName("DetalleFacturas");
    var sheetCobros = ss.getSheetByName("CobrosPagos");
    var sheetProductos = ss.getSheetByName("Productos");

    if (!sheetCompras) throw new Error("No se encontró la hoja 'Compras'.");
    if (!sheetDetalle) throw new Error("No se encontró la hoja 'DetalleFacturas'.");

    var ahora = new Date();
    var fechaStr = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var tipoDoc = String(payload.TipoProveedor || 'Factura').trim();

    var correlativoComp = ("000" + sheetCompras.getLastRow()).slice(-4);
    var idCompra = "COM-" + correlativoComp;

    // 1. REGISTRO EN LA HOJA 'Compras' (31 Columnas)
    var filaCompra = [
      nombreProv,                                  // 1: NombreProveedor
      payload.RifProveedor || "",                  // 2: RifProveedor
      payload.FechaFacturaProveedor || fechaStr,   // 3: FechaFacturaProveedor
      tipoDoc,                                     // 4: TipoProveedor
      nroFact,                                     // 5: NroFacturaProveedor
      payload.NroControlProveedor || "",           // 6: NroControlProveedor
      payload.MonedaProveedor || "Dolares",        // 7: MonedaProveedor
      tasa,                                        // 8: TasaProveedor
      parseFloat(payload.SubTotalProveedor) || 0,  // 9: SubTotalProveedor
      parseFloat(payload.totalExentoCompra) || 0,  // 10: totalExentoCompra
      parseFloat(payload.IvaProveedor) || 0,       // 11: IvaProveedor
      parseFloat(payload.TotalProveedor) || 0,     // 12: TotalProveedor
      parseFloat(payload.SaldoProveedor) || 0,     // 13: SaldoProveedor
      parseFloat(payload.SubTotalProveedorDivisa) || 0, // 14
      parseFloat(payload.IVAProveedorDivisa) || 0, // 15
      parseFloat(payload.TotalProveedorDivisa) || 0, // 16
      parseFloat(payload.SaldoProveedorDivisa) || 0, // 17
      "Compra",                                    // 18
      payload.NroFacturaAfectada || "",            // 19
      idCompra,                                    // 20
      payload.cobradaPagada || "porPagar",         // 21
      0,                                           // 22: Abono
      "",                                          // 23: fechaAbono
      "X",                                         // 24: Contabilidad
      "",                                          // 25: nroAsiento
      payload.descripcionCompra || ("Compra N° " + nroFact), // 26
      "Activa",                                    // 27: Status
      payload.FechaFacturaProveedor || fechaStr,   // 28: fechaRetencion
      parseFloat(payload.RetencionIva) || 0,       // 29: RetencionIva
      parseFloat(payload.RetencionIslr) || 0,      // 30: RetencionISLR
      payload.NroComprobanteRetencion || ""        // 31: nroComprobantes
    ];
    sheetCompras.appendRow(filaCompra);

    // 2. REGISTRO EN 'DetalleFacturas' (19 Columnas)
    var items = Array.isArray(payload.items) ? payload.items : [];
    if (items.length > 0) {
      var filasDet = [];
      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        filasDet.push([
          nroFact,                              // 1: NroFactura
          it.Producto || it.nombre || "",       // 2: Producto
          it.Categoria || "",                   // 3: Categoria
          parseFloat(it.Cantidad) || 0,         // 4: Cantidad
          parseFloat(it.Precio) || 0,           // 5: Precio
          parseFloat(it.PorcentajeIVA) || 0,    // 6: %IVA
          parseFloat(it.MontoIVA) || 0,         // 7: MontoIVA
          parseFloat(it.PrecioTotal) || 0,      // 8: PrecioTotal
          parseFloat(it.PrecioUSD) || 0,        // 9: PrecioUSD
          parseFloat(it.PorcentajeIVAUSD) || 0, // 10: %IVAUSD
          parseFloat(it.MontoIVAUSD) || 0,      // 11: MontoIVAUSD
          parseFloat(it.PrecioTotalUSD) || 0,   // 12: PrecioTotalUSD
          "Compra",                             // 13: Transaccion
          it.CuentaContable || "",              // 14: CuentaContable
          it.Cuenta || "",                      // 15: NombreContable
          payload.descripcionCompra || "",      // 16: DescripcionTransaccion
          payload.usuario || "Admin",           // 17: usuario
          payload.FechaFacturaProveedor || fechaStr, // 18: fecha
          tipoDoc                               // 19: tipo
        ]);
      }
      sheetDetalle.getRange(sheetDetalle.getLastRow() + 1, 1, filasDet.length, 19).setValues(filasDet);
    }

    // 3. REGISTRO EN 'CobrosPagos' (Cuentas por Pagar)
    if (sheetCobros) {
      var correlativoCP = ("000" + sheetCobros.getLastRow()).slice(-4);
      var filaCP = [
        nombreProv,                                // 1: ProveedorCliente
        payload.RifProveedor || "",                // 2: Rif
        nroFact,                                   // 3: NroFactura
        payload.FechaFacturaProveedor || fechaStr, // 4: FechaFactura
        parseFloat(payload.TotalProveedor) || 0,   // 5: TotalFactura
        parseFloat(payload.SaldoProveedor) || 0,   // 6: SaldoFactura
        parseFloat(payload.SaldoProveedorDivisa) || 0, // 7: SaldoFacturaUSD
        payload.cobradaPagada || "porPagar",       // 8: cobradaPagada
        "Compra",                                  // 9: Transaccion
        "CXP-" + correlativoCP,                    // 10: id
        0,                                         // 11: Abono
        0,                                         // 12: AbonoBs
        payload.MonedaProveedor || "Dolares",      // 13: Moneda
        tasa,                                      // 14: Tasa
        payload.descripcionCompra || ("Compra N° " + nroFact), // 15: Descripcion
        "",                                        // 16: fechaAbono
        "",                                        // 17: TipoMovimiento
        "",                                        // 18: CuentaDestino
        "",                                        // 19: NroCuentaDestino
        "",                                        // 20: BancoDestino
        "",                                        // 21: fechaVencimiento
        payload.NroComprobanteRetencion || "",     // 22: nroComprobantesRetIVA
        payload.NroComprobanteRetencionISLR || ""  // 23: nroComprobantesRetISLR
      ];
      sheetCobros.appendRow(filaCP);
    }

    // 4. ACTUALIZACIÓN DE STOCK
    if (sheetProductos && items.length > 0) {
      var pData = sheetProductos.getDataRange().getValues();
      for (var k = 0; k < items.length; k++) {
        var itId = String(items[k].IdProducto || items[k].codigo || '').trim().toUpperCase();
        var itCant = parseFloat(items[k].Cantidad) || 0;
        if (!itId || itCant <= 0) continue;

        for (var r = 1; r < pData.length; r++) {
          if (String(pData[r][0] || '').trim().toUpperCase() === itId) {
            var stActual = parseFloat(pData[r][5]) || 0;
            sheetProductos.getRange(r + 1, 6).setValue(stActual + itCant);
            break;
          }
        }
      }
    }

    // 5. GENERAR ASIENTO CONTABLE
    var nroAsientoGenerado = "";
    try {
      var asientoRes = generarAsientoCompra({
        FechaFacturaProveedor: payload.FechaFacturaProveedor || fechaStr,
        NombreProveedor: nombreProv,
        NroFacturaProveedor: nroFact,
        tipoTransaccion: "Compra",
        TotalProveedor: payload.TotalProveedor,
        IvaProveedor: payload.IvaProveedor,
        totalExentoCompra: payload.totalExentoCompra,
        RetencionIva: payload.RetencionIva,
        RetencionIslr: payload.RetencionIslr
      });
      nroAsientoGenerado = asientoRes;
      
      // Opcional: Actualizar la columna nroAsiento en la hoja Compras
      var lastRowCompras = sheetCompras.getLastRow();
      var headersCompras = sheetCompras.getRange(1, 1, 1, sheetCompras.getLastColumn()).getValues()[0];
      var colNroAsiento = headersCompras.indexOf("nroAsiento") + 1;
      if (colNroAsiento > 0) {
        sheetCompras.getRange(lastRowCompras, colNroAsiento).setValue(nroAsientoGenerado);
      }
    } catch (eAsiento) {
      Logger.log("⚠️ Advertencia: No se pudo generar el asiento, pero la compra se guardó. " + eAsiento);
    }

    return {
      status: "success",
      idCompra: idCompra,
      nroAsiento: nroAsientoGenerado, // ✅ AHORA DEVOLVEMOS EL ASIENTO
      message: "Compra " + nroFact + " registrada exitosamente."
    };

  } catch(err) {
    Logger.log("Error en guardarDatosFacturaCompras: " + err);
    return { status: "error", message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
function obtenerFacturaCompra(nroFacturaBusqueda) {
  try {
    var idSpreadsheet = '';
    try {
      var emp = JSON.parse(PropertiesService.getUserProperties().getProperty('EMPRESA_ACTIVA_ID') || '{}');
      idSpreadsheet = emp.idSpreadsheet || PropertiesService.getUserProperties().getProperty('EMPRESA_ACTIVA_ID');
    } catch(e) {}

    var ss = _resolverSSCompras_(idSpreadsheet);
    var sheetCompras = ss.getSheetByName('Compras');
    var sheetDetalle = ss.getSheetByName('DetalleFacturas');
    
    if (!sheetCompras || !sheetDetalle) return { status: 'error', message: 'Hojas no encontradas.' };

    var dataCompras = sheetCompras.getDataRange().getValues();
    var headersCompras = dataCompras[0].map(function(h) { return String(h).trim(); });

    var idx = {
      proveedor: headersCompras.indexOf('NombreProveedor'),
      rif: headersCompras.indexOf('RifProveedor'),
      fecha: headersCompras.indexOf('FechaFacturaProveedor'),
      tipoDoc: headersCompras.indexOf('TipoProveedor'),
      factura: headersCompras.indexOf('NroFacturaProveedor'),
      control: headersCompras.indexOf('NroControlProveedor'),
      moneda: headersCompras.indexOf('MonedaProveedor'),
      tasa: headersCompras.indexOf('TasaProveedor'),
      descripcion: headersCompras.indexOf('descripcionCompra'),
      afectada: headersCompras.indexOf('NroFacturaAfectada'),
      retencion: headersCompras.indexOf('RetencionIva'),
      comprobante: headersCompras.indexOf('nroComprobantes'),
      retencionIslr: headersCompras.indexOf('RetencionISLR'),
      comprobanteIslr: headersCompras.indexOf('nroComprobantesISLR') // Si tienes esta columna
    };

    var facturaEncontrada = null;
    var busquedaNorm = String(nroFacturaBusqueda).trim().replace(/^0+/, '');

    // Buscar de abajo hacia arriba (la más reciente primero)
    for (var i = dataCompras.length - 1; i > 0; i--) {
      var facturaNorm = String(dataCompras[i][idx.factura] || "").trim().replace(/^0+/, '');
      if (facturaNorm === busquedaNorm) {
        facturaEncontrada = dataCompras[i];
        break;
      }
    }

    if (!facturaEncontrada) return { status: 'error', message: 'No se encontró la factura.' };

    // Obtener Detalle
    var dataDetalle = sheetDetalle.getDataRange().getValues();
    var headersDetalle = dataDetalle[0].map(function(h) { return String(h).trim(); });
    var idxDet = {
      factura: headersDetalle.indexOf('NroFactura'),
      producto: headersDetalle.indexOf('Producto'),
      cantidad: headersDetalle.indexOf('Cantidad'),
      precioUSD: headersDetalle.indexOf('PrecioUSD'),
      iva: headersDetalle.indexOf('%IVA'),
      cuentaContable: headersDetalle.indexOf('CuentaContable'),
      transaccion: headersDetalle.indexOf('Transaccion')
    };

    var itemsFactura = [];
    for (var j = 1; j < dataDetalle.length; j++) {
      var row = dataDetalle[j];
      var factDetNorm = String(row[idxDet.factura] || "").trim().replace(/^0+/, '');
      
      if (factDetNorm === busquedaNorm && String(row[idxDet.transaccion]).trim().toLowerCase() === 'compra') {
        itemsFactura.push({
          codigo: 'N/A', // Si tienes un campo de código en detalle, cámbialo aquí
          nombre: row[idxDet.producto],
          cantidad: parseFloat(row[idxDet.cantidad]) || 0,
          precioUSD: parseFloat(row[idxDet.precioUSD]) || 0,
          tasaIva: parseFloat(row[idxDet.iva]) || 0,
          codcontable: row[idxDet.cuentaContable] || ''
        });
      }
    }

    // Formatear fecha
    var fechaRaw = facturaEncontrada[idx.fecha];
    var fechaFormateada = "";
    if (fechaRaw instanceof Date) {
      var tzDate = new Date(fechaRaw.getTime() - (fechaRaw.getTimezoneOffset() * 60000));
      fechaFormateada = tzDate.toISOString().split('T')[0];
    } else if (typeof fechaRaw === 'string' && fechaRaw.includes('/')) {
      var p = fechaRaw.split('/');
      if (p.length === 3) fechaFormateada = p[2] + '-' + p[1].padStart(2,'0') + '-' + p[0].padStart(2,'0');
    } else if (fechaRaw) {
      fechaFormateada = String(fechaRaw);
    }

    var responseData = {
      nroFactura: String(facturaEncontrada[idx.factura] || ''),
      nroControl: String(facturaEncontrada[idx.control] || ''), 
      proveedor: String(facturaEncontrada[idx.proveedor] || ''),
      rif: String(facturaEncontrada[idx.rif] || ''),
      moneda: String(facturaEncontrada[idx.moneda] || "Dolares"),
      tasa: parseFloat(facturaEncontrada[idx.tasa]) || 1,
      tipoDoc: String(facturaEncontrada[idx.tipoDoc] || "Factura").trim(),
      fechaEmision: fechaFormateada,
      fechaRegistro: fechaFormateada,
      descripcion: String(facturaEncontrada[idx.descripcion] || ""),
      nroFacturaAfectada: idx.afectada !== -1 ? String(facturaEncontrada[idx.afectada] || '') : "",
      retencionIva: idx.retencion !== -1 ? parseFloat(facturaEncontrada[idx.retencion]) || 0 : 0,
      retencionIslr: idx.retencionIslr !== -1 ? parseFloat(facturaEncontrada[idx.retencionIslr]) || 0 : 0,
      nroComprobanteRet: idx.comprobante !== -1 ? String(facturaEncontrada[idx.comprobante] || '') : "",
      nroComprobanteRetISLR: idx.comprobanteIslr !== -1 ? String(facturaEncontrada[idx.comprobanteIslr] || '') : "",
      items: itemsFactura
    };

    return { status: 'success', data: responseData };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}
function buscarFacturaCompra(hoja, idFactura) {
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return null;

  const headers = data[0].map(h => String(h).trim());
  const colId = headers.indexOf("Id");
  const colNroFactura = headers.indexOf("NroFacturaProveedor");
  
  const searchStr = String(idFactura || "").trim();
  const searchNum = parseInt(searchStr, 10);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = colId !== -1 ? String(row[colId] || "").trim() : "";
    const valNro = colNroFactura !== -1 ? row[colNroFactura] : "";
    const rowNroStr = String(valNro || "").trim();
    const rowNroNum = parseInt(rowNroStr, 10);

    const matchId = (rowId.toLowerCase() === searchStr.toLowerCase());
    const matchNroTexto = (rowNroStr === searchStr);
    const matchNroNum = (!isNaN(searchNum) && !isNaN(rowNroNum) && searchNum === rowNroNum);

    if (matchId || matchNroTexto || matchNroNum) {
      return {
        rowIndex: i + 1,
        Id: rowId,
        NroFacturaProveedor: !isNaN(rowNroNum) ? String(rowNroNum).padStart(6, '0') : rowNroStr,
        NombreProveedor: row[headers.indexOf("NombreProveedor")] || "",
        RifProveedor: row[headers.indexOf("RifProveedor")] || "",
        FechaFacturaProveedor: row[headers.indexOf("FechaFacturaProveedor")] || "",
        TotalProveedor: Number(row[headers.indexOf("TotalProveedor")]) || 0,
        TotalProveedorDivisa: Number(row[headers.indexOf("TotalProveedorDivisa")]) || 0,
        SaldoProveedor: Number(row[headers.indexOf("SaldoProveedor")]) || 0,
        SaldoProveedorDivisa: Number(row[headers.indexOf("SaldoProveedorDivisa")]) || 0
      };
    }
  }
  
  return null;
}
function obtenerSiguienteCorrelativoCompra(sheetCompras) {
  const lastRow = sheetCompras.getLastRow();
  if (lastRow <= 1) return "C0000001";

  const lastCol = sheetCompras.getLastColumn();
  const headers = sheetCompras.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

  let colId = headers.indexOf('Id');
  if (colId === -1) colId = headers.indexOf('ID');
  if (colId === -1) throw new Error("No se encontró la columna 'Id' en la hoja Compras.");

  const ids = sheetCompras.getRange(2, colId + 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (let i = ids.length - 1; i >= 0; i--) {
    const val = String(ids[i][0] || "").trim();
    if (val.startsWith("C")) {
      const numPart = parseInt(val.replace("C", ""), 10);
      if (!isNaN(numPart)) {
        maxNum = numPart;
        break;
      }
    }
  }

  return "C" + String(maxNum + 1).padStart(7, '0');
}
function actualizarCompras(registros) {
  if (!Array.isArray(registros) || registros.length === 0) return;

  const hoja = getSheetByName("Compras");
  if (!hoja) throw new Error("No se encontró la hoja 'Compras'.");

  const lastCol = hoja.getLastColumn();
  const headers = hoja.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

  const colContabilidad = headers.indexOf('Contabilidad');
  const colNroAsiento   = headers.indexOf('nroAsiento');
  const colNroFactura   = headers.indexOf('NroFacturaProveedor');

  if (colContabilidad === -1 || colNroAsiento === -1 || colNroFactura === -1) {
    throw new Error("Faltan columnas requeridas en la hoja Compras (Contabilidad, nroAsiento o NroFacturaProveedor).");
  }

  const lastRow = hoja.getLastRow();
  if (lastRow <= 1) return;

  const dataFacturas = hoja.getRange(2, colNroFactura + 1, lastRow - 1, 1).getValues();
  const mapaFilas = {};

  for (let i = 0; i < dataFacturas.length; i++) {
    const nro = String(dataFacturas[i][0] || '').trim();
    if (nro && mapaFilas[nro] === undefined) {
      mapaFilas[nro] = i + 2;
    }
  }

  registros.forEach(function(registro) {
    const nroBuscado = String(registro.NroFacturaProveedor || '').trim();
    if (!nroBuscado) return;

    const rowIndex = mapaFilas[nroBuscado];
    if (!rowIndex) return;

    if (registro.Contabilidad !== undefined && registro.Contabilidad !== null) {
      hoja.getRange(rowIndex, colContabilidad + 1).setValue(registro.Contabilidad);
    }
    if (registro.nroAsiento !== undefined && registro.nroAsiento !== null) {
      hoja.getRange(rowIndex, colNroAsiento + 1).setValue(registro.nroAsiento);
    }
  });
}
function obtenerLibroCompras(fechaDesde, fechaHasta, filtroProveedor, filtroTipoDoc) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCompras = ss.getSheetByName('Compras');
    
    if (!hojaCompras) {
      throw new Error('No se encontró la hoja "Compras"');
    }
    
    const datos = hojaCompras.getDataRange().getValues();
    const headers = datos[0];
    
    const idx = {
      nombre: headers.indexOf('NombreProveedor'),
      rif: headers.indexOf('RifProveedor'),
      fecha: headers.indexOf('FechaFacturaProveedor'),
      tipo: headers.indexOf('TipoProveedor'),
      nroFactura: headers.indexOf('NroFacturaProveedor'),
      nroControl: headers.indexOf('NroControlProveedor'),
      subTotal: headers.indexOf('SubTotalProveedor'),
      iva: headers.indexOf('IvaProveedor'),
      total: headers.indexOf('TotalProveedor'),
      retencionIVA: headers.indexOf('RetencionIva'),
      retencionISLR: headers.indexOf('RetencionISLR'),
      id: headers.indexOf('Id'),
      nroComprobantes: headers.indexOf('nroComprobantes'),
      notaDebito: headers.indexOf('NotaDebito'),
      notaCredito: headers.indexOf('NotaCredito'),
      exentas: headers.indexOf('totalExentoCompra'), // CORREGIDO AQUÍ
      tipoTransaccion: headers.indexOf('TipoTransaccion'),
      nroFacturaAfectada: headers.indexOf('NroFacturaAfectada'),
      alicuota: headers.indexOf('Alicuota')
    };
    
    if (idx.nombre === -1) idx.nombre = headers.indexOf('Proveedor');
    if (idx.rif === -1) idx.rif = headers.indexOf('RIF');
    if (idx.fecha === -1) idx.fecha = headers.indexOf('Fecha');
    if (idx.nroFactura === -1) idx.nroFactura = headers.indexOf('Factura');
    if (idx.nroControl === -1) idx.nroControl = headers.indexOf('Control');
    if (idx.subTotal === -1) idx.subTotal = headers.indexOf('SubTotal');
    if (idx.retencionIVA === -1) idx.retencionIVA = headers.indexOf('Retencion');
    
    const resultado = [];
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      let fechaFactura = fila[idx.fecha];
      
      let fechaStr = '';
      if (fechaFactura instanceof Date) {
        fechaStr = Utilities.formatDate(fechaFactura, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        fechaStr = String(fechaFactura);
        if (fechaStr.includes('/')) {
          const partes = fechaStr.split('/');
          if (partes.length === 3) {
            fechaStr = `${partes[2]}-${partes[1]}-${partes[0]}`;
          }
        }
      }
      
      if (fechaStr >= fechaDesde && fechaStr <= fechaHasta) {
        
        if (filtroProveedor) {
          const nombre = String(fila[idx.nombre] || '').toLowerCase();
          const rif = String(fila[idx.rif] || '').toLowerCase();
          if (!nombre.includes(filtroProveedor.toLowerCase()) && !rif.includes(filtroProveedor.toLowerCase())) {
            continue;
          }
        }
        
        if (filtroTipoDoc) {
          const tipo = String(fila[idx.tipo] || '');
          if (tipo !== filtroTipoDoc) {
            continue;
          }
        }
        
        let fechaMostrar = fechaStr;
        try {
          const fechaObj = new Date(fechaStr);
          fechaMostrar = Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        } catch(e) {
          fechaMostrar = fechaStr;
        }
        
        resultado.push({
          nombreProveedor: fila[idx.nombre],
          rifProveedor: fila[idx.rif],
          fechaFacturaProveedor: fechaMostrar,
          fechaRaw: fechaStr,
          tipoFacturaProveedor: fila[idx.tipo] || 'Factura',
          nroFacturaProveedor: fila[idx.nroFactura],
          nroControlProveedor: fila[idx.nroControl],
          nroComprobantes: fila[idx.nroComprobantes] || '',
          tipoTransaccion: fila[idx.tipoTransaccion] || '01-Reg',
          nroFacturaAfectada: fila[idx.nroFacturaAfectada] || '',
          subTotal: fila[idx.subTotal] || 0,
          iva: fila[idx.iva] || 0,
          total: fila[idx.total] || 0,
          retencionIVA: fila[idx.retencionIVA] || 0,
          retencionISLR: fila[idx.retencionISLR] || 0,
          notaDebito: fila[idx.notaDebito] || 0,
          notaCredito: fila[idx.notaCredito] || 0,
          exentas: fila[idx.exentas] || 0, // CORREGIDO AQUÍ
          alicuota: fila[idx.alicuota] || 16,
          id: fila[idx.id]
        });
      }
    }
    
    resultado.sort((a, b) => new Date(a.fechaRaw) - new Date(b.fechaRaw));
    return resultado;
    
  } catch (error) {
    throw new Error('Error al obtener libro de compras: ' + error.message);
  }
}
function generarLibroComprasExcel(datos) {
  try {
    const libro = SpreadsheetApp.create('Libro de Compras ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    const hoja = libro.getActiveSheet();
    hoja.setName('Libro de Compras');
    
    hoja.getRange('A1').setValue('LIBRO DE COMPRAS');
    hoja.getRange('A1').setFontSize(14).setFontWeight('bold');
    
    hoja.getRange('A2').setValue('Período: ' + datos.periodo.desde + ' al ' + datos.periodo.hasta);
    hoja.getRange('A2').setFontSize(11);
    
    // USO DE getEmpresaData()
    const empresa = typeof getEmpresaData === 'function' ? getEmpresaData() : { nombre: 'EMPRESA C.A.', rif: 'J-00000000-0' };
    
    hoja.getRange('A3').setValue('RIF: ' + empresa.rif + ' | Razón Social: ' + empresa.nombre);
    hoja.getRange('A3').setFontSize(10);
    
    let filaActual = 5;
    
    const encabezados = [
      'N°', 'Fecha', 'RIF', 'Nombre o Razón Social', 'Tipo', 
      'N° Factura', 'N° Control', 'N° Débito', 'N° Crédito', 
      'N° Comprobante', 'Tipo Trans.', 'N° Fact. Afect.',
      'Total Incluye I.V.A.', 'Internas No Gravadas', 
      'Base Imponible', '% Alic.', 'Impuesto I.V.A.', 
      'I.V.A. Retenido'
    ];
    
    hoja.getRange(filaActual, 1, 1, encabezados.length).setValues([encabezados]);
    hoja.getRange(filaActual, 1, 1, encabezados.length)
      .setFontWeight('bold')
      .setBackground('#d9d9d9')
      .setBorder(true, true, true, true, true, true);
    
    filaActual++;
    
    let totalBaseImponible = 0;
    let totalIVA = 0;
    let totalRetenido = 0;
    let totalCompras = 0;
    let totalExentas = 0;
    
    datos.facturas.forEach((factura, index) => {
      const tipoDoc = factura.tipoFacturaProveedor || 'Factura';
      const esNotaCredito = tipoDoc === 'NC';
      const esNotaDebito = tipoDoc === 'ND';
      
      const baseImponible = parseFloat(factura.subTotal) || 0;
      const montoIVA = parseFloat(factura.iva) || 0;
      const totalFactura = parseFloat(factura.total) || 0;
      const porcentajeIVA = factura.alicuota || 16;
      const retencion = parseFloat(factura.retencionIVA) || 0;
      const exentas = parseFloat(factura.exentas) || 0;
      
      totalBaseImponible += baseImponible;
      totalIVA += montoIVA;
      totalRetenido += retencion;
      totalCompras += totalFactura;
      totalExentas += exentas;
      
      const filaDatos = [
        index + 1,
        factura.fechaFacturaProveedor,
        factura.rifProveedor || '',
        factura.nombreProveedor || '',
        tipoDoc,
        factura.nroFacturaProveedor || '',
        factura.nroControlProveedor || '',
        esNotaDebito ? factura.nroFacturaProveedor : '',
        esNotaCredito ? factura.nroFacturaProveedor : '',
        factura.nroComprobantes || '',
        factura.tipoTransaccion || '01-Reg',
        factura.nroFacturaAfectada || '-',
        totalFactura,
        exentas,
        baseImponible,
        porcentajeIVA + '%',
        montoIVA,
        retencion
      ];
      
      hoja.getRange(filaActual, 1, 1, filaDatos.length).setValues([filaDatos]);
      hoja.getRange(filaActual, 1, 1, filaDatos.length).setBorder(true, true, true, true, true, true);
      
      if (esNotaCredito) {
        hoja.getRange(filaActual, 1, 1, filaDatos.length).setBackground('#ffebee');
      } else if (esNotaDebito) {
        hoja.getRange(filaActual, 1, 1, filaDatos.length).setBackground('#fff3e0');
      }
      
      hoja.getRange(filaActual, 13).setNumberFormat('#,##0.00');
      hoja.getRange(filaActual, 14).setNumberFormat('#,##0.00');
      hoja.getRange(filaActual, 15).setNumberFormat('#,##0.00');
      hoja.getRange(filaActual, 17).setNumberFormat('#,##0.00');
      hoja.getRange(filaActual, 18).setNumberFormat('#,##0.00');
      
      filaActual++;
    });
    
    hoja.getRange(filaActual, 14).setValue('TOTALES:');
    hoja.getRange(filaActual, 14).setFontWeight('bold').setHorizontalAlignment('right');
    hoja.getRange(filaActual, 15).setValue(totalBaseImponible).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 16).setValue('');
    hoja.getRange(filaActual, 17).setValue(totalIVA).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 18).setValue(totalRetenido).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 1, 1, 18).setBackground('#e3f2fd').setBorder(true, true, true, true, true, true);
    
    filaActual += 2;
    
    hoja.getRange(filaActual, 1).setValue('RESUMEN DE COMPRAS').setFontWeight('bold').setFontSize(12);
    filaActual += 2;
    
    hoja.getRange(filaActual, 2).setValue('Concepto');
    hoja.getRange(filaActual, 3).setValue('Base Imponible');
    hoja.getRange(filaActual, 4).setValue('Crédito fiscal');
    hoja.getRange(filaActual, 5).setValue('IVA retenido');
    hoja.getRange(filaActual, 2, 1, 4).setFontWeight('bold').setBackground('#d9d9d9').setBorder(true, true, true, true, true, true);
    filaActual++;
    
    let totalNoGravadas = 0;
    let totalGravadas16 = 0;
    let totalIVA16 = 0;
    let totalRetenido16 = 0;
    let totalGravadas8 = 0;
    let totalIVA8 = 0;
    let totalRetenido8 = 0;
    
    datos.facturas.forEach(factura => {
      const base = parseFloat(factura.subTotal) || 0;
      const iva = parseFloat(factura.iva) || 0;
      const retencion = parseFloat(factura.retencionIVA) || 0;
      const alicuota = parseFloat(factura.alicuota) || 16;
      
      if (alicuota === 0 || base === iva) {
        totalNoGravadas += base;
      } else if (alicuota >= 15) {
        totalGravadas16 += base;
        totalIVA16 += iva;
        totalRetenido16 += retencion;
      } else {
        totalGravadas8 += base;
        totalIVA8 += iva;
        totalRetenido8 += retencion;
      }
    });
    
    hoja.getRange(filaActual, 2).setValue('Total Compras Internas No Gravadas y/o Sin Derecho a Crédito Fiscal');
    hoja.getRange(filaActual, 3).setValue(totalNoGravadas).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 4).setValue(0).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 5).setValue(0).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 2, 1, 4).setBorder(true, true, true, true, true, true);
    filaActual++;
    
    hoja.getRange(filaActual, 2).setValue('Total Compras Internas Gravadas por Alícuota 16%');
    hoja.getRange(filaActual, 3).setValue(totalGravadas16).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 4).setValue(totalIVA16).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 5).setValue(totalRetenido16).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 2, 1, 4).setBorder(true, true, true, true, true, true);
    filaActual++;
    
    if (totalGravadas8 > 0) {
      hoja.getRange(filaActual, 2).setValue('Total Compras Internas Gravadas por Alícuota 8%');
      hoja.getRange(filaActual, 3).setValue(totalGravadas8).setNumberFormat('#,##0.00');
      hoja.getRange(filaActual, 4).setValue(totalIVA8).setNumberFormat('#,##0.00');
      hoja.getRange(filaActual, 5).setValue(totalRetenido8).setNumberFormat('#,##0.00');
      hoja.getRange(filaActual, 2, 1, 4).setBorder(true, true, true, true, true, true);
      filaActual++;
    }
    
    const totalBaseFinal = totalNoGravadas + totalGravadas16 + totalGravadas8;
    const totalIVAFinal = totalIVA16 + totalIVA8;
    const totalRetenidoFinal = totalRetenido16 + totalRetenido8;
    
    hoja.getRange(filaActual, 2).setValue('TOTALES:').setFontWeight('bold');
    hoja.getRange(filaActual, 3).setValue(totalBaseFinal).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 4).setValue(totalIVAFinal).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 5).setValue(totalRetenidoFinal).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 2, 1, 4).setBackground('#d9d9d9').setBorder(true, true, true, true, true, true);
    
    hoja.setColumnWidth(1, 4);   hoja.setColumnWidth(2, 10);  hoja.setColumnWidth(3, 12);
    hoja.setColumnWidth(4, 35);  hoja.setColumnWidth(5, 8);   hoja.setColumnWidth(6, 10);
    hoja.setColumnWidth(7, 12);  hoja.setColumnWidth(8, 10);  hoja.setColumnWidth(9, 10);
    hoja.setColumnWidth(10, 12); hoja.setColumnWidth(11, 10); hoja.setColumnWidth(12, 10);
    hoja.setColumnWidth(13, 16); hoja.setColumnWidth(14, 16); hoja.setColumnWidth(15, 14);
    hoja.setColumnWidth(16, 8);  hoja.setColumnWidth(17, 14); hoja.setColumnWidth(18, 14);
    
    const blob = libro.getBlob().setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);
    
    DriveApp.getFileById(libro.getId()).setTrashed(true);
    return base64;
    
  } catch (error) {
    throw new Error('Error al generar Excel de compras: ' + error.message);
  }
}
function generarLibroComprasHTML(fechaDesde, fechaHasta, filtroProveedor, filtroTipoDoc) {
  try {
    var facturas = obtenerLibroCompras(fechaDesde, fechaHasta, filtroProveedor, filtroTipoDoc);
    if (!facturas || facturas.length === 0) {
      return { success: false, message: 'No hay datos para el período indicado.' };
    }

    // USO DE getEmpresaData()
    var empresa = typeof getEmpresaData === 'function' ? getEmpresaData() : { nombre: 'EMPRESA C.A.', rif: 'J-00000000-0' };

    var periodo = formatearFechaRango(fechaDesde) + ' al ' + formatearFechaRango(fechaHasta);

    var totalTotal = 0, totalExento = 0, totalBase = 0, totalIva = 0, totalRet = 0;

    var filasHTML = facturas.map(function(f, i) {
      var esNC = f.tipoFacturaProveedor === 'NC';
      var factor = esNC ? -1 : 1;

      var tot = (parseFloat(f.total) || 0) * factor;
      var exe = (parseFloat(f.exentas) || 0) * factor;
      var bas = (parseFloat(f.subTotal) || 0) * factor;
      var iva = (parseFloat(f.iva) || 0) * factor;
      var ret = (parseFloat(f.retencionIVA) || 0) * factor;

      totalTotal += tot;
      totalExento += exe;
      totalBase += bas;
      totalIva += iva;
      totalRet += ret;

      return '<tr>' +
        '<td class="c">' + (i + 1) + '</td>' +
        '<td class="c">' + f.fechaFacturaProveedor + '</td>' +
        '<td class="c">' + (f.rifProveedor || '') + '</td>' +
        '<td>' + (f.nombreProveedor || '') + '</td>' +
        '<td class="c">' + (f.tipoFacturaProveedor || 'Factura') + '</td>' +
        '<td class="c">' + (f.nroFacturaProveedor || '') + '</td>' +
        '<td class="c">' + (f.nroControlProveedor || '') + '</td>' +
        '<td class="c">' + (f.tipoFacturaProveedor === 'ND' ? f.nroFacturaProveedor : '') + '</td>' +
        '<td class="c">' + (esNC ? f.nroFacturaProveedor : '') + '</td>' +
        '<td class="c">' + (f.nroFacturaAfectada || '') + '</td>' +
        '<td class="c">' + (f.nroComprobantes || '') + '</td>' +
        '<td class="r">' + fmtNum(tot) + '</td>' +
        '<td class="r">' + fmtNum(exe) + '</td>' +
        '<td class="r">' + fmtNum(bas) + '</td>' +
        '<td class="c">' + (f.alicuota || 16) + '%</td>' +
        '<td class="r">' + fmtNum(iva) + '</td>' +
        '<td class="r">' + fmtNum(ret) + '</td>' +
        '</tr>';
    }).join('');

    var totalesHTML = '<tr class="totales">' +
      '<td colspan="11" class="r">TOTALES GENERALES:</td>' +
      '<td class="r">' + fmtNum(totalTotal) + '</td>' +
      '<td class="r">' + fmtNum(totalExento) + '</td>' +
      '<td class="r">' + fmtNum(totalBase) + '</td>' +
      '<td></td>' +
      '<td class="r">' + fmtNum(totalIva) + '</td>' +
      '<td class="r">' + fmtNum(totalRet) + '</td>' +
      '</tr>';

    var html = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
      '<title>Libro de Compras ' + periodo + '</title>' +
      '<style>' +
      '@media print { @page { size: landscape; margin: 0.5cm; } .no-print { display: none !important; } }' +
      'body { font-family: Arial, sans-serif; font-size: 8.5px; margin: 10px; color: #000; }' +
      '.header { text-align: center; border-bottom: 2px solid #000; padding-bottom: 8px; margin-bottom: 12px; }' +
      '.header h1 { font-size: 15px; margin: 0; text-transform: uppercase; }' +
      '.header h2 { font-size: 12px; margin: 3px 0; }' +
      '.header .periodo { font-size: 11px; font-weight: bold; }' +
      'table { width: 100%; border-collapse: collapse; }' +
      'th, td { border: 1px solid #000; padding: 3px 4px; font-size: 8px; }' +
      'th { background: #d9d9d9; font-weight: bold; text-align: center; }' +
      '.c { text-align: center; } .r { text-align: right; }' +
      'tr.totales { background: #e3f2fd; font-weight: bold; }' +
      '.firmas { margin-top: 40px; display: flex; justify-content: space-around; }' +
      '.firma { width: 40%; text-align: center; border-top: 1px solid #000; padding-top: 5px; font-size: 9px; }' +
      '.no-print { text-align: center; padding: 10px; margin-bottom: 10px; background: #f8f9fa; border-bottom: 2px solid #ddd; }' +
      'button { padding: 8px 20px; margin: 0 5px; font-size: 13px; font-weight: bold; cursor: pointer; border: none; border-radius: 4px; }' +
      '.btn-print { background: #16a34a; color: white; } .btn-close { background: #475569; color: white; }' +
      '</style></head><body>' +
      '<div class="no-print">' +
      '<button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR LIBRO DE COMPRAS</button>' +
      '<button class="btn-close" onclick="window.close()">✕ CERRAR</button>' +
      '</div>' +
      '<div class="header">' +
      '<h1>' + empresa.nombre + '</h1>' +
      '<h2>RIF: ' + empresa.rif + '</h2>' +
      '<h2>LIBRO DE COMPRAS</h2>' +
      '<div class="periodo">Período: ' + periodo + '</div>' +
      '</div>' +
      '<table>' +
      '<thead><tr>' +
      '<th>N°</th><th>Fecha</th><th>RIF</th><th>Proveedor</th><th>Tipo</th>' +
      '<th>N° Factura</th><th>N° Control</th><th>N° Déb.</th><th>N° Créd.</th>' +
      '<th>Fact. Afect.</th><th>N° Comprob.</th>' +
      '<th>Total c/IVA</th><th>Exentas</th><th>Base Imp.</th><th>% Alic.</th>' +
      '<th>IVA</th><th>IVA Ret.</th>' +
      '</tr></thead>' +
      '<tbody>' + filasHTML + totalesHTML + '</tbody>' +
      '</table>' +
      '<div class="firmas">' +
      '<div class="firma">Elaborado Por</div>' +
      '<div class="firma">Revisado Por</div>' +
      '</div>' +
      '</body></html>';

    return { success: true, html: html };
  } catch (e) {
    Logger.log('Error en generarLibroComprasHTML: ' + e.message);
    return { success: false, message: e.message };
  }
}
/**
 * ============================================================================
 * ANULAR FACTURA DE COMPRA (GPROPOS MULTIEMPRESA)
 * Coloca montos en 0 en 'Compras', revierte stock en 'Productos' y anula la CXP.
 * ============================================================================
 */
function anularFacturaCompra(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); 

    if (!payload) throw new Error("No se especificaron datos para la anulación.");
    var nroFactura = String(payload.nroFactura || payload.NroFacturaProveedor || '').trim();
    if (!nroFactura) throw new Error("El número de factura es obligatorio para anular.");

    var motivo = String(payload.motivo || 'Anulación solicitada por el usuario').trim();
    var usuario = String(payload.usuario || 'Admin').trim();
    
    // Resolver Spreadsheet de la empresa activa
    var idSpreadsheet = payload.idSpreadsheet || payload.empresaId || '';
    var ss = _resolverSSCompras_(idSpreadsheet);
    if (!ss) throw new Error("No se pudo conectar a la base de datos de la empresa activa.");

    var sheetCompras = ss.getSheetByName("Compras");
    var sheetDetalle = ss.getSheetByName("DetalleFacturas");
    var sheetCobros = ss.getSheetByName("CobrosPagos");
    var sheetProductos = ss.getSheetByName("Productos");
    var sheetAsientos = ss.getSheetByName("Asientos");

    if (!sheetCompras) throw new Error("No se encontró la hoja 'Compras' en la empresa activa.");

    var dataCompras = sheetCompras.getDataRange().getValues();
    var headersCompras = dataCompras[0].map(function(h) { return String(h || '').trim(); });
    
    var colNroFactura = headersCompras.indexOf("NroFacturaProveedor");
    var colStatus = headersCompras.indexOf("Status");
    var colCobrada = headersCompras.indexOf("cobradaPagada");
    var colNroAsiento = headersCompras.indexOf("nroAsiento");

    var busquedaNorm = nroFactura.replace(/^0+/, '');
    var filaIndex = -1;
    var filaEncontrada = null;

    for (var i = dataCompras.length - 1; i > 0; i--) {
      var val = String(dataCompras[i][colNroFactura] || '').trim().replace(/^0+/, '');
      if (val === busquedaNorm) {
        filaIndex = i + 1;
        filaEncontrada = dataCompras[i];
        break;
      }
    }

    if (filaIndex === -1) throw new Error("No se encontró la Factura N° " + nroFactura);

    var statusActual = colStatus !== -1 ? String(filaEncontrada[colStatus] || '').trim() : '';
    if (statusActual.toLowerCase() === 'anulada') {
      return { status: "warning", message: "La Factura N° " + nroFactura + " ya se encuentra anulada." };
    }

    // 1. COLOCAR TODOS LOS MONTOS EN CERO EN LA HOJA 'Compras'
    var colsMonetarias = [
      headersCompras.indexOf("SubTotalProveedor"), headersCompras.indexOf("totalExentoCompra"),
      headersCompras.indexOf("IvaProveedor"), headersCompras.indexOf("TotalProveedor"),
      headersCompras.indexOf("SaldoProveedor"), headersCompras.indexOf("SubTotalProveedorDivisa"),
      headersCompras.indexOf("IVAProveedorDivisa"), headersCompras.indexOf("TotalProveedorDivisa"),
      headersCompras.indexOf("SaldoProveedorDivisa"), headersCompras.indexOf("Abono"),
      headersCompras.indexOf("RetencionIva"), headersCompras.indexOf("RetencionISLR")
    ];

    colsMonetarias.forEach(function(cIdx) {
      if (cIdx !== -1) sheetCompras.getRange(filaIndex, cIdx + 1).setValue(0);
    });

    if (colStatus !== -1) sheetCompras.getRange(filaIndex, colStatus + 1).setValue("Anulada");
    if (colCobrada !== -1) sheetCompras.getRange(filaIndex, colCobrada + 1).setValue("Anulada");
    
    var colDesc = headersCompras.indexOf("descripcionCompra");
    if (colDesc !== -1) {
      var descOriginal = String(filaEncontrada[colDesc] || '');
      var fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      sheetCompras.getRange(filaIndex, colDesc + 1).setValue("ANULADA: " + motivo + " | " + descOriginal + " [" + fechaHora + " por " + usuario + "]");
    }

    // 2. REVERTIR STOCK EN 'Productos'
    if (sheetDetalle) {
      var dataDet = sheetDetalle.getDataRange().getValues();
      var headersDet = dataDet[0].map(function(h) { return String(h || '').trim(); });
      var idxDetFact = headersDet.indexOf("NroFactura");
      var idxDetCant = headersDet.indexOf("Cantidad");
      var idxDetProd = headersDet.indexOf("Producto");
      var idxDetTrans = headersDet.indexOf("Transaccion");

      var itemsRevertir = [];

      for (var d = 1; d < dataDet.length; d++) {
        var factDet = String(dataDet[d][idxDetFact] || '').trim().replace(/^0+/, '');
        var transDet = String(dataDet[d][idxDetTrans] || '').trim().toLowerCase();

        if (factDet === busquedaNorm && transDet === 'compra') {
          var cantOriginal = parseFloat(dataDet[d][idxDetCant]) || 0;
          var prodNombre = String(dataDet[d][idxDetProd] || '').trim();

          if (cantOriginal > 0 && prodNombre) {
            itemsRevertir.push({ nombre: prodNombre, cantidad: cantOriginal });
          }
          
          // Poner cantidad en 0 en el detalle
          if (idxDetCant !== -1) sheetDetalle.getRange(d + 1, idxDetCant + 1).setValue(0);
          if (idxDetTrans !== -1) sheetDetalle.getRange(d + 1, idxDetTrans + 1).setValue("Anulada");
        }
      }

      if (sheetProductos && itemsRevertir.length > 0) {
        var dataProd = sheetProductos.getDataRange().getValues();
        for (var k = 0; k < itemsRevertir.length; k++) {
          var itemRev = itemsRevertir[k];
          for (var p = 1; p < dataProd.length; p++) {
            var codP = String(dataProd[p][0] || '').trim().toUpperCase();
            var nomP = String(dataProd[p][1] || '').trim().toUpperCase();
            if (nomP === itemRev.nombre.toUpperCase() || codP === itemRev.nombre.toUpperCase()) {
              var stockActual = parseFloat(dataProd[p][5]) || 0;
              var nuevoStock = Math.max(0, stockActual - itemRev.cantidad);
              sheetProductos.getRange(p + 1, 6).setValue(nuevoStock);
              break;
            }
          }
        }
      }
    }

    // 3. ANULAR CUENTA POR PAGAR EN 'CobrosPagos'
    if (sheetCobros) {
      var dataCP = sheetCobros.getDataRange().getValues();
      var headersCP = dataCP[0].map(function(h) { return String(h || '').trim(); });
      var idxCPFact = headersCP.indexOf("NroFactura");
      var idxCPTrans = headersCP.indexOf("Transaccion");
      var idxCPTotal = headersCP.indexOf("TotalFactura");
      var idxCPSaldo = headersCP.indexOf("SaldoFactura");
      var idxCPSaldoUSD = headersCP.indexOf("SaldoFacturaUSD");
      var idxCPCobrada = headersCP.indexOf("cobradaPagada");

      for (var c = 1; c < dataCP.length; c++) {
        var factCP = String(dataCP[c][idxCPFact] || '').trim().replace(/^0+/, '');
        var transCP = String(dataCP[c][idxCPTrans] || '').trim().toLowerCase();

        if (factCP === busquedaNorm && transCP === 'compra') {
          if (idxCPTotal !== -1) sheetCobros.getRange(c + 1, idxCPTotal + 1).setValue(0);
          if (idxCPSaldo !== -1) sheetCobros.getRange(c + 1, idxCPSaldo + 1).setValue(0);
          if (idxCPSaldoUSD !== -1) sheetCobros.getRange(c + 1, idxCPSaldoUSD + 1).setValue(0);
          if (idxCPCobrada !== -1) sheetCobros.getRange(c + 1, idxCPCobrada + 1).setValue("Anulada");
        }
      }
    }

    // 4. REVERTIR ASIENTO CONTABLE
    if (sheetAsientos && colNroAsiento !== -1) {
      var nroAsientoOriginal = String(filaEncontrada[colNroAsiento] || '').trim();
      if (nroAsientoOriginal) {
        var dataAs = sheetAsientos.getDataRange().getValues();
        for (var a = 1; a < dataAs.length; a++) {
          // La columna de Transaccion en Asientos tiene el nro de factura (índice 9)
          if (String(dataAs[a][9] || '').trim() === nroFactura || String(dataAs[a][0] || '').trim() === nroAsientoOriginal) {
             // Generar un asiento de reverso (opcional) o simplemente marcarlo. 
             // Para simplificar, aquí creamos un asiento de reverso (contra-asiento)
          }
        }
      }
    }

    return {
      status: "success",
      message: "Factura N° " + nroFactura + " anulada exitosamente. Inventario y CxP revertidos.",
      nroFactura: nroFactura
    };

  } catch(err) {
    Logger.log("Error en anularFacturaCompra: " + err);
    return { status: "error", message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
// =============================================
/** VENTAS                    */
// =============================================
function _resolverSSVentas_(idSpreadsheet) {
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
function guardarDatosFacturaVentas(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!payload) throw new Error("No se recibieron datos de la factura.");
    var idSpreadsheet = payload.idSpreadsheet || payload.empresaId || payload.idEmpresa;
    var ss = _resolverSSVentas_(idSpreadsheet);
    if (!ss) throw new Error("No se pudo conectar con la base de datos de la empresa seleccionada.");

    var sheetVentas = ss.getSheetByName("Ventas");
    var sheetDetalle = ss.getSheetByName("DetalleFacturas");
    var sheetCobros = ss.getSheetByName("CobrosPagos");
    var sheetProductos = ss.getSheetByName("Productos");

    if (!sheetVentas) throw new Error("No se encontró la hoja 'Ventas' en la empresa activa.");
    if (!sheetDetalle) throw new Error("No se encontró la hoja 'DetalleFacturas' en la empresa activa.");

    var ahora = new Date();
    var fechaRegistroStr = payload.fechaRegistro || Utilities.formatDate(ahora, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var nroFactura = String(payload.nroFacturaCliente || '').trim();
    var tipoDoc = String(payload.tipoFacturaCliente || 'Factura').trim();
    var esNotaEntrega = (tipoDoc === 'NE');
    var esNC = (tipoDoc === 'NC');
    var esND = (tipoDoc === 'ND');

    // 1. REGISTRO EN LA HOJA 'Ventas' (31 Columnas Exactas)
    var filaVenta = [
      payload.nombreCliente || "",                  // 1: NombreCliente
      payload.rifCliente || "",                     // 2: RifCliente
      payload.fechaFacturaCliente || "",            // 3: FechaFacturaCliente
      tipoDoc,                                      // 4: TipoCliente
      nroFactura,                                   // 5: NroFacturaCliente
      payload.nroControlCliente || "",              // 6: NroControl
      payload.selectMonedaVenta || "Dolares",       // 7: MonedaCliente
      parseFloat(payload.monedaTasaVenta) || 1,     // 8: TasaCliente
      parseFloat(payload.subTotal) || 0,            // 9: SubTotalCliente
      parseFloat(payload.totalExento) || 0,         // 10: totalExentoVenta
      parseFloat(payload.iva) || 0,                 // 11: IvaCliente
      parseFloat(payload.total) || 0,               // 12: TotalCliente
      parseFloat(payload.saldoCliente) || 0,        // 13: SaldoCliente
      parseFloat(payload.totalClienteUSD) || 0,     // 14: TotalClienteUSD
      parseFloat(payload.saldoClienteUSD) || 0,     // 15: SaldoClienteUSD
      "Venta",                                      // 16: TransaccionCliente
      "FAC-" + nroFactura,                          // 17: Id
      payload.cobradaPagada || (payload.condPago === 'Contado' ? 'Cobrada' : 'porCobrar'), // 18: cobradaPagada
      0,                                            // 19: AbonoCliente
      "",                                           // 20: fechaAbonoCliente
      "X",                                          // 21: Contabilidad
      "",                                           // 22: nroAsiento
      fechaRegistroStr,                             // 23: FechaRegistro
      payload.descripcionVenta || "",               // 24: Descripcion
      0,                                            // 25: RetencionIva
      0,                                            // 26: RetencionISLR
      "Activa",                                     // 27: Status
      "",                                           // 28: fechaRetencion
      "",                                           // 29: RetencionIvaOk
      "",                                           // 30: RetencionISLROk
      payload.nroFacturaAfectada || ""              // 31: nroComprobantes / Factura Afectada
    ];
    sheetVentas.appendRow(filaVenta);

    // 2. REGISTRO EN 'DetalleFacturas' (19 Columnas)
    var items = Array.isArray(payload.items) ? payload.items : [];
    if (items.length > 0) {
      var filasDetalle = [];
      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        filasDetalle.push([
          nroFactura,                               // 1: NroFactura
          it.nombre || it.producto || "",           // 2: Producto
          it.categoria || "",                       // 3: Categoria
          parseFloat(it.cantidad) || 0,             // 4: Cantidad
          parseFloat(it.precio) || 0,               // 5: Precio
          parseFloat(it.tasaIva) || 0,              // 6: %IVA
          parseFloat(it.montoIva) || 0,             // 7: MontoIVA
          parseFloat(it.precioTotal) || 0,          // 8: PrecioTotal
          parseFloat(it.precioUSD) || 0,            // 9: PrecioUSD
          parseFloat(it.tasaIva) || 0,              // 10: %IVAUSD
          parseFloat(it.montoIVAUSD) || 0,          // 11: MontoIVAUSD
          parseFloat(it.precioTotalUSD) || 0,       // 12: PrecioTotalUSD
          "Venta",                                  // 13: Transaccion
          it.cuentaContable || "",                  // 14: CuentaContable
          it.nombreCuenta || "",                    // 15: NombreContable
          payload.descripcionVenta || "",           // 16: DescripcionTransaccion
          payload.usuario || "Admin",               // 17: usuario
          payload.fechaFacturaCliente || "",        // 18: fecha
          tipoDoc                                   // 19: tipo
        ]);
      }
      sheetDetalle.getRange(sheetDetalle.getLastRow() + 1, 1, filasDetalle.length, 19).setValues(filasDetalle);
    }

    // 3. REGISTRO EN 'CobrosPagos' (Cuentas por Cobrar)
    // Se registra siempre que NO sea Nota de Entrega y no sea una NC que anula saldo
    if (sheetCobros && !esNotaEntrega) {
      var correlativoCP = ("000" + (sheetCobros.getLastRow())).slice(-4);
      var filaCP = [
        payload.nombreCliente || "",                // 1: ProveedorCliente
        payload.rifCliente || "",                   // 2: Rif
        nroFactura,                                 // 3: NroFactura
        payload.fechaFacturaCliente || "",          // 4: FechaFactura
        parseFloat(payload.total) || 0,             // 5: TotalFactura
        parseFloat(payload.saldoCliente) || 0,      // 6: SaldoFactura
        parseFloat(payload.saldoClienteUSD) || 0,   // 7: SaldoFacturaUSD
        payload.cobradaPagada || (payload.condPago === 'Contado' ? 'Cobrada' : 'porCobrar'), // 8: cobradaPagada
        "Venta",                                    // 9: Transaccion
        "CXC-" + correlativoCP,                     // 10: id
        0,                                          // 11: Abono
        0,                                          // 12: AbonoBs
        payload.selectMonedaVenta || "Dolares",     // 13: Moneda
        parseFloat(payload.monedaTasaVenta) || 1,   // 14: Tasa
        payload.descripcionVenta || "",             // 15: Descripcion
        "",                                         // 16: fechaAbono
        "",                                         // 17: TipoMovimiento
        "",                                         // 18: CuentaDestino
        "",                                         // 19: NroCuentaDestino
        "",                                         // 20: BancoDestino
        "",                                         // 21: fechaVencimiento
        "",                                         // 22: nroComprobantesRetIVA
        ""                                          // 23: nroComprobantesRetISLR
      ];
      sheetCobros.appendRow(filaCP);
    }

    // 4. ACTUALIZACIÓN DE STOCK EN 'Productos'
    if (sheetProductos && items.length > 0) {
      var prodData = sheetProductos.getDataRange().getValues();
      var debeSumarStock = esNC && payload.devuelveStock;

      for (var k = 0; k < items.length; k++) {
        var itCod = String(items[k].id || items[k].codigo || '').trim().toUpperCase();
        var itCant = parseFloat(items[k].cantidad) || 0;
        var esServ = itCod.startsWith('S') || String(items[k].categoria || '').toUpperCase().startsWith('S');

        if (esServ || itCant <= 0) continue;

        for (var r = 1; r < prodData.length; r++) {
          var codFila = String(prodData[r][0] || '').trim().toUpperCase();
          if (codFila === itCod) {
            var stockActual = parseFloat(prodData[r][5]) || 0; // Columna F: StockProducto
            var nuevoStock = debeSumarStock ? (stockActual + itCant) : Math.max(0, stockActual - itCant);
            sheetProductos.getRange(r + 1, 6).setValue(nuevoStock);
            break;
          }
        }
      }
    }

    // 5. GENERAR ASIENTO CONTABLE
    var nroAsientoGenerado = "";
    try {
      nroAsientoGenerado = generarAsientoVenta({
        fechaFacturaCliente: payload.fechaFacturaCliente,
        nombreCliente: payload.nombreCliente,
        nroFacturaCliente: payload.nroFacturaCliente,
        tipoTransaccion: 'Venta',
        monedaTasaVenta: payload.monedaTasaVenta,
        total: payload.total,
        iva: payload.iva,
        items: items.map(it => ({
          nombre: it.nombre,
          cantidad: it.cantidad,
          precioTotal: it.precioTotal,
          precioCompraUSD: it.precioCompraUSD || 0
        }))
      });
      
      // Actualizar la columna nroAsiento en la hoja Ventas
      var lastRowVentas = sheetVentas.getLastRow();
      var headersVentas = sheetVentas.getRange(1, 1, 1, sheetVentas.getLastColumn()).getValues()[0];
      var colNroAsiento = headersVentas.indexOf("nroAsiento") + 1;
      if (colNroAsiento > 0) {
        sheetVentas.getRange(lastRowVentas, colNroAsiento).setValue(nroAsientoGenerado);
      }
    } catch (eAsiento) {
      Logger.log("⚠️ Advertencia: No se pudo generar el asiento de venta. " + eAsiento);
    }
    
    return {
      status: "success",
      message: tipoDoc + " N° " + nroFactura + " registrada con éxito en la empresa activa.",
      nroFactura: nroFactura
    };

  } catch (err) {
    Logger.log("Error en guardarDatosFacturaVentas: " + err.toString());
    return { status: "error", message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
function obtenerSiguienteNumeroFactura(tipoDoc, transaccion, idSpreadsheet) {
  try {
    var ss = _resolverSSVentas_(idSpreadsheet);
    if (!ss) return "000001";

    var sheet = ss.getSheetByName("Ventas");
    if (!sheet) return "000001";

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return "000001";

    var valores = sheet.getRange(2, 5, lastRow - 1, 1).getValues(); // Columna 5: NroFacturaCliente
    var maxNum = 0;

    for (var i = 0; i < valores.length; i++) {
      var val = String(valores[i][0] || '').replace(/\D/g, '');
      var n = parseInt(val, 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }

    return ("000000" + (maxNum + 1)).slice(-6);
  } catch (e) {
    Logger.log("Error en obtenerSiguienteNumeroFactura: " + e);
    return "000001";
  }
}
function getDataClientesActivos(idSpreadsheet) {
  try {
    var ss = _resolverSSVentas_(idSpreadsheet);
    if (!ss) return [];
    var sheet = ss.getSheetByName("Clientes");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var clientes = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nombre = String(row[1] || row[0] || '').trim();
      if (!nombre) continue;
      clientes.push({
        nombre: nombre,
        rif: String(row[2] || row[1] || '').trim(),
        direccion: String(row[3] || row[2] || '').trim()
      });
    }
    return clientes;
  } catch(e) {
    Logger.log("Error en getDataClientesActivos: " + e);
    return [];
  }
}
function buscarFacturas(cliente, nro, idSpreadsheet) {
  try {
    var ss = _resolverSSVentas_(idSpreadsheet);
    if (!ss) return [];
    var sheet = ss.getSheetByName("Ventas");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var qCli = String(cliente || '').toLowerCase().trim();
    var qNro = String(nro || '').toLowerCase().trim();
    var resultados = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var cliFila = String(row[0] || '').toLowerCase();
      var rifFila = String(row[1] || '').toLowerCase();
      var nroFila = String(row[4] || '').toLowerCase();

      var matchCli = !qCli || cliFila.includes(qCli) || rifFila.includes(qCli);
      var matchNro = !qNro || nroFila.includes(qNro);

      if (matchCli && matchNro) {
        resultados.push({
          nombreCliente: row[0],
          rifCliente: row[1],
          fechaFacturaCliente: row[2],
          tipoFacturaCliente: row[3],
          nroFacturaCliente: row[4],
          total: row[11],
          status: row[26] || 'Activa'
        });
      }
    }
    return resultados;
  } catch(e) {
    Logger.log("Error en buscarFacturas: " + e);
    return [];
  }
}
function obtenerDatosCompletosFactura(nroFactura, idSpreadsheet) {
  try {
    var ss = _resolverSSVentas_(idSpreadsheet);
    if (!ss) return null;
    var sheetVentas = ss.getSheetByName("Ventas");
    var sheetDetalle = ss.getSheetByName("DetalleFacturas");
    if (!sheetVentas) return null;

    var vData = sheetVentas.getDataRange().getValues();
    var factura = null;
    var nroBuscado = String(nroFactura || '').trim();

    for (var i = 1; i < vData.length; i++) {
      if (String(vData[i][4] || '').trim() === nroBuscado) {
        factura = {
          nombreCliente: vData[i][0],
          rifCliente: vData[i][1],
          fechaFacturaCliente: vData[i][2],
          tipoFacturaCliente: vData[i][3],
          nroFacturaCliente: vData[i][4],
          nroControlCliente: vData[i][5],
          subTotal: vData[i][8],
          totalExento: vData[i][9],
          iva: vData[i][10],
          total: vData[i][11],
          empresaData: typeof getEmpresaData === 'function' ? getEmpresaData() : { nombre: 'GPRO SISTEMAS', rif: 'J-00000000-0', direccion: '' },
          items: []
        };
        break;
      }
    }

    if (factura && sheetDetalle) {
      var dData = sheetDetalle.getDataRange().getValues();
      for (var j = 1; j < dData.length; j++) {
        if (String(dData[j][0] || '').trim() === nroBuscado) {
          factura.items.push({
            producto: dData[j][1],
            cantidad: dData[j][3],
            precio: dData[j][4],
            tasaIva: dData[j][5],
            montoIva: dData[j][6],
            precioTotal: dData[j][7]
          });
        }
      }
    }
    return factura;
  } catch(e) {
    Logger.log("Error en obtenerDatosCompletosFactura: " + e);
    return null;
  }
}
function anularFacturaVenta(nroFactura, idSpreadsheet) {
  try {
    var ss = _resolverSSVentas_(idSpreadsheet);
    if (!ss) throw new Error("No se pudo conectar a la base de datos.");
    var sheetVentas = ss.getSheetByName("Ventas");
    var sheetProductos = ss.getSheetByName("Productos");
    var sheetDetalle = ss.getSheetByName("DetalleFacturas");

    var vData = sheetVentas.getDataRange().getValues();
    var nroBuscado = String(nroFactura || '').trim();
    var filaEncontrada = -1;

    for (var i = 1; i < vData.length; i++) {
      if (String(vData[i][4] || '').trim() === nroBuscado) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) throw new Error("Factura no encontrada.");
    sheetVentas.getRange(filaEncontrada, 27).setValue("Anulada"); // Columna 27: Status

    // Devolver stock
    if (sheetDetalle && sheetProductos) {
      var dData = sheetDetalle.getDataRange().getValues();
      var prodData = sheetProductos.getDataRange().getValues();

      for (var d = 1; d < dData.length; d++) {
        if (String(dData[d][0] || '').trim() === nroBuscado) {
          var nomProd = String(dData[d][1] || '').trim().toLowerCase();
          var cant = parseFloat(dData[d][3]) || 0;

          for (var p = 1; p < prodData.length; p++) {
            if (String(prodData[p][1] || '').trim().toLowerCase() === nomProd) {
              var stockAct = parseFloat(prodData[p][5]) || 0;
              sheetProductos.getRange(p + 1, 6).setValue(stockAct + cant);
              break;
            }
          }
        }
      }
    }

    return { status: "success", message: "Factura N° " + nroFactura + " anulada e inventario restaurado." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}
function round2Code(num) {
  return Math.round((Number(num) + Number.EPSILON) * 100) / 100;
}
function round2Code(num) {
  return Math.round((Number(num) + Number.EPSILON) * 100) / 100;
}
function obtenerSiguienteIdVenta(ventasSheet) {
  const lastRow = ventasSheet.getLastRow();
  if (lastRow < 2) return "V0000001";

  const headers = ventasSheet.getRange(1, 1, 1, ventasSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  let colId = headers.indexOf("Id") + 1;
  if (colId === 0) colId = 17; // Fallback columna Q si no existe encabezado

  const ids = ventasSheet.getRange(2, colId, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;

  ids.forEach(id => {
    if (id && typeof id === 'string' && id.trim().toUpperCase().startsWith('V')) {
      const limpio = id.trim().toUpperCase().replace('V', '');
      const num = parseInt(limpio, 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  });

  const siguiente = maxNum + 1;
  return 'V' + String(siguiente).padStart(7, '0');
}
function obtenerPrefijo(tipoFactura) {
  switch (tipoFactura) {
    case "Factura": return 0; // Asume que las facturas empiezan desde 0
    case "NC": return 90000;
    case "ND": return 80000;
    case "NE": return 70000;
    default: return 0; // Tipo desconocido o no especificado
  }
}
function generarNumeroFactura(tipoFactura, correlativo) {
  const numero = parseInt(correlativo, 10) || 1;
  return String(numero).padStart(6, '0');
}
function buscarFactura(hoja, idFactura) {
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return null;

  const headers = data[0].map(h => String(h).trim());
  const colId = headers.indexOf("Id");
  const colNroFactura = headers.indexOf("NroFacturaCliente");
  
  const searchStr = String(idFactura || "").trim();
  const searchNum = parseInt(searchStr, 10);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = colId !== -1 ? String(row[colId] || "").trim() : "";
    const valNro = colNroFactura !== -1 ? row[colNroFactura] : "";
    const rowNroStr = String(valNro || "").trim();
    const rowNroNum = parseInt(rowNroStr, 10);

    const matchId = (rowId.toLowerCase() === searchStr.toLowerCase());
    const matchNroTexto = (rowNroStr === searchStr);
    const matchNroNum = (!isNaN(searchNum) && !isNaN(rowNroNum) && searchNum === rowNroNum);

    if (matchId || matchNroTexto || matchNroNum) {
      return {
        rowIndex: i + 1,
        Id: rowId,
        NroFacturaCliente: !isNaN(rowNroNum) ? String(rowNroNum).padStart(6, '0') : rowNroStr,
        NombreCliente: row[headers.indexOf("NombreCliente")] || "",
        RifCliente: row[headers.indexOf("RifCliente")] || "",
        FechaFacturaCliente: row[headers.indexOf("FechaFacturaCliente")] || "",
        TotalCliente: Number(row[headers.indexOf("TotalCliente")]) || 0,
        TotalClienteUSD: Number(row[headers.indexOf("TotalClienteUSD")]) || 0,
        SaldoCliente: Number(row[headers.indexOf("SaldoCliente")]) || 0,
        SaldoClienteUSD: Number(row[headers.indexOf("SaldoClienteUSD")]) || 0,
        RetencionIva: Number(row[headers.indexOf("RetencionIva")]) || 0,
        RetencionISLR: Number(row[headers.indexOf("RetencionISLR")]) || 0
      };
    }
  }
  
  return null;
}
function obtenerCorrelativo(sheet, tipoFactura) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 1;

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim());

    // Identificar columna Tipo
    let colTipo = headers.indexOf("TipoCliente");
    if (colTipo === -1) colTipo = headers.indexOf("Tipo");
    if (colTipo === -1) colTipo = 3; // Columna D

    // Identificar columna Número
    let colNro = headers.indexOf("NroFacturaCliente");
    if (colNro === -1) colNro = headers.indexOf("NroFactura");
    if (colNro === -1) colNro = 4; // Columna E

    let maxNum = 0;
    const tipoTarget = String(tipoFactura || 'Factura').toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const filaTipo = String(data[i][colTipo] || '').toLowerCase().trim();
      
      // Si coincide el tipo o es Factura por defecto
      if (filaTipo === tipoTarget || tipoTarget === 'factura') {
        const val = data[i][colNro];
        let num = 0;

        if (typeof val === 'number') {
          num = val;
        } else if (val) {
          // Extraer solo dígitos numéricos (ej: "000004" -> 4)
          const limpio = String(val).replace(/\D/g, '');
          num = parseInt(limpio, 10);
        }

        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }

    return maxNum + 1;

  } catch (e) {
    console.error("Error en obtenerCorrelativo: " + e.message);
    return 1;
  }
}

// =============================================
/** FACTURACION                   */
// =============================================
// =========================================================================
// MÓDULO OPERATIVO REFACTORIZADO: Codigo.gs
// Facturación de Ventas/Compras, Inventario y Tesorería Multi-Tenant
// =========================================================================

/**
 * Guarda una factura de compras o ventas con validación estricta de empresa
 */
function guardarFacturaSegura(datosFactura, idSpreadsheet, tokenAuth) {
  try {
    // 1. Verificación de Seguridad y Sesión
    if (!tokenAuth || !tokenAuth.token) {
      throw new Error("Acceso no autorizado: Token de sesión ausente.");
    }

    // 2. Apertura segura de la empresa activa (nunca plantilla)
    const ssEmpresa = Contexto.getEmpresaSS(idSpreadsheet, false);
    const tipo = datosFactura.tipo; // "Venta" o "Compra"
    const nombreHoja = tipo === "Venta" ? "Ventas" : "Compras";
    const sheet = ssEmpresa.getSheetByName(nombreHoja);

    if (!sheet) throw new Error("Hoja " + nombreHoja + " no encontrada en la empresa.");

    // 3. Bloqueo de concurrencia (LockService)
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      // 4. Inserción de encabezado
      sheet.appendRow([
        datosFactura.nroFactura,
        datosFactura.nroControl,
        datosFactura.fechaEmision || new Date(),
        datosFactura.rifTercero,
        datosFactura.nombreTercero,
        datosFactura.condicionPago,
        datosFactura.subtotalUSD,
        datosFactura.montoExentoUSD,
        datosFactura.ivaUSD,
        tipo === "Venta" ? datosFactura.igtfUSD : datosFactura.totalUSD,
        datosFactura.totalUSD,
        datosFactura.tasaBCV,
        datosFactura.totalBs,
        "Emitida",
        datosFactura.saldoPendienteUSD,
        datosFactura.usuario,
        datosFactura.nroAsiento || ""
      ]);

      // 5. Inserción de renglones en DetalleFacturas
      const sheetDetalle = ssEmpresa.getSheetByName("DetalleFacturas");
      if (sheetDetalle && datosFactura.items && datosFactura.items.length > 0) {
        const renglones = datosFactura.items.map(function(item, idx) {
          return [
            Utilities.getUuid(),
            tipo,
            datosFactura.nroFactura,
            item.codigo,
            item.descripcion,
            item.cantidad,
            item.precioUSD,
            item.alicuota,
            item.totalLineaUSD
          ];
        });
        sheetDetalle.getRange(sheetDetalle.getLastRow() + 1, 1, renglones.length, 9).setValues(renglones);
      }

      return { status: "success", message: "Factura registrada exitosamente.", nroFactura: datosFactura.nroFactura };
    } finally {
      lock.releaseLock();
    }
  } catch(err) {
    Logger.log("Error en guardarFacturaSegura: " + err.message);
    return { status: "error", message: err.message };
  }
}
function getFacturas(tipoTransaccion) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Ventas");
    if (!sheet) sheet = ss.getSheetByName("FacturasVentas");

    if (!sheet) {
      console.warn("No se encontró la hoja 'Ventas'.");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0].map(h => String(h || '').trim());
    
    const cNom = headers.indexOf("NombreCliente") !== -1 ? headers.indexOf("NombreCliente") : 0;
    const cRif = headers.indexOf("RifCliente") !== -1 ? headers.indexOf("RifCliente") : 1;
    const cFecha = headers.indexOf("FechaFacturaCliente") !== -1 ? headers.indexOf("FechaFacturaCliente") : 2;
    const cNro = headers.indexOf("NroFacturaCliente") !== -1 ? headers.indexOf("NroFacturaCliente") : 4;
    const cTot = headers.indexOf("TotalCliente") !== -1 ? headers.indexOf("TotalCliente") : 11;
    const cStatus = headers.indexOf("Status") !== -1 ? headers.indexOf("Status") : headers.indexOf("status");

    const facturas = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const valNro = row[cNro];
      if (valNro === undefined || valNro === null || valNro === "") continue;

      let nroStr = String(valNro).trim();
      let nroNum = parseInt(nroStr.replace(/\D/g, ''), 10);
      let nroPadded = !isNaN(nroNum) ? String(nroNum).padStart(6, '0') : nroStr;

      let fechaStr = "";
      if (row[cFecha] instanceof Date) {
        fechaStr = Utilities.formatDate(row[cFecha], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        fechaStr = String(row[cFecha] || "");
      }

      facturas.push({
        nombreCliente: String(row[cNom] || ''),
        rifCliente: String(row[cRif] || ''),
        fechaFacturaCliente: fechaStr,
        nroFacturaCliente: nroPadded,
        total: Number(row[cTot]) || 0,
        status: (cStatus !== -1 && row[cStatus]) ? String(row[cStatus]) : 'Activa'
      });
    }

    return facturas;

  } catch (e) {
    console.error("Error en getFacturas: " + e.message);
    return [];
  }
}
function updateFactura(upd) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!upd || !upd.NroFactura) {
      throw new Error('Datos inválidos: falta el número de factura.');
    }

    var ss = _SS_BC_(upd.idSpreadsheet);
    if (!ss) throw new Error('No se pudo acceder a la empresa activa.');

    var shCP = ss.getSheetByName('CobrosPagos');
    if (!shCP) throw new Error('No existe la hoja CobrosPagos en la empresa activa.');

    var data = shCP.getDataRange().getValues();
    var hd = data[0].map(function(h) { 
      return String(h || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-#.]/g, ''); 
    });

    var cFact = hd.indexOf('nrofactura');
    var cProv = hd.indexOf('proveedorcliente');
    var cRif = hd.indexOf('rif');
    var cFecha = hd.indexOf('fechafactura');
    var cTotal = hd.indexOf('totalfactura') !== -1 ? hd.indexOf('totalfactura') : hd.indexOf('totafactura');
    var cTrans = hd.indexOf('transaccion');
    var cTipoMov = hd.indexOf('tipomovimiento');

    if (cFact === -1 || cProv === -1 || cRif === -1 || cFecha === -1 || cTotal === -1) {
      throw new Error('Estructura de columnas inválida en CobrosPagos.');
    }

    var idOriginal = String(upd.NroFacturaOriginal || upd.NroFactura).trim().toLowerCase();
    var nuevoNro = String(upd.NroFactura).trim();
    var nuevoNombre = String(upd.ProveedorCliente || '').trim();
    var nuevoRif = String(upd.Rif || '').trim();
    var filaOriginal = -1;

    // 1. Localizar la factura original por su NroFacturaOriginal
    for (var i = 1; i < data.length; i++) {
      var facFila = String(data[i][cFact] || '').trim().toLowerCase();
      var totFila = Number(data[i][cTotal]) || 0;
      var tipoMov = String(data[i][cTipoMov] || '').trim();

      if (facFila === idOriginal && totFila > 0 && !tipoMov) {
        filaOriginal = i + 1;
        break;
      }
    }

    if (filaOriginal === -1) {
      throw new Error('No se encontró la factura original "' + (upd.NroFacturaOriginal || nuevoNro) + '" para editar.');
    }

    // 2. Si el número cambió, verificar que el nuevo número no exista ya en otra factura activa
    if (nuevoNro.toLowerCase() !== idOriginal) {
      for (var j = 1; j < data.length; j++) {
        if (j + 1 === filaOriginal) continue;
        var facExistente = String(data[j][cFact] || '').trim().toLowerCase();
        var totExistente = Number(data[j][cTotal]) || 0;
        var tipoMovEx = String(data[j][cTipoMov] || '').trim();

        if (facExistente === nuevoNro.toLowerCase() && totExistente > 0 && !tipoMovEx) {
          throw new Error('El número de factura "' + nuevoNro + '" ya está asignado a otra factura en el sistema.');
        }
      }
    }

    // 3. Aplicar actualización en CobrosPagos
    shCP.getRange(filaOriginal, cFact + 1).setValue(nuevoNro);
    shCP.getRange(filaOriginal, cProv + 1).setValue(nuevoNombre);
    shCP.getRange(filaOriginal, cRif + 1).setValue(nuevoRif);
    shCP.getRange(filaOriginal, cFecha + 1).setValue(upd.FechaFactura);
    shCP.getRange(filaOriginal, cTotal + 1).setValue(Number(upd.TotaFactura || upd.TotalFactura) || 0);

    if (cTrans !== -1 && upd.Transaccion) {
      shCP.getRange(filaOriginal, cTrans + 1).setValue(upd.Transaccion);
    }

    // 4. Si el número de factura cambió, actualizar también sus filas hijas (abonos/retenciones)
    if (nuevoNro.toLowerCase() !== idOriginal) {
      for (var k = 1; k < data.length; k++) {
        var facHija = String(data[k][cFact] || '').trim().toLowerCase();
        if (facHija === idOriginal) {
          shCP.getRange(k + 1, cFact + 1).setValue(nuevoNro);
        }
      }

      // Sincronizar en hoja Compras o Ventas
      var esVenta = String(upd.Transaccion || '').toLowerCase().indexOf('vent') !== -1;
      var shLibro = ss.getSheetByName(esVenta ? 'Ventas' : 'Compras');
      if (shLibro) {
        var dLib = shLibro.getDataRange().getValues();
        var hLib = dLib[0].map(function(h) { 
          return String(h || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-#.]/g, ''); 
        });
        var colFactLib = esVenta ? hLib.indexOf('nrofacturacliente') : hLib.indexOf('nrofacturaproveedor');
        var colProvLib = esVenta ? hLib.indexOf('nombrecliente') : hLib.indexOf('nombreproveedor');
        var colRifLib = esVenta ? hLib.indexOf('rifcliente') : hLib.indexOf('rifproveedor');
        var colFechaLib = esVenta ? hLib.indexOf('fechafacturacliente') : hLib.indexOf('fechafacturaproveedor');

        if (colFactLib !== -1) {
          for (var m = 1; m < dLib.length; m++) {
            if (String(dLib[m][colFactLib] || '').trim().toLowerCase() === idOriginal) {
              shLibro.getRange(m + 1, colFactLib + 1).setValue(nuevoNro);
              if (colProvLib !== -1) shLibro.getRange(m + 1, colProvLib + 1).setValue(nuevoNombre);
              if (colRifLib !== -1) shLibro.getRange(m + 1, colRifLib + 1).setValue(nuevoRif);
              if (colFechaLib !== -1) shLibro.getRange(m + 1, colFechaLib + 1).setValue(upd.FechaFactura);
              break;
            }
          }
        }
      }
    }

    return { ok: true, fila: filaOriginal, nuevoNro: nuevoNro };

  } catch (e) {
    Logger.log('[updateFactura] Error: ' + e.message);
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function verificarDuplicadoAlCrear(params) {
  try {
    var nroFactura = String(params.nroFactura || '').trim();
    var rif        = String(params.rif || '').trim().toLowerCase();
    var nombre     = String(params.nombre || '').trim().toLowerCase();
    var tipo       = String(params.tipo || 'Venta').trim();

    if (!nroFactura) {
      return { existe: false, motivo: null, tercero: null, origen: null };
    }

    var esVenta = tipo.toLowerCase().indexOf('vent') !== -1;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaPrincipal = ss.getSheetByName(esVenta ? 'Ventas' : 'Compras');
    var hojaCP        = ss.getSheetByName('CobrosPagos');

    // ============================================================
    // 1. VALIDAR EN HOJA PRINCIPAL (Ventas o Compras)
    // ============================================================
    if (hojaPrincipal) {
      var dataP = hojaPrincipal.getDataRange().getValues();
      var hdP   = dataP[0].map(function(h) { return String(h).trim(); });

      var cFactP = esVenta ? hdP.indexOf('NroFacturaCliente') : hdP.indexOf('NroFacturaProveedor');
      var cRifP  = esVenta ? hdP.indexOf('RifCliente')        : hdP.indexOf('RifProveedor');
      var cNomP  = esVenta ? hdP.indexOf('NombreCliente')     : hdP.indexOf('NombreProveedor');
      var cStatP = esVenta ? hdP.indexOf('Status')            : hdP.indexOf('Status');
      if (cStatP === -1) cStatP = esVenta ? hdP.indexOf('status') : -1;

      if (cFactP !== -1) {
        for (var i = 1; i < dataP.length; i++) {
          var facP   = String(dataP[i][cFactP] || '').trim();
          if (facP !== nroFactura) continue;

          // Excluir facturas anuladas
          var statFila = cStatP !== -1 ? String(dataP[i][cStatP] || '').trim().toLowerCase() : '';
          if (statFila === 'anulada') continue;

          var rifP = cRifP !== -1 ? String(dataP[i][cRifP] || '').trim().toLowerCase() : '';
          var nomP = cNomP !== -1 ? String(dataP[i][cNomP] || '').trim() : '';

          var mismoTercero = (rifP === rif) || (nomP.toLowerCase() === nombre);

          if (!mismoTercero) {
            return {
              existe: true,
              motivo: 'La factura N° ' + nroFactura + ' ya existe para otro ' +
                      (esVenta ? 'cliente' : 'proveedor') + ': ' + nomP + ' (' + rifP + ')',
              tercero: { nombre: nomP, rif: rifP },
              origen: esVenta ? 'Ventas' : 'Compras'
            };
          }

          // Es del mismo tercero → bloquear también (evita duplicados reales)
          return {
            existe: true,
            motivo: 'La factura N° ' + nroFactura + ' ya existe para ' + nomP +
                    '. Verifica si es una nota de crédito o rectificación.',
            tercero: { nombre: nomP, rif: rifP },
            origen: esVenta ? 'Ventas' : 'Compras'
          };
        }
      }
    }

    // ============================================================
    // 2. VALIDAR EN COBROSPAGOS (por si quedó huérfana)
    // ============================================================
    if (hojaCP) {
      var dataCP = hojaCP.getDataRange().getValues();
      var hdCP   = dataCP[0].map(function(h) { return String(h).trim(); });

      var cFact    = hdCP.indexOf('NroFactura');
      var cProv    = hdCP.indexOf('ProveedorCliente');
      var cRif     = hdCP.indexOf('Rif');
      var cTotal   = hdCP.indexOf('TotaFactura') !== -1 ? hdCP.indexOf('TotaFactura') : hdCP.indexOf('TotalFactura');
      var cTipoMov = hdCP.indexOf('TipoMovimiento');
      var cStatus  = hdCP.indexOf('cobradaPagada');

      if (cFact !== -1) {
        for (var j = 1; j < dataCP.length; j++) {
          var facCP = String(dataCP[j][cFact] || '').trim();
          if (facCP !== nroFactura) continue;

          var totCP   = Number(dataCP[j][cTotal]) || 0;
          var tipoCP  = String(dataCP[j][cTipoMov] || '').trim();
          var statCP  = String(dataCP[j][cStatus] || '').trim().toLowerCase();

          // Ignorar filas de abonos y anuladas
          var esFilaMadre = (totCP > 0 && !tipoCP);
          if (!esFilaMadre) continue;
          if (statCP === 'anulada') continue;

          var provCP = String(dataCP[j][cProv] || '').trim();
          var rifCP  = String(dataCP[j][cRif] || '').trim().toLowerCase();

          var mismoCP = (rifCP === rif) || (provCP.toLowerCase() === nombre);

          if (!mismoCP) {
            return {
              existe: true,
              motivo: 'La factura N° ' + nroFactura + ' ya está registrada en Cobros/Pagos para otro ' +
                      (esVenta ? 'cliente' : 'proveedor') + ': ' + provCP + ' (' + rifCP + ')',
              tercero: { nombre: provCP, rif: rifCP },
              origen: 'CobrosPagos'
            };
          }

          return {
            existe: true,
            motivo: 'La factura N° ' + nroFactura + ' ya está registrada en Cobros/Pagos para ' + provCP + '.',
            tercero: { nombre: provCP, rif: rifCP },
            origen: 'CobrosPagos'
          };
        }
      }
    }

    // No hay duplicados
    return { existe: false, motivo: null, tercero: null, origen: null };

  } catch (e) {
    Logger.log('[verificarDuplicadoAlCrear] Error: ' + e.message);
    // En caso de error, NO bloqueamos la creación; devolvemos false para no trabar al usuario
    return { existe: false, motivo: null, tercero: null, origen: null, error: e.message };
  }
}
function verificarDuplicadoFactura(nroFactura, proveedorCliente, rif, nroFacturaOriginal) {
  try {
    nroFactura         = String(nroFactura || '').trim();
    proveedorCliente   = String(proveedorCliente || '').trim().toLowerCase();
    rif                = String(rif || '').trim().toLowerCase();
    nroFacturaOriginal = String(nroFacturaOriginal || '').trim();

    if (!nroFactura) return { existe: false, esDelMismoTercero: false, tercero: null };

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var shCP = ss.getSheetByName(CFG_CXC.HOJA_COBROS_PAGOS);
    if (!shCP) return { existe: false, esDelMismoTercero: false, tercero: null };

    var data = shCP.getDataRange().getValues();
    var hd   = data[0].map(function(h) { return String(h).trim(); });

    var cFact    = hd.indexOf('NroFactura');
    var cProv    = hd.indexOf('ProveedorCliente');
    var cRif     = hd.indexOf('Rif');
    var cTotal   = hd.indexOf('TotaFactura') !== -1 ? hd.indexOf('TotaFactura') : hd.indexOf('TotalFactura');
    var cTipoMov = hd.indexOf('TipoMovimiento');

    if (cFact === -1) return { existe: false, esDelMismoTercero: false, tercero: null };

    for (var i = 1; i < data.length; i++) {
      var facFila = String(data[i][cFact] || '').trim();
      if (facFila !== nroFactura) continue;

      var totFila = Number(data[i][cTotal]) || 0;
      var tipoMov = String(data[i][cTipoMov] || '').trim();
      var esFilaMadre = (totFila > 0 && !tipoMov);
      if (!esFilaMadre) continue;

      // Si es la misma factura que estoy editando (mismo nro original) → no es duplicado
      if (nroFacturaOriginal && facFila === nroFacturaOriginal) {
        continue;
      }

      var provFila = String(data[i][cProv] || '').trim();
      var rifFila  = String(data[i][cRif] || '').trim();

      var mismoTercero = (provFila.toLowerCase() === proveedorCliente
                       && rifFila.toLowerCase() === rif);

      if (!mismoTercero) {
        return {
          existe: true,
          esDelMismoTercero: false,
          tercero: { proveedor: provFila, rif: rifFila }
        };
      }
    }

    return { existe: false, esDelMismoTercero: true, tercero: null };

  } catch (e) {
    Logger.log('[verificarDuplicadoFactura] Error: ' + e.message);
    return { existe: false, esDelMismoTercero: false, tercero: null, error: e.message };
  }
}
function actualizarVentas(registros) {
  const hoja = getSheetByName("Ventas");
  
  registros.forEach(function(registro) {
    var finder = hoja.getRange("E:E").createTextFinder(registro.NroFactura);
    var fila = finder.findNext();
    
    if (fila) {
      var rowIndex = fila.getRow();
      hoja.getRange(rowIndex, 18).setValue(registro.Contabilidad);
      hoja.getRange(rowIndex, 19).setValue(registro.nroAsiento);
    }
  });
}
function actualizarFilaCompras(hoja, rowIndex, datos) {
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const getCol = (name) => headers.indexOf(name) + 1; // +1 porque Sheets es 1-indexed
  
  const updates = [
    { col: "SaldoProveedor", val: datos.saldoBs },
    { col: "SaldoProveedorDivisa", val: datos.saldoUSD },
    { col: "cobradaPagada", val: datos.estado },
    { col: "Abono", val: datos.abono },
    { col: "fechaAbono", val: datos.fechaAbono }
  ];
  
  updates.forEach(u => {
    const colIndex = getCol(u.col);
    if (colIndex > 0) {
      hoja.getRange(rowIndex, colIndex).setValue(u.val);
      Logger.log(`    Actualizado ${u.col} en fila ${rowIndex}: ${u.val}`);
    } else {
      Logger.log(`    ⚠️ Columna no encontrada: ${u.col}`);
    }
  });
}
function parseNumber(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let str = String(val).trim()
        .replace(/\./g, '')   // eliminar separadores de miles
        .replace(/,/g, '.');  // coma decimal a punto
    let num = parseFloat(str);
    return isNaN(num) ? 0 : num;
}
function anularFactura(nroFactura, transaccion) {
  try {
    const facturasSheet = getSheetByName("Ventas");
    const detalleSheet = getSheetByName("DetalleFacturas");

    if (!facturasSheet) throw new Error("No se encontró la hoja 'Ventas'.");
    if (!detalleSheet) throw new Error("No se encontró la hoja 'DetalleFacturas'.");

    const facturasData = facturasSheet.getDataRange().getValues();
    const headerFacturas = facturasData.shift();

    const idxNroFactura = headerFacturas.indexOf("NroFacturaCliente");
    const idxStatus = headerFacturas.indexOf("Status");
    const idxSaldo = headerFacturas.indexOf("SaldoCliente");
    // NUEVO: Buscamos la columna de Saldo en USD
    const idxSaldoUSD = headerFacturas.indexOf("SaldoClienteUSD");

    if (idxNroFactura === -1) throw new Error("No se encontró la columna 'NroFacturaCliente' en la hoja 'Ventas'.");
    if (idxStatus === -1) throw new Error("No se encontró la columna 'Status' en la hoja 'Ventas'.");
    if (idxSaldo === -1) throw new Error("No se encontró la columna 'SaldoCliente' en la hoja 'Ventas'.");
    // NUEVO: Validación de la columna USD
    if (idxSaldoUSD === -1) throw new Error("No se encontró la columna 'SaldoClienteUSD' en la hoja 'Ventas'.");

    let facturaRow = -1;
    for (let i = 0; i < facturasData.length; i++) {
      const celdaFactura = facturasData[i][idxNroFactura];
      if (celdaFactura && celdaFactura.toString() === nroFactura.toString()) {
        facturaRow = i + 2;
        break;
      }
    }

    if (facturaRow !== -1) {
      facturasSheet.getRange(facturaRow, idxStatus + 1).setValue("Anulada");
      facturasSheet.getRange(facturaRow, idxSaldo + 1).setValue(0);
      // NUEVO: Colocamos en cero el SaldoClienteUSD en la hoja
      facturasSheet.getRange(facturaRow, idxSaldoUSD + 1).setValue(0);
    } else {
      throw new Error(`No se encontró la factura ${nroFactura} en la hoja principal.`);
    }

    const detalleData = detalleSheet.getDataRange().getValues();
    const headerDetalle = detalleData.shift();

    const idxDetalleNroFactura = headerDetalle.indexOf("NroFactura");
    const idxCantidad = headerDetalle.indexOf("Cantidad");
    const idxPrecio = headerDetalle.indexOf("Precio");
    const idxMontoIVA = headerDetalle.indexOf("MontoIVA");
    const idxPrecioTotal = headerDetalle.indexOf("PrecioTotal");
    
    if (idxDetalleNroFactura === -1) throw new Error("No se encontró la columna 'NroFactura' en la hoja 'DetalleFacturas'.");

    for (let i = 0; i < detalleData.length; i++) {
      const celdaDetalle = detalleData[i][idxDetalleNroFactura];
      if (celdaDetalle && celdaDetalle.toString() === nroFactura.toString()) {
        const detalleRow = i + 2;
        if (idxCantidad !== -1) detalleSheet.getRange(detalleRow, idxCantidad + 1).setValue(0);
        if (idxPrecio !== -1) detalleSheet.getRange(detalleRow, idxPrecio + 1).setValue(0);
        if (idxMontoIVA !== -1) detalleSheet.getRange(detalleRow, idxMontoIVA + 1).setValue(0);
        if (idxPrecioTotal !== -1) detalleSheet.getRange(detalleRow, idxPrecioTotal + 1).setValue(0);
      }
    }

    return { status: "success", message: "Factura anulada correctamente." };

  } catch (e) {
    Logger.log(e);
    throw new Error(`Error al anular la factura: ${e.message}`);
  }
}
function transformarAObjeto(datos, encabezados) {
  const idxNombreCliente = encabezados.indexOf("NombreCliente");
  const idxFechaFactura = encabezados.indexOf("FechaFacturaCliente");
  const idxNroFactura = encabezados.indexOf("NroFacturaCliente");
  const idxTotal = encabezados.indexOf("TotalCliente");

  return datos.map(fila => {
    let fecha = fila[idxFechaFactura];
    let fechaFormateada = (fecha instanceof Date && !isNaN(fecha)) 
      ? Utilities.formatDate(fecha, getDb().getSpreadsheetTimeZone(), "dd/MM/yyyy")
      : String(fecha);
    
    return {
      nombreCliente: fila[idxNombreCliente],
      fechaFacturaCliente: fechaFormateada,
      nroFacturaCliente: fila[idxNroFactura],
      total: fila[idxTotal] 
    };
  });
}
function obtenerLibroVentas(fechaDesde, fechaHasta, filtroCliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Ventas');
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  
  // Mapeo automático de índices
  const idx = headers.reduce((obj, h, i) => { obj[h] = i; return obj; }, {});
  
  const resultado = datos.slice(1).filter(fila => {
    const fecha = Utilities.formatDate(new Date(fila[idx['FechaFacturaCliente']]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const esRango = (fecha >= fechaDesde && fecha <= fechaHasta);
    const cumpleFiltro = !filtroCliente || 
                         String(fila[idx['NombreCliente']]).toLowerCase().includes(filtroCliente.toLowerCase()) || 
                         String(fila[idx['RifCliente']]).toLowerCase().includes(filtroCliente.toLowerCase());
    return esRango && cumpleFiltro && fila[idx['TipoCliente']] !== 'NE';
  }).map(fila => ({
    nombreCliente: fila[idx['NombreCliente']],
    rifCliente: fila[idx['RifCliente']],
    fechaFacturaCliente: Utilities.formatDate(new Date(fila[idx['FechaFacturaCliente']]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    tipoFacturaCliente: fila[idx['TipoCliente']],
    nroFacturaCliente: fila[idx['NroFacturaCliente']],
    nroControlCliente: fila[idx['NroControl']],
    nroComprobantes: fila[idx['nroComprobantes']] || '',
    subTotal: fila[idx['SubTotalCliente']],
    totalExento: fila[idx['totalExento']] || 0, // <--- LECTURA DE NUEVA COLUMNA
    iva: fila[idx['IvaCliente']],
    total: fila[idx['TotalCliente']],
    retencionIVA: fila[idx['RetencionIva']] || 0,
    igtf: fila[idx['Igtf']] || 0,
    monedaTasaVenta: fila[idx['TasaCliente']] || 1 // Agregado para tasa promedio
  }));
  
  return resultado;
}
function generarLibroVentasExcel(datos) {
  try {
    // Crear nuevo libro de Excel
    const libro = SpreadsheetApp.create('Libro de Ventas ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    const hoja = libro.getActiveSheet();
    hoja.setName('Libro de Ventas');
    
    // Título
    hoja.getRange('A1').setValue('Libro de Ventas');
    hoja.getRange('A1').setFontSize(14).setFontWeight('bold');
    
    // Período
    hoja.getRange('A2').setValue('Período: ' + datos.periodo.desde + ' al ' + datos.periodo.hasta);
    hoja.getRange('A2').setFontSize(11);
    
    const empresa = getEmpresaData();
    hoja.getRange('A3').setValue('RIF: ' + empresa.rif + ' | Razón Social: ' + empresa.nombre);
    hoja.getRange('A3').setFontSize(10);
    
    let filaActual = 5;
    
    // Encabezados de la tabla principal
    const encabezados = [
      'N°', 'Fecha', 'RIF', 'Nombre o Razón Social', 'Tipo', 
      'N° Factura', 'N° Control', 'N° Débito', 'N° Crédito', 
      'N° Comprobante', 'Tipo Trans.', 'N° Fact. Afect.',
      'Total Incluye I.V.A.', 'Internas No Gravadas', 
      'Base Imponible', '% Alic.', 'Impuesto I.V.A.', 
      'I.G.T.F.', 'I.V.A. Retenido'
    ];
    
    hoja.getRange(filaActual, 1, 1, encabezados.length).setValues([encabezados]);
    hoja.getRange(filaActual, 1, 1, encabezados.length)
      .setFontWeight('bold')
      .setBackground('#d9d9d9')
      .setBorder(true, true, true, true, true, true);
    
    filaActual++;
    
    // Datos de facturas
    datos.facturas.forEach((factura, index) => {
      const tipoDoc = factura.tipoFacturaCliente || 'Factura';
      const esNotaCredito = tipoDoc === 'NC';
      const esNotaDebito = tipoDoc === 'ND';
      
      const baseImponible = parseFloat(factura.subTotal) || 0;
      const exentoFactura = parseFloat(factura.totalExento) || 0; // <--- LECTURA EXENTO
      const montoIVA = parseFloat(factura.iva) || 0;
      const totalFactura = parseFloat(factura.total) || 0;
      const porcentajeIVA = baseImponible > 0 ? ((montoIVA / baseImponible) * 100).toFixed(0) : '0';
      const retencion = parseFloat(factura.retencionIVA) || 0;
      const igtf = parseFloat(factura.igtf) || 0;
      
      const filaDatos = [
        index + 1,
        factura.fechaFacturaCliente,
        factura.rifCliente || '',
        factura.nombreCliente || '',
        tipoDoc,
        factura.nroFacturaCliente || '',
        factura.nroControlCliente || '',
        esNotaDebito ? factura.nroFacturaCliente : '',
        esNotaCredito ? factura.nroFacturaCliente : '',
        factura.nroComprobantes || '',
        '01-reg',
        '-',
        totalFactura,
        exentoFactura, // <--- PONER VALOR EXENTO EN LUGAR DE 0
        baseImponible,
        porcentajeIVA + '%',
        montoIVA,
        igtf,
        retencion
      ];
      
      hoja.getRange(filaActual, 1, 1, filaDatos.length).setValues([filaDatos]);
      hoja.getRange(filaActual, 1, 1, filaDatos.length).setBorder(true, true, true, true, true, true);
      
      hoja.getRange(filaActual, 13).setNumberFormat('#,##0.00'); // Total IVA
      hoja.getRange(filaActual, 14).setNumberFormat('#,##0.00'); // No gravadas
      hoja.getRange(filaActual, 15).setNumberFormat('#,##0.00'); // Base
      hoja.getRange(filaActual, 17).setNumberFormat('#,##0.00'); // IVA
      hoja.getRange(filaActual, 18).setNumberFormat('#,##0.00'); // IGTF
      hoja.getRange(filaActual, 19).setNumberFormat('#,##0.00'); // Retenido
      
      filaActual++;
    });
    
    filaActual += 2;
    
    // Tabla de totales
    hoja.getRange(filaActual, 4).setValue('Base Imponible');
    hoja.getRange(filaActual, 5).setValue('Débito fiscal');
    hoja.getRange(filaActual, 6).setValue('IVA retenido por el comprador');
    hoja.getRange(filaActual, 4, 1, 3).setFontWeight('bold').setBackground('#d9d9d9').setBorder(true, true, true, true, true, true);
    
    filaActual++;
    
    // Calcular totales reales usando el campo totalExento
    let totalNoGravadas = 0;
    let totalGravadas16 = 0;
    let totalIVA16 = 0;
    let totalRetenido16 = 0;
    
    datos.facturas.forEach(factura => {
      const base = parseFloat(factura.subTotal) || 0;
      const exento = parseFloat(factura.totalExento) || 0; 
      const iva = parseFloat(factura.iva) || 0;
      const retencion = parseFloat(factura.retencionIVA) || 0;
      
      totalNoGravadas += exento;
      totalGravadas16 += base;
      totalIVA16 += iva;
      totalRetenido16 += retencion;
    });
    
    const totalBase = totalNoGravadas + totalGravadas16;
    const totalIVA = totalIVA16;
    const totalRetenido = totalRetenido16;
    
    hoja.getRange(filaActual, 1).setValue('Total Ventas Internas No Gravadas');
    hoja.getRange(filaActual, 4).setValue(totalNoGravadas).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 5).setValue(0).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 6).setValue(0).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 1, 1, 6).setBorder(true, true, true, true, true, true);
    filaActual++;
    
    hoja.getRange(filaActual, 1).setValue('Total Ventas Internas afectadas sólo alícuota 16.00');
    hoja.getRange(filaActual, 4).setValue(totalGravadas16).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 5).setValue(totalIVA16).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 6).setValue(totalRetenido16).setNumberFormat('#,##0.00');
    hoja.getRange(filaActual, 1, 1, 6).setBorder(true, true, true, true, true, true);
    filaActual++;
    
    hoja.getRange(filaActual, 1).setValue('Totales:').setFontWeight('bold');
    hoja.getRange(filaActual, 4).setValue(totalBase).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 5).setValue(totalIVA).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 6).setValue(totalRetenido).setNumberFormat('#,##0.00').setFontWeight('bold');
    hoja.getRange(filaActual, 1, 1, 6).setBackground('#d9d9d9').setBorder(true, true, true, true, true, true);
    
    // Ajustar anchos de columnas
    hoja.setColumnWidth(1, 4);   hoja.setColumnWidth(2, 10);  hoja.setColumnWidth(3, 12);  hoja.setColumnWidth(4, 35);
    hoja.setColumnWidth(5, 8);   hoja.setColumnWidth(6, 10);  hoja.setColumnWidth(7, 12);  hoja.setColumnWidth(8, 10);
    hoja.setColumnWidth(9, 10);  hoja.setColumnWidth(10, 12); hoja.setColumnWidth(11, 10); hoja.setColumnWidth(12, 10);
    hoja.setColumnWidth(13, 16); hoja.setColumnWidth(14, 16); hoja.setColumnWidth(15, 12); hoja.setColumnWidth(16, 6);
    hoja.setColumnWidth(17, 12); hoja.setColumnWidth(18, 10); hoja.setColumnWidth(19, 12);
    
    const blob = libro.getBlob().setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);
    
    DriveApp.getFileById(libro.getId()).setTrashed(true);
    return base64;
    
  } catch (error) {
    throw new Error('Error al generar Excel: ' + error.message);
  }
}
function guardarNuevoProveedor(data) {
  // Función deprecada. Se mantiene como wrapper para compatibilidad con código antiguo.
  try {
    return guardarProveedor({
      id: 'P' + new Date().getTime(),
      nombre: data.nombre,
      rif: data.rif,
      direccion: data.direccion || '',
      correo: data.correo || '',
      telefono: data.telefono || '',
      tipoContribuyente: data.tipoContribuyente || 'Formal',
      retIVA: data.retIVA || '0',
      retISLR: data.retISLR || '0',
      status: 'Activo'
    });
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}
function calcularFechaVencimiento(fechaFactura, condPago) {
  try {
    if (!fechaFactura) return '';

    let base;
    if (fechaFactura instanceof Date) {
      base = new Date(fechaFactura.getTime());
    } else {
      const str = String(fechaFactura).trim();
      // Acepta yyyy-MM-dd o dd/MM/yyyy
      if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
        base = new Date(str.substring(0, 10) + 'T00:00:00');
      } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
        const [d, m, y] = str.split('/');
        base = new Date(`${y}-${m}-${d}T00:00:00`);
      } else {
        base = new Date(str);
      }
    }
    if (isNaN(base.getTime())) return '';

    const cond = String(condPago || 'Contado').trim();
    let dias = 0;

    if (cond === 'Contado') {
      dias = 0;
    } else if (cond === 'Credito15') {
      dias = 15;
    } else if (cond === 'Credito30') {
      dias = 30;
    } else if (cond === 'Credito45') {
      dias = 45;
    } else if (cond === 'Credito60' || cond === 'Credito60+') {
      dias = 60;
    } else {
      dias = 0;
    }

    base.setDate(base.getDate() + dias);
    return Utilities.formatDate(base, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    Logger.log('calcularFechaVencimiento error: ' + e.message);
    return '';
  }
}
function insertarFilaDinamica(sheet, objetoDatos) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    console.error("Error: La hoja está completamente vacía o no tiene encabezados.");
    return;
  }

  // Leer encabezados de la primera fila
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  
  // Crear la nueva fila con los datos en el orden correcto de los encabezados
  const nuevaFila = headers.map(header => {
    // Si el objeto tiene una propiedad con ese nombre de encabezado, usarla. Si no, dejar vacío.
    return objetoDatos.hasOwnProperty(header) ? objetoDatos[header] : "";
  });

  // Escribir la fila completa al final de la hoja
  sheet.appendRow(nuevaFila);
}
// ==============================================================================
/** ==== FUNCIONES PRINCIPALES PARA PDF Y CORREO ==============================*/
// ==============================================================================
function generarHtmlFacturaUnificado(datosFactura, paraCorreo = false) {
  
  const formatCurrency = (value) => {
    if (typeof value !== 'number') {
      value = parseFloat(value) || 0;
    }
    return value.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };

  let subtotalGravable = 0;
  let totalExento = 0;
  let totalIva = 0;

  datosFactura.items.forEach(item => {
    const montoBase = item.precioTotal || 0;
    const tasaIvaNum = item.tasaIva || 0;
    
    if (tasaIvaNum === 0) {
      totalExento += montoBase;
    } else {
      subtotalGravable += montoBase;
      totalIva += item.montoIva || 0;
    }
  });

  const totalFactura = subtotalGravable + totalExento + totalIva;

  const itemsHtml = datosFactura.items.map(item => `
    <tr>
      <td style="padding: 10px; border-bottom: 1px solid #ddd;">${item.producto}</td>
      <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: center;">${item.cantidad}</td>
      <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: right;">${formatCurrency(item.precio)}</td>
      <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: right;">${formatCurrency(item.precioTotal)}</td>
    </tr>`
  ).join('');

  // --- ADAPTACIÓN CON VARIABLES GLOBALES ---
  const encabezadoHtml = paraCorreo ? '' : `
    <tr style="vertical-align: top;">
      <td colspan="2" style="padding-bottom: 20px;">
        <img src="${LOGO_EMPRESA}" style="max-width: 150px; height: auto; margin-bottom: 10px;"><br>
        <h2 style="margin: 0; color: #333;">${NOM_EMPRESA}</h2>
        <b>RIF:</b> ${RIF_EMPRESA}<br>
        <small style="color: #666; display: block; max-width: 300px;">${DIR_EMPRESA}</small>
      </td>
      <td colspan="2" style="padding-bottom: 20px; text-align: right;">
        <h1 style="margin: 0; color: #007bff; font-size: 40px;">FACTURA</h1>
        <div style="margin-top: 10px;">
          <b>N° Factura:</b> <span style="color: #d9534f;">${datosFactura.nroFacturaCliente || 'N/A'}</span><br>
          <b>N° Control:</b> ${datosFactura.nroControlCliente || 'N/A'}<br>
          <b>Fecha:</b> ${datosFactura.fechaFacturaCliente || 'N/A'}
        </div>
      </td>
    </tr>`;

  const facturaHtml = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #555; }
      .invoice-box { max-width: 800px; margin: auto; padding: 30px; border: 1px solid #eee; box-shadow: 0 0 10px rgba(0, 0, 0, 0.15); font-size: 16px; line-height: 24px; background-color: #fff; }
      .invoice-box table { width: 100%; line-height: inherit; text-align: left; border-collapse: collapse; }
    </style>
  </head>
  <body>
    <div class="invoice-box">
      <table cellpadding="0" cellspacing="0">
        ${encabezadoHtml}
        <tr style="vertical-align: top;">
          <td colspan="4" style="padding-top: 20px; padding-bottom: 20px; border-top: 1px solid #eee;">
            <b style="color: #333; text-transform: uppercase; font-size: 13px;">Facturar a:</b><br>
            <strong style="font-size: 18px; color: #000;">${datosFactura.nombreCliente || 'N/A'}</strong><br>
            <b>RIF:</b> ${datosFactura.rifCliente || 'N/A'}<br>
            <b>Dirección:</b> ${datosFactura.DireccionCliente || 'N/A'}
          </td>
        </tr>
        <tr style="background-color: #007bff; font-weight: bold; color: #ffffff;">
          <td style="padding: 10px;">Descripción</td>
          <td style="padding: 10px; text-align: center;">Cant.</td>
          <td style="padding: 10px; text-align: right;">Precio Unit.</td>
          <td style="padding: 10px; text-align: right;">Total</td>
        </tr>
        ${itemsHtml}
      </table>
      
      <table style="width: 100%; margin-top: 30px;">
        <tr>
          <td style="width: 60%;"></td>
          <td style="width: 40%;">
            <table style="width: 100%;">
              <tr>
                <td style="text-align: right; padding: 5px;">Sub-Total Gravable:</td>
                <td style="text-align: right; width: 120px; padding: 5px;">${formatCurrency(subtotalGravable)} Bs.</td>
              </tr>
              <tr>
                <td style="text-align: right; padding: 5px;">Base Exenta:</td>
                <td style="text-align: right; padding: 5px;">${formatCurrency(totalExento)} Bs.</td>
              </tr>
              <tr>
                <td style="text-align: right; padding: 5px;">IVA:</td>
                <td style="text-align: right; padding: 5px;">${formatCurrency(totalIva)} Bs.</td>
              </tr>
              <tr style="font-weight: bold; font-size: 1.2em; color: #007bff;">
                <td style="text-align: right; padding: 10px; border-top: 2px solid #007bff;">TOTAL A PAGAR:</td>
                <td style="text-align: right; padding: 10px; border-top: 2px solid #007bff;">${formatCurrency(totalFactura)} Bs.</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
      <div style="margin-top: 50px; text-align: center; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 10px;">
        ${NOM_EMPRESA} - RIF: ${RIF_EMPRESA}<br>
        ${DIR_EMPRESA}
      </div>
    </div>
  </body>
  </html>`;
  
  return facturaHtml;
}
function generarPdfFacturaBase64(nroFactura) {
  try {
    const datosFactura = obtenerDatosCompletosFactura(nroFactura);
    if (!datosFactura) { 
      throw new Error(`No se encontró la factura N° ${nroFactura}`); 
    }

    const htmlCompleto = generarHtmlFacturaUnificado(datosFactura, false);
    const nroDoc = String(datosFactura.nroFacturaCliente || nroFactura).padStart(6, '0');
    const pdfBlob = Utilities.newBlob(htmlCompleto, MimeType.HTML, `Factura-${nroDoc}.html`).getAs(MimeType.PDF);
    pdfBlob.setName(`Factura-${nroDoc}.pdf`);

    return Utilities.base64Encode(pdfBlob.getBytes());
  } catch (error) {
    Logger.log("Error en generarPdfFacturaBase64: " + error.toString());
    throw new Error("No se pudo generar el PDF: " + error.message);
  }
}
function enviarFacturaPorCorreo(datosCorreo) {
  try {
    if (!datosCorreo || !datosCorreo.nroFactura) {
      throw new Error("Parámetros insuficientes: falta el número de factura.");
    }

    const datosFactura = obtenerDatosCompletosFactura(datosCorreo.nroFactura);
    if (!datosFactura) {
      throw new Error(`No se encontró la factura N° ${datosCorreo.nroFactura}.`);
    }

    const nroDoc = String(datosFactura.nroFacturaCliente || datosCorreo.nroFactura).padStart(6, '0');

    // 1. Crear PDF adjunto
    const htmlParaPdf = generarHtmlFacturaUnificado(datosFactura, false);
    const pdfBlob = Utilities.newBlob(htmlParaPdf, MimeType.HTML, `Factura-${nroDoc}.html`).getAs(MimeType.PDF);
    pdfBlob.setName(`Factura-${nroDoc}.pdf`);

    // 2. Crear cuerpo del correo
    const htmlParaCorreo = generarHtmlFacturaUnificado(datosFactura, true);

    // 3. Resolver destinatarios
    let destinatarios = new Set();
    if (datosFactura.correoCliente && String(datosFactura.correoCliente).includes('@')) {
      destinatarios.add(datosFactura.correoCliente.trim());
    }
    if (datosCorreo.destinatarioAdicional && String(datosCorreo.destinatarioAdicional).includes('@')) {
      destinatarios.add(datosCorreo.destinatarioAdicional.trim());
    }

    const listaDestinatarios = Array.from(destinatarios).join(',');
    if (!listaDestinatarios) {
      throw new Error("No se especificó un correo electrónico válido para el envío.");
    }
    
    const saludo = typeof getSaludoSegunHora === 'function' ? getSaludoSegunHora() : "Estimado(a)";
    const comentarios = datosCorreo.comentarios ? `<p><b>Nota adicional:</b><br>${datosCorreo.comentarios}</p>` : '';

    const cuerpoCorreoPrincipal = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <p>${saludo} <b>${datosFactura.nombreCliente}</b>,</p>
        <p>Adjunto a este correo encontrará su factura digital N° <b>${nroDoc}</b> en formato PDF.</p>
        ${comentarios}
        <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
        ${htmlParaCorreo}
      </div>`;

    const empresaData = datosFactura.empresaData || {};
    const remitente = empresaData.nombre || 'Departamento de Facturación';

    MailApp.sendEmail({
      to: listaDestinatarios,
      subject: datosCorreo.asunto || `Comprobante de Factura N° ${nroDoc}`,
      htmlBody: cuerpoCorreoPrincipal,
      attachments: [pdfBlob],
      name: remitente
    });
    
    return `Factura enviada exitosamente a: ${listaDestinatarios}`;

  } catch (error) {
    Logger.log("Error en enviarFacturaPorCorreo: " + error.toString());
    throw new Error("No se pudo enviar el correo: " + error.message);
  }
}
function enviarTicketPorEmail(params) {
  try {
    const { idPedido, html, emailDestino } = params;
    
    if (!emailDestino || !emailDestino.includes('@')) {
      return {
        status: 'error',
        message: '❌ Email no válido'
      };
    }
    
    MailApp.sendEmail({
      to: emailDestino,
      subject: `Ticket de Venta #${idPedido}`,
      htmlBody: html
    });
    
    return {
      status: 'success',
      message: `✅ Ticket enviado a ${emailDestino}`
    };
    
  } catch (error) {
    return {
      status: 'error',
      message: '❌ Error: ' + error.toString()
    };
  }
}
function sendPlanCuentasByEmail(params) {
  try {
    var to = params.to;
    var cc = params.cc || '';
    var subject = params.subject || 'Plan de Cuentas Contable';
    var description = params.description || '';
    var data = params.data || [];
    
    var htmlBody = '<h2>Plan de Cuentas Contable</h2>';
    htmlBody += '<p>Fecha: ' + new Date().toLocaleDateString('es-VE') + '</p>';
    if (description) htmlBody += '<p><em>' + description + '</em></p>';
    
    htmlBody += '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;font-family:Arial;font-size:12px">';
    htmlBody += '<thead style="background:#2980b9;color:white">';
    htmlBody += '<tr><th>Agrupación</th><th>Cuenta Contable</th><th>Cuenta</th><th>EEFF</th><th>Tipo</th><th>Tipo Cuenta</th><th>Grupo Balance</th></tr>';
    htmlBody += '</thead><tbody>';
    
    data.forEach(function(row) {
      htmlBody += '<tr>';
      htmlBody += '<td>' + (row.Agrupacion || row.agrupacion || '') + '</td>';
      htmlBody += '<td>' + (row.CuentaContable || row['Cuenta Contable'] || '') + '</td>';
      htmlBody += '<td>' + (row.Cuenta || row.cuenta || '') + '</td>';
      htmlBody += '<td style="text-align:center">' + (row.EEFF || row.eeff || '') + '</td>';
      htmlBody += '<td style="text-align:center">' + (row.Tipo || row.tipo || '') + '</td>';
      htmlBody += '<td style="text-align:center">' + (row.TipoCuenta || row['Tipo Cuenta'] || '') + '</td>';
      htmlBody += '<td>' + (row.GrupoBalance || row['Grupo Balance'] || '') + '</td>';
      htmlBody += '</tr>';
    });
    
    htmlBody += '</tbody></table>';
    htmlBody += '<p style="font-size:11px;color:#999;margin-top:20px">Generado por Sistema GPro</p>';
    
    GmailApp.sendEmail(to, subject, '', { htmlBody: htmlBody, cc: cc });
    
    return { success: true, message: 'Correo enviado a ' + to };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}
// ==============================================================================
/**  ====== FUNCIONES GENERADORAS DE HTML Y AUXILIARES ======================  */
// ==============================================================================

function generarHtmlCompletoFactura(facturaData, incluirBotonPdf = false) {
    const formatNumber = (num) => (parseFloat(num) || 0).toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    
    const fecha = facturaData.FechaFacturaCliente ? new Date(facturaData.FechaFacturaCliente).toLocaleDateString('es-VE') : 'N/A';
    const tasa = parseFloat(facturaData.TasaCliente) || 1;
    const subTotalUSD = (parseFloat(facturaData.SubTotalCliente) || 0) / tasa;
    const ivaUSD = (parseFloat(facturaData.IvaCliente) || 0) / tasa;
    const totalUSD = parseFloat(facturaData.TotalClienteUSD) || 0;
    
    const filasItems = (facturaData.items || []).map(item => {
        const precioUnitBs = (parseFloat(item.precio) || 0) * tasa;
        const montoIvaBs = (parseFloat(item.montoIva) || 0);
        const precioTotalBs = (parseFloat(item.precioTotal) || 0) * tasa;
        return `<tr><td>${item.cantidad}</td><td>${item.producto}</td><td class="text-right">${formatNumber(precioUnitBs)}</td><td class="text-right">${formatNumber(montoIvaBs)}</td><td class="text-right">${formatNumber(precioTotalBs)}</td></tr>`;
    }).join('');

    const botonPdfHtml = incluirBotonPdf ? `<div class="action-buttons"><button id="btnPdf" class="btn-pdf">Descargar PDF</button></div>` : '';
    const scriptPdfHtml = incluirBotonPdf ? `<script>document.getElementById('btnPdf').addEventListener('click',()=>{const e=document.getElementById('invoiceContent');const t={margin:.5,filename:\`Factura-${facturaData.NroFacturaCliente}.pdf\`,image:{type:'jpeg',quality:.98},html2canvas:{scale:2},jsPDF:{unit:'in',format:'letter',orientation:'portrait'}};html2pdf().set(t).from(e).save()});<\/script>` : '';

    return `<!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8"><title>Factura ${facturaData.NroFacturaCliente}</title>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.9.2/html2pdf.bundle.min.js"><\/script>
        <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 20px; background-color: #f4f7fc; color: #333; }
            .invoice-container { 
                max-width: 950px; margin: 20px auto; background: #fff; border-radius: 8px; 
                box-shadow: 0 4px 12px rgba(0,0,0,.08); padding: 40px;
                font-size: 10px;
            }
            .invoice-header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #007bff; padding-bottom: 20px; margin-bottom: 15px; }
            .company-details p { margin: 2px 0 0 0; }
            .logo-img { max-height: 60px; margin-bottom: 10px; }
            .invoice-header .invoice-details { text-align: right; }
            .invoice-header h2 { margin: 0; font-size: 1.8em; }
            .client-info { padding: 20px; border: 1px solid #eee; border-radius: 8px; margin-bottom: 20px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 10px; border-bottom: 1px solid #dee2e6; text-align: left; }
            th { background-color: #f8f9fa; }
            .text-right { text-align: right; }
            .invoice-bottom-section { display: grid; grid-template-columns: 220px 1fr; gap: 40px; margin-top: 30px; }
            .invoice-summary { padding: 10px; border: 1px solid #eee; border-radius: 8px; background-color: #f8f9fa; font-size: 0.85em; }
            .invoice-summary h4 { font-size: 0.9em; margin-top: 0; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
            .summary-item { display: flex; justify-content: space-between; margin-bottom: 6px; }
            .summary-item.total { font-size: 1em; padding-top: 8px; }
            .bcv-rate { margin-top: 15px; font-size: 0.75em; text-align: center; }
            .totals-section { display: flex; justify-content: flex-end; }
            .totals-table { width: 100%; max-width: 350px; }
            .totals-table .total-final td { font-size: 1.2em; font-weight: 700; color: #007bff; border-top: 2px solid #007bff; }
            .footer { text-align: center; margin-top: 40px; font-size: .9em; color: #888; }
            .action-buttons { text-align: center; padding: 20px 0; display: flex; justify-content: center; gap: 10px; }
            @media print { .action-buttons { display: none; } }
            .btn-pdf { background-color: #d32f2f; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }
        </style>
    </head>
    <body>
        ${botonPdfHtml}
        <div id="invoiceContent" class="invoice-container">
            <header class="invoice-header">
                <div>
                    <img src="${facturaData.logoEmpresa || ''}" alt="Logo" class="logo-img">
                    <div class="company-details">
                        <p><strong>${facturaData.nomEmpresa || ''}</strong></p>
                        <p>${facturaData.rifEmpresa || ''}</p>
                        <p>${facturaData.direccionEmpresa || ''}</p>
                    </div>
                </div>
                <div class="invoice-details">
                    <h2>FACTURA</h2>
                    <p><strong>Nro Factura:</strong> ${facturaData.NroFacturaCliente}</p>
                    <p><strong>Nro Control:</strong> ${facturaData.NroControl}</p>
                    <p><strong>Fecha:</strong> ${fecha}</p>
                </div>
            </header>
            <main>
                <div class="client-info">
                    <h4>Datos del Cliente</h4>
                    <p><strong>Nombre:</strong> ${facturaData.NombreCliente}</p>
                    <p><strong>RIF:</strong> ${facturaData.RifCliente}</p>
                    <p><strong>Dirección:</strong> ${facturaData.DireccionCliente}</p>
                </div>
                <table>
                    <thead><tr><th>Cant.</th><th>Descripción</th><th class="text-right">Precio Unit. (Bs.)</th><th class="text-right">IVA (Bs.)</th><th class="text-right">Total (Bs.)</th></tr></thead>
                    <tbody>${filasItems}</tbody>
                </table>
                <section class="invoice-bottom-section">
                    <aside class="invoice-summary">
                        <h4>Resumen ($)</h4>
                        <div class="summary-item"><span>Sub-Total:</span><span>$ ${formatNumber(subTotalUSD)}</span></div>
                        <div class="summary-item"><span>IVA:</span><span>$ ${formatNumber(ivaUSD)}</span></div>
                        <div class="summary-item total"><span>Total:</span><span>$ ${formatNumber(totalUSD)}</span></div>
                        <div class="bcv-rate">Tasa BCV:<br><strong>${formatNumber(tasa)} Bs.</strong></div>
                    </aside>
                    <div class="totals-section">
                        <table class="totals-table">
                            <tr><td>Sub-Total:</td><td class="text-right">${formatNumber(facturaData.SubTotalCliente)} Bs.</td></tr>
                            <tr><td>IVA:</td><td class="text-right">${formatNumber(facturaData.IvaCliente)} Bs.</td></tr>
                            <tr class="total-final"><td>TOTAL:</td><td class="text-right">${formatNumber(facturaData.TotalCliente)} Bs.</td></tr>
                        </table>
                    </div>
                </section>
            </main>
            <footer class="footer"><p>Gracias por su compra.</p></footer>
        </div>
        ${scriptPdfHtml}
    </body>
    </html>`;
}
function generarHtmlFactura(datosFactura) {
  let tablaHtml = `<style>table{width:100%;border-collapse:collapse}th,td{border:1px solid #ccc;padding:8px;font-size:12px}th{background-color:#f2f2f2;}.text-right{text-align:right}</style><table><thead><tr><th>Descripción</th><th class="text-right">Cantidad</th><th class="text-right">Total (USD)</th></tr></thead><tbody>`;
  datosFactura.items.forEach(item => { tablaHtml += `<tr><td>${item.producto}</td><td class="text-right">${item.cantidad}</td><td class="text-right">${parseFloat(item.precioTotal).toFixed(2)}</td></tr>`; });
  tablaHtml += `</tbody></table><h3 class="text-right">Total: ${parseFloat(datosFactura.TotalClienteUSD).toFixed(2)} USD</h3>`;
  return tablaHtml;
}
function getSaludoSegunHora() {
  const hora = new Date().getHours();
  if (hora >= 5 && hora < 12) { return "Buenos días"; } 
  else if (hora >= 12 && hora < 19) { return "Buenas tardes"; }
  else { return "Buenas noches"; }
}

// =============================================
/**             DASHBOARD                       */
// =============================================
function parseNumeroEspanolV2(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  
  let str = String(val).trim();
  // Limpiar espacios y símbolos monetarios
  str = str.replace(/[^0-9.,-]/g, '');
  
  if (str.includes(',') && str.includes('.')) {
    // Si tiene puntos de miles y coma decimal (estilo latinoamericano)
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      // Estilo US con comas de miles y punto decimal
      str = str.replace(/,/g, '');
    }
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}
function parsearFechaFlexible(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;

  if (typeof val === 'string') {
    val = val.trim();
    const partes = val.split(/[\/\-]/);
    if (partes.length === 3) {
      if (partes[0].length === 4) {
        // YYYY-MM-DD
        return new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
      } else {
        // DD/MM/YYYY
        return new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
      }
    }
  }

  const d = new Date(val);
  return !isNaN(d.getTime()) ? d : null;
}
function getTasaHistory() {
  try {
    var tasasSheet = getSheetByName("Tasas");
    
    if (!tasasSheet) throw new Error("No se encontró la hoja 'Tasas'");
    
    var dataRange = tasasSheet.getRange("G3:N14");
    var values = dataRange.getValues();
    
    var filteredValues = values.filter(row => row[4] > 0);
    
    if (filteredValues.length === 0) {
      return {
        summary: { ultima: 0, promedio: 0, maximo: 0, minimo: 0 },
        history: values
      };
    }
    
    var sum = 0;
    var max = 0;
    var min = Number.MAX_VALUE;
    
    for (var i = 0; i < filteredValues.length; i++) {
      var rate = filteredValues[i][4] || 0;
      sum += rate;
      if (rate > max) max = rate;
      if (rate < min && rate > 0) min = rate;
    }
    
    var avg = sum / filteredValues.length;
    var lastRate = filteredValues[filteredValues.length - 1][4];
    
    return {
      summary: {
        ultima: lastRate,
        promedio: avg,
        maximo: max,
        minimo: min
      },
      history: values
    };
  } catch (error) {
    return { error: error.toString() };
  }
}
function getTasaEuroHistory() {
  try {
    var tasasSheet = getSheetByName("Tasas");
    
    if (!tasasSheet) throw new Error("No se encontró la hoja 'Tasas'");
    
    var dataRange = tasasSheet.getRange("G3:N14");
    var values = dataRange.getValues();
    
    var filteredValues = values.filter(row => row[7] > 0);
    
    if (filteredValues.length === 0) {
      return {
        summary: { ultima: 0, promedio: 0, maximo: 0, minimo: 0 },
        history: values
      };
    }
    
    var sum = 0;
    var max = 0;
    var min = Number.MAX_VALUE;
    
    for (var i = 0; i < filteredValues.length; i++) {
      var rate = filteredValues[i][7] || 0;
      sum += rate;
      if (rate > max) max = rate;
      if (rate < min && rate > 0) min = rate;
    }
    
    var avg = sum / filteredValues.length;
    var lastRate = filteredValues[filteredValues.length - 1][7];
    
    return {
      summary: {
        ultima: lastRate,
        promedio: avg,
        maximo: max,
        minimo: min
      },
      history: values
    };
  } catch (error) {
    return { error: error.toString() };
  }
}
function getData() {
  let ws = getSheetByName("Items");
  let res = {};
  res.items = ws.getRange(2, 1, ws.getLastRow()-1, 5).getValues();
  ws = getSheetByName("Sales");
  res.sales = ws.getRange(2, 1, ws.getLastRow()-1, 6).getValues();
  Logger.log(res);
  return res;
}
function setData(data) {
  const importedData = JSON.parse(data);
  const orderLength = importedData.order.length;
  ws = getSheetByName("Sales");
  range = ws.getRange(ws.getLastRow()+1, 1, orderLength, 6);
  range.setValues(importedData.order);

  ws = getSheetByName("Payments");
  range = ws.getRange(ws.getLastRow()+1, 1, 1, 5);
  range.setValues(importedData.payment);
}
function getDataDash(){
  let ws = getSheetByName('Resumen');
  var dataDash = ws.getDataRange().getValues().slice(0);
  var result = dataDash.map(r=> r[2]);
  console.log(result);
  return result;
}
function getComprasPorMes() {
  try {
    const datos = actualizarResumenComprasPorMes();

    return datos.map(item => ({
      mes: item[0],
      bs: item[1],
      usd: item[2]
    }));
  } catch (error) {
    console.error("❌ Error en getComprasPorMes:", error);
    return [];
  }
}
function getVentasPorMes() {
  try {
    const datos = actualizarResumenVentasPorMes();

    return datos.map(item => ({
      mes: item[0],
      bs: item[1],
      usd: item[2]
    }));
  } catch (error) {
    console.error("❌ Error en getVentasPorMes:", error);
    return [];
  }
}
// =============================================
/**             DATOS CLIENTES                */
// =============================================
function obtenerClientes() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const clientes = [];
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      if (fila[0] === '' || fila[0] === null || fila[0] === undefined) continue;
      const cliente = {};
      headers.forEach((header, index) => { cliente[header] = fila[index]; });
      clientes.push(cliente);
    }
    return clientes;
  } catch (e) {
    Logger.log('Error en obtenerClientes: ' + e.toString());
    return [];
  }
}
function guardarCliente(clienteData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = _SS_(clienteData ? clienteData.idSpreadsheet : null);
    var sheet = ss.getSheetByName('Clientes');
    if (!sheet) throw new Error('Hoja Clientes no encontrada');

    var lastRow = sheet.getLastRow();
    var nuevoId = 'CLI-' + ('000' + lastRow).slice(-4);

    var nuevaFila = [
      nuevoId,
      clienteData.NombreClientes || '',
      clienteData.RifClientes || '',
      clienteData.DireccionCliente || '',
      clienteData.TelefonoCliente || '',
      clienteData.CorreoCliente || '',
      'Activo',
      clienteData.Vendedor || '',
      new Date()
    ];

    sheet.appendRow(nuevaFila);
    return { status: 'success', id: nuevoId, message: 'Cliente registrado correctamente' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
function actualizarCliente(cliente) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Clientes'); 
    if (!sheet) throw new Error("No se encontró la hoja Clientes");

    var data = sheet.getDataRange().getValues();
    var idBuscado = String(cliente.idClientes || cliente.id || '').trim();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === idBuscado) {
        var values = [
          idBuscado,
          cliente.NombreClientes || cliente.nombre || '',
          cliente.RifClientes || cliente.rif || '',
          cliente.DireccionCliente || cliente.direccion || '',
          cliente.CorreoCliente || cliente.correo || '',
          cliente.TelefonoCliente || cliente.telefono || '',
          cliente.TipoContribuyenteCliente || cliente.tipoContribuyente || 'Formal',
          cliente.RetIVACliente || cliente.retIVA || 0,
          cliente.RetISLRCliente || cliente.retISLR || 0,
          cliente.Responsable || cliente.responsable || 'Sin asignar',
          cliente.Vendedor || data[i][10] || 'Vendedor Móvil',
          cliente.Status || cliente.status || 'Activo'
        ];
        sheet.getRange(i + 1, 1, 1, values.length).setValues([values]);
        return { success: true, message: 'Cliente actualizado exitosamente' };
      }
    }
    return { success: false, message: 'Cliente no encontrado' };
  } catch (err) {
    Logger.log("Error en actualizarCliente: " + err.toString());
    return { success: false, message: err.message };
  }
}
function eliminarCliente(idClientes) {
  const sheet = _SS_().getSheetByName('Clientes'); 
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == idClientes) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Cliente eliminado' };
    }
  }
  return { success: false, message: 'Cliente no encontrado' };
}
function getSheetByNameSeguro(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  return sheet;
}
function verificarValorDuplicado(sheetName, columnHeaderName, valueToCheck, excludeId) {
  try {
    if (!valueToCheck) return false;
    const sheet = getSheetByNameSeguro(sheetName);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return false;

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const targetColIdx = headers.indexOf(columnHeaderName.toLowerCase());
    const idColIdx = 0;

    if (targetColIdx === -1) return false;

    const valorNormalizado = String(valueToCheck).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][idColIdx]);
      const rowValue = String(data[i][targetColIdx]).trim().toLowerCase();

      // Si coincide el valor y no es el mismo registro que estamos editando
      if (rowValue === valorNormalizado && (!excludeId || rowId !== String(excludeId))) {
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log('Error en verificarValorDuplicado: ' + e.toString());
    return false;
  }
}
function getDataClientes() {
  try {
    var sheet = getSheetByName("Clientes");
    return sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
  } catch (e) {
    console.error("Error en getDataClientes: " + e.toString());
    return [];
  }
}
function getDataListaClientes() {
  const sheetClientes = getSheetByName("Clientes");
  const clientesData = sheetClientes.getDataRange().getDisplayValues();
  const clientesHeaders = clientesData.shift();
  
  let clientesDataArray = [];

  clientesData.forEach(dataRow => {
    const temporaryObject = {};
    clientesHeaders.forEach((headerTitle, index) => {
      temporaryObject[headerTitle] = dataRow[index];
    });
    clientesDataArray.push(temporaryObject);
  });

  return clientesDataArray;
}
function obtenerClientesParaPedidos() {
  try {
    const sheetClientes = getSheetByName('Clientes');
    if (!sheetClientes) throw new Error("La hoja 'Clientes' no fue encontrada.");
    
    const values = sheetClientes.getDataRange().getValues();
    const headers = values.shift();
    
    const getIndex = (name) => {
      const index = headers.findIndex(header => 
        header.toString().trim().toLowerCase() === name.toLowerCase()
      );
      return index !== -1 ? index : 
        headers.findIndex(header => header.toString().includes(name));
    };
    
    const idIndex = getIndex('idClientes');
    const nombreIndex = getIndex('Nombre') || getIndex('Cliente') || getIndex('Razon Social');
    const rifIndex = getIndex('RIF') || getIndex('Rif');
    const emailIndex = getIndex('Email') || getIndex('Correo') || getIndex('Mail');
    const telefonoIndex = getIndex('Teléfono') || getIndex('Telefono') || getIndex('Phone');
    const direccionIndex = getIndex('Dirección') || getIndex('Direccion') || getIndex('Address');
    
    const clientes = values.map(row => ({
      idClientes: row[idIndex] || '',
      Nombre: row[nombreIndex] || row[rifIndex] || 'Cliente',
      RIF: row[rifIndex] || '',
      Email: row[emailIndex] || '',
      Telefono: row[telefonoIndex] || '',
      Direccion: row[direccionIndex] || ''
    })).filter(cliente => cliente.idClientes && cliente.idClientes.toString().trim() !== "");
    
    return clientes;
  } catch (e) {
    Logger.log('Error en obtenerClientesParaPedidos: ' + e.toString());
    return [];
  }
}
function verificarRifClienteExistente(rif) {
  try {
    const sheet = getSheetByName(SHEET_NAME_CLIENTES);
    const rifColumnIndex = 3;
    const rifValues = sheet.getRange(2, rifColumnIndex, sheet.getLastRow(), 1).getValues();
    return rifValues.some(row => row[0] && row[0].toUpperCase().trim() === rif.toUpperCase().trim());
  } catch (e) {
    Logger.log('Error al verificar RIF de cliente: ' + e.toString());
    return false;
  }
}
function verificarCorreoClienteExistente(correo) {
  try {
    const sheet = getSheetByName(SHEET_NAME_CLIENTES);
    const correoColumnIndex = 5;
    const correoValues = sheet.getRange(2, correoColumnIndex, sheet.getLastRow(), 1).getValues();
    return correoValues.some(row => row[0] && row[0].toLowerCase().trim() === correo.toLowerCase().trim());
  } catch (e) {
    Logger.log('Error al verificar correo de cliente: ' + e.toString());
    return false;
  }
}
function getDataFacturas() {
  const sheetfacturas = getSheetByName("Compras");
  const facturasData = sheetfacturas.getDataRange().getDisplayValues();
  const facturasHeaders = facturasData.shift();
  
  let facturasDataArray = [];

  facturasData.forEach( dataRow =>{
    const temporaryObject = {};
    facturasHeaders.forEach((headerTitle, index) =>{
      temporaryObject[headerTitle] = dataRow[index];
    })
    facturasDataArray.push(temporaryObject);
  })

  return facturasDataArray;
}
function obtenerClientesConSaldos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetClientes = ss.getSheetByName('Clientes');
    const sheetCobros = ss.getSheetByName('CobrosPagos');

    if (!sheetClientes) {
      throw new Error("No se encontró la hoja 'Clientes'");
    }

    // 1. Obtener Datos de Clientes
    const dataClientes = sheetClientes.getDataRange().getValues();
    if (dataClientes.length <= 1) return [];

    const headersCli = dataClientes[0].map(h => String(h).trim().toLowerCase());
    const idxCli = {
      id: headersCli.indexOf('idclientes'),
      nombre: headersCli.indexOf('nombreclientes'),
      rif: headersCli.indexOf('rifclientes'),
      direccion: headersCli.indexOf('direccioncliente'),
      correo: headersCli.indexOf('correocliente'),
      telefono: headersCli.indexOf('telefonocliente'),
      tipo: headersCli.indexOf('tipocontribuyentecliente'),
      retIVA: headersCli.indexOf('retivacliente'),
      retISLR: headersCli.indexOf('retislrcliente'),
      responsable: headersCli.indexOf('responsable'),
      status: headersCli.indexOf('status')
    };

    // Función auxiliar para normalizar RIFs (ej: "J-12345678-9" -> "J123456789")
    const limpiarRif = (rif) => String(rif || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

    // 2. Mapear Saldos Pendientes desde CobrosPagos
    const saldosMap = {}; // Clave: RIF normalizado -> { saldoBs, saldoUSD, pendientes }

    if (sheetCobros) {
      const dataCobros = sheetCobros.getDataRange().getValues();
      if (dataCobros.length > 1) {
        const headersCobros = dataCobros[0].map(h => String(h).trim().toLowerCase());
        const idxCob = {
          rif: headersCobros.indexOf('rif'),
          saldoBs: headersCobros.indexOf('saldofactura'),
          saldoUSD: headersCobros.indexOf('saldofacturausd'),
          estado: headersCobros.indexOf('cobradapagada')
        };

        for (let i = 1; i < dataCobros.length; i++) {
          const row = dataCobros[i];
          const rifCobro = limpiarRif(row[idxCob.rif]);
          const estado = String(row[idxCob.estado] || '').trim().toLowerCase();

          // Solo sumar si no está completamente pagada
          if (rifCobro && estado !== 'pagada') {
            const parseMonto = (val) => {
              if (typeof val === 'number') return val;
              if (!val) return 0;
              // Manejo de formato latinoamericano: 1.234,56 -> 1234.56
              const normalizado = String(val).replace(/\./g, '').replace(',', '.');
              const num = parseFloat(normalizado);
              return isNaN(num) ? 0 : num;
            };

            const saldoBs = parseMonto(row[idxCob.saldoBs]);
            const saldoUSD = parseMonto(row[idxCob.saldoUSD]);

            if (saldoBs > 0 || saldoUSD > 0) {
              if (!saldosMap[rifCobro]) {
                saldosMap[rifCobro] = { saldoBs: 0, saldoUSD: 0, facturasPendientes: 0 };
              }
              saldosMap[rifCobro].saldoBs += saldoBs;
              saldosMap[rifCobro].saldoUSD += saldoUSD;
              saldosMap[rifCobro].facturasPendientes += 1;
            }
          }
        }
      }
    }

    // 3. Cruzar información de Clientes con sus Saldos
    const clientes = [];
    for (let i = 1; i < dataClientes.length; i++) {
      const row = dataClientes[i];
      const rawRif = row[idxCli.rif];
      const rifKey = limpiarRif(rawRif);
      const saldos = saldosMap[rifKey] || { saldoBs: 0, saldoUSD: 0, facturasPendientes: 0 };

      clientes.push({
        idClientes: String(row[idxCli.id] || ''),
        NombreClientes: String(row[idxCli.nombre] || ''),
        RifClientes: String(rawRif || ''),
        DireccionCliente: String(row[idxCli.direccion] || ''),
        CorreoCliente: String(row[idxCli.correo] || ''),
        TelefonoCliente: String(row[idxCli.telefono] || ''),
        TipoContribuyenteCliente: String(row[idxCli.tipo] || 'Formal'),
        RetIVACliente: String(row[idxCli.retIVA] || '0'),
        RetISLRCliente: String(row[idxCli.retISLR] || '0'),
        Responsable: String(row[idxCli.responsable] || ''),
        Status: String(row[idxCli.status] || 'Activo'),
        // Datos cruzados de cobranza:
        saldoPendienteBs: Number(saldos.saldoBs.toFixed(2)),
        saldoPendienteUSD: Number(saldos.saldoUSD.toFixed(2)),
        facturasPendientes: saldos.facturasPendientes
      });
    }

    return clientes;
  } catch (error) {
    throw new Error("Error en obtenerClientesConSaldos: " + error.message);
  }
}
function obtenerDocumentosPendientesPorCliente(rifCliente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCobros = ss.getSheetByName('CobrosPagos');
    
    if (!sheetCobros) {
      throw new Error("No se encontró la hoja 'CobrosPagos'");
    }

    const displayData = sheetCobros.getDataRange().getDisplayValues();
    if (displayData.length <= 1) return [];

    const headers = displayData[0].map(h => String(h).trim().toLowerCase());
    const idxRif = headers.indexOf('rif');
    const idxNroFactura = headers.indexOf('nrofactura');
    const idxFecha = headers.indexOf('fechafactura');
    const idxTotal = headers.indexOf('totafactura');
    const idxSaldoBs = headers.indexOf('saldofactura');
    const idxSaldoUSD = headers.indexOf('saldofacturausd');
    const idxEstado = headers.indexOf('cobradapagada');
    const idxMoneda = headers.indexOf('moneda');
    const idxDescripcion = headers.indexOf('descripcion');
    const idxVencimiento = headers.indexOf('fechavencimiento');

    const rifLimpio = String(rifCliente || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    const documentos = [];

    for (let i = 1; i < displayData.length; i++) {
      const row = displayData[i];
      if (!row[idxRif]) continue;
      
      const rifRow = String(row[idxRif] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      const estado = String(row[idxEstado] || '').trim().toLowerCase();

      if (rifRow === rifLimpio && (estado === 'porcobrar' || estado === 'por cobrar')) {
        const parseMonto = (val) => {
          if (typeof val === 'number') return val;
          if (!val) return 0;
          const normalizado = String(val).replace(/\./g, '').replace(',', '.');
          const num = parseFloat(normalizado);
          return isNaN(num) ? 0 : num;
        };

        documentos.push({
          nroFactura: String(row[idxNroFactura] || ''),
          fechaFactura: String(row[idxFecha] || ''),
          totalFactura: parseMonto(row[idxTotal]),
          saldoBs: parseMonto(row[idxSaldoBs]),
          saldoUSD: parseMonto(row[idxSaldoUSD]),
          moneda: String(row[idxMoneda] || 'Dolares'),
          descripcion: String(row[idxDescripcion] || ''),
          fechaVencimiento: String(row[idxVencimiento] || ''),
          estado: estado
        });
      }
    }

    return documentos;
  } catch (error) {
    Logger.log("Error en obtenerDocumentosPendientesPorCliente: " + error.message);
    throw new Error("Error en obtenerDocumentosPendientesPorCliente: " + error.message);
  }
}

// =============================================
/**  PROVEEDORES                    */
// =============================================
function getDataProveedores(idSpreadsheet) {
  try {
    var ss = _resolverSSCompras_(idSpreadsheet);
    if (!ss) return [];
    var sheet = ss.getSheetByName("Proveedores");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var provs = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nom = String(row[1] || row[0] || '').trim();
      if (!nom) continue;
      
      // Formato estricto: [ID, NOMBRE, RIF, DIRECCION, RET_IVA, RET_ISLR]
      provs.push([
        String(row[0] || '').trim(),       // 0: ID
        nom,                               // 1: NOMBRE
        String(row[2] || '').trim(),       // 2: RIF
        String(row[3] || '').trim(),       // 3: DIRECCION
        String(row[7] || '0').trim(),      // 4: RET_IVA (%)
        String(row[8] || '0').trim()       // 5: RET_ISLR (%)
      ]);
    }
    return provs;
  } catch(e) {
    Logger.log("Error en getDataProveedores: " + e);
    return [];
  }
}
function normalizarRif(rif) {
  return String(rif || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}
function parsearMonto(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  // Limpia miles "." y pasa coma decimal "," a punto "."
  const normalizado = String(val).trim().replace(/\./g, '').replace(',', '.');
  const num = parseFloat(normalizado);
  return isNaN(num) ? 0 : num;
}
function obtenerProveedoresConSaldos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetProv = ss.getSheetByName('Proveedores');
    const sheetCobros = ss.getSheetByName('CobrosPagos');

    if (!sheetProv) throw new Error("No existe la hoja 'Proveedores'");

    const saldosMap = {};

    if (sheetCobros) {
      const dataCobros = sheetCobros.getDataRange().getValues();
      for (let i = 1; i < dataCobros.length; i++) {
        const fila = dataCobros[i];
        const rifRaw = fila[1]; 
        if (!rifRaw) continue;

        const rifLimpio = normalizarRif(rifRaw);
        const estado = String(fila[7] || '').trim().toLowerCase(); 
        const transaccion = String(fila[8] || '').trim().toLowerCase(); 

        // Acepta 'pago' y 'compra'
        if ((transaccion === 'pago' || transaccion === 'compra') && estado !== 'pagada' && estado !== 'pagado') {
          const saldoBs = parsearMonto(fila[5]);   
          const saldoUSD = parsearMonto(fila[6]);  

          if (!saldosMap[rifLimpio]) {
            saldosMap[rifLimpio] = { saldoBs: 0, saldoUSD: 0, facturas: 0 };
          }

          saldosMap[rifLimpio].saldoBs += saldoBs;
          saldosMap[rifLimpio].saldoUSD += saldoUSD;
          saldosMap[rifLimpio].facturas += 1;
        }
      }
    }

    const dataProv = sheetProv.getDataRange().getValues();
    const proveedores = [];

    for (let i = 1; i < dataProv.length; i++) {
      const fila = dataProv[i];
      const id = String(fila[0] !== undefined ? fila[0] : '').trim();
      const nombre = String(fila[1] !== undefined ? fila[1] : '').trim();

      if (!id && !nombre) continue;

      const rifOriginal = String(fila[2] !== undefined ? fila[2] : '').trim();
      const rifLimpio = normalizarRif(rifOriginal);

      const saldoInfo = saldosMap[rifLimpio] || { saldoBs: 0, saldoUSD: 0, facturas: 0 };

      proveedores.push({
        IdProveedor: id,
        NombreProveedor: nombre,
        RifProveedor: rifOriginal,
        DireccionProveedor: String(fila[3] || '').trim(),
        CorreoProveedor: String(fila[4] || '').trim(),
        TelefonoProveedor: String(fila[5] || '').trim(),
        TipoContribuyente: String(fila[6] || 'Formal').trim(),
        RetIVAProveedor: String(fila[7] || '0').trim(),
        RetISLRProveedor: String(fila[8] || '0').trim(),
        Status: String(fila[9] || 'Activo').trim(),
        saldoPendienteBs: saldoInfo.saldoBs,
        saldoPendienteUSD: saldoInfo.saldoUSD,
        facturasPendientes: saldoInfo.facturas
      });
    }

    return proveedores;

  } catch (err) {
    Logger.log("Error: " + err.toString());
    throw new Error(err.message);
  }
}
function obtenerFacturasPendientesPorProveedor(rifProveedor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCobros = ss.getSheetByName('CobrosPagos');
    if (!sheetCobros) return [];

    const dataCobros = sheetCobros.getDataRange().getValues();
    const rifBuscado = normalizarRif(rifProveedor);
    const facturas = [];

    for (let i = 1; i < dataCobros.length; i++) {
      const fila = dataCobros[i];
      const rifFila = normalizarRif(fila[1]); 
      const transaccion = String(fila[8] || '').trim().toLowerCase(); 
      const estado = String(fila[7] || '').trim().toLowerCase(); 

      // Acepta 'pago' y 'compra'
      if (rifFila === rifBuscado && (transaccion === 'pago' || transaccion === 'compra') && estado !== 'pagada' && estado !== 'pagado') {
        let fecha = fila[3]; 
        if (fecha instanceof Date) {
          fecha = Utilities.formatDate(fecha, Session.getScriptTimeZone() || "GMT-4", "dd/MM/yyyy");
        }

        facturas.push({
          nroFactura: String(fila[2] || ''),       
          fechaFactura: String(fecha || ''),       
          totalFactura: parsearMonto(fila[4]),     
          saldoBs: parsearMonto(fila[5]),          
          saldoUSD: parsearMonto(fila[6]),         
          estado: fila[7]                          
        });
      }
    }

    return facturas;
  } catch (e) {
    throw new Error(e.message);
  }
}
function guardarProveedor(proveedorData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Proveedores');
    if (!sheet) throw new Error("No se encontró la hoja 'Proveedores'");

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    
    const idxNombre = headers.indexOf('nombreproveedor');
    const idxRif = headers.indexOf('rifproveedor');
    const rifNorm = normalizarRif(proveedorData.rif);
    const nomNorm = String(proveedorData.nombre || '').toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      if (nomNorm && String(data[i][idxNombre] || '').toLowerCase().trim() === nomNorm) {
        throw new Error("Ya existe un proveedor con el nombre: " + proveedorData.nombre);
      }
      if (rifNorm && normalizarRif(data[i][idxRif]) === rifNorm) {
        throw new Error("Ya existe un proveedor con el RIF: " + proveedorData.rif);
      }
    }

    sheet.appendRow([
      proveedorData.id,
      proveedorData.nombre,
      proveedorData.rif,
      proveedorData.direccion,
      proveedorData.correo,
      proveedorData.telefono,
      proveedorData.tipoContribuyente,
      proveedorData.retIVA,
      proveedorData.retISLR,
      proveedorData.status
    ]);

    return { success: true, message: "Proveedor guardado" };
  } catch (error) {
    throw new Error(error.message);
  }
}
function actualizarProveedor(proveedorData) {
  try {
    Logger.log("✏️ Actualizando proveedor: " + proveedorData.id);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Proveedores');
    if (!sheet) throw new Error("No se encontró la hoja 'Proveedores'");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error("La hoja 'Proveedores' está vacía");

    // Índices de columnas según tu hoja:
    // [0] IdProveedor, [1] NombreProveedor, [2] RifProveedor, [3] DireccionProveedor,
    // [4] CorreoProveedor, [5] TelefonoProveedor, [6] TipoContribuyenteProveedor,
    // [7] RetIVAProveedor, [8] RetISLRProveedor, [9] Status

    const idBuscado = String(proveedorData.id || '').trim();
    const rifNormalizado = normalizarRif(proveedorData.rif);
    const nombreLower = String(proveedorData.nombre || '').trim().toLowerCase();

    let filaIndex = -1;

    // 1. Buscar la fila del proveedor a actualizar
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === idBuscado) {
        filaIndex = i + 1; // Fila en Sheets (base 1)
        break;
      }
    }

    if (filaIndex === -1) {
      throw new Error("No se encontró el proveedor con ID: " + proveedorData.id);
    }

    // 2. Validar que no exista OTRO proveedor con el mismo Nombre o RIF
    for (let j = 1; j < data.length; j++) {
      const filaSheet = j + 1;
      if (filaSheet === filaIndex) continue; // Ignorar el registro que estamos editando

      const nombreExistente = String(data[j][1] || '').trim().toLowerCase();
      const rifExistente = normalizarRif(data[j][2]);

      if (nombreLower && nombreExistente === nombreLower) {
        throw new Error("Ya existe otro proveedor con el nombre: " + proveedorData.nombre);
      }
      if (rifNormalizado && rifExistente === rifNormalizado) {
        throw new Error("Ya existe otro proveedor con el RIF: " + proveedorData.rif);
      }
    }

    // 3. Escribir los datos actualizados en la fila exacta (Columnas 1 a 10)
    sheet.getRange(filaIndex, 1, 1, 10).setValues([[
      proveedorData.id,
      proveedorData.nombre,
      proveedorData.rif,
      proveedorData.direccion,
      proveedorData.correo,
      proveedorData.telefono,
      proveedorData.tipoContribuyente,
      proveedorData.retIVA,
      proveedorData.retISLR,
      proveedorData.status
    ]]);

    Logger.log("✅ Proveedor " + proveedorData.id + " actualizado con éxito");
    return { success: true, message: "Proveedor actualizado exitosamente" };

  } catch (error) {
    Logger.log("❌ Error en actualizarProveedor: " + error.message);
    throw new Error(error.message);
  }
}
function eliminarProveedor(idProveedor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Proveedores');
    if (!sheet) throw new Error("No se encontró la hoja 'Proveedores'");

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxId = headers.indexOf('idproveedor');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxId]).trim() === String(idProveedor).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Proveedor eliminado" };
      }
    }
    throw new Error("No se encontró el proveedor");
  } catch (error) {
    throw new Error(error.message);
  }
}
function getSheetSafe(name) {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("No se pudo obtener el active spreadsheet: " + e.toString());
  }
  
  if (!ss) {
    throw new Error("No se pudo vincular el Script a la Hoja de Cálculo activa. Asegúrate de que este script esté vinculado (Extensiones -> Apps Script dentro de la hoja de cálculo).");
  }
  
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  
  // Buscar variantes de nombres en caso de que la hoja esté en singular o minúsculas
  const alternates = [
    name,
    "Proveedores",
    "Proveedor",
    "proveedores",
    "proveedor"
  ];
  for (let i = 0; i < alternates.length; i++) {
    sheet = ss.getSheetByName(alternates[i]);
    if (sheet) return sheet;
  }
  
  // Si sigue sin encontrarse, usar la primera pestaña de la hoja de cálculo
  const sheets = ss.getSheets();
  if (sheets.length > 0) {
    return sheets[0];
  }
  return null;
}
function obtenerRolUsuario() {
  try {
    var usuario = Session.getActiveUser().getEmail();
    var adminUsers = ['fmorenocontabilidad@gmail.com', 'fjnmorenod@gmail.com'];
    var consultaUsers = [];
    
    var owner = "";
    try {
      owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
    } catch(e) {
      Logger.log("getOwner no disponible: " + e.toString());
    }

    if (usuario === owner && owner !== "") {
      return 'owner';
    }
    if (adminUsers.includes(usuario)) {
      return 'admin';
    }
    if (consultaUsers.includes(usuario)) {
      return 'consulta';
    }
    return 'usuario';
  } catch (error) {
    Logger.log("Error en obtenerRolUsuario: " + error.toString());
    return 'usuario';
  }
}
function obtenerProveedores() {
  try {
    const sheetProveedor = getSheetSafe('Proveedores');
    if (!sheetProveedor) throw new Error("No se pudo hallar la hoja de cálculo requerida.");
    
    const values = sheetProveedor.getDataRange().getValues();
    if (values.length <= 1) return []; 
    
    const headers = values.shift();
    
    const proveedores = values.map(row => {
      const proveedorObj = {};
      headers.forEach((header, index) => {
        if (header) {
          let val = row[index];
          // Evitar errores de serialización transformando tipos de datos complejos a texto plano
          try {
            if (val instanceof Date) {
              val = val.toISOString();
            } else if (typeof val === 'object' && val !== null) {
              val = val.toString();
            }
          } catch(err) {
            val = "";
          }
          proveedorObj[header] = val;
        }
      });
      return proveedorObj;
    }).filter(proveedor => {
      return proveedor.IdProveedor && 
             proveedor.IdProveedor.toString().trim() !== "" && 
             proveedor.Status !== "Inactivo";
    });
    
    return proveedores;
  } catch (e) {
    Logger.log('Error en obtenerProveedores: ' + e.toString());
    throw new Error('No se pudo acceder a los datos de los proveedores: ' + e.message);
  }
}
function enviarTablaPorCorreo(mensaje, destinatario, asunto) {
  try {
    if (!destinatario || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(destinatario.trim())) {
      throw new Error("El destinatario no es válido.");
    }
    
    MailApp.sendEmail({
      to: destinatario.trim(),
      subject: asunto || "Reporte de Proveedores",
      htmlBody: mensaje,
      name: "Sistema de Gestión TaxFlow"
    });
    
    return { success: true, message: "Correo enviado exitosamente." };
  } catch (e) {
    Logger.log("Error en enviarTablaPorCorreo: " + e.toString());
    return { success: false, message: e.toString() };
  }
}
function guardarIdProveedorEdicion(id) {
  PropertiesService.getDocumentProperties().setProperty('editingProveedorId', id);
  return true;
}
function buscarProveedoresPorTermino(termino) {
  try {
    var proveedores = obtenerProveedores();
    termino = termino.toLowerCase();
    
    var resultados = proveedores.filter(function(proveedor) {
      return (
        (proveedor.NombreProveedor && proveedor.NombreProveedor.toLowerCase().includes(termino)) ||
        (proveedor.RifProveedor && proveedor.RifProveedor.toLowerCase().includes(termino)) ||
        (proveedor.DireccionProveedor && proveedor.DireccionProveedor.toLowerCase().includes(termino)) ||
        (proveedor.CorreoProveedor && proveedor.CorreoProveedor.toLowerCase().includes(termino)) ||
        (proveedor.TelefonoProveedor && proveedor.TelefonoProveedor.toLowerCase().includes(termino))
      );
    });
    
    return resultados;
  } catch (error) {
    Logger.log("Error en buscarProveedoresPorTermino: " + error.toString());
    throw new Error("No se pudo realizar la búsqueda: " + error.message);
  }
}
function verificarRifExistente(rif) {
  try {
    var sheet = getSheetByName("Proveedores");
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rifIndex = headers.indexOf("RifProveedor");
    
    var idActual = "";
    try {
      idActual = PropertiesService.getDocumentProperties().getProperty('editingProveedorId') || "";
    } catch (e) {
      console.error("Error obteniendo ID actual:", e);
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][rifIndex] && data[i][rifIndex].toString().toUpperCase().trim() === rif && 
          data[i][headers.indexOf("IdProveedor")].toString() !== idActual) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    Logger.log("Error en verificarRifExistente: " + error.toString());
    throw new Error("No se pudo verificar el RIF: " + error.message);
  }
}
function verificarCorreoExistente(correo) {
  try {
    const sheet = getSheetByName("Proveedores");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const correoIndex = headers.indexOf("CorreoProveedor");
    const idIndex = headers.indexOf("IdProveedor");
    
    const idActual = PropertiesService.getDocumentProperties().getProperty('editingProveedorId') || "";

    for (var i = 1; i < data.length; i++) {
      if (data[i][correoIndex] && data[i][correoIndex].toString().toLowerCase().trim() === correo.toLowerCase().trim() && 
          data[i][idIndex].toString() !== idActual) {
        return true;
      }
    }
    return false;
  } catch (e) {
    return false; 
  }
}
function verificarUsernameExistente(username) {
  const sheetUsuarios = getSheetByName('Usuarios');
  const data = sheetUsuarios.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === username.toLowerCase()) {
      return true;
    }
  }
  
  return false;
}
function guardarDatosFactura1(facturaData) {
  try {
    let facturasSheet;

    if (facturaData.tipoTransaccion === 'Compra') {
      facturasSheet = getSheetByName("Compras");
    } else if (facturaData.tipoTransaccion === 'Venta') {
      facturasSheet = getSheetByName("Ventas");
    } else {
      throw new Error(`El tipo de transacción '${facturaData.tipoTransaccion}' no es válido.`);
    }

    if (!facturasSheet) {
      throw new Error(`No se pudo encontrar la hoja de destino para '${facturaData.tipoTransaccion}'.`);
    }

    const detalleSheet = getSheetByName("DetalleFacturas");
    if (!detalleSheet) {
      throw new Error("No se pudo encontrar la hoja 'DetalleFacturas'.");
    }
    
    const uniqueId = new Date().getTime() + '-' + facturaData.nroFacturaCliente;

    facturasSheet.appendRow([
      facturaData.nombreCliente, facturaData.rifCliente, facturaData.fechaFacturaCliente,
      facturaData.tipoFacturaCliente, facturaData.nroFacturaCliente, facturaData.nroControlCliente,
      facturaData.selectMonedaVenta, facturaData.monedaTasaVenta, facturaData.subTotal,
      facturaData.iva, facturaData.total, facturaData.saldoCliente, facturaData.totalClienteUSD,
      facturaData.saldoClienteUSD, facturaData.tipoTransaccion, uniqueId,
      facturaData.cobradaPagada
    ]);

    const detallesParaGuardar = facturaData.items.map(item => [
      item.factura, item.producto, item.nombre, item.cantidad, item.precio,
      item.tasaIva, item.montoIva, item.precioTotal, facturaData.tipoTransaccion,
    ]);

    if (detallesParaGuardar.length > 0) {
      const RangoDetalle = detalleSheet.getRange(detalleSheet.getLastRow() + 1, 1, detallesParaGuardar.length, detallesParaGuardar[0].length);
      RangoDetalle.setValues(detallesParaGuardar);
    }
    
    actualizarStockProductos(facturaData.items, facturaData.tipoTransaccion);

    if (facturaData.saldoCliente > 0) {
      const cobrosPagosSheet = getSheetByName("CobrosPagos");
      if (cobrosPagosSheet) {
        cobrosPagosSheet.appendRow([
          facturaData.nombreCliente, facturaData.rifCliente, facturaData.nroFacturaCliente,
          facturaData.fechaFacturaCliente, facturaData.total, facturaData.saldoCliente,facturaData.saldoClienteUSD,
          facturaData.cobradaPagada, facturaData.tipoTransaccion === 'Compra' ? 'Pago' : 'Cobro',
          uniqueId
        ]);
      } else {
        Logger.log("Advertencia: No se encontró la hoja 'CobrosPagos'.");
      }
    }

    return {
      message: `Factura ${facturaData.nroFacturaCliente} guardada y stock actualizado.`,
      empresaData: getEmpresaData()
    };

  } catch (e) {
    Logger.log(e.stack);
    throw new Error('Error en el servidor: ' + e.message);
  }
}
function obtenerFacturasPendientes(idProveedor) {
  try {
    const sheetCompras = getSheetSafe("Compras");
    if (!sheetCompras) return [];
    
    const data = sheetCompras.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    
    // Buscar índices de columnas
    let idProveedorIndex = -1;
    let numeroFacturaIndex = -1;
    let fechaFacturaIndex = -1;
    let montoIndex = -1;
    let estadoIndex = -1;
    
    for (let i = 0; i < headers.length; i++) {
      const header = (headers[i] || "").toString().toLowerCase().trim();
      if (header.includes("idproveedor") || header.includes("id proveedor")) idProveedorIndex = i;
      if (header.includes("nrofactura") || header.includes("numero") || header.includes("factura")) numeroFacturaIndex = i;
      if (header.includes("fecha") && header.includes("factura")) fechaFacturaIndex = i;
      if (header.includes("total") || header.includes("monto")) montoIndex = i;
      if (header.includes("cobrada") || header.includes("pagada") || header.includes("estado")) estadoIndex = i;
    }
    
    // Valores por defecto si no se encuentran
    if (idProveedorIndex === -1) idProveedorIndex = 0;
    if (numeroFacturaIndex === -1) numeroFacturaIndex = 4;
    if (fechaFacturaIndex === -1) fechaFacturaIndex = 2;
    if (montoIndex === -1) montoIndex = 10;
    if (estadoIndex === -1) estadoIndex = 16;
    
    const facturas = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Verificar si pertenece al proveedor
      if (row[idProveedorIndex] && row[idProveedorIndex].toString() === idProveedor.toString()) {
        
        // Verificar si está pendiente
        const estado = (row[estadoIndex] || "").toString().toLowerCase().trim();
        if (estado !== "pagada" && estado !== "cobrada" && estado !== "paid" && estado !== "collected") {
          
          let fecha = row[fechaFacturaIndex] || "";
          if (fecha instanceof Date) {
            fecha = fecha.toISOString().split('T')[0];
          } else if (fecha) {
            fecha = new Date(fecha).toISOString().split('T')[0];
          }
          
          facturas.push({
            numero: row[numeroFacturaIndex] || "N/A",
            fecha: fecha || "N/A",
            monto: parseFloat(row[montoIndex]) || 0,
            estado: "Pendiente"
          });
        }
      }
    }
    
    return facturas;
    
  } catch (e) {
    Logger.log("Error en obtenerFacturasPendientes: " + e.toString());
    return [];
  }
}
// =============================================
/**             PRODUCTOS                     */
// =============================================
function _resolverSSProductos_(idSpreadsheet) {
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
function obtenerProductos(idSpreadsheet) {
  try {
    var ss = _resolverSSProductos_(idSpreadsheet);
    if (!ss) return [];
    var sheet = ss.getSheetByName("Productos");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var lista = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var id = String(r[0] || '').trim();
      var nom = String(r[1] || '').trim();
      if (!id && !nom) continue;

      lista.push({
        IdProducto: id,
        NombreProducto: nom,
        PrecioProducto: parseFloat(r[2]) || 0,
        precioVenta: parseFloat(r[3]) || 0,
        IvaProducto: String(r[4] || "0"),
        StockProducto: parseFloat(r[5]) || 0,
        MinimoProducto: parseFloat(r[6]) || 0,
        CuentaContable: String(r[7] || ""),
        NombreCuenta: String(r[8] || ""),
        Cuenta: String(r[9] || ""),
        SubCuenta: String(r[10] || ""),
        CategoriaProducto: String(r[11] || ""),
        ImagenProducto: String(r[12] || ""),
        Status: String(r[13] || "Activo"),
        Condicion: String(r[14] || ""),
        subida: String(r[15] || ""),
        tipo: String(r[16] || "Simple"),
        Medida: String(r[17] || "UNID"),
        UnidadMedida: String(r[18] || "UNID"),
        FactorConversion: parseFloat(r[19]) || 1,
        Compuesto: String(r[20] || ""),
        CodBarra: String(r[21] || "")
      });
    }
    return lista;
  } catch(e) {
    Logger.log("Error en obtenerProductos: " + e);
    return [];
  }
}
function guardarProducto(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var idSpreadsheet = p.idSpreadsheet || p.empresaId;
    var ss = _resolverSSProductos_(idSpreadsheet);
    if (!ss) throw new Error("No se pudo conectar a la empresa activa.");

    var sheet = ss.getSheetByName("Productos");
    if (!sheet) throw new Error("No se encontró la hoja 'Productos'.");

    var data = sheet.getDataRange().getValues();
    var idProd = String(p.IdProducto || "").trim();
    if (!idProd) throw new Error("Código de producto obligatorio.");

    var filaEncontrada = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toUpperCase() === idProd.toUpperCase()) {
        filaEncontrada = i + 1;
        break;
      }
    }

    var filaValores = [
      idProd,
      p.NombreProducto || "",
      parseFloat(p.PrecioProducto) || 0,
      parseFloat(p.precioVenta) || 0,
      String(p.IvaProducto || "16"),
      parseFloat(p.StockProducto) || 0,
      parseFloat(p.MinimoProducto) || 0,
      p.CuentaContable || "",
      p.NombreCuenta || "",
      p.Cuenta || "",
      p.SubCuenta || "",
      p.CategoriaProducto || "",
      p.ImagenProducto || "",
      p.Status || "Activo",
      p.Condicion || "nuevo",
      p.subida || "",
      p.tipo || "Simple",
      p.Medida || "UNID",
      p.UnidadMedida || "UNID",
      parseFloat(p.FactorConversion) || 1,
      p.Compuesto || "",
      p.CodBarra || ""
    ];

    if (filaEncontrada !== -1) {
      sheet.getRange(filaEncontrada, 1, 1, filaValores.length).setValues([filaValores]);
      return { success: true, message: "Producto " + idProd + " actualizado correctamente." };
    } else {
      sheet.appendRow(filaValores);
      return { success: true, message: "Producto " + idProd + " creado exitosamente en la empresa activa." };
    }

  } catch(err) {
    Logger.log("Error en guardarProducto: " + err);
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
function eliminarProducto(id, idSpreadsheet) {
  try {
    var ss = _resolverSSProductos_(idSpreadsheet);
    if (!ss) throw new Error("No se pudo conectar a la empresa.");
    var sheet = ss.getSheetByName("Productos");
    if (!sheet) throw new Error("Hoja no encontrada.");

    var data = sheet.getDataRange().getValues();
    var idBuscado = String(id || '').trim().toUpperCase();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toUpperCase() === idBuscado) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Producto eliminado correctamente." };
      }
    }
    return { success: false, message: "Producto no encontrado." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}
function diagnosticarProductos() {
  try {
    const sheet = getSheetByName("Productos");
    if (!sheet) {
      return {
        success: false,
        message: "❌ La hoja 'Productos' no fue encontrada",
        sheetsDisponibles: SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName())
      };
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Verificar si hay datos
    if (lastRow <= 1) {
      return {
        success: false,
        message: "⚠️ La hoja 'Productos' está vacía (solo tiene encabezados o nada)",
        lastRow: lastRow,
        headers: headers,
        totalColumnas: lastCol
      };
    }
    
    // Obtener primeras 5 filas para muestra
    const muestra = sheet.getRange(2, 1, Math.min(5, lastRow - 1), lastCol).getValues();
    
    // Verificar columnas críticas
    const columnasRequeridas = ['IdProducto', 'NombreProducto', 'PrecioProducto', 'StockProducto'];
    const columnasEncontradas = {};
    const columnasFaltantes = [];
    
    columnasRequeridas.forEach(col => {
      const found = headers.findIndex(h => 
        h.toString().trim().toLowerCase() === col.toLowerCase()
      );
      if (found === -1) {
        columnasFaltantes.push(col);
      } else {
        columnasEncontradas[col] = found;
      }
    });
    
    return {
      success: true,
      message: "✅ Diagnóstico completado",
      lastRow: lastRow,
      totalColumnas: lastCol,
      headers: headers,
      columnasEncontradas: columnasEncontradas,
      columnasFaltantes: columnasFaltantes,
      muestra: muestra,
      cantidadProductos: lastRow - 1
    };
    
  } catch (e) {
    return {
      success: false,
      message: "❌ Error en diagnóstico: " + e.toString()
    };
  }
}
function getDataProductos(idSpreadsheet) {
  try {
    var ss = _resolverSSVentas_(idSpreadsheet);
    if (!ss) return [];
    var sheet = ss.getSheetByName("Productos");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var productos = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var cod = String(row[0] || '').trim();
      var nom = String(row[1] || '').trim();
      if (!nom) continue;
      
      // Formato estricto: [CODIGO, NOMBRE, PRECIO_COMPRA, PRECIO_VENTA, IVA, STOCK, MINIMO, CUENTA_CONT, NOMBRE_CTA, CUENTA, SUBCUENTA, CATEGORIA]
      productos.push([
        cod,                                                              // 0
        nom,                                                              // 1
        parseFloat(row[2]) || 0,                                          // 2: Costo
        parseFloat(row[3]) || 0,                                          // 3: Venta
        parseFloat(row[4] !== undefined ? row[4] : 16) || 0,              // 4: IVA
        parseFloat(row[5]) || 0,                                          // 5: Stock
        String(row[7] || '').trim(),                                      // 6: Cuenta contable
        String(row[8] || 'INVENTARIO').trim(),                            // 7: Nombre cuenta
        String(row[9] || '').trim(),                                      // 8: Cuenta
        String(row[10] || '').trim(),                                     // 9: Subcuenta
        String(row[11] || 'GENERAL').trim()                               // 10: Categoría
      ]);
    }
    return productos;
  } catch(e) {
    Logger.log("Error en getDataProductos: " + e);
    return [];
  }
}
function verificarProducto(codigo) {
  try {
    var sheet = getSheetByName("Productos");
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    
    var producto = data.find(function(row) {
      return row[0] === codigo;
    });
    
    if (producto) {
      console.log(`Producto ${codigo} encontrado:`);
      console.log(`- Nombre: ${producto[1]}`);
      console.log(`- Stock: ${producto[5]}`);
      console.log(`- Status: ${producto[13]}`);
      console.log(`- Precio Venta: ${producto[3]}`);
      return {
        encontrado: true,
        nombre: producto[1],
        stock: producto[5],
        status: producto[13],
        precioVenta: producto[3]
      };
    } else {
      console.log(`Producto ${codigo} NO encontrado`);
      return { encontrado: false };
    }
  } catch (e) {
    console.error("Error:", e);
    return { encontrado: false, error: e.toString() };
  }
}
function actualizarProducto(productoData) {
  try {
    const sheet = getSheetByName("Productos");
    if (!sheet) {
      return { success: false, message: "La hoja 'Productos' no fue encontrada." };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const colIndex = {};
    headers.forEach((h, i) => { colIndex[h] = i; });

    if (colIndex["IdProducto"] === undefined) {
      return { success: false, message: 'Error: La columna "IdProducto" no se encontró.' };
    }

    const idBuscado = productoData.id || productoData.IdProducto;
    const rowIndexToUpdate = data.findIndex((row, i) =>
      i > 0 && String(row[colIndex["IdProducto"]]) === String(idBuscado)
    );

    if (rowIndexToUpdate === -1) {
      return { success: false, message: "El producto con ID " + idBuscado + " no fue encontrado." };
    }

    const filaActual = data[rowIndexToUpdate];
    const idxCondicion = colIndex["Condicion"];
    const idxSubida = colIndex["subida"];

    const condicionActual = idxCondicion !== undefined ? filaActual[idxCondicion] : "";
    const subidaActual = idxSubida !== undefined ? filaActual[idxSubida] : "";

    const cuentaInfo = getCuentaContableCompleta(
      productoData.cuenta || productoData.CuentaContable || ""
    );

    const mapeoCampos = {
      NombreProducto: (productoData.nombre || productoData.NombreProducto || "").toString().toUpperCase(),
      CategoriaProducto: productoData.categoria || productoData.CategoriaProducto,
      PrecioProducto: parseFloat(productoData.precio || productoData.PrecioProducto) || 0,
      precioVenta: parseFloat(productoData.precioVenta) || 0,
      IvaProducto: String(productoData.iva || productoData.IvaProducto || "0"),
      StockProducto: parseFloat(productoData.stock || productoData.StockProducto) || 0,
      MinimoProducto: parseFloat(productoData.minimo || productoData.MinimoProducto) || 0,
      CuentaContable: cuentaInfo.CuentaContable,
      NombreCuenta: cuentaInfo.NombreCuenta,
      Cuenta: cuentaInfo.Cuenta,
      SubCuenta: cuentaInfo.SubCuenta,
      ImagenProducto: productoData.imagen || productoData.ImagenProducto || "",
      Status: productoData.Status || productoData.status || "Activo",
      Condicion: productoData.Condicion || productoData.condicion || condicionActual || "nuevo",
      subida: productoData.subida || subidaActual || "",
      tipo: productoData.tipo || "Simple",
      Medida: productoData.medida || productoData.Medida || "UNID",
      UnidadMedida: productoData.unidadMedida || productoData.UnidadMedida || "UNID",
      FactorConversion: parseFloat(productoData.factorConversion || productoData.FactorConversion) || 1,
      Compuesto: productoData.compuesto || productoData.Compuesto || "",
      CodBarra: productoData.CodBarra || productoData.codigoBarras || productoData.CodigoBarras || ""
    };

    const rowNumber = rowIndexToUpdate + 1;
    for (let headerName in mapeoCampos) {
      if (colIndex[headerName] !== undefined) {
        sheet.getRange(rowNumber, colIndex[headerName] + 1).setValue(mapeoCampos[headerName]);
      }
    }

// Detectar si cambió el precio para recalcular compuestos afectados
const precioAnterior = parseFloat(filaActual[colIndex["PrecioProducto"]]) || 0;
const precioNuevo = mapeoCampos.PrecioProducto;

if (precioAnterior !== precioNuevo) {
  try {
    // Si es compuesto, recalcular todos
    if (mapeoCampos.tipo === "Compuesto") {
      recalcularCostosCompuestos();
    } else {
      // Si es simple, recalcular compuestos que lo usan
      recalcularCompuestosPorIngrediente(idBuscado);
    }
  } catch (e) {
    console.error("Error al recalcular costos tras actualizar: " + e.toString());
    // No fallar la actualización por esto
  }
}

    return { success: true, message: "Producto actualizado exitosamente." };

  } catch (e) {
    console.error("Error en actualizarProducto: " + e.toString());
    return { success: false, message: "Error en el servidor al actualizar: " + e.message };
  }
}
function convertirUnidad(cantidad, unidadOrigen, unidadDestino, factorConversion) {
  try {
    const cant = parseFloat(cantidad);
    if (isNaN(cant)) {
      return { success: false, message: "Cantidad inválida." };
    }

    const normalizar = function(u) {
      const s = String(u || "").toUpperCase().trim();
      if (s === "G" || s === "GR" || s === "GRS" || s === "GRAMO" || s === "GRAMOS") return "GR";
      if (s === "KG" || s === "KILO" || s === "KILOS" || s === "KILOGRAMO" || s === "KILOGRAMOS") return "KG";
      if (s === "MG" || s === "MILIGRAMO" || s === "MILIGRAMOS") return "MG";
      if (s === "ML" || s === "MILILITRO" || s === "MILILITROS") return "ML";
      if (s === "L" || s === "LT" || s === "LITRO" || s === "LITROS") return "L";
      if (s === "CL" || s === "CENTILITRO" || s === "CENTILITROS") return "CL";
      if (s === "UNID" || s === "UNIDAD" || s === "UNIDADES" || s === "U" || s === "UN") return "UNID";
      if (s === "BULTO" || s === "BULTOS" || s === "BLT") return "BULTO";
      if (s === "PQ" || s === "PAQUETE" || s === "PAQUETES" || s === "PK") return "PQ";
      if (s === "CAJA" || s === "CAJAS" || s === "CJ") return "CAJA";
      if (s === "DOCENA" || s === "DOCENAS" || s === "DOC") return "DOCENA";
      if (s === "PAR" || s === "PARES") return "PAR";
      return s;
    };

    const origen = normalizar(unidadOrigen);
    const destino = normalizar(unidadDestino);

    if (origen === destino) {
      return { success: true, cantidad: cant };
    }

    const factor = parseFloat(factorConversion) || 1;

    // Conversiones directas con factorConversion (producto específico)
    // KG -> GR o L -> ML: multiplicar por factor
    if ((origen === "KG" && (destino === "GR" || destino === "MG")) ||
        (origen === "L" && (destino === "ML" || destino === "CL"))) {
      // Si el factor es el estándar, usar factor
      if (factor > 1) {
        // Ajustar según destino
        if (origen === "KG" && destino === "MG") {
          return { success: true, cantidad: cant * 1000000 };
        }
        if (origen === "L" && destino === "CL") {
          return { success: true, cantidad: cant * 100 };
        }
        return { success: true, cantidad: cant * factor };
      }
    }

    // GR -> KG o ML -> L: dividir por factor
    if ((origen === "GR" && destino === "KG") || (origen === "ML" && destino === "L")) {
      if (factor > 1) {
        return { success: true, cantidad: cant / factor };
      }
    }

    // BULTO -> UNID: multiplicar por factor (24 unidades por bulto)
    if ((origen === "BULTO" || origen === "PQ" || origen === "CAJA") && destino === "UNID") {
      const f = factor > 1 ? factor : 24; // Default 24 si no hay factor
      return { success: true, cantidad: cant * f };
    }

    // UNID -> BULTO (o PQ o CAJA): dividir por factor
    if (origen === "UNID" && (destino === "BULTO" || destino === "PQ" || destino === "CAJA")) {
      const f = factor > 1 ? factor : 24;
      return { success: true, cantidad: cant / f };
    }

    // DOCENA -> UNID
    if (origen === "DOCENA" && destino === "UNID") {
      return { success: true, cantidad: cant * 12 };
    }
    if (origen === "UNID" && destino === "DOCENA") {
      return { success: true, cantidad: cant / 12 };
    }

    // PAR -> UNID
    if (origen === "PAR" && destino === "UNID") {
      return { success: true, cantidad: cant * 2 };
    }
    if (origen === "UNID" && destino === "PAR") {
      return { success: true, cantidad: cant / 2 };
    }

    // Conversiones por familia usando tabla base
    const factoresBase = {
      "MG": 0.001,
      "GR": 1,
      "KG": 1000,
      "CL": 10,
      "ML": 1,
      "L": 1000,
      "UNID": 1
    };

    const familiaMasa = ["MG", "GR", "KG"];
    const familiaVolumen = ["CL", "ML", "L"];

    if (familiaMasa.indexOf(origen) !== -1 && familiaMasa.indexOf(destino) !== -1) {
      const cantBase = cant * factoresBase[origen];
      return { success: true, cantidad: cantBase / factoresBase[destino] };
    }
    if (familiaVolumen.indexOf(origen) !== -1 && familiaVolumen.indexOf(destino) !== -1) {
      const cantBase = cant * factoresBase[origen];
      return { success: true, cantidad: cantBase / factoresBase[destino] };
    }

    // Cruce masa ↔ volumen: no permitido (requiere densidad)
    // Cruce empaque ↔ unidad: ya manejado arriba

    return { success: false, message: "No se puede convertir de " + unidadOrigen + " a " + unidadDestino + "." };

  } catch (e) {
    return { success: false, message: "Error en conversión: " + e.message };
  }
}
function descontarStockCompuesto(idProducto, cantidadProducir) {
  try {
    const cantidad = parseFloat(cantidadProducir);
    if (isNaN(cantidad) || cantidad <= 0) {
      return { success: false, message: "La cantidad a producir debe ser mayor a 0." };
    }

    const sheet = getSheetByName("Productos");
    if (!sheet) {
      return { success: false, message: "La hoja 'Productos' no fue encontrada." };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const idxId = headers.indexOf("IdProducto");
    const idxNombre = headers.indexOf("NombreProducto");
    const idxCompuesto = headers.indexOf("Compuesto");
    const idxStock = headers.indexOf("StockProducto");
    const idxTipo = headers.indexOf("tipo");
    const idxMedida = headers.indexOf("Medida");
    const idxUnidad = headers.indexOf("UnidadMedida");
    const idxFactor = headers.indexOf("FactorConversion");

    if (idxId === -1 || idxCompuesto === -1 || idxStock === -1) {
      return { success: false, message: "Faltan columnas requeridas en la hoja Productos." };
    }

    const rowCompuestoIdx = data.findIndex((row, i) => i > 0 && String(row[idxId]) === String(idProducto));
    if (rowCompuestoIdx === -1) {
      return { success: false, message: "No se encontró el producto compuesto." };
    }

    const rowCompuesto = data[rowCompuestoIdx];
    if (idxTipo !== -1 && rowCompuesto[idxTipo] !== "Compuesto") {
      return { success: false, message: "El producto seleccionado no es de tipo compuesto." };
    }

    let receta;
    try {
      receta = JSON.parse(rowCompuesto[idxCompuesto]);
    } catch (e) {
      return { success: false, message: "La receta del producto no posee un formato válido." };
    }

    if (!receta || receta.length === 0) {
      return { success: false, message: "La receta de este producto compuesto se encuentra vacía." };
    }

    // Validar stock de todos los ingredientes antes de descontar
    const actualizaciones = [];
    const faltantes = [];

    for (let i = 0; i < receta.length; i++) {
      const ingrediente = receta[i];
      const rowIngredienteIdx = data.findIndex((row, idx) => idx > 0 && String(row[idxId]) === String(ingrediente.id));

      if (rowIngredienteIdx === -1) {
        faltantes.push("Ingrediente con ID " + ingrediente.id + " no encontrado en inventario.");
        continue;
      }

      const rowIngrediente = data[rowIngredienteIdx];
      const stockActual = parseFloat(rowIngrediente[idxStock]) || 0;
      const medidaIngrediente = idxMedida !== -1 ? String(rowIngrediente[idxMedida] || "UNID").toUpperCase() : "UNID";
      const factorIngrediente = idxFactor !== -1 ? parseFloat(rowIngrediente[idxFactor]) || 1 : 1;
      const nombreIngrediente = idxNombre !== -1 ? String(rowIngrediente[idxNombre]) : ingrediente.id;

      const qtyReceta = parseFloat(ingrediente.qty) || 0;
      const unidadReceta = String(ingrediente.unit || medidaIngrediente).toUpperCase();
      const cantidadTotalReceta = qtyReceta * cantidad;

      // Convertir la cantidad de la receta a la unidad del ingrediente en inventario
      const conversion = convertirUnidad(cantidadTotalReceta, unidadReceta, medidaIngrediente, factorIngrediente);

      if (!conversion.success) {
        faltantes.push("No se pudo convertir unidad para " + nombreIngrediente + ": " + conversion.message);
        continue;
      }

      const cantidadADescontar = conversion.cantidad;

      if (stockActual < cantidadADescontar) {
        faltantes.push(
          "Stock insuficiente para: " + nombreIngrediente +
          ". Requerido: " + cantidadADescontar.toFixed(2) + " " + medidaIngrediente +
          ", Disponible: " + stockActual.toFixed(2) + " " + medidaIngrediente
        );
        continue;
      }

      actualizaciones.push({
        rowNum: rowIngredienteIdx + 1,
        colNum: idxStock + 1,
        nuevoStock: stockActual - cantidadADescontar,
        nombre: nombreIngrediente
      });
    }

    if (faltantes.length > 0) {
      return { success: false, message: "No se puede producir:\n" + faltantes.join("\n") };
    }

    // Aplicar todas las actualizaciones
    actualizaciones.forEach(function(act) {
      sheet.getRange(act.rowNum, act.colNum).setValue(act.nuevoStock);
    });

    // Sumar al stock del compuesto
    const stockCompuestoActual = parseFloat(rowCompuesto[idxStock]) || 0;
    sheet.getRange(rowCompuestoIdx + 1, idxStock + 1).setValue(stockCompuestoActual + cantidad);

    return {
      success: true,
      message: "Producción finalizada. Se descontaron los ingredientes correctamente.",
      ingredientesDescontados: actualizaciones.length
    };

  } catch (e) {
    console.error("Error en descontarStockCompuesto: " + e.toString());
    return { success: false, message: "Error interno: " + e.message };
  }
}
function calcularDisponibilidadCompuesto(idProducto) {
  try {
    const sheet = getSheetByName("Productos");
    if (!sheet) {
      return { success: false, message: "La hoja 'Productos' no fue encontrada.", maxProducible: 0, desglose: [] };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const idxId = headers.indexOf("IdProducto");
    const idxNombre = headers.indexOf("NombreProducto");
    const idxCompuesto = headers.indexOf("Compuesto");
    const idxStock = headers.indexOf("StockProducto");
    const idxMedida = headers.indexOf("Medida");
    const idxFactor = headers.indexOf("FactorConversion");

    if (idxId === -1 || idxCompuesto === -1 || idxStock === -1) {
      return { success: false, message: "Faltan columnas requeridas.", maxProducible: 0, desglose: [] };
    }

    const rowCompuesto = data.find((row, i) => i > 0 && String(row[idxId]) === String(idProducto));
    if (!rowCompuesto) {
      return { success: false, message: "Producto no encontrado.", maxProducible: 0, desglose: [] };
    }

    let receta;
    try {
      receta = JSON.parse(rowCompuesto[idxCompuesto]);
    } catch (e) {
      return { success: false, message: "Receta inválida.", maxProducible: 0, desglose: [] };
    }

    if (!receta || receta.length === 0) {
      return { success: false, message: "Receta vacía.", maxProducible: 0, desglose: [] };
    }

    let maxProducible = Infinity;
    const desglose = [];

    for (let i = 0; i < receta.length; i++) {
      const ingrediente = receta[i];
      const rowIngredienteIdx = data.findIndex((row, idx) => idx > 0 && String(row[idxId]) === String(ingrediente.id));

      if (rowIngredienteIdx === -1) {
        desglose.push({
          id: ingrediente.id,
          nombre: ingrediente.nombre || ingrediente.id,
          encontrado: false,
          stockDisponible: 0,
          requeridoPorUnidad: parseFloat(ingrediente.qty) || 0,
          unidad: ingrediente.unit,
          maxParaEste: 0
        });
        maxProducible = 0;
        continue;
      }

      const rowIngrediente = data[rowIngredienteIdx];
      const stockDisponible = parseFloat(rowIngrediente[idxStock]) || 0;
      const medidaIngrediente = idxMedida !== -1 ? String(rowIngrediente[idxMedida] || "UNID").toUpperCase() : "UNID";
      const factorIngrediente = idxFactor !== -1 ? parseFloat(rowIngrediente[idxFactor]) || 1 : 1;
      const nombreIngrediente = idxNombre !== -1 ? String(rowIngrediente[idxNombre]) : ingrediente.id;

      const qtyReceta = parseFloat(ingrediente.qty) || 0;
      const unidadReceta = String(ingrediente.unit || medidaIngrediente).toUpperCase();

      // Convertir 1 unidad de receta a la unidad del ingrediente
      const conversion = convertirUnidad(qtyReceta, unidadReceta, medidaIngrediente, factorIngrediente);

      if (!conversion.success) {
        desglose.push({
          id: ingrediente.id,
          nombre: nombreIngrediente,
          encontrado: true,
          error: conversion.message,
          stockDisponible: stockDisponible,
          requeridoPorUnidad: qtyReceta,
          unidad: unidadReceta,
          maxParaEste: 0
        });
        maxProducible = 0;
        continue;
      }

      const requeridoPorUnidad = conversion.cantidad;
      const maxParaEste = requeridoPorUnidad > 0 ? Math.floor(stockDisponible / requeridoPorUnidad) : 0;

      if (maxParaEste < maxProducible) {
        maxProducible = maxParaEste;
      }

      desglose.push({
        id: ingrediente.id,
        nombre: nombreIngrediente,
        encontrado: true,
        stockDisponible: stockDisponible,
        requeridoPorUnidad: qtyReceta,
        unidad: unidadReceta,
        unidadInventario: medidaIngrediente,
        requeridoEnInventario: requeridoPorUnidad,
        maxParaEste: maxParaEste
      });
    }

    if (maxProducible === Infinity) maxProducible = 0;

    return {
      success: true,
      maxProducible: maxProducible,
      desglose: desglose
    };

  } catch (e) {
    console.error("Error en calcularDisponibilidadCompuesto: " + e.toString());
    return { success: false, message: "Error: " + e.message, maxProducible: 0, desglose: [] };
  }
}
function getDataListaProductos(){const sheetProductos=getSheetByName("Productos");const productosData=sheetProductos.getDataRange().getDisplayValues();const productosHeaders=productosData.shift();let productosDataArray=[];productosData.forEach(dataRow=>{const temporaryObject={};productosHeaders.forEach((headerTitle,index)=>{temporaryObject[headerTitle]=dataRow[index];});productosDataArray.push(temporaryObject);});return productosDataArray;}

function actualizarStock(producto, cantidad) {
  try {
    const sheetProductos = getSheetByName("Productos");
    if (!sheetProductos) {
      return { success: false, message: "La hoja 'Productos' no fue encontrada." };
    }

    const data = sheetProductos.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxNombre = headers.indexOf("NombreProducto");
    const idxStock = headers.indexOf("StockProducto");

    if (idxNombre === -1 || idxStock === -1) {
      return { success: false, message: "Faltan columnas 'NombreProducto' o 'StockProducto'." };
    }

    if (!producto || isNaN(parseFloat(cantidad))) {
      return { success: false, message: "Producto o cantidad inválidos." };
    }

    const cantidadNum = parseFloat(cantidad);
    const index = data.findIndex((row, i) => i > 0 && String(row[idxNombre]).trim() === String(producto).trim());

    if (index === -1) {
      return { success: false, message: "Producto no encontrado: " + producto };
    }

    const stockActual = parseFloat(data[index][idxStock]) || 0;
    let nuevoStock = stockActual - cantidadNum;
    if (nuevoStock < 0) nuevoStock = 0;

    sheetProductos.getRange(index + 1, idxStock + 1).setValue(nuevoStock);

    return { success: true, message: "Stock actualizado.", nuevoStock: nuevoStock };

  } catch (e) {
    console.error("Error en actualizarStock: " + e.toString());
    return { success: false, message: "Error: " + e.message };
  }
}
function obtenerStockProducto(nombreProducto) {
  try {
    const sheetProductos = getSheetByName('Productos');
    const data = sheetProductos.getDataRange().getValues();
    
    // Buscar el producto por nombre (índice 1)
    for (let i = 1; i < data.length; i++) { // i=1 para saltar encabezados
      if (data[i][1] === nombreProducto) {
        // Stock ahora está en índice 5
        const stock = parseFloat(data[i][5]) || 0;
        console.log("Stock encontrado para " + nombreProducto + ": " + stock);
        return stock;
      }
    }
    
    console.log("Producto no encontrado: " + nombreProducto);
    return 0;
    
  } catch (e) {
    console.error("Error en obtenerStockProducto: " + e.toString());
    return 0;
  }
}
function obtenerTodosProductos(){const sheet=getSheetByName("Productos");const lastRow=sheet.getLastRow();const lastColumn=sheet.getLastColumn();if(lastRow<=1)return[];const data=sheet.getRange(2,1,lastRow-1,lastColumn).getValues();const headers=sheet.getRange(1,1,1,lastColumn).getValues()[0].map(h=>String(h).trim());return data.map(row=>{const p={};headers.forEach((h,i)=>{p[h]=row[i];});return p;}).filter(p=>p.IdProducto&&p.NombreProducto);}
function buscarProductosPorTermino(searchTerm) {
  try {
    if (!searchTerm || searchTerm.trim().length < 2) {
      return [];
    }

    const sheet = getSheetByName("Productos");
    if (!sheet) {
      throw new Error("No se encontró la hoja 'Productos'");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1);

    const idIndex = headers.indexOf("IdProducto");
    const nombreIndex = headers.indexOf("NombreProducto");
    const codBarraIndex = headers.indexOf("CodBarra");
    const categoriaIndex = headers.indexOf("CategoriaProducto");

    if (idIndex === -1 || nombreIndex === -1) {
      throw new Error("No se encontraron las columnas 'IdProducto' o 'NombreProducto' en la hoja.");
    }

    // Cargar mapa de categorías para buscar por nombre
    let categoriasMap = {};
    try {
      const categorias = obtenerCategoriasConId();
      if (Array.isArray(categorias)) {
        categorias.forEach(c => {
          if (c && c.id) categoriasMap[String(c.id)] = String(c.nombre || "").toLowerCase();
        });
      }
    } catch (e) {
      // Si falla, continuar sin búsqueda por categoría
    }

    const lowerSearch = searchTerm.toLowerCase();
    const resultados = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const idProducto = row[idIndex] ? String(row[idIndex]).toLowerCase() : "";
      const nombreProducto = row[nombreIndex] ? String(row[nombreIndex]).toLowerCase() : "";
      const codBarra = codBarraIndex !== -1 && row[codBarraIndex] ? String(row[codBarraIndex]).toLowerCase() : "";
      const categoriaId = categoriaIndex !== -1 && row[categoriaIndex] ? String(row[categoriaIndex]) : "";
      const categoriaNombre = categoriasMap[categoriaId] || "";

      if (
        idProducto.indexOf(lowerSearch) !== -1 ||
        nombreProducto.indexOf(lowerSearch) !== -1 ||
        codBarra.indexOf(lowerSearch) !== -1 ||
        categoriaNombre.indexOf(lowerSearch) !== -1
      ) {
        const productoEncontrado = {};
        headers.forEach((header, idx) => {
          productoEncontrado[header] = row[idx];
        });
        resultados.push(productoEncontrado);
      }

      if (resultados.length >= 10) {
        break;
      }
    }

    if (DEBUG) console.log('Búsqueda de "' + searchTerm + '" encontró ' + resultados.length + ' resultados.');
    return resultados;

  } catch (e) {
    console.error("Error en buscarProductosPorTermino: " + e.message);
    return { error: e.message };
  }
}
function obtenerTop5Productos() {
  const hojaDetalle = getSheetByName("DetalleFacturas");
  if (!hojaDetalle) throw new Error("No existe la hoja 'DetalleFacturas'");

  const data = hojaDetalle.getDataRange().getValues();
  if (data.length < 2) throw new Error("No hay datos para procesar");

  const productoIndex = 1;
  const cantidadIndex = 3;

  const ventas = {};

  for (let i = 1; i < data.length; i++) {
    const producto = data[i][productoIndex];
    const cantidad = Number(data[i][cantidadIndex]) || 0;
    if (producto) {
      ventas[producto] = (ventas[producto] || 0) + cantidad;
    }
  }

  const topProductos = Object.entries(ventas)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  const hojaResumen = getSheetByName("Resumen");
  if (!hojaResumen) throw new Error("No existe la hoja 'Resumen'");

  const startRow = 30;
  const startCol = 9;

  const output = [["Producto", "Cantidad Vendida"]];
  topProductos.forEach(([producto, cantidad]) => {
    output.push([producto, cantidad]);
  });

  hojaResumen.getRange(startRow, startCol, output.length, 2).clearContent();
  hojaResumen.getRange(startRow, startCol, output.length, 2).setValues(output);
}
function getTop5ProductosParaGrafico() {
  const hojaDetalle = getSheetByName("DetalleFacturas");
  if (!hojaDetalle) return { error: "No existe la hoja 'DetalleFacturas'" };

  const data = hojaDetalle.getDataRange().getValues();
  if (data.length < 2) return { error: "No hay datos para procesar" };

  const productoIndex = 1;
  const cantidadIndex = 3;

  const ventas = {};

  for (let i = 1; i < data.length; i++) {
    const producto = data[i][productoIndex]?.toString().trim();
    const cantidad = Number(data[i][cantidadIndex]) || 0;
    if (producto) {
      ventas[producto] = (ventas[producto] || 0) + cantidad;
    }
  }

  const topProductos = Object.entries(ventas)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  return {
    labels: topProductos.map(p => p[0]),
    data: topProductos.map(p => p[1])
  };
}
function actualizarStockProductos(items, tipoTransaccion, devuelveStock) {
  try {
    const productosSheet = getSheetByName("Productos");
    if (!productosSheet) {
      throw new Error("La hoja 'Productos' no fue encontrada.");
    }

    const lastRow = productosSheet.getLastRow();
    const lastColumn = productosSheet.getLastColumn();

    if (lastRow <= 1 || lastColumn === 0) {
      return { actualizados: 0, noEncontrados: 0, productosNoEncontrados: [], advertencias: [] };
    }

    const dataRange = productosSheet.getRange(2, 1, lastRow - 1, lastColumn);
    const productosData = dataRange.getValues();
    const headers = productosSheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => String(h).trim());

    const idxId = headers.indexOf("IdProducto");
    const idxNombre = headers.indexOf("NombreProducto");
    const idxStock = headers.indexOf("StockProducto");

    if (idxId === -1 || idxNombre === -1 || idxStock === -1) {
      throw new Error("Faltan columnas requeridas: IdProducto, NombreProducto o StockProducto.");
    }

    // Construir mapas por ID y por nombre
    const productosPorId = new Map();
    const productosPorNombre = new Map();
    const nombresDuplicados = new Set();

    productosData.forEach((row, index) => {
      const id = String(row[idxId] || "").trim();
      const nombre = String(row[idxNombre] || "").trim().toUpperCase();
      const info = {
        stock: parseFloat(row[idxStock]) || 0,
        rowIndex: index + 2,
        id: id,
        nombre: nombre
      };

      if (id) productosPorId.set(id, info);

      if (nombre) {
        if (productosPorNombre.has(nombre)) {
          nombresDuplicados.add(nombre);
        } else {
          productosPorNombre.set(nombre, info);
        }
      }
    });

    // Eliminar duplicados del mapa de nombres
    nombresDuplicados.forEach(nombre => {
      productosPorNombre.delete(nombre);
    });

    const updates = [];
    const productosNoEncontrados = [];
    const advertencias = [];

    const esNC = (tipoTransaccion === 'NC');
    const ncDevuelve = (typeof devuelveStock === 'boolean') ? devuelveStock : true;

    items.forEach(item => {
      const idItem = String(item.id || item.IdProducto || item.codigo || "").trim();
      const nombreItem = String(item.producto || item.Producto || item.nombre || item.Nombre || "").trim().toUpperCase();

      if (!idItem && !nombreItem) return;

      let cantidadInput = item.cantidad || item.Cantidad;
      let cantidad = 0;
      if (typeof cantidadInput === "number") {
        cantidad = cantidadInput;
      } else if (cantidadInput) {
        cantidad = parseFloat(String(cantidadInput).replace(/\./g, "").replace(",", "."));
      }
      cantidad = cantidad || 0;

      if (cantidad <= 0) return;

      let info = null;
      let usadoNombre = false;

      // Buscar primero por ID
      if (idItem && productosPorId.has(idItem)) {
        info = productosPorId.get(idItem);
      } else if (nombreItem) {
        if (nombresDuplicados.has(nombreItem)) {
          advertencias.push("Nombre duplicado: '" + nombreItem + "'. No se puede actualizar por nombre. Use el ID.");
          productosNoEncontrados.push(nombreItem);
          return;
        }
        if (productosPorNombre.has(nombreItem)) {
          info = productosPorNombre.get(nombreItem);
          usadoNombre = true;
        }
      }

      if (!info) {
        productosNoEncontrados.push(idItem || nombreItem);
        return;
      }

      if (usadoNombre) {
        advertencias.push("Actualizado por nombre '" + info.nombre + "'. Se recomienda usar el ID " + info.id + ".");
      }

      const stockActual = info.stock;
      let nuevoStock = stockActual;

      if (tipoTransaccion === "Compra") {
        nuevoStock = stockActual + cantidad;
      } else if (tipoTransaccion === "Venta") {
        nuevoStock = stockActual - cantidad;
        if (nuevoStock < 0) {
          advertencias.push("Stock insuficiente para '" + info.nombre + "'. Stock: " + stockActual + ", solicitado: " + cantidad + ". Se ajusta a 0.");
          nuevoStock = 0;
        }
      } else if (esNC) {
        if (!ncDevuelve) {
          // Ajuste de precio sin devolución física: no toca stock
          return;
        }
        nuevoStock = stockActual + cantidad;
      } else {
        return;
      }

      updates.push({
        rowIndex: info.rowIndex,
        colIndex: idxStock + 1,
        value: nuevoStock,
        producto: info.nombre,
        id: info.id,
        stockAnterior: stockActual,
        cantidad: cantidad,
        nuevoStock: nuevoStock
      });
    });

    // Aplicar actualizaciones
    updates.forEach(u => {
      productosSheet.getRange(u.rowIndex, u.colIndex).setValue(u.value);
    });

    return {
      actualizados: updates.length,
      noEncontrados: productosNoEncontrados.length,
      productosNoEncontrados: productosNoEncontrados,
      advertencias: advertencias
    };

  } catch (e) {
    console.error("Error en actualizarStockProductos: " + e.toString());
    throw e;
  }
}
function enviarAlertaStock() {
  try {
    const sheet = getSheetByName("Productos");
    if (!sheet) {
      return { success: false, message: "La hoja 'Productos' no fue encontrada." };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, message: "No hay productos para analizar." };
    }

    const headers = data[0].map(h => String(h).trim());

    const nombreIndex = headers.findIndex(h => h === "NombreProducto" || h === "Producto");
    const skuIndex = headers.findIndex(h => h === "IdProducto" || h === "SKU" || h === "Código");
    const stockIndex = headers.findIndex(h => h === "StockProducto" || h === "Stock" || h === "Cantidad");
    const minimoIndex = headers.findIndex(h => h === "MinimoProducto" || h === "Stock Mínimo" || h === "Mínimo");
    const categoriaIndex = headers.findIndex(h => h === "CategoriaProducto" || h === "Categoría");

    if (nombreIndex === -1 || skuIndex === -1 || stockIndex === -1) {
      return { success: false, message: "No se encontraron las columnas necesarias en la hoja." };
    }

    const productosCero = [];
    const productosUno = [];
    const productosMinimo = [];
    const productosOk = [];

    const totalFilas = data.length - 1;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nombre = row[nombreIndex] || "";
      const sku = row[skuIndex] || "";
      const stock = parseInt(row[stockIndex]) || 0;
      const minimo = minimoIndex !== -1 ? parseInt(row[minimoIndex]) || 0 : 0;
      const categoria = categoriaIndex !== -1 ? row[categoriaIndex] || "" : "";

      if (!nombre && !sku) continue;

      const item = { nombre, sku, stock, minimo, categoria };

      if (stock === 0) {
        productosCero.push(item);
      } else if (stock === 1) {
        productosUno.push(item);
      } else if (stock <= minimo) {
        productosMinimo.push(item);
      } else {
        productosOk.push(item);
      }
    }

    const fecha = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy");
    const hora = Utilities.formatDate(new Date(), "America/Mexico_City", "HH:mm");

    const limitePorSeccion = 20;
    const esParcial = totalFilas > 200;

    let htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; }
    .header { background-color: #001f3f; color: white; padding: 15px; text-align: center; }
    .content { padding: 15px; background-color: #f8f9fa; }
    .resumen { background: white; border: 1px solid #ddd; border-radius: 5px; padding: 15px; margin-bottom: 15px; }
    .resumen-table { width: 100%; border-collapse: collapse; }
    .resumen-table td { padding: 8px; border: 1px solid #ddd; }
    .alerta { border-radius: 5px; padding: 10px; margin-bottom: 10px; }
    .alerta-cero { background-color: #f8d7da; border: 1px solid #f5c6cb; }
    .alerta-uno { background-color: #fff3cd; border: 1px solid #ffeaa7; }
    .alerta-minimo { background-color: #ffe5d0; border: 1px solid #ffcc80; }
    .producto-table { width: 100%; border-collapse: collapse; font-size: 12px; }
    .producto-table td { padding: 6px; border: 1px solid #ddd; }
    .footer { margin-top: 20px; color: #6c757d; font-size: 12px; text-align: center; }
    .badge { display: inline-block; padding: 3px 8px; border-radius: 10px; font-size: 12px; font-weight: bold; }
    .badge-success { background-color: #28a745; color: white; }
    .badge-danger { background-color: #dc3545; color: white; }
    .badge-warning { background-color: #ffc107; color: black; }
    .badge-orange { background-color: #fd7e14; color: white; }
    .aviso { background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px; padding: 10px; margin-bottom: 15px; color: #856404; font-size: 13px; }
  </style>
</head>
<body>
  <div class="header">
    <h1 style="margin: 0; font-size: 20px;">📊 Reporte de Inventario</h1>
    <p style="margin: 5px 0 0 0;">${fecha} - ${hora}</p>
  </div>
  <div class="content">
    ${esParcial ? '<div class="aviso">⚠️ <b>Reporte extenso:</b> Se analizaron ' + totalFilas + ' productos.</div>' : ''}
    <div class="resumen">
      <h3 style="margin-top: 0;">📋 Resumen del Inventario</h3>
      <table class="resumen-table">
        <tr><td><span class="badge badge-success">🟢 Stock Normal</span></td><td style="text-align: center;"><strong>${productosOk.length}</strong></td></tr>
        <tr><td><span class="badge badge-danger">🔴 Stock Cero</span></td><td style="text-align: center;"><strong>${productosCero.length}</strong></td></tr>
        <tr><td><span class="badge badge-warning">🟡 Última Unidad</span></td><td style="text-align: center;"><strong>${productosUno.length}</strong></td></tr>
        <tr><td><span class="badge badge-orange">🟠 Stock Mínimo</span></td><td style="text-align: center;"><strong>${productosMinimo.length}</strong></td></tr>
        <tr><td><b>Total analizado</b></td><td style="text-align: center;"><strong>${totalFilas}</strong></td></tr>
      </table>
    </div>`;

    function generarTabla(items, columnas, encabezados, bgColor) {
      if (items.length === 0) return "";
      let html = '<table class="producto-table"><tr style="background-color: ' + bgColor + ';">';
      encabezados.forEach(e => { html += '<th>' + e + '</th>'; });
      html += '</tr>';
      items.slice(0, limitePorSeccion).forEach(p => {
        html += '<tr>';
        columnas.forEach(c => { html += '<td>' + (p[c] || '') + '</td>'; });
        html += '</tr>';
      });
      if (items.length > limitePorSeccion) {
        html += '<tr><td colspan="' + encabezados.length + '" style="text-align: center; font-style: italic;">... y ' + (items.length - limitePorSeccion) + ' más</td></tr>';
      }
      html += '</table>';
      return html;
    }

    if (productosCero.length > 0) {
      htmlContent += '<div class="alerta alerta-cero"><h4 style="margin: 0 0 8px 0; color: #721c24;">🔴 Productos Agotados (' + productosCero.length + ')</h4>';
      htmlContent += generarTabla(productosCero, ['sku', 'nombre', 'categoria'], ['SKU', 'Producto', 'Categoría'], '#f1aeb5');
      htmlContent += '</div>';
    }

    if (productosUno.length > 0) {
      htmlContent += '<div class="alerta alerta-uno"><h4 style="margin: 0 0 8px 0; color: #856404;">🟡 Última Unidad (' + productosUno.length + ')</h4>';
      htmlContent += generarTabla(productosUno, ['sku', 'nombre', 'categoria', 'stock'], ['SKU', 'Producto', 'Categoría', 'Stock'], '#ffeaa7');
      htmlContent += '</div>';
    }

    if (productosMinimo.length > 0) {
      htmlContent += '<div class="alerta alerta-minimo"><h4 style="margin: 0 0 8px 0; color: #e65100;">🟠 Stock Mínimo (' + productosMinimo.length + ')</h4>';
      htmlContent += generarTabla(productosMinimo, ['sku', 'nombre', 'categoria', 'stock', 'minimo'], ['SKU', 'Producto', 'Categoría', 'Stock', 'Mínimo'], '#ffcc80');
      htmlContent += '</div>';
    }

    htmlContent += `
    <div class="footer">
      <p>Generado automáticamente por el Sistema de Inventario.</p>
      <p>Total de productos analizados: ${totalFilas}</p>
    </div>
  </div>
</body>
</html>`;

    const destinatario = correoADMIN;
    const asunto = `📊 Alerta Stock - ${fecha}${esParcial ? ' (Extenso)' : ''}`;

    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: htmlContent
    });

    return {
      success: true,
      message: "Alerta de stock enviada exitosamente",
      resumen: {
        totalProductos: totalFilas,
        stockCero: productosCero.length,
        stockUno: productosUno.length,
        stockMinimo: productosMinimo.length,
        stockOk: productosOk.length,
        esParcial: esParcial
      }
    };

  } catch (error) {
    console.error("Error en enviarAlertaStock:", error);
    return { success: false, message: "Error al enviar alerta: " + error.message };
  }
}


function procesarCargaMasivaProductos(productos,codigoLote){try{if(!productos||!Array.isArray(productos)||productos.length===0)throw new Error("No se recibieron productos para procesar");if(!codigoLote)codigoLote="LOTE-"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd-HHmmss");const ss=SpreadsheetApp.getActiveSpreadsheet();const hoja=ss.getSheetByName("Productos");if(!hoja)throw new Error("No se encontró la hoja 'Productos'.");const lastColumn=hoja.getLastColumn();const headers=hoja.getRange(1,1,1,lastColumn).getValues()[0].map(h=>String(h).trim());const colIndex={};headers.forEach((h,idx)=>{if(h)colIndex[h]=idx+1;});const columnasRequeridas=["IdProducto","NombreProducto","PrecioProducto"];const faltantes=columnasRequeridas.filter(col=>!colIndex[col]);if(faltantes.length>0)throw new Error("Faltan columnas requeridas: "+faltantes.join(", "));const lastRow=hoja.getLastRow();const idsExistentes={};const nombresExistentes={};if(lastRow>1){const values=hoja.getRange(2,1,lastRow-1,lastColumn).getValues();values.forEach((row,idx)=>{const filaReal=idx+2;const id=row[colIndex["IdProducto"]-1];const nombre=row[colIndex["NombreProducto"]-1];if(id)idsExistentes[String(id).trim()]=filaReal;if(nombre)nombresExistentes[String(nombre).trim().toUpperCase()]=filaReal;});}let nuevos=0;let actualizados=0;let errores=[];productos.forEach((producto,index)=>{try{if(!producto||typeof producto!=="object"){errores.push("Ítem "+(index+1)+": Datos inválidos");return;}const idProducto=String(producto.IdProducto||"").trim();const nombreProducto=String(producto.NombreProducto||"").trim().toUpperCase();if(!idProducto||!nombreProducto){errores.push("Ítem "+(index+1)+": ID o Nombre vacío");return;}const esActualizacion=producto.esActualizacion===true;const filaExistente=idsExistentes[idProducto];const cuentaInfo=getCuentaContableCompleta(producto.CuentaContable||"");if(esActualizacion&&filaExistente){const updates=[];if(producto.PrecioProducto!==undefined&&colIndex["PrecioProducto"])updates.push({col:colIndex["PrecioProducto"],val:parseFloat(producto.PrecioProducto)||0});if(producto.precioVenta!==undefined&&colIndex["precioVenta"])updates.push({col:colIndex["precioVenta"],val:parseFloat(producto.precioVenta)||0});if(producto.IvaProducto!==undefined&&colIndex["IvaProducto"])updates.push({col:colIndex["IvaProducto"],val:String(producto.IvaProducto)});if(cuentaInfo.CuentaContable&&colIndex["CuentaContable"])updates.push({col:colIndex["CuentaContable"],val:cuentaInfo.CuentaContable});if(cuentaInfo.NombreCuenta&&colIndex["NombreCuenta"])updates.push({col:colIndex["NombreCuenta"],val:cuentaInfo.NombreCuenta});if(cuentaInfo.Cuenta&&colIndex["Cuenta"])updates.push({col:colIndex["Cuenta"],val:cuentaInfo.Cuenta});if(cuentaInfo.SubCuenta&&colIndex["SubCuenta"])updates.push({col:colIndex["SubCuenta"],val:cuentaInfo.SubCuenta});if(producto.CategoriaProducto!==undefined&&colIndex["CategoriaProducto"])updates.push({col:colIndex["CategoriaProducto"],val:String(producto.CategoriaProducto)});if(producto.Condicion&&colIndex["Condicion"])updates.push({col:colIndex["Condicion"],val:String(producto.Condicion)});if(colIndex["subida"])updates.push({col:colIndex["subida"],val:codigoLote});updates.forEach(u=>{hoja.getRange(filaExistente,u.col).setValue(u.val);});actualizados++;}else{if(nombresExistentes[nombreProducto]){errores.push(nombreProducto+": Ya existe un producto con este nombre");return;}const nuevaFila=new Array(lastColumn).fill("");if(colIndex["IdProducto"])nuevaFila[colIndex["IdProducto"]-1]=idProducto;if(colIndex["NombreProducto"])nuevaFila[colIndex["NombreProducto"]-1]=nombreProducto;if(colIndex["PrecioProducto"])nuevaFila[colIndex["PrecioProducto"]-1]=parseFloat(producto.PrecioProducto)||0;if(colIndex["precioVenta"])nuevaFila[colIndex["precioVenta"]-1]=parseFloat(producto.precioVenta)||0;if(colIndex["IvaProducto"])nuevaFila[colIndex["IvaProducto"]-1]=String(producto.IvaProducto||"16");if(colIndex["StockProducto"])nuevaFila[colIndex["StockProducto"]-1]=parseInt(producto.StockProducto)||0;if(colIndex["MinimoProducto"])nuevaFila[colIndex["MinimoProducto"]-1]=parseInt(producto.MinimoProducto)||5;if(colIndex["CuentaContable"])nuevaFila[colIndex["CuentaContable"]-1]=cuentaInfo.CuentaContable;if(colIndex["NombreCuenta"])nuevaFila[colIndex["NombreCuenta"]-1]=cuentaInfo.NombreCuenta;if(colIndex["Cuenta"])nuevaFila[colIndex["Cuenta"]-1]=cuentaInfo.Cuenta;if(colIndex["SubCuenta"])nuevaFila[colIndex["SubCuenta"]-1]=cuentaInfo.SubCuenta;if(colIndex["CategoriaProducto"])nuevaFila[colIndex["CategoriaProducto"]-1]=String(producto.CategoriaProducto||"");if(colIndex["ImagenProducto"])nuevaFila[colIndex["ImagenProducto"]-1]=String(producto.ImagenProducto||"");if(colIndex["Status"])nuevaFila[colIndex["Status"]-1]=String(producto.Status||"Activo");if(colIndex["Condicion"])nuevaFila[colIndex["Condicion"]-1]=String(producto.Condicion||"nuevo");if(colIndex["subida"])nuevaFila[colIndex["subida"]-1]=codigoLote;if(colIndex["tipo"])nuevaFila[colIndex["tipo"]-1]=String(producto.tipo||"Simple");if(colIndex["Medida"])nuevaFila[colIndex["Medida"]-1]=String(producto.Medida||"UNID");if(colIndex["UnidadMedida"])nuevaFila[colIndex["UnidadMedida"]-1]=String(producto.UnidadMedida||"UNID");if(colIndex["FactorConversion"])nuevaFila[colIndex["FactorConversion"]-1]=parseFloat(producto.FactorConversion)||1;if(colIndex["Compuesto"])nuevaFila[colIndex["Compuesto"]-1]=String(producto.Compuesto||"");if(colIndex["CodBarra"])nuevaFila[colIndex["CodBarra"]-1]=String(producto.CodBarra||"");hoja.appendRow(nuevaFila);nombresExistentes[nombreProducto]=true;idsExistentes[idProducto]=true;nuevos++;}}catch(itemError){console.error("Error procesando item:",producto,itemError);errores.push((producto.IdProducto||"item "+(index+1))+": "+itemError.message);}});return{success:true,message:"Proceso completado. "+nuevos+" nuevos, "+actualizados+" actualizados.",detalles:{nuevos:nuevos,actualizados:actualizados,errores:errores.length,totalProcesados:nuevos+actualizados},errores:errores.length>0?errores:null,lote:codigoLote};}catch(error){console.error("Error general en carga masiva:",error);return{success:false,message:error.toString(),error:error.stack};}}

// =============================================
/**              CORREO                       */
// =============================================
function sendEmailWithTable(htmlTable) {
    var emailAddress = correoADMIN;
    var subject = "Datos de Cobranzas";
    var body = "Aquí están los datos de cobranzas:<br><br>" + htmlTable;

    MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: body,
    });
}
function guardarAbono(nroFactura, abono) {  
  const sheetCompras = getSheetByName('Compras');  

  const dataRange = sheetCompras.getDataRange();  
  const values = dataRange.getValues();  

  console.log('Número de factura:', nroFactura);  
  console.log('Abono:', abono);  

  let rowIndex = -1;  
  for (let i = 1; i < values.length; i++) {
    if (values[i][4] === nroFactura) {
      rowIndex = i;  
      break;  
    }  
  }  

  if (rowIndex === -1) {  
    return ContentService.createTextOutput(JSON.stringify({ error: "Factura no encontrada" })).setMimeType(ContentService.MimeType.JSON);  
  }  

  values[rowIndex][12] = parseFloat(values[rowIndex][12]) + abono;
  values[rowIndex][13] = new Date();
  sheetCompras.getRange(rowIndex + 1, 13, 1, 2).setValues([values[rowIndex].slice(12, 14)]);

  const saldo = parseFloat(values[rowIndex][11]);  
  values[rowIndex][11] = saldo - abono;  
  sheetCompras.getRange(rowIndex + 1, 12, 1, 1).setValues([[values[rowIndex][11]]]);

  values[rowIndex][14] = "Cobrada";
  sheetCompras.getRange(rowIndex + 1, 15, 1, 1).setValues([[values[rowIndex][14]]]);

  const rowData = values[rowIndex];  

  const transactionSummary = {  
    NombreProveedor: rowData[0],  
    RifProveedor: rowData[1],  
    NroFactura: parseInt(rowData[4]),
    FechaFactura: new Date(rowData[2]),  
    Total: parseFloat(rowData[10]),  
    Saldo: parseFloat(values[rowIndex][11]),
    Abono: parseFloat(values[rowIndex][12]),
    FechaAbono: new Date(rowData[13]),  
    CobradaPagada: values[rowIndex][14]
  };  

  const sheetCobrosPagos = getSheetByName('CobrosPagos');  

  const cobrosPagosDataRange = sheetCobrosPagos.getDataRange();  
  const cobrosPagosValues = cobrosPagosDataRange.getValues();  

  const newRowIndex = cobrosPagosValues.length + 1;  
  sheetCobrosPagos.getRange(newRowIndex, 1, 1, Object.keys(transactionSummary).length)  
    .setValues([Object.values(transactionSummary)]);  

  return ContentService.createTextOutput(JSON.stringify({ result: "success" })).setMimeType(ContentService.MimeType.JSON);  
}
function generarId() {
    return Math.floor(100000 + Math.random() * 900000);
}

function enviarTablaPorCorreo(destinatario, asunto, cuerpo, charts, cc, bcc) {
  try {
    Logger.log(`Recibido para 'destinatario': [${destinatario}] (Tipo: ${typeof destinatario})`);
    
    if (typeof destinatario !== 'string' || !destinatario.trim()) {
      throw new Error("El destinatario está vacío o no es un texto válido.");
    }
    
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(destinatario.trim())) {
      throw new Error(`El formato del correo electrónico '${destinatario}' no es válido.`);
    }

    Logger.log(`--- Iniciando envío de correo a: ${destinatario} ---`);
    Logger.log(`Asunto: ${asunto}`);
    
    const emailOptions = {
      to: destinatario.trim(),
      subject: asunto,
      htmlBody: cuerpo,
      name: "Sistema de Gestión TaxFlow"
    };

    const inlineImages = {};
    if (charts && Array.isArray(charts) && charts.length > 0) {
      Logger.log(`Procesando ${charts.length} gráfico(s)...`);
      charts.forEach(function(chart) {
        if (chart && chart.id && chart.data) {
          try {
            const base64Data = chart.data.split(",")[1];
            const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/png", chart.id);
            inlineImages[chart.id] = blob;
          } catch (e) {
            Logger.log(`Error al procesar el gráfico (ID: ${chart.id}): ${e.message}`);
          }
        }
      });
      if (Object.keys(inlineImages).length > 0) {
        emailOptions.inlineImages = inlineImages;
      }
    }

    if (cc && cc.toString().trim() !== "") emailOptions.cc = cc.trim();
    if (bcc && bcc.toString().trim() !== "") emailOptions.bcc = bcc.trim();
    
    MailApp.sendEmail(emailOptions);
    
    Logger.log("¡Correo enviado exitosamente!");
    return { success: true, message: "Correo enviado exitosamente." };

  } catch (error) {
    Logger.log(`--- ERROR FATAL al enviar correo: ${error.toString()} ---`);
    Logger.log(`Stack Trace: ${error.stack}`);
    throw new Error(`No se pudo enviar el correo. Causa: ${error.message}`);
  }
}
function testEnvioCorreo() {
  const resultado = enviarTablaPorCorreo(
    'tu-email-de-prueba@gmail.com', // Cambiar por tu email
    'Test TaxFlow - ' + new Date().toLocaleString(),
    '<h1>Test de envío</h1><p>Este es un correo de prueba del sistema TaxFlow.</p>',
    null
  );
  
  Logger.log('Resultado test: ' + JSON.stringify(resultado));
  return resultado;
}
function probarEnvioCorreo() {
  try {
    const destinatario = Session.getActiveUser().getEmail();
    const asunto = "Prueba - TaxFlow";
    const cuerpo = `
      <html>
        <body>
          <h1>Prueba de Envío</h1>
          <p>Este es un correo de prueba enviado desde TaxFlow.</p>
          <p>Fecha: ${new Date().toLocaleString()}</p>
          <p>Si recibes este correo, la configuración es correcta.</p>
        </body>
      </html>
    `;
    
    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: cuerpo,
      name: "TaxFlow Test"
    });
    
    return "Correo de prueba enviado a: " + destinatario;
  } catch (error) {
    return "Error: " + error.toString();
  }
}

// =============================================
/**  FUNCIONES TELEGRAM                       */
// =============================================
function getChatIdPorEmail(email) {
  try {
    const sheet = getSheetByName('UsuariosTelegram');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        return data[i][1].toString();
      }
    }
    return null;
  } catch (e) {
    console.error("Error al buscar Chat ID por email: " + e.message);
    return null;
  }
}
function enviarNotificacionTelegram(resumenTabla, mensajeAdicional) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const botToken = properties.getProperty('AAEVhJo1NO_Ja622ojKZxajK1FW6QLNrCyk');
    const adminChatId = properties.getProperty('2009802525');

    if (!botToken || !adminChatId) {
      throw new Error('El Token del Bot o el Chat ID del admin no están configurados en las Propiedades del Script.');
    }

    const emailUsuarioActual = Session.getActiveUser().getEmail();
    const usuarioChatId = getChatIdPorEmail(emailUsuarioActual);

    let textoMensaje = `🔔 *Notificación de Inventario*\n\n`;
    textoMensaje += `*Generado por:* ${emailUsuarioActual}\n\n`;
    textoMensaje += `*Resumen del Reporte:*\n\`\`\`\n${resumenTabla}\n\`\`\`\n\n`;
    if (mensajeAdicional) {
      textoMensaje += `*Mensaje Adicional:*\n_${mensajeAdicional}_`;
    }

    const destinatarios = new Set([adminChatId]);
    if (usuarioChatId) {
      destinatarios.add(usuarioChatId);
    }

    destinatarios.forEach(chatId => {
      enviarMensajeTelegram(botToken, chatId, textoMensaje);
    });

    return { success: true, message: 'Notificación enviada a ' + destinatarios.size + ' destinatario(s).' };
  } catch (error) {
    console.error('Error en enviarNotificacionTelegram:', error);
    throw new Error('Error al enviar la notificación: ' + error.message);
  }
}
function enviarNotificacionNuevoProducto(nombre, precio, stock, categoria) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const botToken = properties.getProperty('TELEGRAM_BOT_TOKEN');
    const miChatId = properties.getProperty('MI_CHAT_ID');

    if (!botToken || !miChatId) {
      console.error("Token del bot o Chat ID no configurados.");
      return;
    }

    const emailUsuario = Session.getActiveUser().getEmail();
    
    const mensaje = `✅ *NUEVO PRODUCTO AGREGADO*\n\n` +
                   `📦 *Nombre:* ${nombre || 'N/A'}\n` +
                   `📋 *Categoría:* ${categoria || 'N/A'}\n` +
                   `💰 *Precio:* $${precio || '0.00'}\n` +
                   `📊 *Stock:* ${stock || '0'}\n` +
                   `👤 *Agregado por:* ${emailUsuario}\n` +
                   `⏰ *Fecha:* ${new Date().toLocaleString()}\n` +
                   `🌐 *Sistema:* Inventario Pro`;

    enviarMensajeTelegram(botToken, miChatId, mensaje);
    console.log("Notificación enviada a tu Chat ID:", miChatId);
    
  } catch (e) {
    console.error("Error en enviarNotificacionNuevoProducto:", e.message);
  }
}
function enviarMensajeTelegram(token, chatId, texto) {
  try {
    const url = `https://api.telegram.org/bot${token}/sendMessage`;
    const payload = {
      'chat_id': chatId,
      'text': texto,
      'parse_mode': 'Markdown'
    };
    
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      console.error("Error de Telegram. Código:", responseCode, "Respuesta:", response.getContentText());
    }
    
    return responseCode === 200;
    
  } catch (e) {
    console.error("Error en enviarMensajeTelegram:", e.message);
    return false;
  }
}

// =============================================
/** FUNCIONES AJUSTES INVENTARIO                      */
// =============================================
function guardarAjusteInventario(dataAjuste) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detalleSheet = ss.getSheetByName("DetalleFacturas");
    const productosSheet = ss.getSheetByName("Productos");
    
    if (!detalleSheet) throw new Error("No se encontró la hoja DetalleFacturas");
    if (!productosSheet) throw new Error("No se encontró la hoja Productos");
    
    // ESTRUCTURA EXACTA DE COLUMNAS:
    // A: NroFactura, B: Producto, C: Categoria, D: Cantidad, E: Precio, 
    // F: %IVA, G: MontoIVA, H: PrecioTotal, I: PrecioUSD, J: %IVAUSD, 
    // K: MontoIVAUSD, L: PrecioTotalUSD, M: Transaccion, N: CuentaContable, 
    // O: NombreContable, P: DescripcionTransaccion, Q: usuario, R: fecha, S: tipo
    
    // Generar número de ajuste
    const ultimaFila = detalleSheet.getLastRow();
    let ultimoNumero = 0;
    
    if (ultimaFila > 1) {
      const colA = detalleSheet.getRange(2, 1, ultimaFila - 1, 1).getValues();
      colA.forEach(function(fila) {
        var numFact = fila[0];
        if (numFact && numFact.toString().startsWith('AJ-') && !numFact.toString().startsWith('ANULADO')) {
          var num = parseInt(numFact.toString().split('-')[1]);
          if (num > ultimoNumero) ultimoNumero = num;
        }
      });
    }
    
    const nuevoNumero = ultimoNumero + 1;
    const nroAjuste = 'AJ-' + String(nuevoNumero).padStart(5, '0');
    
    // Preparar fecha actual
    const fechaActual = new Date();
    const fechaFormateada = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const fechaUsuario = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "dd/MM/yyyy, h:mm:ss a");
    
    // Obtener usuario
    var usuario = "Sistema";
    try {
      usuario = Session.getActiveUser().getEmail();
    } catch(e) {
      usuario = dataAjuste.Usuario || "Sistema";
    }
    
    var filasInsertar = [];
    
    dataAjuste.items.forEach(function(item) {
      // Cálculos correctos
      var cantidad = parseFloat(item.Cantidad) || 0;
      var precioUSD = parseFloat(item.CostoUnitarioUSD) || 0;
      var precioTotalUSD = cantidad * precioUSD;
      
      // Precio en Bs (si la tasa es 1, usa el mismo valor en USD)
      var tasa = parseFloat(dataAjuste.TasaCambio) || 1;
      var precioBs = precioUSD * tasa;
      var precioTotalBs = precioTotalUSD * tasa;
      
      // IVA en 0 para ajustes
      var porcIVA = "0,00%";
      var montoIVA = 0;
      var porcIVAUSD = "0,00%";
      var montoIVAUSD = 0;
      
      // MAPEO EXACTO A COLUMNAS:
      filasInsertar.push([
        nroAjuste,                    // A - NroFactura
        item.Producto,                // B - Producto (nombre del producto)
        dataAjuste.Motivo || "",      // C - Categoria (usamos el motivo aquí)
        cantidad,                     // D - Cantidad
        precioBs,                     // E - Precio (unitario en Bs)
        porcIVA,                      // F - %IVA
        montoIVA,                     // G - MontoIVA
        precioTotalBs,                // H - PrecioTotal (total en Bs)
        precioUSD,                    // I - PrecioUSD (unitario en USD)
        porcIVAUSD,                   // J - %IVAUSD
        montoIVAUSD,                  // K - MontoIVAUSD
        precioTotalUSD,               // L - PrecioTotalUSD (total en USD)
        dataAjuste.Transaccion,       // M - Transaccion (Ajuste_Entrada/Salida)
        item.CuentaContable || "",    // N - CuentaContable
        dataAjuste.Responsable || "", // O - NombreContable (responsable)
        dataAjuste.Motivo || "",      // P - DescripcionTransaccion (motivo/detalle)
        usuario + " | " + fechaUsuario, // Q - usuario (email + fecha/hora)
        fechaFormateada,              // R - fecha (solo fecha YYYY-MM-DD)
        dataAjuste.TipoAjuste         // S - tipo (Entrada o Salida)
      ]);
    });
    
    // Insertar todas las filas (19 columnas A-S)
    if (filasInsertar.length > 0) {
      const rangoInsertar = detalleSheet.getRange(ultimaFila + 1, 1, filasInsertar.length, 19);
      rangoInsertar.setValues(filasInsertar);
    }
    
    // ACTUALIZAR INVENTARIO
    var itemsParaStock = dataAjuste.items.map(function(item) {
      return {
        producto: item.Producto,
        cantidad: item.Cantidad
      };
    });
    
    // Entrada = suma stock (Compra), Salida = resta stock (Venta)
    var tipoTransaccionStock = dataAjuste.TipoAjuste === "Entrada" ? "Compra" : "Venta";
    var resultadoStock = actualizarStockProductos(itemsParaStock, tipoTransaccionStock);
    
    Logger.log("✓ Ajuste guardado: " + nroAjuste);
    Logger.log("  Filas: " + filasInsertar.length);
    Logger.log("  Stock actualizado: " + resultadoStock.actualizados);
    
    return {
      status: "success",
      nroAjuste: nroAjuste,
      message: "Ajuste " + nroAjuste + " procesado. " + dataAjuste.items.length + " productos. Stock: " + resultadoStock.actualizados,
      stockResult: resultadoStock
    };
    
  } catch (error) {
    Logger.log("✗ Error en guardarAjusteInventario: " + error.toString());
    return {
      status: "error",
      message: error.toString()
    };
  }
}
function obtenerHistorialAjustes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detalleSheet = ss.getSheetByName("DetalleFacturas");
    
    if (!detalleSheet) throw new Error("No se encontró la hoja DetalleFacturas");
    
    const ultimaFila = detalleSheet.getLastRow();
    if (ultimaFila <= 1) return [];
    
    // Leer columnas A-S (19 columnas)
    const datos = detalleSheet.getRange(2, 1, ultimaFila - 1, 19).getValues();
    
    var ajustes = [];
    
    datos.forEach(function(fila) {
      var transaccion = fila[12]; // Columna M (índice 12)
      
      // Filtrar solo ajustes
      if (transaccion === "Ajuste_Entrada" || transaccion === "Ajuste_Salida") {
        ajustes.push({
          NroFactura: fila[0] ? fila[0].toString() : "",           // A
          Producto: fila[1] ? fila[1].toString() : "",             // B
          Categoria: fila[2] ? fila[2].toString() : "",            // C
          Cantidad: fila[3],                                       // D
          Precio: fila[4],                                         // E
          PorcIVA: fila[5] ? fila[5].toString() : "",              // F
          MontoIVA: fila[6],                                       // G
          PrecioTotal: fila[7],                                    // H
          PrecioUSD: fila[8],                                      // I
          PorcIVAUSD: fila[9] ? fila[9].toString() : "",           // J
          MontoIVAUSD: fila[10],                                   // K
          PrecioTotalUSD: fila[11],                                // L
          Transaccion: transaccion ? transaccion.toString() : "",  // M
          CuentaContable: fila[13] ? fila[13].toString() : "",      // N
          NombreContable: fila[14] ? fila[14].toString() : "",      // O
          DescripcionTransaccion: fila[15] ? fila[15].toString() : "", // P
          Usuario: fila[16] ? fila[16].toString() : "",             // Q
          Fecha: fila[17] ? fila[17].toString() : "",               // R
          Tipo: fila[18] ? fila[18].toString() : ""                 // S <-- COLUMNA TIPO (Entrada/Salida)
        });
      }
    });
    
    return ajustes;
    
  } catch (error) {
    Logger.log("Error en obtenerHistorialAjustes: " + error.toString());
    throw error;
  }
}
function obtenerAjusteParaImprimir(nroAjuste) {
  try {
    const ajustes = obtenerHistorialAjustes();
    const itemsAjuste = ajustes.filter(function(a) { 
      return a.NroFactura === nroAjuste || a.NroFactura === "ANULADO-" + nroAjuste; 
    });
    
    if (itemsAjuste.length === 0) {
      return null;
    }
    
    var primerItem = itemsAjuste[0];
    
    // Usar el campo Tipo de la columna S, o derivarlo de Transaccion
    var tipoAjuste = primerItem.Tipo || (primerItem.Transaccion === "Ajuste_Entrada" ? "Entrada" : "Salida");
    
    return {
      NroAjuste: nroAjuste,
      Fecha: primerItem.Fecha,
      TipoAjuste: tipoAjuste,  // Este es el campo que usa el HTML
      Tipo: tipoAjuste,        // Por compatibilidad
      Transaccion: primerItem.Transaccion,
      Responsable: primerItem.NombreContable,
      Motivo: primerItem.Categoria || primerItem.DescripcionTransaccion,
      Usuario: primerItem.Usuario,
      Items: itemsAjuste.map(function(item) {
        return {
          Producto: item.Producto,
          Cantidad: item.Cantidad,
          PrecioUSD: item.PrecioUSD,
          PrecioTotalUSD: item.PrecioTotalUSD,
          TipoLinea: item.Tipo || tipoAjuste,  // Usar el tipo de la línea o el general
          Tipo: item.Tipo || tipoAjuste
        };
      })
    };
    
  } catch (error) {
    Logger.log("Error en obtenerAjusteParaImprimir: " + error.toString());
    throw error;
  }
}
function anularAjusteInventario(nroAjuste) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detalleSheet = ss.getSheetByName("DetalleFacturas");
    
    if (!detalleSheet) throw new Error("No se encontró la hoja DetalleFacturas");
    
    const ultimaFila = detalleSheet.getLastRow();
    const datos = detalleSheet.getRange(2, 1, ultimaFila - 1, 19).getValues();
    
    var filasAnular = [];
    var itemsRevertir = [];
    
    // Buscar filas del ajuste
    for (var i = 0; i < datos.length; i++) {
      var nroFact = datos[i][0].toString();
      if (nroFact === nroAjuste || nroFact === "ANULADO-" + nroAjuste) {
        var numFila = i + 2;
        
        // Si ya está anulado, no hacer nada
        if (nroFact.startsWith("ANULADO")) {
          return {
            status: "error",
            message: "El ajuste " + nroAjuste + " ya está anulado"
          };
        }
        
        filasAnular.push(numFila);
        
        itemsRevertir.push({
          producto: datos[i][1],        // B - Producto
          cantidad: datos[i][3],        // D - Cantidad
          tipoTransaccion: datos[i][12], // M - Transaccion
          tipo: datos[i][18]            // S - tipo (Entrada/Salida)
        });
      }
    }
    
    if (filasAnular.length === 0) {
      return {
        status: "error",
        message: "No se encontró el ajuste " + nroAjuste
      };
    }
    
    // Marcar como anulado agregando prefijo
    filasAnular.forEach(function(fila) {
      var valorActual = detalleSheet.getRange(fila, 1).getValue();
      detalleSheet.getRange(fila, 1).setValue("ANULADO-" + valorActual);
    });
    
    // REVERTIR STOCK
    var itemsRevertirStock = itemsRevertir.map(function(item) {
      return {
        producto: item.producto,
        cantidad: item.cantidad
      };
    });
    
    // Si era Entrada (sumó stock), ahora restamos (Venta)
    // Si era Salida (restó stock), ahora sumamos (Compra)
    var tipoRevertir = itemsRevertir[0].tipo === "Entrada" ? "Venta" : "Compra";
    
    var resultadoStock = actualizarStockProductos(itemsRevertirStock, tipoRevertir);
    
    Logger.log("✓ Ajuste anulado: " + nroAjuste);
    Logger.log("  Filas anuladas: " + filasAnular.length);
    Logger.log("  Stock revertido: " + resultadoStock.actualizados);
    
    return {
      status: "success",
      message: "Ajuste " + nroAjuste + " anulado. Stock revertido en " + resultadoStock.actualizados + " productos."
    };
    
  } catch (error) {
    Logger.log("✗ Error en anularAjusteInventario: " + error.toString());
    return {
      status: "error",
      message: error.toString()
    };
  }
}
function getActiveUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch(e) {
    return "Sistema";
  }
}
function generarCorrelativoAjuste(ss) {
  try {
    // Buscar en la hoja DetalleFacturas el último número de ajuste
    var hojaDetalle = ss.getSheetByName("DetalleFacturas");
    
    if (!hojaDetalle) {
      // Si no existe la hoja, empezar desde 1
      return "AJ-00001";
    }
    
    var ultimaFila = hojaDetalle.getLastRow();
    var maxNumero = 0;
    
    if (ultimaFila > 1) {
      // Obtener todos los valores de la columna A (NroFactura)
      var valores = hojaDetalle.getRange(2, 1, ultimaFila - 1, 1).getValues();
      
      for (var i = 0; i < valores.length; i++) {
        var valor = valores[i][0];
        if (valor && typeof valor === 'string' && valor.startsWith('AJ-')) {
          // Extraer el número del correlativo AJ-XXXXX
          var numeroStr = valor.replace('AJ-', '');
          var numero = parseInt(numeroStr);
          if (!isNaN(numero) && numero > maxNumero) {
            maxNumero = numero;
          }
        }
      }
    }
    
    // Generar siguiente número
    var siguienteNumero = maxNumero + 1;
    var correlativo = "AJ-" + String(siguienteNumero).padStart(5, '0');
    
    return correlativo;
    
  } catch (error) {
    Logger.log("Error generando correlativo: " + error.toString());
    // En caso de error, usar timestamp como fallback
    return "AJ-" + Date.now().toString().slice(-5);
  }
}
function formatearFecha(fechaInput) {
  try {
    if (!fechaInput) return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Si viene en formato YYYY-MM-DD (del input date)
    if (typeof fechaInput === 'string' && fechaInput.includes('-')) {
      var partes = fechaInput.split('-');
      if (partes.length === 3) {
        return partes[2] + '/' + partes[1] + '/' + partes[0]; // dd/MM/yyyy
      }
    }
    
    // Si es un objeto Date
    if (fechaInput instanceof Date) {
      return Utilities.formatDate(fechaInput, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    
    return fechaInput;
  } catch (e) {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
}
function generarNroAjuste(hojaAjustes) {
    var datos = hojaAjustes.getDataRange().getValues();
    var maxNum = 0;
    
    for (var i = 1; i < datos.length; i++) {
        var nro = datos[i][0];
        if (nro && nro.toString().indexOf("AJ-") === 0) {
            var num = parseInt(nro.toString().replace("AJ-", ""));
            if (num > maxNum) maxNum = num;
        }
    }
    
    var nuevoNum = (maxNum + 1).toString();
    while (nuevoNum.length < 6) {
        nuevoNum = "0" + nuevoNum;
    }
    return "AJ-" + nuevoNum;
}
function actualizarStockInventario(items, tipo) {
    // Implementar según tu estructura de hoja de productos
    // Esta función debe actualizar el stock en tu hoja de productos
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaProductos = ss.getSheetByName("Productos");
    
    if (!hojaProductos) return;
    
    var datos = hojaProductos.getDataRange().getValues();
    
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        // Buscar producto y actualizar stock
        for (var j = 1; j < datos.length; j++) {
            if (datos[j][1] === item.Producto) { // Asumiendo que el nombre está en columna B
                var stockActual = parseFloat(datos[j][5]) || 0; // Asumiendo stock en columna F
                var nuevoStock = tipo === "Entrada" ? stockActual + item.Cantidad : stockActual - item.Cantidad;
                hojaProductos.getRange(j + 1, 6).setValue(nuevoStock); // +1 porque arrays son 0-based
                break;
            }
        }
    }
}
// =============================================
/** FUNCIONES CATEGORIAS                      */
// =============================================
function obtenerCategoriasDesdeTasas() {
  const sheet = getSheetByName('Tasas');
  if (!sheet) throw new Error('Hoja "Tasas" no encontrada');
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  const data = sheet.getRange(2, 15, lastRow, 2).getValues();
  const categorias = [];
  
  for (let i = 0; i < data.length; i++) {
    const [id, nombre] = data[i];
    if (id && nombre) {
      categorias.push({
        id: String(id).trim(),
        nombre: String(nombre).trim()
      });
    }
  }
  
  return categorias;
}
function guardarCategoriaTasas(id, nombre, esEdicion = false) {
  const sheet = getSheetByName('Tasas');
  if (!sheet) throw new Error('Hoja "Tasas" no encontrada');
  
  const data = sheet.getDataRange().getValues();
  let filaExistente = -1;
  
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][14]).trim() === id) {
      filaExistente = i + 1;
      break;
    }
  }
  
  if (esEdicion && filaExistente > 0) {
    sheet.getRange(filaExistente, 15).setValue(id);
    sheet.getRange(filaExistente, 16).setValue(nombre);
  } else if (!esEdicion && filaExistente === -1) {
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 15).setValue(id);
    sheet.getRange(newRow, 16).setValue(nombre);
  } else if (!esEdicion && filaExistente > 0) {
    throw new Error('El ID de categoría ya existe');
  }
}
function eliminarCategoriaTasas(id) {
  const sheet = getSheetByName('Tasas');
  if (!sheet) throw new Error('Hoja "Tasas" no encontrada');
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][14]).trim() === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}
function obtenerCategoriasConId() {
  return obtenerCategoriasDesdeTasas();
}


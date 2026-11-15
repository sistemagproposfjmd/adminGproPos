// ============================================================================
// SISTEMA ADMINISTRATIVO GPROPOS - BOT TELEGRAM
// Versión: 6.2 (con soporte /empresa y /cancelar)
// Archivo: telegram.gs
// ============================================================================

// ============================================================================
// SECCIÓN 0: CONFIGURACIÓN DEL BOT ⚙️
// ============================================================================
const TELEGRAM_TOKEN = "2009802525:AAEVhJo1NO_Ja622ojKZxajK1FW6QLNrCyk";
const TELEGRAM_API = `https://api.telegram.org/bot${TELEGRAM_TOKEN}`;
const SISTEMA_NOMBRE = "Sistema GproPos";
const TIMEZONE = Session.getScriptTimeZone();
const DEBUG = true;

const BOT_CONFIG = {
  emailsAdmin: ["sistemagpropos@gmail.com"],
  emailsCobranza: [],
  
  alertas: {
    deudaVencidaDias: 30,
    montoDeudaCritica: 500,
    pedidosEstancadosDias: 7,
    saldoMinimoCaja: 100
  },
  
  automatizacion: {
    horaReporteDiario: 8,
    horaBackup: 2,
    horasVerificarAlertas: 4
  },
  
  enviarEmailDiario: true,
  enviarAlertaTelegram: true,
  enviarPdfDiario: true,
  respaldoAutomatico: true,
  
  empresa: {
    nombre: "GproPos",
    slogan: "Gestión Profesional de Ventas",
    rif: "J-12345678-9",
    direccion: "Caracas, Venezuela",
    telefono: "+58 212-1234567",
    email: "contacto@gpropos.com",
    web: "https://gpropos.com",
    logo: "",
    color: "#1a73e8",
    colorSecundario: "#0d47a1"
  },
  
  drive: {
    carpetaRaiz: "GproPos_Backups",
    carpetasPorMes: true,
    mantenerMeses: 12
  }
};

// ============================================================================
// SECCIÓN 1: CONTROLADOR PRINCIPAL (WEBHOOK)
// ============================================================================
function doPost(e) {
  const inicio = Date.now();
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    const update = JSON.parse(e.postData.contents);
    const updateId = String(update.update_id || "");
    const esCallback = !!update.callback_query;
    
    // Anti-duplicados por Cache
    if (!esCallback && updateId) {
      const cache = CacheService.getScriptCache();
      const clave = "upd_" + updateId;
      if (cache.get(clave)) {
        log("⚠️ Update duplicado ignorado: " + updateId);
        return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
      }
      cache.put(clave, "1", 300);
    }

    // Descartar mensajes con más de 3 minutos de retraso
    const msgDate = update.message ? update.message.date : 0;
    const ahora = Math.floor(Date.now() / 1000);
    if (!esCallback && msgDate && (ahora - msgDate) > 180) {
      log("Update retrasado ignorado");
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    if (update.callback_query) {
      log("🔘 Callback: " + update.callback_query.data);
      manejarCallback(update.callback_query);
    } else if (update.message && update.message.text) {
      log("💬 Mensaje: " + update.message.text);
      manejarMensaje(update.message);
    }

    log("⏱️ doPost procesado en " + (Date.now() - inicio) + "ms");
  } catch (error) {
    log("❌ ERROR doPost: " + error.toString() + "\nStack: " + error.stack);
  }

  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================================
// SECCIÓN 2: ROUTER DE MENSAJES
// ============================================================================
function manejarMensaje(message) {
  const chatId = message.chat.id;
  const usuario = (message.from && message.from.first_name) || "Colega";
  const textoOriginal = String(message.text || "").trim();

  const texto = textoOriginal
    .replace(/[\u{1F300}-\u{1F9FF}]/gu, "")
    .replace(/[\u{2600}-\u{27BF}]/gu, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

  const partes = texto.split(" ");
  const comando = partes[0].replace(/^\//, "").split("@")[0];
  const parametro = partes.slice(1).join(" ").trim();

  switch (comando) {
    case "start":
    case "menu":
    case "inicio":
      enviarBienvenida(chatId, usuario);
      break;
    
    case "catalogo":
    case "comprar":
    case "vender":
      mostrarCatalogoTelegram(chatId, 1);
      break;

    case "carrito":
      verCarritoTelegram(chatId);
      break;

    case "cancelar":
    case "vaciar":
      // 🔧 NUEVO: vaciar carrito sin guardar
      vaciarCarritoUsuario(chatId);
      enviarMensaje(chatId, "🗑️ Carrito vaciado. Puedes empezar de nuevo con /catalogo.", tecladoPrincipal());
      break;

    case "buscar_cliente":
      buscarClienteComando(chatId, parametro);
      break;

    case "nuevo_cliente":
      crearClienteRapidoComando(chatId, parametro);
      break;

    case "empresa":
      // 🔧 NUEVO: configurar hoja de la empresa para este chat
      comandoConfigurarEmpresa(chatId, parametro);
      break;

    case "pedidos":
    case "pedido":
      enviarMenuPedidos(chatId);
      break;

    case "cobranzas":
    case "cobranza":
    case "deudores":
      enviarMenuCobranzas(chatId);
      break;

    case "dashboard":
    case "resumen":
      enviarDashboard(chatId);
      break;

    case "tasas":
    case "tasa":
      enviarReporteTasas(chatId);
      break;

    case "reportes":
    case "reporte":
      enviarMenuReportes(chatId);
      break;

    case "estado":
    case "ping":
      enviarEstadoServidor(chatId);
      break;

    case "ayuda":
    case "help":
      enviarAyuda(chatId);
      break;

    case "saldo":
      if (!parametro) {
        enviarMensaje(chatId, "⚠️ <b>Falta el cliente o RIF</b>\n\nEjemplo: <code>/saldo J-12345678</code>", tecladoPrincipal());
      } else {
        enviarMensaje(chatId, consultarSaldoCliente(parametro), tecladoPrincipal());
      }
      break;

    case "proveedor":
      if (!parametro) {
        enviarMensaje(chatId, "⚠️ <b>Falta el proveedor</b>\n\nEjemplo: <code>/proveedor FARMATODO</code>", tecladoPrincipal());
      } else {
        enviarMensaje(chatId, consultarSaldoProveedor(parametro), tecladoPrincipal());
      }
      break;

    case "email":
    case "correo":
      enviarReportePorEmail(chatId);
      break;

    case "pdf":
      enviarPdfPorEmail(chatId);
      break;

    case "backup":
    case "respaldo":
      realizarRespaldoManual(chatId);
      break;

    case "backups":
    case "respaldos":
      enviarHistorialRespaldos(chatId);
      break;

    case "suscribir":
      suscribirUsuario(chatId, message.from);
      break;

    case "desuscribir":
      desuscribirUsuario(chatId);
      break;

    case "config":
      mostrarConfiguracion(chatId);
      break;

    default:
      if (texto.includes("catalogo") || texto.includes("crear pedido") || texto.includes("comprar") || texto.includes("vender")) {
        mostrarCatalogoTelegram(chatId, 1);
      }
      if (texto.includes("pedido")) enviarMenuPedidos(chatId);
      else if (texto.includes("cobrar") || texto.includes("deudor") || texto.includes("cobranza")) enviarMenuCobranzas(chatId);
      else if (texto.includes("resumen") || texto.includes("dashboard") || texto.includes("balance")) enviarDashboard(chatId);
      else if (texto.includes("tasa")) enviarReporteTasas(chatId);
      else if (texto.includes("reporte") && !texto.includes("email")) enviarMenuReportes(chatId);
      else if (texto.includes("estado") || texto.includes("servidor")) enviarEstadoServidor(chatId);
      else if (texto.includes("ayuda") || texto.includes("help")) enviarAyuda(chatId);
      else if (texto.includes("email") || texto.includes("correo")) enviarReportePorEmail(chatId);
      else if (texto.includes("backup") || texto.includes("respaldo")) realizarRespaldoManual(chatId);
      else enviarMenuPrincipal(chatId, usuario);
  }
}

// ============================================================================
// SECCIÓN 3: ROUTER DE CALLBACKS (BOTONES INLINE)
// ============================================================================
function manejarCallback(query) {
  const chatId = query.message.chat.id;
  const messageId = query.message.message_id;
  const dataOriginal = query.data || "";

  responderCallback(query.id);
  
  const data = dataOriginal.trim().toUpperCase();
  const [accion, param] = data.split(":");

  try {
    switch (accion) {
      case "MENU": 
        enviarMenuPrincipal(chatId, null, messageId); 
        break;
      case "CATALOGO":
        mostrarCatalogoTelegram(chatId, param || 1, messageId);
        break;

      case "CART_ADD":
        agregarProductoCarrito(chatId, param, query.id);
        break;

      case "CART_VER":
        verCarritoTelegram(chatId, messageId);
        break;

      case "CART_INC":
        modificarCantidadItem(chatId, param, 1, messageId);
        break;

      case "CART_DEC":
        modificarCantidadItem(chatId, param, -1, messageId);
        break;

      case "CART_DEL":
        eliminarItemCarrito(chatId, param, messageId);
        break;

      case "CART_CLEAR":
        vaciarCarritoUsuario(chatId);
        verCarritoTelegram(chatId, messageId);
        break;

      case "CLIENTE_MENU":
        menuSeleccionCliente(chatId, messageId);
        break;

      case "CLIENTE_SET":
        asignarClienteAlCarrito(chatId, param, messageId);
        break;

      case "PEDIDO_GUARDAR":
        const usuarioNombre = (query.from && query.from.first_name) 
          ? `${query.from.first_name} ${query.from.last_name || ''}`.trim() 
          : "Vendedor Telegram";
        procesarConfirmacionPedido(chatId, usuarioNombre, messageId);
        break;

      case "NOOP":
        // Botón informativo, no hace nada
        break;

      case "PEDIDOS":
      case "PEDIDOS_REFRESH":
        editarMensaje(chatId, messageId, renderPedidos(), tecladoPedidos());
        break;
        
      case "PEDIDOS_VER":
        if (param) editarMensaje(chatId, messageId, renderDetallePedido(param), tecladoVolver("PEDIDOS"));
        break;
        
      case "COBRANZAS":
      case "COBRANZAS_REFRESH":
        editarMensaje(chatId, messageId, renderTopDeudores(), tecladoCobranzas());
        break;
        
      case "PROVEEDORES":
        editarMensaje(chatId, messageId, renderTopProveedores(), tecladoVolver("COBRANZAS"));
        break;
        
      case "DASHBOARD":
        editarMensaje(chatId, messageId, renderDashboard(), tecladoDashboard());
        break;
        
      case "FLUJO":
        editarMensaje(chatId, messageId, renderFlujoCaja(), tecladoVolver("DASHBOARD"));
        break;
        
      case "TENDENCIAS":
        editarMensaje(chatId, messageId, renderTendencias(), tecladoVolver("DASHBOARD"));
        break;
        
      case "ALERTAS":
        editarMensaje(chatId, messageId, renderAlertas(), tecladoVolver("DASHBOARD"));
        break;
        
      case "TASAS":
        editarMensaje(chatId, messageId, renderTasas(), tecladoVolver("MENU"));
        break;
        
      case "ESTADO":
        editarMensaje(chatId, messageId, renderEstado(), tecladoVolver("MENU"));
        break;
        
      case "AYUDA":
        editarMensaje(chatId, messageId, renderAyuda(), tecladoVolver("MENU"));
        break;
        
      case "REPORTES":
        editarMensaje(chatId, messageId, renderMenuReportes(), tecladoReportes());
        break;
        
      case "EMAIL_ENVIAR":
        editarMensaje(chatId, messageId, "📧 <i>Procesando reporte completo y enviando correo...</i>", null);
        const resultado = enviarReporteCompletoPorEmail("manual");
        editarMensaje(chatId, messageId, resultado, tecladoVolver("REPORTES"));
        break;
        
      case "EMAIL_PEDIDOS": 
        enviarReporteSegmentado(chatId, "pedidos", messageId); 
        break;
        
      case "EMAIL_COBRANZAS": 
        enviarReporteSegmentado(chatId, "cobranzas", messageId); 
        break;
        
      case "EMAIL_DASHBOARD": 
        enviarReporteSegmentado(chatId, "dashboard", messageId); 
        break;
        
      case "EMAIL_EXCEL": 
        enviarExcelPorEmail(chatId, messageId); 
        break;
        
      case "PDF_ENVIAR": 
        enviarPdfPorEmail(chatId, messageId); 
        break;
        
      case "BACKUP_CREAR": 
        realizarRespaldoManual(chatId, messageId); 
        break;
        
      case "BACKUP_HISTORIAL": 
        enviarHistorialRespaldos(chatId, messageId); 
        break;
        
      default:
        enviarMensaje(chatId, `⚠️ Acción no reconocida: <code>${escapeHtml(dataOriginal)}</code>`, tecladoPrincipal());
    }
  } catch (err) {
    log("❌ Error en manejarCallback: " + err.toString() + "\n" + err.stack);
    enviarMensaje(chatId, `❌ Error en el proceso: <code>${escapeHtml(err.toString())}</code>`, tecladoPrincipal());
  }
}

// ============================================================================
// SECCIÓN 4: TECLADOS (UI)
// ============================================================================
function tecladoPrincipal() {
  return {
    keyboard: [
      [{ text: "🛍️ Crear Pedido / Catálogo" }, { text: "📦 Pedidos" }],
      [{ text: "💰 Cobranzas" }, { text: "📊 Dashboard" }],
      [{ text: "💵 Tasas" }, { text: "📧 Reportes" }],
      [{ text: "⚡ Estado" }, { text: "❓ Ayuda" }]
    ],
    resize_keyboard: true,
    is_persistent: true
  };
}
function tecladoPedidos() {
  return {
    inline_keyboard: [
      [
        { text: "🔄 Actualizar", callback_data: "PEDIDOS_REFRESH" },
        { text: "📧 Enviar Email", callback_data: "EMAIL_PEDIDOS" }
      ],
      [{ text: "🏠 Menú Principal", callback_data: "MENU" }]
    ]
  };
}
function tecladoCobranzas() {
  return {
    inline_keyboard: [
      [
        { text: "🔄 Actualizar", callback_data: "COBRANZAS_REFRESH" },
        { text: "🏢 Ver Proveedores", callback_data: "PROVEEDORES" }
      ],
      [
        { text: "📧 Enviar Email", callback_data: "EMAIL_COBRANZAS" },
        { text: "🏠 Menú Principal", callback_data: "MENU" }
      ]
    ]
  };
}
function tecladoDashboard() {
  return {
    inline_keyboard: [
      [
        { text: "💸 Flujo de Caja", callback_data: "FLUJO" },
        { text: "📈 Tendencias", callback_data: "TENDENCIAS" }
      ],
      [
        { text: "🚨 Ver Alertas", callback_data: "ALERTAS" },
        { text: "📧 Enviar Email", callback_data: "EMAIL_DASHBOARD" }
      ],
      [{ text: "🏠 Menú Principal", callback_data: "MENU" }]
    ]
  };
}
function tecladoReportes() {
  return {
    inline_keyboard: [
      [{ text: "📧 Reporte Completo por Correo", callback_data: "EMAIL_ENVIAR" }],
      [{ text: "📄 Generar PDF Ejecutivo", callback_data: "PDF_ENVIAR" }],
      [
        { text: "📦 Email Pedidos", callback_data: "EMAIL_PEDIDOS" },
        { text: "💰 Email Cobranzas", callback_data: "EMAIL_COBRANZAS" }
      ],
      [
        { text: "📊 Email Dashboard", callback_data: "EMAIL_DASHBOARD" },
        { text: "📎 Descargar Excel", callback_data: "EMAIL_EXCEL" }
      ],
      [
        { text: "💾 Crear Backup Drive", callback_data: "BACKUP_CREAR" },
        { text: "📜 Ver Historial", callback_data: "BACKUP_HISTORIAL" }
      ],
      [{ text: "🏠 Menú Principal", callback_data: "MENU" }]
    ]
  };
}

function tecladoVolver(destino) {
  const mapa = {
    MENU: "🏠 Menú Principal",
    PEDIDOS: "📦 Volver a Pedidos",
    DASHBOARD: "📊 Volver a Dashboard",
    COBRANZAS: "💰 Volver a Cobranzas",
    REPORTES: "📧 Volver a Reportes"
  };
  return {
    inline_keyboard: [
      [{ text: mapa[destino] || "🏠 Menú", callback_data: destino }]
    ]
  };
}

// ============================================================================
// SECCIÓN 5: VISTAS Y MENÚS
// ============================================================================
function enviarMenuPrincipal(chatId, usuario, messageId) {
  const nombre = escapeHtml(usuario || "Colega");
  const hora = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm");

  const texto =
    `🏠 <b>MENÚ PRINCIPAL — ${escapeHtml(BOT_CONFIG.empresa.nombre)}</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `👋 Hola <b>${nombre}</b>\n` +
    `🕐 <i>${hora}</i>\n\n` +
    `📦 <b>Pedidos:</b> Gestión de pedidos abiertos\n` +
    `💰 <b>Cobranzas:</b> Saldos por cobrar y pagar\n` +
    `📊 <b>Dashboard:</b> Rentabilidad y caja\n` +
    `💵 <b>Tasas:</b> Monitor de cambio\n` +
    `📧 <b>Reportes:</b> PDF, correos y respaldos\n` +
    `⚡ <b>Estado:</b> Diagnóstico del sistema\n\n` +
    `Selecciona una opción del panel:`;

  const teclado = {
    inline_keyboard: [
      [
        { text: "📦 Pedidos", callback_data: "PEDIDOS" },
        { text: "💰 Cobranzas", callback_data: "COBRANZAS" }
      ],
      [
        { text: "📊 Dashboard", callback_data: "DASHBOARD" },
        { text: "💵 Tasas", callback_data: "TASAS" }
      ],
      [
        { text: "📧 Reportes Email", callback_data: "REPORTES" },
        { text: "⚡ Estado", callback_data: "ESTADO" }
      ],
      [{ text: "❓ Ayuda", callback_data: "AYUDA" }]
    ]
  };

  if (messageId) editarMensaje(chatId, messageId, texto, teclado);
  else enviarMensaje(chatId, texto, teclado);
}
function enviarBienvenida(chatId, usuario) {
  const nombre = escapeHtml(usuario || "Usuario");
  const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy - hh:mm a");

  const texto =
    `🚀 <b>¡Bienvenido a ${escapeHtml(SISTEMA_NOMBRE)}!</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `👋 Hola <b>${nombre}</b>\n` +
    `📅 <i>${fecha}</i>\n` +
    `🟢 <i>Sistema en línea</i>\n\n` +
    `<b>Acciones rápidas:</b>\n` +
    `• Consulta de pedidos activos\n` +
    `• Control de cartera por cobrar/pagar\n` +
    `• Generación de PDF ejecutivo\n` +
    `• Envíos automáticos por correo\n` +
    `• Consultas directas: <code>/saldo CLIENTE</code>`;

  enviarMensaje(chatId, texto, null, tecladoPrincipal());
  enviarMenuPrincipal(chatId, usuario);
}
function enviarMenuPedidos(chatId) { enviarMensaje(chatId, renderPedidos(), tecladoPedidos()); }
function enviarMenuCobranzas(chatId) { enviarMensaje(chatId, renderTopDeudores(), tecladoCobranzas()); }
function enviarDashboard(chatId) { enviarMensaje(chatId, renderDashboard(), tecladoDashboard()); }
function enviarReporteTasas(chatId) { enviarMensaje(chatId, renderTasas(), tecladoVolver("MENU")); }
function enviarEstadoServidor(chatId) { enviarMensaje(chatId, renderEstado(), tecladoVolver("MENU")); }
function enviarAyuda(chatId) { enviarMensaje(chatId, renderAyuda(), tecladoVolver("MENU")); }
function enviarMenuReportes(chatId) { enviarMensaje(chatId, renderMenuReportes(), tecladoReportes()); }

// ============================================================================
// SECCIÓN 6: RENDERS DE DATOS
// ============================================================================
function renderPedidos() {
  const datos = leerHoja("Pedidos");
  if (datos.length === 0) return `📦 <b>MÓDULO DE PEDIDOS</b>\n━━━━━━━━━━━━━━━━━━━━\n\nℹ️ No hay pedidos registrados.`;

  const pendientes = datos.filter(p => {
    const st = String(p["StatusPedido"] || "").toLowerCase().trim();
    return st !== "facturado" && st !== "anulado" && st !== "";
  });

  const ultimos = pendientes.slice(-8).reverse();
  let texto = `📦 <b>PEDIDOS PENDIENTES (${pendientes.length})</b>\n━━━━━━━━━━━━━━━━━━━━\n\n`;

  ultimos.forEach(p => {
    const id = escapeHtml(p["idPedido"] || p["IdNroPedido"] || "S/N");
    const cliente = escapeHtml(p["nombreClientePedido"] || "Cliente S/N");
    const montoUSD = parseNumero(p["TotalClienteUSD"]);
    const saldoUSD = parseNumero(p["SaldoClienteUSD"]);
    const status = escapeHtml(p["StatusPedido"] || "Pendiente");
    const fecha = formatFecha(p["FechaRegistroPedido"]);

    texto += `🔹 <b>#${id}</b> — ${cliente}\n`;
    texto += `   💵 <b>$${fmt(montoUSD)}</b>`;
    if (saldoUSD > 0.01 && saldoUSD < montoUSD) texto += ` | ⏳ Saldo: $${fmt(saldoUSD)}`;
    texto += `\n   📌 <code>${status}</code>`;
    if (fecha) texto += ` | 🗓 ${fecha}`;
    texto += `\n────────────────────\n`;
  });

  const totalCartera = pendientes.reduce((s, p) => s + parseNumero(p["TotalClienteUSD"]), 0);
  texto += `\n💼 <b>Total Pedidos Abiertos:</b> <b>$${fmt(totalCartera)} USD</b>`;
  return texto;
}
function renderDetallePedido(idPedido) {
  const datos = leerHoja("Pedidos");
  const p = datos.find(x => String(x["idPedido"] || x["IdNroPedido"]) === String(idPedido));
  if (!p) return `❌ Pedido <b>#${escapeHtml(idPedido)}</b> no encontrado.`;

  return (
    `📦 <b>DETALLE DE PEDIDO #${escapeHtml(idPedido)}</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `👤 <b>Cliente:</b> ${escapeHtml(p["nombreClientePedido"] || "N/A")}\n` +
    `🆔 <b>RIF:</b> <code>${escapeHtml(p["rifClientePedido"] || "N/A")}</code>\n` +
    `🏷️ <b>Tipo:</b> ${escapeHtml(p["TipoClientePedido"] || "N/A")}\n` +
    `📄 <b>Factura:</b> ${escapeHtml(p["NroFacturaClientePedido"] || "N/A")}\n` +
    `💵 <b>Total USD:</b> <b>$${fmt(p["TotalClienteUSD"])}</b>\n` +
    `⏳ <b>Saldo USD:</b> <b>$${fmt(p["SaldoClienteUSD"])}</b>\n` +
    `📌 <b>Status:</b> <code>${escapeHtml(p["StatusPedido"] || "N/A")}</code>\n` +
    `🗓️ <b>Fecha:</b> ${formatFecha(p["FechaRegistroPedido"]) || "N/A"}`
  );
}
function renderTopDeudores() {
  const ventas = leerHoja("Ventas");
  const mapa = {};

  ventas.forEach(v => {
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const estado = String(v["cobradaPagada"] || "").toLowerCase();
    if (saldo > 0.01 && !["si", "sí", "pagada", "cobrada", "cancelada"].includes(estado)) {
      const rif = String(v["RifCliente"] || "S/R").toUpperCase().trim();
      if (!mapa[rif]) mapa[rif] = { nombre: v["NombreCliente"] || "N/A", saldo: 0, facturas: 0 };
      mapa[rif].saldo += saldo;
      mapa[rif].facturas += 1;
    }
  });

  const deudores = Object.entries(mapa).map(([rif, d]) => ({ rif, ...d })).sort((a, b) => b.saldo - a.saldo);
  if (deudores.length === 0) return `💰 <b>CUENTAS POR COBRAR</b>\n━━━━━━━━━━━━━━━━━━━━\n\n🎉 <i>¡No hay cuentas pendientes por cobrar!</i>`;

  const total = deudores.reduce((s, d) => s + d.saldo, 0);
  const medallas = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"];

  let texto = `💰 <b>TOP 5 DEUDORES (CLIENTES)</b>\n━━━━━━━━━━━━━━━━━━━━\n\n`;
  deudores.slice(0, 5).forEach((d, i) => {
    texto += `${medallas[i]} <b>${escapeHtml(d.nombre)}</b>\n`;
    texto += `   🆔 <code>${escapeHtml(d.rif)}</code>\n`;
    texto += `   📄 ${d.facturas} facturas pendientes\n`;
    texto += `   💵 Deuda: <b>$${fmt(d.saldo)} USD</b>\n`;
    texto += `────────────────────\n`;
  });

  texto += `\n📊 <b>Total Cartera por Cobrar:</b> <b>$${fmt(total)} USD</b>\n`;
  texto += `👥 <b>Clientes en mora:</b> ${deudores.length}\n`;
  texto += `<i>Consulta específica:</i> <code>/saldo CLIENTE</code>`;
  return texto;
}
function renderTopProveedores() {
  const compras = leerHoja("Compras");
  const mapa = {};

  compras.forEach(c => {
    const saldo = parseNumero(c["SaldoProveedorDivisa"]);
    const estado = String(c["cobradaPagada"] || "").toLowerCase();
    if (saldo > 0.01 && !["si", "sí", "pagada", "pagado", "cancelada"].includes(estado)) {
      const rif = String(c["RifProveedor"] || "S/R").toUpperCase().trim();
      if (!mapa[rif]) mapa[rif] = { nombre: c["NombreProveedor"] || "N/A", saldo: 0, facturas: 0 };
      mapa[rif].saldo += saldo;
      mapa[rif].facturas += 1;
    }
  });

  const proveedores = Object.entries(mapa).map(([rif, p]) => ({ rif, ...p })).sort((a, b) => b.saldo - a.saldo);
  if (proveedores.length === 0) return `🏢 <b>CUENTAS POR PAGAR</b>\n━━━━━━━━━━━━━━━━━━━━\n\n✅ <i>No hay deudas con proveedores.</i>`;

  const total = proveedores.reduce((s, p) => s + p.saldo, 0);

  let texto = `🏢 <b>TOP 5 CUENTAS POR PAGAR (PROVEEDORES)</b>\n━━━━━━━━━━━━━━━━━━━━\n\n`;
  proveedores.slice(0, 5).forEach((p, i) => {
    texto += `<b>${i + 1}. ${escapeHtml(p.nombre)}</b>\n`;
    texto += `   🆔 <code>${escapeHtml(p.rif)}</code>\n`;
    texto += `   📄 ${p.facturas} facturas pendientes\n`;
    texto += `   💳 Deuda: <b>$${fmt(p.saldo)} USD</b>\n`;
    texto += `────────────────────\n`;
  });

  texto += `\n📊 <b>Total Cuentas por Pagar:</b> <b>$${fmt(total)} USD</b>\n`;
  texto += `<i>Consulta específica:</i> <code>/proveedor NOMBRE</code>`;
  return texto;
}
function calcularBalance() {
  const ventas = leerHoja("Ventas");
  const compras = leerHoja("Compras");

  let porCobrar = 0, cobrado = 0, porPagar = 0, pagado = 0;

  ventas.forEach(v => {
    const total = parseNumero(v["TotalClienteUSD"]);
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const pagada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (pagada || saldo <= 0.01) cobrado += total;
    else porCobrar += saldo;
  });

  compras.forEach(c => {
    const total = parseNumero(c["TotalProveedorDivisa"]);
    const saldo = parseNumero(c["SaldoProveedorDivisa"]);
    const pagada = ["si", "sí", "pagada", "pagado", "cancelada"].includes(String(c["cobradaPagada"] || "").toLowerCase());
    if (pagada || saldo <= 0.01) pagado += total;
    else porPagar += saldo;
  });

  return { porCobrar, cobrado, porPagar, pagado, balance: porCobrar - porPagar };
}
function renderDashboard() {
  const b = calcularBalance();
  const emoji = b.balance >= 0 ? "🟢" : "🔴";
  const estado = b.balance >= 0 ? "Superávit Teórico" : "Déficit";

  return (
    `📊 <b>DASHBOARD FINANCIERO</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `💵 <b>POR COBRAR (Clientes)</b>\n` +
    `   🔴 Pendiente: <b>$${fmt(b.porCobrar)}</b>\n` +
    `   🟢 Cobrado: <b>$${fmt(b.cobrado)}</b>\n\n` +
    `💳 <b>POR PAGAR (Proveedores)</b>\n` +
    `   🔴 Pendiente: <b>$${fmt(b.porPagar)}</b>\n` +
    `   🟢 Pagado: <b>$${fmt(b.pagado)}</b>\n\n` +
    `━━━━━━━━━━━━━━━━━━━━\n` +
    `${emoji} <b>BALANCE: ${estado}</b>\n` +
    `   <b>$${fmt(b.balance)} USD</b>\n\n` +
    `🕒 <i>Corte: ${Utilities.formatDate(new Date(), TIMEZONE, "dd/MM HH:mm")}</i>`
  );
}
function renderFlujoCaja() {
  const ventas = leerHoja("Ventas");
  const compras = leerHoja("Compras");
  const hoy = new Date();
  const mes = hoy.getMonth(), anio = hoy.getFullYear();

  let ingresos = 0, egresos = 0, nIng = 0, nEgr = 0;

  ventas.forEach(v => {
    const f = parseFecha(v["FechaFacturaCliente"]);
    if (!f || f.getMonth() !== mes || f.getFullYear() !== anio) return;
    const abono = parseNumero(v["AbonoCliente"]);
    const cobrada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (abono > 0) { ingresos += abono; nIng++; }
    else if (cobrada) { ingresos += parseNumero(v["TotalClienteUSD"]); nIng++; }
  });

  compras.forEach(c => {
    const f = parseFecha(c["FechaFacturaProveedor"]);
    if (!f || f.getMonth() !== mes || f.getFullYear() !== anio) return;
    const abono = parseNumero(c["Abono"]);
    const pagada = ["si", "sí", "pagada", "pagado", "cancelada"].includes(String(c["cobradaPagada"] || "").toLowerCase());
    if (abono > 0) { egresos += abono; nEgr++; }
    else if (pagada) { egresos += parseNumero(c["TotalProveedorDivisa"]); nEgr++; }
  });

  const neto = ingresos - egresos;
  const mesNombre = Utilities.formatDate(hoy, TIMEZONE, "MMMM yyyy");

  return (
    `💸 <b>FLUJO DE CAJA EFECTIVO — ${mesNombre.toUpperCase()}</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `📥 <b>Ingresos cobrados:</b>\n` +
    `   💵 <b>$${fmt(ingresos)}</b> (${nIng} operaciones)\n\n` +
    `📤 <b>Egresos pagados:</b>\n` +
    `   💳 <b>$${fmt(egresos)}</b> (${nEgr} operaciones)\n\n` +
    `━━━━━━━━━━━━━━━━━━━━\n` +
    `${neto >= 0 ? "🟢" : "🔴"} <b>Flujo Neto Operativo:</b>\n` +
    `   <b>$${fmt(neto)} USD</b>`
  );
}
function renderTendencias() {
  const ventas = leerHoja("Ventas");
  const compras = leerHoja("Compras");

  const hoy = new Date();
  const mesActual = hoy.getMonth(), anioActual = hoy.getFullYear();
  const mesAnt = mesActual === 0 ? 11 : mesActual - 1;
  const anioAnt = mesActual === 0 ? anioActual - 1 : anioActual;

  const calcularMes = (m, a) => {
    let v = 0, c = 0, nv = 0, nc = 0;
    ventas.forEach(x => {
      const f = parseFecha(x["FechaFacturaCliente"]);
      if (f && f.getMonth() === m && f.getFullYear() === a) { v += parseNumero(x["TotalClienteUSD"]); nv++; }
    });
    compras.forEach(x => {
      const f = parseFecha(x["FechaFacturaProveedor"]);
      if (f && f.getMonth() === m && f.getFullYear() === a) { c += parseNumero(x["TotalProveedorDivisa"]); nc++; }
    });
    return { v, c, nv, nc, u: v - c };
  };

  const act = calcularMes(mesActual, anioActual);
  const ant = calcularMes(mesAnt, anioAnt);

  const dV = ant.v > 0 ? ((act.v - ant.v) / ant.v) * 100 : 0;
  const dC = ant.c > 0 ? ((act.c - ant.c) / ant.c) * 100 : 0;
  const dU = ant.u !== 0 ? ((act.u - ant.u) / Math.abs(ant.u)) * 100 : 0;

  const flecha = v => v >= 0 ? "📈" : "📉";
  const nomMesActual = Utilities.formatDate(hoy, TIMEZONE, "MMMM");
  const nomMesAnt = Utilities.formatDate(new Date(anioAnt, mesAnt, 1), TIMEZONE, "MMMM");

  return (
    `📈 <b>TENDENCIAS COMPARATIVAS</b>\n` +
    `<i>${nomMesActual.toUpperCase()} vs ${nomMesAnt.toUpperCase()}</i>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `💵 <b>VENTAS FACTURADAS:</b>\n` +
    `   ${nomMesActual}: <b>$${fmt(act.v)}</b>\n` +
    `   ${nomMesAnt}: $${fmt(ant.v)}\n` +
    `   ${flecha(dV)} Var: <b>${dV.toFixed(1)}%</b>\n\n` +
    `💳 <b>COMPRAS REGISTRADAS:</b>\n` +
    `   ${nomMesActual}: <b>$${fmt(act.c)}</b>\n` +
    `   ${nomMesAnt}: $${fmt(ant.c)}\n` +
    `   ${flecha(-dC)} Var: <b>${dC.toFixed(1)}%</b>\n\n` +
    `━━━━━━━━━━━━━━━━━━━━\n` +
    `${act.u >= 0 ? "🟢" : "🔴"} <b>MARGEN BRUTO OPERATIVO:</b>\n` +
    `   ${nomMesActual}: <b>$${fmt(act.u)}</b>\n` +
    `   ${nomMesAnt}: $${fmt(ant.u)}\n` +
    `   ${flecha(dU)} Var: <b>${dU.toFixed(1)}%</b>`
  );
}
function renderAlertas() {
  const alertas = generarAlertas();
  if (alertas.length === 0) return `🚨 <b>ALERTAS DEL SISTEMA</b>\n━━━━━━━━━━━━━━━━━━━━\n\n✅ <i>No se detectaron anomalías ni deudas críticas.</i>`;

  let texto = `🚨 <b>ALERTAS ACTIVAS (${alertas.length})</b>\n━━━━━━━━━━━━━━━━━━━━\n\n`;
  alertas.forEach((a, i) => {
    texto += `${a.emoji} <b>${a.titulo}</b>\n   ${a.detalle}\n`;
    if (i < alertas.length - 1) texto += `────────────────────\n`;
  });
  return texto;
}
function generarAlertas() {
  const alertas = [];
  const cfg = BOT_CONFIG.alertas;
  const ventas = leerHoja("Ventas");
  const pedidos = leerHoja("Pedidos");
  const hoy = new Date();

  let deudasVencidas = 0, montoVencido = 0;
  ventas.forEach(v => {
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const pagada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      const f = parseFecha(v["FechaFacturaCliente"]);
      if (f) {
        const dias = Math.floor((hoy - f) / 86400000);
        if (dias > cfg.deudaVencidaDias) { deudasVencidas++; montoVencido += saldo; }
      }
    }
  });

  if (deudasVencidas > 0) {
    alertas.push({
      emoji: "⚠️", titulo: "Cartera Vencida",
      detalle: `${deudasVencidas} facturas vencidas (> ${cfg.deudaVencidaDias} días)\n   💵 Total: <b>$${fmt(montoVencido)} USD</b>`
    });
  }

  const deudores = {};
  ventas.forEach(v => {
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const pagada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      const nombre = v["NombreCliente"] || "N/A";
      deudores[nombre] = (deudores[nombre] || 0) + saldo;
    }
  });

  Object.entries(deudores).forEach(([nombre, saldo]) => {
    if (saldo > cfg.montoDeudaCritica) {
      alertas.push({ emoji: "🔴", titulo: "Deuda Crítica", detalle: `<b>${escapeHtml(nombre)}</b> supera el límite\n   💵 Saldo: <b>$${fmt(saldo)} USD</b>` });
    }
  });

  let estancados = 0;
  pedidos.forEach(p => {
    const st = String(p["StatusPedido"] || "").toLowerCase();
    if (st === "facturado" || st === "anulado" || st === "") return;
    const f = parseFecha(p["FechaRegistroPedido"]);
    if (f && Math.floor((hoy - f) / 86400000) > cfg.pedidosEstancadosDias) estancados++;
  });

  if (estancados > 0) {
    alertas.push({ emoji: "⏸️", titulo: "Pedidos Retenidos", detalle: `${estancados} pedidos abiertos sin facturar > ${cfg.pedidosEstancadosDias} días` });
  }

  const b = calcularBalance();
  if (b.balance < 0) {
    alertas.push({ emoji: "🚨", titulo: "Balance Deficitario", detalle: `Déficit proyectado: <b>$${fmt(Math.abs(b.balance))} USD</b>` });
  }

  return alertas;
}
function renderTasas() {
  let texto = `💵 <b>MONITOR OFICIAL DE TASAS</b>\n`;
  texto += `━━━━━━━━━━━━━━━━━━━━\n\n`;

  let tasasCargadas = false;

  // 1. Intentar mostrar el resumen principal (USD y EUR)
  try {
    if (typeof obtenerTasaCambioActual === 'function') {
      const t = obtenerTasaCambioActual();
      if (t.tasaUSD > 0 || t.tasaEUR > 0) {
        texto += `🇺🇸 <b>DÓLAR (USD):</b> <code>${fmt(t.tasaUSD)} Bs.</code>\n`;
        texto += `🇪🇺 <b>EURO (EUR):</b>  <code>${fmt(t.tasaEUR)} Bs.</code>\n\n`;
        tasasCargadas = true;
      }
    }
  } catch (e) {
    log("Error renderTasas (obtenerTasaCambioActual): " + e.toString());
  }

  // 2. Listar el resto de monedas si existen en la hoja Tasas (B2:C)
  try {
    if (typeof getDataMoneda === 'function') {
      const monedas = getDataMoneda();
      const otrasMonedas = monedas.filter(m => 
        !m.nombre.toLowerCase().includes('bolivar') && 
        !m.nombre.toLowerCase().includes('bs') &&
        !m.nombre.toLowerCase().includes('dolar') &&
        !m.nombre.toLowerCase().includes('usd') &&
        !m.nombre.toLowerCase().includes('euro') &&
        !m.nombre.toLowerCase().includes('eur')
      );

      if (otrasMonedas.length > 0) {
        texto += `<b>Otras Divisas Registradas:</b>\n`;
        otrasMonedas.forEach(m => {
          texto += `• <b>${escapeHtml(m.nombre)}:</b> <code>${fmt(m.tasa)} Bs.</code>\n`;
        });
        texto += `\n`;
      }
    }
  } catch (e) {
    log("Error renderTasas (getDataMoneda): " + e.toString());
  }

  if (!tasasCargadas) {
    texto += `⚠️ <i>No se pudieron cargar las tasas en este momento. Verifique la celda C3 de 'Tasas'.</i>\n\n`;
  }

  const fechaActual = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy - hh:mm a");
  texto += `🕒 <i>Actualizado: ${fechaActual}</i>`;

  return texto;
}
function renderEstado() {
  const ahora = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm:ss");
  let ping = "—";
  try {
    const t0 = Date.now();
    UrlFetchApp.fetch(`${TELEGRAM_API}/getMe`, { muteHttpExceptions: true });
    ping = (Date.now() - t0) + " ms";
  } catch (e) {}

  let conteos = "";
  ["Pedidos", "Ventas", "Compras", "Tasas"].forEach(n => {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const h = ss.getSheetByName(n);
      if (h) conteos += `   • ${n}: <code>${Math.max(0, h.getLastRow() - 1)} registros</code>\n`;
    } catch (e) {}
  });

  return (
    `⚡ <b>ESTADO DEL SISTEMA</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `🟢 <b>Servicio:</b> Operativo\n` +
    `🏓 <b>Latencia Telegram:</b> <code>${ping}</code>\n` +
    `🕐 <b>Hora Servidor:</b> <code>${ahora}</code>\n` +
    `🌎 <b>Zona Horaria:</b> <code>${TIMEZONE}</code>\n\n` +
    `📊 <b>Registros en base de datos:</b>\n${conteos}\n` +
    `📧 <b>Admin:</b> <code>${BOT_CONFIG.emailsAdmin.join(", ")}</code>`
  );
}
function renderAyuda() {
  return (
    `❓ <b>CENTRO DE AYUDA Y COMANDOS</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `📋 <b>NAVEGACIÓN BÁSICA:</b>\n` +
    `• <code>/menu</code> — Pantalla principal\n` +
    `• <code>/catalogo</code> — Crear pedido (catálogo)\n` +
    `• <code>/carrito</code> — Ver carrito actual\n` +
    `• <code>/cancelar</code> — Vaciar carrito sin guardar\n` +
    `• <code>/pedidos</code> — Pedidos pendientes\n` +
    `• <code>/cobranzas</code> — Top deudores\n` +
    `• <code>/resumen</code> — Dashboard general\n` +
    `• <code>/tasas</code> — Tasas de cambio\n` +
    `• <code>/estado</code> — Salud del servidor\n\n` +
    `🔍 <b>BÚSQUEDAS Y CLIENTES:</b>\n` +
    `• <code>/saldo NOMBRE_O_RIF</code> (Ej: <code>/saldo CLIENTES2</code>)\n` +
    `• <code>/proveedor NOMBRE_O_RIF</code>\n` +
    `• <code>/buscar_cliente NOMBRE</code> — Buscar y asignar cliente al carrito\n` +
    `• <code>/nuevo_cliente Nombre | RIF | Tel | Correo</code> — Crear cliente rápido\n\n` +
    `⚙️ <b>CONFIGURACIÓN:</b>\n` +
    `• <code>/empresa ID_HOJA</code> — Configurar hoja de la empresa para este chat\n` +
    `• <code>/config</code> — Ver configuración actual\n\n` +
    `📧 <b>EXPORTACIÓN Y REPORTES:</b>\n` +
    `• <code>/email</code> — Envía reporte HTML al correo\n` +
    `• <code>/pdf</code> — Genera y envía PDF al correo\n` +
    `• <code>/backup</code> — Genera respaldo en Google Drive\n\n` +
    `🔔 <b>NOTIFICACIONES:</b>\n` +
    `• <code>/suscribir</code> — Recibir alertas diarias automáticas\n` +
    `• <code>/desuscribir</code> — Cancelar alertas`
  );
}
function renderMenuReportes() {
  return (
    `📧 <b>CENTRO DE REPORTES Y DESCARGAS</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `Elige el formato o acción a ejecutar:\n\n` +
    `📊 <b>Reporte General:</b> Resumen en HTML a tu bandeja\n` +
    `📄 <b>PDF Ejecutivo:</b> Reporte completo consolidado\n` +
    `📎 <b>Excel (.xlsx):</b> Copia idéntica de la hoja activa\n` +
    `💾 <b>Respaldo Drive:</b> Respaldo histórico (XLSX, PDF, JSON)\n\n` +
    `📧 <i>Destino actual:</i> <code>${BOT_CONFIG.emailsAdmin[0]}</code>`
  );
}

// ============================================================================
// SECCIÓN 7: CONSULTAS INDIVIDUALES
// ============================================================================
function consultarSaldoCliente(busqueda) {
  const ventas = leerHoja("Ventas");
  const busq = String(busqueda).toLowerCase().trim();

  const coincidencias = ventas.filter(v => {
    const rif = String(v["RifCliente"] || "").toLowerCase();
    const cli = String(v["NombreCliente"] || "").toLowerCase();
    return rif.includes(busq) || cli.includes(busq);
  });

  if (coincidencias.length === 0) return `❌ <b>Sin resultados:</b>\nNo se encontró ningún cliente con "<b>${escapeHtml(busqueda)}</b>"`;

  const nombre = escapeHtml(coincidencias[0]["NombreCliente"]);
  const rif = escapeHtml(coincidencias[0]["RifCliente"]);
  let deudaTotal = 0;
  const facturas = [];

  coincidencias.forEach(v => {
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const pagada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      deudaTotal += saldo;
      facturas.push({
        nro: escapeHtml(v["NroFacturaCliente"] || v["Id"] || "S/N"),
        saldo: saldo,
        fecha: formatFecha(v["FechaFacturaCliente"])
      });
    }
  });

  let txt = `👤 <b>${nombre}</b>\n━━━━━━━━━━━━━━━━━━━━\n`;
  txt += `🆔 <b>RIF:</b> <code>${rif}</code>\n`;
  txt += `📄 <b>Facturas en base de datos:</b> ${coincidencias.length}\n`;
  txt += `💵 <b>Deuda pendiente total:</b> <b>$${fmt(deudaTotal)} USD</b>\n\n`;

  if (facturas.length > 0) {
    txt += `📋 <b>Facturas impagas (${facturas.length}):</b>\n`;
    facturas.slice(0, 10).forEach(f => {
      txt += `• Doc #${f.nro}: <b>$${fmt(f.saldo)}</b> ${f.fecha ? `(${f.fecha})` : ""}\n`;
    });
    if (facturas.length > 10) txt += `<i>... y ${facturas.length - 10} facturas más.</i>\n`;
  } else {
    txt += `🎉 <i>El cliente se encuentra solvente.</i>`;
  }
  return txt;
}
function consultarSaldoProveedor(busqueda) {
  const compras = leerHoja("Compras");
  const busq = String(busqueda).toLowerCase().trim();

  const coincidencias = compras.filter(c => {
    const rif = String(c["RifProveedor"] || "").toLowerCase();
    const nom = String(c["NombreProveedor"] || "").toLowerCase();
    return rif.includes(busq) || nom.includes(busq);
  });

  if (coincidencias.length === 0) return `❌ <b>Sin resultados:</b>\nProveedor "<b>${escapeHtml(busqueda)}</b>" no encontrado.`;

  const nombre = escapeHtml(coincidencias[0]["NombreProveedor"]);
  const rif = escapeHtml(coincidencias[0]["RifProveedor"]);
  let deudaTotal = 0;
  const facturas = [];

  coincidencias.forEach(c => {
    const saldo = parseNumero(c["SaldoProveedorDivisa"]);
    const pagada = ["si", "sí", "pagada", "pagado", "cancelada"].includes(String(c["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      deudaTotal += saldo;
      facturas.push({
        nro: escapeHtml(c["NroFacturaProveedor"] || c["Id"] || "S/N"),
        saldo: saldo,
        fecha: formatFecha(c["FechaFacturaProveedor"])
      });
    }
  });

  let txt = `🏢 <b>${nombre}</b>\n━━━━━━━━━━━━━━━━━━━━\n`;
  txt += `🆔 <b>RIF:</b> <code>${rif}</code>\n`;
  txt += `📄 <b>Documentos registrados:</b> ${coincidencias.length}\n`;
  txt += `💳 <b>Saldo por pagar:</b> <b>$${fmt(deudaTotal)} USD</b>\n\n`;

  if (facturas.length > 0) {
    txt += `📋 <b>Facturas pendientes (${facturas.length}):</b>\n`;
    facturas.slice(0, 10).forEach(f => {
      txt += `• Doc #${f.nro}: <b>$${fmt(f.saldo)}</b> ${f.fecha ? `(${f.fecha})` : ""}\n`;
    });
    if (facturas.length > 10) txt += `<i>... y ${facturas.length - 10} facturas más.</i>\n`;
  } else {
    txt += `✅ <i>No se le adeuda a este proveedor.</i>`;
  }
  return txt;
}

// ============================================================================
// SECCIÓN 8: SERVICIOS DE CORREO HTML
// ============================================================================
function enviarReportePorEmail(chatId) {
  enviarMensaje(chatId, "📧 <i>Enviando reporte completo por correo...</i>", null);
  const resultado = enviarReporteCompletoPorEmail("manual");
  enviarMensaje(chatId, resultado, tecladoVolver("MENU"));
}
function enviarReporteSegmentado(chatId, tipo, messageId) {
  try {
    let asunto = "", html = "";
    const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm");

    if (tipo === "pedidos") { asunto = `📦 Pedidos Pendientes - ${fecha}`; html = construirHtmlPedidos(); }
    else if (tipo === "cobranzas") { asunto = `💰 Cartera de Cobranzas - ${fecha}`; html = construirHtmlCobranzas(); }
    else if (tipo === "dashboard") { asunto = `📊 Dashboard Financiero - ${fecha}`; html = construirHtmlDashboard(); }

    MailApp.sendEmail({
      to: BOT_CONFIG.emailsAdmin.join(","),
      subject: `${asunto} | ${BOT_CONFIG.empresa.nombre}`,
      htmlBody: html,
      name: BOT_CONFIG.empresa.nombre + " Bot"
    });

    const msg = `✅ Reporte <b>${tipo}</b> enviado satisfactoriamente a:\n<code>${BOT_CONFIG.emailsAdmin.join("\n")}</code>`;
    if (messageId) editarMensaje(chatId, messageId, msg, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, msg, tecladoVolver("REPORTES"));
  } catch (e) {
    const err = `❌ Error al enviar reporte: <code>${escapeHtml(e.toString())}</code>`;
    if (messageId) editarMensaje(chatId, messageId, err, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, err, tecladoVolver("REPORTES"));
  }
}
function enviarReporteCompletoPorEmail(tipo) {
  try {
    const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm");
    const html = construirHtmlCompleto();
    const asunto = `📊 Reporte Ejecutivo ${BOT_CONFIG.empresa.nombre} - ${fecha}`;

    MailApp.sendEmail({
      to: BOT_CONFIG.emailsAdmin.join(","),
      subject: asunto,
      htmlBody: html,
      name: BOT_CONFIG.empresa.nombre + " Bot"
    });

    return `✅ <b>Reporte enviado correctamente</b>\n\n📧 Destinatarios:\n<code>${BOT_CONFIG.emailsAdmin.join("\n")}</code>\n\n🕐 ${fecha}`;
  } catch (e) {
    return `❌ <b>Error al enviar correo:</b>\n<code>${escapeHtml(e.toString())}</code>`;
  }
}
function wrapperHtml(titulo, contenido) {
  const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm");
  const c = BOT_CONFIG.empresa.color;
  const e = BOT_CONFIG.empresa;

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:20px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.06);">
        <tr><td style="background:${c};color:#fff;padding:20px 28px;">
          <h1 style="margin:0;font-size:20px;">📊 ${escapeHtml(e.nombre)}</h1>
          <p style="margin:4px 0 0 0;opacity:0.9;font-size:12px;">${escapeHtml(e.slogan)}</p>
        </td></tr>
        <tr><td style="padding:24px 28px;">
          <h2 style="margin:0 0 16px 0;color:#1a1a1a;font-size:16px;border-bottom:2px solid ${c};padding-bottom:8px;">${titulo}</h2>
          ${contenido}
        </td></tr>
        <tr><td style="background:#f8f9fa;padding:16px 28px;color:#666;font-size:11px;border-top:1px solid #eee;">
          <p style="margin:0;">Generado automáticamente: <b>${fecha}</b></p>
          <p style="margin:4px 0 0 0;">${escapeHtml(e.nombre)} · RIF: ${escapeHtml(e.rif)} · ${escapeHtml(e.telefono)}</p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}
function tarjetaMetrica(titulo, valor, color, emoji) {
  return `<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:10px;">
    <tr><td style="background:#f8f9fa;border-left:4px solid ${color};padding:12px 16px;border-radius:4px;">
      <div style="font-size:11px;color:#666;text-transform:uppercase;letter-spacing:0.5px;">${emoji} ${titulo}</div>
      <div style="font-size:20px;font-weight:bold;color:#1a1a1a;margin-top:2px;">${valor}</div>
    </td></tr></table>`;
}
function tablaHtml(headers, filas) {
  let html = `<table width="100%" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:12px;margin:14px 0;">
    <thead><tr style="background:${BOT_CONFIG.empresa.color};color:#fff;">
      ${headers.map(h => `<th style="text-align:left;padding:8px;font-weight:600;">${h}</th>`).join("")}
    </tr></thead><tbody>`;
  filas.forEach((fila, i) => {
    const bg = i % 2 === 0 ? "#fff" : "#f8f9fa";
    html += `<tr style="background:${bg};border-bottom:1px solid #e0e0e0;">
      ${fila.map(c => `<td style="padding:8px;color:#333;">${c}</td>`).join("")}
    </tr>`;
  });
  html += `</tbody></table>`;
  return html;
}
function construirHtmlCompleto() {
  const b = calcularBalance();
  const alertas = generarAlertas();
  let html = "";

  html += `<h3 style="color:#1a1a1a;margin:16px 0 10px 0;">💼 Balance Resumido</h3>`;
  html += tarjetaMetrica("Por Cobrar", `$${fmt(b.porCobrar)} USD`, "#f59e0b", "🔴");
  html += tarjetaMetrica("Por Pagar", `$${fmt(b.porPagar)} USD`, "#ef4444", "💳");
  html += tarjetaMetrica("Balance Neto", `$${fmt(b.balance)} USD`, b.balance >= 0 ? "#10b981" : "#ef4444", b.balance >= 0 ? "🟢" : "🔴");

  const ventas = leerHoja("Ventas");
  const mapa = {};
  ventas.forEach(v => {
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const pagada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      const n = v["NombreCliente"] || "N/A";
      mapa[n] = (mapa[n] || 0) + saldo;
    }
  });
  const top = Object.entries(mapa).map(([n, s]) => ({ n, s })).sort((a, b) => b.s - a.s).slice(0, 5);

  if (top.length > 0) {
    html += `<h3 style="color:#1a1a1a;margin:20px 0 10px 0;">💰 Top 5 Clientes en Mora</h3>`;
    html += tablaHtml(["Cliente", "Deuda"], top.map(d => [escapeHtml(d.n), `<b>$${fmt(d.s)} USD</b>`]));
  }

  if (alertas.length > 0) {
    html += `<h3 style="color:#1a1a1a;margin:20px 0 10px 0;">🚨 Alertas Detectadas (${alertas.length})</h3>`;
    alertas.forEach(a => {
      html += `<div style="background:#fef3c7;border-left:4px solid #f59e0b;padding:10px 14px;margin-bottom:8px;border-radius:4px;font-size:12px;">
        <b>${a.emoji} ${escapeHtml(a.titulo)}</b><br>
        <span style="color:#555;">${a.detalle.replace(/<[^>]*>/g, "")}</span>
      </div>`;
    });
  }

  return wrapperHtml("Reporte General del Sistema", html);
}
function construirHtmlPedidos() {
  const pedidos = leerHoja("Pedidos").filter(p => {
    const st = String(p["StatusPedido"] || "").toLowerCase();
    return st !== "facturado" && st !== "anulado" && st !== "";
  });

  let html = tarjetaMetrica("Pedidos Pendientes", pedidos.length, "#f59e0b", "📦");

  if (pedidos.length > 0) {
    const total = pedidos.reduce((s, p) => s + parseNumero(p["TotalClienteUSD"]), 0);
    html += tarjetaMetrica("Monto Total Pendiente", `$${fmt(total)} USD`, "#10b981", "💵");
    html += tablaHtml(
      ["#", "Cliente", "RIF", "Total USD", "Status", "Fecha"],
      pedidos.map(p => [
        escapeHtml(p["idPedido"] || p["IdNroPedido"] || "S/N"),
        escapeHtml(p["nombreClientePedido"] || "N/A"),
        `<code>${escapeHtml(p["rifClientePedido"] || "N/A")}</code>`,
        `<b>$${fmt(p["TotalClienteUSD"])}</b>`,
        escapeHtml(p["StatusPedido"] || "Pendiente"),
        formatFecha(p["FechaRegistroPedido"]) || "N/A"
      ])
    );
  }

  return wrapperHtml("Pedidos Abiertos en Sistema", html);
}
function construirHtmlCobranzas() {
  const ventas = leerHoja("Ventas");
  const mapa = {};

  ventas.forEach(v => {
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const pagada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      const rif = String(v["RifCliente"] || "S/R").toUpperCase();
      if (!mapa[rif]) mapa[rif] = { nombre: v["NombreCliente"] || "N/A", saldo: 0, facturas: 0 };
      mapa[rif].saldo += saldo;
      mapa[rif].facturas++;
    }
  });

  const deudores = Object.entries(mapa).map(([rif, d]) => ({ rif, ...d })).sort((a, b) => b.saldo - a.saldo);
  const total = deudores.reduce((s, d) => s + d.saldo, 0);

  let html = tarjetaMetrica("Clientes en Cartera", deudores.length, "#f59e0b", "👥");
  html += tarjetaMetrica("Cartera Total por Cobrar", `$${fmt(total)} USD`, "#ef4444", "💰");

  if (deudores.length > 0) {
    html += tablaHtml(
      ["Cliente", "RIF", "Facturas", "Deuda Total USD"],
      deudores.slice(0, 20).map(d => [escapeHtml(d.nombre), `<code>${escapeHtml(d.rif)}</code>`, d.facturas, `<b>$${fmt(d.saldo)}</b>`])
    );
  }

  return wrapperHtml("Reporte Detallado de Cobranzas", html);
}
function construirHtmlDashboard() {
  const b = calcularBalance();
  const flujo = renderFlujoCaja();

  let html = tarjetaMetrica("Por Cobrar", `$${fmt(b.porCobrar)} USD`, "#f59e0b", "🔴");
  html += tarjetaMetrica("Por Pagar", `$${fmt(b.porPagar)} USD`, "#ef4444", "💳");
  html += tarjetaMetrica("Cobrado Efectivo", `$${fmt(b.cobrado)} USD`, "#10b981", "🟢");
  html += tarjetaMetrica("Pagado Efectivo", `$${fmt(b.pagado)} USD`, "#3b82f6", "🔵");
  html += tarjetaMetrica("Balance Proyectado", `$${fmt(b.balance)} USD`, b.balance >= 0 ? "#10b981" : "#ef4444", b.balance >= 0 ? "📈" : "📉");

  html += `<pre style="background:#f8f9fa;padding:14px;border-radius:4px;font-size:11px;color:#333;white-space:pre-wrap;border:1px solid #e5e7eb;">${flujo.replace(/<[^>]*>/g, "")}</pre>`;

  return wrapperHtml("Dashboard Ejecutivo", html);
}

// ============================================================================
// SECCIÓN 9: EXPORTACIÓN A EXCEL 📎
// ============================================================================
function enviarExcelPorEmail(chatId, messageId) {
  if (messageId) editarMensaje(chatId, messageId, "📎 <i>Generando archivo Excel desde Google Sheets...</i>", null);
  else enviarMensaje(chatId, "📎 <i>Generando archivo Excel...</i>", null);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fecha = Utilities.formatDate(new Date(), TIMEZONE, "yyyyMMdd_HHmm");
    const nombreArchivo = `GproPos_Backup_${fecha}.xlsx`;

    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    }).getBlob().setName(nombreArchivo);

    MailApp.sendEmail({
      to: BOT_CONFIG.emailsAdmin.join(","),
      subject: `📎 Respaldo Excel - ${BOT_CONFIG.empresa.nombre} - ${fecha}`,
      body: `Adjunto encontrarás el libro completo en formato Excel.\n\nGenerado: ${Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm")}`,
      attachments: [blob],
      name: BOT_CONFIG.empresa.nombre + " Bot"
    });

    const exitoMsg = `✅ <b>Archivo Excel enviado por correo</b>\n\n📧 <code>${BOT_CONFIG.emailsAdmin.join(", ")}</code>\n📎 ${nombreArchivo}`;
    if (messageId) editarMensaje(chatId, messageId, exitoMsg, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, exitoMsg, tecladoVolver("REPORTES"));
  } catch (e) {
    const errMsg = `❌ Error generando Excel: <code>${escapeHtml(e.toString())}</code>`;
    if (messageId) editarMensaje(chatId, messageId, errMsg, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, errMsg, tecladoVolver("REPORTES"));
  }
}

// ============================================================================
// SECCIÓN 10: SUSCRIPCIONES Y CONFIGURACIÓN
// ============================================================================
function obtenerSuscripciones() {
  const props = PropertiesService.getScriptProperties();
  const data = props.getProperty("suscripciones");
  return data ? JSON.parse(data) : {};
}
function guardarSuscripciones(data) {
  PropertiesService.getScriptProperties().setProperty("suscripciones", JSON.stringify(data));
}
function suscribirUsuario(chatId, from) {
  const subs = obtenerSuscripciones();
  subs[chatId] = {
    nombre: (from && from.first_name) || "Usuario",
    username: (from && from.username) || "",
    fecha: new Date().toISOString()
  };
  guardarSuscripciones(subs);

  enviarMensaje(chatId,
    `✅ <b>Suscripción activada con éxito</b>\n\n` +
    `Recibirás automáticamente:\n` +
    `• 📊 Resumen operativo a las ${BOT_CONFIG.automatizacion.horaReporteDiario}:00 AM\n` +
    `• 🚨 Alertas críticas automáticas\n\n` +
    `<i>Puedes desactivarlo cuando desees con /desuscribir</i>`,
    tecladoVolver("MENU")
  );
}
function desuscribirUsuario(chatId) {
  const subs = obtenerSuscripciones();
  delete subs[chatId];
  guardarSuscripciones(subs);
  enviarMensaje(chatId, `🔕 <b>Suscripción cancelada</b>. Ya no recibirás reportes automáticos.`, tecladoVolver("MENU"));
}
function mostrarConfiguracion(chatId) {
  const subs = obtenerSuscripciones();
  const props = PropertiesService.getScriptProperties();
  const empresaChat = props.getProperty("empresa_chat_" + chatId) || "(no configurada — usando hoja por defecto)";

  const texto =
    `⚙️ <b>CONFIGURACIÓN DEL SISTEMA</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `📄 <b>Hoja de esta empresa:</b>\n<code>${escapeHtml(empresaChat)}</code>\n\n` +
    `📧 <b>Emails autorizados:</b>\n<code>${BOT_CONFIG.emailsAdmin.join("\n")}</code>\n\n` +
    `⏰ <b>Horarios automáticos:</b>\n` +
    `   • Reporte Diario: ${BOT_CONFIG.automatizacion.horaReporteDiario}:00 AM\n` +
    `   • Backup Nocturno: ${BOT_CONFIG.automatizacion.horaBackup}:00 AM\n\n` +
    `🚨 <b>Parámetros de Alerta:</b>\n` +
    `   • Vencimiento: > ${BOT_CONFIG.alertas.deudaVencidaDias} días\n` +
    `   • Deuda crítica: > $${BOT_CONFIG.alertas.montoDeudaCritica} USD\n` +
    `   • Pedidos estancados: > ${BOT_CONFIG.alertas.pedidosEstancadosDias} días\n\n` +
    `👥 <b>Dispositivos Telegram suscritos:</b> ${Object.keys(subs).length}`;

  enviarMensaje(chatId, texto, tecladoVolver("MENU"));
}

// ============================================================================
// SECCIÓN 11: SERVICIOS DE DATOS Y CACHÉ
// ============================================================================
function leerHoja(nombreHojaBuscada) {
  const cache = CacheService.getScriptCache();
  const claveCache = "hoja_" + nombreHojaBuscada.toLowerCase();
  
  try {
    const cached = cache.get(claveCache);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheets().find(h => 
      h.getName().trim().toLowerCase() === nombreHojaBuscada.trim().toLowerCase()
    );
    if (!hoja) return [];

    const lastRow = hoja.getLastRow();
    const lastCol = hoja.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return [];

    const valores = hoja.getRange(1, 1, lastRow, lastCol).getValues();
    const encabezados = valores[0].map(h => String(h).replace(/^\uFEFF/, "").trim());

    const resultado = valores.slice(1)
      .filter(f => f.some(v => v !== "" && v !== null && v !== undefined))
      .map(f => {
        const obj = {};
        encabezados.forEach((h, i) => { if (h) obj[h] = f[i]; });
        return obj;
      });

    try {
      const json = JSON.stringify(resultado);
      if (json.length < 100000) {
        cache.put(claveCache, json, 45);
      }
    } catch (e) {}

    return resultado;
  } catch (err) {
    log("ERROR leerHoja (" + nombreHojaBuscada + "): " + err.toString());
    return [];
  }
}

// ============================================================================
// SECCIÓN 12: UTILIDADES
// ============================================================================
function parseNumero(v) {
  if (typeof v === "number") return v;
  if (!v && v !== 0) return 0;
  let s = String(v).replace(/\$/g, "").replace(/\s/g, "").trim();
  if (!s) return 0;
  if (s.includes(",") && !s.includes(".")) s = s.replace(",", ".");
  else if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function fmt(v) {
  return parseNumero(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function parseFecha(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim();
  if (!s) return null;
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m1) {
    let [, d, m, y] = m1;
    y = y.length === 2 ? "20" + y : y;
    return new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
  }
  const m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m2) return new Date(parseInt(m2[1]), parseInt(m2[2]) - 1, parseInt(m2[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function formatFecha(v) {
  const d = parseFecha(v);
  return d ? Utilities.formatDate(d, TIMEZONE, "dd/MM/yyyy") : "";
}
function escapeHtml(texto) {
  if (!texto) return "";
  return String(texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
function log(msg) { if (DEBUG) Logger.log(msg); }

// ============================================================================
// SECCIÓN 13: CLIENTE HTTP TELEGRAM
// ============================================================================
function enviarMensaje(chatId, texto, tecladoInline, tecladoReply) {
  const payload = {
    chat_id: chatId,
    text: texto,
    parse_mode: "HTML",
    disable_web_page_preview: true
  };
  if (tecladoInline) payload.reply_markup = tecladoInline;
  else if (tecladoReply) payload.reply_markup = tecladoReply;

  return telegramRequest("sendMessage", payload);
}
function editarMensaje(chatId, messageId, texto, teclado) {
  if (!messageId) return enviarMensaje(chatId, texto, teclado);

  const payload = {
    chat_id: chatId,
    message_id: messageId,
    text: texto,
    parse_mode: "HTML",
    disable_web_page_preview: true
  };
  if (teclado) payload.reply_markup = teclado;

  const res = telegramRequest("editMessageText", payload);
  if (!res || !res.ok) {
    delete payload.message_id;
    return telegramRequest("sendMessage", payload);
  }
  return res;
}
function responderCallback(id) {
  telegramRequest("answerCallbackQuery", { callback_query_id: id });
}
function telegramRequest(metodo, payload) {
  const opts = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(`${TELEGRAM_API}/${metodo}`, opts);
    const json = JSON.parse(res.getContentText());
    if (!json.ok && payload.parse_mode === "HTML" && payload.text) {
      payload.text = String(payload.text).replace(/<[^>]*>/g, "");
      delete payload.parse_mode;
      opts.payload = JSON.stringify(payload);
      const res2 = UrlFetchApp.fetch(`${TELEGRAM_API}/${metodo}`, opts);
      return JSON.parse(res2.getContentText());
    }
    return json;
  } catch (e) {
    log("ERROR en telegramRequest: " + e.toString());
    return { ok: false };
  }
}

// ============================================================================
// SECCIÓN 14: AUTOMATIZACIONES (TRIGGERS)
// ============================================================================
function instalarTriggerDiario() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (["reporteDiarioAutomatico", "reporteSemanalAutomatico", "respaldoAutomatico"].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("reporteDiarioAutomatico")
    .timeBased().atHour(BOT_CONFIG.automatizacion.horaReporteDiario).everyDays(1)
    .inTimezone(TIMEZONE).create();

  ScriptApp.newTrigger("reporteSemanalAutomatico")
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(BOT_CONFIG.automatizacion.horaReporteDiario)
    .inTimezone(TIMEZONE).create();

  if (BOT_CONFIG.respaldoAutomatico) {
    ScriptApp.newTrigger("respaldoAutomatico")
      .timeBased().atHour(BOT_CONFIG.automatizacion.horaBackup).everyDays(1)
      .inTimezone(TIMEZONE).create();
  }

  Logger.log("✅ Triggers diarios instalados");
}
function instalarTriggerAlertas() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "verificarAlertasCriticas") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("verificarAlertasCriticas")
    .timeBased().everyHours(BOT_CONFIG.automatizacion.horasVerificarAlertas).create();

  Logger.log("✅ Trigger de alertas críticas instalado");
}
function reporteDiarioAutomatico() {
  log("⏰ Iniciando reporte diario automático...");
  try {
    const adjuntos = [];
    if (BOT_CONFIG.enviarPdfDiario) {
      try {
        const pdfBlob = generarPdfBlob();
        adjuntos.push(pdfBlob);
      } catch (e) { log("⚠️ Error generando PDF diario: " + e.toString()); }
    }

    if (BOT_CONFIG.enviarEmailDiario) {
      const html = construirHtmlCompleto();
      const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy");
      MailApp.sendEmail({
        to: BOT_CONFIG.emailsAdmin.join(","),
        subject: `📊 Reporte Diario ${BOT_CONFIG.empresa.nombre} - ${fecha}`,
        htmlBody: html,
        attachments: adjuntos,
        name: BOT_CONFIG.empresa.nombre + " Bot"
      });
      log("✅ Correo diario enviado a admins");
    }

    if (BOT_CONFIG.enviarAlertaTelegram) {
      const subs = obtenerSuscripciones();
      const alertas = generarAlertas();
      const b = calcularBalance();

      let mensaje =
        `🌅 <b>RESUMEN MATUTINO — ${escapeHtml(BOT_CONFIG.empresa.nombre)}</b>\n━━━━━━━━━━━━━━━━━━━━\n\n` +
        `📊 <b>Posición Actual:</b>\n` +
        `   💰 Por cobrar: <b>$${fmt(b.porCobrar)}</b>\n` +
        `   💳 Por pagar: <b>$${fmt(b.porPagar)}</b>\n` +
        `   ${b.balance >= 0 ? "🟢" : "🔴"} Balance neto: <b>$${fmt(b.balance)} USD</b>\n\n`;

      if (alertas.length > 0) {
        mensaje += `🚨 <b>Alertas Críticas (${alertas.length}):</b>\n`;
        alertas.slice(0, 5).forEach(a => {
          mensaje += `${a.emoji} <b>${a.titulo}</b>\n   ${a.detalle}\n`;
        });
      } else {
        mensaje += `✅ <i>Sin alertas registradas</i>\n`;
      }

      const teclado = {
        inline_keyboard: [
          [
            { text: "📊 Dashboard", callback_data: "DASHBOARD" },
            { text: "🚨 Alertas", callback_data: "ALERTAS" }
          ],
          [{ text: "🏠 Menú Principal", callback_data: "MENU" }]
        ]
      };

      Object.keys(subs).forEach(chatId => {
        try {
          enviarMensaje(chatId, mensaje, teclado);
          Utilities.sleep(150);
        } catch (e) {}
      });
    }
  } catch (e) {
    log("❌ Error fatal en reporte diario: " + e.toString());
  }
}
function reporteSemanalAutomatico() {
  try {
    const html = construirHtmlCompleto();
    const tendencias = renderTendencias();
    const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy");

    const htmlCompleto = wrapperHtml("Reporte Ejecutivo Semanal", `
      ${html}
      <h3 style="color:#1a1a1a;margin:20px 0 10px 0;">📈 Tendencias del Periodo</h3>
      <pre style="background:#f8f9fa;padding:14px;border-radius:4px;font-size:11px;border:1px solid #e5e7eb;">${tendencias.replace(/<[^>]*>/g, "")}</pre>
    `);

    MailApp.sendEmail({
      to: BOT_CONFIG.emailsAdmin.join(","),
      subject: `📈 Resumen Semanal ${BOT_CONFIG.empresa.nombre} - ${fecha}`,
      htmlBody: htmlCompleto,
      name: BOT_CONFIG.empresa.nombre + " Bot"
    });
  } catch (e) {
    log("❌ Error en reporte semanal: " + e.toString());
  }
}
function verificarAlertasCriticas() {
  try {
    const alertas = generarAlertas();
    const criticas = alertas.filter(a => a.emoji === "🚨" || a.emoji === "🔴");
    if (criticas.length === 0) return;

    const subs = obtenerSuscripciones();
    const mensaje =
      `🚨 <b>NOTIFICACIÓN DE ALERTA CRÍTICA</b>\n━━━━━━━━━━━━━━━━━━━━\n\n` +
      criticas.map(a => `${a.emoji} <b>${a.titulo}</b>\n${a.detalle}`).join("\n\n") +
      `\n\n<i>Accede al bot para revisar el estado.</i>`;

    Object.keys(subs).forEach(chatId => {
      try { 
        enviarMensaje(chatId, mensaje, tecladoVolver("MENU")); 
        Utilities.sleep(150); 
      } catch (e) {}
    });
  } catch (e) { log("Error al verificar alertas: " + e.toString()); }
}
function respaldoAutomatico() {
  log("⏰ Ejecutando respaldo automático...");
  realizarRespaldo("auto");
}

// ============================================================================
// SECCIÓN 15: PDF PROFESIONAL (CONVERSIÓN NATIVA EN MEMORIA)
// ============================================================================
function generarPdfDesdeHtml(html, nombreArchivo) {
  return Utilities.newBlob(html, "text/html", nombreArchivo + ".html")
    .getAs("application/pdf")
    .setName(nombreArchivo + ".pdf");
}
function construirHtmlPdfCompleto() {
  const b = calcularBalance();
  const alertas = generarAlertas();
  const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm");
  const c = BOT_CONFIG.empresa;

  const ventas = leerHoja("Ventas");
  const compras = leerHoja("Compras");
  const pedidos = leerHoja("Pedidos");

  const mapaDeudores = {};
  ventas.forEach(v => {
    const saldo = parseNumero(v["SaldoClienteUSD"]);
    const pagada = ["si", "sí", "pagada", "cobrada", "cancelada"].includes(String(v["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      const rif = String(v["RifCliente"] || "S/R").toUpperCase();
      if (!mapaDeudores[rif]) mapaDeudores[rif] = { nombre: v["NombreCliente"] || "N/A", saldo: 0, facturas: 0 };
      mapaDeudores[rif].saldo += saldo;
      mapaDeudores[rif].facturas++;
    }
  });
  const topDeudores = Object.entries(mapaDeudores).map(([rif, d]) => ({ rif, ...d })).sort((a, b) => b.saldo - a.saldo).slice(0, 10);

  const mapaProv = {};
  compras.forEach(x => {
    const saldo = parseNumero(x["SaldoProveedorDivisa"]);
    const pagada = ["si", "sí", "pagada", "pagado", "cancelada"].includes(String(x["cobradaPagada"] || "").toLowerCase());
    if (saldo > 0.01 && !pagada) {
      const rif = String(x["RifProveedor"] || "S/R").toUpperCase();
      if (!mapaProv[rif]) mapaProv[rif] = { nombre: x["NombreProveedor"] || "N/A", saldo: 0, facturas: 0 };
      mapaProv[rif].saldo += saldo;
      mapaProv[rif].facturas++;
    }
  });
  const topProv = Object.entries(mapaProv).map(([rif, p]) => ({ rif, ...p })).sort((a, b) => b.saldo - a.saldo).slice(0, 10);

  const pedidosPend = pedidos.filter(p => {
    const st = String(p["StatusPedido"] || "").toLowerCase();
    return st !== "facturado" && st !== "anulado" && st !== "";
  });

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  @page { size: A4; margin: 12mm; }
  body { font-family: Helvetica, Arial, sans-serif; color: #1a1a1a; margin: 0; padding: 0; font-size: 10px; line-height: 1.3; }
  .header { background: ${c.color}; color: white; padding: 16px 20px; border-radius: 6px; margin-bottom: 16px; }
  .header h1 { margin: 0 0 4px 0; font-size: 18px; }
  .header p { margin: 2px 0; font-size: 10px; opacity: 0.9; }
  .header .meta { margin-top: 8px; padding-top: 8px; border-top: 1px solid rgba(255,255,255,0.25); font-size: 9px; }
  h2 { color: ${c.color}; font-size: 13px; margin: 16px 0 8px 0; border-bottom: 2px solid ${c.color}; padding-bottom: 4px; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 9px; }
  th { background: ${c.color}; color: white; padding: 6px 8px; text-align: left; }
  td { padding: 5px 8px; border-bottom: 1px solid #eee; }
  tr:nth-child(even) td { background: #fbfbfb; }
  .alert { background: #fef3c7; border-left: 3px solid #f59e0b; padding: 6px 10px; margin-bottom: 6px; font-size: 9px; }
  .footer { margin-top: 24px; padding-top: 10px; border-top: 1px solid #eee; font-size: 8px; color: #888; text-align: center; }
</style></head>
<body>

<div class="header">
  <h1>${escapeHtml(c.nombre)} — Reporte Ejecutivo</h1>
  <p>${escapeHtml(c.slogan)}</p>
  <div class="meta">
    RIF: ${escapeHtml(c.rif)} | Emitido: ${fecha} | Contacto: ${escapeHtml(c.telefono)}
  </div>
</div>

<h2>💼 RESUMEN FINANCIERO CONSOLIDADO</h2>
<table style="margin-bottom: 12px;">
  <tr style="background:#f4f6f8; font-weight:bold;">
    <td style="padding:10px; border-left: 4px solid #f59e0b;">Por Cobrar: $${fmt(b.porCobrar)} USD</td>
    <td style="padding:10px; border-left: 4px solid #ef4444;">Por Pagar: $${fmt(b.porPagar)} USD</td>
    <td style="padding:10px; border-left: 4px solid ${b.balance >= 0 ? '#10b981' : '#ef4444'};">
      Balance Neto: $${fmt(b.balance)} USD
    </td>
  </tr>
</table>

${alertas.length > 0 ? `
<h2>🚨 ALERTAS DEL SISTEMA</h2>
${alertas.map(a => `
  <div class="alert">
    <b>${a.emoji} ${escapeHtml(a.titulo)}:</b> ${a.detalle.replace(/<[^>]*>/g, "")}
  </div>
`).join("")}
` : ''}

${topDeudores.length > 0 ? `
<h2>💰 TOP 10 CUENTAS POR COBRAR (CLIENTES)</h2>
<table>
  <thead>
    <tr>
      <th style="width:25px;">#</th><th>Nombre Fiscal</th><th>RIF</th><th style="text-align:center;">Facturas</th><th style="text-align:right;">Saldo USD</th>
    </tr>
  </thead>
  <tbody>
    ${topDeudores.map((d, i) => `
      <tr>
        <td>${i + 1}</td>
        <td><b>${escapeHtml(d.nombre)}</b></td>
        <td>${escapeHtml(d.rif)}</td>
        <td style="text-align:center;">${d.facturas}</td>
        <td style="text-align:right;"><b>$${fmt(d.saldo)}</b></td>
      </tr>
    `).join("")}
  </tbody>
</table>
` : ''}

${topProv.length > 0 ? `
<h2>🏢 TOP 10 CUENTAS POR PAGAR (PROVEEDORES)</h2>
<table>
  <thead>
    <tr>
      <th style="width:25px;">#</th><th>Proveedor</th><th>RIF</th><th style="text-align:center;">Facturas</th><th style="text-align:right;">Saldo USD</th>
    </tr>
  </thead>
  <tbody>
    ${topProv.map((p, i) => `
      <tr>
        <td>${i + 1}</td>
        <td><b>${escapeHtml(p.nombre)}</b></td>
        <td>${escapeHtml(p.rif)}</td>
        <td style="text-align:center;">${p.facturas}</td>
        <td style="text-align:right;"><b>$${fmt(p.saldo)}</b></td>
      </tr>
    `).join("")}
  </tbody>
</table>
` : ''}

${pedidosPend.length > 0 ? `
<h2>📦 PEDIDOS EN COLA (${pedidosPend.length})</h2>
<table>
  <thead>
    <tr>
      <th>#</th><th>Cliente</th><th>Doc. Factura</th><th style="text-align:right;">Total USD</th><th style="text-align:center;">Estado</th><th>Registro</th>
    </tr>
  </thead>
  <tbody>
    ${pedidosPend.slice(0, 15).map(p => `
      <tr>
        <td>${escapeHtml(p["idPedido"] || p["IdNroPedido"] || "S/N")}</td>
        <td>${escapeHtml(p["nombreClientePedido"] || "N/A")}</td>
        <td>${escapeHtml(p["NroFacturaClientePedido"] || "N/A")}</td>
        <td style="text-align:right;"><b>$${fmt(p["TotalClienteUSD"])}</b></td>
        <td style="text-align:center;">${escapeHtml(p["StatusPedido"] || "Pendiente")}</td>
        <td>${formatFecha(p["FechaRegistroPedido"]) || "N/A"}</td>
      </tr>
    `).join("")}
  </tbody>
</table>
` : ''}

<div class="footer">
  ${escapeHtml(c.nombre)} — Documento generado exclusivamente para uso gerencial interno.
</div>

</body></html>`;
}
function enviarPdfPorEmail(chatId, messageId) {
  try {
    if (messageId) editarMensaje(chatId, messageId, "📄 <i>Compilando y generando PDF profesional...</i>", null);
    else enviarMensaje(chatId, "📄 <i>Compilando PDF...</i>", null);

    const fecha = Utilities.formatDate(new Date(), TIMEZONE, "yyyyMMdd_HHmm");
    const nombreArchivo = `GproPos_Reporte_${fecha}`;

    const html = construirHtmlPdfCompleto();
    const pdfBlob = generarPdfDesdeHtml(html, nombreArchivo);

    MailApp.sendEmail({
      to: BOT_CONFIG.emailsAdmin.join(","),
      subject: `📄 Reporte PDF Ejecutivo - ${BOT_CONFIG.empresa.nombre} - ${fecha}`,
      body: `Adjunto encontrarás el informe gerencial PDF consolidado al corte.\n\nGenerado: ${Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm")}`,
      attachments: [pdfBlob],
      name: BOT_CONFIG.empresa.nombre + " Bot"
    });

    const exitoMsg = 
      `✅ <b>PDF generado y despachado con éxito</b>\n\n` +
      `📧 Destinatarios:\n<code>${BOT_CONFIG.emailsAdmin.join("\n")}</code>\n\n` +
      `📄 Archivo: <code>${nombreArchivo}.pdf</code>`;

    if (messageId) editarMensaje(chatId, messageId, exitoMsg, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, exitoMsg, tecladoVolver("REPORTES"));
  } catch (e) {
    const err = `❌ Error en generación de PDF: <code>${escapeHtml(e.toString())}</code>`;
    if (messageId) editarMensaje(chatId, messageId, err, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, err, tecladoVolver("REPORTES"));
  }
}
function generarPdfBlob() {
  const html = construirHtmlPdfCompleto();
  const fecha = Utilities.formatDate(new Date(), TIMEZONE, "yyyyMMdd");
  return generarPdfDesdeHtml(html, `GproPos_Reporte_${fecha}`);
}

// ============================================================================
// SECCIÓN 16: RESPALDOS A DRIVE ☁️
// ============================================================================
function obtenerCarpetaRespaldos() {
  const nombreRaiz = BOT_CONFIG.drive.carpetaRaiz;
  const carpetas = DriveApp.getFoldersByName(nombreRaiz);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(nombreRaiz);
}

function obtenerCarpetaMes(carpetaRaiz) {
  const ahora = new Date();
  const nombreMes = Utilities.formatDate(ahora, TIMEZONE, "yyyy-MM");
  const subcarpetas = carpetaRaiz.getFoldersByName(nombreMes);
  if (subcarpetas.hasNext()) return subcarpetas.next();
  return carpetaRaiz.createFolder(nombreMes);
}

function realizarRespaldo(tipo) {
  tipo = tipo || "auto";
  try {
    log(`💾 Creando respaldo (${tipo})...`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const carpetaRaiz = obtenerCarpetaRespaldos();
    const carpetaMes = BOT_CONFIG.drive.carpetasPorMes ? obtenerCarpetaMes(carpetaRaiz) : carpetaRaiz;

    const ahora = new Date();
    const timestamp = Utilities.formatDate(ahora, TIMEZONE, "yyyyMMdd_HHmmss");
    const nombreBase = `${BOT_CONFIG.empresa.nombre}_${tipo}_${timestamp}`;

    const urlXlsx = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const blobXlsx = UrlFetchApp.fetch(urlXlsx, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    }).getBlob().setName(nombreBase + ".xlsx");
    const archivoXlsx = carpetaMes.createFile(blobXlsx);

    let archivoPdf = null;
    try {
      const pdfBlob = generarPdfBlob();
      pdfBlob.setName(nombreBase + ".pdf");
      archivoPdf = carpetaMes.createFile(pdfBlob);
    } catch (e) { log("⚠️ Error respaldo PDF: " + e.toString()); }

    let archivoJson = null;
    try {
      const datosBackup = {
        timestamp: ahora.toISOString(),
        tipo: tipo,
        empresa: BOT_CONFIG.empresa.nombre,
        hojas: {}
      };
      ["Pedidos", "Ventas", "Compras", "Tasas"].forEach(nombre => {
        datosBackup.hojas[nombre] = leerHoja(nombre);
      });
      const jsonBlob = Utilities.newBlob(JSON.stringify(datosBackup, null, 2), "application/json", nombreBase + ".json");
      archivoJson = carpetaMes.createFile(jsonBlob);
    } catch (e) { log("⚠️ Error respaldo JSON: " + e.toString()); }

    registrarRespaldoEnIndice({
      fecha: Utilities.formatDate(ahora, TIMEZONE, "dd/MM/yyyy HH:mm:ss"),
      tipo: tipo,
      nombreBase: nombreBase,
      urlXlsx: archivoXlsx.getUrl(),
      urlPdf: archivoPdf ? archivoPdf.getUrl() : "",
      urlJson: archivoJson ? archivoJson.getUrl() : "",
      carpeta: carpetaMes.getName(),
      tamañoXlsx: archivoXlsx.getSize(),
      idXlsx: archivoXlsx.getId()
    });

    limpiarRespaldosAntiguos(carpetaRaiz);

    return {
      exito: true, tipo, nombre: nombreBase,
      urlXlsx: archivoXlsx.getUrl(),
      urlPdf: archivoPdf ? archivoPdf.getUrl() : null,
      urlJson: archivoJson ? archivoJson.getUrl() : null,
      carpeta: carpetaMes.getName()
    };
  } catch (e) {
    log("❌ Error general respaldo: " + e.toString());
    return { exito: false, error: e.toString() };
  }
}

function registrarRespaldoEnIndice(info) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Backups");
    if (!hoja) {
      hoja = ss.insertSheet("Backups");
      const headers = ["Fecha", "Tipo", "Nombre", "Carpeta", "URL XLSX", "URL PDF", "URL JSON", "Tamaño (bytes)", "ID"];
      hoja.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground(BOT_CONFIG.empresa.color).setFontColor("#ffffff").setFontWeight("bold");
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([
      info.fecha, info.tipo, info.nombreBase, info.carpeta,
      info.urlXlsx, info.urlPdf, info.urlJson, info.tamañoXlsx, info.idXlsx
    ]);
  } catch (e) { log("ERROR al registrar índice de backups: " + e.toString()); }
}

function obtenerHistorialRespaldos(limite) {
  limite = limite || 8;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Backups");
    if (!hoja || hoja.getLastRow() <= 1) return [];
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 9).getValues();
    return datos.map(f => ({
      fecha: f[0], tipo: f[1], nombre: f[2], carpeta: f[3],
      urlXlsx: f[4], urlPdf: f[5], urlJson: f[6], tamaño: f[7], id: f[8]
    })).reverse().slice(0, limite);
  } catch (e) { return []; }
}

function limpiarRespaldosAntiguos(carpetaRaiz) {
  try {
    const meses = BOT_CONFIG.drive.mantenerMeses;
    const ahora = new Date();
    const limite = new Date(ahora.getFullYear(), ahora.getMonth() - meses, 1);
    const carpetas = carpetaRaiz.getFolders();

    while (carpetas.hasNext()) {
      const carpeta = carpetas.next();
      const match = carpeta.getName().match(/^(\d{4})-(\d{2})$/);
      if (match) {
        const fecha = new Date(parseInt(match[1]), parseInt(match[2]) - 1, 1);
        if (fecha < limite) carpeta.setTrashed(true);
      }
    }
  } catch (e) { log("Error depurando backups: " + e.toString()); }
}

function enviarHistorialRespaldos(chatId, messageId) {
  const historial = obtenerHistorialRespaldos(6);

  if (historial.length === 0) {
    const msgVacio = `💾 <b>HISTORIAL DE RESPALDOS</b>\n━━━━━━━━━━━━━━━━━━━━\n\nℹ️ Aún no existen respaldos archivados.`;
    if (messageId) editarMensaje(chatId, messageId, msgVacio, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, msgVacio, tecladoVolver("REPORTES"));
    return;
  }

  let texto = `💾 <b>ÚLTIMOS RESPALDOS EN DRIVE</b>\n━━━━━━━━━━━━━━━━━━━━\n\n`;
  historial.forEach(r => {
    const icono = r.tipo === "manual" ? "🔧" : "🤖";
    const tam = r.tamaño ? (r.tamaño / 1024).toFixed(1) + " KB" : "—";
    texto += `${icono} <b>${r.fecha}</b>\n`;
    texto += `   📁 Carpeta: <code>${escapeHtml(r.carpeta)}</code>\n`;
    texto += `   🏷️ ${escapeHtml(r.tipo)} | 💾 ${tam}\n`;
    texto += `   🔗 `;
    if (r.urlXlsx) texto += `<a href="${r.urlXlsx}">[Excel]</a> `;
    if (r.urlPdf) texto += `<a href="${r.urlPdf}">[PDF]</a> `;
    if (r.urlJson) texto += `<a href="${r.urlJson}">[JSON]</a>`;
    texto += `\n────────────────────\n`;
  });

  const teclado = {
    inline_keyboard: [
      [{ text: "💾 Crear Backup Inmediato", callback_data: "BACKUP_CREAR" }],
      [{ text: "📧 Volver a Reportes", callback_data: "REPORTES" }]
    ]
  };

  if (messageId) editarMensaje(chatId, messageId, texto, teclado);
  else enviarMensaje(chatId, texto, teclado);
}

function realizarRespaldoManual(chatId, messageId) {
  if (messageId) editarMensaje(chatId, messageId, "💾 <i>Exportando hojas y creando copias en Drive...</i>", null);
  else enviarMensaje(chatId, "💾 <i>Creando respaldo en Drive...</i>", null);

  const r = realizarRespaldo("manual");

  if (r.exito) {
    const msg =
      `✅ <b>Copia de seguridad finalizada</b>\n\n` +
      `📦 <b>Archivo:</b> <code>${r.nombre}</code>\n` +
      `📁 <b>Carpeta:</b> <code>${r.carpeta}</code>\n\n` +
      `🔗 <b>Acceso a los archivos:</b>\n` +
      `<a href="${r.urlXlsx}">📊 Descargar Excel</a>` +
      (r.urlPdf ? ` | <a href="${r.urlPdf}">📄 Ver PDF</a>` : "");

    if (messageId) editarMensaje(chatId, messageId, msg, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, msg, tecladoVolver("REPORTES"));
  } else {
    const err = `❌ <b>Error al respaldar:</b>\n<code>${escapeHtml(r.error)}</code>`;
    if (messageId) editarMensaje(chatId, messageId, err, tecladoVolver("REPORTES"));
    else enviarMensaje(chatId, err, tecladoVolver("REPORTES"));
  }
}

// ============================================================================
// SECCIÓN 17: INSTALACIÓN Y DIAGNÓSTICO
// ============================================================================
function VERIFICAR_WEBHOOK_ESTABLE() {
  const info = JSON.parse(UrlFetchApp.fetch(`${TELEGRAM_API}/getWebhookInfo`).getContentText());
  
  Logger.log("⏰ Verificación ejecutada: " + new Date().toLocaleTimeString());
  Logger.log("📊 Estado de la conexión:");
  Logger.log("   URL actual: " + info.result.url);
  Logger.log("   Updates pendientes en cola: " + info.result.pending_update_count);
  Logger.log("   Último error reportado: " + (info.result.last_error_message || "Ninguno (100% Estable ✅)"));
  
  if (info.result.pending_update_count === 0 && !info.result.last_error_message) {
    Logger.log("\n✅ EL WEBHOOK ESTÁ COMPLETAMENTE OPERATIVO Y SIN RETRASOS.");
  } else {
    Logger.log("\n⚠️ ATENCIÓN: Hay actualizaciones acumuladas o errores reportados por Telegram.");
  }
}

function REPARAR_Y_LIMPIAR_WEBHOOK() {
  const URL_WEBHOOK = "https://script.google.com/macros/s/AKfycbw2Ju09wlH6nebzXV_UIJJJRkkhKMXlpNRwsnSBINZyYgjP7WZcZ1vKvQXUiID1QlLSSg/exec";
  
  Logger.log("1️⃣ Purgando updates atascados y eliminando webhook previo...");
  UrlFetchApp.fetch(`${TELEGRAM_API}/deleteWebhook?drop_pending_updates=true`);
  Utilities.sleep(2000);

  Logger.log("2️⃣ Registrando nuevo webhook con acceso público...");
  const respuesta = UrlFetchApp.fetch(`${TELEGRAM_API}/setWebhook?url=${encodeURIComponent(URL_WEBHOOK)}&drop_pending_updates=true&max_connections=40`);
  Logger.log("   Respuesta de Telegram: " + respuesta.getContentText());
  Utilities.sleep(2000);

  Logger.log("3️⃣ Comprobando estado final...");
  const info = JSON.parse(UrlFetchApp.fetch(`${TELEGRAM_API}/getWebhookInfo`).getContentText());
  Logger.log("   URL: " + info.result.url);
  Logger.log("   Pendientes: " + info.result.pending_update_count);
  Logger.log("   Error: " + (info.result.last_error_message || "¡Ninguno! Conexión limpia ✅"));
}
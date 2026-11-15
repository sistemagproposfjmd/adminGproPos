// ============================================================================
// SISTEMA ADMINISTRATIVO GPROPOS - MÓDULO DE VENTAS Y PEDIDOS POR TELEGRAM
// Archivo: pedidos_telegram.gs  (v2 - alineado a guardarPedido de code.gs)
// ============================================================================

// 🔧 CAMBIO: cliente genérico virtual, coherente con obtenerDatosClientePorNombre
const CLIENTE_GENERICO = {
  id: "",
  nombre: "CLIENTE GENÉRICO",
  rif: "S/R",
  tipo: "Jurídico",          // 🔧 el server fuerza "Jurídico" en TipoClientePedido
  correo: "",
  direccion: "Mostrador / Venta directa",
  telefono: "",
  vendedor: ""               // 🔧 viene del cliente; genérico no tiene
};

// ============================================================================
// GESTIÓN DEL CARRITO EN MEMORIA (CacheService)
// ============================================================================
function obtenerCarritoUsuario(chatId) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get("cart_" + chatId);
  if (!raw) {
    return { items: [], cliente: CLIENTE_GENERICO, timestamp: Date.now() };
  }
  try {
    return JSON.parse(raw);
  } catch (e) {
    return { items: [], cliente: CLIENTE_GENERICO, timestamp: Date.now() };
  }
}

function guardarCarritoUsuario(chatId, cartObj) {
  const cache = CacheService.getScriptCache();
  cache.put("cart_" + chatId, JSON.stringify(cartObj), 21600); // 6 horas
}

function vaciarCarritoUsuario(chatId) {
  CacheService.getScriptCache().remove("cart_" + chatId);
}

// ============================================================================
// 🔧 CAMBIO: helper de precio de venta (col D) y validación de producto vendible
// ============================================================================
function _precioVentaProducto(p) {
  // Precio de venta real = columna "precioVenta" (D).
  // Si viniera vacío o 0, NO se cae a PrecioProducto (eso sería vender al costo).
  const pv = parseNumero(p["precioVenta"]);
  return pv > 0 ? pv : 0;
}

function _esProductoVendible(p) {
  if (!p) return false;
  // 1. Status estrictamente "activo" (case-insensitive)
  const status = String(p["Status"] || "").trim().toLowerCase();
  if (status !== "activo") return false;
  // 2. Excluir compuestos
  const tipo = String(p["tipo"] || "").trim().toLowerCase();
  if (tipo === "compuesto") return false;
  // 3. Debe tener precio de venta > 0
  if (_precioVentaProducto(p) <= 0) return false;
  // 4. Debe tener stock > 0 (stock puede venir con coma decimal)
  const stock = parseNumero(p["StockProducto"]);
  if (stock <= 0) return false;
  return true;
}

// ============================================================================
// CATÁLOGO INTERACTIVO
// ============================================================================
function mostrarCatalogoTelegram(chatId, pagina, messageId) {
  pagina = parseInt(pagina) || 1;
  const limite = 6;

  // 🔧 CAMBIO: filtro único y correcto
  const prods = leerHoja("Productos").filter(_esProductoVendible);

  if (prods.length === 0) {
    const txt = "📦 <b>CATÁLOGO DE PRODUCTOS</b>\n━━━━━━━━━━━━━━━━━━━━\n\n⚠️ No hay productos disponibles para venta en este momento.";
    if (messageId) editarMensaje(chatId, messageId, txt, tecladoVolver("MENU"));
    else enviarMensaje(chatId, txt, tecladoVolver("MENU"));
    return;
  }

  const totalPaginas = Math.ceil(prods.length / limite);
  pagina = Math.max(1, Math.min(pagina, totalPaginas));
  const inicio = (pagina - 1) * limite;
  const itemsPagina = prods.slice(inicio, inicio + limite);

  const tasa = obtenerTasaActual();
  const cart = obtenerCarritoUsuario(chatId);
  const totalItemsEnCarrito = cart.items.reduce((acc, it) => acc + it.cantidad, 0);

  let texto = `🛍️ <b>CATÁLOGO DE PRODUCTOS</b> (Pág. ${pagina}/${totalPaginas})\n`;
  texto += `💵 Tasa oficial: <b>${fmt(tasa)} Bs/$</b>\n`;
  texto += `🛒 En tu carrito: <b>${totalItemsEnCarrito} producto(s)</b>\n`;
  texto += `━━━━━━━━━━━━━━━━━━━━\n\n`;

  const inlineKeyboard = [];

  itemsPagina.forEach((p) => {
    const id = p["IdProducto"] || "S/N";
    const nombre = p["NombreProducto"] || "Sin descripción";
    const precioUSD = _precioVentaProducto(p); // 🔧 col D
    const stock = parseNumero(p["StockProducto"]);
    const precioBs = precioUSD * tasa;

    texto += `🔹 <b>${escapeHtml(nombre)}</b>\n`;
    texto += `   💵 <b>$${fmt(precioUSD)}</b> | Bs ${fmt(precioBs)}\n`;
    texto += `   📦 Stock disponible: <code>${stock}</code>\n\n`;

    inlineKeyboard.push([
      { text: `➕ Agregar: ${nombre.substring(0, 24)}`, callback_data: `CART_ADD:${id}` }
    ]);
  });

  const navFila = [];
  if (pagina > 1) {
    navFila.push({ text: "⬅️ Ant.", callback_data: `CATALOGO:${pagina - 1}` });
  }
  navFila.push({ text: `🛒 Ver Carrito (${totalItemsEnCarrito})`, callback_data: `CART_VER` });
  if (pagina < totalPaginas) {
    navFila.push({ text: "Sig. ➡️", callback_data: `CATALOGO:${pagina + 1}` });
  }
  inlineKeyboard.push(navFila);
  inlineKeyboard.push([{ text: "🏠 Menú Principal", callback_data: "MENU" }]);

  const teclado = { inline_keyboard: inlineKeyboard };

  if (messageId) editarMensaje(chatId, messageId, texto, teclado);
  else enviarMensaje(chatId, texto, teclado);
}

// ============================================================================
// AGREGAR Y MANIPULAR CARRITO
// ============================================================================
function agregarProductoCarrito(chatId, idProducto, queryId) {
  const prods = leerHoja("Productos");
  const p = prods.find(x => String(x["IdProducto"]).trim() === String(idProducto).trim());

  if (!p || !_esProductoVendible(p)) {
    telegramRequest("answerCallbackQuery", { callback_query_id: queryId, text: "❌ Producto no disponible" });
    return;
  }

  const stockDisp = parseNumero(p["StockProducto"]);
  const cart = obtenerCarritoUsuario(chatId);
  const itemEnCart = cart.items.find(x => String(x.idProducto) === String(idProducto));
  const cantidadActual = itemEnCart ? itemEnCart.cantidad : 0;

  if (cantidadActual + 1 > stockDisp) {
    telegramRequest("answerCallbackQuery", {
      callback_query_id: queryId,
      text: `⚠️ Stock insuficiente. Solo quedan ${stockDisp} unidades.`,
      show_alert: true
    });
    return;
  }

  const precioUSD = _precioVentaProducto(p); // 🔧 col D

  if (itemEnCart) {
    itemEnCart.cantidad++;
  } else {
    cart.items.push({
      idProducto: idProducto,
      nombre: p["NombreProducto"] || "Producto",
      precioUSD: precioUSD,
      cantidad: 1
    });
  }

  guardarCarritoUsuario(chatId, cart);
  telegramRequest("answerCallbackQuery", { callback_query_id: queryId, text: `✅ Agregado: ${p["NombreProducto"]}` });
}

function verCarritoTelegram(chatId, messageId) {
  const cart = obtenerCarritoUsuario(chatId);
  const tasa = obtenerTasaActual();

  if (!cart.items || cart.items.length === 0) {
    const txtVacio = `🛒 <b>CARRITO DE COMPRAS</b>\n━━━━━━━━━━━━━━━━━━━━\n\nTu carrito está vacío actualmente.\nSelecciona productos desde el catálogo para armar un pedido.`;
    const tecladoVacio = {
      inline_keyboard: [
        [{ text: "🛍️ Ir al Catálogo", callback_data: "CATALOGO:1" }],
        [{ text: "🏠 Menú Principal", callback_data: "MENU" }]
      ]
    };
    if (messageId) editarMensaje(chatId, messageId, txtVacio, tecladoVacio);
    else enviarMensaje(chatId, txtVacio, tecladoVacio);
    return;
  }

  // 🔧 CAMBIO: sin IVA. El subtotal ES el total.
  let subtotal = 0;
  let texto = `🛒 <b>CARRITO DE PEDIDO</b>\n`;
  texto += `👤 Cliente: <b>${escapeHtml(cart.cliente.nombre)}</b>\n`;
  texto += `🆔 RIF: <code>${escapeHtml(cart.cliente.rif)}</code>\n`;
  texto += `━━━━━━━━━━━━━━━━━━━━\n\n`;

  const inlineKeyboard = [];

  cart.items.forEach((it, idx) => {
    const totalLinea = it.precioUSD * it.cantidad;
    subtotal += totalLinea;

    texto += `🔹 <b>${escapeHtml(it.nombre)}</b>\n`;
    texto += `   ${it.cantidad} un. x $${fmt(it.precioUSD)} = <b>$${fmt(totalLinea)} USD</b>\n`;

    inlineKeyboard.push([
      { text: `➖`, callback_data: `CART_DEC:${idx}` },
      { text: `${it.cantidad} un.`, callback_data: `NOOP` },
      { text: `➕`, callback_data: `CART_INC:${idx}` },
      { text: `🗑️`, callback_data: `CART_DEL:${idx}` }
    ]);
  });

  const total = subtotal; // 🔧 sin IVA
  const totalBs = total * tasa;

  texto += `━━━━━━━━━━━━━━━━━━━━\n`;
  texto += `💰 <b>TOTAL: $${fmt(total)} USD</b>\n`;
  texto += `🇻🇪 <b>Equivalente: Bs ${fmt(totalBs)}</b> (Tasa: ${fmt(tasa)})\n`;
  texto += `💡 <i>Operación exenta de IVA</i>\n\n`;
  texto += `¿Confirmar y registrar el pedido?`;

  inlineKeyboard.push([
    { text: "👤 Cambiar Cliente", callback_data: "CLIENTE_MENU" },
    { text: "🛍️ + Productos", callback_data: "CATALOGO:1" }
  ]);
  inlineKeyboard.push([
    { text: "✅ CONFIRMAR Y GUARDAR PEDIDO", callback_data: "PEDIDO_GUARDAR" }
  ]);
  inlineKeyboard.push([
    { text: "🗑️ Vaciar Carrito", callback_data: "CART_CLEAR" },
    { text: "🏠 Salir", callback_data: "MENU" }
  ]);

  const teclado = { inline_keyboard: inlineKeyboard };

  if (messageId) editarMensaje(chatId, messageId, texto, teclado);
  else enviarMensaje(chatId, texto, teclado);
}

function modificarCantidadItem(chatId, idx, delta, messageId) {
  const cart = obtenerCarritoUsuario(chatId);
  idx = parseInt(idx);
  if (!cart.items[idx]) return;

  const prods = leerHoja("Productos");
  const p = prods.find(x => String(x["IdProducto"]).trim() === String(cart.items[idx].idProducto).trim());
  const stockDisp = p ? parseNumero(p["StockProducto"]) : 999;

  cart.items[idx].cantidad += delta;

  if (cart.items[idx].cantidad > stockDisp) {
    cart.items[idx].cantidad = stockDisp;
  }

  if (cart.items[idx].cantidad <= 0) {
    cart.items.splice(idx, 1);
  }

  guardarCarritoUsuario(chatId, cart);
  verCarritoTelegram(chatId, messageId);
}

function eliminarItemCarrito(chatId, idx, messageId) {
  const cart = obtenerCarritoUsuario(chatId);
  idx = parseInt(idx);
  if (cart.items[idx]) {
    cart.items.splice(idx, 1);
    guardarCarritoUsuario(chatId, cart);
  }
  verCarritoTelegram(chatId, messageId);
}

// ============================================================================
// GESTIÓN DE CLIENTES
// ============================================================================
function menuSeleccionCliente(chatId, messageId) {
  const clientes = leerHoja("Clientes").filter(c => {
    const st = String(c["Status"] || "Activo").trim().toLowerCase();
    return st !== "inactivo";
  }).slice(0, 6);

  let texto = `👤 <b>SELECCIONAR CLIENTE</b>\n━━━━━━━━━━━━━━━━━━━━\n\n`;
  texto += `Elige un cliente, usa el Genérico, o busca con <code>/buscar_cliente NOMBRE</code>.\n\n`;

  const inlineKeyboard = [
    [{ text: "👤 CONSUMIDOR FINAL (Genérico)", callback_data: "CLIENTE_SET:GENERICO" }]
  ];

  clientes.forEach(c => {
    const id = c["idClientes"] || "";
    const nombre = c["NombreClientes"] || "Cliente";
    inlineKeyboard.push([
      { text: `🏢 ${nombre.substring(0, 26)}`, callback_data: `CLIENTE_SET:${id}` }
    ]);
  });

  inlineKeyboard.push([{ text: "⬅️ Volver al Carrito", callback_data: "CART_VER" }]);

  const teclado = { inline_keyboard: inlineKeyboard };
  if (messageId) editarMensaje(chatId, messageId, texto, teclado);
  else enviarMensaje(chatId, texto, teclado);
}

function asignarClienteAlCarrito(chatId, idCliente, messageId) {
  const cart = obtenerCarritoUsuario(chatId);

  if (idCliente === "GENERICO") {
    cart.cliente = CLIENTE_GENERICO;
  } else {
    const clientes = leerHoja("Clientes");
    const c = clientes.find(x => String(x["idClientes"]).trim() === String(idCliente).trim());
    if (c) {
      cart.cliente = {
        id: c["idClientes"] || "",
        nombre: c["NombreClientes"] || "Cliente",
        rif: c["RifClientes"] || "S/R",
        tipo: "Jurídico",
        correo: c["CorreoCliente"] || "",
        direccion: c["DireccionCliente"] || "",
        telefono: c["TelefonoCliente"] || "",
        vendedor: c["Vendedor"] || ""   // 🔧 se manda al server como vendedor del pedido
      };
    }
  }

  guardarCarritoUsuario(chatId, cart);
  verCarritoTelegram(chatId, messageId);
}

function buscarClienteComando(chatId, criterio) {
  if (!criterio) {
    enviarMensaje(chatId, "⚠️ Escribe el nombre o RIF a buscar:\nEjemplo: <code>/buscar_cliente FARMATODO</code>", tecladoPrincipal());
    return;
  }

  const crit = criterio.toLowerCase().trim();
  const clientes = leerHoja("Clientes").filter(c => {
    const st = String(c["Status"] || "Activo").trim().toLowerCase();
    return st !== "inactivo";
  });
  const matches = clientes.filter(c => {
    const nom = String(c["NombreClientes"] || "").toLowerCase();
    const rif = String(c["RifClientes"] || "").toLowerCase();
    return nom.includes(crit) || rif.includes(crit);
  }).slice(0, 6);

  if (matches.length === 0) {
    enviarMensaje(chatId, `❌ No se encontró cliente con: "<b>${escapeHtml(criterio)}</b>"`, tecladoPrincipal());
    return;
  }

  let txt = `🔍 <b>RESULTADOS PARA "${escapeHtml(criterio)}":</b>\nToca un cliente para asignarlo al carrito:\n\n`;
  const inlineKeyboard = [];

  matches.forEach(c => {
    const id = c["idClientes"] || "";
    const nom = c["NombreClientes"] || "Cliente";
    const rif = c["RifClientes"] || "S/R";
    inlineKeyboard.push([
      { text: `🏢 ${nom} (${rif})`, callback_data: `CLIENTE_SET:${id}` }
    ]);
  });
  inlineKeyboard.push([{ text: "🛒 Volver al Carrito", callback_data: "CART_VER" }]);

  enviarMensaje(chatId, txt, { inline_keyboard: inlineKeyboard });
}

function crearClienteRapidoComando(chatId, textoComando) {
  const partes = textoComando.split("|").map(s => s.trim());
  if (partes.length < 2) {
    enviarMensaje(chatId,
      `⚠️ <b>Formato de creación rápida:</b>\n\n` +
      `<code>/nuevo_cliente Nombre | RIF | Teléfono | Correo</code>\n\n` +
      `Ejemplo:\n<code>/nuevo_cliente Inversiones Caracas | J-50123456-7 | 04141234567 | ventas@caracas.com</code>`,
      tecladoPrincipal()
    );
    return;
  }

  const nombre = partes[0];
  const rif = partes[1];
  const telefono = partes[2] || "N/A";
  const correo = partes[3] || "sistemagpropos@gmail.com";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Clientes");
    if (!hoja) {
      enviarMensaje(chatId, "❌ Error: La hoja 'Clientes' no existe.", tecladoPrincipal());
      return;
    }

    const datos = hoja.getDataRange().getValues();
    const headers = datos[0].map(h => String(h).trim());
    const iRif = headers.indexOf("RifClientes");
    const iId = headers.indexOf("idClientes");

    // 🔧 CAMBIO: verificar RIF duplicado
    if (iRif >= 0) {
      const rifNorm = rif.toLowerCase().replace(/[\s\-\.]/g, "");
      for (let i = 1; i < datos.length; i++) {
        const r = String(datos[i][iRif] || "").toLowerCase().replace(/[\s\-\.]/g, "");
        if (r && r === rifNorm) {
          enviarMensaje(chatId,
            `⚠️ Ya existe un cliente con RIF <code>${escapeHtml(rif)}</code>:\n` +
            `🏢 ${escapeHtml(datos[i][headers.indexOf("NombreClientes")] || "")}\n\n` +
            `Usa <code>/buscar_cliente ${escapeHtml(nombre)}</code> para asignarlo.`,
            tecladoPrincipal()
          );
          return;
        }
      }
    }

    // 🔧 CAMBIO: ID = max CLI###### + 1
    let maxNum = 0;
    if (iId >= 0) {
      for (let i = 1; i < datos.length; i++) {
        const m = String(datos[i][iId] || "").match(/^CLI(\d+)$/);
        if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
      }
    }
    const nuevoId = "CLI" + String(maxNum + 1).padStart(6, "0");

    // 🔧 CAMBIO: escribir por header, no por posición
    const fila = headers.map(h => {
      switch (h) {
        case "idClientes": return nuevoId;
        case "NombreClientes": return nombre;
        case "RifClientes": return rif;
        case "DireccionCliente": return "";
        case "CorreoCliente": return correo;
        case "TelefonoCliente": return telefono;
        case "TipoContribuyenteCliente": return "Formal";
        case "RetIVACliente": return 0;
        case "RetISLRCliente": return 0;
        case "Responsable": return "";
        case "Vendedor": return "";          // 🔧 sin vendedor asignado; que lo asigne el admin
        case "Status": return "Activo";
        default: return "";
      }
    });
    hoja.appendRow(fila);

    const cart = obtenerCarritoUsuario(chatId);
    cart.cliente = {
      id: nuevoId,
      nombre: nombre,
      rif: rif,
      tipo: "Jurídico",
      correo: correo,
      direccion: "",
      telefono: telefono,
      vendedor: ""
    };
    guardarCarritoUsuario(chatId, cart);

    enviarMensaje(chatId,
      `✅ <b>Cliente registrado y asignado:</b>\n` +
      `🏢 <b>${escapeHtml(nombre)}</b>\n🆔 RIF: <code>${escapeHtml(rif)}</code>\n🆔 ID: <code>${nuevoId}</code>\n\n` +
      `Ya puedes confirmar tu pedido.`,
      { inline_keyboard: [[{ text: "🛒 Ir al Carrito", callback_data: "CART_VER" }]] }
    );
  } catch (e) {
    enviarMensaje(chatId, `❌ Error al crear cliente: <code>${escapeHtml(e.toString())}</code>`, tecladoPrincipal());
  }
}

// ============================================================================
// GUARDADO DEL PEDIDO — DELEGA EN guardarPedido() DE code.gs
// ============================================================================
function procesarConfirmacionPedido(chatId, usuarioTelegram, messageId) {
  const cart = obtenerCarritoUsuario(chatId);

  if (!cart.items || cart.items.length === 0) {
    editarMensaje(chatId, messageId, "❌ El carrito está vacío.", tecladoVolver("MENU"));
    return;
  }

  editarMensaje(chatId, messageId, "⏳ <i>Registrando pedido...</i>", null);

  // 🔧 CAMBIO: re-validar stock y disponibilidad contra la hoja fresca (sin descontar)
  try {
    const prods = leerHoja("Productos");
    const problemas = [];
    cart.items.forEach(it => {
      const p = prods.find(x => String(x["IdProducto"]).trim() === String(it.idProducto).trim());
      if (!p || !_esProductoVendible(p)) {
        problemas.push(`❌ ${it.nombre}: ya no está disponible`);
        return;
      }
      const stock = parseNumero(p["StockProducto"]);
      if (it.cantidad > stock) {
        problemas.push(`⚠️ ${it.nombre}: pediste ${it.cantidad}, disponible ${stock}`);
      }
    });
    if (problemas.length > 0) {
      editarMensaje(chatId, messageId,
        `⚠️ <b>No se pudo confirmar el pedido:</b>\n\n` + problemas.join("\n") +
        `\n\nAjusta el carrito o intenta de nuevo.`,
        tecladoVolver("CART_VER")
      );
      return;
    }
  } catch (e) {
    editarMensaje(chatId, messageId, `❌ Error validando stock: <code>${escapeHtml(e.toString())}</code>`, tecladoVolver("CART_VER"));
    return;
  }

  // 🔧 CAMBIO: construir payload tal como lo espera guardarPedido() de code.gs
  const idSpreadsheet = obtenerIdSpreadsheetParaChat(chatId);
  const tasa = obtenerTasaActual();
  const fecha = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy");

  const carritoServer = cart.items.map(it => ({
    id: String(it.idProducto),
    nombre: String(it.nombre),
    precio: Number(it.precioUSD),
    cantidad: parseInt(it.cantidad, 10)
  }));

  const payload = {
    idSpreadsheet: idSpreadsheet,
    idPedido: 0,                                  // 0 = nuevo; el server calcula el consecutivo
    cliente: {
      id: cart.cliente.id || "",
      nombre: cart.cliente.nombre || "CLIENTE GENÉRICO",
      rif: cart.cliente.rif || ""
    },
    fecha: fecha,
    tasa: tasa,
    carrito: carritoServer,
    vendedor: cart.cliente.vendedor || "",        // 🔧 del cliente
    correoDestino: cart.cliente.correo || ""
  };

  // 🔧 CAMBIO: delegar TODO en guardarPedido (ID, NroFactura, Detalles, totales, emailCliente)
  let r;
  try {
    r = guardarPedido(payload);
  } catch (e) {
    editarMensaje(chatId, messageId, `❌ Error al guardar: <code>${escapeHtml(e.toString())}</code>`, tecladoVolver("CART_VER"));
    return;
  }

  if (!r || r.status !== "success") {
    editarMensaje(chatId, messageId,
      `❌ No se pudo guardar el pedido:\n<code>${escapeHtml((r && r.message) || "Error desconocido")}</code>`,
      tecladoVolver("CART_VER")
    );
    return;
  }

  // Vaciar carrito solo si guardó bien
  vaciarCarritoUsuario(chatId);

  const idPedido = r.idPedido;
  const nroFactura = "PED-" + String(idPedido).padStart(4, "0");
  const fechaReg = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm");

  let mensajeExito =
    `🎉 <b>¡PEDIDO ${nroFactura} REGISTRADO!</b>\n` +
    `━━━━━━━━━━━━━━━━━━━━\n\n` +
    `👤 <b>Cliente:</b> ${escapeHtml(r.nombreCliente || cart.cliente.nombre)}\n` +
    `🆔 <b>RIF:</b> <code>${escapeHtml(r.rifCliente || cart.cliente.rif)}</code>\n` +
    `💼 <b>Vendedor:</b> ${escapeHtml(payload.vendedor || "Sin asignar")}\n` +
    `📅 <b>Fecha:</b> ${fechaReg}\n\n` +
    `<b>Renglones:</b>\n`;

  (r.items || cart.items).forEach(it => {
    const precio = Number(it.precio || it.precioUSD || 0);
    const cant = Number(it.cantidad || 0);
    mensajeExito += `• ${cant}x ${escapeHtml(it.nombre)} ($${fmt(precio * cant)})\n`;
  });

  mensajeExito += `\n💵 <b>Total USD:</b> <b>$${fmt(r.totalUSD)}</b>\n`;
  mensajeExito += `🇻🇪 <b>Total Bs:</b> <b>Bs ${fmt(r.totalBs)}</b>\n`;
  mensajeExito += `📌 <b>Estado:</b> <code>Pendiente por Facturar</code>\n`;
  mensajeExito += `💡 <i>Operación exenta de IVA</i>`;

  const tecladoFinal = {
    inline_keyboard: [
      [{ text: "🛍️ Nuevo Pedido", callback_data: "CATALOGO:1" }],
      [{ text: "📦 Ver Pedidos", callback_data: "PEDIDOS" }],
      [{ text: "🏠 Menú Principal", callback_data: "MENU" }]
    ]
  };

  editarMensaje(chatId, messageId, mensajeExito, tecladoFinal);
}

// ============================================================================
// 🔧 NUEVO: resolución del idSpreadsheet (Opción B — configurable por chat)
// ============================================================================
function obtenerIdSpreadsheetParaChat(chatId) {
  // 1. Buscar configuración por chatId
  const props = PropertiesService.getScriptProperties();
  const porChat = props.getProperty("empresa_chat_" + chatId);
  if (porChat && porChat.length > 15) return porChat;

  // 2. Fallback: ID_HOJA_DEFAULT de code.gs (si existe en el proyecto)
  if (typeof ID_HOJA_DEFAULT !== "undefined" && ID_HOJA_DEFAULT) {
    return ID_HOJA_DEFAULT;
  }

  // 3. Último recurso: hoja activa
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function comandoConfigurarEmpresa(chatId, idSpreadsheet) {
  if (!idSpreadsheet || idSpreadsheet.length < 15) {
    enviarMensaje(chatId,
      `⚠️ <b>Falta el ID de la hoja</b>\n\n` +
      `Uso: <code>/empresa ID_DE_LA_HOJA</code>\n\n` +
      `Puedes obtenerlo de la URL de tu Google Sheet:\n` +
      `<code>https://docs.google.com/spreadsheets/d/<b>ESTE_ES_EL_ID</b>/edit</code>`,
      tecladoPrincipal()
    );
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(idSpreadsheet);
    const hojas = ["Productos", "Clientes", "Pedidos", "DetallesPedidos"];
    const faltantes = hojas.filter(h => !ss.getSheetByName(h));
    if (faltantes.length > 0) {
      enviarMensaje(chatId,
        `⚠️ La hoja <code>${escapeHtml(idSpreadsheet)}</code> existe pero le faltan pestañas:\n` +
        faltantes.map(h => "• " + h).join("\n"),
        tecladoPrincipal()
      );
      return;
    }

    PropertiesService.getScriptProperties().setProperty("empresa_chat_" + chatId, idSpreadsheet);
    enviarMensaje(chatId,
      `✅ <b>Empresa configurada para este chat</b>\n\n` +
      `📄 Hoja: <code>${escapeHtml(ss.getName())}</code>\n` +
      `🆔 ID: <code>${escapeHtml(idSpreadsheet)}</code>\n\n` +
      `Los pedidos que confirmes se guardarán en esta hoja.`,
      tecladoPrincipal()
    );
  } catch (e) {
    enviarMensaje(chatId, `❌ No se pudo abrir la hoja: <code>${escapeHtml(e.toString())}</code>`, tecladoPrincipal());
  }
}

// ============================================================================
// TASA — fuente única: obtenerTasaCambioActual() de code.gs
// ============================================================================
function obtenerTasaActual() {
  try {
    if (typeof obtenerTasaCambioActual === "function") {
      const t = obtenerTasaCambioActual();
      if (t && t.tasaUSD > 0) return t.tasaUSD;
    }
  } catch (e) {
    log("⚠️ Error obtenerTasaCambioActual: " + e.toString());
  }
  return 1;
}
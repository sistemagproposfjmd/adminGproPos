/**
 * ============================================================================
 * GESTIONEMPRESAS.GS - Backend para Administración y Configuración de Empresas
 * ============================================================================
 */

function apiObtenerDatosGestionEmpresas(token) {
  try {
    const check = validarSessionToken(token);
    if (!check) return { success: false, message: 'Sesión inválida.' };

    const rol = String(check.role || '').toLowerCase();
    if (rol !== 'admin' && rol !== 'owner') {
      return { success: false, message: 'Acceso exclusivo para Administradores y Owners.' };
    }

    const ssMaster = Contexto.getMasterSS();
    const sheetEmp = ssMaster.getSheetByName('Empresas') || ssMaster.getSheetByName('empresas');
    const sheetUsu = ssMaster.getSheetByName('Usuarios');
    const sheetPerm = ssMaster.getSheetByName('PermisosUsuarioEmpresa');

    // 1. Obtener empresas
    const dataEmp = sheetEmp ? sheetEmp.getDataRange().getValues() : [];
    const empresas = [];
    if (dataEmp.length >= 2) {
      const h = dataEmp[0].map(x => String(x).trim().toLowerCase());
      for (let i = 1; i < dataEmp.length; i++) {
        const r = dataEmp[i];
        const idEmp = String(r[h.indexOf('idempresa')] || '').trim();
        if (!idEmp) continue;

        empresas.push({
          idEmpresa: idEmp,
          nombreEmpresa: String(r[h.indexOf('nombreempresa')] || ''),
          rifEmpresa: String(r[h.indexOf('rifempresa')] || ''),
          direccionEmpresa: String(r[h.indexOf('direccionempresa')] || ''),
          telefonoEmpresa: String(r[h.indexOf('telefonoempresa')] || ''),
          correoEmpresa: String(r[h.indexOf('correoempresa')] || ''),
          idSpreadsheet: String(r[h.indexOf('idspreadsheet')] || ''),
          status: String(r[h.indexOf('status')] || 'Activo'),
          logo: String(r[h.indexOf('logo')] || ''),
          porcentajeIVA: r[h.indexOf('porcentajeiva')] !== undefined && r[h.indexOf('porcentajeiva')] !== '' ? r[h.indexOf('porcentajeiva')] : 75,
          porcentajeISLR: r[h.indexOf('porcentajeislr')] !== undefined && r[h.indexOf('porcentajeislr')] !== '' ? r[h.indexOf('porcentajeislr')] : 2,
          manejaIGTF: String(r[h.indexOf('manejaigtf')]).toUpperCase() === 'TRUE' || String(r[h.indexOf('manejaigtf')]) === 'SI',
          manejaAlcaldia: String(r[h.indexOf('manejaalcaldia')]).toUpperCase() === 'TRUE' || String(r[h.indexOf('manejaalcaldia')]) === 'SI',
          cierreEjercicio: String(r[h.indexOf('cierreejercicio')] || '31/12')
        });
      }
    }

    // 2. Obtener lista de usuarios para asignar
    const usuarios = [];
    if (sheetUsu && sheetUsu.getLastRow() >= 2) {
      const dataU = sheetUsu.getDataRange().getValues();
      const hU = dataU[0].map(x => String(x).trim().toLowerCase());
      const idxUser = hU.indexOf('username');
      const idxNom = hU.indexOf('nombre');
      const idxRol = hU.indexOf('role');
      const idxStat = hU.indexOf('status');

      for (let u = 1; u < dataU.length; u++) {
        const rU = dataU[u];
        const stat = String(rU[idxStat] || 'Activo').toLowerCase();
        if (stat === 'activo') {
          usuarios.push({
            username: String(rU[idxUser] || ''),
            nombre: String(rU[idxNom] || rU[idxUser] || ''),
            role: String(rU[idxRol] || 'Usuario')
          });
        }
      }
    }

    // 3. Obtener relaciones usuario-empresa
    const permisosMap = {}; // { idEmpresa: [ { username, rolEmpresa } ] }
    if (sheetPerm && sheetPerm.getLastRow() >= 2) {
      const dataP = sheetPerm.getDataRange().getValues();
      const hP = dataP[0].map(x => String(x).trim().toLowerCase());
      const idxPUser = hP.indexOf('username');
      const idxPEmp = hP.indexOf('idempresa');
      const idxPRol = hP.indexOf('rolempresa');

      for (let p = 1; p < dataP.length; p++) {
        const rowP = dataP[p];
        const empId = String(rowP[idxPEmp] || '').trim();
        const uName = String(rowP[idxPUser] || '').trim();
        const rEmp = String(rowP[idxPRol] || 'Usuario').trim();

        if (empId && uName) {
          if (!permisosMap[empId]) permisosMap[empId] = [];
          permisosMap[empId].push({ username: uName, rolEmpresa: rEmp });
        }
      }
    }

    // Unir asignaciones a cada empresa
    empresas.forEach(emp => {
      emp.usuariosAsignados = permisosMap[emp.idEmpresa] || [];
    });

    return {
      success: true,
      empresas: empresas,
      todosLosUsuarios: usuarios
    };

  } catch (err) {
    Logger.log('Error en apiObtenerDatosGestionEmpresas: ' + err);
    return { success: false, message: 'Error al consultar empresas: ' + err.message };
  }
}
function apiGuardarDatosEmpresa(token, datosEmpresa, usuariosAsignados) {
  try {
    const check = validarSessionToken(token);
    if (!check) return { success: false, message: 'Sesión inválida.' };

    const ssMaster = Contexto.getMasterSS();
    const sheetEmp = ssMaster.getSheetByName('Empresas') || ssMaster.getSheetByName('empresas');
    const dataEmp = sheetEmp.getDataRange().getValues();
    const h = dataEmp[0].map(x => String(x).trim().toLowerCase());

    const idxId = h.indexOf('idempresa');
    let filaEncontrada = -1;

    for (let i = 1; i < dataEmp.length; i++) {
      if (String(dataEmp[i][idxId]).trim() === String(datosEmpresa.idEmpresa).trim()) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { success: false, message: 'Empresa no encontrada.' };
    }

    // Actualizar columnas en la hoja Empresas
    const setCol = (nombreCol, valor) => {
      const c = h.indexOf(nombreCol.toLowerCase());
      if (c !== -1) sheetEmp.getRange(filaEncontrada, c + 1).setValue(valor);
    };

    setCol('nombreEmpresa', datosEmpresa.nombreEmpresa);
    setCol('rifEmpresa', datosEmpresa.rifEmpresa);
    setCol('direccionEmpresa', datosEmpresa.direccionEmpresa || '');
    setCol('telefonoEmpresa', datosEmpresa.telefonoEmpresa || '');
    setCol('correoEmpresa', datosEmpresa.correoEmpresa || '');
    setCol('status', datosEmpresa.status || 'Activo');
    setCol('logo', datosEmpresa.logo || '');
    setCol('porcentajeIVA', datosEmpresa.porcentajeIVA || 75);
    setCol('porcentajeISLR', datosEmpresa.porcentajeISLR || 2);
    setCol('manejaIGTF', datosEmpresa.manejaIGTF ? 'SI' : 'NO');
    setCol('manejaAlcaldia', datosEmpresa.manejaAlcaldia ? 'SI' : 'NO');
    setCol('cierreEjercicio', datosEmpresa.cierreEjercicio || '31/12');

    // Actualizar asignación de usuarios en PermisosUsuarioEmpresa
    actualizarPermisosEmpresa_(ssMaster, datosEmpresa.idEmpresa, usuariosAsignados);

    registrarAuditoria_('EDITAR_EMPRESA', check.sub, 'OK', 'Modificó: ' + datosEmpresa.nombreEmpresa);

    return { success: true, message: 'Configuración de empresa actualizada correctamente.' };

  } catch (err) {
    Logger.log('Error en apiGuardarDatosEmpresa: ' + err);
    return { success: false, message: 'Error al guardar: ' + err.message };
  }
}
function apiCrearEmpresaCompleta(token, datosEmpresa, usuariosAsignados) {
  try {
    const check = validarSessionToken(token);
    if (!check) return { success: false, message: 'Sesión inválida.' };

    // 1. Crear empresa clonando la plantilla base
    const resCreacion = crearEmpresaBackend_(datosEmpresa, check.sub);
    if (!resCreacion.success) return resCreacion;

    const idNuevaEmpresa = resCreacion.idEmpresa;
    const ssMaster = Contexto.getMasterSS();
    const sheetEmp = ssMaster.getSheetByName('Empresas') || ssMaster.getSheetByName('empresas');
    const dataEmp = sheetEmp.getDataRange().getValues();
    const h = dataEmp[0].map(x => String(x).trim().toLowerCase());

    // Encontrar fila recién agregada
    let fila = sheetEmp.getLastRow();
    const setCol = (nombreCol, valor) => {
      const c = h.indexOf(nombreCol.toLowerCase());
      if (c !== -1) sheetEmp.getRange(fila, c + 1).setValue(valor);
    };

    setCol('porcentajeIVA', datosEmpresa.porcentajeIVA || 75);
    setCol('porcentajeISLR', datosEmpresa.porcentajeISLR || 2);
    setCol('manejaIGTF', datosEmpresa.manejaIGTF ? 'SI' : 'NO');
    setCol('manejaAlcaldia', datosEmpresa.manejaAlcaldia ? 'SI' : 'NO');
    setCol('cierreEjercicio', datosEmpresa.cierreEjercicio || '31/12');

    // 2. Asignar los usuarios seleccionados
    actualizarPermisosEmpresa_(ssMaster, idNuevaEmpresa, usuariosAsignados);

    return {
      success: true,
      message: '¡Empresa "' + datosEmpresa.nombreEmpresa + '" creada con éxito!',
      idEmpresa: idNuevaEmpresa
    };

  } catch (err) {
    Logger.log('Error en apiCrearEmpresaCompleta: ' + err);
    return { success: false, message: 'Error: ' + err.message };
  }
}
function actualizarPermisosEmpresa_(ssMaster, idEmpresa, listaUsuarios) {
  const sheetPerm = ssMaster.getSheetByName('PermisosUsuarioEmpresa');
  if (!sheetPerm) return;

  const dataP = sheetPerm.getDataRange().getValues();
  if (dataP.length >= 2) {
    const h = dataP[0].map(x => String(x).trim().toLowerCase());
    const idxEmp = h.indexOf('idempresa');

    // Borrar permisos viejos de esta empresa de abajo hacia arriba
    for (let i = dataP.length - 1; i >= 1; i--) {
      if (String(dataP[i][idxEmp]).trim() === String(idEmpresa).trim()) {
        sheetPerm.deleteRow(i + 1);
      }
    }
  }

  // Insertar los nuevos permisos seleccionados
  if (Array.isArray(listaUsuarios)) {
    listaUsuarios.forEach(u => {
      sheetPerm.appendRow([
        'PERM_' + Utilities.getUuid().substring(0, 8),
        Utilities.getUuid(),
        u.username,
        idEmpresa,
        u.rolEmpresa || 'Usuario',
        false,
        new Date()
      ]);
    });
  }
}

/**
 * ============================================================================
 * CREAREMPRESA.GS
 * Módulo independiente para crear y dar de alta nuevas empresas automáticamente.
 * ============================================================================
 */

/**
 * Función pública invocable desde la WebApp o desde el Administrador
 * @param {string} token - Token de sesión del usuario logueado
 * @param {Object} datosEmpresa - { nombreEmpresa, rifEmpresa, direccion, telefono, correo, logo }
 */
function apiCrearEmpresa(token, datosEmpresa) {
  try {
    // 1. Validar que el usuario que ejecuta tenga permiso
    const check = validarSessionToken(token);
    if (!check) {
      return { success: false, message: 'Sesión inválida o expirada.' };
    }

    const rol = String(check.role || '').toLowerCase();
    if (rol !== 'admin' && rol !== 'owner') {
      return { success: false, message: 'Acceso denegado: solo Administradores pueden crear empresas.' };
    }

    // 2. Ejecutar la creación con el módulo Contexto
    return crearEmpresaBackend_(datosEmpresa, check.sub);

  } catch (err) {
    Logger.log('Error en apiCrearEmpresa: ' + err);
    return { success: false, message: 'Error al procesar la solicitud: ' + err.message };
  }
}

/**
 * Lógica interna de aprovisionamiento
 */
function crearEmpresaBackend_(datosEmpresa, usernameCreador) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30 seg espera máxima

    const nombre = String(datosEmpresa.nombreEmpresa || '').trim();
    const rif    = String(datosEmpresa.rifEmpresa || '').trim().toUpperCase();

    if (!nombre || !rif) {
      return { success: false, message: 'El nombre y el RIF de la empresa son obligatorios.' };
    }

    const props = PropertiesService.getScriptProperties();
    const masterId   = props.getProperty('MASTER_SPREADSHEET_ID');
    const templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID');

    if (!masterId || !templateId) {
      return { 
        success: false, 
        message: 'Configuración incompleta: falta MASTER_SPREADSHEET_ID o TEMPLATE_SPREADSHEET_ID en ScriptProperties.' 
      };
    }

    const ssMaster = SpreadsheetApp.openById(masterId);
    const sheetEmpresas = ssMaster.getSheetByName('Empresas') || ssMaster.getSheetByName('empresas');
    if (!sheetEmpresas) {
      return { success: false, message: 'Hoja "Empresas" no encontrada en GPRO_MAESTRO.' };
    }

    // 3. Verificar si el RIF ya está registrado
    const dataEmp = sheetEmpresas.getDataRange().getValues();
    if (dataEmp.length >= 2) {
      const headers = dataEmp[0].map(h => String(h).trim().toLowerCase());
      const idxRif = headers.indexOf('rifempresa');
      if (idxRif !== -1) {
        for (let i = 1; i < dataEmp.length; i++) {
          if (String(dataEmp[i][idxRif]).trim().toUpperCase() === rif) {
            return { success: false, message: 'Ya existe una empresa registrada con el RIF: ' + rif };
          }
        }
      }
    }

    // 4. Clonar la plantilla en Drive
    const archivoPlantilla = DriveApp.getFileById(templateId);
    const carpetaContenedora = archivoPlantilla.getParents().hasNext() 
      ? archivoPlantilla.getParents().next() 
      : DriveApp.getRootFolder();

    const nombreArchivoNuevo = 'GPRO_DATA - ' + nombre + ' (' + rif + ')';
    const nuevoArchivoDrive = archivoPlantilla.makeCopy(nombreArchivoNuevo, carpetaContenedora);
    const nuevoSpreadsheetId = nuevoArchivoDrive.getId();

    // 5. Generar ID único de empresa
    const nuevoIdEmpresa = 'EMP_' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // 6. Registrar en hoja Empresas del Maestro
    sheetEmpresas.appendRow([
      nuevoIdEmpresa,
      nombre,
      rif,
      datosEmpresa.direccionEmpresa || '',
      datosEmpresa.telefonoEmpresa || '',
      datosEmpresa.correoEmpresa || '',
      nuevoSpreadsheetId,
      'Activo',
      datosEmpresa.logo || '',
      new Date()
    ]);

    // 7. Asignar al creador como Administrador de la nueva empresa
    const sheetPermisos = ssMaster.getSheetByName('PermisosUsuarioEmpresa');
    if (sheetPermisos) {
      sheetPermisos.appendRow([
        'PERM_' + Utilities.getUuid().substring(0, 8),
        Utilities.getUuid(),
        usernameCreador || 'admin',
        nuevoIdEmpresa,
        'Admin',
        false, // No sobreescribir su predeterminada si ya tiene una
        new Date()
      ]);
    }

    // 8. Registrar en Log de Accesos
    try {
      registrarAuditoria_('CREAR_EMPRESA', usernameCreador, 'OK', 'Creada: ' + nombre + ' [' + nuevoIdEmpresa + ']');
    } catch(e) {}

    Logger.log('✅ Empresa creada exitosamente: ' + nombre + ' | ID: ' + nuevoSpreadsheetId);

    return {
      success: true,
      message: '¡Empresa "' + nombre + '" creada y activada correctamente!',
      idEmpresa: nuevoIdEmpresa,
      idSpreadsheet: nuevoSpreadsheetId,
      url: nuevoArchivoDrive.getUrl()
    };

  } catch (e) {
    Logger.log('❌ Error en crearEmpresaBackend_: ' + e);
    return { success: false, message: 'Error al crear empresa: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
function TEST_CREAR_EMPRESA_2() {
  const datosNuevaEmpresa = {
    nombreEmpresa: 'Comercializadora Beta, C.A.',
    rifEmpresa: 'J-98765432-1',
    direccionEmpresa: 'Valencia, Edo. Carabobo',
    telefonoEmpresa: '0424-5551234',
    correoEmpresa: 'contacto@betaca.com',
    logo: ''
  };

  // Ejecutamos con tu usuario administrador
  const resultado = crearEmpresaBackend_(datosNuevaEmpresa, 'fmorenocontabilidad@gmail.com');
  Logger.log(JSON.stringify(resultado, null, 2));
}


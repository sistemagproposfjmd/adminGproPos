// =========================================================================
// MÓDULO CENTRAL: Contexto.gs
// Gestión de Empresas, Conexión a Base de Datos y Sesión Activa
// =========================================================================

const ID_HOJA_MAESTRO = '1fFyzPgSjBnmnHmJrKAV300B5aShq2ah9BnhlrjeXr2k';
const ID_HOJA_PLANTILLA_DEFAULT = '1bu0s7SQlGHuLkV7nHb_QBIRnXGaZ9IG66bj92rEhUcY';

const Contexto = {
  getMasterSS: function() {
    try { return SpreadsheetApp.openById(ID_HOJA_MAESTRO); } 
    catch(e) { Logger.log("❌ Error al abrir GPRO_MAESTRO: " + e); return null; }
  },

  getEmpresaSS: function(idSpreadsheet) {
    if (!idSpreadsheet) idSpreadsheet = this.getEmpresaActivaId();
    if (!idSpreadsheet) return null;
    try { return SpreadsheetApp.openById(idSpreadsheet); } 
    catch(e) { Logger.log("⚠️ Error al abrir Spreadsheet de empresa: " + e); return null; }
  },

  /**
   * Guarda la empresa activa en las propiedades del usuario.
   * Nota: Como el script corre como "Yo", esto se guarda en la sesión del servidor.
   * El frontend maneja su propia sesión en sessionStorage.
   */
  fijarEmpresaActiva: function(idSpreadsheet, token) {
    try {
      if (!idSpreadsheet) return { success: false, error: 'ID de empresa no proporcionado' };
      
      // SEGURIDAD: Si viene un token, validar que el usuario tenga acceso a esta empresa
      if (token) {
        var sesion = validarSessionToken(token);
        if (!sesion) return { success: false, error: 'Sesión inválida' };
        
        var misEmpresas = this.obtenerEmpresasDeUsuario(sesion.sub).empresas || [];
        var autorizado = misEmpresas.some(function(e) { return e.idSpreadsheet === idSpreadsheet; });
        if (!autorizado) {
          registrarAuditoria_('SEGURIDAD', sesion.sub, 'DENEGADO', 'Intento de acceder a ' + idSpreadsheet);
          return { success: false, error: 'No tienes permisos para esta empresa.' };
        }
      }

      PropertiesService.getUserProperties().setProperty('EMPRESA_ACTIVA_ID', String(idSpreadsheet).trim());
      return { success: true, idSpreadsheet: idSpreadsheet };
    } catch(e) {
      return { success: false, error: e.toString() };
    }
  },

  getEmpresaActivaId: function() {
    try {
      var id = PropertiesService.getUserProperties().getProperty('EMPRESA_ACTIVA_ID');
      return id ? String(id).trim() : null;
    } catch(e) { return null; }
  },

  limpiarEmpresaActiva: function() {
    try {
      PropertiesService.getUserProperties().deleteProperty('EMPRESA_ACTIVA_ID');
      return { success: true };
    } catch(e) { return { success: false, error: e.toString() }; }
  },

  obtenerContextoEmpresa: function(idSpreadsheet) {
    try {
      var targetId = idSpreadsheet || this.getEmpresaActivaId();
      if (!targetId) return { success: false, error: 'No hay empresa seleccionada.' };

      var ssM = this.getMasterSS();
      if (!ssM) return { success: false, error: 'No se pudo conectar con GPRO_MAESTRO' };

      var shEmpresas = ssM.getSheetByName('Empresas');
      if (!shEmpresas) return { success: false, error: 'Hoja Empresas no encontrada' };

      var data = shEmpresas.getDataRange().getValues();
      var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
      var idxSS = headers.indexOf('idspreadsheet');

      if (idxSS === -1) return { success: false, error: 'Columna idSpreadsheet no existe en Empresas' };

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idxSS]).trim() === String(targetId).trim()) {
          var empObj = {};
          data[0].forEach(function(h, colIdx) { empObj[String(h).trim()] = data[i][colIdx]; });
          return { success: true, empresa: empObj };
        }
      }
      return { success: false, error: 'Empresa no encontrada en el catálogo.' };
    } catch(e) {
      return { success: false, error: e.toString() };
    }
  },

  /**
   * Obtiene la lista de empresas asignadas a un usuario.
   * Si el usuario es 'Owner' o 'Admin', garantiza acceso a todas las empresas activas.
   */
  obtenerEmpresasDeUsuario: function(username, role) {
    try {
      var ssM = this.getMasterSS();
      if (!ssM) return { success: false, error: 'No se pudo abrir GPRO_MAESTRO' };

      var shEmpresas = ssM.getSheetByName('Empresas');
      if (!shEmpresas) return { success: false, error: 'Hoja Empresas no encontrada' };

      var dataEmp = shEmpresas.getDataRange().getValues();
      var headEmp = dataEmp[0].map(function(h) { return String(h).trim().toLowerCase(); });
      var idxEmpId = headEmp.indexOf('idempresa');
      var idxEmpNom = headEmp.indexOf('nombreempresa');
      var idxEmpRif = headEmp.indexOf('rifempresa');
      var idxEmpSS = headEmp.indexOf('idspreadsheet');
      var idxEmpStatus = headEmp.indexOf('status');
      var idxEmpLogo = headEmp.indexOf('logo');
      var idxEmpDir = headEmp.indexOf('direccionempresa');
      var idxEmpTel = headEmp.indexOf('telefonoempresa');
      var idxEmpCor = headEmp.indexOf('correoempresa');

      var catalogoEmpresas = {};
      for (var e = 1; e < dataEmp.length; e++) {
        var rowE = dataEmp[e];
        var idEmp = String(rowE[idxEmpId] || '').trim();
        var status = String(rowE[idxEmpStatus] || 'Activo').trim().toLowerCase();
        if (idEmp && (status === 'activo' || status === 'true')) {
          catalogoEmpresas[idEmp] = {
            idEmpresa: idEmp,
            nombreEmpresa: String(rowE[idxEmpNom] || idEmp).trim(),
            rifEmpresa: idxEmpRif !== -1 ? String(rowE[idxEmpRif] || '').trim() : '',
            idSpreadsheet: idxEmpSS !== -1 ? String(rowE[idxEmpSS] || '').trim() : '',
            status: status,
            logo: idxEmpLogo !== -1 ? String(rowE[idxEmpLogo] || '').trim() : '',
            direccionEmpresa: idxEmpDir !== -1 ? String(rowE[idxEmpDir] || '').trim() : '',
            telefonoEmpresa: idxEmpTel !== -1 ? String(rowE[idxEmpTel] || '').trim() : '',
            correoEmpresa: idxEmpCor !== -1 ? String(rowE[idxEmpCor] || '').trim() : '',
            rolEmpresa: 'Usuario',
            esPredeterminada: false
          };
        }
      }

      // Verificar si es Owner o Admin
      var esAdminOwner = false;
      var r = String(role || '').trim().toLowerCase();
      if (r === 'owner' || r === 'admin' || r === 'superadmin') esAdminOwner = true;

      // Buscar permisos explícitos en PermisosUsuarioEmpresa
      var shPerm = ssM.getSheetByName('PermisosUsuarioEmpresa') || ssM.getSheetByName('PermisosUsuarioEpresas');
      var empresasAsignadas = [];
      var idsAsignadosMap = {};

      if (shPerm) {
        var dataP = shPerm.getDataRange().getValues();
        var headP = dataP[0].map(function(h) { return String(h).trim().toLowerCase(); });
        var idxPUser = headP.indexOf('username');
        var idxPEmp = headP.indexOf('idempresa');
        var idxPRol = headP.indexOf('rolempresa');
        var idxPPred = headP.indexOf('espredeterminada');

        var userLower = String(username || '').trim().toLowerCase();

        for (var p = 1; p < dataP.length; p++) {
          var rowP = dataP[p];
          var pUser = idxPUser !== -1 ? String(rowP[idxPUser] || '').trim().toLowerCase() : '';
          var pEmpId = idxPEmp !== -1 ? String(rowP[idxPEmp] || '').trim() : '';

          if (pUser === userLower && pEmpId) {
            if (catalogoEmpresas[pEmpId]) {
              var empObj = Object.assign({}, catalogoEmpresas[pEmpId]);
              empObj.rolEmpresa = idxPRol !== -1 ? String(rowP[idxPRol] || 'Admin').trim() : 'Admin';
              var predVal = idxPPred !== -1 ? String(rowP[idxPPred]).trim().toUpperCase() : 'FALSE';
              empObj.esPredeterminada = (predVal === 'TRUE' || predVal === '1');
              empresasAsignadas.push(empObj);
              idsAsignadosMap[pEmpId] = true;
            }
          }
        }
      }

      // Si es Admin/Owner, agregar las que le falten
      if (esAdminOwner) {
        Object.keys(catalogoEmpresas).forEach(function(k) {
          if (!idsAsignadosMap[k]) {
            var empOwner = Object.assign({}, catalogoEmpresas[k]);
            empOwner.rolEmpresa = 'Admin';
            empresasAsignadas.push(empOwner);
          }
        });
      }

      // Fallback: si no tiene ninguna y hay empresas, darle la primera
      if (empresasAsignadas.length === 0 && Object.keys(catalogoEmpresas).length > 0) {
        var primeraKey = Object.keys(catalogoEmpresas)[0];
        empresasAsignadas.push(catalogoEmpresas[primeraKey]);
      }

      return { success: true, empresas: empresasAsignadas };
    } catch(e) {
      return { success: false, error: e.toString(), empresas: [] };
    }
  },

  /**
   * Router inteligente de hojas (Maestro vs Operacional)
   */
  getSheet: function(name, idSpreadsheetEmpresa) {
    if (!name) return null;
    var nameLower = String(name).trim().toLowerCase();

    // Hojas del Sistema Central
    var hojasMaestro = ['usuarios','empresas','permisosusuarioempresa','permisosusuarioepresas','soporteweb','log_accesos','auditoria'];
    if (hojasMaestro.indexOf(nameLower) !== -1) {
      var ssM = this.getMasterSS();
      if (!ssM) return null;
      var shM = ssM.getSheetByName(name);
      if (shM) return shM;
      var allSheetsM = ssM.getSheets();
      for (var i = 0; i < allSheetsM.length; i++) {
        if (allSheetsM[i].getName().trim().toLowerCase() === nameLower) return allSheetsM[i];
      }
      return null;
    }

    // Hojas Operativas de la Empresa
    var ssEmp = this.getEmpresaSS(idSpreadsheetEmpresa);
    if (ssEmp) {
      var shEmp = ssEmp.getSheetByName(name);
      if (shEmp) return shEmp;
      var allSheetsEmp = ssEmp.getSheets();
      for (var j = 0; j < allSheetsEmp.length; j++) {
        if (allSheetsEmp[j].getName().trim().toLowerCase() === nameLower) return allSheetsEmp[j];
      }
    }

    // Fallback para tasas
    if (nameLower === 'tasas' || nameLower === 'resumen') {
      var ssMaster = this.getMasterSS();
      if (ssMaster) return ssMaster.getSheetByName(name) || ssMaster.getSheetByName('Tasas');
    }

    return null;
  }
};

// =========================================================================
// FUNCIONES GLOBALES EXPUESTAS (google.script.run)
// =========================================================================

function fijarEmpresaActiva(idSpreadsheet, token) {
  return Contexto.fijarEmpresaActiva(idSpreadsheet, token);
}

function getEmpresaActivaId() {
  return Contexto.getEmpresaActivaId();
}

function limpiarEmpresaActiva() {
  return Contexto.limpiarEmpresaActiva();
}

function obtenerEmpresasDeUsuario(username, role) {
  // Nota: Aquí idealmente deberíamos validar un token. 
  // Como el frontend llama esto al hacer login antes de tener token, lo dejamos público.
  return Contexto.obtenerEmpresasDeUsuario(username, role);
}

function obtenerContextoEmpresa(idSpreadsheet) {
  return Contexto.obtenerContextoEmpresa(idSpreadsheet);
}

function getSheetByName(name, idSpreadsheetEmpresa) {
  return Contexto.getSheet(name, idSpreadsheetEmpresa);
}

function _SS_(idSpreadsheetEmpresa) {
  return Contexto.getEmpresaSS(idSpreadsheetEmpresa);
}
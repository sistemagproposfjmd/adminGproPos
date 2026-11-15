/**
 * ============================================================================
 * MIGRACIONINICIAL.GS
 * Ejecuta una sola vez para crear Maestro, Plantilla y Empresa 1 automáticamente.
 * ============================================================================
 */

// ID de tu archivo actual original con todas las hojas
const ARCHIVO_ORIGINAL_ID = '1bu0s7SQlGHuLkV7nHb_QBIRnXGaZ9IG66bj92rEhUcY';

function EJECUTAR_MIGRACION_COMPLETA() {
  const ui = (typeof SpreadsheetApp.getUi === 'function') ? null : null;
  Logger.log('🚀 ============================================================');
  Logger.log('🚀 INICIANDO MIGRACIÓN AUTOMÁTICA GPRO MULTIEMPRESA');
  Logger.log('🚀 ============================================================');

  try {
    const ssOriginal = SpreadsheetApp.openById(ARCHIVO_ORIGINAL_ID);
    const driveOriginal = DriveApp.getFileById(ARCHIVO_ORIGINAL_ID);
    
    // 1. Crear carpeta contenedora en Google Drive
    const carpetaPadre = driveOriginal.getParents().hasNext() 
      ? driveOriginal.getParents().next() 
      : DriveApp.getRootFolder();
    
    const nombreCarpeta = 'SISTEMA_GPRO_MULTIEMPRESA';
    const iteradorCarpetas = carpetaPadre.getFoldersByName(nombreCarpeta);
    const carpetaDestino = iteradorCarpetas.hasNext() 
      ? iteradorCarpetas.next() 
      : carpetaPadre.createFolder(nombreCarpeta);

    Logger.log('📁 Carpeta de trabajo lista: ' + carpetaDestino.getName());

    // ========================================================================
    // PASO A: CREAR ARCHIVO GPRO_MAESTRO
    // ========================================================================
    Logger.log('\n🏛️ [PASO 1/4] Creando GPRO_MAESTRO...');
    const ssMaestro = SpreadsheetApp.create('GPRO_MAESTRO');
    const driveMaestro = DriveApp.getFileById(ssMaestro.getId());
    driveMaestro.moveTo(carpetaDestino);

    // Hojas maestras a copiar
    const hojasMaestrasOriginales = ['Usuarios', 'Log_Accesos', 'Tasas', 'manual_ES', 'manual_EN'];
    
    hojasMaestrasOriginales.forEach(nombre => {
      copiarHojaConDatos_(ssOriginal, ssMaestro, nombre);
    });

    // Crear/Adaptar hoja 'Empresas' en el Maestro
    let sheetEmpresas = ssMaestro.getSheetByName('Empresas');
    if (!sheetEmpresas) sheetEmpresas = ssMaestro.insertSheet('Empresas');
    sheetEmpresas.clear();
    sheetEmpresas.appendRow([
      'idEmpresa', 'nombreEmpresa', 'rifEmpresa', 'direccionEmpresa',
      'telefonoEmpresa', 'correoEmpresa', 'idSpreadsheet', 'status',
      'logo', 'fechaCreacion'
    ]);
    sheetEmpresas.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

    // Crear hoja 'PermisosUsuarioEmpresa' en el Maestro
    let sheetPermisos = ssMaestro.getSheetByName('PermisosUsuarioEmpresa');
    if (!sheetPermisos) sheetPermisos = ssMaestro.insertSheet('PermisosUsuarioEmpresa');
    sheetPermisos.clear();
    sheetPermisos.appendRow([
      'idPermiso', 'idUsuario', 'username', 'idEmpresa',
      'rolEmpresa', 'esPredeterminada', 'fechaAsignacion'
    ]);
    sheetPermisos.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

    // Eliminar la 'Hoja 1' por defecto creada en el nuevo archivo
    eliminarHojaPorDefecto_(ssMaestro);

    Logger.log('✅ GPRO_MAESTRO creado con éxito. ID: ' + ssMaestro.getId());

    // ========================================================================
    // PASO B: CREAR GPRO_PLANTILLA_BASE (Solo encabezados, sin datos viejos)
    // ========================================================================
    Logger.log('\n📋 [PASO 2/4] Creando GPRO_PLANTILLA_BASE (Estructura Limpia)...');
    const ssPlantilla = SpreadsheetApp.create('GPRO_PLANTILLA_BASE');
    const drivePlantilla = DriveApp.getFileById(ssPlantilla.getId());
    drivePlantilla.moveTo(carpetaDestino);

    // Lista oficial de hojas operativas de cada empresa
    const hojasOperativas = [
      'Clientes', 'Proveedores', 'Compras', 'Ventas', 'DetalleFacturas',
      'CobrosPagos', 'Abonos', 'Pedidos', 'DetallesPedidos',
      'datosBancarios', 'MovBancoCaja', 'detalleBancosCajas',
      'Productos', 'Producciones', 'RecetasHistorial',
      'Asientos', 'PlandeCuentas', 'Resumen'
    ];

    hojasOperativas.forEach(nombre => {
      copiarSoloEncabezados_(ssOriginal, ssPlantilla, nombre);
    });

    eliminarHojaPorDefecto_(ssPlantilla);
    Logger.log('✅ GPRO_PLANTILLA_BASE creada con éxito. ID: ' + ssPlantilla.getId());

    // ========================================================================
    // PASO C: CREAR BASE DE DATOS DE LA PRIMERA EMPRESA (EMPRESA 1 con datos)
    // ========================================================================
    Logger.log('\n🏢 [PASO 3/4] Creando base de datos para EMPRESA 1 con tus datos actuales...');
    
    // Obtener datos de la empresa actual desde la hoja 'empresas' original
    let nombreEmp1 = 'Empresa Matriz';
    let rifEmp1 = 'J-00000000-0';
    let dirEmp1 = '';
    let tlfEmp1 = '';
    let corEmp1 = '';
    let logoEmp1 = '';

    const sheetEmpOriginal = ssOriginal.getSheetByName('empresas') || ssOriginal.getSheetByName('Empresas');
    if (sheetEmpOriginal && sheetEmpOriginal.getLastRow() >= 2) {
      const datosEmp = sheetEmpOriginal.getRange(2, 1, 1, 12).getValues()[0];
      nombreEmp1 = String(datosEmp[1] || nombreEmp1).trim();
      rifEmp1 = String(datosEmp[2] || rifEmp1).trim();
      dirEmp1 = String(datosEmp[3] || '');
      tlfEmp1 = String(datosEmp[5] || '');
      corEmp1 = String(datosEmp[6] || '');
      logoEmp1 = String(datosEmp[11] || '');
    }

    const ssEmpresa1 = SpreadsheetApp.create('GPRO_DATA - ' + nombreEmp1);
    const driveEmpresa1 = DriveApp.getFileById(ssEmpresa1.getId());
    driveEmpresa1.moveTo(carpetaDestino);

    // Copiar hojas operativas CON sus datos actuales
    hojasOperativas.forEach(nombre => {
      copiarHojaConDatos_(ssOriginal, ssEmpresa1, nombre);
    });

    eliminarHojaPorDefecto_(ssEmpresa1);
    Logger.log('✅ GPRO_DATA - ' + nombreEmp1 + ' creado. ID: ' + ssEmpresa1.getId());

    // ========================================================================
    // PASO D: VINCULAR EMPRESA 1 Y USUARIOS EN GPRO_MAESTRO
    // ========================================================================
    Logger.log('\n🔗 [PASO 4/4] Vinculando empresa y usuarios en el Maestro...');
    
    const idEmpresa1 = 'EMP_001';
    sheetEmpresas.appendRow([
      idEmpresa1, nombreEmp1, rifEmp1, dirEmp1, tlfEmp1, corEmp1,
      ssEmpresa1.getId(), 'Activo', logoEmp1, new Date()
    ]);

    // Asignar todos los usuarios existentes a la Empresa 1
    const sheetUsuarios = ssMaestro.getSheetByName('Usuarios');
    if (sheetUsuarios && sheetUsuarios.getLastRow() >= 2) {
      const dataUsuarios = sheetUsuarios.getDataRange().getValues();
      const headersU = dataUsuarios[0].map(h => String(h).trim().toLowerCase());
      const idxIdU   = headersU.indexOf('idusuario');
      const idxUserU = headersU.indexOf('username');
      const idxRolU  = headersU.indexOf('role');

      for (let u = 1; u < dataUsuarios.length; u++) {
        const filaU = dataUsuarios[u];
        const idU = idxIdU !== -1 ? filaU[idxIdU] : Utilities.getUuid();
        const usernameU = idxUserU !== -1 ? filaU[idxUserU] : '';
        const rolU = idxRolU !== -1 ? (filaU[idxRolU] || 'Usuario') : 'Usuario';

        if (usernameU) {
          sheetPermisos.appendRow([
            'PERM_' + Utilities.getUuid().substring(0, 8),
            idU,
            usernameU,
            idEmpresa1,
            rolU,
            true, // esPredeterminada
            new Date()
          ]);
        }
      }
    }

    // ========================================================================
    // GUARDAR AUTOMÁTICAMENTE EN LAS PROPIEDADES DEL SCRIPT
    // ========================================================================
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('MASTER_SPREADSHEET_ID', ssMaestro.getId());
    scriptProps.setProperty('TEMPLATE_SPREADSHEET_ID', ssPlantilla.getId());

    Logger.log('\n🎉 ============================================================');
    Logger.log('🎉 ¡MIGRACIÓN COMPLETADA EXITOSAMENTE!');
    Logger.log('🎉 ============================================================');
    Logger.log('📊 GPRO_MAESTRO ID:        ' + ssMaestro.getId());
    Logger.log('📋 GPRO_PLANTILLA_BASE ID:  ' + ssPlantilla.getId());
    Logger.log('🏢 EMPRESA 1 DATA ID:      ' + ssEmpresa1.getId());
    Logger.log('🔗 URL Carpeta en Drive:   ' + carpetaDestino.getUrl());
    Logger.log('⚙️ Propiedades del Script configuradas automáticamente.');
    Logger.log('============================================================');

  } catch (error) {
    Logger.log('❌ ERROR FATAL EN MIGRACIÓN: ' + error.message);
    Logger.log(error.stack);
  }
}

// ============================================================================
// FUNCIONES AUXILIARES INTERNAS DE COPIA Y DEPURACIÓN
// ============================================================================
function copiarHojaConDatos_(ssOrigen, ssDestino, nombreHoja) {
  const hojaOrigen = ssOrigen.getSheetByName(nombreHoja);
  if (!hojaOrigen) {
    Logger.log('   ⚠️ Hoja "' + nombreHoja + '" no encontrada en origen, omitiendo.');
    return;
  }
  const nuevaHoja = hojaOrigen.copyTo(ssDestino);
  nuevaHoja.setName(nombreHoja);
  Logger.log('   ✓ Hoja copiada con datos: ' + nombreHoja);
}

function copiarSoloEncabezados_(ssOrigen, ssDestino, nombreHoja) {
  const hojaOrigen = ssOrigen.getSheetByName(nombreHoja);
  if (!hojaOrigen) {
    Logger.log('   ⚠️ Hoja "' + nombreHoja + '" no encontrada en origen, omitiendo.');
    return;
  }

  // 1. Copiar la hoja completa al nuevo archivo (esto preserva anchos, colores y formatos)
  const nuevaHoja = hojaOrigen.copyTo(ssDestino);
  nuevaHoja.setName(nombreHoja);

  // 2. Si tiene más de 1 fila (es decir, tiene datos), borrar las filas desde la 2 en adelante
  const ultFila = nuevaHoja.getLastRow();
  if (ultFila > 1) {
    // Si la hoja 'Resumen' tiene fórmulas de dashboard que deseas conservar, no se limpian
    if (nombreHoja !== 'Resumen') {
      nuevaHoja.deleteRows(2, ultFila - 1);
    }
  }

  Logger.log('   ✓ Plantilla limpia creada (con formatos y anchos exactos): ' + nombreHoja);
}

function eliminarHojaPorDefecto_(ss) {
  try {
    const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
    if (hoja1 && ss.getSheets().length > 1) {
      ss.deleteSheet(hoja1);
    }
  } catch(e) {}
}

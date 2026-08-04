/**
 * ============================================================================
 * PALM DIAMANTE CRM 2.0 - SERVIDOR CENTRAL (2026)
 * ============================================================================
 * Gestión Integrada:
 * - Prospectos Multi-Desarrollo (Meta Ads / Ventas)
 * - Clientes Activos y Cobranza Palm Diamante
 * - Integración Google Drive / Google Calendar
 * - Historial y Validaciones Manuales de WhatsApp
 */

// --- CONFIGURACIÓN GLOBAL DE ID's Y NOMBRES DE HOJAS ---
const CONFIG = {
  HOJA_PROSPECTOS: 'PROSPECTOS_MULTIDESARROLLO',
  HOJA_PALM: 'CLIENTES_PALM',
  HOJA_CONVENIOS: 'CONVENIOS',
  HOJA_HISTORIAL: 'Historial',
  ID_CARPETA_DRIVE: '1CSmNqlQ9eSdWHIcKqQRBu_aMDBJ0PPAj' // Actualizar si creaste una carpeta nueva en Drive
};

/**
 * Renderizado de la Aplicación Web
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Palm Diamante CRM 2.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Carga Inicial Unificada para la Interfaz (Dashboard / Tablas)
 */
function obtenerDatosCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Obtener o Inicializar Prospectos Multi-Desarrollo
  let sheetProspectos = ss.getSheetByName(CONFIG.HOJA_PROSPECTOS);
  let prospectos = [];
  
  if (sheetProspectos && sheetProspectos.getLastRow() > 1) {
    let dataP = sheetProspectos.getDataRange().getValues();
    for (let i = 1; i < dataP.length; i++) {
      if (dataP[i][0]) { // Verificar que exista nombre
        prospectos.push({
          id: i,
          nombre: dataP[i][0] || '',
          telefono: String(dataP[i][1] || ''),
          desarrollo: dataP[i][2] || 'Palm Diamante',
          origen: dataP[i][3] || 'Meta Ads',
          estatus: dataP[i][4] || 'Primer Contacto',
          fecha: dataP[i][5] ? formatearFecha_(dataP[i][5]) : ''
        });
      }
    }
  } else {
    // Datos de prueba inicial si la hoja es totalmente nueva
    prospectos = [
      { id: 1, nombre: 'Daniel Acevedo', telefono: '5543218765', desarrollo: 'Palm Diamante', origen: 'Meta Ads', estatus: 'Venta Realizada - Pendiente Contable', fecha: '04/08/2026' },
      { id: 2, nombre: 'Sofía Sánchez', telefono: '5598765432', desarrollo: 'La Marquesa', origen: 'Instagram', estatus: 'Cotización Enviada', fecha: '03/08/2026' },
      { id: 3, nombre: 'Diego Hernández', telefono: '7441112233', desarrollo: 'Costa Azul', origen: 'Directo WA', estatus: 'Primer Contacto', fecha: '02/08/2026' }
    ];
  }

  // 2. Obtener o Inicializar Clientes Activos Palm Diamante
  let sheetPalm = ss.getSheetByName(CONFIG.HOJA_PALM) || ss.getSheetByName('CONVENIOS');
  let clientesPalm = [];
  
  if (sheetPalm && sheetPalm.getLastRow() > 1) {
    let dataC = sheetPalm.getDataRange().getValues();
    for (let i = 1; i < dataC.length; i++) {
      if (dataC[i][0] || dataC[i][1]) {
        clientesPalm.push({
          id: i,
          depto: dataC[i][0] || 'N/A',
          nombre: dataC[i][1] || '',
          telefono: String(dataC[i][2] || ''),
          saldo: dataC[i][3] ? '$' + Number(dataC[i][3]).toLocaleString('es-MX', {minimumFractionDigits: 2}) : '$0.00',
          estatus: dataC[i][4] || 'Al Día',
          proximoPago: dataC[i][5] ? formatearFecha_(dataC[i][5]) : 'N/A'
        });
      }
    }
  } else {
    clientesPalm = [
      { id: 1, depto: '1A - 1204', nombre: 'Carlos Mendoza', telefono: '5512345678', saldo: '$15,000.00', estatus: 'Al Día', proximoPago: '15/08/2026' },
      { id: 2, depto: '2B - 302', nombre: 'María Rosa Flores', telefono: '5587654321', saldo: '$32,500.00', estatus: 'Vencido', proximoPago: '01/08/2026' }
    ];
  }

  return { prospectos: prospectos, clientesPalm: clientesPalm };
}

/**
 * MARCA LA VENTA Y COLOCA AL CLIENTE EN ESPERA DE REPORTE MENSUAL DE DIRECCIÓN
 */
function marcarVentaPendienteContable(nombreCliente, desarrollo, depto, precio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetP = ss.getSheetByName(CONFIG.HOJA_PROSPECTOS);
  
  // Actualizar estatus en la hoja si existe
  if (sheetP) {
    let data = sheetP.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nombreCliente) {
        sheetP.getRange(i + 1, 5).setValue('Venta Realizada - Pendiente Contable');
        break;
      }
    }
  }

  // Registrar en el historial de operaciones
  registrarEnHistorial_(nombreCliente, 'Venta Marcada (' + desarrollo + ')', 'En espera de sábana contable de Dirección. Depto: ' + depto + ' | Monto: $' + precio);
  
  return { status: 'success', message: 'Cliente movido a sala de espera contable correctamente.' };
}

/**
 * REGISTRO OBLIGATORIO DE GESTIÓN HUMANA (SOLO TRAS ABRIR/REVISAR WHATSAPP)
 */
function registrarAtendidoManual(clienteNombre, tipoModulo, notaAsesor) {
  registrarEnHistorial_(clienteNombre, 'Seguimiento Atendido (' + tipoModulo + ')', notaAsesor);
  return { status: 'success', message: 'Atención registrada en el historial.' };
}

/**
 * INTEGRACIÓN GOOGLE DRIVE: CREAR CARPETA DE EXPEDIENTE POR CLIENTE
 */
function crearExpedienteDrive(nombreCliente, depto) {
  try {
    const parentFolder = DriveApp.getFolderById(CONFIG.ID_CARPETA_DRIVE);
    const folderName = depto + ' - ' + nombreCliente;
    const newFolder = parentFolder.createFolder(folderName);
    
    registrarEnHistorial_(nombreCliente, 'Carpeta Drive Creada', newFolder.getUrl());
    return { status: 'success', url: newFolder.getUrl() };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// ============================================================================
// FUNCIONES AUXILIARES PRIVADAS
// ============================================================================

function registrarEnHistorial_(cliente, accion, detalle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.HOJA_HISTORIAL);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.HOJA_HISTORIAL);
    sheet.appendRow(['Fecha/Hora', 'Cliente / ID', 'Acción', 'Detalle / Notas']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1A2530').setFontColor('#FFFFFF');
  }
  
  const fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  sheet.appendRow([fechaStr, cliente, accion, detalle]);
}

function formatearFecha_(fecha) {
  if (!fecha) return '';
  try {
    return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (e) {
    return String(fecha);
  }
}
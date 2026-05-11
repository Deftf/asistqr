// ════════════════════════════════════════════════
//  ASISTQR — Google Apps Script
//  Pega este código en Extensiones > Apps Script
//  Luego publica como Web App (acceso: Cualquiera)
// ════════════════════════════════════════════════

// ── RECIBE LOS ESCANEOS DESDE LA APP ──
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Asistencia");
    
    // Crear la hoja si no existe
    if (!sheet) {
      sheet = ss.insertSheet("Asistencia");
      // Encabezados
      sheet.getRange(1, 1, 1, 6).setValues([[
        "Fecha", "Hora", "Alumno / Código QR", "Sección", "Estado", "Timestamp"
      ]]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1a2535").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
    
    // Parsear datos recibidos
    var data = JSON.parse(e.postData.contents);
    var ahora = new Date();
    
    // Agregar fila
    sheet.appendRow([
      Utilities.formatDate(ahora, "America/Lima", "dd/MM/yyyy"),
      Utilities.formatDate(ahora, "America/Lima", "HH:mm:ss"),
      data.alumno || data.codigo || "Sin nombre",
      data.seccion || "Sin sección",
      "Presente",
      data.timestamp || ahora.toISOString()
    ]);
    
    // Colorear la fila nueva (verde claro)
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 5).setBackground("#d4edda").setFontColor("#155724");
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", fila: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── PERMITE VERIFICAR QUE EL SCRIPT ESTÁ ACTIVO ──
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      status: "activo", 
      mensaje: "AsistQR webhook funcionando correctamente",
      fecha: new Date().toLocaleDateString("es-PE")
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GENERA LA HOJA DE ALUMNOS SI NO EXISTE ──
function setupHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Hoja: Alumnos
  var alumnos = ss.getSheetByName("Alumnos");
  if (!alumnos) {
    alumnos = ss.insertSheet("Alumnos");
    alumnos.getRange(1, 1, 1, 5).setValues([[
      "N°", "Apellidos y Nombres", "DNI", "Sección", "Código QR (usar como contenido del QR)"
    ]]);
    alumnos.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#0f1923").setFontColor("#ffffff");
    alumnos.setFrozenRows(1);
    alumnos.setColumnWidth(2, 250);
    alumnos.setColumnWidth(5, 300);
    
    // Ejemplo de alumno
    alumnos.getRange(2, 1, 1, 5).setValues([[
      1, "García López, Ana María", "74123456", "3° A", "ANA GARCIA - 74123456 - 3A"
    ]]);
  }
  
  // Hoja: Asistencia
  var asistencia = ss.getSheetByName("Asistencia");
  if (!asistencia) {
    asistencia = ss.insertSheet("Asistencia");
    asistencia.getRange(1, 1, 1, 6).setValues([[
      "Fecha", "Hora", "Alumno / Código QR", "Sección", "Estado", "Timestamp"
    ]]);
    asistencia.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1a2535").setFontColor("#ffffff");
    asistencia.setFrozenRows(1);
  }
  
  SpreadsheetApp.getUi().alert(
    "✓ Hoja configurada correctamente.\n\n" +
    "Ahora:\n" +
    "1. Llena la hoja 'Alumnos' con tu lista.\n" +
    "2. Publica este script: Implementar > Nueva implementación > Aplicación web.\n" +
    "3. Acceso: 'Cualquier usuario'.\n" +
    "4. Copia la URL y pégala en la app AsistQR."
  );
}

// ── MENÚ PERSONALIZADO EN EL SHEETS ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 AsistQR")
    .addItem("Configurar hojas", "setupHoja")
    .addItem("Ver resumen de hoy", "verResumenHoy")
    .addToUi();
}

// ── RESUMEN DEL DÍA ──
function verResumenHoy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Asistencia");
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("No hay registros de asistencia aún.");
    return;
  }
  
  var hoy = Utilities.formatDate(new Date(), "America/Lima", "dd/MM/yyyy");
  var data = sheet.getDataRange().getValues();
  var registrosHoy = data.filter((row, i) => i > 0 && row[0] === hoy);
  
  var msg = "📅 Asistencia del " + hoy + "\n";
  msg += "Total presentes: " + registrosHoy.length + "\n\n";
  
  // Agrupar por sección
  var porSeccion = {};
  registrosHoy.forEach(row => {
    var sec = row[3] || "Sin sección";
    if (!porSeccion[sec]) porSeccion[sec] = 0;
    porSeccion[sec]++;
  });
  
  for (var sec in porSeccion) {
    msg += "• " + sec + ": " + porSeccion[sec] + " alumnos\n";
  }
  
  SpreadsheetApp.getUi().alert(msg);
}

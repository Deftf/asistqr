// ════════════════════════════════════════════════
//  ASISTQR — Google Apps Script v2
//  Pega en: Extensiones > Apps Script
//  Publica: Implementar > Nueva implementación >
//           Aplicación web > Acceso: Cualquiera
// ════════════════════════════════════════════════

/* ─────────────────────────────────────────
   CONFIGURACIÓN
───────────────────────────────────────── */
const HOJA_ALUMNOS    = "Alumnos";
const HOJA_ASISTENCIA = "Asistencia";
const ZONA_HORARIA    = "America/Lima";

/* ─────────────────────────────────────────
   doGET  →  la app consulta secciones
───────────────────────────────────────── */
function doGet(e) {
  var action = e.parameter.action || "";

  if (action === "getSecciones") {
    return getSecciones();
  }

  // Verificación de que el script está activo
  return jsonOut({ status: "ok", mensaje: "AsistQR activo", fecha: hoy() });
}

function getSecciones() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_ALUMNOS);

    if (!hoja || hoja.getLastRow() < 2) {
      return jsonOut({ secciones: [], alumnos: [], mensaje: "Hoja Alumnos vacía o inexistente" });
    }

    // Columnas: A=N°, B=Nombres, C=DNI, D=Grado y Sección
    var rows = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();

    var secciones = [];
    var alumnos   = [];

    rows.forEach(function(r) {
      var nombre  = String(r[1]).trim();
      var dni     = String(r[2]).trim();
      var seccion = String(r[3]).trim();
      if (!nombre || !seccion) return;

      // El código QR que generó la fórmula es: Nombre|DNI
      var codigo = nombre + "|" + dni;

      alumnos.push({ codigo: codigo, nombre: nombre, dni: dni, grado_seccion: seccion });

      if (secciones.indexOf(seccion) === -1) secciones.push(seccion);
    });

    secciones.sort();

    return jsonOut({ secciones: secciones, alumnos: alumnos });

  } catch (err) {
    return jsonOut({ secciones: [], alumnos: [], error: err.message });
  }
}

/* ─────────────────────────────────────────
   doPOST  →  la app registra un escaneo
───────────────────────────────────────── */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === "registrar") {
      return registrarAsistencia(data);
    }

    return jsonOut({ status: "error", mensaje: "Acción desconocida" });

  } catch (err) {
    return jsonOut({ status: "error", mensaje: err.message });
  }
}

function registrarAsistencia(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(HOJA_ASISTENCIA);
  var ahora = new Date();

  hoja.appendRow([
    Utilities.formatDate(ahora, ZONA_HORARIA, "dd/MM/yyyy"),  // Fecha
    Utilities.formatDate(ahora, ZONA_HORARIA, "HH:mm:ss"),   // Hora
    data.nombre       || data.codigo || "—",                  // Apellidos y Nombres
    data.dni          || "—",                                 // DNI
    data.grado_seccion || "—",                               // Grado y Sección
    "Presente"                                                // Estado
  ]);

  // Colorear fila: verde claro
  var lastRow = hoja.getLastRow();
  hoja.getRange(lastRow, 6).setBackground("#c8e6c9").setFontColor("#1b5e20");
  // Zebra alternado en las demás columnas
  if (lastRow % 2 === 0) {
    hoja.getRange(lastRow, 1, 1, 5).setBackground("#fafafa");
  }

  return jsonOut({ status: "ok", fila: lastRow });
}

/* ─────────────────────────────────────────
   SETUP AUTOMÁTICO DE HOJAS
   Ejecuta manualmente o desde el menú
───────────────────────────────────────── */
function setupHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  crearHojaAlumnos(ss);
  crearHojaAsistencia(ss);

  SpreadsheetApp.getUi().alert(
    "✓ Hojas creadas correctamente.\n\n" +
    "Próximos pasos:\n" +
    "1. Ve a la hoja 'Alumnos' y llena la lista.\n" +
    "   • La columna 'Código QR' usa la fórmula de Google Charts\n" +
    "     (instrucciones en la celda E2).\n\n" +
    "2. Publica el script como aplicación web:\n" +
    "   Implementar > Nueva implementación\n" +
    "   Tipo: Aplicación web\n" +
    "   Ejecutar como: Yo\n" +
    "   Acceso: Cualquier usuario\n\n" +
    "3. Copia la URL y pégala en la app AsistQR."
  );
}

function crearHojaAlumnos(ss) {
  var existe = ss.getSheetByName(HOJA_ALUMNOS);
  if (existe) {
    // Ya existe, no recrear para no perder datos
    SpreadsheetApp.getUi().alert("⚠ La hoja 'Alumnos' ya existe. No se modificó para proteger tus datos.");
    return;
  }

  var h = ss.insertSheet(HOJA_ALUMNOS);

  // Encabezados
  var headers = ["N°", "Apellidos y Nombres", "DNI", "Grado y Sección", "Código QR"];
  h.getRange(1, 1, 1, 5).setValues([headers]);

  // Estilo del encabezado
  var header = h.getRange(1, 1, 1, 5);
  header.setBackground("#0a0a0a")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setFontFamily("Arial")
        .setFontSize(11)
        .setHorizontalAlignment("center");
  h.setFrozenRows(1);

  // Fila de ejemplo con fórmula QR
  // El QR codifica: "NOMBRE COMPLETO|DNI"  →  la app lo separa
  h.getRange(2, 1).setValue(1);
  h.getRange(2, 2).setValue("García López, Ana María");
  h.getRange(2, 3).setValue("74123456");
  h.getRange(2, 4).setValue("3° A");
  // Fórmula QR: genera imagen desde Google Charts API
  // Codifica: Nombre|DNI  (lo que leerá el escáner)
  h.getRange(2, 5).setFormula(
    '=IMAGE("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl="&ENCODEURL(B2&"|"&C2))'
  );

  // Nota explicativa para las siguientes filas (en la celda F2)
  h.getRange(2, 6).setValue(
    "⬅ Fórmula QR: =IMAGE(\"https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=\"&ENCODEURL(B2&\"|\"&C2))\n" +
    "Copia esta fórmula en E3, E4, E5... cambiando B2→B3, C2→C3, etc.\n" +
    "O arrastra la celda E2 hacia abajo para aplicarla a todos."
  );
  h.getRange(2, 6).setFontColor("#888888").setFontSize(9).setWrap(true);

  // Ancho de columnas
  h.setColumnWidth(1, 50);
  h.setColumnWidth(2, 240);
  h.setColumnWidth(3, 110);
  h.setColumnWidth(4, 130);
  h.setColumnWidth(5, 180);
  h.setColumnWidth(6, 350);
  h.setRowHeight(2, 150); // altura para mostrar el QR

  // Borde inferior del encabezado
  header.setBorder(null, null, true, null, null, null, "#ffffff", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  Logger.log("Hoja Alumnos creada correctamente.");
}

function crearHojaAsistencia(ss) {
  var existe = ss.getSheetByName(HOJA_ASISTENCIA);
  if (existe) return; // No recrear si ya tiene datos

  var h = ss.insertSheet(HOJA_ASISTENCIA);

  var headers = ["Fecha", "Hora", "Apellidos y Nombres", "DNI", "Grado y Sección", "Estado"];
  h.getRange(1, 1, 1, 6).setValues([headers]);

  var header = h.getRange(1, 1, 1, 6);
  header.setBackground("#0a0a0a")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
  h.setFrozenRows(1);

  // Ancho de columnas
  h.setColumnWidth(1, 110);
  h.setColumnWidth(2, 90);
  h.setColumnWidth(3, 240);
  h.setColumnWidth(4, 110);
  h.setColumnWidth(5, 130);
  h.setColumnWidth(6, 100);

  // Filtros habilitados desde el inicio
  h.getRange(1, 1, 1, 6).createFilter();

  Logger.log("Hoja Asistencia creada correctamente.");
}

/* ─────────────────────────────────────────
   RESUMEN DEL DÍA (menú)
───────────────────────────────────────── */
function verResumenHoy() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_ASISTENCIA);

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("No hay registros de asistencia aún.");
    return;
  }

  var hoyStr = Utilities.formatDate(new Date(), ZONA_HORARIA, "dd/MM/yyyy");
  var data   = sheet.getDataRange().getValues();
  var hoy    = data.filter((r, i) => i > 0 && r[0] === hoyStr);

  var msg = "📅 Asistencia del " + hoyStr + "\n";
  msg += "Total registros: " + hoy.length + "\n\n";

  var porSeccion = {};
  hoy.forEach(r => {
    var sec = r[4] || "Sin sección";
    porSeccion[sec] = (porSeccion[sec] || 0) + 1;
  });

  Object.keys(porSeccion).sort().forEach(sec => {
    msg += "• " + sec + ": " + porSeccion[sec] + " alumno(s)\n";
  });

  SpreadsheetApp.getUi().alert(msg);
}

/* ─────────────────────────────────────────
   MENÚ PERSONALIZADO
───────────────────────────────────────── */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 AsistQR")
    .addItem("1. Configurar hojas (primera vez)", "setupHojas")
    .addSeparator()
    .addItem("Ver resumen de hoy", "verResumenHoy")
    .addToUi();
}

/* ─────────────────────────────────────────
   UTILIDADES
───────────────────────────────────────── */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hoy() {
  return Utilities.formatDate(new Date(), ZONA_HORARIA, "dd/MM/yyyy");
}

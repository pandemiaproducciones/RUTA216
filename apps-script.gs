// RUTA216 v2.0 - SCRIPT DE CONEXIÓN COMPLETO
// Hoja "usuarios":
//   A usuario | B contraseña | C nombre_real | D instrumento (C, Bb, Eb o F)
//   E nivel actual | F escala actual | G escalas hechas (0 a 216)

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Espera 10 segundos si el script está ocupado

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("usuarios");
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var rows = sheet.getDataRange().getValues();

    // ACCIÓN: LOGIN (igual que antes)
    if (action === "login") {
      for (var i = 1; i < rows.length; i++) {
        // Validación exacta de mayúsculas y minúsculas
        if (rows[i][0].toString() === data.usuario && rows[i][1].toString() === data.contrasena) {
          return res({
            result: "success",
            nombre_real: rows[i][2],
            tonalidad: rows[i][3],
            nivel_ritmo: rows[i][4],
            escala_actual: rows[i][5],
            ejercicio_actual: rows[i][6]
          });
        }
      }
      return res({result: "error", message: "Usuario o contraseña incorrectos"});
    }

    // ACCIÓN: LEER EL AVANCE DE UN ESTUDIANTE
    if (action === "progreso") {
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString() === data.usuario) {
          return res({
            result: "success",
            instrumento: rows[i][3],
            progreso: Number(rows[i][6]) || 0
          });
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: GUARDAR EL AVANCE (lo manda la ventana de estudio)
    if (action === "guardarProgreso") {
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString() === data.usuario) {
          var actual = Number(rows[i][6]) || 0;
          var nuevo = Number(data.progreso) || 0;
          if (data.instrumento) sheet.getRange(i + 1, 4).setValue(data.instrumento);
          // nunca se retrocede: si la hoja tiene más avance, se deja el de la hoja
          if (nuevo >= actual) {
            sheet.getRange(i + 1, 5).setValue(data.nivel || "");
            sheet.getRange(i + 1, 6).setValue(data.escala || "");
            sheet.getRange(i + 1, 7).setValue(nuevo);
          }
          return res({result: "success", progreso: Math.max(nuevo, actual)});
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    return res({result: "error", message: "Acción desconocida"});

  } catch (error) {
    return res({result: "error", message: error.toString()});
  } finally {
    lock.releaseLock();
  }
}

// Función auxiliar para enviar respuestas correctamente
function res(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Para comprobar que el script está activo (abrir la URL en el navegador)
function doGet(e) {
  return ContentService.createTextOutput("RUTA216 activo").setMimeType(ContentService.MimeType.TEXT);
}

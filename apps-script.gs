// RUTA216 v4.0 - SCRIPT DE CONEXIÓN COMPLETO
// Hoja "usuarios":
//   A usuario | B contraseña | C nombre | D instrumento (C, Bb, Eb o F)
//   E nivel | F escala | G escalas_hechas (0 a 216) | H detalle (lo llena la app)
//   I nombre completo | J instrumento real | K grado   (solo referencia)
//   L última escala (fecha) | M semana desde | N escalas esta semana   (lo llena la app)
// La semana empieza el lunes a las 0:00 (hora de Colombia).

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("usuarios");
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var rows = sheet.getDataRange().getValues();

    // ACCIÓN: LOGIN
    if (action === "login") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario) && igual(rows[i][1], data.contrasena)) {
          return res({
            result: "success",
            usuario: String(rows[i][0]),
            nombre_real: rows[i][2],
            instrumento: afinacion(rows[i][3]),
            progreso: Number(rows[i][6]) || 0,
            hechas: String(rows[i][7] || "")
          });
        }
      }
      return res({result: "error", message: "Usuario o contraseña incorrectos"});
    }

    // ACCIÓN: LEER EL AVANCE
    if (action === "progreso") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          return res({
            result: "success",
            instrumento: afinacion(rows[i][3]),
            progreso: Number(rows[i][6]) || 0,
            hechas: String(rows[i][7] || "")
          });
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: GUARDAR EL AVANCE (nunca borra escalas ya hechas)
    if (action === "guardarProgreso") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          var enHoja = aLista(rows[i][7], rows[i][6]);
          var nuevas = aLista(data.hechas, data.progreso);
          var total = 0, recien = 0;
          for (var k = 0; k < 216; k++) {
            if (nuevas[k] && !enHoja[k]) { enHoja[k] = true; recien++; }
            if (enHoja[k]) total++;
          }
          if (data.instrumento) sheet.getRange(i + 1, 4).setValue(data.instrumento);
          sheet.getRange(i + 1, 5, 1, 4).setValues([[
            data.nivel || "",
            data.escala || "",
            total,
            "'" + aTexto(enHoja)
          ]]);
          // Ranking semanal: cuenta las escalas nuevas de esta semana
          if (recien > 0) {
            var semana = inicioSemana();
            var deEstaSemana = String(rows[i][12] || "") === semana ? (Number(rows[i][13]) || 0) : 0;
            sheet.getRange(i + 1, 12, 1, 3).setValues([[
              "'" + Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd HH:mm"),
              "'" + semana,
              deEstaSemana + recien
            ]]);
          }
          return res({result: "success", progreso: total, hechas: aTexto(enHoja)});
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: RANKING SEMANAL
    if (action === "ranking") {
      var semana = inicioSemana();
      var lista = [];
      for (var i = 1; i < rows.length; i++) {
        var usuario = String(rows[i][0] || "");
        if (usuario === "" || limpiar(usuario) === "paisa") continue;   // el profesor no compite
        var n = String(rows[i][12] || "") === semana ? (Number(rows[i][13]) || 0) : 0;
        lista.push({
          nombre: String(rows[i][2] || usuario),
          semana: n,
          total: Number(rows[i][6]) || 0,
          yo: igual(usuario, data.usuario)
        });
      }
      // más escalas esta semana primero; si empatan, el que lleva más en total
      lista.sort(function (a, b) { return (b.semana - a.semana) || (b.total - a.total); });
      return res({result: "success", semana: semana, lista: lista});
    }

    return res({result: "error", message: "Acción desconocida"});

  } catch (error) {
    return res({result: "error", message: error.toString()});
  } finally {
    lock.releaseLock();
  }
}

var ZONA = "America/Bogota";

// Fecha del lunes de esta semana (ej. "2026-11-02")
function inicioSemana() {
  var ahora = new Date();
  var diaSemana = Number(Utilities.formatDate(ahora, ZONA, "u")); // 1 = lunes ... 7 = domingo
  var lunes = new Date(ahora.getTime() - (diaSemana - 1) * 24 * 60 * 60 * 1000);
  return Utilities.formatDate(lunes, ZONA, "yyyy-MM-dd");
}

// Compara sin importar mayúsculas, tildes ni espacios ("Isabel Sofía G" = "isabelsofiag")
function limpiar(t) {
  return String(t || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "");
}
function igual(a, b) {
  return limpiar(a) !== "" && limpiar(a) === limpiar(b);
}

// Texto de 216 "0" y "1" -> lista. Si no hay texto, usa el número (formato viejo).
function aLista(texto, numero) {
  var l = [];
  var t = String(texto || "");
  var esTexto = /^[01]+$/.test(t);
  var n = Math.min(Number(numero) || 0, 216);
  for (var k = 0; k < 216; k++) l.push(esTexto ? t.charAt(k) === "1" : k < n);
  return l;
}

function aTexto(l) {
  var t = "";
  for (var k = 0; k < 216; k++) t += l[k] ? "1" : "0";
  return t;
}

function res(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Acepta la afinación escrita de varias formas en la hoja (C, Do, Bb, Sib, Si bemol...)
function afinacion(valor) {
  var v = String(valor || "").toLowerCase().replace(/\s/g, "").replace(/♭/g, "b").replace("bemol", "b");
  if (v === "c" || v === "do") return "C";
  if (v === "bb" || v === "sib") return "Bb";
  if (v === "eb" || v === "mib") return "Eb";
  if (v === "f" || v === "fa") return "F";
  return "";
}

function doGet(e) {
  return ContentService.createTextOutput("RUTA216 activo").setMimeType(ContentService.MimeType.TEXT);
}

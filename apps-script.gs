// RUTA216 v5.0 - SCRIPT DE CONEXIÓN COMPLETO
// Hoja "usuarios":
//   A usuario | B contraseña | C nombre | D instrumento (C, Bb, Eb o F)
//   E nivel | F escala | G escalas_hechas (0 a 216) | H detalle (lo llena la app)
//   I nombre completo | J instrumento real | K grado   (solo referencia)
//   L última escala (fecha) | M día | N escalas ese día | O vuelta   (lo llena la app)
// El ranking es diario: se reinicia a medianoche (hora de Colombia).

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
            hechas: String(rows[i][7] || ""),
            vuelta: Number(rows[i][14]) || 1
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
            hechas: String(rows[i][7] || ""),
            vuelta: Number(rows[i][14]) || 1
          });
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: GUARDAR EL AVANCE (nunca borra escalas ya hechas)
    if (action === "guardarProgreso") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          var vueltaHoja = Number(rows[i][14]) || 1;
          // si el celular viene de una vuelta anterior, no se mezcla con la nueva
          if ((Number(data.vuelta) || 1) < vueltaHoja) {
            return res({result: "success", progreso: Number(rows[i][6]) || 0, hechas: String(rows[i][7] || ""), vuelta: vueltaHoja});
          }
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
          // Ranking del día: cuenta las escalas nuevas de hoy
          if (recien > 0) {
            var dia = hoy();
            var deHoy = String(rows[i][12] || "") === dia ? (Number(rows[i][13]) || 0) : 0;
            sheet.getRange(i + 1, 12, 1, 3).setValues([[
              "'" + Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd HH:mm"),
              "'" + dia,
              deHoy + recien
            ]]);
          }
          return res({result: "success", progreso: total, hechas: aTexto(enHoja)});
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: VOLVER A EMPEZAR (borra las 216 y suma una vuelta)
    if (action === "reiniciar") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          var actual = Number(rows[i][14]) || 1;
          var nueva = Number(data.vuelta) || (actual + 1);
          // solo reinicia si es una vuelta nueva (si el aviso llega dos veces, no pasa nada)
          if (nueva > actual) {
            sheet.getRange(i + 1, 5, 1, 4).setValues([["Negras a 80", "C mayor", 0, "'" + aTexto([])]]);
            sheet.getRange(i + 1, 15).setValue(nueva);
            actual = nueva;
          }
          return res({result: "success", vuelta: actual});
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: RANKING DEL DÍA
    if (action === "ranking") {
      var dia = hoy();
      var lista = [];
      for (var i = 1; i < rows.length; i++) {
        var usuario = String(rows[i][0] || "");
        if (usuario === "" || limpiar(usuario) === "paisa") continue;   // el profesor no compite
        var n = String(rows[i][12] || "") === dia ? (Number(rows[i][13]) || 0) : 0;
        lista.push({
          nombre: String(rows[i][2] || usuario),
          hoy: n,
          total: Number(rows[i][6]) || 0,
          yo: igual(usuario, data.usuario)
        });
      }
      // más escalas hoy primero; si empatan, el que lleva más en total
      lista.sort(function (a, b) { return (b.hoy - a.hoy) || (b.total - a.total); });
      return res({result: "success", dia: dia, lista: lista});
    }

    return res({result: "error", message: "Acción desconocida"});

  } catch (error) {
    return res({result: "error", message: error.toString()});
  } finally {
    lock.releaseLock();
  }
}

var ZONA = "America/Bogota";

// Fecha de hoy en Colombia (ej. "2026-11-04")
function hoy() {
  return Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd");
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

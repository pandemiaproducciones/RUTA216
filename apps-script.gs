// RUTINA v7.0 (ESCALAS + COMPLEMENTO) - SCRIPT DE CONEXIÓN COMPLETO
// Hoja "usuarios":
//   A usuario | B contraseña | C nombre | D instrumento (C, Bb, Eb o F)
//   E nivel | F ejercicio actual | G ejercicios hechos (0 a 648) | H detalle (lo llena la app)
//   I nombre completo | J instrumento real | K grado   (solo referencia)
//   L último ejercicio (fecha) | M día | N ejercicios ese día | O vuelta   (lo llena la app)
//   COMPLEMENTO: P hechos | Q detalle | R día | S hechos ese día | T ejercicio actual   (lo llena la app)
// El ranking es diario: se reinicia a medianoche (hora de Colombia).
//
// 648 ejercicios = 6 velocidades x 3 familias x 3 tipos (escala, terceras conjuntas,
// terceras disjuntas) x 12 tonalidades.
// El detalle viejo (216 = solo escalas) se convierte solo, sin perder nada.

var TOTAL = 648;
var ZONA = "America/Bogota";

// Usuarios que no aparecen en el ranking (escritos en minúscula, sin tildes ni espacios)
var NO_COMPITEN = ["paisa", "invitado"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("usuarios");
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var rows = sheet.getDataRange().getValues();
    // la página nueva avisa que maneja 648; la vieja no dice nada y recibe el formato de 216
    var nuevo = Number(data.formato) === TOTAL;

    // ACCIÓN: LOGIN
    if (action === "login") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario) && igual(rows[i][1], data.contrasena)) {
          var l = aLista(rows[i][7], rows[i][6]);
          return res({
            result: "success",
            usuario: String(rows[i][0]),
            nombre_real: rows[i][2],
            instrumento: afinacion(rows[i][3]),
            progreso: contar(l),
            hechas: nuevo ? aTexto(l) : aTextoViejo(l),
            vuelta: Number(rows[i][14]) || 1,
            complemento: textoCeros(rows[i][16])
          });
        }
      }
      return res({result: "error", message: "Usuario o contraseña incorrectos"});
    }

    // ACCIÓN: LEER EL AVANCE
    if (action === "progreso") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          var l = aLista(rows[i][7], rows[i][6]);
          return res({
            result: "success",
            instrumento: afinacion(rows[i][3]),
            progreso: contar(l),
            hechas: nuevo ? aTexto(l) : aTextoViejo(l),
            vuelta: Number(rows[i][14]) || 1,
            complemento: textoCeros(rows[i][16])
          });
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: GUARDAR EL AVANCE (nunca borra ejercicios ya hechos)
    if (action === "guardarProgreso") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          var vueltaHoja = Number(rows[i][14]) || 1;
          var enHoja = aLista(rows[i][7], rows[i][6]);
          // si el celular viene de una vuelta anterior, no se mezcla con la nueva
          if ((Number(data.vuelta) || 1) < vueltaHoja) {
            return res({result: "success", progreso: contar(enHoja), hechas: nuevo ? aTexto(enHoja) : aTextoViejo(enHoja), vuelta: vueltaHoja});
          }
          var nuevas = aLista(data.hechas, data.progreso);
          var total = 0, recien = 0;
          for (var k = 0; k < TOTAL; k++) {
            if (nuevas[k] && !enHoja[k]) { enHoja[k] = true; recien++; }
            if (enHoja[k]) total++;
          }
          if (data.instrumento) sheet.getRange(i + 1, 4).setValue(data.instrumento);
          // la página vieja no conoce las terceras: no se le deja cambiar nivel/ejercicio
          sheet.getRange(i + 1, 5, 1, 4).setValues([[
            nuevo ? (data.nivel || "") : rows[i][4],
            nuevo ? (data.escala || "") : rows[i][5],
            total,
            "'" + aTexto(enHoja)
          ]]);
          // Ranking del día: cuenta los ejercicios nuevos de hoy
          if (recien > 0) {
            var dia = hoy();
            var deHoy = String(rows[i][12] || "") === dia ? (Number(rows[i][13]) || 0) : 0;
            sheet.getRange(i + 1, 12, 1, 3).setValues([[
              "'" + Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd HH:mm"),
              "'" + dia,
              deHoy + recien
            ]]);
          }
          return res({result: "success", progreso: total, hechas: nuevo ? aTexto(enHoja) : aTextoViejo(enHoja)});
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: GUARDAR AVANCE DE COMPLEMENTO (nunca borra lo ya hecho)
    if (action === "guardarComplemento") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          var u = unirTextos(textoCeros(rows[i][16]), textoCeros(data.hechas));
          sheet.getRange(i + 1, 16, 1, 2).setValues([[u.total, "'" + u.texto]]);
          sheet.getRange(i + 1, 20).setValue(data.actual || "");
          if (u.nuevos > 0) {
            var dia = hoy();
            var deHoy = String(rows[i][17] || "") === dia ? (Number(rows[i][18]) || 0) : 0;
            sheet.getRange(i + 1, 18, 1, 2).setValues([["'" + dia, deHoy + u.nuevos]]);
          }
          return res({result: "success", hechas: u.texto, progreso: u.total});
        }
      }
      return res({result: "error", message: "Usuario no encontrado"});
    }

    // ACCIÓN: RANKING DEL DÍA DE COMPLEMENTO
    if (action === "rankingComplemento") {
      var dia = hoy();
      var lista = [];
      for (var i = 1; i < rows.length; i++) {
        var usuario = String(rows[i][0] || "");
        if (usuario === "" || NO_COMPITEN.indexOf(limpiar(usuario)) >= 0) continue;
        var n = String(rows[i][17] || "") === dia ? (Number(rows[i][18]) || 0) : 0;
        lista.push({
          nombre: String(rows[i][2] || usuario),
          hoy: n,
          total: Number(rows[i][15]) || 0,
          yo: igual(usuario, data.usuario)
        });
      }
      lista.sort(function (a, b) { return (b.hoy - a.hoy) || (b.total - a.total); });
      return res({result: "success", dia: dia, lista: lista});
    }

    // ACCIÓN: VOLVER A EMPEZAR (borra todo y suma una vuelta)
    if (action === "reiniciar") {
      for (var i = 1; i < rows.length; i++) {
        if (igual(rows[i][0], data.usuario)) {
          var actual = Number(rows[i][14]) || 1;
          var nueva = Number(data.vuelta) || (actual + 1);
          // solo reinicia si es una vuelta nueva (si el aviso llega dos veces, no pasa nada)
          if (nueva > actual) {
            sheet.getRange(i + 1, 5, 1, 4).setValues([["Negras a 80", "C mayor · Escala", 0, "'" + aTexto([])]]);
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
        if (usuario === "" || NO_COMPITEN.indexOf(limpiar(usuario)) >= 0) continue;
        var n = String(rows[i][12] || "") === dia ? (Number(rows[i][13]) || 0) : 0;
        lista.push({
          nombre: String(rows[i][2] || usuario),
          hoy: n,
          total: Number(rows[i][6]) || 0,
          yo: igual(usuario, data.usuario)
        });
      }
      // más ejercicios hoy primero; si empatan, el que lleva más en total
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

// Fecha de hoy en Colombia (ej. "2026-11-09")
function hoy() {
  return Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd");
}

// Compara sin importar mayúsculas, tildes ni espacios ("Isabel Sofía G" = "isabelsofiag")
function limpiar(t) {
  return String(t || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, "");
}
function igual(a, b) {
  return limpiar(a) !== "" && limpiar(a) === limpiar(b);
}

// Posición de una escala del formato viejo (216) dentro del nuevo (648)
//   viejo: velocidad*36 + familia*12 + tonalidad
//   nuevo: velocidad*108 + familia*36 + tipo*12 + tonalidad   (tipo 0 = escala)
function nuevoIndice(i) {
  var velocidad = Math.floor(i / 36), familia = Math.floor((i % 36) / 12), tonalidad = i % 12;
  return velocidad * 108 + familia * 36 + tonalidad;
}

// Texto de "0" y "1" -> lista de 648. Acepta el formato viejo (216) y lo convierte.
function aLista(texto, numero) {
  var l = [];
  for (var k = 0; k < TOTAL; k++) l.push(false);
  var t = String(texto || "");
  if (/^[01]+$/.test(t)) {
    if (t.length === 216) {
      for (var i = 0; i < 216; i++) if (t.charAt(i) === "1") l[nuevoIndice(i)] = true;
    } else {
      for (var k = 0; k < TOTAL && k < t.length; k++) l[k] = t.charAt(k) === "1";
    }
  } else {
    var n = Math.min(Number(numero) || 0, 216);   // formato muy viejo: solo un número
    for (var i = 0; i < n; i++) l[nuevoIndice(i)] = true;
  }
  return l;
}

function aTexto(l) {
  var t = "";
  for (var k = 0; k < TOTAL; k++) t += l[k] ? "1" : "0";
  return t;
}

// Para la página vieja: solo las escalas, en su formato de 216
function aTextoViejo(l) {
  var t = "";
  for (var i = 0; i < 216; i++) t += l[nuevoIndice(i)] ? "1" : "0";
  return t;
}

// COMPLEMENTO: texto de "0" y "1" de cualquier largo (crece cuando se agregan ejercicios)
function textoCeros(t) {
  t = String(t || "");
  return /^[01]+$/.test(t) ? t : "";
}
function unirTextos(a, b) {
  var n = Math.max(a.length, b.length), texto = "", nuevos = 0, total = 0;
  for (var k = 0; k < n; k++) {
    var ya = a.charAt(k) === "1", llega = b.charAt(k) === "1";
    if (llega && !ya) nuevos++;
    if (ya || llega) { texto += "1"; total++; } else texto += "0";
  }
  return {texto: texto, nuevos: nuevos, total: total};
}

function contar(l) {
  var n = 0;
  for (var k = 0; k < TOTAL; k++) if (l[k]) n++;
  return n;
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
  return ContentService.createTextOutput("RUTINA activo").setMimeType(ContentService.MimeType.TEXT);
}

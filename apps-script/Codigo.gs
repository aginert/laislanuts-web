/**
 * La Isla Nuts — formulario de suscripción de laislanuts.com
 *
 * Recibe el formulario de la web, guarda cada suscripción en la hoja
 * «Suscripciones» y te envía un correo por cada persona nueva.
 * Los envíos de bots no te avisan: se apartan a la pestaña «Spam».
 *
 * Columnas de «Suscripciones»: Fecha | Email | Nombre | Consentimiento | Origen
 */

var ID_HOJA = '13oEZpf0uSUs9Urcbe6XJW7lTSnKe4OrzifKTThwhm-Q'; // «La Isla Nuts — Suscripciones»
var HOJA = 'Suscripciones';
var HOJA_SPAM = 'Spam';

function doPost(e) {
  try {
    return recibir(e);
  } catch (err) {
    return respuesta({ result: 'error', message: err.toString() });
  }
}

function recibir(e) {
  var p = (e && e.parameter) || {};
  var email = String(p.email || '').trim();
  var nombre = String(p.nombre || '').trim();
  var consentimiento = String(p.consentimiento || '').trim();
  var origen = String(p.origen || '').trim();

  if (!email) return respuesta({ result: 'error', message: 'sin email' });

  var fila = [new Date(), email, nombre, consentimiento, origen];
  var libro = SpreadsheetApp.openById(ID_HOJA);

  if (esBot(p, email, nombre)) {
    hoja(libro, HOJA_SPAM).appendRow(fila);
    return respuesta({ result: 'ok' });
  }

  var suscripciones = hoja(libro, HOJA);
  var repetido = yaEsta(suscripciones, email);
  suscripciones.appendRow(fila);
  try {
    avisar(libro, email, nombre, origen, repetido);
  } catch (err) {
    // Si el correo falla, la suscripción ya está guardada igualmente.
  }
  return respuesta({ result: 'ok' });
}

/* ── Detección de bots ─────────────────────────────── */
function esBot(p, email, nombre) {
  // 1. Trampa invisible: las personas no ven este campo; los bots lo rellenan.
  if (String(p.web || '').trim() !== '') return true;
  // 2. Enviado demasiado rápido (menos de 3 segundos tras abrir la página).
  var segundos = Number(p.tiempo || 0);
  if (p.tiempo !== undefined && segundos > 0 && segundos < 3) return true;
  // 3. Patrón del spam que llegaba: nombre de letras al azar y Gmail lleno de puntos.
  var usuario = email.split('@')[0];
  var puntos = (usuario.match(/\./g) || []).length;
  var nombreRaro = /^[A-Za-z]{12,}$/.test(nombre) && /[a-z][A-Z]/.test(nombre) && /[A-Z].*[A-Z].*[A-Z]/.test(nombre);
  if (/@gmail\.com$/i.test(email) && puntos >= 3 && (nombreRaro || nombre === '')) return true;
  return false;
}

/* ── Correo de aviso ───────────────────────────────── */
function avisar(libro, email, nombre, origen, repetido) {
  var destino = Session.getEffectiveUser().getEmail(); // la dueña de la hoja
  var total = Math.max(0, hoja(libro, HOJA).getLastRow() - 1);
  var asunto = (repetido ? 'Se ha vuelto a apuntar: ' : 'Nueva suscripción: ') + (nombre || email);
  var cuerpo =
    (nombre ? 'Nombre: ' + nombre + '\n' : '') +
    'Correo: ' + email + '\n' +
    'Llegó desde: ' + (origen || 'directo') + '\n' +
    (repetido ? 'Ojo: este correo ya estaba en la lista.\n' : '') +
    '\nFilas en la hoja: ' + total + '\n' + libro.getUrl();
  MailApp.sendEmail({ to: destino, subject: asunto, body: cuerpo, name: 'La Isla Nuts · Web' });
}

/* ── Utilidades ────────────────────────────────────── */
function hoja(libro, nombre) {
  var h = libro.getSheetByName(nombre);
  if (!h) {
    h = libro.insertSheet(nombre);
    h.appendRow(['Fecha', 'Email', 'Nombre', 'Consentimiento', 'Origen']);
  }
  return h;
}

function yaEsta(h, email) {
  var n = h.getLastRow();
  if (n < 2) return false;
  var correos = h.getRange(2, 2, n - 1, 1).getValues();
  var buscado = email.toLowerCase();
  return correos.some(function (fila) { return String(fila[0]).toLowerCase() === buscado; });
}

function respuesta(datos) {
  return ContentService.createTextOutput(JSON.stringify(datos))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Ejecútala una vez a mano desde el editor para dar permiso de enviar correos y probar el aviso. */
function probarAviso() {
  avisar(SpreadsheetApp.openById(ID_HOJA), 'prueba@laislanuts.com', 'Prueba', 'editor', false);
}

/* ── Limpieza de la hoja (ejecutar a mano desde el editor) ───────────── */
/**
 * Deja en «Suscripciones» solo a las personas reales.
 * Lo demás NO se borra: se mueve a «Descartados» con el motivo (Bot, Prueba o Comercial).
 * También quita de «Spam» las filas de prueba.
 */
function limpiarHoja() {
  var libro = SpreadsheetApp.openById(ID_HOJA);
  var h = hoja(libro, HOJA);
  var n = h.getLastRow();
  if (n < 2) return;
  var filas = h.getRange(2, 1, n - 1, 5).getValues();
  var quedan = [], fuera = [];

  filas.forEach(function (f) {
    var email = String(f[1]).trim(), nombre = String(f[2]).trim();
    var motivo = motivoDescarte(email, nombre);
    if (motivo) fuera.push(f.concat([motivo])); else quedan.push(f);
  });

  if (fuera.length) {
    var d = libro.getSheetByName('Descartados');
    if (!d) {
      d = libro.insertSheet('Descartados');
      d.appendRow(['Fecha', 'Email', 'Nombre', 'Consentimiento', 'Origen', 'Motivo']);
    }
    d.getRange(d.getLastRow() + 1, 1, fuera.length, 6).setValues(fuera);
  }
  h.getRange(2, 1, n - 1, 5).clearContent();
  if (quedan.length) h.getRange(2, 1, quedan.length, 5).setValues(quedan);

  var spam = libro.getSheetByName(HOJA_SPAM);
  if (spam && spam.getLastRow() > 1) {
    for (var i = spam.getLastRow(); i >= 2; i--) {
      if (String(spam.getRange(i, 5).getValue()).indexOf('prueba') === 0) spam.deleteRow(i);
    }
  }
  Logger.log('Quedan ' + quedan.length + ' personas. Movidas a «Descartados»: ' + fuera.length + '.');
}

function motivoDescarte(email, nombre) {
  var e = email.toLowerCase();
  if (/@laislanuts\.com$/.test(e) || e === 'ana.ginert@gmail.com') return 'Prueba';
  if (/@melottogroup\.com$/.test(e)) return 'Comercial';
  if (esBot({}, email, nombre)) return 'Bot';
  return '';
}

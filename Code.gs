/**
 * Calendriers des pompiers — Puiselet-le-Marais
 * Serveur de l'app (Google Apps Script). À coller dans Extensions > Apps Script.
 * Les données sont dans l'onglet "Tournee" du Google Sheet.
 * Codes d'accès : onglet "Reglages" (B1 = code pompiers, B2 = code responsable).
 * © FEVRE Léo — créateur de l'application.
 * Responsable de la tournée : Jules FEVRE (gère les codes d'accès).
 */

var SHEET = 'Tournee';
// Colonnes (1 = A)
var C = { ID: 1, SECTEUR: 2, ORDRE: 3, QUARTIER: 4, NOM: 5, NUMERO: 6, RUE: 7, REMARQUE: 8,
          A2025: 9, STATUT: 10, MONTANT: 11, PAIEMENT: 12, POMPIER: 13, DATE: 14, NOTE: 15 };
var NB_COL = 15;
var STATUTS = ['', 'Donné', 'Absent', 'Refus'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Calendriers Puiselet')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- API pour la page GitHub (appelée en POST, réponse JSON) ---------- */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var fns = { getData: getData, verifyCode: verifyCode, savePassage: savePassage,
                addAdresse: addAdresse, setEquipe: setEquipe, adminSetCodes: adminSetCodes };
    var f = fns[req.fn];
    if (!f) throw new Error('Action inconnue.');
    out = { ok: true, data: f(req.arg || {}) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh) throw new Error('Onglet "' + SHEET + '" introuvable dans le Google Sheet.');
  return sh;
}

function fmtDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, 'Europe/Paris', 'dd/MM HH:mm');
  }
  return String(d);
}

function rowToObj_(r) {
  return {
    id: String(r[C.ID - 1]),
    secteur: String(r[C.SECTEUR - 1]),
    ordre: Number(r[C.ORDRE - 1]) || 0,
    quartier: String(r[C.QUARTIER - 1] || ''),
    nom: String(r[C.NOM - 1] || ''),
    numero: String(r[C.NUMERO - 1] || ''),
    rue: String(r[C.RUE - 1] || ''),
    remarque: String(r[C.REMARQUE - 1] || ''),
    a2025: String(r[C.A2025 - 1] || ''),
    statut: String(r[C.STATUT - 1] || ''),
    montant: r[C.MONTANT - 1] === '' ? '' : Number(r[C.MONTANT - 1]) || 0,
    paiement: String(r[C.PAIEMENT - 1] || ''),
    pompier: String(r[C.POMPIER - 1] || ''),
    date: fmtDate_(r[C.DATE - 1]),
    note: String(r[C.NOTE - 1] || '')
  };
}

/** Toutes les adresses (appelé au chargement puis régulièrement). */
/* ---------- Codes d'accès (onglet "Reglages") ---------- */
function reglages_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Reglages');
  if (!sh) {
    sh = ss.insertSheet('Reglages');
    sh.getRange('B1:B2').setNumberFormat('@');
    sh.getRange('A1:B2').setValues([['Code pompiers', '1818'], ['Code responsable', '9999']]);
  }
  return sh;
}
function codes_() {
  var v = reglages_().getRange('B1:B2').getDisplayValues();
  return { pompiers: String(v[0][0]).trim(), admin: String(v[1][0]).trim() };
}
function guard_() {
  var n = Number(CacheService.getScriptCache().get('fails') || 0);
  if (n >= 30) throw new Error('TROP_ESSAIS : trop de codes faux, réessaie dans 10 minutes.');
}
function fail_() {
  var c = CacheService.getScriptCache(), n = Number(c.get('fails') || 0) + 1;
  c.put('fails', String(n), 600);
}
/** Vérifie le code envoyé par l'app. Renvoie 'admin' ou 'pompier'. */
function checkCode_(code) {
  guard_();
  var c = codes_(), s = String(code == null ? '' : code).trim();
  if (s && c.admin.length >= 4 && s === c.admin) return 'admin';
  if (s && c.pompiers.length >= 4 && s === c.pompiers) return 'pompier';
  fail_();
  throw new Error('CODE_INVALIDE');
}
function verifyCode(a) { return checkCode_(a && a.code); }

/** Espace responsable : change les codes. a = {admin, pompiers, nouvelAdmin} */
function adminSetCodes(a) {
  guard_();
  var c = codes_();
  if (!a || String(a.admin || '').trim() !== c.admin || c.admin.length < 4) { fail_(); throw new Error('Code responsable incorrect.'); }
  var p = String(a.pompiers || '').trim(), n = String(a.nouvelAdmin || '').trim();
  if (p && !/^[0-9A-Za-z]{4,12}$/.test(p)) throw new Error('Le code pompiers doit faire 4 à 12 chiffres ou lettres.');
  if (n && !/^[0-9A-Za-z]{4,12}$/.test(n)) throw new Error('Le code responsable doit faire 4 à 12 chiffres ou lettres.');
  if ((n || c.admin) === (p || c.pompiers)) throw new Error('Les deux codes doivent être différents.');
  var sh = reglages_();
  sh.getRange('B1:B2').setNumberFormat('@');
  if (p) sh.getRange('B1').setValue(p);
  if (n) sh.getRange('B2').setValue(n);
  return true;
}

function getData(a) {
  checkCode_(a && a.code);
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return { rows: [], equipes: getEquipes_(), t: Date.now() };
  var values = sh.getRange(2, 1, last - 1, NB_COL).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === '') continue;
    rows.push(rowToObj_(values[i]));
  }
  return { rows: rows, equipes: getEquipes_(), t: Date.now() };
}

/* ---------- Équipes : qui prend quel secteur (onglet "Equipes") ---------- */
function equipesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Equipes');
  if (!sh) { sh = ss.insertSheet('Equipes'); sh.getRange(1, 1, 1, 2).setValues([['Secteur', 'Équipe']]).setFontWeight('bold'); }
  return sh;
}
function getEquipes_() {
  var sh = equipesSheet_(), last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { if (r[0]) out[String(r[0])] = String(r[1] || ''); });
  return out;
}
/** e = {secteur, equipe} ; equipe vide = secteur libéré */
function setEquipe(e) {
  checkCode_(e && e.code);
  if (!e || !e.secteur) throw new Error('Secteur manquant.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = equipesSheet_(), last = sh.getLastRow(), row = -1;
    if (last >= 2) {
      var v = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(e.secteur)) { row = i + 2; break; }
    }
    if (row < 0) row = last + 1;
    sh.getRange(row, 1, 1, 2).setValues([[clean_(e.secteur, 80), clean_(e.equipe, 60)]]);
    return getEquipes_();
  } finally {
    lock.releaseLock();
  }
}

function clean_(s, max) {
  return String(s == null ? '' : s).replace(/^[=+\-@]+/, '').slice(0, max || 200);
}

/** Enregistre un passage. p = {id, statut, montant, paiement, note, pompier} */
function savePassage(p) {
  checkCode_(p && p.code);
  if (!p || !p.id) throw new Error('Adresse manquante.');
  if (STATUTS.indexOf(p.statut) < 0) throw new Error('Statut inconnu.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_();
    var last = sh.getLastRow();
    var ids = sh.getRange(2, C.ID, last - 1, 1).getValues();
    var row = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(p.id)) { row = i + 2; break; }
    }
    if (row < 0) throw new Error('Adresse introuvable : ' + p.id);
    var montant = p.statut === 'Donné' && p.montant !== '' && p.montant != null ? Number(p.montant) : '';
    if (montant !== '' && (isNaN(montant) || montant < 0 || montant > 10000)) throw new Error('Montant invalide.');
    var vide = p.statut === '';
    sh.getRange(row, C.STATUT, 1, 6).setValues([[
      p.statut,
      vide ? '' : montant,
      vide || p.statut !== 'Donné' ? '' : clean_(p.paiement, 20),
      vide ? '' : clean_(p.pompier, 40),
      vide ? '' : new Date(),
      clean_(p.note, 300)
    ]]);
    return rowToObj_(sh.getRange(row, 1, 1, NB_COL).getValues()[0]);
  } finally {
    lock.releaseLock();
  }
}

/** Ajoute une adresse à la fin d'un secteur. a = {secteur, nom, numero, rue, quartier} */
function addAdresse(a) {
  checkCode_(a && a.code);
  if (!a || !a.secteur) throw new Error('Secteur manquant.');
  if (!a.numero && !a.nom && !a.rue) throw new Error('Indique au moins un nom, un numéro ou une rue.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_();
    var last = sh.getLastRow();
    var data = sh.getRange(2, 1, last - 1, 3).getValues();
    var maxOrdre = 0, prefix = 'NEW', n = 0;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1]) === String(a.secteur)) {
        maxOrdre = Math.max(maxOrdre, Number(data[i][2]) || 0);
        prefix = String(data[i][0]).split('-')[0];
      }
    }
    for (var j = 0; j < data.length; j++) {
      var id = String(data[j][0]);
      if (id.indexOf(prefix + '-A') === 0) n = Math.max(n, Number(id.split('-A')[1]) || 0);
    }
    var newId = prefix + '-A' + ('00' + (n + 1)).slice(-3);
    var line = [newId, a.secteur, maxOrdre + 1, clean_(a.quartier, 60), clean_(a.nom, 80), clean_(a.numero, 20),
                clean_(a.rue, 80), 'Ajoutée par ' + clean_(a.pompier, 40), '', '', '', '', '', '', ''];
    sh.getRange(last + 1, C.NUMERO).setNumberFormat('@');
    sh.getRange(last + 1, 1, 1, NB_COL).setValues([line]);
    return rowToObj_(line);
  } finally {
    lock.releaseLock();
  }
}

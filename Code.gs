// =====================================================================
//  COVOITURAGE KARATÉ — COMPÉTITIONS
//  Application web Google Apps Script, sans inscription.
//  Données : un fichier Google Sheets « Covoiturage Compétitions — Karaté »
//  créé automatiquement dans Google Drive au premier lancement,
//  avec 3 onglets : Compétitions / Conducteurs / Passagers.
// =====================================================================

var SPREADSHEET_NAME = 'Covoiturage Compétitions — Karaté';
var PROP_KEY = 'spreadsheetIdCompetitions';
var SHEET_COMPETITIONS = 'Compétitions';
var SHEET_DRIVERS = 'Conducteurs';
var SHEET_PASSENGERS = 'Passagers';

var COMPETITIONS_HEADER = ['Horodateur', 'ID', 'Nom', 'Date', 'Heure de rendez-vous', 'Lieu'];
var DRIVERS_HEADER = [
  'Horodateur', 'ID', 'ID compétition', 'Prénom', 'Heure de départ',
  'Point de prise en charge', 'Adresse (géolocalisée)', 'Places'
];
var PASSENGERS_HEADER = [
  'Horodateur', 'ID', 'ID compétition', 'ID conducteur', 'Prénom', 'Demande',
  'Point de prise en charge', 'Heure souhaitée'
];

/** Point d'entrée de la web-app. */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Covoiturage Compétitions — Karaté')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Permet d'inclure un fragment HTML : <?!= include('fichier') ?>. */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ---------------------------------------------------------------------
 *  FEUILLE GOOGLE SHEETS
 * ------------------------------------------------------------------- */

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_KEY);
  var ss = null;

  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; }
  }
  if (ss && ss.getSheetByName(SHEET_COMPETITIONS) && ss.getSheetByName(SHEET_DRIVERS) && ss.getSheetByName(SHEET_PASSENGERS)) {
    return ss;
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (id) {
      try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; }
      if (ss && ss.getSheetByName(SHEET_COMPETITIONS) && ss.getSheetByName(SHEET_DRIVERS) && ss.getSheetByName(SHEET_PASSENGERS)) {
        return ss;
      }
    }
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    createSheet_(ss, SHEET_COMPETITIONS, COMPETITIONS_HEADER);
    createSheet_(ss, SHEET_DRIVERS, DRIVERS_HEADER);
    createSheet_(ss, SHEET_PASSENGERS, PASSENGERS_HEADER);
    var first = ss.getSheets()[0];
    if (first && first.getName() !== SHEET_COMPETITIONS && first.getName() !== SHEET_DRIVERS && first.getName() !== SHEET_PASSENGERS) {
      ss.deleteSheet(first);
    }
    props.setProperty(PROP_KEY, ss.getId());
    Logger.log('Fichier Google Sheets créé : ' + ss.getUrl());
    return ss;
  } finally {
    lock.releaseLock();
  }
}

function createSheet_(ss, name, header) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange('A:Z').setNumberFormat('@');
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** État complet, renvoyé au client après chaque action. */
function getData() {
  return buildState_(getSpreadsheet_());
}

/** Normalise une cellule date (éventuellement convertie en Date par Sheets) en "yyyy-MM-dd". */
function dateSheetStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  return m ? m[1] + '-' + m[2] + '-' + m[3] : s;
}

/** Normalise une cellule horaire (éventuellement convertie en Date par Sheets) en "HH:mm". */
function timeSheetStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  }
  var s = String(v || '').trim();
  var m = /^(\d{1,2}):(\d{2})/.exec(s);
  if (m) {
    var hh = ('0' + parseInt(m[1], 10)).slice(-2);
    return hh + ':' + m[2];
  }
  return s;
}

function buildState_(ss) {
  var bookingsByDriver = {};
  var waitingByComp = {};
  var driversByComp = {};
  var compIndex = {};
  var competitions = [];

  var sh = ss.getSheetByName(SHEET_COMPETITIONS);
  var last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, COMPETITIONS_HEADER.length).getValues().forEach(function (r) {
      var comp = {
        id: String(r[1]),
        nom: String(r[2]),
        date: dateSheetStr_(r[3]),
        heureRdv: timeSheetStr_(r[4]),
        lieu: String(r[5])
      };
      if (!comp.id || !comp.nom) return;
      competitions.push(comp);
      compIndex[comp.id] = comp;
    });
  }

  var ph = ss.getSheetByName(SHEET_PASSENGERS);
  last = ph.getLastRow();
  if (last > 1) {
    ph.getRange(2, 1, last - 1, PASSENGERS_HEADER.length).getValues().forEach(function (r) {
      var compId = String(r[2]);
      var driverId = String(r[3]);
      var prenom = String(r[4]).trim();
      var type = String(r[5]).trim();
      if (!prenom) return;
      if (type === 'confirme' && driverId) {
        if (!bookingsByDriver[driverId]) bookingsByDriver[driverId] = [];
        bookingsByDriver[driverId].push({ id: String(r[1]), prenom: prenom });
      } else if (type === 'attente' && compId) {
        if (!waitingByComp[compId]) waitingByComp[compId] = [];
        waitingByComp[compId].push({ id: String(r[1]), prenom: prenom, point: String(r[6]), heure: timeSheetStr_(r[7]) });
      }
    });
  }

  var dh = ss.getSheetByName(SHEET_DRIVERS);
  last = dh.getLastRow();
  if (last > 1) {
    dh.getRange(2, 1, last - 1, DRIVERS_HEADER.length).getValues().forEach(function (r) {
      var id = String(r[1]);
      var compId = String(r[2]);
      var prenom = String(r[3]).trim();
      if (!id || !prenom || !compIndex[compId]) return;
      var pax = bookingsByDriver[id] || [];
      var places = parseInt(String(r[7]), 10);
      if (!(places >= 1)) places = 1;
      var driver = {
        id: id,
        prenom: prenom,
        heure: timeSheetStr_(r[4]),
        point: String(r[5]),
        adresse: String(r[6]),
        places: places,
        passagers: pax,
        restantes: Math.max(0, places - pax.length)
      };
      if (!driversByComp[compId]) driversByComp[compId] = [];
      driversByComp[compId].push(driver);
    });
  }

  competitions.forEach(function (comp) {
    var drivers = driversByComp[comp.id] || [];
    drivers.sort(function (a, b) { return a.heure < b.heure ? -1 : (a.heure > b.heure ? 1 : 0); });
    var waiting = waitingByComp[comp.id] || [];
    waiting.sort(function (a, b) {
      var pa = (a.point || '~').toLowerCase();
      var pb = (b.point || '~').toLowerCase();
      return pa < pb ? -1 : (pa > pb ? 1 : 0);
    });
    comp.drivers = drivers;
    comp.attente = waiting;
    comp.placesRestantes = drivers.reduce(function (sum, d) { return sum + d.restantes; }, 0);
  });

  competitions.sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  return { competitions: competitions };
}

/* ---------------------------------------------------------------------
 *  ACTIONS
 * ------------------------------------------------------------------- */

function findCompetition_(competitions, compId) {
  for (var i = 0; i < competitions.length; i++) {
    if (competitions[i].id === compId) return competitions[i];
  }
  return null;
}

/** Créer une compétition « Doodle ». */
function addCompetition(p) {
  p = p || {};
  var nom = cleanText_(p.nom, 120);
  var date = cleanText_(p.date, 10);
  var heureRdv = cleanText_(p.heureRdv, 10);
  var lieu = cleanText_(p.lieu, 200);

  if (!nom) throw new Error('Merci d’indiquer le nom de la compétition.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Merci d’indiquer la date de la compétition.');
  if (!lieu) throw new Error('Merci d’indiquer le lieu de la compétition.');

  var ss = getSpreadsheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ss.getSheetByName(SHEET_COMPETITIONS)
      .appendRow([new Date().toISOString(), Utilities.getUuid(), nom, date, heureRdv, lieu]);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** Un conducteur propose sa voiture pour une compétition. */
function addDriver(p) {
  p = p || {};
  var prenom = cleanText_(p.prenom, 40);
  var compId = cleanText_(p.competitionId, 60);
  var heure = cleanText_(p.heure, 10);
  var point = cleanText_(p.point, 200);
  var places = parseInt(String(p.places || ''), 10);

  if (!prenom) throw new Error('Merci d’indiquer ton prénom.');
  if (!compId) throw new Error('Merci de choisir la compétition.');
  if (!heure) throw new Error('Merci d’indiquer ton horaire de départ.');
  if (!point) throw new Error('Merci d’indiquer ton point de prise en charge.');
  if (!(places >= 1 && places <= 8)) throw new Error('Le nombre de places doit être entre 1 et 8.');

  var ss = getSpreadsheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var state = buildState_(ss);
    var comp = findCompetition_(state.competitions, compId);
    if (!comp) throw new Error('Cette compétition n’existe plus. Actualise la page.');
    var drivers = comp.drivers || [];
    for (var i = 0; i < drivers.length; i++) {
      if (drivers[i].prenom.toLowerCase() === prenom.toLowerCase()) {
        throw new Error('Tu proposes déjà une voiture pour cette compétition, ' + prenom + '.');
      }
    }
    ss.getSheetByName(SHEET_DRIVERS)
      .appendRow([
        new Date().toISOString(), Utilities.getUuid(), compId, prenom, heure,
        point, cleanText_(p.adresse, 300), String(places)
      ]);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** Un passager réserve une place… ou se déclare en attente. */
function addPassenger(p) {
  p = p || {};
  var prenom = cleanText_(p.prenom, 40);
  var compId = cleanText_(p.competitionId, 60);
  var driverId = cleanText_(p.driverId, 60);
  var point = cleanText_(p.point, 200);
  var heure = cleanText_(p.heure, 10);
  var mode = driverId ? 'confirme' : 'attente';

  if (!prenom) throw new Error('Merci d’indiquer ton prénom.');
  if (!compId) throw new Error('Merci de choisir la compétition.');
  if (mode === 'attente' && !point) {
    throw new Error('Merci d’indiquer un point de prise en charge le plus pratique pour toi.');
  }

  var ss = getSpreadsheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var state = buildState_(ss);
    var comp = findCompetition_(state.competitions, compId);
    if (!comp) throw new Error('Cette compétition n’existe plus. Actualise la page.');

    var i;
    if (mode === 'confirme') {
      var target = null;
      var drivers = comp.drivers || [];
      for (i = 0; i < drivers.length; i++) {
        if (drivers[i].id === driverId) { target = drivers[i]; break; }
      }
      if (!target) throw new Error('Ce conducteur n’existe plus pour cette compétition. Actualise la page.');
      if (target.restantes < 1) throw new Error('Désolé, la voiture de ' + target.prenom + ' est complète.');
      for (i = 0; i < target.passagers.length; i++) {
        if (target.passagers[i].prenom.toLowerCase() === prenom.toLowerCase()) {
          throw new Error('Tu es déjà réservé dans la voiture de ' + target.prenom + '.');
        }
      }
    } else {
      var waiting = comp.attente || [];
      for (i = 0; i < waiting.length; i++) {
        if (waiting[i].prenom.toLowerCase() === prenom.toLowerCase()) {
          throw new Error('Ta demande de place pour cette compétition est déjà enregistrée.');
        }
      }
    }

    ss.getSheetByName(SHEET_PASSENGERS)
      .appendRow([
        new Date().toISOString(), Utilities.getUuid(), compId,
        mode === 'confirme' ? driverId : '', prenom, mode, point, heure
      ]);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** Le conducteur dépose un passager (liste d'attente ou autre voiture) dans sa voiture. */
function movePassenger(p) {
  p = p || {};
  var entryId = cleanText_(p.entryId, 60);
  var driverId = cleanText_(p.driverId, 60);
  if (!entryId || !driverId) throw new Error('Déplacement incomplet. Actualise la page et réessaie.');

  var ss = getSpreadsheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var state = buildState_(ss);
    var target = null;
    for (var c = 0; c < state.competitions.length && !target; c++) {
      var drivers = state.competitions[c].drivers || [];
      for (var k = 0; k < drivers.length; k++) {
        if (drivers[k].id === driverId) { target = drivers[k]; break; }
      }
    }
    if (!target) throw new Error('Voiture introuvable. Actualise la page.');

    var sh = ss.getSheetByName(SHEET_PASSENGERS);
    var last = sh.getLastRow();
    if (last < 2) throw new Error('Inscription introuvable. Actualise la page.');
    var values = sh.getRange(2, 1, last - 1, PASSENGERS_HEADER.length).getValues();
    var rowIndex = -1;
    var rowPrenom = '';
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][1]) === entryId) {
        rowIndex = i + 2;
        rowPrenom = String(values[i][4]).trim();
        break;
      }
    }
    if (rowIndex < 0) throw new Error('Inscription introuvable. Actualise la page.');

    for (i = 0; i < target.passagers.length; i++) {
      if (target.passagers[i].prenom.toLowerCase() === rowPrenom.toLowerCase()) {
        throw new Error(rowPrenom + ' est déjà dans cette voiture.');
      }
    }
    if (target.restantes < 1) throw new Error('La voiture de ' + target.prenom + ' est complète.');

    sh.getRange(rowIndex, 4).setValue(driverId);
    sh.getRange(rowIndex, 6).setValue('confirme');
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** Annulation : pour un conducteur ou un passager, le prénom doit correspondre. */
function removeEntry(p) {
  p = p || {};
  var prenom = cleanText_(p.prenom, 40);
  var type = String(p.type || '');
  var id = String(p.id || '');

  var ss = getSpreadsheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var state = buildState_(ss);

    if (type === 'competition') {
      var comp = findCompetition_(state.competitions, id);
      if (!comp) throw new Error('Compétition introuvable.');
      if ((comp.drivers || []).length > 0 || (comp.attente || []).length > 0) {
        throw new Error('Impossible de supprimer : vide d’abord les voitures et la liste d’attente de cette compétition.');
      }
      deleteRowById_(ss.getSheetByName(SHEET_COMPETITIONS), id, null);
    } else if (type === 'conducteur') {
      if (!prenom) throw new Error('Merci de redonner ton prénom pour confirmer.');
      deleteRowById_(ss.getSheetByName(SHEET_DRIVERS), id, { prenom: prenom });
    } else if (type === 'reservations') {
      if (!prenom) throw new Error('Merci de redonner ton prénom pour confirmer.');
      deleteRowById_(ss.getSheetByName(SHEET_PASSENGERS), id, { prenom: prenom, type: 'confirme', typeIndex: 5 });
    } else if (type === 'attente') {
      if (!prenom) throw new Error('Merci de redonner ton prénom pour confirmer.');
      deleteRowById_(ss.getSheetByName(SHEET_PASSENGERS), id, { prenom: prenom, type: 'attente', typeIndex: 5 });
    } else {
      throw new Error('Type d’entrée inconnu.');
    }
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function deleteRowById_(sheet, id, match) {
  var last = sheet.getLastRow();
  if (last < 2) throw new Error('Entrée introuvable.');
  var width = Math.max(COMPETITIONS_HEADER.length, DRIVERS_HEADER.length, PASSENGERS_HEADER.length);
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][1]) !== id) continue;
    if (match) {
      var prenomIndex = sheet.getName() === SHEET_PASSENGERS ? 4 : 3;
      if (String(values[i][prenomIndex]).trim().toLowerCase() !== match.prenom.toLowerCase()) continue;
      if (match.type !== null && match.typeIndex !== null) {
        if (String(values[i][match.typeIndex]).trim() !== match.type) continue;
      }
    }
    sheet.deleteRow(i + 2);
    return;
  }
  throw new Error(match ? 'Entrée introuvable, ou le prénom ne correspond pas.' : 'Entrée introuvable.');
}

/* ---------------------------------------------------------------------
 *  OUTILS
 * ------------------------------------------------------------------- */

function cleanText_(v, max) {
  var s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  return s.slice(0, max || 100);
}

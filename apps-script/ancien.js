// ============================================================
//  ANCIEN SHEETS — fiches d'inspection d'avant la webapp (lecture seule)
//  Holland Bikes / Montévrain — v14 (25/11/2026)
//
//  Fichier « ancien.gs » dans le même projet Apps Script que summary.gs.
//
//  « Tableau - Rapport d'inspection reconditionnement complet VAE »
//  (environ 60 onglets, un par vélo, nommés « 66296 » ou « 55638  OK »).
//  Ce classeur n'est plus utilisé : le script le LIT seulement, il n'y
//  écrit jamais. Les fiches sont indexées une fois, puis gardées dans les
//  Propriétés du script (réponse rapide, plus besoin de relire le classeur).
//
//  Si une fiche existe dans les DEUX classeurs, celle du Sheets actuel l'emporte.
//
//  À lancer depuis l'éditeur :
//    indexerAncienSheets()  → (re)construit l'index (≈ 30 s, à refaire
//                              seulement si l'ancien classeur est modifié)
// ============================================================

const ARCHIVE_SPREADSHEET_ID = '1Xv6J0Kef-oD3T4WgZdDsRqkeD4b2B-dD4S4cIMTcglM';
const ARCHIVE_PROP_PREFIX    = 'ARCHIVE_IDX_';
const ARCHIVE_CHUNK          = 8000;   // une propriété du script = 9 Ko max

var _archive = null;

// { N° CS → fiche ancienne } — lu dans les Propriétés du script, construit si absent
function getArchive() {
  if (_archive) return _archive;
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty(ARCHIVE_PROP_PREFIX + 'N'), 10);
  if (n) {
    var json = '';
    for (var i = 0; i < n; i++) json += props.getProperty(ARCHIVE_PROP_PREFIX + i) || '';
    try { _archive = JSON.parse(json); return _archive; } catch (e) { Logger.log('Index ancien illisible, reconstruction : ' + e.message); }
  }
  _archive = indexerAncienSheets();
  return _archive;
}

function archiveFor(cs) {
  try { return getArchive()[String(cs).trim()] || null; }
  catch (e) { Logger.log('archiveFor: ' + e.message); return null; }
}

// ── Construction de l'index (lecture seule de l'ancien classeur) ──
function indexerAncienSheets() {
  var ss = SpreadsheetApp.openById(ARCHIVE_SPREADSHEET_ID);
  var out = {};
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    var m = name.trim().match(/^(\d{5})\b/);          // « 66296 », « 55638  OK »
    if (!m) return;
    try {
      var f = parseArchiveSheet_(sheet, m[1]);
      if (f) out[m[1]] = f;
    } catch (e) { Logger.log('Onglet « ' + name + ' » ignoré : ' + e.message); }
  });

  // Sauvegarde en morceaux dans les Propriétés du script
  var props = PropertiesService.getScriptProperties();
  var old = parseInt(props.getProperty(ARCHIVE_PROP_PREFIX + 'N'), 10) || 0;
  for (var k = 0; k < old; k++) props.deleteProperty(ARCHIVE_PROP_PREFIX + k);
  var json = JSON.stringify(out), parts = {}, n = Math.ceil(json.length / ARCHIVE_CHUNK) || 1;
  for (var i = 0; i < n; i++) parts[ARCHIVE_PROP_PREFIX + i] = json.substr(i * ARCHIVE_CHUNK, ARCHIVE_CHUNK);
  parts[ARCHIVE_PROP_PREFIX + 'N'] = String(n);
  props.setProperties(parts);

  _archive = out;
  Logger.log(Object.keys(out).length + ' fiches anciennes indexées (' + json.length + ' caractères) : ' + Object.keys(out).join(', '));
  return out;
}

// ── Lecture d'un onglet de l'ancien format ──
//  En-tête : A « Date » / « Technicien… » / « …Marque / Modèle » / « N° de cadre » / « Kilometrage » → valeur en C
//  Specs   : B « Marque / Modèle du moteur », « Capacité (Wh) » → valeur en C
//  Points  : C « 12. libellé » → statut en D, état en E (lignes variables)
//  D1      : ancien statut (« Vélo prêt pour la vente », « Vélo non-prêt »…)
function parseArchiveSheet_(sheet, cs) {
  var last = Math.min(sheet.getLastRow(), 130);
  if (last < 5) return null;
  var g = sheet.getRange(1, 1, last, 5).getDisplayValues();
  function t(v) { return String(v || '').replace(/\s+/g, ' ').trim(); }

  var f = { cs: cs, tab: sheet.getName(), gid: sheet.getSheetId(), ancienStatut: t(g[0][3]),
            date: '', technicien: '', titre: '', cadre: '', km: '', moteur: '', batterieWh: '',
            ok: 0, pasOk: 0, na: 0, nonRens: 0, pasOkLabels: [] };
  var seen = {};

  g.forEach(function(row, r) {
    var a = t(row[0]).toLowerCase(), b = t(row[1]).toLowerCase(), c = t(row[2]);
    if (r < 8) {
      if (a === 'date') f.date = c;
      else if (a.indexOf('technicien') === 0) f.technicien = c;
      else if (a.indexOf('marque') >= 0) f.titre = c;
      else if (a.indexOf('cadre') >= 0) f.cadre = c;
      else if (a.indexOf('kilom') === 0) f.km = c;
    }
    if (b.indexOf('marque / modèle du moteur') === 0) f.moteur = c;
    if (b.indexOf('capacité') === 0) f.batterieWh = c;

    var m = c.match(/^(\d{1,2})\.\s*(.+)$/);
    if (!m) return;
    var n = parseInt(m[1], 10);
    if (n < 1 || n > 53 || seen[n]) return;
    seen[n] = true;
    var st = t(row[3]).toUpperCase().replace(/^N[\/,.]?A$/, 'NA').replace(/^NON APPLICABLE$/, 'NA');
    var cat = inspCategory(45 + n, st);            // mêmes règles que les fiches actuelles (25, 49, 51 neutres)
    if (cat === 'ok') f.ok++;
    else if (cat === 'na') f.na++;
    else if (cat === 'pasok') { f.pasOk++; f.pasOkLabels.push((n + '. ' + m[2]).substring(0, 70)); }
    else f.nonRens++;
  });
  for (var p = 1; p <= 53; p++) if (!seen[p]) f.nonRens++;

  // « 66296 - Cortina E-FOSS 450Wh / RECONDITIONNE 50cm 2022 … » → Cortina / E-FOSS 450Wh
  var nom = f.titre.replace(/^\s*\d{5}\s*[-–]?\s*/, '').split(/\s*\/\s*|\s+RECONDITIONN|\s+RECO\b/i)[0].trim();
  f.marque = nom.split(' ')[0] || '';
  f.modele = nom.split(' ').slice(1).join(' ');
  var d = toDate(f.date);
  f.dateIso = d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  f.sheetsUrl = 'https://docs.google.com/spreadsheets/d/' + ARCHIVE_SPREADSHEET_ID + '/edit#gid=' + f.gid;
  return f;
}

// Statut d'une fiche ancienne pour le parcours (même forme que getBikeStatus)
function archiveBikeStatus(a, photos, turnover) {
  var parcours = {};
  var t = applyTurnover(a.cs, parcours, turnover);
  var hasPh = !!(photos && photos.found && photos.count > 0);
  var counts = { ok: a.ok, pasOk: a.pasOk };
  var statutCode = computeStatut('RECONDITIONNABLE', counts, hasPh, parcours);
  return {
    enLigne: t.enLigne, vendu: t.vendu,
    exists: true, archived: true, cs: a.cs,
    date: a.date, technicien: a.technicien, typeVelo: '',
    marque: a.marque, modele: a.modele, annee: '', km: a.km,
    moteur: a.moteur, batterieWh: a.batterieWh,
    statutCode: statutCode, statutLabel: statutLabel(statutCode), statutSaisi: 'RECONDITIONNABLE',
    ancienStatut: a.ancienStatut,
    horsParcours: false, recoTerminee: inspectionTerminee(counts),
    notes: '', pasOk: a.pasOkLabels || [], nonRenseignes: a.nonRens,
    parcours: parcours, photos: photos,
    etapeKey: computeEtapeKey('RECONDITIONNABLE', parcours, hasPh, counts),
    etape: computeEtape('RECONDITIONNABLE', parcours, hasPh, counts),
    sheetsUrl: a.sheetsUrl
  };
}

// TEST — lancer depuis l'éditeur : affiche une fiche ancienne (remplacer le N°)
function testAncien() {
  var cs = '66296';
  var a = archiveFor(cs);
  Logger.log(JSON.stringify(a, null, 1));
  if (a) Logger.log(JSON.stringify(archiveBikeStatus(a, findPhotoFolder(cs), readTurnover(false)), null, 1));
}
// ============================================================
//  E-Bike Service Center — Google Apps Script v12
//  Holland Bikes / Montévrain
//
//  v5 (12/11/2026) — Summary corrigé :
//   • Date et Dernière modif. écrites comme de VRAIES dates → tri correct
//   • Tri : date d'inspection la plus récente en haut, puis N° fiche
//   • Colonne Moteur propre (plus de « NA NA », « Bafang BAFANG »…)
//   • Statut affiché en clair (« Facile à reco ») — le code reste dans la fiche
//   • PAS OK et Non renseignés > 0 surlignés en rouge / orange
//   • Filtre automatique sur l'en-tête
//   • Lecture des fiches en un seul appel par onglet → envoi plus rapide
//   • Verrou : deux envois simultanés ne se marchent plus dessus
//  v6 — Parcours du vélo :
//   • action=status&cs=…  → fiche, étapes cochées, photos Drive, étape actuelle
//   • action=setStep&cs=…&key=…&value=1|0 → coche / décoche une étape
//   • Bloc « Parcours mise en ligne » en A105:C109 de chaque fiche,
//     conservé quand la fiche est ré-enregistrée depuis la webapp
//   • Colonne « Étape parcours » dans le Summary
//  v7 — Vélos par étape :
//   • action=bikes → tous les vélos avec étape, statut, photos (filtres de la webapp)
//   • Photos comptées partout de la même façon (dossier Drive non vide)
//  v8 (13/11/2026) — Vitesse :
//   • action=get : la fiche est lue en UN appel (au lieu d'environ 150)
//   • action=list : un seul appel par onglet (au lieu de deux) + mise en cache
//   • action=bikes : mise en cache (10 min) — ajouter &fresh=1 pour forcer
//   • Photos Drive de tous les vélos mises en cache (10 min)
//   • Écriture d'une fiche en quelques appels groupés (au lieu d'environ 110)
//   • Cocher une étape ne reconstruit plus tout le Summary
//   • Le cache est vidé à chaque envoi / coche depuis la webapp
//  v9 (13/11/2026) — Vélos en ligne (fichier « Suivi Turnover Vélos Reco ») :
//   • action=online → vélos en ligne (prix, jours, baisses de prix), vendus, indicateurs
//   • « Publié en ligne » coché automatiquement quand le N° CS est dans « Velos en ligne »
//   • Nouvelle étape « Vendu » quand le N° CS est dans « Velos vendus »
//  v10 (16/11/2026) — Le site fait foi :
//   • Un vélo présent sur le site est toujours à l'étape « En ligne » (ou « Vendu »)
//   • action=bikes inclut aussi les vélos en ligne / vendus SANS fiche d'inspection
//  v11 (24/11/2026) — Parcours 100 % automatique, CS Connect, PDF batterie :
//   • Plus de cases à cocher : l'étape est calculée avec le statut global,
//     le dossier photos Drive et le site (Reconditionnement → Photos →
//     Mise en ligne → En ligne → Vendu). Les anciennes cases restent
//     stockées dans les fiches mais ne servent plus.
//   • action=prefill&cs=… → caractéristiques lues dans CycleSoftware (csconnect.gs)
//   • POST action=uploadPdf → rapport batterie PDF rangé dans le dossier
//     photos du vélo (« A mettre en ligne / N° CS »), créé s'il n'existe pas
//   • action=batteryPdf&cs=… → le PDF batterie du vélo (s'il existe)
//   • Point 25 (« Rapport de test fourni ») : NON ne compte plus comme PAS OK
//     (compté avec les NA)
//   • Vélos en ligne / vendus : lien vers la fiche produit du site (url)
//  v12 (24/11/2026) — Statuts :
//   • « Vélo en ligne » : automatique dès que le vélo est sur le site
//   • « Vélo prêt pour mise en ligne » (code PRET) : automatique quand la fiche
//     est remplie sans aucun PAS OK et que les photos sont dans le Drive
//   • « Facile à reco » supprimé → « Vélo à reconditionner » (code RECONDITIONNABLE)
//     (lancer une fois migrerStatutsV12 pour réécrire les anciennes fiches)
//   • « M2ie » supprimé de la webapp (anciennes fiches lues telles quelles)
//   • Statut affiché = statut calculé ; la fiche garde le statut choisi (B101)
//   • Points 49 et 51 : NON ne compte plus comme PAS OK (comme le point 25)
//  v13 (25/11/2026) : status renvoie recoTerminee (étape « Reconditionnement » du parcours)
//  v15 (27/11/2026) : bloc « Administration Reco » (admin.js) en A111:C125
//     de chaque fiche, action=admin (Sheets navette + CycleSoftware)
//  v14 (25/11/2026) : fiches de l'ancien Sheets lues en lecture seule (ancien.gs) :
//     parcours, « Vélos par étape » et « Vélos en ligne » (archived: true)
// ============================================================

const SPREADSHEET_ID      = '1WpjsysAbQNppB32dd6qV_XeR4Wi3rd5E_7zwl-6PX0E';
// Si tu redéploies et que l'URL change, mets aussi à jour APPS_SCRIPT_URL dans index.html
const TEMPLATE_SHEET_NAME = 'Modele Vierge (à dupliquer)';
const SUMMARY_SHEET_NAME  = 'Summary';
const WEBAPP_URL          = 'https://hollandbikes.github.io/secondlife/';
const PHOTOS_ROOT_FOLDER_ID = '10jewdRqaD8_FpQ5xRXU-GyU0JzLaPrDM'; // Drive — "A mettre en ligne"

// Fiche produit du site public à partir de l'id produit (colonne « id » de _state du fichier Turnover)
const SITE_PRODUCT_URL = 'https://www.hollandbikes.com/index.php?controller=product&id_product=';

// Rapport batterie (v11)
const BATTERY_PDF_SUFFIX = ' - Rapport batterie.pdf';
const BATTERY_PDF_MAX_MB = 20;

// Points d'inspection dont le NON ne compte pas comme PAS OK (ligne = 45 + n° du point)
//   p25 « Le rapport de test est fourni » → 70
//   p49 « Un autre marquage antivol est présent » → 94
//   p51 « Certificat de conformité et notice fournis » → 96
const NEUTRAL_NON_ROWS = [70, 94, 96];

// ── SUIVI TURNOVER (v9) — fichier alimenté par le script qui lit le site Holland Bikes
const TURNOVER_SPREADSHEET_ID = '1ykdUuUYWrnNGML7PVZSKEnv1PshUDY6zB6jU7AFMO-k';
const TURNOVER_SHEETS = {
  online:  'Velos en ligne',   // Reference | Nom | Prix | En ligne depuis | Jours en ligne
  history: 'Historique prix',  // Reference | Nom | Ancien prix | Nouveau prix | Date du changement
  sold:    'Velos vendus',     // Reference | Nom | En ligne depuis | Date de vente | Jours en ligne
  state:   '_state'            // id | name | reference | first_seen | first_seen_source | last_seen | last_price | archived
};
const CACHE_TTL_TURNOVER = 600;

// ── CACHE (v8) — durées en secondes
const CACHE_TTL_LIST   = 600;  // listes « Mes fiches » et « Vélos par étape »
const CACHE_TTL_PHOTOS = 600;  // dossiers photos Drive de tous les vélos
const CACHE_KEYS_DATA  = ['list', 'bikes']; // vidés à chaque modification

// ── PARCOURS — anciennes cases à cocher (A105:C109), conservées dans les fiches.
//    Depuis v11 elles ne servent plus au calcul de l'étape.
const PARCOURS_ROW_TITLE = 105;
const PARCOURS_STEPS = [
  { key: 'travaux', row: 106, label: 'Travaux et pièces terminés' },
  { key: 'desc',    row: 107, label: 'Description FR / EN rédigée' },
  { key: 'prix',    row: 108, label: 'Prix de vente fixé' },
  { key: 'publie',  row: 109, label: 'Publié en ligne' }
];
// Statuts qui sortent du parcours de mise en ligne (pas de photos / annonce)
const STATUTS_HORS_PARCOURS = ['UPWAY', 'M2IE', 'POUR_PIECES'];

// Libellés et couleurs des statuts (v12)
//   Choisis dans la webapp : RECONDITIONNABLE, UPWAY, POUR_PIECES, RECO_INVERSION
//   Automatiques : PRET, EN_LIGNE, VENDU
const STATUTS = {
  'EN_LIGNE':         { label: 'Vélo en ligne',                bg: '#e8f5e9' },
  'PRET':             { label: 'Vélo prêt pour mise en ligne', bg: '#e0f2f1' },
  'RECONDITIONNABLE': { label: 'Vélo à reconditionner',        bg: '#fff3e0' },
  'UPWAY':            { label: 'Upway',                        bg: '#e3f2fd' },
  'POUR_PIECES':      { label: 'Pour pièces',                  bg: '#f5f5f5' },
  'RECO_INVERSION':   { label: 'Reco inversion',               bg: '#fafafa' },
  'VENDU':            { label: 'Vendu',                        bg: '#eceff1' },
  'M2IE':             { label: 'M2ie (ancien statut)',         bg: '#ffebee' }
};

// Anciens codes → codes v12 (statut choisi dans la fiche)
//   PRET n'est plus choisi à la main : il est recalculé (fiche sans PAS OK + photos)
function normalizeStatut(code) {
  code = String(code || '').trim();
  if (code === 'FACILE_RECO' || code === 'PRET' || code === 'EN_LIGNE' || code === 'VENDU') return 'RECONDITIONNABLE';
  return code;
}

// Statut affiché (Summary, parcours, vélos par étape)
//   c = { ok, pasOk } (comptage de l'inspection), hasPhotos, p = parcours (enLigne / vendu du site)
function computeStatut(storedCode, c, hasPhotos, p) {
  var s = normalizeStatut(storedCode);
  if (p && p.enLigne) return 'EN_LIGNE';
  if (p && p.vendu) return 'VENDU';
  if (STATUTS_HORS_PARCOURS.indexOf(s) >= 0) return s;
  if (inspectionTerminee(c) && hasPhotos) return 'PRET';
  return s === 'RECO_INVERSION' ? s : (s || 'RECONDITIONNABLE');
}
// Reconditionnement terminé = fiche remplie (au moins un OK) et plus aucun PAS OK
function inspectionTerminee(c) {
  return !!(c && c.ok > 0 && c.pasOk === 0);
}
function statutLabel(code) { return (STATUTS[code] || {}).label || code; }

// Le classeur n'est ouvert qu'une fois par requête
var _ss = null;
function getSS() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}

// Onglets qui ne sont pas des fiches vélo
function isFicheName(name) {
  return name !== SUMMARY_SHEET_NAME && name !== TEMPLATE_SHEET_NAME && name.indexOf('ANCIEN') !== 0;
}

// N° CS utilisable comme nom de dossier / d'onglet (pas de / ni de \)
function cleanCs(cs) {
  return String(cs || '').trim().replace(/[\/\\]/g, '-');
}

// ============================================================
//  doPost — reçoit le JSON de la webapp (via formulaire iframe)
//           ou un PDF batterie (action=uploadPdf)
// ============================================================
function doPost(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  // Rapport batterie PDF (v11)
  if (p.action === 'uploadPdf') {
    try {
      return jsonResponse(Object.assign({ success: true }, uploadBatteryPdf(p.cs, p.data, p.name)));
    } catch (err) {
      Logger.log('uploadPdf error: ' + err.message);
      return jsonResponse({ success: false, error: err.message });
    }
  }

  var lock = LockService.getScriptLock();
  try {
    var raw = '';

    if (p.payload) {
      raw = p.payload;
    } else if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }

    if (!raw) return jsonResponse({ success: false, error: 'Aucune donnée reçue' });

    var data = JSON.parse(raw);

    // Un seul envoi à la fois (sinon deux Summary se reconstruisent en même temps)
    lock.waitLock(60000);
    var sheetName = createInspectionSheet(data);
    SpreadsheetApp.flush();
    clearDataCache();
    updateSummary();
    return jsonResponse({ success: true, sheetName: sheetName });

  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return jsonResponse({ success: false, error: err.message });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// ============================================================
//  doGet — ping / get / list / bikes / online / status / setStep /
//          prefill / batteryPdf
//  Ajouter &fresh=1 à list, bikes ou prefill pour ignorer le cache
// ============================================================
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || 'ping';
    const fresh  = params.fresh === '1' || params.fresh === 'true';

    if (action === 'ping') {
      return jsonResponse({ status: 'ok', version: 15 });
    }

    if (action === 'get') {
      const cs = (params.cs || '').toString().trim();
      if (!cs) return jsonResponse({ success: false, error: 'cs manquant' });
      const data = readSheetData(cs);
      return jsonResponse({ success: true, data: data });
    }

    if (action === 'list') {
      return jsonResponse({ success: true, list: cached('list', CACHE_TTL_LIST, listSheets, fresh) });
    }

    // Tous les vélos avec leur étape (vue « Vélos par étape »)
    if (action === 'bikes') {
      return jsonResponse({ success: true, bikes: cached('bikes', CACHE_TTL_LIST, function() { return listBikes(fresh); }, fresh) });
    }

    // Vélos en ligne sur le site + vendus + indicateurs (fichier Turnover)
    if (action === 'online') {
      return jsonResponse(Object.assign({ success: true }, getOnlineView(fresh)));
    }

    // Parcours d'un vélo : fiche, photos Drive, site
    if (action === 'status') {
      const cs = (params.cs || '').toString().trim();
      if (!cs) return jsonResponse({ success: false, error: 'cs manquant' });
      return jsonResponse({ success: true, status: getBikeStatus(cs) });
    }

    // Caractéristiques lues dans CycleSoftware (csconnect.gs)
    if (action === 'prefill') {
      const cs = (params.cs || '').toString().trim();
      if (!cs) return jsonResponse({ success: false, error: 'cs manquant' });
      return jsonResponse(Object.assign({ success: true }, getCSPrefill(cs, fresh)));
    }

    // Administration Reco : Sheets navette + CycleSoftware (admin.js) — v15
    if (action === 'admin') {
      const cs = (params.cs || '').toString().trim();
      if (!cs) return jsonResponse({ success: false, error: 'cs manquant' });
      return jsonResponse(Object.assign({ success: true }, getAdminInfo(cs, fresh)));
    }

    // Rapport batterie PDF déjà envoyé pour ce vélo ?
    if (action === 'batteryPdf') {
      const cs = (params.cs || '').toString().trim();
      if (!cs) return jsonResponse({ success: false, error: 'cs manquant' });
      return jsonResponse(Object.assign({ success: true }, findBatteryPdf(cs)));
    }

    // Ancienne action (cases du parcours) — conservée pour compatibilité
    if (action === 'setStep') {
      const cs  = (params.cs || '').toString().trim();
      const key = (params.key || '').toString();
      const done = params.value === '1' || params.value === 'true';
      if (!cs || !key) return jsonResponse({ success: false, error: 'cs ou key manquant' });
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      var status;
      try {
        setParcoursStep(cs, key, done);
        SpreadsheetApp.flush();
        status = getBikeStatus(cs);
        updateSummaryEtape(cs, status.etape);
        clearDataCache();
      } finally {
        lock.releaseLock();
      }
      return jsonResponse({ success: true, status: status });
    }

    return jsonResponse({ success: false, error: 'action inconnue: ' + action });

  } catch (err) {
    Logger.log('doGet error: ' + err.message);
    return jsonResponse({ success: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  CACHE — CacheService (100 Ko max par clé → découpage en morceaux)
// ============================================================
function cachePut(key, obj, ttl) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(obj);
    var size = 90000, parts = {}, n = Math.ceil(json.length / size) || 1;
    for (var i = 0; i < n; i++) parts[key + '_' + i] = json.substr(i * size, size);
    parts[key + '_n'] = String(n);
    cache.putAll(parts, ttl);
  } catch (e) { Logger.log('cachePut ' + key + ': ' + e.message); }
}

function cacheGet(key) {
  try {
    var cache = CacheService.getScriptCache();
    var n = parseInt(cache.get(key + '_n'), 10);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    var parts = cache.getAll(keys), json = '';
    for (var j = 0; j < n; j++) {
      if (parts[keys[j]] == null) return null; // morceau expiré → on relit tout
      json += parts[keys[j]];
    }
    return JSON.parse(json);
  } catch (e) { Logger.log('cacheGet ' + key + ': ' + e.message); return null; }
}

function cacheRemove(key) {
  try { CacheService.getScriptCache().remove(key + '_n'); } catch (e) {}
}

function cached(key, ttl, compute, fresh) {
  if (!fresh) {
    var hit = cacheGet(key);
    if (hit !== null) return hit;
  }
  var val = compute();
  cachePut(key, val, ttl);
  return val;
}

// Vidé après chaque envoi / coche depuis la webapp
function clearDataCache() {
  CACHE_KEYS_DATA.forEach(cacheRemove);
}

// ============================================================
//  CRÉATION D'UNE FICHE
// ============================================================
function createInspectionSheet(data) {
  const ss = getSS();
  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!template) throw new Error('Template introuvable: "' + TEMPLATE_SHEET_NAME + '"');

  const h      = data.header || {};
  const objCS  = (h.objetCS || '').toString().trim();
  const date   = (h.date || '').replace(/-/g, '') || formatDateCompact(new Date());
  const tech   = (h.technicien || 'INC').toString().substring(0, 10).replace(/[\s\/\\]/g, '-');
  const sheetName = objCS ? objCS : date + '-' + tech;

  // Supprimer si déjà existant — en gardant le bloc des anciennes cases du parcours
  const existing = ss.getSheetByName(sheetName);
  let parcoursSaved = null, adminSaved = null;
  if (existing) {
    parcoursSaved = existing.getRange(PARCOURS_ROW_TITLE, 1, PARCOURS_STEPS.length + 1, 3).getValues();
    if (existing.getMaxRows() >= ADMIN_ROW_TITLE + ADMIN_ROWS - 1) {
      adminSaved = existing.getRange(ADMIN_ROW_TITLE, 1, ADMIN_ROWS, 3).getValues();
      if (String(adminSaved[0][0]).indexOf('ADMINISTRATION') !== 0) adminSaved = null;
    }
    ss.deleteSheet(existing);
  }

  // Copier le template
  const newSheet = template.copyTo(ss);
  newSheet.setName(sheetName);

  // Remplir les cellules
  fillSheet(newSheet, data);

  // Bloc Parcours : restauré tel quel, ou créé vide
  if (parcoursSaved && parcoursSaved[0][0]) {
    newSheet.getRange(PARCOURS_ROW_TITLE, 1, parcoursSaved.length, 3).setValues(parcoursSaved);
  } else {
    writeParcoursBlock(newSheet);
  }

  // Bloc Administration Reco (v15) : valeurs de la webapp, sinon l'ancien bloc
  try {
    writeAdminBlock(newSheet, data.admin || null, adminSaved, objCS);
  } catch (admErr) {
    Logger.log('writeAdminBlock error: ' + admErr.message);
  }

  // Repositionner juste après le template
  try {
    const sheets    = ss.getSheets();
    const templateIdx = sheets.findIndex(function(s) { return s.getName() === TEMPLATE_SHEET_NAME; });
    if (templateIdx >= 0) {
      ss.setActiveSheet(newSheet);
      ss.moveActiveSheet(templateIdx + 2); // +2 car positions sont 1-indexées
    }
  } catch(posErr) {
    Logger.log('repositionSheet non-critical error: ' + posErr.message);
  }

  return sheetName;
}

// ============================================================
//  REMPLISSAGE DES CELLULES — écritures groupées (v8)
//  v11 : poids, couleur secondaire, poids net/brut et état batterie ne sont
//  plus saisis dans la webapp, mais leurs lignes (C11, C15, C16, C26) sont
//  gardées : la webapp renvoie les anciennes valeurs, rien ne se décale.
// ============================================================
function fillSheet(sheet, data) {
  var h    = data.header           || {};
  var c    = data.caracteristiques || {};
  var insp = data.inspection        || {};
  var gen  = (c.general    || {});
  var cad  = (c.cadre      || {});
  var elec = (c.electrique || {});
  var mec  = (c.mecanique  || {});

  function v(x) { return x || ''; }
  function col(arr) { return arr.map(function(x) { return [v(x)]; }); }

  // En-tête C2:C4
  sheet.getRange('C2:C4').setValues(col([h.date, h.technicien, h.objetCS]));
  sheet.getRange('D4').setValue(gen.typeVelo === 'MECANIQUE' ? 'Mécanique' : gen.typeVelo === 'ELECTRIQUE' ? 'Électrique' : '');

  // Caractéristiques C6:C44 (C5 = titre du template, on n'y touche pas)
  sheet.getRange('C6:C44').setValues(col([
    // Général C6–C11
    gen.numeroCadre, gen.marque, gen.modele, gen.anneeModele, gen.kilometrage, gen.poids,
    // Cadre C12–C16
    cad.materiau, cad.hauteur, cad.couleurPrincipale, cad.couleurSecondaire, cad.poidsNetBrut,
    // Électrique C17–C28
    elec.moteurMarque, elec.moteurModele, elec.moteurPosition, elec.moteurCouple,
    elec.assistanceType, elec.assistanceVitesseMax, elec.batterieCapacite, elec.batteriePosition,
    elec.batterieAmovible, elec.batterieEtat, elec.batterieCycles, elec.display,
    // Mécanique C29–C44
    mec.vitessesNombre, mec.passageVitesse, mec.derailleurType, mec.derailleurMarque, mec.derailleurModele,
    mec.freinAvantSysteme, mec.freinAvantMarque, mec.freinAvantModele,
    mec.freinArriereSysteme, mec.freinArriereMarque, mec.freinArriereModele,
    mec.pneusModele, mec.pneusTaille, mec.roueTaille, mec.fourche, mec.bequille
  ]));

  // Inspection — 53 points : col D = Statut, col F = État (lignes 46 à 98)
  // E et G = rappels visuels du template, on n'y touche pas
  var statuts = [], aspects = [];
  for (var i = 1; i <= 53; i++) {
    var val = insp['p' + i] || {};
    statuts.push([(val.statut || '').toString().trim()]);
    aspects.push([(val.aspect !== undefined && val.aspect !== null) ? String(val.aspect).trim() : '']);
  }
  sheet.getRange('D46:D98').clearDataValidations().setValues(statuts);
  sheet.getRange('F46:F98').clearDataValidations().setValues(aspects);

  // Notes, statut global, dernière modif., lien retour webapp (A100:B103)
  var csVal   = (h.objetCS || '').toString().trim();
  var editUrl = WEBAPP_URL + '?cs=' + encodeURIComponent(csVal);
  sheet.getRange('A100:B103').setValues([
    ['Notes',                 v(data.notes)],
    ['Statut global',         v(data.statutGlobal)],
    ['Dernière modification', new Date()],
    ['✏ Modifier cette fiche', editUrl]
  ]);
  sheet.getRange('B102').setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange('A103:B103').setFontColor('#00335b').setFontWeight('bold');
  sheet.getRange('B103').setFontLine('underline');
}

// Comptage d'un statut d'inspection (ligne de la fiche) → 'ok' | 'pasok' | 'na' | ''
//   Le NON des lignes « neutres » (point 25) compte comme NA.
function inspCategory(row, statut) {
  var st = String(statut || '').trim().toUpperCase();
  if (st === 'OK' || st === 'OUI') return 'ok';
  if (st === 'PAS OK' || st === 'NON') return NEUTRAL_NON_ROWS.indexOf(row) >= 0 && st === 'NON' ? 'na' : 'pasok';
  if (st === 'NA') return 'na';
  return '';
}

// ============================================================
//  PHOTOS — dossier Drive correspondant au N° Objet CS
//  Cherche, dans "A mettre en ligne" puis dans son sous-dossier
//  "Archive", un dossier dont le nom est identique au N° Objet CS.
// ============================================================
// Index { nom du dossier → Folder } (racine avant Archive)
function buildPhotoFolderIndex() {
  var index = {};
  var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
  addFoldersToPhotoIndex(root, index);
  var archiveIt = root.getFoldersByName('Archive');
  if (archiveIt.hasNext()) addFoldersToPhotoIndex(archiveIt.next(), index);
  return index;
}

function addFoldersToPhotoIndex(parentFolder, index) {
  var it = parentFolder.getFolders();
  while (it.hasNext()) {
    var folder = it.next();
    var name = folder.getName().trim();
    if (name === 'Archive' || name === 'TEMP') continue;
    if (!(name in index)) index[name] = folder;
  }
}

// Nombre d'images dans un dossier (plafonné) — le PDF batterie n'est pas compté
function countImages(folder) {
  var count = 0, files = folder.getFiles();
  while (files.hasNext() && count < 500) {
    if (/^image\//.test(files.next().getMimeType())) count++;
  }
  return count;
}

// Photos de TOUS les dossiers Drive, en un passage — mis en cache (v8)
//   { nom du dossier → { url, count } }
function allPhotoFolders(fresh) {
  return cached('photos', CACHE_TTL_PHOTOS, function() {
    var out = {};
    try {
      var index = buildPhotoFolderIndex();
      Object.keys(index).forEach(function(name) {
        out[name] = { url: index[name].getUrl(), count: countImages(index[name]) };
      });
    } catch (e) { Logger.log('allPhotoFolders: ' + e.message); }
    return out;
  }, fresh);
}

// Photos d'une liste de N° CS
function photosForNames(names, fresh) {
  var all = allPhotoFolders(fresh);
  var out = {};
  names.forEach(function(n) {
    var p = all[n];
    out[n] = p ? { found: true, url: p.url, count: p.count }
               : { found: false, url: '', count: 0 };
  });
  return out;
}

// Dossier Drive d'un vélo (racine puis Archive) — Folder ou null
function getPhotoFolderObj(cs) {
  var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
  var parents = [root];
  var archiveIt = root.getFoldersByName('Archive');
  if (archiveIt.hasNext()) parents.push(archiveIt.next());
  for (var i = 0; i < parents.length; i++) {
    var it = parents[i].getFoldersByName(cs);
    if (it.hasNext()) return { folder: it.next(), archived: i === 1 };
  }
  return null;
}

// Dossier photos d'un seul vélo — toujours lu en direct (photos fraîches dans le Parcours)
function findPhotoFolder(cs) {
  var f = getPhotoFolderObj(cs);
  if (!f) return { found: false, url: '', count: 0, archived: false };
  return { found: true, url: f.folder.getUrl(), count: countImages(f.folder), archived: f.archived };
}

// ============================================================
//  RAPPORT BATTERIE PDF (v11)
//  Rangé dans le même dossier que les photos (« A mettre en ligne / N° CS »),
//  créé s'il n'existe pas. Un nouvel envoi remplace l'ancien PDF.
// ============================================================
function uploadBatteryPdf(cs, base64, originalName) {
  cs = cleanCs(cs);
  if (!cs) throw new Error('N° CS manquant');
  if (!base64) throw new Error('Fichier vide');
  var bytes = Utilities.base64Decode(String(base64).replace(/^data:[^,]*,/, ''));
  if (bytes.length > BATTERY_PDF_MAX_MB * 1024 * 1024) throw new Error('PDF trop lourd (max ' + BATTERY_PDF_MAX_MB + ' Mo)');
  // Vérifie la signature PDF (%PDF)
  if (!(bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46)) throw new Error('Ce fichier n\'est pas un PDF');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var found = getPhotoFolderObj(cs);
    var created = false, folder;
    if (found) folder = found.folder;
    else { folder = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID).createFolder(cs); created = true; }

    var name = cs + BATTERY_PDF_SUFFIX;
    var old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);

    var file = folder.createFile(Utilities.newBlob(bytes, MimeType.PDF, name));
    if (originalName) file.setDescription('Fichier d\'origine : ' + originalName + ' — envoyé depuis la webapp Second Life');
    cacheRemove('photos');
    clearDataCache();
    return { found: true, name: name, url: file.getUrl(), folderUrl: folder.getUrl(),
             folderCreated: created, archived: !!(found && found.archived),
             date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') };
  } finally {
    lock.releaseLock();
  }
}

function findBatteryPdf(cs) {
  cs = cleanCs(cs);
  var f = getPhotoFolderObj(cs);
  if (!f) return { found: false, folderFound: false };
  var it = f.folder.getFilesByName(cs + BATTERY_PDF_SUFFIX);
  if (!it.hasNext()) return { found: false, folderFound: true, folderUrl: f.folder.getUrl() };
  var file = it.next();
  return { found: true, folderFound: true, name: file.getName(), url: file.getUrl(), folderUrl: f.folder.getUrl(),
           archived: f.archived,
           date: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') };
}

// ============================================================
//  PARCOURS — anciennes cases (conservées pour compatibilité)
// ============================================================
function writeParcoursBlock(sheet) {
  var rows = [['Parcours mise en ligne', 'Fait', 'Le']];
  PARCOURS_STEPS.forEach(function(s) { rows.push([s.label, '', '']); });
  sheet.getRange(PARCOURS_ROW_TITLE, 1, rows.length, 3).setValues(rows);
  sheet.getRange(PARCOURS_ROW_TITLE, 1, 1, 3).setFontWeight('bold').setFontColor('#00335b');
  sheet.getRange(PARCOURS_ROW_TITLE + 1, 3, PARCOURS_STEPS.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
}

function setParcoursStep(cs, key, done) {
  var step = PARCOURS_STEPS.filter(function(s) { return s.key === key; })[0];
  if (!step) throw new Error('Étape inconnue : ' + key);
  var sheet = getSS().getSheetByName(cs);
  if (!sheet) throw new Error('Fiche "' + cs + '" introuvable — remplis d\'abord la fiche d\'inspection');
  if (!sheet.getRange(PARCOURS_ROW_TITLE, 1).getValue()) writeParcoursBlock(sheet);
  sheet.getRange(step.row, 2, 1, 2).setValues([[done ? '✓' : '', done ? new Date() : '']]);
}

// Lit les anciennes cases depuis une grille de valeurs (A1:C109) d'une fiche.
function readParcours(grid, statutCode) {
  var hasBlock = !!(grid[PARCOURS_ROW_TITLE - 1] && grid[PARCOURS_ROW_TITLE - 1][0]);
  var out = {};
  PARCOURS_STEPS.forEach(function(s) {
    var row = grid[s.row - 1] || [];
    out[s.key] = hasBlock ? !!String(row[1] || '').trim() : (statutCode === 'PRET');
    out[s.key + 'At'] = (row[2] instanceof Date)
      ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
  });
  return out;
}

// Étape actuelle du vélo (même logique que la webapp) — v12 : 100 % automatique
//   clé : 'vendu' | 'enligne' | 'hors' | 'reco' | 'photos' | 'miseenligne'
//   • sur le site → en ligne / vendu (le site fait foi)
//   • statut Upway / M2ie / Pour pièces → hors parcours
//   • fiche avec des PAS OK (ou vide) → reconditionnement
//   • reconditionnement fini, pas de photos → photos
//   • reconditionnement fini + photos → mise en ligne (statut « prêt pour mise en ligne »)
function computeEtapeKey(statutCode, p, hasPhotos, c) {
  if (p.vendu) return 'vendu';
  if (p.enLigne) return 'enligne';
  if (STATUTS_HORS_PARCOURS.indexOf(normalizeStatut(statutCode)) >= 0) return 'hors';
  if (!inspectionTerminee(c)) return 'reco';
  if (!hasPhotos) return 'photos';
  return 'miseenligne';
}
function computeEtape(statutCode, p, hasPhotos, c) {
  var key = computeEtapeKey(statutCode, p, hasPhotos, c);
  if (key === 'hors') return statutLabel(normalizeStatut(statutCode));
  return { reco: 'Reconditionnement', photos: 'Photos', miseenligne: 'Mise en ligne', enligne: 'En ligne', vendu: '✓ Vendu' }[key];
}

// Ajoute au parcours ce que dit le fichier Turnover (site Holland Bikes)
function applyTurnover(cs, parcours, turnover) {
  var online = turnover.onlineByRef[cs] || null;
  var sold   = turnover.soldByRef[cs] || null;
  parcours.publieAuto = !!(online || sold);
  if (parcours.publieAuto) {
    parcours.publie = true;
    if (!parcours.publieAt) parcours.publieAt = (online || sold).depuis || '';
  }
  parcours.vendu = !!sold && !online;   // remis en ligne après une vente → on le considère en ligne
  parcours.enLigne = !!online;
  parcours.siteAuto = parcours.publieAuto;
  if (parcours.siteAuto) {
    ['travaux', 'desc', 'prix'].forEach(function(k) { parcours[k] = true; });
  }
  return { enLigne: online, vendu: parcours.vendu ? sold : null };
}

// Tous les vélos (une fiche = un vélo) pour la vue « Vélos par étape »
function listBikes(fresh) {
  var ss = getSS();
  var fiches = [];
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (!isFicheName(name)) return;
    var f = extractSummaryRow(sheet, name);
    if (f) fiches.push(f);
  });
  var archive = {};
  try { archive = getArchive(); } catch (e) { Logger.log('getArchive: ' + e.message); }
  var current = {};
  fiches.forEach(function(f) { current[f.name] = true; });
  var oldOnly = Object.keys(archive).filter(function(cs) { return !current[cs]; });
  var photos = photosForNames(fiches.map(function(f) { return f.name; }).concat(oldOnly), fresh);
  var turnover = readTurnover(fresh);
  var tz = Session.getScriptTimeZone();
  var out = fiches.map(function(f) {
    var ph = photos[f.name];
    var t = applyTurnover(f.name, f.parcours, turnover);
    var hasPh = ph.found && ph.count > 0;
    var counts = { ok: f.ok, pasOk: f.pasOk };
    var eff = computeStatut(f.statutCode, counts, hasPh, f.parcours);
    return {
      enLigne:      t.enLigne ? { prix: t.enLigne.prix, jours: t.enLigne.jours, approx: t.enLigne.approx, url: t.enLigne.url } : null,
      vendu:        t.vendu ? { dateVente: t.vendu.dateVente, jours: t.vendu.jours } : null,
      cs:           f.name,
      date:         f.date ? Utilities.formatDate(f.date, tz, 'dd/MM/yyyy') : '',
      dateIso:      f.date ? Utilities.formatDate(f.date, tz, 'yyyy-MM-dd') : '',
      technicien:   f.technicien,
      marque:       f.marque,
      modele:       f.modele,
      typeVelo:     f.typeVelo,
      statutCode:   eff,
      statutLabel:  statutLabel(eff),
      statutSaisi:  normalizeStatut(f.statutCode),
      horsParcours: STATUTS_HORS_PARCOURS.indexOf(eff) >= 0,
      parcours:     f.parcours,
      photos:       ph,
      pasOk:        f.pasOk,
      etapeKey:     computeEtapeKey(f.statutCode, f.parcours, hasPh, counts),
      hasFiche:     true
    };
  });
  // Fiches de l'ancien Sheets (absentes du Sheets actuel)
  oldOnly.forEach(function(cs) {
    var st = archiveBikeStatus(archive[cs], photos[cs], turnover);
    out.push({
      enLigne: st.enLigne ? { prix: st.enLigne.prix, jours: st.enLigne.jours, approx: st.enLigne.approx, url: st.enLigne.url } : null,
      vendu:   st.vendu ? { dateVente: st.vendu.dateVente, jours: st.vendu.jours } : null,
      cs: cs, date: archive[cs].date, dateIso: archive[cs].dateIso || '',
      technicien: archive[cs].technicien, marque: st.marque, modele: st.modele, typeVelo: '',
      statutCode: st.statutCode, statutLabel: st.statutLabel, statutSaisi: 'RECONDITIONNABLE',
      horsParcours: false, parcours: st.parcours, photos: photos[cs], pasOk: archive[cs].pasOk,
      etapeKey: st.etapeKey, hasFiche: true, archived: true
    });
  });

  // Vélos du site sans fiche d'inspection (mis en ligne avant l'app, ou référence ≠ N° CS)
  var known = {};
  out.forEach(function(b) { known[b.cs] = true; });
  function siteOnly(b, isSold) {
    return {
      cs: b.cs, nom: b.nom, marque: '', modele: '', technicien: '', typeVelo: '',
      date: b.depuis || '', dateIso: b.depuisIso || '',
      statutCode: isSold ? 'VENDU' : 'EN_LIGNE', statutLabel: statutLabel(isSold ? 'VENDU' : 'EN_LIGNE'), horsParcours: false,
      parcours: {}, photos: { found: false, url: '', count: 0 }, pasOk: 0,
      enLigne: isSold ? null : { prix: b.prix, jours: b.jours, approx: b.approx, url: b.url },
      vendu:   isSold ? { dateVente: b.dateVente, jours: b.jours } : null,
      etapeKey: isSold ? 'vendu' : 'enligne',
      hasFiche: false
    };
  }
  (turnover.online || []).forEach(function(b) {
    if (!known[b.cs]) { known[b.cs] = true; out.push(siteOnly(b, false)); }
  });
  Object.keys(turnover.soldByRef || {}).forEach(function(cs) {
    if (!known[cs]) { known[cs] = true; out.push(siteOnly(turnover.soldByRef[cs], true)); }
  });
  return out;
}

// Tout ce que la webapp affiche dans le « Parcours du vélo »
function getBikeStatus(cs) {
  var ss = getSS();
  var sheet = ss.getSheetByName(cs);
  var photos = { found: false, url: '', count: 0 };
  try { photos = findPhotoFolder(cs); } catch (e) { Logger.log('findPhotoFolder: ' + e.message); }

  var turnover = { onlineByRef: {}, soldByRef: {} };
  try { turnover = readTurnover(false); } catch (e) { Logger.log('readTurnover: ' + e.message); }

  if (!sheet || !isFicheName(cs)) {
    var old = archiveFor(cs);                       // ancien Sheets (lecture seule)
    if (old) return archiveBikeStatus(old, photos, turnover);
    var onl = turnover.onlineByRef[cs] || null;
    return { exists: false, cs: cs, photos: photos,
             enLigne: onl, vendu: onl ? null : (turnover.soldByRef[cs] || null) };
  }

  var grid = sheet.getRange(1, 1, PARCOURS_STEPS[PARCOURS_STEPS.length - 1].row, 7).getValues();
  function v(r, c) {
    var x = grid[r - 1][c - 1];
    if (x instanceof Date) return Utilities.formatDate(x, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    return (x === null || x === undefined) ? '' : String(x).trim();
  }

  var pasOk = [], nonRens = 0, nbOk = 0;
  for (var r = 46; r <= 98; r++) {
    var cat = inspCategory(r, v(r, 4));
    if (cat === 'pasok') pasOk.push(v(r, 3));
    else if (cat === 'ok') nbOk++;
    else if (!cat) nonRens++;
  }

  var statutSaisi = v(101, 2);
  var parcours = readParcours(grid, statutSaisi);
  var t = applyTurnover(cs, parcours, turnover);
  var hasPh = photos.found && photos.count > 0;
  var counts = { ok: nbOk, pasOk: pasOk.length };
  var statutCode = computeStatut(statutSaisi, counts, hasPh, parcours);

  return {
    enLigne:     t.enLigne,
    vendu:       t.vendu,
    exists:      true,
    cs:          cs,
    date:        v(2, 3),
    technicien:  v(3, 3),
    typeVelo:    v(4, 4),
    marque:      v(7, 3),
    modele:      v(8, 3),
    annee:       v(9, 3),
    km:          v(10, 3),
    moteur:      buildMoteur(v(17, 3), v(18, 3), v(4, 4)),
    batterieWh:  v(23, 3),
    statutCode:  statutCode,
    statutLabel: statutLabel(statutCode),
    statutSaisi: normalizeStatut(statutSaisi),
    horsParcours: STATUTS_HORS_PARCOURS.indexOf(statutCode) >= 0,
    recoTerminee: inspectionTerminee(counts),   // v13 : plus aucun PAS OK (et au moins un OK)
    notes:       v(100, 2),
    pasOk:       pasOk,
    nonRenseignes: nonRens,
    parcours:    parcours,
    photos:      photos,
    etapeKey:    computeEtapeKey(statutSaisi, parcours, hasPh, counts),
    etape:       computeEtape(statutSaisi, parcours, hasPh, counts),
    sheetsUrl:   'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit#gid=' + sheet.getSheetId()
  };
}

// ============================================================
//  VÉLOS EN LIGNE — fichier « Suivi Turnover Vélos Reco » (v9)
//  Lecture seule : le script de ce fichier reste seul à l'alimenter.
//  « Reference » = N° Objet CS (nom de l'onglet de la fiche d'inspection).
// ============================================================
function parsePrix(v) {
  if (typeof v === 'number') return v;
  var s = String(v || '').replace(/[^\d.,-]/g, '');
  if (!s) return null;
  if (s.indexOf('.') >= 0) s = s.replace(/,/g, '');        // « 1,199.00 » → 1199.00
  else s = s.replace(/\s/g, '').replace(',', '.');           // « 1199,00 »  → 1199.00
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function isoDate(v) {
  var d = toDate(v);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}
function frDate(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
}
function daysBetween(isoA, isoB) {
  if (!isoA || !isoB) return null;
  var a = new Date(isoA + 'T00:00:00Z'), b = new Date(isoB + 'T00:00:00Z');
  return Math.round((b - a) / 86400000);
}

function readTurnover(fresh) {
  try {
    return cached('turnover', CACHE_TTL_TURNOVER, readTurnoverFromSheet, fresh);
  } catch (e) {
    Logger.log('readTurnover: ' + e.message);
    return { online: [], sold: [], onlineByRef: {}, soldByRef: {}, error: e.message };
  }
}

function readTurnoverFromSheet() {
  var ss = SpreadsheetApp.openById(TURNOVER_SPREADSHEET_ID);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  function rows(name, keyCol) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(5, sh.getLastColumn())).getValues()
      .filter(function(r) { return String(r[keyCol || 0]).trim() !== ''; });
  }
  function ref(v) { return String(v).trim().replace(/\.0+$/, ''); }

  // _state : date de 1re apparition « au plus tard » + id produit du site (lien vers la fiche)
  //   Plusieurs lignes pour une même référence (remis en ligne) → la plus récente (last_seen) l'emporte
  var approxRefs = {}, productByRef = {};
  rows(TURNOVER_SHEETS.state, 2).forEach(function(r) {
    var k = ref(r[2]);
    if (String(r[4]).trim() === 'wayback_lower_bound') approxRefs[k] = true;
    var id = ref(r[0]), seen = isoDate(r[5]);
    if (id && (!productByRef[k] || seen >= productByRef[k].seen)) productByRef[k] = { id: id, seen: seen };
  });
  function productUrl(k) { return productByRef[k] ? SITE_PRODUCT_URL + encodeURIComponent(productByRef[k].id) : ''; }

  // Historique des prix, du plus ancien au plus récent
  var histByRef = {};
  rows(TURNOVER_SHEETS.history).forEach(function(r) {
    var k = ref(r[0]);
    (histByRef[k] = histByRef[k] || []).push({
      dateIso: isoDate(r[4]), ancien: parsePrix(r[2]), nouveau: parsePrix(r[3])
    });
  });
  Object.keys(histByRef).forEach(function(k) {
    histByRef[k].sort(function(a, b) { return a.dateIso.localeCompare(b.dateIso); });
    histByRef[k].forEach(function(h) { h.date = frDate(h.dateIso); });
  });

  var online = rows(TURNOVER_SHEETS.online).map(function(r) {
    var k = ref(r[0]);
    var depuisIso = isoDate(r[3]);
    var jours = depuisIso ? daysBetween(depuisIso, today) : (typeof r[4] === 'number' ? r[4] : parseInt(r[4], 10) || null);
    var histo = histByRef[k] || [];
    var last = histo.length ? histo[histo.length - 1] : null;
    return {
      cs: k, nom: String(r[1]).trim(), prix: parsePrix(r[2]), url: productUrl(k),
      depuisIso: depuisIso, depuis: frDate(depuisIso), jours: jours, approx: !!approxRefs[k],
      prixInitial: histo.length ? histo[0].ancien : parsePrix(r[2]),
      nbChangements: histo.length,
      dernierChangement: last ? { date: last.date, dateIso: last.dateIso, ancien: last.ancien, nouveau: last.nouveau,
                                  joursDepuis: daysBetween(last.dateIso, today) } : null,
      histo: histo
    };
  });

  var sold = rows(TURNOVER_SHEETS.sold).map(function(r) {
    var k = ref(r[0]);
    var depuisIso = isoDate(r[2]), venteIso = isoDate(r[3]);
    var jours = (depuisIso && venteIso) ? daysBetween(depuisIso, venteIso)
              : (typeof r[4] === 'number' ? r[4] : null);
    var histo = histByRef[k] || [];
    return {
      cs: k, nom: String(r[1]).trim(),
      depuisIso: depuisIso, depuis: frDate(depuisIso),
      dateVenteIso: venteIso, dateVente: frDate(venteIso),
      jours: jours, approx: !!approxRefs[k],
      nbChangements: histo.length,
      dernierPrix: histo.length ? histo[histo.length - 1].nouveau : null,
      histo: histo
    };
  });

  var onlineByRef = {}, soldByRef = {};
  online.forEach(function(b) { onlineByRef[b.cs] = b; });
  sold.forEach(function(b) {
    var prev = soldByRef[b.cs];
    if (!prev || (b.dateVenteIso || '') > (prev.dateVenteIso || '')) soldByRef[b.cs] = b;
  });
  return { online: online, sold: sold, onlineByRef: onlineByRef, soldByRef: soldByRef, updatedAt: today };
}

// Données de l'écran « Vélos en ligne »
function getOnlineView(fresh) {
  var t = readTurnover(fresh);
  if (t.error) throw new Error('Fichier Turnover illisible : ' + t.error);

  var fiches = {};
  getSS().getSheets().forEach(function(s) { fiches[s.getName()] = true; });
  var archive = {};
  try { archive = getArchive(); } catch (e) { Logger.log('getArchive: ' + e.message); }
  [t.online, t.sold].forEach(function(list) {
    list.forEach(function(b) {
      b.ficheArchive = !fiches[b.cs] && !!archive[b.cs];
      b.hasFiche = !!fiches[b.cs] || b.ficheArchive;
    });
  });

  var ventesConnues = t.sold.filter(function(b) { return typeof b.jours === 'number'; });
  var joursVente = ventesConnues.map(function(b) { return b.jours; });
  var totalPrix = t.online.reduce(function(s, b) { return s + (b.prix || 0); }, 0);
  return {
    online: t.online,
    sold: t.sold,
    updatedAt: t.updatedAt,
    stats: {
      nbOnline:  t.online.length,
      nbOrange:  t.online.filter(function(b) { return b.jours >= 45 && b.jours <= 90; }).length,
      nbRouge:   t.online.filter(function(b) { return b.jours > 90; }).length,
      valeurEnLigne: Math.round(totalPrix),
      nbVendus:  t.sold.length,
      nbVendusAvecDelai: joursVente.length,
      delaiMoyenVente: joursVente.length ? Math.round(joursVente.reduce(function(a, b) { return a + b; }, 0) / joursVente.length) : null,
      delaiApprox: ventesConnues.some(function(b) { return b.approx; })
    },
    sheetsUrl: 'https://docs.google.com/spreadsheets/d/' + TURNOVER_SPREADSHEET_ID + '/edit'
  };
}

// ============================================================
//  SUMMARY — reconstruit à chaque soumission
// ============================================================
var SUMMARY_HEADERS = [
  'N° Fiche', '📷 Photos', '✏ Modifier', '📊 Fiche Sheets', 'Statut global', 'Étape parcours',
  'Date', 'Technicien', 'N° Objet CS',
  'Marque', 'Modèle', 'Type de vélo', 'N° Cadre', 'Année', 'Km',
  'Moteur', 'Batterie (Wh)', 'État batterie', 'Cycles',
  '✓ OK', '✗ PAS OK', '— NA', 'Non renseignés',
  'Dernière modif.', 'Notes'
];
var SUMMARY_WIDTHS = [80, 95, 85, 105, 125, 140, 90, 100, 100, 110, 160, 95, 150, 60, 65, 170, 90, 90, 65, 50, 70, 50, 95, 125, 260];
var SUMMARY_COL_ETAPE = 6;

// Met à jour la seule cellule « Étape parcours » d'un vélo
function updateSummaryEtape(cs, etape) {
  var summary = getSS().getSheetByName(SUMMARY_SHEET_NAME);
  if (!summary || summary.getLastRow() < 2) return;
  var names = summary.getRange(2, 1, summary.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === cs) {
      summary.getRange(i + 2, SUMMARY_COL_ETAPE).setValue(etape);
      return;
    }
  }
}

function updateSummary(fresh) {
  var ss = getSS();
  var summary = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summary) {
    summary = ss.insertSheet(SUMMARY_SHEET_NAME);
  } else {
    if (summary.getFilter()) summary.getFilter().remove();
    summary.clear(); // contenu + formats
  }
  var nCols = SUMMARY_HEADERS.length;

  // ── En-tête ────────────────────────────────────────────────
  summary.getRange(1, 1, 1, nCols).setValues([SUMMARY_HEADERS])
    .setBackground('#00335b').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setWrap(true);
  summary.setRowHeight(1, 32);
  summary.setFrozenRows(1);
  summary.setFrozenColumns(1);
  SUMMARY_WIDTHS.forEach(function(w, i) { summary.setColumnWidth(i + 1, w); });

  // ── Collecte des fiches ────────────────────────────────────
  var fiches = [];
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (!isFicheName(name)) return;
    var f = extractSummaryRow(sheet, name);
    if (f) fiches.push(f);
  });

  // ── Dossiers photos (Drive, via cache) : lien + étape ──────
  var photos = photosForNames(fiches.map(function(f) { return f.name; }), fresh);
  var turnover = { onlineByRef: {}, soldByRef: {} };
  try { turnover = readTurnover(fresh); } catch (e) { Logger.log('readTurnover: ' + e.message); }
  fiches.forEach(function(f) {
    var ph = photos[f.name];
    applyTurnover(f.name, f.parcours, turnover);
    var hasPh = ph.found && ph.count > 0;
    var counts = { ok: f.ok, pasOk: f.pasOk };
    f.photoUrl = ph.url;
    f.etape = computeEtape(f.statutCode, f.parcours, hasPh, counts);
    f.statutCode = computeStatut(f.statutCode, counts, hasPh, f.parcours);   // statut affiché (calculé)
    f.statutLabel = statutLabel(f.statutCode);
  });

  fiches.sort(function(a, b) {
    var da = a.date ? a.date.getTime() : 0, db = b.date ? b.date.getTime() : 0;
    if (db !== da) return db - da;
    var ma = a.derniereModif ? a.derniereModif.getTime() : 0, mb = b.derniereModif ? b.derniereModif.getTime() : 0;
    if (mb !== ma) return mb - ma;
    return String(b.name).localeCompare(String(a.name));
  });

  if (fiches.length === 0) {
    ss.setActiveSheet(summary);
    ss.moveActiveSheet(1);
    return;
  }

  var n = fiches.length;

  var values = fiches.map(function(f) {
    return [
      f.name, '', '', '', f.statutLabel, f.etape,
      f.date || '', f.technicien, f.objetCS,
      f.marque, f.modele, f.typeVelo, f.numeroCadre, f.annee, f.km,
      f.moteur, f.batterieWh, f.batterieEtat, f.cycles,
      f.ok, f.pasOk, f.na, f.nonRens,
      f.derniereModif || '', f.notes
    ];
  });
  var data = summary.getRange(2, 1, n, nCols);
  data.setValues(values);

  function link(text, url) {
    var b = SpreadsheetApp.newRichTextValue().setText(url ? text : '');
    if (url) b.setLinkUrl(url);
    return b.build();
  }
  summary.getRange(2, 2, n, 3).setRichTextValues(fiches.map(function(f) {
    return [link('Voir photos', f.photoUrl), link('Modifier', f.editUrl), link('Voir fiche', f.sheetsUrl)];
  }));

  data.setFontSize(10).setVerticalAlignment('middle');
  data.setBackgrounds(fiches.map(function(f) {
    var bg = (STATUTS[f.statutCode] || {}).bg || '#ffffff';
    return SUMMARY_HEADERS.map(function() { return bg; });
  }));
  summary.getRange(2, 1, n, 1).setFontWeight('bold').setHorizontalAlignment('left');
  summary.getRange(2, 2, n, 1).setFontColor('#1a73e8').setFontWeight('bold');
  summary.getRange(2, 3, n, 1).setFontColor('#00335b').setFontWeight('bold');
  summary.getRange(2, 4, n, 1).setFontColor('#1a73e8').setFontWeight('bold');
  summary.getRange(2, 5, n, 1).setFontWeight('bold');
  summary.getRange(2, 6, n, 1).setFontWeight('bold').setFontColor('#00335b');
  summary.getRange(2, 7, n, 1).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');
  summary.getRange(2, 20, n, 4).setNumberFormat('0').setHorizontalAlignment('center');
  summary.getRange(2, 24, n, 1).setNumberFormat('dd/mm/yyyy hh:mm').setHorizontalAlignment('center');
  summary.getRange(2, 25, n, 1).setWrap(true);

  var rangePasOk   = summary.getRange(2, 21, n, 1);
  var rangeNonRens = summary.getRange(2, 23, n, 1);
  summary.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#c62828').setBold(true).setBackground('#ffcdd2')
      .setRanges([rangePasOk]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#e65100').setBold(true).setBackground('#ffe0b2')
      .setRanges([rangeNonRens]).build()
  ]);

  summary.getRange(1, 1, n + 1, nCols).createFilter();

  ss.setActiveSheet(summary);
  ss.moveActiveSheet(1);
}

// ── Extraction d'une fiche pour le Summary ────────────────────
function extractSummaryRow(sheet, sheetName) {
  try {
    var grid = sheet.getRange(1, 1, PARCOURS_STEPS[PARCOURS_STEPS.length - 1].row, 7).getValues();
    var COLS = { A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6 };
    function raw(a1) {
      var m = a1.match(/^([A-G])(\d+)$/);
      return grid[parseInt(m[2], 10) - 1][COLS[m[1]]];
    }
    function g(a1) {
      var v = raw(a1);
      if (v === null || v === undefined || v === '') return '';
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      return String(v).trim();
    }

    // Compter les résultats d'inspection (col D, lignes 46–98)
    var ok = 0, pasOk = 0, na = 0, nonRens = 0;
    for (var r = 46; r <= 98; r++) {
      var cat = inspCategory(r, g('D' + r));
      if (cat === 'ok')         ok++;
      else if (cat === 'pasok') pasOk++;
      else if (cat === 'na')    na++;
      else                      nonRens++;
    }

    var typeVelo   = g('D4');
    var statutCode = g('B101');

    return {
      name:          sheetName,
      editUrl:       g('B103'),
      sheetsUrl:     'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit#gid=' + sheet.getSheetId(),
      statutCode:    statutCode,
      statutLabel:   statutLabel(normalizeStatut(statutCode)),
      parcours:      readParcours(grid, statutCode),
      date:          toDate(raw('C2')),
      technicien:    g('C3'),
      objetCS:       g('C4'),
      marque:        g('C7'),
      modele:        g('C8'),
      typeVelo:      typeVelo,
      numeroCadre:   g('C6'),
      annee:         g('C9'),
      km:            g('C10'),
      moteur:        buildMoteur(g('C17'), g('C18'), typeVelo),
      batterieWh:    g('C23'),
      batterieEtat:  g('C26'),
      cycles:        g('C27'),
      ok: ok, pasOk: pasOk, na: na, nonRens: nonRens,
      derniereModif: toDate(raw('B102')),
      notes:         g('B100')
    };
  } catch (e) {
    Logger.log('extractSummaryRow error for ' + sheetName + ': ' + e.message);
    return null;
  }
}

// Moteur lisible : ignore les « NA », évite les doublons (« Bafang BAFANG »)
function buildMoteur(marque, modele, typeVelo) {
  function clean(s) {
    s = String(s || '').trim();
    return (s.toUpperCase() === 'NA' || s === '-') ? '' : s;
  }
  if (typeVelo === 'Mécanique') return '—';
  marque = clean(marque);
  modele = clean(modele);
  if (!marque) return modele;
  if (!modele) return marque;
  var mU = marque.toUpperCase(), oU = modele.toUpperCase();
  if (oU === mU) return marque;
  if (oU.indexOf(mU + ' ') === 0) return modele;
  return marque + ' ' + modele;
}

// Convertit une valeur de cellule en vraie Date (ou '' si impossible)
function toDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v || '').trim();
  var m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) {
    return new Date(+m[1], +m[2] - 1, +m[3]);
  }
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/))) {
    return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0));
  }
  return '';
}

// ============================================================
//  LECTURE D'UNE FICHE (pour rechargement dans la webapp)
// ============================================================
function readSheetData(cs) {
  var sheet = getSS().getSheetByName(cs);
  if (!sheet) throw new Error('Fiche "' + cs + '" introuvable');

  var grid = sheet.getRange(1, 1, 101, 6).getValues(); // A1:F101
  var COLS = { A: 0, B: 1, C: 2, D: 3, E: 4, F: 5 };
  var tz = Session.getScriptTimeZone();

  function g(cell) {
    var m = cell.match(/^([A-F])(\d+)$/);
    var val = grid[parseInt(m[2], 10) - 1][COLS[m[1]]];
    if (val === null || val === undefined || val === '') return '';
    if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
    return String(val).trim();
  }

  var inspection = {};
  for (var i = 1; i <= 53; i++) {
    var r = 45 + i;
    var statut = g('D' + r);
    var aspect = g('F' + r);
    if (statut || aspect) inspection['p' + i] = { statut: statut, aspect: aspect };
  }

  return {
    header: {
      date:       g('C2'),
      technicien: g('C3'),
      objetCS:    g('C4')
    },
    caracteristiques: {
      general: {
        typeVelo: g('D4') === 'Mécanique' ? 'MECANIQUE' : g('D4') === 'Électrique' ? 'ELECTRIQUE' : '',
        numeroCadre: g('C6'), marque: g('C7'), modele: g('C8'),
        anneeModele: g('C9'), kilometrage: g('C10'), poids: g('C11')
      },
      cadre: {
        materiau: g('C12'), hauteur: g('C13'),
        couleurPrincipale: g('C14'), couleurSecondaire: g('C15'), poidsNetBrut: g('C16')
      },
      electrique: {
        moteurMarque: g('C17'), moteurModele: g('C18'), moteurPosition: g('C19'),
        moteurCouple: g('C20'), assistanceType: g('C21'), assistanceVitesseMax: g('C22'),
        batterieCapacite: g('C23'), batteriePosition: g('C24'), batterieAmovible: g('C25'),
        batterieEtat: g('C26'), batterieCycles: g('C27'), display: g('C28')
      },
      mecanique: {
        vitessesNombre: g('C29'), passageVitesse: g('C30'), derailleurType: g('C31'),
        derailleurMarque: g('C32'), derailleurModele: g('C33'),
        freinAvantSysteme: g('C34'), freinAvantMarque: g('C35'), freinAvantModele: g('C36'),
        freinArriereSysteme: g('C37'), freinArriereMarque: g('C38'), freinArriereModele: g('C39'),
        pneusModele: g('C40'), pneusTaille: g('C41'), roueTaille: g('C42'),
        fourche: g('C43'), bequille: g('C44')
      }
    },
    inspection:   inspection,
    statutGlobal: normalizeStatut(g('B101')),
    notes:        g('B100'),
    admin:        readAdminBlock_(sheet)   // v15
  };
}

// ============================================================
//  LISTE DES FICHES (pour le panneau "Mes fiches" de la webapp)
// ============================================================
function listSheets() {
  var ss = getSS();
  var tz = Session.getScriptTimeZone();
  var result = [];
  function fmt(val) {
    if (val === null || val === undefined || val === '') return '';
    if (val instanceof Date) return Utilities.formatDate(val, tz, 'dd/MM/yyyy');
    return val.toString().trim();
  }
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (!isFicheName(name)) return;
    var head = sheet.getRange('A1:C101').getValues();
    var d = toDate(head[1][2]);
    result.push({
      sheetName:  name,
      cs:         name,
      date:       fmt(head[1][2]),
      dateIso:    d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '',
      technicien: fmt(head[2][2]),
      marque:     fmt(head[6][2]),
      modele:     fmt(head[7][2]),
      statut:     normalizeStatut(fmt(head[100][1])),
      sheetsUrl:  'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit#gid=' + sheet.getSheetId()
    });
  });
  return result;
}

// ============================================================
//  UTILITAIRES
// ============================================================
function formatDateCompact(date) {
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, '0');
  var d = String(date.getDate()).padStart(2, '0');
  return '' + y + m + d;
}

// ============================================================
//  REBUILD SUMMARY MANUEL — à lancer après chaque mise à jour du script
// ============================================================
function rebuildSummaryManual() {
  Logger.log('Reconstruction du Summary…');
  clearDataCache();
  cacheRemove('turnover');
  updateSummary(true);
  Logger.log('Summary reconstruit. ✅');
}

// ============================================================
//  MIGRATION v12 — à lancer UNE fois depuis l'éditeur
//  « Facile à reco » (FACILE_RECO) et l'ancien « Prêt (en ligne) » (PRET)
//  → « Vélo à reconditionner » (RECONDITIONNABLE) dans la cellule B101 de
//  chaque fiche. Le statut affiché est ensuite recalculé (prêt / en ligne).
// ============================================================
function migrerStatutsV12() {
  var ss = getSS(), n = 0, noms = [];
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (!isFicheName(name)) return;
    var cell = sheet.getRange('B101');
    var code = String(cell.getValue() || '').trim();
    if (code === 'FACILE_RECO' || code === 'PRET') {
      cell.setValue('RECONDITIONNABLE');
      n++; noms.push(name + ' (' + code + ')');
    }
  });
  SpreadsheetApp.flush();
  Logger.log(n + ' fiche(s) passée(s) en RECONDITIONNABLE : ' + noms.join(', '));
  clearDataCache();
  cacheRemove('turnover');
  updateSummary(true);
  Logger.log('Summary reconstruit. ✅');
}

// ============================================================
//  TEST DE VITESSE — à lancer depuis l'éditeur (remplace le N° CS)
// ============================================================
function testVitesse() {
  var cs = 'METTRE_UN_N_CS_ICI';
  var t = Date.now();
  readSheetData(cs);           Logger.log('get    : ' + (Date.now() - t) + ' ms'); t = Date.now();
  getBikeStatus(cs);           Logger.log('status : ' + (Date.now() - t) + ' ms'); t = Date.now();
  listSheets();                Logger.log('list   : ' + (Date.now() - t) + ' ms'); t = Date.now();
  listBikes(true);             Logger.log('bikes (sans cache) : ' + (Date.now() - t) + ' ms'); t = Date.now();
  listBikes(false);            Logger.log('bikes (cache photos) : ' + (Date.now() - t) + ' ms');
}
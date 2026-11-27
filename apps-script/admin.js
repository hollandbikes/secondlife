// ============================================================
//  ADMINISTRATION RECO — v15 (27/11/2026)
//  Holland Bikes / Montévrain
//
//  Bloc rempli par Geoffroy à l'arrivée du vélo à EBSC (étape 1 de la webapp).
//  Pré-remplissage pour lui éviter d'aller chercher une info qui existe déjà :
//    • Sheets navette « HB LOG Navette / Expedition » : site d'origine (A),
//      destination (B), N° CS (I), n° de BLI (K), date de chargement (O)
//    • CycleSoftware (csconnect.js) : statut neuf / occasion / location,
//      code-barres, filiale, et garantie / TVA / emplacement si CS les renvoie
//  Tout reste modifiable dans la webapp.
//
//  Prix de reprise : JAMAIS renvoyé à la webapp (adresse publique, accès
//  sans connexion). Il est écrit uniquement dans la fiche Sheets, à chaque
//  enregistrement (writeAdminBlock).
//
//  Appel webapp : ?action=admin&cs=82036  (&fresh=1 pour ignorer le cache)
// ============================================================

const NAVETTE_SPREADSHEET_ID = '1-wc-Qgo1qpEBNOHrTobVFz84zZpjQMGQFeqDWW2UuWI';
const NAVETTE_TABS = ['NAVETTE HBLOG IDF', 'ARCHIVE Navette'];  // en-têtes sur 2 lignes
const CACHE_TTL_NAVETTE = 600;

// Bloc dans chaque fiche : A111:C125 (A = libellé, B = valeur, C = source)
const ADMIN_ROW_TITLE = 111;
const ADMIN_FIELDS = [
  { key: 'origine',        row: 112, label: "Site d'origine" },
  { key: 'dateTransport',  row: 113, label: 'Date du transport (navette)' },
  { key: 'destination',    row: 114, label: 'Destination (navette)' },
  { key: 'lienNavette',    row: 115, label: 'Ligne dans le Sheets navette' },
  { key: 'statutCS',       row: 116, label: 'Statut CS (neuf / occasion / location)' },
  { key: 'garantie',       row: 117, label: 'Garantie CS' },
  { key: 'codeBarre',      row: 118, label: 'Code-barres CS' },
  { key: 'tva',            row: 119, label: 'Statut TVA' },
  { key: 'factureOrigine', row: 120, label: "Facture d'origine (TVA marge)" },
  { key: 'prixReprise',    row: 121, label: 'Prix de reprise € HT (prix d\'achat CS)' },
  { key: 'bli',            row: 122, label: 'BLI transfert EBSC' },
  { key: 'filiale',        row: 123, label: 'Filiale CS' },
  { key: 'emplacement',    row: 124, label: 'Emplacement CS' },
  { key: 'factureBli',     row: 125, label: 'Facturation BLI (n° facture)' }
];
const ADMIN_ROWS = ADMIN_FIELDS.length + 1;   // titre + champs

// Filiales CS : noms lus dans CycleSoftware (horaires par magasin → store_id + store_name),
// cache 24 h. Constaté le 27/11 : 2 = EBSC (vélos transférés), 6 = boutique d'origine HB17
// (81460, BLI pas encore fait). Ce tableau ne sert que si l'API ne répond pas.
const CS_STORE_NAMES = {   // relevé dans CS le 27/11/2026 (testFilialesCS)
  1: 'Holland Bikes Lyon', 2: 'EBSC', 3: 'Atelier Paris 15', 4: 'Holland Bikes Paris 15',
  5: 'Holland Bikes Paris 05', 6: 'Holland Bikes Paris 17', 7: 'Holland Bikes Webshop', 9: 'Holland Bikes Lagny'
};
function csStoreNames_() {
  return cached('cs_stores', 86400, function() {
    var names = {};
    try {
      var r = csFetch_('/api/v2/workshop/trading-hours-per-store.json', { day_count: 1 });
      var list = r.json ? (Array.isArray(r.json) ? r.json : (r.json.data || [])) : [];
      list.forEach(function(x) { if (x && x.store_id !== undefined && x.store_name) names[x.store_id] = String(x.store_name).trim(); });
      if (!list.length) Logger.log('csStoreNames_ : HTTP ' + r.status + ' ' + String(r.text).substring(0, 300));
    } catch (e) { Logger.log('csStoreNames_ : ' + e.message); }
    return names;
  }, false);
}
// Test depuis l'éditeur : liste des filiales CS
function testFilialesCS() {
  cacheRemove('cs_stores');
  Logger.log(JSON.stringify(csStoreNames_(), null, 1));
}

// ============================================================
//  POINT D'ENTRÉE — doGet(action=admin)
// ============================================================
function getAdminInfo(cs, fresh) {
  cs = cleanCs(cs);
  var out = { cs: cs, navette: null, csAdmin: null };
  try { out.navette = findNavette(cs, fresh); }
  catch (e) { out.navetteError = e.message; }
  try {
    var pre = getCSPrefill(cs, fresh);
    out.csAdmin = pre.found ? (pre.admin || null) : null;
    if (!pre.found) out.csMessage = pre.message || 'Vélo introuvable dans CycleSoftware';
  } catch (e) { out.csError = e.message; }
  return out;
}

// ============================================================
//  SHEETS NAVETTE — index N° CS → transports (cache 10 min)
// ============================================================
function navetteIndex_(fresh) {
  return cached('navette_idx', CACHE_TTL_NAVETTE, function() {
    var ss = SpreadsheetApp.openById(NAVETTE_SPREADSHEET_ID);
    var tz = Session.getScriptTimeZone();
    var idx = {};
    NAVETTE_TABS.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return;
      var last = sh.getLastRow();
      if (last < 3) return;
      var vals = sh.getRange(3, 1, last - 2, 15).getValues();  // A3:O
      var gid = sh.getSheetId();
      vals.forEach(function(r, i) {
        var m = String(r[8] || '').match(/(\d{5})\s*$/);          // I : # Réf CS Objet
        if (!m) return;
        var d = r[14];                                             // O : chargement programmé
        idx[m[1]] = idx[m[1]] || [];
        idx[m[1]].push({
          onglet: name, gid: gid, row: i + 3,
          origine: String(r[0] || '').trim(),                      // A
          destination: String(r[1] || '').trim(),                  // B
          statut: String(r[4] || '').trim(),                       // E
          bli: String(r[10] || '').trim(),                         // K : # BL interne
          remarques: String(r[11] || '').trim().substring(0, 200), // L
          date: d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''
        });
      });
    });
    return idx;
  }, fresh);
}

function findNavette(cs, fresh) {
  var list = (navetteIndex_(fresh)[String(cs)] || []).slice();
  if (!list.length) return { found: false };
  // Plus récent d'abord (date de chargement, sinon onglet en cours avant l'archive)
  list.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return a.onglet === NAVETTE_TABS[0] ? -1 : 1;
  });
  list.forEach(function(t) {
    t.url = 'https://docs.google.com/spreadsheets/d/' + NAVETTE_SPREADSHEET_ID + '/edit#gid=' + t.gid + '&range=A' + t.row;
    t.versEBSC = /EBSC/i.test(t.destination);
    t.bliNumero = /^\d{5,8}$/.test(t.bli) ? t.bli : '';           // « ? », « NA »… → pas de n°
  });
  var ebsc = list.filter(function(t) { return t.versEBSC && !/ANNUL/i.test(t.statut); })[0] || null;
  return { found: true, dernierVersEBSC: ebsc, transports: list.slice(0, 10) };
}

// ============================================================
//  CYCLESOFTWARE — infos administratives de l'objet stock
//  Appelé par csMapToFiche_ (csconnect.js). Jamais de prix ici.
// ============================================================
function csAdminFromStock_(cs, raw) {
  var so = raw.stockObj || {}, st = so.status || {}, o = so.object || {};

  // Statut : location > occasion > démo > neuf
  var statutCS = '';
  // (61402 est à la fois is_demo et is_used : l'occasion l'emporte)
  if (st.is_rental === true) statutCS = 'LOCATION';
  else if (st.is_used === true) statutCS = 'OCCASION';
  else if (st.is_demo === true) statutCS = 'DEMO';
  else if (st.is_used === false) statutCS = 'NEUF';

  // Code-barres de l'objet : doit être le N° Objet CS
  var barcode = String(raw.barcode || o.barcode || so.barcode || '').trim();
  var bcNum = (barcode.match(/(\d+)\s*$/) || [])[1] || '';
  var codeBarreOk = barcode ? (bcNum === String(raw.objectId) || barcode === String(cs)) : null;

  // Filiale
  var storeId = so.store_id !== undefined ? so.store_id : (o.store_id !== undefined ? o.store_id : '');
  var stores = storeId === '' ? {} : csStoreNames_();
  var filiale = storeId === '' ? '' : (stores[storeId] || CS_STORE_NAMES[storeId] || ('Filiale CS n° ' + storeId));
  // Le vélo est-il déjà dans la filiale EBSC dans CS ? (sinon : BLI de transfert pas fait)
  var ebscId = Object.keys(stores).filter(function(k) { return /EBSC|MONTEVRAIN|MONTÉVRAIN/i.test(stores[k]); })[0] || '2';
  var enFilialeEBSC = storeId === '' ? null : String(storeId) === String(ebscId);

  // TVA et emplacement : absents de l'objet stock (vérifié le 27/11 sur 61402 et 82036).
  // Recherche par nom de clé gardée au cas où CS les ajoute.
  var flat = csFlatten_(so);
  var gar = csFindKey_(flat, /warrant|garant|guarantee/i);
  var garDuree = csFindKey_(flat, /(warrant|garant).*(month|mois|duration|duree|period)|(month|duration|duree).*(warrant|garant)/i);
  var tvaRaw = csFindKey_(flat, /(^|\.)(vat|tva|btw)|margin|marge/i);
  var empl = csFindKey_(flat, /location|emplacement|locatie|shelf|bin_|stock_place/i);

  // Champs réels (27/11) : object.warranty_type (1 = Garantie…) + object.warranty_duration_months
  var GAR_TYPES = { 1: 'Garantie', 2: 'Garantie Bovag', 3: 'Garantie du revendeur', 4: 'Garantie du fabricant', 5: 'Garantie de la marque', 6: 'Label Bike Totaal' };
  var garantie = '';
  var gMois = o.warranty_duration_months, gType = o.warranty_type;
  if (gMois) garantie = (gType && gType !== 1 && GAR_TYPES[gType] ? GAR_TYPES[gType] + ' · ' : '') + gMois + ' mois';
  else if (gType && GAR_TYPES[gType]) garantie = GAR_TYPES[gType] + ' (durée non renseignée)';
  else if (garDuree && /^\d+$/.test(garDuree.v)) garantie = garDuree.v + ' mois';

  // TVA : 20 % si CS l'indique, sinon occasion → TVA MARGE par défaut
  var tva = '', tvaSource = '';
  if (tvaRaw) {
    var tv = String(tvaRaw.v).toUpperCase();
    if (/MARG/.test(tv) || /margin|marge/i.test(tvaRaw.k) && tv === 'TRUE') { tva = 'TVA MARGE'; tvaSource = 'CS'; }
    else if (/20|HIGH|HAUT|HOOG/.test(tv)) { tva = 'TVA 20%'; tvaSource = 'CS'; }
  }
  if (!tva && statutCS === 'OCCASION') { tva = 'TVA MARGE'; tvaSource = 'regle'; }

  return {
    statutCS: statutCS,
    codeBarre: barcode, codeBarreOk: codeBarreOk,
    storeId: storeId, filiale: filiale, enFilialeEBSC: enFilialeEBSC,
    garantie: garantie,
    tva: tva, tvaSource: tvaSource,
    emplacement: empl ? empl.v : '',
    // Pour caler les noms de champs (jamais de valeurs de prix)
    champsTrouves: [gar, garDuree, tvaRaw, empl].filter(Boolean).map(function(x) { return x.k; }),
    venduFacture: so.sales_transaction_number || ''
  };
}

// Aplatit un objet en [{k: 'a.b.c', v: 'texte'}] (sans les clés de prix ni images)
function csFlatten_(x, pfx, out) {
  out = out || [];
  if (x === null || x === undefined) return out;
  if (typeof x !== 'object') { out.push({ k: pfx, v: String(x) }); return out; }
  Object.keys(x).forEach(function(k) {
    var p = pfx ? pfx + '.' + k : k;
    if (/price|cost|cents|purchase|achat|amount|images/i.test(k)) return;
    csFlatten_(x[k], p, out);
  });
  return out;
}
function csFindKey_(flat, re) {
  for (var i = 0; i < flat.length; i++) {
    var leaf = flat[i].k.split('.').slice(-2).join('.');
    if (re.test(leaf) && flat[i].v !== '' && flat[i].v !== 'null') return flat[i];
  }
  return null;
}

// Prix de reprise (prix d'achat de l'objet dans CS), en € — écrit seulement dans la fiche
function csPrixReprise_(cs) {
  var key = 'prixreprise_' + cs, cache = CacheService.getScriptCache();
  var hit = cache.get(key);
  if (hit !== null) return hit === '' ? '' : Number(hit);
  try {
    var raw = csLookup_(cs);
    if (!raw.found) return '';
    var so = raw.stockObj || {}, o = so.object || {}, obj = raw.obj || {};
    var c = [o.purchase_price_cents, so.purchase_price_cents, obj.purchase_price_cents]
      .filter(function(v) { return v !== undefined && v !== null && v !== ''; })[0];
    var prix = c === undefined ? '' : Math.round(Number(c)) / 100;
    cache.put(key, String(prix), CS_CACHE_TTL);
    return prix;
  } catch (e) {
    Logger.log('csPrixReprise_ ' + cs + ': ' + e.message);
    return '';
  }
}

// ============================================================
//  FICHE — écriture / lecture du bloc A111:C125
// ============================================================
// saved = valeurs de l'ancien bloc (avant recréation de l'onglet) ou null
function writeAdminBlock(sheet, admin, saved, cs) {
  var old = {};
  if (saved) ADMIN_FIELDS.forEach(function(f, i) { old[f.key] = saved[i + 1] ? saved[i + 1][1] : ''; });
  var a = admin || null;
  var need = ADMIN_ROW_TITLE + ADMIN_ROWS - 1;
  if (sheet.getMaxRows() < need) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());

  var rows = [['ADMINISTRATION RECO', '', 'Source']];
  ADMIN_FIELDS.forEach(function(f) {
    var v = a ? (a[f.key] !== undefined ? a[f.key] : '') : (old[f.key] !== undefined ? old[f.key] : '');
    var src = a && a.sources ? (a.sources[f.key] || '') : (saved ? (saved[ADMIN_FIELDS.indexOf(f) + 1] || [])[2] || '' : '');
    if (f.key === 'prixReprise') {
      var p = csPrixReprise_(cs);
      if (p !== '') { v = p; src = 'CS (prix d\'achat)'; } else { v = old.prixReprise || ''; src = v ? (src || 'CS') : ''; }
    }
    rows.push([f.label, v, src]);
  });
  sheet.getRange(ADMIN_ROW_TITLE, 1, rows.length, 3).setValues(rows);
  sheet.getRange(ADMIN_ROW_TITLE, 1, 1, 3).setFontWeight('bold').setFontColor('#00335b');
  sheet.getRange(ADMIN_ROW_TITLE + 1, 3, ADMIN_FIELDS.length, 1).setFontColor('#888888');
  var prixRow = ADMIN_FIELDS.filter(function(f) { return f.key === 'prixReprise'; })[0].row;
  sheet.getRange(prixRow, 2).setNumberFormat('#,##0.00 €');
}

function readAdminBlock_(sheet) {
  if (sheet.getMaxRows() < ADMIN_ROW_TITLE + ADMIN_ROWS - 1) return null;
  var vals = sheet.getRange(ADMIN_ROW_TITLE, 1, ADMIN_ROWS, 3).getValues();
  if (String(vals[0][0]).indexOf('ADMINISTRATION') !== 0) return null;
  var tz = Session.getScriptTimeZone();
  var out = { sources: {} };
  ADMIN_FIELDS.forEach(function(f, i) {
    if (f.key === 'prixReprise') return;   // jamais renvoyé à la webapp
    var v = vals[i + 1][1];
    out[f.key] = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v === null ? '' : v).trim();
    if (vals[i + 1][2]) out.sources[f.key] = String(vals[i + 1][2]);
  });
  out.prixRepriseRenseigne = vals[ADMIN_FIELDS.map(function(f) { return f.key; }).indexOf('prixReprise') + 1][1] !== '';
  return out;
}

// ============================================================
//  TEST — où CS range-t-il « Emplacement du stock » ?
//  L'objet stock ne le contient pas (vérifié le 27/11). On essaie les autres
//  routes de l'API qui parlent d'emplacement, et on n'affiche QUE les clés
//  liées au stock / magasin / emplacement (pas de prix, pas de clients).
//  Mettre dans « numeros » un vélo dont l'emplacement est rempli dans CS.
// ============================================================
function testEmplacementCS() {
  var numeros = ['82036'];
  var garder = /locat|emplac|stock|store|warehouse|shelf|bin|position|quantity|object_id|^data\.\d+\.barcode$/i;
  var jeter  = /price|cost|cents|purchase|amount|customer|email|phone|street|city|postcode|description/i;
  function aplatir(x, pfx, out) {
    out = out || [];
    if (x === null || typeof x !== 'object') { out.push(pfx + ' = ' + JSON.stringify(x)); return out; }
    Object.keys(x).forEach(function(k) { aplatir(x[k], pfx ? pfx + '.' + k : k, out); });
    return out;
  }
  function essai(titre, path, params) {
    try {
      var r = csFetch_(path, params);
      var lignes = r.json ? aplatir(r.json, '').filter(function(l) {
        var k = l.split(' = ')[0];
        return garder.test(k) && !jeter.test(k);
      }) : [String(r.text).substring(0, 300)];
      Logger.log('── ' + titre + ' → HTTP ' + r.status + '\n' + lignes.slice(0, 120).join('\n'));
    } catch (e) { Logger.log('── ' + titre + ' → ERREUR ' + e.message); }
  }
  numeros.forEach(function(cs) {
    Logger.log('================ VÉLO ' + cs + ' ================');
    essai('Article (articledata v4)', '/api/v4/articledata/ecommerce.json', { barcode: cs });
    essai('Stockinfo', '/app/api/whm/stockinfo/' + cs + '/', {});
    essai('Stock POS', '/api/v2/warehouse/stock/pos-items.json', {});
  });
}

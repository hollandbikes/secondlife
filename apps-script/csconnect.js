// ============================================================
//  CS CONNECT — pré-remplissage de la fiche depuis CycleSoftware
//  Holland Bikes / Montévrain — v11.1 (24/11/2026)
//
//  Fichier « csconnect.gs » dans le même projet Apps Script que summary.gs
//  (les fichiers .gs d'un projet partagent tout).
//
//  POURQUOI ÇA MARCHE DEPUIS N'IMPORTE QUEL ORDINATEUR
//  La webapp (GitHub Pages) n'appelle jamais CycleSoftware elle-même :
//  elle appelle ce script Apps Script, qui tourne sur les serveurs Google
//  avec TON compte, lit la clé dans les Propriétés du script et interroge
//  CS Connect. Aucun ordinateur n'a besoin de la clé.
//
//  Identifiants : JAMAIS dans le code ni dans index.html.
//  Projet Apps Script → ⚙ Paramètres du projet → Propriétés du script :
//     CS_API_USER   nom d'utilisateur CS Connect
//     CS_API_PASS   mot de passe CS Connect
//     CS_API_KEY    clé API (si CycleSoftware en a fourni une)
//     CS_API_BASE   (facultatif) https://api.cyclesoftware.fr par défaut
//
//  Appel webapp : ?action=prefill&cs=51846
//  → renvoie { found, prefill: { header, caracteristiques }, extras, source }
//    même forme que ?action=get ; la webapp ne remplit QUE les champs vides.
//
//  v11 : N° Objet CS à 5 chiffres uniquement (plus de recherche par EAN),
//        valeurs alignées sur les listes de la webapp, prix d'achat retiré
//        de la réponse (l'adresse du script est publique), champs retirés
//        de la fiche (poids, couleur secondaire, état batterie) non renvoyés.
//
//  v11.1 : lecture directe de l'objet stock (object / status / identification),
//        format réel constaté sur le vélo 61402. La fiche article n'est appelée
//        que si l'objet porte un vrai EAN.
//
//  Premier essai : lancer testCS() depuis l'éditeur avec un vrai N° CS,
//  puis envoyer le journal (Exécutions) pour caler les noms de champs.
// ============================================================

const CS_CACHE_TTL = 21600; // 6 h — les specs d'un vélo ne bougent pas

function csConfig_() {
  var p = PropertiesService.getScriptProperties();
  var user = p.getProperty('CS_API_USER');
  var pass = p.getProperty('CS_API_PASS');
  var key  = p.getProperty('CS_API_KEY') || '';
  if (!key && (!user || !pass)) throw new Error('Identifiants CS Connect absents (Propriétés du script CS_API_KEY ou CS_API_USER / CS_API_PASS)');
  return {
    user: user || '', pass: pass || '', key: key,
    base: (p.getProperty('CS_API_BASE') || 'https://api.cyclesoftware.fr').replace(/\/+$/, '')
  };
}

// GET JSON sur l'API CS — renvoie { status, json, text }
function csFetch_(path, params) {
  var cfg = csConfig_();
  var q = Object.assign({}, params || {});
  if (cfg.key && !q.api_key) q.api_key = cfg.key;
  var qs = Object.keys(q).map(function(k) { return encodeURIComponent(k) + '=' + encodeURIComponent(q[k]); }).join('&');
  var url = cfg.base + path + (qs ? '?' + qs : '');
  var headers = { Accept: 'application/json' };
  if (cfg.user && cfg.pass) headers.Authorization = 'Basic ' + Utilities.base64Encode(cfg.user + ':' + cfg.pass);
  if (cfg.key) headers['api-key'] = cfg.key;
  var res = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
  var status = res.getResponseCode(), text = res.getContentText();
  var json = null;
  try { json = JSON.parse(text); } catch (e) {}
  if (status === 401 || status === 403) throw new Error('CS Connect refuse l\'accès (' + status + ') — vérifier identifiants / scopes');
  return { status: status, json: json, text: text };
}

// N° saisi → id d'objet stock : 5 chiffres uniquement (« F.1-51846 » accepté)
function csObjectId_(cs) {
  var s = String(cs || '').trim();
  var m = s.match(/-(\d{5})$/);
  if (m) return m[1];
  return /^\d{5}$/.test(s) ? s : '';
}

// ============================================================
//  POINT D'ENTRÉE — appelé par doGet(action=prefill)
// ============================================================
function getCSPrefill(cs, fresh) {
  cs = String(cs || '').trim();
  if (!cs) throw new Error('cs manquant');
  if (!csObjectId_(cs)) return { found: false, cs: cs, message: 'Le N° Objet CS doit faire 5 chiffres (pas l\'EAN)' };
  var cacheKey = 'csprefill_' + cs;
  if (!fresh) { var hit = cacheGet(cacheKey); if (hit) return hit; }

  var raw = csLookup_(cs);
  var out = raw.found ? csMapToFiche_(cs, raw) : { found: false, cs: cs, message: raw.message || 'Vélo introuvable dans CycleSoftware' };
  if (out.found) cachePut(cacheKey, out, CS_CACHE_TTL);
  return out;
}

// 1) objet stock (N° Objet CS)  2) fiche article par le code-barres de cet objet
function csLookup_(cs) {
  var objectId = csObjectId_(cs);
  var stockObj = null, article = null, barcode = '';

  var r = csFetch_('/api/v1/stock-objects/' + encodeURIComponent(objectId) + '.json');
  if (r.status === 200 && r.json) {
    stockObj = (r.json.data && !Array.isArray(r.json.data)) ? r.json.data : r.json;
    var so = stockObj.object || stockObj;
    barcode = String(so.ean || so.barcode || '').trim();
  }

  // Fiche article (specs fabricant) : seulement si l'objet porte un vrai EAN.
  // Chez HB, « barcode » de l'objet = le N° Objet CS lui-même → pas d'appel inutile.
  if (/^\d{8,14}$/.test(barcode)) {
    var a = csFetch_('/api/v4/articledata/ecommerce.json', { barcode: barcode });
    if (a.status === 200 && a.json) {
      var list = Array.isArray(a.json) ? a.json : (a.json.data || []);
      if (list.length) article = list[0];
    }
  }

  if (!stockObj && !article) return { found: false, message: 'Aucun objet CS n° ' + objectId + ' trouvé dans CycleSoftware' };

  // L'exemplaire précis dans article.objects[] (cadre, km, prix…)
  var obj = null;
  if (article && Array.isArray(article.objects)) {
    obj = article.objects.filter(function(o) {
      return String(o.object_id) === objectId || String(o.object_code || '').indexOf('-' + objectId) > 0;
    })[0] || (article.objects.length === 1 ? article.objects[0] : null);
  }
  return { found: true, objectId: objectId, barcode: barcode, stockObj: stockObj, article: article, obj: obj };
}

// ============================================================
//  LECTURE DES PROPRIÉTÉS (format Articledata V4)
//  properties.<nom> = { type, unit, values: [{ object_ids, supplier_ids, code, value }] }
// ============================================================
function csPropText_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v !== 'object') return String(v).trim();
  if (Array.isArray(v)) return v.map(csPropText_).filter(String).join(', ');
  var langs = ['fr_FR', 'fr', 'FR', 'en_GB', 'en_US', 'en', 'nl_NL', 'nl', 'description', 'name', 'label', 'value'];
  for (var i = 0; i < langs.length; i++) if (v[langs[i]] !== undefined && typeof v[langs[i]] !== 'object') return String(v[langs[i]]).trim();
  for (var k in v) if (typeof v[k] === 'string' && v[k].trim()) return v[k].trim();
  return '';
}

// Valeur d'une propriété : priorité à l'exemplaire (object_ids) puis 1re valeur
function csProp_(props, names, objectId) {
  if (!props) return '';
  for (var i = 0; i < names.length; i++) {
    var p = props[names[i]];
    if (!p) continue;
    var vals = Array.isArray(p.values) ? p.values : (p.value !== undefined ? [p] : []);
    if (!vals.length) continue;
    var mine = vals.filter(function(x) { return (x.object_ids || []).map(String).indexOf(String(objectId)) >= 0; })[0];
    var x = mine || vals[0];
    var txt = csPropText_(x.value !== undefined ? x.value : x);
    if (!txt && x.code) txt = String(x.code);
    if (txt) return p.unit && /^\d+([.,]\d+)?$/.test(txt) ? { v: txt, unit: p.unit, code: x.code || '' } : { v: txt, code: x.code || '' };
  }
  return '';
}
function csVal_(r) { return r ? r.v : ''; }
function csCode_(r) { return r ? String(r.code || r.v).toUpperCase() : ''; }

// Recherche souple : 1re propriété dont le nom contient tous les mots
function csPropLike_(props, words, objectId) {
  if (!props) return '';
  var keys = Object.keys(props).filter(function(k) {
    return words.every(function(w) { return k.toLowerCase().indexOf(w) >= 0; });
  });
  return keys.length ? csProp_(props, keys, objectId) : '';
}

// ============================================================
//  CORRESPONDANCE → champs de la fiche (mêmes clés que fillSheet)
//  Les valeurs des listes déroulantes sont celles de la webapp.
// ============================================================
var CS_MOTOR_BRANDS = ['Bosch', 'Shimano', 'Brose', 'Yamaha', 'Bafang', 'Panasonic', 'Ananda', 'Mahle',
                       'Specialized', 'Giant', 'TQ', 'Fazua', 'Neodrives', 'Impulse', 'Valeo', 'Polini', 'Promovec', 'Sachs', 'Oli', 'Enviolo', 'DAPU', 'SR Suntour'];

// Codes CS des systèmes de freinage (type_brake_system_*) → liste webapp
//   SH = disque hydraulique (vu sur 61402). Les autres codes sont déduits : à confirmer.
var CS_BRAKE_CODES = {
  SH: 'Hydraulique disque', HS: 'Hydraulique disque', HD: 'Hydraulique disque',
  SM: 'Mécanique', MS: 'Mécanique', S: 'Mécanique',
  VH: 'Hydraulique sur jante', HV: 'Hydraulique sur jante', H: 'Hydraulique sur jante',
  V: 'V-Brake', VB: 'V-Brake',
  R: 'Mécanique', RO: 'Mécanique', T: 'Mécanique', TR: 'Mécanique', N: 'Mécanique', NR: 'Mécanique'
};
// « GAZELLE » → « Gazelle » (seulement les textes tout en majuscules)
function csNice_(s) {
  s = String(s || '').trim();
  if (!s || s !== s.toUpperCase() || !/[A-Z]{3}/.test(s)) return s;
  return s.toLowerCase().replace(/(^|[\s\-\/])([a-zà-ÿ])/g, function(m, a, b) { return a + b.toUpperCase(); });
}

function csMapToFiche_(cs, raw) {
  var art = raw.article || {}, so = raw.stockObj || {}, obj = raw.obj || {};
  var P = art.properties || {}, oid = raw.objectId;
  var o  = so.object || {};          // caractéristiques de l'objet stock
  var st = so.status || {};          // km, occasion, disponible…
  var ident = so.identification || obj.identification || {};
  function p(names) { return csProp_(P, names, oid); }
  function like(words) { return csPropLike_(P, words, oid); }
  function first() { for (var i = 0; i < arguments.length; i++) { var a = arguments[i]; if (a !== undefined && a !== null && String(a).trim() !== '') return String(a).trim(); } return ''; }

  // ── Général
  var nomCS   = first(o.model_name, so.model_name);
  var marque  = csNice_(first(o.brand_name, csVal_(p(['brand'])), so.brand_name));
  // Le nom CS se termine par « / RECONDITIONNE » une fois le vélo traité : on ne le reprend pas dans le modèle
  var modele  = first(nomCS, csVal_(p(['model']))).replace(/\s*\/\s*RECONDITIONN[EÉ]\s*$/i, '').trim();
  var annee   = first(o.model_year, csVal_(p(['model_year'])));
  var km      = first(st.km_age, obj.km_age, so.km_age);

  // ── Électrique ?
  var systeme  = csVal_(p(['ebike_system', 'engine', 'motor'])) || csVal_(like(['engine'])) || csVal_(like(['motor', 'brand']));
  var batt     = p(['battery_capacity']) || like(['battery', 'capacity']);
  // Capacité écrite dans le nom CS (« Makki Load 400wh ») ou la description courte
  var whTxt    = (String(nomCS + ' ' + (o.description_short || '')).match(/(\d{3,4})\s?wh\b/i) || [])[1] || '';
  if (!batt && whTxt) batt = { v: whTxt };
  var ebType   = csCode_(p(['electric_bicycle_type']));
  var isElec   = !!(systeme || batt || /PEDELEC|ELECTRIC|EBIKE|E-BIKE/.test(ebType) || ident.battery_number || ident.engine_number);

  var moteurMarque = '', moteurModele = '';
  if (systeme) {
    var brand = CS_MOTOR_BRANDS.filter(function(b) { return new RegExp('\\b' + b + '\\b', 'i').test(systeme); })[0];
    moteurMarque = brand || systeme.split(/\s+/)[0];
    moteurModele = systeme.replace(new RegExp('\\b' + moteurMarque + '\\b', 'i'), '').replace(/\s+/g, ' ').trim();
  }
  var mBrand = csVal_(like(['motor', 'brand'])) || csVal_(like(['engine', 'brand']));
  if (!moteurMarque && mBrand) moteurMarque = mBrand;

  var motorPos = csCode_(p(['motor_position']));
  var moteurPosition = /MID|CENT|CRANK|BOTTOM|TRAPP/.test(motorPos) ? 'Central'
                     : /FRONT|AVANT|VOOR/.test(motorPos) ? 'Moyeu avant'
                     : /REAR|BACK|ARRI|ACHTER/.test(motorPos) ? 'Moyeu arrière' : '';

  var couple = like(['torque']); couple = couple ? couple.v + ' Nm' : '';

  var sensor = csCode_(p(['sensor_type']));
  var assistanceType = /TORQUE|COUPLE|KRACHT/.test(sensor) && /CADENCE|ROTAT|SPEED|TRAP/.test(sensor) ? 'Mixte'
                     : /TORQUE|COUPLE|KRACHT/.test(sensor) ? 'Couple'
                     : /CADENCE|ROTAT|TRAP/.test(sensor) ? 'Pédalage' : '';
  var vmax = /SPEED|S-PEDELEC|SPEEDPEDELEC|45/.test(ebType) ? '45' : (isElec ? '25' : '');

  // Liste webapp : Cadre / Porte-bagage / Externe
  var battPos = csCode_(p(['battery_position', 'battery_location']));
  var batteriePosition = /RACK|CARRIER|BAGAGE|DRAGER/.test(battPos) ? 'Porte-bagage'
                       : /INTEG|FRAME|INTUBE|DOWN|BENEDEN|DIAG|SEAT|ZADEL|SELLE|CADRE/.test(battPos) ? 'Cadre'
                       : /EXTERN|EXT/.test(battPos) ? 'Externe' : '';
  var amov = csCode_(p(['battery_removable'])) || csCode_(like(['battery', 'remov']));
  var batterieAmovible = /TRUE|YES|OUI|JA|1/.test(amov) ? 'Oui' : /FALSE|NO|NON|NEE|0/.test(amov) ? 'Non' : '';

  // ── Cadre
  var mat = String(o.frame_material || '').toUpperCase() || csCode_(p(['frame_material']));
  var materiau = /ALU/.test(mat) ? 'Aluminium' : /CARB/.test(mat) ? 'Carbone' : /TITAN/.test(mat) ? 'Titane'
               : /STEEL|ACIER|STAAL|CRMO|CHRO/.test(mat) ? 'Acier' : (mat ? 'Autre' : '');
  var hauteur = first(o.frame_height ? o.frame_height + ' cm' : '', o.frame_size_supplier,
                      csVal_(p(['framesize_description', 'frame_height_description'])),
                      (function() { var h = p(['framesize', 'frame_height']); return h ? h.v + ' ' + (h.unit || 'cm') : ''; })());
  var couleur = first(o.color, csNice_(o.primary_color), csVal_(p(['color_description'])), csVal_(p(['basic_color', 'base_color'])));

  // ── Mécanique
  var gsys = String(o.type_gear_system || '').toUpperCase() || csCode_(p(['gear_system']));
  var gearModel = first(o.model_gear_system, csVal_(p(['gearsystem_rear_model', 'gear_model'])));
  var gearBrand = csNice_(first(o.brand_gear_system, csVal_(p(['gearsystem_rear_brand', 'gear_brand'])), csVal_(like(['gear', 'brand']))));
  var gears = first(o.gear_count, csVal_(p(['gears', 'gear_count'])));
  // Moyeu à variation continue (Enviolo…) : CS met 1 vitesse → on n'en déduit rien
  var cvt = /ENVIOLO|NUVINCI|CVT/.test((gearBrand + ' ' + gearModel).toUpperCase());
  if (cvt && gears === '1') gears = 'Continu';
  if (!gearBrand && gearModel) {
    var gb = ['Shimano', 'SRAM', 'Enviolo', 'Rohloff', 'Sturmey', 'Microshift', 'Campagnolo', 'Pinion', 'Kindernay']
      .filter(function(b) { return new RegExp(b, 'i').test(gearModel); })[0];
    if (gb) gearBrand = gb;
  }
  // Liste webapp « Dérailleur (Type) » : Arrière / Avant + Arrière / Interne / Aucun
  var hubLike = /HUB|NABE|MOYEU|CVT|NEXUS|ALFINE|ENVIOLO|ROHLOFF|INTER/.test(gsys + ' ' + gearModel.toUpperCase());
  var single  = /SINGLE/.test(gsys) || (gears === '1' && !hubLike);
  var derailleurType = single ? 'Aucun' : hubLike ? 'Interne'
                     : (gsys || gearModel) ? (/FRONT|VOOR|AVANT/.test(gsys) ? 'Avant + Arrière' : 'Arrière') : '';
  var passageVitesse = single ? 'Single speed' : ''; // manette / poignée : pas fiable dans CS

  // Liste webapp « Frein (Système) » : Hydraulique disque / Hydraulique sur jante / Mécanique / V-Brake
  function brakeSys(code) {
    code = String(code || '').toUpperCase();
    if (!code) return '';
    if (/HYD/.test(code) && /DISC|SCHIJF|DISQUE/.test(code)) return 'Hydraulique disque';
    if (/HYD/.test(code)) return 'Hydraulique sur jante';
    if (/V-?BRAKE|V_BRAKE/.test(code)) return 'V-Brake';
    if (/DISC|SCHIJF|DISQUE|ROLLER|ROLLEN|ROULEAU|COASTER|TERUGTRAP|RETRO|TORPEDO|RIM|VELG|PATIN|CALIPER/.test(code)) return 'Mécanique';
    return '';
  }
  var bsAll = csCode_(p(['brake_system']));
  var bsF = csCode_(p(['brake_system_front'])) || bsAll, bsR = csCode_(p(['brake_system_rear'])) || bsAll;
  function brakeFromObj(code) { code = String(code || '').toUpperCase().trim(); return code ? (CS_BRAKE_CODES[code] || '') : ''; }
  var freinAvSys = brakeFromObj(o.type_brake_system_front) || brakeSys(bsF);
  var freinArSys = brakeFromObj(o.type_brake_system_rear)  || brakeSys(bsR);

  // Roues : ETRTO → pouces
  var etrto = first(csVal_(p(['etrto_wheelsize', 'wheelsize', 'wheel_size'])));
  var ws = String(o.wheel_size || '').replace(',', '.');   // pouces dans l'objet stock (ex. 26)
  var bead = (String(etrto).match(/(\d{3})\s*$/) || String(etrto).match(/-(\d{3})/) || [])[1] || '';
  var inch = (String(etrto).match(/(\d{2}(?:[.,]5)?)\s*("|''|inch|in\b|pouces)/i) || [])[1] || '';
  var roueTaille = inch ? (inch.replace(',', '.') === '28' ? '28"/700c' : inch.replace(',', '.') + '"')
                 : ({ '622': '28"/700c', '584': '27.5"', '559': '26"', '406': '20"', '305': '16"' }[bead] || '');
  if (bead === '622' && /29/.test(etrto)) roueTaille = '29"';
  if (ws) roueTaille = ws === '28' ? '28"/700c' : ws + '"';

  var prefill = {
    header: { objetCS: cs },
    caracteristiques: {
      general: {
        typeVelo: isElec ? 'ELECTRIQUE' : (raw.article ? 'MECANIQUE' : ''),
        numeroCadre: first(ident.frame_number, so.frame_number, so.frame_id),
        marque: marque, modele: modele, anneeModele: annee, kilometrage: km
      },
      cadre: {
        materiau: materiau, hauteur: hauteur, couleurPrincipale: couleur
      },
      electrique: isElec ? {
        moteurMarque: moteurMarque, moteurModele: moteurModele, moteurPosition: moteurPosition, moteurCouple: couple,
        assistanceType: assistanceType, assistanceVitesseMax: vmax,
        batterieCapacite: batt ? String(batt.v).replace(/\s*wh$/i, '') : '', batteriePosition: batteriePosition, batterieAmovible: batterieAmovible,
        batterieCycles: '', display: first(csVal_(p(['display_type', 'display'])))
      } : {},
      mecanique: {
        vitessesNombre: gears, passageVitesse: passageVitesse, derailleurType: derailleurType,
        derailleurMarque: gearBrand, derailleurModele: gearModel,
        freinAvantSysteme: freinAvSys,
        freinAvantMarque: csNice_(first(o.brand_brake_system_front, csVal_(p(['brake_front_brand', 'brakes_front_brand'])), csVal_(like(['brake', 'front', 'brand'])))),
        freinAvantModele: first(o.model_brake_system_front, csVal_(p(['brake_front_model', 'brakes_front_model'])), csVal_(like(['brake', 'front', 'model']))),
        freinArriereSysteme: freinArSys,
        freinArriereMarque: csNice_(first(o.brand_brake_system_rear, csVal_(p(['brake_rear_brand', 'brakes_rear_brand'])), csVal_(like(['brake', 'rear', 'brand'])))),
        freinArriereModele: first(o.model_brake_system_rear, csVal_(p(['brake_rear_model', 'brakes_rear_model'])), csVal_(like(['brake', 'rear', 'model']))),
        pneusModele: first(csVal_(like(['tyre', 'model'])), csVal_(like(['tire', 'model']))),
        pneusTaille: /\d{2}-\d{3}/.test(etrto) ? etrto : '',
        roueTaille: roueTaille,
        fourche: o.has_front_suspension === true ? 'Suspension' : first(csVal_(like(['fork']))),
        bequille: ''
      }
    }
  };

  // Infos utiles hors fiche (affichées dans la webapp, pas écrites dans Sheets)
  // Pas de prix d'achat : l'adresse du script est publique.
  function eur(c) { return (c === undefined || c === null || c === '') ? null : Math.round(Number(c)) / 100; }
  var extras = {
    objectCode: first(o.object_reference, obj.object_code),
    isUsed: st.is_used !== undefined ? !!st.is_used : null,
    etatCS: first(o.custom_variable_1),                       // ex. « Bon État »
    prixMagasin: eur(first(o.sales_price_cents, obj.pos_sales_price_cents)),
    prixWeb: eur(first(o.ecommerce_price_cents, obj.ecommerce_price_cents)),
    numeroBatterie: first(ident.battery_number), numeroBatterie2: first(ident.battery_number_2),
    numeroMoteur: first(ident.engine_number), numeroCle: first(ident.key_number),
    numeroAntivol: first(ident.lock_number), numeroSerie: first(ident.serial_number),
    nomCS: nomCS
    // Pas de description : elle contient des notes internes (prix, initiales)
  };

  // Champs remplis (pour le message « 18 champs pré-remplis »)
  var n = 0;
  ['general', 'cadre', 'electrique', 'mecanique'].forEach(function(s) {
    var o = prefill.caracteristiques[s];
    Object.keys(o).forEach(function(k) { if (k !== 'typeVelo' && o[k]) n++; });
  });

  return { found: true, cs: cs, nbChamps: n, prefill: prefill, extras: extras,
           source: { stockObject: !!raw.stockObj, article: !!raw.article } };
}

// ============================================================
//  TEST — lancer depuis l'éditeur (remplacer le N° CS)
//  Le journal montre les réponses brutes de CS Connect + le résultat.
//  (Le journal n'est visible que dans l'éditeur : il peut contenir les prix.)
// ============================================================
function testCS() {
  var cs = '61402';
  var raw = csLookup_(cs);
  Logger.log('objet stock : ' + JSON.stringify(raw.stockObj, null, 1));
  Logger.log('article     : ' + JSON.stringify(raw.article, null, 1).substring(0, 20000));
  if (raw.article) Logger.log('propriétés CS : ' + Object.keys(raw.article.properties || {}).join(', '));
  if (raw.found) Logger.log('PRÉ-REMPLISSAGE : ' + JSON.stringify(csMapToFiche_(cs, raw), null, 1));
  else Logger.log(raw.message);
}
// ============================================================
//  EXPLORATION « Administration Reco » (27/11/2026)
//  À lancer UNE fois depuis l'éditeur (menu Exécuter) avec 2 ou 3 vrais N° CS
//  de vélos reco (un en TVA 20 %, un en TVA marge, un déjà transféré à EBSC).
//  Le journal liste tous les champs renvoyés par CS (garantie, TVA, emplacement,
//  filiale, prix de reprise…) pour caler la synchro. Il peut contenir des prix :
//  il reste dans l'éditeur, rien n'est renvoyé à la webapp.
// ============================================================
function explorerAdminCS() {
  var numeros = ['61402', '82036'];   // ← mettre ici 2 ou 3 N° Objet CS

  // Aplatit un objet JSON en lignes « chemin = valeur »
  function aplatir(x, pfx, out) {
    out = out || [];
    if (x === null || typeof x !== 'object') { out.push(pfx + ' = ' + JSON.stringify(x)); return out; }
    var ks = Object.keys(x);
    if (!ks.length) out.push(pfx + ' = ' + (Array.isArray(x) ? '[]' : '{}'));
    ks.forEach(function(k) { aplatir(x[k], pfx ? pfx + '.' + k : k, out); });
    return out;
  }
  function essai(titre, path, params) {
    try {
      var r = csFetch_(path, params);
      Logger.log('── ' + titre + ' (' + path + ') → HTTP ' + r.status);
      if (r.json) Logger.log(aplatir(r.json, '').slice(0, 400).join('\n'));
      else Logger.log(String(r.text).substring(0, 1500));
    } catch (e) { Logger.log('── ' + titre + ' → ERREUR ' + e.message); }
  }

  numeros.forEach(function(cs) {
    Logger.log('================ VÉLO ' + cs + ' ================');
    essai('Objet stock', '/api/v1/stock-objects/' + cs + '.json');
  });
  // Listes de codes : types de garantie, codes TVA, liste des enums disponibles
  essai('Liste des enums', '/api/v1/common/enum.json');
  essai('Types de garantie', '/api/v1/common/enum/warranty_types.json');
  essai('Codes TVA', '/api/v4/articledata/codelist/vat.json');

  // Factures : transactions du 26/11/2026 (jour du BLI 676607 → facture 20007533 pour le 82036)
  // Le journal montre si les factures internes (BLI) apparaissent, et sous quelle forme.
  essai('Transactions 26/11/2026', '/api/v3/salesdata/transactions.json',
        { start_date: '2026-11-26', end_date: '2026-11-26' });
}

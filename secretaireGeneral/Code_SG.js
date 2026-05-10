/**
 * Code_SG.js — Version allégée :
 * Ne conserve que la configuration et les helpers partagés nécessaires aux modules.
 */

// ────────────────────────────────────────────────────────────
//  CONFIGURATION
// ────────────────────────────────────────────────────────────

var SG_CONFIG = {
  driveFolderId: '0AOEs1AL5uGXpUk9PVA',
  sheetName: 'Base de données SG',
  sheets: {
    membres: 'Membres',
    reunions: 'Réunions',
    votes: 'Votes',
    archives: 'Archives',
    licences: 'Licences',
    rgpd: 'RGPD',
    config: 'Config'
  },
  junior: {
    nom: 'Junior INSA Services',
    nom_ecole: 'INSA Toulouse',
    adresse1: '135 avenue de Rangueil',
    code_postal: '31400',
    ville: 'Toulouse'
  },
  cacheTtl: 60
};

// ────────────────────────────────────────────────────────────
//  HELPERS PARTAGÉS
// ────────────────────────────────────────────────────────────

function getCacheKey_(key) {
  return 'SG_' + key;
}
function fromCache_(key) {
  try {
    var cached = CacheService.getScriptCache().get(getCacheKey_(key));
    return cached ? JSON.parse(cached) : null;
  } catch (e) {
    return null;
  }
}
function toCache_(key, value) {
  try {
    CacheService.getScriptCache().put(getCacheKey_(key), JSON.stringify(value), SG_CONFIG.cacheTtl);
  } catch (e) {}
}
function invalidateCache_(key) {
  try {
    CacheService.getScriptCache().remove(getCacheKey_(key));
  } catch (e) {}
}
function getSGSpreadsheet_() {
  var folder;
  try {
    folder = DriveApp.getFolderById(SG_CONFIG.driveFolderId);
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }
  var files = folder.getFilesByName(SG_CONFIG.sheetName);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  var ss = SpreadsheetApp.create(SG_CONFIG.sheetName);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  return ss;
}
function getSGSheet_(sheetName) {
  var ss = getSGSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

// Ajoutez ici d'autres helpers strictement nécessaires aux modules si besoin.

function _initLicencesSheet_(sheet) {
  try {
    sheet.clearContents();
    // Colonnes principales reprenant le tableau fourni
    sheet.appendRow(['id', 'logiciel', 'type_de_licence', 'date_installation', 'date_debut_validite', 'date_fin_validite', 'proprietaire', 'contexte_utilisation', 'preuve_paiement', 'conditions_generales', 'remarques', 'data_json']);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
  } catch (e) {}
}

/**
 * Lecture simple de la config (clé unique)
 */

function _generateId_() {
  return String(Date.now()) + String(Math.floor(Math.random() * 10000));
}

// ────────────────────────────────────────────────────────────
//  MEMBRES
// ────────────────────────────────────────────────────────────

/**
 * Checklist membres (v2) : BeeQuick supprimé → 3 étapes
 * kiwi | ba_signe | cotisation
 */

// ────────────────────────────────────────────────────────────
//  GÉNÉRATION BA (Bulletin d'Adhésion)
// ────────────────────────────────────────────────────────────

/**
 * Génère un BA pour un membre depuis le template "ba-membre" ou "ba-intervenant".
 * @param {string} membreId  ID du membre dans la BDD SG
 * @param {string} type      'membre' | 'intervenant'
 */

// ────────────────────────────────────────────────────────────
//  GÉNÉRATION ODJ
// ────────────────────────────────────────────────────────────

/**
 * Génère un ODJ pour une réunion.
 * @param {string} reunionId
 * @param {boolean} avecIndicateurs  Si true → template "odj-ca-indicateurs"
 */

// ────────────────────────────────────────────────────────────
//  BALOTILO — Import résultats
// ────────────────────────────────────────────────────────────

/**
 * Récupère les résultats d'un vote Balotilo depuis son URL publique.
 * Retourne { pour, contre, abstentions, total, titre, success }
 */
function fetchBalotiloResults(url) {
  if (!url || !url.trim()) return {
    success: false,
    error: 'URL vide'
  };
  try {
    var cleanUrl = url.trim();
    // Normaliser l'URL (supporte toute occurrence de 'balotilo')
    if (cleanUrl.toLowerCase().indexOf('balotilo') === -1) {
      return {
        success: false,
        error: "URL invalide — doit contenir 'balotilo'"
      };
    }
    var response = UrlFetchApp.fetch(cleanUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0'
      }
    });
    if (response.getResponseCode() !== 200) {
      return {
        success: false,
        error: 'Erreur HTTP ' + response.getResponseCode()
      };
    }
    var html = response.getContentText();
    var result = _parseBalotiloHtml_(html);
    result.url = cleanUrl;
    return result;
  } catch (e) {
    return {
      success: false,
      error: 'Erreur réseau: ' + e.message
    };
  }
}

/**
 * Parse le HTML Balotilo pour extraire les comptages.
 * Balotilo affiche les options avec leurs scores.
 */
function _parseBalotiloHtml_(html) {
  var result = {
    success: false,
    pour: 0,
    contre: 0,
    abstentions: 0,
    total: 0,
    titre: '',
    options: []
  };
  if (!html) return result;

  // Titre de la question
  var titreMatch = html.match(/<h1[^>]*>([^<]+)<\/h1>/i) || html.match(/class="[^"]*title[^"]*"[^>]*>([^<]+)</i) || html.match(/<title>([^<]+)<\/title>/i);
  if (titreMatch) result.titre = titreMatch[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();

  // Chercher les scores — Balotilo affiche généralement des nombres avec %
  // Pattern : "X votes" ou des chiffres dans des éléments de résultats
  var scorePatterns = [/(\d+)\s*vote/gi, /score[^>]*>[\s]*(\d+)/gi, /"count"\s*:\s*(\d+)/gi, /data-votes="(\d+)"/gi, /class="[^"]*result[^"]*"[^>]*>[\s]*(\d+)/gi];
  var numbers = [];
  for (var p = 0; p < scorePatterns.length; p++) {
    var pattern = scorePatterns[p];
    var match;
    while ((match = pattern.exec(html)) !== null) {
      var n = parseInt(match[1], 10);
      if (!isNaN(n)) numbers.push(n);
    }
    if (numbers.length >= 3) break;
  }

  // Fallback: chercher des objets JSON inline contenant count/score
  if (numbers.length < 3) {
    var jsonMatch = html.match(/\{[^}]{0,400}\}/g);
    if (jsonMatch) {
      for (var j = 0; j < jsonMatch.length && numbers.length < 3; j++) {
        try {
          var obj = JSON.parse(jsonMatch[j]);
          if (obj && typeof obj === 'object') {
            if (obj.count && !isNaN(parseInt(obj.count, 10))) numbers.push(parseInt(obj.count, 10));
            if (obj.votes && !isNaN(parseInt(obj.votes, 10))) numbers.push(parseInt(obj.votes, 10));
            if (obj.score && !isNaN(parseInt(obj.score, 10))) numbers.push(parseInt(obj.score, 10));
          }
        } catch (e) {}
      }
    }
  }

  // Normalize numbers: keep first 3 meaningful numbers
  if (numbers.length > 0) {
    var uniq = [];
    for (var ii = 0; ii < numbers.length; ii++) {
      var val = numbers[ii];
      if (uniq.indexOf(val) === -1) uniq.push(val);
    }
    numbers = uniq.slice(0, 3);
  }

  // Chercher les options (Pour / Contre / Abstention)
  var options = [];
  var optionPattern = /(?:Oui|Pour|Yes|Non|Contre|No|Abstention|Neutre)[^<]*/gi;
  var optMatch;
  while ((optMatch = optionPattern.exec(html)) !== null) {
    options.push(optMatch[0].trim());
  }

  // Mapping standard Balotilo : première option = Pour, deuxième = Contre, troisième = Abstention
  if (numbers.length >= 1) {
    result.pour = numbers[0] || 0;
    result.contre = numbers[1] || 0;
    result.abstentions = numbers[2] || 0;
    result.total = result.pour + result.contre + result.abstentions;
    result.success = true;
    result.options = options;
  }

  // Si aucun nombre n'a été trouvé, tenter une extraction heuristique près des labels
  if (!result.success || result.total === 0) {
    try {
      var labelPattern = /(?:Oui|Pour|Yes|Non|Contre|No|Abstention|Neutre)/gi;
      var match;
      var extracted = [];
      while ((match = labelPattern.exec(html)) !== null && extracted.length < 3) {
        var start = Math.max(0, match.index - 40);
        var slice = html.substr(start, 200);
        var nm = slice.match(/(\d{1,6})/);
        if (nm) extracted.push(parseInt(nm[1], 10));
      }
      if (extracted.length) {
        result.pour = extracted[0] || 0;
        result.contre = extracted[1] || 0;
        result.abstentions = extracted[2] || 0;
        result.total = result.pour + result.contre + result.abstentions;
        result.success = true;
      }
    } catch (e) {/* ignore fallback errors */}
  }
  return result;
}

// ────────────────────────────────────────────────────────────
//  RÉUNIONS / CA
// ────────────────────────────────────────────────────────────

// ────────────────────────────────────────────────────────────
//  VOTES
// ────────────────────────────────────────────────────────────

// ────────────────────────────────────────────────────────────
//  GÉNÉRATION CR / DOCUMENTS
// ────────────────────────────────────────────────────────────

// ────────────────────────────────────────────────────────────
//  CHARGEMENT GROUPÉ (performance)
// ────────────────────────────────────────────────────────────

/**
 * Charge membres + réunions en un seul appel GAS pour réduire la latence.
 */
function chargerDonneesSG() {
  return {
    membres: getMembres(),
    reunions: getReunions(),
    config: getAllConfig()
  };
}
/** Retourne le token OAuth pour Google Picker (côté client) */
function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

/**
 * Envoie des rappels pour les licences arrivant à expiration.
 * Utilise le webhook Google Chat configuré dans la feuille Config (key: google_chat_webhook).
 * Optionnellement définit la fenêtre de rappel en jours (key: licence_reminder_days).
 */

function _clearTriggersForFunction_(fnName) {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) {
    if (ts[i].getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(ts[i]);
    }
  }
}

/** Crée un déclencheur journalier à 8h pour envoyer les rappels de licences. */

/** Copie un fichier Drive sélectionné par l'utilisateur dans le dossier Licences configuré et retourne l'URL */

// ────────────────────────────────────────────────────────────
//  PLACEHOLDERS — fonctionnalités listées dans README (scaffolds)
// ────────────────────────────────────────────────────────────

/** Prépare automatiquement un ODJ en se basant sur les études en cours */
function prepareAutomaticODJ(options) {
  return {
    success: false,
    error: 'Not implemented',
    note: 'prepareAutomaticODJ to be implemented'
  };
}

/** Envoie l'ODJ par mail aux membres du CA */
function sendODJToCA(reunionId, options) {
  return {
    success: false,
    error: 'Not implemented',
    note: 'sendODJToCA to be implemented'
  };
}

/** Planification de rappels automatiques pour une réunion (H-24, H-1) */
function scheduleMeetingReminders(reunionId, reminders) {
  return {
    success: false,
    error: 'Not implemented',
    note: 'scheduleMeetingReminders to be implemented'
  };
}

/** Suivi des licences des membres — scaffold pour future intégration */
function trackMemberLicenses() {
  return {
    success: false,
    error: 'Not implemented',
    note: 'trackMemberLicenses to be implemented'
  };
}

/** Génère une feuille d'émargement depuis un template */
function generateAttendanceSheet(reunionId) {
  return {
    success: false,
    error: 'Not implemented',
    note: 'generateAttendanceSheet to be implemented'
  };
}

// ────────────────────────────────────────────────────────────
//  ALIAS FR
// ────────────────────────────────────────────────────────────

function obtenirMembres() {
  return getMembres();
}
function enregistrerMembre(d) {
  return saveMembre(d);
}
function supprimerMembre(id) {
  return deleteMembre(id);
}
function mettreAJourChecklistMembre(id, cl) {
  return updateMembreChecklist(id, cl);
}
function obtenirReunions() {
  return getReunions();
}
function enregistrerReunion(d) {
  return saveReunion(d);
}
function supprimerReunion(id) {
  return deleteReunion(id);
}
function obtenirVotes(reunionId) {
  return getVotes(reunionId);
}
function enregistrerVote(d) {
  return saveVote(d);
}
function genererCRReunion(id) {
  return genererCRCA(id);
}
function genererODJReunion(id, avecInd) {
  return genererODJ(id, avecInd);
}
function genererBAMembre(id, type) {
  return genererBA(id, type);
}
function importerResultatsBalotilo(url) {
  return fetchBalotiloResults(url);
}
function chargerToutDonneesSG() {
  return chargerDonneesSG();
}
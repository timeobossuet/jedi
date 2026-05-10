/**
 * Récupère la liste des études depuis Firestore (alimenté par sync Drive)
 */
function getStudiesFromDrive() {
  try {
    return getStudiesFromFirestore_();
  } catch (firestoreErr) {
    var firestoreMsg = firestoreErr && firestoreErr.message ? firestoreErr.message : String(firestoreErr || 'Erreur inconnue');
    Logger.log('getStudiesFromDrive: fallback Drive scan (Firestore indisponible): ' + firestoreMsg);

    try {
      return listStudiesDirectFromDrive_();
    } catch (driveErr) {
      var driveMsg = driveErr && driveErr.message ? driveErr.message : String(driveErr || 'Erreur inconnue');
      throw new Error('Chargement études impossible (Firestore + Drive). Firestore: ' + firestoreMsg + ' | Drive: ' + driveMsg);
    }
  }
}

function getStudiesFromFirestore_() {
  var studies = [];
  var resp = firestoreListDocuments_(fsEtudesCollection_(), 500);
  var docs = resp && resp.documents ? resp.documents : [];

  for (var i = 0; i < docs.length; i++) {
    var rawDoc = docs[i] || {};
    var etude = fromFirestoreDocument_(rawDoc) || {};

    var docName = String(rawDoc.name || '');
    var id = docName ? docName.split('/').pop() : '';
    if (!id) continue;

    var year = etude.ref && etude.ref.year ? String(etude.ref.year) : '';
    var code = etude.ref && etude.ref.code ? String(etude.ref.code) : '';
    var fullRef = [year, code].filter(function(v) { return !!String(v).trim(); }).join('_');
    var updatedIso = etude.drive && etude.drive.lastUpdated ? etude.drive.lastUpdated : etude.updatedAt;

    studies.push({
      id: id,
      fullRef: fullRef,
      year: year,
      ref: code,
      name: etude.name || (etude.etudeInfo && etude.etudeInfo.nom) || 'Sans titre',
      url: etude.drive && etude.drive.url ? etude.drive.url : '',
      lastUpdated: formatDateForStudiesList_(updatedIso),
      updatedAt: etude.updatedAt || '',
      status: etude.status || 'En cours',
      progress: typeof etude.progress === 'number' ? etude.progress : 0
    });
  }

  studies.sort(function(a, b) {
    return String(b.fullRef || '').localeCompare(String(a.fullRef || ''));
  });

  return studies;
}

function listStudiesDirectFromDrive_() {
  var rootFolderId = getStudiesRootFolderIdSafe_();
  var rootFolder = DriveApp.getFolderById(rootFolderId);
  var folders = rootFolder.getFolders();
  var studies = [];

  while (folders.hasNext()) {
    var folder = folders.next();
    var parsed = parseStudyFolderNameSafe_(folder.getName());
    if (!parsed) continue;

    var updatedIso = '';
    try { updatedIso = folder.getLastUpdated().toISOString(); } catch (e) { updatedIso = ''; }

    studies.push({
      id: folder.getId(),
      fullRef: parsed.fullRef,
      year: parsed.year,
      ref: parsed.code,
      name: parsed.studyName || 'Sans titre',
      url: folder.getUrl(),
      lastUpdated: formatDateForStudiesList_(updatedIso),
      updatedAt: updatedIso,
      status: 'En cours',
      progress: 0
    });
  }

  studies.sort(function(a, b) {
    return String(b.fullRef || '').localeCompare(String(a.fullRef || ''));
  });

  return studies;
}

function getStudiesRootFolderIdSafe_() {
  try {
    if (typeof getStudiesRootFolderId_ === 'function') {
      var configured = String(getStudiesRootFolderId_() || '').trim();
      if (configured) return configured;
    }
  } catch (e) {}
  return '1M4e3ZeYrEKSLc5B84QS94lovZLdji63o';
}

function parseStudyFolderNameSafe_(name) {
  try {
    if (typeof parseStudyFolderName_ === 'function') {
      var parsed = parseStudyFolderName_(name);
      if (parsed) return parsed;
    }
  } catch (e) {}

  var folderName = String(name || '').trim();
  var m = folderName.match(/^(\d{4})_(\d+)(?:\s+(.*))?$/);
  if (!m) return null;

  return {
    year: String(m[1] || ''),
    code: String(m[2] || ''),
    fullRef: String(m[1] || '') + '_' + String(m[2] || ''),
    studyName: String(m[3] || 'Sans titre').trim() || 'Sans titre'
  };
}

function formatDateForStudiesList_(isoDate) {
  if (!isoDate) return '';
  var d = new Date(isoDate);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, 'Europe/Paris', 'dd/MM/yyyy');
}

/**
 * Récupère la progression d'une seule étude (appel léger)
 */
function getStudyProgress(studyId) {
  try {
    var fsDoc = fsEtudeGet(studyId);
    if (fsDoc && fsDoc.hasOwnProperty('progress')) {
      var p = parseInt(fsDoc.progress, 10);
      return isNaN(p) ? 0 : p;
    }
    return 0;
  } catch(e) {
    return 0;
  }
}

/**
 * Récupère le dossier d'échanges/CR d'une étude.
 * Nouvelle structure prioritaire: "5. [REF] CR échanges"
 * Remplacement de la structure legacy "1. Echanges" par la nouvelle structure "5. [REF] CR échanges".
 */
function getStudyExchangeFolder(studyId) {
  var studyFolder = DriveApp.getFolderById(studyId);
  var subfolders = studyFolder.getFolders();

  function normalize_(value) {
    return String(value || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  while (subfolders.hasNext()) {
    var folder = subfolders.next();
    var folderName = String(folder.getName() || '');
    var normalized = normalize_(folderName);

    // Nouvelle structure: 5. [REF] CR échanges
    if (/^5\s/.test(normalized) && normalized.indexOf('cr') >= 0 && normalized.indexOf('echange') >= 0) {
      return folder;
    }
  }

  var exactNew = studyFolder.getFoldersByName('5. CR échanges');
  if (exactNew.hasNext()) return exactNew.next();

  var exactNewAscii = studyFolder.getFoldersByName('5. CR echanges');
  if (exactNewAscii.hasNext()) return exactNewAscii.next();

  throw new Error('Dossier de CR/échanges introuvable (attendu: "5. [REF] CR échanges").');
}

/**
 * Donne le lien du dossier "1. Echanges" pour l'étude
 */
function getStudyExchangeFolderLink(studyId) {
  try {
    var fsDoc = fsEtudeGet(studyId) || {};
    var cached = fsDoc.drive && fsDoc.drive.exchangeFolder ? fsDoc.drive.exchangeFolder : null;
    if (cached && cached.url) {
      var cachedOk = false;
      try {
        if (cached.id) {
          DriveApp.getFolderById(String(cached.id));
          cachedOk = true;
        }
      } catch (eCheck) {
        cachedOk = false;
      }

      if (cachedOk) {
      return {
        success: true,
        url: cached.url,
        id: cached.id || '',
          name: cached.name || '5. [REF] CR échanges'
        };
      }
    }
  } catch (e) {}

  var exchangeFolder = getStudyExchangeFolder(studyId);
  var payload = {
    success: true,
    url: exchangeFolder.getUrl(),
    id: exchangeFolder.getId(),
    name: exchangeFolder.getName()
  };

  try {
    var fsDocToPatch = fsEtudeGet(studyId) || {};
    var driveInfo = fsDocToPatch.drive && typeof fsDocToPatch.drive === 'object' ? fsDocToPatch.drive : {};
    driveInfo.exchangeFolder = {
      id: payload.id,
      name: payload.name,
      url: payload.url
    };
    fsEtudePatch(studyId, {
      drive: driveInfo,
      updatedAt: new Date().toISOString()
    });
  } catch (e2) {}

  return payload;
}

function getChildFolderByPredicate(parentFolder, predicateFn) {
  var folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    var f = folders.next();
    if (predicateFn(f.getName())) return f;
  }
  return null;
}

function getStudyYearAndCode(studyFolder) {
  var folderName = studyFolder.getName() || '';
  var m = folderName.match(/^(\d{4})_(\d+)/);
  if (!m) {
    throw new Error("Nom de dossier étude invalide pour extraire year/code: " + folderName);
  }
  return { year: m[1], code: m[2] };
}

function renameStudyAndSyncName(studyId, newStudyName) {
  if (!studyId) throw new Error('ID étude manquant.');

  var folder = DriveApp.getFolderById(studyId);
  var currentFolderName = String(folder.getName() || '').trim();
  var m = currentFolderName.match(/^(\d{4}_\d+)(?:\s+(.*))?$/);
  if (!m) {
    throw new Error('Nom de dossier étude invalide: ' + currentFolderName);
  }

  var reference = String(m[1] || '').trim();
  var cleanName = String(newStudyName || '').replace(/\s+/g, ' ').trim();
  if (!cleanName) {
    throw new Error('Le titre de l\'étude ne peut pas être vide.');
  }

  var newFolderName = reference + ' ' + cleanName;
  if (newFolderName !== currentFolderName) {
    folder.setName(newFolderName);
  }

  var etudeInfo = {};
  try {
    var existing = fsEtudeGet(studyId) || {};
    etudeInfo = existing.etudeInfo && typeof existing.etudeInfo === 'object' ? existing.etudeInfo : {};
  } catch (e) {
    etudeInfo = {};
  }

  etudeInfo.nom = cleanName;
  etudeInfo.reference = reference;

  fsEtudePatch(studyId, {
    name: cleanName,
    etudeInfo: etudeInfo,
    updatedAt: new Date().toISOString()
  });

  return {
    success: true,
    id: studyId,
    name: cleanName,
    reference: reference,
    folderName: newFolderName,
    url: folder.getUrl()
  };
}

// --- Alias FR (plus explicites) ---
function listerEtudesDrive() {
  return getStudiesFromDrive();
}

function obtenirProgressionEtude(studyId) {
  return getStudyProgress(studyId);
}

function obtenirDossierEchangesEtude(studyId) {
  return getStudyExchangeFolder(studyId);
}

function obtenirLienDossierEchangesEtude(studyId) {
  return getStudyExchangeFolderLink(studyId);
}

function trouverSousDossierParNom(parentFolder, predicateFn) {
  return getChildFolderByPredicate(parentFolder, predicateFn);
}

function extraireAnneeEtCodeEtude(studyFolder) {
  return getStudyYearAndCode(studyFolder);
}

function renommerEtudeEtSynchroniserNom(studyId, newStudyName) {
  return renameStudyAndSyncName(studyId, newStudyName);
}

function supprimerFichierDrive(fileId) {
  try {
    const activeUser = Session.getActiveUser().getEmail();
    Logger.log("[" + activeUser + "] tentative de suppression du fichier " + fileId);
    let file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    return { success: true, message: "Fichier supprimé avec succès." };
  } catch(e) {
    Logger.log("Erreur suppression: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

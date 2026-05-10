/**
 * Synchronisation lecture Drive -> Firestore (sans écriture métier sur Drive)
 *
 * Objectif:
 * - Scanner périodiquement les dossiers d'études sur Drive
 * - Mettre à jour Firestore avec les métadonnées de lecture (nom, ref, url, date maj)
 * - Mettre à jour le statut devis détecté via arborescence/fichiers
 *
 * Les opérations de création, déplacement, suppression de documents restent gérées
 * dans les scripts métiers existants (async côté UI).
 */

function getStudiesRootFolderId_() {
  var configured = '';
  try {
    configured = PropertiesService.getScriptProperties().getProperty('STUDIES_DRIVE_ROOT_FOLDER_ID') || '';
  } catch (e) {
    configured = '';
  }
  return String(configured || '1M4e3ZeYrEKSLc5B84QS94lovZLdji63o').trim();
}

function setLastDriveReadSyncAt_(isoDate) {
  try {
    PropertiesService.getScriptProperties().setProperty('LAST_DRIVE_READ_SYNC_AT', String(isoDate || new Date().toISOString()));
  } catch (e) {}
}

function getLastDriveReadSyncAt_() {
  try {
    return PropertiesService.getScriptProperties().getProperty('LAST_DRIVE_READ_SYNC_AT') || '';
  } catch (e) {
    return '';
  }
}

function getDriveReadSyncStatus() {
  return {
    success: true,
    lastSyncAt: getLastDriveReadSyncAt_(),
    rootFolderId: getStudiesRootFolderId_()
  };
}

function setStudiesDriveRootFolderId(folderId) {
  if (!folderId || !String(folderId).trim()) {
    throw new Error('ID du dossier racine Drive des études manquant.');
  }
  PropertiesService.getScriptProperties().setProperty('STUDIES_DRIVE_ROOT_FOLDER_ID', String(folderId).trim());
  return {
    success: true,
    folderId: String(folderId).trim()
  };
}

function parseStudyFolderName_(name) {
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

function computeDevisStatusFromDrive_(studyFolder, year, code) {
  var devisName = (year || 'XXXX') + '_' + (code || 'XXX') + '_DEVIS';
  var pdfName = devisName + '_relecture.pdf';
  var result = {
    hasDraft: false,
    draftUrl: '',
    draftName: devisName,
    hasPdf: false,
    pdfUrl: '',
    hasValide: false,
    valideUrl: '',
    devisName: devisName,
    relectureDemandee: false,
    chatUrl: 'https://chat.google.com/room/AAQApAX9EAQ'
  };

  var docsFolder = getChildFolderByPredicate(studyFolder, function(n) { return /^2\.\s*Documents$/i.test(n); });
  if (!docsFolder) return result;

  var reviewFolder = getChildFolderByPredicate(docsFolder, function(n) { return /^1\.\s*En\s*relecture$/i.test(n); });
  if (reviewFolder) {
    var em = reviewFolder.getFilesByName(devisName);
    if (em.hasNext()) {
      result.hasDraft = true;
      result.draftUrl = em.next().getUrl();
    }

    var pm = reviewFolder.getFilesByName(pdfName);
    if (pm.hasNext()) {
      result.hasPdf = true;
      result.pdfUrl = pm.next().getUrl();
      result.relectureDemandee = true;
    }
  }

  var valideFolder = getChildFolderByPredicate(docsFolder, function(n) { return /^2\.\s*Valid/i.test(n); });
  if (valideFolder) {
    var vf = valideFolder.getFiles();
    while (vf.hasNext()) {
      var f = vf.next();
      if (f.getName().indexOf(devisName) !== -1) {
        result.hasValide = true;
        result.valideUrl = f.getUrl();
        break;
      }
    }
  }

  return result;
}

function computeCdcStatusFromDrive_(studyFolder, year, code) {
  var cdcName = (year || 'XXXX') + '_' + (code || 'XXX') + '_CDC';
  var result = {
    hasDraft: false,
    draftUrl: '',
    cdcName: cdcName
  };

  var reviewFolder = null;

  if (typeof findEditableDocumentsFolder_ === 'function') {
    reviewFolder = findEditableDocumentsFolder_(studyFolder);
  }

  if (!reviewFolder) {
    var docsFolder = getChildFolderByPredicate(studyFolder, function(n) { return /^2\.\s*Documents$/i.test(n); });
    if (docsFolder) {
      reviewFolder = getChildFolderByPredicate(docsFolder, function(n) { return /^1\.\s*En\s*relecture$/i.test(n); });
    }
  }

  if (!reviewFolder) return result;

  var latest = null;
  var highestVersion = 0;
  var regex = new RegExp('^' + cdcName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:_V(\\d+))?(?:\\.[a-z0-9]+)?$', 'i');
  var files = reviewFolder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var mime = String(file.getMimeType() || '').toLowerCase();
    if (mime === 'application/pdf') continue;

    var match = String(file.getName() || '').match(regex);
    if (!match) continue;

    var version = match[1] ? parseInt(match[1], 10) : 1;
    if (!Number.isFinite(version)) version = 1;
    if (!latest || version >= highestVersion) {
      latest = file;
      highestVersion = version;
    }
  }

  if (latest) {
    result.hasDraft = true;
    result.draftUrl = latest.getUrl();
    result.cdcName = latest.getName() || cdcName;
  }

  return result;
}

function buildStudyDriveSnapshot_(folder, existing, forceFullScan) {
  var parsed = parseStudyFolderName_(folder.getName());
  if (!parsed) return null;

  var folderLastUpdated = folder.getLastUpdated().toISOString();
  var sameAsCached = !!(
    !forceFullScan &&
    existing && existing.drive &&
    existing.drive.lastUpdated &&
    String(existing.drive.lastUpdated) === String(folderLastUpdated)
  );

  var exchangeFolderMeta = { id: '', name: '', url: '' };
  if (sameAsCached && existing.drive && existing.drive.exchangeFolder) {
    exchangeFolderMeta = existing.drive.exchangeFolder;
  } else {
    try {
      var exchangeFolder = getStudyExchangeFolder(folder.getId());
      if (exchangeFolder) {
        exchangeFolderMeta = {
          id: exchangeFolder.getId(),
          name: exchangeFolder.getName(),
          url: exchangeFolder.getUrl()
        };
      }
    } catch (e) {}
  }

  var devisStatus = null;
  var cdcStatus = null;
  if (sameAsCached && existing.devisStatus && typeof existing.devisStatus === 'object') {
    devisStatus = existing.devisStatus;
    cdcStatus = existing.cdcStatus;
  } else {
    devisStatus = computeDevisStatusFromDrive_(folder, parsed.year, parsed.code);
    cdcStatus = computeCdcStatusFromDrive_(folder, parsed.year, parsed.code);
  }

  return {
    id: folder.getId(),
    name: parsed.studyName,
    ref: {
      year: parsed.year,
      code: parsed.code
    },
    fullRef: parsed.fullRef,
    drive: {
      folderId: folder.getId(),
      folderName: folder.getName(),
      url: folder.getUrl(),
      lastUpdated: folderLastUpdated,
      exchangeFolder: exchangeFolderMeta
    },
    devisStatus: devisStatus,
    cdcStatus: cdcStatus,
    reusedCachedDriveReads: sameAsCached
  };
}

function syncStudyFromDriveToFirestore_(folder, forceFullScan) {
  var existing = {};
  var folderId = '';
  try {
    folderId = folder.getId();
    existing = fsEtudeGet(folderId) || {};
  } catch (e) {
    existing = {};
  }

  var snapshot = buildStudyDriveSnapshot_(folder, existing, forceFullScan === true);
  if (!snapshot) return { skipped: true, reason: 'invalid-folder-name' };

  var etudeInfo = existing.etudeInfo && typeof existing.etudeInfo === 'object' ? existing.etudeInfo : {};
  etudeInfo.nom = snapshot.name;
  etudeInfo.reference = snapshot.fullRef;

  if (!etudeInfo.categories && etudeInfo.categorie) {
    etudeInfo.categories = String(etudeInfo.categorie)
      .split(',')
      .map(function(v) { return String(v || '').trim(); })
      .filter(function(v) { return !!v; });
  }

  var overviewSnapshot = {
    name: snapshot.name,
    reference: snapshot.fullRef,
    status: existing.status || 'En cours de réalisation',
    progress: typeof existing.progress === 'number' ? existing.progress : 0,
    contextShort: etudeInfo.contexte ? String(etudeInfo.contexte).substring(0, 280) : '',
    problematiqueShort: etudeInfo.problematique ? String(etudeInfo.problematique).substring(0, 280) : '',
    categories: Array.isArray(etudeInfo.categories)
      ? etudeInfo.categories
      : (etudeInfo.categorie ? String(etudeInfo.categorie).split(',').map(function(v) { return String(v || '').trim(); }).filter(function(v) { return !!v; }) : []),
    devisStatus: snapshot.devisStatus,
    cdcStatus: snapshot.cdcStatus,
    driveLastUpdated: snapshot.drive.lastUpdated,
    updatedAt: new Date().toISOString()
  };

  var patch = {
    name: snapshot.name,
    ref: snapshot.ref,
    drive: snapshot.drive,
    etudeInfo: etudeInfo,
    studyOverviewSnapshot: overviewSnapshot,
    devisStatus: snapshot.devisStatus,
    cdcStatus: snapshot.cdcStatus,
    source: 'drive-sync',
    lastDriveSyncAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  if (!existing.status) patch.status = 'En cours de réalisation';
  if (typeof existing.progress === 'undefined' || existing.progress === null) patch.progress = 0;

  fsEtudePatch(snapshot.id, patch);
  try {
    if (typeof invalidateStudyLoadCaches_ === 'function') {
      invalidateStudyLoadCaches_(snapshot.id);
    }
  } catch (e) {}

  return {
    skipped: false,
    studyId: snapshot.id,
    name: snapshot.name,
    fullRef: snapshot.fullRef,
    reusedCachedDriveReads: !!snapshot.reusedCachedDriveReads
  };
}

function syncAllDriveReadsToFirestore(forceFullScan) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Une synchronisation Drive est déjà en cours. Réessayez dans quelques secondes.');
  }

  try {
  var doFullScan = forceFullScan === true;
  var rootFolderId = getStudiesRootFolderId_();
  var rootFolder = DriveApp.getFolderById(rootFolderId);
  var folders = rootFolder.getFolders();

  var scanned = 0;
  var synced = 0;
  var skipped = 0;
  var reusedCachedDriveReads = 0;

  while (folders.hasNext()) {
    var folder = folders.next();
    scanned++;

    var result = syncStudyFromDriveToFirestore_(folder, doFullScan);
    if (result && result.skipped) skipped++;
    else {
      synced++;
      if (result.reusedCachedDriveReads) reusedCachedDriveReads++;
    }
  }

  var syncedAt = new Date().toISOString();
  setLastDriveReadSyncAt_(syncedAt);

  return {
    success: true,
    rootFolderId: rootFolderId,
    fullScan: doFullScan,
    scanned: scanned,
    synced: synced,
    skipped: skipped,
    reusedCachedDriveReads: reusedCachedDriveReads,
    syncedAt: syncedAt
  };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function syncSingleStudyDriveReadToFirestore(studyId, forceFullScan) {
  if (!studyId) throw new Error('ID étude manquant.');
  var folder = DriveApp.getFolderById(studyId);
  var result = syncStudyFromDriveToFirestore_(folder, forceFullScan === true);
  var syncedAt = new Date().toISOString();
  setLastDriveReadSyncAt_(syncedAt);
  return { success: true, result: result, syncedAt: syncedAt };
}

function runDriveReadSyncToFirestore() {
  return syncAllDriveReadsToFirestore(false);
}

function createDriveReadSyncTrigger(minutes) {
  var everyMinutes = Math.max(1, parseInt(minutes, 10) || 5);
  ScriptApp.newTrigger('runDriveReadSyncToFirestore')
    .timeBased()
    .everyMinutes(everyMinutes)
    .create();

  return {
    success: true,
    handler: 'runDriveReadSyncToFirestore',
    everyMinutes: everyMinutes
  };
}

function clearDriveReadSyncTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;

  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (t.getHandlerFunction && t.getHandlerFunction() === 'runDriveReadSyncToFirestore') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }

  return { success: true, removed: removed };
}

function runDriveReadFullSyncToFirestore() {
  return syncAllDriveReadsToFirestore(true);
}

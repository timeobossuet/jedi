function studyRawDataCacheKey_(studyId) {
  return 'STUDY_RAW_DATA_' + String(studyId);
}

function studyOverviewNativeCacheKey_(studyId) {
  return 'STUDY_OVERVIEW_NATIVE_' + String(studyId);
}

function createGoogleMeetEvent_(data) {
  try {
    var calendarId = 'primary';
    var event = {
      summary: data.title,
      description: data.notes || '',
      start: {
        dateTime: data.startIso,
        timeZone: Session.getScriptTimeZone()
      },
      end: {
        dateTime: data.endIso,
        timeZone: Session.getScriptTimeZone()
      },
      attendees: data.attendees.map(function(email) { return {email: email}; }),
    };

    if (data.withMeet) {
      event.conferenceData = {
        createRequest: {
          requestId: Utilities.getUuid(),
          conferenceSolutionKey: {
            type: 'hangoutsMeet'
          }
        }
      };
    } else if (data.location) {
      event.location = data.location;
    }

    if (data.addCalendar) {
      var createdEvent = Calendar.Events.insert(event, calendarId, {conferenceDataVersion: 1, sendUpdates: 'all'});
      return {
        success: true,
        eventId: createdEvent.id,
        eventUrl: createdEvent.htmlLink,
        meetLink: (data.withMeet && createdEvent.conferenceData && createdEvent.conferenceData.entryPoints) ? createdEvent.conferenceData.entryPoints[0].uri : ''
      };
    } else {
      return { success: true, eventId: '', eventUrl: '', meetLink: '' };
    }
  } catch (e) {
    Logger.log('Erreur création Meet: ' + e.message);
    return { success: false, error: e.message };
  }
}

function studySectionsNativeCacheKey_(studyId, sections) {
  return 'STUDY_SECTIONS_NATIVE_' + String(studyId) + '_' + String(sections || '').toLowerCase();
}

function getCachedJsonByKey_(cacheKey) {
  try {
    var raw = CacheService.getScriptCache().get(cacheKey);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function setCachedJsonByKey_(cacheKey, value, ttlSec) {
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(value), Math.max(30, parseInt(ttlSec, 10) || 60));
  } catch (e) {}
}

function invalidateStudyLoadCaches_(studyId) {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(studyRawDataCacheKey_(studyId));
    cache.remove(studyOverviewNativeCacheKey_(studyId));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'todos'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'planning'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'devis'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'cdc'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'categories'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'todos,planning'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'todos,devis'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'todos,cdc'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'planning,devis'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'planning,cdc'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'devis,cdc'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'todos,planning,devis,categories'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'planning,devis,cdc'));
    cache.remove(studySectionsNativeCacheKey_(studyId, 'todos,planning,devis,cdc,categories'));
  } catch (e) {}
}

function buildStudyOverviewSnapshotNative_(doc) {
  var safe = doc || {};
  var etudeInfo = safe.etudeInfo || {};
  var defaults = {
    Description: safe.description || '',
    Status: safe.status || 'En cours de réalisation',
    Planning_StartDate: safe.planningStartDate || '',
    EtudeInfo: etudeInfo,
    ClientInfo: safe.clientInfo || {},
    AvantProjet_Steps: safe.avantProjetSteps || { client: {}, interv: {} },
    Negociation_Steps: safe.negociationSteps || etudeInfo.negociationSteps || { client: {}, interv: {} },
    progress: typeof safe.progress === 'number' ? safe.progress : Number(etudeInfo.progress || 0) || 0,
    updatedAt: safe.updatedAt || '',
    studyOverviewSnapshot: safe.studyOverviewSnapshot || {}
  };
  return defaults;
}

function mapFirestoreStudyToRawData_(full) {
  var year = full.ref && full.ref.year ? String(full.ref.year) : '';
  var code = full.ref && full.ref.code ? String(full.ref.code) : '';
  var reference = [year, code].filter(function(v) { return !!String(v).trim(); }).join('_');

  return {
    Description: full.description || '',
    Status: full.status || '',
    Planning_StartDate: full.planningStartDate || '',
    ClientInfo: full.clientInfo || {},
    EtudeInfo: full.etudeInfo || {},
    Planning_Phases: full.phases || [],
    Todo_Chargé: full.todos || [],
    AvantProjet_Steps: full.avantProjetSteps || { client: {}, interv: {} },
    Negociation_Steps: full.negociationSteps || { client: {}, interv: {} },
    negociationDocStates: full.negociationDocStates || {},
    year: year,
    code: code,
    reference: reference,
    nom: full.name || ''
  };
}

function getStudyFirestoreFullSafe_(studyId) {
  var full = fsEtudeGetFull(studyId);
  if (!full || typeof full !== 'object') {
    throw new Error('Étude introuvable dans Firestore: ' + studyId);
  }
  return full;
}

/**
 * Sauvegarde une donnée spécifique pour une étude (Clé/Valeur) avec buffer temporel global
 */
function saveStudyData(studyId, key, value, forceFlush) {
  var props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  
  // Tente d'obtenir le lock (max 3 secondes d'attente)
  var hasLock = lock.tryLock(3000);
  
  if (!hasLock) {
    // Si lock échoue, on écrit en direct
    if (key === 'FLUSH') return { success: true, backend: 'firestore_direct_flush', mappedTo: '' };
    invalidateStudyLoadCaches_(studyId);
    var fsResult = enregistrerDonneeEtudeFirestore(studyId, key, value);
    invalidateStudyLoadCaches_(studyId);
    return {
      success: true,
      backend: 'firestore_direct',
      mappedTo: fsResult && fsResult.mappedTo ? fsResult.mappedTo : ''
    };
  }
  
  try {
    var bufferKey = 'FS_BUF_' + studyId;
    var timerKey = 'FS_TIM_' + studyId;
    
    var existingBuf = props.getProperty(bufferKey);
    var buf = existingBuf ? JSON.parse(existingBuf) : [];
    
    if (key !== 'FLUSH') {
      buf.push({ key: key, value: value });
    }
    
    var now = Date.now();
    var lastPush = parseInt(props.getProperty(timerKey) || '0', 10);
    var mappedPaths = [];
    
    if (forceFlush || (key === 'FLUSH') || (now - lastPush > 3500) || buf.length >= 20) {
       if (buf.length > 0) {
         invalidateStudyLoadCaches_(studyId);
         if (typeof enregistrerDonneeEtudeFirestoreBatch === 'function') {
           var batchResult = enregistrerDonneeEtudeFirestoreBatch(studyId, buf);
           if (batchResult && batchResult.mappedTo) {
             mappedPaths.push(batchResult.mappedTo);
           }
         } else {
           // Fallback if batch is not available
           for (var i = 0; i < buf.length; i++) {
             var res = enregistrerDonneeEtudeFirestore(studyId, buf[i].key, buf[i].value);
             if (res && res.mappedTo) mappedPaths.push(res.mappedTo);
           }
         }
         invalidateStudyLoadCaches_(studyId);
       }
       props.deleteProperty(bufferKey);
       props.setProperty(timerKey, now.toString());
       
       return {
         success: true,
         backend: 'firestore_batch',
         mappedTo: mappedPaths.join(',')
       };
    } else {
       props.setProperty(bufferKey, JSON.stringify(buf));
       return {
         success: true,
         backend: 'firestore_buffered',
         mappedTo: 'pending'
       };
    }
  } finally {
    lock.releaseLock();
  }
}

// --- Helper: Get All Study Data ---
function getStudyData(studyId) {
  var cacheKey = studyRawDataCacheKey_(studyId);
  var cached = getCachedJsonByKey_(cacheKey);
  if (cached && typeof cached === 'object') {
    return cached;
  }

  var full = getStudyFirestoreFullSafe_(studyId);
  var mapped = mapFirestoreStudyToRawData_(full);
  setCachedJsonByKey_(cacheKey, mapped, 90);
  return mapped;
}

function getStudyOverviewSnapshotNative(studyId) {
  var cacheKey = studyOverviewNativeCacheKey_(studyId);
  var cached = getCachedJsonByKey_(cacheKey);
  if (cached && typeof cached === 'object') return cached;

  var doc = fsEtudeGet(studyId) || {};
  if (!doc || typeof doc !== 'object') {
    throw new Error('Étude introuvable dans Firestore: ' + studyId);
  }

  var payload = buildStudyOverviewSnapshotNative_(doc);
  setCachedJsonByKey_(cacheKey, payload, 120);
  return payload;
}

function getStudySectionsNative(studyId, sections) {
  var requested = Object.prototype.toString.call(sections) === '[object Array]' ? sections : [];
  var normalized = requested
    .map(function(s) { return String(s || '').trim().toLowerCase(); })
    .filter(function(s) { return !!s; })
    .sort();
  var key = normalized.join(',');
  var cacheKey = studySectionsNativeCacheKey_(studyId, key);
  var cached = getCachedJsonByKey_(cacheKey);
  if (cached && typeof cached === 'object') return cached;

  var doc = fsEtudeGet(studyId) || {};
  var result = {
    updatedAt: doc.updatedAt || '',
    todos: null,
    planning: null,
    devisStatus: null,
    cdcStatus: null,
    categories: null
  };

  if (normalized.indexOf('todos') >= 0) {
    var todos = Object.prototype.toString.call(doc.todosSnapshot) === '[object Array]' ? doc.todosSnapshot : null;
    if (!todos) {
      try { todos = (fsEtudeGetFull(studyId) || {}).todos || []; }
      catch (e) { todos = []; }
    }
    result.todos = todos;
  }

  if (normalized.indexOf('planning') >= 0) {
    var phases = Object.prototype.toString.call(doc.phasesSnapshot) === '[object Array]' ? doc.phasesSnapshot : null;
    if (!phases) {
      try { phases = (fsEtudeGetFull(studyId) || {}).phases || []; }
      catch (e2) { phases = []; }
    }
    result.planning = {
      phases: phases,
      startDate: doc.planningStartDate || ''
    };
  }

  if (normalized.indexOf('devis') >= 0) {
    if (doc.devisStatus && typeof doc.devisStatus === 'object') {
      result.devisStatus = doc.devisStatus;
    } else {
      result.devisStatus = checkDevisStatus(studyId);
    }
  }

  if (normalized.indexOf('cdc') >= 0) {
    if (doc.cdcStatus && typeof doc.cdcStatus === 'object') {
      result.cdcStatus = doc.cdcStatus;
    } else {
      if (typeof checkCdcStatus === 'function') {
        result.cdcStatus = checkCdcStatus(studyId);
      } else {
        result.cdcStatus = null;
      }
    }
  }

  if (normalized.indexOf('categories') >= 0) {
    result.categories = listerCategoriesEtudesFirestore();
  }

  setCachedJsonByKey_(cacheKey, result, 120);
  return result;
}

// --- Alias FR (plus explicites) ---
function obtenirFeuilleDonneesEtude(studyId) {
  throw new Error('Stockage legacy supprimé: utiliser Firestore uniquement.');
}

function enregistrerDonneeEtude(studyId, key, value) {
  return saveStudyData(studyId, key, value);
}

function obtenirDonneesEtude(studyId) {
  return getStudyData(studyId);
}

function obtenirOverviewEtude(studyId) {
  return getStudyOverviewSnapshotNative(studyId);
}

function obtenirSectionsEtude(studyId, sections) {
  return getStudySectionsNative(studyId, sections);
}


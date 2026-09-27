function buildDefaultDocWorkflowStatus_(rawData, docKey) {
  var cfg = getDocWorkflowConfig_(docKey);
  var reference = getStudyReferenceForDocs_(rawData);
  var docName = reference + '_' + cfg.suffix;
  var status = {
    key: cfg.key,
    label: cfg.label,
    suffix: cfg.suffix,
    templateName: '',
    hasDraft: false,
    draftUrl: '',
    draftName: docName,
    hasPdf: false,
    pdfUrl: '',
    hasValide: false,
    valideUrl: '',
    archiveRequired: false,
    signedAt: '',
    docName: docName,
    relectureDemandee: false,
    chatUrl: DOC_WORKFLOW_CHAT_URL_
  };

  if (cfg.key === 'devis') {
    status.devisName = docName;
  }
  return status;
}

function docWorkflowStatusCacheKey_(studyId, docKey) {
  return 'DOC_WORKFLOW_STATUS_' + String(docKey || '').toLowerCase() + '_' + String(studyId || '');
}

function getCachedDocWorkflowStatus_(studyId, docKey) {
  try {
    var raw = CacheService.getScriptCache().get(docWorkflowStatusCacheKey_(studyId, docKey));
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function setCachedDocWorkflowStatus_(studyId, docKey, statusObj) {
  try {
    CacheService.getScriptCache().put(docWorkflowStatusCacheKey_(studyId, docKey), JSON.stringify(statusObj || {}), 60);
  } catch (e) {}
}

function invalidateDocWorkflowStatusCache_(studyId, docKey) {
  var cache = CacheService.getScriptCache();
  try {
    cache.remove(docWorkflowStatusCacheKey_(studyId, docKey));
    if (typeof invalidateStudyLoadCaches_ === 'function') invalidateStudyLoadCaches_(studyId);
  } catch (e) {}
}

function adaptDevisStatusLegacyShape_(status) {
  var s = cloneFlatObject_(status || {});
  var docName = String(s.docName || s.devisName || s.draftName || '').trim();
  if (docName) {
    s.docName = docName;
    s.devisName = docName;
    if (!s.draftName) s.draftName = docName;
  }
  return s;
}

function patchDocWorkflowStatusFirestore_(studyId, docKey, updateObj, rawData) {
  try {
    var etudeData = fsEtudeGet(studyId) || {};
    var negDocStates = etudeData.negociationDocStates || {};
    var existingState = negDocStates[docKey] || {};

    for (var k in updateObj) { existingState[k] = updateObj[k]; }

    negDocStates[docKey] = existingState;
    fsEtudePatch(studyId, { negociationDocStates: negDocStates });
  } catch (e) {
    console.error('Erreur mise à jour statuts doc negociation : ' + e.message);
  }
}

function checkWorkflowDocumentStatus_(studyId, docKey) {
  var cached = getCachedDocWorkflowStatus_(studyId, docKey);
  if (cached) return cached;

  var etudeData = {};
  try {
    etudeData = fsEtudeGet(studyId) || {};
  } catch (e) {
    console.error('Erreur checkWorkflowDocumentStatus_ (fsEtudeGet):', e);
  }

  var states = etudeData.negociationDocStates || {};
  var status = states[docKey];
  var rawDataForDefault = getStudyData(studyId) || {};

  if (!status || typeof status !== 'object') {
    status = buildDefaultDocWorkflowStatus_(rawDataForDefault, docKey);
  } else {
    var def = buildDefaultDocWorkflowStatus_(rawDataForDefault, docKey);
    for (var k in def) {
      if (status[k] === undefined) status[k] = def[k];
    }
  }

  setCachedDocWorkflowStatus_(studyId, docKey, status);
  return status;
}
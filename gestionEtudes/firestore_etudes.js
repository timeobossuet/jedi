/**
 * Firestore bridge (Apps Script) for ERP Études
 *
 * Ce module fournit:
 * - Auth service account -> OAuth2 access token (sans lib externe)
 * - CRUD basique collection etudes
 *
 * Script Properties attendues:
 * - FIREBASE_PROJECT_ID
 * - FIREBASE_CLIENT_EMAIL
 * - FIREBASE_PRIVATE_KEY
 * - FIRESTORE_DATABASE_ID (optionnel, défaut: (default))
 * - FIREBASE_TOKEN_URI (optionnel, défaut: https://oauth2.googleapis.com/token)
 */

var FIRESTORE_CONFIG = {
  baseUrl: 'https://firestore.googleapis.com/v1',
  tokenScope: 'https://www.googleapis.com/auth/datastore',
  etudesCollection: 'etudes'
};

function setFirestoreSettings(projectId, clientEmail, privateKey, databaseId) {
  if (!projectId || !clientEmail || !privateKey) {
    throw new Error('Paramètres requis: projectId, clientEmail, privateKey');
  }
  var props = PropertiesService.getScriptProperties();
  props.setProperty('FIREBASE_PROJECT_ID', String(projectId).trim());
  props.setProperty('FIREBASE_CLIENT_EMAIL', String(clientEmail).trim());
  props.setProperty('FIREBASE_PRIVATE_KEY', String(privateKey));
  props.setProperty('FIRESTORE_DATABASE_ID', databaseId ? String(databaseId).trim() : '(default)');
  if (!props.getProperty('FIREBASE_TOKEN_URI')) {
    props.setProperty('FIREBASE_TOKEN_URI', 'https://oauth2.googleapis.com/token');
  }
  return { success: true };
}

function validateFirestoreSettings() {
  var cfg = getFirestoreSettings_();
  return {
    success: true,
    projectId: cfg.projectId,
    databaseId: cfg.databaseId,
    hasClientEmail: !!cfg.clientEmail,
    hasPrivateKey: !!cfg.privateKey
  };
}

function getFirestoreSettings_() {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty('FIREBASE_PROJECT_ID');
  var clientEmail = props.getProperty('FIREBASE_CLIENT_EMAIL');
  var privateKey = props.getProperty('FIREBASE_PRIVATE_KEY');
  var databaseId = props.getProperty('FIRESTORE_DATABASE_ID') || '(default)';
  var tokenUri = props.getProperty('FIREBASE_TOKEN_URI') || 'https://oauth2.googleapis.com/token';

  if (!projectId) throw new Error('ScriptProperty manquante: FIREBASE_PROJECT_ID');
  if (!clientEmail) throw new Error('ScriptProperty manquante: FIREBASE_CLIENT_EMAIL');
  if (!privateKey) throw new Error('ScriptProperty manquante: FIREBASE_PRIVATE_KEY');

  return {
    projectId: String(projectId).trim(),
    clientEmail: String(clientEmail).trim(),
    privateKey: normalizePrivateKey_(privateKey),
    databaseId: String(databaseId).trim(),
    tokenUri: String(tokenUri).trim()
  };
}

function normalizePrivateKey_(privateKeyRaw) {
  var key = String(privateKeyRaw || '');
  return key.indexOf('\\n') >= 0 ? key.replace(/\\n/g, '\n') : key;
}

function getFirestoreAccessToken_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'FS_ACCESS_TOKEN';
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var cfg = getFirestoreSettings_();
  var now = Math.floor(Date.now() / 1000);
  var header = { alg: 'RS256', typ: 'JWT' };
  var claim = {
    iss: cfg.clientEmail,
    sub: cfg.clientEmail,
    aud: cfg.tokenUri,
    scope: FIRESTORE_CONFIG.tokenScope,
    iat: now,
    exp: now + 3600
  };

  var encodedHeader = base64UrlEncodeJson_(header);
  var encodedClaim = base64UrlEncodeJson_(claim);
  var unsignedJwt = encodedHeader + '.' + encodedClaim;

  var signatureBytes;
  try {
    signatureBytes = Utilities.computeRsaSha256Signature(unsignedJwt, cfg.privateKey);
  } catch (sigErr) {
    var sigMsg = String((sigErr && sigErr.message) || sigErr || '');
    var normalized = sigMsg.toLowerCase();
    if (normalized.indexOf('argument non valide : key') !== -1 || normalized.indexOf('invalid argument: key') !== -1) {
      throw new Error(
        'FIREBASE_PRIVATE_KEY invalide ou mal formatée. ' +
        'Vérifiez la Script Property FIREBASE_PRIVATE_KEY (clé complète avec -----BEGIN PRIVATE KEY----- / -----END PRIVATE KEY-----, ' +
        'et retours à la ligne correctement encodés).'
      );
    }
    throw sigErr;
  }
  var signature = Utilities.base64EncodeWebSafe(signatureBytes).replace(/=+$/g, '');
  var assertion = unsignedJwt + '.' + signature;

  var tokenResp = UrlFetchApp.fetch(cfg.tokenUri, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: assertion
    },
    muteHttpExceptions: true
  });

  var code = tokenResp.getResponseCode();
  var body = tokenResp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Token OAuth Firestore invalide (' + code + '): ' + body);
  }

  var parsed = JSON.parse(body);
  var token = parsed.access_token;
  var expiresIn = Number(parsed.expires_in) || 3600;
  if (!token) throw new Error('Réponse token sans access_token');

  cache.put(cacheKey, token, Math.max(60, Math.min(3500, expiresIn - 60)));
  return token;
}

function base64UrlEncodeJson_(obj) {
  var json = JSON.stringify(obj);
  return Utilities.base64EncodeWebSafe(json).replace(/=+$/g, '');
}

function firestoreDbBaseUrl_() {
  var cfg = getFirestoreSettings_();
  return FIRESTORE_CONFIG.baseUrl + '/projects/' + encodeURIComponent(cfg.projectId) + '/databases/' + encodeURIComponent(cfg.databaseId);
}

function firestoreDbResourceRoot_() {
  var cfg = getFirestoreSettings_();
  return 'projects/' + cfg.projectId + '/databases/' + cfg.databaseId;
}

function firestoreRequest_(method, path, payload, queryParams) {
  var token = getFirestoreAccessToken_();
  var url = firestoreDbBaseUrl_() + path;
  if (queryParams) {
    var qs = [];
    for (var key in queryParams) {
      if (!queryParams.hasOwnProperty(key)) continue;
      var val = queryParams[key];
      if (val === undefined || val === null) continue;
      if (Object.prototype.toString.call(val) === '[object Array]') {
        for (var i = 0; i < val.length; i++) {
          qs.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(val[i])));
        }
      } else {
        qs.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(val)));
      }
    }
    if (qs.length) url += '?' + qs.join('&');
  }

  var options = {
    method: method,
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  if (payload !== undefined && payload !== null) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var text = resp.getContentText();

  if (code < 200 || code >= 300) {
    var details = extractFirestoreErrorDetails_(text);
    var baseMsg = 'Firestore ' + method + ' ' + path + ' -> ' + code + ' : ' + text;
    if (code === 404 && details.indexOf('database (default) does not exist') !== -1) {
      var cfg = getFirestoreSettings_();
      throw new Error(
        'Base Firestore introuvable pour le projet "' + cfg.projectId + '" (databaseId=' + cfg.databaseId + '). ' +
        'Ouvrez Firebase Console et activez Firestore pour CE projet, puis réessayez. ' +
        'Si vous utilisez un autre projet, mettez à jour FIREBASE_PROJECT_ID via setFirestoreSettings(...). ' +
        'Détail API: ' + text
      );
    }
    if (code === 403 && (details.indexOf('missing or insufficient permissions') !== -1 || details.indexOf('permission_denied') !== -1)) {
      var cfg403 = getFirestoreSettings_();
      throw new Error(
        'Permission refusée sur Firestore. ' +
        'Service account configuré: ' + cfg403.clientEmail + '. ' +
        'Projet ciblé: ' + cfg403.projectId + '. ' +
        'Vérifiez que CE compte de service a un rôle IAM sur le projet Firestore (minimum: Cloud Datastore User) ' +
        'et que l\'API Cloud Firestore est activée sur ce projet. ' +
        'Détail API: ' + text
      );
    }
    throw new Error(baseMsg);
  }
  if (!text) return {};
  return JSON.parse(text);
}

function extractFirestoreErrorDetails_(rawText) {
  try {
    var parsed = JSON.parse(rawText || '{}');
    var message = parsed && parsed.error && parsed.error.message ? String(parsed.error.message) : '';
    return message.toLowerCase();
  } catch (e) {
    return String(rawText || '').toLowerCase();
  }
}

function toFirestoreFields_(obj) {
  var fields = {};
  if (!obj || typeof obj !== 'object') return fields;
  for (var key in obj) {
    if (!obj.hasOwnProperty(key)) continue;
    fields[key] = toFirestoreValue_(obj[key]);
  }
  return fields;
}

function toFirestoreValue_(value) {
  if (value === null || value === undefined) return { nullValue: null };

  var t = typeof value;
  if (t === 'string') return { stringValue: value };
  if (t === 'boolean') return { booleanValue: value };
  if (t === 'number') {
    if (!isFinite(value)) return { nullValue: null };
    if (Math.floor(value) === value) return { integerValue: String(value) };
    return { doubleValue: value };
  }

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return { timestampValue: value.toISOString() };
  }

  if (Object.prototype.toString.call(value) === '[object Array]') {
    var arr = [];
    for (var i = 0; i < value.length; i++) {
      arr.push(toFirestoreValue_(value[i]));
    }
    return { arrayValue: { values: arr } };
  }

  var mapFields = {};
  for (var key in value) {
    if (!value.hasOwnProperty(key)) continue;
    mapFields[key] = toFirestoreValue_(value[key]);
  }
  return { mapValue: { fields: mapFields } };
}

function fromFirestoreDocument_(doc) {
  if (!doc || !doc.fields) return {};
  return fromFirestoreFields_(doc.fields);
}

function fromFirestoreFields_(fields) {
  var out = {};
  for (var key in fields) {
    if (!fields.hasOwnProperty(key)) continue;
    out[key] = fromFirestoreValue_(fields[key]);
  }
  return out;
}

function fromFirestoreValue_(wrapped) {
  if (!wrapped || typeof wrapped !== 'object') return null;
  if (wrapped.hasOwnProperty('nullValue')) return null;
  if (wrapped.hasOwnProperty('stringValue')) return wrapped.stringValue;
  if (wrapped.hasOwnProperty('booleanValue')) return !!wrapped.booleanValue;
  if (wrapped.hasOwnProperty('integerValue')) return Number(wrapped.integerValue);
  if (wrapped.hasOwnProperty('doubleValue')) return Number(wrapped.doubleValue);
  if (wrapped.hasOwnProperty('timestampValue')) return wrapped.timestampValue;
  if (wrapped.hasOwnProperty('mapValue')) return fromFirestoreFields_(wrapped.mapValue.fields || {});
  if (wrapped.hasOwnProperty('arrayValue')) {
    var values = (wrapped.arrayValue && wrapped.arrayValue.values) ? wrapped.arrayValue.values : [];
    var out = [];
    for (var i = 0; i < values.length; i++) out.push(fromFirestoreValue_(values[i]));
    return out;
  }
  return null;
}

function firestoreCommitWrites_(writes) {
  return firestoreRequest_('post', '/documents:commit', { writes: writes || [] });
}

function firestoreDocName_(docPath) {
  return firestoreDbResourceRoot_() + '/documents/' + docPath;
}

function firestoreSetDocument_(docPath, data) {
  var write = {
    update: {
      name: firestoreDocName_(docPath),
      fields: toFirestoreFields_(data || {})
    }
  };
  return firestoreCommitWrites_([write]);
}

function firestorePatchDocument_(docPath, partialData) {
  var keys = [];
  for (var k in partialData) {
    if (partialData.hasOwnProperty(k)) keys.push(k);
  }
  if (!keys.length) return { success: true, skipped: true };

  var query = { 'updateMask.fieldPaths': keys };
  var payload = { 
    name: firestoreDocName_(docPath),
    fields: toFirestoreFields_(partialData) 
  };
  return firestoreRequest_('patch', '/documents/' + docPath, payload, query);
}

function firestoreGetDocument_(docPath) {
  return firestoreRequest_('get', '/documents/' + docPath);
}

function firestoreDeleteDocument_(docPath) {
  return firestoreRequest_('delete', '/documents/' + docPath);
}

function firestoreListDocuments_(collectionPath, pageSize) {
  var q = {};
  if (pageSize) q.pageSize = pageSize;
  return firestoreRequest_('get', '/documents/' + collectionPath, null, q);
}

function normalizePhases_(phases) {
  var list = Object.prototype.toString.call(phases) === '[object Array]' ? phases : [];
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var p = list[i] || {};
    var id = p.id ? String(p.id) : ('phase_' + (i + 1));
    out.push({
      phaseId: id,
      title: String(p.title || ('Phase ' + (i + 1))),
      description: String(p.description || ''),
      start: Math.max(1, Number(p.start || 1)),
      duration: Math.max(1, Number(p.duration || 1)),
      jeh: Math.max(0, Number(p.jeh || 0)),
      priceJeh: Math.max(0, Number(p.priceJeh || 0)),
      done: !!p.done,
      order: i + 1
    });
  }
  return out;
}

function normalizeTodos_(todos) {
  var list = Object.prototype.toString.call(todos) === '[object Array]' ? todos : [];
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var t = list[i] || {};
    var id = t.id ? String(t.id) : ('todo_' + (i + 1));
    out.push({
      todoId: id,
      text: String(t.text || ''),
      done: !!t.done,
      date: String(t.date || ''),
      order: i + 1
    });
  }
  return out;
}

function extractJsonObject_(value, fallback) {
  if (value && typeof value === 'object' && Object.prototype.toString.call(value) !== '[object Array]') return value;
  if (typeof value === 'string' && value) {
    try {
      var parsed = JSON.parse(value);
      if (parsed && typeof parsed === 'object' && Object.prototype.toString.call(parsed) !== '[object Array]') return parsed;
    } catch (e) {}
  }
  return fallback || {};
}

function extractJsonArray_(value, fallback) {
  if (Object.prototype.toString.call(value) === '[object Array]') return value;
  if (typeof value === 'string' && value) {
    try {
      var parsed = JSON.parse(value);
      if (Object.prototype.toString.call(parsed) === '[object Array]') return parsed;
    } catch (e) {}
  }
  return fallback || [];
}

function round2_(n) {
  return Math.round(Number(n || 0) * 100) / 100;
}

function fsEtudesCollection_() {
  return FIRESTORE_CONFIG.etudesCollection;
}

function fsEtudeDocPath_(studyId) {
  return fsEtudesCollection_() + '/' + String(studyId);
}

function fsJuniorDocPath_() {
  return 'meta/junior';
}

function getDefaultJuniorProfile_() {
  return {
    nom: 'Junior INSA Services',
    signature: "Junior-Entreprise de l'INSA de Toulouse",
    statut_juridique: 'Association régie par la loi de 1901 - Membre de la Confédération Nationale des Junior-Entreprises',
    adresse: '135 Avenue de Rangueil 31077 TOULOUSE Cedex 04',
    'numero-telephone': '05 67 04 88 99',
    mail: 'contact@juniorinsaservices.fr',
    siret: '341 760 528 00016',
    ape: '9499Z',
    tva_intracommunautaire: 'FR35 341760528',
    siteweb: 'https://juniorinsaservices.fr',
    president: { prenom: 'Jean-Gabriel', nom: 'Virgone', mail: '', sexe: 'H' },
    vicepresident: { prenom: 'Victor', nom: 'Schalk', mail: '', sexe: 'H' },
    tresorier: { prenom: 'William', nom: 'Woodward', mail: '', sexe: 'H' },
    updatedAt: new Date().toISOString()
  };
}

function normalizeJuniorRolePerson_(value, fallback) {
  var fb = fallback && typeof fallback === 'object' ? fallback : { prenom: '', nom: '', mail: '', sexe: '' };

  if (value && typeof value === 'object' && Object.prototype.toString.call(value) !== '[object Array]') {
    return {
      prenom: String(value.prenom || fb.prenom || '').trim(),
      nom: String(value.nom || fb.nom || '').trim(),
      mail: String(value.mail || fb.mail || '').trim(),
      sexe: String(value.sexe || fb.sexe || '').trim()
    };
  }

  if (typeof value === 'string' && String(value).trim()) {
    var raw = String(value).trim();
    var parts = raw.split(/\s+/);
    if (parts.length === 1) {
      return { prenom: '', nom: parts[0], mail: String(fb.mail || '').trim(), sexe: String(fb.sexe || '').trim() };
    }
    return {
      prenom: parts.slice(0, parts.length - 1).join(' '),
      nom: parts[parts.length - 1],
      mail: String(fb.mail || '').trim(),
      sexe: String(fb.sexe || '').trim()
    };
  }

  return {
    prenom: String(fb.prenom || '').trim(),
    nom: String(fb.nom || '').trim(),
    mail: String(fb.mail || '').trim(),
    sexe: String(fb.sexe || '').trim()
  };
}

function normalizeJuniorProfile_(profile, defaults) {
  var base = profile && typeof profile === 'object' ? profile : {};
  var fb = defaults || getDefaultJuniorProfile_();

  base.president = normalizeJuniorRolePerson_(base.president, fb.president);
  base.vicepresident = normalizeJuniorRolePerson_(base.vicepresident, fb.vicepresident);
  base.tresorier = normalizeJuniorRolePerson_(base.tresorier, fb.tresorier);
  return base;
}

function fsJuniorGetProfile() {
  var defaults = getDefaultJuniorProfile_();
  try {
    var doc = firestoreGetDocument_(fsJuniorDocPath_());
    var data = fromFirestoreDocument_(doc) || {};
    var merged = {};
    for (var dk in defaults) {
      if (defaults.hasOwnProperty(dk)) merged[dk] = defaults[dk];
    }
    for (var k in data) {
      if (data.hasOwnProperty(k) && data[k] !== null && data[k] !== undefined && data[k] !== '') {
        merged[k] = data[k];
      }
    }
    return normalizeJuniorProfile_(merged, defaults);
  } catch (e) {
    var msg = String((e && e.message) || e || '');
    if (msg.indexOf(' -> 404 ') !== -1) {
      firestoreSetDocument_(fsJuniorDocPath_(), defaults);
      return defaults;
    }
    throw e;
  }
}

function fsJuniorUpsertProfile(partialData) {
  var existing = fsJuniorGetProfile();
  var patch = partialData && typeof partialData === 'object' ? partialData : {};
  var merged = {};

  for (var ek in existing) {
    if (existing.hasOwnProperty(ek)) merged[ek] = existing[ek];
  }
  for (var pk in patch) {
    if (patch.hasOwnProperty(pk)) merged[pk] = patch[pk];
  }
  merged = normalizeJuniorProfile_(merged, getDefaultJuniorProfile_());
  merged.updatedAt = new Date().toISOString();

  firestoreSetDocument_(fsJuniorDocPath_(), merged);
  return { success: true, data: merged };
}

function fsEtudeFullCacheKey_(studyId) {
  return 'FS_ETUDE_FULL_' + String(studyId);
}

function getCachedFsEtudeFull_(studyId) {
  try {
    var raw = CacheService.getScriptCache().get(fsEtudeFullCacheKey_(studyId));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function setCachedFsEtudeFull_(studyId, fullData) {
  try {
    CacheService.getScriptCache().put(fsEtudeFullCacheKey_(studyId), JSON.stringify(fullData || {}), 60);
  } catch (e) {}
}

function invalidateFsEtudeCaches_(studyId) {
  try {
    CacheService.getScriptCache().remove(fsEtudeFullCacheKey_(studyId));
  } catch (e) {}
}

function fsEtudesCategoriesCacheKey_() {
  return 'FS_ETUDES_CATEGORIES';
}

function invalidateFsEtudesCategoriesCache_() {
  try {
    CacheService.getScriptCache().remove(fsEtudesCategoriesCacheKey_());
  } catch (e) {}
}

function getCachedFsEtudesCategories_() {
  try {
    var raw = CacheService.getScriptCache().get(fsEtudesCategoriesCacheKey_());
    if (!raw) return null;
    var parsed = JSON.parse(raw);
    return Object.prototype.toString.call(parsed) === '[object Array]' ? parsed : null;
  } catch (e) {
    return null;
  }
}

function setCachedFsEtudesCategories_(categories) {
  try {
    CacheService.getScriptCache().put(fsEtudesCategoriesCacheKey_(), JSON.stringify(categories || []), 300);
  } catch (e) {}
}

function fsEtudeGet(studyId) {
  var doc = firestoreGetDocument_(fsEtudeDocPath_(studyId));
  return fromFirestoreDocument_(doc);
}

function fsEtudeGetFull(studyId) {
  var cached = getCachedFsEtudeFull_(studyId);
  if (cached) return cached;

  var base = fsEtudeGet(studyId);
  var phases = Object.prototype.toString.call(base.phasesSnapshot) === '[object Array]' ? base.phasesSnapshot : null;
  var todos = Object.prototype.toString.call(base.todosSnapshot) === '[object Array]' ? base.todosSnapshot : null;

  if (!phases) {
    var phasesResp = firestoreListDocuments_(fsEtudeDocPath_(studyId) + '/phases', 500);
    phases = [];
    var docsA = phasesResp.documents || [];
    for (var i = 0; i < docsA.length; i++) phases.push(fromFirestoreDocument_(docsA[i]));
  }

  if (!todos) {
    var todosResp = firestoreListDocuments_(fsEtudeDocPath_(studyId) + '/todos', 500);
    todos = [];
    var docsB = todosResp.documents || [];
    for (var j = 0; j < docsB.length; j++) todos.push(fromFirestoreDocument_(docsB[j]));
  }

  base.phases = phases;
  base.todos = todos;
  setCachedFsEtudeFull_(studyId, base);
  return base;
}

function fsEtudeUpsert(studyId, data) {
  firestoreSetDocument_(fsEtudeDocPath_(studyId), data || {});
  invalidateFsEtudeCaches_(studyId);
  invalidateFsEtudesCategoriesCache_();
  return { success: true, studyId: String(studyId) };
}

function fsEtudePatch(studyId, partialData) {
  firestorePatchDocument_(fsEtudeDocPath_(studyId), partialData || {});
  invalidateFsEtudeCaches_(studyId);
  invalidateFsEtudesCategoriesCache_();
  return { success: true, studyId: String(studyId) };
}

function fsEtudeDelete(studyId) {
  firestoreDeleteDocument_(fsEtudeDocPath_(studyId));
  invalidateFsEtudeCaches_(studyId);
  invalidateFsEtudesCategoriesCache_();
  return { success: true, studyId: String(studyId) };
}

function fsEtudeReplacePhases(studyId, phases) {
  replaceSubcollectionDocs_(fsEtudeDocPath_(studyId), 'phases', phases || [], 'phaseId');
  firestorePatchDocument_(fsEtudeDocPath_(studyId), {
    phasesSnapshot: phases || [],
    updatedAt: new Date().toISOString()
  });
  invalidateFsEtudeCaches_(studyId);
  return { success: true, count: (phases || []).length };
}

function fsEtudeReplaceTodos(studyId, todos) {
  replaceSubcollectionDocs_(fsEtudeDocPath_(studyId), 'todos', todos || [], 'todoId');
  firestorePatchDocument_(fsEtudeDocPath_(studyId), {
    todosSnapshot: todos || [],
    updatedAt: new Date().toISOString()
  });
  invalidateFsEtudeCaches_(studyId);
  return { success: true, count: (todos || []).length };
}

function replaceSubcollectionDocs_(parentDocPath, subcollectionName, docs, idField) {
  var collectionPath = parentDocPath + '/' + subcollectionName;
  var existing = [];
  try {
    var resp = firestoreListDocuments_(collectionPath, 500);
    existing = resp.documents || [];
  } catch (e) {
    existing = [];
  }

  var writes = [];
  var i;
  for (i = 0; i < existing.length; i++) {
    writes.push({ delete: existing[i].name });
    if (writes.length >= 450) {
      firestoreCommitWrites_(writes);
      writes = [];
    }
  }

  for (i = 0; i < docs.length; i++) {
    var row = docs[i] || {};
    var docId = row[idField] ? String(row[idField]) : (subcollectionName + '_' + (i + 1));
    writes.push({
      update: {
        name: firestoreDocName_(collectionPath + '/' + docId),
        fields: toFirestoreFields_(row)
      }
    });
    if (writes.length >= 450) {
      firestoreCommitWrites_(writes);
      writes = [];
    }
  }

  if (writes.length) firestoreCommitWrites_(writes);
}

function enregistrerDonneeEtudeFirestore(studyId, key, value) {
  var k = String(key || '');
  if (!k) throw new Error('Clé vide');

  var patch = { updatedAt: new Date().toISOString() };

  if (k === 'Description') {
    patch.description = String(value || '');
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'description' };
  }
  if (k === 'Status') {
    patch.status = String(value || '');
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'status' };
  }
  if (k === 'Planning_StartDate') {
    patch.planningStartDate = String(value || '');
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'planningStartDate' };
  }
  if (k === 'ClientInfo') {
    patch.clientInfo = extractJsonObject_(value, {});
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'clientInfo' };
  }
  if (k === 'AvantProjet_Steps') {
    patch.avantProjetSteps = extractJsonObject_(value, { client: {}, interv: {} });
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'avantProjetSteps' };
  }
  if (k === 'Negociation_Steps') {
    patch.negociationSteps = extractJsonObject_(value, { client: {}, interv: {} });
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'negociationSteps' };
  }
  if (k === 'EtudeInfo') {
    var etudeInfo = extractJsonObject_(value, {});
    patch.etudeInfo = etudeInfo;
    if (etudeInfo.hasOwnProperty('progress')) patch.progress = Number(etudeInfo.progress || 0);
    if (etudeInfo.nom) patch.name = String(etudeInfo.nom);

    if (etudeInfo.reference && typeof etudeInfo.reference === 'string') {
      var refParts = etudeInfo.reference.split('_');
      if (refParts.length >= 2) {
        patch.ref = {
          year: String(refParts[0] || ''),
          code: String(refParts.slice(1).join('_') || '')
        };
      }
    }

    var prix = Number(etudeInfo.prix || 0);
    var frais = Number(etudeInfo.fraisEtude || 0);
    var totalJeh = Number(etudeInfo.totalJeh || 0);
    patch.financials = {
      prixHT: round2_(Math.max(0, prix)),
      fraisEtude: round2_(Math.max(0, frais)),
      totalJeh: round2_(Math.max(0, totalJeh)),
      prixTotalHT: round2_(Math.max(0, prix) + Math.max(0, frais)),
      tva: round2_((Math.max(0, prix) + Math.max(0, frais)) * 0.2),
      prixTotalTTC: round2_((Math.max(0, prix) + Math.max(0, frais)) * 1.2)
    };

    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'etudeInfo+financials+progress' };
  }
  if (k === 'Planning_Phases') {
    var phases = normalizePhases_(extractJsonArray_(value, []));
    fsEtudeReplacePhases(studyId, phases);
    patch.updatedAt = new Date().toISOString();
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'phases' };
  }
  if (k === 'Todo_Chargé') {
    var todos = normalizeTodos_(extractJsonArray_(value, []));
    fsEtudeReplaceTodos(studyId, todos);
    patch.updatedAt = new Date().toISOString();
    fsEtudePatch(studyId, patch);
    return { success: true, mappedTo: 'todos' };
  }

  patch.extraData = patch.extraData || {};
  patch.extraData[sanitizeKey_(k)] = String(value || '');
  fsEtudePatch(studyId, patch);
  return { success: true, mappedTo: 'extraData.' + sanitizeKey_(k) };
}


function enregistrerDonneeEtudeFirestoreBatch(studyId, updates) {
  var patch = { updatedAt: new Date().toISOString() };
  var mappedTo = [];
  var hasPatch = false;

  for (var i = 0; i < updates.length; i++) {
    var k = String(updates[i].key || '');
    if (!k) continue;
    var value = updates[i].value;

    if (k === 'Description') {
      patch.description = String(value || '');
      mappedTo.push('description');
      hasPatch = true;
    }
    else if (k === 'Status') {
      patch.status = String(value || '');
      mappedTo.push('status');
      hasPatch = true;
    }
    else if (k === 'Planning_StartDate') {
      patch.planningStartDate = String(value || '');
      mappedTo.push('planningStartDate');
      hasPatch = true;
    }
    else if (k === 'ClientInfo') {
      patch.clientInfo = extractJsonObject_(value, {});
      mappedTo.push('clientInfo');
      hasPatch = true;
    }
    else if (k === 'AvantProjet_Steps') {
      patch.avantProjetSteps = extractJsonObject_(value, { client: {}, interv: {} });
      mappedTo.push('avantProjetSteps');
      hasPatch = true;
    }
    else if (k === 'Negociation_Steps') {
      patch.negociationSteps = extractJsonObject_(value, { client: {}, interv: {} });
      mappedTo.push('negociationSteps');
      hasPatch = true;
    }
    else if (k === 'EtudeInfo') {
      var etudeInfo = extractJsonObject_(value, {});
      patch.etudeInfo = etudeInfo;
      if (etudeInfo.hasOwnProperty('progress')) patch.progress = Number(etudeInfo.progress || 0);
      if (etudeInfo.nom) patch.name = String(etudeInfo.nom);

      if (etudeInfo.reference && typeof etudeInfo.reference === 'string') {
        var refParts = etudeInfo.reference.split('_');
        if (refParts.length >= 2) {
          patch.ref = {
            year: String(refParts[0] || ''),
            code: String(refParts.slice(1).join('_') || '')
          };
        }
      }

      var prix = Number(etudeInfo.prix || 0);
      var frais = Number(etudeInfo.fraisEtude || 0);
      var totalJeh = Number(etudeInfo.totalJeh || 0);
      patch.financials = {
        prixHT: round2_(Math.max(0, prix)),
        fraisEtude: round2_(Math.max(0, frais)),
        totalJeh: round2_(Math.max(0, totalJeh)),
        prixTotalHT: round2_(Math.max(0, prix) + Math.max(0, frais)),
        tva: round2_((Math.max(0, prix) + Math.max(0, frais)) * 0.2),
        prixTotalTTC: round2_((Math.max(0, prix) + Math.max(0, frais)) * 1.2)
      };

      mappedTo.push('etudeInfo+financials+progress');
      hasPatch = true;
    }
    else if (k === 'Planning_Phases') {
      var phases = normalizePhases_(extractJsonArray_(value, []));
      fsEtudeReplacePhases(studyId, phases);
      mappedTo.push('phases');
      hasPatch = true;
    }
    else if (k === 'Todo_Chargé') {
      var todos = normalizeTodos_(extractJsonArray_(value, []));
      fsEtudeReplaceTodos(studyId, todos);
      mappedTo.push('todos');
      hasPatch = true;
    }
    else {
      patch.extraData = patch.extraData || {};
      patch.extraData[sanitizeKey_(k)] = String(value || '');
      mappedTo.push('extraData.' + sanitizeKey_(k));
      hasPatch = true;
    }
  }

  if (hasPatch) {
    fsEtudePatch(studyId, patch);
  }
  
  return { success: true, mappedTo: mappedTo.join(',') };
}

function sanitizeKey_(key) {
  return String(key || '').toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '') || 'key';
}

function testConnexionFirestore() {
  var cfg = getFirestoreSettings_();
  var token = getFirestoreAccessToken_();
  if (!token) throw new Error('Token vide');

  var resp = firestoreRequest_('get', '/documents/' + fsEtudesCollection_(), null, { pageSize: 1 });
  return {
    success: true,
    projectId: cfg.projectId,
    databaseId: cfg.databaseId,
    collection: fsEtudesCollection_(),
    sampleCount: (resp.documents || []).length
  };
}

function listerCategoriesEtudesFirestore() {
  var cached = getCachedFsEtudesCategories_();
  if (cached) return cached;

  var resp = firestoreListDocuments_(fsEtudesCollection_(), 500);
  var docs = resp && resp.documents ? resp.documents : [];
  var categoriesSet = {};
  var categories = [];

  for (var i = 0; i < docs.length; i++) {
    var etude = fromFirestoreDocument_(docs[i]) || {};
    var etudeInfo = etude.etudeInfo || {};
    var raw = etudeInfo.categories;

    if (!raw) raw = etudeInfo.categorie;
    var list = [];
    if (Object.prototype.toString.call(raw) === '[object Array]') {
      list = raw;
    } else if (typeof raw === 'string' && raw) {
      list = raw.split(',');
    }

    for (var j = 0; j < list.length; j++) {
      var cat = String(list[j] || '').trim();
      if (!cat) continue;
      var key = cat.toLowerCase();
      if (categoriesSet[key]) continue;
      categoriesSet[key] = true;
      categories.push(cat);
    }
  }

  categories.sort(function(a, b) { return a.localeCompare(b, 'fr'); });
  setCachedFsEtudesCategories_(categories);
  return categories;
}

/**
 * Générateur de documents généraliste (Google Apps Script).
 *
 * Entrées minimales:
 * - documentName: nom du document généré
 * - destinationFolderId: ID du dossier de sauvegarde
 * - templateName: nom de la template dans l'annuaire
 *
 * Entrées optionnelles:
 * - studyId: ID de l'étude pour alimenter automatiquement les variables
 * - extraVariables: objet (ou JSON string) fusionné dans les variables disponibles
 */
function generateGenericDocument(documentName, destinationFolderId, templateName, studyId, extraVariables) {
  if (!documentName || !String(documentName).trim()) {
    throw new Error('Le nom du document est obligatoire.');
  }
  if (!destinationFolderId || !String(destinationFolderId).trim()) {
    throw new Error("L'ID du dossier de destination est obligatoire.");
  }
  if (!templateName || !String(templateName).trim()) {
    throw new Error('Le nom de la template est obligatoire.');
  }

  var requestedName = String(documentName).trim();
  var templateId = getTemplateFileIdByName(String(templateName).trim());
  var templateFile = DriveApp.getFileById(templateId);
  var destinationFolder = DriveApp.getFolderById(String(destinationFolderId).trim());
  var newFile = templateFile.makeCopy(requestedName, destinationFolder);
  newFile = maybeConvertOfficeTemplateCopy_(newFile, destinationFolder.getId(), requestedName);

  var context = buildAutofillVariablesMap_(studyId, extraVariables);
  var variables = context.variables;
  var phases = context.phases;
  var mimeType = newFile.getMimeType();

  if (mimeType === MimeType.GOOGLE_SLIDES) {
    var presentation = SlidesApp.openById(newFile.getId());
    applyPhaseTokensInPresentation_(presentation, variables, phases);
    replaceTokensInPresentation_(presentation, variables);
    presentation.saveAndClose();
  } else if (mimeType === MimeType.GOOGLE_DOCS) {
    var doc = DocumentApp.openById(newFile.getId());
    if (!doc || typeof doc.getBody !== 'function') {
      throw new Error('Le document genere ne peut pas etre lu comme Google Docs. Verifiez la template source "' + String(templateName) + '".');
    }
    replaceTokensInDocument_(doc, variables, phases);
    doc.saveAndClose();
  } else {
    throw new Error('Type de template non supporté: ' + mimeType + '. Utilisez Google Docs ou Google Slides.');
  }

  return {
    success: true,
    id: newFile.getId(),
    name: newFile.getName(),
    url: newFile.getUrl(),
    folderUrl: destinationFolder.getUrl(),
    mimeType: mimeType,
    templateName: templateName
  };
}

function maybeConvertOfficeTemplateCopy_(fileObj, destinationFolderId, requestedName) {
  var mimeType = String((fileObj && fileObj.getMimeType && fileObj.getMimeType()) || '').toLowerCase();
  var isWord = mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || mimeType === 'application/msword';
  var isPowerPoint = mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || mimeType === 'application/vnd.ms-powerpoint';

  if (!isWord && !isPowerPoint) {
    return fileObj;
  }

  if (!canUseDriveAdvancedService_()) {
    var officeType = isWord ? 'Word (.docx/.doc)' : 'PowerPoint (.pptx/.ppt)';
    throw new Error(
      'Le template est un fichier ' + officeType + '. ' +
      'Conversion automatique indisponible: activez le service avance Drive ou convertissez ce template en Google Docs/Slides.'
    );
  }

  var resource = {
    title: String(requestedName || fileObj.getName() || 'Document').trim(),
    parents: [{ id: String(destinationFolderId || '').trim() }]
  };

  var copied;
  try {
    copied = Drive.Files.copy(resource, fileObj.getId(), {
      convert: true,
      supportsAllDrives: true,
      supportsTeamDrives: true
    });
  } catch (err) {
    throw new Error('Conversion automatique du template Office impossible: ' + (err && err.message ? err.message : err));
  }

  var newId = copied && copied.id ? String(copied.id) : '';
  if (!newId) {
    throw new Error('Conversion automatique du template Office echouee: ID converti introuvable.');
  }

  try {
    fileObj.setTrashed(true);
  } catch (trashErr) {}

  return DriveApp.getFileById(newId);
}

function canUseDriveAdvancedService_() {
  return typeof Drive !== 'undefined' && Drive && Drive.Files && typeof Drive.Files.copy === 'function';
}

// Alias FR explicite
function genererDocumentGeneraliste(documentName, destinationFolderId, templateName, studyId, extraVariables) {
  return generateGenericDocument(documentName, destinationFolderId, templateName, studyId, extraVariables);
}

var DOC_WORKFLOW_CHAT_URL_ = 'https://chat.google.com/room/AAQApAX9EAQ';
var DOC_WORKFLOW_WEBHOOK_URL_ = 'https://chat.googleapis.com/v1/spaces/AAQApAX9EAQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=gvhEGBSUOfSq6NN-tleGBs4-NQNG1GGFuah5SYG5W18';

function getDocWorkflowConfig_(docKey) {
  var key = String(docKey || '').trim().toLowerCase();
  var configs = {
    devis: {
      key: 'devis',
      suffix: 'DEVIS',
      label: 'Devis',
      templateCandidates: ['devis'],
      reviewEnabled: true
    },
    ce: {
      key: 'ce',
      suffix: 'CE',
      label: 'Convention etude',
      templateCandidates: ['ce', 'convention-etude', 'convention etude', 'convention_etude'],
      reviewEnabled: true
    },
    cc: {
      key: 'cc',
      suffix: 'CC',
      label: 'Convention client',
      templateCandidates: ['cc', 'convention-client', 'convention client', 'convention_client'],
      reviewEnabled: true
    },
    bdc: {
      key: 'bdc',
      suffix: 'BDC',
      label: 'Bon de commande',
      templateCandidates: ['bdc', 'bon-de-commande', 'bon de commande', 'bon_de_commande'],
      reviewEnabled: true
    },
    rm: {
      key: 'rm',
      suffix: 'RM',
      label: 'RM',
      templateCandidates: ['rm', 'releve-mission', 'releve mission', 'releve_mission'],
      reviewEnabled: true
    },
    arm: {
      key: 'arm',
      suffix: 'ARM',
      label: 'ARM',
      templateCandidates: ['arm', 'accord-realisation-mission', 'accord realisation mission', 'accord_realisation_mission'],
      reviewEnabled: true
    },
    arrm: {
      key: 'arrm',
      suffix: 'ARRM',
      label: 'ARRM',
      templateCandidates: ['arrm', 'avenant-realisation-rm', 'avenant realisation rm', 'avenant_realisation_rm'],
      reviewEnabled: true
    },
    arce: {
      key: 'arce',
      suffix: 'ARCE',
      label: 'ARCE',
      templateCandidates: ['arce', 'accord-realisation-ce', 'accord realisation ce', 'accord_realisation_ce'],
      reviewEnabled: true
    },
    ace: {
      key: 'ace',
      suffix: 'ACE',
      label: 'ACE',
      templateCandidates: ['ace', 'avenant-ce', 'avenant ce', 'avenant_ce'],
      reviewEnabled: true
    },
    acc: {
      key: 'acc',
      suffix: 'ACC',
      label: 'ACC',
      templateCandidates: ['acc', 'avenant-cc', 'avenant cc', 'avenant_cc'],
      reviewEnabled: true
    },
    bdcr: {
      key: 'bdcr',
      suffix: 'BDCR',
      label: 'BDCR',
      templateCandidates: ['bdcr', 'bon-de-commande-realisation', 'bon de commande realisation', 'bon_de_commande_realisation'],
      reviewEnabled: true
    }
  };

  if (!configs[key]) {
    throw new Error('Type de document non supporte: ' + docKey);
  }
  return configs[key];
}

function cloneFlatObject_(source) {
  var out = {};
  var s = source && typeof source === 'object' ? source : {};
  for (var key in s) {
    if (s.hasOwnProperty(key)) out[key] = s[key];
  }
  return out;
}

function resolveFirstAvailableTemplateName_(candidates) {
  var names = Object.prototype.toString.call(candidates) === '[object Array]' ? candidates : [candidates];
  var tested = [];

  for (var i = 0; i < names.length; i++) {
    var candidate = String(names[i] || '').trim();
    if (!candidate) continue;
    tested.push(candidate);
    try {
      getTemplateFileIdByName(candidate);
      return candidate;
    } catch (e) {}
  }

  throw new Error('Aucun template trouve parmi: ' + tested.join(', '));
}

function getStudyReferenceForDocs_(rawData) {
  var source = rawData && typeof rawData === 'object' ? rawData : {};
  var etudeInfo = source.EtudeInfo && typeof source.EtudeInfo === 'object' ? source.EtudeInfo : {};
  var explicitRef = String(etudeInfo.reference || source.reference || '').trim();
  if (explicitRef) return explicitRef;

  var year = String(source.year || 'XXXX').trim() || 'XXXX';
  var code = String(source.code || 'XXX').trim() || 'XXX';
  return year + '_' + code;
}

function normalizeFolderNameForLookup_(name) {
  return String(name || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findEditableDocumentsFolder_(studyFolder) {
  return getChildFolderByPredicate(studyFolder, function(name) {
    var normalized = normalizeFolderNameForLookup_(name);
    if (!/^2\s/.test(normalized)) return false;
    if (normalized.indexOf('document') < 0) return false;
    return normalized.indexOf('editable') >= 0;
  });
}

function findSignedDocumentsFolder_(studyFolder) {
  return getChildFolderByPredicate(studyFolder, function(name) {
    var normalized = normalizeFolderNameForLookup_(name);
    if (!/^1\s/.test(normalized)) return false;
    if (normalized.indexOf('document') < 0) return false;
    return normalized.indexOf('sign') >= 0;
  });
}

function getStudyDocumentsFolders_(studyId) {
  var studyFolder = DriveApp.getFolderById(studyId);
  var legacyDocsFolder = getChildFolderByPredicate(studyFolder, function(name) {
    return /^2\.\s*Documents$/i.test(String(name || '').trim());
  });
  if (!legacyDocsFolder) {
    var docsByName = studyFolder.getFoldersByName('2. Documents');
    legacyDocsFolder = docsByName.hasNext() ? docsByName.next() : null;
  }

  var legacyReviewFolder = null;
  var legacyValidatedFolder = null;
  if (legacyDocsFolder) {
    legacyReviewFolder = getChildFolderByPredicate(legacyDocsFolder, function(name) { return /^1\.\s*En\s*relecture$/i.test(name); });
    if (!legacyReviewFolder) {
      var reviewByName = legacyDocsFolder.getFoldersByName('1. En relecture');
      legacyReviewFolder = reviewByName.hasNext() ? reviewByName.next() : null;
    }

    legacyValidatedFolder = getChildFolderByPredicate(legacyDocsFolder, function(name) { return /^2\.\s*Valid/i.test(name); });
    if (!legacyValidatedFolder) {
      var validByName = legacyDocsFolder.getFoldersByName('2. Valides');
      if (!validByName.hasNext()) validByName = legacyDocsFolder.getFoldersByName('2. Validés');
      legacyValidatedFolder = validByName.hasNext() ? validByName.next() : null;
    }
  }

  var editableFolder = findEditableDocumentsFolder_(studyFolder) || legacyReviewFolder;
  if (!editableFolder) {
    throw new Error('Dossier "2. [REF] Documents éditables" (ou legacy "2. Documents/1. En relecture") introuvable.');
  }

  var signedFolder = findSignedDocumentsFolder_(studyFolder) || legacyValidatedFolder || null;

  return {
    studyFolder: studyFolder,
    docsFolder: legacyDocsFolder || editableFolder,
    reviewFolder: editableFolder,
    validatedFolder: signedFolder,
    signedFolder: signedFolder,
    legacyDocsFolder: legacyDocsFolder
  };
}

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
    console.error("Erreur mise à jour statuts doc negociation : " + e.message);
  }
}

function checkWorkflowDocumentStatus_(studyId, docKey) {
  var cached = getCachedDocWorkflowStatus_(studyId, docKey);
  if (cached) return cached;

  var etudeData = {};
  try {
    etudeData = fsEtudeGet(studyId) || {};
  } catch(e) {
    console.error("Erreur checkWorkflowDocumentStatus_ (fsEtudeGet):", e);
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

function findWorkflowDocumentFile_(folder, expectedBaseName) {
  var file = null;
  var highest = 0;
  var safeBase = String(expectedBaseName).replace(/[.*+?^\$\{\}()|[\]\\]/g, '\\$&');
  var regex = new RegExp('^' + safeBase + '(?:_v(\\d+))?(?:\\.[a-z0-9]+)?$', 'i');

  if (!folder) return null;
  var all = folder.getFiles();
  while (all.hasNext()) {
    var current = all.next();
    var name = String(current.getName() || '');
    var match = name.match(regex);
    if (match) {
      var v = match[1] ? parseInt(match[1], 10) : 1;
      if (typeof v === 'number' && !isNaN(v) && v >= highest) {
        highest = v;
        file = current;
      }
    }
  }

  return file;
}

function getNextWorkflowDocumentName_(folders, expectedBaseName, docKey, isNewBdc) {
  var safeBase = String(expectedBaseName).replace(/[.*+?^\$\{\}()|[\]\\]/g, '\\$&');
  var highestNum = 0;
  var highestV = 0;
  var foundAny = false;
  var regex;
  
  if (docKey === 'bdc') {
    regex = new RegExp('^' + safeBase + '(\\d+)?(?:_v(\\d+))?(?:\\.[a-z0-9]+)?$', 'i');
  } else {
    regex = new RegExp('^' + safeBase + '(?:_v(\\d+))?(?:\\.[a-z0-9]+)?$', 'i');
  }

  function scanFolder_(folder) {
    if (!folder) return;
    var all = folder.getFiles();
    while (all.hasNext()) {
      var current = all.next();
      var mime = String(current.getMimeType() || '').toLowerCase();
      if (mime === 'application/pdf') continue;
      var match = String(current.getName() || '').match(regex);
      if (match) {
        foundAny = true;
        if (docKey === 'bdc') {
          var n = match[1] ? parseInt(match[1], 10) : 1;
          var v = match[2] ? parseInt(match[2], 10) : 0;
          if (n > highestNum) {
            highestNum = n;
            highestV = v;
          } else if (n === highestNum && v > highestV) {
            highestV = v;
          }
        } else {
          var v = match[1] ? parseInt(match[1], 10) : 0;
          if (v > highestV) highestV = v;
        }
      }
    }
  }

  scanFolder_(folders.reviewFolder);
  scanFolder_(folders.docsFolder);

  if (!foundAny) {
    if (docKey === 'bdc') return expectedBaseName + '01';
    return expectedBaseName;
  }

  if (docKey === 'bdc') {
    if (isNewBdc) {
      highestNum++;
      var numStr = highestNum < 10 ? '0' + highestNum : String(highestNum);
      return expectedBaseName + numStr;
    } else {
      highestV++;
      var numStr = highestNum < 10 ? '0' + highestNum : String(highestNum);
      return expectedBaseName + numStr + (highestV > 0 ? '_v' + highestV : '');
    }
  } else {
    return expectedBaseName + '_v' + (highestV + 1);
  }
}

function generateWorkflowDocument_(studyId, docKey, forceRegenerate, isNew) {
  var cfg = getDocWorkflowConfig_(docKey);
  var rawData = getStudyData(studyId) || {};
  var reference = getStudyReferenceForDocs_(rawData);
  var baseName = reference + '_' + cfg.suffix;
  var folders = getStudyDocumentsFolders_(studyId);

  var targetName = getNextWorkflowDocumentName_(folders, baseName, cfg.key, isNew);

  if (forceRegenerate) {
    try {
      PropertiesService.getScriptProperties().deleteProperty('relecture_' + cfg.key + '_' + studyId);
      if (cfg.key === 'devis') {
        PropertiesService.getScriptProperties().deleteProperty('relecture_' + studyId);
      }
    } catch (e) {}
  }
  var templateName = resolveFirstAvailableTemplateName_(cfg.templateCandidates || []);
  
  var extraVars = {};
  if (cfg.key === 'bdc') {
    var bdcNumMatch = targetName.match(/BDC(\d+)/i);
    if (bdcNumMatch && bdcNumMatch[1]) {
      extraVars['bdc.num_bdc'] = Number(bdcNumMatch[1]).toString(); // '1', '2' etc.
    }
    
    var ccBase = reference + '_CC';
    var latestCcFile = findWorkflowDocumentFile_(folders.reviewFolder, ccBase);
    if (!latestCcFile) {
      latestCcFile = findWorkflowDocumentFile_(folders.docsFolder, ccBase);
    }
    
    extraVars['bdc.ref_cc'] = latestCcFile ? String(latestCcFile.getName()).replace(/\.[a-z0-9]+$/i, '') : ccBase;
  }

  var generated = generateGenericDocument(targetName, folders.reviewFolder.getId(), templateName, studyId, extraVars);

  patchDocWorkflowStatusFirestore_(studyId, cfg.key, {
    key: cfg.key,
    label: cfg.label,
    suffix: cfg.suffix,
    templateName: templateName,
    hasDraft: true,
    draftUrl: generated.url || '',
    draftName: generated.name || targetName,
    hasPdf: false,
    pdfUrl: '',
    hasValide: false,
    valideUrl: '',
    archiveRequired: false,
    signedAt: '',
    docName: generated.name || targetName,
    relectureDemandee: false,
    chatUrl: DOC_WORKFLOW_CHAT_URL_
  }, rawData);

  invalidateDocWorkflowStatusCache_(studyId, cfg.key);

  return {
    success: true,
    id: generated.id,
    name: generated.name || targetName,
    url: generated.url,
    folderUrl: folders.reviewFolder.getUrl(),
    templateName: templateName,
    key: cfg.key,
    label: cfg.label
  };
}

function sendWorkflowDocumentRelectureChat_(studyId, docKey, options) {
  var cfg = getDocWorkflowConfig_(docKey);
  if (!cfg.reviewEnabled) {
    throw new Error('La relecture nest pas activee pour ' + cfg.label + '.');
  }

  var rawData = getStudyData(studyId) || {};
  var reference = getStudyReferenceForDocs_(rawData);
  var baseName = reference + '_' + cfg.suffix;
  var folders = getStudyDocumentsFolders_(studyId);

  var sourceFile = findWorkflowDocumentFile_(folders.reviewFolder, baseName);
  if (!sourceFile && folders.docsFolder) {
    sourceFile = findWorkflowDocumentFile_(folders.docsFolder, baseName);
  }
  var pdfName = String(sourceFile ? sourceFile.getName() : baseName).replace(/\.[^/.]+$/, '') + '_relecture.pdf';

  if (!sourceFile) {
    throw new Error('Aucun document "' + cfg.label + '" trouvé dans le dossier "Documents éditables" de l\'étude.');
  }

  var pdfBlob = sourceFile.getAs('application/pdf');
  pdfBlob.setName(pdfName);

  var oldPdf = folders.reviewFolder.getFilesByName(pdfName);
  while (oldPdf.hasNext()) { oldPdf.next().setTrashed(true); }

  var pdfFile = folders.reviewFolder.createFile(pdfBlob);
  pdfFile.setName(pdfName);

  var opts = options && typeof options === 'object' ? options : {};
  var withChecklist = opts.showChecklist !== false;
  var checklistText = withChecklist
    ? '\n\n*Retour attendu:*' +
      '\n☐ Valide par le pole qualite' +
      '\n☐ Erreurs a corriger (iteration Modif ↔ Relecture)'
    : '';

  var studyLabel = reference + (rawData.nom ? ' - ' + rawData.nom : '');
  var editableFolderLabel = folders.reviewFolder ? folders.reviewFolder.getName() : 'Documents éditables';
  var payload = {
    text: '📄 *' + cfg.label + ' en relecture (PDF)* - ' + studyLabel +
      '\n🔗 <' + pdfFile.getUrl() + '|' + pdfName + '>' +
      '\n_Dossier : ' + editableFolderLabel + '_' +
      checklistText
  };

  var response = UrlFetchApp.fetch(DOC_WORKFLOW_WEBHOOK_URL_, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Erreur envoi webhook Chat (' + code + ') : ' + response.getContentText());
  }

  try {
    PropertiesService.getScriptProperties().setProperty(
      'relecture_' + cfg.key + '_' + studyId,
      JSON.stringify({ chatUrl: DOC_WORKFLOW_CHAT_URL_, sentAt: new Date().toISOString() })
    );
    if (cfg.key === 'devis') {
      PropertiesService.getScriptProperties().setProperty(
        'relecture_' + studyId,
        JSON.stringify({ chatUrl: DOC_WORKFLOW_CHAT_URL_, sentAt: new Date().toISOString() })
      );
    }
  } catch (e) {}

  patchDocWorkflowStatusFirestore_(studyId, cfg.key, {
    hasPdf: true,
    pdfUrl: pdfFile.getUrl(),
    hasValide: false,
    valideUrl: '',
    archiveRequired: false,
    signedAt: '',
    relectureDemandee: true,
    chatUrl: DOC_WORKFLOW_CHAT_URL_
  }, rawData);

  invalidateDocWorkflowStatusCache_(studyId, cfg.key);

  return {
    success: true,
    key: cfg.key,
    chatUrl: DOC_WORKFLOW_CHAT_URL_,
    pdfUrl: pdfFile.getUrl()
  };
}

function validateWorkflowDocument_(studyId, docKey) {
  return setWorkflowDocumentSignedFlag_(studyId, docKey, true);
}

function setWorkflowDocumentSignedFlag_(studyId, docKey, isSigned) {
  var cfg = getDocWorkflowConfig_(docKey);
  var rawData = getStudyData(studyId) || {};
  var signed = !!isSigned;
  var reference = getStudyReferenceForDocs_(rawData);
  var baseName = reference + '_' + cfg.suffix;
  var folders = getStudyDocumentsFolders_(studyId);

  var sourceFile = findWorkflowDocumentFile_(folders.reviewFolder, baseName);
  if (!sourceFile && folders.docsFolder) {
    sourceFile = findWorkflowDocumentFile_(folders.docsFolder, baseName);
  }
  if (signed && !sourceFile) {
    throw new Error('Aucun fichier "' + cfg.label + '" trouvé dans le dossier "Documents éditables".');
  }

  var patch = {
    hasValide: signed,
    archiveRequired: signed,
    signedAt: signed ? new Date().toISOString() : '',
    relectureDemandee: signed ? false : undefined,
    valideUrl: signed && sourceFile ? sourceFile.getUrl() : ''
  };

  if (!signed) {
    delete patch.relectureDemandee;
  }

  patchDocWorkflowStatusFirestore_(studyId, cfg.key, patch, rawData);

  invalidateDocWorkflowStatusCache_(studyId, cfg.key);

  return {
    success: true,
    key: cfg.key,
    hasValide: signed,
    archiveRequired: signed,
    signedAt: patch.signedAt,
    valideUrl: signed && sourceFile ? sourceFile.getUrl() : '',
    signedFolderName: folders.signedFolder ? folders.signedFolder.getName() : '',
    signedFolderUrl: folders.signedFolder ? folders.signedFolder.getUrl() : ''
  };
}

function setWorkflowDocumentSignedFlag(studyId, docKey, isSigned) {
  return setWorkflowDocumentSignedFlag_(studyId, docKey, !!isSigned);
}

function getNegociationDocumentsStatuses(studyId) {
  return {
    success: true,
    cc: checkWorkflowDocumentStatus_(studyId, 'cc'),
    bdc: checkWorkflowDocumentStatus_(studyId, 'bdc'),
    ce: checkWorkflowDocumentStatus_(studyId, 'ce'),
    rm: checkWorkflowDocumentStatus_(studyId, 'rm')
  };
}

function getWorkflowDocumentsStatuses(studyId, docKeys) {
  var list = Object.prototype.toString.call(docKeys) === '[object Array]'
    ? docKeys
    : [docKeys];
  var out = { success: true };

  for (var i = 0; i < list.length; i++) {
    var key = String(list[i] || '').trim().toLowerCase();
    if (!key) continue;
    out[key] = checkWorkflowDocumentStatus_(studyId, key);
  }

  return out;
}

function obtenirStatutsNegociationEtude(studyId) {
  return getNegociationDocumentsStatuses(studyId);
}

function obtenirStatutsDocumentsWorkflowEtude(studyId, docKeys) {
  return getWorkflowDocumentsStatuses(studyId, docKeys);
}

function genererDocumentNegociationEtude(studyId, docKey, forceRegenerate, isNewBdc) {
  return generateWorkflowDocument_(studyId, docKey, !!forceRegenerate, !!isNewBdc);
}

function envoyerDocumentNegociationRelectureChat(studyId, docKey, options) {
  return sendWorkflowDocumentRelectureChat_(studyId, docKey, options);
}

function validerDocumentNegociation(studyId, docKey) {
  return validateWorkflowDocument_(studyId, docKey);
}

function marquerDocumentWorkflowSigne(studyId, docKey, isSigned) {
  return setWorkflowDocumentSignedFlag_(studyId, docKey, !!isSigned);
}

function obtenirStatutDocumentWorkflowEtude(studyId, docKey) {
  return checkWorkflowDocumentStatus_(studyId, docKey);
}

function genererConventionClientEtude(studyId, forceRegenerate) {
  return generateWorkflowDocument_(studyId, 'cc', !!forceRegenerate);
}

function genererBonDeCommandeEtude(studyId, forceRegenerate) {
  return generateWorkflowDocument_(studyId, 'bdc', !!forceRegenerate);
}

function genererConventionEtudeEtude(studyId, forceRegenerate) {
  return generateWorkflowDocument_(studyId, 'ce', !!forceRegenerate);
}

function prepareMailFromTemplate(studyId, templateName, defaultSubject, fromAlias, extraVariables) {
  if (!templateName || !String(templateName).trim()) {
    throw new Error('Le nom de la template mail est obligatoire.');
  }

  var templateId = getTemplateFileIdByName(String(templateName).trim());
  var templateFile = DriveApp.getFileById(templateId);
  var templateText = getMailTemplateText_(templateFile);

  var context = buildAutofillVariablesMap_(studyId, extraVariables);
  var rendered = replaceTokensInText_(templateText, context.variables || {});
  var parsed = extractMailSubjectAndBody_(rendered, defaultSubject || '');

  return {
    success: true,
    templateName: templateName,
    subject: parsed.subject || defaultSubject || '',
    body: parsed.body || '',
    from: fromAlias || ''
  };
}

function prepareIntervenantSearchMail(studyId) {
  if (!studyId) throw new Error('ID étude manquant.');

  var rawData = {};
  try {
    rawData = getStudyData(studyId) || {};
  } catch (e) {
    rawData = {};
  }

  var studyRef = rawData.reference || [rawData.year || '', rawData.code || ''].filter(function(v) { return !!String(v).trim(); }).join('_');
  var defaultSubject = 'Recherche intervenant' + (studyRef ? ' — Étude ' + studyRef : '');

  var context = buildAutofillVariablesMap_(studyId, null);
  var templateHtml = getProjectHtmlFileContent_('mail-recherche-intervenant');
  var renderedHtml = replaceTokensInText_(templateHtml, context.variables || {});
  var parsed = extractMailSubjectAndBodyFromHtml_(renderedHtml, defaultSubject);
  var textBody = htmlToText_(parsed.bodyHtml || '');

  return {
    success: true,
    templateName: 'mail-recherche-intervenant.html',
    subject: parsed.subject || defaultSubject,
    body: textBody || '',
    htmlBody: parsed.bodyHtml || renderedHtml || '',
    from: 'etudes@juniorinsaservices.fr'
  };
}

function normalizeMailRecipientsList_(raw) {
  var values = [];

  if (Object.prototype.toString.call(raw) === '[object Array]') {
    values = raw;
  } else if (typeof raw === 'string') {
    values = raw.split(/[;,\n]/);
  }

  var seen = {};
  var out = [];
  values.forEach(function(item) {
    var email = String(item || '').trim().toLowerCase();
    if (!email || email.indexOf('@') < 1 || seen[email]) return;
    seen[email] = true;
    out.push(email);
  });
  return out;
}

function resolveIntervenantDraftRecipients_(studyId, recipientsFromUi) {
  var directRecipients = normalizeMailRecipientsList_(recipientsFromUi);
  if (directRecipients.length) return directRecipients;

  var rawData = {};
  try {
    rawData = getStudyData(studyId) || {};
  } catch (e) {
    rawData = {};
  }

  var etudeInfo = rawData && rawData.EtudeInfo && typeof rawData.EtudeInfo === 'object'
    ? rawData.EtudeInfo
    : {};

  return normalizeMailRecipientsList_(etudeInfo.diffusionListEmails || []);
}

function createIntervenantSearchMailDraft(studyId, recipientsFromUi) {
  var prepared = prepareIntervenantSearchMail(studyId);

  var subject = prepared.subject || 'Recherche intervenant';
  var textBody = prepared.body || '';
  var htmlBody = prepared.htmlBody || '';
  var fromAlias = prepared.from || '';

  var options = { htmlBody: htmlBody };
  try {
    var aliases = GmailApp.getAliases();
    if (fromAlias && aliases && aliases.indexOf(fromAlias) !== -1) {
      options.from = fromAlias;
    }
  } catch (e) {}

  var recipients = resolveIntervenantDraftRecipients_(studyId, recipientsFromUi);
  var toValue = recipients.length ? recipients.join(',') : 'etudes@juniorinsaservices.fr';

  var draft = GmailApp.createDraft(toValue, subject, textBody, options);
  var draftId = draft.getId();

  return {
    success: true,
    draftId: draftId,
    recipients: recipients,
    draftUrl: 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftId)
  };
}

function createCahierDesChargesDoc(studyId) {
  if (!studyId) throw new Error('ID étude manquant.');

  var rawData = getStudyData(studyId) || {};
  var folders = getStudyDocumentsFolders_(studyId);

  var reference = rawData.reference
    ? String(rawData.reference)
    : [rawData.year || 'XXXX', rawData.code || 'XXX'].join('_');
  var documentName = reference + '_CDC';

  var generated = generateGenericDocument(documentName, folders.reviewFolder.getId(), 'cahier-des-charges', studyId);
  var cdcStatus = {
    hasDraft: !!(generated && generated.url),
    draftUrl: generated && generated.url ? generated.url : '',
    cdcName: generated && generated.name ? generated.name : documentName
  };

  try {
    if (typeof fsEtudePatch === 'function') {
      fsEtudePatch(studyId, {
        cdcStatus: cdcStatus,
        updatedAt: new Date().toISOString()
      });
    }
    if (typeof invalidateStudyLoadCaches_ === 'function') {
      invalidateStudyLoadCaches_(studyId);
    }
  } catch (e) {}

  return Object.assign({}, generated, { cdcStatus: cdcStatus });
}

function checkCdcStatus(studyId) {
  if (!studyId) {
    return { hasDraft: false, draftUrl: '', cdcName: '' };
  }

  var rawData = {};
  try {
    rawData = getStudyData(studyId) || {};
  } catch (e) {
    rawData = {};
  }

  var reference = rawData.reference
    ? String(rawData.reference)
    : [rawData.year || 'XXXX', rawData.code || 'XXX'].join('_');
  var cdcName = reference + '_CDC';
  var status = { hasDraft: false, draftUrl: '', cdcName: cdcName };

  try {
    var folders = getStudyDocumentsFolders_(studyId);
    var file = null;

    if (folders.reviewFolder) {
      file = findWorkflowDocumentFile_(folders.reviewFolder, cdcName);
    }
    if (!file && folders.docsFolder) {
      file = findWorkflowDocumentFile_(folders.docsFolder, cdcName);
    }

    if (file) {
      status.hasDraft = true;
      status.draftUrl = file.getUrl();
      status.cdcName = file.getName() || cdcName;
    }
  } catch (e2) {}

  return status;
}

function getProjectHtmlFileContent_(fileName) {
  if (!fileName || !String(fileName).trim()) {
    throw new Error('Nom de fichier HTML manquant.');
  }

  var candidates = buildProjectHtmlFileCandidates_(String(fileName).trim());
  var lastError = null;

  for (var i = 0; i < candidates.length; i++) {
    try {
      var output = HtmlService.createHtmlOutputFromFile(candidates[i]);
      return String(output.getContent() || '');
    } catch (err) {
      lastError = err;
    }
  }

  throw new Error(
    'Fichier HTML introuvable dans le projet: ' + fileName + '.html' +
    (candidates.length ? ' (noms testés: ' + candidates.join(', ') + ')' : '')
  );
}

function buildProjectHtmlFileCandidates_(fileName) {
  var name = String(fileName || '').trim();
  if (!name) return [];

  name = name.replace(/\.html$/i, '');

  var out = [];
  function addCandidate_(value) {
    var v = String(value || '').trim();
    if (!v) return;
    if (out.indexOf(v) === -1) out.push(v);
  }

  addCandidate_(name);
  addCandidate_(name.replace(/\//g, '_'));

  if (name.indexOf('/') !== -1) {
    var parts = name.split('/');
    var base = parts[parts.length - 1] || '';
    addCandidate_(base);
  } else {
    addCandidate_('gestionEtudes/templates/' + name);
    addCandidate_('gestionEtudes_templates_' + name);
    addCandidate_('templates/' + name);
    addCandidate_('templates_' + name);
    addCandidate_('gestionEtudes/' + name);
    addCandidate_('gestionEtudes_' + name);
  }

  return out;
}

function getMailTemplateText_(templateFile) {
  if (!templateFile) throw new Error('Template mail introuvable.');

  var mimeType = templateFile.getMimeType();
  if (mimeType === MimeType.GOOGLE_DOCS) {
    var doc = DocumentApp.openById(templateFile.getId());
    var body = getDocumentBodySafe_(doc);
    return (body.getText() || '').trim();
  }

  if (mimeType === 'application/vnd.google-apps.mail-layout') {
    var plain = exportGoogleWorkspaceFileAsText_(templateFile.getId(), 'text/plain');
    if (plain && String(plain).trim()) return String(plain).trim();

    var html = exportGoogleWorkspaceFileAsText_(templateFile.getId(), 'text/html');
    if (html && String(html).trim()) return htmlToText_(String(html));

    throw new Error('Impossible d\'extraire le contenu de la template mail-layout.');
  }

  if (mimeType && mimeType.indexOf('text/') === 0) {
    return String(templateFile.getBlob().getDataAsString() || '').trim();
  }

  throw new Error('Type de template mail non supporté: ' + mimeType + '. Utilisez Google Docs ou un fichier texte.');
}

function exportGoogleWorkspaceFileAsText_(fileId, exportMimeType) {
  try {
    var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/export?mimeType=' + encodeURIComponent(exportMimeType);
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      return resp.getContentText();
    }
    return '';
  } catch (e) {
    return '';
  }
}

function htmlToText_(html) {
  var text = String(html || '');
  text = text.replace(/<style[\s\S]*?<\/style>/gi, '');
  text = text.replace(/<script[\s\S]*?<\/script>/gi, '');
  text = text.replace(/<\s*br\s*\/?\s*>/gi, '\n');
  text = text.replace(/<\s*\/p\s*>/gi, '\n\n');
  text = text.replace(/<[^>]+>/g, '');

  text = text
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  text = text.replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n');
  return text.trim();
}

function extractMailSubjectAndBody_(renderedText, fallbackSubject) {
  var text = String(renderedText || '').replace(/\r\n/g, '\n');
  var lines = text.split('\n');
  var subject = '';
  var subjectLineIndex = -1;

  for (var i = 0; i < lines.length; i++) {
    var trimmed = String(lines[i] || '').trim();
    if (!trimmed) continue;

    var match = trimmed.match(/^(objet|subject)\s*:\s*(.+)$/i);
    if (match) {
      subject = String(match[2] || '').trim();
      subjectLineIndex = i;
    }
    break;
  }

  if (subjectLineIndex >= 0) {
    lines.splice(subjectLineIndex, 1);
  }

  return {
    subject: subject || String(fallbackSubject || ''),
    body: lines.join('\n').replace(/^\s+|\s+$/g, '')
  };
}

function extractMailSubjectAndBodyFromHtml_(renderedHtml, fallbackSubject) {
  var html = String(renderedHtml || '');
  var subject = '';

  var commentMatch = html.match(/<!--\s*(?:objet|subject)\s*:\s*([^>]+?)\s*-->/i);
  if (commentMatch && commentMatch[1]) {
    subject = String(commentMatch[1]).trim();
    html = html.replace(commentMatch[0], '');
  }

  if (!subject) {
    var h1Match = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
    if (h1Match && h1Match[1]) {
      subject = htmlToText_(h1Match[1]).trim();
    }
  }

  return {
    subject: subject || String(fallbackSubject || ''),
    bodyHtml: html
  };
}

// Alias FR explicites
function preparerMailDepuisTemplate(studyId, templateName, defaultSubject, fromAlias, extraVariables) {
  return prepareMailFromTemplate(studyId, templateName, defaultSubject, fromAlias, extraVariables);
}

function preparerMailRechercheIntervenant(studyId) {
  return prepareIntervenantSearchMail(studyId);
}

function creerBrouillonMailRechercheIntervenant(studyId, recipientsFromUi) {
  return createIntervenantSearchMailDraft(studyId, recipientsFromUi);
}

function genererCahierDesChargesEtude(studyId) {
  return createCahierDesChargesDoc(studyId);
}

function buildAutofillVariablesMap_(studyId, extraVariables) {
  var map = {};
  var now = new Date();
  addVariable_(map, 'today', now);

  var juniorInfo = getJuniorAutofillInfo_();
  addVariable_(map, 'junior.nom', juniorInfo.nom || 'Junior INSA Services');
  addVariable_(map, 'junior.signature', juniorInfo.signature || "Junior-Entreprise de l'INSA de Toulouse");
  addVariable_(map, 'junior.statut_juridique', juniorInfo.statut_juridique || 'Association régie par la loi de 1901 - Membre de la Confédération Nationale des Junior-Entreprises');
  addVariable_(map, 'junior.numero-telephone', juniorInfo['numero-telephone'] || '05 67 04 88 99');
  addVariable_(map, 'junior.telephone', juniorInfo['numero-telephone'] || '05 67 04 88 99');
  var president = normalizeJuniorPerson_(juniorInfo.president, { prenom: 'Jean-Gabriel', nom: 'Virgone' });
  var vicepresident = normalizeJuniorPerson_(juniorInfo.vicepresident, { prenom: 'Victor', nom: 'Schalk' });
  var tresorier = normalizeJuniorPerson_(juniorInfo.tresorier, { prenom: 'William', nom: 'Woodward' });

  addVariable_(map, 'junior.president.prenom', president.prenom);
  addVariable_(map, 'junior.president.nom', president.nom);
  addVariable_(map, 'junior.president.nom_complet', joinFullName_(president.prenom, president.nom));
  addVariable_(map, 'junior.president', joinFullName_(president.prenom, president.nom));
  addVariable_(map, 'junior.president.mail', president.mail || '');
  addVariable_(map, 'junior.president.sexe', president.sexe);
  addVariable_(map, 'junior.president.sexe', president.sexe);

  addVariable_(map, 'junior.vicepresident.prenom', vicepresident.prenom);
  addVariable_(map, 'junior.vicepresident.nom', vicepresident.nom);
  addVariable_(map, 'junior.vicepresident.nom_complet', joinFullName_(vicepresident.prenom, vicepresident.nom));
  addVariable_(map, 'junior.vicepresident', joinFullName_(vicepresident.prenom, vicepresident.nom));
  addVariable_(map, 'junior.vicepresident.mail', vicepresident.mail || '');
  addVariable_(map, 'junior.vicepresident.email', vicepresident.mail || '');
  addVariable_(map, 'junior.vicepresident.portable', vicepresident.portable || vicepresident.phone || juniorInfo['numero-telephone'] || '05 67 04 88 99');
  addVariable_(map, 'junior.vicepresident.sexe', vicepresident.sexe);

  addVariable_(map, 'junior.tresorier.prenom', tresorier.prenom);
  addVariable_(map, 'junior.tresorier.nom', tresorier.nom);
  addVariable_(map, 'junior.tresorier.nom_complet', joinFullName_(tresorier.prenom, tresorier.nom));
  addVariable_(map, 'junior.tresorier', joinFullName_(tresorier.prenom, tresorier.nom));
  addVariable_(map, 'junior.tresorier.mail', tresorier.mail || '');
  addVariable_(map, 'junior.tresorier.sexe', tresorier.sexe);
  addVariable_(map, 'junior.tresorier.sexe', tresorier.sexe);

  addVariable_(map, 'president.prenom', president.prenom);
  addVariable_(map, 'president.nom', president.nom);
  addVariable_(map, 'president.mail', president.mail || '');
  addVariable_(map, 'president.sexe', president.sexe);
  addVariable_(map, 'president.sexe', president.sexe);
  addVariable_(map, 'vicepresident.prenom', vicepresident.prenom);
  addVariable_(map, 'vicepresident.nom', vicepresident.nom);
  addVariable_(map, 'vicepresident.mail', vicepresident.mail || '');
  addVariable_(map, 'vicepresident.sexe', vicepresident.sexe);
  addVariable_(map, 'vicepresident.sexe', vicepresident.sexe);
  addVariable_(map, 'tresorier.prenom', tresorier.prenom);
  addVariable_(map, 'tresorier.nom', tresorier.nom);
  addVariable_(map, 'tresorier.mail', tresorier.mail || '');
  addVariable_(map, 'tresorier.sexe', tresorier.sexe);
  addVariable_(map, 'tresorier.sexe', tresorier.sexe);
  addVariable_(map, 'junior.adresse', juniorInfo.adresse || '135 Avenue de Rangueil 31077 TOULOUSE Cedex 04');
  addVariable_(map, 'junior.siteweb', juniorInfo.siteweb || 'https://juniorinsaservices.fr');
  addVariable_(map, 'junior.mail', juniorInfo.mail || 'contact@juniorinsaservices.fr');
  addVariable_(map, 'junior.email', juniorInfo.mail || 'contact@juniorinsaservices.fr');
  addVariable_(map, 'junior.siret', juniorInfo.siret || '341 760 528 00016');
  addVariable_(map, 'junior.ape', juniorInfo.ape || '9499Z');
  addVariable_(map, 'junior.tva_intracommunautaire', juniorInfo.tva_intracommunautaire || 'FR35 341760528');
  addVariable_(map, 'junior.numero-tva', juniorInfo.tva_intracommunautaire || 'FR35 341760528');
  addVariable_(map, 'junior.numero_tva', juniorInfo.tva_intracommunautaire || 'FR35 341760528');
  addVariable_(map, 'junior.tva', juniorInfo.tva_intracommunautaire || 'FR35 341760528');

  var rawData = {};
  if (studyId) {
    try {
      rawData = getStudyData(studyId) || {};
    } catch (e) {
      rawData = {};
    }
  }

  var clientInfo = rawData.ClientInfo || {};
  var etudeInfo = rawData.EtudeInfo || {};
  var phases = getSortedPhases_(Array.isArray(rawData.Planning_Phases) ? rawData.Planning_Phases : []);

  var contact = {};
  if (clientInfo.type === 'Particulier') {
    contact = clientInfo.particulier || {};
  } else {
    var contacts = Array.isArray(clientInfo.contacts) ? clientInfo.contacts : [];
    if (contacts.length > 0) {
      var signatory = contacts.filter(function(c) { return c && c.isSignatory; });
      contact = signatory.length > 0 ? signatory[0] : contacts[0];
    } else {
      contact = {};
    }
  }

  var clientType = String(clientInfo.type || '').trim();
  var clientTypeLower = clientType.toLowerCase();

  addVariable_(map, 'client.type', clientType);
  addVariable_(map, 'client.type_particulier', clientTypeLower === 'particulier' ? 1 : 0);
  addVariable_(map, 'client.type_entreprise', clientTypeLower === 'entreprise' ? 1 : 0);
  addVariable_(map, 'client.type_public', clientTypeLower === 'public' ? 1 : 0);
  addVariable_(map, 'client.type_je', clientTypeLower === 'je' ? 1 : 0);

  // Client / structure
  addVariable_(map, 'client.structure.nom', clientInfo.name || '');
  addVariable_(map, 'client.structure.adresse', clientInfo.address || '');
  addVariable_(map, 'client.structure.codepostal', clientInfo.postalCode || '');
  addVariable_(map, 'client.structure.ville', clientInfo.city || '');
  addVariable_(map, 'client.structure.capitalsocial', parseNumber_(clientInfo.capitalSocial));
  addVariable_(map, 'client.structure.description', clientInfo.description || '');

  // Signataire / représentant
  addVariable_(map, 'client.representant.nom', contact.lastName || '');
  addVariable_(map, 'client.representant.prenom', contact.firstName || '');
  addVariable_(map, 'client.representant.poste', contact.position || contact.poste || '');
  addVariable_(map, 'client.representant.civilite', contact.civility || contact.civilite || '');
  addVariable_(map, 'client.representant.mail', contact.email || '');
  addVariable_(map, 'client.representant.portable', contact.phone || '');
  addVariable_(map, 'client.nom', clientInfo.name || '');
  addVariable_(map, 'client.adresse', clientInfo.address || '');
  addVariable_(map, 'client.codepostal', clientInfo.postalCode || '');
  addVariable_(map, 'client.ville', clientInfo.city || '');
  addVariable_(map, 'client.capitalsocial', parseNumber_(clientInfo.capitalSocial));
  addVariable_(map, 'client.description', clientInfo.description || '');
  addVariable_(map, 'client.signataire.nom', contact.lastName || '');
  addVariable_(map, 'client.signataire.prenom', contact.firstName || '');
  addVariable_(map, 'client.signataire.civilite', contact.civility || contact.civilite || '');
  addVariable_(map, 'client.signataire.mail', contact.email || '');
  addVariable_(map, 'client.signataire.email', contact.email || '');
  addVariable_(map, 'client.signataire.poste', contact.position || contact.poste || '');
  addVariable_(map, 'client.signataire.portable', contact.phone || '');

  // Chargé d'étude
  addVariable_(map, 'charge.nom', etudeInfo.chargeNom || '');
  addVariable_(map, 'charge.prenom', etudeInfo.chargePrenom || '');
  addVariable_(map, 'charge.civilite', etudeInfo.chargeCivilite || '');
  addVariable_(map, 'charge.sexe', etudeInfo.chargeSexe || '');
  addVariable_(map, 'charge.annee', etudeInfo.chargePromo || etudeInfo.chargeAnnee || '');
  addVariable_(map, 'charge.specialite', etudeInfo.chargeSpecialite || '');
  addVariable_(map, 'charge.mail', etudeInfo.chargeEmail || '');
  addVariable_(map, 'charge.portable', etudeInfo.chargePortable || '');

  // Intervenant (souvent fourni ensuite via extraVariables)
  addVariable_(map, 'intervenant.nom', etudeInfo.intervenantNom || '');
  addVariable_(map, 'intervenant.prenom', etudeInfo.intervenantPrenom || '');
  addVariable_(map, 'intervenant.civilite', etudeInfo.intervenantCivilite || '');
  addVariable_(map, 'intervenant.sexe', etudeInfo.intervenantSexe || '');
  addVariable_(map, 'intervenant.annee', etudeInfo.intervenantAnnee || '');
  addVariable_(map, 'intervenant.specialite', etudeInfo.intervenantSpecialite || '');
  addVariable_(map, 'intervenant.mail', etudeInfo.intervenantMail || '');
  addVariable_(map, 'intervenant.portable', etudeInfo.intervenantPortable || '');
  addVariable_(map, 'intervenant.adresse', etudeInfo.intervenantAdresse || '');
  addVariable_(map, 'intervenant.email', etudeInfo.intervenantMail || '');
  addVariable_(map, 'intervenant.telephone', etudeInfo.intervenantPortable || '');
  addVariable_(map, 'intervenant.ecole', etudeInfo.intervenantEcole || 'INSA Toulouse');
  addVariable_(map, 'intervenant.niveau', etudeInfo.intervenantAnnee || '');
  addVariable_(map, 'intervenant.option', etudeInfo.intervenantOption || '');
  addVariable_(map, 'intervenant.postal_code', etudeInfo.intervenantCodePostal || '');
  addVariable_(map, 'intervenant.ville', etudeInfo.intervenantVille || '');

  var intervenantNomComplet = joinFullName_(etudeInfo.intervenantPrenom || '', etudeInfo.intervenantNom || '');
  addVariable_(map, 'intervenant.nom_complet', intervenantNomComplet);
  addVariable_(map, 'intervenant.prenom_nom', intervenantNomComplet);
  addVariable_(map, 'intervenant.inscription_sg_message', buildSgInscriptionMessage_(etudeInfo));

  // Étude
  var etudeNom = etudeInfo.nom || rawData.nom || '';
  var etudeReference = etudeInfo.reference || rawData.reference || [rawData.year || '', rawData.code || ''].filter(function(v) { return !!String(v).trim(); }).join('_');
  var categories = normalizeEtudeCategories_(etudeInfo.categories || etudeInfo.categorie || []);
  var categoriesText = categories.join(', ');

  addVariable_(map, 'etude.nom', etudeNom);
  addVariable_(map, 'etude.reference', etudeReference);
  addVariable_(map, 'etude.contexte', etudeInfo.contexte || '');
  addVariable_(map, 'etude.problematique', etudeInfo.problematique || '');
  addVariable_(map, 'etude.categories', categoriesText);
  addVariable_(map, 'etude.categorie', categoriesText);
  addVariable_(map, 'etude.reference_convention_client', etudeReference);
  addVariable_(map, 'etude.reference_cc', etudeReference ? (etudeReference + '_CC') : '');
  addVariable_(map, 'etude.reference_bdc', etudeReference ? (etudeReference + '_BDC') : '');
  addVariable_(map, 'etude.reference_ce', etudeReference ? (etudeReference + '_CE') : '');
  addVariable_(map, 'etude.reference_devis', etudeReference ? (etudeReference + '_DEVIS') : '');
  addVariable_(map, 'etude.reference_cdc', etudeReference ? (etudeReference + '_CDC') : '');

  var negociationMode = String(etudeInfo.negociationClientMode || etudeInfo.negociationMode || '').trim().toLowerCase();
  if (negociationMode === 'ccbdc' || negociationMode === 'cc+bdc' || negociationMode === 'cc+bcd' || negociationMode === 'cc/bdc') {
    negociationMode = 'cc_bdc';
  }
  addVariable_(map, 'negociation.mode', negociationMode);
  addVariable_(map, 'negociation.mode_ce', negociationMode === 'ce' ? 1 : 0);
  addVariable_(map, 'negociation.mode_cc_bdc', negociationMode === 'cc_bdc' ? 1 : 0);
  for (var ci = 0; ci < categories.length; ci++) {
    addVariable_(map, 'etude.categories[' + ci + ']', categories[ci]);
    addVariable_(map, 'etude.categorie[' + ci + ']', categories[ci]);
    addVariable_(map, 'etudes.categorie[' + ci + ']', categories[ci]);
  }

  // Prix
  var prixHT = parseNumber_(etudeInfo.prix);
  var frais = parseNumber_(etudeInfo.fraisEtude);
  var nbJEH = parseNumber_(etudeInfo.totalJeh || etudeInfo.totalJEH || etudeInfo.jehs || 0);

  var prixHTFrais = prixHT + frais;
  var tva = prixHTFrais * 0.2;
  var prixTTC = prixHTFrais + tva;

  addVariable_(map, 'etude.nbjeh', nbJEH);
  addVariable_(map, 'etude.nbJEH', nbJEH);
  addVariable_(map, 'etude.prix.ht', prixHT);
  addVariable_(map, 'etude.prix.frais', frais);
  addVariable_(map, 'etude.prix.ht.frais', prixHTFrais);
  addVariable_(map, 'etude.prix.tva', tva);
  addVariable_(map, 'etude.prix.ttc', prixTTC);

  var maxWeek = getMaxWeek_(phases);
  var etudeDuree = parseInt(etudeInfo.semaines || etudeInfo.duree, 10);
  if (!Number.isFinite(etudeDuree) || etudeDuree <= 0) {
    etudeDuree = maxWeek > 1 ? maxWeek : 0;
  }
  addVariable_(map, 'etude.duree', etudeDuree);
  addVariable_(map, 'etude.semaines', etudeDuree);

  // Phases (variables gérées dynamiquement dans les slides/tableaux)
  addVariable_(map, 'phases.planning', generatePhasesPlanningText_(phases));
  addVariable_(map, 'phase.planning', generatePhasesPlanningText_(phases));

  // Variables extra (objet ou JSON string), avec priorité finale
  var extras = parseExtraVariables_(extraVariables);
  flattenObjectToMap_(extras, '', map);

  return { variables: map, phases: phases };
}

function getJuniorAutofillInfo_() {
  var defaults = {
    nom: 'Junior INSA Services',
    signature: "Junior-Entreprise de l'INSA de Toulouse",
    statut_juridique: 'Association régie par la loi de 1901 - Membre de la Confédération Nationale des Junior-Entreprises',
    'numero-telephone': '05 67 04 88 99',
    president: { prenom: 'Jean-Gabriel', nom: 'Virgone', mail: '', sexe: 'H' },
    vicepresident: { prenom: 'Victor', nom: 'Schalk', mail: '', sexe: 'H' },
    tresorier: { prenom: 'William', nom: 'Woodward', mail: '', sexe: 'H' },
    adresse: '135 Avenue de Rangueil 31077 TOULOUSE Cedex 04',
    siteweb: 'https://juniorinsaservices.fr',
    mail: 'contact@juniorinsaservices.fr',
    siret: '341 760 528 00016',
    ape: '9499Z',
    tva_intracommunautaire: 'FR35 341760528'
  };

  try {
    if (typeof fsJuniorGetProfile === 'function') {
      var profile = fsJuniorGetProfile();
      if (profile && typeof profile === 'object') {
        for (var key in profile) {
          if (profile.hasOwnProperty(key) && profile[key] !== null && profile[key] !== undefined && profile[key] !== '') {
            defaults[key] = profile[key];
          }
        }
      }
    }
  } catch (e) {}

  defaults.president = normalizeJuniorPerson_(defaults.president, { prenom: 'Jean-Gabriel', nom: 'Virgone', sexe: 'H' });
  defaults.vicepresident = normalizeJuniorPerson_(defaults.vicepresident, { prenom: 'Victor', nom: 'Schalk', sexe: 'H' });
  defaults.tresorier = normalizeJuniorPerson_(defaults.tresorier, { prenom: 'William', nom: 'Woodward', sexe: 'H' });

  return defaults;
}

function normalizeJuniorPerson_(value, fallback) {
  var fb = fallback && typeof fallback === 'object' ? fallback : { prenom: '', nom: '', mail: '', sexe: '', portable: '' };

  if (value && typeof value === 'object' && !Array.isArray(value)) {
    return {
      prenom: String(value.prenom || fb.prenom || '').trim(),
      nom: String(value.nom || fb.nom || '').trim(),
      mail: String(value.mail || fb.mail || '').trim(),
      portable: String(value.portable || fallback.portable || '').trim(),
      sexe: String(value.sexe || fb.sexe || '').trim()
    };
  }

  if (typeof value === 'string' && value.trim()) {
    var raw = value.trim();
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

function joinFullName_(prenom, nom) {
  return [String(prenom || '').trim(), String(nom || '').trim()].filter(function(part) { return !!part; }).join(' ');
}

function buildSgInscriptionMessage_(etudeInfo) {
  var info = etudeInfo && typeof etudeInfo === 'object' ? etudeInfo : {};
  var fullName = joinFullName_(info.intervenantPrenom || '', info.intervenantNom || '');
  var email = String(info.intervenantMail || '').trim();
  var suffix = email ? (', email ' + email) : '';
  return 'Bonjour Secretaire General, merci dinscrire ' + (fullName || 'lintervenant') + suffix + '.';
}

function normalizeEtudeCategories_(raw) {
  var values = [];
  if (Array.isArray(raw)) {
    values = raw;
  } else if (typeof raw === 'string') {
    values = raw.split(',');
  } else if (raw !== null && raw !== undefined) {
    values = [String(raw)];
  }

  var seen = {};
  var normalized = [];
  for (var i = 0; i < values.length; i++) {
    var cat = String(values[i] || '').trim();
    if (!cat) continue;
    var key = cat.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    normalized.push(cat);
  }
  return normalized;
}

function replaceTokensInPresentation_(presentation, variables) {
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    replaceTokensInSlide_(slides[i], variables);
  }
}

function applyPhaseTokensInPresentation_(presentation, baseVariables, phases) {
  if (!Array.isArray(phases) || phases.length === 0) return;

  var originalSlides = presentation.getSlides().slice();
  for (var i = 0; i < originalSlides.length; i++) {
    var slide = originalSlides[i];

    if (slideHasPhaseTokensOutsideTables_(slide)) {
      var phaseSlides = [slide];
      for (var d = 1; d < phases.length; d++) {
        phaseSlides.push(slide.duplicate());
      }

      // Garantit l'ordre des slides: phase 1 (original), puis phase 2, phase 3...
      reorderPhaseSlidesAfterTemplate_(presentation, phaseSlides);

      for (var p = 0; p < phases.length; p++) {
        var targetSlide = phaseSlides[p];
        var phaseVars = buildPhaseVariables_(baseVariables, phases[p], p);
        replaceTokensInSlide_(targetSlide, phaseVars);
      }
    } else {
      duplicatePhaseRowsInSlide_(slide, baseVariables, phases);
    }
  }

  // Render planning placeholder as a visual table
  var allSlides = presentation.getSlides();
  for (var s = 0; s < allSlides.length; s++) {
    renderPlanningTableIfNeeded_(allSlides[s], phases);
  }
}

function reorderPhaseSlidesAfterTemplate_(presentation, phaseSlides) {
  if (!presentation || !phaseSlides || phaseSlides.length <= 1) return;

  var templateSlide = phaseSlides[0];
  var templateId;
  try {
    templateId = templateSlide.getObjectId();
  } catch (e) {
    return;
  }

  var templateIndex = getSlideIndexByObjectId_(presentation, templateId);
  if (templateIndex < 0) return;

  for (var i = 1; i < phaseSlides.length; i++) {
    try {
      phaseSlides[i].move(templateIndex + i);
    } catch (e2) {
      // En cas de limitation API, on conserve l'ordre existant
    }
  }
}

function getSlideIndexByObjectId_(presentation, objectId) {
  if (!presentation || !objectId) return -1;
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    try {
      if (slides[i].getObjectId() === objectId) return i;
    } catch (e) {}
  }
  return -1;
}

function replaceTokensInSlide_(slide, variables) {
  var elements = slide.getPageElements();
  for (var j = 0; j < elements.length; j++) {
    var type = elements[j].getPageElementType();

    if (type === SlidesApp.PageElementType.SHAPE) {
      try {
        var txt = elements[j].asShape().getText();
        replaceTokensInSlidesTextRange_(txt, variables);
      } catch (e) {}
    } else if (type === SlidesApp.PageElementType.TABLE) {
      try {
        var table = elements[j].asTable();
        for (var r = 0; r < table.getNumRows(); r++) {
          for (var c = 0; c < table.getNumColumns(); c++) {
            var cellTextRange;
            try {
              cellTextRange = table.getCell(r, c).getText();
            } catch (e2) {
              continue;
            }
            replaceTokensInSlidesTextRange_(cellTextRange, variables);
          }
        }
      } catch (e3) {}
    }
  }
}

function duplicatePhaseRowsInSlide_(slide, baseVariables, phases) {
  if (!Array.isArray(phases) || phases.length === 0) return;

  var elements = slide.getPageElements();
  for (var j = 0; j < elements.length; j++) {
    if (elements[j].getPageElementType() !== SlidesApp.PageElementType.TABLE) continue;

    var table;
    try {
      table = elements[j].asTable();
    } catch (e) {
      continue;
    }

    for (var r = 0; r < table.getNumRows(); r++) {
      var rowHasPhaseToken = false;
      var rowTexts = [];

      for (var c = 0; c < table.getNumColumns(); c++) {
        var cellText = safeGetTableCellText_(table, r, c);
        rowTexts.push(cellText);
        if (containsPhaseToken_(cellText)) rowHasPhaseToken = true;
      }

      if (!rowHasPhaseToken) continue;

      for (var p = 0; p < phases.length; p++) {
        var targetRow = r + p;
        if (p > 0) {
          table.insertRow(targetRow);
        }

        var phaseVars = buildPhaseVariables_(baseVariables, phases[p], p);
        for (var cc = 0; cc < table.getNumColumns(); cc++) {
          var content = rowTexts[cc] || '';
          var replaced = replaceTokensInText_(content, phaseVars);
          safeSetTableCellText_(table, targetRow, cc, replaced);
        }
      }

      r += phases.length - 1;
    }
  }
}

function slideHasPhaseTokensOutsideTables_(slide) {
  var elements = slide.getPageElements();
  for (var j = 0; j < elements.length; j++) {
    var type = elements[j].getPageElementType();
    if (type !== SlidesApp.PageElementType.SHAPE) continue;
    try {
      var txt = elements[j].asShape().getText().asString();
      if (containsPhaseToken_(txt)) return true;
    } catch (e) {}
  }
  return false;
}

function containsPhaseToken_(text) {
  return /\{[^{}]*\bphase\.[^{}]*\}/i.test(String(text || ''));
}

function getSortedPhases_(phases) {
  if (!Array.isArray(phases)) return [];

  return phases
    .map(function(phase, index) {
      return {
        phase: phase && typeof phase === 'object' ? phase : {},
        index: index,
        sortKey: getPhaseSortKey_(phase, index)
      };
    })
    .sort(function(a, b) {
      if (a.sortKey !== b.sortKey) return a.sortKey - b.sortKey;
      return a.index - b.index;
    })
    .map(function(item) {
      return item.phase;
    });
}

function getPhaseSortKey_(phase, index) {
  var explicit = parseInt(
    phase && (phase.numero || phase.num || phase.number || phase.order || phase.ordre),
    10
  );
  if (Number.isFinite(explicit) && explicit > 0) return explicit;

  var start = parseInt(phase && (phase.start || phase.debut || phase.semaineDebut), 10);
  if (Number.isFinite(start) && start > 0) return start + ((index || 0) * 0.0001);

  return (index || 0) + 1;
}

function resolvePhaseNumber_(phase, index) {
  var explicit = parseInt(
    phase && (phase.numero || phase.num || phase.number || phase.order || phase.ordre),
    10
  );
  if (Number.isFinite(explicit) && explicit > 0) return explicit;
  return (index || 0) + 1;
}

function resolvePhaseTitle_(phase, index) {
  if (phase && (phase.title || phase.titre || phase.nom)) {
    return String(phase.title || phase.titre || phase.nom);
  }
  return 'Phase ' + resolvePhaseNumber_(phase, index);
}

function resolvePhaseDurationWeeks_(phase) {
  var duration = parseInt(phase && (phase.duration || phase.duree || phase.semaines), 10);
  if (!Number.isFinite(duration) || duration <= 0) return 1;
  return duration;
}

function resolvePhaseStartWeek_(phase) {
  var start = parseInt(phase && (phase.start || phase.debut || phase.semaineDebut), 10);
  if (!Number.isFinite(start) || start <= 0) return 1;
  return start;
}

function buildPhaseVariables_(baseVariables, phase, index) {
  var map = {};
  for (var key in baseVariables) {
    if (baseVariables.hasOwnProperty(key)) map[key] = baseVariables[key];
  }

  var pIndex = resolvePhaseNumber_(phase, index);
  var pTitle = resolvePhaseTitle_(phase, index);
  var pDescription = (phase && (phase.description || phase.resume)) || '';
  var pDuration = resolvePhaseDurationWeeks_(phase);
  var pJeh = parseNumber_(
    phase && (phase.jeh || phase.nbJEH || phase.nbJeh || phase.nombreJeh || phase.nombre_jeh)
  );
  var pPrice = parseNumber_(
    phase && (
      phase.priceJeh || phase.prixParJEH || phase.prixJeH || phase.prixParJeh ||
      (phase.prix && (phase.prix.parJEH || phase.prix.parJeh || phase.prix.parjeh))
    )
  );
  var pTotalHt = parseNumber_(phase && phase.prix && (phase.prix.totalHT || phase.prix.totalht));
  if (!pTotalHt) pTotalHt = pJeh * pPrice;

  addVariable_(map, 'phase.nom', pTitle);
  addVariable_(map, 'phase.titre', pTitle);
  addVariable_(map, 'phase.title', pTitle);
  addVariable_(map, 'phase.description', pDescription);
  addVariable_(map, 'phase.numero', pIndex);
  addVariable_(map, 'phase.duree', pDuration);
  addVariable_(map, 'phase.duration', pDuration);
  addVariable_(map, 'phase.nbjeh', pJeh);
  addVariable_(map, 'phase.nombre-jeh', pJeh);
  addVariable_(map, 'phase.nombre_jeh', pJeh);
  addVariable_(map, 'phase.prix.parjeh', pPrice);
  addVariable_(map, 'phase.prix-jeh', pPrice);
  addVariable_(map, 'phase.prix.ht', pTotalHt);
  addVariable_(map, 'phase.prix.totalht', pTotalHt);
  addVariable_(map, 'phase.livrable', (phase && phase.livrable) || '');

  return map;
}

function safeGetTableCellText_(table, row, col) {
  try {
    return table.getCell(row, col).getText().asString();
  } catch (e) {
    return '';
  }
}

function safeSetTableCellText_(table, row, col, value) {
  try {
    table.getCell(row, col).getText().setText(value);
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    if (msg.indexOf('fusionnée') !== -1 || msg.indexOf('merged') !== -1) {
      return;
    }
    throw e;
  }
}

function findTemplateTokens_(text) {
  return String(text || '').match(/\{\{[^{}]+\}\}|\{[^{}]+\}/g) || null;
}

function extractTokenInside_(token) {
  var t = String(token || '');
  if (!t) return '';
  if (/^\{\{[\s\S]+\}\}$/.test(t)) return t.substring(2, t.length - 2);
  if (/^\{[\s\S]+\}$/.test(t)) return t.substring(1, t.length - 1);
  return t;
}

function replaceTokensInSlidesTextRange_(textRange, variables) {
  var source;
  try {
    source = textRange.asString();
  } catch (e) {
    return;
  }

  if (!source || source.indexOf('{') === -1) return;

  var tokens = findTemplateTokens_(source);
  if (!tokens || !tokens.length) return;

  var unique = {};
  for (var i = 0; i < tokens.length; i++) unique[tokens[i]] = true;

  for (var token in unique) {
    if (!unique.hasOwnProperty(token)) continue;
    var inside = extractTokenInside_(token);
    var value = evaluateToken_(inside, variables);
    var replacement = value === null || typeof value === 'undefined' ? '' : String(value);

    if (isLikelyHtmlContent_(replacement)) {
      if (typeof replaceHtmlTokenInSlidesTextRange_ === 'function' && replaceHtmlTokenInSlidesTextRange_(textRange, token, replacement)) {
        continue;
      }
      replacement = typeof convertHtmlToPlainText === 'function' ? convertHtmlToPlainText(replacement) : replacement;
    }

    try {
      textRange.replaceAllText(token, replacement);
    } catch (e2) {
      var msg = String(e2 && e2.message ? e2.message : e2);
      if (msg.indexOf('fusionnée') !== -1 || msg.indexOf('merged') !== -1) {
        return;
      }
      throw e2;
    }
  }
}

function replaceTokensInDocument_(doc, variables, phases) {
  var body = getDocumentBodySafe_(doc);
  
  applyPhaseTokensInDocument_(body, variables || {}, phases || []);
  replaceTokensInDocElement_(body, variables);
  replacePlanningTokenInDocument_(body, variables || {});

  try {
    var header = doc.getHeader();
    if (header) {
      applyPhaseTokensInDocument_(header, variables || {}, phases || []);
      replaceTokensInDocElement_(header, variables);
      replacePlanningTokenInDocument_(header, variables || {});
    }
  } catch(e) {}
  
  try {
    var footer = doc.getFooter();
    if (footer) {
       applyPhaseTokensInDocument_(footer, variables || {}, phases || []);
       replaceTokensInDocElement_(footer, variables);
       replacePlanningTokenInDocument_(footer, variables || {});
    }
  } catch(e) {}
}

function getDocumentBodySafe_(docOrBody) {
  if (docOrBody && typeof docOrBody.getBody === 'function') {
    var bodyFromDoc = docOrBody.getBody();
    if (bodyFromDoc && typeof bodyFromDoc.getNumChildren === 'function') {
      return bodyFromDoc;
    }
  }

  if (docOrBody && typeof docOrBody.getNumChildren === 'function' && typeof docOrBody.getTables === 'function') {
    return docOrBody;
  }

  throw new Error('Le moteur de generation attend un Google Doc valide (body introuvable). Verifiez le type de la template utilisee.');
}

function applyPhaseTokensInDocument_(body, baseVariables, phases) {
  if (!body) return;

  var orderedPhases = getSortedPhases_(phases || []);

  if (!orderedPhases.length) {
    try {
      body.replaceText('\\{\\s*phases\\s*\\}', '');
      body.replaceText('\\{\\s*/\\s*phases\\s*\\}', '');
    } catch (e) {}
    return;
  }

  duplicatePhaseRowsInDocumentBody_(body, baseVariables, orderedPhases);
  duplicatePhaseBlocksInDocumentBody_(body, baseVariables, orderedPhases);
  replaceInlinePhaseTokensInDocElement_(body, baseVariables, orderedPhases);

  try {
    body.replaceText('\\{\\s*phases\\s*\\}', '');
    body.replaceText('\\{\\s*/\\s*phases\\s*\\}', '');
  } catch (e2) {}
}

function duplicatePhaseRowsInDocumentBody_(body, baseVariables, phases) {
  if (!body || !Array.isArray(phases) || !phases.length) return;

  var tables = body.getTables();
  for (var t = 0; t < tables.length; t++) {
    duplicatePhaseRowsInDocumentTable_(tables[t], baseVariables, phases);
  }
}

function duplicatePhaseRowsInDocumentTable_(table, baseVariables, phases) {
  var groups = [];
  var rowCount = table.getNumRows();

  for (var r = 0; r < rowCount; r++) {
    var row = table.getRow(r);
    var hasPhaseToken = containsPhaseToken_(row.getText ? row.getText() : '');
    if (!hasPhaseToken) continue;

    if (groups.length && groups[groups.length - 1].end === r - 1) {
      groups[groups.length - 1].end = r;
    } else {
      groups.push({ start: r, end: r });
    }
  }

  for (var g = groups.length - 1; g >= 0; g--) {
    var group = groups[g];
    var rowCopies = [];
    for (var rr = group.start; rr <= group.end; rr++) {
      rowCopies.push(table.getRow(rr).copy());
    }

    for (var del = group.end; del >= group.start; del--) {
      table.removeRow(del);
    }

    var insertAt = group.start;
    for (var p = 0; p < phases.length; p++) {
      var phaseVars = buildPhaseVariables_(baseVariables, phases[p], p);
      for (var rc = 0; rc < rowCopies.length; rc++) {
        var insertedRow = table.insertTableRow(insertAt, rowCopies[rc].copy());
        replaceTokensInDocElement_(insertedRow, phaseVars);
        insertAt++;
      }
    }
  }
}

function duplicatePhaseBlocksInDocumentBody_(body, baseVariables, phases) {
  var blocks = [];
  var currentBlock = null;
  var count = body.getNumChildren();

  for (var i = 0; i < count; i++) {
    var child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      if (currentBlock) {
        blocks.push(currentBlock);
        currentBlock = null;
      }
      continue;
    }

    var text = '';
    try { text = child.getText ? child.getText() : ''; } catch (e) { text = ''; }

    var hasPhaseToken = containsPhaseToken_(text);
    if (hasPhaseToken) {
      if (!currentBlock) currentBlock = { start: i, end: i };
      else currentBlock.end = i;
    } else if (currentBlock) {
      blocks.push(currentBlock);
      currentBlock = null;
    }
  }

  if (currentBlock) blocks.push(currentBlock);

  for (var b = blocks.length - 1; b >= 0; b--) {
    var block = blocks[b];
    var copies = [];
    for (var ci = block.start; ci <= block.end; ci++) {
      copies.push(body.getChild(ci).copy());
    }

    for (var rem = block.end; rem >= block.start; rem--) {
      body.getChild(rem).removeFromParent();
    }

    var insertIndex = block.start;
    for (var p = 0; p < phases.length; p++) {
      var phaseVars = buildPhaseVariables_(baseVariables, phases[p], p);
      for (var c = 0; c < copies.length; c++) {
        var inserted = insertCopiedDocumentChild_(body, insertIndex, copies[c].copy());
        if (inserted) {
          replaceTokensInDocElement_(inserted, phaseVars);
          insertIndex++;
        }
      }
    }
  }
}

function insertCopiedDocumentChild_(body, childIndex, elementCopy) {
  if (!body || !elementCopy) return null;

  var type = elementCopy.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) {
    return body.insertParagraph(childIndex, elementCopy.asParagraph());
  }
  if (type === DocumentApp.ElementType.LIST_ITEM) {
    return body.insertListItem(childIndex, elementCopy.asListItem());
  }
  if (type === DocumentApp.ElementType.TABLE) {
    return body.insertTable(childIndex, elementCopy.asTable());
  }

  var txt = '';
  try { txt = elementCopy.getText ? elementCopy.getText() : ''; } catch (e) { txt = ''; }
  return body.insertParagraph(childIndex, txt);
}

function replaceInlinePhaseTokensInDocElement_(element, baseVariables, phases) {
  var type = element.getType();

  if (type === DocumentApp.ElementType.TEXT) {
    var textNode = element.asText();
    var source = textNode.getText() || '';

    var expanded = expandPhaseTemplateBlocksInText_(source, baseVariables, phases);
    if (expanded !== source) {
      textNode.setText(expanded);
      source = expanded;
    }

    if (containsPhaseToken_(source)) {
      var rendered = [];
      for (var i = 0; i < phases.length; i++) {
        var vars = buildPhaseVariables_(baseVariables, phases[i], i);
        rendered.push(replaceTokensInText_(source, vars));
      }
      textNode.setText(rendered.join('\n'));
    }
    return;
  }

  if (typeof element.getNumChildren !== 'function') return;
  var childCount = element.getNumChildren();
  for (var c = 0; c < childCount; c++) {
    replaceInlinePhaseTokensInDocElement_(element.getChild(c), baseVariables, phases);
  }
}

function expandPhaseTemplateBlocksInText_(text, baseVariables, phases) {
  var source = String(text || '');
  if (source.indexOf('{phases}') === -1 && source.indexOf('{/phases}') === -1) {
    return source;
  }

  return source.replace(/\{\s*phases\s*\}([\s\S]*?)\{\s*\/\s*phases\s*\}/gi, function(_m, templateBlock) {
    if (!Array.isArray(phases) || !phases.length) return '';

    var chunks = [];
    for (var i = 0; i < phases.length; i++) {
      var vars = buildPhaseVariables_(baseVariables, phases[i], i);
      chunks.push(replaceTokensInText_(templateBlock, vars));
    }
    return chunks.join('');
  });
}

function replacePlanningTokenInDocument_(body, variables) {
  if (!body || !variables) return;
  var planningValue = getVariable_(variables, 'phases.planning');
  if (planningValue === null || typeof planningValue === 'undefined') return;

  try {
    body.replaceText('\\{\\s*phases\\.planning\\s*\\}', String(planningValue));
  } catch (e) {}
}

function replaceTokensInDocElement_(element, variables) {
  var type = element.getType();

  if (type === DocumentApp.ElementType.TEXT) {
    var textNode = element.asText();
    var text = textNode.getText();
    var tokens = text ? findTemplateTokens_(text) : null;
    if (tokens && tokens.length) {
      var unique = {};
      for (var u = 0; u < tokens.length; u++) unique[tokens[u]] = true;
      for (var token in unique) {
        if (!unique.hasOwnProperty(token)) continue;
        var inside = extractTokenInside_(token);
        var value = evaluateToken_(inside, variables);
        var replacement = value === null || typeof value === 'undefined' ? '' : String(value);

        if (isLikelyHtmlContent_(replacement)) {
          if (typeof replaceHtmlTokenInDocsTextNode_ === 'function' && replaceHtmlTokenInDocsTextNode_(textNode, token, replacement)) {
            return;
          }
          replacement = typeof convertHtmlToPlainText === 'function' ? convertHtmlToPlainText(replacement) : replacement;
        }

        textNode.replaceText(escapeRegex_(token), replacement);
      }
    }
    return;
  }

  if (typeof element.getNumChildren !== 'function') {
    return;
  }

  var childCount = element.getNumChildren();
  for (var i = 0; i < childCount; i++) {
    replaceTokensInDocElement_(element.getChild(i), variables);
  }
}

function replaceTokensInText_(text, variables) {
  if (!text || text.indexOf('{') === -1) return text;

  return String(text).replace(/\{\{([^{}]+)\}\}|\{([^{}]+)\}/g, function(match, insideDouble, insideSingle) {
    var inside = insideDouble || insideSingle || '';
    var value = evaluateToken_(inside, variables);
    var strValue = value === null || typeof value === 'undefined' ? '' : String(value);
    if (isLikelyHtmlContent_(strValue) && typeof convertHtmlToPlainText === 'function') {
      strValue = convertHtmlToPlainText(strValue);
    }
    return strValue;
  });
}

function evaluateToken_(tokenInside, variables) {
  var raw = normalizeQuotes_(String(tokenInside || '').trim());
  raw = raw.replace(/^([a-zA-Z0-9_.-]+)\s*"/, '$1,"');
  if (!raw) return '';

  var parts = splitByCommaOutsideQuotes_(raw);
  if (!parts.length) return '';

  var expression = (parts[0] || '').trim();
  if (!expression) return '';

  // Condition explicite: {if:condition, "si vrai", "si faux"}
  if (/^if\s*:/i.test(expression)) {
    var conditionExpr = expression.replace(/^if\s*:/i, '').trim();
    var whenTrue = parts.length >= 2 ? resolveConditionOutput_(parts[1], variables) : '';
    var whenFalse = parts.length >= 3 ? resolveConditionOutput_(parts[2], variables) : '';
    return evaluateBooleanCondition_(conditionExpr, variables) ? whenTrue : whenFalse;
  }

  // Condition implicite: {variable == "x", "si vrai", "si faux"}
  if (parts.length >= 3 && isConditionalExpression_(expression)) {
    var condTrue = resolveConditionOutput_(parts[1], variables);
    var condFalse = resolveConditionOutput_(parts[2], variables);
    return evaluateBooleanCondition_(expression, variables) ? condTrue : condFalse;
  }

  // Condition de genre/pluriel: {variable, "A", "B"}
  var isTernaryOperator = parts.length >= 3 && String(parts[1]).indexOf('=') === -1 && String(parts[2]).indexOf('=') === -1;

  if (isTernaryOperator || (parts.length >= 3 && isQuotedText_(parts[1]) && isQuotedText_(parts[2]))) {
    var resolvedValue = resolveValue_(expression, variables);
    var optionA = stripWrappingQuotes_(parts[1]);
    var optionB = stripWrappingQuotes_(parts[2]);

    if (isGenderExpression_(expression) || isGenderValue_(resolvedValue)) {
      return selectTextByGender_(resolvedValue, optionA, optionB);
    }

    if (isNumericLikeValue_(resolvedValue)) {
      return selectTextByCount_(resolvedValue, optionA, optionB);
    }

    var truthy = !(resolvedValue === null || typeof resolvedValue === 'undefined' || resolvedValue === '' || resolvedValue === 0 || resolvedValue === false);
    return truthy ? optionA : optionB;
  }

  
  // Switch/Case mapping: {variable, key="value", key2="value2"}
  var isMapping = false;
  var mapping = {};
  for (var i = 1; i < parts.length; i++) {
    var p = parts[i].trim();
    var eqIdx = p.indexOf('=');
    if (eqIdx > 0) {
      var k = normalizeMappingKey_(p.substring(0, eqIdx).trim());
      var v = p.substring(eqIdx + 1).trim();
      if (isQuotedText_(v)) {
        mapping[k] = stripWrappingQuotes_(v);
        isMapping = true;
      }
    }
  }
  
  if (isMapping) {
    var resolved = resolveValue_(expression, variables);
    var varValue = (resolved === null || typeof resolved === 'undefined' || resolved === '') ? expression : resolved;
    var normalizedValue = normalizeMappingKey_(varValue);

    if (mapping.hasOwnProperty(normalizedValue)) {
      return mapping[normalizedValue];
    }
    if (mapping.hasOwnProperty('default')) {
      return mapping['default'];
    }
    return '';
  }

  // Options: format / decimales / lettres

  if (parts.length >= 2) {
    var rawOption = (parts[1] || '').trim();

    if (/^format\s*:/i.test(rawOption)) {
      var rawFormatString = parts.slice(1).join(',').trim();
      var fmt = rawFormatString.replace(/^format\s*:\s*/i, '');
      var dateValue = resolveDateValue_(expression, variables);
      return dateValue ? formatDateWithPattern_(dateValue, fmt) : '';
    }

    var option = rawOption.toLowerCase();

    if (/^decimales\s*:\s*\d+$/i.test(option)) {
      var dec = parseInt(option.split(':')[1], 10);
      var numDec = resolveNumericValue_(expression, variables);
      return formatNumberWithDecimals_(numDec, dec);
    }

    if (option === 'lettres') {
      var numRaw = resolveNumericValue_(expression, variables);
      return numberToFrenchWordsGeneric_(numRaw);
    }
    
    if (option === 'lettres_euros' || option === 'lettres_prix') {
      var numMoney = resolveNumericValue_(expression, variables);
      return numberToFrenchMoneyWords_(numMoney);
    }
  }

  // S'il y a une valeur directe correspondante à la clé avec la même casse, on ne passe pas en arithmétique
  var directVal = getVariable_(variables, expression);
  if (directVal !== undefined && directVal !== null && directVal !== '') {
    if (directVal instanceof Date) {
      return formatDateWithPattern_(directVal, 'dd/MM/yyyy');
    }
    if (typeof directVal === 'number' && Number.isFinite(directVal)) {
      return String(directVal).replace('.', ',');
    }
    return String(directVal);
  }

  // Valeur simple ou expression arithmétique générique
  if (isLikelyNumericExpression_(expression)) {
    var n = resolveNumericValue_(expression, variables);
    return Number.isFinite(n) ? String(n).replace('.', ',') : '';
  }

  return resolveValue_(expression, variables);
}

function resolveConditionOutput_(part, variables) {
  var v = String(part || '').trim();
  if (!v) return '';
  if (isQuotedText_(v)) return stripWrappingQuotes_(v);
  var resolved = resolveValue_(v, variables);
  return resolved === null || typeof resolved === 'undefined' ? '' : String(resolved);
}

function isConditionalExpression_(expression) {
  return /(==|!=|>=|<=|>|<)/.test(String(expression || ''));
}

function resolveConditionOperand_(operand, variables) {
  var op = String(operand || '').trim();
  if (!op) return '';

  if (isQuotedText_(op)) {
    return stripWrappingQuotes_(op);
  }

  if (/^-?\d+(?:[.,]\d+)?$/.test(op)) {
    return parseNumber_(op);
  }

  var byValue = resolveValue_(op, variables);
  if (typeof byValue !== 'undefined' && byValue !== null && byValue !== '') {
    return byValue;
  }

  return op;
}

function evaluateBooleanCondition_(conditionExpr, variables) {
  var expr = String(conditionExpr || '').trim();
  if (!expr) return false;

  var m = expr.match(/^(.*?)\s*(==|!=|>=|<=|>|<)\s*(.*?)$/);
  if (!m) {
    var unary = expr;
    var negate = false;
    if (unary.charAt(0) === '!') {
      negate = true;
      unary = unary.substring(1).trim();
    }
    var value = resolveConditionOperand_(unary, variables);
    var truthy = !(value === null || typeof value === 'undefined' || value === '' || value === 0 || value === false || value === 'false');
    return negate ? !truthy : truthy;
  }

  var left = resolveConditionOperand_(m[1], variables);
  var right = resolveConditionOperand_(m[3], variables);
  var operator = m[2];

  var leftNum = parseNumber_(left);
  var rightNum = parseNumber_(right);
  var bothNumbers = Number.isFinite(leftNum) && Number.isFinite(rightNum) &&
    !(String(left).trim() === '' || String(right).trim() === '');

  if (bothNumbers) {
    if (operator === '==') return leftNum === rightNum;
    if (operator === '!=') return leftNum !== rightNum;
    if (operator === '>=') return leftNum >= rightNum;
    if (operator === '<=') return leftNum <= rightNum;
    if (operator === '>') return leftNum > rightNum;
    if (operator === '<') return leftNum < rightNum;
    return false;
  }

  var leftStr = String(left || '').toLowerCase();
  var rightStr = String(right || '').toLowerCase();

  if (operator === '==') return leftStr === rightStr;
  if (operator === '!=') return leftStr !== rightStr;
  if (operator === '>=') return leftStr >= rightStr;
  if (operator === '<=') return leftStr <= rightStr;
  if (operator === '>') return leftStr > rightStr;
  if (operator === '<') return leftStr < rightStr;
  return false;
}

function resolveValue_(expression, variables) {
  var expr = String(expression || '').trim();
  if (!expr) return '';

  if (isQuotedText_(expr)) {
    return stripWrappingQuotes_(expr);
  }

  var byKey = getVariable_(variables, expr);
  if (typeof byKey !== 'undefined' && byKey !== null) {
    if (byKey instanceof Date) {
      return formatDateWithPattern_(byKey, 'dd/MM/yyyy');
    }
    return byKey;
  }

  if (isLikelyNumericExpression_(expr)) {
    var n = evaluateNumericExpression_(expr, variables);
    return Number.isFinite(n) ? n : '';
  }

  return '';
}

function resolveDateValue_(expression, variables) {
  var expr = String(expression || '').trim();
  if (!expr) return null;

  var value = getVariable_(variables, expr);
  if (value instanceof Date) return value;

  if (expr.toLowerCase() === 'today') return new Date();

  if (typeof value === 'string' && value) {
    var d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}

function resolveNumericValue_(expression, variables) {
  var value = resolveValue_(expression, variables);
  if (typeof value === 'number') return value;
  return parseNumber_(value);
}

function evaluateNumericExpression_(expression, variables) {
  var expr = String(expression || '').trim();
  if (!expr) return NaN;

  // Convertit les décimales FR en format JS (ex: 10,5 -> 10.5)
  expr = expr.replace(/(\d),(\d)/g, '$1.$2');

  var replaced = expr.replace(/[a-zA-Z_][a-zA-Z0-9_.-]*/g, function(symbol) {
    var value = getVariable_(variables, symbol);
    if (typeof value === 'undefined' || value === null || value === '') {
      return '0';
    }
    return String(parseNumber_(value));
  });

  if (/[^0-9+\-*/().\s]/.test(replaced)) {
    return NaN;
  }

  try {
    var result = Function('"use strict"; return (' + replaced + ');')();
    return Number(result);
  } catch (e) {
    return NaN;
  }
}

function isLikelyNumericExpression_(expression) {
  return /[+\-*/()]/.test(expression) || /^\d+(?:[.,]\d+)?$/.test(String(expression || '').trim());
}

function splitByCommaOutsideQuotes_(text) {
  var result = [];
  var current = '';
  var quote = ''; // peut être ", ', “, ou ‘
  var escaped = false;

  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);

    if (escaped) {
      current += ch;
      escaped = false;
      continue;
    }

    if (ch === '\\') {
      current += ch;
      escaped = true;
      continue;
    }

    if (quote) {
      var isPartner = false;
      if (quote === '"' && ch === '"') isPartner = true;
      else if (quote === "'" && ch === "'") isPartner = true;
      else if (quote === '“' && (ch === '“' || ch === '”')) isPartner = true;
      else if (quote === '‘' && (ch === '‘' || ch === '’')) isPartner = true;
      
      if (isPartner) {
        quote = '';
        current += ch;
        continue;
      }
    } else {
      if (ch === '"' || ch === "'") { quote = ch; current += ch; continue; }
      else if (ch === '“' || ch === '”') { quote = '“'; current += ch; continue; }
      else if (ch === '‘' || ch === '’') { quote = '‘'; current += ch; continue; }
    }

    if (ch === ',' && !quote) {
      result.push(current.trim());
      current = '';
      continue;
    }

    current += ch;
  }

  if (current || text.endsWith(',')) {
    result.push(current.trim());
  }

  return result;
}

function normalizeQuotes_(text) {
  return String(text || '')
    .replace(/[“”«»]/g, '"')
    .replace(/[’‘]/g, "'");
}

function isQuotedText_(value) {
  var v = String(value || '').trim();
  if (v.length < 2) return false;
  var s = v.charAt(0);
  var e = v.charAt(v.length - 1);
  if (s === '"' && e === '"') return true;
  if (s === "'" && e === "'") return true;
  if ((s === '“' || s === '”') && (e === '“' || e === '”')) return true;
  if ((s === '‘' || s === '’') && (e === '‘' || e === '’')) return true;
  return false;
}

function stripWrappingQuotes_(value) {
  var v = String(value || '').trim();
  if (isQuotedText_(v)) {
    return v.substring(1, v.length - 1);
  }
  // Aggressive strip for unbalanced quotes created by document typos
  return v.replace(/^["'“‘”’]+|["'“‘”’]+$/g, '').trim();
}

function isGenderExpression_(expression) {
  return /(?:^|\.)sexe$/i.test(String(expression || '').trim());
}

function isGenderValue_(value) {
  var g = String(value || '').trim().toLowerCase();
  return g === 'femme' || g === 'f' || g === 'female' || g === 'feminin' || g === 'féminin' ||
         g === 'homme' || g === 'h' || g === 'm' || g === 'male' || g === 'masculin';
}

function isNumericLikeValue_(value) {
  if (typeof value === 'number') return Number.isFinite(value);
  if (value === null || typeof value === 'undefined') return false;
  return /^-?\d+(?:[.,]\d+)?$/.test(String(value).trim());
}

function selectTextByCount_(countValue, singularText, pluralText) {
  var n = parseNumber_(countValue);
  return Math.abs(n) === 1 ? singularText : pluralText;
}

function normalizeMappingKey_(value) {
  var key = String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/[._-]/g, '');

  if (!key) return '';

  if (key === 'public' || key === 'publique' || key === 'publiqueadministration') return 'publique';
  if (key === 'prive' || key === 'privee' || key === 'particulier' || key === 'entreprise' || key === 'societe' || key === 'private') return 'prive';
  if (key === 'je' || key === 'juniorentreprise' || key === 'junior') return 'je';

  return key;
}

function selectTextByGender_(genderValue, femaleText, maleText) {
  var g = String(genderValue || '').trim().toLowerCase();
  if (g === 'femme' || g === 'f' || g === 'female' || g === 'feminin' || g === 'féminin') return femaleText;
  if (g === 'homme' || g === 'h' || g === 'm' || g === 'male' || g === 'masculin') return maleText;
  return femaleText || maleText || '';
}

function addVariable_(map, key, value) {
  map[normalizeVariableKey_(key)] = value;
}

function getVariable_(map, key) {
  var val = map[normalizeVariableKey_(key)];
  if (typeof val === 'string') {
    return val.trim();
  }
  return val;
}

function normalizeVariableKey_(key) {
  return String(key || '').trim().toLowerCase();
}

function parseNumber_(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (value === null || typeof value === 'undefined') return 0;

  var s = String(value).trim();
  if (!s) return 0;

  s = s.replace(/\s/g, '').replace(',', '.');
  var n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function formatDateWithPattern_(dateObj, pattern) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    return '';
  }
  var p = String(pattern || '').trim();
  if (!p) p = 'dd/MM/yyyy';
  var normalized = normalizeDatePattern_(p);
  return Utilities.formatDate(dateObj, 'Europe/Paris', normalized);
}

function normalizeDatePattern_(pattern) {
  var p = String(pattern || '').trim();
  if (!p) return 'dd/MM/yyyy';
  return p
    .replace(/YYYY/g, 'yyyy')
    .replace(/YY/g, 'yy')
    .replace(/DD/g, 'dd');
}

function escapeRegex_(text) {
  return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function formatNumberWithDecimals_(num, decimals) {
  var d = Math.max(0, parseInt(decimals, 10) || 0);
  var n = Number(num);
  if (!Number.isFinite(n)) n = 0;
  return n.toFixed(d).replace('.', ',');
}

function parseExtraVariables_(extraVariables) {
  if (!extraVariables) return {};
  if (typeof extraVariables === 'string') {
    try {
      return JSON.parse(extraVariables) || {};
    } catch (e) {
      return {};
    }
  }
  if (typeof extraVariables === 'object') return extraVariables;
  return {};
}

function flattenObjectToMap_(obj, prefix, targetMap) {
  if (!obj || typeof obj !== 'object') return;

  if (Array.isArray(obj)) {
    return;
  }

  for (var key in obj) {
    if (!obj.hasOwnProperty(key)) continue;
    var nextPath = prefix ? (prefix + '.' + key) : key;
    var value = obj[key];

    if (value && typeof value === 'object' && !Array.isArray(value)) {
      flattenObjectToMap_(value, nextPath, targetMap);
    } else {
      addVariable_(targetMap, nextPath, value);
    }
  }
}

function generatePhasesPlanningText_(phases) {
  var ordered = getSortedPhases_(phases || []);
  if (!ordered.length) return '';

  var lines = [];
  for (var i = 0; i < ordered.length; i++) {
    var p = ordered[i] || {};
    var phaseNumber = resolvePhaseNumber_(p, i);
    var title = resolvePhaseTitle_(p, i);
    var start = resolvePhaseStartWeek_(p);
    var duration = resolvePhaseDurationWeeks_(p);
    var end = start + duration - 1;
    lines.push('Phase ' + phaseNumber + ' - ' + title + ' : S' + start + ' à S' + end);
  }

  return lines.join('\n');
}

function numberToFrenchLiteral_(num) {
  var value = Number(num);
  if (!Number.isFinite(value)) value = 0;

  var negative = value < 0;
  var absValue = Math.abs(value);
  var integerPart = Math.floor(absValue);
  var decimalPart = Math.round((absValue - integerPart) * 100);

  var text = numberToFrenchWordsGeneric_(integerPart);
  if (decimalPart > 0) {
    text += ' virgule ' + numberToFrenchWordsGeneric_(decimalPart);
  }

  if (negative) text = 'moins ' + text;
  return text;
}

function numberToFrenchMoneyWords_(num) {
  var value = Number(num);
  if (!Number.isFinite(value)) value = 0;
  var euros = Math.floor(Math.abs(value));
  var cents = Math.round((Math.abs(value) - euros) * 100);

  var euroPart = numberToFrenchWordsGeneric_(euros) + (euros > 1 ? ' euros' : ' euro');
  if (cents > 0) {
    return euroPart + ' et ' + numberToFrenchWordsGeneric_(cents) + (cents > 1 ? ' centimes' : ' centime');
  }
  return euroPart;
}

function renderPlanningTableIfNeeded_(slide, phases) {
  if (!Array.isArray(phases) || phases.length === 0) return;

  var elements = slide.getPageElements();
  for (var i = 0; i < elements.length; i++) {
    if (elements[i].getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;

    var shape;
    var sourceText;
    try {
      shape = elements[i].asShape();
      sourceText = shape.getText().asString();
    } catch (e) {
      continue;
    }

    if (!/\{\s*phases\.planning\s*\}/i.test(sourceText || '')) continue;

    var left = shape.getLeft();
    var top = shape.getTop();
    var width = shape.getWidth();
    var height = shape.getHeight();

    var maxWeek = getMaxWeek_(phases);
    var rows = phases.length + 1;
    var cols = maxWeek + 1;

    // Agrandissement automatique pour un rendu lisible
    var minWidth = Math.max(560, 95 + (maxWeek * 40));
    var minHeight = Math.max(240, 34 + (rows * 26));
    width = Math.max(width, minWidth);
    height = Math.max(height, minHeight);

    var table = slide.insertTable(rows, cols, left, top, width, height);

    // Widths (first col wider for labels)
    try {
      var firstColWidth = Math.max(70, width * 0.2);
      table.getColumn(0).setWidth(firstColWidth);
      var otherW = (width - firstColWidth) / Math.max(1, maxWeek);
      for (var c = 1; c < cols; c++) {
        table.getColumn(c).setWidth(otherW);
      }
    } catch (e2) {}

    // Header
    setPlanningCell_(table.getCell(0, 0), 'Phase', '#F3F4F6', '#111111', true);
    for (var w = 1; w <= maxWeek; w++) {
      setPlanningCell_(table.getCell(0, w), 'S' + w, '#F3F4F6', '#111111', true);
    }

    // Body
    for (var p = 0; p < phases.length; p++) {
      var phase = phases[p] || {};
      var start = Math.max(1, parseInt(phase.start, 10) || 1);
      var duration = Math.max(1, parseInt(phase.duration, 10) || 1);
      var end = start + duration - 1;

      setPlanningCell_(table.getCell(p + 1, 0), 'P' + (p + 1), '#FFFFFF', '#111111', true);

      for (var ww = 1; ww <= maxWeek; ww++) {
        var active = ww >= start && ww <= end;
        setPlanningCell_(
          table.getCell(p + 1, ww),
          '',
          active ? '#000000' : '#FFFFFF',
          active ? '#FFFFFF' : '#111111',
          false
        );
      }
    }

    try { shape.remove(); } catch (e3) { try { shape.getText().setText(''); } catch (e4) {} }
  }
}

function getMaxWeek_(phases) {
  var maxWeek = 1;
  for (var i = 0; i < phases.length; i++) {
    var p = phases[i] || {};
    var start = Math.max(1, parseInt(p.start, 10) || 1);
    var duration = Math.max(1, parseInt(p.duration, 10) || 1);
    maxWeek = Math.max(maxWeek, start + duration - 1);
  }
  return maxWeek;
}

function setPlanningCell_(cell, text, bgColor, fgColor, bold) {
  try {
    cell.getText().setText(String(text || ''));
    var txtStyle = cell.getText().getTextStyle();
    txtStyle.setBold(!!bold);
    txtStyle.setForegroundColor(fgColor || '#111111');
    try { txtStyle.setFontSize(11); } catch (e) {}
  } catch (e2) {}

  try {
    cell.getFill().setSolidFill(bgColor || '#FFFFFF');
  } catch (e3) {}
}

function numberToFrenchWordsGeneric_(num) {
  num = Math.floor(Number(num || 0));
  if (num === 0) return 'zéro';

  var units = ['', 'un', 'deux', 'trois', 'quatre', 'cinq', 'six', 'sept', 'huit', 'neuf'];
  var teens = ['dix', 'onze', 'douze', 'treize', 'quatorze', 'quinze', 'seize', 'dix-sept', 'dix-huit', 'dix-neuf'];
  var tens = ['', '', 'vingt', 'trente', 'quarante', 'cinquante', 'soixante'];

  function underHundred(n) {
    if (n < 10) return units[n];
    if (n < 20) return teens[n - 10];
    if (n < 70) {
      var t = Math.floor(n / 10);
      var u = n % 10;
      if (u === 1) return tens[t] + ' et un';
      return tens[t] + (u ? '-' + units[u] : '');
    }
    if (n < 80) {
      if (n === 71) return 'soixante et onze';
      return 'soixante-' + underHundred(n - 60);
    }
    if (n === 80) return 'quatre-vingts';
    return 'quatre-vingt-' + underHundred(n - 80);
  }

  function underThousand(n) {
    if (n < 100) return underHundred(n);
    var c = Math.floor(n / 100);
    var r = n % 100;
    var prefix = c === 1 ? 'cent' : units[c] + ' cent';
    if (r === 0 && c > 1) prefix += 's';
    return r ? prefix + ' ' + underHundred(r) : prefix;
  }

  var parts = [];
  var millions = Math.floor(num / 1000000);
  var thousands = Math.floor((num % 1000000) / 1000);
  var rest = num % 1000;

  if (millions) {
    parts.push((millions === 1 ? 'un million' : underThousand(millions) + ' millions'));
  }
  if (thousands) {
    parts.push(thousands === 1 ? 'mille' : underThousand(thousands) + ' mille');
  }
  if (rest) {
    parts.push(underThousand(rest));
  }

  return parts.join(' ');
}


function obtenirListeVersionsAvenantsEtude(studyId, docKey) {
  var cfg = getDocWorkflowConfig_(docKey);
  var rawData = getStudyData(studyId) || {};
  var reference = getStudyReferenceForDocs_(rawData);
  var baseName = reference + '_' + cfg.suffix;
  
  var folders = getStudyDocumentsFolders_(studyId);
  var folderReview = folders.reviewFolder;
  var folderDocs = folders.docsFolder;

  var safeBase = String(baseName).replace(/[.*+?^${()|[\]\\]/g, '\\$&');
  var regex;
  if (docKey === 'bdc') {
    regex = new RegExp('^' + safeBase + '(\\d+)?(?:\\.[a-z0-9]+)?$', 'i');
  } else {
    regex = new RegExp('^' + safeBase + '(?:_v(\\d+))?(?:\\.[a-z0-9]+)?$', 'i');
  }

  var results = [];
  
  function scanFolder_(folder) {
    if (!folder) return;
    var all = folder.getFiles();
    while (all.hasNext()) {
      var file = all.next();
      var mime = String(file.getMimeType() || '').toLowerCase();
      var name = file.getName();
      if (mime !== 'application/pdf' && regex.test(name)) {
        var match = name.match(regex);
        var v = match && match[1] ? parseInt(match[1], 10) : 1;
        results.push({
          id: file.getId(),
          name: name,
          url: file.getUrl(),
          version: isNaN(v) ? 1 : v,
          updated: file.getLastUpdated().getTime(),
          updatedStr: Utilities.formatDate(file.getLastUpdated(), "Europe/Paris", "dd/MM/yyyy HH:mm")
        });
      }
    }
  }
  
  scanFolder_(folderReview);
  scanFolder_(folderDocs);
  
  results.sort(function(a, b) {
    if (b.version !== a.version) return b.version - a.version;
    return b.updated - a.updated;
  });
  
  return results;
}

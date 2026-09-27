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
/**
 * Generation de documents generiques et conversion des templates Office.
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
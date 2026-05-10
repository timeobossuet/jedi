function formatMeetingStampForCr_(isoLikeValue) {
  var d = new Date(String(isoLikeValue || '').trim());
  if (!(d instanceof Date) || isNaN(d.getTime())) {
    d = new Date();
  }

  var tz = Session.getScriptTimeZone() || 'Europe/Paris';
  return Utilities.formatDate(d, tz, 'yyyyMMdd_HHmm');
}

function sanitizeMeetingIdToken_(value) {
  return String(value || '')
    .trim()
    .replace(/[^a-zA-Z0-9_-]/g, '')
    .slice(-16);
}

function findExistingFileByName_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function buildMeetingReportName_(studyId, options) {
  var opts = options && typeof options === 'object' ? options : {};
  var rawData = getStudyData(studyId) || {};
  var reference = getStudyReferenceForDocs_(rawData);

  var customName = String(opts.fileName || '').trim();
  if (customName) return customName;

  var stamp = formatMeetingStampForCr_(opts.meetingIso || opts.startIso || opts.dateIso || '');
  var base = reference + '_CR_REUNION_' + stamp;

  if (opts.customTitle && opts.customTitle.trim()) {
    var safeTitle = opts.customTitle.replace(/[^a-zA-Z0-9_ -]/g, '_').replace(/\s+/g, '_');
    base += '_' + safeTitle;
  } else {
    var token = sanitizeMeetingIdToken_(opts.meetingId || '');
    if (token) {
      base += '_' + token;
    }
  }

  return base;
}

/**
 * Crée un compte-rendu à partir du template "cr-reunion-etude"
 * Nommage par défaut: [REF]_CR_REUNION_YYYYMMDD_HHmm(_meetingId)
 * Si le fichier existe déjà pour ce nom, il est réutilisé (idempotent).
 */
function createMeetingReportDoc(studyId, options) {
  if (!studyId) throw new Error('studyId manquant.');

  var opts = options && typeof options === 'object' ? options : {};
  var templateId = getTemplateFileIdByName('cr-reunion-etude');
  var templateFile = DriveApp.getFileById(templateId);
  var exchangeFolder = getStudyExchangeFolder(studyId);

  var baseName = buildMeetingReportName_(studyId, opts);
  var existing = findExistingFileByName_(exchangeFolder, baseName);
  if (existing) {
    return {
      success: true,
      id: existing.getId(),
      name: existing.getName(),
      url: existing.getUrl(),
      folderUrl: exchangeFolder.getUrl(),
      meetingId: String(opts.meetingId || ''),
      meetingIso: String(opts.meetingIso || opts.startIso || opts.dateIso || ''),
      reused: true
    };
  }

  var finalName = baseName;
  var seq = 2;
  while (findExistingFileByName_(exchangeFolder, finalName)) {
    finalName = baseName + '_v' + seq;
    seq++;
  }

  var newFile = templateFile.makeCopy(finalName, exchangeFolder);

  return {
    success: true,
    id: newFile.getId(),
    name: finalName,
    url: newFile.getUrl(),
    folderUrl: exchangeFolder.getUrl(),
    meetingId: String(opts.meetingId || ''),
    meetingIso: String(opts.meetingIso || opts.startIso || opts.dateIso || ''),
    reused: false
  };
}

// --- Alias FR (plus explicite) ---
function creerCompteRenduReunionEtude(studyId, options) {
  return createMeetingReportDoc(studyId, options);
}

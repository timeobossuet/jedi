function obtenirListeVersionsAvenantsEtude(studyId, docKey) {
  var cfg = getDocWorkflowConfig_(docKey);
  var rawData = getStudyData(studyId) || {};
  var reference = getStudyReferenceForDocs_(rawData);
  var baseName = reference + '_' + cfg.suffix;
  var folders = getStudyDocumentsFolders_(studyId);
  var folderReview = folders.reviewFolder;
  var folderDocs = folders.docsFolder;

  var safeBase = String(baseName).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
        var version = match && match[1] ? parseInt(match[1], 10) : 1;
        results.push({
          id: file.getId(),
          name: name,
          url: file.getUrl(),
          version: isNaN(version) ? 1 : version,
          updated: file.getLastUpdated().getTime(),
          updatedStr: Utilities.formatDate(file.getLastUpdated(), 'Europe/Paris', 'dd/MM/yyyy HH:mm')
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
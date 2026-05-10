// Fonctions RGPD — gestion des demandes

function getRGPDRequests() {
  var sheet = getSGSheet_(SG_CONFIG.sheets.rgpd);
  var rows = readRows(sheet);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    out.push({
      id: String(row[0]),
      date: row[1] || '',
      demandeur: row[2] || '',
      type: row[3] || '',
      statut: row[4] || 'En attente',
      commentaire: row[5] || '',
      dateTraitement: row[6] || ''
    });
  }
  return out;
}

function addRGPDRequest(data) {
  var sheet = getSGSheet_(SG_CONFIG.sheets.rgpd);
  var id = data.id || _generateId_();
  var row = [
    id,
    data.date || fmtDateForSheet(new Date(), 'Europe/Paris', 'yyyy-MM-dd'),
    data.demandeur || '',
    data.type || '',
    data.statut || 'En attente',
    data.commentaire || '',
    data.dateTraitement || ''
  ];
  var res = upsertRowById(sheet, id, row);
  return { success: true, id: id, action: res.action };
}

function updateRGPDRequest(data) {
  var sheet = getSGSheet_(SG_CONFIG.sheets.rgpd);
  var id = data.id;
  var rowIdx = findRowIndexById(sheet, id);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 4).setValue(data.type || '');
    sheet.getRange(rowIdx, 5).setValue(data.statut || '');
    sheet.getRange(rowIdx, 6).setValue(data.commentaire || '');
    sheet.getRange(rowIdx, 7).setValue(data.dateTraitement || '');
    return { success: true };
  }
  return { success: false };
}
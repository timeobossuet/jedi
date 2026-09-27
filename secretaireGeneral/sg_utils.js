// Utilities for SG module
function safeJsonParse(text, fallback) {
  try { return text ? JSON.parse(text) : (fallback === undefined ? null : fallback); } catch (e) { return fallback === undefined ? null : fallback; }
}

function fmtDateForSheet(value, tz, pattern) {
  try {
    if (!value) return '';
    var d = (value instanceof Date) ? value : new Date(value);
    return Utilities.formatDate(d, tz || 'Europe/Paris', pattern || 'dd/MM/yyyy');
  } catch (e) { return '';
  }
}

function ensureHeaderRow(sheet, headers) {
  try {
    var rng = sheet.getRange(1,1,1,headers.length);
    rng.setValues([headers]);
    rng.setFontWeight('bold');
  } catch(e) {}
}

function findRowIndexById(sheet, id) {
  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) return i+1; // 1-based
    }
    return -1;
  } catch(e) { return -1; }
}

function upsertRowById(sheet, id, rowValues) {
  var idx = findRowIndexById(sheet, id);
  if (idx > 0) {
    sheet.getRange(idx, 1, 1, rowValues.length).setValues([rowValues]);
    return { action: 'updated', row: idx };
  } else {
    sheet.appendRow(rowValues);
    return { action: 'created', row: sheet.getLastRow() };
  }
}

function deleteRowById(sheet, id) {
  var idx = findRowIndexById(sheet, id);
  if (idx > 0) { sheet.deleteRow(idx); return true; }
  return false;
}

function readRows(sheet) {
  try {
    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];
    return data.slice(1);
  } catch(e) { return []; }
}

function findById(items, id) {
  if (!items || !items.length) return null;
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].id) === String(id)) return items[i];
  }
  return null;
}

function readExtraJson(sheet, rowIdx, colIdx) {
  try {
    var raw = sheet.getRange(rowIdx, colIdx).getValue();
    return safeJsonParse(raw, {});
  } catch (e) { return {} }
}

function writeExtraJson(sheet, rowIdx, colIdx, obj) {
  try {
    sheet.getRange(rowIdx, colIdx).setValue(JSON.stringify(obj || {}));
    return true;
  } catch (e) { return false; }
}

function getFolderByIdOrDefault(id) {
  try { return DriveApp.getFolderById(id); }
  catch (e) {
    try { return DriveApp.getFolderById(getSGDriveFolderId_()); }
    catch (e2) { return DriveApp.getRootFolder(); }
  }
}

function getFolderIdOrDefault(id) {
  var f = getFolderByIdOrDefault(id);
  try { return f.getId(); } catch (e) { return DriveApp.getRootFolder().getId(); }
}

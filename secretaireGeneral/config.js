

// Fonctions liées à la configuration
function getConfigValue(key, fallback) {
	try {
		var ss = getSGSpreadsheet_();
		var sheet = ss.getSheetByName(SG_CONFIG.sheets.config);
		if (!sheet) return fallback;
		var rowIdx = findRowIndexById(sheet, key);
		if (rowIdx > 0) return sheet.getRange(rowIdx, 2).getValue();
		return fallback;
	} catch (e) { return fallback; }
}

function setConfigValue(key, value) {
	try {
		var ss = getSGSpreadsheet_();
		var sheet = ss.getSheetByName(SG_CONFIG.sheets.config);
		if (!sheet) sheet = ss.insertSheet(SG_CONFIG.sheets.config);
		var res = upsertRowById(sheet, key, [key, value]);
		return { success: true, action: res.action };
	} catch (e) { throw new Error('Erreur setConfigValue: ' + e.message); }
}

function getAllConfig() {
	var out = {};
	try {
		var ss = getSGSpreadsheet_();
		var sheet = ss.getSheetByName(SG_CONFIG.sheets.config);
		if (!sheet) return out;
		var rows = readRows(sheet);
		for (var i = 0; i < rows.length; i++) {
			var k = rows[i][0];
			var v = rows[i][1];
			if (!k) continue;
			try { out[k] = safeJsonParse(v, v); }
			catch (e) { out[k] = v; }
		}
	} catch (e) { }
	return out;
}

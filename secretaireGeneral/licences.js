

// Fonctions liées aux licences
function getLicences() {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.licences);
		var rows  = readRows(sheet);
		var out = [];
		for (var i = 0; i < rows.length; i++) {
			var row = rows[i];
			if (!row[0]) continue;
			var item = {
				id: String(row[0]),
				logiciel: row[1] || '',
				type: row[2] || '',
				dateInstallation: row[3] ? fmtDateForSheet(row[3], 'Europe/Paris', 'dd/MM/yyyy') : '',
				dateDebutValidite: row[4] ? fmtDateForSheet(row[4], 'Europe/Paris', 'dd/MM/yyyy') : '',
				dateFinValidite: row[5] ? fmtDateForSheet(row[5], 'Europe/Paris', 'dd/MM/yyyy') : '',
				proprietaire: row[6] || '',
				contexte: row[7] || '',
				preuve: row[8] || '',
				conditions: row[9] || '',
				remarques: row[10] || ''
			};
			if (row[11]) {
				var extra = safeJsonParse(row[11], {});
				for (var k in extra) { if (extra.hasOwnProperty(k)) item[k] = extra[k]; }
			}
			out.push(item);
		}
		return out;
	} catch (e) { throw new Error('Erreur getLicences: ' + e.message); }
}

function saveLicence(licenceData) {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.licences);
		var id = licenceData.id || _generateId_();
		var row = [
			id,
			licenceData.logiciel || '',
			licenceData.type || '',
			licenceData.dateInstallation || '',
			licenceData.dateDebutValidite || '',
			licenceData.dateFinValidite || '',
			licenceData.proprietaire || '',
			licenceData.contexte || '',
			licenceData.preuve || '',
			licenceData.conditions || '',
			licenceData.remarques || '',
			JSON.stringify(licenceData.data || {})
		];
		var res = upsertRowById(sheet, id, row);
		return { success: true, id: id, action: res.action };
	} catch (e) { throw new Error('Erreur saveLicence: ' + e.message); }
}

function deleteLicence(id) {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.licences);
		var ok = deleteRowById(sheet, id);
		return { success: ok };
	} catch (e) { throw new Error('Erreur deleteLicence: ' + e.message); }
}

function attachLicenceProof(id, proofValue) {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.licences);
		var rowIdx = findRowIndexById(sheet, id);
		if (rowIdx > 0) {
			sheet.getRange(rowIdx, 9).setValue(proofValue);
			return { success: true };
		}
		return { success: false };
	} catch (e) { throw new Error('Erreur attachLicenceProof: ' + e.message); }
}

function sendLicenceRemindersNow() {
	try {
		var webhook = getConfigValue('google_chat_webhook', '');
		var daysStr = getConfigValue('licence_reminder_days', '30');
		var days = parseInt(daysStr, 10) || 30;
		if (!webhook || !webhook.toString().match(/^https?:\/\//i)) {
			return { success: false, error: 'Webhook Google Chat non configuré (key: google_chat_webhook).' };
		}
		var licences = getLicences();
		var now = new Date();
		var threshold = new Date(now.getTime() + days * 24 * 3600 * 1000);
		var warn = [];
		licences.forEach(function(l) {
			if (!l.dateFinValidite) return;
			var d;
			try { d = new Date(l.dateFinValidite); } catch (e) { d = null; }
			if (!d || isNaN(d.getTime())) {
				var parts = String(l.dateFinValidite).split('/');
				if (parts.length === 3) {
					d = new Date(parts[2], parts[1]-1, parts[0]);
				}
			}
			if (!d || isNaN(d.getTime())) return;
			if (d <= threshold) {
				warn.push({ licence: l.logiciel || l.logiciel, fin: fmtDateForSheet(d, 'Europe/Paris', 'dd/MM/yyyy'), id: l.id });
			}
		});
		if (warn.length === 0) {
			return { success: true, message: 'Aucune licence à alerter.' };
		}
		var text = 'Alerte licences — prochaines expirations (dans ' + days + ' jours):\n';
		warn.forEach(function(w) { text += '- ' + (w.licence || '—') + ' : ' + w.fin + '\n'; });
		var payload = JSON.stringify({ text: text });
		var opts = { method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true };
		var resp = UrlFetchApp.fetch(webhook, opts);
		return { success: true, sent: warn.length, responseCode: resp.getResponseCode() };
	} catch (e) {
		return { success: false, error: e.message };
	}
}

function createLicenceReminderTrigger() {
	try {
		_clearTriggersForFunction_('sendLicenceRemindersNow');
		ScriptApp.newTrigger('sendLicenceRemindersNow').timeBased().everyDays(1).atHour(8).create();
		return { success: true };
	} catch (e) { return { success: false, error: e.message }; }
}

function deleteLicenceReminderTriggers() {
	try { _clearTriggersForFunction_('sendLicenceRemindersNow'); return { success: true }; } catch (e) { return { success: false, error: e.message }; }
}

function copyFileToSgFolder(fileId) {
	try {
		if (!fileId) throw new Error('fileId manquant');
		var targetFolderId = getConfigValue('licences_folder_id', SG_CONFIG.driveFolderId);
		var folder = getFolderByIdOrDefault(targetFolderId);
		var file = DriveApp.getFileById(fileId);
		var copyName = file.getName() + ' (copie pour licences)';
		var newFile = file.makeCopy(copyName, folder);
		return { success: true, id: newFile.getId(), url: newFile.getUrl(), name: newFile.getName() };
	} catch (e) { return { success: false, error: e.message }; }
}

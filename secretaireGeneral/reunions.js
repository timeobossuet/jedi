
// Fonctions liées aux réunions

// Fonctions liées aux réunions

function getReunions() {
	var cached = fromCache_('reunions');
	if (cached) return cached;
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.reunions);
		var rows  = readRows(sheet);
		var reunions = [];
		for (var i = 0; i < rows.length; i++) {
			var row = rows[i];
			if (!row[0]) continue;
			var reunion = {
				id: String(row[0]),
				date: row[1] || '',
				heure: row[2] || '',
				lieu: row[3] || '',
				type: row[4] || 'CA',
				statut: row[5] || 'Planifié',
				data: row[6] ? safeJsonParse(row[6], {}) : {}
			};
			reunions.push(reunion);
		}
		toCache_('reunions', reunions);
		return reunions;
	} catch (e) {
		throw new Error('Impossible de charger les réunions: ' + e.message);
	}
}

function saveReunion(reunionData) {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.reunions);
		var id    = reunionData.id || _generateId_();
		var extra = reunionData.data || {};
		var row = [
			id,
			reunionData.date || '',
			reunionData.heure || '',
			reunionData.lieu || '',
			reunionData.type || 'CA',
			reunionData.statut || 'Planifié',
			JSON.stringify(extra)
		];
		var res = upsertRowById(sheet, id, row);
		invalidateCache_('reunions');
		return { success: true, id: id, action: res.action };
	} catch (e) {
		throw new Error('Erreur sauvegarde réunion: ' + e.message);
	}
}

function deleteReunion(reunionId) {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.reunions);
		var ok = deleteRowById(sheet, reunionId);
		if (ok) invalidateCache_('reunions');
		return { success: ok };
	} catch (e) { throw new Error('Erreur suppression réunion: ' + e.message); }
}

// Pour genererODJ et genererCRCA, il faut probablement réutiliser la logique de génération de documents
// Si tu veux que je les recode aussi, dis-le moi !

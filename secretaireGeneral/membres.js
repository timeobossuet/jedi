// Fonctions liées aux membres
// Importé depuis Code_SG.js

function _defaultChecklist_() {
	return { kiwi: false, ba_signe: false, cotisation: false };
}

function getMembres() {
	var cached = fromCache_('membres');
	if (cached) return cached;

	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.membres);
		var rows  = readRows(sheet);
		var membres = [];
		for (var i = 0; i < rows.length; i++) {
			var row = rows[i];
			if (!row[0]) continue;
			var membre = {
				id:              String(row[0]),
				civilite:        row[1] || '',
				nom:             row[2] || '',
				prenom:          row[3] || '',
				email:           row[4] || '',
				telephone:       row[5] || '',
				promo:           row[6] || '',
				specialite:      row[7] || '',
				roleCA:          row[8] || '',
				statut:          row[9] || 'Actif',
				dateInscription: row[10] ? fmtDateForSheet(row[10], 'Europe/Paris', 'dd/MM/yyyy') : ''
			};
			if (row[11]) {
				var extra = safeJsonParse(row[11], {});
				// Migration : supprimer beequick si présent
				if (extra.checklist && extra.checklist.hasOwnProperty('beequick')) {
					delete extra.checklist.beequick;
				}
				for (var key in extra) { if (extra.hasOwnProperty(key)) membre[key] = extra[key]; }
			}
			if (!membre.checklist) membre.checklist = _defaultChecklist_();
			membres.push(membre);
		}
		toCache_('membres', membres);
		return membres;
	} catch (e) {
		throw new Error('Impossible de charger les membres: ' + e.message);
	}
}

function saveMembre(membreData) {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.membres);
		var data  = sheet.getDataRange().getValues();
		var id    = membreData.id || _generateId_();

		// Nettoyage checklist (on retire beequick si présent)
		var cl = membreData.checklist || _defaultChecklist_();
		if (cl.hasOwnProperty('beequick')) delete cl.beequick;

		var extra = {
			checklist: cl,
			notes:     membreData.notes   || '',
			adresse:   membreData.adresse || ''
		};

		var row = [
			id,
			membreData.civilite        || '',
			(membreData.nom || '').toUpperCase(),
			membreData.prenom          || '',
			membreData.email           || '',
			membreData.telephone       || '',
			membreData.promo           || '',
			membreData.specialite      || '',
			membreData.roleCA          || '',
			membreData.statut          || 'Actif',
			membreData.dateInscription || fmtDateForSheet(new Date(), 'Europe/Paris', 'yyyy-MM-dd'),
			JSON.stringify(extra)
		];

		var res = upsertRowById(sheet, id, row);
		invalidateCache_('membres');
		return { success: true, id: id, action: res.action };
	} catch (e) {
		throw new Error('Erreur sauvegarde membre: ' + e.message);
	}
}

function deleteMembre(membreId) {
	try {
		var sheet = getSGSheet_(SG_CONFIG.sheets.membres);
		var ok = deleteRowById(sheet, membreId);
		if (ok) invalidateCache_('membres');
		return { success: ok };
	} catch (e) { throw new Error('Erreur suppression membre: ' + e.message); }
}

function updateMembreChecklist(membreId, checklist) {
	try {
		// Nettoyage beequick
		if (checklist && checklist.hasOwnProperty('beequick')) delete checklist.beequick;
		var sheet = getSGSheet_(SG_CONFIG.sheets.membres);
		var rowIdx = findRowIndexById(sheet, membreId);
		if (rowIdx > 0) {
			var extra = readExtraJson(sheet, rowIdx, 12);
			extra.checklist = checklist;
			writeExtraJson(sheet, rowIdx, 12, extra);
			invalidateCache_('membres');
			return { success: true };
		}
		return { success: false };
	} catch (e) { throw new Error('Erreur update checklist: ' + e.message); }
}

function saveMembre(membreData) {
	// ...existing code...
	var res = _saveMembreImpl && _saveMembreImpl(membreData);
	// Génération automatique du BA si nouveau membre
	if (res && res.success && (!membreData.id || res.action === 'insert')) {
		try {
			genererBA(res.id || membreData.id, 'membre');
		} catch (e) {
			// Log erreur mais ne bloque pas la sauvegarde
			Logger && Logger.log && Logger.log('Erreur génération BA auto: ' + e.message);
		}
	}
	return res;
}

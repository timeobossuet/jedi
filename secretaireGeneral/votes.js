

// Fonctions liées aux votes
function getVotes(reunionId) {
	try {
		var cacheKey = 'votes' + (reunionId ? ('_' + reunionId) : '');
		var cached = fromCache_(cacheKey);
		if (cached) return cached;
		var sheet = getSGSheet_(SG_CONFIG.sheets.votes);
		var rows  = readRows(sheet);
		var votes = [];
		for (var i = 0; i < rows.length; i++) {
			var row = rows[i];
			if (!row[0]) continue;
			if (reunionId && String(row[1]) !== String(reunionId)) continue;
			votes.push({
				id:          String(row[0]),
				reunionId:   String(row[1] || ''),
				date:        row[2] ? fmtDateForSheet(row[2], 'Europe/Paris', 'dd/MM/yyyy') : '',
				sujet:       row[3]  || '',
				type:        row[4]  || 'CA',
				pour:        Number(row[5])  || 0,
				contre:      Number(row[6])  || 0,
				abstentions: Number(row[7])  || 0,
				presents:    Number(row[8])  || 0,
				resultat:    row[9]  || 'En attente'
			});
		}
		toCache_(cacheKey, votes);
		return votes;
	} catch (e) { throw new Error('Erreur chargement votes: ' + e.message); }
}

function saveVote(voteData) {
	try {
		var sheet       = getSGSheet_(SG_CONFIG.sheets.votes);
		var id          = voteData.id || _generateId_();
		var pour        = Number(voteData.pour)         || 0;
		var contre      = Number(voteData.contre)       || 0;
		var abstentions = Number(voteData.abstentions)  || 0;
		var presents    = Number(voteData.presents)     || 0;
		var nbTotal     = Number(voteData.nbTotal)      || 0;
		var type        = voteData.type                  || 'CA';
		var quorumMin    = type === 'CA' || type === 'Bureau'
			? Math.ceil(nbTotal / 2)
			: Math.ceil(nbTotal / 3);
		var quorumAtteint = presents >= quorumMin;
		var votants       = pour + contre;
		var resultat;
		if (type === 'AG_Dissolution') {
			resultat = (quorumAtteint && votants > 0 && pour >= votants * 0.75) ? 'Adopté' : 'Rejeté';
		} else if (!quorumAtteint) {
			resultat = 'Quorum non atteint';
		} else if (pour > contre) {
			resultat = 'Adopté';
		} else if (contre > pour) {
			resultat = 'Rejeté';
		} else {
			resultat = 'Égalité';
		}
		var row = [
			id,
			voteData.reunionId || '',
			voteData.date || fmtDateForSheet(new Date(), 'Europe/Paris', 'yyyy-MM-dd'),
			voteData.sujet || '',
			type,
			pour,
			contre,
			abstentions,
			presents,
			resultat,
			JSON.stringify(voteData.data || {})
		];
		var res = upsertRowById(sheet, id, row);
		return { success: true, id: id, action: res.action };
	} catch (e) { throw new Error('Erreur sauvegarde vote: ' + e.message); }
}

function getJuniorAutofillInfo_() {
  var defaults = {
    nom: 'Junior INSA Services',
    signature: "Junior-Entreprise de l'INSA de Toulouse",
    statut_juridique: 'Association régie par la loi de 1901 - Membre de la Confédération Nationale des Junior-Entreprises',
    'numero-telephone': '05 67 04 88 99',
    president: { prenom: 'Jean-Gabriel', nom: 'Virgone', mail: '', sexe: 'H' },
    vicepresident: { prenom: 'Victor', nom: 'Schalk', mail: '', sexe: 'H' },
    tresorier: { prenom: 'William', nom: 'Woodward', mail: '', sexe: 'H' },
    adresse: '135 Avenue de Rangueil 31077 TOULOUSE Cedex 04',
    siteweb: 'https://juniorinsaservices.fr',
    mail: 'contact@juniorinsaservices.fr',
    siret: '341 760 528 00016',
    ape: '9499Z',
    tva_intracommunautaire: 'FR35 341760528'
  };

  try {
    if (typeof fsJuniorGetProfile === 'function') {
      var profile = fsJuniorGetProfile();
      if (profile && typeof profile === 'object') {
        for (var key in profile) {
          if (profile.hasOwnProperty(key) && profile[key] !== null && profile[key] !== undefined && profile[key] !== '') {
            defaults[key] = profile[key];
          }
        }
      }
    }
  } catch (e) {}

  defaults.president = normalizeJuniorPerson_(defaults.president, { prenom: 'Jean-Gabriel', nom: 'Virgone', sexe: 'H' });
  defaults.vicepresident = normalizeJuniorPerson_(defaults.vicepresident, { prenom: 'Victor', nom: 'Schalk', sexe: 'H' });
  defaults.tresorier = normalizeJuniorPerson_(defaults.tresorier, { prenom: 'William', nom: 'Woodward', sexe: 'H' });
  return defaults;
}

function normalizeJuniorPerson_(value, fallback) {
  var fb = fallback && typeof fallback === 'object' ? fallback : { prenom: '', nom: '', mail: '', sexe: '', portable: '' };
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    return {
      prenom: String(value.prenom || fb.prenom || '').trim(),
      nom: String(value.nom || fb.nom || '').trim(),
      mail: String(value.mail || fb.mail || '').trim(),
      portable: String(value.portable || fallback.portable || '').trim(),
      sexe: String(value.sexe || fb.sexe || '').trim()
    };
  }

  if (typeof value === 'string' && value.trim()) {
    var parts = value.trim().split(/\s+/);
    if (parts.length === 1) return { prenom: '', nom: parts[0], mail: String(fb.mail || '').trim(), sexe: String(fb.sexe || '').trim() };
    return {
      prenom: parts.slice(0, parts.length - 1).join(' '),
      nom: parts[parts.length - 1],
      mail: String(fb.mail || '').trim(),
      sexe: String(fb.sexe || '').trim()
    };
  }

  return {
    prenom: String(fb.prenom || '').trim(),
    nom: String(fb.nom || '').trim(),
    mail: String(fb.mail || '').trim(),
    sexe: String(fb.sexe || '').trim()
  };
}

function joinFullName_(prenom, nom) {
  return [String(prenom || '').trim(), String(nom || '').trim()].filter(function(part) { return !!part; }).join(' ');
}

function buildSgInscriptionMessage_(etudeInfo) {
  var info = etudeInfo && typeof etudeInfo === 'object' ? etudeInfo : {};
  var fullName = joinFullName_(info.intervenantPrenom || '', info.intervenantNom || '');
  var email = String(info.intervenantMail || '').trim();
  var suffix = email ? (', email ' + email) : '';
  return 'Bonjour Secretaire General, merci dinscrire ' + (fullName || 'lintervenant') + suffix + '.';
}

function normalizeEtudeCategories_(raw) {
  var values = Array.isArray(raw) ? raw : typeof raw === 'string' ? raw.split(',') : raw !== null && raw !== undefined ? [String(raw)] : [];
  var seen = {};
  var normalized = [];
  for (var i = 0; i < values.length; i++) {
    var category = String(values[i] || '').trim();
    if (!category || seen[category.toLowerCase()]) continue;
    seen[category.toLowerCase()] = true;
    normalized.push(category);
  }
  return normalized;
}
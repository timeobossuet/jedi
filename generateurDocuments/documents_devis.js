/**
 * Crée un devis à partir du template "devis" dans:
 * 2. [REF] Documents éditables
 * Nom: [year]_[code]_DEVIS
 * Si forceRegenerate=true, supprime l'ancien avant de recréer.
 */
function devisStatusCacheKey_(studyId) {
  return 'DEVIS_STATUS_' + String(studyId);
}

function getCachedDevisStatus_(studyId) {
  try {
    var raw = CacheService.getScriptCache().get(devisStatusCacheKey_(studyId));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function setCachedDevisStatus_(studyId, statusObj) {
  try {
    CacheService.getScriptCache().put(devisStatusCacheKey_(studyId), JSON.stringify(statusObj || {}), 45);
  } catch (e) {}
}

function invalidateDevisStatusCache_(studyId) {
  try {
    CacheService.getScriptCache().remove(devisStatusCacheKey_(studyId));
  } catch (e) {}
}

function buildDefaultDevisStatus_(rawData) {
  var reference = rawData && rawData.reference ? String(rawData.reference) : [rawData.year || 'XXXX', rawData.code || 'XXX'].join('_');
  var devisName = reference + '_DEVIS';
  return {
    hasDraft: false,
    draftUrl: '',
    draftName: devisName,
    hasPdf: false,
    pdfUrl: '',
    hasValide: false,
    valideUrl: '',
    devisName: devisName,
    relectureDemandee: false,
    chatUrl: 'https://chat.google.com/room/AAQApAX9EAQ'
  };
}

function patchDevisStatusFirestore_(studyId, partialStatus) {
  var merged = patchDocWorkflowStatusFirestore_(studyId, 'devis', partialStatus || {});
  var legacy = adaptDevisStatusLegacyShape_(merged);
  setCachedDevisStatus_(studyId, legacy);
  return legacy;
}

function createDevisDoc(studyId, forceRegenerate) {
  var generated = generateWorkflowDocument_(studyId, 'devis', !!forceRegenerate);
  invalidateDevisStatusCache_(studyId);
  return generated;
}

// --- Alias FR ---
function genererDevisEtude(studyId, forceRegenerate) {
  return createDevisDoc(studyId, forceRegenerate || false);
}

/**
 * Vérifie l'état du devis pour une étude :
 * - présence dans 2. [REF] Documents éditables (draft Google Slides + PDF _relecture)
 * - état de signature/archivage (sans déplacement automatique)
 * - état de relecture persisté
 */
function checkDevisStatus(studyId) {
  try {
    var result = adaptDevisStatusLegacyShape_(checkWorkflowDocumentStatus_(studyId, 'devis'));
    setCachedDevisStatus_(studyId, result);
    return result;
  } catch(e) {
    return { hasDraft: false, hasPdf: false, hasValide: false, draftUrl: '', pdfUrl: '', valideUrl: '',
             devisName: '', relectureDemandee: false, chatUrl: 'https://chat.google.com/room/AAQApAX9EAQ', error: e.message };
  }
}

/**
 * Exporte le devis en PDF (nom: YYYY_XXX_DEVIS_relecture.pdf) dans «2. [REF] Documents éditables»
 * puis envoie le lien via le webhook entrant de l'espace Google Chat.
 * Stocke l'état de relecture dans PropertiesService.
 */
function envoyerDevisRelectureChat(studyId, options) {
  var result = sendWorkflowDocumentRelectureChat_(studyId, 'devis', options);
  invalidateDevisStatusCache_(studyId);
  return {
    success: !!(result && result.success),
    chatUrl: result && result.chatUrl ? result.chatUrl : 'https://chat.google.com/room/AAQApAX9EAQ',
    pdfUrl: result && result.pdfUrl ? result.pdfUrl : ''
  };
}

/**
 * Marque le devis comme signé (archivage manuel requis), sans déplacement automatique.
 */
function validerDevisClient(studyId) {
  var result = validateWorkflowDocument_(studyId, 'devis');
  invalidateDevisStatusCache_(studyId);
  return result;
}

/**
 * Crée un événement Google Calendar "Présentation du devis" avec Meet intégré
 * et retourne le lien Meet + lien événement.
 * Nécessite l'API Calendar v3 activée dans appsscript.json.
 */
function creerEvenementPresentationDevis(studyId) {
  var rawData = getStudyData(studyId);
  var code  = rawData.code || '';
  var nom   = rawData.nom  || '';
  var title = 'Présentation du devis — ' + [code, nom].filter(Boolean).join(' ');

  var tz = 'Europe/Paris';
  var startDate = new Date();
  startDate.setDate(startDate.getDate() + 1);
  startDate.setHours(10, 0, 0, 0);
  var endDate = new Date(startDate.getTime() + 60 * 60 * 1000);

  var event = {
    summary: title,
    description: 'Présentation et discussion du devis avec le client.',
    start: { dateTime: startDate.toISOString(), timeZone: tz },
    end:   { dateTime: endDate.toISOString(),   timeZone: tz },
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: 'hangoutsMeet' }
      }
    }
  };

  var created = Calendar.Events.insert(event, 'primary', { conferenceDataVersion: 1 });

  var meetLink = '';
  if (created.conferenceData && created.conferenceData.entryPoints) {
    for (var i = 0; i < created.conferenceData.entryPoints.length; i++) {
      var ep = created.conferenceData.entryPoints[i];
      if (ep.entryPointType === 'video') { meetLink = ep.uri; break; }
    }
  }

  return {
    success: true,
    meetLink: meetLink,
    eventUrl: created.htmlLink || '',
    title: title
  };
}

function _extractClientEmailsForMeeting_(clientInfo) {
  var info = clientInfo && typeof clientInfo === 'object' ? clientInfo : {};
  var out = [];

  function pushEmail_(value) {
    var email = String(value || '').trim().toLowerCase();
    if (!email || email.indexOf('@') < 1) return;
    if (out.indexOf(email) >= 0) return;
    out.push(email);
  }

  if (String(info.type || '').toLowerCase() === 'particulier') {
    var p = info.particulier && typeof info.particulier === 'object' ? info.particulier : {};
    pushEmail_(p.email);
    return out;
  }

  var contacts = Object.prototype.toString.call(info.contacts) === '[object Array]' ? info.contacts : [];
  for (var i = 0; i < contacts.length; i++) {
    var c = contacts[i] && typeof contacts[i] === 'object' ? contacts[i] : {};
    pushEmail_(c.email);
  }

  return out;
}

function _safeIntervenantEmailForMeeting_(etudeInfo) {
  var info = etudeInfo && typeof etudeInfo === 'object' ? etudeInfo : {};
  var email = String(info.intervenantMail || '').trim().toLowerCase();
  return email.indexOf('@') > 0 ? email : '';
}

function _formatMeetingDateLabel_(dateObj) {
  var tz = Session.getScriptTimeZone() || 'Europe/Paris';
  return Utilities.formatDate(dateObj, tz, 'dd/MM/yyyy HH:mm');
}

function _pushUniqueEmail_(out, value) {
  var email = String(value || '').trim().toLowerCase();
  if (!email || email.indexOf('@') < 1) return;
  if (out.indexOf(email) >= 0) return;
  out.push(email);
}

function planifierReunionEtudeCalendar(studyId, options) {
  if (!studyId) throw new Error('studyId manquant.');
  var opts = options && typeof options === 'object' ? options : {};

  var startIso = String(opts.startIso || opts.isoDateTime || '').trim();
  if (!startIso) throw new Error('Date de reunion manquante.');

  var startDate = new Date(startIso);
  if (!(startDate instanceof Date) || isNaN(startDate.getTime())) {
    throw new Error('Date de reunion invalide.');
  }

  var duration = Math.max(15, parseInt(opts.durationMinutes, 10) || 60);
  var endDate = new Date(startDate.getTime() + duration * 60 * 1000);
  var withMeet = opts.withMeet !== false;
  var includeClient = opts.includeClient !== false;
  var includeIntervenant = !!opts.includeIntervenant;
  var location = String(opts.location || '').trim();
  var notes = String(opts.notes || '').trim();
  var tz = 'Europe/Paris';

  var rawData = getStudyData(studyId) || {};
  var etudeInfo = rawData.EtudeInfo && typeof rawData.EtudeInfo === 'object' ? rawData.EtudeInfo : {};
  var clientInfo = rawData.ClientInfo && typeof rawData.ClientInfo === 'object' ? rawData.ClientInfo : {};
  var ref = String(etudeInfo.reference || rawData.reference || [rawData.year || '', rawData.code || ''].filter(Boolean).join('_')).trim();
  var nomEtude = String(rawData.nom || etudeInfo.nom || '').trim();

  var clientEmails = includeClient ? _extractClientEmailsForMeeting_(clientInfo) : [];
  var intervenantEmail = includeIntervenant ? _safeIntervenantEmailForMeeting_(etudeInfo) : '';
  var attendees = [];
  for (var i = 0; i < clientEmails.length; i++) {
    _pushUniqueEmail_(attendees, clientEmails[i]);
  }
  if (intervenantEmail) _pushUniqueEmail_(attendees, intervenantEmail);

  var defaultTitle = 'Reunion etude ' + (ref || studyId) + (nomEtude ? (' - ' + nomEtude) : '');
  var title = String(opts.title || defaultTitle).trim();

  var descriptionLines = [];
  descriptionLines.push(withMeet ? 'Reunion de suivi (visio Google Meet).' : 'Reunion de suivi (physique).');
  if (notes) descriptionLines.push('Notes: ' + notes);

  var event = {
    summary: title,
    description: descriptionLines.join('\n'),
    start: { dateTime: startDate.toISOString(), timeZone: tz },
    end: { dateTime: endDate.toISOString(), timeZone: tz }
  };

  if (!withMeet && location) {
    event.location = location;
  }

  if (attendees.length) {
    event.attendees = attendees.map(function(email) { return { email: email }; });
  }

  if (withMeet) {
    event.conferenceData = {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: 'hangoutsMeet' }
      }
    };
  }

  var created = Calendar.Events.insert(event, 'primary', withMeet ? { conferenceDataVersion: 1 } : {});
  var meetLink = '';
  if (withMeet && created.conferenceData && created.conferenceData.entryPoints) {
    for (var j = 0; j < created.conferenceData.entryPoints.length; j++) {
      var ep = created.conferenceData.entryPoints[j];
      if (ep.entryPointType === 'video') {
        meetLink = ep.uri;
        break;
      }
    }
  }

  return {
    success: true,
    eventUrl: created.htmlLink || '',
    meetLink: meetLink,
    scheduledAt: _formatMeetingDateLabel_(startDate),
    startIso: startDate.toISOString(),
    durationMinutes: duration,
    withMeet: withMeet,
    location: withMeet ? '' : location,
    notes: notes,
    clientEmails: clientEmails,
    intervenantEmail: intervenantEmail,
    attendees: attendees
  };
}

function planifierReunionClientIntervenantEtude(studyId, isoDateTime, durationMinutes) {
  var result = planifierReunionEtudeCalendar(studyId, {
    startIso: isoDateTime,
    durationMinutes: durationMinutes,
    withMeet: true,
    includeClient: true,
    includeIntervenant: true,
    notes: 'Reunion de suivi etude (client + intervenant).'
  });

  return {
    success: true,
    eventUrl: result.eventUrl || '',
    meetLink: result.meetLink || '',
    scheduledAt: result.scheduledAt || '',
    clientEmailSent: false,
    intervenantEmailSent: false,
    clientEmails: result.clientEmails || [],
    intervenantEmail: result.intervenantEmail || '',
    sendErrors: []
  };
}

function applyLiteralReplacements(presentation, replacements) {
  function buildTokenVariants_(token) {
    var t = String(token || '');
    if (!t) return [];
    if (/^\{\{[^{}]+\}\}$/.test(t)) {
      var innerD = t.substring(2, t.length - 2);
      return [t, '{' + innerD + '}'];
    }
    if (/^\{[^{}]+\}$/.test(t)) {
      var inner = t.substring(1, t.length - 1);
      return ['{{' + inner + '}}', t];
    }
    return [t];
  }

  for (var key in replacements) {
    if (!replacements.hasOwnProperty(key)) continue;
    var replacement = replacements[key] === null || typeof replacements[key] === 'undefined'
      ? ''
      : String(replacements[key]);

    var tokenVariants = buildTokenVariants_(key);

    if (isLikelyHtmlContent_(replacement)) {
      var replacedRich = false;
      if (typeof replaceHtmlTokenInPresentation_ === 'function') {
        for (var v = 0; v < tokenVariants.length; v++) {
          replacedRich = replaceHtmlTokenInPresentation_(presentation, tokenVariants[v], replacement) || replacedRich;
        }
      }

      if (!replacedRich) {
        var plain = typeof convertHtmlToPlainText === 'function' ? convertHtmlToPlainText(replacement) : replacement;
        for (var p = 0; p < tokenVariants.length; p++) {
          presentation.replaceAllText(tokenVariants[p], plain);
        }
      }
      continue;
    }

    for (var n = 0; n < tokenVariants.length; n++) {
      presentation.replaceAllText(tokenVariants[n], replacement);
    }
  }
}

function forEachPresentationTextRange(presentation, callback) {
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    var elements = slides[i].getPageElements();
    for (var j = 0; j < elements.length; j++) {
      var type = elements[j].getPageElementType();
      if (type === SlidesApp.PageElementType.SHAPE) {
        try {
          callback(elements[j].asShape().getText(), slides[i], elements[j]);
        } catch (e) {}
      } else if (type === SlidesApp.PageElementType.TABLE) {
        try {
          var table = elements[j].asTable();
          for (var r = 0; r < table.getNumRows(); r++) {
            for (var c = 0; c < table.getNumColumns(); c++) {
              callback(table.getCell(r, c).getText(), slides[i], elements[j]);
            }
          }
        } catch (e2) {}
      }
    }
  }
}

function replaceRegexAcrossPresentation(presentation, regex, replacer) {
  forEachPresentationTextRange(presentation, function(textRange) {
    var source = textRange.asString();
    var next = source.replace(regex, replacer);
    if (next !== source) {
      textRange.setText(next);
    }
  });
}

function normalizeTypographicQuotes(value) {
  return String(value || '').replace(/[’‘]/g, "'").replace(/[“”]/g, '"');
}

function replaceDateTokensInPresentation(presentation, dateObj) {
  var formatted = Utilities.formatDate(dateObj || new Date(), 'Europe/Paris', 'dd/MM/yyyy');
  var regex = /\{\s*today\s*\|\s*formatDate\s*:\s*['’][^'’]+['’]\s*\}/gi;
  replaceRegexAcrossPresentation(presentation, regex, function() { return formatted; });
}

function resolveGenderVariant(options, sexe) {
  var s = String(sexe || '').toLowerCase();
  var female = s === 'femme' || s === 'f';
  var male = s === 'homme' || s === 'h';

  if (options.length >= 3) {
    if (female) return options[1];
    if (male) return options[2];
    return options[0];
  }

  if (options.length >= 2) {
    if (female) return options[0];
    if (male) return options[1];
    return options[0];
  }

  return options.length ? options[0] : '';
}

function replaceGenderTokensInPresentation(presentation, chargeSexe) {
  replaceRegexAcrossPresentation(
    presentation,
    /\{\s*suiveur\s*\|\s*sexe\s*:\s*([^}]*)\}/gi,
    function(fullMatch, inside) {
      var normalized = normalizeTypographicQuotes(inside);
      var parts = [];
      var re = /'([^']*)'/g;
      var m;
      while ((m = re.exec(normalized)) !== null) {
        parts.push(m[1]);
      }
      if (!parts.length) return fullMatch;
      return resolveGenderVariant(parts, chargeSexe);
    }
  );
}

function replaceFinancialTokensInPresentation(presentation, values) {
  var replacements = [
    { regex: /\{\s*\(\s*etude\.prix\s*-\s*etude\.frais\s*\)\s*\|\s*decimales:2\s*\}/gi, value: values.prixHTHorsFrais - values.frais },
    { regex: /\{\s*\(\s*etude\.prix\s*\*\s*20\s*\/\s*100\s*\)\s*\|\s*decimales:2\s*\}/gi, value: values.prixHTHorsFrais * 0.2 },
    { regex: /\{\s*\(?\s*etude\.prix\s*\*\s*\(\s*1\s*\+\s*20\s*\/\s*100\s*\)\s*\)?\s*\|\s*decimales:2\s*\}/gi, value: values.prixHTHorsFrais * 1.2 },
    { regex: /\{\s*etude\.frais\s*\|\s*decimales:2\s*\}/gi, value: values.frais },
    { regex: /\{\s*etude\.prix[-–]?total\s*\|\s*decimales:2\s*\}/gi, value: values.prixTotalHT },
    { regex: /\{\s*etudes\.prix[-–]?total\s*\|\s*decimales:2\s*\}/gi, value: values.prixTotalHT },
    { regex: /\{\s*etude\.prix[-–]?total[-–]?tva\s*\|\s*decimales:2\s*\}/gi, value: values.tvaAmount },
    { regex: /\{\s*etude\.prix[-–]?total\s*\*\s*20\s*\/\s*100\s*\|\s*decimales:2\s*\}/gi, value: values.tvaAmount },
    { regex: /\{\s*etude\.prix[-–]?total[-–]?ttc\s*\|\s*decimales:2\s*\}/gi, value: values.prixTotalTTC },
    { regex: /\{\s*\(\s*etude\.prix[-–]?total\s*\*\s*0\.3\s*\)\s*\|\s*decimales:2\s*\}/gi, value: values.acompteHT },
    { regex: /\{\s*\(\s*etude\.prix[-–]?total\s*\*\s*0\.7\s*\)\s*\|\s*decimales:2\s*\}/gi, value: values.soldeHT },
    { regex: /\{\s*etude\.prix[-–]?ttc\s*\*\s*0\.3\s*\|\s*decimales:2\s*\}/gi, value: values.acompteTTC },
    { regex: /\{\s*etude\.prix[-–]?ttc\s*\*\s*0\.7\s*\|\s*decimales:2\s*\}/gi, value: values.soldeTTC },
    { regex: /\{\s*phase\.prix-jeh\s*\|\s*decimales:2\s*\}/gi, value: null }
  ];

  for (var i = 0; i < replacements.length; i++) {
    if (replacements[i].value !== null) {
      replaceRegexAcrossPresentation(presentation, replacements[i].regex, formatNumber(replacements[i].value));
    }
  }
}

function replaceLetterTokensInPresentation(presentation, values) {
  var map = [
    { regex: /\{\s*etude\.frais\s*\|\s*lettres\s*\}/gi, value: values.frais },
    { regex: /\{\s*etude\.prix\s*\|\s*lettres\s*\}/gi, value: values.prixHTHorsFrais },
    { regex: /\{\s*etude\.prix[-–]?total\s*\|\s*lettres\s*\}/gi, value: values.prixTotalHT },
    { regex: /\{\s*etude\.prix[-–]?total[-–]?ttc\s*\|\s*lettres\s*\}/gi, value: values.prixTotalTTC },
    { regex: /\{\s*\(?\s*etude\.prix\s*\*\s*\(\s*1\s*\+\s*20\s*\/\s*100\s*\)\s*\)?\s*\|\s*lettres\s*\}/gi, value: values.prixHTHorsFrais * 1.2 }
  ];

  for (var i = 0; i < map.length; i++) {
    var textValue = numberToFrenchMoneyWords(map[i].value);
    replaceRegexAcrossPresentation(presentation, map[i].regex, textValue);
  }
}

function numberToFrenchMoneyWords(num) {
  var value = Number(num || 0);
  var euros = Math.floor(Math.abs(value));
  var cents = Math.round((Math.abs(value) - euros) * 100);

  var euroPart = numberToFrenchWords(euros) + (euros > 1 ? ' euros' : ' euro');
  if (cents > 0) {
    return euroPart + ' et ' + numberToFrenchWords(cents) + (cents > 1 ? ' centimes' : ' centime');
  }
  return euroPart;
}

function numberToFrenchWords(num) {
  num = Math.floor(Number(num || 0));
  if (num === 0) return 'zéro';

  var units = ['', 'un', 'deux', 'trois', 'quatre', 'cinq', 'six', 'sept', 'huit', 'neuf'];
  var teens = ['dix', 'onze', 'douze', 'treize', 'quatorze', 'quinze', 'seize', 'dix-sept', 'dix-huit', 'dix-neuf'];
  var tens = ['', '', 'vingt', 'trente', 'quarante', 'cinquante', 'soixante'];

  function underHundred(n) {
    if (n < 10) return units[n];
    if (n < 20) return teens[n - 10];
    if (n < 70) {
      var t = Math.floor(n / 10);
      var u = n % 10;
      if (u === 1) return tens[t] + ' et un';
      return tens[t] + (u ? '-' + units[u] : '');
    }
    if (n < 80) {
      if (n === 71) return 'soixante et onze';
      return 'soixante-' + underHundred(n - 60);
    }
    if (n === 80) return 'quatre-vingts';
    return 'quatre-vingt-' + underHundred(n - 80);
  }

  function underThousand(n) {
    if (n < 100) return underHundred(n);
    var c = Math.floor(n / 100);
    var r = n % 100;
    var prefix = c === 1 ? 'cent' : units[c] + ' cent';
    if (r === 0 && c > 1) prefix += 's';
    return r ? prefix + ' ' + underHundred(r) : prefix;
  }

  var parts = [];
  var millions = Math.floor(num / 1000000);
  var thousands = Math.floor((num % 1000000) / 1000);
  var rest = num % 1000;

  if (millions) {
    parts.push((millions === 1 ? 'un million' : underThousand(millions) + ' millions'));
  }
  if (thousands) {
    parts.push(thousands === 1 ? 'mille' : underThousand(thousands) + ' mille');
  }
  if (rest) {
    parts.push(underThousand(rest));
  }

  return parts.join(' ');
}

function handleSlidesPhaseMethodology(presentation, phases) {
    var slides = presentation.getSlides();
    var templateSlide = null;
    
    // Find the slide containing the phase marker
    for (var i = 0; i < slides.length; i++) {
        var elements = slides[i].getPageElements();
        for (var j = 0; j < elements.length; j++) {
            if (elements[j].getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                var shape = elements[j].asShape();
                if (shape.getText().asString().indexOf("-Phase {phase.numero}") !== -1) {
                    templateSlide = slides[i];
                    break;
                }
            }
        }
        if (templateSlide) break;
    }
    
    if (!templateSlide) return;
    
    // Process phases
    for (var i = 0; i < phases.length; i++) {
        var phase = phases[i];
        var currentSlide;
        
        if (i === 0) {
            currentSlide = templateSlide;
        } else {
            currentSlide = templateSlide.duplicate();
            // Move duplicated slide to after the previous one
            // Apps Script Duplicate acts usually puts it at end or after?
            // "The duplicate is created at the end of the presentation."
            // We want it contiguous.
            // Move logic: presentation.getSlides()[index].move(newIndex)
            // But let's just leave at end or optimize later. To be safe, we leave it.
            // Actually, methodology pages should probably be sequential.
            // Let's assume user accepts them appended or check indices.
        }
        
        // Replacements on this specific slide
        // SlidesApp does not allow replaceAllText on a single slide easily (it's global or shape based).
        // `slide.replaceAllText(findText, replaceText)` ensures scoping to slide.
        
        var pNum = (i + 1).toString();
        var pName = phase.title || 'Phase ' + pNum;
        var pJeh = (phase.jeh || 0).toString();
        var pDesc = phase.description || '';
        
        currentSlide.replaceAllText("-Phase {phase.numero}", "-Phase " + pNum);
        currentSlide.replaceAllText("-Phase {{phase.numero}}", "-Phase " + pNum);
        currentSlide.replaceAllText("{phase.nom}", pName);
        currentSlide.replaceAllText("{{phase.nom}}", pName);
        currentSlide.replaceAllText("{phase.nombre-jeh}", pJeh);
        currentSlide.replaceAllText("{{phase.nombre-jeh}}", pJeh);
        if (isLikelyHtmlContent_(pDesc)) {
          var richApplied = false;
          if (typeof replaceHtmlTokenInSlide_ === 'function') {
            richApplied = replaceHtmlTokenInSlide_(currentSlide, '{phase.description}', pDesc);
            richApplied = replaceHtmlTokenInSlide_(currentSlide, '{{phase.description}}', pDesc) || richApplied;
          }
          if (!richApplied) {
            var pDescPlain = typeof convertHtmlToPlainText === 'function' ? convertHtmlToPlainText(pDesc) : String(pDesc || '');
            currentSlide.replaceAllText('{phase.description}', pDescPlain);
            currentSlide.replaceAllText('{{phase.description}}', pDescPlain);
          }
        } else {
          currentSlide.replaceAllText('{phase.description}', String(pDesc || ''));
          currentSlide.replaceAllText('{{phase.description}}', String(pDesc || ''));
        }

        // Force phase title in white (template issue: text may keep red style)
        try {
          var els = currentSlide.getPageElements();
          for (var k = 0; k < els.length; k++) {
            if (els[k].getPageElementType() === SlidesApp.PageElementType.SHAPE) {
              var shp = els[k].asShape();
              var t = shp.getText().asString();
              if (t.indexOf('-Phase ') !== -1) {
                shp.getText().getTextStyle().setForegroundColor('#FFFFFF');
              }
            }
          }
        } catch (e) {}
    }
}

function handleSlidesPhasePricingTable(presentation, phases) {
    var slides = presentation.getSlides();
    var targetSlide = null;
    var targetTable = null;
    var rowTemplateIndex = -1;
    
    for (var i = 0; i < slides.length; i++) {
        var elements = slides[i].getPageElements();
        for (var j = 0; j < elements.length; j++) {
            if (elements[j].getPageElementType() === SlidesApp.PageElementType.TABLE) {
                var table;
                try {
                    table = elements[j].asTable();
                } catch (e) {
                    continue; // Skip if casting fails for some reason
                }
                
                for (var r = 0; r < table.getNumRows(); r++) {
                     var rowText = "";
                     for (var c=0; c<table.getNumColumns(); c++) {
                     rowText += safeGetTableCellText(table, r, c);
                     }
                     if (rowText.indexOf("{phases}") !== -1) {
                         targetTable = table; rowTemplateIndex = r; break;
                     }
                }
            }
            if (targetTable) break;
        }
        if (targetTable) break;
    }
    
    if (!targetTable) return;
    
    // 1. Capture Template Logic WITHOUT modifying the table yet
    // Read the cell texts and clean markers in memory
    var templateCells = [];
    for (var c = 0; c < targetTable.getNumColumns(); c++) {
      var rawText = safeGetTableCellText(targetTable, rowTemplateIndex, c);
        rawText = rawText.replace("{phases}", "").replace("{/phases}", "");
        templateCells.push(rawText); 
    }
    
    // 2. Expand Table
    // Iterate phases
    for (var i = 0; i < phases.length; i++) {
        var phase = phases[i];
        var currentIndex = rowTemplateIndex + i;
        
        if (i > 0) {
            targetTable.insertRow(currentIndex);
        }
        
        // Fill Cells
        for (var c = 0; c < targetTable.getNumColumns(); c++) {
            var cellContent = templateCells[c];
            
            // Calc values
            var pNum = (i + 1).toString();
            var pName = phase.title || '';
            var pJeh = (phase.jeh || 0).toString();
            var pPriceJeh = parseFloat(phase.priceJeh) || 0;
            var pTotal = pPriceJeh * (parseFloat(phase.jeh) || 0);

            // Replacements
            cellContent = cellContent.replace(/{numero}/g, pNum);
            cellContent = cellContent.replace(/{nom}/g, pName);
            cellContent = cellContent.replace(/{phase.nombre-jeh}/g, pJeh);
            cellContent = cellContent.replace(/\{\s*phase\.prix-jeh\s*\|\s*decimales:2\s*\}/g, formatNumber(pPriceJeh));
            
            // Complex
            cellContent = cellContent.replace(/\{\s*\(\s*nombre_jeh\s*\*\s*prix_jeh\)\s*\|\s*decimales:2\s*\}/g, formatNumber(pTotal));
            
            // Update Cell
            safeSetTableCellText(targetTable, currentIndex, c, cellContent);
        }
    }
}

      function handleSlidesPlanningTable(presentation, phases) {
        var slides = presentation.getSlides();
        var targetTable = null;

        for (var i = 0; i < slides.length; i++) {
          var elements = slides[i].getPageElements();
          for (var j = 0; j < elements.length; j++) {
            if (elements[j].getPageElementType() !== SlidesApp.PageElementType.TABLE) continue;
            var table = elements[j].asTable();
            var found = false;
            for (var r = 0; r < table.getNumRows() && !found; r++) {
              for (var c = 0; c < table.getNumColumns() && !found; c++) {
                var txt = safeGetTableCellText(table, r, c);
                if (txt.indexOf('{phases.planning}') !== -1) {
                  found = true;
                  targetTable = table;
                }
              }
            }
            if (targetTable) break;
          }
          if (targetTable) break;
        }

        // Fallback: replace placeholder by textual planning if no table contains it
        if (!targetTable) {
          var planningText = generatePlanningText(phases);
          replaceRegexAcrossPresentation(presentation, /\{\s*phases\.planning\s*\}/gi, planningText);
          return;
        }

        var maxWeek = 1;
        for (var p = 0; p < phases.length; p++) {
          var s = Math.max(1, parseInt(phases[p].start, 10) || 1);
          var d = Math.max(1, parseInt(phases[p].duration, 10) || 1);
          maxWeek = Math.max(maxWeek, s + d - 1);
        }

        var requiredRows = 1 + phases.length;
        var requiredCols = 1 + maxWeek;

        while (targetTable.getNumRows() < requiredRows) {
          targetTable.appendRow();
        }
        while (targetTable.getNumColumns() < requiredCols) {
          targetTable.appendColumn();
        }

        // Clear all existing cells
        for (var rr = 0; rr < targetTable.getNumRows(); rr++) {
          for (var cc = 0; cc < targetTable.getNumColumns(); cc++) {
            safeSetTableCellText(targetTable, rr, cc, '');
          }
        }

        // Header
        safeSetTableCellText(targetTable, 0, 0, 'Semaine');
        for (var w = 1; w <= maxWeek; w++) {
          safeSetTableCellText(targetTable, 0, w, String(w));
        }

        // Rows
        for (var idx = 0; idx < phases.length; idx++) {
          var phase = phases[idx] || {};
          var start = Math.max(1, parseInt(phase.start, 10) || 1);
          var duration = Math.max(1, parseInt(phase.duration, 10) || 1);
          var end = start + duration - 1;

          safeSetTableCellText(targetTable, idx + 1, 0, (idx + 1) + '. ' + (phase.title || ('Phase ' + (idx + 1))));
          for (var ww = 1; ww <= maxWeek; ww++) {
            var mark = (ww >= start && ww <= end) ? '■' : '';
            safeSetTableCellText(targetTable, idx + 1, ww, mark);
          }
        }

        // Remove placeholder if it still exists elsewhere
        replaceRegexAcrossPresentation(presentation, /\{\s*phases\.planning\s*\}/gi, '');
      }

      function generatePlanningText(phases) {
        if (!phases || !phases.length) return '';
        var maxWeek = 1;
        for (var i = 0; i < phases.length; i++) {
          var s = Math.max(1, parseInt(phases[i].start, 10) || 1);
          var d = Math.max(1, parseInt(phases[i].duration, 10) || 1);
          maxWeek = Math.max(maxWeek, s + d - 1);
        }

        var lines = [];
        var header = ['Semaine'];
        for (var w = 1; w <= maxWeek; w++) header.push(String(w));
        lines.push(header.join(' | '));

        for (var p = 0; p < phases.length; p++) {
          var phase = phases[p] || {};
          var start = Math.max(1, parseInt(phase.start, 10) || 1);
          var duration = Math.max(1, parseInt(phase.duration, 10) || 1);
          var end = start + duration - 1;
          var row = [(p + 1) + '. ' + (phase.title || ('Phase ' + (p + 1)))];
          for (var ww = 1; ww <= maxWeek; ww++) {
            row.push((ww >= start && ww <= end) ? '■' : '');
          }
          lines.push(row.join(' | '));
        }
        return lines.join('\n');
      }

function safeGetTableCellText(table, row, col) {
  try {
    return table.getCell(row, col).getText().asString();
  } catch (e) {
    return '';
  }
}

function safeSetTableCellText(table, row, col, value) {
  try {
    table.getCell(row, col).getText().setText(value);
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    // Ignore writes on merged cells that are not top-left anchor cells
    if (msg.indexOf('cellules fusionnées') !== -1 || msg.indexOf('merged') !== -1) {
      return;
    }
    throw e;
  }
}

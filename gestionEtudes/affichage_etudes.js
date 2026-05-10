function doGet(e) {
  // Sert le fichier HTML lors de l'accès à l'application web
  var params = e && e.parameter ? e.parameter : {};
  var page = (params.page || '').toLowerCase();
  var templateCandidates;

  switch(page) {
    case 'sg':
      templateCandidates = ['secretaireGeneral/dashboard_sg', 'secretaireGeneral_dashboard_sg', 'dashboard_sg'];
    break;
    case 'etudes':
      templateCandidates = ['gestionEtudes/dashboard_etudes', 'gestionEtudes_dashboard_etudes', 'dashboard_etudes'];
    break;
    default:
      templateCandidates = ['gestionEtudes/dashboard_etudes', 'gestionEtudes_dashboard_etudes', 'dashboard_etudes'];
  }

  try {
    var template = createTemplateWithFallback(expandTemplateCandidates(templateCandidates));
    template.studyId = params.id || '';
    template.navUrlSg = buildAppPageUrl_('sg');
    template.navUrlEtudes = buildAppPageUrl_('etudes');
    template.dsiChatUrl = (typeof getDsiChatRoomUrl_ === 'function')
      ? getDsiChatRoomUrl_()
      : 'https://chat.google.com/room/AAQAMCtOHjs';
    template.dsiTicketsUrl = (typeof getDsiTicketsRoomUrl_ === 'function')
      ? getDsiTicketsRoomUrl_()
      : template.dsiChatUrl;

    return template
      .evaluate()
      .setTitle('Jedi ERP')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    Logger.log('Erreur doGet: ' + (err && err.stack ? err.stack : err));
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Jedi - Erreur</title></head><body style="font-family:Arial,sans-serif;padding:24px"><h2>Erreur de chargement du dashboard</h2><p>' +
      escapeHtml(err && err.message ? err.message : String(err)) +
      '</p><p>Ouvre Executions dans Apps Script pour la stack complete.</p></body></html>'
    );
  }
}

function include(filename) {
  try {
    return createHtmlOutputWithFallback(expandIncludeCandidates(filename)).getContent();
  } catch (err) {
    Logger.log('Include introuvable pour "' + filename + '": ' + (err && err.message ? err.message : err));
    return '<!-- include manquant: ' + escapeHtml(filename) + ' -->';
  }
}

function buildAppPageUrl_(page) {
  var query = 'page=' + encodeURIComponent(String(page || '').toLowerCase());
  var baseUrl = '';
  try {
    baseUrl = ScriptApp.getService().getUrl() || '';
  } catch (e) {
    baseUrl = '';
  }

  if (baseUrl) {
    return baseUrl + (baseUrl.indexOf('?') === -1 ? '?' : '&') + query;
  }

  return '?' + query;
}

function expandTemplateCandidates(names) {
  var out = [];
  for (var i = 0; i < names.length; i++) {
    var name = String(names[i] || '').trim();
    if (!name) continue;
    pushUnique_(out, name);

    if (name.indexOf('/') !== -1) {
      pushUnique_(out, name.replace(/\//g, '_'));
    }
  }
  return out;
}

function expandIncludeCandidates(filename) {
  var name = String(filename || '').trim();
  var out = [];
  if (!name) return out;

  if (name.indexOf('/') !== -1) {
    var parts = name.split('/');
    var baseName = parts[parts.length - 1] || '';

    pushUnique_(out, name);
    pushUnique_(out, name.replace(/\//g, '_'));
    if (baseName) pushUnique_(out, baseName);
    return out;
  }

  // Prefer namespaced files to avoid loading stale root-level homonyms.
  pushUnique_(out, 'gestionEtudes/' + name);
  pushUnique_(out, 'gestionEtudes_' + name);
  pushUnique_(out, 'secretaireGeneral/' + name);
  pushUnique_(out, 'secretaireGeneral_' + name);
  pushUnique_(out, 'generateurDocuments/' + name);
  pushUnique_(out, 'generateurDocuments_' + name);
  pushUnique_(out, name);

  return out;
}

function pushUnique_(arr, value) {
  if (arr.indexOf(value) === -1) arr.push(value);
}

function createTemplateWithFallback(names) {
  var lastError = null;
  for (var i = 0; i < names.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(names[i]);
    } catch (err) {
      lastError = err;
    }
  }
  throw lastError || new Error('Template HTML introuvable.');
}

function createHtmlOutputWithFallback(names) {
  var lastError = null;
  for (var i = 0; i < names.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(names[i]);
    } catch (err) {
      lastError = err;
    }
  }
  throw lastError || new Error('Fichier HTML introuvable.');
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

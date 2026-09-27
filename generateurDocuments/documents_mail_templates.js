function getProjectHtmlFileContent_(fileName) {
  if (!fileName || !String(fileName).trim()) {
    throw new Error('Nom de fichier HTML manquant.');
  }

  var candidates = buildProjectHtmlFileCandidates_(String(fileName).trim());
  for (var i = 0; i < candidates.length; i++) {
    try {
      var output = HtmlService.createHtmlOutputFromFile(candidates[i]);
      return String(output.getContent() || '');
    } catch (err) {}
  }

  throw new Error(
    'Fichier HTML introuvable dans le projet: ' + fileName + '.html' +
    (candidates.length ? ' (noms testés: ' + candidates.join(', ') + ')' : '')
  );
}

function buildProjectHtmlFileCandidates_(fileName) {
  var name = String(fileName || '').trim();
  if (!name) return [];

  name = name.replace(/\.html$/i, '');
  var out = [];
  function addCandidate_(value) {
    var v = String(value || '').trim();
    if (v && out.indexOf(v) === -1) out.push(v);
  }

  addCandidate_(name);
  addCandidate_(name.replace(/\//g, '_'));

  if (name.indexOf('/') !== -1) {
    var parts = name.split('/');
    addCandidate_(parts[parts.length - 1] || '');
  } else {
    addCandidate_('gestionEtudes/templates/' + name);
    addCandidate_('gestionEtudes_templates_' + name);
    addCandidate_('templates/' + name);
    addCandidate_('templates_' + name);
    addCandidate_('gestionEtudes/' + name);
    addCandidate_('gestionEtudes_' + name);
  }

  return out;
}

function getMailTemplateText_(templateFile) {
  if (!templateFile) throw new Error('Template mail introuvable.');

  var mimeType = templateFile.getMimeType();
  if (mimeType === MimeType.GOOGLE_DOCS) {
    var doc = DocumentApp.openById(templateFile.getId());
    var body = getDocumentBodySafe_(doc);
    return (body.getText() || '').trim();
  }

  if (mimeType === 'application/vnd.google-apps.mail-layout') {
    var plain = exportGoogleWorkspaceFileAsText_(templateFile.getId(), 'text/plain');
    if (plain && String(plain).trim()) return String(plain).trim();
    var html = exportGoogleWorkspaceFileAsText_(templateFile.getId(), 'text/html');
    if (html && String(html).trim()) return htmlToText_(String(html));
    throw new Error('Impossible d\'extraire le contenu de la template mail-layout.');
  }

  if (mimeType && mimeType.indexOf('text/') === 0) {
    return String(templateFile.getBlob().getDataAsString() || '').trim();
  }

  throw new Error('Type de template mail non supporté: ' + mimeType + '. Utilisez Google Docs ou un fichier texte.');
}

function exportGoogleWorkspaceFileAsText_(fileId, exportMimeType) {
  try {
    var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/export?mimeType=' + encodeURIComponent(exportMimeType);
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    return response.getResponseCode() >= 200 && response.getResponseCode() < 300
      ? response.getContentText()
      : '';
  } catch (e) {
    return '';
  }
}

function htmlToText_(html) {
  var text = String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\s*\/p\s*>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
  return text.replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

function extractMailSubjectAndBody_(renderedText, fallbackSubject) {
  var text = String(renderedText || '').replace(/\r\n/g, '\n');
  var lines = text.split('\n');
  var subject = '';
  var subjectLineIndex = -1;

  for (var i = 0; i < lines.length; i++) {
    var trimmed = String(lines[i] || '').trim();
    if (!trimmed) continue;
    var match = trimmed.match(/^(objet|subject)\s*:\s*(.+)$/i);
    if (match) {
      subject = String(match[2] || '').trim();
      subjectLineIndex = i;
    }
    break;
  }

  if (subjectLineIndex >= 0) lines.splice(subjectLineIndex, 1);
  return {
    subject: subject || String(fallbackSubject || ''),
    body: lines.join('\n').replace(/^\s+|\s+$/g, '')
  };
}

function extractMailSubjectAndBodyFromHtml_(renderedHtml, fallbackSubject) {
  var html = String(renderedHtml || '');
  var subject = '';
  var commentMatch = html.match(/<!--\s*(?:objet|subject)\s*:\s*([^>]+?)\s*-->/i);
  if (commentMatch && commentMatch[1]) {
    subject = String(commentMatch[1]).trim();
    html = html.replace(commentMatch[0], '');
  }

  if (!subject) {
    var h1Match = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
    if (h1Match && h1Match[1]) subject = htmlToText_(h1Match[1]).trim();
  }

  return { subject: subject || String(fallbackSubject || ''), bodyHtml: html };
}
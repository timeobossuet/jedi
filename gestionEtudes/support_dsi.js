var DSI_CHAT_SPACE_ID_ = 'AAQAMCtOHjs';
var DSI_CHAT_ROOM_URL_ = 'https://chat.google.com/room/' + DSI_CHAT_SPACE_ID_;
var DSI_CHAT_TICKETS_URL_ = DSI_CHAT_ROOM_URL_;
var DSI_CHAT_DEFAULT_THREAD_KEY_ = 'jedi-erp-incidents';
var SG_INTERVENANT_CHAT_SPACE_ID_ = 'AAQAI-FHX64';
var SG_INTERVENANT_CHAT_ROOM_URL_ = 'https://chat.google.com/room/' + SG_INTERVENANT_CHAT_SPACE_ID_;
var SG_INTERVENANT_CHAT_DEFAULT_THREAD_KEY_ = 'jedi-inscriptions-intervenants';
var SG_CONFIG_WEBHOOK_KEY_ = 'google_chat_webhook';

function getDsiChatRoomUrl_() {
  return DSI_CHAT_ROOM_URL_;
}

function getDsiTicketsRoomUrl_() {
  return DSI_CHAT_TICKETS_URL_;
}

function getDsiSupportLinks() {
  return {
    chatUrl: DSI_CHAT_ROOM_URL_,
    ticketsUrl: DSI_CHAT_TICKETS_URL_
  };
}

function signalerErreurDsi(payload) {
  try {
    var ticket = buildDsiTicketPayload_(payload || {});
    var sendResult = sendDsiTicketToChat_(ticket);

    return {
      success: true,
      ticketId: ticket.ticketId,
      chatUrl: DSI_CHAT_ROOM_URL_,
      ticketsUrl: DSI_CHAT_TICKETS_URL_,
      sentVia: sendResult.via || '',
      responseCode: sendResult.responseCode || 0,
      messageName: sendResult.messageName || ''
    };
  } catch (err) {
    var errMessage = err && err.message ? err.message : String(err || 'Erreur inconnue');
    Logger.log('signalerErreurDsi failed: ' + errMessage);

    return {
      success: false,
      error: errMessage,
      chatUrl: DSI_CHAT_ROOM_URL_,
      ticketsUrl: DSI_CHAT_TICKETS_URL_
    };
  }
}

function envoyerInfosIntervenantSgChat(studyId, intervenantPayload) {
  try {
    var payload = intervenantPayload && typeof intervenantPayload === 'object' ? intervenantPayload : {};
    var rawData = {};
    try {
      rawData = studyId ? (getStudyData(studyId) || {}) : {};
    } catch (e) {
      rawData = {};
    }

    var etudeInfo = rawData.EtudeInfo && typeof rawData.EtudeInfo === 'object' ? rawData.EtudeInfo : {};
    var ref = sanitizeDsiField_(
      payload.reference ||
      etudeInfo.reference ||
      rawData.reference ||
      [rawData.year || '', rawData.code || ''].filter(function(v) { return !!String(v || '').trim(); }).join('_'),
      80
    );
    var etudeNom = sanitizeDsiField_(payload.etudeNom || rawData.nom || etudeInfo.nom || '', 180);
    var civilite = sanitizeDsiField_(payload.civilite || etudeInfo.intervenantCivilite || '', 40);
    var prenom = sanitizeDsiField_(payload.prenom || etudeInfo.intervenantPrenom || '', 80);
    var nom = sanitizeDsiField_(payload.nom || etudeInfo.intervenantNom || '', 120);
    var email = sanitizeDsiField_(payload.email || etudeInfo.intervenantMail || '', 180);
    var portable = sanitizeDsiField_(payload.portable || etudeInfo.intervenantPortable || '', 40);
    var fullName = sanitizeDsiField_([civilite, prenom, nom].filter(function(v) { return !!String(v || '').trim(); }).join(' '), 180);

    if (!fullName && !email) {
      throw new Error('Informations intervenant insuffisantes pour contacter le SG.');
    }

    var studyLabel = [ref, etudeNom].filter(function(v) { return !!String(v || '').trim(); }).join(' - ');
    var lines = [];
    lines.push('[Inscription intervenant]');
    if (studyLabel) lines.push('Etude: ' + studyLabel);
    lines.push('Intervenant: ' + (fullName || '(nom non renseigne)'));
    if (email) lines.push('Email: ' + email);
    if (portable) lines.push('Portable: ' + portable);
    lines.push('Action demandee: inscription administrative SG.');
    lines.push('Date: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Paris', 'dd/MM/yyyy HH:mm:ss'));

    var text = sanitizeDsiField_(lines.join('\n'), 3500);
    var webhookUrl = getConfiguredSgIntervenantWebhookUrl_();
    if (!webhookUrl) {
      throw new Error(
        'Webhook Google Chat SG non configuré. ' +
        'Renseignez SG_INTERVENANT_CHAT_WEBHOOK_URL (Script Properties) ' +
        'ou la clé "' + SG_CONFIG_WEBHOOK_KEY_ + '" dans la config SG.'
      );
    }

    var sendResult = postDsiChatMessage_(webhookUrl, { text: text }, 'sg-webhook');

    return {
      success: true,
      chatUrl: SG_INTERVENANT_CHAT_ROOM_URL_,
      sentVia: sendResult.via || '',
      responseCode: sendResult.responseCode || 0,
      messageName: sendResult.messageName || ''
    };
  } catch (err) {
    return {
      success: false,
      error: err && err.message ? err.message : String(err || 'Erreur inconnue'),
      chatUrl: SG_INTERVENANT_CHAT_ROOM_URL_
    };
  }
}

function buildDsiTicketPayload_(payload) {
  var now = new Date();
  var tz = Session.getScriptTimeZone() || 'Europe/Paris';
  var userEmail = '';

  try {
    userEmail = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    userEmail = '';
  }

  return {
    ticketId: buildDsiTicketId_(now),
    createdAtIso: now.toISOString(),
    createdAtLabel: Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss'),
    userEmail: sanitizeDsiField_(userEmail || 'inconnu', 120),
    page: sanitizeDsiField_(payload.page || '', 120),
    pageTitle: sanitizeDsiField_(payload.pageTitle || '', 180),
    source: sanitizeDsiField_(payload.source || 'frontend', 120),
    location: sanitizeDsiField_(payload.location || '', 220),
    url: sanitizeDsiField_(payload.url || '', 350),
    message: sanitizeDsiField_(payload.message || payload.userMessage || 'Erreur inconnue', 600),
    userMessage: sanitizeDsiField_(payload.userMessage || '', 700),
    stack: sanitizeDsiField_(payload.stack || '', 1800),
    userAgent: sanitizeDsiField_(payload.userAgent || '', 300),
    build: sanitizeDsiField_(payload.build || '', 120),
    currentStudyId: sanitizeDsiField_(payload.currentStudyId || '', 80),
    currentTab: sanitizeDsiField_(payload.currentTab || '', 80),
    logs: normalizeDsiLogs_(payload.logs)
  };
}

function buildDsiTicketId_(dateObj) {
  var d = dateObj instanceof Date ? dateObj : new Date();
  var tz = Session.getScriptTimeZone() || 'Europe/Paris';
  var stamp = Utilities.formatDate(d, tz, 'yyyyMMdd-HHmmss');
  var suffix = String(Math.floor(Math.random() * 900) + 100);
  return 'JEDI-' + stamp + '-' + suffix;
}

function normalizeDsiLogs_(rawLogs) {
  var entries = Object.prototype.toString.call(rawLogs) === '[object Array]' ? rawLogs : [];
  var out = [];

  for (var i = 0; i < entries.length; i++) {
    var item = entries[i];
    if (!item) continue;

    if (typeof item === 'string') {
      out.push({
        at: '',
        level: 'info',
        message: sanitizeDsiField_(item, 260),
        location: ''
      });
      continue;
    }

    if (typeof item === 'object') {
      out.push({
        at: sanitizeDsiField_(item.at || '', 40),
        level: sanitizeDsiField_(item.level || 'info', 16),
        message: sanitizeDsiField_(item.message || '', 260),
        location: sanitizeDsiField_(item.location || '', 120)
      });
    }
  }

  if (out.length > 25) {
    out = out.slice(out.length - 25);
  }

  return out;
}

function buildDsiChatMessageText_(ticket) {
  var lines = [];

  lines.push('[ALERTE] Ticket incident ERP');
  lines.push('ID: ' + ticket.ticketId);
  lines.push('Date: ' + ticket.createdAtLabel);
  lines.push('Utilisateur: ' + ticket.userEmail);

  if (ticket.page) {
    var pageLabel = ticket.page + (ticket.pageTitle ? ' (' + ticket.pageTitle + ')' : '');
    lines.push('Page: ' + pageLabel);
  }
  if (ticket.location) lines.push('Emplacement: ' + ticket.location);
  if (ticket.source) lines.push('Source: ' + ticket.source);
  if (ticket.url) lines.push('URL: ' + ticket.url);
  if (ticket.currentStudyId) {
    var etudeLabel = ticket.currentStudyId + (ticket.currentTab ? ' / onglet ' + ticket.currentTab : '');
    lines.push('Contexte etude: ' + etudeLabel);
  }
  if (ticket.build) lines.push('Build: ' + ticket.build);

  lines.push('Message: ' + ticket.message);

  if (ticket.userMessage) {
    lines.push('Contexte utilisateur: ' + ticket.userMessage);
  }

  if (ticket.stack) {
    lines.push('Stack:');
    lines.push(ticket.stack);
  }

  if (ticket.logs && ticket.logs.length) {
    lines.push('Logs recents:');
    for (var i = 0; i < ticket.logs.length; i++) {
      var log = ticket.logs[i];
      var prefix = '[' + (log.level || 'info') + ']';
      var at = log.at ? (' ' + log.at) : '';
      var loc = log.location ? (' @ ' + log.location) : '';
      lines.push(prefix + at + ' ' + (log.message || '') + loc);
    }
  }

  return sanitizeDsiField_(lines.join('\n'), 3800);
}

function sendDsiTicketToChat_(ticket) {
  var text = buildDsiChatMessageText_(ticket);

  var webhookUrl = getConfiguredDsiWebhookUrl_();
  if (webhookUrl) {
    return postDsiChatMessage_(webhookUrl, {
      text: text
    }, 'webhook');
  }

  var apiUrl = 'https://chat.googleapis.com/v1/spaces/' + encodeURIComponent(DSI_CHAT_SPACE_ID_) + '/messages';
  var apiPayload = {
    text: text,
    thread: {
      threadKey: DSI_CHAT_DEFAULT_THREAD_KEY_
    }
  };

  var response = UrlFetchApp.fetch(apiUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(apiPayload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var rawBody = response.getContentText() || '';

  if (code >= 200 && code < 300) {
    var parsed = safeJsonParse_(rawBody) || {};
    return {
      via: 'chat-api',
      responseCode: code,
      messageName: parsed.name || ''
    };
  }

  throw new Error(
    'Envoi Google Chat impossible (' + code + '). ' +
    'Configurez DSI_CHAT_WEBHOOK_URL ou DSI_CHAT_WEBHOOK_KEY/DSI_CHAT_WEBHOOK_TOKEN dans Script Properties. ' +
    'Detail: ' + sanitizeDsiField_(rawBody, 300)
  );
}

function postDsiChatMessage_(url, payload, viaLabel) {
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var rawBody = response.getContentText() || '';

  if (code >= 200 && code < 300) {
    var parsed = safeJsonParse_(rawBody) || {};
    return {
      via: viaLabel || 'webhook',
      responseCode: code,
      messageName: parsed.name || ''
    };
  }

  throw new Error('Webhook Chat DSI en echec (' + code + '): ' + sanitizeDsiField_(rawBody, 300));
}

function getConfiguredDsiWebhookUrl_() {
  var props = PropertiesService.getScriptProperties();

  var directUrl = sanitizeDsiField_(props.getProperty('DSI_CHAT_WEBHOOK_URL') || '', 600);
  if (directUrl) return directUrl;

  var altUrl = sanitizeDsiField_(props.getProperty('CHAT_DSI_WEBHOOK_URL') || '', 600);
  if (altUrl) return altUrl;

  var key = sanitizeDsiField_(props.getProperty('DSI_CHAT_WEBHOOK_KEY') || '', 200);
  var token = sanitizeDsiField_(props.getProperty('DSI_CHAT_WEBHOOK_TOKEN') || '', 400);

  if (key && token) {
    return 'https://chat.googleapis.com/v1/spaces/' + DSI_CHAT_SPACE_ID_ + '/messages?key=' + encodeURIComponent(key) + '&token=' + encodeURIComponent(token);
  }

  return '';
}

function getConfiguredSgIntervenantWebhookUrl_() {
  var props = PropertiesService.getScriptProperties();

  var directUrl = normalizeChatWebhookUrl_(props.getProperty('SG_INTERVENANT_CHAT_WEBHOOK_URL') || '');
  if (directUrl) return directUrl;

  var altUrl = normalizeChatWebhookUrl_(props.getProperty('SG_CHAT_WEBHOOK_URL') || '');
  if (altUrl) return altUrl;

  var genericUrl = normalizeChatWebhookUrl_(props.getProperty('GOOGLE_CHAT_WEBHOOK_URL') || '');
  if (genericUrl) return genericUrl;

  var sgConfigWebhook = '';
  try {
    if (typeof getConfigValue === 'function') {
      sgConfigWebhook = normalizeChatWebhookUrl_(getConfigValue(SG_CONFIG_WEBHOOK_KEY_, ''));
    }
  } catch (e) {
    sgConfigWebhook = '';
  }
  if (sgConfigWebhook) return sgConfigWebhook;

  var key = sanitizeDsiField_(props.getProperty('SG_INTERVENANT_CHAT_WEBHOOK_KEY') || '', 200);
  var token = sanitizeDsiField_(props.getProperty('SG_INTERVENANT_CHAT_WEBHOOK_TOKEN') || '', 400);
  if (!key || !token) {
    key = sanitizeDsiField_(props.getProperty('SG_CHAT_WEBHOOK_KEY') || '', 200);
    token = sanitizeDsiField_(props.getProperty('SG_CHAT_WEBHOOK_TOKEN') || '', 400);
  }

  if (key && token) {
    return 'https://chat.googleapis.com/v1/spaces/' + SG_INTERVENANT_CHAT_SPACE_ID_ + '/messages?key=' + encodeURIComponent(key) + '&token=' + encodeURIComponent(token);
  }

  return '';
}

function normalizeChatWebhookUrl_(rawValue) {
  var value = sanitizeDsiField_(rawValue || '', 700);
  if (!value) return '';
  if (!/^https:\/\/chat\.googleapis\.com\/v1\/spaces\/[^/]+\/messages(\?.*)?$/i.test(value)) return '';
  return value;
}

function sanitizeDsiField_(value, maxLen) {
  var text = String(value == null ? '' : value)
    .replace(/\u0000/g, '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();

  if (maxLen && text.length > maxLen) {
    return text.slice(0, Math.max(0, maxLen - 3)) + '...';
  }

  return text;
}

function safeJsonParse_(value) {
  try {
    return JSON.parse(value);
  } catch (e) {
    return null;
  }
}

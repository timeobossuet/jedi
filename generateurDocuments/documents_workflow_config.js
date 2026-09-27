var DOC_WORKFLOW_CHAT_URL_ = 'https://chat.google.com/room/AAQApAX9EAQ';

function getDocWorkflowWebhookUrl_() {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('DOC_WORKFLOW_WEBHOOK_URL');
  if (!webhookUrl) {
    throw new Error('Script Property manquante: DOC_WORKFLOW_WEBHOOK_URL');
  }
  return String(webhookUrl).trim();
}

function getDocWorkflowConfig_(docKey) {
  var key = String(docKey || '').trim().toLowerCase();
  var configs = {
    devis: {
      key: 'devis',
      suffix: 'DEVIS',
      label: 'Devis',
      templateCandidates: ['devis'],
      reviewEnabled: true
    },
    ce: {
      key: 'ce',
      suffix: 'CE',
      label: 'Convention etude',
      templateCandidates: ['ce', 'convention-etude', 'convention etude', 'convention_etude'],
      reviewEnabled: true
    },
    cc: {
      key: 'cc',
      suffix: 'CC',
      label: 'Convention client',
      templateCandidates: ['cc', 'convention-client', 'convention client', 'convention_client'],
      reviewEnabled: true
    },
    bdc: {
      key: 'bdc',
      suffix: 'BDC',
      label: 'Bon de commande',
      templateCandidates: ['bdc', 'bon-de-commande', 'bon de commande', 'bon_de_commande'],
      reviewEnabled: true
    },
    rm: {
      key: 'rm',
      suffix: 'RM',
      label: 'RM',
      templateCandidates: ['rm', 'releve-mission', 'releve mission', 'releve_mission'],
      reviewEnabled: true
    },
    arm: {
      key: 'arm',
      suffix: 'ARM',
      label: 'ARM',
      templateCandidates: ['arm', 'accord-realisation-mission', 'accord realisation mission', 'accord_realisation_mission'],
      reviewEnabled: true
    },
    arrm: {
      key: 'arrm',
      suffix: 'ARRM',
      label: 'ARRM',
      templateCandidates: ['arrm', 'avenant-realisation-rm', 'avenant realisation rm', 'avenant_realisation_rm'],
      reviewEnabled: true
    },
    arce: {
      key: 'arce',
      suffix: 'ARCE',
      label: 'ARCE',
      templateCandidates: ['arce', 'accord-realisation-ce', 'accord realisation ce', 'accord_realisation_ce'],
      reviewEnabled: true
    },
    ace: {
      key: 'ace',
      suffix: 'ACE',
      label: 'ACE',
      templateCandidates: ['ace', 'avenant-ce', 'avenant ce', 'avenant_ce'],
      reviewEnabled: true
    },
    acc: {
      key: 'acc',
      suffix: 'ACC',
      label: 'ACC',
      templateCandidates: ['acc', 'avenant-cc', 'avenant cc', 'avenant_cc'],
      reviewEnabled: true
    },
    bdcr: {
      key: 'bdcr',
      suffix: 'BDCR',
      label: 'BDCR',
      templateCandidates: ['bdcr', 'bon-de-commande-realisation', 'bon de commande realisation', 'bon_de_commande_realisation'],
      reviewEnabled: true
    }
  };

  if (!configs[key]) {
    throw new Error('Type de document non supporte: ' + docKey);
  }
  return configs[key];
}
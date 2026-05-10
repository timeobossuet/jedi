// Script d'initialisation des feuilles nécessaires à l'ERP JIS
function initAllSheets() {
  var ss = getSGSpreadsheet_();
  var sheets = SG_CONFIG.sheets;
  var headers = {
    membres: [
      'id','civilite','nom','prenom','email','telephone','promo','specialite','roleCA','statut','dateInscription','data_json'
    ],
    reunions: [
      'id','type','date','titre','ordre_du_jour','pv','participants','data_json'
    ],
    votes: [
      'id','objet','date','type','resultat','data_json'
    ],
    archives: [
      'id','type','date','titre','contenu','data_json'
    ],
    licences: [
      'id','logiciel','type_de_licence','date_installation','date_debut_validite','date_fin_validite','proprietaire','contexte_utilisation','preuve_paiement','conditions_generales','remarques','data_json'
    ],
    rgpd: [
      'id','date','demandeur','type','statut','commentaire','dateTraitement'
    ],
    config: [
      'key','value'
    ]
  };
  for (var key in sheets) {
    var name = sheets[key];
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (headers[key]) {
      sheet.clearContents();
      sheet.appendRow(headers[key]);
      sheet.getRange(1,1,1,headers[key].length).setFontWeight('bold');
    }
  }
}

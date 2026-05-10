/**
 * Convertit des balises HTML basiques (issues de Quill) en formatage Google Docs.
 * @param {GoogleAppsScript.Document.Body} body ou element
 * @param {string} textToReplace - ie "{{Contexte}}"
 * @param {string} htmlText - Le texte enrichi "<b>Gras</b>"
 */
function replaceHtmlInDocs(element, searchPattern, htmlText) {
  // Une fonction très basique pour l'exemple.
  // Nettoyage rapide (stripping simple) ou logique de parseur.
  const plainText = htmlText.replace(/<[^>]+>/g, '');
  element.replaceText(searchPattern, plainText);
}

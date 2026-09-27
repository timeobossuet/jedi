function findTemplateTokens_(text) {
  return String(text || '').match(/\{\{[^{}]+\}\}|\{[^{}]+\}/g) || null;
}

function extractTokenInside_(token) {
  var t = String(token || '');
  if (!t) return '';
  if (/^\{\{[\s\S]+\}\}$/.test(t)) return t.substring(2, t.length - 2);
  if (/^\{[\s\S]+\}$/.test(t)) return t.substring(1, t.length - 1);
  return t;
}

function replaceTokensInSlidesTextRange_(textRange, variables) {
  var source;
  try { source = textRange.asString(); } catch (e) { return; }
  if (!source || source.indexOf('{') === -1) return;
  var tokens = findTemplateTokens_(source);
  if (!tokens || !tokens.length) return;
  var unique = {};
  for (var i = 0; i < tokens.length; i++) unique[tokens[i]] = true;
  for (var token in unique) {
    if (!unique.hasOwnProperty(token)) continue;
    var value = evaluateToken_(extractTokenInside_(token), variables);
    var replacement = value === null || typeof value === 'undefined' ? '' : String(value);
    if (isLikelyHtmlContent_(replacement)) {
      if (typeof replaceHtmlTokenInSlidesTextRange_ === 'function' && replaceHtmlTokenInSlidesTextRange_(textRange, token, replacement)) continue;
      replacement = typeof convertHtmlToPlainText === 'function' ? convertHtmlToPlainText(replacement) : replacement;
    }
    try { textRange.replaceAllText(token, replacement); } catch (e2) {
      var msg = String(e2 && e2.message ? e2.message : e2);
      if (msg.indexOf('fusionnée') !== -1 || msg.indexOf('merged') !== -1) return;
      throw e2;
    }
  }
}

function replaceInlinePhaseTokensInDocElement_(element, baseVariables, phases) {
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var textNode = element.asText();
    var source = textNode.getText() || '';
    var expanded = expandPhaseTemplateBlocksInText_(source, baseVariables, phases);
    if (expanded !== source) { textNode.setText(expanded); source = expanded; }
    if (containsPhaseToken_(source)) {
      var rendered = [];
      for (var i = 0; i < phases.length; i++) rendered.push(replaceTokensInText_(source, buildPhaseVariables_(baseVariables, phases[i], i)));
      textNode.setText(rendered.join('\n'));
    }
    return;
  }
  if (typeof element.getNumChildren !== 'function') return;
  for (var c = 0; c < element.getNumChildren(); c++) replaceInlinePhaseTokensInDocElement_(element.getChild(c), baseVariables, phases);
}

function expandPhaseTemplateBlocksInText_(text, baseVariables, phases) {
  var source = String(text || '');
  if (source.indexOf('{phases}') === -1 && source.indexOf('{/phases}') === -1) return source;
  return source.replace(/\{\s*phases\s*\}([\s\S]*?)\{\s*\/\s*phases\s*\}/gi, function(_match, templateBlock) {
    if (!Array.isArray(phases) || !phases.length) return '';
    var chunks = [];
    for (var i = 0; i < phases.length; i++) chunks.push(replaceTokensInText_(templateBlock, buildPhaseVariables_(baseVariables, phases[i], i)));
    return chunks.join('');
  });
}

function replacePlanningTokenInDocument_(body, variables) {
  if (!body || !variables) return;
  var planningValue = getVariable_(variables, 'phases.planning');
  if (planningValue === null || typeof planningValue === 'undefined') return;
  try { body.replaceText('\\{\\s*phases\\.planning\\s*\\}', String(planningValue)); } catch (e) {}
}

function replaceTokensInDocElement_(element, variables) {
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var textNode = element.asText();
    var text = textNode.getText();
    var tokens = text ? findTemplateTokens_(text) : null;
    if (tokens && tokens.length) {
      var unique = {};
      for (var u = 0; u < tokens.length; u++) unique[tokens[u]] = true;
      for (var token in unique) {
        if (!unique.hasOwnProperty(token)) continue;
        var value = evaluateToken_(extractTokenInside_(token), variables);
        var replacement = value === null || typeof value === 'undefined' ? '' : String(value);
        if (isLikelyHtmlContent_(replacement)) {
          if (typeof replaceHtmlTokenInDocsTextNode_ === 'function' && replaceHtmlTokenInDocsTextNode_(textNode, token, replacement)) return;
          replacement = typeof convertHtmlToPlainText === 'function' ? convertHtmlToPlainText(replacement) : replacement;
        }
        textNode.replaceText(escapeRegex_(token), replacement);
      }
    }
    return;
  }
  if (typeof element.getNumChildren !== 'function') return;
  for (var i = 0; i < element.getNumChildren(); i++) replaceTokensInDocElement_(element.getChild(i), variables);
}

function replaceTokensInText_(text, variables) {
  if (!text || text.indexOf('{') === -1) return text;
  return String(text).replace(/\{\{([^{}]+)\}\}|\{([^{}]+)\}/g, function(match, insideDouble, insideSingle) {
    var value = evaluateToken_(insideDouble || insideSingle || '', variables);
    var strValue = value === null || typeof value === 'undefined' ? '' : String(value);
    if (isLikelyHtmlContent_(strValue) && typeof convertHtmlToPlainText === 'function') strValue = convertHtmlToPlainText(strValue);
    return strValue;
  });
}
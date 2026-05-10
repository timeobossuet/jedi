/**
 * Récupère l'ID d'un template Drive à partir de son nom
 * dans l'annuaire de templates (colonne A = nom, B = fileId)
 */
function getTemplateFileIdByName(templateName) {
  var directorySpreadsheetId = '1PUFv22FmxLHcYOUQq1RBZxgiwUJejumbGsbU_5vV-QY';
  var ss = SpreadsheetApp.openById(directorySpreadsheetId);
  var sheet = ss.getSheets()[0];
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();

  for (var i = 0; i < values.length; i++) {
    var name = values[i][0] ? String(values[i][0]).trim() : '';
    var fileIdRaw = values[i][1] ? String(values[i][1]).trim() : '';
    if (name === templateName) {
      if (!fileIdRaw) throw new Error("Template trouvé mais ID de fichier vide: " + templateName);
      var normalizedId = normalizeDriveFileId(fileIdRaw);
      if (!normalizedId) {
        throw new Error("ID template invalide pour '" + templateName + "': " + fileIdRaw);
      }
      return normalizedId;
    }
  }

  throw new Error("Template introuvable dans l'annuaire: " + templateName);
}

function normalizeDriveFileId(value) {
  if (!value) return '';
  var v = String(value).trim();

  // Cas 1: URL Drive/Docs/Sheets
  var m = v.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];

  // Cas 2: query param ?id=
  m = v.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];

  // Cas 3: déjà un ID brut
  if (/^[a-zA-Z0-9_-]{20,}$/.test(v)) return v;

  return '';
}

function formatNumber(num) {
    if (typeof num !== 'number' || isNaN(num)) return "0,00";
    return num.toFixed(2).replace('.', ',');
}

// --- Alias FR (plus explicites) ---
function obtenirIdModeleParNom(templateName) {
  return getTemplateFileIdByName(templateName);
}

function normaliserIdDrive(value) {
  return normalizeDriveFileId(value);
}

function formaterNombreFrancais(num) {
  return formatNumber(num);
}

/**
 * Détermine si une valeur ressemble à du HTML enrichi (tags réels ou encodés).
 */
function isLikelyHtmlContent_(value) {
  var text = String(value || '');
  if (!text) return false;
  if (/<\/?[a-z][^>]*>/i.test(text)) return true;
  if (/&lt;\/?[a-z][^&]*&gt;/i.test(text)) return true;
  return false;
}

function decodeHtmlEntities_(input) {
  var text = String(input || '');
  if (!text) return '';

  text = text.replace(/&#(\d+);/g, function(_m, dec) {
    var code = parseInt(dec, 10);
    if (isNaN(code)) return _m;
    try {
      return String.fromCharCode(code);
    } catch (e) {
      return _m;
    }
  });

  text = text.replace(/&#x([0-9a-f]+);/gi, function(_m, hex) {
    var code = parseInt(hex, 16);
    if (isNaN(code)) return _m;
    try {
      return String.fromCharCode(code);
    } catch (e) {
      return _m;
    }
  });

  var named = {
    nbsp: ' ',
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'",
    '#39': "'"
  };

  text = text.replace(/&([a-zA-Z0-9#]+);/g, function(_m, name) {
    return Object.prototype.hasOwnProperty.call(named, name) ? named[name] : _m;
  });

  return text;
}

function normalizeHtmlInput_(htmlContent) {
  var html = decodeHtmlEntities_(htmlContent || '');
  if (!html) return '';

  html = html
    .replace(/\r\n?/g, '\n')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?\s*>/gi, '<br/>');

  return html;
}

function isBlockTagName_(name) {
  return /^(p|div|section|article|header|footer|h[1-6]|blockquote)$/i.test(String(name || ''));
}

function isInlineStyleTagName_(name) {
  return /^(strong|b|em|i|u|s|strike|del|a|span)$/i.test(String(name || ''));
}

function extractTagName_(rawTag) {
  var m = String(rawTag || '').match(/^<\/?\s*([a-zA-Z0-9]+)/);
  return m && m[1] ? m[1].toLowerCase() : '';
}

function isClosingTag_(rawTag) {
  return /^<\//.test(String(rawTag || ''));
}

function isSelfClosingTag_(rawTag) {
  return /\/>$/.test(String(rawTag || ''));
}

function getAttrValue_(rawTag, attrName) {
  var re = new RegExp(attrName + '\\s*=\\s*"([^"]*)"|' + attrName + '\\s*=\\s*\'([^\']*)\'|' + attrName + '\\s*=\\s*([^\\s>]+)', 'i');
  var m = String(rawTag || '').match(re);
  return m ? (m[1] || m[2] || m[3] || '') : '';
}

function cloneStyleState_(style) {
  return {
    bold: !!(style && style.bold),
    italic: !!(style && style.italic),
    underline: !!(style && style.underline),
    strike: !!(style && style.strike),
    link: style && style.link ? String(style.link) : ''
  };
}

function createRichTextBlock_(options) {
  var opts = options || {};
  return {
    text: '',
    spans: [],
    listType: opts.listType || null,
    listLevel: opts.listLevel || 0,
    listIndex: opts.listIndex || 0,
    headingLevel: opts.headingLevel || 0
  };
}

function appendStyledTextToBlock_(block, text, style) {
  if (!block) return;
  var chunk = String(text || '');
  if (!chunk) return;

  var start = block.text.length;
  block.text += chunk;
  var endExclusive = block.text.length;

  var hasStyle = !!(style && (style.bold || style.italic || style.underline || style.strike || style.link));
  if (!hasStyle) return;

  var last = block.spans.length ? block.spans[block.spans.length - 1] : null;
  var sameAsLast = last &&
    last.endExclusive === start &&
    !!last.bold === !!style.bold &&
    !!last.italic === !!style.italic &&
    !!last.underline === !!style.underline &&
    !!last.strike === !!style.strike &&
    String(last.link || '') === String(style.link || '');

  if (sameAsLast) {
    last.endExclusive = endExclusive;
    return;
  }

  block.spans.push({
    start: start,
    endExclusive: endExclusive,
    bold: !!style.bold,
    italic: !!style.italic,
    underline: !!style.underline,
    strike: !!style.strike,
    link: style.link ? String(style.link) : ''
  });
}

/**
 * Parse un sous-ensemble HTML (Quill) en blocs + styles inline.
 */
function parseHtmlToRichBlocks_(htmlContent) {
  var html = normalizeHtmlInput_(htmlContent);
  if (!html) return [];

  if (html.indexOf('<') === -1) {
    var plain = createRichTextBlock_();
    plain.text = decodeHtmlEntities_(html);
    return [plain];
  }

  var tokens = html.match(/<[^>]+>|[^<]+/g) || [];
  var blocks = [];
  var currentBlock = createRichTextBlock_();
  var styleStack = [{ tag: '__root__', style: cloneStyleState_({}) }];
  var listStack = [];
  var headingStack = [];

  function currentStyle_() {
    return styleStack[styleStack.length - 1].style;
  }

  function openInlineStyle_(tagName, rawTag) {
    var next = cloneStyleState_(currentStyle_());
    if (tagName === 'strong' || tagName === 'b') next.bold = true;
    else if (tagName === 'em' || tagName === 'i') next.italic = true;
    else if (tagName === 'u') next.underline = true;
    else if (tagName === 's' || tagName === 'strike' || tagName === 'del') next.strike = true;
    else if (tagName === 'a') {
      var href = getAttrValue_(rawTag, 'href');
      next.link = href ? String(href) : '';
    } else if (tagName === 'span') {
      var styleAttr = String(getAttrValue_(rawTag, 'style') || '').toLowerCase();
      if (/font-weight\s*:\s*(bold|[6-9]00)/.test(styleAttr)) next.bold = true;
      if (/font-style\s*:\s*italic/.test(styleAttr)) next.italic = true;
      if (/text-decoration\s*:\s*[^;]*underline/.test(styleAttr)) next.underline = true;
      if (/text-decoration\s*:\s*[^;]*(line-through|strike)/.test(styleAttr)) next.strike = true;
    }
    styleStack.push({ tag: tagName, style: next });
  }

  function closeInlineStyle_(tagName) {
    for (var i = styleStack.length - 1; i > 0; i--) {
      if (styleStack[i].tag === tagName) {
        styleStack.splice(i, 1);
        return;
      }
    }
  }

  function flushBlock_(force) {
    var hasText = currentBlock.text.length > 0;
    if (force || hasText) {
      blocks.push(currentBlock);
    }
    currentBlock = createRichTextBlock_();
  }

  function ensureBlockForListItem_() {
    if (!listStack.length) return;
    if (currentBlock.listType) return;

    var top = listStack[listStack.length - 1];
    currentBlock.listType = top.type;
    currentBlock.listLevel = listStack.length;
    currentBlock.listIndex = top.type === 'ol' ? top.index : 0;
  }

  for (var t = 0; t < tokens.length; t++) {
    var token = tokens[t];
    if (!token) continue;

    if (token.charAt(0) !== '<') {
      var textChunk = decodeHtmlEntities_(token);
      if (!textChunk) continue;
      ensureBlockForListItem_();
      appendStyledTextToBlock_(currentBlock, textChunk, currentStyle_());
      continue;
    }

    var tagName = extractTagName_(token);
    if (!tagName) continue;
    var closing = isClosingTag_(token);
    var selfClosing = isSelfClosingTag_(token);

    if (tagName === 'br') {
      ensureBlockForListItem_();
      appendStyledTextToBlock_(currentBlock, '\n', currentStyle_());
      continue;
    }

    if (!closing && (tagName === 'ul' || tagName === 'ol')) {
      flushBlock_(false);
      listStack.push({ type: tagName, index: 0 });
      continue;
    }

    if (closing && (tagName === 'ul' || tagName === 'ol')) {
      flushBlock_(false);
      if (listStack.length) listStack.pop();
      continue;
    }

    if (!closing && tagName === 'li') {
      flushBlock_(false);
      if (!listStack.length) listStack.push({ type: 'ul', index: 0 });

      var resolvedType = listStack[listStack.length - 1].type;
      var dataList = String(getAttrValue_(token, 'data-list') || '').toLowerCase();
      if (dataList === 'bullet') resolvedType = 'ul';
      else if (dataList === 'ordered') resolvedType = 'ol';

      if (resolvedType === 'ol') {
        listStack[listStack.length - 1].index += 1;
      }

      currentBlock = createRichTextBlock_({
        listType: resolvedType,
        listLevel: listStack.length,
        listIndex: resolvedType === 'ol' ? listStack[listStack.length - 1].index : 0
      });
      continue;
    }

    if (closing && tagName === 'li') {
      flushBlock_(false);
      continue;
    }

    if (isBlockTagName_(tagName)) {
      if (!closing) {
        flushBlock_(false);
        currentBlock = createRichTextBlock_();
        if (/^h[1-6]$/.test(tagName)) {
          var level = parseInt(tagName.charAt(1), 10);
          currentBlock.headingLevel = isNaN(level) ? 0 : level;
          openInlineStyle_(tagName, '<strong>');
          headingStack.push(tagName);
        }
      } else {
        if (headingStack.length && headingStack[headingStack.length - 1] === tagName) {
          headingStack.pop();
          closeInlineStyle_(tagName);
        }
        flushBlock_(false);
      }
      continue;
    }

    if (isInlineStyleTagName_(tagName)) {
      if (!closing) {
        openInlineStyle_(tagName, token);
        if (selfClosing) closeInlineStyle_(tagName);
      } else {
        closeInlineStyle_(tagName);
      }
      continue;
    }
  }

  flushBlock_(false);

  var cleaned = [];
  for (var b = 0; b < blocks.length; b++) {
    var block = blocks[b];
    if (!block) continue;
    var keep = block.text.length > 0 || !!block.listType;
    if (keep) cleaned.push(block);
  }

  return cleaned;
}

function buildListPrefix_(block) {
  if (!block || !block.listType) return '';
  var indent = '';
  var level = Math.max(1, parseInt(block.listLevel || 1, 10));
  for (var i = 1; i < level; i++) indent += '  ';

  if (block.listType === 'ol') {
    var n = parseInt(block.listIndex || 1, 10);
    if (!n || n < 1) n = 1;
    return indent + n + '. ';
  }
  return indent + '• ';
}

function buildSlidesHtmlRenderPlan_(htmlContent) {
  var blocks = parseHtmlToRichBlocks_(htmlContent);
  var parts = [];
  var spans = [];
  var cursor = 0;

  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    var prefix = buildListPrefix_(block);
    var text = prefix + String(block.text || '');
    parts.push(text);

    for (var s = 0; s < block.spans.length; s++) {
      var span = block.spans[s];
      spans.push({
        start: cursor + prefix.length + span.start,
        endExclusive: cursor + prefix.length + span.endExclusive,
        bold: !!span.bold,
        italic: !!span.italic,
        underline: !!span.underline,
        strike: !!span.strike,
        link: span.link ? String(span.link) : ''
      });
    }

    cursor += text.length;
    if (i < blocks.length - 1) {
      parts.push('\n');
      cursor += 1;
    }
  }

  return {
    text: parts.join(''),
    spans: spans
  };
}

function applySlidesInlineStyles_(textRange, spans) {
  if (!textRange || !spans || !spans.length) return;

  for (var i = 0; i < spans.length; i++) {
    var span = spans[i];
    var start = parseInt(span.start, 10);
    var end = parseInt(span.endExclusive, 10) - 1;
    if (!(start >= 0) || !(end >= start)) continue;

    try {
      var sub = textRange.getRange(start, end);
      var style = sub.getTextStyle();
      if (span.bold) style.setBold(true);
      if (span.italic) style.setItalic(true);
      if (span.underline) style.setUnderline(true);
      if (span.strike && typeof style.setStrikethrough === 'function') style.setStrikethrough(true);
      if (span.link && typeof style.setLinkUrl === 'function') style.setLinkUrl(span.link);
    } catch (e) {}
  }
}

function collectTokenRanges_(source, token) {
  var text = String(source || '');
  var tk = String(token || '');
  var ranges = [];
  if (!text || !tk) return ranges;

  var cursor = 0;
  while (true) {
    var idx = text.indexOf(tk, cursor);
    if (idx < 0) break;
    ranges.push({ start: idx, endExclusive: idx + tk.length });
    cursor = idx + tk.length;
  }
  return ranges;
}

function safeExec_(fn, fallback) {
  try {
    return fn();
  } catch (e) {
    return fallback;
  }
}

function captureSlidesBaseStyleAt_(textRange, offset) {
  var o = parseInt(offset, 10);
  if (!textRange || !(o >= 0)) return null;

  try {
    var style = textRange.getRange(o, o).getTextStyle();
    if (!style) return null;

    var colorHex = '';
    try {
      var colorObj = style.getForegroundColor();
      if (colorObj && typeof colorObj.asRgbColor === 'function') {
        var rgb = colorObj.asRgbColor();
        if (rgb && typeof rgb.asHexString === 'function') colorHex = String(rgb.asHexString() || '');
      }
    } catch (eColor) {}

    return {
      bold: safeExec_(function() { return style.isBold(); }, null),
      italic: safeExec_(function() { return style.isItalic(); }, null),
      underline: safeExec_(function() { return style.isUnderline(); }, null),
      strike: safeExec_(function() { return style.isStrikethrough(); }, null),
      fontFamily: safeExec_(function() { return style.getFontFamily(); }, ''),
      fontSize: safeExec_(function() { return style.getFontSize(); }, null),
      foregroundColor: colorHex
    };
  } catch (e) {
    return null;
  }
}

function applySlidesBaseStyle_(textRange, start, end, baseStyle) {
  if (!textRange || !(start >= 0) || !(end >= start) || !baseStyle) return;

  try {
    var sub = textRange.getRange(start, end);
    var style = sub.getTextStyle();

    if (baseStyle.bold !== null && typeof baseStyle.bold !== 'undefined') style.setBold(!!baseStyle.bold);
    if (baseStyle.italic !== null && typeof baseStyle.italic !== 'undefined') style.setItalic(!!baseStyle.italic);
    if (baseStyle.underline !== null && typeof baseStyle.underline !== 'undefined') style.setUnderline(!!baseStyle.underline);
    if (baseStyle.strike !== null && typeof baseStyle.strike !== 'undefined' && typeof style.setStrikethrough === 'function') {
      style.setStrikethrough(!!baseStyle.strike);
    }
    if (baseStyle.fontFamily) style.setFontFamily(String(baseStyle.fontFamily));
    if (baseStyle.fontSize && typeof style.setFontSize === 'function') style.setFontSize(baseStyle.fontSize);
    if (baseStyle.foregroundColor) style.setForegroundColor(baseStyle.foregroundColor);
  } catch (e) {}
}

function applySlidesInlineStylesAtOffset_(textRange, spans, offset) {
  if (!textRange || !spans || !spans.length) return;
  var baseOffset = parseInt(offset || 0, 10) || 0;

  for (var i = 0; i < spans.length; i++) {
    var span = spans[i];
    var start = baseOffset + parseInt(span.start, 10);
    var end = baseOffset + parseInt(span.endExclusive, 10) - 1;
    if (!(start >= 0) || !(end >= start)) continue;

    try {
      var sub = textRange.getRange(start, end);
      var style = sub.getTextStyle();
      if (span.bold) style.setBold(true);
      if (span.italic) style.setItalic(true);
      if (span.underline) style.setUnderline(true);
      if (span.strike && typeof style.setStrikethrough === 'function') style.setStrikethrough(true);
      if (span.link && typeof style.setLinkUrl === 'function') style.setLinkUrl(span.link);
    } catch (e) {}
  }
}

function mergeHtmlPlanIntoSourceText_(source, token, plan) {
  var src = String(source || '');
  var tk = String(token || '');
  var injectedText = plan && typeof plan.text === 'string' ? plan.text : '';
  var injectedSpans = plan && Object.prototype.toString.call(plan.spans) === '[object Array]' ? plan.spans : [];

  if (!src || !tk) {
    return { text: src, spans: [], replaced: false };
  }

  var pieces = [];
  var spans = [];
  var srcCursor = 0;
  var dstCursor = 0;
  var replaced = false;

  while (true) {
    var idx = src.indexOf(tk, srcCursor);
    if (idx < 0) break;

    var before = src.substring(srcCursor, idx);
    if (before) {
      pieces.push(before);
      dstCursor += before.length;
    }

    pieces.push(injectedText);
    for (var i = 0; i < injectedSpans.length; i++) {
      var span = injectedSpans[i];
      spans.push({
        start: dstCursor + parseInt(span.start, 10),
        endExclusive: dstCursor + parseInt(span.endExclusive, 10),
        bold: !!span.bold,
        italic: !!span.italic,
        underline: !!span.underline,
        strike: !!span.strike,
        link: span.link ? String(span.link) : ''
      });
    }
    dstCursor += injectedText.length;

    srcCursor = idx + tk.length;
    replaced = true;
  }

  var tail = src.substring(srcCursor);
  if (tail) pieces.push(tail);

  return {
    text: pieces.join(''),
    spans: spans,
    replaced: replaced
  };
}

/**
 * Remplacement HTML enrichi pour un token dans une zone de texte Slides.
 * Gère les tokens seuls ou inline (avec texte avant/après).
 */
function replaceHtmlTokenInSlidesTextRange_(textRange, token, htmlContent) {
  if (!textRange || !token) return false;

  var source = '';
  try {
    source = String(textRange.asString() || '');
  } catch (e) {
    return false;
  }

  var ranges = collectTokenRanges_(source, token);
  if (!ranges.length) return false;

  var plan = buildSlidesHtmlRenderPlan_(htmlContent);
  var replacementText = plan && typeof plan.text === 'string' ? plan.text : '';
  var replacementSpans = plan && Object.prototype.toString.call(plan.spans) === '[object Array]' ? plan.spans : [];

  for (var i = ranges.length - 1; i >= 0; i--) {
    var r = ranges[i];
    var start = r.start;
    var end = r.endExclusive - 1;

    var baseStyle = captureSlidesBaseStyleAt_(textRange, start);

    try {
      textRange.getRange(start, end).setText(replacementText);
    } catch (eSet) {
      return false;
    }

    if (replacementText) {
      var newEnd = start + replacementText.length - 1;
      applySlidesBaseStyle_(textRange, start, newEnd, baseStyle);
      applySlidesInlineStylesAtOffset_(textRange, replacementSpans, start);
    }
  }

  return true;
}

function forEachPresentationTextRange_(presentation, callback) {
  if (!presentation || typeof callback !== 'function') return;
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

function replaceHtmlTokenInPresentation_(presentation, token, htmlContent) {
  if (!presentation || !token) return false;
  var replacedRich = false;

  forEachPresentationTextRange_(presentation, function(textRange) {
    try {
      if (replaceHtmlTokenInSlidesTextRange_(textRange, token, htmlContent)) {
        replacedRich = true;
      }
    } catch (e) {}
  });

  return replacedRich;
}

function replaceHtmlTokenInSlide_(slide, token, htmlContent) {
  if (!slide || !token) return false;
  var replacedRich = false;
  var elements = slide.getPageElements();

  for (var i = 0; i < elements.length; i++) {
    var type = elements[i].getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) {
      try {
        if (replaceHtmlTokenInSlidesTextRange_(elements[i].asShape().getText(), token, htmlContent)) {
          replacedRich = true;
        }
      } catch (e) {}
    } else if (type === SlidesApp.PageElementType.TABLE) {
      try {
        var table = elements[i].asTable();
        for (var r = 0; r < table.getNumRows(); r++) {
          for (var c = 0; c < table.getNumColumns(); c++) {
            if (replaceHtmlTokenInSlidesTextRange_(table.getCell(r, c).getText(), token, htmlContent)) {
              replacedRich = true;
            }
          }
        }
      } catch (e2) {}
    }
  }

  return replacedRich;
}

function applyDocsInlineStyles_(textElement, spans, offset) {
  if (!textElement || !spans || !spans.length) return;
  var baseOffset = parseInt(offset || 0, 10) || 0;

  for (var i = 0; i < spans.length; i++) {
    var span = spans[i];
    var start = baseOffset + parseInt(span.start, 10);
    var end = baseOffset + parseInt(span.endExclusive, 10) - 1;
    if (!(start >= 0) || !(end >= start)) continue;

    try {
      if (span.bold) textElement.setBold(start, end, true);
      if (span.italic) textElement.setItalic(start, end, true);
      if (span.underline) textElement.setUnderline(start, end, true);
      if (span.strike && typeof textElement.setStrikethrough === 'function') textElement.setStrikethrough(start, end, true);
      if (span.link && typeof textElement.setLinkUrl === 'function') textElement.setLinkUrl(start, end, span.link);
    } catch (e) {}
  }
}

function captureDocsBaseStyleAt_(textNode, offset) {
  var o = parseInt(offset, 10);
  if (!textNode || !(o >= 0)) return null;

  return {
    bold: safeExec_(function() { return textNode.isBold(o); }, null),
    italic: safeExec_(function() { return textNode.isItalic(o); }, null),
    underline: safeExec_(function() { return textNode.isUnderline(o); }, null),
    strike: safeExec_(function() { return textNode.isStrikethrough(o); }, null),
    fontFamily: safeExec_(function() { return textNode.getFontFamily(o); }, ''),
    fontSize: safeExec_(function() { return textNode.getFontSize(o); }, null),
    foregroundColor: safeExec_(function() { return textNode.getForegroundColor(o); }, ''),
    backgroundColor: safeExec_(function() { return textNode.getBackgroundColor(o); }, ''),
    link: safeExec_(function() { return textNode.getLinkUrl(o); }, '')
  };
}

function applyDocsBaseStyleSnapshot_(textNode, start, end, baseStyle) {
  if (!textNode || !(start >= 0) || !(end >= start) || !baseStyle) return;

  try {
    if (baseStyle.bold !== null && typeof baseStyle.bold !== 'undefined') textNode.setBold(start, end, !!baseStyle.bold);
    if (baseStyle.italic !== null && typeof baseStyle.italic !== 'undefined') textNode.setItalic(start, end, !!baseStyle.italic);
    if (baseStyle.underline !== null && typeof baseStyle.underline !== 'undefined') textNode.setUnderline(start, end, !!baseStyle.underline);
    if (baseStyle.strike !== null && typeof baseStyle.strike !== 'undefined' && typeof textNode.setStrikethrough === 'function') {
      textNode.setStrikethrough(start, end, !!baseStyle.strike);
    }
    if (baseStyle.fontFamily) textNode.setFontFamily(start, end, String(baseStyle.fontFamily));
    if (baseStyle.fontSize) textNode.setFontSize(start, end, baseStyle.fontSize);
    if (baseStyle.foregroundColor) textNode.setForegroundColor(start, end, baseStyle.foregroundColor);
    if (baseStyle.backgroundColor) textNode.setBackgroundColor(start, end, baseStyle.backgroundColor);
    if (baseStyle.link && typeof textNode.setLinkUrl === 'function') textNode.setLinkUrl(start, end, baseStyle.link);
  } catch (e) {}
}

function replaceHtmlTokenInlineInDocsTextNode_(textNode, token, htmlContent) {
  if (!textNode || !token) return false;

  var source = String(textNode.getText() || '');
  var ranges = collectTokenRanges_(source, token);
  if (!ranges.length) return false;

  var plan = buildSlidesHtmlRenderPlan_(htmlContent);
  var replacementText = plan && typeof plan.text === 'string' ? plan.text : '';
  var replacementSpans = plan && Object.prototype.toString.call(plan.spans) === '[object Array]' ? plan.spans : [];

  try {
    for (var i = ranges.length - 1; i >= 0; i--) {
      var r = ranges[i];
      var start = r.start;
      var end = r.endExclusive - 1;
      var baseStyle = captureDocsBaseStyleAt_(textNode, start);

      textNode.deleteText(start, end);
      if (replacementText) {
        textNode.insertText(start, replacementText);
        var insertedEnd = start + replacementText.length - 1;
        applyDocsBaseStyleSnapshot_(textNode, start, insertedEnd, baseStyle);
        applyDocsInlineStyles_(textNode, replacementSpans, start);
      }
    }
    return true;
  } catch (e) {
    return false;
  }
}

function insertRichHtmlInContainerAtIndex_(container, index, htmlContent, baseStyle) {
  if (!container || typeof container.insertParagraph !== 'function') return false;

  var blocks = parseHtmlToRichBlocks_(htmlContent);
  if (!blocks.length) {
    container.insertParagraph(index, '');
    return true;
  }

  var cursor = index;
  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    var prefix = '';
    var text = String(block.text || '');
    var element;

    if (block.listType === 'ul' && typeof container.insertListItem === 'function') {
      element = container.insertListItem(cursor, text);
    } else {
      if (block.listType === 'ol') {
        prefix = buildListPrefix_(block);
      }
      element = container.insertParagraph(cursor, prefix + text);
      if (!block.listType && block.headingLevel && typeof element.setHeading === 'function') {
        var headingKey = 'HEADING' + parseInt(block.headingLevel, 10);
        if (DocumentApp.ParagraphHeading && DocumentApp.ParagraphHeading[headingKey]) {
          try {
            element.setHeading(DocumentApp.ParagraphHeading[headingKey]);
          } catch (e) {}
        }
      }
    }

    try {
      var textElement = element.editAsText();
      if (text && text.length) {
        applyDocsBaseStyleSnapshot_(textElement, 0, (prefix + text).length - 1, baseStyle || null);
      }
      applyDocsInlineStyles_(textElement, block.spans, prefix.length);
    } catch (e2) {}

    cursor++;
  }

  return true;
}

/**
 * Remplacement HTML enrichi pour Google Docs.
 * Ne s'applique que si le placeholder occupe entièrement le paragraphe/list-item.
 */
function replaceHtmlTokenInDocsTextNode_(textNode, token, htmlContent) {
  if (!textNode || !token) return false;

  var source = String(textNode.getText() || '');
  if (!source || source.indexOf(token) === -1) return false;

  if (source.trim() !== token) {
    return replaceHtmlTokenInlineInDocsTextNode_(textNode, token, htmlContent);
  }

  var tokenStart = source.indexOf(token);
  var baseStyle = tokenStart >= 0 ? captureDocsBaseStyleAt_(textNode, tokenStart) : null;

  var parent = textNode.getParent();
  if (!parent || typeof parent.getParent !== 'function') return false;
  var container = parent.getParent();
  if (!container || typeof container.getChildIndex !== 'function') return false;

  var index;
  try {
    index = container.getChildIndex(parent);
  } catch (e) {
    return false;
  }

  try {
    parent.removeFromParent();
  } catch (e2) {
    return false;
  }

  var replacedBlock = insertRichHtmlInContainerAtIndex_(container, index, htmlContent, baseStyle);
  if (replacedBlock) return true;

  // Fallback si insertion en blocs impossible
  try {
    var fallback = container.insertParagraph(index, convertHtmlToPlainText(htmlContent));
    if (fallback && baseStyle) {
      try {
        var fallbackText = fallback.editAsText();
        var len = String(fallbackText.getText() || '').length;
        if (len > 0) applyDocsBaseStyleSnapshot_(fallbackText, 0, len - 1, baseStyle);
      } catch (eFb) {}
    }
    return !!fallback;
  } catch (e3) {
    return false;
  }
}

function escapeDocRegexPattern_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Remplace un marqueur dans un body Docs avec rendu HTML enrichi quand possible.
 */
function replaceHtmlInDocument(body, marker, htmlContent) {
  if (!body || !marker) return;
  if (!htmlContent) {
    body.replaceText(marker, '');
    return;
  }

  var replacedRich = false;
  var markerPattern = escapeDocRegexPattern_(marker);
  try {
    var found = body.findText(markerPattern);
    while (found) {
      var elem = found.getElement();
      if (elem && elem.getType && elem.getType() === DocumentApp.ElementType.TEXT) {
        if (replaceHtmlTokenInDocsTextNode_(elem.asText(), marker, htmlContent)) {
          replacedRich = true;
        }
      }
      found = body.findText(markerPattern, found);
    }
  } catch (e) {}

  if (!replacedRich) {
    body.replaceText(markerPattern, convertHtmlToPlainText(htmlContent));
  }
}

function convertHtmlToPlainText(htmlContent) {
  if (!htmlContent) return '';

  var blocks = parseHtmlToRichBlocks_(htmlContent);
  if (!blocks.length) {
    return decodeHtmlEntities_(String(htmlContent || ''))
      .replace(/<[^>]+>/g, '')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  var lines = [];
  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    lines.push(buildListPrefix_(block) + String(block.text || ''));
  }

  return lines.join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

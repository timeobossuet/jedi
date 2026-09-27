function replaceTokensInPresentation_(presentation, variables) {
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) replaceTokensInSlide_(slides[i], variables);
}

function applyPhaseTokensInPresentation_(presentation, baseVariables, phases) {
  if (!Array.isArray(phases) || phases.length === 0) return;
  var originalSlides = presentation.getSlides().slice();
  for (var i = 0; i < originalSlides.length; i++) {
    var slide = originalSlides[i];
    if (slideHasPhaseTokensOutsideTables_(slide)) {
      var phaseSlides = [slide];
      for (var d = 1; d < phases.length; d++) phaseSlides.push(slide.duplicate());
      reorderPhaseSlidesAfterTemplate_(presentation, phaseSlides);
      for (var p = 0; p < phases.length; p++) {
        replaceTokensInSlide_(phaseSlides[p], buildPhaseVariables_(baseVariables, phases[p], p));
      }
    } else {
      duplicatePhaseRowsInSlide_(slide, baseVariables, phases);
    }
  }

  var allSlides = presentation.getSlides();
  for (var s = 0; s < allSlides.length; s++) renderPlanningTableIfNeeded_(allSlides[s], phases);
}

function reorderPhaseSlidesAfterTemplate_(presentation, phaseSlides) {
  if (!presentation || !phaseSlides || phaseSlides.length <= 1) return;
  var templateId;
  try { templateId = phaseSlides[0].getObjectId(); } catch (e) { return; }
  var templateIndex = getSlideIndexByObjectId_(presentation, templateId);
  if (templateIndex < 0) return;
  for (var i = 1; i < phaseSlides.length; i++) {
    try { phaseSlides[i].move(templateIndex + i); } catch (e2) {}
  }
}

function getSlideIndexByObjectId_(presentation, objectId) {
  if (!presentation || !objectId) return -1;
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    try { if (slides[i].getObjectId() === objectId) return i; } catch (e) {}
  }
  return -1;
}

function replaceTokensInSlide_(slide, variables) {
  var elements = slide.getPageElements();
  for (var j = 0; j < elements.length; j++) {
    var type = elements[j].getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) {
      try { replaceTokensInSlidesTextRange_(elements[j].asShape().getText(), variables); } catch (e) {}
    } else if (type === SlidesApp.PageElementType.TABLE) {
      try {
        var table = elements[j].asTable();
        for (var r = 0; r < table.getNumRows(); r++) {
          for (var c = 0; c < table.getNumColumns(); c++) {
            try { replaceTokensInSlidesTextRange_(table.getCell(r, c).getText(), variables); } catch (e2) {}
          }
        }
      } catch (e3) {}
    }
  }
}

function duplicatePhaseRowsInSlide_(slide, baseVariables, phases) {
  if (!Array.isArray(phases) || phases.length === 0) return;
  var elements = slide.getPageElements();
  for (var j = 0; j < elements.length; j++) {
    if (elements[j].getPageElementType() !== SlidesApp.PageElementType.TABLE) continue;
    var table;
    try { table = elements[j].asTable(); } catch (e) { continue; }
    for (var r = 0; r < table.getNumRows(); r++) {
      var rowHasPhaseToken = false;
      var rowTexts = [];
      for (var c = 0; c < table.getNumColumns(); c++) {
        var cellText = safeGetTableCellText_(table, r, c);
        rowTexts.push(cellText);
        if (containsPhaseToken_(cellText)) rowHasPhaseToken = true;
      }
      if (!rowHasPhaseToken) continue;
      for (var p = 0; p < phases.length; p++) {
        var targetRow = r + p;
        if (p > 0) table.insertRow(targetRow);
        var phaseVars = buildPhaseVariables_(baseVariables, phases[p], p);
        for (var cc = 0; cc < table.getNumColumns(); cc++) {
          safeSetTableCellText_(table, targetRow, cc, replaceTokensInText_(rowTexts[cc] || '', phaseVars));
        }
      }
      r += phases.length - 1;
    }
  }
}

function slideHasPhaseTokensOutsideTables_(slide) {
  var elements = slide.getPageElements();
  for (var j = 0; j < elements.length; j++) {
    if (elements[j].getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
    try {
      if (containsPhaseToken_(elements[j].asShape().getText().asString())) return true;
    } catch (e) {}
  }
  return false;
}

function safeGetTableCellText_(table, row, col) {
  try { return table.getCell(row, col).getText().asString(); } catch (e) { return ''; }
}

function safeSetTableCellText_(table, row, col, value) {
  try {
    table.getCell(row, col).getText().setText(value);
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    if (msg.indexOf('fusionnée') !== -1 || msg.indexOf('merged') !== -1) return;
    throw e;
  }
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
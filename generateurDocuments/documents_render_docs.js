function replaceTokensInDocument_(doc, variables, phases) {
  var body = getDocumentBodySafe_(doc);
  applyPhaseTokensInDocument_(body, variables || {}, phases || []);
  replaceTokensInDocElement_(body, variables);
  replacePlanningTokenInDocument_(body, variables || {});
  try {
    var header = doc.getHeader();
    if (header) {
      applyPhaseTokensInDocument_(header, variables || {}, phases || []);
      replaceTokensInDocElement_(header, variables);
      replacePlanningTokenInDocument_(header, variables || {});
    }
  } catch (e) {}
  try {
    var footer = doc.getFooter();
    if (footer) {
      applyPhaseTokensInDocument_(footer, variables || {}, phases || []);
      replaceTokensInDocElement_(footer, variables);
      replacePlanningTokenInDocument_(footer, variables || {});
    }
  } catch (e2) {}
}

function getDocumentBodySafe_(docOrBody) {
  if (docOrBody && typeof docOrBody.getBody === 'function') {
    var bodyFromDoc = docOrBody.getBody();
    if (bodyFromDoc && typeof bodyFromDoc.getNumChildren === 'function') return bodyFromDoc;
  }
  if (docOrBody && typeof docOrBody.getNumChildren === 'function' && typeof docOrBody.getTables === 'function') return docOrBody;
  throw new Error('Le moteur de generation attend un Google Doc valide (body introuvable). Verifiez le type de la template utilisee.');
}

function applyPhaseTokensInDocument_(body, baseVariables, phases) {
  if (!body) return;
  var orderedPhases = getSortedPhases_(phases || []);
  if (!orderedPhases.length) {
    try {
      body.replaceText('\\{\\s*phases\\s*\\}', '');
      body.replaceText('\\{\\s*/\\s*phases\\s*\\}', '');
    } catch (e) {}
    return;
  }
  duplicatePhaseRowsInDocumentBody_(body, baseVariables, orderedPhases);
  duplicatePhaseBlocksInDocumentBody_(body, baseVariables, orderedPhases);
  replaceInlinePhaseTokensInDocElement_(body, baseVariables, orderedPhases);
  try {
    body.replaceText('\\{\\s*phases\\s*\\}', '');
    body.replaceText('\\{\\s*/\\s*phases\\s*\\}', '');
  } catch (e2) {}
}

function duplicatePhaseRowsInDocumentBody_(body, baseVariables, phases) {
  if (!body || !Array.isArray(phases) || !phases.length) return;
  var tables = body.getTables();
  for (var t = 0; t < tables.length; t++) duplicatePhaseRowsInDocumentTable_(tables[t], baseVariables, phases);
}

function duplicatePhaseRowsInDocumentTable_(table, baseVariables, phases) {
  var groups = [];
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    if (!containsPhaseToken_(row.getText ? row.getText() : '')) continue;
    if (groups.length && groups[groups.length - 1].end === r - 1) groups[groups.length - 1].end = r;
    else groups.push({ start: r, end: r });
  }

  for (var g = groups.length - 1; g >= 0; g--) {
    var group = groups[g];
    var rowCopies = [];
    for (var rr = group.start; rr <= group.end; rr++) rowCopies.push(table.getRow(rr).copy());
    for (var del = group.end; del >= group.start; del--) table.removeRow(del);
    var insertAt = group.start;
    for (var p = 0; p < phases.length; p++) {
      var phaseVars = buildPhaseVariables_(baseVariables, phases[p], p);
      for (var rc = 0; rc < rowCopies.length; rc++) {
        var insertedRow = table.insertTableRow(insertAt, rowCopies[rc].copy());
        replaceTokensInDocElement_(insertedRow, phaseVars);
        insertAt++;
      }
    }
  }
}

function duplicatePhaseBlocksInDocumentBody_(body, baseVariables, phases) {
  var blocks = [], currentBlock = null;
  for (var i = 0; i < body.getNumChildren(); i++) {
    var child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      if (currentBlock) blocks.push(currentBlock);
      currentBlock = null;
      continue;
    }
    var text = '';
    try { text = child.getText ? child.getText() : ''; } catch (e) {}
    if (containsPhaseToken_(text)) {
      if (!currentBlock) currentBlock = { start: i, end: i };
      else currentBlock.end = i;
    } else if (currentBlock) {
      blocks.push(currentBlock);
      currentBlock = null;
    }
  }
  if (currentBlock) blocks.push(currentBlock);

  for (var b = blocks.length - 1; b >= 0; b--) {
    var block = blocks[b], copies = [];
    for (var ci = block.start; ci <= block.end; ci++) copies.push(body.getChild(ci).copy());
    for (var rem = block.end; rem >= block.start; rem--) body.getChild(rem).removeFromParent();
    var insertIndex = block.start;
    for (var p = 0; p < phases.length; p++) {
      var phaseVars = buildPhaseVariables_(baseVariables, phases[p], p);
      for (var c = 0; c < copies.length; c++) {
        var inserted = insertCopiedDocumentChild_(body, insertIndex, copies[c].copy());
        if (inserted) {
          replaceTokensInDocElement_(inserted, phaseVars);
          insertIndex++;
        }
      }
    }
  }
}

function insertCopiedDocumentChild_(body, childIndex, elementCopy) {
  if (!body || !elementCopy) return null;
  var type = elementCopy.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) return body.insertParagraph(childIndex, elementCopy.asParagraph());
  if (type === DocumentApp.ElementType.LIST_ITEM) return body.insertListItem(childIndex, elementCopy.asListItem());
  if (type === DocumentApp.ElementType.TABLE) return body.insertTable(childIndex, elementCopy.asTable());
  var text = '';
  try { text = elementCopy.getText ? elementCopy.getText() : ''; } catch (e) {}
  return body.insertParagraph(childIndex, text);
}
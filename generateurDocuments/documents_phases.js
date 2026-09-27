function containsPhaseToken_(text) {
  return /\{[^{}]*\bphase\.[^{}]*\}/i.test(String(text || ''));
}

function getSortedPhases_(phases) {
  if (!Array.isArray(phases)) return [];
  return phases
    .map(function(phase, index) {
      return { phase: phase && typeof phase === 'object' ? phase : {}, index: index, sortKey: getPhaseSortKey_(phase, index) };
    })
    .sort(function(a, b) { return a.sortKey !== b.sortKey ? a.sortKey - b.sortKey : a.index - b.index; })
    .map(function(item) { return item.phase; });
}

function getPhaseSortKey_(phase, index) {
  var explicit = parseInt(phase && (phase.numero || phase.num || phase.number || phase.order || phase.ordre), 10);
  if (Number.isFinite(explicit) && explicit > 0) return explicit;
  var start = parseInt(phase && (phase.start || phase.debut || phase.semaineDebut), 10);
  if (Number.isFinite(start) && start > 0) return start + ((index || 0) * 0.0001);
  return (index || 0) + 1;
}

function resolvePhaseNumber_(phase, index) {
  var explicit = parseInt(phase && (phase.numero || phase.num || phase.number || phase.order || phase.ordre), 10);
  return Number.isFinite(explicit) && explicit > 0 ? explicit : (index || 0) + 1;
}

function resolvePhaseTitle_(phase, index) {
  return phase && (phase.title || phase.titre || phase.nom)
    ? String(phase.title || phase.titre || phase.nom)
    : 'Phase ' + resolvePhaseNumber_(phase, index);
}

function resolvePhaseDurationWeeks_(phase) {
  var duration = parseInt(phase && (phase.duration || phase.duree || phase.semaines), 10);
  return Number.isFinite(duration) && duration > 0 ? duration : 1;
}

function resolvePhaseStartWeek_(phase) {
  var start = parseInt(phase && (phase.start || phase.debut || phase.semaineDebut), 10);
  return Number.isFinite(start) && start > 0 ? start : 1;
}

function buildPhaseVariables_(baseVariables, phase, index) {
  var map = {};
  for (var key in baseVariables) if (baseVariables.hasOwnProperty(key)) map[key] = baseVariables[key];

  var pIndex = resolvePhaseNumber_(phase, index);
  var pTitle = resolvePhaseTitle_(phase, index);
  var pDuration = resolvePhaseDurationWeeks_(phase);
  var pJeh = parseNumber_(phase && (phase.jeh || phase.nbJEH || phase.nbJeh || phase.nombreJeh || phase.nombre_jeh));
  var pPrice = parseNumber_(phase && (phase.priceJeh || phase.prixParJEH || phase.prixJeH || phase.prixParJeh || (phase.prix && (phase.prix.parJEH || phase.prix.parJeh || phase.prix.parjeh))));
  var pTotalHt = parseNumber_(phase && phase.prix && (phase.prix.totalHT || phase.prix.totalht));
  if (!pTotalHt) pTotalHt = pJeh * pPrice;

  addVariable_(map, 'phase.nom', pTitle);
  addVariable_(map, 'phase.titre', pTitle);
  addVariable_(map, 'phase.title', pTitle);
  addVariable_(map, 'phase.description', (phase && (phase.description || phase.resume)) || '');
  addVariable_(map, 'phase.numero', pIndex);
  addVariable_(map, 'phase.duree', pDuration);
  addVariable_(map, 'phase.duration', pDuration);
  addVariable_(map, 'phase.nbjeh', pJeh);
  addVariable_(map, 'phase.nombre-jeh', pJeh);
  addVariable_(map, 'phase.nombre_jeh', pJeh);
  addVariable_(map, 'phase.prix.parjeh', pPrice);
  addVariable_(map, 'phase.prix-jeh', pPrice);
  addVariable_(map, 'phase.prix.ht', pTotalHt);
  addVariable_(map, 'phase.prix.totalht', pTotalHt);
  addVariable_(map, 'phase.livrable', (phase && phase.livrable) || '');
  return map;
}
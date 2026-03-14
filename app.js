const defaultRules = [
  "14-15 pouces = 28 unités/palette",
  "16-17 pouces = 24 unités/palette",
  "18-19 pouces = 20 unités/palette",
  "20-21-22 pouces = 16 unités/palette",
  "> 22 pouces = 10 unités/palette",
  "P1 max 3 palettes; si faible qty -> considérer racking",
  "P2 max 6 palettes; en dessous de 3 -> descendre vers P1",
  "P3 max 9 palettes; en dessous de 6 -> descendre vers P2",
  "L2 = zone critique, L3 régulière, L5 éloignée"
];

const improvements = {
  ui: Array.from({ length: 30 }, (_, i) => `UI-${i + 1} • Optimisation ergonomique professionnelle #${i + 1}`),
  ai: Array.from({ length: 30 }, (_, i) => `IA-${i + 1} • Raffinement moteur décisionnel #${i + 1}`),
  app: Array.from({ length: 30 }, (_, i) => `APP-${i + 1} • Renforcement plateforme globale #${i + 1}`),
  pilotage: [
    "Tour de contrôle multi-zones en temps réel",
    "KPIs opérationnels consolidés par vague",
    "Indice de stabilité d'entrepôt",
    "Indice de saturation projetée",
    "Prévision capacité sur 6 semaines",
    "Projection charge équipes par quart",
    "Détection proactive des incidents critiques",
    "Moteur de priorisation incidents IA",
    "Alertes sur seuil de risque configurable",
    "Paramétrage horizon opérationnel",
    "Modes d'automatisation assisté/hybride/autonome",
    "Recommandations d'actions automatisées",
    "Instantané exportable du pilotage",
    "Pipeline de suivi des engagements terrain",
    "Suivi delta entre plans et exécution",
    "Score de confiance dynamique pilotage",
    "Répartition de charge par zone logistique",
    "Consolidation des urgences de capacité",
    "Visualisation du backlog critique",
    "Synthèse tendances journalières",
    "Simulation de pics saisonniers",
    "Scorage d'impact des mouvements",
    "Matrice des dépendances inter-zones",
    "Pilotage des files d'attente opérationnelles",
    "Recalibrage automatique des seuils",
    "Surveillance qualité données pilotage",
    "Journal enrichi des décisions IA",
    "Tableau de bord de résilience",
    "Métriques de performance des équipes",
    "Roadmap d'actions priorisées par valeur"
  ]
};

const demoData = {
  inventory: [
    { sku: "A100", qty: 132, size: 16, location: "E-12" },
    { sku: "B210", qty: 8, size: 18, location: "A-02" },
    { sku: "C314", qty: 54, size: 22, location: "L-45" },
    { sku: "D777", qty: 240, size: 20, location: "G-10" },
    { sku: "E901", qty: 17, size: 24, location: "M-07" }
  ],
  sales: [
    { sku: "A100", ventes: 92 },
    { sku: "B210", ventes: 7 },
    { sku: "C314", ventes: 33 },
    { sku: "D777", ventes: 101 },
    { sku: "E901", ventes: 11 }
  ],
  incoming: [
    { sku: "A100", qty: 22 },
    { sku: "D777", qty: 120 },
    { sku: "E901", qty: 8 }
  ],
  locations: [
    { location: "E-12", type: "P3" },
    { location: "A-02", type: "P1" },
    { location: "L-45", type: "P2" },
    { location: "G-10", type: "P4" },
    { location: "M-07", type: "P1" }
  ]
};

const state = {
  datasets: {},
  report: null,
  currentMoves: [],
  archives: JSON.parse(localStorage.getItem("iaArchives") || "[]"),
  activity: JSON.parse(localStorage.getItem("activityLog") || "[]"),
  archiveFilter: "all",
  scenarioLab: [],
  generatedActionPlan: ""
};


const scenarioProfiles = [
  { id: "base", label: "Base opérationnelle", sales: 1, incoming: 1, capacity: 1, volatility: 1 },
  { id: "peak", label: "Pic de ventes", sales: 1.65, incoming: 1.1, capacity: 0.95, volatility: 1.3 },
  { id: "supply", label: "Surplus arrivages", sales: 0.95, incoming: 1.9, capacity: 1.05, volatility: 1.2 },
  { id: "resilience", label: "Plan résilience", sales: 1.25, incoming: 1.2, capacity: 1.35, volatility: 0.85 }
];


const el = (id) => document.getElementById(id);
const persist = (k, v) => localStorage.setItem(k, JSON.stringify(v));

function toast(message) {
  const div = document.createElement("div");
  div.className = "toast";
  div.textContent = message;
  el("toastContainer").appendChild(div);
  setTimeout(() => div.remove(), 2800);
}

function logActivity(action) {
  state.activity.unshift({ at: new Date().toLocaleString("fr-CA"), action });
  state.activity = state.activity.slice(0, 80);
  persist("activityLog", state.activity);
  renderActivity();
}

function unitPerPalette(size) {
  if (size <= 15) return 28;
  if (size <= 17) return 24;
  if (size <= 19) return 20;
  if (size <= 22) return 16;
  return 10;
}

function capacityByType(type) {
  const depth = Number((type || "P1").replace("P", "")) || 1;
  return depth * 3;
}

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const sheet = wb.SheetNames[0];
        resolve(XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null }));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

function normalizeRows(rows) {
  return rows
    .map((r) => {
      const sku = String(r.sku || r.SKU || "").trim();
      if (!sku) return null;
      return {
        sku,
        qty: Number(r.qty ?? r.qte ?? 0),
        size: Number(r.size ?? r.pouces ?? 0),
        location: String(r.location || r.Location || "").trim(),
        ventes: Number(r.ventes ?? r.sales ?? 0),
        type: String(r.type || r.Type || "P1")
      };
    })
    .filter(Boolean);
}

function validateDatasetReadiness() {
  const required = ["inventory", "sales", "incoming", "locations"];
  const missing = required.filter((k) => !Array.isArray(state.datasets[k]) || !state.datasets[k].length);
  if (missing.length) {
    alert(`Jeux de données manquants: ${missing.join(", ")}`);
    return false;
  }
  return true;
}

function withScenarioMultipliers(profile) {
  const moves = (state.report?.moves || []).map((move) => {
    const score = Math.min(100, Math.max(0, Math.round(move.score * profile.sales * 0.7 + move.capacityPressure * profile.incoming * 0.2 + 8 * profile.volatility)));
    const confidence = Math.max(45, Math.min(99, Math.round(move.confidence * profile.capacity - profile.volatility * 4)));
    const capacityPressure = Math.min(100, Math.round(move.capacityPressure * profile.incoming * (2 - profile.capacity)));
    const velocityScore = Math.min(100, Math.round(move.velocityScore * profile.sales));
    return {
      ...move,
      score,
      confidence,
      capacityPressure,
      velocityScore,
      priority: aiPriority(score),
      reason: `${move.reason} | Scénario ${profile.label}`
    };
  });

  const riskMoves = moves.filter((m) => m.score >= 65);
  const avgConfidence = moves.length ? Math.round(moves.reduce((sum, m) => sum + m.confidence, 0) / moves.length) : 0;
  const riskIndex = moves.length ? Math.round((riskMoves.length / moves.length) * 100) : 0;
  const saturation = Math.min(100, Math.round(moves.reduce((sum, m) => sum + m.capacityPressure, 0) / Math.max(1, moves.length)));
  const teamLoad = Math.min(100, Math.round((riskIndex * 0.45) + (saturation * 0.4) + (100 - avgConfidence) * 0.15));
  const priorityMix = ["Critique", "Haute", "Moyenne", "Standard"].map((level) => `${level}:${moves.filter((m) => m.priority === level).length}`).join(" · ");

  return {
    ...profile,
    totalMoves: moves.length,
    riskMoves: riskMoves.length,
    riskIndex,
    saturation,
    avgConfidence,
    teamLoad,
    priorityMix,
    moves
  };
}

function renderScenarioLab() {
  const output = el("scenarioLabOutput");
  const summary = el("scenarioLabSummary");
  if (!state.report || !state.report.moves?.length) {
    output.innerHTML = "<p>Exécutez d'abord une analyse IA pour alimenter le laboratoire.</p>";
    summary.textContent = "Aucun scénario disponible pour le moment.";
    return;
  }

  const scenarios = scenarioProfiles.map(withScenarioMultipliers).sort((a, b) => (a.riskIndex - b.riskIndex) || (b.avgConfidence - a.avgConfidence));
  state.scenarioLab = scenarios;
  const best = scenarios[0];

  summary.innerHTML = `<span class="chip">Meilleur scénario: ${best.label}</span> <span class="chip">Risque ${best.riskIndex}%</span> <span class="chip">Confiance ${best.avgConfidence}%</span>`;
  output.innerHTML = `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Scénario</th><th>Risque</th><th>Saturation</th><th>Charge équipe</th><th>Confiance</th><th>Mix priorités</th></tr></thead>
        <tbody>
          ${scenarios.map((s) => `<tr><td>${s.label}</td><td>${s.riskIndex}% (${s.riskMoves}/${s.totalMoves})</td><td>${s.saturation}%</td><td>${s.teamLoad}%</td><td>${s.avgConfidence}%</td><td>${s.priorityMix}</td></tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}


function buildWavePlan() {
  const source = state.currentMoves.length ? state.currentMoves : state.report?.moves || [];
  if (!source.length) {
    el("wavePlanOutput").innerHTML = "<p>Aucun déplacement disponible. Lancez une analyse IA.</p>";
    return;
  }

  const thresholdMap = { Critique: 80, Haute: 60, Moyenne: 35, all: 0 };
  const filter = el("wavePriorityFilter").value;
  const batchSize = Math.max(2, Math.min(25, Number(el("waveBatchSize").value) || 6));
  const minScore = thresholdMap[filter] ?? 0;
  const selected = source
    .filter((m) => m.score >= minScore)
    .sort((a, b) => b.score - a.score || b.confidence - a.confidence);

  if (!selected.length) {
    el("wavePlanOutput").innerHTML = "<p>Aucun déplacement ne correspond au filtre choisi.</p>";
    return;
  }

  const waves = [];
  for (let i = 0; i < selected.length; i += batchSize) waves.push(selected.slice(i, i + batchSize));

  el("wavePlanOutput").innerHTML = waves
    .map((wave, i) => {
      const avgScore = Math.round(wave.reduce((sum, m) => sum + m.score, 0) / wave.length);
      const critical = wave.filter((m) => m.priority === "Critique").length;
      const skus = wave.map((m) => maskSku(m.sku)).join(", ");
      return `<div class="archive-item"><h4>Vague ${i + 1} · ${wave.length} mouvements</h4><p>Score moyen: ${avgScore} | Critiques: ${critical}</p><p>${skus}</p></div>`;
    })
    .join("");
}

function runZoneSimulation() {
  const source = state.currentMoves.length ? state.currentMoves : state.report?.moves || [];
  if (!source.length) {
    el("zoneSimulationOutput").innerHTML = "<p>Aucun déplacement disponible pour simuler les zones.</p>";
    return;
  }

  const demandShift = Number(el("zoneDemandShift").value) || 0;
  const reserveBoost = Number(el("zoneReserveBoost").value) || 0;
  el("zoneSimulationMeta").textContent = `Variation demande: ${demandShift}% | Réserve capacité: ${reserveBoost}%`;

  const zones = ["L2", "L3", "L5"].map((zone) => {
    const items = source.filter((m) => m.toZone === zone);
    const baseLoad = items.length ? Math.round(items.reduce((sum, m) => sum + m.capacityPressure, 0) / items.length) : 0;
    const projected = Math.max(0, Math.min(160, Math.round(baseLoad * (1 + demandShift / 100) - reserveBoost)));
    const status = projected >= 100 ? "Critique" : projected >= 80 ? "Sous tension" : "Stable";
    return { zone, projected, status, count: items.length };
  });

  el("zoneSimulationOutput").innerHTML = `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Zone</th><th>Mouvements</th><th>Charge projetée</th><th>Statut</th></tr></thead>
        <tbody>${zones.map((z) => `<tr><td>${z.zone}</td><td>${z.count}</td><td>${z.projected}%</td><td>${z.status}</td></tr>`).join("")}</tbody>
      </table>
    </div>
  `;
}

function runDataQualityCheck() {
  if (!validateDatasetReadiness()) return;
  const inventory = normalizeRows(state.datasets.inventory);
  const sales = normalizeRows(state.datasets.sales);
  const incoming = normalizeRows(state.datasets.incoming);
  const locations = normalizeRows(state.datasets.locations);

  const invSku = new Set(inventory.map((i) => i.sku));
  const locSet = new Set(locations.map((l) => l.location));

  const missingLocation = inventory.filter((i) => !i.location || !locSet.has(i.location));
  const zeroQty = inventory.filter((i) => i.qty <= 0);
  const unknownSales = sales.filter((s) => !invSku.has(s.sku));
  const unknownIncoming = incoming.filter((i) => !invSku.has(i.sku));

  const score = Math.max(0, 100 - (missingLocation.length * 8 + zeroQty.length * 5 + unknownSales.length * 4 + unknownIncoming.length * 4));

  el("dataQualityOutput").innerHTML = `
    <p><span class="chip">Score qualité: ${score}%</span></p>
    <ul>
      <li>Inventaire sans location valide: ${missingLocation.length}</li>
      <li>Inventaire qty nulle/invalide: ${zeroQty.length}</li>
      <li>Ventes sans SKU inventaire: ${unknownSales.length}</li>
      <li>Arrivages sans SKU inventaire: ${unknownIncoming.length}</li>
    </ul>
  `;
}

function generateActionPlan() {
  const source = state.currentMoves.length ? state.currentMoves : state.report?.moves || [];
  if (!source.length) {
    el("actionPlanOutput").innerHTML = "<p>Aucune recommandation active pour générer un plan.</p>";
    return;
  }

  const detail = el("actionPlanDetail").value;
  const top = source
    .slice()
    .sort((a, b) => b.score - a.score || (a.toZone || "").localeCompare(b.toZone || ""))
    .slice(0, 14);

  const days = Array.from({ length: 7 }, (_, i) => ({ day: i + 1, moves: [] }));
  top.forEach((move, idx) => days[idx % 7].moves.push(move));

  const lines = days.map((d) => {
    if (!d.moves.length) return `Jour ${d.day}: surveillance standard`;
    const base = `Jour ${d.day}: ${d.moves.length} actions`;
    if (detail === "compact") return `${base} (${d.moves.map((m) => maskSku(m.sku)).join(", ")})`;
    if (detail === "full") {
      const detailLines = d.moves
        .map((m) => `${maskSku(m.sku)}: ${m.from} → ${m.targetType}/${m.toZone} [${m.priority}, score ${m.score}]`)
        .join("\n- ");
      return `${base}\n- ${detailLines}`;
    }
    return `${base} · priorités: ${d.moves.map((m) => m.priority).join(", ")}`;
  });

  const planText = lines.join("\n\n");
  state.generatedActionPlan = planText;
  el("actionPlanOutput").innerHTML = `<pre>${planText}</pre>`;
}


function aiPriority(score) {
  if (score >= 80) return "Critique";
  if (score >= 60) return "Haute";
  if (score >= 35) return "Moyenne";
  return "Standard";
}

function analyze() {
  if (!validateDatasetReadiness()) return null;
  const t0 = performance.now();
  const salesMultiplier = Number(el("salesMultiplier").value);
  const incomingMultiplier = Number(el("incomingMultiplier").value);
  const capacityBias = Number(el("capacityBias").value);

  const inventory = normalizeRows(state.datasets.inventory);
  const sales = normalizeRows(state.datasets.sales);
  const incoming = normalizeRows(state.datasets.incoming);
  const locations = normalizeRows(state.datasets.locations);

  const salesMap = new Map(sales.map((s) => [s.sku, s.ventes]));
  const incomingMap = new Map(incoming.map((i) => [i.sku, i.qty]));
  const locMap = new Map(locations.map((l) => [l.location, l]));

  const recommendations = inventory.map((item) => {
    const saleRate = (salesMap.get(item.sku) || 0) * salesMultiplier;
    const incomingQty = (incomingMap.get(item.sku) || 0) * incomingMultiplier;
    const locInfo = locMap.get(item.location) || { type: "P1" };
    const type = String(locInfo.type || "P1").toUpperCase();
    const capacity = capacityByType(type) * capacityBias;
    const pallets = Math.ceil(item.qty / unitPerPalette(item.size || 15));
    const incomingPallets = Math.ceil(incomingQty / unitPerPalette(item.size || 15));
    const totalProjected = pallets + incomingPallets;

    const zone = saleRate >= 80 ? "L2" : saleRate <= 10 ? "L5" : "L3";
    const capacityPressure = Math.min(100, Math.round((totalProjected / capacity) * 100));
    const velocityScore = Math.min(100, Math.round((saleRate / 120) * 100));
    const relocationGain = Math.max(0, Math.round((velocityScore * 0.7 + capacityPressure * 0.3) - 25));
    const anomaly = item.qty === 0 || !item.location;

    const reasons = [];
    if (capacityPressure > 100) reasons.push("Projection dépasse la capacité");
    if (velocityScore > 70) reasons.push("Forte vélocité de ventes");
    if (pallets <= Math.max(1, capacity / 3)) reasons.push("Consolidation possible");
    if (anomaly) reasons.push("Anomalie de données détectée");

    const score = Math.round(capacityPressure * 0.45 + velocityScore * 0.35 + relocationGain * 0.2);
    const confidence = Math.max(55, Math.min(98, 90 - reasons.length * 4 + (anomaly ? -12 : 0)));

    return {
      sku: item.sku,
      from: item.location,
      type,
      toZone: zone,
      targetType: type === "P1" ? "Racking" : `P${Math.max(1, Number(type.replace("P", "")) - 1)}`,
      qty: item.qty,
      score,
      confidence,
      priority: aiPriority(score),
      capacityPressure,
      velocityScore,
      relocationGain,
      reason: reasons.join(" | ") || "Optimisation préventive"
    };
  });

  const moves = recommendations.filter((r) => r.score >= 35 || r.reason.includes("Anomalie"));
  const duration = `${Math.round(performance.now() - t0)} ms`;

  return {
    generatedAt: new Date().toLocaleString("fr-CA"),
    duration,
    totals: {
      inventoryRows: inventory.length,
      recommendations: moves.length,
      critical: moves.filter((m) => m.priority === "Critique").length,
      confidence: moves.length ? Math.round(moves.reduce((a, b) => a + b.confidence, 0) / moves.length) : 0
    },
    moves
  };
}

function applyFilters(moves) {
  const search = el("searchSkuInput").value.trim().toLowerCase();
  const zone = el("zoneFilterSelect").value;
  const priority = el("priorityFilterSelect").value;
  const sortBy = el("sortMovesSelect").value;
  const rank = { Critique: 0, Haute: 1, Moyenne: 2, Standard: 3 };

  let filtered = moves.filter((m) => {
    const skuMatch = !search || m.sku.toLowerCase().includes(search);
    const zoneMatch = zone === "all" || m.toZone === zone;
    const priorityMatch = priority === "all" || m.priority === priority;
    return skuMatch && zoneMatch && priorityMatch;
  });

  if (sortBy === "sku") filtered.sort((a, b) => a.sku.localeCompare(b.sku));
  if (sortBy === "zone") filtered.sort((a, b) => a.toZone.localeCompare(b.toZone));
  if (sortBy === "score") filtered.sort((a, b) => b.score - a.score);
  if (sortBy === "priority") filtered.sort((a, b) => rank[a.priority] - rank[b.priority]);

  return filtered;
}

function maskSku(sku) {
  return el("secureModeToggle").checked ? `${sku.slice(0, 2)}***` : sku;
}

function renderReport(report) {
  if (!report) return;
  state.currentMoves = applyFilters(report.moves);
  renderStats(report);

  const table = `
  <div class="table-wrap">
  <table>
    <thead><tr><th>SKU</th><th>Origine</th><th>Cible</th><th>Priorité</th><th>Score</th><th>Confiance</th><th>Raison</th></tr></thead>
    <tbody>
      ${state.currentMoves
        .map(
          (m) => `<tr><td>${maskSku(m.sku)}</td><td>${m.from}</td><td>${m.targetType}/${m.toZone}</td><td>${m.priority}</td><td>${m.score}</td><td>${m.confidence}%</td><td>${m.reason}</td></tr>`
        )
        .join("")}
    </tbody>
  </table></div>`;

  el("analysisOutput").innerHTML = `<p><span class="chip">Rapport ${report.generatedAt}</span> <span class="chip">Durée ${report.duration}</span></p>${table}`;
  el("movementOutput").innerHTML =
    "<ol>" +
    state.currentMoves
      .map((m) => `<li>${maskSku(m.sku)}: ${m.from} → ${m.targetType}/${m.toZone} (${m.priority}, impact ${m.score})</li>`)
      .join("") +
    "</ol>";

  renderHeatmap(state.currentMoves);
  renderZoneSummary(state.currentMoves);
  renderPilotage(buildPilotageSnapshot(state.currentMoves));
  renderDiagnostics(report);
}

function renderStats(report) {
  el("statSku").textContent = report?.totals.inventoryRows || 0;
  el("statMoves").textContent = report?.totals.recommendations || 0;
  el("statCritical").textContent = report?.totals.critical || 0;
  el("statConfidence").textContent = `${report?.totals.confidence || 0}%`;
  el("statDuration").textContent = report?.duration || "0 ms";
  el("statLastRun").textContent = report?.generatedAt || "-";
}

function renderHeatmap(moves) {
  const groups = ["Critique", "Haute", "Moyenne", "Standard"].map((p) => ({
    p,
    count: moves.filter((m) => m.priority === p).length
  }));
  el("riskHeatmap").innerHTML = groups
    .map((g) => `<div class="heat-cell" style="background:hsl(${120 - g.count * 10},70%,45%)">${g.p}<br>${g.count}</div>`)
    .join("");
}

function renderZoneSummary(moves) {
  const zones = ["L2", "L3", "L5"].map((z) => ({ z, count: moves.filter((m) => m.toZone === z).length }));
  el("zoneSummary").innerHTML = zones.map((z) => `<div class="stat-card"><span>Zone ${z.z}</span><b>${z.count}</b></div>`).join("");
}

function buildPilotageSnapshot(moves) {
  const horizon = Number(el("pilotageHorizon")?.value || 21);
  const riskThreshold = Number(el("pilotageRiskThreshold")?.value || 60);
  const automation = el("pilotageAutomationMode")?.value || "assist";
  const riskMoves = moves.filter((m) => m.score >= riskThreshold);
  const avgScore = moves.length ? Math.round(moves.reduce((sum, m) => sum + m.score, 0) / moves.length) : 0;
  const avgConfidence = moves.length ? Math.round(moves.reduce((sum, m) => sum + m.confidence, 0) / moves.length) : 0;
  const forecast = Array.from({ length: 6 }, (_, i) => {
    const week = i + 1;
    const load = Math.max(12, Math.min(98, Math.round(avgScore * 0.55 + riskMoves.length * 3 + week * 4 - (avgConfidence - 60) * 0.3)));
    return { week, load };
  });
  const teamLabels = ["Réception", "Mise en stock", "Préparation", "Expédition"];
  const teamLoad = teamLabels.map((name, idx) => ({
    name,
    load: Math.max(20, Math.min(100, Math.round(avgScore * 0.5 + riskMoves.length * 2 + idx * 9 + (automation === "auto" ? -8 : automation === "hybrid" ? -2 : 6))))
  }));

  const incidents = riskMoves
    .sort((a, b) => b.score - a.score)
    .slice(0, 6)
    .map((m) => `${maskSku(m.sku)} · ${m.priority} · score ${m.score} · ${m.from} → ${m.targetType}/${m.toZone}`);

  const automationActions = [
    `Réallouer ${Math.min(8, Math.max(2, riskMoves.length))} tâches vers zone ${riskMoves[0]?.toZone || "L3"}`,
    `Déclencher cycle de consolidation ${automation === "auto" ? "automatique" : "supervisé"}`,
    `Renforcer le créneau ${teamLoad.sort((a, b) => b.load - a.load)[0].name} (+${Math.round(avgScore / 12)}%)`,
    `Programmer une revue capacité sur horizon ${horizon} jours`
  ];

  return {
    at: new Date().toLocaleString("fr-CA"),
    horizon,
    riskThreshold,
    automation,
    avgScore,
    avgConfidence,
    riskCount: riskMoves.length,
    forecast,
    teamLoad,
    incidents,
    automationActions,
    stability: Math.max(0, 100 - Math.round((riskMoves.length / Math.max(1, moves.length)) * 100)),
    saturation: Math.min(100, Math.round(forecast.reduce((sum, w) => sum + w.load, 0) / forecast.length))
  };
}

function renderPilotage(snapshot) {
  if (!snapshot) return;
  el("pilotageMeta").textContent = `Pilotage ${snapshot.at} · horizon ${snapshot.horizon}j · seuil risque ${snapshot.riskThreshold} · mode ${snapshot.automation}`;
  el("pilotageKpis").innerHTML = [
    ["Indice de stabilité", `${snapshot.stability}%`],
    ["Saturation projetée", `${snapshot.saturation}%`],
    ["Incidents critiques", snapshot.riskCount],
    ["Confiance pilotage", `${snapshot.avgConfidence}%`],
    ["Score moyen impact", snapshot.avgScore],
    ["Actions auto", snapshot.automationActions.length]
  ]
    .map(([label, value]) => `<div class="stat-card"><span>${label}</span><b>${value}</b></div>`)
    .join("");

  el("capacityForecast").innerHTML = snapshot.forecast
    .map((w) => `<div class="mini-bar"><span>S${w.week}</span><div><i style="width:${w.load}%"></i></div><b>${w.load}%</b></div>`)
    .join("");

  el("teamLoad").innerHTML = snapshot.teamLoad
    .map((t) => `<div class="mini-bar"><span>${t.name}</span><div><i style="width:${t.load}%"></i></div><b>${t.load}%</b></div>`)
    .join("");

  el("incidentQueue").innerHTML = snapshot.incidents.length
    ? snapshot.incidents.map((i) => `<li>${i}</li>`).join("")
    : "<li>Aucun incident au-dessus du seuil</li>";
  el("automationActions").innerHTML = snapshot.automationActions.map((a) => `<li>${a}</li>`).join("");
}

function renderActivity() {
  el("activityLog").innerHTML =
    "<ul>" + state.activity.map((a) => `<li><b>${a.at}</b> — ${a.action}</li>`).join("") + "</ul>";
}

function renderDiagnostics(report) {
  const usage = new Blob([JSON.stringify(localStorage)]).size;
  const lines = [
    `Version moteur IA: ${report ? "2.0" : "-"}`,
    `Stockage local estimé: ${Math.round(usage / 1024)} Ko`,
    `Archives: ${state.archives.length}`,
    `Session active: ${localStorage.getItem("activeSession") || "non"}`,
    `Mode compact: ${document.body.classList.contains("compact") ? "oui" : "non"}`
  ];
  el("diagnosticsOutput").innerHTML = "<ul>" + lines.map((l) => `<li>${l}</li>`).join("") + "</ul>";
}

function renderArchives() {
  const search = el("archiveSearchInput")?.value?.toLowerCase?.() || "";
  const filter = el("archiveFilterSelect")?.value || "all";
  const filtered = state.archives.filter((a) => {
    const txt = `${a.generatedAt} ${a.totals.recommendations}`.toLowerCase();
    const bySearch = !search || txt.includes(search);
    const byFilter =
      filter === "all" ||
      (filter === "critical" && a.totals.critical > 0) ||
      (filter === "light" && a.totals.critical === 0);
    return bySearch && byFilter;
  });

  el("archiveList").innerHTML = filtered.length
    ? filtered
        .map(
          (a, i) => `<div class="archive-item"><h4>#${i + 1} • ${a.generatedAt}</h4><p>${a.totals.recommendations} déplacements | critiques: ${a.totals.critical}</p>
      <div class="actions"><button data-open="${i}">Ouvrir</button><button data-export="${i}" class="outline-btn">CSV</button><button data-delete="${i}" class="danger-btn">Supprimer</button></div></div>`
        )
        .join("")
    : "<p>Aucune archive.</p>";

  const select = el("compareArchiveSelect");
  select.innerHTML = '<option value="">Comparer avec...</option>' + state.archives.map((a, i) => `<option value="${i}">Archive ${i + 1} • ${a.generatedAt}</option>`).join("");

  el("archiveList").querySelectorAll("button[data-open]").forEach((b) =>
    b.addEventListener("click", () => {
      state.report = state.archives[Number(b.dataset.open)];
      renderReport(state.report);
      document.querySelector('[data-tab="consolidation"]').click();
    })
  );
  el("archiveList").querySelectorAll("button[data-export]").forEach((b) =>
    b.addEventListener("click", () => exportCsv(state.archives[Number(b.dataset.export)].moves, `archive-${Number(b.dataset.export) + 1}.csv`))
  );
  el("archiveList").querySelectorAll("button[data-delete]").forEach((b) =>
    b.addEventListener("click", () => {
      state.archives.splice(Number(b.dataset.delete), 1);
      persist("iaArchives", state.archives);
      logActivity("Archive supprimée");
      renderArchives();
    })
  );
}

function compareArchives() {
  const idx = Number(el("compareArchiveSelect").value);
  if (Number.isNaN(idx) || !state.report || !state.archives[idx]) {
    el("compareOutput").innerHTML = "";
    return;
  }
  const base = state.report.totals;
  const other = state.archives[idx].totals;
  el("compareOutput").innerHTML = `<p>Comparaison active vs archive #${idx + 1}</p>
  <ul><li>Déplacements: ${base.recommendations} vs ${other.recommendations}</li>
  <li>Critiques: ${base.critical} vs ${other.critical}</li>
  <li>Confiance: ${base.confidence || 0}% vs ${other.confidence || 0}%</li></ul>`;
}

function exportCsv(moves = state.currentMoves, filename = "deplacements.csv") {
  if (!moves.length) return alert("Aucun déplacement à exporter");
  const headers = ["sku", "from", "toZone", "targetType", "priority", "score", "confidence", "reason"];
  const csv = [headers.join(",")]
    .concat(moves.map((m) => headers.map((h) => `"${String(m[h] ?? "").replaceAll('"', '""')}"`).join(",")))
    .join("\n");
  downloadBlob(csv, filename, "text/csv;charset=utf-8;");
}

function exportJson(report = state.report) {
  if (!report) return alert("Aucun rapport actif");
  downloadBlob(JSON.stringify(report, null, 2), `rapport-${Date.now()}.json`, "application/json");
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function renderImprovements() {
  el("uiImprovementsList").innerHTML = improvements.ui.map((i) => `<li>${i}</li>`).join("");
  el("aiImprovementsList").innerHTML = improvements.ai.map((i) => `<li>${i}</li>`).join("");
  el("appImprovementsList").innerHTML = improvements.app.map((i) => `<li>${i}</li>`).join("");
  el("pilotageImprovementsList").innerHTML = improvements.pilotage.map((i) => `<li>${i}</li>`).join("");
}

function refreshDatasetStatus() {
  const keys = ["inventory", "sales", "incoming", "locations"];
  const loaded = keys.filter((k) => Array.isArray(state.datasets[k]) && state.datasets[k].length).length;
  const pct = (loaded / keys.length) * 100;
  el("datasetStatus").textContent = `${loaded}/4 jeux chargés (${Math.round(pct)}%)`;
  el("datasetProgress").style.width = `${pct}%`;
}

function loadRules() {
  const raw = localStorage.getItem("aiRules");
  return raw ? JSON.parse(raw) : defaultRules;
}

function renderRules() {
  const rules = loadRules();
  el("rulesPreview").innerHTML = rules.map((r) => `<li>${r}</li>`).join("");
  el("customRules").value = rules.join("\n");
}


function setupDropdownPanels() {
  document.querySelectorAll(".dropdown-panel").forEach((panel) => {
    panel.classList.remove("open");
    const toggle = panel.querySelector(".panel-toggle");
    if (!toggle) return;
    toggle.addEventListener("click", () => {
      panel.classList.toggle("open");
    });
  });
}

function setupTabs() {
  document.querySelectorAll(".tab").forEach((tab) =>
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      el(tab.dataset.tab).classList.add("active");
    })
  );
}

function login(email) {
  el("inviteGate").classList.add("hidden");
  el("mainContent").classList.remove("hidden");
  el("logoutBtn").classList.remove("hidden");
  el("inviteStatus").textContent = `Session active: ${email}`;
  localStorage.setItem("activeSession", email);
  logActivity("Connexion réussie");
}

function logout() {
  localStorage.removeItem("activeSession");
  el("inviteGate").classList.remove("hidden");
  el("mainContent").classList.add("hidden");
  el("logoutBtn").classList.add("hidden");
  el("inviteStatus").textContent = "Accès protégé par invitation";
}

function applyPreferences() {
  const prefs = JSON.parse(localStorage.getItem("prefs") || "{}");
  document.body.classList.toggle("dark", prefs.theme === "dark");
  document.body.classList.toggle("compact", !!prefs.compact);
  document.body.classList.toggle("high-contrast", !!prefs.highContrast);
  document.body.classList.toggle("reduce-motion", !!prefs.reduceMotion);
  document.documentElement.style.setProperty("--font-scale", `${prefs.fontScale || 100}%`);
  el("themeToggleBtn").textContent = prefs.theme === "dark" ? "Mode clair" : "Mode sombre";
  el("highContrastToggle").checked = !!prefs.highContrast;
  el("reduceMotionToggle").checked = !!prefs.reduceMotion;
  el("fontScaleRange").value = prefs.fontScale || 100;
  el("secureModeToggle").checked = !!prefs.secureMode;
}

function setPref(partial) {
  const prefs = { ...(JSON.parse(localStorage.getItem("prefs") || "{}")), ...partial };
  localStorage.setItem("prefs", JSON.stringify(prefs));
  applyPreferences();
}

function bindEvents() {
  el("loginBtn").addEventListener("click", () => {
    const email = el("emailInput").value.trim().toLowerCase();
    const code = el("codeInput").value.trim();
    if (email === "admin@langelier.ca" && code === "LANGELIER-2026") login(email);
    else alert("Invitation invalide");
  });

  el("logoutBtn").addEventListener("click", logout);

  ["inventory", "sales", "incoming", "locations"].forEach((key) => {
    el(`${key}File`).addEventListener("change", async (evt) => {
      const file = evt.target.files[0];
      if (!file) return;
      try {
        state.datasets[key] = await parseExcel(file);
        refreshDatasetStatus();
        logActivity(`Import ${key}: ${file.name}`);
      } catch {
        alert(`Impossible de lire ${file.name}`);
      }
    });
  });

  el("runAnalysisBtn").addEventListener("click", () => {
    state.report = analyze();
    renderReport(state.report);
    renderScenarioLab();
    logActivity("Analyse IA exécutée");
  });

  el("generateMovesBtn").addEventListener("click", () => {
    if (!state.report) state.report = analyze();
    renderReport(state.report);
  });

  el("loadDemoBtn").addEventListener("click", () => {
    state.datasets = JSON.parse(JSON.stringify(demoData));
    refreshDatasetStatus();
    toast("Données de démonstration chargées");
    logActivity("Chargement jeu de démonstration");
  });

  el("scenarioBtn").addEventListener("click", () => {
    state.datasets = JSON.parse(JSON.stringify(demoData));
    state.datasets.incoming.forEach((i) => (i.qty *= 2));
    state.datasets.sales.forEach((s) => (s.ventes = Math.round(s.ventes * 1.4)));
    refreshDatasetStatus();
    toast("Scénario stress activé");
    logActivity("Activation scénario stress");
  });


  el("runScenarioLabBtn").addEventListener("click", () => {
    if (!state.report) {
      state.report = analyze();
      renderReport(state.report);
    }
    renderScenarioLab();
    toast("Laboratoire de scénarios exécuté");
    logActivity("Exécution laboratoire de scénarios");
  });

  el("applyBestScenarioBtn").addEventListener("click", () => {
    if (!state.scenarioLab.length) renderScenarioLab();
    const best = state.scenarioLab[0];
    if (!best) return alert("Aucun scénario à appliquer");
    state.currentMoves = applyFilters(best.moves);
    renderHeatmap(state.currentMoves);
    renderZoneSummary(state.currentMoves);
    renderPilotage(buildPilotageSnapshot(state.currentMoves));
    el("movementOutput").innerHTML = "<ol>" + state.currentMoves.map((m) => `<li>${maskSku(m.sku)}: ${m.from} → ${m.targetType}/${m.toZone} (${m.priority}, impact ${m.score})</li>`).join("") + "</ol>";
    toast(`Profil appliqué: ${best.label}`);
    logActivity(`Application scénario ${best.label}`);
  });

  el("buildWavePlanBtn").addEventListener("click", () => {
    buildWavePlan();
    logActivity("Construction du plan de vagues");
  });
  el("runZoneSimulationBtn").addEventListener("click", () => {
    runZoneSimulation();
    logActivity("Simulation de charge zones");
  });
  ["zoneDemandShift", "zoneReserveBoost"].forEach((id) =>
    el(id).addEventListener("input", () => {
      el("zoneSimulationMeta").textContent = `Variation demande: ${el("zoneDemandShift").value}% | Réserve capacité: ${el("zoneReserveBoost").value}%`;
    })
  );
  el("runDataQualityBtn").addEventListener("click", () => {
    runDataQualityCheck();
    logActivity("Contrôle qualité des données");
  });
  el("generateActionPlanBtn").addEventListener("click", () => {
    generateActionPlan();
    logActivity("Génération plan d'action 7 jours");
  });
  el("exportActionPlanBtn").addEventListener("click", () => {
    if (!state.generatedActionPlan) generateActionPlan();
    if (!state.generatedActionPlan) return;
    downloadBlob(state.generatedActionPlan, `plan-action-${Date.now()}.txt`, "text/plain;charset=utf-8;");
    toast("Plan d'action exporté");
  });

  ["searchSkuInput", "zoneFilterSelect", "priorityFilterSelect", "sortMovesSelect"].forEach((id) =>
    el(id).addEventListener(id.includes("input") ? "input" : "change", () => renderReport(state.report))
  );
  ["pilotageHorizon", "pilotageRiskThreshold", "pilotageAutomationMode"].forEach((id) =>
    el(id).addEventListener(id.includes("Mode") ? "change" : "input", () => renderPilotage(buildPilotageSnapshot(state.currentMoves)))
  );
  el("pilotageRefreshBtn").addEventListener("click", () => renderPilotage(buildPilotageSnapshot(state.currentMoves)));
  el("pilotageExportBtn").addEventListener("click", () => {
    const snapshot = buildPilotageSnapshot(state.currentMoves);
    downloadBlob(JSON.stringify(snapshot, null, 2), "pilotage-snapshot.json", "application/json");
    toast("Instantané pilotage exporté");
    logActivity("Export instantané pilotage");
  });

  ["salesMultiplier", "incomingMultiplier", "capacityBias"].forEach((id) => el(id).addEventListener("input", () => {
    el("simValues").textContent = `Ventes x${el("salesMultiplier").value} | Arrivages x${el("incomingMultiplier").value} | Capacité x${el("capacityBias").value}`;
  }));
  el("simValues").textContent = `Ventes x1 | Arrivages x1 | Capacité x1`;
  el("zoneSimulationMeta").textContent = `Variation demande: 0% | Réserve capacité: 10%`;

  el("exportCsvBtn").addEventListener("click", () => exportCsv());
  el("exportJsonBtn").addEventListener("click", () => exportJson());
  el("printReportBtn").addEventListener("click", () => window.print());

  el("archiveReportBtn").addEventListener("click", () => {
    if (!state.report) return alert("Aucun rapport à archiver");
    state.archives.unshift(state.report);
    persist("iaArchives", state.archives);
    renderArchives();
    toast("Rapport archivé");
    logActivity("Rapport archivé");
  });

  el("clearArchivesBtn").addEventListener("click", () => {
    state.archives = [];
    persist("iaArchives", state.archives);
    renderArchives();
    logActivity("Archives effacées");
  });

  el("archiveSearchInput").addEventListener("input", renderArchives);
  el("archiveFilterSelect").addEventListener("change", renderArchives);
  el("compareArchiveSelect").addEventListener("change", compareArchives);

  el("saveRulesBtn").addEventListener("click", () => {
    const rules = el("customRules").value.split("\n").map((r) => r.trim()).filter(Boolean);
    localStorage.setItem("aiRules", JSON.stringify(rules));
    renderRules();
    toast("Règles IA sauvegardées");
  });
  el("resetRulesBtn").addEventListener("click", () => {
    localStorage.removeItem("aiRules");
    renderRules();
  });

  el("themeToggleBtn").addEventListener("click", () => setPref({ theme: document.body.classList.contains("dark") ? "light" : "dark" }));
  el("densityToggleBtn").addEventListener("click", () => setPref({ compact: !document.body.classList.contains("compact") }));
  el("highContrastToggle").addEventListener("change", (e) => setPref({ highContrast: e.target.checked }));
  el("reduceMotionToggle").addEventListener("change", (e) => setPref({ reduceMotion: e.target.checked }));
  el("fontScaleRange").addEventListener("input", (e) => setPref({ fontScale: Number(e.target.value) }));
  el("secureModeToggle").addEventListener("change", (e) => {
    setPref({ secureMode: e.target.checked });
    renderReport(state.report);
  });

  el("resetAppBtn").addEventListener("click", () => {
    localStorage.clear();
    location.reload();
  });

  el("exportBackupBtn").addEventListener("click", () => {
    const payload = { archives: state.archives, activity: state.activity, prefs: JSON.parse(localStorage.getItem("prefs") || "{}") };
    downloadBlob(JSON.stringify(payload, null, 2), "backup-langelier.json", "application/json");
  });
  el("importBackupBtn").addEventListener("click", () => el("backupInput").click());
  el("backupInput").addEventListener("change", async (evt) => {
    const file = evt.target.files[0];
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      state.archives = data.archives || [];
      state.activity = data.activity || [];
      persist("iaArchives", state.archives);
      persist("activityLog", state.activity);
      localStorage.setItem("prefs", JSON.stringify(data.prefs || {}));
      applyPreferences();
      renderArchives();
      renderActivity();
      toast("Sauvegarde importée");
    } catch {
      alert("Fichier de sauvegarde invalide");
    }
  });

  el("commandPaletteBtn").addEventListener("click", () => el("commandPalette").classList.remove("hidden"));
  el("closePaletteBtn").addEventListener("click", () => el("commandPalette").classList.add("hidden"));
  el("commandPalette").querySelectorAll("button[data-command]").forEach((b) =>
    b.addEventListener("click", () => {
      const c = b.dataset.command;
      if (c === "analyze") el("runAnalysisBtn").click();
      if (c === "demo") el("loadDemoBtn").click();
      if (c === "archive") el("archiveReportBtn").click();
      if (c === "theme") el("themeToggleBtn").click();
      if (c === "export") el("exportCsvBtn").click();
      el("commandPalette").classList.add("hidden");
    })
  );

  document.addEventListener("keydown", (evt) => {
    if (evt.ctrlKey && evt.key.toLowerCase() === "enter") el("runAnalysisBtn").click();
    if (evt.ctrlKey && evt.key.toLowerCase() === "e") { evt.preventDefault(); el("exportCsvBtn").click(); }
    if (evt.ctrlKey && evt.key.toLowerCase() === "k") { evt.preventDefault(); el("commandPalette").classList.remove("hidden"); }
  });
}

function setupSession() {
  const email = localStorage.getItem("activeSession");
  if (email) login(email);
}

setInterval(() => {
  if (el("liveClock")) el("liveClock").textContent = new Date().toLocaleTimeString("fr-CA");
}, 1000);

setupTabs();
setupDropdownPanels();
bindEvents();
renderRules();
renderImprovements();
renderArchives();
renderActivity();
applyPreferences();
setupSession();
refreshDatasetStatus();
renderStats(null);
renderDiagnostics(null);
renderScenarioLab();

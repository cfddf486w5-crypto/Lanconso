const defaultRules = [
  "14-15 pouces = 28 unités/palette",
  "16-17 pouces = 24 unités/palette",
  "18-19 pouces = 20 unités/palette",
  "20-21-22 pouces = 16 unités/palette",
  "> 22 pouces = 10 unités/palette",
  "P1 = 1 profondeur, max 3 palettes; si <10 unités d'un article => analyser déplacement racking",
  "P2 = max 6 palettes; si 3 palettes ou moins => déplacer en P1",
  "P3 = max 9 palettes; si 6 palettes ou moins => peut devenir P2 s'il y a de la place",
  "Si P2 contient moins de 3 palettes => doit devenir P1",
  "Étendre la logique jusqu'à P7 (max profondeur)",
  "L2 = zone critique (meilleurs vendeurs), L3 = pige régulière, L5 = éloigné",
  "L2H/L2J = zones magasin distinctes"
];

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
  files: {},
  datasets: {},
  report: null,
  currentMoves: [],
  archiveFilter: "all",
  archives: JSON.parse(localStorage.getItem("iaArchives") || "[]")
};

const el = (id) => document.getElementById(id);
const persist = (key, value) => localStorage.setItem(key, JSON.stringify(value));

function loadRules() {
  const raw = localStorage.getItem("aiRules");
  return raw ? JSON.parse(raw) : defaultRules;
}

function getTheme() {
  return localStorage.getItem("themeMode") || "light";
}

function setTheme(mode) {
  document.body.classList.toggle("dark", mode === "dark");
  localStorage.setItem("themeMode", mode);
  el("themeToggleBtn").textContent = mode === "dark" ? "Mode clair" : "Mode sombre";
}

function refreshDatasetStatus() {
  const keys = ["inventory", "sales", "incoming", "locations"];
  const loaded = keys.filter((k) => Array.isArray(state.datasets[k]) && state.datasets[k].length).length;
  el("datasetStatus").textContent = `${loaded}/4 jeux de données chargés.`;
}

function renderRulesPreview() {
  const rules = loadRules();
  el("rulesPreview").innerHTML = rules.map((r) => `<li>${r}</li>`).join("");
  el("customRules").value = rules.join("\n");
}

function setupTabs() {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      el(tab.dataset.tab).classList.add("active");
    });
  });
}

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

function unitPerPalette(size) {
  if (size <= 15) return 28;
  if (size <= 17) return 24;
  if (size <= 19) return 20;
  if (size <= 22) return 16;
  return 10;
}

function capacityByType(type) {
  const depth = Number((type || "P1").replace("P", ""));
  return Math.max(1, depth) * 3;
}

function normalizeRows(rows) {
  return rows
    .map((row) => {
      const sku = String(row.sku || row.SKU || "").trim();
      if (!sku) return null;
      return {
        ...row,
        sku,
        qty: Number(row.qty ?? row.qte ?? 0),
        size: Number(row.size ?? row.pouces ?? 0),
        location: String(row.location || row.Location || "").trim()
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

function priorityFromReasons(reasons) {
  if (reasons.some((r) => r.includes("dépasse capacité"))) return "Critique";
  if (reasons.some((r) => r.includes("Meilleur vendeur"))) return "Haute";
  if (reasons.some((r) => r.includes("Occupation faible"))) return "Moyenne";
  return "Standard";
}

function analyze() {
  if (!validateDatasetReadiness()) return null;

  const inventory = normalizeRows(state.datasets.inventory || []);
  const sales = normalizeRows(state.datasets.sales || []);
  const incoming = normalizeRows(state.datasets.incoming || []);
  const locations = normalizeRows(state.datasets.locations || []);

  const salesMap = new Map(sales.map((s) => [String(s.sku || s.SKU), Number(s.ventes || s.sales || 0)]));
  const incomingMap = new Map(incoming.map((i) => [String(i.sku || i.SKU), Number(i.qty || i.qte || 0)]));
  const locationMap = new Map(locations.map((l) => [String(l.location || l.Location), l]));

  const recommendations = inventory.map((item) => {
    const sku = String(item.sku || item.SKU);
    const qty = Number(item.qty || item.qte || 0);
    const size = Number(item.size || item.pouces || 0);
    const locCode = String(item.location || item.Location || "");
    const locInfo = locationMap.get(locCode) || {};
    const type = String(locInfo.type || locInfo.Type || "P1").toUpperCase();
    const pallets = Math.ceil(qty / unitPerPalette(size || 15));
    const cap = capacityByType(type);
    const saleRate = salesMap.get(sku) || 0;
    const incomingQty = incomingMap.get(sku) || 0;

    let zone = "L3";
    if (saleRate >= 80) zone = "L2";
    else if (saleRate <= 10) zone = "L5";

    const reasons = [];
    if (type === "P1" && qty < 10) reasons.push("P1 < 10 unités: analyser déplacement racking");

    const threshold = Math.max(1, cap - 3);
    if (pallets <= threshold && type !== "P1") {
      const targetDepth = Math.max(1, Number(type.replace("P", "")) - 1);
      reasons.push(`Occupation faible (${pallets}/${cap}): réduire vers P${targetDepth}`);
    }
    if (incomingQty > 0 && pallets + Math.ceil(incomingQty / unitPerPalette(size || 15)) > cap) {
      reasons.push("Réception attendue dépasse capacité actuelle");
    }
    if (zone === "L2") reasons.push("Meilleur vendeur: prioriser zone critique L2");

    return {
      sku,
      qty,
      size,
      location: locCode,
      type,
      pallets,
      cap,
      zone,
      reasons,
      priority: priorityFromReasons(reasons)
    };
  });

  const toMove = recommendations.filter((r) => r.reasons.length);
  return {
    generatedAt: new Date().toLocaleString("fr-CA"),
    totals: {
      inventoryRows: inventory.length,
      recommendations: toMove.length,
      critical: toMove.filter((m) => m.priority === "Critique").length
    },
    recommendations,
    moves: toMove.map((r) => ({
      sku: r.sku,
      from: r.location,
      toZone: r.zone,
      priority: r.priority,
      targetType: r.type === "P1" ? "Racking" : `P${Math.max(1, Number(r.type.replace("P", "")) - 1)}`,
      reason: r.reasons.join(" | ")
    }))
  };
}

function applyMoveFilters(moves) {
  const search = el("searchSkuInput").value.trim().toLowerCase();
  const zone = el("zoneFilterSelect").value;
  const sortBy = el("sortMovesSelect").value;

  let filtered = moves.filter((m) => {
    const skuMatch = !search || m.sku.toLowerCase().includes(search);
    const zoneMatch = zone === "all" || m.toZone === zone;
    return skuMatch && zoneMatch;
  });

  if (sortBy === "sku") filtered = filtered.sort((a, b) => a.sku.localeCompare(b.sku));
  if (sortBy === "zone") filtered = filtered.sort((a, b) => a.toZone.localeCompare(b.toZone));
  if (sortBy === "priority") {
    const rank = { Critique: 0, Haute: 1, Moyenne: 2, Standard: 3 };
    filtered = filtered.sort((a, b) => rank[a.priority] - rank[b.priority]);
  }
  return filtered;
}

function renderStats(report) {
  el("statSku").textContent = report?.totals?.inventoryRows || 0;
  el("statMoves").textContent = report?.totals?.recommendations || 0;
  el("statCritical").textContent = report?.totals?.critical || 0;
  el("statLastRun").textContent = report?.generatedAt || "-";
}

function renderReport(report) {
  if (!report) return;
  state.currentMoves = applyMoveFilters(report.moves);
  const lines = [
    `<p><span class="badge">Rapport IA</span><span class="badge">${report.generatedAt}</span></p>`,
    `<p>Lignes inventaire: <b>${report.totals.inventoryRows}</b> • Recommandations: <b>${report.totals.recommendations}</b> • Critiques: <b>${report.totals.critical}</b></p>`,
    "<h3>Raisons des déplacements</h3>",
    "<ul>" +
      state.currentMoves
        .map((m) => `<li><b>${m.sku}</b> (${m.from} → ${m.targetType}/${m.toZone}) [${m.priority}]: ${m.reason}</li>`)
        .join("") +
      "</ul>"
  ];
  el("analysisOutput").innerHTML = lines.join("");
  el("movementOutput").innerHTML =
    "<ol>" +
    state.currentMoves
      .map((m) => `<li>${m.sku}: Déplacer ${m.from} vers ${m.targetType} (${m.toZone}) — priorité ${m.priority}</li>`)
      .join("") +
    "</ol>";
  renderStats(report);
}

function renderArchives() {
  const list = el("archiveList");
  if (!state.archives.length) {
    list.innerHTML = "<p>Aucun rapport archivé pour le moment.</p>";
    return;
  }

  list.innerHTML = state.archives
    .map(
      (a, i) => `<div class="archive-item"><h4>#${i + 1} • ${a.generatedAt}</h4>
      <p>${a.totals.recommendations} déplacements recommandés (${a.totals.critical || 0} critiques).</p>
      <div class="actions">
        <button data-archive="${i}">Ouvrir</button>
        <button data-print="${i}">Imprimer</button>
        <button class="outline-btn" data-export="${i}">Exporter CSV</button>
        <button class="danger-btn" data-delete="${i}">Supprimer</button>
      </div>
    </div>`
    )
    .join("");

  list.querySelectorAll("button[data-archive]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.archive);
      state.report = state.archives[idx];
      renderReport(state.report);
      document.querySelector('[data-tab="consolidation"]').click();
    });
  });

  list.querySelectorAll("button[data-print]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.print);
      state.report = state.archives[idx];
      renderReport(state.report);
      window.print();
    });
  });

  list.querySelectorAll("button[data-export]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.export);
      exportMovesCsv(state.archives[idx].moves, `archive-${idx + 1}.csv`);
    });
  });

  list.querySelectorAll("button[data-delete]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.delete);
      state.archives.splice(idx, 1);
      persist("iaArchives", state.archives);
      renderArchives();
    });
  });
}

function exportMovesCsv(moves = state.currentMoves, filename = "deplacements.csv") {
  if (!moves?.length) return alert("Aucun déplacement à exporter");
  const headers = ["sku", "from", "toZone", "targetType", "priority", "reason"];
  const csv = [headers.join(",")]
    .concat(
      moves.map((m) => headers.map((h) => `"${String(m[h] || "").replaceAll('"', '""')}"`).join(","))
    )
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function login(email) {
  el("inviteGate").classList.add("hidden");
  el("mainContent").classList.remove("hidden");
  el("logoutBtn").classList.remove("hidden");
  el("inviteStatus").textContent = `Session active: ${email}`;
  localStorage.setItem("activeSession", email);
}

function logout() {
  localStorage.removeItem("activeSession");
  el("inviteGate").classList.remove("hidden");
  el("mainContent").classList.add("hidden");
  el("logoutBtn").classList.add("hidden");
  el("inviteStatus").textContent = "Accès protégé par invitation";
}

function setupSession() {
  const email = localStorage.getItem("activeSession");
  if (email) login(email);
}

function setupShortcuts() {
  document.addEventListener("keydown", (evt) => {
    if (evt.ctrlKey && evt.key.toLowerCase() === "enter") {
      evt.preventDefault();
      state.report = analyze();
      renderReport(state.report);
    }
    if (evt.ctrlKey && evt.key.toLowerCase() === "e") {
      evt.preventDefault();
      exportMovesCsv();
    }
  });
}

function bindEvents() {
  el("loginBtn").addEventListener("click", () => {
    const email = el("emailInput").value.trim().toLowerCase();
    const code = el("codeInput").value.trim();
    if (email === "admin@langelier.ca" && code === "LANGELIER-2026") {
      login(email);
    } else {
      alert("Invitation invalide");
    }
  });

  el("logoutBtn").addEventListener("click", logout);

  ["inventory", "sales", "incoming", "locations"].forEach((key) => {
    el(`${key}File`).addEventListener("change", async (evt) => {
      const file = evt.target.files[0];
      if (!file) return;
      state.files[key] = file;
      try {
        state.datasets[key] = await parseExcel(file);
        refreshDatasetStatus();
      } catch (error) {
        alert(`Impossible de lire ${file.name}`);
      }
    });
  });

  el("loadDemoBtn").addEventListener("click", () => {
    state.datasets = JSON.parse(JSON.stringify(demoData));
    refreshDatasetStatus();
    alert("Données de démonstration chargées.");
  });

  el("runAnalysisBtn").addEventListener("click", () => {
    state.report = analyze();
    renderReport(state.report);
  });

  el("generateMovesBtn").addEventListener("click", () => {
    if (!state.report) state.report = analyze();
    renderReport(state.report);
  });

  el("searchSkuInput").addEventListener("input", () => renderReport(state.report));
  el("zoneFilterSelect").addEventListener("change", () => renderReport(state.report));
  el("sortMovesSelect").addEventListener("change", () => renderReport(state.report));

  el("exportCsvBtn").addEventListener("click", () => exportMovesCsv());

  el("printReportBtn").addEventListener("click", () => {
    if (!state.report) return alert("Aucun rapport à imprimer");
    window.print();
  });

  el("archiveReportBtn").addEventListener("click", () => {
    if (!state.report) return alert("Aucun rapport à archiver");
    state.archives.unshift(state.report);
    persist("iaArchives", state.archives);
    renderArchives();
    alert("Rapport archivé");
  });

  el("clearArchivesBtn").addEventListener("click", () => {
    if (!state.archives.length) return;
    if (!window.confirm("Supprimer toutes les archives ?")) return;
    state.archives = [];
    persist("iaArchives", state.archives);
    renderArchives();
  });

  el("saveRulesBtn").addEventListener("click", () => {
    const rules = el("customRules")
      .value.split("\n")
      .map((r) => r.trim())
      .filter(Boolean);
    localStorage.setItem("aiRules", JSON.stringify(rules));
    renderRulesPreview();
    alert("Règles mises à jour");
  });

  el("resetRulesBtn").addEventListener("click", () => {
    localStorage.removeItem("aiRules");
    renderRulesPreview();
  });

  el("themeToggleBtn").addEventListener("click", () => {
    setTheme(document.body.classList.contains("dark") ? "light" : "dark");
  });
}

setupTabs();
renderRulesPreview();
renderArchives();
bindEvents();
setupSession();
setupShortcuts();
setTheme(getTheme());
refreshDatasetStatus();
renderStats(null);

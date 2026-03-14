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

const state = {
  files: {},
  datasets: {},
  report: null,
  archives: JSON.parse(localStorage.getItem("iaArchives") || "[]")
};

const el = (id) => document.getElementById(id);

function loadRules() {
  const raw = localStorage.getItem("aiRules");
  return raw ? JSON.parse(raw) : defaultRules;
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
  return depth * 3;
}

function analyze() {
  const inventory = state.datasets.inventory || [];
  const sales = state.datasets.sales || [];
  const incoming = state.datasets.incoming || [];
  const locations = state.datasets.locations || [];

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
    const pallets = Math.ceil(qty / unitPerPalette(size));
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
    if (incomingQty > 0 && pallets + Math.ceil(incomingQty / unitPerPalette(size)) > cap) {
      reasons.push("Réception attendue dépasse capacité actuelle");
    }
    if (zone === "L2") reasons.push("Meilleur vendeur: prioriser zone critique L2");

    return { sku, qty, size, location: locCode, type, pallets, cap, zone, reasons };
  });

  const toMove = recommendations.filter((r) => r.reasons.length);
  return {
    generatedAt: new Date().toLocaleString("fr-CA"),
    totals: {
      inventoryRows: inventory.length,
      recommendations: toMove.length
    },
    recommendations,
    moves: toMove.map((r) => ({
      sku: r.sku,
      from: r.location,
      toZone: r.zone,
      targetType: r.type === "P1" ? "Racking" : `P${Math.max(1, Number(r.type.replace("P", "")) - 1)}`,
      reason: r.reasons.join(" | ")
    }))
  };
}

function renderReport(report) {
  const lines = [
    `<p><span class="badge">Rapport IA</span><span class="badge">${report.generatedAt}</span></p>`,
    `<p>Lignes inventaire: <b>${report.totals.inventoryRows}</b> • Recommandations: <b>${report.totals.recommendations}</b></p>`,
    "<h3>Raisons des déplacements</h3>",
    "<ul>" + report.moves.map((m) => `<li><b>${m.sku}</b> (${m.from} → ${m.targetType}/${m.toZone}): ${m.reason}</li>`).join("") + "</ul>"
  ];
  el("analysisOutput").innerHTML = lines.join("");
  el("movementOutput").innerHTML = "<ol>" + report.moves.map((m) => `<li>${m.sku}: Déplacer ${m.from} vers ${m.targetType} (${m.toZone})</li>`).join("") + "</ol>";
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
      <p>${a.totals.recommendations} déplacements recommandés.</p>
      <button data-archive="${i}">Ouvrir</button>
      <button data-print="${i}">Imprimer</button>
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
}

function bindEvents() {
  el("loginBtn").addEventListener("click", () => {
    const email = el("emailInput").value.trim().toLowerCase();
    const code = el("codeInput").value.trim();
    if (email === "admin@langelier.ca" && code === "LANGELIER-2026") {
      el("inviteGate").classList.add("hidden");
      el("mainContent").classList.remove("hidden");
      el("inviteStatus").textContent = `Session active: ${email}`;
    } else {
      alert("Invitation invalide");
    }
  });

  ["inventory", "sales", "incoming", "locations"].forEach((key) => {
    el(`${key}File`).addEventListener("change", async (evt) => {
      const file = evt.target.files[0];
      if (!file) return;
      state.files[key] = file;
      state.datasets[key] = await parseExcel(file);
    });
  });

  el("runAnalysisBtn").addEventListener("click", () => {
    state.report = analyze();
    renderReport(state.report);
  });

  el("generateMovesBtn").addEventListener("click", () => {
    if (!state.report) state.report = analyze();
    renderReport(state.report);
  });

  el("printReportBtn").addEventListener("click", () => {
    if (!state.report) return alert("Aucun rapport à imprimer");
    window.print();
  });

  el("archiveReportBtn").addEventListener("click", () => {
    if (!state.report) return alert("Aucun rapport à archiver");
    state.archives.unshift(state.report);
    localStorage.setItem("iaArchives", JSON.stringify(state.archives));
    renderArchives();
    alert("Rapport archivé");
  });

  el("saveRulesBtn").addEventListener("click", () => {
    const rules = el("customRules").value.split("\n").map((r) => r.trim()).filter(Boolean);
    localStorage.setItem("aiRules", JSON.stringify(rules));
    renderRulesPreview();
    alert("Règles mises à jour");
  });

  el("resetRulesBtn").addEventListener("click", () => {
    localStorage.removeItem("aiRules");
    renderRulesPreview();
  });
}

setupTabs();
renderRulesPreview();
renderArchives();
bindEvents();

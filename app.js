
let systemUsers = {};

const BADGES = {
  "MAC001": "Anna Rossi",
  "MAC002": "Maria Bianchi",
  "MAC003": "Giulia Esposito"
};

let currentUser = null;
let controls = [];
let anomalies = [];
let gammaOrders = [];
let currentAttestControlId = null;
let currentAnomalyId = null;
let html5QrCode = null;
let scannerRunning = false;

let masterData = {
  postazioni: [],
  linee: [],
  codici: [],
  macchiniste: [],
  lineeDettaglio: {},
  postazioneToLinea: {},
  tolleranze: {},
  materialiSettings: {}
};
let syncQueue = [];
let syncMeta = { lastError: "", lastAttemptAt: null, lastSuccessAt: null };
let syncInFlight = false;
const DEFAULT_SYNC_ENDPOINT = "";

function $(id){ return document.getElementById(id); }

function todayString() {
  return new Date().toLocaleDateString("it-IT");
}
function dateTimeString() {
  return new Date().toLocaleString("it-IT");
}
function setFastChoiceValue(targetId, value) {
  const input = $(targetId);
  if (!input) return;
  input.value = value || "";
  document.querySelectorAll(`.fast-choice[data-target="${targetId}"] .choice-btn`).forEach((btn) => {
    btn.classList.toggle("active", String(btn.dataset.value || "") === String(value || ""));
  });
}
function initFastChoices() {
  document.querySelectorAll('.fast-choice').forEach((group) => {
    const targetId = group.dataset.target;
    const defaultValue = group.dataset.default || "";
    group.querySelectorAll('.choice-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        setFastChoiceValue(targetId, btn.dataset.value || "");
      });
    });
    const current = $(targetId)?.value || defaultValue;
    setFastChoiceValue(targetId, current);
  });
}
function getSyncEndpoint() {
  return (localStorage.getItem("qc_sync_endpoint") || DEFAULT_SYNC_ENDPOINT || "").trim();
}
function saveStorage() {
  localStorage.setItem("qc_controls", JSON.stringify(controls));
  localStorage.setItem("qc_anomalies", JSON.stringify(anomalies));
  localStorage.setItem("qc_master_data", JSON.stringify(masterData));
  localStorage.setItem("qc_sync_queue", JSON.stringify(syncQueue));
  localStorage.setItem("qc_sync_meta", JSON.stringify(syncMeta));
  localStorage.setItem("qc_gamma_orders", JSON.stringify(gammaOrders));
}
function loadStorage() {
  controls = JSON.parse(localStorage.getItem("qc_controls") || "[]");
  anomalies = JSON.parse(localStorage.getItem("qc_anomalies") || "[]").map(a => ({
    actions: [],
    status: "Aperta",
    ...a,
    actions: Array.isArray(a.actions) ? a.actions : []
  }));
  // reset rigoroso a ogni avvio: Gamma e QC non vengono considerati caricati
  masterData = {
    postazioni: [], linee: [], codici: [], macchiniste: [],
    lineeDettaglio: {}, postazioneToLinea: {}, tolleranze: {}, materialiSettings: {}
  };
  syncQueue = JSON.parse(localStorage.getItem("qc_sync_queue") || "[]");
  syncMeta = JSON.parse(localStorage.getItem("qc_sync_meta") || '{"lastError":"","lastAttemptAt":null,"lastSuccessAt":null}');
  gammaOrders = [];
  try {
    localStorage.removeItem("qc_master_data");
    localStorage.removeItem("qc_gamma_orders");
  } catch (e) {}
}
function queueControlForSync(control) {
  if (!control || !control.id) return;
  const relatedAnomalies = anomalies.filter(a => a.controlId === control.id);
  const item = {
    id: `sync_${control.id}`,
    controlId: control.id,
    type: "control_upsert",
    generatedAt: new Date().toISOString(),
    payload: {
      control: JSON.parse(JSON.stringify(control)),
      anomalies: JSON.parse(JSON.stringify(relatedAnomalies))
    }
  };
  syncQueue = syncQueue.filter(x => x.controlId !== control.id);
  syncQueue.push(item);
  syncMeta.lastError = "";
  saveStorage();
  renderSyncStatus();
  scheduleSync();
}
function renderSyncStatus() {
  const pendingBadge = $("badgePending");
  const storageBadge = $("badgeStorage");
  const endpointBadge = $("badgeEndpoint");
  const retry = $("btnRetrySync");
  if (!pendingBadge || !storageBadge || !endpointBadge || !retry) return;

  const pending = syncQueue.length;
  const online = navigator.onLine !== false;
  const hasError = !!syncMeta.lastError;
  const endpointConfigured = !!getSyncEndpoint();

  pendingBadge.className = "top-badge";
  storageBadge.className = "top-badge badge-neutral-soft";
  endpointBadge.className = "top-badge badge-neutral-soft";

  if (hasError) {
    pendingBadge.textContent = "Sync: errore";
    pendingBadge.classList.add("badge-danger-soft");
    storageBadge.textContent = "Salvataggio: locale";
    endpointBadge.textContent = syncMeta.lastError;
    endpointBadge.classList.add("badge-danger-soft");
    retry.classList.remove("hidden");
    return;
  }

  if (pending > 0) {
    pendingBadge.textContent = `Attesa: ${pending}`;
    pendingBadge.classList.add("badge-warning-soft");
    storageBadge.textContent = online ? "Salvataggio: locale" : "Offline: locale";
    endpointBadge.textContent = endpointConfigured ? (syncInFlight ? "Endpoint: sincronizzazione" : "Endpoint: pronto") : "Endpoint: OFF";
    endpointBadge.classList.add(endpointConfigured ? "badge-success-soft" : "badge-neutral-soft");
    retry.classList.toggle("hidden", !online || !endpointConfigured);
    return;
  }

  pendingBadge.textContent = "Attesa: 0";
  pendingBadge.classList.add("badge-success-soft");
  storageBadge.textContent = syncMeta.lastSuccessAt ? `Ultima sync: ${syncMeta.lastSuccessAt}` : "Salvataggio: locale";
  endpointBadge.textContent = endpointConfigured ? "Endpoint: pronto" : "Endpoint: OFF";
  endpointBadge.classList.add(endpointConfigured ? "badge-success-soft" : "badge-neutral-soft");
  retry.classList.add("hidden");
}
function scheduleSync() {
  if (!syncQueue.length) return;
  if (navigator.onLine === false) {
    renderSyncStatus();
    return;
  }
  if (!getSyncEndpoint()) {
    renderSyncStatus();
    return;
  }
  setTimeout(() => { attemptSync(); }, 150);
}
async function attemptSync() {
  if (syncInFlight || !syncQueue.length) {
    renderSyncStatus();
    return;
  }
  if (navigator.onLine === false) {
    renderSyncStatus();
    return;
  }
  const endpoint = getSyncEndpoint();
  if (!endpoint) {
    renderSyncStatus();
    return;
  }

  syncInFlight = true;
  syncMeta.lastAttemptAt = dateTimeString();
  syncMeta.lastError = "";
  saveStorage();
  renderSyncStatus();

  try {
    while (syncQueue.length) {
      const item = syncQueue[0];
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          clientRequestId: item.id,
          type: item.type,
          generatedAt: item.generatedAt,
          payload: item.payload
        })
      });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      syncQueue.shift();
      syncMeta.lastSuccessAt = dateTimeString();
      saveStorage();
      renderSyncStatus();
    }
  } catch (err) {
    syncMeta.lastError = `Sync fallita: ${String(err.message || err)}`;
    saveStorage();
    renderSyncStatus();
  } finally {
    syncInFlight = false;
    saveStorage();
    renderSyncStatus();
  }
}

function loadStoredUsers() {
  systemUsers = JSON.parse(localStorage.getItem("qc_login_users") || "{}");
}
function saveStoredUsers(users, importedAt=null) {
  systemUsers = users || {};
  localStorage.setItem("qc_login_users", JSON.stringify(systemUsers));
  if (importedAt) localStorage.setItem("qc_login_users_updated_at", importedAt);
}
function getStoredUsersUpdatedAt() {
  return localStorage.getItem("qc_login_users_updated_at") || "";
}
function hasStoredUsers() {
  return Object.keys(systemUsers || {}).length > 0;
}
function parseQcUsers(workbook) {
  const userRows = sheetRows(workbook, ["controllo qualità","Controllo Qualità","Controllo_Qualita","Controllo Qualita"]);
  if (!userRows.length) throw new Error("Foglio controllo qualità mancante o vuoto");
  const users = {};
  userRows.forEach((r) => {
    const username = String(r["Utente"] || r.Utente || "").trim();
    const password = String(r["Password"] ?? r.password ?? "").trim();
    const active = String(r["Attiva (SI/NO)"] || r.Attiva || "SI").trim().toUpperCase();
    if (!username) return;
    if (!password) throw new Error(`Password mancante per utente ${username}`);
    if (active !== "SI" && active !== "NO") throw new Error(`Valore Attiva non valido per utente ${username}`);
    if (users[username]) throw new Error(`Utente duplicato: ${username}`);
    if (active === "NO") return;
    users[username] = {
      password,
      name: String(r["Nome e cognome"] || r.Nome || "").trim() || username,
      badge: String(r["ID_Badge"] || r.Badge || "").trim()
    };
  });
  if (!Object.keys(users).length) throw new Error("Nessun utente attivo trovato nel foglio controllo qualità");
  return users;
}
function updateLoginScreenState() {
  const bootstrap = $("bootstrapBox");
  const loginForm = $("loginFormBox");
  const info = $("loginUsersInfo");
  if (!bootstrap || !loginForm || !info) return;
  const usersReady = hasStoredUsers();
  bootstrap.classList.toggle("hidden", usersReady);
  loginForm.classList.toggle("hidden", !usersReady);
  if (usersReady) {
    const count = Object.keys(systemUsers).length;
    const updatedAt = getStoredUsersUpdatedAt();
    info.textContent = updatedAt ? `Utenti attivi: ${count} · Anagrafica aggiornata il ${updatedAt}` : `Utenti attivi: ${count}`;
  } else {
    info.textContent = "";
  }
}

function normalizeCode(raw) {
  if (!raw) return "";
  const txt = String(raw).trim().toUpperCase();
  const jsonBadge = txt.match(/"badge"\s*:\s*"?([^",}\s]+)"?/i);
  if (jsonBadge && jsonBadge[1]) return normalizeCode(jsonBadge[1]);
  const pipeBadge = txt.match(/^QC\|(?:MAC|QC)\|([^|]+)/i);
  if (pipeBadge && pipeBadge[1]) return normalizeCode(pipeBadge[1]);
  const genericBadge = txt.match(/BADGE\s*[=:]\s*([A-Z0-9]+)/i);
  if (genericBadge && genericBadge[1]) return normalizeCode(genericBadge[1]);
  const macMatch = txt.match(/MAC\d{3,}/);
  if (macMatch) return macMatch[0];
  const digits = txt.replace(/\D/g, "");
  if (digits) {
    const stripped = digits.replace(/^0+/, "");
    return stripped || "0";
  }
  return txt;
}
function badgeKey(raw) {
  return normalizeCode(raw);
}
function deriveTypeFromWorkstation(workstation) {
  const code = String(workstation || "").toUpperCase().trim();
  if (code.startsWith("MCD")) return "IM";
  if (code.startsWith("TC")) return "SORF";
  if (code.startsWith("ML")) return "CU";
  return "";
}
function uniqueClean(arr) {
  return [...new Set(arr.map(v => String(v ?? "").trim()).filter(Boolean))];
}
function splitAllowedValues(raw) {
  return uniqueClean(String(raw ?? "").split("/").map(v => v.trim()));
}
function parseDelimitedLine(line, delimiter=";") {
  const out = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delimiter && !inQuotes) {
      out.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out.map(v => String(v || "").trim());
}
function normalizeHeaderKey(value) {
  return String(value || "").trim().toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}
function ensureSelectOption(selectId, value, label=null) {
  const sel = $(selectId);
  const normalized = String(value || "").trim();
  if (!sel || !normalized) return;
  if (![...sel.options].some(o => String(o.value || "").trim() === normalized)) {
    const opt = document.createElement("option");
    opt.value = normalized;
    opt.textContent = label || normalized;
    sel.appendChild(opt);
  }
  sel.value = normalized;
}
function getCurrentQcLine() {
  return String(getSelectedLinea() || masterData.postazioneToLinea?.[$("workstationSelect")?.value] || "").trim();
}
function getSelectedGammaOrder() {
  const lot = String($("lot")?.value || "").trim();
  if (!lot) return null;
  return gammaOrders.find(x => x.lotto === lot || x.idOrdine === lot) || null;
}
function toggleManualEntryPanel(forceOpen = null) {
  const panel = $("manualEntryPanel");
  const btn = $("btnToggleManualEntry");
  if (!panel || !btn) return;
  const shouldOpen = forceOpen === null ? panel.classList.contains("hidden") : !!forceOpen;
  panel.classList.toggle("hidden", !shouldOpen);
  btn.textContent = shouldOpen ? "− Nascondi inserimento manuale" : "+ Inserimento manuale";
}
function resetGammaDependentFields(keepLot=false) {
  if (!keepLot && $("lot")) $("lot").value = "";
  if ($("setupCodeSelect")) $("setupCodeSelect").value = "";
  if ($("setupCodeNew")) $("setupCodeNew").value = "";
  if ($("particularAuto")) $("particularAuto").value = "";
  if ($("particularNew")) $("particularNew").value = "";
  if ($("threadNeedle")) $("threadNeedle").value = "";
  if ($("threadCrochet")) $("threadCrochet").value = "";
  if ($("needleType")) $("needleType").value = "";
  setStatusField("threadNeedleStatus", "");
  setStatusField("threadCrochetStatus", "");
  setStatusField("needleTypeStatus", "");
}
function populateGammaLots() {
  const datalist = $("lotSuggestions");
  if (!datalist) return;
  datalist.innerHTML = "";
  const currentLine = getCurrentQcLine();
  gammaOrders
    .filter(order => !currentLine || String(order.linea || "").trim() === currentLine)
    .forEach(order => {
      const opt = document.createElement("option");
      opt.value = order.lotto;
      opt.label = `${order.codice} — ${order.descrizione}`;
      datalist.appendChild(opt);
    });
}
function applyGammaOrderToForm(order) {
  if (!order) return;
  const currentLine = getCurrentQcLine();
  const orderLine = String(order.linea || "").trim();
  if (currentLine && orderLine && currentLine !== orderLine) {
    resetGammaDependentFields();
    $("saveMsg").className = "msg err";
    $("saveMsg").textContent = "Lotto non compatibile con la linea della postazione.";
    focusIfEnabled("lot");
    return;
  }
  $("lot").value = order.lotto || "";
  if (order.codice) {
    ensureSelectOption("setupCodeSelect", order.codice);
    $("setupCodeNew").value = "";
  }
  if (order.particolare) {
    $("particularAuto").value = order.particolare;
    $("particularNew").value = "";
  } else {
    updateParticularAuto();
  }
  updateToleranceInfo();
  populateMaterialSelects();
  if (order.filatoAgo) ensureSelectOption("threadNeedle", order.filatoAgo);
  if (order.filatoCrochet) ensureSelectOption("threadCrochet", order.filatoCrochet);
  if (order.ago) ensureSelectOption("needleType", order.ago);
  applyMaterialValidation();
  refreshDerivedFieldStates();
}
function handleGammaLotInput() {
  const lot = String($("lot").value || "").trim();
  if (!lot) {
    resetGammaDependentFields(true)
  }
  if (!lot) return;
  const order = gammaOrders.find(x => x.lotto === lot || x.idOrdine === lot);
  if (!order) return;
  applyGammaOrderToForm(order);
}
function importGammaFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const text = String(e.target.result || "").replace(/^\ufeff/, "");
      const lines = text.replace(/\r/g, "").split("\n").filter(l => String(l || "").trim());
      if (lines.length < 2) throw new Error("CSV vuoto o non valido");
      const headers = parseDelimitedLine(lines[0]).map(normalizeHeaderKey);
      const idx = (name) => headers.indexOf(normalizeHeaderKey(name));
      const iId = idx("ID_ORDINE");
      const iLotto = idx("LOTTO");
      const iCodice = idx("CODICE_PRODOTTO");
      const iDesc = idx("DESCRIZIONE_PRODOTTO");
      const iLinea = idx("ID_LINEA");
      const iPart = idx("PARTICOLARE");
      const iFilAgo = idx("CODICE_FILATO_AGO");
      const iFilCro = idx("CODICE_FILATO_CROCHET");
      const iAgo = idx("CODICE_AGO");
      const iStato = idx("STATO_ORDINE");
      if ([iId, iLotto, iCodice, iLinea].some(v => v < 0)) throw new Error("Colonne Gamma obbligatorie mancanti");

      gammaOrders = lines.slice(1).map(line => parseDelimitedLine(line)).map(cols => ({
        idOrdine: cols[iId] || "",
        lotto: cols[iLotto] || "",
        codice: cols[iCodice] || "",
        descrizione: iDesc >= 0 ? (cols[iDesc] || "") : "",
        linea: iLinea >= 0 ? (cols[iLinea] || "") : "",
        particolare: iPart >= 0 ? (cols[iPart] || "") : "",
        filatoAgo: iFilAgo >= 0 ? (cols[iFilAgo] || "") : "",
        filatoCrochet: iFilCro >= 0 ? (cols[iFilCro] || "") : "",
        ago: iAgo >= 0 ? (cols[iAgo] || "") : "",
        statoOrdine: iStato >= 0 ? (cols[iStato] || "") : ""
      })).filter(r => r.lotto && r.codice).filter(r => !r.statoOrdine || !["CHIUSO","ANNULLATO"].includes(String(r.statoOrdine).toUpperCase()));

      saveStorage();
      populateGammaLots();
      resetGammaDependentFields();
      refreshDerivedFieldStates();
      $("saveMsg").className = "msg ok";
      $("saveMsg").textContent = `Import Gamma OK: ${gammaOrders.length} ordini attivi caricati.`;
      updateDataSourceStatus();
    } catch (err) {
      $("saveMsg").className = "msg err";
      $("saveMsg").textContent = `Import Gamma fallito: ${String(err.message || err)}`;
    }
  };
  reader.readAsText(file, "utf-8");
}
function getMaterialRules(linea) {
  return masterData.materialiSettings?.[String(linea || "").trim()] || null;
}
function populateSelect(selectId, values, formatter=null) {
  const sel = $(selectId);
  if (!sel) return;
  const current = sel.value;
  sel.innerHTML = '<option value="">Seleziona</option>';
  values.forEach(v => {
    const opt = document.createElement("option");
    if (typeof v === "string") {
      opt.value = v;
      opt.textContent = formatter ? formatter(v) : v;
    } else {
      const fallbackValue = v.value ?? v.badge ?? v.codice ?? v.id ?? "";
      opt.value = String(fallbackValue).trim();
      opt.textContent = formatter ? formatter(v) : opt.value;
    }
    sel.appendChild(opt);
  });
  if ([...sel.options].some(o => o.value === current)) sel.value = current;
}
function populateCodeSelect() {
  const linea = getSelectedLinea();
  const codiciFiltrati = (masterData.codici || [])
    .filter(item => String(item.linea || "").trim() === linea)
    .map(item => item.codice);
  populateSelect("setupCodeSelect", codiciFiltrati);
}
function populateMasterSelects() {
  populateSelect("workstationSelect", masterData.postazioni);
  populateSelect("lineaSelect", masterData.linee);
  populateCodeSelect();
  populateSelect("machineBadgeSelect", masterData.macchiniste, item => item.nome ? `${item.badge} — ${item.nome}` : item.badge);
  populateMaterialSelects();
  populateGammaLots();
}
function sheetRows(workbook, names) {
  for (const n of names) {
    if (workbook.Sheets[n]) {
      return XLSX.utils.sheet_to_json(workbook.Sheets[n], {defval:""});
    }
  }
  return [];
}
function importAnagraficheFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const postRows = sheetRows(workbook, ["Postazioni"]);
      const lineRows = sheetRows(workbook, ["Elenco Linee","Elenco_Linee","Linee"]);
      const macRows = sheetRows(workbook, ["Macchiniste"]);
      const codRows = sheetRows(workbook, ["Codici_Prodotto","Codici Prodotto","Codici"]);
      const tollRows = sheetRows(workbook, ["Misure","Tolleranze","Tolleranze_Misure"]);
      const materialRows = sheetRows(workbook, ["Materiali e settaggi","Materiali_e_settaggi","Materiali"]);
      let qcUsers = {};

      const errors = [];
      try {
        qcUsers = parseQcUsers(workbook);
      } catch (userErr) {
        errors.push(String(userErr.message || userErr));
      }
      if (!postRows.length) errors.push("Foglio Postazioni mancante o vuoto");
      if (!lineRows.length) errors.push("Foglio Elenco Linee mancante o vuoto");
      if (!tollRows.length) errors.push("Foglio Misure mancante o vuoto");
      if (errors.length) {
        $("saveMsg").className = "msg err";
        $("saveMsg").innerHTML = "Import anagrafiche KO:<br>" + errors.join("<br>");
        return;
      }

      const activePostRows = postRows
        .filter(r => String(r["Attiva (SI/NO)"] || r.Attiva || "SI").toUpperCase() !== "NO");
      const activeLineRows = lineRows
        .filter(r => String(r["Attiva (SI/NO)"] || r.Attiva || "SI").toUpperCase() !== "NO");
      const activeCodRows = codRows
        .filter(r => String(r["Attivo (SI/NO)"] || r.Attivo || "SI").toUpperCase() !== "NO");
      const activeMacRows = macRows
        .filter(r => String(r["Attiva (SI/NO)"] || r.Attiva || "SI").toUpperCase() !== "NO");

      const candidate = {
        postazioni: activePostRows
          .map(r => String(r.ID_Postazione || r.Postazione || "").trim())
          .filter(Boolean),
        linee: uniqueClean(activeLineRows.map(r => r.ID_Linea || r.Id_Linea || r.id_linea || "")),
        lineeDettaglio: {},
        postazioneToLinea: {},
        codici: activeCodRows
          .map(r => ({
            codice: String(r.Codice || "").trim(),
            linea: String(r.ID_Linea || r.Id_Linea || r.id_linea || "").trim()
          }))
          .filter(x => x.codice),
        macchiniste: activeMacRows
          .map(r => ({
            badge: String(r.ID_Badge || r.Badge || "").trim(),
            nome: String(r["Nome e cognome"] || r.Nome || "").trim() || [r.Nome || "", r.Cognome || ""].join(" ").trim()
          }))
          .filter(x => x.badge),
        tolleranze: {},
        materialiSettings: {}
      };

      activePostRows.forEach(r => {
        const postazione = String(r.ID_Postazione || r.Postazione || "").trim();
        const linea = String(r.Linea || r.linea || r.ID_Linea || r.Id_Linea || r.id_linea || "").trim();
        if (postazione && linea) candidate.postazioneToLinea[postazione] = linea;
      });

      activeLineRows.forEach(r => {
        const idLinea = String(r.ID_Linea || r.Id_Linea || r.id_linea || "").trim();
        const particolare = String(r.Linea || r.linea || "").trim();
        if (idLinea && particolare) candidate.lineeDettaglio[idLinea] = particolare;
      });

      tollRows.forEach(r => {
        const linea = String(r.ID_Linea || r.Id_Linea || r.id_linea || r.Linea || r.linea || "").trim();
        if (!linea) return;
        const values = {
          bordoMin: parseFloat(r["Bordo cucitura (mm) - Valore Min"] ?? r.Bordo_Min ?? r.BordoMin ?? r.bordo_min ?? ""),
          bordoMax: parseFloat(r["Bordo cucitura (mm) - Valore Max "] ?? r["Bordo cucitura (mm) - Valore Max"] ?? r.Bordo_Max ?? r.BordoMax ?? r.bordo_max ?? ""),
          passoMin: parseFloat(r["Passo punto  (mm)- Valore Min"] ?? r.Passo_Min ?? r.PassoMin ?? r.passo_min ?? ""),
          passoMax: parseFloat(r["Passo punto  (mm)-  Valore Max "] ?? r["Passo punto  (mm)-  Valore Max"] ?? r.Passo_Max ?? r.PassoMax ?? r.passo_max ?? ""),
          taccheMin: parseFloat(r["Disallineamento tacche (mm)  - Valore Min"] ?? r.Tacche_Min ?? r.TaccheMin ?? r.tacche_min ?? ""),
          taccheMax: parseFloat(r["Disallineamento tacche (mm)  - Valore Max "] ?? r["Disallineamento tacche (mm)  - Valore Max"] ?? r.Tacche_Max ?? r.TaccheMax ?? r.tacche_max ?? "")
        };
        const hasAtLeastOneNumber = Object.values(values).some(v => Number.isFinite(v));
        if (!hasAtLeastOneNumber) return;
        candidate.tolleranze[linea] = values;
      });

      materialRows.forEach(r => {
        const linea = String(r.ID_Linea || r.Id_Linea || r.id_linea || "").trim();
        if (!linea) return;
        const settings = {
          threadNeedle: splitAllowedValues(r["Codice filato ago"] ?? r.Codice_filato_ago ?? r.thread_needle ?? ""),
          threadCrochet: splitAllowedValues(r["Codice filato crochet"] ?? r.Codice_filato_crochet ?? r.thread_crochet ?? ""),
          needleType: splitAllowedValues(r["Ago"] ?? r.ago ?? r.Needle ?? "")
        };
        const hasMaterialData = settings.threadNeedle.length || settings.threadCrochet.length || settings.needleType.length;
        if (!hasMaterialData) return;
        candidate.materialiSettings[linea] = settings;
      });

      const lineeSet = new Set(candidate.linee);
      const postMismatch = [...new Set(Object.entries(candidate.postazioneToLinea)
        .filter(([, linea]) => linea && !lineeSet.has(linea))
        .map(([postazione, linea]) => `${postazione} → ${linea}`))];
      if (postMismatch.length) {
        errors.push("Postazioni con linea non presente in Elenco Linee: " + postMismatch.join(", "));
      }

      const codMismatch = [...new Set(candidate.codici
        .filter(x => x.linea && !lineeSet.has(x.linea))
        .map(x => `${x.codice} → ${x.linea}`))];
      if (codMismatch.length) {
        errors.push("Codici prodotto con ID_Linea non presente in Elenco Linee: " + codMismatch.join(", "));
      }

      const tollMismatch = [...new Set(Object.keys(candidate.tolleranze)
        .filter(linea => linea && !lineeSet.has(linea)))];
      if (tollMismatch.length) {
        errors.push("Misure con ID_Linea non presente in Elenco Linee: " + tollMismatch.join(", "));
      }

      const materialiMismatch = [...new Set(Object.keys(candidate.materialiSettings)
        .filter(linea => linea && !lineeSet.has(linea)))];
      if (materialiMismatch.length) {
        errors.push("Materiali e settaggi con ID_Linea non presente in Elenco Linee: " + materialiMismatch.join(", "));
      }

      const lineeSenzaParticolare = candidate.linee.filter(linea => !String(candidate.lineeDettaglio[linea] || "").trim());
      if (lineeSenzaParticolare.length) {
        errors.push("Linee senza particolare associato: " + lineeSenzaParticolare.join(", "));
      }

      if (errors.length) {
        $("saveMsg").className = "msg err";
        $("saveMsg").innerHTML = "Import anagrafiche KO:<br>" + errors.join("<br>");
        return;
      }

      masterData = candidate;
      saveStoredUsers(qcUsers, dateTimeString());
      saveStorage();
      populateMasterSelects();
      updateAutoType();
      populateMasterSelects();
      updateToleranceInfo();
      refreshDerivedFieldStates();
      $("saveMsg").className = "msg ok";
      $("saveMsg").textContent = `Import anagrafiche OK. Utenti attivi caricati: ${Object.keys(qcUsers).length}.`;
      updateDataSourceStatus();
      updateLoginScreenState();
    } catch (err) {
      $("saveMsg").className = "msg err";
      $("saveMsg").textContent = "Import anagrafiche fallito.";
    }
  };
  reader.readAsArrayBuffer(file);
}


function importBootstrapAnagraficheFile(file) {
  if (!file) return;
  const msg = $("bootstrapMsg");
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const qcUsers = parseQcUsers(workbook);
      saveStoredUsers(qcUsers, dateTimeString());
      if (msg) {
        msg.className = "msg ok";
        msg.textContent = `Anagrafica utenti caricata. Utenti attivi: ${Object.keys(qcUsers).length}. Ora puoi accedere.`;
      }
      updateLoginScreenState();
      focusIfEnabled("loginUser");
    } catch (err) {
      if (msg) {
        msg.className = "msg err";
        msg.textContent = `Import configurazione fallito: ${String(err.message || err)}`;
      }
    }
  };
  reader.readAsArrayBuffer(file);
}


function refreshDerivedFieldStates() {
  const lineField = $("lineaSelect");
  const codeField = $("setupCodeSelect");
  const stitchField = $("stitchTypeAuto");
  const particularField = $("particularAuto");
  if (lineField) lineField.disabled = true;
  if (codeField) codeField.disabled = true;
  if (stitchField) stitchField.readOnly = true;
  if (particularField) particularField.readOnly = true;
}

function focusIfEnabled(id) {
  const el = $(id);
  if (!el || el.disabled || el.type === "hidden") return;
  requestAnimationFrame(() => el.focus());
}
function handleEnterAdvance(event, nextId) {
  if (event.key !== "Enter") return;
  event.preventDefault();
  focusIfEnabled(nextId);
}
function syncMachineSelectionPreview() {
  const sel = $("machineBadgeSelect");
  const nameField = $("manualMachineName");
  const badgeField = $("manualBadge");
  if (!sel || !nameField || !badgeField) return;
  const selectedBadge = String(sel.value || "").trim();
  if (!selectedBadge) return;
  const found = (masterData.macchiniste || []).find(x => String(x.badge || "").trim() === selectedBadge);
  if (!found) return;
  nameField.value = found.nome || "";
  badgeField.value = selectedBadge;
}

function initSmartFocus() {
  $("workstationSelect").addEventListener("change", () => { populateGammaLots(); resetGammaDependentFields(); focusIfEnabled("lot"); });
  $("machineBadgeSelect")?.addEventListener("change", syncMachineSelectionPreview);

  const lotInput = $("lot");
  const moveLotToSetupCode = () => {
    if (!lotInput.value.trim()) return;
    focusIfEnabled("setupCodeSelect");
  };
  lotInput.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    moveLotToSetupCode();
  });
  lotInput.addEventListener("change", () => { handleGammaLotInput(); moveLotToSetupCode(); });
  lotInput.addEventListener("blur", () => { handleGammaLotInput(); moveLotToSetupCode(); });
  lotInput.addEventListener("input", handleGammaLotInput);

  $("setupCodeSelect").addEventListener("change", () => focusIfEnabled("measureBorder"));
  $("measureBorder").addEventListener("keydown", (event) => handleEnterAdvance(event, "measurePitch"));
  $("measurePitch").addEventListener("keydown", (event) => handleEnterAdvance(event, "measureNotch"));
  $("measureNotch").addEventListener("keydown", (event) => handleEnterAdvance(event, "btnSave"));
}

function getSelectedLinea() {
  return ($("lineaNew").value || $("lineaSelect").value || "").trim();
}
function updateAutoType() {
  const workstation = ($("workstationNew").value || $("workstationSelect").value || "").trim();
  $("stitchTypeAuto").value = deriveTypeFromWorkstation(workstation);
}
function syncLineFromWorkstation() {
  const workstation = ($("workstationNew").value || $("workstationSelect").value || "").trim();
  const mapped = masterData.postazioneToLinea?.[workstation] || "";
  if (!mapped) return;
  $("lineaSelect").value = mapped;
  if ($("lineaSelect").value !== mapped) {
    const opt = document.createElement("option");
    opt.value = mapped;
    opt.textContent = mapped;
    $("lineaSelect").appendChild(opt);
    $("lineaSelect").value = mapped;
  }
  populateCodeSelect();
  refreshDerivedFieldStates();
  updateToleranceInfo();
}

function updateParticularAuto() {
  const linea = String(getSelectedLinea() || "").trim();
  const value = linea ? (masterData.lineeDettaglio?.[linea] || "") : "";
  $("particularAuto").value = value;
}
function autoMeasureStatus(value, min, max) {
  const num = parseFloat(value);
  if (Number.isNaN(num) || Number.isNaN(min) || Number.isNaN(max)) return null;
  return num >= min && num <= max ? "OK" : "KO";
}
function setStatusField(id, value) {
  const el = $(id);
  el.value = value || "";
  el.classList.remove("status-ok","status-ko");
  if (value === "OK") el.classList.add("status-ok");
  if (value === "KO") el.classList.add("status-ko");
}
function applyToleranceAuto() {
  const linea = String(getSelectedLinea() || "").trim();
  const tol = masterData.tolleranze?.[linea];
  if (!tol) {
    setStatusField("measureBorderStatus", "");
    setStatusField("measurePitchStatus", "");
    setStatusField("measureNotchStatus", "");
    return;
  }
  const border = autoMeasureStatus($("measureBorder").value, tol.bordoMin, tol.bordoMax);
  const pitch = autoMeasureStatus($("measurePitch").value, tol.passoMin, tol.passoMax);
  const notch = autoMeasureStatus($("measureNotch").value, tol.taccheMin, tol.taccheMax);
  setStatusField("measureBorderStatus", border);
  setStatusField("measurePitchStatus", pitch);
  setStatusField("measureNotchStatus", notch);
}
function populateMaterialSelects() {
  const gammaOrder = getSelectedGammaOrder();
  if (gammaOrder) {
    populateSelect("threadNeedle", uniqueClean([gammaOrder.filatoAgo].filter(Boolean)));
    populateSelect("threadCrochet", uniqueClean([gammaOrder.filatoCrochet].filter(Boolean)));
    populateSelect("needleType", uniqueClean([gammaOrder.ago].filter(Boolean)));
    applyMaterialValidation();
    return;
  }
  const linea = String(getSelectedLinea() || "").trim();
  const rules = getMaterialRules(linea);
  populateSelect("threadNeedle", rules?.threadNeedle || []);
  populateSelect("threadCrochet", rules?.threadCrochet || []);
  populateSelect("needleType", rules?.needleType || []);
  applyMaterialValidation();
  refreshDerivedFieldStates();
}
function applyMaterialValidation() {
  const statuses = ["threadNeedleStatus","threadCrochetStatus","needleTypeStatus"];
  statuses.forEach(id => {
    const el = $(id);
    if (!el) return;
    const current = String(el.value || "").trim();
    if (!["OK","KO"].includes(current)) {
      setFastChoiceValue(id, "");
    } else {
      setFastChoiceValue(id, current);
    }
  });
}
function updateToleranceInfo() {
  const linea = getSelectedLinea();
  updateParticularAuto();
  populateMaterialSelects();
  const tol = masterData.tolleranze?.[linea];
  const box = $("toleranceInfo");
  if (!tol) {
    box.textContent = linea ? `Nessuna tolleranza trovata per la linea ${linea}` : "";
    setStatusField("measureBorderStatus", "");
    setStatusField("measurePitchStatus", "");
    setStatusField("measureNotchStatus", "");
    return;
  }
  box.textContent = `Tolleranze linea ${linea} — Bordo: ${tol.bordoMin} ÷ ${tol.bordoMax} · Passo: ${tol.passoMin} ÷ ${tol.passoMax} · Tacche: ${tol.taccheMin} ÷ ${tol.taccheMax}`;
  applyToleranceAuto();
}
function overallOutcome(c) {
  const values = [
    c.measureBorderStatus, c.measurePitchStatus, c.measureNotchStatus,
    c.qualityStop, c.qualitySkipped, c.qualityNeedle,
    c.threadNeedleStatus, c.threadCrochetStatus, c.needleTypeStatus,
    c.notesStatus
  ];
  return values.includes("KO") ? "KO" : "OK";
}
function createAnomaliesFor(control) {
  const created = [];
  if (control.measureBorderStatus === "KO") created.push({ type: "Bordo cucitura KO", severity: "Grave" });
  if (control.measurePitchStatus === "KO") created.push({ type: "Passo punto KO", severity: "Grave" });
  if (control.measureNotchStatus === "KO") created.push({ type: "Disallineamento tacche KO", severity: "Grave" });
  if (control.qualityStop === "KO") created.push({ type: "Fermapunto non conforme", severity: "Media" });
  if (control.qualitySkipped === "KO") created.push({ type: "Punti saltati", severity: "Media" });
  if (control.qualityNeedle === "KO") created.push({ type: "Ago spuntato", severity: "Media" });
  if (control.threadNeedleStatus === "KO") created.push({ type: "Codice filato ago non conforme", severity: "Media" });
  if (control.threadCrochetStatus === "KO") created.push({ type: "Codice filato crochet non conforme", severity: "Media" });
  if (control.needleTypeStatus === "KO") created.push({ type: "Ago non conforme", severity: "Media" });
  if (control.notesStatus === "KO") created.push({ type: "Note KO", severity: "Media" });

  created.forEach(a => {
    anomalies.push({
      id: crypto.randomUUID(),
      controlId: control.id,
      openedAt: control.createdAt,
      lot: control.lot,
      type: a.type,
      severity: a.severity,
      status: "Aperta"
    });
  });
}
function renderHistory() {
  const body = $("historyBody");
  body.innerHTML = "";
  controls.slice().reverse().forEach(c => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${c.createdAt}</td>
      <td>${c.qualityUser}</td>
      <td>${c.shift}</td>
      <td>${c.workstation}</td>
      <td>${c.lot}</td>
      <td><span class="badge ${c.outcome === "OK" ? "badge-ok":"badge-ko"}">${c.outcome}</span></td>
      <td><span class="badge ${c.status === "Validato" ? "badge-valid":"badge-state"}">${c.status}</span></td>
      <td>${c.machineBadge ? c.machineBadge : "Non attestata"}</td>
      <td>
        <button class="secondary" onclick="openDetailModal('${c.id}')">Dettagli</button>
        ${c.status !== "Validato" ? `<button class="warning" onclick="openAttestModal('${c.id}')">Attesta QR</button>` : ``}
      </td>
    `;
    body.appendChild(tr);
  });
}
function renderAnomalies() {
  const body = $("anomalyBody");
  body.innerHTML = "";
  anomalies.slice().reverse().forEach(a => {
    const tr = document.createElement("tr");
    const effectiveStatus = getDerivedAnomalyStatus(a);
    const badgeClass = effectiveStatus === "Chiusa" ? "badge-valid" : (effectiveStatus === "In gestione" ? "badge-state" : "badge-open");
    tr.innerHTML = `
      <td>${a.openedAt}</td>
      <td>${a.lot}</td>
      <td>${a.type}</td>
      <td><span class="badge ${a.severity === "Grave" ? "badge-sev-high":"badge-sev-medium"}">${a.severity}</span></td>
      <td><span class="badge ${badgeClass}">${effectiveStatus}</span></td>
      <td><button class="secondary" onclick="openAnomalyModal('${a.id}')">Dettagli</button></td>
    `;
    body.appendChild(tr);
  });
}
function renderKpis() {
  const total = controls.length;
  const ko = controls.filter(c => c.outcome === "KO").length;
  const rate = total ? ((ko / total) * 100).toFixed(1).replace(".", ",") + "%" : "0%";
  const open = anomalies.filter(a => getDerivedAnomalyStatus(a) === "Aperta").length;
  const inProgress = anomalies.filter(a => getDerivedAnomalyStatus(a) === "In gestione").length;
  const closed = anomalies.filter(a => getDerivedAnomalyStatus(a) === "Chiusa").length;
  const resolvedBase = open + inProgress + closed;
  const resolutionRate = resolvedBase ? ((closed / resolvedBase) * 100).toFixed(1).replace(".", ",") + "%" : "0%";
  $("kpiTotalControls").textContent = total;
  $("kpiTotalKo").textContent = ko;
  $("kpiKoRate").textContent = rate;
  $("kpiOpenAnomalies").textContent = open;
  if ($("kpiInProgressAnomalies")) $("kpiInProgressAnomalies").textContent = inProgress;
  if ($("kpiClosedAnomalies")) $("kpiClosedAnomalies").textContent = closed;
  if ($("kpiResolutionRate")) $("kpiResolutionRate").textContent = resolutionRate + " risolte";
}
function clearForm() {
  ["workstationNew","stitchTypeNew","lineaNew","lot","setupCodeNew","particularNew","measureBorder","measurePitch","measureNotch","notes"].forEach(id => $(id).value = "");
  ["needleChange","notesStatus","threadNeedle","threadCrochet","needleType"].forEach(id => $(id).selectedIndex = 0);
  ["qualityStop","qualitySkipped","qualityNeedle"].forEach(id => setFastChoiceValue(id, "OK"));
  ["threadNeedleStatus","threadCrochetStatus","needleTypeStatus"].forEach(id => setFastChoiceValue(id, ""));
  $("shift").selectedIndex = 0;
  $("workstationSelect").selectedIndex = 0;
  $("lineaSelect").selectedIndex = 0;
  $("setupCodeSelect").selectedIndex = 0;
  $("stitchTypeAuto").value = "";
  $("particularAuto").value = "";
  setStatusField("measureBorderStatus", "");
  setStatusField("measurePitchStatus", "");
  setStatusField("measureNotchStatus", "");
  setStatusField("threadNeedleStatus", "");
  setStatusField("threadCrochetStatus", "");
  setStatusField("needleTypeStatus", "");
  focusIfEnabled("workstationSelect");
}
function clearForContinuousControl() {
  ["lot","measureBorder","measurePitch","measureNotch","notes"].forEach(id => $(id).value = "");
  $("notesStatus").selectedIndex = 0;
  ["threadNeedleStatus","threadCrochetStatus","needleTypeStatus"].forEach(id => setFastChoiceValue(id, ""));
  setStatusField("measureBorderStatus", "");
  setStatusField("measurePitchStatus", "");
  setStatusField("measureNotchStatus", "");
  applyMaterialValidation();
  focusIfEnabled("lot");
}
function saveControl() {
  applyToleranceAuto();
  applyMaterialValidation();
  const control = {
    id: crypto.randomUUID(),
    createdAt: dateTimeString(),
    qualityUser: currentUser.name,
    shift: $("shift").value,
    workstation: ($("workstationNew").value || $("workstationSelect").value || "").trim(),
    stitchType: ($("stitchTypeNew").value || $("stitchTypeAuto").value || "").trim(),
    linea: getSelectedLinea(),
    lot: $("lot").value.trim(),
    setupCode: ($("setupCodeNew").value || $("setupCodeSelect").value || "").trim(),
    particular: ($("particularNew").value || $("particularAuto").value || "").trim(),
    measureBorder: $("measureBorder").value,
    measureBorderStatus: $("measureBorderStatus").value,
    measurePitch: $("measurePitch").value,
    measurePitchStatus: $("measurePitchStatus").value,
    measureNotch: $("measureNotch").value,
    measureNotchStatus: $("measureNotchStatus").value,
    qualityStop: $("qualityStop").value,
    qualitySkipped: $("qualitySkipped").value,
    qualityNeedle: $("qualityNeedle").value,
    threadNeedle: $("threadNeedle").value.trim(),
    threadNeedleStatus: $("threadNeedleStatus").value,
    threadCrochet: $("threadCrochet").value.trim(),
    threadCrochetStatus: $("threadCrochetStatus").value,
    needleType: $("needleType").value.trim(),
    needleTypeStatus: $("needleTypeStatus").value,
    needleChange: $("needleChange").value,
    notes: $("notes").value.trim(),
    notesStatus: $("notesStatus").value,
    machineOperator: null,
    machineBadge: null,
    status: "Da validare"
  };

  const validationErrors = [];
  if (!control.workstation) validationErrors.push({ field: "workstationSelect", message: "Postazione obbligatoria." });
  if (!control.linea) validationErrors.push({ field: "lineaSelect", message: "Linea lavorazione obbligatoria." });
  if (!control.setupCode) validationErrors.push({ field: "setupCodeSelect", message: "Codice prodotto obbligatorio." });
  if (!control.lot) validationErrors.push({ field: "lot", message: "Lotto obbligatorio." });
  if (control.measureBorder === "") validationErrors.push({ field: "measureBorder", message: "Bordo cucitura obbligatorio." });
  if (control.measurePitch === "") validationErrors.push({ field: "measurePitch", message: "Passo punto obbligatorio." });
  if (control.measureNotch === "") validationErrors.push({ field: "measureNotch", message: "Disallineamento tacche obbligatorio." });
  if (!["OK","KO"].includes(control.qualityStop)) validationErrors.push({ field: "qualityStop", message: "Fermapunto obbligatorio." });
  if (!["OK","KO"].includes(control.qualitySkipped)) validationErrors.push({ field: "qualitySkipped", message: "Punti saltati obbligatorio." });
  if (!["OK","KO"].includes(control.qualityNeedle)) validationErrors.push({ field: "qualityNeedle", message: "Ago spuntato obbligatorio." });

  const availableLines = new Set((masterData.linee || []).map(v => String(v || "").trim()).filter(Boolean));
  if (control.linea && availableLines.size && !availableLines.has(control.linea)) {
    validationErrors.push({ field: "lineaSelect", message: "Linea non presente nelle anagrafiche." });
  }

  if (control.linea) {
    const tolerance = masterData.tolleranze?.[control.linea];
    if (!tolerance) {
      validationErrors.push({ field: "lineaSelect", message: "Linea senza tolleranze valide." });
    } else {
      if (control.measureBorder !== "" && !control.measureBorderStatus) {
        validationErrors.push({ field: "measureBorder", message: "Esito bordo non calcolato." });
      }
      if (control.measurePitch !== "" && !control.measurePitchStatus) {
        validationErrors.push({ field: "measurePitch", message: "Esito passo punto non calcolato." });
      }
      if (control.measureNotch !== "" && !control.measureNotchStatus) {
        validationErrors.push({ field: "measureNotch", message: "Esito disallineamento tacche non calcolato." });
      }
    }
  }

  if (control.workstation && control.linea) {
    const expectedLine = String(masterData.postazioneToLinea?.[control.workstation] || "").trim();
    if (expectedLine && expectedLine !== control.linea) {
      validationErrors.push({ field: "workstationSelect", message: `Mismatch postazione/linea: ${control.workstation} appartiene a ${expectedLine}.` });
    }
  }

  const gammaOrder = gammaOrders.find(x => x.lotto === control.lot || x.idOrdine === control.lot) || null;

  if (gammaOrder) {
    const gammaLine = String(gammaOrder.linea || "").trim();
    if (control.linea && gammaLine && control.linea !== gammaLine) {
      validationErrors.push({ field: "lot", message: "Lotto non compatibile con la linea della postazione." });
    }
    if (control.setupCode && String(gammaOrder.codice || "").trim() && control.setupCode !== String(gammaOrder.codice || "").trim()) {
      validationErrors.push({ field: "setupCodeSelect", message: "Il codice prodotto non coincide con l'ordine Gamma selezionato." });
    }
  } else if (control.setupCode && control.linea) {
    const codeMatch = (masterData.codici || []).some(item => String(item.codice || "").trim() === control.setupCode && String(item.linea || "").trim() === control.linea);
    if (!codeMatch) {
      validationErrors.push({ field: "setupCodeSelect", message: "Il codice prodotto non appartiene alla linea selezionata." });
    }
  }

  if (gammaOrder) {
    if (!control.threadNeedle) validationErrors.push({ field: "threadNeedle", message: "Codice filato ago obbligatorio." });
    if (!control.threadCrochet) validationErrors.push({ field: "threadCrochet", message: "Codice filato crochet obbligatorio." });
    if (!control.needleType) validationErrors.push({ field: "needleType", message: "Ago obbligatorio." });
    if (control.threadNeedle && !control.threadNeedleStatus) validationErrors.push({ field: "threadNeedle", message: "Esito codice filato ago non calcolato." });
    if (control.threadCrochet && !control.threadCrochetStatus) validationErrors.push({ field: "threadCrochet", message: "Esito codice filato crochet non calcolato." });
    if (control.needleType && !control.needleTypeStatus) validationErrors.push({ field: "needleType", message: "Esito ago non calcolato." });
  } else {
    const materialRules = getMaterialRules(control.linea);
    if (control.linea && materialRules) {
      if (!control.threadNeedle) validationErrors.push({ field: "threadNeedle", message: "Codice filato ago obbligatorio." });
      if (!control.threadCrochet) validationErrors.push({ field: "threadCrochet", message: "Codice filato crochet obbligatorio." });
      if (!control.needleType) validationErrors.push({ field: "needleType", message: "Ago obbligatorio." });
      if (control.threadNeedle && !control.threadNeedleStatus) validationErrors.push({ field: "threadNeedle", message: "Esito codice filato ago non calcolato." });
      if (control.threadCrochet && !control.threadCrochetStatus) validationErrors.push({ field: "threadCrochet", message: "Esito codice filato crochet non calcolato." });
      if (control.needleType && !control.needleTypeStatus) validationErrors.push({ field: "needleType", message: "Esito ago non calcolato." });
    }
  }

  if (validationErrors.length) {
    const first = validationErrors[0];
    $("saveMsg").className = "msg err";
    $("saveMsg").textContent = first.message;
    focusIfEnabled(first.field);
    return;
  }

  control.outcome = overallOutcome(control);
  controls.push(control);
  createAnomaliesFor(control);
  saveStorage();
  queueControlForSync(control);
  renderHistory();
  renderAnomalies();
  renderKpis();
  clearForContinuousControl();
  $("saveMsg").className = "msg ok";
  $("saveMsg").textContent = "Controllo salvato.";
}
function toIsoDateParts(localDateTime) {
  if (!localDateTime) {
    return { date: "", time: "", year: "", month: "", day: "", hour: "", minute: "" };
  }
  const [datePart = "", timePart = ""] = String(localDateTime).split(", ");
  const [day = "", month = "", year = ""] = datePart.split("/");
  const [hour = "", minute = ""] = timePart.split(":");
  const pad = (v) => String(v || "").padStart(2, "0");
  return {
    date: year && month && day ? `${year}-${pad(month)}-${pad(day)}` : "",
    time: hour && minute ? `${pad(hour)}:${pad(minute)}` : "",
    year,
    month: pad(month),
    day: pad(day),
    hour: pad(hour),
    minute: pad(minute)
  };
}
function csvValue(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}
function badgeForCsv(value) {
  const badge = String(value ?? "").trim();
  return badge ? `'${badge}` : "";
}
function notesStatusForCsv(notes) {
  return String(notes ?? "").trim() ? "OK" : "";
}
function koCounters(control) {
  const measureStatuses = [control.measureBorderStatus, control.measurePitchStatus, control.measureNotchStatus];
  const qualityStatuses = [control.qualityStop, control.qualitySkipped, control.qualityNeedle];
  const materialStatuses = [control.threadNeedleStatus, control.threadCrochetStatus, control.needleTypeStatus];
  const countKo = (arr) => arr.filter(v => String(v || "").trim().toUpperCase() === "KO").length;
  const measures = countKo(measureStatuses);
  const quality = countKo(qualityStatuses);
  const materials = countKo(materialStatuses);
  return { measures, quality, materials, total: measures + quality + materials };
}
function exportCsv() {
  const delimiter = ";";
  const header = [
    "Control_ID","DataOra_Locale","Data_ISO","Ora","Anno","Mese","Giorno","Ora_Num","Minuto_Num",
    "Operatrice_Qualita","Turno","Postazione","Tipo_Cucitura","Linea","Particolare_in_Produzione","Codice_Prodotto","Lotto",
    "Bordo_Cucitura_Valore","Bordo_Cucitura_Esito","Passo_Punto_Valore","Passo_Punto_Esito","Disallineamento_Tacche_Valore","Disallineamento_Tacche_Esito",
    "Fermapunto_Esito","Punti_Saltati_Esito","Ago_Spuntato_Esito",
    "Codice_Filato_Ago","Codice_Filato_Ago_Esito","Codice_Filato_Crochet","Codice_Filato_Crochet_Esito","Ago_Codice","Ago_Esito","Cambio_Ago",
    "Note","Esito_Note","Esito_Complessivo","KO_Flag","OK_Flag",
    "Numero_KO_Misure","Numero_KO_Qualitativi","Numero_KO_Materiali","Numero_KO_Totali",
    "Stato_Validazione","Validato_Flag","Badge_Macchinista","DataOra_Attestazione","Attestato_Flag",
    "Numero_Anomalie","Anomalia_Bordo","Anomalia_Passo","Anomalia_Tacche","Anomalia_Fermapunto","Anomalia_Punti_Saltati","Anomalia_Ago_Spuntato","Anomalia_Filato_Ago","Anomalia_Filato_Crochet","Anomalia_Ago","Anomalia_Note"
  ];

  const rows = controls.map(c => {
    const parts = toIsoDateParts(c.createdAt);
    const anomaliesForControl = anomalies.filter(a => a.controlId === c.id);
    const anomalyTypes = new Set(anomaliesForControl.map(a => String(a.type || "").trim()));
    const ko = koCounters(c);
    return [
      c.id || "",
      c.createdAt || "",
      parts.date,
      parts.time,
      parts.year,
      parts.month,
      parts.day,
      parts.hour,
      parts.minute,
      c.qualityUser || "",
      c.shift || "",
      c.workstation || "",
      c.stitchType || "",
      c.linea || "",
      c.particular || "",
      c.setupCode || "",
      c.lot || "",
      c.measureBorder || "",
      c.measureBorderStatus || "",
      c.measurePitch || "",
      c.measurePitchStatus || "",
      c.measureNotch || "",
      c.measureNotchStatus || "",
      c.qualityStop || "",
      c.qualitySkipped || "",
      c.qualityNeedle || "",
      c.threadNeedle || "",
      c.threadNeedleStatus || "",
      c.threadCrochet || "",
      c.threadCrochetStatus || "",
      c.needleType || "",
      c.needleTypeStatus || "",
      c.needleChange || "",
      c.notes || "",
      notesStatusForCsv(c.notes),
      c.outcome || "",
      c.outcome === "KO" ? 1 : 0,
      c.outcome === "OK" ? 1 : 0,
      ko.measures,
      ko.quality,
      ko.materials,
      ko.total,
      c.status || "",
      c.status === "Validato" ? 1 : 0,
      badgeForCsv(c.machineBadge),
      c.machineAttestedAt || "",
      c.machineBadge ? 1 : 0,
      anomaliesForControl.length,
      anomalyTypes.has("Bordo cucitura KO") ? 1 : 0,
      anomalyTypes.has("Passo punto KO") ? 1 : 0,
      anomalyTypes.has("Disallineamento tacche KO") ? 1 : 0,
      anomalyTypes.has("Fermapunto non conforme") ? 1 : 0,
      anomalyTypes.has("Punti saltati") ? 1 : 0,
      anomalyTypes.has("Ago spuntato") ? 1 : 0,
      anomalyTypes.has("Codice filato ago non conforme") ? 1 : 0,
      anomalyTypes.has("Codice filato crochet non conforme") ? 1 : 0,
      anomalyTypes.has("Ago non conforme") ? 1 : 0,
      anomalyTypes.has("Note KO") ? 1 : 0
    ].map(csvValue).join(delimiter);
  });

  const csv = "\ufeff" + [header.join(delimiter), ...rows].join("\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8;"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `qc_controlli_kpi_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}

function exportKpi() {
  const delimiter = ";";
  const header = [
    "Data","Ora","Turno","Linea","Postazione","Macchina","Risorsa_Qualita","Badge_Macchinista",
    "Codice_Prodotto","Particolare","Lotto","Esito_Complessivo","Stato_Controllo",
    "Categoria_KO","Tipo_KO","Dettaglio_KO","Valore_Rilevato","Valore_Atteso","Severita","Anomalia_Stato"
  ];

  const koDetailsForControl = (c) => {
    const linea = c.linea || "";
    const tol = masterData.tolleranze?.[linea];
    const material = masterData.materiali?.[linea] || {threadNeedle: [], threadCrochet: [], needleType: []};
    const details = [];
    const add = (categoria, tipo, dettaglio, valore, atteso, severita) => details.push({categoria, tipo, dettaglio, valore, atteso, severita});

    if (c.measureBorderStatus === "KO") add("MISURE", "BORDO_CUCITURA", "Bordo cucitura fuori tolleranza", c.measureBorder || "", tol ? `${tol.bordoMin} ÷ ${tol.bordoMax}` : "", "Grave");
    if (c.measurePitchStatus === "KO") add("MISURE", "PASSO_PUNTO", "Passo punto fuori tolleranza", c.measurePitch || "", tol ? `${tol.passoMin} ÷ ${tol.passoMax}` : "", "Grave");
    if (c.measureNotchStatus === "KO") add("MISURE", "DISALLINEAMENTO_TACCHE", "Disallineamento tacche fuori tolleranza", c.measureNotch || "", tol ? `${tol.taccheMin} ÷ ${tol.taccheMax}` : "", "Grave");

    if (c.qualityStop === "KO") add("QUALITATIVI", "FERMAPUNTO", "Fermapunto non conforme", c.qualityStop || "", "OK", "Media");
    if (c.qualitySkipped === "KO") add("QUALITATIVI", "PUNTI_SALTATI", "Punti saltati rilevati", c.qualitySkipped || "", "OK", "Media");
    if (c.qualityNeedle === "KO") add("QUALITATIVI", "AGO_SPUNTATO", "Ago spuntato rilevato", c.qualityNeedle || "", "OK", "Media");

    if (c.threadNeedleStatus === "KO") add("MATERIALI", "FILATO_AGO", "Codice filato ago non conforme", c.threadNeedle || "", material.threadNeedle?.join("/") || "", "Media");
    if (c.threadCrochetStatus === "KO") add("MATERIALI", "FILATO_CROCHET", "Codice filato crochet non conforme", c.threadCrochet || "", material.threadCrochet?.join("/") || "", "Media");
    if (c.needleTypeStatus === "KO") add("MATERIALI", "AGO", "Ago non conforme", c.needleType || "", material.needleType?.join("/") || "", "Media");

    return details;
  };

  const rows = [];
  controls.forEach(c => {
    const parts = toIsoDateParts(c.createdAt);
    const anomaliesForControl = anomalies.filter(a => a.controlId === c.id);
    const detailRows = koDetailsForControl(c);

    if (!detailRows.length) {
      rows.push([
        parts.date,
        parts.time,
        c.shift || "",
        c.linea || "",
        c.workstation || "",
        c.workstation || "",
        c.qualityUser || "",
        badgeForCsv(c.machineBadge),
        c.setupCode || "",
        c.particular || "",
        c.lot || "",
        c.outcome || "",
        c.status || "",
        "",
        "",
        "",
        "",
        "",
        "",
        ""
      ].map(csvValue).join(delimiter));
      return;
    }

    detailRows.forEach(d => {
      rows.push([
        parts.date,
        parts.time,
        c.shift || "",
        c.linea || "",
        c.workstation || "",
        c.workstation || "",
        c.qualityUser || "",
        badgeForCsv(c.machineBadge),
        c.setupCode || "",
        c.particular || "",
        c.lot || "",
        c.outcome || "",
        c.status || "",
        d.categoria,
        d.tipo,
        d.dettaglio,
        d.valore,
        d.atteso,
        d.severita,
        anomaliesForControl.length ? anomaliesForControl[0].status || "Aperta" : (c.outcome === "KO" ? "Aperta" : "")
      ].map(csvValue).join(delimiter));
    });
  });

  const csv = "﻿" + [header.join(delimiter), ...rows].join("\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8;"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `qc_kpi_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}

function printApp() {
  window.print();
}
function backupJson() {
  const blob = new Blob([JSON.stringify({controls, anomalies, masterData}, null, 2)], {type:"application/json"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "qc_backup.json";
  a.click();
}

function getAnomalyById(id) {
  return anomalies.find(x => x.id === id) || null;
}
function renderAnomalyActionsHistory(anomaly) {
  const box = $("anomalyActionsHistory");
  if (!box) return;
  const actions = Array.isArray(anomaly?.actions) ? anomaly.actions : [];
  if (!actions.length) {
    box.innerHTML = '<div class="detail-row"><div class="detail-key">Storico</div><div>Nessuna azione registrata.</div></div>';
    return;
  }
  box.innerHTML = actions.slice().reverse().map(a => `
    <div class="detail-row"><div class="detail-key">${a.createdAt}</div><div><strong>${a.owner || '-'} </strong><br>${a.text || '-'}${a.verifyOutcome ? `<br>Verifica dopo azione: <span class="badge ${a.verifyOutcome === "OK" ? "badge-ok" : (a.verifyOutcome === "KO" ? "badge-ko" : "badge-state")}">${a.verifyOutcome}</span>` : ''}${a.notes ? `<br>Note: ${a.notes}` : ''}</div></div>
  `).join('');
}
function applyVerifyOutcomeVisualState() {
  const sel = $("anomalyVerifyOutcome");
  if (!sel) return;
  sel.classList.remove("state-ok","state-ko","state-progress");
  const v = String(sel.value || "").trim().toUpperCase();
  if (v === "OK") sel.classList.add("state-ok");
  else if (v === "KO") sel.classList.add("state-ko");
  else sel.classList.add("state-progress");
}

function getDerivedAnomalyStatus(anomaly) {
  if (!anomaly) return "Aperta";
  if (anomaly.status === "Chiusa" || anomaly.closedAt) return "Chiusa";
  const hasActions = Array.isArray(anomaly.actions) && anomaly.actions.length > 0;
  return hasActions ? "In gestione" : "Aperta";
}
function refreshAnomalyCloseButton() {
  const anomaly = getAnomalyById(currentAnomalyId);
  if (!anomaly) return;
  const text = String($("anomalyActionText").value || "").trim();
  const verifyOutcome = String($("anomalyVerifyOutcome").value || "IN CORSO").trim().toUpperCase();
  $("btnCloseAnomaly").disabled = anomaly.status === "Chiusa" || !text || verifyOutcome !== "OK";
}
function refreshAnomalyModal() {
  const anomaly = getAnomalyById(currentAnomalyId);
  if (!anomaly) return;
  anomaly.status = getDerivedAnomalyStatus(anomaly);
  $("anomalyStatusView").value = anomaly.status || "Aperta";
  const related = controls.find(c => c.id === anomaly.controlId);
  const section = (title, rows, full=false) => `
    <div class="detail-section ${full ? 'full' : ''}">
      <h4>${title}</h4>
      ${rows.map(([k,v]) => `<div class="detail-row"><div class="detail-key">${k}</div><div>${v || '-'}</div></div>`).join('')}
    </div>`;
  $("anomalyContent").innerHTML = [
    section("Anomalia", [
      ["Aperta il", anomaly.openedAt],
      ["Lotto", anomaly.lot],
      ["Tipo", anomaly.type],
      ["Severità", anomaly.severity],
      ["Stato", anomaly.status],
      ["Chiusa il", anomaly.closedAt || '-']
    ]),
    section("Controllo collegato", related ? [
      ["Postazione", related.workstation],
      ["Linea", related.linea],
      ["Codice prodotto", related.setupCode],
      ["Esito controllo", related.outcome]
    ] : [["Controllo", "Non trovato"]])
  ].join('');
  renderAnomalyActionsHistory(anomaly);
  applyVerifyOutcomeVisualState();
  refreshAnomalyCloseButton();
}
function openAnomalyModal(id) {
  currentAnomalyId = id;
  $("anomalyActionText").value = "";
  $("anomalyActionNotes").value = "";
  $("anomalyVerifyOutcome").value = "IN CORSO";
  $("anomalyOwner").value = currentUser?.name || "";
  $("anomalyMsg").textContent = "";
  applyVerifyOutcomeVisualState();
  refreshAnomalyModal();
  $("anomalyModal").classList.remove("hidden");
}
function closeAnomalyModal() {
  currentAnomalyId = null;
  $("anomalyModal").classList.add("hidden");
}
function saveAnomalyAction(shouldClose=false) {
  const anomaly = getAnomalyById(currentAnomalyId);
  if (!anomaly) return;
  const text = String($("anomalyActionText").value || "").trim();
  const owner = String($("anomalyOwner").value || "").trim();
  const verifyOutcome = String($("anomalyVerifyOutcome").value || "IN CORSO").trim().toUpperCase();
  const notes = String($("anomalyActionNotes").value || "").trim();
  if (!text) {
    $("anomalyMsg").className = "msg err";
    $("anomalyMsg").textContent = "Azione correttiva obbligatoria.";
    return;
  }
  if (!owner) {
    $("anomalyMsg").className = "msg err";
    $("anomalyMsg").textContent = "Responsabile obbligatorio.";
    return;
  }
  if (shouldClose && verifyOutcome !== "OK") {
    $("anomalyMsg").className = "msg err";
    $("anomalyMsg").textContent = "Per chiudere devi impostare 'Verifica dopo azione' su OK.";
    return;
  }
  anomaly.actions = Array.isArray(anomaly.actions) ? anomaly.actions : [];
  anomaly.actions.push({
    id: crypto.randomUUID(),
    text,
    owner,
    verifyOutcome,
    notes,
    createdAt: dateTimeString(),
    createdBy: currentUser?.name || ""
  });
  anomaly.takenInChargeAt = anomaly.takenInChargeAt || dateTimeString();
  anomaly.takenInChargeBy = anomaly.takenInChargeBy || owner;
  anomaly.status = shouldClose ? "Chiusa" : "In gestione";
  if (shouldClose) {
    anomaly.closedAt = dateTimeString();
    anomaly.closedBy = currentUser?.name || owner;
  } else {
    anomaly.closedAt = "";
    anomaly.closedBy = "";
  }
  saveStorage();
  const related = controls.find(c => c.id === anomaly.controlId);
  if (related) queueControlForSync(related);
  renderAnomalies();
  renderKpis();
  $("anomalyActionText").value = "";
  $("anomalyActionNotes").value = "";
  $("anomalyVerifyOutcome").value = "IN CORSO";
  refreshAnomalyModal();
  $("anomalyMsg").className = "msg ok";
  $("anomalyMsg").textContent = shouldClose ? "Anomalia chiusa." : "Aggiornamento salvato.";
}
function closeAnomalyRecord() {
  saveAnomalyAction(true);
}

function openDetailModal(id) {
  const c = controls.find(x => x.id === id);
  if (!c) return;

  const section = (title, rows, full=false) => `
    <div class="detail-section ${full ? 'full' : ''}">
      <h4>${title}</h4>
      ${rows.map(([k,v]) => `<div class="detail-row"><div class="detail-key">${k}</div><div>${v || "-"}</div></div>`).join("")}
    </div>
  `;

  const html = [
    section("Intestazione", [
      ["Data/Ora", c.createdAt],
      ["Operatrice qualità", c.qualityUser],
      ["Turno", c.shift],
      ["Postazione", c.workstation],
      ["Tipo cucitura", c.stitchType],
      ["Linea lavorazione", c.linea || "-"],
      ["Lotto", c.lot],
      ["Codice prodotto", c.setupCode],
      ["Particolare in produzione", c.particular || "-"]
    ]),
    section("Misure", [
      ["Bordo cucitura", `${c.measureBorder || "-"} · ${c.measureBorderStatus}`],
      ["Passo punto", `${c.measurePitch || "-"} · ${c.measurePitchStatus}`],
      ["Disallineamento tacche", `${c.measureNotch || "-"} · ${c.measureNotchStatus}`]
    ]),
    section("Controlli qualitativi", [
      ["Fermapunto", c.qualityStop],
      ["Punti saltati", c.qualitySkipped],
      ["Ago spuntato", c.qualityNeedle]
    ]),
    section("Materiali e settaggi", [
      ["Codice filato ago", `${c.threadNeedle || "-"} · ${c.threadNeedleStatus || "-"}`],
      ["Codice filato crochet", `${c.threadCrochet || "-"} · ${c.threadCrochetStatus || "-"}`],
      ["Ago", `${c.needleType || "-"} · ${c.needleTypeStatus || "-"}`],
      ["Cambio ago", c.needleChange]
    ]),
    section("Chiusura e attestazione", [
      ["Esito complessivo", c.outcome],
      ["Stato validazione", c.status],
      ["Badge macchinista", c.machineBadge || "Non attestata"],
      ["Data/Ora attestazione", c.machineAttestedAt || "-"]
    ]),
    section("Note", [
      ["Note", c.notes || "-"],
      ["Esito note", c.notesStatus]
    ], true)
  ].join("");

  $("detailContent").innerHTML = html;
  $("detailModal").classList.remove("hidden");
}
function closeDetailModal() {
  $("detailModal").classList.add("hidden");
}
function openAttestModal(id) {
  currentAttestControlId = id;
  $("scanStatus").textContent = "Stato: pronto";
  $("scanStatus").className = "msg";
  $("scanResult").textContent = "";
  $("scanResult").className = "msg";
  $("manualBadge").value = "";
  $("manualMachineName").value = "";
  $("manualMachineBadgeNew").value = "";
  $("machineBadgeSelect").selectedIndex = 0;
  $("attestModal").classList.remove("hidden");
}
async function closeAttestModal() {
  await stopScanner();
  $("attestModal").classList.add("hidden");
}
function confirmBadge(raw) {
  const selectedBadge = String($("machineBadgeSelect").value || "").trim();
  const manualNewBadge = String($("manualMachineBadgeNew").value || "").trim();
  const rawInput = raw || manualNewBadge;
  const codeKey = selectedBadge ? badgeKey(selectedBadge) : badgeKey(rawInput);
  const control = controls.find(x => x.id === currentAttestControlId);
  if (!control) {
    $("scanResult").textContent = "Controllo non trovato.";
    $("scanResult").className = "msg err";
    return;
  }

  let operatorName = null;
  let resolvedBadge = selectedBadge || String(rawInput || "").trim();
  const found = (masterData.macchiniste || []).find(x => badgeKey(x.badge) === codeKey) || null;
  if (found) {
    operatorName = found.nome;
    resolvedBadge = String(found.badge || resolvedBadge).trim();
    $("machineBadgeSelect").value = resolvedBadge;
    $("manualMachineName").value = operatorName;
    $("manualBadge").value = codeKey;
  }
  if (!found && BADGES[codeKey]) operatorName = BADGES[codeKey];
  if (!operatorName && $("manualMachineName").value.trim() && $("manualMachineBadgeNew").value.trim()) {
    operatorName = $("manualMachineName").value.trim();
    resolvedBadge = String($("manualMachineBadgeNew").value || resolvedBadge).trim();
  }

  if (!codeKey || !operatorName) {
    $("scanStatus").textContent = "Stato: codice letto ma non associato";
    $("scanResult").textContent = `Codice acquisito: ${codeKey}. Macchinista non riconosciuta.`;
    $("scanResult").className = "msg err";
    return;
  }

  control.machineBadge = resolvedBadge;
  control.machineOperator = operatorName;
  control.machineAttestedAt = dateTimeString();
  control.status = "Validato";
  saveStorage();
  queueControlForSync(control);
  renderHistory();
  renderKpis();
  $("scanStatus").textContent = "Stato: codice acquisito ✓";
  $("scanResult").textContent = `Attestazione acquisita: ${resolvedBadge} · ${operatorName}`;
  $("scanResult").className = "msg ok";
  closeAttestModal();
}
let scanLock = false;
let scanFeedbackTimer = null;
function pulseReaderFeedback() {
  const reader = $("reader");
  if (!reader) return;
  reader.classList.add("scan-ok");
  clearTimeout(scanFeedbackTimer);
  scanFeedbackTimer = setTimeout(() => reader.classList.remove("scan-ok"), 700);
}
async function startScanner() {
  $("scanResult").textContent = "";
  $("scanResult").className = "msg";
  $("scanStatus").textContent = "Stato: avvio scanner...";
  $("scanStatus").className = "msg";
  if (typeof Html5Qrcode === "undefined") {
    $("scanStatus").textContent = "Stato: libreria QR non caricata";
    $("scanStatus").className = "msg err";
    return;
  }
  if (scannerRunning) return;
  try {
    html5QrCode = new Html5Qrcode("reader");
    const cameras = await Html5Qrcode.getCameras();
    if (!cameras || cameras.length === 0) {
      $("scanStatus").textContent = "Stato: nessuna camera trovata";
      $("scanStatus").className = "msg err";
      return;
    }
    const preferredCamera = cameras.find(c => /back|rear|environment/i.test(`${c.label || ""} ${c.id || ""}`)) || cameras[0];
    scannerRunning = true;
    scanLock = false;
    $("scanStatus").textContent = "Stato: camera attiva";
    await html5QrCode.start(
      preferredCamera.id,
      { fps: 10, qrbox: { width: 260, height: 260 } },
      decodedText => {
        if (scanLock) return;
        scanLock = true;
        const normalized = badgeKey(decodedText);
        $("scanStatus").textContent = "Stato: codice letto ✓";
        $("scanStatus").className = "msg ok";
        $("scanResult").textContent = `Codice acquisito: ${normalized}`;
        $("scanResult").className = "msg ok";
        $("manualBadge").value = normalized;
        pulseReaderFeedback();
        setTimeout(() => {
          confirmBadge(decodedText);
          scanLock = false;
        }, 250);
      },
      _err => {}
    );
  } catch (err) {
    scannerRunning = false;
    $("scanStatus").textContent = "Stato: errore scanner";
    $("scanStatus").className = "msg err";
    $("scanResult").textContent = String(err);
    $("scanResult").className = "msg err";
  }
}
async function stopScanner() {
  if (!html5QrCode || !scannerRunning) return;
  try {
    await html5QrCode.stop();
    await html5QrCode.clear();
  } catch(e) {
  } finally {
    scannerRunning = false;
    $("scanStatus").textContent = "Stato: scanner fermato";
    $("reader").innerHTML = "";
  }
}
function importBackupFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const parsed = JSON.parse(e.target.result);
      if (!parsed.controls || !parsed.anomalies) throw new Error("File non valido");
      controls = parsed.controls;
      anomalies = (parsed.anomalies || []).map(a => ({ actions: [], status: "Aperta", ...a, actions: Array.isArray(a.actions) ? a.actions : [] }));
      if (parsed.masterData) masterData = parsed.masterData;
      syncQueue = [];
      syncMeta.lastError = "";
      saveStorage();
      populateMasterSelects();
      updateAutoType();
      updateToleranceInfo();
      renderHistory();
      renderAnomalies();
      renderKpis();
      renderSyncStatus();
      $("saveMsg").className = "msg ok";
      $("saveMsg").textContent = "Backup importato correttamente.";
    } catch (err) {
      $("saveMsg").className = "msg err";
      $("saveMsg").textContent = "Importazione fallita.";
    }
  };
  reader.readAsText(file);
}
function login() {
  const u = $("loginUser").value.trim();
  const p = $("loginPass").value;
  const user = systemUsers[u];
  if (!user || String(user.password) !== String(p)) {
    $("loginMsg").textContent = "Credenziali non valide.";
    return;
  }
  $("loginMsg").textContent = "";
  currentUser = user;
  $("loginView").classList.add("hidden");
  $("appView").classList.remove("hidden");
  $("loggedUser").textContent = currentUser.name;
  $("todayDate").textContent = todayString();
  renderSyncStatus();
  updateDataSourceStatus();
  focusIfEnabled("workstationSelect");
}
function logout() {
  currentUser = null;
  $("appView").classList.add("hidden");
  $("loginView").classList.remove("hidden");
  $("loginPass").value = "";
  $("loginMsg").textContent = "";
  updateLoginScreenState();
}
window.openDetailModal = openDetailModal;
window.closeDetailModal = closeDetailModal;
window.openAnomalyModal = openAnomalyModal;
window.closeAnomalyModal = closeAnomalyModal;
window.openAttestModal = openAttestModal;
window.closeAttestModal = closeAttestModal;

window.addEventListener("DOMContentLoaded", () => {
  // Bind critical actions first
  $("btnLogin").addEventListener("click", login);
  $("btnBootstrapImport").addEventListener("click", () => $("bootstrapAnagraficheFile").click());
  $("bootstrapAnagraficheFile").addEventListener("change", (e) => importBootstrapAnagraficheFile(e.target.files[0]));
  $("btnLogout").addEventListener("click", logout);
  $("loginPass").addEventListener("keydown", (e) => { if (e.key === "Enter") login(); });
  $("btnSave").addEventListener("click", saveControl);
  $("btnSaveAnomalyAction").addEventListener("click", () => saveAnomalyAction(false));
  $("btnCloseAnomaly").addEventListener("click", closeAnomalyRecord);
  $("anomalyVerifyOutcome").addEventListener("change", () => { applyVerifyOutcomeVisualState(); refreshAnomalyCloseButton(); });
  $("anomalyActionText").addEventListener("input", refreshAnomalyCloseButton);
  $("btnExportCsv").addEventListener("click", exportCsv);
  $("btnExportKpi").addEventListener("click", exportKpi);
  $("btnPrint").addEventListener("click", printApp);
  $("btnImportGamma").addEventListener("click", () => $("importGammaFile").click());
  $("importGammaFile").addEventListener("change", (e) => importGammaFile(e.target.files[0]));
  $("btnImportAnagrafiche").addEventListener("click", () => $("importAnagraficheFile").click());
  if ($("btnToggleManualEntry")) $("btnToggleManualEntry").addEventListener("click", () => toggleManualEntryPanel());
  $("importAnagraficheFile").addEventListener("change", (e) => importAnagraficheFile(e.target.files[0]));
  $("workstationSelect").addEventListener("change", () => { updateAutoType(); syncLineFromWorkstation(); populateGammaLots(); resetGammaDependentFields(); });
  $("workstationNew").addEventListener("input", () => { updateAutoType(); syncLineFromWorkstation(); populateGammaLots(); resetGammaDependentFields(); });
  $("lineaSelect").addEventListener("change", () => { populateCodeSelect(); updateToleranceInfo(); populateGammaLots(); resetGammaDependentFields(); });
  $("lineaNew").addEventListener("input", () => { populateCodeSelect(); updateToleranceInfo(); populateGammaLots(); resetGammaDependentFields(); });
  $("measureBorder").addEventListener("input", applyToleranceAuto);
  $("measurePitch").addEventListener("input", applyToleranceAuto);
  $("measureNotch").addEventListener("input", applyToleranceAuto);
  $("threadNeedle").addEventListener("change", applyMaterialValidation);
  $("threadCrochet").addEventListener("change", applyMaterialValidation);
  $("needleType").addEventListener("change", applyMaterialValidation);
  $("btnStartScan").addEventListener("click", startScanner);
  $("btnStopScan").addEventListener("click", stopScanner);
  $("btnConfirmBadge").addEventListener("click", () => confirmBadge($("manualBadge").value));
  $("btnRetrySync").addEventListener("click", attemptSync);
  window.addEventListener("online", () => { syncMeta.lastError = ""; renderSyncStatus(); scheduleSync(); });
  window.addEventListener("offline", renderSyncStatus);
  initSmartFocus();
  initFastChoices();

  try {
    loadStoredUsers();
    loadStorage();
    populateMasterSelects();
    updateAutoType();
    syncLineFromWorkstation();
    updateParticularAuto();
    updateToleranceInfo();
    renderHistory();
    renderAnomalies();
    renderKpis();
    renderSyncStatus();
    scheduleSync();
    updateDataSourceStatus();
    updateLoginScreenState();
    toggleManualEntryPanel(false);
  } catch (err) {
    console.error("Init error:", err);
    const msg = $("loginMsg");
    if (msg) msg.textContent = "Inizializzazione parziale. Il login resta disponibile.";
  }
});

// ===== Data source status (Gamma + QC) =====
window.__gammaLoaded = false;
window.__qcLoaded = false;

function updateDataSourceStatus(){
  const gammaBadge = $("badgeGamma");
  const qcBadge = $("badgeQc");
  if (!gammaBadge || !qcBadge) return;

  window.__gammaLoaded = Array.isArray(gammaOrders) && gammaOrders.length > 0;
  window.__qcLoaded = Array.isArray(masterData.postazioni) && masterData.postazioni.length > 0
    && Array.isArray(masterData.linee) && masterData.linee.length > 0
    && masterData.tolleranze && Object.keys(masterData.tolleranze).length > 0;

  gammaBadge.className = "top-badge";
  qcBadge.className = "top-badge";
  gammaBadge.textContent = window.__gammaLoaded ? `Gamma: OK (${gammaOrders.length})` : "Gamma: NO";
  qcBadge.textContent = window.__qcLoaded ? "QC: OK" : "QC: NO";
  gammaBadge.classList.add(window.__gammaLoaded ? "badge-success-soft" : "badge-neutral-soft");
  qcBadge.classList.add(window.__qcLoaded ? "badge-success-soft" : "badge-neutral-soft");

  const saveBtn = $("btnSave");
  if (saveBtn) saveBtn.disabled = !(window.__gammaLoaded && window.__qcLoaded);
}

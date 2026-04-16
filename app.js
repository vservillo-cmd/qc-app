
const USERS = {
  "qualita1": { password: "1234", name: "Operatrice Controllo Qualità 1" }
};

const BADGES = {
  "MAC001": "Anna Rossi",
  "MAC002": "Maria Bianchi",
  "MAC003": "Giulia Esposito"
};

let currentUser = null;
let controls = [];
let anomalies = [];
let currentAttestControlId = null;
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
function getSyncEndpoint() {
  return (localStorage.getItem("qc_sync_endpoint") || DEFAULT_SYNC_ENDPOINT || "").trim();
}
function saveStorage() {
  localStorage.setItem("qc_controls", JSON.stringify(controls));
  localStorage.setItem("qc_anomalies", JSON.stringify(anomalies));
  localStorage.setItem("qc_master_data", JSON.stringify(masterData));
  localStorage.setItem("qc_sync_queue", JSON.stringify(syncQueue));
  localStorage.setItem("qc_sync_meta", JSON.stringify(syncMeta));
}
function loadStorage() {
  controls = JSON.parse(localStorage.getItem("qc_controls") || "[]");
  anomalies = JSON.parse(localStorage.getItem("qc_anomalies") || "[]");
  masterData = JSON.parse(localStorage.getItem("qc_master_data") || '{"postazioni":[],"linee":[],"codici":[],"macchiniste":[],"lineeDettaglio":{},"postazioneToLinea":{},"tolleranze":{}}');
  syncQueue = JSON.parse(localStorage.getItem("qc_sync_queue") || "[]");
  syncMeta = JSON.parse(localStorage.getItem("qc_sync_meta") || '{"lastError":"","lastAttemptAt":null,"lastSuccessAt":null}');
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
  const dot = $("syncDot");
  const label = $("syncLabel");
  const detail = $("syncDetail");
  const retry = $("btnRetrySync");
  if (!dot || !label || !detail || !retry) return;

  dot.classList.remove("sync-ok","sync-pending","sync-error");
  const pending = syncQueue.length;
  const online = navigator.onLine !== false;
  const hasError = !!syncMeta.lastError;

  if (hasError) {
    dot.classList.add("sync-error");
    label.textContent = "Errore sync";
    detail.textContent = syncMeta.lastError;
    retry.classList.remove("hidden");
    return;
  }

  if (pending > 0) {
    dot.classList.add("sync-pending");
    label.textContent = `In attesa (${pending} controll${pending === 1 ? "o" : "i"})`;
    if (!online) {
      detail.textContent = `Offline: ${pending} controll${pending === 1 ? "o" : "i"} salvati in locale.`;
    } else if (!getSyncEndpoint()) {
      detail.textContent = `Salvati in locale. Endpoint sync non configurato.`;
    } else if (syncInFlight) {
      detail.textContent = `Sincronizzazione in corso...`;
    } else {
      detail.textContent = `In coda per la sincronizzazione.`;
    }
    retry.classList.toggle("hidden", !online || !getSyncEndpoint());
    return;
  }

  dot.classList.add("sync-ok");
  label.textContent = "Sincronizzato";
  detail.textContent = syncMeta.lastSuccessAt
    ? `Ultima sync OK: ${syncMeta.lastSuccessAt}`
    : "Nessun controllo in attesa.";
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
function normalizeCode(raw) {
  if (!raw) return "";
  const txt = String(raw).trim().toUpperCase();
  const match = txt.match(/MAC\d{3}/);
  return match ? match[0] : txt;
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

      const errors = [];
      if (!postRows.length) errors.push("Foglio Postazioni mancante o vuoto");
      if (!lineRows.length) errors.push("Foglio Elenco Linee mancante o vuoto");
      if (!codRows.length) errors.push("Foglio Codici_Prodotto mancante o vuoto");
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
      saveStorage();
      populateMasterSelects();
      updateAutoType();
      populateMasterSelects();
      updateToleranceInfo();
      $("saveMsg").className = "msg";
      $("saveMsg").textContent = "";
    } catch (err) {
      $("saveMsg").className = "msg err";
      $("saveMsg").textContent = "Import anagrafiche fallito.";
    }
  };
  reader.readAsArrayBuffer(file);
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
  $("workstationSelect").addEventListener("change", () => focusIfEnabled("lot"));
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
  lotInput.addEventListener("change", moveLotToSetupCode);
  lotInput.addEventListener("blur", moveLotToSetupCode);

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
  const linea = String(getSelectedLinea() || "").trim();
  const rules = getMaterialRules(linea);
  populateSelect("threadNeedle", rules?.threadNeedle || []);
  populateSelect("threadCrochet", rules?.threadCrochet || []);
  populateSelect("needleType", rules?.needleType || []);
  applyMaterialValidation();
}
function applyMaterialValidation() {
  const linea = String(getSelectedLinea() || "").trim();
  const rules = getMaterialRules(linea);
  if (!rules) {
    setStatusField("threadNeedleStatus", "");
    setStatusField("threadCrochetStatus", "");
    setStatusField("needleTypeStatus", "");
    return;
  }
  const needle = String($("threadNeedle").value || "").trim();
  const crochet = String($("threadCrochet").value || "").trim();
  const ago = String($("needleType").value || "").trim();
  setStatusField("threadNeedleStatus", needle ? (rules.threadNeedle.includes(needle) ? "OK" : "KO") : "");
  setStatusField("threadCrochetStatus", crochet ? (rules.threadCrochet.includes(crochet) ? "OK" : "KO") : "");
  setStatusField("needleTypeStatus", ago ? (rules.needleType.includes(ago) ? "OK" : "KO") : "");
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
    tr.innerHTML = `
      <td>${a.openedAt}</td>
      <td>${a.lot}</td>
      <td>${a.type}</td>
      <td><span class="badge ${a.severity === "Grave" ? "badge-sev-high":"badge-sev-medium"}">${a.severity}</span></td>
      <td><span class="badge badge-open">${a.status}</span></td>
    `;
    body.appendChild(tr);
  });
}
function renderKpis() {
  const total = controls.length;
  const ko = controls.filter(c => c.outcome === "KO").length;
  const rate = total ? ((ko / total) * 100).toFixed(1).replace(".", ",") + "%" : "0%";
  const open = anomalies.filter(a => a.status === "Aperta").length;
  $("kpiTotalControls").textContent = total;
  $("kpiTotalKo").textContent = ko;
  $("kpiKoRate").textContent = rate;
  $("kpiOpenAnomalies").textContent = open;
}
function clearForm() {
  ["workstationNew","stitchTypeNew","lineaNew","lot","setupCodeNew","particularNew","measureBorder","measurePitch","measureNotch","notes"].forEach(id => $(id).value = "");
  ["qualityStop","qualitySkipped","qualityNeedle","needleChange","notesStatus","threadNeedle","threadCrochet","needleType"].forEach(id => $(id).selectedIndex = 0);
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

  if (control.setupCode && control.linea) {
    const codeMatch = (masterData.codici || []).some(item => String(item.codice || "").trim() === control.setupCode && String(item.linea || "").trim() === control.linea);
    if (!codeMatch) {
      validationErrors.push({ field: "setupCodeSelect", message: "Il codice prodotto non appartiene alla linea selezionata." });
    }
  }

  const materialRules = getMaterialRules(control.linea);
  if (control.linea && materialRules) {
    if (!control.threadNeedle) validationErrors.push({ field: "threadNeedle", message: "Codice filato ago obbligatorio." });
    if (!control.threadCrochet) validationErrors.push({ field: "threadCrochet", message: "Codice filato crochet obbligatorio." });
    if (!control.needleType) validationErrors.push({ field: "needleType", message: "Ago obbligatorio." });
    if (control.threadNeedle && !control.threadNeedleStatus) validationErrors.push({ field: "threadNeedle", message: "Esito codice filato ago non calcolato." });
    if (control.threadCrochet && !control.threadCrochetStatus) validationErrors.push({ field: "threadCrochet", message: "Esito codice filato crochet non calcolato." });
    if (control.needleType && !control.needleTypeStatus) validationErrors.push({ field: "needleType", message: "Esito ago non calcolato." });
    if (control.threadNeedleStatus === "KO") validationErrors.push({ field: "threadNeedle", message: "Codice filato ago non conforme per la linea selezionata." });
    if (control.threadCrochetStatus === "KO") validationErrors.push({ field: "threadCrochet", message: "Codice filato crochet non conforme per la linea selezionata." });
    if (control.needleTypeStatus === "KO") validationErrors.push({ field: "needleType", message: "Ago non conforme per la linea selezionata." });
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
  $("scanResult").textContent = "";
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
  const code = selectedBadge || normalizeCode(raw || manualNewBadge);
  const control = controls.find(x => x.id === currentAttestControlId);
  if (!control) {
    $("scanResult").textContent = "Controllo non trovato.";
    return;
  }

  let operatorName = null;
  const found = masterData.macchiniste.find(x => x.badge === code) || null;
  if (found) operatorName = found.nome;
  if (!found && BADGES[code]) operatorName = BADGES[code];
  if (!operatorName && $("manualMachineName").value.trim() && $("manualMachineBadgeNew").value.trim()) {
    operatorName = $("manualMachineName").value.trim();
  }

  if (!code || !operatorName) {
    $("scanResult").textContent = "Macchinista non riconosciuta.";
    return;
  }

  control.machineBadge = code;
  control.machineOperator = operatorName;
  control.machineAttestedAt = dateTimeString();
  control.status = "Validato";
  saveStorage();
  queueControlForSync(control);
  renderHistory();
  renderKpis();
  $("scanResult").textContent = `Attestazione acquisita: ${code}`;
  closeAttestModal();
}
async function startScanner() {
  $("scanResult").textContent = "";
  $("scanStatus").textContent = "Stato: avvio scanner...";
  if (typeof Html5Qrcode === "undefined") {
    $("scanStatus").textContent = "Stato: libreria QR non caricata";
    return;
  }
  if (scannerRunning) return;
  try {
    html5QrCode = new Html5Qrcode("reader");
    const cameras = await Html5Qrcode.getCameras();
    if (!cameras || cameras.length === 0) {
      $("scanStatus").textContent = "Stato: nessuna camera trovata";
      return;
    }
    scannerRunning = true;
    $("scanStatus").textContent = "Stato: camera attiva";
    await html5QrCode.start(
      cameras[0].id,
      { fps: 10, qrbox: { width: 220, height: 220 } },
      decodedText => confirmBadge(decodedText),
      _err => {}
    );
  } catch (err) {
    scannerRunning = false;
    $("scanStatus").textContent = "Stato: errore scanner";
    $("scanResult").textContent = String(err);
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
      anomalies = parsed.anomalies;
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
  if (!USERS[u] || USERS[u].password !== p) {
    $("loginMsg").textContent = "Credenziali non valide.";
    return;
  }
  currentUser = USERS[u];
  $("loginView").classList.add("hidden");
  $("appView").classList.remove("hidden");
  $("loggedUser").textContent = currentUser.name;
  $("todayDate").textContent = todayString();
  renderSyncStatus();
  focusIfEnabled("workstationSelect");
}
function logout() {
  currentUser = null;
  $("appView").classList.add("hidden");
  $("loginView").classList.remove("hidden");
}
window.openDetailModal = openDetailModal;
window.closeDetailModal = closeDetailModal;
window.openAttestModal = openAttestModal;
window.closeAttestModal = closeAttestModal;

window.addEventListener("DOMContentLoaded", () => {
  // Bind critical actions first
  $("btnLogin").addEventListener("click", login);
  $("btnLogout").addEventListener("click", logout);
  $("btnSave").addEventListener("click", saveControl);
  $("btnExportCsv").addEventListener("click", exportCsv);
  $("btnExportKpi").addEventListener("click", exportKpi);
  $("btnPrint").addEventListener("click", printApp);
  $("btnImportAnagrafiche").addEventListener("click", () => $("importAnagraficheFile").click());
  $("importAnagraficheFile").addEventListener("change", (e) => importAnagraficheFile(e.target.files[0]));
  $("workstationSelect").addEventListener("change", () => { updateAutoType(); syncLineFromWorkstation(); });
  $("workstationNew").addEventListener("input", () => { updateAutoType(); syncLineFromWorkstation(); });
  $("lineaSelect").addEventListener("change", () => { populateCodeSelect(); updateToleranceInfo(); });
  $("lineaNew").addEventListener("input", () => { populateCodeSelect(); updateToleranceInfo(); });
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

  try {
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
  } catch (err) {
    console.error("Init error:", err);
    const msg = $("loginMsg");
    if (msg) msg.textContent = "Inizializzazione parziale. Il login resta disponibile.";
  }
});
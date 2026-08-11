/* global JSZip */

(() => {
  "use strict";

  const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const XML_NS = "http://www.w3.org/XML/1998/namespace";
  const STATE_KEY = "docu-chantier-v1";
  const DB_NAME = "docu-chantier-local";
  const DB_STORE = "custom-templates";

  const BUILTIN_TEMPLATES = [
    {
      id: "isf",
      kind: "isf",
      title: "Instruction de sécurité ferroviaire",
      shortTitle: "ISF",
      filename: "TRAME_ISF.docx",
      source: "templates/TRAME_ISF.docx",
      description: "Instruction opérationnelle : organisation, périmètre, mesures de prévention et validation.",
      builtIn: true,
    },
    {
      id: "csf",
      kind: "csf",
      title: "Consigne de sécurité ferroviaire",
      shortTitle: "CSF",
      filename: "TRAME_CSF_2026.docx",
      source: "templates/TRAME_CSF_2026.docx",
      description: "Consigne complète avec périmètre, prévention, tableaux de travaux et annexes.",
      builtIn: true,
    },
    {
      id: "ppsps",
      kind: "ppsps",
      title: "Plan particulier de sécurité et de protection de la santé",
      shortTitle: "PPSPS",
      filename: "TRAME_PPSPS.docx",
      source: "templates/TRAME_PPSPS.docx",
      description: "Plan complet de prévention : renseignements, intervenants, analyses de risques et annexes.",
      builtIn: true,
    },
  ];

  const COMMON_FIELDS = [
    ["referenceCsf", "Référence CSF complète"],
    ["numeroIsf", "N° ISF"],
    ["annee", "Année"],
    ["referencePgc", "Référence PGC / indice"],
    ["operation", "Opération"],
    ["lieu", "Ligne / lieu"],
    ["natureTravaux", "Nature des travaux"],
    ["baseTravaux", "Base travaux"],
    ["baseArriere", "Base arrière"],
    ["periode", "Période d’intervention"],
    ["semaines", "Semaines concernées"],
    ["entreprise", "Entreprise"],
    ["lieuRdv", "Lieu de rendez-vous"],
    ["horairesJour", "Horaires de jour"],
    ["horairesNuit", "Horaires de nuit"],
    ["effectifMin", "Effectif minimum"],
    ["effectifMax", "Effectif maximum"],
    ["moeNom", "MOE Tx — nom"],
    ["moeTelephone", "MOE Tx — téléphone"],
    ["rsoNom", "RSO — nom"],
    ["rsoTelephone", "RSO — téléphone"],
    ["aspNom", "ASP — nom"],
    ["aspTelephone", "ASP — téléphone"],
    ["entrepriseNom", "Référent entreprise — nom"],
    ["entrepriseTelephone", "Référent entreprise — téléphone"],
  ];

  const state = {
    activeView: "documents",
    activeTemplateId: "isf",
    templates: [],
    dossier: {},
    drafts: {},
    scans: {},
    exportSelected: { isf: true },
    deferredInstallPrompt: null,
    pendingImport: null,
    confirmResolve: null,
    toastTimer: null,
  };

  const els = {};

  document.addEventListener("DOMContentLoaded", init);

  async function init() {
    cacheElements();
    bindEvents();
    await restoreLocalState();
    renderStaticAreas();
    setConnectivityIndicator();
    registerServiceWorker();
    ensureTemplateScanned(state.activeTemplateId).catch((error) => {
      console.error(error);
      showToast("La trame ISF n’a pas pu être analysée. Vérifie la connexion puis recharge l’application.", "error");
    });
  }

  function cacheElements() {
    [
      "templateGrid",
      "dossierForm",
      "saveDossierButton",
      "fieldTemplateSelect",
      "fieldSearch",
      "fieldFilter",
      "fieldList",
      "fieldEmpty",
      "fieldStats",
      "fieldIntro",
      "autofillButton",
      "verifyVisibleButton",
      "exportTemplateChoices",
      "exportSummary",
      "removeHighlights",
      "exportOnlyChanged",
      "generateButton",
      "offlineStatus",
      "installButton",
      "importDialog",
      "openImportButton",
      "closeImportButton",
      "templateFileInput",
      "templateNameInput",
      "importTemplateButton",
      "importMessage",
      "clearLocalDataButton",
      "confirmDialog",
      "confirmTitle",
      "confirmText",
      "confirmCancel",
      "confirmAction",
      "toast",
    ].forEach((id) => {
      els[id] = document.getElementById(id);
    });
  }

  function bindEvents() {
    document.querySelectorAll(".nav-link").forEach((button) => {
      button.addEventListener("click", () => switchView(button.dataset.view));
    });

    els.templateGrid.addEventListener("click", async (event) => {
      const cardAction = event.target.closest("[data-template-action]");
      if (!cardAction) return;
      const id = cardAction.dataset.templateId;
      const action = cardAction.dataset.templateAction;
      if (action === "open") {
        await activateTemplate(id, "champs");
      }
      if (action === "select") {
        await activateTemplate(id, "documents");
      }
      if (action === "delete") {
        await removeCustomTemplate(id);
      }
    });

    els.dossierForm.addEventListener("input", (event) => {
      const input = event.target;
      if (!input.name) return;
      state.dossier[input.name] = input.value;
      persistState();
    });

    els.saveDossierButton.addEventListener("click", () => {
      syncDossierFromForm();
      persistState();
      showToast("La fiche dossier est enregistrée sur cet appareil.", "success");
      renderExport();
    });

    els.fieldTemplateSelect.addEventListener("change", async (event) => {
      await activateTemplate(event.target.value, "champs");
    });
    els.fieldSearch.addEventListener("input", renderFields);
    els.fieldFilter.addEventListener("change", renderFields);
    els.autofillButton.addEventListener("click", prefillFromDossier);
    els.verifyVisibleButton.addEventListener("click", verifyVisibleFields);

    els.fieldList.addEventListener("input", (event) => {
      const input = event.target;
      const key = input.dataset.fieldValue;
      if (!key) return;
      const field = getActiveFields().find((item) => item.key === key);
      if (!field) return;
      const entry = getDraftEntry(state.activeTemplateId, field);
      entry.value = input.value;
      entry.origin = "manual";
      entry.verified = false;
      persistState();
      updateFieldCardState(input.closest(".field-card"), field, entry);
      renderFieldStats();
      renderExport();
    });

    els.fieldList.addEventListener("change", (event) => {
      const target = event.target;
      const key = target.dataset.fieldKey;
      if (!key) return;
      const field = getActiveFields().find((item) => item.key === key);
      if (!field) return;
      const entry = getDraftEntry(state.activeTemplateId, field);

      if (target.dataset.action === "marker") {
        entry.value = target.checked ? "X" : "";
        entry.origin = "manual";
        entry.verified = false;
        persistState();
        updateFieldCardState(target.closest(".field-card"), field, entry);
        renderFieldStats();
        renderExport();
      }

      if (target.dataset.action === "link") {
        entry.linkedKey = target.value || "";
        const commonValue = state.dossier[entry.linkedKey];
        if (entry.linkedKey && commonValue && commonValue.trim()) {
          entry.value = commonValue;
          entry.origin = "dossier";
          entry.verified = false;
        }
        persistState();
        renderFields();
        renderExport();
      }
    });

    els.fieldList.addEventListener("click", (event) => {
      const button = event.target.closest("[data-field-button]");
      if (!button) return;
      const key = button.dataset.fieldKey;
      const field = getActiveFields().find((item) => item.key === key);
      if (!field) return;
      const entry = getDraftEntry(state.activeTemplateId, field);
      if (button.dataset.fieldButton === "verify") entry.verified = !entry.verified;
      if (button.dataset.fieldButton === "restore") {
        entry.value = field.source;
        entry.origin = "manual";
        entry.verified = false;
        entry.linkedKey = "";
      }
      persistState();
      renderFields();
      renderExport();
    });

    els.exportTemplateChoices.addEventListener("change", (event) => {
      const id = event.target.dataset.exportTemplate;
      if (!id) return;
      state.exportSelected[id] = event.target.checked;
      persistState();
      renderExport();
    });

    els.generateButton.addEventListener("click", generateDocuments);

    els.openImportButton.addEventListener("click", openImportDialog);
    els.closeImportButton.addEventListener("click", closeImportDialog);
    els.templateFileInput.addEventListener("change", inspectImportFile);
    els.templateNameInput.addEventListener("input", updateImportButton);
    els.importTemplateButton.addEventListener("click", addCustomTemplate);

    els.clearLocalDataButton.addEventListener("click", clearLocalData);
    els.confirmCancel.addEventListener("click", () => settleConfirm(false));
    els.confirmAction.addEventListener("click", () => settleConfirm(true));

    window.addEventListener("online", setConnectivityIndicator);
    window.addEventListener("offline", setConnectivityIndicator);
    window.addEventListener("beforeinstallprompt", (event) => {
      event.preventDefault();
      state.deferredInstallPrompt = event;
      els.installButton.hidden = false;
    });
    els.installButton.addEventListener("click", installApplication);
  }

  async function restoreLocalState() {
    const saved = readStoredState();
    state.dossier = saved.dossier || {};
    state.drafts = saved.drafts || {};
    state.activeTemplateId = saved.activeTemplateId || "isf";
    state.activeView = saved.activeView || "documents";
    state.exportSelected = saved.exportSelected || { [state.activeTemplateId]: true };

    const customTemplates = await getCustomTemplates();
    state.templates = [...BUILTIN_TEMPLATES, ...customTemplates.map(normalizeCustomTemplate)];
    if (!state.templates.some((template) => template.id === state.activeTemplateId)) {
      state.activeTemplateId = state.templates[0]?.id || "isf";
    }
  }

  function normalizeCustomTemplate(record) {
    return {
      id: record.id,
      kind: "custom",
      title: record.title,
      shortTitle: record.shortTitle || "Trame",
      filename: record.filename,
      description: record.description || "Trame ajoutée localement depuis cet appareil.",
      fieldCount: record.fieldCount || null,
      custom: true,
      createdAt: record.createdAt,
    };
  }

  function readStoredState() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STATE_KEY) || "{}");
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch {
      return {};
    }
  }

  function persistState() {
    const data = {
      dossier: state.dossier,
      drafts: state.drafts,
      activeTemplateId: state.activeTemplateId,
      activeView: state.activeView,
      exportSelected: state.exportSelected,
    };
    try {
      localStorage.setItem(STATE_KEY, JSON.stringify(data));
    } catch (error) {
      console.error(error);
      showToast("Le navigateur n’a pas pu sauvegarder le brouillon local.", "warn");
    }
  }

  function renderStaticAreas() {
    switchView(state.activeView, true);
    renderDocumentCards();
    renderDossierForm();
    renderFieldSelectors();
    renderFields();
    renderExport();
  }

  function switchView(viewName, silent = false) {
    if (!document.getElementById(`view-${viewName}`)) return;
    state.activeView = viewName;
    document.querySelectorAll(".view").forEach((view) => {
      const active = view.dataset.viewPanel === viewName;
      view.hidden = !active;
      view.classList.toggle("active", active);
    });
    document.querySelectorAll(".nav-link").forEach((button) => {
      button.classList.toggle("active", button.dataset.view === viewName);
    });
    if (!silent) persistState();
    if (viewName === "champs") {
      ensureTemplateScanned(state.activeTemplateId).catch(handleTemplateError);
    }
    if (viewName === "export") renderExport();
    if (viewName === "dossier") renderDossierForm();
  }

  async function activateTemplate(templateId, destination) {
    if (!getTemplate(templateId)) return;
    state.activeTemplateId = templateId;
    state.exportSelected[templateId] = true;
    persistState();
    renderDocumentCards();
    renderFieldSelectors();
    renderExport();
    if (destination) switchView(destination);
    try {
      await ensureTemplateScanned(templateId);
    } catch (error) {
      handleTemplateError(error);
    }
  }

  function getTemplate(id) {
    return state.templates.find((template) => template.id === id);
  }

  function getActiveTemplate() {
    return getTemplate(state.activeTemplateId);
  }

  function renderDocumentCards() {
    els.templateGrid.innerHTML = state.templates.map((template) => {
      const scanned = state.scans[template.id];
      const count = scanned?.fields.length ?? template.fieldCount;
      const selected = template.id === state.activeTemplateId;
      const countLabel = Number.isFinite(count) ? `${count} zones repérées` : "Analyse au premier usage";
      return `
        <article class="template-card ${escapeHtml(template.kind)} ${selected ? "selected" : ""}">
          <div class="template-top">
            <span class="template-kind">${escapeHtml(template.shortTitle)}</span>
            <span class="field-count">${escapeHtml(countLabel)}</span>
          </div>
          <h2>${escapeHtml(template.title)}</h2>
          <p>${escapeHtml(template.description)}</p>
          <div class="template-meta">${escapeHtml(template.filename)}</div>
          <div class="template-actions">
            <button class="button ${selected ? "secondary" : "primary"}" type="button" data-template-action="open" data-template-id="${escapeHtml(template.id)}">${selected ? "Continuer" : "Préparer"}</button>
            ${template.custom ? `<button class="card-delete" type="button" title="Supprimer cette trame locale" aria-label="Supprimer ${escapeHtml(template.title)}" data-template-action="delete" data-template-id="${escapeHtml(template.id)}">⌫</button>` : ""}
          </div>
        </article>`;
    }).join("");
  }

  function renderDossierForm() {
    Object.entries(state.dossier).forEach(([key, value]) => {
      const input = els.dossierForm.elements.namedItem(key);
      if (input && document.activeElement !== input) input.value = value ?? "";
    });
  }

  function syncDossierFromForm() {
    const formData = new FormData(els.dossierForm);
    formData.forEach((value, key) => { state.dossier[key] = String(value); });
  }

  function renderFieldSelectors() {
    const activeId = state.activeTemplateId;
    els.fieldTemplateSelect.innerHTML = state.templates.map((template) => `<option value="${escapeHtml(template.id)}" ${template.id === activeId ? "selected" : ""}>${escapeHtml(template.shortTitle)} — ${escapeHtml(template.title)}</option>`).join("");
  }

  function getActiveFields() {
    return state.scans[state.activeTemplateId]?.fields || [];
  }

  function renderFields() {
    const template = getActiveTemplate();
    const fields = getActiveFields();
    const isReady = Boolean(state.scans[state.activeTemplateId]);
    els.autofillButton.disabled = !isReady;
    els.verifyVisibleButton.disabled = !isReady;

    if (!template) {
      els.fieldIntro.textContent = "Aucune trame n’est disponible.";
      els.fieldList.innerHTML = "";
      els.fieldEmpty.hidden = false;
      return;
    }
    if (!isReady) {
      els.fieldIntro.textContent = `Analyse en cours de la trame ${template.shortTitle}…`;
      els.fieldList.innerHTML = `<div class="empty-state"><span aria-hidden="true">◌</span><h2>Analyse de la trame en cours</h2><p>Les zones surlignées apparaîtront ici dans quelques instants.</p></div>`;
      els.fieldEmpty.hidden = true;
      els.fieldStats.innerHTML = "";
      return;
    }

    els.fieldIntro.textContent = `${template.title} · ${fields.length} zones surlignées détectées dans le document et les pieds de page.`;
    const visible = getFilteredFields(fields);
    renderFieldStats(fields, visible.length);
    els.fieldList.innerHTML = visible.map((field, index) => renderFieldCard(field, index + 1)).join("");
    els.fieldEmpty.hidden = visible.length !== 0;
  }

  function getFilteredFields(fields) {
    const search = normalizeForSearch(els.fieldSearch.value);
    const filter = els.fieldFilter.value;
    return fields.filter((field) => {
      const entry = getDraftEntry(state.activeTemplateId, field);
      const isChanged = hasChanged(field, entry);
      const linked = Boolean(entry.linkedKey);
      const matchesSearch = !search || normalizeForSearch(`${field.source} ${field.context} ${field.title}`).includes(search);
      if (!matchesSearch) return false;
      if (filter === "all") return true;
      if (filter === "review") return !entry.verified;
      if (filter === "changed") return isChanged;
      if (filter === "markers") return field.type === "marker";
      if (filter === "linked") return linked;
      if (filter === "footer") return /footer/i.test(field.part);
      return isPriorityField(field, entry);
    });
  }

  function isPriorityField(field, entry) {
    if (entry.linkedKey || field.suggestedKey) return true;
    if (/\b(xxx|a remplir|à remplir|mettre|remplacer|compl[eé]ter|neant|néant)\b/i.test(`${field.source} ${field.context}`)) return true;
    if (field.type === "marker" && /travaux|phase|voies|faisceau/i.test(field.context)) return true;
    return field.order < 65 && field.type !== "marker";
  }

  function renderFieldStats(fields = getActiveFields(), visibleCount = null) {
    const total = fields.length;
    const entries = fields.map((field) => getDraftEntry(state.activeTemplateId, field));
    const changed = fields.filter((field, index) => hasChanged(field, entries[index])).length;
    const verified = entries.filter((entry) => entry.verified).length;
    const pending = total - verified;
    const markers = fields.filter((field) => field.type === "marker").length;
    const shown = visibleCount ?? getFilteredFields(fields).length;
    els.fieldStats.innerHTML = `
      <span class="stat-pill"><strong>${shown}</strong> affichées</span>
      <span class="stat-pill"><strong>${total}</strong> zones</span>
      <span class="stat-pill ${changed ? "good" : ""}"><strong>${changed}</strong> modifiées</span>
      <span class="stat-pill ${pending ? "warn" : "good"}"><strong>${pending}</strong> à vérifier</span>
      <span class="stat-pill"><strong>${markers}</strong> cases « X »</span>`;
  }

  function renderFieldCard(field, index) {
    const entry = getDraftEntry(state.activeTemplateId, field);
    const changed = hasChanged(field, entry);
    const classes = ["field-card", field.type, changed ? "changed" : "", entry.verified ? "verified" : ""].filter(Boolean).join(" ");
    const linkOptions = [`<option value="">Saisie manuelle</option>`]
      .concat(COMMON_FIELDS.map(([key, label]) => `<option value="${key}" ${entry.linkedKey === key ? "selected" : ""}>${escapeHtml(label)}${field.suggestedKey === key ? " · suggéré" : ""}</option>`))
      .join("");
    const isLong = field.source.length > 54 || /observation|commentaire|mesures|disposition/i.test(`${field.title} ${field.context}`);
    const valueInput = field.type === "marker"
      ? `<label class="marker-control"><input type="checkbox" data-field-key="${escapeHtml(field.key)}" data-action="marker" ${entry.value.trim() ? "checked" : ""} /><span><strong>${entry.value.trim() ? "Case cochée" : "Case non cochée"}</strong><small>Décoche pour supprimer le « X ».</small></span></label>`
      : `<label>${isLong ? "Texte à insérer" : "Valeur à insérer"}${isLong
        ? `<textarea class="field-input" data-field-value="${escapeHtml(field.key)}" rows="3">${escapeHtml(entry.value)}</textarea>`
        : `<input class="field-input" data-field-value="${escapeHtml(field.key)}" value="${escapeHtml(entry.value)}" />`}</label>`;
    const verifyLabel = entry.verified ? "Annuler la validation" : "Valider";
    return `
      <article class="${classes}" data-field-card="${escapeHtml(field.key)}">
        <div class="field-description">
          <div class="field-position"><span class="field-index">${index}</span><span>${escapeHtml(formatPart(field.part))} · zone ${field.paragraph + 1}</span></div>
          <strong class="field-title">${escapeHtml(field.title)}</strong>
          <p class="field-context">${escapeContext(field.context, field.source)}</p>
        </div>
        <div class="field-edit">
          ${valueInput}
          <div class="field-actions">
            <button type="button" class="field-action ${entry.verified ? "unverify" : "verify"}" data-field-button="verify" data-field-key="${escapeHtml(field.key)}">${verifyLabel}</button>
            ${changed ? `<button type="button" class="field-action" data-field-button="restore" data-field-key="${escapeHtml(field.key)}">Rétablir la trame</button>` : ""}
          </div>
        </div>
        <div class="field-link">
          <label>Lier à la fiche dossier<select class="field-link-select" data-field-key="${escapeHtml(field.key)}" data-action="link">${linkOptions}</select></label>
        </div>
      </article>`;
  }

  function updateFieldCardState(card, field, entry) {
    if (!card) return;
    card.classList.toggle("changed", hasChanged(field, entry));
    card.classList.toggle("verified", Boolean(entry.verified));
  }

  async function prefillFromDossier() {
    syncDossierFromForm();
    const fields = getActiveFields();
    let changed = 0;
    fields.forEach((field) => {
      const entry = getDraftEntry(state.activeTemplateId, field);
      const suggestedKey = entry.linkedKey || field.suggestedKey;
      const commonValue = suggestedKey ? state.dossier[suggestedKey] : "";
      if (!suggestedKey || !commonValue || !commonValue.trim() || !canAutofill(field, suggestedKey)) return;
      if (entry.origin === "manual" && hasChanged(field, entry)) return;
      if (entry.value !== commonValue || entry.linkedKey !== suggestedKey) {
        entry.value = commonValue;
        entry.linkedKey = suggestedKey;
        entry.origin = "dossier";
        entry.verified = false;
        changed += 1;
      }
    });
    persistState();
    renderFields();
    renderExport();
    showToast(changed ? `${changed} zone${changed > 1 ? "s" : ""} a été préremplie depuis la fiche dossier.` : "Aucune zone sûre à préremplir : complète la fiche ou lie une zone manuellement.", changed ? "success" : "warn");
  }

  function canAutofill(field, key) {
    const source = normalizeText(field.source);
    if (!source || field.type === "marker") return false;
    if (["referenceCsf", "numeroIsf", "annee"].includes(key) && /[/\d]/.test(source)) return false;
    if (["effectifMin", "effectifMax"].includes(key) && !/^\d{1,3}$/.test(source)) return false;
    if (key === "periode" && source.length < 6) return false;
    return true;
  }

  function verifyVisibleFields() {
    const visible = getFilteredFields(getActiveFields());
    if (!visible.length) return;
    visible.forEach((field) => { getDraftEntry(state.activeTemplateId, field).verified = true; });
    persistState();
    renderFields();
    renderExport();
    showToast(`${visible.length} zone${visible.length > 1 ? "s" : ""} visible${visible.length > 1 ? "s" : ""} validée${visible.length > 1 ? "s" : ""}.`, "success");
  }

  function getDraftEntry(templateId, field) {
    if (!state.drafts[templateId]) state.drafts[templateId] = {};
    if (!state.drafts[templateId][field.key]) {
      state.drafts[templateId][field.key] = {
        value: field.source,
        verified: false,
        linkedKey: "",
        origin: "template",
      };
    }
    return state.drafts[templateId][field.key];
  }

  function hasChanged(field, entry) {
    return entry.value !== field.source;
  }

  function renderExport() {
    const templates = state.templates;
    templates.forEach((template) => {
      if (!(template.id in state.exportSelected)) state.exportSelected[template.id] = template.id === state.activeTemplateId;
    });
    els.exportTemplateChoices.innerHTML = templates.map((template) => {
      const fields = state.scans[template.id]?.fields;
      const changed = fields ? fields.filter((field) => hasChanged(field, getDraftEntry(template.id, field))).length : null;
      const details = fields ? `${fields.length} zones · ${changed} modification${changed !== 1 ? "s" : ""}` : "Analyse au moment de la génération";
      return `<label class="choice-card"><input type="checkbox" data-export-template="${escapeHtml(template.id)}" ${state.exportSelected[template.id] ? "checked" : ""} /><span><strong>${escapeHtml(template.shortTitle)} — ${escapeHtml(template.title)}</strong><small>${escapeHtml(details)}</small></span></label>`;
    }).join("");

    const selectedTemplates = getSelectedTemplates();
    const summaries = selectedTemplates.map((template) => createExportSummary(template));
    els.exportSummary.innerHTML = summaries.length
      ? summaries.join("")
      : `<div class="summary-row warn"><span>Aucun document choisi</span><strong>Sélectionne au moins une trame.</strong></div>`;
    els.generateButton.disabled = selectedTemplates.length === 0;
    els.generateButton.textContent = selectedTemplates.length > 1 ? "Générer les documents Word" : "Générer et télécharger le Word";
  }

  function createExportSummary(template) {
    const fields = state.scans[template.id]?.fields;
    if (!fields) {
      return `<div class="summary-row"><span>${escapeHtml(template.shortTitle)}</span><strong>À analyser</strong></div>`;
    }
    const changed = fields.filter((field) => hasChanged(field, getDraftEntry(template.id, field))).length;
    const verified = fields.filter((field) => getDraftEntry(template.id, field).verified).length;
    const pending = fields.length - verified;
    return `
      <div class="summary-row"><span>${escapeHtml(template.shortTitle)} · zones détectées</span><strong>${fields.length}</strong></div>
      <div class="summary-row good"><span>${escapeHtml(template.shortTitle)} · modifiées</span><strong>${changed}</strong></div>
      <div class="summary-row ${pending ? "warn" : "good"}"><span>${escapeHtml(template.shortTitle)} · à vérifier</span><strong>${pending}</strong></div>`;
  }

  function getSelectedTemplates() {
    return state.templates.filter((template) => state.exportSelected[template.id]);
  }

  async function generateDocuments() {
    const selected = getSelectedTemplates();
    if (!selected.length) return;
    syncDossierFromForm();

    for (const template of selected) {
      try {
        await ensureTemplateScanned(template.id);
      } catch (error) {
        handleTemplateError(error);
        return;
      }
    }

    const onlyChanged = els.exportOnlyChanged.checked;
    const candidates = selected.filter((template) => {
      const fields = state.scans[template.id].fields;
      const changes = fields.some((field) => hasChanged(field, getDraftEntry(template.id, field)));
      return !onlyChanged || changes;
    });
    if (!candidates.length) {
      showToast("Aucun document n’a été généré : aucune zone n’est modifiée avec ce filtre.", "warn");
      return;
    }

    const pending = candidates.reduce((total, template) => {
      const fields = state.scans[template.id].fields;
      return total + fields.filter((field) => !getDraftEntry(template.id, field).verified).length;
    }, 0);
    if (pending) {
      const proceed = await confirmAction(
        "Zones encore à vérifier",
        `Il reste ${pending} zone${pending > 1 ? "s" : ""} non validée${pending > 1 ? "s" : ""}. Tu peux tout de même générer le document, mais vérifie bien les données avant diffusion.`,
        "Générer quand même",
      );
      if (!proceed) return;
    }

    els.generateButton.disabled = true;
    const originalButtonText = els.generateButton.textContent;
    try {
      const outputs = [];
      for (let index = 0; index < candidates.length; index += 1) {
        const template = candidates[index];
        els.generateButton.textContent = `Génération ${index + 1}/${candidates.length}…`;
        const blob = await buildDocument(template, els.removeHighlights.checked);
        outputs.push({ blob, filename: createOutputName(template) });
      }
      if (outputs.length === 1) {
        downloadBlob(outputs[0].blob, outputs[0].filename);
        showToast("Le document Word est généré et téléchargé.", "success");
      } else {
        els.generateButton.textContent = "Création de l’archive…";
        const archive = new JSZip();
        outputs.forEach(({ blob, filename }) => archive.file(filename, blob));
        const archiveBlob = await archive.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
        downloadBlob(archiveBlob, `Docu_Chantier_${new Date().toISOString().slice(0, 10)}.zip`);
        showToast(`${outputs.length} documents Word sont générés dans une archive ZIP téléchargée.`, "success");
      }
    } catch (error) {
      console.error(error);
      showToast("La génération a échoué. Le brouillon est conservé ; essaie à nouveau.", "error");
    } finally {
      els.generateButton.disabled = false;
      els.generateButton.textContent = originalButtonText;
      persistState();
    }
  }

  async function buildDocument(template, removeHighlights) {
    const scan = await ensureTemplateScanned(template.id);
    const zip = await JSZip.loadAsync(scan.buffer);
    const fieldsByPart = groupBy(scan.fields, (field) => field.part);

    for (const [part, scannedFields] of fieldsByPart.entries()) {
      const file = zip.file(part);
      if (!file) continue;
      const xml = await file.async("text");
      const doc = parseXml(xml, part);
      const generatedFields = collectEditableFields(doc, part);
      const byKey = new Map(generatedFields.map((field) => [field.key, field]));

      scannedFields.forEach((scannedField) => {
        const field = byKey.get(scannedField.key);
        if (!field) return;
        const entry = getDraftEntry(template.id, scannedField);
        if (hasChanged(scannedField, entry)) setGroupText(field, entry.value);
        if (removeHighlights) field.runs.forEach((run) => clearRunHighlight(run));
      });
      zip.file(part, new XMLSerializer().serializeToString(doc));
    }
    return zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 45_000);
  }

  function createOutputName(template) {
    const project = sanitizeFilename(state.dossier.operation || state.dossier.lieu || template.shortTitle);
    const reference = sanitizeFilename(state.dossier.referenceCsf || state.dossier.numeroIsf || "brouillon");
    const stamp = new Date().toISOString().slice(0, 10);
    return `${template.shortTitle}_${project}_${reference}_${stamp}.docx`;
  }

  async function ensureTemplateScanned(templateId) {
    if (state.scans[templateId]) return state.scans[templateId];
    const template = getTemplate(templateId);
    if (!template) throw new Error("Trame introuvable");
    const buffer = await getTemplateBuffer(template);
    const fields = await scanDocx(buffer);
    state.scans[templateId] = { buffer, fields };
    template.fieldCount = fields.length;
    renderDocumentCards();
    renderFieldSelectors();
    if (templateId === state.activeTemplateId) renderFields();
    renderExport();
    return state.scans[templateId];
  }

  async function getTemplateBuffer(template) {
    if (template.builtIn) {
      const response = await fetch(template.source, { cache: "no-cache" });
      if (!response.ok) throw new Error(`Impossible de charger ${template.filename}`);
      return response.arrayBuffer();
    }
    const record = await getCustomTemplate(template.id);
    if (!record?.data) throw new Error("Cette trame locale est introuvable sur cet appareil.");
    if (record.data instanceof Blob) return record.data.arrayBuffer();
    return record.data;
  }

  async function scanDocx(buffer) {
    if (!window.JSZip) throw new Error("Le moteur de génération Word n’est pas disponible.");
    const zip = await JSZip.loadAsync(buffer);
    const xmlParts = Object.keys(zip.files)
      .filter((name) => name.startsWith("word/") && name.endsWith(".xml"))
      .filter((name) => /(?:document|header|footer|footnotes|endnotes|comments)/i.test(name));
    const allFields = [];
    for (const part of xmlParts) {
      const file = zip.file(part);
      if (!file) continue;
      const doc = parseXml(await file.async("text"), part);
      allFields.push(...collectEditableFields(doc, part));
    }
    return allFields.map((field, order) => ({ ...field, order }));
  }

  function parseXml(xml, partName) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error(`Le contenu Word ${partName} est illisible.`);
    return doc;
  }

  function collectEditableFields(doc, part) {
    const paragraphs = Array.from(doc.getElementsByTagNameNS(WORD_NS, "p"));
    const previousText = paragraphs.map((paragraph) => normalizeText(getParagraphText(paragraph)));
    const results = [];
    paragraphs.forEach((paragraph, paragraphIndex) => {
      const runs = Array.from(paragraph.getElementsByTagNameNS(WORD_NS, "r"));
      let currentGroup = null;
      let groupIndex = 0;
      let previousRunIndex = -2;
      const flush = () => {
        if (!currentGroup || !currentGroup.runs.length) return;
        const source = currentGroup.runs.map((run) => getRunText(run)).join("");
        if (source) {
          const context = buildContext(previousText, paragraphIndex, source);
          const field = {
            key: `${part}::${paragraphIndex}::${groupIndex}`,
            part,
            paragraph: paragraphIndex,
            group: groupIndex,
            source,
            context,
            title: inferFieldTitle(previousText, paragraphIndex, source),
            type: normalizeText(source).toLowerCase() === "x" ? "marker" : "text",
            suggestedKey: inferSuggestedCommonKey(previousText, paragraphIndex, source),
            runs: currentGroup.runs,
          };
          results.push(field);
          groupIndex += 1;
        }
        currentGroup = null;
      };

      runs.forEach((run, runIndex) => {
        const editable = isEditableRun(run);
        const hasText = Boolean(getRunText(run));
        if (editable && hasText) {
          if (!currentGroup || runIndex !== previousRunIndex + 1) {
            flush();
            currentGroup = { runs: [] };
          }
          currentGroup.runs.push(run);
          previousRunIndex = runIndex;
        } else {
          flush();
          previousRunIndex = runIndex;
        }
      });
      flush();
    });
    return results;
  }

  function isEditableRun(run) {
    const properties = directChild(run, "rPr");
    if (!properties) return false;
    const highlight = properties.getElementsByTagNameNS(WORD_NS, "highlight")[0];
    if (highlight) return true;
    const shading = properties.getElementsByTagNameNS(WORD_NS, "shd")[0];
    if (!shading) return false;
    const fill = (shading.getAttributeNS(WORD_NS, "fill") || shading.getAttribute("w:fill") || shading.getAttribute("fill") || "").toLowerCase();
    return fill && !["auto", "ffffff", "fff"].includes(fill);
  }

  function directChild(element, localName) {
    return Array.from(element.childNodes).find((node) => node.nodeType === Node.ELEMENT_NODE && node.namespaceURI === WORD_NS && node.localName === localName) || null;
  }

  function getRunText(run) {
    return Array.from(run.getElementsByTagNameNS(WORD_NS, "t")).map((node) => node.textContent || "").join("");
  }

  function getParagraphText(paragraph) {
    return Array.from(paragraph.getElementsByTagNameNS(WORD_NS, "t")).map((node) => node.textContent || "").join("");
  }

  function buildContext(paragraphTexts, index, source) {
    const own = paragraphTexts[index] || "";
    const previous = paragraphTexts.slice(Math.max(0, index - 4), index).filter(Boolean).slice(-3);
    const combined = [own, ...previous.reverse()].filter(Boolean).join(" · ");
    return cropAround(combined || source, source, 210);
  }

  function inferFieldTitle(paragraphTexts, index, source) {
    const own = paragraphTexts[index] || "";
    const sourceNormalized = normalizeText(source);
    const remaining = normalizeText(own.replace(sourceNormalized, ""));
    if (isUsefulLabel(remaining)) return remaining;
    const nearby = paragraphTexts.slice(Math.max(0, index - 5), index).reverse();
    const label = nearby.find((text) => isUsefulLabel(text));
    if (label) return cropText(label, 88);
    return sourceNormalized.length <= 84 ? sourceNormalized : "Zone surlignée";
  }

  function isUsefulLabel(text) {
    const value = normalizeText(text);
    if (value.length < 3 || value.length > 100) return false;
    if (/^(x|oui|non|observations?|signature)$/i.test(value)) return false;
    return true;
  }

  function inferSuggestedCommonKey(paragraphTexts, index, source) {
    const nearby = [paragraphTexts[index], ...paragraphTexts.slice(Math.max(0, index - 6), index)].join(" ");
    const text = normalizeForSearch(nearby);
    const value = normalizeText(source);
    if (/lieu de (rdv|rendez)/.test(text)) return "lieuRdv";
    if (/operation/.test(text)) return "operation";
    if (/nature des travaux/.test(text)) return "natureTravaux";
    if (/base travaux/.test(text)) return "baseTravaux";
    if (/base arriere/.test(text)) return "baseArriere";
    if (/duree d.?intervention|periode previsionnelle|a partir du/.test(text)) return "periode";
    if (/semaines? concerne/.test(text)) return "semaines";
    if (/horaires? de travail.*jour|de jour/.test(text) && /horaires?/.test(text)) return "horairesJour";
    if (/horaires? de travail.*nuit|de nuit/.test(text) && /horaires?/.test(text)) return "horairesNuit";
    if (/effectif.*min/.test(text) || (/min/.test(text) && /^\d{1,3}$/.test(value))) return "effectifMin";
    if (/effectif.*max/.test(text) || (/max/.test(text) && /^\d{1,3}$/.test(value))) return "effectifMax";
    if (/entreprise\s*:|l.?entreprise electrique/.test(text) && value.length > 3) return "entreprise";
    if (/\blieu\b|ligne\s+\d{3}/.test(text) && value.length > 4) return "lieu";
    if (/moe.*contact|moe tx/.test(text) && /tel|telephone/.test(text)) return "moeTelephone";
    if (/moe.*contact|moe tx/.test(text)) return "moeNom";
    if (/\brso\b/.test(text) && /tel|telephone/.test(text)) return "rsoTelephone";
    if (/\brso\b/.test(text)) return "rsoNom";
    if (/\basp\b/.test(text) && /tel|telephone/.test(text)) return "aspTelephone";
    if (/\basp\b/.test(text)) return "aspNom";
    if (/representant chantier entreprise/.test(text) && /tel|telephone/.test(text)) return "entrepriseTelephone";
    if (/representant chantier entreprise/.test(text)) return "entrepriseNom";
    return "";
  }

  function setGroupText(field, value) {
    field.runs.forEach((run, index) => setRunText(run, index === 0 ? value : ""));
  }

  function setRunText(run, value) {
    Array.from(run.childNodes)
      .filter((node) => node.nodeType === Node.ELEMENT_NODE && node.namespaceURI === WORD_NS && ["t", "delText", "instrText"].includes(node.localName))
      .forEach((node) => node.remove());
    const doc = run.ownerDocument;
    const properties = directChild(run, "rPr");
    let cursor = properties ? properties.nextSibling : run.firstChild;
    const add = (node) => {
      if (cursor) run.insertBefore(node, cursor);
      else run.appendChild(node);
    };
    const lines = String(value ?? "").replace(/\r\n/g, "\n").split("\n");
    lines.forEach((line, index) => {
      if (index) add(doc.createElementNS(WORD_NS, "w:br"));
      if (!line && lines.length === 1) return;
      const text = doc.createElementNS(WORD_NS, "w:t");
      if (/^\s|\s$/.test(line)) text.setAttributeNS(XML_NS, "xml:space", "preserve");
      text.textContent = line;
      add(text);
    });
  }

  function clearRunHighlight(run) {
    const properties = directChild(run, "rPr");
    if (!properties) return;
    Array.from(properties.childNodes)
      .filter((node) => node.nodeType === Node.ELEMENT_NODE && node.namespaceURI === WORD_NS && node.localName === "highlight")
      .forEach((node) => node.remove());
  }

  function openImportDialog() {
    state.pendingImport = null;
    els.templateFileInput.value = "";
    els.templateNameInput.value = "";
    els.importMessage.textContent = "";
    els.importMessage.className = "form-message";
    els.importTemplateButton.disabled = true;
    if (typeof els.importDialog.showModal === "function") els.importDialog.showModal();
  }

  function closeImportDialog() {
    if (els.importDialog.open) els.importDialog.close();
  }

  async function inspectImportFile() {
    const file = els.templateFileInput.files?.[0];
    state.pendingImport = null;
    if (!file) {
      updateImportButton();
      return;
    }
    if (!/\.docx$/i.test(file.name)) {
      setImportMessage("Choisis un fichier Word au format .docx.");
      updateImportButton();
      return;
    }
    setImportMessage("Analyse locale de la trame…", true);
    try {
      const buffer = await file.arrayBuffer();
      const fields = await scanDocx(buffer);
      if (!fields.length) throw new Error("Aucune zone surlignée n’a été détectée.");
      state.pendingImport = { file, buffer, fields };
      if (!els.templateNameInput.value) els.templateNameInput.value = file.name.replace(/\.docx$/i, "").replace(/[_-]+/g, " ").trim();
      setImportMessage(`${fields.length} zones surlignées détectées. Donne un nom à cette trame puis ajoute-la.`, true);
    } catch (error) {
      console.error(error);
      setImportMessage(error.message || "Cette trame Word ne peut pas être analysée.");
    }
    updateImportButton();
  }

  function updateImportButton() {
    els.importTemplateButton.disabled = !state.pendingImport || !els.templateNameInput.value.trim();
  }

  function setImportMessage(message, success = false) {
    els.importMessage.textContent = message;
    els.importMessage.className = `form-message${success ? " success" : ""}`;
  }

  async function addCustomTemplate() {
    const pending = state.pendingImport;
    const title = els.templateNameInput.value.trim();
    if (!pending || !title) return;
    const id = `custom-${createId()}`;
    const record = {
      id,
      title,
      shortTitle: "Trame",
      filename: pending.file.name,
      description: "Trame ajoutée localement depuis cet appareil.",
      fieldCount: pending.fields.length,
      data: pending.buffer,
      createdAt: Date.now(),
    };
    try {
      await putCustomTemplate(record);
      const template = normalizeCustomTemplate(record);
      state.templates.push(template);
      state.scans[id] = { buffer: pending.buffer, fields: pending.fields };
      state.drafts[id] = {};
      state.activeTemplateId = id;
      state.exportSelected[id] = true;
      persistState();
      closeImportDialog();
      renderStaticAreas();
      switchView("champs");
      showToast("La nouvelle trame est ajoutée et prête à être complétée.", "success");
    } catch (error) {
      console.error(error);
      setImportMessage("La trame n’a pas pu être enregistrée localement.");
    }
  }

  async function removeCustomTemplate(templateId) {
    const template = getTemplate(templateId);
    if (!template?.custom) return;
    const confirmed = await confirmAction("Supprimer la trame locale", `Supprimer « ${template.title} » de cet appareil ? Les documents Word d’origine ne sont pas touchés, mais son brouillon local sera supprimé.`, "Supprimer");
    if (!confirmed) return;
    try {
      await deleteCustomTemplate(templateId);
      state.templates = state.templates.filter((item) => item.id !== templateId);
      delete state.scans[templateId];
      delete state.drafts[templateId];
      delete state.exportSelected[templateId];
      if (state.activeTemplateId === templateId) state.activeTemplateId = state.templates[0]?.id || "isf";
      persistState();
      renderStaticAreas();
      showToast("La trame locale et son brouillon sont supprimés de cet appareil.", "success");
    } catch (error) {
      console.error(error);
      showToast("La trame n’a pas pu être supprimée.", "error");
    }
  }

  async function clearLocalData() {
    const confirmed = await confirmAction("Effacer les données locales", "Tous les brouillons, la fiche dossier et les trames ajoutées seront effacés de cet appareil. Les deux trames fournies par l’application resteront disponibles.", "Tout effacer");
    if (!confirmed) return;
    try {
      localStorage.removeItem(STATE_KEY);
      await clearCustomTemplates();
      location.reload();
    } catch (error) {
      console.error(error);
      showToast("Les données locales n’ont pas pu être entièrement effacées.", "error");
    }
  }

  function confirmAction(title, text, actionLabel) {
    els.confirmTitle.textContent = title;
    els.confirmText.textContent = text;
    els.confirmAction.textContent = actionLabel;
    els.confirmDialog.hidden = false;
    return new Promise((resolve) => { state.confirmResolve = resolve; });
  }

  function settleConfirm(value) {
    els.confirmDialog.hidden = true;
    const resolve = state.confirmResolve;
    state.confirmResolve = null;
    if (resolve) resolve(value);
  }

  function setConnectivityIndicator() {
    const online = navigator.onLine;
    els.offlineStatus.textContent = online ? "Local et hors connexion" : "Mode hors connexion";
    els.offlineStatus.classList.toggle("offline", !online);
  }

  function registerServiceWorker() {
    if (!("serviceWorker" in navigator)) return;
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("./sw.js").catch((error) => console.warn("Service worker indisponible", error));
    });
  }

  async function installApplication() {
    if (!state.deferredInstallPrompt) return;
    state.deferredInstallPrompt.prompt();
    await state.deferredInstallPrompt.userChoice;
    state.deferredInstallPrompt = null;
    els.installButton.hidden = true;
  }

  function handleTemplateError(error) {
    console.error(error);
    showToast(error?.message || "La trame ne peut pas être analysée.", "error");
  }

  function showToast(message, type = "") {
    window.clearTimeout(state.toastTimer);
    els.toast.textContent = message;
    els.toast.className = `toast ${type}`;
    els.toast.hidden = false;
    state.toastTimer = window.setTimeout(() => { els.toast.hidden = true; }, 5200);
  }

  function formatPart(part) {
    if (/footer/i.test(part)) return "Pied de page";
    if (/header/i.test(part)) return "En-tête";
    if (/footnotes|endnotes/i.test(part)) return "Note";
    return "Document";
  }

  function groupBy(items, keyFunction) {
    return items.reduce((map, item) => {
      const key = keyFunction(item);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(item);
      return map;
    }, new Map());
  }

  function normalizeText(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim();
  }

  function normalizeForSearch(value) {
    return normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  function cropText(value, length) {
    const text = normalizeText(value);
    return text.length > length ? `${text.slice(0, length - 1).trimEnd()}…` : text;
  }

  function cropAround(value, needle, length) {
    const text = normalizeText(value);
    const target = normalizeText(needle);
    if (text.length <= length) return text;
    const index = text.toLowerCase().indexOf(target.toLowerCase());
    if (index < 0) return cropText(text, length);
    const start = Math.max(0, index - Math.floor((length - target.length) / 2));
    const end = Math.min(text.length, start + length);
    return `${start ? "…" : ""}${text.slice(start, end).trim()}${end < text.length ? "…" : ""}`;
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", "\"": "&quot;" }[char]));
  }

  function escapeContext(context, source) {
    const safeContext = escapeHtml(context);
    const safeSource = escapeHtml(normalizeText(source));
    if (!safeSource) return safeContext;
    const pattern = new RegExp(escapeRegExp(safeSource), "i");
    return safeContext.replace(pattern, (match) => `<mark>${match}</mark>`);
  }

  function escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function sanitizeFilename(value) {
    const cleaned = normalizeText(value).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
    return (cleaned || "document").slice(0, 54);
  }

  function createId() {
    if (window.crypto?.randomUUID) return window.crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function openDatabase() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE, { keyPath: "id" });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function getCustomTemplates() {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const request = db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).getAll();
      request.onsuccess = () => { db.close(); resolve(request.result || []); };
      request.onerror = () => { db.close(); reject(request.error); };
    });
  }

  async function getCustomTemplate(id) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const request = db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).get(id);
      request.onsuccess = () => { db.close(); resolve(request.result); };
      request.onerror = () => { db.close(); reject(request.error); };
    });
  }

  async function putCustomTemplate(record) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const request = db.transaction(DB_STORE, "readwrite").objectStore(DB_STORE).put(record);
      request.onsuccess = () => { db.close(); resolve(); };
      request.onerror = () => { db.close(); reject(request.error); };
    });
  }

  async function deleteCustomTemplate(id) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const request = db.transaction(DB_STORE, "readwrite").objectStore(DB_STORE).delete(id);
      request.onsuccess = () => { db.close(); resolve(); };
      request.onerror = () => { db.close(); reject(request.error); };
    });
  }

  async function clearCustomTemplates() {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const request = db.transaction(DB_STORE, "readwrite").objectStore(DB_STORE).clear();
      request.onsuccess = () => { db.close(); resolve(); };
      request.onerror = () => { db.close(); reject(request.error); };
    });
  }
})();

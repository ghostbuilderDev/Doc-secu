/* global ISFStorage, DocxEngine */

(() => {
  "use strict";

  const APP_VERSION = "ISF Express V1";
  const LEGACY_KEY = "docu-chantier-v1";
  const STEPS = [
    { id: "chantier", label: "Chantier" },
    { id: "intervention", label: "Intervention" },
    { id: "phases", label: "Phases et protections" },
    { id: "review", label: "Vérification" },
    { id: "generate", label: "Génération" },
  ];
  const PROTECTIONS = [
    ["consignation", "Consignation caténaire"],
    ["interception", "Interception circulation"],
    ["annonce", "Annonce"],
    ["cloture", "Clôture limitative"],
    ["barriere", "Barrière défensive"],
  ];
  const QUESTION_LABELS = [
    "Le RSO (passant gris) reste en permanence sur le chantier",
    "En l’absence du RSO, l’entreprise est autorisée à travailler dans les emprises",
    "Le responsable chantier de l’entreprise, repérable par un passant vert, reste sur place pendant toute la session de travail",
    "Le personnel entreprise reste sur place jusqu’à l’autorisation de départ du RSO / ASP",
  ];
  const LIBRARY_TYPES = [
    ["contact", "Intervenants"],
    ["activity", "Activités"],
    ["activitySet", "Ensembles"],
    ["observation", "Observations"],
    ["company", "Entreprises"],
    ["line", "Lignes"],
    ["csf", "CSF"],
    ["place", "Lieux"],
    ["phaseTemplate", "Phases"],
  ];

  const state = {
    view: "dashboard",
    editorStep: 1,
    isfs: [],
    chantiers: [],
    library: [],
    settings: {},
    currentIsf: null,
    editingChantier: null,
    libraryFilter: "contact",
    isfFilter: "all",
    filters: { isfSearch: "", chantierSearch: "", librarySearch: "" },
    autosaveTimer: null,
    toastTimer: null,
    confirmResolve: null,
    contactTarget: null,
    libraryTarget: null,
    pendingWorker: null,
  };

  const els = {};

  document.addEventListener("DOMContentLoaded", init);

  async function init() {
    cacheElements();
    bindEvents();
    await ISFStorage.open();
    await loadData();
    await seedLibrary();
    await migrateLegacyData();
    sortCollections();
    render();
    setConnectionStatus();
    registerServiceWorker();
    exposeTestHooks();
  }

  function cacheElements() {
    [
      "appMain", "mainNav", "saveStatus", "connectionStatus", "updateBanner", "updateAppButton",
      "newIsfDialog", "newIsfPicker", "contactDialog", "contactForm", "contactDialogTitle",
      "libraryDialog", "libraryForm", "libraryDialogTitle", "libraryValueField",
      "confirmDialog", "confirmTitle", "confirmText", "confirmButton", "backupFileInput", "toast",
    ].forEach((id) => { els[id] = document.getElementById(id); });
  }

  function bindEvents() {
    document.addEventListener("click", handleClick);
    document.addEventListener("keydown", (event) => {
      const actionable = event.target.closest?.('[role="button"][data-action]');
      if (!actionable || !["Enter", " "].includes(event.key)) return;
      event.preventDefault();
      actionable.click();
    });
    els.appMain.addEventListener("input", handleMainInput);
    els.appMain.addEventListener("change", handleMainChange);
    els.contactForm.addEventListener("submit", handleContactSubmit);
    els.libraryForm.addEventListener("submit", handleLibrarySubmit);
    els.backupFileInput.addEventListener("change", importBackupFile);
    els.confirmDialog.addEventListener("close", () => {
      const resolve = state.confirmResolve;
      state.confirmResolve = null;
      if (resolve) resolve(els.confirmDialog.returnValue === "confirm");
    });
    window.addEventListener("online", setConnectionStatus);
    window.addEventListener("offline", setConnectionStatus);
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden") flushCurrentSave();
    });
    window.addEventListener("beforeunload", () => flushCurrentSave());
  }

  async function loadData() {
    const [isfs, chantiers, library, settings] = await Promise.all([
      ISFStorage.getAll("isfs"),
      ISFStorage.getAll("chantiers"),
      ISFStorage.getAll("library"),
      ISFStorage.getAll("settings"),
    ]);
    state.isfs = isfs;
    state.chantiers = chantiers;
    state.library = library;
    state.settings = Object.fromEntries(settings.map((item) => [item.id, item.value]));
  }

  async function seedLibrary() {
    const seeds = [
      { id: "seed-contact-figueras", type: "contact", lastName: "FIGUERAS", firstName: "Laurent", phone: "07 77 96 08 58", company: "SNCF", function: "RLT SES", role: "moe", source: "TRAME ISF.docx" },
      { id: "seed-contact-haider", type: "contact", lastName: "HAIDER", firstName: "David", phone: "06 51 37 31 01", company: "SNCF", function: "RSO / ASP", role: "rso", source: "TRAME ISF.docx" },
      { id: "seed-contact-elmahfoudi", type: "contact", lastName: "ELMAHFOUDI", firstName: "", phone: "06 23 23 23 23", company: "", function: "Représentant chantier", role: "representative", source: "TRAME ISF.docx" },
      { id: "seed-contact-martin", type: "contact", lastName: "MARTIN", firstName: "", phone: "06 23 23 23 23", company: "", function: "Représentant chantier", role: "representative", source: "TRAME ISF.docx" },
      { id: "seed-activity-installation", type: "activity", label: "Installation de chantier", source: "TRAME ISF.docx" },
      { id: "seed-activity-forage", type: "activity", label: "Réalisation de forages", source: "TRAME ISF.docx" },
      { id: "seed-activity-tranchee", type: "activity", label: "Réalisation de tranchées", source: "TRAME ISF.docx" },
      { id: "seed-activity-pots", type: "activity", label: "Pose de pots de terre", source: "TRAME ISF.docx" },
      { id: "seed-activity-mesure", type: "activity", label: "Mesure des terres", source: "TRAME ISF.docx" },
      { id: "seed-set-terre", type: "activitySet", label: "Travaux prise de terre", activities: ["Installation de chantier", "Réalisation de forages", "Réalisation de tranchées", "Pose de pots de terre", "Mesure des terres"], source: "TRAME ISF.docx" },
      { id: "seed-observation-arf", type: "observation", label: "Présence Agent d’Activité et remise ARF", value: "L’entreprise ne pourra commencer son activité qu’en présence de l’Agent d’Activité et après remise de l’ARF.", source: "TRAME ISF.docx" },
    ];
    const ids = new Set(state.library.map((item) => item.id));
    const missing = seeds.filter((item) => !ids.has(item.id));
    if (!missing.length) return;
    for (const item of missing) {
      item.createdAt = Date.now();
      await ISFStorage.put("library", item);
      state.library.push(item);
    }
  }

  async function migrateLegacyData() {
    if (state.settings.legacyMigrationDone) return;
    let legacy = null;
    try { legacy = JSON.parse(localStorage.getItem(LEGACY_KEY) || "null"); } catch { legacy = null; }
    if (legacy?.dossier && Object.values(legacy.dossier).some((value) => normalizeText(value))) {
      const chantier = chantierFromLegacy(legacy.dossier);
      const isf = createBlankIsf(chantier);
      isf.number = legacy.dossier.numeroIsf || "";
      isf.year = legacy.dossier.annee || String(new Date().getFullYear());
      isf.phases = Array.isArray(legacy.isfPlan?.phases) ? legacy.isfPlan.phases.map(convertLegacyPhase) : [];
      isf.migrationSource = "Docu Chantier V3";
      await ISFStorage.put("chantiers", chantier);
      await ISFStorage.put("isfs", isf);
      state.chantiers.push(chantier);
      state.isfs.push(isf);
    }
    await saveSetting("legacyMigrationDone", true);
  }

  function chantierFromLegacy(dossier) {
    return {
      id: createId("chantier"),
      name: dossier.operation || dossier.lieu || "Chantier importé",
      line: dossier.lieu || "",
      csfReference: dossier.referenceCsf || "",
      operation: dossier.operation || "",
      location: dossier.lieu || "",
      workDescription: dossier.natureTravaux || "",
      contacts: { moeId: "", rsoId: "", aspId: "", representativeIds: [] },
      organization: {
        meetingPlace: dossier.lieuRdv || dossier.baseTravaux || "",
        meetingDayTime: "", meetingNightTime: "", careBy: "rso",
        dayHours: dossier.horairesJour || "", nightHours: dossier.horairesNuit || "",
        staffMin: dossier.effectifMin || "", staffMax: dossier.effectifMax || "",
        questions: [null, null, null, null],
      },
      perimeters: [],
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };
  }

  function convertLegacyPhase(phase, index) {
    return {
      id: phase.id || createId("phase"),
      code: phase.code || `A${index + 1}`,
      title: phase.title || "Phase d’activité",
      period: phase.day && phase.night ? "both" : phase.night ? "night" : "day",
      zone: phase.zone || "",
      activities: String(phase.operations || "").split(/\r?\n/).map(normalizeText).filter(Boolean),
      company: phase.company || "",
      observations: phase.observations || "",
      sourceLabel: "Importée depuis le brouillon précédent — protections à vérifier",
      open: index === 0,
      tracks: (phase.tracks || []).map((track) => ({ id: track.id || createId("track"), name: track.name || "Voie à préciser", protections: normalizeProtections(track.protections) })),
    };
  }

  function sortCollections() {
    state.isfs.sort((a, b) => Number(b.updatedAt || 0) - Number(a.updatedAt || 0));
    state.chantiers.sort((a, b) => normalizeSearch(a.name).localeCompare(normalizeSearch(b.name), "fr"));
    state.library.sort((a, b) => libraryLabel(a).localeCompare(libraryLabel(b), "fr"));
  }

  function render() {
    const editor = state.view === "editor";
    document.body.classList.toggle("editor-mode", editor);
    els.mainNav.hidden = editor;
    document.querySelectorAll("[data-nav]").forEach((button) => button.classList.toggle("active", !editor && button.dataset.nav === state.view));
    if (state.view === "dashboard") els.appMain.innerHTML = renderDashboard();
    else if (state.view === "isfs") els.appMain.innerHTML = renderIsfsPage();
    else if (state.view === "chantiers") els.appMain.innerHTML = renderChantiersPage();
    else if (state.view === "chantier-edit") els.appMain.innerHTML = renderChantierEdit();
    else if (state.view === "library") els.appMain.innerHTML = renderLibraryPage();
    else if (state.view === "settings") els.appMain.innerHTML = renderSettingsPage();
    else if (state.view === "editor") els.appMain.innerHTML = renderEditor();
    else els.appMain.innerHTML = renderDashboard();
    els.appMain.focus({ preventScroll: true });
  }

  function renderDashboard() {
    const recent = state.isfs.slice(0, 4);
    const drafts = state.isfs.filter((item) => item.status !== "archived").length;
    return `
      <section class="hero">
        <p class="eyebrow">Rédaction terrain</p>
        <h1>Une ISF complète en quelques minutes</h1>
        <p>Sélectionne un chantier, ajuste ce qui change et vérifie les protections. Le Word original fait le reste.</p>
        <button class="new-isf-button" type="button" data-action="new-isf">＋ NOUVELLE ISF</button>
        <div class="hero-stats"><span>${state.chantiers.length} chantier${plural(state.chantiers.length)}</span><span>${drafts} brouillon${plural(drafts)}</span><span>Sauvegarde automatique</span></div>
      </section>
      <section class="dashboard-grid">
        ${quickCard("isfs", "▤", "Mes ISF", `${state.isfs.length} document${plural(state.isfs.length)}`)}
        ${quickCard("chantiers", "◇", "Chantiers", "Réutiliser les données")}
        ${quickCard("library", "☆", "Bibliothèque", "Contacts et activités")}
        ${quickCard("settings", "⚙", "Paramètres", "Sauvegarde et PWA")}
      </section>
      <div class="section-title"><h2>Dernières ISF</h2>${state.isfs.length > 4 ? `<button type="button" data-nav="isfs">Tout voir</button>` : ""}</div>
      <div class="list-stack">${recent.length ? recent.map(renderIsfListCard).join("") : emptyState("▤", "Aucune ISF pour le moment", "La première peut partir d’un chantier ou d’une opération entièrement nouvelle.", `<button class="primary-button compact-button" type="button" data-action="new-isf">Créer la première ISF</button>`)}</div>`;
  }

  function quickCard(view, icon, title, subtitle) {
    return `<button class="quick-card" type="button" data-nav="${view}"><span aria-hidden="true">${icon}</span><span><strong>${escapeHtml(title)}</strong><small>${escapeHtml(subtitle)}</small></span></button>`;
  }

  function renderIsfsPage() {
    return `
      <div class="page-head"><div><p class="eyebrow">Documents</p><h1>Mes ISF</h1><p>Reprends un brouillon, duplique une ancienne ISF ou retrouve une version archivée.</p></div><button class="round-add" type="button" data-action="new-isf" aria-label="Nouvelle ISF">＋</button></div>
      <div class="search-row"><label class="search-box"><input id="isfSearch" type="search" value="${escapeHtml(state.filters.isfSearch)}" placeholder="Numéro, chantier, ligne, opération…" aria-label="Rechercher une ISF" /></label><button class="primary-button compact-button" type="button" data-action="new-isf">Nouvelle</button></div>
      <div class="filter-chips">${[["all","Toutes"],["draft","Brouillons"],["archived","Archivées"]].map(([key,label]) => `<button class="filter-chip ${state.isfFilter === key ? "active" : ""}" type="button" data-action="set-isf-filter" data-filter="${key}">${label}</button>`).join("")}</div>
      <div id="isfResults" class="list-stack">${renderIsfResults()}</div>`;
  }

  function renderIsfResults() {
    const query = normalizeSearch(state.filters.isfSearch);
    const items = state.isfs.filter((item) => {
      if (state.isfFilter === "draft" && item.status === "archived") return false;
      if (state.isfFilter === "archived" && item.status !== "archived") return false;
      return !query || normalizeSearch(`${isfReference(item)} ${item.chantierName} ${item.line} ${item.operation} ${item.location}`).includes(query);
    });
    return items.length ? items.map(renderIsfListCard).join("") : emptyState("⌕", "Aucune ISF trouvée", "Modifie la recherche ou crée une nouvelle ISF.");
  }

  function renderIsfListCard(isf) {
    const reference = isf.number ? `ISF ${isfReference(isf)} · IND ${isf.index ?? "0"}` : "ISF sans numéro";
    const status = isf.status === "archived" ? "Archivée" : isf.lastGeneratedAt ? "Générée" : "Brouillon";
    return `<article class="list-card" data-action="open-isf" data-isf-id="${escapeHtml(isf.id)}" role="button" tabindex="0"><span class="list-icon ${isf.status === "archived" ? "archived" : ""}">${escapeHtml(isf.index ?? "0")}</span><span><strong>${escapeHtml(reference)}</strong><small>${escapeHtml(isf.chantierName || isf.operation || isf.location || "Chantier à préciser")} · ${formatRelativeDate(isf.updatedAt)}</small></span><em class="${isf.status === "archived" || isf.lastGeneratedAt ? "done" : ""}">${status}</em></article>`;
  }

  function renderChantiersPage() {
    return `
      <div class="page-head"><div><p class="eyebrow">Réutilisation</p><h1>Chantiers</h1><p>Un chantier mémorise la ligne, la CSF, les contacts, l’organisation et le périmètre habituel.</p></div><button class="round-add" type="button" data-action="new-chantier" aria-label="Nouveau chantier">＋</button></div>
      <div class="search-row"><label class="search-box"><input id="chantierSearch" type="search" value="${escapeHtml(state.filters.chantierSearch)}" placeholder="Nom, ligne, CSF, opération…" aria-label="Rechercher un chantier" /></label><button class="primary-button compact-button" type="button" data-action="new-chantier">Nouveau</button></div>
      <div id="chantierResults" class="smart-choice-list">${renderChantierResults()}</div>`;
  }

  function renderChantierResults() {
    const query = normalizeSearch(state.filters.chantierSearch);
    const items = state.chantiers.filter((item) => !query || normalizeSearch(`${item.name} ${item.line} ${item.csfReference} ${item.operation} ${item.location}`).includes(query));
    if (!items.length) return emptyState("◇", "Aucun chantier trouvé", "Crée un chantier une seule fois, puis réutilise-le dans toutes ses ISF.", `<button class="primary-button compact-button" type="button" data-action="new-chantier">Créer un chantier</button>`);
    return items.map((chantier) => {
      const lastIsf = lastIsfForChantier(chantier.id);
      return `<article class="section-card"><div class="card-head"><div><h2>${escapeHtml(chantier.name || "Chantier sans nom")}</h2><p>${escapeHtml(chantier.line || chantier.location || "Ligne à préciser")} · ${escapeHtml(chantier.csfReference || "CSF à préciser")}</p></div><span class="status-pill">${countIsfsForChantier(chantier.id)} ISF</span></div><div class="settings-actions"><button class="primary-button compact-button" type="button" data-action="new-from-chantier" data-chantier-id="${escapeHtml(chantier.id)}">＋ Nouvelle ISF</button>${lastIsf ? `<button class="secondary-button compact-button" type="button" data-action="duplicate-last-chantier" data-chantier-id="${escapeHtml(chantier.id)}">⧉ Reprendre la dernière</button>` : ""}<button class="secondary-button compact-button" type="button" data-action="edit-chantier" data-chantier-id="${escapeHtml(chantier.id)}">Modifier</button><button class="ghost-button compact-button" type="button" data-action="delete-chantier" data-chantier-id="${escapeHtml(chantier.id)}">Supprimer</button></div></article>`;
    }).join("");
  }

  function renderChantierEdit() {
    const chantier = state.editingChantier;
    if (!chantier) return renderChantiersPage();
    return `
      <div class="page-head"><div><p class="eyebrow">Chantier</p><h1>${escapeHtml(chantier.name || "Nouveau chantier")}</h1><p>Ces valeurs seront proposées automatiquement dans les prochaines ISF.</p></div><button class="secondary-button compact-button" type="button" data-action="close-chantier-edit">Terminer</button></div>
      <section class="section-card highlight"><div class="card-head"><div><h2>Identité du chantier</h2><p>Les informations les plus réutilisées.</p></div><span class="status-pill good">Sauvegarde auto</span></div><div class="field-grid two">
        ${chantierField("Nom du chantier", "name", chantier.name, "Ex. Montereau — renouvellement 2026", true)}
        ${chantierField("Ligne ferroviaire", "line", chantier.line, "Ex. Ligne Corbeil à Melun")}
        ${chantierField("Référence CSF", "csfReference", chantier.csfReference, "Ex. 194 / 022 / 2026")}
        ${chantierField("Opération", "operation", chantier.operation, "Intitulé de l’opération")}
        ${chantierField("Lieu", "location", chantier.location, "Gare, commune, zone")}
        ${chantierField("Travaux habituels", "workDescription", chantier.workDescription, "Nature des travaux")}
      </div></section>
      <section class="section-card"><div class="card-head"><div><h2>Organisation habituelle</h2><p>Modifiable ensuite dans chaque ISF.</p></div></div><div class="field-grid two">
        ${chantierField("Lieu de rendez-vous", "organization.meetingPlace", chantier.organization?.meetingPlace, "Base travaux, adresse…")}
        ${chantierField("Horaires de jour", "organization.dayHours", chantier.organization?.dayHours, "Ex. 08h00 – 17h00")}
        ${chantierField("Horaires de nuit", "organization.nightHours", chantier.organization?.nightHours, "Ex. 22h00 – 06h00")}
        ${chantierField("Effectif minimum", "organization.staffMin", chantier.organization?.staffMin, "3", false, "number")}
        ${chantierField("Effectif maximum", "organization.staffMax", chantier.organization?.staffMax, "15", false, "number")}
      </div></section>
      <section class="section-card"><div class="card-head"><div><h2>Intervenants habituels</h2><p>Choisis dans la bibliothèque.</p></div></div>${renderChantierContactSelectors(chantier)}</section>`;
  }

  function chantierField(label, path, value, placeholder, wide = false, type = "text") {
    return `<label class="field ${wide ? "wide" : ""}">${escapeHtml(label)}<input type="${type}" data-chantier-path="${escapeHtml(path)}" value="${escapeHtml(value || "")}" placeholder="${escapeHtml(placeholder || "")}" /></label>`;
  }

  function renderChantierContactSelectors(chantier) {
    return `<div class="contact-role-list">${[["MOE Tx","moeId","moe"],["RSO","rsoId","rso"],["ASP","aspId","asp"]].map(([label,key,role]) => `<div class="contact-role"><strong>${label}</strong><label class="field">Intervenant<select data-chantier-path="contacts.${key}">${contactOptions(chantier.contacts?.[key], role)}</select></label><button type="button" data-action="open-contact-dialog" data-contact-target="chantier:contacts.${key}" data-role="${role}" aria-label="Ajouter un intervenant">＋</button></div>`).join("")}</div>`;
  }

  function renderLibraryPage() {
    return `
      <div class="page-head"><div><p class="eyebrow">Saisir une fois</p><h1>Bibliothèque</h1><p>Les intervenants, activités et phrases fréquentes sont proposés dans toutes les ISF.</p></div><button class="round-add" type="button" data-action="add-library-item" data-type="${state.libraryFilter}" aria-label="Ajouter">＋</button></div>
      <div class="search-row"><label class="search-box"><input id="librarySearch" type="search" value="${escapeHtml(state.filters.librarySearch)}" placeholder="Recherche instantanée…" aria-label="Rechercher dans la bibliothèque" /></label><button class="primary-button compact-button" type="button" data-action="add-library-item" data-type="${state.libraryFilter}">Ajouter</button></div>
      <div class="filter-chips">${LIBRARY_TYPES.map(([key,label]) => `<button class="filter-chip ${state.libraryFilter === key ? "active" : ""}" type="button" data-action="set-library-filter" data-filter="${key}">${label}</button>`).join("")}</div>
      <div id="libraryResults" class="library-grid">${renderLibraryResults()}</div>`;
  }

  function renderLibraryResults() {
    const query = normalizeSearch(state.filters.librarySearch);
    const items = state.library.filter((item) => item.type === state.libraryFilter && (!query || normalizeSearch(`${libraryLabel(item)} ${librarySubtitle(item)}`).includes(query)));
    if (!items.length) return emptyState("☆", "Aucun élément trouvé", "Ajoute une valeur que tu réutilises souvent.");
    return items.map((item) => `<article class="library-card"><span>${libraryIcon(item.type)}</span><div><strong>${escapeHtml(libraryLabel(item))}</strong><small>${escapeHtml(librarySubtitle(item))}</small></div><div class="library-actions"><button type="button" data-action="edit-library-item" data-library-id="${escapeHtml(item.id)}" aria-label="Modifier">✎</button><button class="delete" type="button" data-action="delete-library-item" data-library-id="${escapeHtml(item.id)}" aria-label="Supprimer">×</button></div></article>`).join("");
  }

  function renderSettingsPage() {
    const size = state.isfs.length + state.chantiers.length + state.library.length;
    return `
      <div class="page-head"><div><p class="eyebrow">Application</p><h1>Paramètres</h1><p>Sauvegarde, transfert des données et informations sur la version installée.</p></div></div>
      <div class="settings-list">
        <section class="settings-card"><h2>Sauvegarde des données</h2><p>Exporte les ISF, chantiers et bibliothèques dans un fichier JSON pour les transférer ou les sécuriser.</p><div class="settings-actions"><button class="primary-button compact-button" type="button" data-action="export-backup">Exporter la sauvegarde</button><button class="secondary-button compact-button" type="button" data-action="import-backup">Importer</button></div></section>
        <section class="settings-card"><h2>Stockage local</h2><p>${size} élément${plural(size)} enregistré${plural(size)} sur cet appareil. Les mises à jour de l’application ne suppriment pas IndexedDB.</p><div class="technical-note">${state.isfs.length} ISF · ${state.chantiers.length} chantier${plural(state.chantiers.length)} · ${state.library.length} élément${plural(state.library.length)} de bibliothèque</div></section>
        <section class="settings-card"><h2>Modèle Word</h2><p>La génération utilise directement la vraie trame ISF incluse dans l’application.</p><div class="technical-note">templates/TRAME_ISF.docx · mise en page, logos, annexes et pages de remplacement du RSO conservés.</div></section>
        <section class="settings-card"><h2>Version</h2><p>${APP_VERSION} · PWA compatible Android, Chrome et Edge · fonctionnement hors connexion après le premier chargement.</p><div class="settings-actions"><button class="ghost-button compact-button" type="button" data-action="check-update">Vérifier les mises à jour</button></div></section>
        <section class="settings-card"><h2>Réinitialisation</h2><p>Cette action efface les données locales de l’application sur cet appareil. Une confirmation est demandée.</p><button class="danger-button compact-button" type="button" data-action="clear-all-data">Tout effacer</button></section>
      </div>`;
  }

  function renderEditor() {
    const isf = state.currentIsf;
    if (!isf) return renderDashboard();
    const step = STEPS[state.editorStep - 1];
    const completeSteps = getCompleteSteps(isf);
    return `
      <section class="editor-shell">
        <header class="editor-head">
          <div class="editor-topline"><button class="back-button" type="button" data-action="close-editor" aria-label="Fermer l’éditeur">←</button><div class="editor-title"><strong>${escapeHtml(isf.number ? `ISF ${isfReference(isf)} · IND ${isf.index}` : "Nouvelle ISF")}</strong><small>${escapeHtml(isf.chantierName || isf.operation || "Chantier à définir")}</small></div><span class="autosave-badge">Auto</span></div>
          <div class="progress-track" aria-hidden="true"><span style="width:${state.editorStep * 20}%"></span></div>
          <nav class="stepper" aria-label="Étapes de rédaction">${STEPS.map((item,index) => `<button class="step-tab ${state.editorStep === index + 1 ? "active" : ""} ${completeSteps.has(index + 1) ? "complete" : ""}" type="button" data-action="go-step" data-step="${index + 1}"><span>${completeSteps.has(index + 1) ? "✓" : index + 1}</span><strong>${escapeHtml(item.label)}</strong></button>`).join("")}</nav>
        </header>
        <section class="step-content" data-step-panel="${state.editorStep}">${renderEditorStep(isf, step)}</section>
      </section>
      <footer class="editor-footer"><button class="secondary-button" type="button" data-action="previous-step" ${state.editorStep === 1 ? "disabled" : ""}>← Précédent</button><button class="primary-button" type="button" data-action="next-step">${state.editorStep === 5 ? "Terminer" : state.editorStep === 4 ? "Préparer le Word →" : "Continuer →"}</button></footer>
      ${renderDatalists()}`;
  }

  function renderEditorStep(isf) {
    if (state.editorStep === 1) return renderStepChantier(isf);
    if (state.editorStep === 2) return renderStepIntervention(isf);
    if (state.editorStep === 3) return renderStepPhases(isf);
    if (state.editorStep === 4) return renderStepReview(isf);
    return renderStepGeneration(isf);
  }

  function renderStepChantier(isf) {
    return `
      <div class="step-heading"><p class="eyebrow">Étape 1 sur 5</p><h1>Quel chantier&nbsp;?</h1><p>Sélectionne un chantier connu ou complète uniquement les informations essentielles.</p></div>
      ${state.chantiers.length ? `<section class="section-card"><div class="card-head"><div><h2>Reprendre un chantier existant</h2><p>Toutes les informations associées seront reprises.</p></div></div><label class="field">Chantier<select data-action-change="apply-chantier"><option value="">Choisir…</option>${state.chantiers.map((item) => `<option value="${escapeHtml(item.id)}" ${item.id === isf.chantierId ? "selected" : ""}>${escapeHtml(item.name || item.operation || item.location)}</option>`).join("")}</select></label></section>` : ""}
      <section class="section-card highlight" id="field-chantier"><div class="card-head"><div><h2>Informations essentielles</h2><p>Une seule saisie pour toute l’ISF.</p></div><span class="status-pill ${isf.chantierId ? "good" : "warn"}">${isf.chantierId ? "Chantier lié" : "Nouveau"}</span></div><div class="field-grid two">
        ${editorInput("Nom du chantier", "chantierName", isf.chantierName, "Ex. Montereau — renouvellement 2026", "text", true, "chantierName")}
        ${editorInput("Ligne ferroviaire", "line", isf.line, "Ex. Ligne Corbeil à Melun", "text", false, "line", "lines")}
        ${editorInput("Référence CSF", "csfReference", isf.csfReference, "Ex. 194 / 022 / 2026", "text", false, "csfReference", "csfs")}
        ${editorInput("Opération", "operation", isf.operation, "Intitulé de l’opération", "text", false, "operation")}
        ${editorInput("Lieu", "location", isf.location, "Gare, commune, secteur…", "text", false, "location", "places")}
        ${editorTextarea("Nature habituelle des travaux", "workDescription", isf.workDescription, "Travaux réalisés sur ce chantier", 2, true, "workDescription")}
      </div><div class="settings-actions" style="margin-top:13px"><button class="secondary-button compact-button" type="button" data-action="save-current-chantier">${isf.chantierId ? "Mettre à jour le chantier" : "Enregistrer comme chantier"}</button></div></section>
      <section class="section-card"><div class="card-head"><div><h2>Résumé repris automatiquement</h2><p>Tu pourras tout ajuster à l’étape suivante.</p></div></div><div class="review-grid">${miniSummary("Intervenants", `${countSelectedContacts(isf)} sélectionné${plural(countSelectedContacts(isf))}`)}${miniSummary("Organisation", isf.organization?.meetingPlace ? "Lieu de rendez-vous connu" : "À compléter")}${miniSummary("Périmètre", `${isf.perimeters.length} zone${plural(isf.perimeters.length)}`)}</div></section>`;
  }

  function renderStepIntervention(isf) {
    return `
      <div class="step-heading"><p class="eyebrow">Étape 2 sur 5</p><h1>Qu’est-ce qui change&nbsp;?</h1><p>Numéro, indice, dates, intervenants et périmètre de cette intervention.</p></div>
      <section class="section-card highlight" id="field-reference"><div class="card-head"><div><h2>Référence et dates</h2><p>Les éléments propres à cette ISF.</p></div><span class="status-pill source">IND ${escapeHtml(isf.index ?? "0")}</span></div><div class="field-grid three">
        ${editorInput("Série", "series", isf.series, "194", "text", false, "series")}
        ${editorInput("Numéro ISF", "number", isf.number, "Ex. 31", "text", false, "number")}
        ${editorInput("Année", "year", isf.year, String(new Date().getFullYear()), "number", false, "year")}
        ${editorInput("Indice", "index", isf.index, "0", "text", false, "index")}
        ${editorInput("Début", "startDate", isf.startDate, "", "date", false, "startDate")}
        ${editorInput("Fin", "endDate", isf.endDate, "", "date", false, "endDate")}
        ${editorInput("Date de l’indice", "revisionDate", isf.revisionDate, "", "date", false, "revisionDate")}
        ${editorInput("Semaines", "weeks", isf.weeks, "Ex. 8 et 9", "text", false, "weeks")}
        ${editorInput("Modification apportée", "modification", isf.modification, "Ex. Adaptation du périmètre", "text", true, "modification")}
        ${editorTextarea("Travaux de cette ISF", "workDescription", isf.workDescription, "Ce qui sera réalisé pendant l’intervention", 2, true, "workDescription")}
      </div></section>
      <section class="section-card" id="field-contacts"><div class="card-head"><div><h2>Intervenants</h2><p>Choisis un nom : téléphone, entreprise et fonction sont repris automatiquement.</p></div><button class="secondary-button compact-button" type="button" data-action="open-contact-dialog" data-role="">＋ Intervenant</button></div>${renderContactSelectors(isf)}</section>
      <details class="collapsible" open><summary><span><strong>Organisation générale</strong><small>Rendez-vous, horaires, effectifs et questions OUI/NON</small></span></summary><div class="collapsible-content">${renderOrganization(isf)}</div></details>
      <details class="collapsible" open id="field-perimeter"><summary><span><strong>Périmètre</strong><small>${isf.perimeters.length} zone${plural(isf.perimeters.length)} · collage multiple possible</small></span></summary><div class="collapsible-content">${renderPerimeter(isf)}</div></details>
      <details class="collapsible"><summary><span><strong>Historique des indices</strong><small>${isf.revisions.length} version${plural(isf.revisions.length)} conservée${plural(isf.revisions.length)}</small></span></summary><div class="collapsible-content"><div class="list-stack">${isf.revisions.map((revision) => `<div class="review-item"><span>${escapeHtml(revision.index ?? "0")}</span><span><strong>IND ${escapeHtml(revision.index ?? "0")} · ${escapeHtml(formatDate(revision.date))}</strong><small>${escapeHtml(revision.change || "Sans commentaire")}</small></span><em>Conservé</em></div>`).join("")}</div></div></details>`;
  }

  function renderContactSelectors(isf) {
    const roles = [["MOE Tx","moeId","moe"],["RSO","rsoId","rso"],["ASP","aspId","asp"]];
    return `<div class="contact-role-list">${roles.map(([label,key,role]) => `<div class="contact-role"><strong>${label}</strong><label class="field">Intervenant<select data-path="contacts.${key}" data-contact-role="${role}">${contactOptions(isf.contacts?.[key], role)}</select></label><button type="button" data-action="open-contact-dialog" data-contact-target="isf:contacts.${key}" data-role="${role}" aria-label="Ajouter ${label}">＋</button></div>`).join("")}</div><div class="card-head" style="margin:14px 0 7px"><div><h2>Représentants entreprise</h2><p>Nombre variable selon le chantier.</p></div><button class="ghost-button compact-button" type="button" data-action="add-representative">＋ Ajouter</button></div><div class="representative-list">${isf.contacts.representativeIds.length ? isf.contacts.representativeIds.map((id,index) => `<div class="representative-row"><select data-representative-index="${index}" aria-label="Représentant ${index + 1}">${contactOptions(id, "representative")}</select><button type="button" data-action="remove-representative" data-index="${index}" aria-label="Retirer">×</button></div>`).join("") : `<p class="helper">Aucun représentant sélectionné.</p>`}</div>`;
  }

  function renderOrganization(isf) {
    const org = isf.organization;
    return `<div class="field-grid two">
      ${editorInput("Lieu de rendez-vous", "organization.meetingPlace", org.meetingPlace, "Base travaux, adresse…", "text", true, "meetingPlace", "places")}
      ${editorInput("Heure rendez-vous jour", "organization.meetingDayTime", org.meetingDayTime, "Ex. 08h00", "text")}
      ${editorInput("Heure rendez-vous nuit", "organization.meetingNightTime", org.meetingNightTime, "Ex. 22h00", "text")}
      ${editorInput("Horaires de travail jour", "organization.dayHours", org.dayHours, "Ex. 08h00 – 17h00", "text")}
      ${editorInput("Horaires de travail nuit", "organization.nightHours", org.nightHours, "Ex. 22h00 – 06h00", "text")}
      ${editorInput("Effectif minimum", "organization.staffMin", org.staffMin, "3", "number")}
      ${editorInput("Effectif maximum", "organization.staffMax", org.staffMax, "15", "number")}
      <div class="field wide">Prise en charge<div class="segmented three">${[["rso","RSO"],["asp","ASP"],["autre","Autre"]].map(([value,label]) => `<label><input type="radio" name="careBy" data-path="organization.careBy" value="${value}" ${org.careBy === value ? "checked" : ""} />${label}</label>`).join("")}</div></div>
    </div><div class="yes-no-list" style="margin-top:13px">${QUESTION_LABELS.map((label,index) => `<div class="yes-no-row"><span>${escapeHtml(label)}</span><div class="segmented"><label><input type="radio" name="question-${index}" data-path="organization.questions.${index}" value="true" ${org.questions[index] === true ? "checked" : ""} />OUI</label><label><input type="radio" name="question-${index}" data-path="organization.questions.${index}" value="false" ${org.questions[index] === false ? "checked" : ""} />NON</label></div></div>`).join("")}</div>`;
  }

  function renderPerimeter(isf) {
    return `<div class="perimeter-add"><input id="newPerimeterCenter" placeholder="Centre / zone" aria-label="Centre ou zone" /><input id="newPerimeterPk" placeholder="PK" aria-label="PK" /><button type="button" data-action="add-perimeter" aria-label="Ajouter">＋</button></div>
      <div class="perimeter-list">${isf.perimeters.map((item,index) => `<div class="perimeter-row"><input data-perimeter-index="${index}" data-perimeter-field="center" value="${escapeHtml(item.center || "")}" aria-label="Centre ${index + 1}" /><input data-perimeter-index="${index}" data-perimeter-field="pk" value="${escapeHtml(item.pk || "")}" aria-label="PK ${index + 1}" /><button class="duplicate" type="button" data-action="duplicate-perimeter" data-index="${index}" aria-label="Dupliquer la zone" title="Dupliquer">⧉</button><button type="button" data-action="remove-perimeter" data-index="${index}" aria-label="Supprimer">×</button></div>`).join("")}</div>
      <details class="bulk-paste"><summary>Coller plusieurs lignes en une fois</summary><textarea id="bulkPerimeter" placeholder="HT 35,0 ; 35,000\nHT 36,5 ; 36,500\nAr HT.BT + SE ; 37,955"></textarea><button class="secondary-button compact-button full" type="button" data-action="parse-perimeters" style="margin-top:7px">Créer les zones</button></details>
      <div class="field" style="margin-top:12px">Situation géographique jointe<div class="segmented"><label><input type="radio" name="geo" data-path="geographicAnnex" value="true" ${isf.geographicAnnex === true ? "checked" : ""} />OUI</label><label><input type="radio" name="geo" data-path="geographicAnnex" value="false" ${isf.geographicAnnex === false ? "checked" : ""} />NON</label></div></div>`;
  }

  function renderStepPhases(isf) {
    const activitySets = state.library.filter((item) => item.type === "activitySet");
    const phaseTemplates = state.library.filter((item) => item.type === "phaseTemplate");
    return `
      <div class="step-heading"><p class="eyebrow">Étape 3 sur 5</p><h1>Phases et protections</h1><p>Les protections restent toujours visibles. L’application ne décide jamais à ta place.</p></div>
      <section class="phase-quick-add"><strong>Ajouter une phase rapidement</strong><small>Un ensemble insère uniquement les activités. Les protections restent désactivées tant que tu ne les choisis pas.</small><div class="preset-row"><button class="preset-chip" type="button" data-action="add-phase">＋ Phase libre</button>${activitySets.map((item) => `<button class="preset-chip" type="button" data-action="add-phase-set" data-library-id="${escapeHtml(item.id)}">＋ ${escapeHtml(item.label)}</button>`).join("")}${phaseTemplates.map((item) => `<button class="preset-chip" type="button" data-action="add-phase-template" data-library-id="${escapeHtml(item.id)}">⧉ ${escapeHtml(item.label)}</button>`).join("")}</div></section>
      <div class="safety-warning"><strong>Règle de sécurité :</strong> les configurations copiées sont seulement reprises pour gagner du temps. Elles doivent être contrôlées puis validées explicitement à l’étape 4.</div>
      <div style="margin-top:10px">${isf.phases.length ? isf.phases.map(renderPhaseCard).join("") : emptyState("1", "Aucune phase", "Ajoute une phase libre ou l’ensemble « Travaux prise de terre ».")}</div>`;
  }

  function renderPhaseCard(phase, index) {
    const activitySuggestions = state.library.filter((item) => item.type === "activity").slice(0, 8);
    const observationSuggestions = state.library.filter((item) => item.type === "observation").slice(0, 6);
    const period = phase.period || "day";
    return `<article class="phase-card ${phase.open ? "open" : "collapsed"}" id="phase-${escapeHtml(phase.id)}">
      <header class="phase-header"><button class="phase-toggle" type="button" data-action="toggle-phase" data-phase-id="${escapeHtml(phase.id)}"><span class="phase-number">${escapeHtml(phase.code || `A${index + 1}`)}</span><span><strong>${escapeHtml(phase.title || "Phase à nommer")}</strong><small>${escapeHtml(periodLabel(period))} · ${phase.activities.length} activité${plural(phase.activities.length)} · ${phase.tracks.length} voie${plural(phase.tracks.length)}</small></span><span class="phase-chevron">⌄</span></button><div class="phase-actions"><button type="button" data-action="duplicate-phase" data-phase-id="${escapeHtml(phase.id)}">Dupliquer</button><button class="remove" type="button" data-action="remove-phase" data-phase-id="${escapeHtml(phase.id)}" aria-label="Supprimer">×</button></div></header>
      <div class="phase-body">${phase.sourceLabel ? `<div class="phase-source"><strong>${escapeHtml(phase.sourceLabel)}</strong>Les protections sont affichées ci-dessous et doivent être vérifiées.</div>` : ""}
        <div class="field-grid two">${phaseInput("Code",phase,"code",phase.code,"A1")}${phaseInput("Intitulé",phase,"title",phase.title,"Ex. Installation de chantier",true)}${phaseInput("Zone de travail",phase,"zone",phase.zone,"Voies, PK, enveloppe de briefing",true)}${phaseInput("Entreprise",phase,"company",phase.company,"Entreprise intervenante",true,"companies")}</div>
        <div class="field period-choice">Travaux réalisés<div class="segmented three">${[["day","Jour"],["night","Nuit"],["both","Jour + nuit"]].map(([value,label]) => `<label><input type="radio" name="period-${escapeHtml(phase.id)}" data-phase-id="${escapeHtml(phase.id)}" data-phase-field="period" value="${value}" ${period === value ? "checked" : ""} />${label}</label>`).join("")}</div></div>
        <section class="activity-section"><h3>Activités</h3><p>Ajoute depuis les favoris ou saisis une activité. Aucune grille Word à remplir.</p><div class="activity-chips">${phase.activities.map((label,activityIndex) => `<span class="activity-chip">${escapeHtml(label)}<button type="button" data-action="remove-activity" data-phase-id="${escapeHtml(phase.id)}" data-index="${activityIndex}" aria-label="Retirer">×</button></span>`).join("")}</div><div class="activity-add"><input data-new-activity="${escapeHtml(phase.id)}" placeholder="Ajouter une activité" list="activities" /><button type="button" data-action="add-activity" data-phase-id="${escapeHtml(phase.id)}" aria-label="Ajouter l’activité">＋</button></div><div class="suggestion-scroller">${activitySuggestions.map((item) => `<button type="button" data-action="add-suggested-activity" data-phase-id="${escapeHtml(phase.id)}" data-label="${escapeHtml(item.label)}">＋ ${escapeHtml(item.label)}</button>`).join("")}</div></section>
        <section class="tracks-section"><h3>Voies et mesures de protection</h3><p>Ajoute plusieurs voies avec une virgule. Chaque voie conserve sa propre configuration.</p><div class="track-add"><input data-new-tracks="${escapeHtml(phase.id)}" placeholder="V1C, V2C, Faisceau pair…" /><button type="button" data-action="add-tracks" data-phase-id="${escapeHtml(phase.id)}">Ajouter les voies</button></div>${phase.tracks.length > 1 ? `<div class="track-tools"><button type="button" data-action="apply-first-track" data-phase-id="${escapeHtml(phase.id)}">Appliquer la 1re voie à toutes</button><button type="button" data-action="clear-phase-protections" data-phase-id="${escapeHtml(phase.id)}">Tout désactiver</button></div>` : ""}<div class="track-list">${phase.tracks.map((track,trackIndex) => renderTrackCard(phase,track,trackIndex)).join("")}</div></section>
        <details class="collapsible" style="margin-top:12px"><summary><span><strong>Observations</strong><small>${phase.observations ? "Renseignées" : "Facultatif"}</small></span></summary><div class="collapsible-content"><label class="field">Texte<textarea rows="3" data-phase-id="${escapeHtml(phase.id)}" data-phase-field="observations" placeholder="Observations et restrictions…">${escapeHtml(phase.observations || "")}</textarea></label><div class="suggestion-scroller">${observationSuggestions.map((item) => `<button type="button" data-action="add-observation" data-phase-id="${escapeHtml(phase.id)}" data-label="${escapeHtml(item.value || item.label)}">＋ ${escapeHtml(item.label)}</button>`).join("")}</div></div></details>
        <button class="text-button" type="button" data-action="save-phase-template" data-phase-id="${escapeHtml(phase.id)}">☆ Enregistrer cette phase dans la bibliothèque</button>
      </div></article>`;
  }

  function renderTrackCard(phase, track, trackIndex) {
    return `<article class="track-card"><div class="track-head ${trackIndex > 0 ? "has-copy" : ""}"><input data-phase-id="${escapeHtml(phase.id)}" data-track-id="${escapeHtml(track.id)}" data-track-field="name" value="${escapeHtml(track.name || "")}" aria-label="Nom de la voie" />${trackIndex > 0 ? `<button class="copy-track" type="button" data-action="copy-previous-track" data-phase-id="${escapeHtml(phase.id)}" data-track-id="${escapeHtml(track.id)}" title="Copier les protections de la voie précédente">Copier ↑</button>` : ""}<button type="button" data-action="remove-track" data-phase-id="${escapeHtml(phase.id)}" data-track-id="${escapeHtml(track.id)}" aria-label="Supprimer la voie">×</button></div><div class="protection-list">${PROTECTIONS.map(([key,label]) => `<label class="protection-row"><span>${escapeHtml(label)}</span><input type="checkbox" data-phase-id="${escapeHtml(phase.id)}" data-track-id="${escapeHtml(track.id)}" data-protection="${key}" ${track.protections?.[key] ? "checked" : ""} /><i class="toggle" aria-hidden="true"></i></label>`).join("")}</div></article>`;
  }

  function renderStepReview(isf) {
    syncValidationParticipants(isf);
    const anomalies = collectAnomalies(isf);
    const dataAnomalies = anomalies.filter((item) => item.key !== "protectionsVerified");
    const tracks = countTracks(isf);
    return `
      <div class="step-heading"><p class="eyebrow">Étape 4 sur 5</p><h1>Vérification</h1><p>Un résumé utile, sans réafficher tout le formulaire.</p></div>
      <section class="section-card"><div class="card-head"><div><h2>Résumé de l’ISF</h2><p>Les rubriques essentielles du document.</p></div><span class="status-pill ${dataAnomalies.length ? "warn" : "good"}">${dataAnomalies.length ? `${dataAnomalies.length} à compléter` : "Complet"}</span></div><div class="review-grid">
        ${reviewRow("Numéro ISF", isf.number ? isfReference(isf) : "Manquant", Boolean(isf.number))}
        ${reviewRow("Chantier", isf.chantierName || isf.operation || "Manquant", Boolean(isf.chantierName || isf.operation))}
        ${reviewRow("CSF", isf.csfReference || "Manquante", Boolean(isf.csfReference))}
        ${reviewRow("Dates", isf.startDate && isf.endDate ? `${formatDate(isf.startDate)} → ${formatDate(isf.endDate)}` : "À compléter", Boolean(isf.startDate && isf.endDate))}
        ${reviewRow("Intervenants", `${countSelectedContacts(isf)} sélectionné${plural(countSelectedContacts(isf))}`, Boolean(isf.contacts.moeId && isf.contacts.rsoId && isf.contacts.aspId))}
        ${reviewRow("Périmètre", `${isf.perimeters.length} zone${plural(isf.perimeters.length)}`, isf.perimeters.length > 0)}
        ${reviewRow("Phases d’activité", `${isf.phases.length} phase${plural(isf.phases.length)}`, isf.phases.length > 0)}
        ${reviewRow("Voies", `${tracks} voie${plural(tracks)}`, tracks > 0)}
      </div></section>
      ${dataAnomalies.length ? `<section class="section-card"><div class="card-head"><div><h2>Anomalies à corriger</h2><p>Appuie sur une ligne pour aller directement à l’information.</p></div></div><div class="anomaly-list">${dataAnomalies.map((item) => `<button class="anomaly-button" type="button" data-action="jump-anomaly" data-step="${item.step}" data-target="${escapeHtml(item.target || "")}"><span>⚠</span><strong>${escapeHtml(item.label)}</strong><em>›</em></button>`).join("")}</div></section>` : ""}
      <section class="section-card"><div class="card-head"><div><h2>Validation prévue dans le document</h2><p>La signature peut être ajoutée ultérieurement. Les données connues sont préremplies.</p></div></div><div class="review-grid">${isf.validationParticipants.map((person,index) => {
        const contact = getContact(person.contactId);
        return `<div class="review-item"><span>${person.included !== false ? "✓" : "—"}</span><span><strong>${escapeHtml(contactName(contact) || "Intervenant inconnu")}</strong><small>${escapeHtml(contact?.company || "")} · ${escapeHtml(contact?.function || "Fonction à préciser")}</small></span><label class="check-confirm" style="min-height:34px;padding:5px 8px"><input type="checkbox" data-validation-index="${index}" data-validation-field="included" ${person.included !== false ? "checked" : ""} /><span>Inclure</span></label></div>`;
      }).join("") || `<p class="helper">Les intervenants sélectionnés seront proposés ici.</p>`}</div></section>
      <section class="safety-gate ${isf.protectionsVerified ? "verified" : ""}" id="field-safety"><h2>Mesures de protection vérifiées</h2><p>Cette validation est obligatoire. Elle confirme uniquement ton contrôle ; l’application ne détermine aucune mesure ferroviaire.</p><label class="check-confirm"><input type="checkbox" data-path="protectionsVerified" ${isf.protectionsVerified ? "checked" : ""} /><span>J’ai vérifié les mesures de protection de chaque voie et de chaque phase.</span></label>${isf.protectionsVerified ? `<label class="field" style="margin-top:10px">Vérifié par<input data-path="protectionsVerifiedBy" value="${escapeHtml(isf.protectionsVerifiedBy || "")}" placeholder="Nom / fonction (facultatif)" /></label>` : ""}</section>
      <section class="section-card" style="margin-top:11px"><div class="card-head"><div><h2>Type de document</h2><p>Le brouillon conserve le jaune ; la version finale le retire.</p></div></div><div class="segmented"><label><input type="radio" name="generationMode" data-path="generationMode" value="draft" ${isf.generationMode === "draft" ? "checked" : ""} />Brouillon</label><label><input type="radio" name="generationMode" data-path="generationMode" value="final" ${isf.generationMode !== "draft" ? "checked" : ""} />Version finale</label></div></section>`;
  }

  function renderStepGeneration(isf) {
    const anomalies = collectAnomalies(isf);
    const hardAnomalies = anomalies.filter((item) => item.key !== "protectionsVerified");
    const finalBlocked = isf.generationMode !== "draft" && hardAnomalies.length > 0;
    const safetyBlocked = !isf.protectionsVerified;
    const filename = DocxEngine.createFilename(isf, isf.generationMode || "final");
    return `
      <div class="step-heading"><p class="eyebrow">Étape 5 sur 5</p><h1>Générer le document</h1><p>Le fichier est construit directement depuis TRAME ISF.docx, sur ton appareil.</p></div>
      <section class="generation-hero"><p class="eyebrow">${isf.generationMode === "draft" ? "Mode brouillon" : "Version finale"}</p><h1>${safetyBlocked || finalBlocked ? "Le Word attend une vérification" : "Ton ISF est prête"}</h1><p>${safetyBlocked ? "Valide d’abord les mesures de protection." : finalBlocked ? `${hardAnomalies.length} information${plural(hardAnomalies.length)} obligatoire${plural(hardAnomalies.length)} reste${hardAnomalies.length > 1 ? "nt" : ""} à compléter.` : "La mise en page, les tableaux, logos, annexes et pages de remplacement du RSO sont conservés."}</p><div class="generation-name">${escapeHtml(filename)}</div><div class="generation-actions"><button class="primary-button" type="button" data-action="generate-word" ${safetyBlocked || finalBlocked ? "disabled" : ""}>↓ GÉNÉRER WORD</button>${safetyBlocked || finalBlocked ? `<button class="secondary-button" type="button" data-action="go-step" data-step="4">Retour à la vérification</button>` : `<button class="secondary-button" type="button" data-action="generate-draft">Générer aussi un brouillon</button>`}</div></section>
      <div class="action-grid"><button class="action-card" type="button" data-action="duplicate-current"><span>⧉</span><strong>Dupliquer l’ISF</strong><small>Indice suivant proposé</small></button><button class="action-card" type="button" data-action="archive-current"><span>▣</span><strong>${isf.status === "archived" ? "Désarchiver" : "Archiver"}</strong><small>La version reste consultable</small></button><button class="action-card" type="button" data-action="save-current-chantier"><span>◇</span><strong>Mettre à jour le chantier</strong><small>Réutiliser ces valeurs</small></button><button class="action-card" type="button" data-action="export-current-json"><span>{ }</span><strong>Exporter les données</strong><small>Cette ISF au format JSON</small></button></div>
      <section class="section-card" style="margin-top:11px"><div class="card-head"><div><h2>Fidélité documentaire</h2><p>Contrôles intégrés au moteur de génération.</p></div><span class="status-pill good">Trame originale</span></div><div class="review-grid">${reviewRow("Modèle", "TRAME ISF.docx", true)}${reviewRow("Phases", `${isf.phases.length} tableau${plural(isf.phases.length)} dynamique${plural(isf.phases.length)}`, true)}${reviewRow("Remplacement RSO", "2 pages conservées", true)}${reviewRow("Surlignage", isf.generationMode === "draft" ? "Conservé" : "Retiré", true)}</div></section>`;
  }

  function renderDatalists() {
    const values = (type) => state.library.filter((item) => item.type === type).map((item) => item.label || item.value).filter(Boolean);
    return `<datalist id="activities">${values("activity").map((value) => `<option value="${escapeHtml(value)}"></option>`).join("")}</datalist><datalist id="companies">${values("company").map((value) => `<option value="${escapeHtml(value)}"></option>`).join("")}</datalist><datalist id="lines">${values("line").map((value) => `<option value="${escapeHtml(value)}"></option>`).join("")}</datalist><datalist id="csfs">${values("csf").map((value) => `<option value="${escapeHtml(value)}"></option>`).join("")}</datalist><datalist id="places">${values("place").map((value) => `<option value="${escapeHtml(value)}"></option>`).join("")}</datalist>`;
  }

  function editorInput(label, path, value, placeholder = "", type = "text", wide = false, target = "", list = "") {
    return `<label class="field ${wide ? "wide" : ""}" ${target ? `id="field-${escapeHtml(target)}"` : ""}>${escapeHtml(label)}<input type="${type}" data-path="${escapeHtml(path)}" value="${escapeHtml(value || "")}" placeholder="${escapeHtml(placeholder)}" ${list ? `list="${list}"` : ""} /></label>`;
  }

  function editorTextarea(label, path, value, placeholder = "", rows = 3, wide = false, target = "") {
    return `<label class="field ${wide ? "wide" : ""}" ${target ? `id="field-${escapeHtml(target)}"` : ""}>${escapeHtml(label)}<textarea data-path="${escapeHtml(path)}" rows="${rows}" placeholder="${escapeHtml(placeholder)}">${escapeHtml(value || "")}</textarea></label>`;
  }

  function phaseInput(label, phase, field, value, placeholder = "", wide = false, list = "") {
    return `<label class="field ${wide ? "wide" : ""}">${escapeHtml(label)}<input data-phase-id="${escapeHtml(phase.id)}" data-phase-field="${field}" value="${escapeHtml(value || "")}" placeholder="${escapeHtml(placeholder)}" ${list ? `list="${list}"` : ""} /></label>`;
  }

  function miniSummary(label, value) { return `<div class="review-item"><span>✓</span><span><strong>${escapeHtml(label)}</strong><small>${escapeHtml(value)}</small></span><em>Repris</em></div>`; }
  function reviewRow(label, value, good) { return `<div class="review-item ${good ? "" : "warn"}"><span>${good ? "✓" : "!"}</span><span><strong>${escapeHtml(label)}</strong><small>${escapeHtml(value)}</small></span><em>${good ? "OK" : "À faire"}</em></div>`; }
  function emptyState(icon, title, text, action = "") { return `<div class="empty-card"><span>${icon}</span><h2>${escapeHtml(title)}</h2><p>${escapeHtml(text)}</p>${action}</div>`; }

  function handleClick(event) {
    const nav = event.target.closest("[data-nav]");
    if (nav) {
      navigate(nav.dataset.nav);
      return;
    }
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const action = button.dataset.action;
    if (action === "new-isf") openNewIsfDialog();
    else if (action === "start-blank") startBlankIsf();
    else if (action === "start-from-chantier") showNewPicker("chantier");
    else if (action === "start-from-isf") showNewPicker("isf");
    else if (action === "confirm-start-chantier") confirmStartFromChantier();
    else if (action === "confirm-start-isf") confirmStartFromIsf();
    else if (action === "open-isf") openIsf(button.dataset.isfId || button.closest("[data-isf-id]")?.dataset.isfId);
    else if (action === "new-from-chantier") createFromChantierId(button.dataset.chantierId);
    else if (action === "duplicate-last-chantier") duplicateLastForChantier(button.dataset.chantierId);
    else if (action === "new-chantier") createNewChantier();
    else if (action === "edit-chantier") editChantier(button.dataset.chantierId);
    else if (action === "close-chantier-edit") closeChantierEdit();
    else if (action === "delete-chantier") deleteChantier(button.dataset.chantierId);
    else if (action === "set-isf-filter") { state.isfFilter = button.dataset.filter; render(); }
    else if (action === "set-library-filter") { state.libraryFilter = button.dataset.filter; render(); }
    else if (action === "add-library-item") openLibraryDialog(button.dataset.type || state.libraryFilter);
    else if (action === "edit-library-item") openLibraryDialog(null, button.dataset.libraryId);
    else if (action === "delete-library-item") deleteLibraryItem(button.dataset.libraryId);
    else if (action === "open-contact-dialog") openContactDialog(button.dataset.role || "", button.dataset.contactTarget || "");
    else if (action === "close-editor") closeEditor();
    else if (action === "go-step") goStep(Number(button.dataset.step));
    else if (action === "previous-step") goStep(Math.max(1, state.editorStep - 1));
    else if (action === "next-step") nextStep();
    else if (action === "save-current-chantier") saveCurrentAsChantier();
    else if (action === "add-representative") addRepresentative();
    else if (action === "remove-representative") removeRepresentative(Number(button.dataset.index));
    else if (action === "add-perimeter") addPerimeter();
    else if (action === "duplicate-perimeter") duplicatePerimeter(Number(button.dataset.index));
    else if (action === "remove-perimeter") removePerimeter(Number(button.dataset.index));
    else if (action === "parse-perimeters") parseBulkPerimeters();
    else if (action === "add-phase") addPhase();
    else if (action === "add-phase-set") addPhaseFromSet(button.dataset.libraryId);
    else if (action === "add-phase-template") addPhaseFromTemplate(button.dataset.libraryId);
    else if (action === "toggle-phase") togglePhase(button.dataset.phaseId);
    else if (action === "duplicate-phase") duplicatePhase(button.dataset.phaseId);
    else if (action === "remove-phase") removePhase(button.dataset.phaseId);
    else if (action === "add-activity") addActivity(button.dataset.phaseId);
    else if (action === "add-suggested-activity") addSuggestedActivity(button.dataset.phaseId, button.dataset.label);
    else if (action === "remove-activity") removeActivity(button.dataset.phaseId, Number(button.dataset.index));
    else if (action === "add-tracks") addTracks(button.dataset.phaseId);
    else if (action === "remove-track") removeTrack(button.dataset.phaseId, button.dataset.trackId);
    else if (action === "copy-previous-track") copyPreviousTrack(button.dataset.phaseId, button.dataset.trackId);
    else if (action === "apply-first-track") applyFirstTrack(button.dataset.phaseId);
    else if (action === "clear-phase-protections") clearPhaseProtections(button.dataset.phaseId);
    else if (action === "add-observation") addObservation(button.dataset.phaseId, button.dataset.label);
    else if (action === "save-phase-template") savePhaseTemplate(button.dataset.phaseId);
    else if (action === "jump-anomaly") jumpToAnomaly(Number(button.dataset.step), button.dataset.target);
    else if (action === "generate-word") generateWord();
    else if (action === "generate-draft") generateWord("draft");
    else if (action === "duplicate-current") duplicateCurrentIsf();
    else if (action === "archive-current") toggleArchiveCurrent();
    else if (action === "export-current-json") exportCurrentJson();
    else if (action === "export-backup") exportBackup();
    else if (action === "import-backup") els.backupFileInput.click();
    else if (action === "clear-all-data") clearAllData();
    else if (action === "check-update") checkForUpdate();
  }

  function handleMainInput(event) {
    const target = event.target;
    if (target.id === "isfSearch") {
      state.filters.isfSearch = target.value;
      const result = document.getElementById("isfResults");
      if (result) result.innerHTML = renderIsfResults();
      return;
    }
    if (target.id === "chantierSearch") {
      state.filters.chantierSearch = target.value;
      const result = document.getElementById("chantierResults");
      if (result) result.innerHTML = renderChantierResults();
      return;
    }
    if (target.id === "librarySearch") {
      state.filters.librarySearch = target.value;
      const result = document.getElementById("libraryResults");
      if (result) result.innerHTML = renderLibraryResults();
      return;
    }
    if (target.dataset.chantierPath && state.editingChantier) {
      setByPath(state.editingChantier, target.dataset.chantierPath, inputValue(target));
      saveEditingChantier();
      return;
    }
    if (!state.currentIsf) return;
    if (target.dataset.path) {
      setByPath(state.currentIsf, target.dataset.path, inputValue(target));
      if (["index", "revisionDate", "modification"].includes(target.dataset.path)) syncCurrentRevision();
      if (target.dataset.path === "protectionsVerified" && target.checked) {
        state.currentIsf.protectionsVerifiedAt = new Date().toISOString();
      }
      scheduleCurrentSave();
    }
    if (target.dataset.phaseField) {
      const phase = findPhase(target.dataset.phaseId);
      if (phase) {
        phase[target.dataset.phaseField] = inputValue(target);
        scheduleCurrentSave();
      }
    }
    if (target.dataset.trackField) {
      const track = findTrack(target.dataset.phaseId, target.dataset.trackId);
      if (track) {
        track[target.dataset.trackField] = target.value;
        scheduleCurrentSave();
      }
    }
    if (target.dataset.perimeterField) {
      const item = state.currentIsf.perimeters[Number(target.dataset.perimeterIndex)];
      if (item) {
        item[target.dataset.perimeterField] = target.value;
        scheduleCurrentSave();
      }
    }
  }

  function handleMainChange(event) {
    const target = event.target;
    if (target.dataset.actionChange === "apply-chantier") {
      applyChantierToCurrent(target.value);
      return;
    }
    if (target.dataset.contactRole && state.currentIsf) {
      syncValidationParticipants(state.currentIsf);
      scheduleCurrentSave();
    }
    if (target.dataset.representativeIndex && state.currentIsf) {
      state.currentIsf.contacts.representativeIds[Number(target.dataset.representativeIndex)] = target.value;
      syncValidationParticipants(state.currentIsf);
      scheduleCurrentSave();
    }
    if (target.dataset.protection) {
      const track = findTrack(target.dataset.phaseId, target.dataset.trackId);
      if (track) {
        track.protections[target.dataset.protection] = target.checked;
        invalidateSafety("Protection modifiée manuellement — à vérifier");
        scheduleCurrentSave();
        render();
      }
    }
    if (target.dataset.validationField) {
      const participant = state.currentIsf.validationParticipants[Number(target.dataset.validationIndex)];
      if (participant) {
        participant[target.dataset.validationField] = inputValue(target);
        scheduleCurrentSave();
      }
    }
    if (target.dataset.path === "protectionsVerified") render();
  }

  function navigate(view) {
    flushCurrentSave();
    state.view = view;
    state.editingChantier = null;
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function openNewIsfDialog() {
    els.newIsfPicker.hidden = true;
    els.newIsfPicker.innerHTML = "";
    if (!els.newIsfDialog.open) els.newIsfDialog.showModal();
  }

  function showNewPicker(type) {
    if (type === "chantier") {
      if (!state.chantiers.length) {
        els.newIsfDialog.close();
        startBlankIsf();
        showToast("Aucun chantier enregistré : crée le premier pendant cette ISF.", "warn");
        return;
      }
      els.newIsfPicker.innerHTML = `<label>Chantier<select id="newChantierSelect">${state.chantiers.map((item) => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.name || item.operation || item.location)}</option>`).join("")}</select></label><button class="primary-button" type="button" data-action="confirm-start-chantier">Continuer avec ce chantier</button>`;
    } else {
      if (!state.isfs.length) {
        showToast("Aucune ISF existante à dupliquer.", "warn");
        return;
      }
      els.newIsfPicker.innerHTML = `<label>ISF à dupliquer<select id="newBaseIsfSelect">${state.isfs.map((item) => `<option value="${escapeHtml(item.id)}">${escapeHtml(`ISF ${isfReference(item)} · IND ${item.index} · ${item.chantierName || item.operation || "Sans chantier"}`)}</option>`).join("")}</select></label><button class="primary-button" type="button" data-action="confirm-start-isf">Dupliquer et modifier</button>`;
    }
    els.newIsfPicker.hidden = false;
  }

  function confirmStartFromChantier() {
    const id = document.getElementById("newChantierSelect")?.value;
    if (id) createFromChantierId(id);
  }

  function confirmStartFromIsf() {
    const id = document.getElementById("newBaseIsfSelect")?.value;
    const source = state.isfs.find((item) => item.id === id);
    if (source) {
      els.newIsfDialog.close();
      createDuplicate(source);
    }
  }

  function startBlankIsf() {
    if (els.newIsfDialog.open) els.newIsfDialog.close();
    openNewDraft(createBlankIsf());
  }

  function createFromChantierId(id) {
    const chantier = state.chantiers.find((item) => item.id === id);
    if (!chantier) return;
    if (els.newIsfDialog.open) els.newIsfDialog.close();
    openNewDraft(createBlankIsf(chantier));
  }

  function duplicateLastForChantier(id) {
    const source = lastIsfForChantier(id);
    if (source) createDuplicate(source);
    else createFromChantierId(id);
  }

  async function openNewDraft(isf) {
    state.isfs.unshift(isf);
    await ISFStorage.put("isfs", isf);
    state.currentIsf = clone(isf);
    state.view = "editor";
    state.editorStep = 1;
    render();
    window.scrollTo(0, 0);
  }

  function openIsf(id) {
    const isf = state.isfs.find((item) => item.id === id);
    if (!isf) return;
    state.currentIsf = normalizeIsfModel(clone(isf));
    state.view = "editor";
    state.editorStep = isf.lastEditorStep || 1;
    render();
    window.scrollTo(0, 0);
  }

  async function closeEditor() {
    await flushCurrentSave();
    state.currentIsf = null;
    navigate("dashboard");
  }

  function goStep(step) {
    if (!state.currentIsf || step < 1 || step > 5) return;
    state.editorStep = step;
    state.currentIsf.lastEditorStep = step;
    scheduleCurrentSave();
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function nextStep() {
    if (state.editorStep === 5) {
      closeEditor();
      return;
    }
    goStep(state.editorStep + 1);
  }

  function createBlankIsf(chantier = null) {
    const now = Date.now();
    const today = todayIso();
    const base = chantier || {};
    return normalizeIsfModel({
      id: createId("isf"),
      status: "draft",
      sourceIsfId: "",
      chantierId: base.id || "",
      chantierName: base.name || "",
      series: referenceSeries(base.csfReference) || "194",
      number: "",
      year: String(new Date().getFullYear()),
      index: "0",
      csfReference: base.csfReference || "",
      operation: base.operation || "",
      line: base.line || "",
      location: base.location || "",
      workDescription: base.workDescription || "",
      startDate: "",
      endDate: "",
      weeks: "",
      revisionDate: today,
      modification: "Initialisation",
      revisions: [{ index: "0", date: today, change: "Initialisation" }],
      contacts: clone(base.contacts || { moeId: "", rsoId: "", aspId: "", representativeIds: [] }),
      organization: clone(base.organization || defaultOrganization()),
      perimeters: clone(base.perimeters || []),
      geographicAnnex: null,
      phases: [],
      protectionsVerified: false,
      protectionsVerifiedAt: "",
      protectionsVerifiedBy: "",
      validationParticipants: [],
      rsoComments: "",
      generationMode: "final",
      createdAt: now,
      updatedAt: now,
      lastEditorStep: 1,
    });
  }

  function normalizeIsfModel(isf) {
    isf.contacts = { moeId: "", rsoId: "", aspId: "", representativeIds: [], ...(isf.contacts || {}) };
    if (!Array.isArray(isf.contacts.representativeIds)) isf.contacts.representativeIds = [];
    isf.organization = { ...defaultOrganization(), ...(isf.organization || {}) };
    if (!Array.isArray(isf.organization.questions)) isf.organization.questions = [null, null, null, null];
    isf.perimeters = Array.isArray(isf.perimeters) ? isf.perimeters : [];
    isf.phases = Array.isArray(isf.phases) ? isf.phases.map((phase,index) => normalizePhase(phase,index)) : [];
    isf.revisions = Array.isArray(isf.revisions) ? isf.revisions : [];
    isf.validationParticipants = Array.isArray(isf.validationParticipants) ? isf.validationParticipants : [];
    return isf;
  }

  function defaultOrganization() {
    return { meetingPlace: "", meetingDayTime: "", meetingNightTime: "", careBy: "rso", dayHours: "", nightHours: "", staffMin: "", staffMax: "", questions: [null, null, null, null] };
  }

  function normalizePhase(phase = {}, index = 0) {
    return {
      id: phase.id || createId("phase"),
      code: phase.code || `A${index + 1}`,
      title: phase.title || "Nouvelle phase",
      period: ["day","night","both"].includes(phase.period) ? phase.period : "day",
      zone: phase.zone || "",
      company: phase.company || "",
      observations: phase.observations || "",
      activities: Array.isArray(phase.activities) ? phase.activities.map((item) => typeof item === "string" ? item : item?.label).filter(Boolean) : [],
      sourceLabel: phase.sourceLabel || "Créée manuellement — protections à définir",
      open: phase.open !== false,
      tracks: Array.isArray(phase.tracks) ? phase.tracks.map((track,index2) => ({ id: track.id || createId("track"), name: track.name || `Voie ${index2 + 1}`, protections: normalizeProtections(track.protections) })) : [],
    };
  }

  function normalizeProtections(value = {}) {
    return PROTECTIONS.reduce((result,[key]) => ({ ...result, [key]: Boolean(value?.[key]) }), {});
  }

  function addPhase() {
    const phase = normalizePhase({ code: `A${state.currentIsf.phases.length + 1}`, title: "Nouvelle phase", tracks: [] }, state.currentIsf.phases.length);
    closePhases();
    state.currentIsf.phases.push(phase);
    invalidateSafety("Nouvelle phase — protections à définir");
    scheduleCurrentSave();
    render();
    scrollSoon(`phase-${phase.id}`);
  }

  function addPhaseFromSet(id) {
    const set = state.library.find((item) => item.id === id && item.type === "activitySet");
    if (!set) return;
    const phase = normalizePhase({ code: `A${state.currentIsf.phases.length + 1}`, title: set.label, activities: clone(set.activities || []), sourceLabel: `Activités reprises de l’ensemble « ${set.label} » — protections à définir`, tracks: [] }, state.currentIsf.phases.length);
    closePhases();
    state.currentIsf.phases.push(phase);
    invalidateSafety(phase.sourceLabel);
    scheduleCurrentSave();
    render();
    scrollSoon(`phase-${phase.id}`);
  }

  function addPhaseFromTemplate(id) {
    const template = state.library.find((item) => item.id === id && item.type === "phaseTemplate");
    if (!template?.data) return;
    const phase = normalizePhase({ ...clone(template.data), id: createId("phase"), code: `A${state.currentIsf.phases.length + 1}`, open: true, sourceLabel: `Reprise du modèle « ${template.label} » — protections à vérifier`, tracks: (template.data.tracks || []).map((track) => ({ ...clone(track), id: createId("track") })) }, state.currentIsf.phases.length);
    closePhases();
    state.currentIsf.phases.push(phase);
    invalidateSafety(phase.sourceLabel);
    scheduleCurrentSave();
    render();
    scrollSoon(`phase-${phase.id}`);
  }

  function togglePhase(id) {
    const phase = findPhase(id);
    if (!phase) return;
    const open = !phase.open;
    closePhases();
    phase.open = open;
    scheduleCurrentSave();
    render();
  }

  function duplicatePhase(id) {
    const source = findPhase(id);
    if (!source) return;
    closePhases();
    const copy = normalizePhase({ ...clone(source), id: createId("phase"), code: `${source.code || "A"} bis`, period: source.period === "day" ? "night" : source.period, open: true, sourceLabel: "Reprise de la phase précédente — protections à vérifier", tracks: source.tracks.map((track) => ({ ...clone(track), id: createId("track") })) }, state.currentIsf.phases.length);
    state.currentIsf.phases.splice(state.currentIsf.phases.indexOf(source) + 1, 0, copy);
    invalidateSafety(copy.sourceLabel);
    scheduleCurrentSave();
    render();
    scrollSoon(`phase-${copy.id}`);
  }

  function removePhase(id) {
    state.currentIsf.phases = state.currentIsf.phases.filter((phase) => phase.id !== id);
    invalidateSafety("Phase supprimée — protections à revérifier");
    scheduleCurrentSave();
    render();
    showToast("Phase supprimée.", "warn");
  }

  function addActivity(phaseId) {
    const input = document.querySelector(`[data-new-activity="${cssEscape(phaseId)}"]`);
    addSuggestedActivity(phaseId, input?.value || "");
  }

  function addSuggestedActivity(phaseId, label) {
    const phase = findPhase(phaseId);
    const clean = normalizeText(label);
    if (!phase || !clean) return;
    if (!phase.activities.some((item) => normalizeSearch(item) === normalizeSearch(clean))) phase.activities.push(clean);
    scheduleCurrentSave();
    render();
  }

  function removeActivity(phaseId, index) {
    const phase = findPhase(phaseId);
    if (!phase) return;
    phase.activities.splice(index, 1);
    scheduleCurrentSave();
    render();
  }

  function addTracks(phaseId) {
    const phase = findPhase(phaseId);
    const input = document.querySelector(`[data-new-tracks="${cssEscape(phaseId)}"]`);
    if (!phase || !input) return;
    const known = new Set(phase.tracks.map((track) => normalizeSearch(track.name)));
    const names = input.value.split(/[,;\n]+/).map(normalizeText).filter((name) => {
      const key = normalizeSearch(name);
      if (!key || known.has(key)) return false;
      known.add(key);
      return true;
    });
    if (!names.length) { input.focus(); return; }
    names.forEach((name) => phase.tracks.push({ id: createId("track"), name, protections: normalizeProtections() }));
    invalidateSafety("Voies ajoutées — protections à définir");
    scheduleCurrentSave();
    render();
  }

  function removeTrack(phaseId, trackId) {
    const phase = findPhase(phaseId);
    if (!phase) return;
    phase.tracks = phase.tracks.filter((track) => track.id !== trackId);
    invalidateSafety("Voie supprimée — protections à revérifier");
    scheduleCurrentSave();
    render();
  }

  function copyPreviousTrack(phaseId, trackId) {
    const phase = findPhase(phaseId);
    const index = phase?.tracks.findIndex((track) => track.id === trackId) ?? -1;
    if (!phase || index < 1) return;
    phase.tracks[index].protections = clone(phase.tracks[index - 1].protections);
    phase.sourceLabel = "Protections reprises depuis la voie précédente — à vérifier";
    invalidateSafety(phase.sourceLabel);
    scheduleCurrentSave();
    render();
  }

  function applyFirstTrack(phaseId) {
    const phase = findPhase(phaseId);
    if (!phase || phase.tracks.length < 2) return;
    const source = clone(phase.tracks[0].protections);
    phase.tracks.slice(1).forEach((track) => { track.protections = clone(source); });
    phase.sourceLabel = "Configuration de la première voie recopiée — protections à vérifier";
    invalidateSafety(phase.sourceLabel);
    scheduleCurrentSave();
    render();
  }

  function clearPhaseProtections(phaseId) {
    const phase = findPhase(phaseId);
    if (!phase) return;
    phase.tracks.forEach((track) => { track.protections = normalizeProtections(); });
    phase.sourceLabel = "Protections désactivées manuellement — à vérifier";
    invalidateSafety(phase.sourceLabel);
    scheduleCurrentSave();
    render();
  }

  function addObservation(phaseId, text) {
    const phase = findPhase(phaseId);
    if (!phase || !normalizeText(text)) return;
    const existing = normalizeText(phase.observations);
    if (!normalizeSearch(existing).includes(normalizeSearch(text))) phase.observations = [existing, text].filter(Boolean).join("\n");
    scheduleCurrentSave();
    render();
  }

  async function savePhaseTemplate(phaseId) {
    const phase = findPhase(phaseId);
    if (!phase) return;
    const item = { id: createId("library"), type: "phaseTemplate", label: phase.title || phase.code || "Phase", data: clone(phase), createdAt: Date.now() };
    await ISFStorage.put("library", item);
    state.library.push(item);
    sortCollections();
    showToast("Phase enregistrée dans la bibliothèque. Les protections devront toujours être vérifiées.", "success");
  }

  function addPerimeter() {
    const center = normalizeText(document.getElementById("newPerimeterCenter")?.value);
    const pk = normalizeText(document.getElementById("newPerimeterPk")?.value);
    if (!center && !pk) return;
    state.currentIsf.perimeters.push({ id: createId("perimeter"), center, pk });
    scheduleCurrentSave();
    render();
  }

  function duplicatePerimeter(index) {
    const source = state.currentIsf.perimeters[index];
    if (!source) return;
    state.currentIsf.perimeters.splice(index + 1, 0, { ...clone(source), id: createId("perimeter") });
    scheduleCurrentSave();
    render();
  }

  function removePerimeter(index) {
    state.currentIsf.perimeters.splice(index, 1);
    scheduleCurrentSave();
    render();
  }

  function parseBulkPerimeters() {
    const value = document.getElementById("bulkPerimeter")?.value || "";
    const entries = value.split(/\r?\n/).map((line) => line.trim()).filter(Boolean).map((line) => {
      const parts = line.split(/\s*[;\t]\s*|\s{2,}/).map(normalizeText);
      if (parts.length === 1) {
        const match = /^(.*?)(\d+[,.]\d+|\d+)\s*$/.exec(parts[0]);
        return match ? { center: normalizeText(match[1]), pk: normalizeText(match[2]) } : { center: parts[0], pk: "" };
      }
      return { center: parts[0], pk: parts.slice(1).join(" ") };
    });
    if (!entries.length) return;
    const known = new Set(state.currentIsf.perimeters.map((item) => normalizeSearch(`${item.center}|${item.pk}`)));
    let added = 0;
    entries.forEach((entry) => {
      const key = normalizeSearch(`${entry.center}|${entry.pk}`);
      if (!known.has(key)) {
        state.currentIsf.perimeters.push({ id: createId("perimeter"), ...entry });
        known.add(key);
        added += 1;
      }
    });
    scheduleCurrentSave();
    render();
    showToast(`${added} zone${plural(added)} ajoutée${plural(added)} au périmètre.`, "success");
  }

  function addRepresentative() {
    state.contactTarget = "isf:representative";
    const reps = getContacts().filter((item) => item.role === "representative");
    if (reps.length) {
      const unused = reps.find((item) => !state.currentIsf.contacts.representativeIds.includes(item.id));
      state.currentIsf.contacts.representativeIds.push(unused?.id || "");
      scheduleCurrentSave();
      render();
    } else {
      openContactDialog("representative", "isf:representative");
    }
  }

  function removeRepresentative(index) {
    state.currentIsf.contacts.representativeIds.splice(index, 1);
    syncValidationParticipants(state.currentIsf);
    scheduleCurrentSave();
    render();
  }

  function applyChantierToCurrent(id) {
    const chantier = state.chantiers.find((item) => item.id === id);
    if (!chantier || !state.currentIsf) return;
    const isf = state.currentIsf;
    isf.chantierId = chantier.id;
    isf.chantierName = chantier.name || "";
    isf.line = chantier.line || "";
    isf.csfReference = chantier.csfReference || "";
    isf.series = referenceSeries(chantier.csfReference) || isf.series;
    isf.operation = chantier.operation || "";
    isf.location = chantier.location || "";
    isf.workDescription = chantier.workDescription || isf.workDescription;
    isf.contacts = clone(chantier.contacts || isf.contacts);
    isf.organization = { ...defaultOrganization(), ...clone(chantier.organization || {}) };
    isf.perimeters = clone(chantier.perimeters || []);
    syncValidationParticipants(isf);
    scheduleCurrentSave();
    render();
    showToast("Le chantier a prérempli l’ISF.", "success");
  }

  async function saveCurrentAsChantier() {
    const isf = state.currentIsf;
    if (!isf) return;
    let chantier = state.chantiers.find((item) => item.id === isf.chantierId);
    if (!chantier) {
      chantier = { id: createId("chantier"), createdAt: Date.now() };
      state.chantiers.push(chantier);
      isf.chantierId = chantier.id;
    }
    Object.assign(chantier, {
      name: isf.chantierName || isf.operation || isf.location || "Chantier sans nom",
      line: isf.line || "",
      csfReference: isf.csfReference || "",
      operation: isf.operation || "",
      location: isf.location || "",
      workDescription: isf.workDescription || "",
      contacts: clone(isf.contacts),
      organization: clone(isf.organization),
      perimeters: clone(isf.perimeters),
      lastIsfId: isf.id,
      updatedAt: Date.now(),
    });
    isf.chantierName = chantier.name;
    await ISFStorage.put("chantiers", chantier);
    scheduleCurrentSave();
    sortCollections();
    showToast("Le chantier est enregistré pour les prochaines ISF.", "success");
  }

  function createNewChantier() {
    const chantier = { id: createId("chantier"), name: "", line: "", csfReference: "", operation: "", location: "", workDescription: "", contacts: { moeId: "", rsoId: "", aspId: "", representativeIds: [] }, organization: defaultOrganization(), perimeters: [], createdAt: Date.now(), updatedAt: Date.now() };
    state.editingChantier = chantier;
    state.chantiers.push(chantier);
    state.view = "chantier-edit";
    ISFStorage.put("chantiers", chantier);
    render();
  }

  function editChantier(id) {
    const chantier = state.chantiers.find((item) => item.id === id);
    if (!chantier) return;
    state.editingChantier = clone(chantier);
    state.editingChantier.organization = { ...defaultOrganization(), ...(state.editingChantier.organization || {}) };
    state.editingChantier.contacts = { moeId: "", rsoId: "", aspId: "", representativeIds: [], ...(state.editingChantier.contacts || {}) };
    state.view = "chantier-edit";
    render();
  }

  async function saveEditingChantier() {
    if (!state.editingChantier) return;
    state.editingChantier.updatedAt = Date.now();
    setSaveStatus(true);
    await ISFStorage.put("chantiers", state.editingChantier);
    const index = state.chantiers.findIndex((item) => item.id === state.editingChantier.id);
    if (index >= 0) state.chantiers[index] = clone(state.editingChantier);
    setSaveStatus(false);
  }

  async function closeChantierEdit() {
    await saveEditingChantier();
    state.editingChantier = null;
    navigate("chantiers");
  }

  async function deleteChantier(id) {
    const chantier = state.chantiers.find((item) => item.id === id);
    if (!chantier) return;
    const confirmed = await confirmAction("Supprimer ce chantier", `Supprimer « ${chantier.name || "Chantier sans nom"} » ? Les ISF déjà créées restent conservées.`, "Supprimer");
    if (!confirmed) return;
    await ISFStorage.remove("chantiers", id);
    state.chantiers = state.chantiers.filter((item) => item.id !== id);
    render();
    showToast("Chantier supprimé.", "success");
  }

  function openContactDialog(role = "", target = "", existingId = "") {
    state.contactTarget = target;
    const contact = getContact(existingId);
    els.contactForm.reset();
    els.contactDialogTitle.textContent = contact ? "Modifier l’intervenant" : "Ajouter un intervenant";
    ["id","lastName","firstName","phone","company","function","role"].forEach((key) => { if (els.contactForm.elements[key]) els.contactForm.elements[key].value = contact?.[key] || (key === "role" ? role : ""); });
    if (!els.contactDialog.open) els.contactDialog.showModal();
  }

  async function handleContactSubmit(event) {
    if (event.submitter?.value === "cancel") return;
    event.preventDefault();
    const data = Object.fromEntries(new FormData(els.contactForm));
    if (!normalizeText(data.lastName)) { showToast("Le nom est nécessaire.", "warn"); return; }
    const contact = { id: data.id || createId("contact"), type: "contact", lastName: normalizeText(data.lastName).toUpperCase(), firstName: normalizeText(data.firstName), phone: normalizeText(data.phone), company: normalizeText(data.company), function: normalizeText(data.function), role: data.role || "", updatedAt: Date.now(), createdAt: getContact(data.id)?.createdAt || Date.now() };
    await ISFStorage.put("library", contact);
    const index = state.library.findIndex((item) => item.id === contact.id);
    if (index >= 0) state.library[index] = contact;
    else state.library.push(contact);
    assignCreatedContact(contact.id);
    sortCollections();
    els.contactDialog.close();
    render();
    showToast("Intervenant enregistré et réutilisable.", "success");
  }

  function assignCreatedContact(id) {
    const target = state.contactTarget;
    state.contactTarget = null;
    if (!target) return;
    const [scope,path] = target.split(":");
    if (scope === "isf" && state.currentIsf) {
      if (path === "representative") state.currentIsf.contacts.representativeIds.push(id);
      else setByPath(state.currentIsf, path, id);
      syncValidationParticipants(state.currentIsf);
      scheduleCurrentSave();
    }
    if (scope === "chantier" && state.editingChantier) {
      setByPath(state.editingChantier, path, id);
      saveEditingChantier();
    }
  }

  function openLibraryDialog(type, id = "") {
    const item = state.library.find((entry) => entry.id === id);
    const effectiveType = item?.type || type || "activity";
    if (effectiveType === "contact") {
      openContactDialog(item?.role || "", "", id);
      return;
    }
    if (effectiveType === "phaseTemplate" && !item) {
      showToast("Une phase se sauvegarde directement depuis l’éditeur ISF.", "warn");
      return;
    }
    els.libraryForm.reset();
    els.libraryDialogTitle.textContent = item ? "Modifier l’élément" : `Ajouter — ${libraryTypeLabel(effectiveType)}`;
    els.libraryForm.elements.id.value = item?.id || "";
    els.libraryForm.elements.type.value = effectiveType;
    els.libraryForm.elements.label.value = item?.label || "";
    els.libraryForm.elements.value.value = effectiveType === "activitySet" ? (item?.activities || []).join("\n") : item?.value || "";
    els.libraryValueField.firstChild.textContent = effectiveType === "activitySet" ? "Activités — une par ligne" : effectiveType === "observation" ? "Phrase complète" : "Détail facultatif";
    els.libraryValueField.hidden = ["activity","company","line","csf","place"].includes(effectiveType);
    if (!els.libraryDialog.open) els.libraryDialog.showModal();
  }

  async function handleLibrarySubmit(event) {
    if (event.submitter?.value === "cancel") return;
    event.preventDefault();
    const data = Object.fromEntries(new FormData(els.libraryForm));
    const label = normalizeText(data.label);
    if (!label) { showToast("Le libellé est nécessaire.", "warn"); return; }
    const previous = state.library.find((item) => item.id === data.id);
    const item = { ...previous, id: data.id || createId("library"), type: data.type, label, value: normalizeText(data.value), updatedAt: Date.now(), createdAt: previous?.createdAt || Date.now() };
    if (data.type === "activitySet") item.activities = String(data.value || "").split(/\r?\n/).map(normalizeText).filter(Boolean);
    await ISFStorage.put("library", item);
    const index = state.library.findIndex((entry) => entry.id === item.id);
    if (index >= 0) state.library[index] = item;
    else state.library.push(item);
    sortCollections();
    els.libraryDialog.close();
    render();
    showToast("Bibliothèque mise à jour.", "success");
  }

  async function deleteLibraryItem(id) {
    const item = state.library.find((entry) => entry.id === id);
    if (!item) return;
    const confirmed = await confirmAction("Supprimer de la bibliothèque", `Supprimer « ${libraryLabel(item)} » ? Les ISF existantes ne seront pas modifiées.`, "Supprimer");
    if (!confirmed) return;
    await ISFStorage.remove("library", id);
    state.library = state.library.filter((entry) => entry.id !== id);
    render();
  }

  function syncValidationParticipants(isf) {
    const ids = [isf.contacts.moeId, isf.contacts.rsoId, isf.contacts.aspId, ...isf.contacts.representativeIds].filter(Boolean);
    const unique = [...new Set(ids)];
    const existing = new Map((isf.validationParticipants || []).map((item) => [item.contactId, item]));
    isf.validationParticipants = unique.map((id) => existing.get(id) || { contactId: id, included: true, date: "" });
  }

  function collectAnomalies(isf) {
    const anomalies = [];
    const add = (condition, key, label, step, target) => { if (!condition) anomalies.push({ key, label, step, target }); };
    add(normalizeText(isf.number), "number", "Numéro ISF à compléter", 2, "field-reference");
    add(normalizeText(isf.year), "year", "Année de l’ISF à compléter", 2, "field-reference");
    add(normalizeText(isf.chantierName || isf.operation), "chantier", "Chantier ou opération à préciser", 1, "field-chantier");
    add(normalizeText(isf.csfReference), "csf", "Référence CSF à compléter", 1, "field-csfReference");
    add(normalizeText(isf.line), "line", "Ligne ferroviaire à compléter", 1, "field-line");
    add(normalizeText(isf.startDate) && normalizeText(isf.endDate), "dates", "Dates de début et de fin à compléter", 2, "field-reference");
    add(!isf.startDate || !isf.endDate || isf.endDate >= isf.startDate, "dateOrder", "La date de fin précède la date de début", 2, "field-reference");
    add(Boolean(isf.contacts.moeId), "moe", "MOE Tx à sélectionner", 2, "field-contacts");
    add(Boolean(isf.contacts.rsoId), "rso", "RSO à sélectionner", 2, "field-contacts");
    add(Boolean(isf.contacts.aspId), "asp", "ASP à sélectionner", 2, "field-contacts");
    const asp = getContact(isf.contacts.aspId);
    add(!isf.contacts.aspId || normalizeText(asp?.phone), "aspPhone", "Téléphone ASP manquant", 2, "field-contacts");
    add(isf.perimeters.length > 0, "perimeter", "Ajouter au moins une zone au périmètre", 2, "field-perimeter");
    isf.perimeters.forEach((item,index) => add(normalizeText(item.center) && normalizeText(item.pk), `perimeter-${index}`, `Centre ou PK manquant dans la zone ${index + 1}`, 2, "field-perimeter"));
    add(isf.phases.length > 0, "phases", "Ajouter au moins une phase d’activité", 3, "");
    isf.phases.forEach((phase,index) => {
      add(normalizeText(phase.title) && normalizeSearch(phase.title) !== "nouvelle phase", `phase-title-${index}`, `Intitulé à préciser dans la phase ${index + 1}`, 3, `phase-${phase.id}`);
      add(phase.activities.length > 0, `phase-activities-${index}`, `Aucune activité dans la phase ${index + 1}`, 3, `phase-${phase.id}`);
      add(phase.tracks.length > 0, `phase-tracks-${index}`, `Aucune voie dans la phase ${index + 1}`, 3, `phase-${phase.id}`);
      add(normalizeText(phase.company), `phase-company-${index}`, `Entreprise manquante dans la phase ${index + 1}`, 3, `phase-${phase.id}`);
    });
    add(Boolean(isf.protectionsVerified), "protectionsVerified", "Valider explicitement les mesures de protection", 4, "field-safety");
    return anomalies;
  }

  function getCompleteSteps(isf) {
    const anomalies = collectAnomalies(isf);
    const complete = new Set();
    if (!anomalies.some((item) => item.step === 1)) complete.add(1);
    if (!anomalies.some((item) => item.step === 2)) complete.add(2);
    if (!anomalies.some((item) => item.step === 3)) complete.add(3);
    if (!anomalies.length) complete.add(4);
    if (isf.lastGeneratedAt) complete.add(5);
    return complete;
  }

  function jumpToAnomaly(step, target) {
    goStep(step);
    if (target) scrollSoon(target);
  }

  function invalidateSafety(sourceLabel = "Configuration modifiée — protections à vérifier") {
    if (!state.currentIsf) return;
    state.currentIsf.protectionsVerified = false;
    state.currentIsf.protectionsVerifiedAt = "";
    state.currentIsf.phases.forEach((phase) => {
      if (!phase.sourceLabel || /vérifi/i.test(phase.sourceLabel)) phase.sourceLabel = sourceLabel;
    });
  }

  async function generateWord(forceMode = "") {
    const isf = state.currentIsf;
    if (!isf) return;
    const mode = forceMode || isf.generationMode || "final";
    const anomalies = collectAnomalies(isf);
    if (!isf.protectionsVerified) {
      goStep(4);
      showToast("La vérification explicite des protections est obligatoire.", "warn");
      return;
    }
    if (mode === "final" && anomalies.some((item) => item.key !== "protectionsVerified")) {
      goStep(4);
      showToast("Complète les informations obligatoires avant la version finale.", "warn");
      return;
    }
    const button = document.querySelector('[data-action="generate-word"]');
    if (button) { button.disabled = true; button.textContent = "Génération du Word…"; }
    try {
      syncValidationParticipants(isf);
      const result = await DocxEngine.generate(isf, { contacts: getContacts() }, { mode });
      downloadBlob(result.blob, result.filename);
      isf.lastGeneratedAt = new Date().toISOString();
      isf.lastGeneratedMode = mode;
      await flushCurrentSave();
      showToast(`Le Word ${mode === "draft" ? "brouillon " : ""}est téléchargé.`, "success");
      render();
    } catch (error) {
      console.error(error);
      showToast(error.message || "La génération du Word a échoué. Le brouillon est conservé.", "error");
      if (button) { button.disabled = false; button.textContent = "↓ GÉNÉRER WORD"; }
    }
  }

  function createDuplicate(source) {
    const copy = normalizeIsfModel(clone(source));
    copy.id = createId("isf");
    copy.sourceIsfId = source.id;
    copy.status = "draft";
    copy.index = nextIndex(source.index);
    copy.revisionDate = todayIso();
    copy.modification = "Nouvel indice — à préciser";
    copy.revisions = [...(source.revisions || []), { index: copy.index, date: copy.revisionDate, change: copy.modification }];
    copy.phases = copy.phases.map((phase,index) => normalizePhase({ ...phase, id: createId("phase"), open: index === 0, sourceLabel: `Reprise de l’ISF ${isfReference(source)} IND ${source.index} — protections à vérifier`, tracks: phase.tracks.map((track) => ({ ...track, id: createId("track") })) }, index));
    copy.protectionsVerified = false;
    copy.protectionsVerifiedAt = "";
    copy.protectionsVerifiedBy = "";
    copy.lastGeneratedAt = "";
    copy.lastEditorStep = 2;
    copy.createdAt = Date.now();
    copy.updatedAt = Date.now();
    openNewDraft(copy).then(() => {
      state.editorStep = 2;
      render();
      showToast(`Copie créée. L’indice ${copy.index} est proposé sans écraser l’ancienne version.`, "success");
    });
  }

  function duplicateCurrentIsf() { if (state.currentIsf) createDuplicate(state.currentIsf); }

  async function toggleArchiveCurrent() {
    if (!state.currentIsf) return;
    state.currentIsf.status = state.currentIsf.status === "archived" ? "draft" : "archived";
    await flushCurrentSave();
    render();
    showToast(state.currentIsf.status === "archived" ? "ISF archivée." : "ISF replacée dans les brouillons.", "success");
  }

  function exportCurrentJson() {
    if (!state.currentIsf) return;
    const blob = new Blob([JSON.stringify({ schema: "docu-secu-single-isf", version: 1, exportedAt: new Date().toISOString(), isf: state.currentIsf }, null, 2)], { type: "application/json" });
    downloadBlob(blob, `${DocxEngine.createFilename(state.currentIsf, "final").replace(/\.docx$/i, "")}.json`);
  }

  async function exportBackup() {
    const payload = await ISFStorage.exportAll();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    downloadBlob(blob, `Sauvegarde_Docu_Secu_${todayIso()}.json`);
    showToast("Sauvegarde JSON téléchargée.", "success");
  }

  async function importBackupFile() {
    const file = els.backupFileInput.files?.[0];
    els.backupFileInput.value = "";
    if (!file) return;
    try {
      const payload = JSON.parse(await file.text());
      const confirmed = await confirmAction("Importer la sauvegarde", "Les éléments du fichier seront fusionnés avec les données présentes. Les identifiants identiques seront mis à jour.", "Importer");
      if (!confirmed) return;
      await ISFStorage.importAll(payload, { replace: false });
      await loadData();
      sortCollections();
      render();
      showToast("Sauvegarde importée.", "success");
    } catch (error) {
      console.error(error);
      showToast(error.message || "Le fichier de sauvegarde est invalide.", "error");
    }
  }

  async function clearAllData() {
    const confirmed = await confirmAction("Effacer toutes les données", "Toutes les ISF, les chantiers et les éléments ajoutés seront supprimés de cet appareil. Exporte une sauvegarde avant si nécessaire.", "Tout effacer");
    if (!confirmed) return;
    for (const store of ISFStorage.stores) await ISFStorage.clear(store);
    localStorage.removeItem(LEGACY_KEY);
    state.isfs = [];
    state.chantiers = [];
    state.library = [];
    state.settings = {};
    await seedLibrary();
    render();
    showToast("Données locales effacées. Les favoris issus de la trame ont été restaurés.", "success");
  }

  function scheduleCurrentSave() {
    if (!state.currentIsf) return;
    state.currentIsf.updatedAt = Date.now();
    setSaveStatus(true);
    window.clearTimeout(state.autosaveTimer);
    state.autosaveTimer = window.setTimeout(() => flushCurrentSave(), 280);
  }

  async function flushCurrentSave() {
    window.clearTimeout(state.autosaveTimer);
    state.autosaveTimer = null;
    if (!state.currentIsf) return;
    state.currentIsf.updatedAt = Date.now();
    await ISFStorage.put("isfs", state.currentIsf);
    const index = state.isfs.findIndex((item) => item.id === state.currentIsf.id);
    if (index >= 0) state.isfs[index] = clone(state.currentIsf);
    else state.isfs.unshift(clone(state.currentIsf));
    sortCollections();
    setSaveStatus(false);
  }

  function syncCurrentRevision() {
    const isf = state.currentIsf;
    if (!isf) return;
    const key = String(isf.index ?? "0");
    let revision = isf.revisions.find((item) => String(item.index) === key);
    if (!revision) {
      revision = { index: key, date: isf.revisionDate || todayIso(), change: isf.modification || "" };
      isf.revisions.push(revision);
    } else {
      revision.date = isf.revisionDate || revision.date;
      revision.change = isf.modification || revision.change;
    }
  }

  function setSaveStatus(saving) {
    els.saveStatus.classList.toggle("saving", saving);
    els.saveStatus.lastChild.textContent = saving ? "Enregistrement…" : "Sauvegardé";
  }

  function setConnectionStatus() {
    const online = navigator.onLine;
    els.connectionStatus.textContent = online ? "En ligne" : "Hors connexion";
    els.connectionStatus.classList.toggle("offline", !online);
  }

  function registerServiceWorker() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.register("./sw.js").then((registration) => {
      if (registration.waiting) showUpdate(registration.waiting);
      registration.addEventListener("updatefound", () => {
        const worker = registration.installing;
        worker?.addEventListener("statechange", () => {
          if (worker.state === "installed" && navigator.serviceWorker.controller) showUpdate(worker);
        });
      });
      els.updateAppButton.addEventListener("click", () => state.pendingWorker?.postMessage({ type: "SKIP_WAITING" }));
      navigator.serviceWorker.addEventListener("controllerchange", () => location.reload());
    }).catch((error) => console.warn("Service worker indisponible", error));
  }

  function showUpdate(worker) {
    state.pendingWorker = worker;
    els.updateBanner.hidden = false;
  }

  async function checkForUpdate() {
    if (!("serviceWorker" in navigator)) { showToast("Le service worker n’est pas disponible dans ce navigateur.", "warn"); return; }
    const registration = await navigator.serviceWorker.getRegistration();
    await registration?.update();
    showToast(registration?.waiting ? "Une mise à jour est prête." : "L’application est à jour.", "success");
  }

  function confirmAction(title, text, actionLabel) {
    els.confirmTitle.textContent = title;
    els.confirmText.textContent = text;
    els.confirmButton.textContent = actionLabel;
    els.confirmDialog.returnValue = "";
    if (!els.confirmDialog.open) els.confirmDialog.showModal();
    return new Promise((resolve) => { state.confirmResolve = resolve; });
  }

  function showToast(message, type = "") {
    window.clearTimeout(state.toastTimer);
    els.toast.textContent = message;
    els.toast.className = `toast ${type}`;
    els.toast.hidden = false;
    state.toastTimer = window.setTimeout(() => { els.toast.hidden = true; }, 4800);
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

  function findPhase(id) { return state.currentIsf?.phases.find((phase) => phase.id === id); }
  function findTrack(phaseId, trackId) { return findPhase(phaseId)?.tracks.find((track) => track.id === trackId); }
  function closePhases() { state.currentIsf.phases.forEach((phase) => { phase.open = false; }); }
  function countTracks(isf) { return isf.phases.reduce((total, phase) => total + phase.tracks.length, 0); }
  function countSelectedContacts(isf) { return [isf.contacts.moeId, isf.contacts.rsoId, isf.contacts.aspId, ...isf.contacts.representativeIds].filter(Boolean).length; }
  function countIsfsForChantier(id) { return state.isfs.filter((item) => item.chantierId === id).length; }
  function lastIsfForChantier(id) { return state.isfs.find((item) => item.chantierId === id) || null; }
  function getContacts() { return state.library.filter((item) => item.type === "contact"); }
  function getContact(id) { return getContacts().find((item) => item.id === id) || null; }
  function contactName(contact) { return contact ? [contact.lastName, contact.firstName].filter(Boolean).join(" ").trim() : ""; }

  function contactOptions(selectedId, role = "") {
    const contacts = getContacts().slice().sort((a,b) => {
      const aRole = a.role === role ? 0 : 1;
      const bRole = b.role === role ? 0 : 1;
      return aRole - bRole || contactName(a).localeCompare(contactName(b), "fr");
    });
    return `<option value="">Choisir…</option>${contacts.map((contact) => `<option value="${escapeHtml(contact.id)}" ${contact.id === selectedId ? "selected" : ""}>${escapeHtml(contactName(contact))}${contact.phone ? ` · ${escapeHtml(contact.phone)}` : ""}</option>`).join("")}`;
  }

  function libraryLabel(item) { return item.type === "contact" ? contactName(item) : item.label || item.value || "Sans libellé"; }
  function librarySubtitle(item) {
    if (item.type === "contact") return [item.function, item.company, item.phone].filter(Boolean).join(" · ") || "Coordonnées à compléter";
    if (item.type === "activitySet") return `${(item.activities || []).length} activités`;
    if (item.type === "phaseTemplate") return `${item.data?.tracks?.length || 0} voies · protections à vérifier à chaque reprise`;
    return item.value || item.source || libraryTypeLabel(item.type);
  }
  function libraryIcon(type) { return ({ contact: "ID", activity: "A", activitySet: "A+", observation: "\"", company: "E", line: "L", csf: "C", place: "PK", phaseTemplate: "P" })[type] || "☆"; }
  function libraryTypeLabel(type) { return LIBRARY_TYPES.find(([key]) => key === type)?.[1] || "Élément"; }

  function inputValue(input) {
    if (input.type === "checkbox") return input.checked;
    if (input.type === "radio") return input.value === "true" ? true : input.value === "false" ? false : input.value;
    return input.value;
  }

  function setByPath(object, path, value) {
    const keys = String(path).split(".");
    let cursor = object;
    for (let index = 0; index < keys.length - 1; index += 1) {
      const key = keys[index];
      const nextKey = keys[index + 1];
      if (cursor[key] == null) cursor[key] = /^\d+$/.test(nextKey) ? [] : {};
      cursor = cursor[key];
    }
    cursor[keys[keys.length - 1]] = value;
  }

  function nextIndex(value) {
    const match = /^(.*?)(\d+)$/.exec(String(value ?? "0"));
    return match ? `${match[1]}${Number(match[2]) + 1}` : `${value || "0"}.1`;
  }
  function referenceSeries(csf) { return normalizeText(csf).split(/[\/-]/)[0]?.trim() || ""; }
  function isfReference(isf) { return `${isf.series || "194"}-${isf.number || "?"}-${isf.year || "?"}`; }
  function periodLabel(period) { return period === "night" ? "Nuit" : period === "both" ? "Jour et nuit" : "Jour"; }
  function todayIso() { return new Date().toISOString().slice(0,10); }
  function formatDate(value) { const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || ""); return match ? `${match[3]}/${match[2]}/${match[1]}` : value || ""; }
  function formatRelativeDate(value) { if (!value) return "date inconnue"; const days = Math.floor((Date.now() - Number(new Date(value))) / 86400000); return days <= 0 ? "aujourd’hui" : days === 1 ? "hier" : `il y a ${days} jours`; }
  function plural(count) { return Number(count) === 1 ? "" : "s"; }
  function normalizeText(value) { return String(value ?? "").replace(/\s+/g," ").trim(); }
  function normalizeSearch(value) { return normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase(); }
  function escapeHtml(value) { return String(value ?? "").replace(/[&<>'"]/g,(char) => ({ "&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;" })[char]); }
  function cssEscape(value) { return window.CSS?.escape ? CSS.escape(value) : String(value).replace(/[^a-zA-Z0-9_-]/g,"\\$&"); }
  function createId(prefix) { return `${prefix}-${Date.now().toString(36)}-${crypto.getRandomValues(new Uint32Array(1))[0].toString(36)}`; }
  function clone(value) { return typeof structuredClone === "function" ? structuredClone(value) : JSON.parse(JSON.stringify(value)); }
  function scrollSoon(id) { window.setTimeout(() => document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "center" }), 60); }

  async function saveSetting(id, value) {
    state.settings[id] = value;
    await ISFStorage.put("settings", { id, value });
  }

  function exposeTestHooks() {
    window.__ISF_TEST__ = { createBlankIsf, normalizeIsfModel, collectAnomalies, parsePerimeterText: (text) => text.split(/\r?\n/).filter(Boolean), nextIndex };
  }
})();

/* global JSZip */

(() => {
  "use strict";

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const XML = "http://www.w3.org/XML/1998/namespace";
  const TEMPLATE_URL = "TRAME_ISF.docx";
  const PROTECTIONS = [
    ["consignation", "Consignation\ncaténaire"],
    ["interception", "Interception\ncirculations"],
    ["annonce", "Annonce"],
    ["cloture", "Clôture\nlimitative"],
    ["barriere", "Barrière\ndéfensive"],
  ];

  let templateBufferPromise = null;

  async function generate(isf, context = {}, options = {}) {
    const mode = options.mode === "draft" ? "draft" : "final";
    const draft = mode === "draft";
    const templateBuffer = await getTemplateBuffer(options.templateUrl || TEMPLATE_URL);
    const zip = await JSZip.loadAsync(templateBuffer.slice(0));
    const documentFile = zip.file("word/document.xml");
    if (!documentFile) throw new Error("Le document principal de la trame ISF est introuvable.");

    const doc = parseXml(await documentFile.async("text"), "word/document.xml");
    fillDocument(doc, isf, context, draft);
    if (!draft) removeAllHighlights(doc);
    zip.file("word/document.xml", serializeXml(doc));

    const filename = createFilename(isf, mode);
    const footer = zip.file("word/footer1.xml");
    if (footer) {
      const footerDoc = parseXml(await footer.async("text"), "word/footer1.xml");
      if (!draft) removeAllHighlights(footerDoc);
      zip.file("word/footer1.xml", serializeXml(footerDoc));
    }

    const blob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    return { blob, filename };
  }

  async function getTemplateBuffer(url) {
    if (!templateBufferPromise) {
      templateBufferPromise = fetch(url, { cache: "no-store" }).then(async (response) => {
        if (!response.ok) throw new Error(`Impossible de charger la trame ISF (${response.status}).`);
        return response.arrayBuffer();
      }).catch((error) => {
        templateBufferPromise = null;
        throw error;
      });
    }
    return templateBufferPromise;
  }

  function fillDocument(doc, rawIsf, context, draft) {
    const isf = normalizeIsf(rawIsf);
    const contacts = Array.isArray(context.contacts) ? context.contacts : [];
    const bodyParagraphs = directChildren(doc.getElementsByTagNameNS(W, "body")[0], "p");
    const allParagraphs = Array.from(doc.getElementsByTagNameNS(W, "p"));
    const tables = Array.from(doc.getElementsByTagNameNS(W, "tbl"));

    setMatchingParagraph(bodyParagraphs, (text) => /^n°\s*194\s*\//i.test(text), `N° ${valueOr(isf.series, "194")} / ${requiredOr(isf.number)} / ${requiredOr(isf.year)}`, draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^en référence à la csf n°/i.test(text), `En référence à la CSF n° ${requiredOr(isf.csfReference)}`, draft);
    setFollowingParagraph(bodyParagraphs, /^opération$/i, requiredOr(isf.operation), draft);
    setFollowingParagraph(bodyParagraphs, /^lieu$/i, requiredOr(isf.location || isf.line), draft);
    setFollowingParagraph(bodyParagraphs, /^durée d[’']intervention$/i, formatInterventionPeriod(isf), draft);

    setMatchingParagraph(bodyParagraphs, (text) => /^lieu de rdv\s*:/i.test(text), "Lieu de RDV :", draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^base travaux[,\s]/i.test(text), requiredOr(isf.organization.meetingPlace), draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^heure\s*:/i.test(text), formatMeetingHours(isf.organization), draft);
    setMatchingParagraph(bodyParagraphs, (text) => /sera pris en charge par/i.test(text), `sera pris en charge par : ${careByLabel(isf.organization.careBy)}`, draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^horaires de travail\s*:/i.test(text), "Horaires de travail :", draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^de jour\s/i.test(text), `De jour ${requiredOr(isf.organization.dayHours)}`, draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^de nuit\s/i.test(text), `De nuit ${requiredOr(isf.organization.nightHours)}`, draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^effectif prévisible\s*:/i.test(text), `Effectif prévisible : min : ${requiredOr(isf.organization.staffMin)} / max : ${requiredOr(isf.organization.staffMax)}`, draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^les travaux se trouvent sur la ligne/i.test(text), `Les travaux se trouvent sur ${requiredOr(isf.line)}, ci-dessous la liste des centres concernés`, draft);
    // In the source template, the word "jointe" is split across several Word
    // runs ("joint" + "e"). Match that OOXML reality instead of relying on a
    // contiguous plain-text replacement.
    setMatchingParagraph(allParagraphs, (text) => /^une situation géographique est joint\s*e/i.test(text), `Une situation géographique est jointe en annexe à cette ISF : ${yesNo(isf.geographicAnnex)}`, draft);
    setMatchingParagraph(bodyParagraphs, (text) => /^travaux\s*:/i.test(text), `TRAVAUX : ${requiredOr(isf.workDescription || isf.operation)}`, draft);

    const commentsParagraph = allParagraphs.find((paragraph) => normalizeSearch(paragraphText(paragraph)).startsWith("points positifs / a ameliorer"));
    if (commentsParagraph && normalizeText(isf.rsoComments)) {
      setParagraphText(commentsParagraph, `Points positifs / à améliorer - Gestion des aléas - Incidents / accidents :\n${isf.rsoComments}`, { highlight: draft });
    }

    const contactsTable = tables.find((table) => normalizeSearch(tableText(table)).includes("moe tx a contacter en cas d'urgence"));
    if (contactsTable) fillContactsTable(contactsTable, isf, contacts, draft);

    const revisionTable = tables.find((table) => {
      const text = normalizeSearch(firstRowText(table));
      return text.includes("indice") && text.includes("modifications apportees");
    });
    if (revisionTable) fillRevisionTable(revisionTable, isf, draft);

    const questionsTable = tables.find((table) => normalizeSearch(tableText(table)).includes("le rso (passant gris) reste en permanence"));
    if (questionsTable) fillQuestionsTable(questionsTable, isf.organization.questions, draft);

    const perimeterTable = tables.find((table) => normalizeSearch(firstRowText(table)) === "centre pk");
    if (perimeterTable) fillPerimeterTable(perimeterTable, isf.perimeters, draft);

    const validationTable = tables.find((table) => {
      const text = normalizeSearch(firstRowText(table));
      return text.includes("nom prenom entreprise fonction date et signature");
    });
    if (validationTable) fillValidationTable(validationTable, isf.validationParticipants, contacts, draft);

    injectPhaseTables(doc, isf.phases, isf, draft);
  }

  function normalizeIsf(value = {}) {
    return {
      ...value,
      series: normalizeText(value.series) || "194",
      number: String(value.number || ""),
      year: String(value.year || new Date().getFullYear()),
      index: String(value.index ?? "0"),
      organization: {
        meetingPlace: "",
        meetingDayTime: "",
        meetingNightTime: "",
        careBy: "rso",
        dayHours: "",
        nightHours: "",
        staffMin: "",
        staffMax: "",
        questions: [null, null, null, null],
        ...(value.organization || {}),
      },
      contacts: {
        moeId: "",
        rsoId: "",
        aspId: "",
        representativeIds: [],
        ...(value.contacts || {}),
      },
      perimeters: Array.isArray(value.perimeters) ? value.perimeters : [],
      phases: Array.isArray(value.phases) ? value.phases : [],
      revisions: Array.isArray(value.revisions) ? value.revisions : [],
      validationParticipants: Array.isArray(value.validationParticipants) ? value.validationParticipants : [],
    };
  }

  function fillContactsTable(table, isf, contacts, draft) {
    const rows = directRows(table);
    if (rows.length < 12) return;
    const moe = getContact(contacts, isf.contacts.moeId);
    const rso = getContact(contacts, isf.contacts.rsoId);
    const asp = getContact(contacts, isf.contacts.aspId);
    setCell(rows[1], 1, contactFullName(moe), draft);
    setCell(rows[2], 1, moe?.phone, draft);
    setCell(rows[4], 1, contactFullName(rso), draft);
    setCell(rows[5], 1, rso?.phone, draft);
    setCell(rows[7], 1, contactFullName(asp), draft);
    setCell(rows[8], 1, asp?.phone, draft);

    const representativeIds = Array.isArray(isf.contacts.representativeIds) ? isf.contacts.representativeIds : [];
    const representatives = representativeIds.map((id) => getContact(contacts, id)).filter(Boolean);
    const headerTemplate = rows[9].cloneNode(true);
    const nameTemplate = rows[10].cloneNode(true);
    const phoneTemplate = rows[11].cloneNode(true);
    rows.slice(9).forEach((row) => row.parentNode?.removeChild(row));
    const count = Math.max(1, representatives.length);
    for (let index = 0; index < count; index += 1) {
      const person = representatives[index];
      const header = headerTemplate.cloneNode(true);
      const nameRow = nameTemplate.cloneNode(true);
      const phoneRow = phoneTemplate.cloneNode(true);
      setCell(header, 0, `Représentant Chantier Entreprise ${person?.company || ""}`.trim(), false);
      setCell(nameRow, 1, contactFullName(person), draft);
      setCell(phoneRow, 1, person?.phone, draft);
      table.appendChild(header);
      table.appendChild(nameRow);
      table.appendChild(phoneRow);
    }
  }

  function fillRevisionTable(table, isf, draft) {
    const rows = directRows(table);
    if (rows.length < 2) return;
    const template = rows[1].cloneNode(true);
    rows.slice(1).forEach((row) => row.parentNode?.removeChild(row));
    const revisions = isf.revisions.length ? isf.revisions : [{ index: isf.index, date: isf.revisionDate || todayIso(), change: isf.modification || "Initialisation" }];
    const count = Math.max(2, revisions.length);
    for (let index = 0; index < count; index += 1) {
      const revision = revisions[index];
      const row = template.cloneNode(true);
      setCell(row, 0, revision?.index ?? "", draft && Boolean(revision));
      setCell(row, 1, revision ? formatDate(revision.date) : "", draft && Boolean(revision));
      setCell(row, 2, revision?.change || "", draft && Boolean(revision));
      table.appendChild(row);
    }
  }

  function fillQuestionsTable(table, rawQuestions, draft) {
    const questions = Array.isArray(rawQuestions) ? rawQuestions : [];
    directRows(table).forEach((row, index) => setCell(row, 1, yesNo(questions[index]), draft));
  }

  function fillPerimeterTable(table, rawPerimeters, draft) {
    const rows = directRows(table);
    if (rows.length < 2) return;
    const template = rows[1].cloneNode(true);
    rows.slice(1).forEach((row) => row.parentNode?.removeChild(row));
    const perimeters = (Array.isArray(rawPerimeters) ? rawPerimeters : []).filter((item) => normalizeText(item?.center) || normalizeText(item?.pk));
    const entries = perimeters.length ? perimeters : [{ center: "", pk: "" }];
    entries.forEach((item) => {
      const row = template.cloneNode(true);
      setCell(row, 0, item.center, draft);
      setCell(row, 1, item.pk, draft);
      table.appendChild(row);
    });
  }

  function fillValidationTable(table, rawParticipants, contacts, draft) {
    const rows = directRows(table);
    if (rows.length < 2) return;
    const template = rows[1].cloneNode(true);
    rows.slice(1).forEach((row) => row.parentNode?.removeChild(row));
    const participants = (Array.isArray(rawParticipants) ? rawParticipants : [])
      .filter((item) => item?.included !== false)
      .map((item) => ({ ...getContact(contacts, item.contactId), ...item }));
    const count = Math.max(6, participants.length);
    for (let index = 0; index < count; index += 1) {
      const person = participants[index];
      const row = template.cloneNode(true);
      setCell(row, 0, person?.lastName || "", draft && Boolean(person));
      setCell(row, 1, person?.firstName || "", draft && Boolean(person));
      setCell(row, 2, person?.company || "", draft && Boolean(person));
      setCell(row, 3, person?.function || "", draft && Boolean(person));
      setCell(row, 4, person?.date ? formatDate(person.date) : "", draft && Boolean(person));
      table.appendChild(row);
    }
  }

  function injectPhaseTables(doc, rawPhases, isf, draft) {
    const legacy = Array.from(doc.getElementsByTagNameNS(W, "tbl")).filter(isLegacyPhaseTable);
    if (!legacy.length) throw new Error("Les tableaux de phases d’activité sont introuvables dans la trame ISF.");
    const anchor = legacy[0];
    const parent = anchor.parentNode;
    const phases = rawPhases.map((phase, index) => normalizePhase(phase, index));
    phases.forEach((phase, index) => {
      if (index) parent.insertBefore(createPageBreak(doc), anchor);
      parent.insertBefore(createPhaseTable(doc, phase, isf, draft), anchor);
    });
    legacy.forEach((table) => table.parentNode?.removeChild(table));
  }

  function normalizePhase(value = {}, index = 0) {
    const activities = Array.isArray(value.activities)
      ? value.activities.map((item) => normalizeText(typeof item === "string" ? item : item?.label)).filter(Boolean)
      : String(value.operations || "").split(/\r?\n/).map(normalizeText).filter(Boolean);
    const tracks = Array.isArray(value.tracks) ? value.tracks : [];
    return {
      code: normalizeText(value.code) || `A${index + 1}`,
      title: normalizeText(value.title) || "Phase à préciser",
      period: ["day", "night", "both"].includes(value.period) ? value.period : value.night && value.day ? "both" : value.night ? "night" : "day",
      zone: normalizeText(value.zone),
      company: normalizeText(value.company),
      observations: normalizeText(value.observations),
      weeks: normalizeText(value.weeks),
      activities,
      tracks: tracks.map((track, trackIndex) => ({
        name: normalizeText(track.name) || `Voie ${trackIndex + 1}`,
        protections: PROTECTIONS.reduce((result, [key]) => ({ ...result, [key]: Boolean(track.protections?.[key]) }), {}),
      })),
    };
  }

  function createPhaseTable(doc, phase, isf, draft) {
    const widths = [600, 3500, 900, 800, 800, 800, 800, 800];
    const table = wElement(doc, "tbl");
    const properties = wElement(doc, "tblPr");
    properties.appendChild(wElement(doc, "tblW", { w: sum(widths), type: "dxa" }));
    properties.appendChild(wElement(doc, "tblLayout", { type: "fixed" }));
    properties.appendChild(createTableBorders(doc));
    properties.appendChild(createCellMargins(doc));
    properties.appendChild(wElement(doc, "tblLook", { val: "04A0", firstRow: "1", firstColumn: "0", noHBand: "1", noVBand: "1" }));
    table.appendChild(properties);
    const grid = wElement(doc, "tblGrid");
    widths.forEach((width) => grid.appendChild(wElement(doc, "gridCol", { w: width })));
    table.appendChild(grid);

    const plum = "7D1F4E";
    const fills = ["D9EAF7", "FCE4D6", "FFF2CC", "D9EAD3", "D9D2E9"];
    table.appendChild(createRow(doc, [
      createCell(doc, "N°", { width: widths[0], fill: plum, color: "FFFFFF", bold: true, align: "center", size: 16 }),
      createCell(doc, "OPÉRATION OU PHASE D’ACTIVITÉ", { width: sum(widths, 1, 2), span: 2, fill: plum, color: "FFFFFF", bold: true, align: "center", size: 16 }),
      createCell(doc, "DISPOSITIF DE SÉCURITÉ", { width: sum(widths, 3, 5), span: 5, fill: plum, color: "FFFFFF", bold: true, align: "center", size: 16 }),
    ], true));
    table.appendChild(createRow(doc, [
      createCell(doc, phase.code, { width: widths[0], fill: "E8D8E1", color: plum, bold: true, align: "center", size: 20, highlight: draft }),
      createCell(doc, `${phase.title}\nTravaux réalisés de ${periodLabel(phase.period).toUpperCase()}`, { width: sum(widths, 1, 2), span: 2, fill: "EEE6EC", bold: true, size: 17, highlight: draft }),
      ...PROTECTIONS.map(([, label], index) => createCell(doc, label, { width: widths[index + 3], fill: fills[index], bold: true, align: "center", size: 12 })),
    ], true));

    const tracks = phase.tracks.length ? phase.tracks : [{ name: "Voie à préciser", protections: {} }];
    const activities = phase.activities.length ? phase.activities.join("\n") : "À préciser";
    const zone = phase.zone || isf.location || "À préciser";
    tracks.forEach((track, index) => {
      const details = index === 0 ? `ZONE DE TRAVAIL\n${zone}\n\nACTIVITÉS\n${activities}` : "";
      table.appendChild(createRow(doc, [
        createCell(doc, "", { width: widths[0] }),
        createCell(doc, details, { width: widths[1], fill: index === 0 ? "F8F6F8" : "FFFFFF", size: 16, highlight: draft && index === 0 }),
        createCell(doc, track.name, { width: widths[2], bold: true, align: "center", size: 16, highlight: draft }),
        ...PROTECTIONS.map(([key], protectionIndex) => createCell(doc, track.protections?.[key] ? "X" : "", { width: widths[protectionIndex + 3], fill: track.protections?.[key] ? fills[protectionIndex] : "FFFFFF", bold: true, align: "center", size: 20, highlight: draft })),
      ]));
    });
    const company = phase.company || "À préciser";
    const observationParts = [];
    if (phase.weeks || isf.weeks) observationParts.push(`Semaines concernées : ${phase.weeks || isf.weeks}`);
    if (phase.observations) observationParts.push(phase.observations);
    table.appendChild(createRow(doc, [
      createCell(doc, "", { width: widths[0] }),
      createCell(doc, `ENTREPRISE : ${company}`, { width: sum(widths, 1, 7), span: 7, fill: "F3F0F2", bold: true, size: 16, highlight: draft }),
    ]));
    table.appendChild(createRow(doc, [
      createCell(doc, "", { width: widths[0] }),
      createCell(doc, `OBSERVATIONS : ${observationParts.join("\n") || "À préciser"}`, { width: sum(widths, 1, 7), span: 7, size: 16, highlight: draft }),
    ]));
    return table;
  }

  function createTableBorders(doc) {
    const borders = wElement(doc, "tblBorders");
    ["top", "left", "bottom", "right", "insideH", "insideV"].forEach((edge) => borders.appendChild(wElement(doc, edge, { val: "single", sz: "6", space: "0", color: "6E6570" })));
    return borders;
  }

  function createCellMargins(doc) {
    const margins = wElement(doc, "tblCellMar");
    [["top", 80], ["left", 105], ["bottom", 80], ["right", 105]].forEach(([side, size]) => margins.appendChild(wElement(doc, side, { w: size, type: "dxa" })));
    return margins;
  }

  function createRow(doc, cells, header = false) {
    const row = wElement(doc, "tr");
    if (header) {
      const properties = wElement(doc, "trPr");
      properties.appendChild(wElement(doc, "tblHeader", { val: "true" }));
      row.appendChild(properties);
    }
    cells.forEach((cell) => row.appendChild(cell));
    return row;
  }

  function createCell(doc, text, options = {}) {
    const cell = wElement(doc, "tc");
    const properties = wElement(doc, "tcPr");
    properties.appendChild(wElement(doc, "tcW", { w: options.width || 0, type: "dxa" }));
    if (options.span > 1) properties.appendChild(wElement(doc, "gridSpan", { val: options.span }));
    if (options.fill && options.fill !== "FFFFFF") properties.appendChild(wElement(doc, "shd", { val: "clear", color: "auto", fill: options.fill }));
    properties.appendChild(wElement(doc, "vAlign", { val: "center" }));
    cell.appendChild(properties);
    cell.appendChild(createParagraph(doc, text, options));
    return cell;
  }

  function createParagraph(doc, text, options = {}) {
    const paragraph = wElement(doc, "p");
    const properties = wElement(doc, "pPr");
    properties.appendChild(wElement(doc, "spacing", { before: "0", after: "0", line: "220", lineRule: "auto" }));
    if (options.align) properties.appendChild(wElement(doc, "jc", { val: options.align }));
    paragraph.appendChild(properties);
    const run = wElement(doc, "r");
    const runProperties = wElement(doc, "rPr");
    runProperties.appendChild(wElement(doc, "rFonts", { ascii: "Arial", hAnsi: "Arial", cs: "Arial" }));
    runProperties.appendChild(wElement(doc, "sz", { val: options.size || 16 }));
    if (options.bold) runProperties.appendChild(wElement(doc, "b"));
    if (options.color) runProperties.appendChild(wElement(doc, "color", { val: options.color }));
    if (options.highlight) runProperties.appendChild(wElement(doc, "highlight", { val: "yellow" }));
    run.appendChild(runProperties);
    appendRunText(doc, run, text);
    paragraph.appendChild(run);
    return paragraph;
  }

  function createPageBreak(doc) {
    const paragraph = wElement(doc, "p");
    const run = wElement(doc, "r");
    run.appendChild(wElement(doc, "br", { type: "page" }));
    paragraph.appendChild(run);
    return paragraph;
  }

  function setMatchingParagraph(paragraphs, predicate, text, highlight) {
    const paragraph = paragraphs.find((item) => predicate(normalizeText(paragraphText(item))));
    if (paragraph) setParagraphText(paragraph, text, { highlight });
    return paragraph;
  }

  function setFollowingParagraph(paragraphs, headingPattern, text, highlight) {
    const index = paragraphs.findIndex((paragraph) => headingPattern.test(normalizeText(paragraphText(paragraph))));
    if (index < 0) return null;
    const paragraph = paragraphs.slice(index + 1).find((item) => normalizeText(paragraphText(item)));
    if (paragraph) setParagraphText(paragraph, text, { highlight });
    return paragraph;
  }

  function setParagraphText(paragraph, text, options = {}) {
    const doc = paragraph.ownerDocument;
    const runs = Array.from(paragraph.getElementsByTagNameNS(W, "r"));
    const templateRun = runs[0];
    const run = wElement(doc, "r");
    const sourceProperties = templateRun ? directChildren(templateRun, "rPr")[0] : null;
    const properties = sourceProperties ? sourceProperties.cloneNode(true) : wElement(doc, "rPr");
    removeDirectChildren(properties, "highlight");
    if (options.highlight) properties.appendChild(wElement(doc, "highlight", { val: "yellow" }));
    run.appendChild(properties);
    appendRunText(doc, run, text);
    Array.from(paragraph.childNodes).forEach((child) => {
      if (!(child.nodeType === 1 && child.namespaceURI === W && child.localName === "pPr")) paragraph.removeChild(child);
    });
    paragraph.appendChild(run);
  }

  function appendRunText(doc, run, value) {
    String(value ?? "").replace(/\r\n/g, "\n").split("\n").forEach((line, index) => {
      if (index) run.appendChild(wElement(doc, "br"));
      const text = wElement(doc, "t");
      if (/^\s|\s$/.test(line)) text.setAttributeNS(XML, "xml:space", "preserve");
      text.textContent = line;
      run.appendChild(text);
    });
  }

  function setCell(row, cellIndex, text, highlight = false) {
    const cells = directChildren(row, "tc");
    const cell = cells[cellIndex];
    if (!cell) return;
    let paragraphs = directChildren(cell, "p");
    if (!paragraphs.length) {
      const paragraph = wElement(cell.ownerDocument, "p");
      cell.appendChild(paragraph);
      paragraphs = [paragraph];
    }
    setParagraphText(paragraphs[0], text ?? "", { highlight });
    paragraphs.slice(1).forEach((paragraph) => setParagraphText(paragraph, "", { highlight: false }));
  }

  function updateFooterFilename(doc, filename, draft) {
    const highlightedRuns = Array.from(doc.getElementsByTagNameNS(W, "r")).filter((run) => directChildren(run, "rPr")[0]?.getElementsByTagNameNS(W, "highlight").length);
    if (!highlightedRuns.length) return;
    const first = highlightedRuns[0];
    Array.from(first.childNodes).filter((node) => node.nodeType === 1 && node.namespaceURI === W && ["t", "br"].includes(node.localName)).forEach((node) => first.removeChild(node));
    appendRunText(doc, first, `Diffusable SNCF RESEAU ${filename}`);
    highlightedRuns.slice(1).forEach((run) => Array.from(run.getElementsByTagNameNS(W, "t")).forEach((text) => { text.textContent = ""; }));
    if (!draft) removeAllHighlights(doc);
  }

  function removeAllHighlights(doc) {
    Array.from(doc.getElementsByTagNameNS(W, "highlight")).forEach((node) => node.parentNode?.removeChild(node));
  }

  function isLegacyPhaseTable(table) {
    const text = normalizeSearch(tableText(table));
    return text.includes("mesures de protection") && /phases\s+d.{0,10}activit/.test(text);
  }

  function directRows(table) { return directChildren(table, "tr"); }
  function firstRowText(table) { return tableText(directRows(table)[0]); }
  function tableText(table) { return paragraphText(table); }
  function paragraphText(node) { return Array.from(node?.getElementsByTagNameNS?.(W, "t") || []).map((item) => item.textContent || "").join(" "); }
  function directChildren(node, localName) { return Array.from(node?.childNodes || []).filter((child) => child.nodeType === 1 && child.namespaceURI === W && child.localName === localName); }
  function removeDirectChildren(node, localName) { directChildren(node, localName).forEach((child) => node.removeChild(child)); }

  function wElement(doc, localName, attributes = {}) {
    const element = doc.createElementNS(W, `w:${localName}`);
    Object.entries(attributes).forEach(([name, value]) => element.setAttributeNS(W, `w:${name}`, String(value)));
    return element;
  }

  function parseXml(xml, part) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const parserError = doc.getElementsByTagName("parsererror")[0];
    if (parserError) throw new Error(`Le XML de ${part} est invalide.`);
    return doc;
  }

  function serializeXml(doc) { return new XMLSerializer().serializeToString(doc); }
  function getContact(contacts, id) { return contacts.find((item) => item.id === id) || null; }
  function contactFullName(person) { return person ? [person.lastName, person.firstName].filter(Boolean).join(" ").trim() : ""; }
  function requiredOr(value) { return normalizeText(value) || "À COMPLÉTER"; }
  function valueOr(value, fallback) { return normalizeText(value) || fallback; }
  function yesNo(value) { return value === true ? "OUI" : value === false ? "NON" : "OUI / NON"; }
  function careByLabel(value) { return ({ rso: "RSO", asp: "ASP", autre: "Autre" })[value] || "À COMPLÉTER"; }
  function periodLabel(value) { return value === "night" ? "Nuit" : value === "both" ? "Jour et nuit" : "Jour"; }
  function formatMeetingHours(org) { return `Heure : ${requiredOr(org.meetingDayTime)} de jour et ${requiredOr(org.meetingNightTime)} de nuit`; }
  function formatInterventionPeriod(isf) { return `DU ${requiredOr(formatDate(isf.startDate))} AU ${requiredOr(formatDate(isf.endDate))}`; }
  function formatDate(value) {
    const text = normalizeText(value);
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
    return match ? `${match[3]}/${match[2]}/${match[1]}` : text;
  }
  function todayIso() { return new Date().toISOString().slice(0, 10); }
  function sum(values, start = 0, count = values.length) { return values.slice(start, start + count).reduce((total, value) => total + value, 0); }
  function normalizeText(value) { return String(value ?? "").replace(/\s+/g, " ").trim(); }
  function normalizeSearch(value) { return normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[’]/g, "'").toLowerCase(); }
  function sanitizeFilename(value) {
    const safe = normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
    return safe || "ISF";
  }
  function createFilename(isf, mode) {
    const reference = `${sanitizeFilename(isf.series || "194")}-${sanitizeFilename(isf.number || "brouillon")}-${sanitizeFilename(isf.year || new Date().getFullYear())}`;
    const place = sanitizeFilename(isf.location || isf.line || "chantier");
    const suffix = mode === "draft" ? "_BROUILLON" : "";
    return `ISF_${reference}_IND_${sanitizeFilename(isf.index ?? "0")}_${place}${suffix}.docx`;
  }

  window.DocxEngine = { generate, createFilename, normalizeIsf };
})();

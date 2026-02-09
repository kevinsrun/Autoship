// PUSH TEST: 2026-02-07 11:xxclasp status

// @ts-nocheck
/*************************************************
 * GLOBAL CONFIG
 *************************************************/
 // ===== GLOBAL CONSTANTS (LOAD FIRST) =====
const MAIN_SHEET = "Main Menu";
const RESULTS_SHEET = "Results";
const HEADER_ROW = 1;

function testToast_() {
  SpreadsheetApp.getActive().toast("Autoship is LIVE", "OK", 3);
}

const GEMINI_MODEL = "gemini-2.5-flash"; // model string you chose

// Your Requirements multi-select token that indicates writing
const ESSAY_REQUIREMENT_TOKEN = "Essay(s)";

// Headers (must match your sheet)
const COL_APPLICATION_PORTAL = "Application Portal";
const COL_REQUIREMENTS = "Requirements";
const COL_APPLICATION_FILE = "Application File";
const COL_ADDITIONAL_APPLICATION_FILE = "Additional Application Essay File";

// If Gemini detects PDF upload requirement, we generate a PDF and put it in Additional Application File
const PDF_REQUIRED_KEYWORDS = ["pdf", "upload a pdf", "submit a pdf", "pdf format"];

// Performance tuning
const MAX_HTML_CHARS = 12000; // reduce for speed

/*************************************************
 * MENU
 *************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Scholarship Tools")
    .addItem("Sync Satisfied → Results", "syncSatisfiedToResults")
    .addItem("Sync Surrendered → Surrendered", "syncSurrenderedToSheet")
    .addSeparator()
    .addItem("Refresh Triage (Main Menu)", "refreshTriageForMainMenu")
    .addSeparator()
    .addItem("Create Essay Doc(s) for Selected Rows (FAST)", "createEssayDocsForSelection")
    .addItem("Fill Essay Prompt(s) with AI for Selected Rows (SLOW)", "fillEssayPromptsForSelection")
    .addSeparator()
    .addItem("Delete Essay Docs for Selected Rows", "deleteEssayDocsForSelection")
    .addItem("Debug: Count Triage Values", "debugTriageValues")
    .addToUi();

  ui.createMenu("Scholarship AI")
    .addItem("Open Add Sidebar", "openAddSidebar")
    .addToUi();
}

function openAddSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Add Scholarship (AI)");
  SpreadsheetApp.getUi().showSidebar(html);
}

function aiAddScholarship(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    payload = payload || {};
    const url = (payload.url || "").toString().trim();
    const pastedText = (payload.pastedText || "").toString();
    const extraNotes = (payload.extraNotes || "").toString();

    if (!url && pastedText.trim().length < 20) {
      throw new Error("Provide a URL or paste at least ~20 characters of scholarship text.");
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(MAIN_SHEET);
    if (!sh) throw new Error(`Sheet not found: ${MAIN_SHEET}`);

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const h = buildHeaderIndex(headers);

    // REQUIRED (based on your header row)
    const idxName = mustIndex(h, "Scholarship Name");
    const idxPortal = mustIndex(h, "Application Portal");

    // OPTIONAL (based on your header row)
    const idxDue = optionalIndex_(h, "Due Date");
    const idxReq = optionalIndex_(h, "Requirements");
    const idxNotes = optionalIndex_(h, "Notes");
    const idxTriage = optionalIndex_(h, "Triage");
    const idxStatus = optionalIndex_(h, "Status");
    const idxStart = optionalIndex_(h, "Start date");
    const idxDifficulty = optionalIndex_(h, "Difficulty");
    const idxTheme = optionalIndex_(h, COL_THEME);

    /// Fast, non-hanging extraction
const extracted = basicExtractFromText_(pastedText, url);

// ✅ Create row FIRST
const row = new Array(lastCol).fill("");
row[idxName] = extracted.name || "Scholarship";
row[idxPortal] = url || "";

// Detect + fill Theme + Requirements (now row exists)
if (idxTheme !== null) {
  row[idxTheme] = joinMultiSelect_(detectThemesFromText_(pastedText + "\n" + extraNotes));
}

if (idxReq !== null) {
  // Merge anything extractor already set + our detector
  const detectedReq = detectRequirementsFromText_(pastedText + "\n" + extraNotes);
  const existingReq = (row[idxReq] || "").toString();

  const merged = uniq_(
    existingReq
      ? existingReq.split(/[,;\n]+/).map(s => s.trim()).filter(Boolean).concat(detectedReq)
      : detectedReq
  );
  row[idxReq] = joinMultiSelect_(merged);
}

// Rest of your assignments
if (idxNotes !== null) row[idxNotes] = [extracted.notes, extraNotes].filter(Boolean).join("\n").trim();
if (idxReq !== null && extracted.requirements) {
  const current = (row[idxReq] || "").toString();
  row[idxReq] = current ? joinMultiSelect_(uniq_(current.split(/[,;\n]+/).map(s=>s.trim()).filter(Boolean)
    .concat(extracted.requirements.split(/[,;\n]+/).map(s=>s.trim()).filter(Boolean)))) : extracted.requirements;
}

if (idxDue !== null && extracted.due) row[idxDue] = extracted.due;

    // optional: today
    if (idxStart !== null) row[idxStart] = new Date();

    // --- DROPDOWNS (deterministic, matches your exact options) ---
    if (idxDifficulty !== null) {
      row[idxDifficulty] = computeDifficulty_(extracted.difficulty); // defaults to "No Idea..."
    }

    if (idxStatus !== null) {
      // Sidebar add = you/AI scanned requirements
      row[idxStatus] = computeStatus_({
        currentStatus: row[idxStatus],
        scanned: true,
        prepped: false,
        completed: false,
        surrendered: false
      });
    }

    if (idxTriage !== null) {
      const dueDate = (idxDue !== null) ? parseSheetDate_(row[idxDue]) : null;
      row[idxTriage] = computeTriage_(dueDate, row[idxTriage]); // Immediate/Urgent/Non-Urgent from due date
    }

    sh.appendRow(row);
    const addedRow = sh.getLastRow();

    // Make portal clickable
    if (row[idxPortal] && looksLikeUrl_(row[idxPortal])) {
      const rt = SpreadsheetApp.newRichTextValue()
        .setText(row[idxPortal])
        .setLinkUrl(row[idxPortal])
        .build();
      sh.getRange(addedRow, idxPortal + 1).setRichTextValue(rt);
    }

    return { addedRow };
  } finally {
    lock.releaseLock();
  }
}

function basicExtractFromText_(text, url) {
  const t = (text || "").replace(/\u00A0/g, " ").trim();
  const firstLine = (t.split(/\r?\n/).map(s => s.trim()).filter(Boolean)[0] || "").slice(0, 80);

  const dueMatch = t.match(/\b(due|deadline)\b[:\s-]*([^\n]{0,60})/i);
  const req = /essay|personal statement|short answer/i.test(t) ? "Essay(s)" : "";

  return {
    name: firstLine || "Scholarship",
    due: dueMatch ? dueMatch[2].trim() : "",
    requirements: req,
    difficulty: "",
    notes: ""
  };
}

// same simple extractor as before
function basicExtractFromText_(text, url) {
  const t = (text || "").replace(/\u00A0/g, " ").trim();
  const firstLine = (t.split(/\r?\n/).map(s => s.trim()).filter(Boolean)[0] || "").slice(0, 80);
  const dueMatch = t.match(/\b(due|deadline)\b[:\s-]*([^\n]{0,50})/i);
  const req = /essay/i.test(t) ? "Essay(s)" : "";
  return { name: firstLine || "Scholarship", portal: url || "", due: dueMatch ? dueMatch[2].trim() : "", requirements: req, difficulty: "", notes: "" };
}

/**
 * Very fast “good enough” extractor so the sidebar never hangs.
 * We can replace this with Gemini once everything is stable.
 */
function basicExtractFromText_(text, url) {
  const t = (text || "").replace(/\u00A0/g, " ").trim();

  // Name guess: first non-empty line up to 80 chars
  const firstLine = (t.split(/\r?\n/).map(s => s.trim()).filter(Boolean)[0] || "").slice(0, 80);

  // Due guess
  // (kept simple; better parsing later)
  const dueMatch = t.match(/\b(due|deadline)\b[:\s-]*([^\n]{0,50})/i);

  // Requirements guess
  const req = /essay/i.test(t) ? "Essay(s)" : "";

  return {
    name: firstLine || (url ? "Scholarship" : "Scholarship"),
    portal: url || "",
    due: dueMatch ? dueMatch[2].trim() : "",
    requirements: req,
    difficulty: "",
    notes: ""
  };
}

/*************************************************
 * 1) FAST: Create docs instantly for selected rows
 * - Only if Requirements includes "Essay(s)"
 * - Attaches Google Doc link into Application File (append, not overwrite)
 *************************************************/
function createEssayDocsForSelection() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(MAIN_SHEET);
  if (!sheet) throw new Error(`Missing sheet: ${MAIN_SHEET}`);

  const range = sheet.getActiveRange();
  if (!range) throw new Error("Select at least one cell in the row(s) you want.");

  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxReq    = optionalIndex_(h, COL_REQUIREMENTS);
  const idxPortal = optionalIndex_(h, COL_APPLICATION_PORTAL);
  const idxAppFile= optionalIndex_(h, COL_APPLICATION_FILE);
  const idxName   = optionalIndex_(h, "Scholarship Name");

  // Optional columns
  const idxDocId  = optionalIndex_(h, "Essay Doc ID");
  const idxStatus = optionalIndex_(h, "Status");
  const idxDue    = optionalIndex_(h, "Due Date");
  const idxTriage = optionalIndex_(h, "Triage");

  if (idxReq === null || idxAppFile === null || idxName === null) {
    throw new Error(
      `Missing required columns on "${MAIN_SHEET}": "Scholarship Name", "${COL_REQUIREMENTS}", "${COL_APPLICATION_FILE}"`
    );
  }

  const startRow = range.getRow();
  const endRow = range.getLastRow();
  const lastCol = sheet.getLastColumn();

  const numRows = endRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const rich = sheet.getRange(startRow, 1, numRows, lastCol).getRichTextValues();
  const formulas = sheet.getRange(startRow, 1, numRows, lastCol).getFormulas();

  let made = 0;

  for (let r = 0; r < data.length; r++) {
    const absoluteRow = startRow + r;
    if (absoluteRow <= HEADER_ROW) continue;

    const reqText = (data[r][idxReq] ?? "").toString();
    if (!requirementsIncludes_(reqText, ESSAY_REQUIREMENT_TOKEN)) continue;

    const name = (data[r][idxName] ?? "").toString().trim() || "Scholarship";

    const portalUrl = (idxPortal !== null)
      ? extractBestUrlFromCell_(
          formulas[r][idxPortal],
          rich[r][idxPortal],
          (data[r][idxPortal] ?? "").toString()
        )
      : "";

    const appFileCell = sheet.getRange(absoluteRow, idxAppFile + 1);

    // If already has any doc link, skip
    const existingUrl = extractDocUrlFromCell_(appFileCell);
    if (existingUrl) continue;

    const docInfo = createEssayPrepDocOneDraft_({
      scholarshipName: name,
      portalUrl,
      promptText: "",
      wordLimit: ""
    });

    // Store doc id if column exists
    if (idxDocId !== null && docInfo && docInfo.id) {
      sheet.getRange(absoluteRow, idxDocId + 1).setValue(docInfo.id);
    }

    // Put link into Application File
    appendLinkIntoCell_(sheet, absoluteRow, idxAppFile + 1, "Essay Draft (Google Doc)", docInfo.url);

    // Mark status as Prepped if Status column exists
    if (idxStatus !== null) {
      sheet.getRange(absoluteRow, idxStatus + 1).setValue("Prepped");
    }

    // Update triage if both columns exist
    if (idxDue !== null && idxTriage !== null) {
      const due = parseSheetDate_(data[r][idxDue]);
      sheet.getRange(absoluteRow, idxTriage + 1).setValue(computeTriage_(due, data[r][idxTriage]));
    }

    made++;
  }

  ss.toast(
    made ? `Created ${made} essay doc(s).` : "No rows needed an essay doc (Requirements must include Essay(s)).",
    "Scholarship Tools",
    6
  );
}

/*************************************************
 * 2) SLOW: Fill prompt(s) with AI for selected rows
 * - Only if Requirements includes "Essay(s)"
 * - Fetches portal (HTML/PDF) and writes prompt into the linked Google Doc
 * - If PDF upload required, creates PDF and attaches to Additional Application File
 *************************************************/
function fillEssayPromptsForSelection() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(MAIN_SHEET);

  const range = sheet.getActiveRange();
  if (!range) throw new Error("Select at least one cell in the row(s) you want.");

  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxReq = optionalIndex_(h, COL_REQUIREMENTS);
  const idxPortal = optionalIndex_(h, COL_APPLICATION_PORTAL);
  const idxAppFile = optionalIndex_(h, COL_APPLICATION_FILE);
  const idxAddl = optionalIndex_(h, COL_ADDITIONAL_APPLICATION_FILE);
  const idxName = optionalIndex_(h, "Scholarship Name");

  if (idxReq === null || idxPortal === null || idxAppFile === null || idxName === null) {
    throw new Error(`Missing required columns on active sheet: "Scholarship Name", "${COL_REQUIREMENTS}", "${COL_APPLICATION_PORTAL}", "${COL_APPLICATION_FILE}"`);
  }

  const startRow = range.getRow();
  const endRow = range.getLastRow();
  const lastCol = sheet.getLastColumn();

  const values = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getValues();
  const rich = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getRichTextValues();
  const formulas = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getFormulas();

  let filled = 0;

  for (let r = 0; r < values.length; r++) {
    const absoluteRow = startRow + r;
    if (absoluteRow <= HEADER_ROW) continue;

    const reqText = (values[r][idxReq] ?? "").toString();
    if (!requirementsIncludes_(reqText, ESSAY_REQUIREMENT_TOKEN)) continue;

    const name = (values[r][idxName] ?? "").toString().trim() || "Scholarship";
    const portalUrl = extractBestUrlFromCell_(formulas[r][idxPortal], rich[r][idxPortal], (values[r][idxPortal] ?? "").toString());

    const appFileCell = sheet.getRange(absoluteRow, idxAppFile + 1);

    // Create doc if missing
    let docUrl = extractDocUrlFromCell_(appFileCell);
    if (!docUrl) {
      const docInfo = createEssayPrepDocOneDraft_({ scholarshipName: name, portalUrl, promptText: "", wordLimit: "" });
      appendLinkIntoCell_(sheet, absoluteRow, idxAppFile + 1, "Essay Draft (Google Doc)", docInfo.url);
      docUrl = docInfo.url;
    }

    const docId = extractDriveIdFromUrl_(docUrl);
    if (!docId) continue;

    const promptPack = extractEssayPromptFromUrlOrPdf_(portalUrl);

    writePromptIntoEssayDoc_(docId, portalUrl, promptPack.prompt_text, promptPack.word_limit);

    if (idxAddl !== null && promptPack.pdf_required === true) {
      const pdfUrl = createPdfFromDoc_(docId, `Essay PDF — ${name}`);
      appendLinkIntoCell_(sheet, absoluteRow, idxAddl + 1, "Essay PDF (upload)", pdfUrl);
    }

    filled++;
  }

  ss.toast(
    filled ? `Filled ${filled} essay prompt(s).` : "No prompts filled (did you select rows with Essay(s)?)",
    "Scholarship Tools",
    6
  );
}

/*************************************************
 * DELETE: Trashes linked docs for selected rows
 * - Detects link via: formula, rich text runs, whole-cell link, raw text URL
 *************************************************/
function deleteEssayDocsForSelection() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(MAIN_SHEET);

  const range = sheet.getActiveRange();
  if (!range) throw new Error("Select at least one row.");

  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxAppFile = optionalIndex_(h, COL_APPLICATION_FILE);
  const idxDocId = optionalIndex_(h, "Essay Doc ID"); // NEW

  if (idxAppFile === null && idxDocId === null) {
    throw new Error(`Need "${COL_APPLICATION_FILE}" or "Essay Doc ID" column to delete.`);
  }

  const startRow = range.getRow();
  const endRow = range.getLastRow();

  let detected = 0;
  let trashed = 0;

  for (let row = startRow; row <= endRow; row++) {
    if (row <= HEADER_ROW) continue;

    // 1) Prefer stored ID
    let docId = "";
    if (idxDocId !== null) {
      docId = (sheet.getRange(row, idxDocId + 1).getValue() || "").toString().trim();
    }

    // 2) Fallback: try to extract from Application File cell
    let docUrl = "";
    if (!docId && idxAppFile !== null) {
      const cell = sheet.getRange(row, idxAppFile + 1);
      docUrl = extractDocUrlFromCell_(cell); // your robust extractor
      docId = extractDriveIdFromUrl_(docUrl);
    }

    if (!docId) {
      Logger.log(`[delete] Row ${row}: no docId found`);
      continue;
    }

    detected++;
    const ok = trashDriveFileBestEffort_(docId);
    Logger.log(`[delete] Row ${row}: docId=${docId} trashOk=${ok}`);
    if (ok) trashed++;

    // Clean up sheet cells (optional)
    if (idxDocId !== null) sheet.getRange(row, idxDocId + 1).clearContent();

    if (idxAppFile !== null) {
      const cell = sheet.getRange(row, idxAppFile + 1);
      const txt = (cell.getDisplayValue() || "").toString();
      const cleaned = txt
        .split(/\n+/).map(s => s.trim()).filter(Boolean)
        .filter(line => !/docs\.google\.com\/document\/d\//i.test(line))
        .filter(line => normalize_(line) !== normalize_("Essay Draft (Google Doc)"))
        .join("\n");
      if (!cleaned) cell.clearContent();
      else cell.setValue(cleaned);
    }
  }

  ss.toast(
    trashed
      ? `Trashed ${trashed} doc(s). (Detected ${detected})`
      : (detected ? `Detected ${detected} doc(s) but trash failed (permissions/API). Check Logs.` : "No docs detected in selected rows."),
    "Scholarship Tools",
    8
  );
}

/*************************************************
 * SYNC SATISFIED → RESULTS
 *************************************************/
function syncSatisfiedToResults() {
  const ss = SpreadsheetApp.getActive();
  const main = ss.getSheetByName(MAIN_SHEET);
  const results = ss.getSheetByName(RESULTS_SHEET);
  if (!main || !results) throw new Error("Missing required sheet(s).");

  const lastRow = main.getLastRow();
  const lastCol = main.getLastColumn();
  if (lastRow <= HEADER_ROW) {
    ss.toast("No data rows below header.", "Scholarship Tools", 5);
    return;
  }

  const numRows = lastRow - HEADER_ROW + 1;
  const mainRange = main.getRange(HEADER_ROW, 1, numRows, lastCol);

  const mainValues = mainRange.getValues();
  const mainRich = mainRange.getRichTextValues();
  const mainFormulas = mainRange.getFormulas();

  const h = buildHeaderIndex(mainValues[0]);

  const idxName = mustIndex(h, "Scholarship Name");
  const idxDue = mustIndex(h, "Due Date");
  const idxPortal = mustIndex(h, COL_APPLICATION_PORTAL);
  const idxNotes = mustIndex(h, "Notes");
  const idxDifficulty = mustIndex(h, "Difficulty");
  const idxTriage = mustIndex(h, "Triage");

  const idxReq = optionalIndex_(h, COL_REQUIREMENTS);
  const idxAppFile = optionalIndex_(h, COL_APPLICATION_FILE);
  const idxAddl = optionalIndex_(h, COL_ADDITIONAL_APPLICATION_FILE);

  const resValues = results.getDataRange().getValues();
  if (!resValues.length) throw new Error("Results sheet has no header row.");
  const rh = buildHeaderIndex(resValues[0]);

  const rName = mustIndex(rh, "Scholarship Name");
  const rDue = mustIndex(rh, "Due Date");
  const rPortal = mustIndex(rh, COL_APPLICATION_PORTAL);
  const rNotes = mustIndex(rh, "Notes");
  const rDifficulty = mustIndex(rh, "Difficulty");

  const rReq = optionalIndex_(rh, COL_REQUIREMENTS);
  const rAppFile = optionalIndex_(rh, COL_APPLICATION_FILE);
  const rAddl = optionalIndex_(rh, COL_ADDITIONAL_APPLICATION_FILE);

  const existing = new Set();
  for (let i = 1; i < resValues.length; i++) {
    const nm = (resValues[i][rName] || "").toString().trim();
    if (nm) existing.add(nm.toLowerCase());
  }

  const linkWrites = [];
  const rowsToRemove = [];
  let moved = 0;

 for (let i = 1; i < mainValues.length; i++) {
  const row = mainValues[i];
  const name = (row[idxName] || "").toString().trim();
  if (!name) continue;

  // ✅ Results sync is based on TRIAGE == Satisfied
  const triageNorm = normalize_((row[idxTriage] ?? "").toString());
  if (triageNorm !== "satisfied") continue;

  const key = name.toLowerCase();
  if (existing.has(key)) continue;

  const out = new Array(resValues[0].length).fill("");

  out[rName] = row[idxName];
  out[rDue] = row[idxDue];
  out[rDifficulty] = row[idxDifficulty];
  out[rNotes] = row[idxNotes];
  out[rPortal] = row[idxPortal];

  if (rReq !== null && idxReq !== null) out[rReq] = row[idxReq];

  // --- Application File chips ---
  if (rAppFile !== null && idxAppFile !== null) {
    out[rAppFile] = chipFormulaFromCell_(
      row[idxAppFile],
      mainRich[i][idxAppFile],
      mainFormulas[i][idxAppFile]
    );
  }

  if (rAddl !== null && idxAddl !== null) {
    out[rAddl] = chipFormulaFromCell_(
      row[idxAddl],
      mainRich[i][idxAddl],
      mainFormulas[i][idxAddl]
    );
  }

  results.appendRow(out);
  const appendedRow = results.getLastRow();

  // Preserve rich links in portal/files (if present)
  enqueuePreservedLinkWrite_({
    linkWrites,
    srcValues: row,
    srcRichRow: mainRich[i],
    srcFormulaRow: mainFormulas[i],
    srcIdx: idxPortal,
    dstRow: appendedRow,
    dstCol1Based: rPortal + 1
  });

  if (idxAppFile !== null && rAppFile !== null) {
    enqueuePreservedLinkWrite_({
      linkWrites,
      srcValues: row,
      srcRichRow: mainRich[i],
      srcFormulaRow: mainFormulas[i],
      srcIdx: idxAppFile,
      dstRow: appendedRow,
      dstCol1Based: rAppFile + 1
    });
  }

  if (idxAddl !== null && rAddl !== null) {
    enqueuePreservedLinkWrite_({
      linkWrites,
      srcValues: row,
      srcRichRow: mainRich[i],
      srcFormulaRow: mainFormulas[i],
      srcIdx: idxAddl,
      dstRow: appendedRow,
      dstCol1Based: rAddl + 1
    });
  }

  existing.add(key);
  moved++;
  rowsToRemove.push(HEADER_ROW + i);
}
// <-- Add this closing brace to end syncSatisfiedToResults
}

function syncSurrenderedToSheet() {
  const ss = SpreadsheetApp.getActive();
  const main = ss.getSheetByName(MAIN_SHEET);
  if (!main) throw new Error("Missing Main sheet.");

  const SURRENDERED_SHEET = "Surrendered";
  let surrendered = ss.getSheetByName(SURRENDERED_SHEET);
  if (!surrendered) surrendered = ss.insertSheet(SURRENDERED_SHEET);

  const HEADERS = ["Scholarship Name", "Date of Surrender", "Reason", "Lesson"];

  // Ensure Surrendered header row is exactly the 4 headers
  const sLastRow = surrendered.getLastRow();
  if (sLastRow < 1) {
    surrendered.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    const existingHeader = surrendered.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    const same =
      existingHeader.length === HEADERS.length &&
      existingHeader.every((v, i) => String(v) === String(HEADERS[i]));
    if (!same) {
      surrendered.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  }

  const lastRow = main.getLastRow();
  const lastCol = main.getLastColumn();
  if (lastRow <= HEADER_ROW) {
    ss.toast("No data rows below header.", "Scholarship Tools", 5);
    return;
  }

  const numRows = lastRow - HEADER_ROW + 1;
  const mainRange = main.getRange(HEADER_ROW, 1, numRows, lastCol);
  const mainValues = mainRange.getValues();
  const h = buildHeaderIndex(mainValues[0]);
  const mainDisplay = mainRange.getDisplayValues();

  const idxName = mustIndex(h, "Scholarship Name");
  const idxStatus = mustIndex(h, "Status");


  // If you have dedicated columns in Main for Reason/Lesson, use them.
  // If not, we’ll pull from Notes as a fallback.
  const idxNotes = optionalIndex_(h, "Notes");
  const idxReason = optionalIndex_(h, "Reason");
  const idxLesson = optionalIndex_(h, "Lesson");

  // Build existing set in Surrendered by scholarship name (avoid duplicates)
  const sData = surrendered.getDataRange().getValues();
  const existing = new Set();
  for (let r = 1; r < sData.length; r++) {
    const nm = (sData[r][0] || "").toString().trim();
    if (nm) existing.add(nm.toLowerCase());
  }

  const rowsToDelete = [];
  const rowsToAppend = [];

  const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  for (let i = 1; i < mainValues.length; i++) {
  // ✅ Surrendered sync is based on STATUS == Surrendered (use display values)
  const statusDisp = String(mainDisplay[i][idxStatus] ?? "").trim();
  const statusNorm = normalize_(statusDisp);
  if (statusNorm !== "surrendered") continue;

  const name = (mainValues[i][idxName] || "").toString().trim();
  if (!name) continue;

  // Pull reason/lesson (prefer dedicated cols; fallback to Notes)
  const reason =
    (idxReason !== null ? (mainValues[i][idxReason] || "") : "") ||
    (idxNotes !== null ? (mainValues[i][idxNotes] || "") : "");

  const lesson =
    (idxLesson !== null ? (mainValues[i][idxLesson] || "") : "");

  if (!existing.has(name.toLowerCase())) {
    rowsToAppend.push([name, todayStr, String(reason || ""), String(lesson || "")]);
    existing.add(name.toLowerCase());
  }

  rowsToDelete.push(HEADER_ROW + i);
}

  if (!rowsToDelete.length) {
    ss.toast('No rows marked "Surrendered".', "Scholarship Tools", 5);
    return;
  }

  // Append to Surrendered
  if (rowsToAppend.length) {
    const startRow = Math.max(surrendered.getLastRow(), 1) + 1;
    surrendered.getRange(startRow, 1, rowsToAppend.length, HEADERS.length).setValues(rowsToAppend);
  }

  // Delete from Main bottom-up
  rowsToDelete.sort((a, b) => b - a);
  for (const r of rowsToDelete) main.deleteRow(r);

  ss.toast(
    `Moved ${rowsToAppend.length} new row(s) to Surrendered; removed ${rowsToDelete.length} from Main.`,
    "Scholarship Tools",
    6
  );
}
/*************************************************
 * AI extraction (URL/PDF) for prompt
 *************************************************/
function extractEssayPromptFromUrlOrPdf_(url) {
  if (!url) return { prompt_text: "", word_limit: "", pdf_required: false };

  let sourceText = "";
  const lower = url.toLowerCase();

  if (lower.endsWith(".pdf") || lower.includes(".pdf?")) sourceText = fetchPdfAsTextViaDriveConvert_(url);
  else sourceText = fetchHtmlAsText_(url);

  if (!sourceText) return { prompt_text: "", word_limit: "", pdf_required: false };

  const snippet = extractLikelyPromptSection_(sourceText);
  const pack = geminiExtractEssayPrompt_(snippet, url);

  if (!pack.prompt_text || pack.prompt_text.trim().length < 20) {
    return {
      prompt_text: heuristicPromptFallback_(snippet) || heuristicPromptFallback_(sourceText) || "",
      word_limit: pack.word_limit || heuristicLimitFallback_(snippet) || heuristicLimitFallback_(sourceText) || "",
      pdf_required: pack.pdf_required === true
    };
  }

  return pack;
}

function fetchHtmlAsText_(url) {
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) return "";

    const html = res.getContentText() || "";

    const text = html
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<(br|BR)\s*\/?>/g, "\n")
      .replace(/<\/(p|div|li|h1|h2|h3|h4|tr)>/gi, "\n")
      .replace(/<li[^>]*>/gi, "• ")
      .replace(/<[^>]+>/g, " ")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .replace(/[ \t]{2,}/g, " ")
      .trim();

    return text.slice(0, MAX_HTML_CHARS);
  } catch (e) {
    Logger.log("fetchHtmlAsText_ error: " + (e?.message || e));
    return "";
  }
}

/**
 * PDF->Text by: download -> temp Drive file -> Drive convert to Google Doc -> extract text -> cleanup
 * Requires enabling Advanced Google Service: Drive API (Drive.Files)
 */
function fetchPdfAsTextViaDriveConvert_(pdfUrl) {
  try {
    const fileId = extractDriveIdFromUrl_(pdfUrl);
    let blob;

    if (fileId) {
      blob = DriveApp.getFileById(fileId).getBlob();
    } else {
      const res = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true, followRedirects: true });
      const code = res.getResponseCode();
      if (code < 200 || code >= 300) return "";
      blob = res.getBlob();
    }

    const tmpPdf = DriveApp.createFile(blob).setName("tmp_pdf_" + Date.now() + ".pdf");

    const converted = Drive.Files.copy(
      { title: "tmp_pdf_doc_" + Date.now(), mimeType: MimeType.GOOGLE_DOCS },
      tmpPdf.getId()
    );

    const docId = converted.id;
    const text = DocumentApp.openById(docId).getBody().getText();

    try { DriveApp.getFileById(tmpPdf.getId()).setTrashed(true); } catch (_) {}
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (_) {}

    return (text || "").replace(/\s+/g, " ").trim().slice(0, MAX_HTML_CHARS);
  } catch (e) {
    Logger.log("fetchPdfAsTextViaDriveConvert_ error: " + (e?.message || e));
    return "";
  }
}

function geminiExtractEssayPrompt_(sourceText, url) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("Missing Script Property: GEMINI_API_KEY");

  const endpoint =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(GEMINI_MODEL) +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  const schema = {
    type: "object",
    additionalProperties: false,
    properties: {
      prompt_text: { type: "string" },
      word_limit: { type: "string" },
      pdf_required: { type: "boolean" }
    },
    required: ["prompt_text", "word_limit", "pdf_required"]
  };

  const userText = [
    "Extract ONLY the scholarship essay/writing prompt(s) and limit from the source text.",
    "Return ONLY valid JSON (no markdown).",
    "pdf_required must be true ONLY if it explicitly says upload/submit the essay as a PDF.",
    "",
    "URL:",
    url,
    "",
    "JSON Schema:",
    JSON.stringify(schema),
    "",
    "SOURCE TEXT:",
    sourceText
  ].join("\n");

  const body = {
    contents: [{ role: "user", parts: [{ text: userText }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.1 }
  };

  const res = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const raw = res.getContentText() || "";

  if (code < 200 || code >= 300) {
    Logger.log("Gemini HTTP " + code + " for " + url);
    Logger.log("Gemini raw error: " + raw.slice(0, 2000));
    return { prompt_text: "", word_limit: "", pdf_required: false };
  }

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    Logger.log("Gemini response not JSON. Raw: " + raw.slice(0, 2000));
    return { prompt_text: "", word_limit: "", pdf_required: false };
  }

  const text = (parsed?.candidates?.[0]?.content?.parts || [])
    .map(p => p.text || "")
    .join("")
    .trim();

  if (!text) {
    Logger.log("Gemini returned empty candidate text for " + url);
    return { prompt_text: "", word_limit: "", pdf_required: false };
  }

  try {
    const obj = JSON.parse(text);
    const combined = (obj.prompt_text || "").toLowerCase();
    const pdf = (obj.pdf_required === true) || PDF_REQUIRED_KEYWORDS.some(k => combined.includes(k));
    return { prompt_text: obj.prompt_text || "", word_limit: obj.word_limit || "", pdf_required: pdf };
  } catch (e) {
    Logger.log("Gemini candidate not valid JSON. Candidate: " + text.slice(0, 2000));
    return { prompt_text: "", word_limit: "", pdf_required: false };
  }
}

function heuristicLimitFallback_(t) {
  const m = (t || "").match(/\b(\d{2,4})\s*(word|words|character|characters)\b/i);
  return m ? `${m[1]} ${m[2].toLowerCase()}` : "";
}

function heuristicPromptFallback_(t) {
  if (!t) return "";
  const matches = t.match(/(?:^|\n|\s)(Describe|Explain|Discuss|Tell us|Reflect|Write|Respond)[^\n?]{0,400}\?/gi);
  if (matches && matches.length) return matches.slice(0, 3).join("\n\n").trim();
  return "";
}

function extractLikelyPromptSection_(text) {
  if (!text) return "";

  const t = text
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const lower = t.toLowerCase();

  const strongAnchors = [
    "essay question", "essay questions", "essay prompt", "writing prompt",
    "writing requirement", "writing requirements",
    "short answer", "short-answer", "short answers",
    "personal statement", "statement of purpose",
    "topic", "topics", "prompt", "prompts",
    "question", "questions"
  ];

  let best = null; // { start, end, score }

  function considerWindow(centerIdx, score) {
    const start = Math.max(0, centerIdx - 1800);
    const end = Math.min(t.length, centerIdx + 9000);
    if (!best || score > best.score) best = { start, end, score };
  }

  for (const a of strongAnchors) {
    const idx = lower.indexOf(a);
    if (idx !== -1) { considerWindow(idx, 100); break; }
  }

  if (!best) {
    const headingColonRegex = /\b(Essay Question\(s\)|Essay Questions|Essay Prompt|Writing Prompt|Short Answer|Personal Statement|Prompt|Question\(s\)|Questions)\s*:\s*/i;
    const m = t.match(headingColonRegex);
    if (m && m.index != null) considerWindow(m.index, 95);
  }

  if (!best) {
    const qRegex = /\b(Describe|Explain|Discuss|Tell us|Tell me|Reflect|Share|In \d+ words|In \d+ characters|Write|Respond)\b[^.?!]{0,220}[?]/i;
    const qm = t.match(qRegex);
    if (qm && qm.index != null) considerWindow(qm.index, 90);
  }

  if (!best) {
    const limitRegex = /\b(\d{2,4})\s*(word|words|character|characters)\b/i;
    const lm = t.match(limitRegex);
    if (lm && lm.index != null) considerWindow(lm.index, 80);
  }

  if (!best) {
    const softNeedles = ["essay", "prompt", "writing", "short answer", "personal statement", "respond", "describe", "word limit", "characters"];
    let idx = -1;
    for (const n of softNeedles) {
      const i = lower.indexOf(n);
      if (i !== -1 && (idx === -1 || i < idx)) idx = i;
    }
    if (idx !== -1) considerWindow(idx, 60);
  }

  if (!best) return t.slice(0, 12000);

  const snippet = t.slice(best.start, best.end);

  const tightenRegex = /\b(Essay Question\(s\)|Essay Questions|Essay Prompt|Writing Prompt|Short Answer|Personal Statement|Prompt|Question\(s\)|Questions)\b/i;
  const tm = snippet.match(tightenRegex);
  if (tm && tm.index != null && tm.index > 0) return snippet.slice(tm.index);

  return snippet;
}

/*************************************************
 * DOC: create + write prompt (Times New Roman, 1 draft)
 *************************************************/
function createEssayPrepDocOneDraft_(opts) {
  const { scholarshipName, portalUrl, promptText, wordLimit } = opts;
  const title = `Kevin Srun - ${scholarshipName}`;
  const doc = DocumentApp.create(title);
  writePromptIntoEssayDoc_(doc.getId(), portalUrl, promptText, wordLimit);
  return { id: doc.getId(), url: doc.getUrl() };
}

function writePromptIntoEssayDoc_(docId, portalUrl, promptText, wordLimit) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  body.clear();

  const font = "Times New Roman";
  const size = 12;

  function addPara(text, heading) {
    const p = body.appendParagraph(text);
    if (heading) p.setHeading(heading);
    p.setFontFamily(font).setFontSize(size);
    return p;
  }

  addPara(doc.getName(), DocumentApp.ParagraphHeading.HEADING1);
  if (portalUrl) addPara("Application Portal: " + portalUrl, null);

  addPara("Prompt", DocumentApp.ParagraphHeading.HEADING2);
  if (wordLimit) addPara("Word/Character Limit: " + wordLimit, null);

  addPara(promptText || "(Prompt not detected — paste it here manually.)", null);

  addPara("Draft", DocumentApp.ParagraphHeading.HEADING2);
  addPara("", null);

  doc.saveAndClose();
}

function createPdfFromDoc_(docId, pdfName) {
  const file = DriveApp.getFileById(docId);
  const pdfBlob = file.getBlob().getAs(MimeType.PDF).setName((pdfName || "Essay") + ".pdf");
  const pdfFile = DriveApp.createFile(pdfBlob);
  return pdfFile.getUrl();
}

/*************************************************
 * DETECT DOC URL IN A CELL (robust)
 *************************************************/
function extractDocUrlFromCell_(cell) {
  // 1) HYPERLINK formula
  const f = (cell.getFormula() || "").trim();
  if (f && /^=HYPERLINK\(/i.test(f)) {
    const m = f.match(/=HYPERLINK\(\s*"([^"]+)"/i);
    if (m && m[1] && isGoogleDocUrl_(m[1])) return m[1];
  }

  // 2) RichText runs
  const rt = cell.getRichTextValue();
  if (rt) {
    const runs = rt.getRuns();
    for (const run of runs) {
      const url = run.getLinkUrl();
      if (isGoogleDocUrl_(url)) return url;
    }
    const whole = rt.getLinkUrl();
    if (isGoogleDocUrl_(whole)) return whole;
  }

  // 3) Plain text regex
  const text = (cell.getDisplayValue() || "").toString();
  const m2 = text.match(/https:\/\/docs\.google\.com\/document\/d\/[a-zA-Z0-9_-]+/i);
  if (m2 && m2[0]) return m2[0];

  return "";
}

function isGoogleDocUrl_(url) {
  return /https:\/\/docs\.google\.com\/document\/d\//i.test(url || "");
}

function extractDriveIdFromUrl_(url) {
  if (!url) return "";
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : "";
}

/*************************************************
 * TRASH FILE (DriveApp first, Drive API fallback)
 *************************************************/
function trashDriveFileBestEffort_(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return true;
  } catch (e1) {
    try {
      // Requires Advanced Google Service: Drive API enabled
      Drive.Files.update({ trashed: true }, fileId);
      return true;
    } catch (e2) {
      Logger.log(`trashDriveFileBestEffort_ failed for ${fileId}: ${e2?.message || e2}`);
      return false;
    }
  }
}

/*************************************************
 * Append link into cell without overwriting
 *************************************************/
function appendLinkIntoCell_(sheet, row, col, label, url) {
  const cell = sheet.getRange(row, col);

  const existingRT = cell.getRichTextValue();
  const existingText = (existingRT ? existingRT.getText() : (cell.getDisplayValue() || "")).toString();

  const sep = existingText.trim() ? "\n" : "";
  const newText = existingText + sep + label;

  const builder = SpreadsheetApp.newRichTextValue().setText(newText);

  // preserve a single existing whole-cell link if present
  if (existingRT && existingRT.getLinkUrl() && existingText.length > 0) {
    builder.setLinkUrl(0, existingText.length, existingRT.getLinkUrl());
  }

  const start = (existingText + sep).length;
  const end = start + label.length;
  builder.setLinkUrl(start, end, url);

  cell.setRichTextValue(builder.build());
}

/*************************************************
 * Requirements multi-select parsing
 *************************************************/
function requirementsIncludes_(reqCellText, token) {
  if (!reqCellText || !token) return false;

  const normToken = normalize_(token);
  const parts = reqCellText
    .toString()
    .replace(/\u00A0/g, " ")
    .split(/[,;\n]+/)
    .map(s => normalize_(s))
    .filter(Boolean);

  return parts.includes(normToken);
}

/*************************************************
 * Link preserving copy for Sync
 *************************************************/
function enqueuePreservedLinkWrite_(opts) {
  const { linkWrites, srcValues, srcRichRow, srcFormulaRow, srcIdx, dstRow, dstCol1Based } = opts;
  if (srcIdx === null || srcIdx === undefined) return;

  const formula = (srcFormulaRow[srcIdx] || "").trim();
  const rich = srcRichRow[srcIdx];
  const value = (srcValues[srcIdx] ?? "").toString().trim();

  if (formula) {
    linkWrites.push({ row: dstRow, col: dstCol1Based, formula });
    return;
  }
  if (rich && rich.getLinkUrl()) {
    linkWrites.push({ row: dstRow, col: dstCol1Based, richText: rich });
    return;
  }
  if (looksLikeUrl_(value)) {
    const rt = SpreadsheetApp.newRichTextValue().setText(value).setLinkUrl(value).build();
    linkWrites.push({ row: dstRow, col: dstCol1Based, richText: rt });
    return;
  }
  if (rich) linkWrites.push({ row: dstRow, col: dstCol1Based, richText: rich });
}

function extractBestUrlFromCell_(formula, rich, value) {
  if (rich && rich.getLinkUrl()) return rich.getLinkUrl();
  if (formula && formula.toUpperCase().startsWith("=HYPERLINK(")) {
    const m = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
    if (m && m[1]) return m[1];
  }
  if (looksLikeUrl_(value || "")) return value.trim();
  return "";
}

/*************************************************
 * Debug triage
 *************************************************/
function debugTriageValues() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MAIN_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${MAIN_SHEET}`);

  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const h = buildHeaderIndex(headerRow);

  const triageCol = mustIndex(h, "Triage") + 1;
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) {
    ss.toast("No data rows below header.", "Scholarship Tools", 5);
    return;
  }

  const vals = sh.getRange(HEADER_ROW + 1, triageCol, lastRow - HEADER_ROW, 1).getValues().flat();
  const counts = {};
  vals.forEach(v => {
    const raw = (v ?? "").toString();
    if (!raw) return;
    const key = normalize_(raw);
    counts[key] = (counts[key] || 0) + 1;
  });

  Logger.log("Triage values (normalized) + counts:");
  Object.keys(counts).sort().forEach(k => Logger.log(`${k} => ${counts[k]}`));

  ss.toast("Check Apps Script → Executions → Logs for triage values.", "Scholarship Tools", 6);
}

/*************************************************
 * TRIAGE: Recompute for all rows in Main Menu
 * - Uses Due Date + computeTriage_()
 * - Preserves terminal triage (Satisfied, Cooked)
 * - If Status == Completed => Triage = Satisfied (optional but recommended)
 *************************************************/
function refreshTriageForMainMenu() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(MAIN_SHEET);
    if (!sh) throw new Error(`Sheet not found: ${MAIN_SHEET}`);

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= HEADER_ROW) {
      ss.toast("No data rows below header.", "Scholarship Tools", 5);
      return;
    }

    const range = sh.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, lastCol);
    const values = range.getValues();
    const display = range.getDisplayValues(); // helps with dropdown-rendered text
    const headers = values[0];
    const h = buildHeaderIndex(headers);

    const idxDue = optionalIndex_(h, "Due Date");
    const idxTriage = optionalIndex_(h, "Triage");
    const idxStatus = optionalIndex_(h, "Status");

    if (idxTriage === null) throw new Error('Missing column header: "Triage"');
    if (idxDue === null) throw new Error('Missing column header: "Due Date"');

    const triageUpdates = [];
    let changed = 0;

    for (let r = 1; r < values.length; r++) {
      const status = (idxStatus !== null ? String(display[r][idxStatus] ?? "") : "").trim();
      const statusNorm = normalize_(status);

      // If completed, triage should be satisfied (recommended behavior)
      if (statusNorm === "completed") {
        const current = String(values[r][idxTriage] ?? "").trim();
        if (normalize_(current) !== "satisfied") changed++;
        triageUpdates.push(["Satisfied"]);
        continue;
      }

      // Preserve terminal triage values
      const currentTriageRaw = String(values[r][idxTriage] ?? "").trim();
      const currentTriageNorm = normalize_(currentTriageRaw);

      if (currentTriageNorm === "satisfied" || currentTriageNorm === "cooked" || currentTriageNorm === "cooked") {
        triageUpdates.push([currentTriageRaw]); // keep exactly as-is
        continue;
      }

      const due = parseSheetDate_(values[r][idxDue]);
      const nextTriage = computeTriage_(due, currentTriageRaw);

      if (normalize_(nextTriage) !== currentTriageNorm) changed++;
      triageUpdates.push([nextTriage]);
    }

    // Write back in one shot
    sh.getRange(HEADER_ROW + 1, idxTriage + 1, triageUpdates.length, 1).setValues(triageUpdates);

    ss.toast(`Triage refreshed. Updated ${changed} row(s).`, "Scholarship Tools", 6);
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * Shared helpers
 *************************************************/
function buildHeaderIndex(headerRow) {
  const m = {};
  for (let i = 0; i < headerRow.length; i++) {
    const key = (headerRow[i] || "").toString().trim();
    if (key) m[key] = i;
  }
  return m;
}

function mustIndex(map, name) {
  if (!(name in map)) throw new Error(`Missing column header: "${name}"`);
  return map[name];
}

function optionalIndex_(map, name) {
  return (name in map) ? map[name] : null;
}

function looksLikeUrl_(s) {
  return /^https?:\/\/\S+$/i.test(s);
}

function normalize_(s) {
  return (s ?? "")
    .toString()
    .replace(/\u00A0/g, " ")
    .trim()
    .toLowerCase();
}
function normHeader_(s) {
  return (s ?? "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}

function findHeaderIndex_(headers, candidates) {
  const map = new Map();
  headers.forEach((h, i) => map.set(normHeader_(h), i));

  for (const c of candidates) {
    const idx = map.get(normHeader_(c));
    if (idx != null) return idx;
  }

  // fallback: contains match (handles headers like "Scholarship Name (Official)")
  const lower = headers.map(h => normHeader_(h));
  for (const c of candidates) {
    const needle = normHeader_(c);
    const idx = lower.findIndex(h => h.includes(needle));
    if (idx !== -1) return idx;
  }
  return null;
}

function logHeaders_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MAIN_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${MAIN_SHEET}`);
  const headers = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  Logger.log("HEADERS: " + headers.join(" | "));
  return headers;
}
function debugHeaders() {
  logHeaders_();
}

const DIFFICULTY_OPTIONS = ["Easy", "Medium", "Hard", "Fuck it", "No Idea..."];
const TRIAGE_OPTIONS = ["Immediate", "Urgent", "Non-Urgent", "Satisfied", "Cooked"];
const STATUS_OPTIONS = ["Not started", "Scanned", "Prepped", "In Progress", "Completed", "Surrendered"];
const COL_THEME = "Theme";

// Dropdown options (keep EXACT spelling to match your sheet)
const THEME_OPTIONS = [
  "Biomed.", "Comp Sci.", "Business/Entrepreneurship", "Pol. Sci.", "Herritage",
  "Engineering", "Math/Sci", "First Gen", "Broke", "English/Humanities",
  "Leadership/Service", "None", "Other"
];

const REQUIREMENT_OPTIONS = [
  "Essay(s)", "Rec. Letter(s)", "App./ECs)", "Academics/Transcript",
  "Field of Study", "Exam(s)", "Interview(s)", "Broke/Finances",
  "Video/Recording", "Other"
];

// --- normalize + clamp ---
function norm_(s) {
  return (s ?? "").toString().trim().toLowerCase().replace(/\u00A0/g, " ").replace(/\s+/g, " ");
}

function clampToOptions_(raw, options, fallback) {
  const r = norm_(raw);
  if (!r) return fallback;

  // exact match
  for (const o of options) if (norm_(o) === r) return o;

  // contains match
  for (const o of options) {
    const no = norm_(o);
    if (no.includes(r) || r.includes(no)) return o;
  }
  return fallback;
}

// --- parse Due Date robustly ---
function parseSheetDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  // handle strings like "2/7/2026" or "2026-02-07"
  const s = v.toString().trim();
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  // try mm/dd/yyyy
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const mm = Number(m[1]), dd = Number(m[2]), yy = Number(m[3]);
    const yyyy = yy < 100 ? (2000 + yy) : yy;
    const d2 = new Date(yyyy, mm - 1, dd);
    if (!isNaN(d2.getTime())) return d2;
  }
  return null;
}

function daysUntil_(dueDate) {
  if (!dueDate) return null;
  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfDue = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
  const ms = startOfDue.getTime() - startOfToday.getTime();
  return Math.floor(ms / (24 * 3600 * 1000));
}

// --- triage rules ---
function computeTriage_(dueDate, currentTriage) {
  // Preserve terminal states
  const ct = clampToOptions_(currentTriage, TRIAGE_OPTIONS, "");
  if (ct === "Satisfied" || ct === "Cooked") return ct;

  const d = daysUntil_(dueDate);
  if (d == null) return "Non-Urgent"; // no due date -> default

  if (d < 0) return "Cooked";            // past due
  if (d <= 7) return "Immediate";         // < 1 week
  if (d <= 28) return "Urgent";           // 2–4 weeks (we include 8–28 days)
  return "Non-Urgent";                    // > 4 weeks
}

function uniq_(arr) {
  const seen = new Set();
  const out = [];
  for (const x of (arr || [])) {
    const k = norm_(x);
    if (!k || seen.has(k)) continue;
    seen.add(k);
    out.push(x);
  }
  return out;
}

function joinMultiSelect_(items) {
  return uniq_(items).join(", ");
}

// ===== Helpers (bottom of file) =====

function fileChip_(driveId) {
  return `=FILE("${driveId}")`;
}

function linkChip_(url, label) {
  return `=HYPERLINK("${url}", "${label}")`;
}

function driveIdFromUrl_(url) {
  if (!url) return "";
  const u = String(url);

  // /d/<id>
  let m = u.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];

  // ?id=<id>
  m = u.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];

  return "";
}

function chipFormulaFromCell_(plainValue, richText, formula) {
  // Prefer a clickable URL from rich text link if present
  let url = "";
  try {
    if (richText && typeof richText.getLinkUrl === "function") {
      url = richText.getLinkUrl() || "";
    }
  } catch (e) {}

  // If no rich link, try formula like HYPERLINK("url","label")
  if (!url && formula) {
    const m = String(formula).match(/HYPERLINK\("([^"]+)"/i);
    if (m) url = m[1];
  }

  // If still no url, maybe the plain value is a URL
  if (!url && plainValue && /^https?:\/\//i.test(String(plainValue))) {
    url = String(plainValue);
  }

  // Build chip
  const id = driveIdFromUrl_(url);
  if (id) return `=FILE("${id}")`;         // ✅ true Drive file chip
  if (url) return `=HYPERLINK("${url}","Open")`; // fallback link
  return ""; // nothing
}

function detectThemesFromText_(text) {
  const t = norm_(text);

  // Fast keyword buckets
  const hits = [];

  if (/\b(biomed|medical|medicine|health|hospital|nursing|patient|clinical|physician|biology|biochem|bioengineering|premed)\b/.test(t))
    hits.push("Biomed.");

  if (/\b(computer science|comp sci|software|programming|coding|developer|cs|machine learning|ai|data science|cyber|security)\b/.test(t))
    hits.push("Comp Sci.");

  if (/\b(engineering|mechanical|electrical|civil|chemical engineering|a\/e\/c|construction|architecture)\b/.test(t))
    hits.push("Engineering");

  if (/\b(math|mathematics|calculus|statistics|physics|chemistry|stem)\b/.test(t))
    hits.push("Math/Sci");

  if (/\b(english|humanities|literature|history|philosophy|writing)\b/.test(t))
    hits.push("English/Humanities");

  if (/\b(business|entrepreneur|entrepreneurship|startup|marketing|finance|economics)\b/.test(t))
    hits.push("Business/Entrepreneurship");

  if (/\b(political|politics|government|public policy|policy|civics|international relations|law|legal)\b/.test(t))
    hits.push("Pol. Sci.");

  if (/\b(first[- ]gen|first generation)\b/.test(t))
    hits.push("First Gen");

  if (/\b(financial need|need[- ]based|low[- ]income|income|fafsa|pell|scholarship need)\b/.test(t))
    hits.push("Broke");

  if (/\b(heritage|cultural|ethnic|minority|refugee|immigrant|asian|latino|black|african american|native)\b/.test(t))
    hits.push("Herritage");

  if (/\b(leadership|service|volunteer|community|nonprofit|mentor|tutor)\b/.test(t))
    hits.push("Leadership/Service");

  // If nothing matched, set Other (never blank)
  const finalHits = uniq_(hits);
  return finalHits.length ? finalHits : ["Other"];
}

function detectRequirementsFromText_(text) {
  const t = norm_(text);
  const req = [];

  if (/\b(essay|personal statement|short answer|writing prompt)\b/.test(t))
    req.push("Essay(s)");

  if (/\b(letter of recommendation|recommendation letter|rec letter|reference letter)\b/.test(t))
    req.push("Rec. Letter(s)");

  if (/\b(resume|activities|extracurricular|ec|experience|work experience|volunteer hours|community service)\b/.test(t))
    req.push("App./ECs"); // match your dropdown text exactly

  if (/\b(transcript|gpa|grades|academic record|class rank)\b/.test(t))
    req.push("Academics/Transcript");

  if (/\b(major|field of study|must be enrolled in|degree program|pursuing)\b/.test(t))
    req.push("Field of Study");

  if (/\b(exam|test score|sat|act|ap score)\b/.test(t))
    req.push("Exam(s)");

  if (/\b(interview)\b/.test(t))
    req.push("Interview(s)");

  if (/\b(financial|income|tax return|fafsa|need[- ]based|household|budget)\b/.test(t))
    req.push("Broke/Finances");

  if (/\b(video|recording|youtube|tiktok|reel|submit a video)\b/.test(t))
    req.push("Video/Recording");

  const finalReq = uniq_(req);
  return finalReq.length ? finalReq : [];
}

// --- status rules (based on what actions happened) ---
function computeStatus_(opts) {
  // opts: { currentStatus, scanned, prepped, completed, surrendered }
  const cs = clampToOptions_(opts.currentStatus, STATUS_OPTIONS, "Not started");

  // Preserve terminal
  if (cs === "Completed" || opts.completed) return "Completed";
  if (cs === "Surrendered" || opts.surrendered) return "Surrendered";

  // Progression
  if (opts.prepped) return "Prepped";
  if (opts.scanned) return "Scanned";
  return "Not started";
}

// difficulty: default to No idea... unless user/you explicitly set
function computeDifficulty_(rawDifficulty) {
  return clampToOptions_(rawDifficulty, DIFFICULTY_OPTIONS, "No Idea...");
}

function installDailyTriageRefreshTrigger() {
  // runs around 6am in your spreadsheet’s timezone
  ScriptApp.newTrigger("refreshTriageForMainMenu")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}

function _pushProof() {
  SpreadsheetApp.getActive().toast("CLASP IS FIXED", "Autoship", 4);
}

function findColByNorm_(headerRow, targetNorm) {
  const t = normalize_(targetNorm);
  for (let c = 0; c < headerRow.length; c++) {
    if (normalize_(String(headerRow[c] ?? "")) === t) return c;
  }
  return null;
}
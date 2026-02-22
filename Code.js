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

// Reminder schedule (days before due date)
const SCHOLARSHIP_REMINDER_DAYS = [7, 3, 1]; // change if you want 14,10,5,3,1 etc.

// Status words that should REMOVE the event
const REMOVE_EVENT_TRIAGE_VALUES = ["surrendered"];
const REMOVE_EVENT_STATUS_VALUES = ["submitted", "complete", "completed", "won", "not applying"];
/*************************************************
 * MENU
 *************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Scholarship Tools")
    .addItem("Sync Satisfied → Results", "syncSatisfiedToResults")
    .addItem("Sync Surrendered → Surrendered", "syncSurrenderedToSheet")
    .addItem("Sync Deadlines → Google Calendar", "syncScholarshipDeadlinesToCalendar")
    .addSeparator()
    .addItem("Refresh Triage (Main Menu)", "refreshTriageForMainMenu")
    .addSeparator()
    .addItem("Create Essay Doc(s) for Selected Rows (FAST)", "createEssayDocsForSelection")
    .addItem("Fill Essay Prompt(s) with AI for Selected Rows (SLOW)", "fillEssayPromptsForSelection")
    .addSeparator()
    .addItem("Pipe Canvas Intake → Main Menu", "pipeCanvasHitsToMainMenu")
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

    // Auto-create Application File (essay doc) if applicable
    autoCreateEssayDocForRow_(sh, addedRow, extracted, ctx);

    // Draft recommender emails
    draftRecEmailsForRow_(sh, addedRow);

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

function aiSidebarUploadFile(payload) {
  payload = payload || {};
  const name = (payload.name || ("upload_" + Date.now())).toString();
  const mimeType = (payload.mimeType || "application/octet-stream").toString();
  const data = payload.data || [];
  if (!data.length) throw new Error("No file bytes received.");

  const blob = Utilities.newBlob(data, mimeType, name);
  const f = DriveApp.createFile(blob);
  return {
    driveFileId: f.getId(),
    driveUrl: f.getUrl(),
    name: f.getName(),
    mimeType: f.getMimeType()
  };
}

function detectEssayPrompt_(text) {
  const t = (text || "").replace(/\u00A0/g, " ").trim();
  if (t.length < 80) return "";

  // Strong signals
  const hasEssaySignal = /essay|personal statement|short answer|prompt|respond|write about|topic|question:/i.test(t);
  if (!hasEssaySignal) return "";

  // Try to pull a “prompt-like” chunk:
  // 1) If there is a "Prompt:" section
  const m1 = t.match(/(?:^|\n)\s*(prompt|essay prompt|writing prompt)\s*[:\-]\s*([\s\S]{40,1200})/i);
  if (m1 && m1[2]) return m1[2].trim();

  // 2) If there is a question mark cluster, take surrounding text
  const qIdx = t.search(/\?/);
  if (qIdx !== -1) {
    const start = Math.max(0, qIdx - 250);
    const end = Math.min(t.length, qIdx + 900);
    return t.slice(start, end).trim();
  }

  // 3) Fallback: first 900 chars if essay signal exists
  return t.slice(0, 900).trim();
}

function buildMainMenuDedupeSets_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MAIN_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${MAIN_SHEET}`);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const h = buildHeaderIndex(headers);
  const idxName = mustIndex(h, "Scholarship Name");
  const idxPortal = mustIndex(h, "Application Portal");

  const portal = new Set();
  const name = new Set();

  const lr = sh.getLastRow();
  if (lr >= HEADER_ROW + 1) {
    const vals = sh.getRange(HEADER_ROW + 1, 1, lr - HEADER_ROW, lastCol).getValues();
    vals.forEach(r => {
      const nm = String(r[idxName] || "").trim().toLowerCase();
      const po = String(r[idxPortal] || "").trim().toLowerCase();
      if (nm) name.add(nm);
      if (po) portal.add(po);
    });
  }

  return { portal, name };
}

function aiSidebarIngest(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    payload = payload || {};
    const url = (payload.url || "").toString().trim();
    const pastedText = (payload.pastedText || "").toString();
    const extraNotes = (payload.extraNotes || "").toString();
    const uploadedFile = payload.uploadedFile || null;

    // 1) Build sourceText priority: paste > upload text > URL > dataset URL
    let sourceText = (pastedText || "").trim();
    let uploadedUrl = "";

    if (!sourceText && uploadedFile && uploadedFile.driveFileId) {
      uploadedUrl = uploadedFile.driveUrl || "";
      sourceText = extractTextFromDriveFile_(uploadedFile.driveFileId, uploadedFile.mimeType);
    }

    if (!sourceText && url) {
      sourceText = fetchUrlAsTextSmart_(url) || "";
    }

    // Dataset fallback (sheet/csv)
    if (!sourceText && url) {
      sourceText = extractTextFromDatasetUrl_(url) || "";
    }

    if (!sourceText || sourceText.trim().length < 40) {
      throw new Error("Need more text. Paste text, upload a PDF, or provide a working URL.");
    }

    // 2) Batch extract (with chunking if needed)
    const meta = { url, uploadedUrl, extraNotes };

    const chunks = chunkText_(sourceText, 12000);
    let all = [];
    for (let i = 0; i < chunks.length; i++) {
      const batch = geminiExtractScholarshipsBatch_(chunks[i], Object.assign({}, meta, { chunk_index: i + 1, chunks_total: chunks.length }));
      all = all.concat(batch.items || []);
    }
    const items = mergeBatchItems_(all);

    // 3) If URL exists and some items are missing critical fields, do one URL “assist” pass
    // (lightweight: only if URL fetch succeeds and missing fields exist)
    if (url) {
      const missingAny = items.some(it =>
        (!it.due_date || it.due_date.length < 4) ||
        (!it.portal_url || it.portal_url.length < 8)
      );
      if (missingAny) {
        const urlText = fetchUrlAsTextSmart_(url);
        if (urlText && urlText.length > 200) {
          const assist = geminiExtractScholarshipsBatch_(urlText.slice(0, 12000), Object.assign({}, meta, { assist_pass: true }));
          const mergedAssist = mergeBatchItems_(items.concat(assist.items || []));
          // Prefer original items, but fill blanks from assist
          const filled = mergedAssist.map(it => it); // already merged; ok for most docs
          items.length = 0;
          Array.prototype.push.apply(items, filled);
        }
      }
    }

    // 4) Build sheet dedupe sets once
    const dedupe = buildMainMenuDedupeSets_();

    // 5) Apply gates + append
    const results = {
      action: "BATCH_DONE",
      total_extracted: items.length,
      appended: 0,
      appended_rows: [],
      skipped: {
        past_due: [],
        not_eligible: [],
        duplicate_portal: [],
        duplicate_name: [],
        invalid: []
      }
    };

    items.forEach(it => {
      const nm = String(it.name || "").trim();
      const po = String(it.portal_url || url || "").trim();

      if (!nm || nm.length < 3) {
        results.skipped.invalid.push({ name: nm, portal: po, reason: "missing_name" });
        return;
      }

      const dueObj = parseDueDate_(it.due_date || "");
      if (dueObj && isPastDue_(dueObj)) {
        results.skipped.past_due.push({ name: nm, due_date: it.due_date || "" });
        return;
      }

      if (String(it.is_eligible || "unknown").toLowerCase() === "no") {
        results.skipped.not_eligible.push({ name: nm, reason: it.eligibility || "" });
        return;
      }

      const portalKey = po.toLowerCase();
      const nameKey = nm.toLowerCase();

      if (po && dedupe.portal.has(portalKey)) {
        results.skipped.duplicate_portal.push({ name: nm, portal: po });
        return;
      }
      if (dedupe.name.has(nameKey)) {
        results.skipped.duplicate_name.push({ name: nm });
        return;
      }

      const writeRes = upsertScholarshipToMainMenu_(it, {
        url,
        uploadedUrl,
        extraNotes,
        sourceText: "" // don’t store full mega text
      });

      if (writeRes.action === "APPENDED") {
        results.appended++;
        results.appended_rows.push(writeRes.row);

        // Update dedupe sets so duplicates in same ingest don't pass
        if (po) dedupe.portal.add(portalKey);
        dedupe.name.add(nameKey);
      } else {
        // upsert function already dedupes; classify roughly
        if (writeRes.action === "SKIP_DUPLICATE_PORTAL") results.skipped.duplicate_portal.push({ name: nm, portal: po });
        else if (writeRes.action === "SKIP_DUPLICATE_NAME") results.skipped.duplicate_name.push({ name: nm });
        else results.skipped.invalid.push({ name: nm, portal: po, reason: writeRes.action });
      }
    });

    return results;
  } finally {
    lock.releaseLock();
  }
}

function aiSidebarChat(payload) {
  payload = payload || {};
  const msg = String(payload.message || "").trim();
  if (!msg) return { type: "reply", text: "" };

  const tools = getAutoshipToolRegistry_();
  const plan = geminiDecideChatOrTool_(msg, tools);

  if (!plan || plan.type === "reply") {
    return { type: "reply", text: String(plan?.text || "") };
  }

  if (plan.type === "choose_recommenders") {
    return {
      type: "choose_recommenders",
      text: String(plan.text || "Select which recommender(s) to draft to:"),
      row: Number(plan.row || 0),
      options: Array.isArray(plan.options) ? plan.options : listRecommenders_()
    };
  }

  if (plan.type === "tool_call") {
    const toolName = String(plan.tool || "");
    const args = plan.args || {};
    const execRes = executeAutoshipTool_(toolName, args, tools);
    const finalText = geminiSummarizeToolResult_(msg, toolName, args, execRes);
    return { type: "reply", text: finalText };
  }

  return { type: "reply", text: "I couldn't determine an action." };
}

function geminiDecideChatOrTool_(userMsg, tools) {
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
      type: { type: "string", enum: ["reply", "tool_call", "choose_recommenders"] },
      text: { type: "string" },
      tool: { type: "string" },
      args: { type: "object" },
      row: { type: "number" },
      options: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            name: { type: "string" },
            email: { type: "string" },
            role: { type: "string" }
          },
          required: ["name", "email", "role"]
        }
      }
    },
    required: ["type"]
  };

  const toolList = tools.map(t => ({
    name: t.name,
    description: t.description,
    argsSchema: t.argsSchema
  }));

  const recs = listRecommenders_();
  const prompt = [
    "You are Autoship Sidebar AI. Decide whether to answer normally, run a tool, or ask the user to choose recommenders.",
    "If the user requests drafting recommendation emails, respond with type=choose_recommenders and include options from the recommender list.",
    "Only include recommenders that exist in the list.",
    "",
    "Recommender list:",
    JSON.stringify(recs),
    "",
    "Tool list:",
    JSON.stringify(toolList),
    "",
    "Return ONLY JSON matching this schema:",
    JSON.stringify(schema),
    "",
    "User message:",
    userMsg
  ].join("\n");

  const body = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.1 }
  };

  const res = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    return { type: "reply", text: "AI tool router error." };
  }

  const raw = JSON.parse(res.getContentText() || "{}");
  const text = (raw?.candidates?.[0]?.content?.parts || []).map(p => p.text || "").join("").trim();
  if (!text) return { type: "reply", text: "" };

  try {
    return JSON.parse(text);
  } catch (e) {
    return { type: "reply", text: text };
  }
}

function executeAutoshipTool_(toolName, args, tools) {
  const reg = tools.find(t => t.name === toolName);
  if (!reg) throw new Error("Tool not allowed: " + toolName);

  // Minimal safety: args must be an object
  if (args === null || typeof args !== "object" || Array.isArray(args)) {
    throw new Error("Invalid args for tool: " + toolName);
  }

  // Execute by name (allowlisted)
  const fn = this[toolName];
  if (typeof fn !== "function") throw new Error("Tool function not found: " + toolName);

  // Tools can be no-arg or single-arg (your aiSidebarIngest takes payload)
  try {
    const arity = fn.length;
    if (arity === 0) return fn();
    return fn(args);
  } catch (e) {
    return { error: String(e && e.message ? e.message : e) };
  }
}

function geminiSummarizeToolResult_(userMsg, toolName, args, toolResult) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return JSON.stringify(toolResult, null, 2);

  const endpoint =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(GEMINI_MODEL) +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  const prompt = [
    "You are Autoship Sidebar AI.",
    "Explain the result of running a scholarship tool clearly and briefly.",
    "If there is an error, tell the user what to do next.",
    "",
    "User message:",
    userMsg,
    "",
    "Tool executed:",
    toolName,
    "Args:",
    JSON.stringify(args),
    "",
    "Tool result:",
    JSON.stringify(toolResult)
  ].join("\n");

  const body = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2 }
  };

  const res = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    return JSON.stringify(toolResult, null, 2);
  }

  const raw = JSON.parse(res.getContentText() || "{}");
  const text = (raw?.candidates?.[0]?.content?.parts || []).map(p => p.text || "").join("").trim();
  return text || JSON.stringify(toolResult, null, 2);
}

function geminiSummarizeToolResult_(userMsg, toolName, args, toolResult) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return JSON.stringify(toolResult, null, 2);

  const endpoint =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(GEMINI_MODEL) +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  const prompt = [
    "You are Autoship Sidebar AI.",
    "Explain the result of running a scholarship tool clearly and briefly.",
    "If there is an error, tell the user what to do next.",
    "",
    "User message:",
    userMsg,
    "",
    "Tool executed:",
    toolName,
    "Args:",
    JSON.stringify(args),
    "",
    "Tool result:",
    JSON.stringify(toolResult)
  ].join("\n");

  const body = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2 }
  };

  const res = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    return JSON.stringify(toolResult, null, 2);
  }

  const raw = JSON.parse(res.getContentText() || "{}");
  const text = (raw?.candidates?.[0]?.content?.parts || []).map(p => p.text || "").join("").trim();
  return text || JSON.stringify(toolResult, null, 2);
}

function geminiExtractScholarshipsBatch_(sourceText, meta) {
  meta = meta || {};
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("Missing Script Property: GEMINI_API_KEY");

  const endpoint =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(GEMINI_MODEL) +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);

  const itemSchema = {
    type: "object",
    additionalProperties: false,
    properties: {
      name: { type: "string" },
      portal_url: { type: "string" },
      due_date: { type: "string" },
      amount: { type: "string" },
      eligibility: { type: "string" },
      is_eligible: { type: "string", enum: ["yes", "no", "unknown"] },
      requirements: { type: "array", items: { type: "string" } },
      theme: { type: "array", items: { type: "string" } },
      notes: { type: "string" }
    },
    required: [
      "name", "portal_url", "due_date", "amount", "eligibility",
      "is_eligible", "requirements", "theme", "notes"
    ]
  };

  const schema = {
    type: "object",
    additionalProperties: false,
    properties: {
      items: { type: "array", items: itemSchema },
      notes: { type: "string" }
    },
    required: ["items", "notes"]
  };

  const prompt = [
    "Extract ALL scholarships from the source text.",
    "Return ONLY valid JSON (no markdown).",
    "",
    "Rules:",
    "- Output must match schema exactly.",
    "- Each scholarship should be one item in items[].",
    "- If portal_url missing, use meta.url if it clearly applies; otherwise empty.",
    "- due_date: use exact date if present; else empty.",
    "- requirements/theme: short tags (sheet dropdown vocabulary if possible).",
    "- is_eligible: decide based on eligibility criteria vs applicant profile below.",
    "- If one document lists many scholarships, include them all.",
    "",
    "Applicant profile:",
    "- US high school senior class of 2026 in Arlington, TX",
    "",
    "Meta:",
    JSON.stringify(meta),
    "",
    "JSON Schema:",
    JSON.stringify(schema),
    "",
    "SOURCE TEXT:",
    (sourceText || "").slice(0, 12000)
  ].join("\n");

  const body = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
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
  if (code < 200 || code >= 300) throw new Error("Gemini batch extraction failed: HTTP " + code);

  const parsed = JSON.parse(raw);
  const text = (parsed?.candidates?.[0]?.content?.parts || []).map(p => p.text || "").join("").trim();
  if (!text) throw new Error("Gemini returned empty batch extraction.");

  const obj = JSON.parse(text);
  obj.items = Array.isArray(obj.items) ? obj.items : [];
  obj.items.forEach(it => {
    it.requirements = uniq_((it.requirements || []).map(String));
    it.theme = uniq_((it.theme || []).map(String));
    it.name = String(it.name || "").trim();
    it.portal_url = String(it.portal_url || "").trim();
    it.due_date = String(it.due_date || "").trim();
    it.amount = String(it.amount || "").trim();
    it.eligibility = String(it.eligibility || "").trim();
    it.is_eligible = String(it.is_eligible || "unknown").toLowerCase();
    it.notes = String(it.notes || "").trim();
  });

  return obj;
}

function chunkText_(text, maxChars) {
  text = String(text || "");
  maxChars = Math.max(2000, Number(maxChars) || 12000);

  const chunks = [];
  let i = 0;
  while (i < text.length) {
    const end = Math.min(i + maxChars, text.length);
    chunks.push(text.slice(i, end));
    i = end;
  }
  return chunks;
}

function mergeBatchItems_(allItems) {
  // Dedupe within batch by portal_url then name (case-insensitive)
  const byPortal = new Map();
  const byName = new Map();
  const out = [];

  allItems.forEach(it => {
    const portalKey = (it.portal_url || "").trim().toLowerCase();
    const nameKey = (it.name || "").trim().toLowerCase();
    if (portalKey && byPortal.has(portalKey)) return;
    if (!portalKey && nameKey && byName.has(nameKey)) return;

    out.push(it);
    if (portalKey) byPortal.set(portalKey, true);
    if (nameKey) byName.set(nameKey, true);
  });

  return out;
}
function fetchUrlAsTextSmart_(url) {
  const u = String(url || "").trim();
  if (!u) return "";
  const lower = u.toLowerCase();
  if (lower.endsWith(".pdf") || lower.includes(".pdf?")) return fetchPdfAsTextViaDriveConvert_(u);
  return fetchHtmlAsText_(u);
}

function extractTextFromDatasetUrl_(url) {
  const u = String(url || "").trim();
  if (!u) return "";

  // Google Sheet
  if (/docs\.google\.com\/spreadsheets\/d\//i.test(u)) {
    const id = extractDriveIdFromUrl_(u);
    if (!id) return "";
    const ss = SpreadsheetApp.openById(id);
    const sh = ss.getSheets()[0];
    const values = sh.getDataRange().getDisplayValues();
    // Turn first N rows into a compact text blob
    const maxRows = Math.min(values.length, 200);
    const maxCols = Math.min(values[0].length, 20);
    const lines = [];
    for (let r = 0; r < maxRows; r++) {
      lines.push(values[r].slice(0, maxCols).join(" | "));
    }
    return lines.join("\n");
  }

  // CSV (basic)
  if (u.toLowerCase().includes(".csv")) {
    const res = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return "";
    const csv = res.getContentText() || "";
    const rows = Utilities.parseCsv(csv);
    const maxRows = Math.min(rows.length, 200);
    const maxCols = Math.min(rows[0].length, 20);
    return rows.slice(0, maxRows).map(r => r.slice(0, maxCols).join(" | ")).join("\n");
  }

  return "";
}

function createEssayDocForRow_(sheet, absoluteRow, opts) {
  opts = opts || {};
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxReq     = optionalIndex_(h, COL_REQUIREMENTS);
  const idxPortal  = optionalIndex_(h, COL_APPLICATION_PORTAL);
  const idxAppFile = optionalIndex_(h, COL_APPLICATION_FILE);
  const idxName    = optionalIndex_(h, "Scholarship Name");

  const idxDocId   = optionalIndex_(h, "Essay Doc ID");
  const idxStatus  = optionalIndex_(h, "Status");
  const idxDue     = optionalIndex_(h, "Due Date");
  const idxTriage  = optionalIndex_(h, "Triage");

  if (idxReq === null || idxAppFile === null || idxName === null) {
    throw new Error(`Missing required columns on "${MAIN_SHEET}"`);
  }

  const lastCol = sheet.getLastColumn();
  const row = sheet.getRange(absoluteRow, 1, 1, lastCol).getValues()[0];
  const rich = sheet.getRange(absoluteRow, 1, 1, lastCol).getRichTextValues()[0];
  const formulas = sheet.getRange(absoluteRow, 1, 1, lastCol).getFormulas()[0];

  const reqText = String(row[idxReq] ?? "");
  const name = String(row[idxName] ?? "").trim() || "Scholarship";

  // Skip if requirements doesn't include Essay(s), unless forced
  const needsEssay = requirementsIncludes_(reqText, ESSAY_REQUIREMENT_TOKEN);
  if (!needsEssay && !opts.force) return { action: "SKIP_NO_ESSAY_REQUIREMENT" };

  // Skip if already has a doc
  const appFileCell = sheet.getRange(absoluteRow, idxAppFile + 1);
  const existingUrl = extractDocUrlFromCell_(appFileCell);
  if (existingUrl) return { action: "SKIP_ALREADY_HAS_DOC", url: existingUrl };

  const portalUrl = (idxPortal !== null)
    ? extractBestUrlFromCell_(formulas[idxPortal], rich[idxPortal], String(row[idxPortal] ?? ""))
    : "";

  // Prompt extraction: use provided sourceText if present, else fetch portal
  let sourceText = String(opts.sourceText || "").trim();
  if ((!sourceText || sourceText.length < 40) && portalUrl) {
    sourceText = fetchUrlAsTextSmart_(portalUrl) || "";
  }

  const ep = (sourceText && sourceText.length >= 40)
    ? geminiExtractEssayPrompt_(sourceText.slice(0, 12000), portalUrl)
    : { prompt_text: "", word_limit: "", pdf_required: false };

  const docInfo = createEssayPrepDocOneDraft_({
    scholarshipName: name,
    portalUrl,
    promptText: ep.prompt_text || "",
    wordLimit: ep.word_limit || ""
  });

  if (idxDocId !== null && docInfo && docInfo.id) {
    sheet.getRange(absoluteRow, idxDocId + 1).setValue(docInfo.id);
  }

  appendLinkIntoCell_(sheet, absoluteRow, idxAppFile + 1, "Essay Draft (Google Doc)", docInfo.url);

  if (idxStatus !== null) sheet.getRange(absoluteRow, idxStatus + 1).setValue("Prepped");

  if (idxDue !== null && idxTriage !== null) {
    const due = parseSheetDate_(row[idxDue]);
    sheet.getRange(absoluteRow, idxTriage + 1).setValue(computeTriage_(due, row[idxTriage]));
  }

  return { action: "CREATED_DOC", url: docInfo.url, id: docInfo.id || "" };
}
/*******************************
 * CANVAS → AUTOSHIP INTAKE
 * Focus:
 *  - HEAVY: MHS Counselor (announcements + modules)
 *  - LIGHT: Class of 2026 (announcements + modules, fewer items)
 *******************************/

const CANVAS_INTAKE_SHEET = "Canvas Intake";

// Course matching (flexible): edit if your actual course names differ
const COURSE_MATCH = {
  // Exact match for your counselor course name
  COUNSELOR_EXACT: [/^MHS Counselors$/i],

  // Keep Class of 2026 broad unless you tell me the exact name
  CLASS_2026: [/class/i, /2026/i]
};

// Scholarship-ish keywords (expand whenever)
const SCHOLAR_KEYWORDS = [
  "scholarship", "award", "grant", "fellowship",
  "financial", "money", "tuition", "apply", "application",
  "deadline", "due", "senior", "class of", "$", "dollars",
  "local scholarship", "foundation", "community scholarship",
  "FAFSA", "TASFA"
];

// How “hard” to scan each course
const SCAN_PROFILE = {
  COUNSELOR: {
    announcementsLookbackDays: 120,
    modulesLookbackDays: 365,
    maxModuleItems: 250   // high
  },
  CLASS_2026: {
    announcementsLookbackDays: 45,
    modulesLookbackDays: 120,
    maxModuleItems: 60    // light
  }
};

/***************
 * ENTRY POINT
 ***************/
function scanCanvasForScholarships() {
  const courses = canvasListActiveCourses_();

  const counselorCourses = pickCourses_(courses, COURSE_MATCH.COUNSELOR_EXACT);
  const class2026Courses = pickCourses_(courses, COURSE_MATCH.CLASS_2026);


  const targets = [];
  counselorCourses.forEach(c => targets.push({ course: c, profile: "COUNSELOR" }));
  class2026Courses.forEach(c => targets.push({ course: c, profile: "CLASS_2026" }));

  if (!targets.length) {
    throw new Error("No matching Canvas courses found. Rename match patterns in COURSE_MATCH.");
  }

  ensureCanvasIntakeSheet_();

  // Load dedupe set from existing sheet
  const seen = loadSeenKeys_();

  // Process heavy first
  targets.sort((a, b) => (a.profile === "COUNSELOR" ? -1 : 1));

  const rowsToAppend = [];

  targets.forEach(t => {
    const { course, profile } = t;
    const cfg = SCAN_PROFILE[profile];

    // Announcements (always)
    const ann = canvasFetchAnnouncements_(course.id, cfg.announcementsLookbackDays);
    ann.forEach(a => {
      const hit = scholarshipHit_(a.title, a.message);
      if (!hit) return;

      const key = `announcement:${a.id}`;
      if (seen.has(key)) return;
      seen.add(key);

      rowsToAppend.push(canvasRow_({
        key,
        profile,
        course,
        type: "Announcement",
        title: a.title,
        body: stripHtml_(a.message),
        postedAt: a.posted_at,
        url: a.html_url || "",
        hit
      }));
    });

    // Modules + items
    const modItems = canvasFetchModuleItems_(course.id, cfg.modulesLookbackDays, cfg.maxModuleItems);
    modItems.forEach(mi => {
      // Cheap first pass: title-only filter to save quota
      const hitTitleOnly = scholarshipHit_(mi.title, "");
      if (!hitTitleOnly) return;

      // Deep fetch content from Canvas API (page/assignment/discussion/file)
      const deep = canvasDeepFetchTextForModuleItem_(course.id, mi);

      // Real test on fetched body
      const hit = scholarshipHit_(deep.title, deep.body);
      if (!hit) return;

      const key = `module_item:${mi.id}`;
      if (seen.has(key)) return;
      seen.add(key);

      rowsToAppend.push(canvasRow_({
        key,
        profile,
        course,
        type: `Module Item (${mi.type || "unknown"})`,
        title: deep.title,
        body: deep.body,
        postedAt: mi.updated_at || mi.published_at || "",
        url: deep.url,
        hit
      }));
    });
  }); // <-- Add this closing brace to end targets.forEach
}

/***********************
 * CANVAS API HELPERS
 ***********************/
function canvasBaseUrl_() {
  const v = PropertiesService.getScriptProperties().getProperty("CANVAS_BASE_URL");
  if (!v) throw new Error("Missing Script Property: CANVAS_BASE_URL");
  return v.replace(/\/+$/, "");
}
function canvasToken_() {
  const v = PropertiesService.getScriptProperties().getProperty("CANVAS_TOKEN");
  if (!v) throw new Error("Missing Script Property: CANVAS_TOKEN");
  return v.trim();
}

function canvasFetchJson_(path, params) {
  const base = canvasBaseUrl_();
  const url = buildUrl_(base + path, params || {});
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + canvasToken_() }
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Canvas API error ${code}: ${res.getContentText().slice(0, 200)}`);
  }
  return JSON.parse(res.getContentText() || "null");
}

function canvasListActiveCourses_() {
  // enrollment_state=active keeps it clean; per_page helps reduce pagination needs
  const data = canvasFetchJson_("/api/v1/courses", {
    enrollment_state: "active",
    per_page: 100
  });
  // data can include many; keep minimal fields
  return (data || []).map(c => ({
    id: c.id,
    name: c.name || c.course_code || `Course ${c.id}`
  }));
}

function canvasFetchAnnouncements_(courseId, lookbackDays) {
  const start = new Date(Date.now() - lookbackDays * 24 * 3600 * 1000).toISOString();
  const data = canvasFetchJson_("/api/v1/announcements", {
    "context_codes[]": `course_${courseId}`,
    start_date: start,
    per_page: 100
  });
  return (data || []).map(a => ({
    id: a.id,
    title: a.title || "(no title)",
    message: a.message || "",
    posted_at: a.posted_at || "",
    html_url: a.html_url || ""
  }));
}

function canvasFetchModuleItems_(courseId, lookbackDays, maxItems) {
  const cutoff = Date.now() - lookbackDays * 24 * 3600 * 1000;

  const modules = canvasFetchJson_(`/api/v1/courses/${courseId}/modules`, { per_page: 100 }) || [];

  const out = [];
  for (let i = 0; i < modules.length; i++) {
    if (out.length >= maxItems) break;

    const m = modules[i];
    const items = canvasFetchJson_(`/api/v1/courses/${courseId}/modules/${m.id}/items`, { per_page: 100 }) || [];

    for (let j = 0; j < items.length; j++) {
      if (out.length >= maxItems) break;

      const it = items[j];
      const updated = parseDate_(it.updated_at) || parseDate_(it.published_at);
      if (updated && updated.getTime() < cutoff) continue;

      out.push({
        id: it.id,
        title: it.title || "(no title)",
        type: it.type || "",               // <- key for deep fetch
        content_id: it.content_id || null, // <- key for deep fetch
        page_url: it.page_url || "",       // <- some items include this
        url: it.url || "",
        html_url: it.html_url || "",
        updated_at: it.updated_at || "",
        published_at: it.published_at || ""
      });
    }
  }
  return out;
}

/***********************
 * FILTERING + OUTPUT
 ***********************/
function scholarshipHit_(title, body) {
  const t = (title || "").toLowerCase();
  const b = (body || "").toLowerCase();

  // strong signals
  for (const k of SCHOLAR_KEYWORDS) {
    if (k === "$") {
      if (t.includes("$") || b.includes("$")) return "contains:$";
      continue;
    }
    if (t.includes(k) || b.includes(k)) return `contains:${k}`;
  }

  // quick extra pattern: “due DATE”
  if (/\bdue\b/.test(t + " " + b) && /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b/.test(t + " " + b)) {
    return "pattern:due+month";
  }

  return "";
}

function canvasRow_({ key, profile, course, type, title, body, postedAt, url, hit }) {
  return [
    new Date(),            // Imported At
    key,                   // Unique Key (dedupe)
    profile,               // COUNSELOR / CLASS_2026
    course.name,           // Course
    type,                  // Announcement / Module Item
    title || "",
    postedAt || "",
    url || "",
    hit || "",
    body || ""             // Snippet/body
  ];
}

function ensureCanvasIntakeSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CANVAS_INTAKE_SHEET);
  if (!sh) sh = ss.insertSheet(CANVAS_INTAKE_SHEET);

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "Imported At",
      "Unique Key",
      "Profile",
      "Course",
      "Type",
      "Title",
      "Posted/Updated At",
      "URL",
      "Hit",
      "Snippet/Body"
    ]);
    sh.setFrozenRows(1);
  }
}

function appendCanvasRows_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CANVAS_INTAKE_SHEET);
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function loadSeenKeys_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CANVAS_INTAKE_SHEET);
  const seen = new Set();
  if (!sh || sh.getLastRow() < 2) return seen;

  // Unique Key is column 2
  const values = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  values.forEach(r => {
    const k = (r[0] || "").toString().trim();
    if (k) seen.add(k);
  });
  return seen;
}

/***********************
 * UTIL
 ***********************/
function pickCourses_(courses, regexList) {
  return (courses || []).filter(c => {
    const name = (c.name || "").toString();
    return regexList.every(rx => rx.test(name));
  });
}

function buildUrl_(base, params) {
  const q = Object.keys(params).map(k => {
    const v = params[k];
    return encodeURIComponent(k) + "=" + encodeURIComponent(v);
  }).join("&");
  return q ? `${base}?${q}` : base;
}

function stripHtml_(html) {
  if (!html) return "";
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseDate_(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function basicExtractFromText_(text, url) {
  const t = (text || "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const firstLine =
    (t.split(/\r?\n/)
      .map(s => s.trim())
      .filter(Boolean)[0] || "").slice(0, 80);

  const dueMatch = t.match(
    /\b(due|deadline)\b[:\s-]*([^\n]{0,60})/i
  );

  const requirements =
    /essay|personal statement|short answer/i.test(t)
      ? "Essay(s)"
      : "";

  return {
    name: firstLine || "Scholarship",
    portal: url || "",
    due: dueMatch ? dueMatch[2].trim() : "",
    requirements,
    difficulty: "",
    notes: ""
  };
}

function pipeCanvasHitsToMainMenu() {
  const ss = SpreadsheetApp.getActive();

  const intake = ss.getSheetByName("Canvas Intake");
  if (!intake) throw new Error('Missing sheet: "Canvas Intake"');

  const main = ss.getSheetByName(MAIN_SHEET); // "Main Menu" from your config
  if (!main) throw new Error(`Missing sheet: "${MAIN_SHEET}"`);

  // --- Read headers ---
  const intakeLastCol = intake.getLastColumn();
  const intakeHeaders = intake.getRange(1, 1, 1, intakeLastCol).getValues()[0];
  const ih = buildHeaderIndex(intakeHeaders);

  const idxImportedAt = optionalIndex_(ih, "Imported At");
  const idxKey = mustIndex(ih, "Unique Key");
  const idxProfile = optionalIndex_(ih, "Profile");
  const idxCourse = optionalIndex_(ih, "Course");
  const idxType = optionalIndex_(ih, "Type");
  const idxTitle = mustIndex(ih, "Title");
  const idxPosted = optionalIndex_(ih, "Posted/Updated At");
  const idxUrl = mustIndex(ih, "URL");
  const idxBody = optionalIndex_(ih, "Snippet/Body");

  // Ensure "Piped" column exists on intake
  let idxPiped = optionalIndex_(ih, "Piped");
  if (idxPiped === null) {
    intake.getRange(1, intakeLastCol + 1).setValue("Piped");
    idxPiped = intakeLastCol; // 0-based index for new column
  }

  const intakeLastRow = intake.getLastRow();
  if (intakeLastRow < 2) return;

  const intakeValues = intake.getRange(2, 1, intakeLastRow - 1, intake.getLastColumn()).getValues();

  // --- Main Menu header map ---
  const mainLastCol = main.getLastColumn();
  const mainHeaders = main.getRange(HEADER_ROW, 1, 1, mainLastCol).getValues()[0];
  const mh = buildHeaderIndex(mainHeaders);

  const mName = mustIndex(mh, "Scholarship Name");
  const mPortal = mustIndex(mh, "Application Portal");
  const mDue = optionalIndex_(mh, "Due Date");
  const mReq = optionalIndex_(mh, "Requirements");
  const mNotes = optionalIndex_(mh, "Notes");
  const mTriage = optionalIndex_(mh, "Triage");
  const mDifficulty = optionalIndex_(mh, "Difficulty");

  // Build a quick dedupe set from Main Menu notes + portal
  const existing = new Set();
  const mainLastRow = main.getLastRow();
  if (mainLastRow >= 2) {
    const portalCol = mPortal + 1;
    const notesCol = mNotes !== null ? (mNotes + 1) : null;

    const portals = main.getRange(2, portalCol, mainLastRow - 1, 1).getValues();
    portals.forEach(r => {
      const u = (r[0] || "").toString().trim();
      if (u) existing.add("portal:" + u);
    });

    if (notesCol) {
      const notes = main.getRange(2, notesCol, mainLastRow - 1, 1).getValues();
      notes.forEach(r => {
        const n = (r[0] || "").toString();
        const m = n.match(/\[CanvasKey:([^\]]+)\]/i);
        if (m && m[1]) existing.add("key:" + m[1].trim());
      });
    }
  }

  const rowsToAppend = [];
  const pipedUpdates = []; // [rowIndexWithinIntakeValues, "YES"]

  for (let r = 0; r < intakeValues.length; r++) {
    const row = intakeValues[r];

    const alreadyPiped = String(row[idxPiped] || "").trim();
    if (alreadyPiped) continue;

    const key = String(row[idxKey] || "").trim();
    const title = String(row[idxTitle] || "").trim();
    const url = String(row[idxUrl] || "").trim();
    const body = idxBody !== null ? String(row[idxBody] || "") : "";

    if (!title || !url || !key) continue;

    // hard dedupe: key + portal
    if (existing.has("key:" + key) || existing.has("portal:" + url)) {
      pipedUpdates.push([r, "DUPLICATE"]);
      continue;
    }

    const extracted = basicExtractFromText_(body, url);
    // --- Essay prompt detection + doc creation ---
  let essayDocUrl = "";
  const promptText = detectEssayPrompt_(body);

  if (promptText && promptText.length >= 80) {
    essayDocUrl = createEssayDoc_(
      extracted.name || title || "Scholarship",
      url,
      promptText
    );
  }

    const newRow = new Array(mainLastCol).fill("");
    newRow[mName] = extracted.name || title || "Scholarship";
    newRow[mPortal] = url;
    newRow[mNotes] = essayDocUrl ? `[Essay Doc: ${essayDocUrl}]` : "";

    if (mDue !== null) newRow[mDue] = extracted.due || "";
    if (mReq !== null) newRow[mReq] = extracted.requirements || "";
    if (mDifficulty !== null) newRow[mDifficulty] = extracted.difficulty || "";
    if (mTriage !== null) newRow[mTriage] = ""; // let refreshTriageForMainMenu compute it

    if (mNotes !== null) {
      const course = idxCourse !== null ? String(row[idxCourse] || "").trim() : "";
      const type = idxType !== null ? String(row[idxType] || "").trim() : "";
      const posted = idxPosted !== null ? String(row[idxPosted] || "").trim() : "";
      const profile = idxProfile !== null ? String(row[idxProfile] || "").trim() : "";
      const importedAt = idxImportedAt !== null ? row[idxImportedAt] : "";

      const meta = [
        `[CanvasKey:${key}]`,
        profile ? `Profile=${profile}` : "",
        course ? `Course=${course}` : "",
        type ? `Type=${type}` : "",
        posted ? `Posted=${posted}` : "",
        importedAt ? `Imported=${importedAt}` : ""
      ].filter(Boolean).join(" | ");

      const snippet = body ? `\n\nCanvas Snippet: ${body}` : "";
      newRow[mNotes] = meta + snippet;
    }

    rowsToAppend.push(newRow);
    pipedUpdates.push([r, "YES"]);

    existing.add("key:" + key);
    existing.add("portal:" + url);
  }

  if (rowsToAppend.length) {
    main.getRange(main.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  if (pipedUpdates.length) {
    // Write back "Piped" status
    const pipedCol = idxPiped + 1; // 1-based for Range
    const updatesRange = intake.getRange(2, pipedCol, intakeValues.length, 1);
    const colVals = updatesRange.getValues();

    pipedUpdates.forEach(([i, val]) => colVals[i][0] = val);
    updatesRange.setValues(colVals);
  }

  // Optional: refresh triage after adding
  // refreshTriageForMainMenu();

  ss.toast(`Piped ${rowsToAppend.length} Canvas hit(s) into Main Menu.`, "Autoship", 6);
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

function upsertScholarshipToMainMenu_(extracted, ctx) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MAIN_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${MAIN_SHEET}`);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxName = mustIndex(h, "Scholarship Name");
  const idxPortal = mustIndex(h, "Application Portal");
  const idxDue = optionalIndex_(h, "Due Date");
  const idxReq = optionalIndex_(h, "Requirements");
  const idxTheme = optionalIndex_(h, COL_THEME);
  const idxNotes = optionalIndex_(h, "Notes");
  const idxStatus = optionalIndex_(h, "Status");
  const idxTriage = optionalIndex_(h, "Triage");
  const idxStart = optionalIndex_(h, "Start date");
  const idxDifficulty = optionalIndex_(h, "Difficulty");

  // Build dedupe set from existing portal + name
  const existingPortal = new Set();
  const existingName = new Set();
  const lr = sh.getLastRow();
  if (lr >= HEADER_ROW + 1) {
    const vals = sh.getRange(HEADER_ROW + 1, 1, lr - HEADER_ROW, lastCol).getValues();
    vals.forEach(r => {
      const nm = String(r[idxName] || "").trim().toLowerCase();
      const po = String(r[idxPortal] || "").trim().toLowerCase();
      if (nm) existingName.add(nm);
      if (po) existingPortal.add(po);
    });
  }

  const portal = String(extracted.portal_url || ctx.url || "").trim();
  const name = String(extracted.name || "Scholarship").trim();

  if (portal && existingPortal.has(portal.toLowerCase())) return { action: "SKIP_DUPLICATE_PORTAL" };
  if (name && existingName.has(name.toLowerCase())) return { action: "SKIP_DUPLICATE_NAME" };

  const row = new Array(lastCol).fill("");
  row[idxName] = name;
  row[idxPortal] = portal;

  if (idxDue !== null && extracted.due_date) row[idxDue] = extracted.due_date;

  if (idxReq !== null) {
    // You already have detectRequirementsFromText_() + joinMultiSelect_()
    const req = uniq_([].concat(extracted.requirements || []));
    row[idxReq] = joinMultiSelect_(req);
  }

  if (idxTheme !== null) {
    const th = uniq_([].concat(extracted.theme || []));
    row[idxTheme] = joinMultiSelect_(th);
  }

  if (idxNotes !== null) {
    const sources = []
      .concat(ctx.url ? ["URL=" + ctx.url] : [])
      .concat(ctx.uploadedUrl ? ["Upload=" + ctx.uploadedUrl] : [])
      .filter(Boolean)
      .join(" | ");

    const eligibility = extracted.eligibility ? ("Eligibility: " + extracted.eligibility) : "";
    const amount = extracted.amount ? ("Amount: " + extracted.amount) : "";
    const extra = (ctx.extraNotes || "").trim();

    row[idxNotes] = [sources, amount, eligibility, extra].filter(Boolean).join("\n");
  }

  if (idxStart !== null) row[idxStart] = new Date();
  if (idxDifficulty !== null) row[idxDifficulty] = "No Idea...";
  if (idxStatus !== null) row[idxStatus] = "Scanned";

   if (idxTriage !== null) {
    const dueObj = parseDueDate_(row[idxDue]);
    row[idxTriage] = computeTriage_(dueObj, "");
  }

  sh.appendRow(row);
  const addedRow = sh.getLastRow();

  // === AUTO: create essay doc after append (only if prompt exists) ===
  autoCreateEssayDocForRow_(sh, addedRow, extracted, ctx);

  // Make portal clickable
  if (portal && looksLikeUrl_(portal)) {
    const rt = SpreadsheetApp.newRichTextValue().setText(portal).setLinkUrl(portal).build();
    sh.getRange(addedRow, idxPortal + 1).setRichTextValue(rt);
  }

  return { action: "APPENDED", row: addedRow };
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

function deleteEssayDocForRow_(sheet, absoluteRow) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxAppFile = optionalIndex_(h, COL_APPLICATION_FILE);
  const idxDocId   = optionalIndex_(h, "Essay Doc ID");
  if (idxAppFile === null) throw new Error("Missing Application File column");

  let docId = "";
  if (idxDocId !== null) {
    docId = String(sheet.getRange(absoluteRow, idxDocId + 1).getDisplayValue() || "").trim();
  }

  if (!docId) {
    const url = extractDocUrlFromCell_(sheet.getRange(absoluteRow, idxAppFile + 1)) || "";
    docId = extractDriveIdFromUrl_(url) || "";
  }

  if (!docId) return { action: "SKIP_NO_DOC_FOUND" };

  DriveApp.getFileById(docId).setTrashed(true);

  // Clear cells
  sheet.getRange(absoluteRow, idxAppFile + 1).clearContent();
  if (idxDocId !== null) sheet.getRange(absoluteRow, idxDocId + 1).clearContent();

  return { action: "TRASHED_DOC", id: docId };
}

function deleteEssayDocForRowTool(args) {
  args = args || {};
  const row = Number(args.row);
  if (!row || row <= HEADER_ROW) throw new Error("Provide a valid row > header.");

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(MAIN_SHEET);
  if (!sheet) throw new Error(`Missing sheet: ${MAIN_SHEET}`);

  return deleteEssayDocForRow_(sheet, row);
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
 * *Email Drafter
 * ***********************************************/

function getRecommenders_() {
  const raw = PropertiesService.getScriptProperties().getProperty("RECOMMENDERS_JSON") || "[]";
  try { return JSON.parse(raw); } catch (e) { return []; }
}

function listRecommenders_() {
  const raw = PropertiesService.getScriptProperties().getProperty("RECOMMENDERS_JSON") || "[]";
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function draftRecEmailsForRow_(sheet, absoluteRow) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxName   = mustIndex(h, "Scholarship Name");
  const idxPortal = optionalIndex_(h, COL_APPLICATION_PORTAL);
  const idxDue    = optionalIndex_(h, "Due Date");

  const row = sheet.getRange(absoluteRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const scholarship = String(row[idxName] || "").trim() || "Scholarship";
  const portal = (idxPortal !== null) ? String(row[idxPortal] || "").trim() : "";
  const due = (idxDue !== null) ? String(row[idxDue] || "").trim() : "";

  const recs = getRecommenders_();
  if (!recs.length) return { action: "SKIP_NO_RECOMMENDERS_CONFIGURED" };

  const subject = `Recommendation letter request: ${scholarship}`;

  const draftIds = [];
  recs.forEach(r => {
    if (!r || !r.email) return;

    const name = String(r.name || "").trim() || "there";
    const role = String(r.role || "").trim();

    const body =
`Dear ${name}${role ? " (" + role + ")" : ""},

I’m applying to ${scholarship} and would like to request a recommendation letter.

Due date: ${due || "N/A"}

If you’re able to support this, I can send any required forms and a short brag sheet immediately.

Thank you so much for your support,
Kevin Srun`;

    const draft = GmailApp.createDraft(String(r.email).trim(), subject, body);
    draftIds.push(draft.getId());
  });

  return { action: "DRAFTED", count: draftIds.length, draftIds };
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

function getAutoshipToolRegistry_() {
  return [
    {
      name: "refreshTriageForMainMenu",
      description: "Recomputes triage values for the Main Menu sheet.",
      argsSchema: { type: "object", additionalProperties: false, properties: {} }
    },
    {
      name: "createEssayDocsForSelection",
      description: "Creates application file docs for selected rows (manual tool).",
      argsSchema: { type: "object", additionalProperties: false, properties: {} }
    },
    {
      name: "fillEssayPromptsForSelection",
      description: "Uses AI to fill essay prompts for selected rows (slow).",
      argsSchema: { type: "object", additionalProperties: false, properties: {} }
    },
    {
      name: "deleteEssayDocForRowTool",
      description: "Delete the essay Google Doc for a specific Main Menu row.",
      argsSchema: {
        type: "object",
        additionalProperties: false,
        properties: {
          row: { type: "number" }
        },
        required: ["row"]
      }
    },
    {
      name: "syncSatisfiedToResults",
      description: "Sync Satisfied scholarships to Results sheet.",
      argsSchema: { type: "object", additionalProperties: false, properties: {} }
    },
    {
      name: "syncSurrenderedToSheet",
      description: "Sync Surrendered list to Surrendered sheet.",
      argsSchema: { type: "object", additionalProperties: false, properties: {} }
    },
    {
      name: "draftRecEmailsForRowToRecipientsTool",
      description: "Create Gmail drafts for selected recommenders for a given Main Menu row.",
      argsSchema: {
        type: "object",
        additionalProperties: false,
        properties: {
          row: { type: "number" },
          recipients: { type: "array", items: { type: "string" } }
        },
        required: ["row", "recipients"]
      }
    },
    {
      name: "aiSidebarIngest",
      description: "Ingest scholarship dataset input (URL/paste/upload) and add to Main Menu.",
      argsSchema: {
        type: "object",
        additionalProperties: false,
        properties: {
          url: { type: "string" },
          pastedText: { type: "string" },
          extraNotes: { type: "string" },
          uploadedFile: { type: ["object", "null"] }
        }
      }
    }
  ];
}

function geminiExtractEssayPrompts_(sourceText, meta) {
  meta = meta || {};
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
      has_essay: { type: "string", enum: ["yes", "no", "unknown"] },
      prompts: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            prompt: { type: "string" },
            word_limit: { type: "string" },
            notes: { type: "string" }
          },
          required: ["prompt", "word_limit", "notes"]
        }
      },
      overall_notes: { type: "string" }
    },
    required: ["has_essay", "prompts", "overall_notes"]
  };

  const prompt = [
    "Extract scholarship essay prompts from the SOURCE TEXT.",
    "Return ONLY valid JSON (no markdown).",
    "",
    "Rules:",
    "- has_essay = yes if any essay/personal statement/short answer prompt is required or optional.",
    "- prompts[] should include each distinct prompt found (1+).",
    "- word_limit: put the number if stated (e.g., '500'), else empty.",
    "- If prompt is implied but not explicitly stated, set has_essay=unknown and keep prompts empty.",
    "",
    "Meta:",
    JSON.stringify(meta),
    "",
    "JSON Schema:",
    JSON.stringify(schema),
    "",
    "SOURCE TEXT:",
    (sourceText || "").slice(0, 12000)
  ].join("\n");

  const body = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
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
  if (code < 200 || code >= 300) throw new Error("Essay prompt extraction failed: HTTP " + code);

  const parsed = JSON.parse(raw);
  const text = (parsed?.candidates?.[0]?.content?.parts || []).map(p => p.text || "").join("").trim();
  if (!text) throw new Error("Gemini returned empty essay extraction.");

  const obj = JSON.parse(text);
  obj.prompts = Array.isArray(obj.prompts) ? obj.prompts : [];
  obj.has_essay = String(obj.has_essay || "unknown").toLowerCase();
  obj.overall_notes = String(obj.overall_notes || "").trim();

  obj.prompts = obj.prompts
    .map(p => ({
      prompt: String(p.prompt || "").trim(),
      word_limit: String(p.word_limit || "").trim(),
      notes: String(p.notes || "").trim()
    }))
    .filter(p => p.prompt && p.prompt.length >= 10);

  return obj;
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

/***********************
 * MAIN MENU → GOOGLE CALENDAR
 ***********************/

const CAL_EVENT_COL_NAME = "Calendar Event Id";

function syncScholarshipDeadlinesToCalendar() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MAIN_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${MAIN_SHEET}`);

  const calId = PropertiesService.getScriptProperties()
    .getProperty("SCHOLARSHIP_CALENDAR_ID") || "primary";
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) throw new Error(`Could not open calendar: ${calId}`);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxName = mustIndex(h, "Scholarship Name");
  const idxPortal = mustIndex(h, "Application Portal");
  const idxDue = mustIndex(h, "Due Date");

  // Optional
  const idxTriage = optionalIndex_(h, "Triage");
  const idxStatus = optionalIndex_(h, "Status");
  const idxNotes = optionalIndex_(h, "Notes");

  let idxEvent = optionalIndex_(h, CAL_EVENT_COL_NAME);
  if (idxEvent === null) {
    sh.getRange(HEADER_ROW, lastCol + 1).setValue(CAL_EVENT_COL_NAME);
    idxEvent = lastCol; // 0-based
  }

  const lastRow = sh.getLastRow();
  if (lastRow < HEADER_ROW + 1) return;

  const dataRange = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sh.getLastColumn());
  const data = dataRange.getValues();

  let created = 0, updated = 0, removed = 0, skipped = 0;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    const name = String(row[idxName] || "").trim();
    const portal = String(row[idxPortal] || "").trim();
    const dueRaw = row[idxDue];
    let eventId = String(row[idxEvent] || "").trim();

    // Decide status-aware removal
    const triage = idxTriage !== null ? String(row[idxTriage] || "").trim().toLowerCase() : "";
    const status = idxStatus !== null ? String(row[idxStatus] || "").trim().toLowerCase() : "";

    const shouldRemove =
      (triage && REMOVE_EVENT_TRIAGE_VALUES.includes(triage)) ||
      (status && REMOVE_EVENT_STATUS_VALUES.includes(status));

    // If marked submitted/surrendered: delete calendar event and clear id
    if (shouldRemove) {
      if (eventId) {
        const ev = safeGetEventById_(cal, eventId);
        if (ev) {
          ev.deleteEvent();
          removed++;
        }
        row[idxEvent] = ""; // clear stored id either way
        eventId = "";
      } else {
        skipped++;
      }
      continue;
    }

    // Need minimum fields to create/update
    if (!name || !portal || !dueRaw) {
      skipped++;
      continue;
    }

    const dueDate = parseDueDate_(dueRaw);
    if (!dueDate) {
      skipped++;
      continue;
    }

    const allDay = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
    const eventTitle = `Scholarship Due: ${name}`;

    const notes = idxNotes !== null ? String(row[idxNotes] || "") : "";
    const eventDesc =
      `Scholarship: ${name}\n` +
      `Portal: ${portal}\n` +
      `Sheet: ${MAIN_SHEET}\n` +
      (notes ? `\nNotes:\n${notes}\n` : "");

    if (!eventId) {
      const ev = cal.createAllDayEvent(eventTitle, allDay, {
        description: eventDesc,
        location: portal
      });
      applyScholarshipReminders_(ev);
      row[idxEvent] = ev.getId();
      created++;
      continue;
    }

    // Update existing
    let ev = safeGetEventById_(cal, eventId);
    if (!ev) {
      const newEv = cal.createAllDayEvent(eventTitle, allDay, {
        description: eventDesc,
        location: portal
      });
      applyScholarshipReminders_(newEv);
      row[idxEvent] = newEv.getId();
      created++;
      continue;
    }

    // Check date changes for all-day events
    const evDate = ev.getAllDayStartDate ? ev.getAllDayStartDate() : ev.getStartTime();
    const evDay = new Date(evDate.getFullYear(), evDate.getMonth(), evDate.getDate());

    const needsDate = evDay.getTime() !== allDay.getTime();
    const needsTitle = ev.getTitle() !== eventTitle;
    const needsLoc = (ev.getLocation() || "") !== portal;
    const needsDesc = (ev.getDescription() || "") !== eventDesc;

    if (needsDate) {
      // safest for all-day: recreate
      ev.deleteEvent();
      const repl = cal.createAllDayEvent(eventTitle, allDay, {
        description: eventDesc,
        location: portal
      });
      applyScholarshipReminders_(repl);
      row[idxEvent] = repl.getId();
      updated++;
      continue;
    }

    if (needsTitle) ev.setTitle(eventTitle);
    if (needsLoc) ev.setLocation(portal);
    if (needsDesc) ev.setDescription(eventDesc);

    // Ensure reminders are present (idempotent-ish)
    ensureScholarshipReminders_(ev);

    if (needsTitle || needsLoc || needsDesc) updated++;
    else skipped++;
  }

  dataRange.setValues(data);
  ss.toast(`Calendar: created ${created}, updated ${updated}, removed ${removed}, skipped ${skipped}`, "Autoship", 7);
}

/***********************
 * Due date parser
 * Accepts: Date objects or strings like "Mar 1, 2026", "3/1/26", "2026-03-01"
 ***********************/
function parseDueDate_(dueRaw) {
  if (Object.prototype.toString.call(dueRaw) === "[object Date]" && !isNaN(dueRaw.getTime())) {
    return dueRaw;
  }

  const s = String(dueRaw || "").trim();
  if (!s) return null;

  // ISO-ish
  const iso = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (iso) {
    const y = +iso[1], m = +iso[2] - 1, d = +iso[3];
    const dt = new Date(y, m, d);
    return isNaN(dt.getTime()) ? null : dt;
  }

  // MM/DD/YY or MM/DD/YYYY
  const md = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (md) {
    const m = +md[1] - 1, d = +md[2];
    let y = +md[3];
    if (y < 100) y += 2000;
    const dt = new Date(y, m, d);
    return isNaN(dt.getTime()) ? null : dt;
  }

  // Let Date.parse try (handles "March 1, 2026" etc.)
  const p = Date.parse(s);
  if (!isNaN(p)) return new Date(p);

  // If text contains a date-like substring, try to extract
  const sub = s.match(/(\w+\s+\d{1,2},\s+\d{4})/);
  if (sub) {
    const p2 = Date.parse(sub[1]);
    if (!isNaN(p2)) return new Date(p2);
  }

  return null;
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

function canvasDeepFetchTextForModuleItem_(courseId, mi) {
  // Returns: { title, body, url }  (body is plain text)
  const type = (mi.type || "").toLowerCase();
  const cid = mi.content_id;

  // default fallback: what we already have
  const fallback = {
    title: mi.title || "",
    body: "",
    url: mi.html_url || mi.url || ""
  };

  try {
    // Page
    if (type === "page") {
      // Canvas sometimes provides page_url; if not, we can’t resolve reliably
      const pageUrl = mi.page_url;
      if (!pageUrl) return fallback;

      const page = canvasFetchJson_(`/api/v1/courses/${courseId}/pages/${encodeURIComponent(pageUrl)}`, {});
      return {
        title: page.title || fallback.title,
        body: stripHtml_(page.body || ""),
        url: page.html_url || fallback.url
      };
    }

    // Assignment
    if (type === "assignment" && cid) {
      const a = canvasFetchJson_(`/api/v1/courses/${courseId}/assignments/${cid}`, {});
      return {
        title: a.name || fallback.title,
        body: stripHtml_(a.description || ""),
        url: a.html_url || fallback.url
      };
    }

    // Discussion topic
    if ((type === "discussion" || type === "discussion_topic") && cid) {
      const d = canvasFetchJson_(`/api/v1/courses/${courseId}/discussion_topics/${cid}`, {});
      return {
        title: d.title || fallback.title,
        body: stripHtml_(d.message || ""),
        url: d.html_url || fallback.url
      };
    }

    // File (metadata only; you can store download link)
    if (type === "file" && cid) {
      const f = canvasFetchJson_(`/api/v1/files/${cid}`, {});
      const bestUrl = f.url || f.html_url || fallback.url; // f.url is often a direct download URL
      return {
        title: f.display_name || fallback.title,
        body: `File posted: ${f.display_name || ""}`.trim(),
        url: bestUrl
      };
    }

    // External URL
    if (type === "external_url") {
      return fallback;
    }

    return fallback;
  } catch (e) {
    // Don’t kill the whole run if one item errors
    Logger.log(`deepFetch fail course=${courseId} item=${mi.id} type=${mi.type}: ${e && e.message ? e.message : e}`);
    return fallback;
  }
}

function autoCreateEssayDocForRow_(sheet, absoluteRow, extracted, ctx) {
  ctx = ctx || {};

  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxReq     = optionalIndex_(h, COL_REQUIREMENTS);
  const idxPortal  = optionalIndex_(h, COL_APPLICATION_PORTAL);
  const idxAppFile = optionalIndex_(h, COL_APPLICATION_FILE);
  const idxName    = optionalIndex_(h, "Scholarship Name");

  const idxDocId   = optionalIndex_(h, "Essay Doc ID");
  const idxStatus  = optionalIndex_(h, "Status");
  const idxDue     = optionalIndex_(h, "Due Date");
  const idxTriage  = optionalIndex_(h, "Triage");

  if (idxReq === null || idxAppFile === null || idxName === null) {
    throw new Error(
      `Missing required columns on "${MAIN_SHEET}": "Scholarship Name", "${COL_REQUIREMENTS}", "${COL_APPLICATION_FILE}"`
    );
  }

  const rowValues = sheet.getRange(absoluteRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const reqText = String(rowValues[idxReq] ?? "");
  const name = String(rowValues[idxName] ?? "").trim() || (extracted?.name || "Scholarship");

  // Portal URL (prefer extracted value from intake; fallback to sheet cell text)
  let portalUrl = String(extracted?.portal_url || ctx.url || "").trim();
  if (!portalUrl && idxPortal !== null) {
    portalUrl = String(rowValues[idxPortal] ?? "").trim();
  }

  // If already has any doc link in Application File, skip
  const appFileCell = sheet.getRange(absoluteRow, idxAppFile + 1);
  const existingUrl = extractDocUrlFromCell_(appFileCell);
  if (existingUrl) return { action: "SKIP_ALREADY_HAS_DOC", url: existingUrl };

  // Decide whether to create doc:
  // 1) If Requirements includes Essay(s) -> create.
  // 2) Otherwise, try to detect an essay prompt; if found -> create.
  let shouldCreate = requirementsIncludes_(reqText, ESSAY_REQUIREMENT_TOKEN);

  // Build best available text for prompt extraction
  let sourceText = String(ctx.sourceText || "").trim();
  if (!sourceText || sourceText.length < 40) {
    // last-resort: fetch from portal (safe: only done when needed)
    if (portalUrl) {
      try { sourceText = fetchUrlAsTextSmart_(portalUrl) || ""; } catch (e) {}
    }
  }

  // Extract prompt (only if we have text)
  let promptText = "";
  let wordLimit = "";
  let pdfRequired = false;

  if (sourceText && sourceText.length >= 40) {
    const ep = geminiExtractEssayPrompt_(sourceText.slice(0, 12000), portalUrl);
    promptText = String(ep?.prompt_text || "").trim();
    wordLimit = String(ep?.word_limit || "").trim();
    pdfRequired = !!ep?.pdf_required;

    // If we found a real prompt, create even if Requirements didn't include Essay(s)
    if (promptText.length >= 20) shouldCreate = true;
  }

  if (!shouldCreate) return { action: "SKIP_NO_ESSAY_SIGNAL" };

  // Create the doc using your existing creator (now with prompt + limit)
  const docInfo = createEssayPrepDocOneDraft_({
    scholarshipName: name,
    portalUrl,
    promptText,
    wordLimit
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
    const due = parseSheetDate_(rowValues[idxDue]);
    sheet.getRange(absoluteRow, idxTriage + 1).setValue(computeTriage_(due, rowValues[idxTriage]));
  }

  // Optional: note PDF-required somewhere if you have a column (not shown)

  return { action: "CREATED_DOC", url: docInfo.url, prompt_len: promptText.length, pdf_required: pdfRequired };
}

function deleteEssayDocForRow_(sheet, row) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const h = buildHeaderIndex(headers);

  const idxAppFile = optionalIndex_(h, COL_APPLICATION_FILE);
  const idxDocId = optionalIndex_(h, "Essay Doc ID");

  if (idxAppFile === null && idxDocId === null) {
    throw new Error(`Need "${COL_APPLICATION_FILE}" or "Essay Doc ID" column to delete.`);
  }

  // Prefer stored ID
  let docId = "";
  if (idxDocId !== null) {
    docId = String(sheet.getRange(row, idxDocId + 1).getValue() || "").trim();
  }

  // Fallback: extract from Application File cell
  if (!docId && idxAppFile !== null) {
    const cell = sheet.getRange(row, idxAppFile + 1);
    const docUrl = extractDocUrlFromCell_(cell);
    docId = extractDriveIdFromUrl_(docUrl);
  }

  if (!docId) return { action: "SKIP_NO_DOC_FOUND", row };

  const ok = trashDriveFileBestEffort_(docId);

  // Clean up sheet cells
  if (idxDocId !== null) sheet.getRange(row, idxDocId + 1).clearContent();

  if (idxAppFile !== null) {
    const cell = sheet.getRange(row, idxAppFile + 1);
    const txt = String(cell.getDisplayValue() || "");
    const cleaned = txt
      .split(/\n+/).map(s => s.trim()).filter(Boolean)
      .filter(line => !/docs\.google\.com\/document\/d\//i.test(line))
      .filter(line => normalize_(line) !== normalize_("Essay Draft (Google Doc)"))
      .join("\n");
    if (!cleaned) cell.clearContent();
    else cell.setValue(cleaned);
  }

  return { action: ok ? "TRASHED_DOC" : "TRASH_FAILED", row, docId };
}

function draftRecEmailsForRowTool(args) {
  const row = Number(args.row);
  if (!row || row <= HEADER_ROW) throw new Error("Provide a valid row > header.");
  const sh = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET);
  return draftRecEmailsForRow_(sh, row);
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

function safeGetEventById_(cal, eventId) {
  try {
    return cal.getEventById(eventId);
  } catch (e) {
    return null;
  }
}

function isPastDue_(dueDate) {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const d = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
  return d.getTime() < today.getTime();
}

function applyScholarshipReminders_(ev) {
  // Clear existing reminders then add ours
  try {
    ev.removeAllReminders();
  } catch (e) {
    // Some events may not support removal in edge cases; ignore
  }

  // Add popup reminders at midnight-local offsets
  // (Days -> minutes)
  SCHOLARSHIP_REMINDER_DAYS.forEach(d => {
    const minutes = d * 24 * 60;
    try {
      ev.addPopupReminder(minutes);
    } catch (e) {
      // If popup reminder fails, you can switch to email reminders:
      // ev.addEmailReminder(minutes);
    }
  });
}

function ensureScholarshipReminders_(ev) {
  // Simple strategy: always reset to the correct schedule.
  // Keeps it consistent even if user edits reminders manually.
  applyScholarshipReminders_(ev);
}
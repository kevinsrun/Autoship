const MAIN_SHEET = "Main Menu";
const RESULTS_SHEET = "Results";
const SURRENDERED_SHEET = "Surrendered";
const HEADER_ROW = 1;

// Canvas config
const CANVAS_COURSES_CACHE_KEY = "CANVAS_ACTIVE_COURSES_v1";
const CANVAS_COURSES_CACHE_SECONDS = 6 * 60 * 60; // 6 hours
const CANVAS_MAX_COURSES_PER_RUN = 40;            // hard cap
const CANVAS_SLEEP_MS = 350;

// Force/target courses
const FORCE_CANVAS_COURSE_IDS = [302960];

const TARGET_CANVAS_COURSE_IDS = [
  138471, // MHS Counselors
  239025  // Class of 2026
];

const TARGET_COURSE_IDS = {
  COUNSELOR: [138471],
  CLASS_2026: [239025]
};

// ===== ELIGIBILITY PROFILE (Kevin) =====
const USER_PROFILE = {
  grade_level: "Senior",
  hs_class_year: 2026,
  age: 18,
  state: "TX",
  city: "Arlington",
  gender: "Male",
  ethnicity: "asian",
  intended_majors: ["biochem", "biochemistry", "biomedical engineering", "bme", "bioengineering"],
  gpa_unweighted: 3.98,
  gpa_weighted: 4.78,
  family_income: 133000,
  sai: 20000,

  // "within 10000 is fine" rule for income + SAI cutoffs
  tolerance_income: 10000,
  tolerance_sai: 10000
};

// ✅ IMPORTANT: your code references APPLICANT_PROFILE in multiple places
const APPLICANT_PROFILE = {
  gradeLevel: USER_PROFILE.grade_level,
  hsClassYear: USER_PROFILE.hs_class_year,
  age: USER_PROFILE.age,
  state: USER_PROFILE.state,
  city: USER_PROFILE.city,
  gender: USER_PROFILE.gender,
  ethnicity: USER_PROFILE.ethnicity,
  intendedMajors: USER_PROFILE.intended_majors,
  gpaUnweighted: USER_PROFILE.gpa_unweighted,
  gpaWeighted: USER_PROFILE.gpa_weighted,
  familyIncome: USER_PROFILE.family_income,
  sai: USER_PROFILE.sai,
  incomeTolerance: USER_PROFILE.tolerance_income,
  saiTolerance: USER_PROFILE.tolerance_sai
};

// Gemini
const GEMINI_MODEL = "gemini-2.5-flash";

// Your Requirements multi-select token that indicates writing
const ESSAY_REQUIREMENT_TOKEN = "Essay(s)";

// Headers (must match your sheet)
const COL_APPLICATION_PORTAL = "Application Portal";
const COL_REQUIREMENTS = "Requirements";
const COL_APPLICATION_FILE = "Application File";
const COL_ADDITIONAL_APPLICATION_FILE = "Additional Application Essay File";
const COL_THEME = "Theme"; // ✅ must be defined BEFORE any usage

// If Gemini detects PDF upload requirement
const PDF_REQUIRED_KEYWORDS = ["pdf", "upload a pdf", "submit a pdf", "pdf format"];

// Performance tuning
const MAX_HTML_CHARS = 12000;

// Reminder schedule (days before due date)
const SCHOLARSHIP_REMINDER_DAYS = [7, 3, 1];

// Status words that should REMOVE the calendar event
const REMOVE_EVENT_TRIAGE_VALUES = ["surrendered"];
const REMOVE_EVENT_STATUS_VALUES = ["submitted", "complete", "completed", "won", "not applying"];

// Calendar columns / properties
const CAL_EVENT_COL_NAME = "Calendar Event Id";
const DEFAULT_CAL_ID_PROP = "SCHOLARSHIP_CALENDAR_ID";
const CAL_WATCH_HEADERS = ["Due Date", "Triage", "Status", "Scholarship Name", "Application Portal", "Notes"];

// Dropdown options (keep EXACT spelling to match your sheet)
const DIFFICULTY_OPTIONS = ["Easy", "Medium", "Hard", "Fuck it", "No Idea..."];
const TRIAGE_OPTIONS = ["Immediate", "Urgent", "Non-Urgent", "Satisfied", "Cooked"];
const STATUS_OPTIONS = ["Not started", "Scanned", "Prepped", "In Progress", "Completed", "Surrendered"];

const THEME_OPTIONS = [
  "Biomed.", "Comp Sci.", "Business/Entrepreneurship", "Pol. Sci.", "Herritage",
  "Engineering", "Math/Sci", "First Gen", "Broke", "English/Humanities",
  "Leadership/Service", "None", "Other"
];

const REQUIREMENT_OPTIONS = [
  "Essay(s)", "Rec. Letter(s)", "App./ECs", "Academics/Transcript",
  "Field of Study", "Exam(s)", "Interview(s)", "Broke/Finances",
  "Video/Recording", "Other"
];

/*************************************************
 * QUICK TOAST / DEBUG
 *************************************************/
function testToast_() {
  SpreadsheetApp.getActive().toast("Autoship is LIVE", "OK", 3);
}

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

/*************************************************
 * CORE HELPERS (USED ABOVE/BELOW)
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
  return /^https?:\/\/\S+$/i.test(String(s || ""));
}

function normalize_(s) {
  return (s ?? "")
    .toString()
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function norm_(s) {
  return (s ?? "").toString().trim().toLowerCase().replace(/\u00A0/g, " ").replace(/\s+/g, " ");
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

function clampToOptions_(raw, options, fallback) {
  const r = norm_(raw);
  if (!r) return fallback;

  for (const o of options) if (norm_(o) === r) return o;
  for (const o of options) {
    const no = norm_(o);
    if (no.includes(r) || r.includes(no)) return o;
  }
  return fallback;
}

function parseDate_(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function stripHtml_(html) {
  if (!html) return "";
  return String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractDriveIdFromUrl_(url) {
  if (!url) return "";
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : "";
}

/*************************************************
 * TRIAGE HELPERS
 *************************************************/
function parseSheetDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  const s = String(v).trim();
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;

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

function computeTriage_(dueDate, currentTriage) {
  const ct = clampToOptions_(currentTriage, TRIAGE_OPTIONS, "");
  if (ct === "Satisfied" || ct === "Cooked") return ct;

  const d = daysUntil_(dueDate);
  if (d == null) return "Non-Urgent";

  if (d < 0) return "Cooked";
  if (d <= 7) return "Immediate";
  if (d <= 28) return "Urgent";
  return "Non-Urgent";
}

function computeStatus_(opts) {
  const cs = clampToOptions_(opts.currentStatus, STATUS_OPTIONS, "Not started");
  if (cs === "Completed" || opts.completed) return "Completed";
  if (cs === "Surrendered" || opts.surrendered) return "Surrendered";
  if (opts.prepped) return "Prepped";
  if (opts.scanned) return "Scanned";
  return "Not started";
}

function computeDifficulty_(rawDifficulty) {
  return clampToOptions_(rawDifficulty, DIFFICULTY_OPTIONS, "No Idea...");
}

/*************************************************
 * ELIGIBILITY (USED BY CANVAS SCAN)
 *************************************************/
function moneyToNumber_(s) {
  const raw = String(s || "").toLowerCase().replace(/\$/g, "").replace(/,/g, "").trim();
  if (!raw) return null;
  const k = raw.match(/^(\d+(?:\.\d+)?)\s*k$/);
  if (k) return Math.round(parseFloat(k[1]) * 1000);
  const n = parseFloat(raw);
  return isNaN(n) ? null : Math.round(n);
}

function extractIncomeCap_(t) {
  const m = String(t || "").match(/\b(household|family)?\s*income\s*(under|less than|below|<=?)\s*\$?\s*([\d,]{4,})\b/i);
  if (!m) return null;
  const n = Number(String(m[3]).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}

function extractSaiCap_(t) {
  const m = String(t || "").match(/\b(sai|student aid index)\s*(under|less than|below|<=?)\s*([\d,]{3,})\b/i);
  if (!m) return null;
  const n = Number(String(m[3]).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}

function extractAnyDate_(t) {
  const s = String(t || "");

  let m = s.match(/\b(20\d{2})[-\/](\d{1,2})[-\/](\d{1,2})\b/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(d.getTime()) ? null : d;
  }

  m = s.match(/\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/);
  if (m) {
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    const d = new Date(y, Number(m[1]) - 1, Number(m[2]));
    return isNaN(d.getTime()) ? null : d;
  }

  m = s.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+(\d{1,2})(,)?\s+(20\d{2})\b/i);
  if (m) {
    const monthMap = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,sept:8,oct:9,nov:10,dec:11 };
    const mm = monthMap[m[1].toLowerCase()];
    const d = new Date(Number(m[4]), mm, Number(m[2]));
    return isNaN(d.getTime()) ? null : d;
  }

  return null;
}

function eligibilityGate_(text, profile) {
  const t = String(text || "").toLowerCase();

  // deadline passed (very conservative)
  const dt = extractAnyDate_(t);
  if (dt && !isNaN(dt.getTime())) {
    const today = new Date();
    const startToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const startDt = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    if (startDt.getTime() < startToday.getTime()) {
      return { eligible: false, reason: "deadline passed" };
    }
  }

  // Texas resident requirement
  if (/\b(texas residents?|resident of texas|must be (a )?texas resident)\b/.test(t)) {
    if (String(profile.state || "").toUpperCase() !== "TX") return { eligible: false, reason: "not Texas resident" };
  }

  // Arlington requirement
  if (/\b(arlington(,)? texas|city of arlington|arlington isd|arlington resident)\b/.test(t)) {
    if (String(profile.city || "").toLowerCase() !== "arlington") return { eligible: false, reason: "not Arlington" };
  }

  // gender requirements
  if (/\b(male applicants only|men only|for men|must be male)\b/.test(t)) {
    if (String(profile.gender || "").toLowerCase() !== "male") return { eligible: false, reason: "requires male" };
  }
  if (/\b(female applicants only|women only|for women|must be female)\b/.test(t)) {
    if (String(profile.gender || "").toLowerCase() === "male") return { eligible: false, reason: "requires female" };
  }

  // senior requirement
  if (/\b(high school senior|hs senior|12th grade|grade 12)\b/.test(t)) {
    if (String(profile.gradeLevel || "").toLowerCase() !== "senior") return { eligible: false, reason: "requires senior" };
  }

  // age strict
  if (/\b(must be 18|age 18 required)\b/.test(t)) {
    if (Number(profile.age) !== 18) return { eligible: false, reason: "requires age 18" };
  }

  // ethnicity strict
  if (/\b(asian (students|applicants)|aapi (students|applicants))\b/.test(t)) {
    if (String(profile.ethnicity || "").toLowerCase() !== "asian") return { eligible: false, reason: "requires Asian/AAPI" };
  }

  // major strict
  const wantsBiochem = /\b(biochem|biochemistry)\b/.test(t);
  const wantsBme = /\b(biomedical engineering|biomed eng|bme)\b/.test(t);
  if (wantsBiochem || wantsBme) {
    const majors = (profile.intendedMajors || []).map(x => String(x).toLowerCase());
    if (wantsBiochem && !majors.some(m => m.includes("biochem"))) return { eligible: false, reason: "requires biochem" };
    if (wantsBme && !majors.some(m => m.includes("biomedical"))) return { eligible: false, reason: "requires BME" };
  }

  // income cap
  const incomeCap = extractIncomeCap_(t);
  if (incomeCap != null) {
    if (Number(profile.familyIncome) > incomeCap) return { eligible: false, reason: `income>${incomeCap}` };
  }

  // SAI cap with tolerance
  const saiCap = extractSaiCap_(t);
  if (saiCap != null) {
    const sai = Number(profile.sai);
    const tol = Number(profile.saiTolerance || 0);
    if (sai > (saiCap + tol)) return { eligible: false, reason: `sai>${saiCap}+tol` };
  }

  return { eligible: true, reason: "" };
}

/*************************************************
 * FINGERPRINT (USED BY CANVAS DEDUPE)
 *************************************************/
function canvasFingerprint_(title, body, url) {
  const s = normalize_([title, body, url].join(" | ")).slice(0, 2000);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

/*******************************
 * CANVAS → AUTOSHIP INTAKE (FIXED)
 *******************************/

const CANVAS_INTAKE_SHEET = "Canvas Intake";

// Course matching (optional — not required if using fixed IDs)
const COURSE_MATCH = {
  COUNSELOR: [/counselor/i, /mhs/i],
  CLASS_2026: [/(class|senior).*(2026|26)/i]
};

const SCHOLAR_KEYWORDS = [
  "scholarship", "award", "grant", "fellowship",
  "financial", "money", "tuition", "apply", "application",
  "deadline", "due", "senior", "class of", "$", "dollars",
  "local scholarship", "foundation", "community scholarship",
  "fafsa", "tasfa"
];

const SCAN_PROFILE = {
  COUNSELOR: {
    announcementsLookbackDays: 120,
    modulesLookbackDays: 365,
    maxModuleItems: 250
  },
  CLASS_2026: {
    announcementsLookbackDays: 45,
    modulesLookbackDays: 120,
    maxModuleItems: 60
  }
};

/**
 * Entry point (menu item calls this)
 */
function scanCanvasForScholarships() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  // 5-minute cooldown
  const last = Number(props.getProperty("canvas:lastScanMs") || "0");
  const now = Date.now();
  if (now - last < 5 * 60 * 1000) {
    ss.toast("Canvas scan blocked (cooldown 5 min).", "Autoship", 5);
    return;
  }
  props.setProperty("canvas:lastScanMs", String(now));

  ensureCanvasIntakeSheet_();
  const intake = ss.getSheetByName(CANVAS_INTAKE_SHEET);
  if (!intake) throw new Error(`Missing sheet: "${CANVAS_INTAKE_SHEET}"`);

  const seen = loadSeenKeys_(); // contains "key:<...>" and "fp:<...>"

  // Only scan the IDs you explicitly set
  const courses = canvasListActiveCourses_();

  const targets = [];
  courses.forEach(c => {
    if (TARGET_COURSE_IDS.COUNSELOR.includes(c.id)) targets.push({ course: c, profile: "COUNSELOR" });
    if (TARGET_COURSE_IDS.CLASS_2026.includes(c.id)) targets.push({ course: c, profile: "CLASS_2026" });
  });

  // prioritize counselor scan
  targets.sort((a, b) => (a.profile === "COUNSELOR" ? -1 : 1));

  const rowsToAppend = [];

  targets.forEach(t => {
    const { course, profile } = t;
    const cfg = SCAN_PROFILE[profile];
    if (!cfg) return;

    // --- Announcements ---
    const ann = canvasFetchAnnouncements_(course.id, cfg.announcementsLookbackDays);

    ann.forEach(a => {
      const plain = stripHtml_(a.message || "");
      const hit = scholarshipHit_(a.title, a.message);
      if (!hit) return;

      const key = `announcement:${a.id}`;
      const fp = canvasFingerprint_(a.title || "", plain, a.html_url || "");

      if (seen.has("key:" + key) || seen.has("fp:" + fp)) return;

      // Eligibility gate
      const gate = eligibilityGate_((a.title || "") + "\n" + plain, APPLICANT_PROFILE);
      if (!gate.eligible) return;

      seen.add("key:" + key);
      seen.add("fp:" + fp);

      rowsToAppend.push(canvasRow_({
        key,
        fingerprint: fp,
        profile,
        course,
        type: "Announcement",
        title: a.title || "",
        body: plain,
        postedAt: a.posted_at || "",
        url: a.html_url || "",
        hit
      }));
    });

    // --- Modules (Sunday only) ---
    const doModules = (new Date().getDay() === 0);
    if (!doModules) return;

    const modItems = canvasFetchModuleItems_(course.id, cfg.modulesLookbackDays, cfg.maxModuleItems);

    modItems.forEach(mi => {
      // Cheap filter
      const hitTitleOnly = scholarshipHit_(mi.title || "", "");
      if (!hitTitleOnly) return;

      const deep = canvasDeepFetchTextForModuleItem_(course.id, mi);
      const hit = scholarshipHit_(deep.title, deep.body);
      if (!hit) return;

      const key = `module_item:${mi.id}`;
      const fp = canvasFingerprint_(deep.title || "", deep.body || "", deep.url || "");

      if (seen.has("key:" + key) || seen.has("fp:" + fp)) return;

      const gate = eligibilityGate_((deep.title || "") + "\n" + (deep.body || ""), APPLICANT_PROFILE);
      if (!gate.eligible) return;

      seen.add("key:" + key);
      seen.add("fp:" + fp);

      rowsToAppend.push(canvasRow_({
        key,
        fingerprint: fp,
        profile,
        course,
        type: `Module Item (${mi.type || "unknown"})`,
        title: deep.title || "",
        body: deep.body || "",
        postedAt: mi.updated_at || mi.published_at || "",
        url: deep.url || "",
        hit
      }));
    });
  });

  if (rowsToAppend.length) {
    appendCanvasRows_(rowsToAppend);
    ss.toast(`Canvas scan: added ${rowsToAppend.length} hit(s).`, "Autoship", 6);
  } else {
    ss.toast("Canvas scan: no new hits.", "Autoship", 5);
  }
}

/***********************
 * CANVAS API HELPERS
 ***********************/
function canvasBaseUrl_() {
  const v = PropertiesService.getScriptProperties().getProperty("CANVAS_BASE_URL");
  if (!v) throw new Error("Missing Script Property: CANVAS_BASE_URL");

  const base = v.replace(/\/+$/, "");
  if (/canvas\.arlingtonisd\.org$/i.test(base)) {
    throw new Error("Wrong CANVAS_BASE_URL. Use https://arlington.instructure.com");
  }
  return base;
}

function canvasToken_() {
  const v = PropertiesService.getScriptProperties().getProperty("CANVAS_TOKEN");
  if (!v) throw new Error("Missing Script Property: CANVAS_TOKEN");
  return v.trim();
}

// Handles arrays like {"fields[]": ["id","name"]}
function buildUrl_(base, params) {
  const parts = [];
  Object.keys(params || {}).forEach(k => {
    const v = params[k];
    if (v === null || v === undefined || v === "") return;

    if (Array.isArray(v)) {
      v.forEach(item => parts.push(encodeURIComponent(k) + "=" + encodeURIComponent(String(item))));
    } else {
      parts.push(encodeURIComponent(k) + "=" + encodeURIComponent(String(v)));
    }
  });
  return parts.length ? `${base}?${parts.join("&")}` : base;
}

function canvasFetchJson_(path, params) {
  const base = canvasBaseUrl_();
  const url = buildUrl_(base + path, params || {});
  Logger.log("Canvas URL: " + url);

  const res = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + canvasToken_() }
  });

  const code = res.getResponseCode();
  const text = res.getContentText() || "";

  if (code < 200 || code >= 300) {
    throw new Error(`Canvas HTTP ${code}: ${text.slice(0, 300)}`);
  }
  return JSON.parse(text || "null");
}

function canvasListActiveCourses_() {
  const out = [];

  TARGET_CANVAS_COURSE_IDS.forEach(id => {
    try {
      const c = canvasFetchJson_(`/api/v1/courses/${id}`, {
        "fields[]": ["id", "name", "course_code"]
      });
      if (c && c.id) out.push(c);
    } catch (e) {
      Logger.log(`Could not fetch course ${id}: ${e && e.message ? e.message : e}`);
    }
  });

  return out;
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

  const ITEMS_PER_PAGE = 30;
  const MAX_MODULES_PER_RUN = 4;

  const modules = canvasFetchJson_(`/api/v1/courses/${courseId}/modules`, {
    per_page: ITEMS_PER_PAGE
  }) || [];

  const out = [];

  for (let i = 0; i < modules.length && i < MAX_MODULES_PER_RUN; i++) {
    if (out.length >= maxItems) break;

    const m = modules[i];

    const items = canvasFetchJson_(`/api/v1/courses/${courseId}/modules/${m.id}/items`, {
      per_page: ITEMS_PER_PAGE
    }) || [];

    Utilities.sleep(250);

    for (let j = 0; j < items.length; j++) {
      if (out.length >= maxItems) break;

      const it = items[j];
      const updated = parseDate_(it.updated_at) || parseDate_(it.published_at);
      if (updated && updated.getTime() < cutoff) continue;

      out.push({
        id: it.id,
        title: it.title || "(no title)",
        type: it.type || "",
        content_id: it.content_id || null,
        page_url: it.page_url || "",
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

  for (const k of SCHOLAR_KEYWORDS) {
    if (k === "$") {
      if (t.includes("$") || b.includes("$")) return "contains:$";
      continue;
    }
    if (t.includes(k) || b.includes(k)) return `contains:${k}`;
  }

  if (/\bdue\b/.test(t + " " + b) && /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b/.test(t + " " + b)) {
    return "pattern:due+month";
  }

  return "";
}

function canvasRow_({ key, fingerprint, profile, course, type, title, body, postedAt, url, hit }) {
  return [
    new Date(),
    key,
    fingerprint || "",
    profile || "",
    (course && course.name) ? course.name : "",
    type || "",
    title || "",
    postedAt || "",
    url || "",
    hit || "",
    body || ""
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
      "Fingerprint",
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

  const lastRow = sh.getLastRow();
  const vals = sh.getRange(2, 2, lastRow - 1, 2).getValues(); // Unique Key + Fingerprint

  vals.forEach(r => {
    const k = String(r[0] || "").trim();
    const f = String(r[1] || "").trim();
    if (k) seen.add("key:" + k);
    if (f) seen.add("fp:" + f);
  });

  return seen;
}

/***********************
 * UTIL
 ***********************/
function pickCoursesAll_(courses, regexList) {
  return (courses || []).filter(c => {
    const name = String(c.name || "");
    return regexList.every(rx => rx.test(name));
  });
}

function pickCoursesAny_(courses, regexList) {
  return (courses || []).filter(c => {
    const name = String(c.name || "");
    return regexList.some(rx => rx.test(name));
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
  const existing = buildExistingScholarshipKeySet_();
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
    const startRow = main.getLastRow() + 1;

    // 1) fast append
    main.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);

    // 2) rebuild RichText for link-sensitive columns (Notes + Portal)
    for (let i = 0; i < rowsToAppend.length; i++) {
      const rOut = startRow + i;

      // Portal: make clickable if it’s a URL
      if (mPortal !== null) {
        const portal = String(rowsToAppend[i][mPortal] || "").trim();
        if (portal && /^https?:\/\//i.test(portal)) {
          const rt = SpreadsheetApp.newRichTextValue()
            .setText(portal)
            .setLinkUrl(portal)
            .build();
          main.getRange(rOut, mPortal + 1).setRichTextValue(rt);
        }
      }

      // Notes: auto-link any raw URLs inside the text
      if (mNotes !== null) {
        const notesText = String(rowsToAppend[i][mNotes] || "");
        const rtNotes = buildRichTextPreservingLinks_(notesText, null);
        main.getRange(rOut, mNotes + 1).setRichTextValue(rtNotes);
      }
    }
  }

  // ✅ write back "Piped" status (THIS must be inside the function)
  if (pipedUpdates.length) {
    const pipedCol = idxPiped + 1; // 1-based
    const updatesRange = intake.getRange(2, pipedCol, intakeValues.length, 1);
    const colVals = updatesRange.getValues();

    pipedUpdates.forEach(([i, val]) => {
      colVals[i][0] = val;
    });

    updatesRange.setValues(colVals);
  }

  ss.toast(`Piped ${rowsToAppend.length} Canvas hit(s) into Main Menu.`, "Autoship", 6);
} // ✅ end pipeCanvasHitsToMainMenu


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
  const mainDisplay = mainRange.getDisplayValues(); // ✅ dropdown-safe
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
  const idxEvent = optionalIndex_(h, CAL_EVENT_COL_NAME); // may be null


    // --- Calendar (optional but recommended) ---
  const calId = PropertiesService.getScriptProperties()
    .getProperty(DEFAULT_CAL_ID_PROP) || "primary";
  const cal = CalendarApp.getCalendarById(calId);

  const existing = new Set();
  for (let i = 1; i < resValues.length; i++) {
    const nm = (resValues[i][rName] || "").toString().trim();
    if (nm) existing.add(nm.toLowerCase());
  }

  const linkWrites = [];
  const rowsToRemove = [];
  let moved = 0;

  for (let i = 1; i < mainValues.length; i++) {
    const name = String(mainValues[i][idxName] || "").trim();
    if (!name) continue;

    const triageDisp = String(mainDisplay[i][idxTriage] || "");
    const triageNorm = normalize_(triageDisp);

    // DEBUG: log first 25 rows
    if (i <= 25) Logger.log(`Row ${HEADER_ROW + i}: triageDisp="${triageDisp}" triageNorm="${triageNorm}"`);

    if (triageNorm !== "satisfied") continue;

    const key = name.toLowerCase();
    const row = mainValues[i];

    const out = new Array(resValues[0].length).fill("");

    out[rName] = row[idxName];
    out[rDue] = row[idxDue];
    out[rDifficulty] = row[idxDifficulty];
    out[rNotes] = row[idxNotes];
    out[rPortal] = row[idxPortal];

    if (rReq !== null && idxReq !== null) out[rReq] = row[idxReq];

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

    // ✅ delete calendar event for satisfied rows we’re moving
    if (idxEvent !== null && cal) {
      const eventId = String(mainValues[i][idxEvent] || "").trim();
      if (eventId) {
        const ev = safeGetEventById_(cal, eventId);
        if (ev) ev.deleteEvent();
        // optional: clear it in-memory so when you write main back later it’s blank
        // mainValues[i][idxEvent] = "";
      }
    }

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

  // Apply link writes in Results
  if (linkWrites.length) {
    linkWrites.forEach(w => {
      const cell = results.getRange(w.row, w.col);
      if (w.formula) cell.setFormula(w.formula);
      else if (w.richText) cell.setRichTextValue(w.richText);
    });
  }

  // ✅ Delete from Main bottom-up
  if (rowsToRemove.length) {
    rowsToRemove.sort((a, b) => b - a);
    for (const r of rowsToRemove) main.deleteRow(r);
  }

  ss.toast(
    moved ? `Moved ${moved} satisfied row(s) to Results.` : "No rows with Triage = Satisfied.",
    "Scholarship Tools",
    6
  );
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
  //  Surrendered sync is based on STATUS == Surrendered (use display values)
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

function extractLinkRunsFromRichText_(rt) {
  const out = [];
  if (!rt) return out;

  const runs = rt.getRuns ? rt.getRuns() : [];
  runs.forEach(run => {
    const url = run.getLinkUrl();
    const txt = run.getText();
    if (url && txt) out.push({ text: txt, url: url });
  });

  // whole-cell link fallback
  try {
    const whole = rt.getLinkUrl && rt.getLinkUrl();
    const wholeText = rt.getText && rt.getText();
    if (whole && wholeText) out.push({ text: wholeText, url: whole });
  } catch (_) {}

  // de-dupe
  const seen = new Set();
  return out.filter(x => {
    const k = x.url + "||" + x.text;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

function buildRichTextPreservingLinks_(newText, oldRichText) {
  const text = String(newText ?? "");
  const builder = SpreadsheetApp.newRichTextValue().setText(text);

  const runs = extractLinkRunsFromRichText_(oldRichText);

  // Re-apply old links where the same linked text appears in new text
  runs.forEach(r => {
    const needle = String(r.text || "");
    if (!needle) return;

    let idx = 0;
    while (idx < text.length) {
      const at = text.indexOf(needle, idx);
      if (at === -1) break;
      builder.setLinkUrl(at, at + needle.length, r.url);
      idx = at + needle.length;
    }
  });

  // Also auto-link any raw URLs that appear in new text
  const urlRegex = /https?:\/\/[^\s)"'>]+/g;
  let m;
  while ((m = urlRegex.exec(text)) !== null) {
    const u = m[0];
    builder.setLinkUrl(m.index, m.index + u.length, u);
  }

  return builder.build();
}

function setCellTextPreserveLinks_(sheet, row, col, newText) {
  const cell = sheet.getRange(row, col);
  const oldRT = cell.getRichTextValue();
  const rt = buildRichTextPreservingLinks_(newText, oldRT);
  cell.setRichTextValue(rt);
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

/***********************
 * MAIN MENU → GOOGLE CALENDAR
 ***********************/

/**
 * Installable trigger (required to call CalendarApp reliably)
 * Run ONCE manually from Apps Script.
 */
function installCalendarAutoSyncTrigger() {
  const ss = SpreadsheetApp.getActive();

  // remove old triggers pointing to this handler to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === "calendarAutoSyncOnEdit") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("calendarAutoSyncOnEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast("Installed: Calendar auto-sync on edit ✅", "Autoship", 5);
}

/**
 * Trigger handler: when you edit Main Menu, auto-sync only that row.
 */
function calendarAutoSyncOnEdit(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (!sh) return;

    // Only react on Main Menu
    if (sh.getName() !== MAIN_SHEET) return;

    // Ignore header edits
    const row = e.range.getRow();
    if (row <= HEADER_ROW) return;

    // Only react if the edited column is one of our watch columns
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const h = buildHeaderIndex(headers);

    const editedCol = e.range.getColumn(); // 1-based
    const editedHeader = headers[editedCol - 1] ? String(headers[editedCol - 1]).trim() : "";

    if (editedHeader && !CAL_WATCH_HEADERS.includes(editedHeader)) return;

    // Sync just this one row
    syncScholarshipCalendarForRow_(sh, row, h);
  } catch (err) {
    Logger.log("calendarAutoSyncOnEdit error: " + (err && err.message ? err.message : err));
  }
}

/**
 * Sync ONE row of Main Menu to Calendar (create/update/delete).
 * Deletes event if:
 *  - Triage == Satisfied
 *  - Due Date blank
 *  - Status in REMOVE_EVENT_STATUS_VALUES
 *  - Triage in REMOVE_EVENT_TRIAGE_VALUES
 */
function syncScholarshipCalendarForRow_(sh, rowNumber, headerIndexMap) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const h = headerIndexMap || buildHeaderIndex(
      sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0]
    );

    const idxName   = optionalIndex_(h, "Scholarship Name");
    const idxPortal = optionalIndex_(h, COL_APPLICATION_PORTAL);
    const idxDue    = optionalIndex_(h, "Due Date");
    const idxTriage = optionalIndex_(h, "Triage");
    const idxStatus = optionalIndex_(h, "Status");
    const idxNotes  = optionalIndex_(h, "Notes");
    const idxEvent  = optionalIndex_(h, CAL_EVENT_COL_NAME);

    if (idxName === null || idxPortal === null || idxDue === null) return;
    if (idxEvent === null) return; // if you don't have the column, nothing to store

    const lastCol = sh.getLastColumn();
    const rng = sh.getRange(rowNumber, 1, 1, lastCol);

    const values  = rng.getValues()[0];
    const display = rng.getDisplayValues()[0];

    const name   = String(values[idxName] || "").trim();
    const portal = String(values[idxPortal] || "").trim();
    const dueRaw = values[idxDue];
    const triage = idxTriage !== null ? String(display[idxTriage] || "").trim().toLowerCase() : "";
    const status = idxStatus !== null ? String(display[idxStatus] || "").trim().toLowerCase() : "";
    const notes  = idxNotes !== null ? String(values[idxNotes] || "") : "";
    let eventId  = String(values[idxEvent] || "").trim();

    // Calendar
    const calId = PropertiesService.getScriptProperties().getProperty(DEFAULT_CAL_ID_PROP) || "primary";
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) throw new Error(`Could not open calendar: ${calId}`);

    // Removal conditions
    const dueDate = parseDueDate_(dueRaw); // you already have this helper
    const shouldRemove =
      !name ||
      !portal ||
      !dueDate ||
      (triage && REMOVE_EVENT_TRIAGE_VALUES.includes(triage)) ||
      (status && REMOVE_EVENT_STATUS_VALUES.includes(status)) ||
      (triage === "satisfied"); // <- your new rule

    if (shouldRemove) {
      if (eventId) {
        const ev = safeGetEventById_(cal, eventId);
        if (ev) ev.deleteEvent();
        sh.getRange(rowNumber, idxEvent + 1).setValue(""); // clear stored id
      }
      return;
    }

    // Build event
    const allDay = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
    const eventTitle = `Scholarship Due: ${name}`;
    const eventDesc =
      `Scholarship: ${name}\n` +
      `Portal: ${portal}\n` +
      `Sheet: ${MAIN_SHEET}\n` +
      (notes ? `\nNotes:\n${notes}\n` : "");

    // Create if missing
    if (!eventId) {
      const ev = cal.createAllDayEvent(eventTitle, allDay, {
        description: eventDesc,
        location: portal
      });
      applyScholarshipReminders_(ev);
      sh.getRange(rowNumber, idxEvent + 1).setValue(ev.getId());
      return;
    }

    // Update existing (or recreate if date changed)
    let ev = safeGetEventById_(cal, eventId);
    if (!ev) {
      const newEv = cal.createAllDayEvent(eventTitle, allDay, {
        description: eventDesc,
        location: portal
      });
      applyScholarshipReminders_(newEv);
      sh.getRange(rowNumber, idxEvent + 1).setValue(newEv.getId());
      return;
    }

    const evDate = ev.getAllDayStartDate ? ev.getAllDayStartDate() : ev.getStartTime();
    const evDay = new Date(evDate.getFullYear(), evDate.getMonth(), evDate.getDate());

    const needsDate = evDay.getTime() !== allDay.getTime();
    const needsTitle = ev.getTitle() !== eventTitle;
    const needsLoc = (ev.getLocation() || "") !== portal;
    const needsDesc = (ev.getDescription() || "") !== eventDesc;

    if (needsDate) {
      ev.deleteEvent();
      const repl = cal.createAllDayEvent(eventTitle, allDay, {
        description: eventDesc,
        location: portal
      });
      applyScholarshipReminders_(repl);
      sh.getRange(rowNumber, idxEvent + 1).setValue(repl.getId());
      return;
    }

    if (needsTitle) ev.setTitle(eventTitle);
    if (needsLoc) ev.setLocation(portal);
    if (needsDesc) ev.setDescription(eventDesc);

    ensureScholarshipReminders_(ev);
  } finally {
    lock.releaseLock();
  }
}


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
      (status && REMOVE_EVENT_STATUS_VALUES.includes(status)) ||
      (triage === "satisfied") ||   
      (!dueRaw);                    


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
    .replace(/[\u200B-\u200D\uFEFF]/g, "") // zero-width chars
    .replace(/\s+/g, " ")
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

function buildExistingScholarshipKeySet_() {
  const ss = SpreadsheetApp.getActive();
  const set = new Set();

  const sheets = [MAIN_SHEET, RESULTS_SHEET, "Surrendered"];

  sheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastCol = sh.getLastColumn();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const h = buildHeaderIndex(headers);

    const idxName = h["Scholarship Name"];
    const idxPortal = h["Application Portal"];

    if (idxName != null) {
      const names = sh.getRange(2, idxName + 1, lastRow - 1, 1).getValues();
      names.forEach(r => {
        const v = String(r[0] || "").trim().toLowerCase();
        if (v) set.add("name:" + v);
      });
    }

    if (idxPortal != null) {
      const portals = sh.getRange(2, idxPortal + 1, lastRow - 1, 1).getValues();
      portals.forEach(r => {
        const v = String(r[0] || "").trim();
        if (v) set.add("portal:" + v);
      });
    }
  });

  return set;
}

function canvasFingerprint_(title, body, url) {
  const s = normalize_([title, body, url].join(" | ")).slice(0, 2000);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function extractFirstUrl_(text) {
  const m = String(text || "").match(/https?:\/\/[^\s)"'>]+/i);
  return m ? m[0] : "";
}

function debugWriteCanvasIntake_() {
  const ss = SpreadsheetApp.getActive();
  const intake = ss.getSheetByName("Canvas Intake");
  if (!intake) throw new Error('Missing sheet: "Canvas Intake"');

  intake.appendRow([
    new Date(),
    "debug:key",
    "debug profile",
    "debug course",
    "Announcement",
    "Debug intake write works",
    new Date().toISOString(),
    "https://example.com",
    "debug body"
  ]);

  Logger.log("✅ Debug row appended to Canvas Intake.");
}


function canonicalizeUrl_(u) {
  if (!u) return "";
  let s = String(u).trim();
  if (!/^https?:\/\//i.test(s)) return s;

  // strip common tracking query params
  const parts = s.split("?");
  let base = parts[0].replace(/\/+$/, ""); // no trailing slash
  if (parts.length === 1) return base;

  const query = parts[1]
    .split("&")
    .filter(kv => kv && !/^utm_/i.test(kv) && !/^fbclid=/i.test(kv) && !/^gclid=/i.test(kv))
    .join("&");

  return query ? (base + "?" + query) : base;
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

function cacheSetJson_(key, obj, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  const str = JSON.stringify(obj);
  // Cache has size limits; if too big, fallback to Properties (chunked)
  try {
    cache.put(key, str, ttlSeconds);
    return true;
  } catch (e) {
    // fallback: chunk into ScriptProperties
    const props = PropertiesService.getScriptProperties();
    const chunkSize = 8000; // safe-ish
    const chunks = [];
    for (let i = 0; i < str.length; i += chunkSize) chunks.push(str.slice(i, i + chunkSize));
    props.setProperty(key + ":n", String(chunks.length));
    chunks.forEach((c, i) => props.setProperty(key + ":" + i, c));
    props.setProperty(key + ":exp", String(Date.now() + ttlSeconds * 1000));
    return true;
  }
}

function cacheGetJson_(key) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const props = PropertiesService.getScriptProperties();
  const exp = Number(props.getProperty(key + ":exp") || "0");
  if (!exp || Date.now() > exp) return null;

  const n = Number(props.getProperty(key + ":n") || "0");
  if (!n) return null;

  let str = "";
  for (let i = 0; i < n; i++) str += (props.getProperty(key + ":" + i) || "");
  return str ? JSON.parse(str) : null;
}

function extractIncomeCap_(t) {
  // matches: "income under $80,000" / "household income less than 100000"
  const m = t.match(/\b(household|family)?\s*income\s*(under|less than|below|<=?)\s*\$?\s*([\d,]{4,})\b/i);
  if (!m) return null;
  const n = Number(String(m[3]).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}

function extractSaiCap_(t) {
  // matches: "SAI under 15000" / "Student Aid Index <= 20000"
  const m = t.match(/\b(sai|student aid index)\s*(under|less than|below|<=?)\s*([\d,]{3,})\b/i);
  if (!m) return null;
  const n = Number(String(m[3]).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}

function extractAnyDate_(t) {
  // Conservative: looks for YYYY-MM-DD or MM/DD/YYYY or "Mar 1, 2026"
  // Returns Date or null.
  let m = t.match(/\b(20\d{2})[-\/](\d{1,2})[-\/](\d{1,2})\b/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(d.getTime()) ? null : d;
  }

  m = t.match(/\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/);
  if (m) {
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    const d = new Date(y, Number(m[1]) - 1, Number(m[2]));
    return isNaN(d.getTime()) ? null : d;
  }

  m = t.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+(\d{1,2})(,)?\s+(20\d{2})\b/i);
  if (m) {
    const monthMap = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,sept:8,oct:9,nov:10,dec:11};
    const mm = monthMap[m[1].toLowerCase()];
    const d = new Date(Number(m[4]), mm, Number(m[2]));
    return isNaN(d.getTime()) ? null : d;
  }

  return null;
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

function debugCanvasCourseNames() {
  const url = CANVAS_BASE + "/api/v1/courses?per_page=100&fields[]=id&fields[]=name";
  const courses = canvasFetchJson_(url) || [];
  Logger.log("Course count: " + courses.length);
  courses.forEach(c => Logger.log(`${c.id} | ${c.name}`));
}

function safeGetEventById_(cal, eventId) {
  if (!cal || !eventId) return null;
  try {
    return cal.getEventById(eventId);
  } catch (e) {
    return null; // prevents crash if event doesn't exist
  }
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
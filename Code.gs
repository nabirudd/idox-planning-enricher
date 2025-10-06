/***** Ealing (Idox Public Access) scraper — runs until canceled, column J output
 * Put planning URLs (?keyVal=...) in column B (from B2 down).
 *
 * Menus:
 *  - Run now (continuous pass)          → runs ~5.9 minutes straight, then stops
 *  - Start "Run until canceled"         → auto-runs every minute until you stop it
 *  - Stop "Run until canceled"
 *  - Scrape active row                  → only the selected row
 *
 * Behavior:
 *  - Outputs J→O:  Application Received, Decision Issued Date, Documents sentence,
 *                  Agent Name, Agent Email, Agent Telephone
 *  - Helpers P/Q:  Status, Next Try After (for 1-minute backoff on slow/error rows)
 *  - Skips Documents tab unless Summary lacks the documents sentence (faster)
 ****************************************************************************************/

// Columns
const COL_B_URL = 2;                 // B
const OUTPUT_START_COLUMN = 10;      // J
const STATUS_COL = 16;               // P
const NEXTTRY_COL = 17;              // Q

// Run control
const TIME_BUDGET_MS = 355000;       // ~5.9 minutes per run (safe under Apps Script limit)
const MAX_ROWS_PER_RUN = 1500;       // hard cap per run (safety)
const POINTER_KEY = "NEXT_ROW_PTR";  // script property to remember where to resume
const RUN_UNTIL_FLAG = "RUN_UNTIL_CANCELED"; // "1" while background loop is enabled

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Planning Scraper")
    .addItem("Run now (continuous pass)", "runContinuouslyNow_")
    .addSeparator()
    .addItem("Start \"Run until canceled\"", "startRunUntilCanceled_")
    .addItem("Stop \"Run until canceled\"", "stopRunUntilCanceled_")
    .addSeparator()
    .addItem("Scrape active row", "scrapeActiveRow")
    .addToUi();
}

/* ========================= Public actions ========================= */

function scrapeActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) throw new Error("Please select a row ≥ 2.");
  writeHeaders_(sheet);
  tryScrapeRow_(sheet, row);
}

function runContinuouslyNow_() {
  const sheet = SpreadsheetApp.getActiveSheet();
  writeHeaders_(sheet);
  const res = runBatch_(sheet, /*respectTime*/ true);
  SpreadsheetApp.getActive().toast(
    `Processed ${res.processed} row(s). Next pointer: row ${res.nextPtr}. ${res.done ? "All done ✅" : ""}`
  );
}

function startRunUntilCanceled_() {
  // Turn on flag and ensure a 1-minute trigger exists
  PropertiesService.getScriptProperties().setProperty(RUN_UNTIL_FLAG, "1");
  ensureMinuteTrigger_();
  SpreadsheetApp.getActive().toast("Background loop started (runs every minute until you stop).");
}

function stopRunUntilCanceled_() {
  PropertiesService.getScriptProperties().deleteProperty(RUN_UNTIL_FLAG);
  removeMinuteTriggers_();
  SpreadsheetApp.getActive().toast("Background loop stopped.");
}

// Called by the time-based trigger every minute when enabled
function autoRunner_() {
  const flag = PropertiesService.getScriptProperties().getProperty(RUN_UNTIL_FLAG);
  if (flag !== "1") { removeMinuteTriggers_(); return; }

  // Prevent overlapping runs
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // another run is in progress
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    writeHeaders_(sheet);
    const res = runBatch_(sheet, /*respectTime*/ true);
    if (res.done) {
      // If you want it to keep watching for new rows, leave triggers running.
      // If you prefer to stop when done, uncomment the next two lines:
      // PropertiesService.getScriptProperties().deleteProperty(RUN_UNTIL_FLAG);
      // removeMinuteTriggers_();
    }
  } finally {
    lock.releaseLock();
  }
}

/* ========================= Core batch engine ========================= */

function runBatch_(sheet, respectTime) {
  const props = PropertiesService.getScriptProperties();
  let startRow = Number(props.getProperty(POINTER_KEY) || 2);
  const lastUrlRow = getLastUrlRow_(sheet);
  const startTime = Date.now();

  if (lastUrlRow < 2 || startRow > lastUrlRow) {
    props.setProperty(POINTER_KEY, String(lastUrlRow + 1));
    return { processed: 0, nextPtr: lastUrlRow + 1, done: true };
  }

  let processed = 0;
  let r = startRow;

  while (r <= lastUrlRow && processed < MAX_ROWS_PER_RUN) {
    // Stop before we hit Apps Script’s wall
    if (respectTime && Date.now() - startTime > TIME_BUDGET_MS) break;

    const url = String(sheet.getRange(r, COL_B_URL).getValue() || "").trim();
    if (url) {
      // Skip if “Next Try After” is in the future
      const nextTryVal = sheet.getRange(r, NEXTTRY_COL).getValue();
      if (nextTryVal && new Date(nextTryVal).getTime() > Date.now()) {
        r++;
        continue;
      }
      // Skip if already has output (remove this check to re-scrape filled rows)
      const hasOutput = String(sheet.getRange(r, OUTPUT_START_COLUMN).getValue() || "").trim() !== "";
      if (!hasOutput) {
        tryScrapeRow_(sheet, r);
        processed++;
      }
    }
    r++;
  }

  const nextPtr = r;
  props.setProperty(POINTER_KEY, String(nextPtr));
  const done = nextPtr > lastUrlRow && processed === 0;
  return { processed: processed, nextPtr: nextPtr, done: done };
}

/* ========================= Single-row scrape ========================= */

function tryScrapeRow_(sheet, row) {
  const rawUrl = String(sheet.getRange(row, COL_B_URL).getValue() || "").trim();
  if (!rawUrl) return false;

  const keyVal = getKeyVal_(rawUrl);
  if (!keyVal) {
    setStatus_(sheet, row, "ERROR: missing keyVal", 1);
    return false;
  }

  const base = "https://pam.ealing.gov.uk/online-applications/applicationDetails.do";
  const urlSummary  = `${base}?activeTab=summary&keyVal=${encodeURIComponent(keyVal)}`;
  const urlContacts = `${base}?activeTab=contacts&keyVal=${encodeURIComponent(keyVal)}`;
  const urlDocs     = `${base}?activeTab=documents&keyVal=${encodeURIComponent(keyVal)}`;

  let htmlSummary, htmlContacts, docsSentence = "";

  try {
    // SUMMARY first (fastest)
    htmlSummary = fetchWithRetry_(urlSummary);
  } catch (e) {
    setStatus_(sheet, row, `RETRY 1m (Summary: ${e.message})`, 1);
    return false;
  }

  // Pull from SUMMARY
  const appReceived    = extractLabel_(htmlSummary, "Application Received");
  const decisionIssued = extractLabel_(htmlSummary, "Decision Issued Date");
  docsSentence         = extractDocsSentenceFromOnePage_(htmlSummary);

  // Only hit DOCUMENTS tab if needed
  if (!docsSentence) {
    try {
      const htmlDocs = fetchWithRetry_(urlDocs);
      docsSentence = extractDocsSentenceFromOnePage_(htmlDocs);
    } catch (e) {
      // If docs fetch fails, we’ll still continue; leave docs blank this pass
      docsSentence = "";
    }
  }

  // CONTACTS always needed
  try {
    htmlContacts = fetchWithRetry_(urlContacts);
  } catch (e) {
    setStatus_(sheet, row, `RETRY 1m (Contacts: ${e.message})`, 1);
    return false;
  }
  const agent = extractAgentDetails_(htmlContacts);

  // Write outputs J→O
  const rowValues = [
    appReceived || "",
    decisionIssued || "",
    docsSentence || "",
    agent.name || "",
    agent.email || "",
    agent.telephone || ""
  ];
  sheet.getRange(row, OUTPUT_START_COLUMN, 1, rowValues.length).setValues([rowValues]);

  // Clear helper columns and mark OK
  sheet.getRange(row, NEXTTRY_COL).clearContent();
  sheet.getRange(row, STATUS_COL).setValue("OK");
  return true;
}

/* ========================= Headers & helpers ========================= */

function writeHeaders_(sheet) {
  const headers = [
    "Application Received",
    "Decision Issued Date",
    "Documents Associated",
    "Agent Name",
    "Agent Email",
    "Agent Telephone",
    "Status",            // P
    "Next Try After"     // Q
  ];
  sheet.getRange(1, OUTPUT_START_COLUMN, 1, headers.length).setValues([headers]);
}

function setStatus_(sheet, row, msg, retryInMinutes) {
  sheet.getRange(row, STATUS_COL).setValue(msg);
  if (retryInMinutes && retryInMinutes > 0) {
    const next = new Date(Date.now() + retryInMinutes * 60 * 1000);
    sheet.getRange(row, NEXTTRY_COL).setValue(next);
  } else {
    sheet.getRange(row, NEXTTRY_COL).clearContent();
  }
}

function getLastUrlRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const colB = sheet.getRange(2, COL_B_URL, lastRow - 1, 1).getValues();
  for (let i = colB.length - 1; i >= 0; i--) {
    if (String(colB[i][0]).trim() !== "") return i + 2;
  }
  return 1;
}

function getKeyVal_(url) {
  const m = url.match(/[?&]keyVal=([^&#]+)/i);
  return m ? decodeURIComponent(m[1]) : "";
}

/* ========================= Networking ========================= */

function fetchWithRetry_(url) {
  const opts = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    }
  };
  const waits = [0, 120, 300, 600]; // ms backoff (short, so we don't block a row)
  let lastErr = null;

  for (let i = 0; i < waits.length; i++) {
    try {
      const resp = UrlFetchApp.fetch(url, opts);
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) return resp.getContentText();
      lastErr = new Error("HTTP " + code);
    } catch (e) {
      lastErr = e;
    }
    if (waits[i] > 0) Utilities.sleep(waits[i]);
  }
  throw new Error(lastErr && lastErr.message ? lastErr.message : "fetch failed");
}

/* ========================= Parsers ========================= */

function extractLabel_(html, label) {
  const esc = label.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
  const patterns = [
    new RegExp("<dt[^>]*>\\s*" + esc + "\\s*<\/dt>\\s*<dd[^>]*>([\\s\\S]*?)<\/dd>", "i"),
    new RegExp("<th[^>]*>\\s*" + esc + "\\s*<\/th>\\s*<td[^>]*>([\\s\\S]*?)<\/td>", "i"),
    new RegExp("<td[^>]*>\\s*" + esc + "\\s*<\/td>\\s*<td[^>]*>([\\s\\S]*?)<\/td>", "i")
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = html.match(patterns[i]);
    if (m && m[1]) return cleanText_(m[1]);
  }
  return "";
}

function cleanText_(raw) {
  let txt = String(raw)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, " ");
  const ents = { "&nbsp;": " ", "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'" };
  txt = txt.replace(/(&nbsp;|&amp;|&lt;|&gt;|&quot;|&#39;)/g, m => ents[m] || m);
  return txt.replace(/\s+/g, " ").trim();
}

function extractByLabels_(html, labels) {
  for (let i = 0; i < labels.length; i++) {
    const esc = labels[i].replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
    let m = html.match(new RegExp("<dt[^>]*>\\s*" + esc + "\\s*:?\\s*<\/dt>\\s*<dd[^>]*>([\\s\\S]*?)<\/dd>", "i"));
    if (m && m[1]) return cleanText_(m[1]);
    m = html.match(new RegExp("<th[^>]*>\\s*" + esc + "\\s*:?\\s*<\/th>\\s*<td[^>]*>([\\s\\S]*?)<\/td>", "i"));
    if (m && m[1]) return cleanText_(m[1]);
    m = html.match(new RegExp("<td[^>]*>\\s*" + esc + "\\s*:?\\s*<\/td>\\s*<td[^>]*>([\\s\\S]*?)<\/td>", "i"));
    if (m && m[1]) return cleanText_(m[1]);
  }
  return "";
}

function extractEmail_(html) {
  const m = html.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0] : "";
}

function extractTelephone_(html) {
  // tel: link
  let m = html.match(/href=["']tel:([^"']+)["']/i);
  if (m && m[1]) return m[1].replace(/\s+/g, " ").trim();

  // Loose UK pattern
  const text = cleanText_(html);
  m = text.match(/(?:\+44|0)\s?\d(?:[\s\-\(\)]?\d){8,}/);
  if (m && m[0]) return m[0].replace(/\s{2,}/g, " ").trim();

  return "";
}

function extractAgentDetails_(html) {
  const out = { name: "", email: "", telephone: "" };

  // Section after "Agent" heading
  const afterHeading = (html.split(/<h[23][^>]*>\s*Agent\s*<\/h[23]>/i)[1] || "");
  const untilNextBlock = (afterHeading.match(/^([\s\S]*?)(?:<table|<dl|<div|<h[23])/i) || ["",""])[1];
  const plain = cleanText_(untilNextBlock).split(/\n/).map(s => s.trim()).filter(Boolean);
  if (plain.length) out.name = plain[0];

  if (!out.name) {
    const labeledName = extractByLabels_(afterHeading, ["Name","Agent Name","Contact Name"]);
    if (labeledName) out.name = labeledName;
  }

  let email = extractByLabels_(afterHeading, ["Email","Agent Email"]);
  if (!email) email = extractEmail_(afterHeading || html);
  if (email) out.email = email;

  let tel = extractByLabels_(afterHeading, ["Telephone number","Telephone","Phone","Tel","Mobile","Agent Telephone"]);
  if (!tel) tel = extractTelephone_(afterHeading || html);
  if (tel) out.telephone = tel;

  return out;
}

function extractDocsSentenceFromOnePage_(html) {
  const text = cleanText_(html);
  let m = text.match(/There\s+(?:is|are)\s+(\d+)\s+document(?:s)?\s+associated\s+with\s+this\s+application\.?/i);
  if (m && m[1]) {
    const n = Number(m[1]);
    const verb = n === 1 ? "is" : "are";
    const plural = n === 1 ? "document" : "documents";
    return `There ${verb} ${n} ${plural} associated with this application.`;
  }
  // Link form like “14 documents”
  m = html.match(/>\s*(\d+)\s+documents\s*<\/a>/i);
  if (m && m[1]) {
    const n2 = Number(m[1]);
    const verb2 = n2 === 1 ? "is" : "are";
    const plural2 = n2 === 1 ? "document" : "documents";
    return `There ${verb2} ${n2} ${plural2} associated with this application.`;
  }
  return "";
}

/* ========================= Trigger helpers ========================= */

function ensureMinuteTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === "autoRunner_");
  if (!exists) ScriptApp.newTrigger("autoRunner_").timeBased().everyMinutes(1).create();
}
function removeMinuteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autoRunner_") ScriptApp.deleteTrigger(t);
  });
}

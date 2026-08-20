// ─── 69 Event Tracker — Google Apps Script Backend ───────────────────────────
// Paste ALL of this into your Apps Script editor, then deploy as a Web App.
// See Sheets_Setup_Guide.md for the full step-by-step.
//
// Storage: the entire app's data (pins, divisions, every Mini/Main event's
// tips + last-3 history, and any pending submissions) lives as one JSON
// string in cell AppData!A1. At this app's scale — a few dozen events, each
// keeping only its last 3 scores — that comfortably stays under Google's
// 50,000-character-per-cell limit, so there's no need for the multi-cell
// splitting some larger clan tools use.
// ───────────────────────────────────────────────────────────────────────────

const SHEET_NAME = "AppData";

// ── Read all data ─────────────────────────────────────────────────────────
function doGet(e) {
  try { return jsonResponse(readData()); }
  catch (err) { return jsonResponse({ error: err.message }); }
}

// ── Write all data ────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const incoming = JSON.parse(e.postData.contents);
    writeData(incoming);
    return jsonResponse({ ok: true });
  } catch (err) { return jsonResponse({ error: err.message }); }
}

// ── Sheet helper ──────────────────────────────────────────────────────────
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue("");
  }
  return sheet;
}

function readData() {
  var sheet = getSheet();
  var raw = sheet.getRange("A1").getValue();
  if (!raw) return {};
  try { return JSON.parse(raw); }
  catch (err) { return { error: "Stored data in AppData!A1 is not valid JSON" }; }
}

function writeData(data) {
  var sheet = getSheet();
  var json = JSON.stringify(data);
  if (json.length > 45000) {
    // Well under the 50K cell limit at normal usage — this just flags it
    // early if it ever grows a lot (e.g. a big backlog of never-approved
    // pending submissions piling up).
    Logger.log("Warning: AppData!A1 is " + json.length + " chars — approaching the 50,000 cell limit.");
  }
  sheet.getRange("A1").setValue(json);
}

function jsonResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── Diagnostic — run manually from the Apps Script editor's function
// dropdown if you ever want to check how big the stored data has grown. ────
function diagnose() {
  var sheet = getSheet();
  var raw = sheet.getRange("A1").getValue();
  var len = raw ? raw.length : 0;
  Logger.log("AppData!A1: " + len + " chars (" + Math.round(len / 500) + "% of the 50,000 limit)");
  try {
    var d = JSON.parse(raw);
    Logger.log("miniEvents: " + Object.keys(d.miniEvents || {}).length);
    Logger.log("mainEvents: " + Object.keys(d.mainEvents || {}).length);
    Logger.log("divisions: " + (d.divisions || []).length);
    Logger.log("pending: " + (d.pending || []).length);
  } catch (err) {
    Logger.log("Could not parse stored JSON: " + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// EVENT TIMING SCRAPER
// Pulls "running now" / "next occurrence" for every event from
// totalcalculator.org/events.php and caches it into AppData!A1's `schedule`
// field, so the app itself never has to fetch that site directly.
//
// IMPORTANT: this was built without being able to inspect the site's raw
// HTML directly (only an AI-summarized read of it), so the parsing below is
// a best effort based on the visible text patterns it reported. The FIRST
// TIME you run scrapeEventSchedule() manually, check the execution log
// (View → Logs, or the "Executions" panel) — it logs the exact text
// snippet it matched for every single event. If any snippet looks wrong
// (garbled, empty, or clearly not a duration), something about the site's
// actual markup differs from what was assumed — copy the log output back so
// the parsing regex can be corrected.
// ───────────────────────────────────────────────────────────────────────────

const SCRAPE_URL = "https://totalcalculator.org/events.php";

// Keep these two lists in sync with MINI_EVENTS / MAIN_EVENTS in index.html —
// they're duplicated here because Apps Script and the HTML file can't share
// a JS module. Checked against the live site on 20 Aug 2026.
const SCRAPE_MINI_EVENTS = [
  "Castle Development", "Scientific Progress", "Capital Challenge",
  "Blessing of the Gods", "Officer Academy", "Hammer and Anvil",
  "Tar Mastery", "Regular Decrees", "Battle Training", "Power Points",
  "Silver Rush", "Wargames", "Gold Rush", "Call of Duty", "Beastslayer",
  "War Tools", "Crypt Raiders", "The Quest for Chests", "The King's Mercy",
];
// "Thirst for Battle" is a Weekly event in index.html's MAIN_EVENTS but is
// deliberately NOT scraped here — it always runs at the exact same time as
// "Clash for the Throne" (confirmed by Kirsty), so index.html aliases its
// timing lookup to that event's scraped data instead (see
// EVENT_TIMING_ALIAS in index.html) rather than scraping it twice.
const SCRAPE_MAIN_EVENTS_BY_CATEGORY = {
  monthly:  ["Ragnarok", "Armageddon", "Dark Omens", "Shadow Invasion", "Hellforge", "Trials of Olympus"],
  weekly:   ["Doomsday", "Arachne", "Ancient's Treasure", "Rise of the Ancients"],
  biweekly: ["Clash for the Throne", "Clash of Kingdoms"],
};

// The page's sections appear in this order in the text (order confirmed
// live 20 Aug 2026), but we locate each by name and sort by position rather
// than assuming a fixed order, so a reshuffle on the site doesn't break this.
const SCRAPE_SECTION_HEADERS = [
  { key: "monthly",  label: "Monthly events" },
  { key: "biweekly", label: "Biweekly events" },
  { key: "weekly",   label: "Weekly events" },
  { key: "mini",     label: "Mini events" },
  { key: "summon",   label: "Summon mastery" }, // boundary marker only — not scraped
];

function stripHtmlToText(html) {
  var text = html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ");
  text = text
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&#39;/g, "'")
    .replace(/&apos;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");
  return text;
}

// "1 day 4h56m" / "4h26m" / "56m" → milliseconds. Returns null if nothing
// matched at all (vs. 0, which would be a real "starting now" case).
//
// Each unit is searched for independently (not as one sequential pattern)
// because a single combined regex with every part optional can match an
// empty/partial prefix and stop — e.g. it would silently match just the
// whitespace padding before "1 day 1h26m" and never get to the numbers at
// all. Searching for \d+ days? / \d+h / \d+m separately sidesteps that.
function parseDurationMs(snippet) {
  var dMatch = snippet.match(/(\d+)\s*days?/i);
  var hMatch = snippet.match(/(\d+)h/i);
  var mMatch = snippet.match(/(\d+)m/i);
  var days = dMatch ? parseInt(dMatch[1], 10) : 0;
  var hours = hMatch ? parseInt(hMatch[1], 10) : 0;
  var mins = mMatch ? parseInt(mMatch[1], 10) : 0;
  if (!days && !hours && !mins) return null;
  return ((days * 24 + hours) * 60 + mins) * 60000;
}

// Slice `text` into { sectionKey: sectionText } using SCRAPE_SECTION_HEADERS,
// locating each header by searching for it rather than assuming positions.
function splitIntoSections(text) {
  var found = SCRAPE_SECTION_HEADERS
    .map(function(h) { return { key: h.key, idx: text.indexOf(h.label), labelLen: h.label.length }; })
    .filter(function(h) { return h.idx !== -1; })
    .sort(function(a, b) { return a.idx - b.idx; });
  var sections = {};
  for (var i = 0; i < found.length; i++) {
    var start = found[i].idx + found[i].labelLen;
    var end = (i + 1 < found.length) ? found[i + 1].idx : text.length;
    sections[found[i].key] = text.slice(start, end);
  }
  return sections;
}

// For a section's text and its list of known event names, return
// { name: durationSnippet } — the text between each name and whichever
// comes next (next known name, "CURRENT:", or end of section).
function extractDurationSnippets(sectionText, names) {
  var currentIdx = sectionText.indexOf("CURRENT:");
  var listText = currentIdx !== -1 ? sectionText.slice(0, currentIdx) : sectionText;

  var positions = names
    .map(function(n) { return { name: n, idx: listText.indexOf(n) }; })
    .filter(function(p) { return p.idx !== -1; })
    .sort(function(a, b) { return a.idx - b.idx; });

  var out = {};
  for (var i = 0; i < positions.length; i++) {
    var start = positions[i].idx + positions[i].name.length;
    var end = (i + 1 < positions.length) ? positions[i + 1].idx : listText.length;
    out[positions[i].name] = listText.slice(start, end);
  }
  return out;
}

// Which of `names` are currently running, per this section's "CURRENT:"
// marker(s). Matches by simple substring search in the text after the
// first "CURRENT:", which is robust whether multiple current events are
// separated by "and", a comma, or back-to-back "CURRENT:" markers.
function findRunningNames(sectionText, names) {
  var running = {};
  var idx = sectionText.indexOf("CURRENT:");
  if (idx === -1) return running;
  var tail = sectionText.slice(idx);
  names.forEach(function(name) {
    if (tail.indexOf(name) !== -1) running[name] = true;
  });
  return running;
}

function scrapeEventSchedule() {
  var html = UrlFetchApp.fetch(SCRAPE_URL, { muteHttpExceptions: true }).getContentText();
  var text = stripHtmlToText(html).replace(/\s+/g, " ");
  Logger.log("Fetched " + html.length + " raw HTML chars → " + text.length + " plain-text chars after stripping tags.");

  var now = new Date();
  var sections = splitIntoSections(text);
  Logger.log("Sections found: " + Object.keys(sections).join(", "));

  var items = [];

  function processSection(sectionKey, names, category) {
    var sectionText = sections[sectionKey];
    if (!sectionText) {
      Logger.log("⚠️ Section \"" + sectionKey + "\" not found on the page — skipping " + category + " events.");
      return;
    }
    var snippets = extractDurationSnippets(sectionText, names);
    var running = findRunningNames(sectionText, names);
    names.forEach(function(name) {
      var snippet = snippets[name];
      if (snippet === undefined) {
        Logger.log("⚠️ " + category + " | \"" + name + "\" — name not found on the page at all. Check spelling/casing matches the site.");
        items.push({ name: name, category: category, running: !!running[name], nextStart: null });
        return;
      }
      var ms = parseDurationMs(snippet);
      var nextStart = ms != null ? new Date(now.getTime() + ms).toISOString() : null;
      Logger.log(category + " | " + name + " | snippet=\"" + snippet.slice(0, 24).trim() + "\" | parsedMs=" + ms + " | running=" + !!running[name]);
      items.push({ name: name, category: category, running: !!running[name], nextStart: nextStart });
    });
  }

  processSection("mini", SCRAPE_MINI_EVENTS, "mini");
  processSection("monthly", SCRAPE_MAIN_EVENTS_BY_CATEGORY.monthly, "monthly");
  processSection("weekly", SCRAPE_MAIN_EVENTS_BY_CATEGORY.weekly, "weekly");
  processSection("biweekly", SCRAPE_MAIN_EVENTS_BY_CATEGORY.biweekly, "biweekly");

  var data = readData();
  data.schedule = { lastScraped: now.toISOString(), items: items };
  writeData(data);

  Logger.log("Done — cached " + items.length + " events. lastScraped=" + now.toISOString());
  return items;
}

// ── Run once manually (function dropdown → createHourlyScrapeTrigger → Run)
// to set up automatic hourly scraping. Safe to re-run — clears any existing
// scrape trigger first so it never creates duplicates. ─────────────────────
function createHourlyScrapeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "scrapeEventSchedule") ScriptApp.deleteTrigger(t);
  });
  // .nearMinute(2) asks Apps Script to fire close to 2 minutes past each hour
  // instead of at a random offset anywhere in the hour (the default for
  // plain everyHours(1)). That keeps the cached schedule close behind the
  // source site's own hourly reset — a couple of minutes' margin so the
  // page has settled, instead of drifting up to ~59 minutes stale.
  // Apps Script only guarantees this within roughly a 15-minute window, not
  // to the exact minute, but it's the closest control it offers.
  ScriptApp.newTrigger("scrapeEventSchedule").timeBased().everyHours(1).nearMinute(2).create();
  Logger.log("Hourly scrape trigger created — scrapeEventSchedule() will now run automatically near the top of every hour (~2 minutes past).");
}

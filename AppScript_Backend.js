// ─── 69 Event Tracker — Google Apps Script Backend ───────────────────────────
// Paste ALL of this into your Apps Script editor, then deploy as a Web App.
// See Sheets_Setup_Guide.md for the full step-by-step.
//
// Storage (rewritten 2026-08-23 — was one JSON blob in a single AppData!A1
// cell; Kirsty asked for it split into separate rows so no single write can
// ever hold everything at once). The app's data now lives across several
// sheets (tabs) in this spreadsheet, one row per record:
//   Pins            — role | pin                              (4 rows)
//   Meta            — key | value                              (lastScraped, savedAt)
//   Divisions       — division                                 (one row each)
//   MiniEvents      — name | tip | rewardFirst | rewardRest     (one row per event)
//   MainEvents      — same columns as MiniEvents
//   EventHistory    — type | eventName | id | date | highest | lowest | submittedBy | division
//   Schedule        — category | name | running | nextStart    (rewritten by the scraper)
//   Pending         — id | category | eventName | highest | lowest | division | submittedBy | submittedAt
//   Epics           — id | event | day | name | captainHero | numCaptains | meat | chests
//   EpicMonsters    — epicId | id | order | name | type1 | type2 | type3 | strengthAgainst |
//                      pctStrengthAgainst | initiative | strength | health | leadership |
//                      quantity | totalHealth | totalStrength
//   Troops          — id | type1 | level | name | types | type2 | strengthAgainst1 |
//                      strengthAgainst2 | strength | health | leadership | authority |
//                      dominance | costPerUnit | groupOrder
//   Faqs            — id | category | question | answer
//   FaqCategories   — category
//   Tools           — id | category | name | url | guide | comingSoon
//   ToolCategories  — category
//
// Every cell holds JSON.stringify(value) for that one field, in a column
// forced to plain-text number format. That's deliberate, not decorative: it's
// what stops Sheets from "helpfully" reinterpreting a value as a number,
// date, or boolean on its own — e.g. a PIN like "0500" silently losing its
// leading zero, an ISO timestamp turning into a Sheets date serial, or a
// troop stat like "110,200" losing its comma. JSON-encoding + text format
// makes every field round-trip byte-for-byte regardless of what it looks
// like, without having to hand-classify each column's "real" type. The
// doGet/doPost contract (and the JSON shape they read/write) is UNCHANGED
// from before — index.html needed no changes for this — only how that JSON
// is actually stored in the spreadsheet changed.
//
// UPGRADING FROM THE OLD SINGLE-CELL LAYOUT: after pasting this new version
// in, run migrateFromSingleCell() once (function dropdown → run) to copy your
// existing AppData!A1 data into the new sheets. It does NOT delete anything —
// the old sheet is renamed to "AppData_OLD_BACKUP" and left in place as a
// safety copy; delete it yourself once you've checked the new sheets look
// right.
// ───────────────────────────────────────────────────────────────────────────

const OLD_SHEET_NAME = "AppData"; // only used by migrateFromSingleCell() below

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

function jsonResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ═══ Generic per-record-rows storage helpers ═══════════════════════════════
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

// Reads every data row (below the header) into an array of plain objects
// keyed by `headers`. Blank cells become `undefined`; everything else is
// JSON.parse()d back to its original type (falls back to the raw text if a
// cell somehow isn't valid JSON, rather than throwing and losing the row).
function readRows(sheetName, headers) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var rows = [];
  values.forEach(function(row) {
    if (row.every(function(c) { return c === "" || c === null; })) return; // skip blank rows
    var obj = {};
    headers.forEach(function(h, i) {
      var raw = row[i];
      if (raw === "" || raw === null || raw === undefined) { obj[h] = undefined; return; }
      try { obj[h] = JSON.parse(raw); } catch (err) { obj[h] = raw; }
    });
    rows.push(obj);
  });
  return rows;
}

// Replaces every data row in `sheetName` with `rows` (array of plain
// objects keyed by `headers`) — a full clear + rewrite each save, same
// "whole set replaces whole set" semantics the old single-cell blob had, just
// scoped to one sheet/data-type at a time instead of the entire app at once.
function writeRows(sheetName, headers, rows) {
  var sheet = getOrCreateSheet(sheetName, headers);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length).clearContent();
  }
  if (!rows || rows.length === 0) return;
  sheet.getRange(2, 1, rows.length, headers.length).setNumberFormat("@");
  var values = rows.map(function(row) {
    return headers.map(function(h) {
      var v = row[h];
      return v === undefined ? "" : JSON.stringify(v);
    });
  });
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

// Key/value sheets (Pins, Meta) — same idea as readRows/writeRows but for a
// flat {key: value} object instead of a list of records.
function readKeyValue(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var out = {};
  values.forEach(function(row) {
    var key = row[0];
    if (!key) return;
    var raw = row[1];
    if (raw === "" || raw === null) { out[key] = undefined; return; }
    try { out[key] = JSON.parse(raw); } catch (err) { out[key] = raw; }
  });
  return out;
}

function writeKeyValue(sheetName, obj) {
  var sheet = getOrCreateSheet(sheetName, ["key", "value"]);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 2).clearContent();
  }
  var keys = Object.keys(obj || {});
  if (keys.length === 0) return;
  sheet.getRange(2, 1, keys.length, 2).setNumberFormat("@");
  var values = keys.map(function(k) { return [k, JSON.stringify(obj[k])]; });
  sheet.getRange(2, 1, values.length, 2).setValues(values);
}

// ═══ Per-sheet column layouts ═══════════════════════════════════════════════
const SHEET_SCHEMAS = {
  Divisions:      ["division"],
  MiniEvents:     ["name", "tip", "rewardFirst", "rewardRest"],
  MainEvents:     ["name", "tip", "rewardFirst", "rewardRest"],
  EventHistory:   ["type", "eventName", "id", "date", "highest", "lowest", "submittedBy", "division"],
  Schedule:       ["category", "name", "running", "nextStart"],
  Pending:        ["id", "category", "eventName", "highest", "lowest", "division", "submittedBy", "submittedAt"],
  Epics:          ["id", "event", "day", "name", "captainHero", "numCaptains", "meat", "chests"],
  EpicMonsters:   ["epicId", "id", "order", "name", "type1", "type2", "type3", "strengthAgainst",
                    "pctStrengthAgainst", "initiative", "strength", "health", "leadership",
                    "quantity", "totalHealth", "totalStrength"],
  Troops:         ["id", "type1", "level", "name", "types", "type2", "strengthAgainst1", "strengthAgainst2",
                    "strength", "health", "leadership", "authority", "dominance", "costPerUnit", "groupOrder"],
  Faqs:           ["id", "category", "question", "answer"],
  FaqCategories:  ["category"],
  Tools:          ["id", "category", "name", "url", "guide", "comingSoon"],
  ToolCategories: ["category"],
};

// ═══ Read — reassembles the same combined JSON shape index.html always
// expected from doGet, just sourced from many sheets instead of one cell. ═══
function readData() {
  var pins = readKeyValue("Pins");
  var meta = readKeyValue("Meta");

  var divisions = readRows("Divisions", SHEET_SCHEMAS.Divisions).map(function(r) { return r.division; });

  var historyRows = readRows("EventHistory", SHEET_SCHEMAS.EventHistory);

  function buildEvents(rows, type) {
    var out = {};
    rows.forEach(function(r) {
      out[r.name] = { tip: r.tip || "", rewards: { first: r.rewardFirst || "", rest: r.rewardRest || "" }, history: [] };
    });
    historyRows.filter(function(h) { return h.type === type; }).forEach(function(h) {
      if (!out[h.eventName]) out[h.eventName] = { tip: "", rewards: { first: "", rest: "" }, history: [] };
      out[h.eventName].history.push({
        id: h.id, date: h.date, highest: h.highest, lowest: h.lowest,
        submittedBy: h.submittedBy, division: h.division === undefined ? null : h.division,
      });
    });
    return out;
  }

  var miniEvents = buildEvents(readRows("MiniEvents", SHEET_SCHEMAS.MiniEvents), "mini");
  var mainEvents = buildEvents(readRows("MainEvents", SHEET_SCHEMAS.MainEvents), "main");

  var schedule = { lastScraped: meta.lastScraped || null, items: readRows("Schedule", SHEET_SCHEMAS.Schedule) };

  var pending = readRows("Pending", SHEET_SCHEMAS.Pending);

  var epicRows = readRows("Epics", SHEET_SCHEMAS.Epics);
  var monsterRows = readRows("EpicMonsters", SHEET_SCHEMAS.EpicMonsters);
  var epics = epicRows.map(function(e) {
    var monsters = monsterRows
      .filter(function(m) { return m.epicId === e.id; })
      .map(function(m) { var c = Object.assign({}, m); delete c.epicId; return c; });
    return Object.assign({}, e, { monsters: monsters });
  });

  var troops = readRows("Troops", SHEET_SCHEMAS.Troops);
  var faqs = readRows("Faqs", SHEET_SCHEMAS.Faqs);
  var faqCategories = readRows("FaqCategories", SHEET_SCHEMAS.FaqCategories).map(function(r) { return r.category; });
  var tools = readRows("Tools", SHEET_SCHEMAS.Tools);
  var toolCategories = readRows("ToolCategories", SHEET_SCHEMAS.ToolCategories).map(function(r) { return r.category; });

  return {
    pins: pins,
    divisions: divisions,
    miniEvents: miniEvents,
    mainEvents: mainEvents,
    schedule: schedule,
    pending: pending,
    epics: epics,
    troops: troops,
    faqs: faqs,
    faqCategories: faqCategories,
    tools: tools,
    toolCategories: toolCategories,
    savedAt: meta.savedAt || 0,
  };
}

// ═══ Write — decomposes the same combined JSON shape index.html always sent
// to doPost across the sheets above. Each sheet is fully replaced with the
// current set of records for that data type (same semantics as the old
// single-cell "whole blob replaces whole blob" write, just per-sheet). ═════
function writeData(data) {
  writeKeyValue("Pins", data.pins || {});
  writeKeyValue("Meta", {
    lastScraped: (data.schedule && data.schedule.lastScraped) || null,
    savedAt: data.savedAt || Date.now(),
  });

  writeRows("Divisions", SHEET_SCHEMAS.Divisions, (data.divisions || []).map(function(d) { return { division: d }; }));

  function eventsToRows(events) {
    return Object.keys(events || {}).map(function(name) {
      var e = events[name] || {};
      var r = e.rewards || {};
      return { name: name, tip: e.tip || "", rewardFirst: r.first || "", rewardRest: r.rest || "" };
    });
  }
  writeRows("MiniEvents", SHEET_SCHEMAS.MiniEvents, eventsToRows(data.miniEvents));
  writeRows("MainEvents", SHEET_SCHEMAS.MainEvents, eventsToRows(data.mainEvents));

  var historyRows = [];
  ["mini", "main"].forEach(function(type) {
    var events = (type === "mini" ? data.miniEvents : data.mainEvents) || {};
    Object.keys(events).forEach(function(name) {
      (events[name].history || []).forEach(function(h) {
        historyRows.push({
          type: type, eventName: name, id: h.id, date: h.date,
          highest: h.highest, lowest: h.lowest, submittedBy: h.submittedBy,
          division: h.division === undefined ? null : h.division,
        });
      });
    });
  });
  writeRows("EventHistory", SHEET_SCHEMAS.EventHistory, historyRows);

  writeRows("Schedule", SHEET_SCHEMAS.Schedule, (data.schedule && data.schedule.items) || []);

  writeRows("Pending", SHEET_SCHEMAS.Pending, (data.pending || []).map(function(p) {
    return {
      id: p.id, category: p.category, eventName: p.eventName,
      highest: p.highest, lowest: p.lowest, division: p.division === undefined ? null : p.division,
      submittedBy: p.submittedBy, submittedAt: p.submittedAt,
    };
  }));

  writeRows("Epics", SHEET_SCHEMAS.Epics, (data.epics || []).map(function(e) {
    return {
      id: e.id, event: e.event, day: e.day, name: e.name,
      captainHero: e.captainHero, numCaptains: e.numCaptains, meat: e.meat, chests: e.chests,
    };
  }));
  var monsterRows = [];
  (data.epics || []).forEach(function(e) {
    (e.monsters || []).forEach(function(m) { monsterRows.push(Object.assign({ epicId: e.id }, m)); });
  });
  writeRows("EpicMonsters", SHEET_SCHEMAS.EpicMonsters, monsterRows);

  writeRows("Troops", SHEET_SCHEMAS.Troops, data.troops || []);
  writeRows("Faqs", SHEET_SCHEMAS.Faqs, data.faqs || []);
  writeRows("FaqCategories", SHEET_SCHEMAS.FaqCategories, (data.faqCategories || []).map(function(c) { return { category: c }; }));
  writeRows("Tools", SHEET_SCHEMAS.Tools, data.tools || []);
  writeRows("ToolCategories", SHEET_SCHEMAS.ToolCategories, (data.toolCategories || []).map(function(c) { return { category: c }; }));
}

// ── One-time migration from the old single-cell layout (2026-08-23) ────────
// Run this ONCE from the function dropdown (▸ Run, with migrateFromSingleCell
// selected) after pasting this new version of the script in. Safe to re-run —
// if the old "AppData" sheet is already gone/renamed it just logs that
// there's nothing to do. Does NOT delete your existing data: the old sheet is
// renamed to "AppData_OLD_BACKUP" and left in the spreadsheet untouched.
function migrateFromSingleCell() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var oldSheet = ss.getSheetByName(OLD_SHEET_NAME);
  if (!oldSheet) {
    Logger.log("No old '" + OLD_SHEET_NAME + "' sheet found — nothing to migrate (already migrated, or this is a fresh setup).");
    return;
  }
  var raw = oldSheet.getRange("A1").getValue();
  if (!raw) {
    Logger.log("'" + OLD_SHEET_NAME + "' sheet exists but A1 is empty — nothing to migrate.");
    return;
  }
  var oldData;
  try { oldData = JSON.parse(raw); }
  catch (err) { Logger.log("Could not parse old " + OLD_SHEET_NAME + "!A1 as JSON — migration aborted, nothing changed: " + err.message); return; }

  Logger.log("Migrating from the old single-cell layout...");
  Logger.log("Found: " + Object.keys(oldData.miniEvents || {}).length + " mini events, "
    + Object.keys(oldData.mainEvents || {}).length + " main events, "
    + (oldData.divisions || []).length + " divisions, "
    + (oldData.pending || []).length + " pending, "
    + (oldData.epics || []).length + " epics, "
    + (oldData.troops || []).length + " troops, "
    + (oldData.faqs || []).length + " FAQs, "
    + (oldData.tools || []).length + " tools & links.");

  writeData(oldData);

  var backupName = "AppData_OLD_BACKUP";
  if (!ss.getSheetByName(backupName)) {
    oldSheet.setName(backupName);
    Logger.log("Old sheet renamed to '" + backupName + "' — kept as a safety copy, not deleted. Delete it yourself once you've confirmed the new sheets look right.");
  } else {
    Logger.log("A sheet named '" + backupName + "' already exists — old '" + OLD_SHEET_NAME + "' sheet left as-is (not renamed) to avoid overwriting that backup. Check both manually.");
  }

  Logger.log("Migration complete. Check the new sheets (Pins, Divisions, MiniEvents, MainEvents, EventHistory, Schedule, Pending, Epics, EpicMonsters, Troops, Faqs, FaqCategories, Tools, ToolCategories, Meta) before deleting the backup.");
}

// ── Diagnostic — run manually from the function dropdown to see row counts
// across every sheet. ───────────────────────────────────────────────────────
function diagnose() {
  var names = ["Pins", "Meta", "Divisions", "MiniEvents", "MainEvents", "EventHistory", "Schedule",
    "Pending", "Epics", "EpicMonsters", "Troops", "Faqs", "FaqCategories", "Tools", "ToolCategories"];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  names.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log(name + ": sheet does not exist yet (created on first save/migration)."); return; }
    Logger.log(name + ": " + Math.max(sheet.getLastRow() - 1, 0) + " row(s).");
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// EVENT TIMING SCRAPER
// Pulls "running now" / "next occurrence" for every event from
// totalcalculator.org/events.php and caches it into the Schedule sheet (via
// readData()/writeData() above), so the app itself never has to fetch that
// site directly. Unchanged by the 2026-08-23 storage restructure — it only
// ever talked to readData()/writeData(), never the sheet layout directly.
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
// BOTH "Clash for the Throne" and "Clash of Kingdoms" (confirmed by
// Kirsty), so index.html aliases its timing lookup to whichever of those
// two events' scraped data is currently running (or soonest upcoming) —
// see EVENT_TIMING_ALIAS in index.html — rather than scraping it a third
// time.
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

// ── Run once manually (function dropdown → createScrapeTrigger → Run) to
// set up automatic scraping. Safe to re-run — clears any existing scrape
// trigger first so it never creates duplicates. ────────────────────────────
//
// Was createHourlyScrapeTrigger() / .everyHours(1).nearMinute(2) — Kirsty
// observed an event ending at :30 past the hour, i.e. not every event's own
// cycle resets on the hour the way "Clash of Kingdoms"/"Clash for the
// Throne" do. A once-an-hour scrape left the cache up to ~57 minutes stale
// around a boundary like that (worse if the ~15-minute trigger jitter fell
// the wrong way). index.html's getEffectiveSchedule() already fills in the
// gaps BETWEEN scrapes live from the client's own clock, using each item's
// last-scraped `nextStart` as ground truth — but it can only be as accurate
// as that last scrape's data, so halving the scrape interval halves the
// worst-case staleness window too. 30 minutes is the shortest interval
// ScriptApp's minute-based triggers support without stepping down to 15/10/5
// (which would just burn quota for little real benefit given the client-side
// smoothing already in place).
function createScrapeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "scrapeEventSchedule") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("scrapeEventSchedule").timeBased().everyMinutes(30).create();
  Logger.log("Scrape trigger created — scrapeEventSchedule() will now run automatically about every 30 minutes.");
}

/**
 * Our206 — combined sheet/calendar automation
 *
 * Includes:
 * - Debounced future-event calendar sync
 * - Dry-run sync
 * - Daily maintenance: move past events -> Past Concerts, sort, compact, sync
 * - Incoming Raw processor: parse messy source lines, normalize, dedupe, append to Concerts
 * - Venue normalization via reference tab "Venue Map" if present, with baked-in fallback map
 * - Automatic copying of Google Form input to main concert tab
 * - Logging + custom menu
 *
 * Tabs used:
 * - Concerts
 * - Past Concerts
 * - Incoming Raw   (optional, for ingestion)
 * - Venue Map      (optional, for venue KV normalization)
 *
 * Venue Map format:
 *   A: Raw Venue
 *   B: Normalized Venue
 * Starting at row 2. Header row 1 can be anything.
 */

const SHEET_KEYS = {
  CONCERTS: "concerts",
  PAST_CONCERTS: "pastConcerts",
  INCOMING_RAW: "incomingRaw",
  VENUE_MAP: "venueMap"
};

const CFG = {
  calendarId: "our206wa@gmail.com",
  sheets: {
    [SHEET_KEYS.CONCERTS]: {
      name: "Concerts",
      headerRow: 3
    },
    [SHEET_KEYS.PAST_CONCERTS]: {
      name: "Past Concerts",
      headerRow: 3
    },
    [SHEET_KEYS.INCOMING_RAW]: {
      name: "Incoming Raw",
      headerRow: 1
    },
    [SHEET_KEYS.VENUE_MAP]: {
      name: "Venue Map",
      headerRow: 1
    }
  },
  headerRowFallback: 3,
  debounceMinutes: 10,
  debounceGuardMinutes: 8,
  horizonYears: 2,
  uidHeader: "UID",
  uidMarkerPrefix: "[our206_uid]:",
  headerMatchers: {
    date: "date",
    artist: "artist",
    venue: "venue",
    rating: "skoi",
    notes: "notes",
    ticket: "ticket",
    addedOn: "added"
  },
  keepLastColumnHeader: "cap"
};

// Fallback venue normalization map.
// If Venue Map tab exists, its values take precedence.
const VENUE_ADDRESS = {
  "Chop Suey": "1325 E Madison St, Seattle, WA 98122",
  "Clock-Out Lounge": "4864 Beacon Ave S, Seattle, WA 98108",
  "Edmonds Center for the Arts": "410 4th Ave N, Edmonds, WA 98020",
  "Hidden Hall": "400 N 35th St, Seattle, WA 98103",
  "Moore Theatre": "1932 2nd Ave, Seattle, WA 98101",
  "Nectar Lounge": "412 N 36th St, Seattle, WA 98103",
  "Neptune Theatre": "1303 NE 45th St, Seattle, WA 98105",
  "Neumos": "925 E Pike St, Seattle, WA 98122",
  "Paramount Theatre": "911 Pine St, Seattle, WA 98101",
  "Pony": "1221 E Madison St, Seattle, WA 98122",
  "Q Nightclub": "1426 Broadway, Seattle, WA 98122",
  "Showbox SoDo": "1700 1st Ave S, Seattle, WA 98134",
  "Substation Seattle": "645 NW 45th St, Seattle, WA 98107",
  "The Chapel": "4649 Sunnyside Ave N, Seattle, WA 98103",
  "The Crocodile": "2505 1st Ave, Seattle, WA 98121",
  "The Showbox": "1426 1st Ave, Seattle, WA 98101",
  "Town Hall Seattle": "1119 8th Ave, Seattle, WA 98101",
  "Tractor Tavern": "5213 Ballard Ave NW, Seattle, WA 98107",
  "Wheelie Pop Brewing": "",
  "Darrell’s Tavern": "",
  "The Triple Door": "",
  "El Corazon": "",
  "Belltown Yacht Club": "",
  "Airport Tavern": "",
  "Massive": "",
  "Real Art Tacoma": "",
  "WaMu Theater": "",
  "T-Mobile Park": "",
  "Barboza": "",
  "Stumpfest PDX": ""
};

const FALLBACK_VENUE_NORMALIZATION = {
  "sodo showbox": "Showbox SoDo",
  "showbox sodo": "Showbox SoDo",
  "showbox": "The Showbox",
  "croc": "The Crocodile",
  "the crocodile": "The Crocodile",
  "neptune": "Neptune Theatre",
  "paramount": "Paramount Theatre",
  "substation": "Substation Seattle",
  "nectar": "Nectar Lounge",
  "tractor": "Tractor Tavern",
  "chop": "Chop Suey",
  "clock-out": "Clock-Out Lounge",
  "clock-out lounge": "Clock-Out Lounge",
  "wheelie pop": "Wheelie Pop Brewing",
  "sunset": "The Sunset Tavern",
  "the moore": "Moore Theatre",
  "q": "Q Nightclub",
  "edmonds arts center": "Edmonds Center for the Arts",
  "edmonds center for the arts": "Edmonds Center for the Arts",
  "wamu": "WaMu Theater",
  "darrel’s tavern": "Darrell’s Tavern",
  "darrel's tavern": "Darrell’s Tavern",
  "massive": "Massive",
  "real art tacoma": "Real Art Tacoma",
  "airport tavern": "Airport Tavern",
  "t-mobile park": "T-Mobile Park",
  "barboza": "Barboza",
  "hidden hall": "Hidden Hall",
  "the triple door": "The Triple Door",
  "el corazon": "El Corazon",
  "belltown yacht club": "Belltown Yacht Club",
  "stumpfest pdx": "Stumpfest PDX",
  "neumos": "Neumos",
  "paramount theatre": "Paramount Theatre",
  "neptune theatre": "Neptune Theatre",
  "tractor tavern": "Tractor Tavern",
  "nectar lounge": "Nectar Lounge",
  "substation seattle": "Substation Seattle",
  "the showbox": "The Showbox",
  "showbox sodo": "Showbox SoDo"
};

// ---------------------- MENU ----------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Our206")
    .addItem("Set up (install triggers)", "setUpOur206")
    .addItem("Ensure Venue Map tab", "ensureVenueMapTab_our206")
    .addSeparator()
    .addItem("Process Incoming Raw", "processIncomingRaw_our206")
    .addSeparator()
    .addItem("Sync now", "syncUpcomingEvents")
    .addItem("Dry run sync (no calendar changes)", "dryRunSync")
    .addSeparator()
    .addItem("Move past events to Past Concerts", "movePastEvents")
    .addItem("Move past events + Sync now", "movePastEventsAndSync")
    .addSeparator()
    .addItem("Show last run log", "showLastRunLog")
    .addToUi();
}

function showLastRunLog() {
  const msg = PropertiesService.getScriptProperties().getProperty("LAST_RUN_LOG") || "(no log yet)";
  SpreadsheetApp.getUi().alert(msg);
}

// ---------------------- SETUP ----------------------

function setUpOur206() { return setupOur206(); }

function setupOur206() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureUidColumn_(getSheet_(ss, SHEET_KEYS.CONCERTS));
  ensureUidColumn_(ensureSheetByKey_(ss, SHEET_KEYS.PAST_CONCERTS));
  installOnEditTriggerIfMissing_();
  installDailyTriggerIfMissing_("our206_dailyMaintenance", 3);
  toast_("Our206 setup complete.");
}

function ensureVenueMapTab_our206() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = getOptionalSheet_(ss, SHEET_KEYS.VENUE_MAP);
  if (!sheet) sheet = insertSheetByKey_(ss, SHEET_KEYS.VENUE_MAP);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([["Raw Venue", "Normalized Venue", "Address (optional)"]]);
  } else if (sheet.getRange(1,1).getValue() === "") {
    sheet.getRange(1, 1, 1, 3).setValues([["Raw Venue", "Normalized Venue", "Address (optional)"]]);
  }

  // Seed with fallback normalization values if sheet is mostly empty
  const existingLastRow = sheet.getLastRow();
  if (existingLastRow <= 1) {
    const seeded = Object.keys(FALLBACK_VENUE_NORMALIZATION)
      .sort()
      .map(k => {
        const normalized = FALLBACK_VENUE_NORMALIZATION[k];
        const addr = VENUE_ADDRESS[normalized] || "";
        return [k, normalized, addr];
      });
    if (seeded.length) {
      sheet.getRange(2, 1, seeded.length, 3).setValues(seeded);
    }
  }

  toast_("Venue Map tab is ready.");
}

// ---------------------- TRIGGERS ----------------------

// Handle form submissions: copy form values from form sheet into main concert/event sheet.
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = getSheet_(ss, SHEET_KEYS.CONCERTS);

  const [
    timestamp,
    artist,
    date,
    venue,
    ticketLink,
    notes
  ] = e.range.getValues()[0];

  targetSheet.appendRow([
    coerceDate_(date) || date,
    artist,
    normalizeVenueName_(venue),
    "", // Skoi Rating
    notes,
    ticketLink,
    "", // UID reserved for existing calendar sync logic
    coerceTimestamp_(timestamp) || new Date() // Added On
  ]);

  applyConcertRowFormats_(targetSheet, targetSheet.getLastRow());
}

function our206_onEdit(e) {
  maybeSetAddedOn_(e);
  scheduleDebouncedSync_(e);
}

function our206_debouncedSync() {
  try {
    const lastEdit = Number(PropertiesService.getScriptProperties().getProperty("LAST_EDIT_TS") || "0");
    if (Date.now() - lastEdit < CFG.debounceGuardMinutes * 60 * 1000) return;
    syncUpcomingEvents();
  } finally {
    // Keep the "Triggers" UI clean, since debounced sync triggers are handled within the script.
    clearTriggersByHandler_("our206_debouncedSync");
  }
}

function our206_dailyMaintenance() {
  movePastEvents();
  syncUpcomingEvents();
}

// Add "Added On" date when someone directly adds an event to the
// sheet (as opposed to using the form to add), if at least date,
// event/artist, and venue are set and it's not already set.
function maybeSetAddedOn_(e) {
  const sheet = e.range.getSheet();
  if (!isSheet_(sheet, SHEET_KEYS.CONCERTS)) return;

  const row = e.range.getRow();
  if (row <= getHeaderRowForSheet_(sheet)) return;

  const { idx } = getColumnIndexes_(sheet);
  const addedOnCol = findOptionalColumnIndex_(sheet, CFG.headerMatchers.addedOn);
  if (addedOnCol === null) return;

  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const date = rowValues[idx.date];
  const artist = rowValues[idx.artist];
  const venue = rowValues[idx.venue];
  const addedOn = rowValues[addedOnCol];

  if (date && artist && venue && !addedOn) {
    sheet.getRange(row, addedOnCol + 1).setValue(new Date());
    applyConcertRowFormats_(sheet, row);
  }
}

function scheduleDebouncedSync_(e) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("LAST_EDIT_TS", String(Date.now()));
  clearTriggersByHandler_("our206_debouncedSync");
  ScriptApp.newTrigger("our206_debouncedSync")
    .timeBased()
    .after(CFG.debounceMinutes * 60 * 1000)
    .create();
}

// ---------------------- INCOMING RAW PROCESSOR ----------------------

function processIncomingRaw_our206() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incoming = getSheet_(ss, SHEET_KEYS.INCOMING_RAW);
  const concerts = getSheet_(ss, SHEET_KEYS.CONCERTS);

  const incomingLastRow = incoming.getLastRow();
  if (incomingLastRow < 2) {
    toast_("Incoming Raw is empty.");
    return;
  }

  const rawValues = incoming.getRange(2, 1, incomingLastRow - 1, 1).getValues().map(r => String(r[0] || "").trim());
  const concertInfo = getConcertSheetInfo_(concerts);
  const existingRows = concertInfo.rows;

  const exactSet = new Set();
  const fuzzySet = new Set();

  existingRows.forEach(r => {
    const dateKey = sheetDateToKey_(r.date);
    const venueKey = normalizeVenueName_(r.venue);
    const artistNorm = normalizeArtistForCompare_(r.artist);
    const primaryArtist = primaryArtistKey_(r.artist);
    exactSet.add(`${dateKey}|${venueKey}|${artistNorm}`);
    fuzzySet.add(`${dateKey}|${venueKey}|${primaryArtist}`);
  });

  const toAppend = [];
  const statusOut = [];
  const summary = { appended: 0, exactDropped: 0, fuzzyDropped: 0, ignored: 0, changed: [] };

  rawValues.forEach(line => {
    if (!line) {
      statusOut.push(["", ""]);
      return;
    }

    const parsed = parseIncomingLine_(line);
    if (!parsed) {
      statusOut.push(["IGNORED", "Could not parse or no venue (@ missing)"]);
      summary.ignored++;
      return;
    }

    const venue = normalizeVenueName_(parsed.venueRaw);
    const artist = deShoutifyArtist_(parsed.artistRaw);
    const artistNorm = normalizeArtistForCompare_(artist);
    const primaryArtist = primaryArtistKey_(artist);

    const exactKey = `${parsed.dateKey}|${venue}|${artistNorm}`;
    const fuzzyKey = `${parsed.dateKey}|${venue}|${primaryArtist}`;

    if (exactSet.has(exactKey)) {
      statusOut.push(["DROPPED", "Exact duplicate in Concerts"]);
      summary.exactDropped++;
      return;
    }

    if (fuzzySet.has(fuzzyKey)) {
      statusOut.push(["DROPPED", "Likely duplicate in Concerts (same date/venue/primary artist)"]);
      summary.fuzzyDropped++;
      return;
    }

    const score = flamesFromRawScore_(parsed.scoreRaw);

    toAppend.push([parsed.displayDate, artist, venue, score, "", ""]);
    statusOut.push(["APPENDED", ""]);
    summary.appended++;

    exactSet.add(exactKey);
    fuzzySet.add(fuzzyKey);

    if (artist !== parsed.artistRaw) summary.changed.push(`Artist: "${parsed.artistRaw}" → "${artist}"`);
    if (venue !== parsed.venueRaw) summary.changed.push(`Venue: "${parsed.venueRaw}" → "${venue}"`);
    if (parsed.scoreRaw) summary.changed.push(`Score normalized for "${artist}"`);
  });

  incoming.getRange(1, 2, 1, 2).setValues([["Status", "Notes"]]);
  if (statusOut.length) incoming.getRange(2, 2, statusOut.length, 2).setValues(statusOut);

  if (toAppend.length) {
    const targetRow = Math.max(concertInfo.firstDataRow, concerts.getLastRow() + 1);
    concerts.getRange(targetRow, 1, toAppend.length, 6).setValues(toAppend);
    sortConcertsByDateIfPossible_(concerts);
  }

  const logLines = [];
  logLines.push(`Incoming processing complete: appended=${summary.appended}, exactDropped=${summary.exactDropped}, fuzzyDropped=${summary.fuzzyDropped}, ignored=${summary.ignored}`);
  if (summary.changed.length) {
    logLines.push("Changes:");
    summary.changed.slice(0, 50).forEach(x => logLines.push(`- ${x}`));
    if (summary.changed.length > 50) logLines.push(`- …and ${summary.changed.length - 50} more`);
  }
  saveLog_(logLines);
  toast_(`Processed Incoming Raw: appended ${summary.appended}, dropped ${summary.exactDropped + summary.fuzzyDropped}, ignored ${summary.ignored}`);
}

function parseIncomingLine_(line) {
  const dateMatch = line.match(/^\s*(\d{1,2})\/(\d{1,2})/);
  if (!dateMatch || !/@/.test(line)) return null;

  const month = Number(dateMatch[1]);
  const day = Number(dateMatch[2]);
  const year = 2026;
  let rest = line.replace(/^\s*\d{1,2}\/\d{1,2}:?\s*/, "").trim();
  const scoreRaw = extractScoreRaw_(rest);

  rest = rest.replace(/[✅!]+/g, "").trim();
  rest = rest.replace(/\s+-\s+.*$/, "").trim();

  const atIdx = rest.lastIndexOf("@");
  if (atIdx === -1) return null;

  const artistRaw = rest.slice(0, atIdx).trim();
  let venueRaw = rest.slice(atIdx + 1).trim();
  venueRaw = venueRaw.replace(/\s*\([^)]*\)\s*$/, "").trim();

  if (!artistRaw || !venueRaw) return null;

  const date = new Date(year, month - 1, day, 12, 0, 0, 0);
  const tz = spreadsheetTimeZone_();
  return {
    displayDate: Utilities.formatDate(date, tz, "dd-MMM-yyyy"),
    dateKey: Utilities.formatDate(date, tz, "yyyy-MM-dd"),
    artistRaw,
    venueRaw,
    scoreRaw
  };
}

function extractScoreRaw_(line) {
  const count = (line.match(/[✅!]/g) || []).length;
  return count ? "!".repeat(count) : "";
}

function flamesFromRawScore_(scoreRaw) {
  const count = Math.min(4, (String(scoreRaw || "").match(/[✅!]/g) || []).length);
  return count ? "🔥".repeat(count) : "";
}

function deShoutifyArtist_(artist) {
  const s = String(artist || "").trim();
  const letters = s.replace(/[^A-Za-z]/g, "");
  if (!letters) return s;
  if (letters === letters.toUpperCase()) return toTitleCase_(s.toLowerCase());
  return fixKnownArtistNames_(s);
}

function toTitleCase_(s) {
  return s.replace(/\b([a-z])([a-z']*)/g, (_, a, b) => a.toUpperCase() + b);
}

function fixKnownArtistNames_(artist) {
  const map = {
    "Royksopp (DJ Set)": "Röyksopp (DJ Set)",
    "Royksopp": "Röyksopp",
    "Devin The Due": "Devin The Dude"
  };
  return map[artist] || artist;
}

function normalizeArtistForCompare_(artist) {
  return String(artist || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/\bx\b/g, " and ")
    .replace(/[^\w\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function primaryArtistKey_(artist) {
  let s = normalizeArtistForCompare_(artist).replace(/\band friends\b/g, "").trim();
  const split = s.split(/\s+(?:and)\s+/);
  return split[0].trim();
}

// ---------------------- CALENDAR SYNC ----------------------

function dryRunSync() { syncUpcomingEvents_({ dryRun: true }); }
function syncUpcomingEvents() { syncUpcomingEvents_({ dryRun: false }); }

function movePastEvents() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;
  const log = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const concerts = getSheet_(ss, SHEET_KEYS.CONCERTS);
    const past = ensureSheetByKey_(ss, SHEET_KEYS.PAST_CONCERTS);

    ensureUidColumn_(concerts);
    ensureUidColumn_(past);

    compactDataSheet_(concerts);
    sortSheetByDate_(concerts);

    const { idx, headerRow } = getColumnIndexes_(concerts);
    const dataRange = concerts.getDataRange();
    const values = dataRange.getValues();
    const displayValues = dataRange.getDisplayValues();
    const firstDataIdx0 = headerRow;

    const today = new Date();
    today.setHours(0,0,0,0);

    const rowsToMove = [];
    for (let r = firstDataIdx0; r < values.length; r++) {
      const row = values[r];
      const dateCell = displayValues[r][idx.date];
      const artist = String(row[idx.artist] || "").trim();
      const venue = String(row[idx.venue] || "").trim();
      if (!dateCell || !artist || !venue) continue;

      const d = normalizeToAllDayDate_(dateCell);
      if (!d) continue;
      if (d < today) rowsToMove.push({ r0: r, row });
    }

    if (!rowsToMove.length) {
      sortSheetByDate_(past);
      log.push("Move past events: none to move.");
      saveLog_(log);
      toast_("No past events to move.");
      return;
    }

    const lastCol = concerts.getLastColumn();
    ensureColumnCount_(past, lastCol);
    const toAppend = rowsToMove.map(x => x.row.slice(0, lastCol));
    past.getRange(past.getLastRow() + 1, 1, toAppend.length, lastCol).setValues(toAppend);
    rowsToMove.sort((a,b) => b.r0 - a.r0).forEach(x => concerts.deleteRow(x.r0 + 1));

    sortSheetByDate_(past);
    compactDataSheet_(concerts);
    sortSheetByDate_(concerts);

    log.push(`Move past events: moved ${rowsToMove.length} row(s) to "${getSheetSpec_(SHEET_KEYS.PAST_CONCERTS).name}".`);
    saveLog_(log);
    toast_(`Moved ${rowsToMove.length} past row(s).`);
  } finally {
    lock.releaseLock();
  }
}

function movePastEventsAndSync() {
  movePastEvents();
  syncUpcomingEvents();
}

function syncUpcomingEvents_(opts) {
  const dryRun = !!(opts && opts.dryRun);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;

  const log = [];
  try {
    log.push(dryRun ? "DRY RUN (no calendar changes will be made)" : "LIVE RUN");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheet_(ss, SHEET_KEYS.CONCERTS);
    ensureUidColumn_(sheet);
    compactDataSheet_(sheet);
    sortSheetByDate_(sheet);

    const cal = CalendarApp.getCalendarById(CFG.calendarId);
    if (!cal) throw new Error(`Calendar not found for ID "${CFG.calendarId}".`);

    const { idx, uidColIndex, headerRow } = getColumnIndexes_(sheet);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const displayValues = dataRange.getDisplayValues();    const firstDataIdx0 = headerRow;

    const today = new Date();
    today.setHours(0,0,0,0);
    const horizon = new Date(today);
    horizon.setFullYear(horizon.getFullYear() + CFG.horizonYears);

    const desired = new Map();
    const uidWrites = [];

    for (let r = firstDataIdx0; r < values.length; r++) {
      const row = values[r];
      if (row.every(v => String(v || "").trim() === "")) break;

      const dateCell = displayValues[r][idx.date];
      const artist = String(row[idx.artist] || "").trim();
      const venue = String(row[idx.venue] || "").trim();
      if (!dateCell || !artist || !venue) continue;

      const start = normalizeToAllDayDate_(dateCell);
      if (!start) continue;
      if (start < today) continue;

      const uid = buildUid_(start, artist, venue);
      if (uidColIndex !== null) {
        const currentUid = String(row[uidColIndex] || "").trim();
        if (currentUid !== uid) uidWrites.push({ row: r + 1, col: uidColIndex + 1, value: uid });
      }

      const rating = String(row[idx.rating] || "").trim();
      const notes = String(row[idx.notes] || "").trim();
      const ticket = String(row[idx.ticket] || "").trim();

      desired.set(uid, {
        uid,
        title: artist,
        start,
        location: buildLocation_(venue),
        description: buildDescription_(notes, rating, ticket)
      });
    }

    uidWrites.forEach(w => sheet.getRange(w.row, w.col).setValue(w.value));
    if (uidWrites.length) log.push(`UID updates written to sheet: ${uidWrites.length}`);

    const existing = cal.getEvents(today, horizon);
    const existingByUid = new Map();
    existing.forEach(ev => {
      const uid = extractUidFromDescription_(ev.getDescription());
      if (uid) existingByUid.set(uid, ev);
    });

    let created = 0, updated = 0, deleted = 0, dateFixed = 0;

    desired.forEach(d => {
      const ev = existingByUid.get(d.uid);
      const newDesc = attachUidToDescription_(d.uid, d.description);

      if (ev) {
        const curStart = normalizeToAllDayDate_(ev.getAllDayStartDate ? ev.getAllDayStartDate() : ev.getStartTime());
        const want = normalizeToAllDayDate_(d.start);
        const needsDateFix = curStart.getTime() !== want.getTime();
        const changed = ev.getTitle() !== d.title || ev.getLocation() !== d.location || ev.getDescription() !== newDesc || needsDateFix;

        if (changed) {
          updated++;
          log.push(`UPDATED: ${formatDate_(d.start)} — ${d.title} @ ${firstLine_(d.location)}`);
          if (!dryRun) {
            calendarUpdateEvent_(ev, {
              date: needsDateFix ? want : null,
              title: d.title,
              location: d.location,
              description: newDesc
            }, `update ${d.title}`);
            if (needsDateFix) dateFixed++;          
          }
        }
      } else {
        created++;
        log.push(`CREATED: ${formatDate_(d.start)} — ${d.title} @ ${firstLine_(d.location)}`);
        if (!dryRun) {
            calendarCreateAllDayEvent_(cal, d.title, d.start, {
              location: d.location,
              description: newDesc
            }, `create ${d.title}`);
        }
      }
    });

    existingByUid.forEach((ev, uid) => {
      if (!desired.has(uid)) {
        const title = ev.getTitle();
        const when = ev.getAllDayStartDate ? ev.getAllDayStartDate() : ev.getStartTime();
        deleted++;
        log.push(`DELETED: ${formatDate_(when)} — ${title}`);
        if (!dryRun) calendarDeleteEvent_(ev, `delete ${title}`);
      }
    });

    log.unshift(`Sync complete: created=${created}, updated=${updated}, deleted=${deleted}, desired=${desired.size}, existingTagged=${existingByUid.size}, dateFixed=${dateFixed}`);
    saveLog_(log);
    toast_(dryRun ? "Dry run complete — see log." : "Sync complete — see log.");
  } finally {
    lock.releaseLock();
  }
}

// ---------------------- SHEET TIDYING ----------------------

function compactDataSheet_(sheet) {
  const headerRow = getHeaderRowForSheet_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRow) return;

  const dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol);
  const data = dataRange.getValues();
  const kept = data.filter(row => row.some(v => String(v || "").trim() !== ""));
  if (kept.length === data.length) return;

  dataRange.clearContent();
  if (kept.length) sheet.getRange(headerRow + 1, 1, kept.length, lastCol).setValues(kept);
}

function sortSheetByDate_(sheet) {
  const { idx, headerRow } = getColumnIndexes_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRow) return;
  sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).sort({ column: idx.date + 1, ascending: true });
}

// ---------------------- COMMON HELPERS ----------------------

function getSheetSpec_(sheetKey) {
  const spec = CFG.sheets[sheetKey];
  if (!spec) throw new Error(`Unknown sheet key: ${sheetKey}`);
  return spec;
}

function getOptionalSheet_(ss, sheetKey) {
  return ss.getSheetByName(getSheetSpec_(sheetKey).name);
}

function getSheet_(ss, sheetKey) {
  const spec = getSheetSpec_(sheetKey);
  const sheet = ss.getSheetByName(spec.name);
  if (!sheet) throw new Error(`Sheet "${spec.name}" not found.`);
  return sheet;
}

function insertSheetByKey_(ss, sheetKey) {
  return ss.insertSheet(getSheetSpec_(sheetKey).name);
}

function ensureSheetByKey_(ss, sheetKey) {
  return getOptionalSheet_(ss, sheetKey) || insertSheetByKey_(ss, sheetKey);
}

function getSheetKeyByName_(sheetName) {
  return Object.keys(CFG.sheets).find(key => CFG.sheets[key].name === sheetName) || null;
}

function isSheet_(sheet, sheetKey) {
  return sheet.getName() === getSheetSpec_(sheetKey).name;
}

function getHeaderRowForSheet_(sheet) {
  const sheetKey = getSheetKeyByName_(sheet.getName());
  if (sheetKey && CFG.sheets[sheetKey].headerRow) return CFG.sheets[sheetKey].headerRow;
  return detectHeaderRow_(sheet);
}

function applyConcertRowFormats_(sheet, row) {
  sheet.getRange(row, 1).setNumberFormat("dd-MMM-yyyy");

  const addedOnCol = findOptionalColumnIndex_(sheet, CFG.headerMatchers.addedOn);
  if (addedOnCol !== null) {
    sheet.getRange(row, addedOnCol + 1).setNumberFormat("dd-MMM-yyyy HH:mm:ss");
  }
}

function coerceTimestamp_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) return value;
  if (value == null || value === "") return null;
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function getHeaders_(sheet) {
  const headerRow = getHeaderRowForSheet_(sheet);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  return { headers, headerRow };
}

function findOptionalColumnIndex_(sheet, matcher) {
  const { headers } = getHeaders_(sheet);
  const needle = String(matcher || "").toLowerCase();
  const i = headers.map(h => h.toLowerCase()).findIndex(h => h.includes(needle));
  return i === -1 ? null : i;
}

function spreadsheetTimeZone_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
}

function normalizeToAllDayDate_(v) {
  const d = coerceDate_(v);
  if (!d) return null;
  const tz = spreadsheetTimeZone_();
  const y = Number(Utilities.formatDate(d, tz, "yyyy"));
  const m = Number(Utilities.formatDate(d, tz, "M"));
  const day = Number(Utilities.formatDate(d, tz, "d"));
  // Anchor at local noon-equivalent to avoid previous-day shifts when Calendar interprets the Date.
  return new Date(y, m - 1, day, 12, 0, 0, 0);
}


function buildDescription_(notes, rating, ticket) {
  const parts = [];
  if (notes) parts.push(notes);
  if (rating) parts.push(`Skoi rating: ${rating}`);
  if (ticket) parts.push(`Ticket link: ${ticket}`);
  return parts.join("\n");
}

function buildLocation_(venue) {
  const normalized = normalizeVenueName_(venue);
  const addr = getVenueAddress_(normalized);
  return addr ? `${normalized}\n${addr}` : normalized;
}

function getVenueAddress_(normalizedVenue) {
  const custom = getVenueMapData_();
  const lower = String(normalizedVenue || "").trim().toLowerCase();
  if (custom.byNormalized[lower] && custom.byNormalized[lower].address) return custom.byNormalized[lower].address;
  return VENUE_ADDRESS[normalizedVenue] || "";
}

function normalizeVenueName_(raw) {
  const s = String(raw || "").trim();
  const lower = s.toLowerCase();
  const custom = getVenueMapData_();
  if (custom.byRaw[lower]) return custom.byRaw[lower].normalized;
  return FALLBACK_VENUE_NORMALIZATION[lower] || s;
}

function getVenueMapData_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("our206_venue_map_v1");
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOptionalSheet_(ss, SHEET_KEYS.VENUE_MAP);
  const result = { byRaw: {}, byNormalized: {} };

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(3, sheet.getLastColumn())).getValues();
    rows.forEach(r => {
      const raw = String(r[0] || "").trim().toLowerCase();
      const normalized = String(r[1] || "").trim();
      const address = String(r[2] || "").trim();
      if (!raw || !normalized) return;
      result.byRaw[raw] = { normalized, address };
      const nkey = normalized.toLowerCase();
      if (!result.byNormalized[nkey]) result.byNormalized[nkey] = { address };
      if (address) result.byNormalized[nkey].address = address;
    });
  }

  cache.put("our206_venue_map_v1", JSON.stringify(result), 300);
  return result;
}

function clearVenueMapCache_our206() {
  CacheService.getScriptCache().remove("our206_venue_map_v1");
  toast_("Venue map cache cleared.");
}

function detectHeaderRow_(sheet) {
  const sheetKey = getSheetKeyByName_(sheet.getName());
  if (sheetKey && CFG.sheets[sheetKey].headerRow) return CFG.sheets[sheetKey].headerRow;

  const maxScan = Math.min(15, sheet.getLastRow() || 15);
  const lastCol = sheet.getLastColumn() || 20;
  const scan = sheet.getRange(1, 1, maxScan, lastCol).getValues();
  const need = [CFG.headerMatchers.date, CFG.headerMatchers.artist, CFG.headerMatchers.venue];

  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map(v => String(v || "").trim().toLowerCase());
    const hits = need.every(k => row.some(cell => cell.includes(k)));
    if (hits) return r + 1;
  }
  return CFG.headerRowFallback;
}

function ensureUidColumn_(sheet) {
  const headerRow = getHeaderRowForSheet_(sheet);
  const lastCol = sheet.getLastColumn() || 1;
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const lc = headers.map(h => h.toLowerCase());
  if (lc.findIndex(h => h === CFG.uidHeader.toLowerCase()) !== -1) return;

  let insertAfterCol = lastCol;
  if (CFG.keepLastColumnHeader && CFG.keepLastColumnHeader.trim()) {
    const keepIdx = lc.findIndex(h => h === CFG.keepLastColumnHeader.trim().toLowerCase());
    if (keepIdx !== -1) insertAfterCol = Math.max(1, keepIdx);
  }

  if (insertAfterCol >= lastCol) {
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(headerRow, lastCol + 1).setValue(CFG.uidHeader);
  } else {
    sheet.insertColumnAfter(insertAfterCol);
    sheet.getRange(headerRow, insertAfterCol + 1).setValue(CFG.uidHeader);
  }
}

function getColumnIndexes_(sheet) {
  const { headers, headerRow } = getHeaders_(sheet);
  const lc = headers.map(h => h.toLowerCase());

  function find(sub) {
    const i = lc.findIndex(h => h.includes(sub.toLowerCase()));
    if (i === -1) throw new Error(`Missing column whose header includes "${sub}". Headers: ${headers.join(", ")}`);
    return i;
  }

  return {
    idx: {
      date: find(CFG.headerMatchers.date),
      artist: find(CFG.headerMatchers.artist),
      venue: find(CFG.headerMatchers.venue),
      rating: find(CFG.headerMatchers.rating),
      notes: find(CFG.headerMatchers.notes),
      ticket: find(CFG.headerMatchers.ticket)
    },
    uidColIndex: lc.findIndex(h => h === CFG.uidHeader.toLowerCase()),
    headerRow
  };
}

function getConcertSheetInfo_(sheet) {
  const headerRow = getHeaderRowForSheet_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 6);

  if (lastRow <= headerRow) return { rows: [], headerRow, firstDataRow: headerRow + 1 };

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim().toLowerCase());
  const idxDate = headers.findIndex(h => h.includes("date"));
  const idxArtist = headers.findIndex(h => h.includes("artist"));
  const idxVenue = headers.findIndex(h => h.includes("venue"));

  const range = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  const rows = values
    .map((r, i) => ({
      raw: r,
      display: displayValues[i]
    }))
    .filter(x => x.raw.some(v => String(v || "").trim() !== ""))
    .map(x => ({
      date: x.display[idxDate],
      artist: x.raw[idxArtist],
      venue: x.raw[idxVenue]
    }));

  return { rows, headerRow, firstDataRow: Math.max(headerRow + 1, lastRow + 1) };
}

function ensureColumnCount_(sheet, neededCols) {
  const have = sheet.getLastColumn();
  if (have < neededCols) sheet.insertColumnsAfter(have, neededCols - have);
}

function buildUid_(date, artist, venue) {
  const seed = `${sheetDateToKey_(date)}|${normalizeArtistForCompare_(artist)}|${normalizeVenueName_(venue).toLowerCase()}`;
  return sha1_(seed).slice(0, 24);
}

function attachUidToDescription_(uid, userDescription) {
  const desc = String(userDescription || "").trim();
  return desc ? `${desc}\n\n${CFG.uidMarkerPrefix}${uid}` : `${CFG.uidMarkerPrefix}${uid}`;
}

function extractUidFromDescription_(description) {
  const d = String(description || "");
  const re = new RegExp(`\\${CFG.uidMarkerPrefix}(\\w{16,64})`);
  const m = d.match(re);
  return m ? m[1] : null;
}

function sha1_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, s, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function localDateFromKey_(key) {
  const m = String(key).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;

  // Noon local avoids DST/UTC rollover weirdness for all-day events
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
}

function coerceDate_(v) {
  const key = sheetDateToKey_(v);
  if (!key) return null;
  return localDateFromKey_(key);
}

function sheetDateToKey_(value) {
  if (value == null || value === "") return "";

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  const s = String(value).trim();
  if (!s) return "";

  // DD-MMM-YYYY
  let m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (m) {
    const tmp = new Date(`${m[2]} ${m[1]}, ${m[3]} 12:00:00`);
    if (!isNaN(tmp)) {
      return Utilities.formatDate(tmp, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  }

  // M/D/YYYY or MM/DD/YYYY
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const month = Number(m[1]);
    const day = Number(m[2]);
    const year = Number(m[3]);
    const tmp = new Date(year, month - 1, day, 12, 0, 0, 0);
    return Utilities.formatDate(tmp, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  // YYYY-MM-DD
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return s;
  }

  const tmp = new Date(s);
  if (isNaN(tmp)) return "";
  return Utilities.formatDate(tmp, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDate_(d) {
  return sheetDateToKey_(d) || "";
}

function firstLine_(s) {
  return String(s || "").split(/\r?\n/)[0];
}

function addDays_(d, days) {
  const x = new Date(d);
  x.setDate(x.getDate() + days);
  return x;
}

function sortConcertsByDateIfPossible_(sheet) {
  try { sortSheetByDate_(sheet); } catch (err) { Logger.log(`Sort skipped: ${err}`); }
}

function saveLog_(lines) {
  const msg = lines.join("\n");
  Logger.log(msg);
  PropertiesService.getScriptProperties().setProperty("LAST_RUN_LOG", msg);
}

function toast_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, "Our206", 8); } catch (e) { Logger.log(msg); }
}

// ---------------------- TRIGGER INSTALL ----------------------

function installOnEditTriggerIfMissing_() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "our206_onEdit");
  if (!exists) {
    ScriptApp.newTrigger("our206_onEdit")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
  }
}

function installDailyTriggerIfMissing_(handlerFunction, hourLocal) {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === handlerFunction);
  if (!exists) {
    ScriptApp.newTrigger(handlerFunction)
      .timeBased()
      .everyDays(1)
      .atHour(hourLocal)
      .create();
  }
}

function clearTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
  });
}

// See attached note in chat: this file is a patch block to add rate-limit-aware calendar writes.
// Paste these helpers into your script and replace the create/update/delete calls as noted.

function calendarWrite_(fn, label) {
  const MAX_ATTEMPTS = 6;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const result = fn();
      Utilities.sleep(250);
      return result;
    } catch (err) {
      const msg = String(err && err.message ? err.message : err);
      if (msg.includes("too many calendars or calendar events")) {
        const delay = Math.min(30000, 1000 * Math.pow(2, attempt - 1));
        Logger.log(`Rate limited during ${label || "calendar write"}; retrying in ${delay}ms (attempt ${attempt}/${MAX_ATTEMPTS})`);
        Utilities.sleep(delay);
        continue;
      }
      throw err;
    }
  }
  throw new Error(`Calendar write failed after retries: ${label || "unknown operation"}`);
}

function calendarDeleteEvent_(ev, label) {
  return calendarWrite_(function() {
    ev.deleteEvent();
  }, label || `delete ${ev.getTitle()}`);
}

function calendarCreateAllDayEvent_(cal, title, start, options, label) {
  return calendarWrite_(function() {
    return cal.createAllDayEvent(title, start, options);
  }, label || `create ${title}`);
}

function calendarUpdateEvent_(ev, data, label) {
  return calendarWrite_(function() {
    if (data.date) ev.setAllDayDate(data.date);
    if (data.title != null) ev.setTitle(data.title);
    if (data.location != null) ev.setLocation(data.location);
    if (data.description != null) ev.setDescription(data.description);
  }, label || `update ${ev.getTitle()}`);
}

function debugHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, SHEET_KEYS.CONCERTS);

  const lastCol = sheet.getLastColumn();
  const headerRow = 3;
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

  console.log("Sheet:", sheet.getName());
  console.log("Last col:", lastCol);
  headers.forEach((h, i) => console.log(`${i + 1}: [${h}]`));
}


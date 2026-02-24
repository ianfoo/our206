/**
 * Our206 — Sheet ↔ Google Calendar sync (debounced, forward-only, cancellable)
 *
 * Notes:
 * - Uses all-day events only.
 * - UID marker is stable: [our206_uid]:<hash>
 * - Date parsing is timezone-safe for YYYY-MM-DD and sheet Date values.
 * - Calendar ID can be set in Script Properties as OUR206_CALENDAR_ID.
 */

const CFG = {
  // Prefer Script Property OUR206_CALENDAR_ID; this fallback can be left blank in public repos.
  calendarId: "",
  concertsSheetName: "Concerts",
  pastSheetName: "Past Concerts",
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
    uid: "uid"
  },
  keepLastColumnHeader: "cap"
};

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
  "Tractor Tavern": "5213 Ballard Ave NW, Seattle, WA 98107"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Our206")
    .addItem("Set up (install triggers)", "setUpOur206")
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
  const props = PropertiesService.getScriptProperties();
  const msg = props.getProperty("LAST_RUN_LOG") || "(no log yet)";
  SpreadsheetApp.getUi().alert(msg);
}

function setUpOur206() { return setupOur206(); }

function setupOur206() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const concerts = mustGetSheet_(ss, CFG.concertsSheetName);
  const past = ensureSheet_(ss, CFG.pastSheetName);

  ensureUidColumn_(concerts);
  ensureUidColumn_(past);

  installOnEditTriggerIfMissing_();
  installDailyTriggerIfMissing_("our206_dailyMaintenance", 3);

  toast_("Our206 setup complete. Hide the UID column in both sheets if you want.");
}

function our206_onEdit(e) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("LAST_EDIT_TS", String(Date.now()));

  clearTriggersByHandler_("our206_debouncedSync");

  ScriptApp.newTrigger("our206_debouncedSync")
    .timeBased()
    .after(CFG.debounceMinutes * 60 * 1000)
    .create();
}

function our206_debouncedSync() {
  const props = PropertiesService.getScriptProperties();
  const lastEdit = Number(props.getProperty("LAST_EDIT_TS") || "0");
  const now = Date.now();
  if (now - lastEdit < CFG.debounceGuardMinutes * 60 * 1000) return;

  syncUpcomingEvents();
}

function our206_dailyMaintenance() {
  movePastEvents();
  syncUpcomingEvents();
}

function movePastEvents() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;

  const log = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const concerts = mustGetSheet_(ss, CFG.concertsSheetName);
    const past = ensureSheet_(ss, CFG.pastSheetName);

    ensureUidColumn_(concerts);
    ensureUidColumn_(past);

    compactDataSheet_(concerts);
    sortSheetByDate_(concerts);

    const { idx, headerRow } = getColumnIndexes_(concerts);

    const dataRange = concerts.getDataRange();
    const values = dataRange.getValues();
    const displayValues = dataRange.getDisplayValues();
    const firstDataIdx0 = headerRow;

    const today = startOfDay_(new Date());
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    const rowsToMove = [];
    for (let r = firstDataIdx0; r < values.length; r++) {
      const row = values[r];
      const dateCell = row[idx.date];
      const dateDisplay = displayValues[r][idx.date];
      const artist = String(row[idx.artist] || "").trim();
      const venue = String(row[idx.venue] || "").trim();
      if (!dateCell || !artist || !venue) continue;

      const dayKey = coerceSheetDayKey_(dateCell, dateDisplay, tz);
      const d = dayKey ? parseYmdLocal_(dayKey) : null;
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
    const toAppend = rowsToMove.map(x => x.row.slice(0, lastCol));

    ensureColumnCount_(past, lastCol);
    past.getRange(past.getLastRow() + 1, 1, toAppend.length, lastCol).setValues(toAppend);

    rowsToMove.sort((a, b) => b.r0 - a.r0).forEach(x => concerts.deleteRow(x.r0 + 1));

    ensureUidColumn_(past);
    sortSheetByDate_(past);
    compactDataSheet_(concerts);
    sortSheetByDate_(concerts);

    log.push(`Move past events: moved ${rowsToMove.length} row(s) to "${CFG.pastSheetName}".`);
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

function dryRunSync() { syncUpcomingEvents_({ dryRun: true }); }
function syncUpcomingEvents() { syncUpcomingEvents_({ dryRun: false }); }

function syncUpcomingEvents_(opts) {
  const dryRun = !!(opts && opts.dryRun);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;

  const log = [];
  try {
    log.push(dryRun ? "DRY RUN (no calendar changes will be made)" : "LIVE RUN");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = mustGetSheet_(ss, CFG.concertsSheetName);
    const tz = ss.getSpreadsheetTimeZone();

    ensureUidColumn_(sheet);
    compactDataSheet_(sheet);
    sortSheetByDate_(sheet);

    const cal = CalendarApp.getCalendarById(getCalendarId_());
    if (!cal) throw new Error(`Calendar not found for ID \"${getCalendarId_()}\".`);

    const { idx, uidColIndex, headerRow } = getColumnIndexes_(sheet);

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const displayValues = dataRange.getDisplayValues();
    const firstDataIdx0 = headerRow;

    const today = startOfDay_(new Date());
    const horizon = new Date(today);
    horizon.setFullYear(horizon.getFullYear() + CFG.horizonYears);

    const desired = new Map();
    const uidWrites = [];

    for (let r = firstDataIdx0; r < values.length; r++) {
      const row = values[r];
      if (row.every(v => String(v || "").trim() === "")) break;

      const dateCell = row[idx.date];
      const dateDisplay = displayValues[r][idx.date];
      const artist = String(row[idx.artist] || "").trim();
      const venue = String(row[idx.venue] || "").trim();
      if (!dateCell || !artist || !venue) continue;

      const dayKey = coerceSheetDayKey_(dateCell, dateDisplay, tz);
      const start = dayKey ? parseYmdLocal_(dayKey) : null;
      if (!start || start < today) continue;

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
        dayKey,
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

    let created = 0;
    let updated = 0;
    let deleted = 0;

    desired.forEach(d => {
      const ev = existingByUid.get(d.uid);
      const newDesc = attachUidToDescription_(d.uid, d.description);

      if (ev) {
        const changed =
          ev.getTitle() !== d.title ||
          ev.getLocation() !== d.location ||
          ev.getDescription() !== newDesc;

        if (changed) {
          updated++;
          log.push(`UPDATED: ${d.dayKey} — ${d.title} @ ${firstLine_(d.location)}`);
          if (!dryRun) {
            ev.setTitle(d.title);
            ev.setLocation(d.location);
            ev.setDescription(newDesc);
          }
        }
      } else {
        created++;
        log.push(`CREATED: ${d.dayKey} — ${d.title} @ ${firstLine_(d.location)}`);
        if (!dryRun) {
          // Single-day all-day event (no explicit end date) to avoid day-shift bugs.
          cal.createAllDayEvent(d.title, d.start, { location: d.location, description: newDesc });
        }
      }
    });

    existingByUid.forEach((ev, uid) => {
      if (!desired.has(uid)) {
        const title = ev.getTitle();
        const when = ev.isAllDayEvent && ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime();
        deleted++;
        log.push(`DELETED: ${formatDate_(when)} — ${title}`);
        if (!dryRun) ev.deleteEvent();
      }
    });

    log.unshift(`Sync complete: created=${created}, updated=${updated}, deleted=${deleted}, desired=${desired.size}, existingTagged=${existingByUid.size}`);
    saveLog_(log);

    toast_(dryRun ? "Dry run complete — see log." : "Sync complete — see log.");
  } finally {
    lock.releaseLock();
  }
}

function compactDataSheet_(sheet) {
  const headerRow = detectHeaderRow_(sheet);
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

  const dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol);
  dataRange.sort({ column: idx.date + 1, ascending: true });
}

function buildDescription_(notes, rating, ticket) {
  const parts = [];
  if (notes) parts.push(notes);
  if (rating) parts.push(`Skoi rating: ${rating}`);
  if (ticket) parts.push(`Ticket link: ${ticket}`);
  return parts.join("\n");
}

function buildLocation_(venue) {
  const v = String(venue || "").trim();
  const addr = VENUE_ADDRESS[v];
  return addr ? `${v}\n${addr}` : v;
}

function normalize_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\w\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function buildUid_(date, artist, venue) {
  const seed = `${formatDate_(date)}|${normalize_(artist)}|${normalize_(venue)}`;
  return sha1_(seed).slice(0, 24);
}

function attachUidToDescription_(uid, userDescription, uidMarkerPrefix) {
  const marker = uidMarkerPrefix || CFG.uidMarkerPrefix;
  const desc = String(userDescription || "").trim();
  return desc ? `${desc}\n\n${marker}${uid}` : `${marker}${uid}`;
}

function extractUidFromDescription_(description, uidMarkerPrefix) {
  const marker = uidMarkerPrefix || CFG.uidMarkerPrefix;
  const d = String(description || "");
  const re = new RegExp(`${escapeRegex_(marker)}(\\w{16,64})`);
  const m = d.match(re);
  return m ? m[1] : null;
}

function sha1_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, s, Utilities.Charset.UTF_8);
  return raw.map(b => ("0" + (b & 0xFF).toString(16)).slice(-2)).join("");
}

function coerceDate_(v, tz) {
  const timezone = tz || Session.getScriptTimeZone();

  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
    // Google Sheets date cells often arrive as UTC-midnight Date objects.
    // Using UTC components preserves the intended spreadsheet day.
    const y = v.getUTCFullYear();
    const m = v.getUTCMonth() + 1;
    const d = v.getUTCDate();
    const ymd = `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    return parseYmdLocal_(ymd);
  }

  const s = String(v || "").trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return parseYmdLocal_(s);

  const m1 = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (m1) {
    const month = {
      jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
      jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
    }[m1[2].toLowerCase()];
    if (!month) return null;
    return new Date(Number(m1[3]), month - 1, Number(m1[1]), 12, 0, 0, 0);
  }

  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m2) return new Date(Number(m2[3]), Number(m2[1]) - 1, Number(m2[2]), 12, 0, 0, 0);

  const tmp = new Date(s);
  if (isNaN(tmp)) return null;
  return new Date(tmp.getFullYear(), tmp.getMonth(), tmp.getDate(), 12, 0, 0, 0);
}

function coerceSheetDate_(rawValue, displayValue, tz) {
  const shown = String(displayValue || "").trim();
  if (shown) {
    const fromShown = coerceDate_(shown, tz);
    if (fromShown) return fromShown;
  }
  return coerceDate_(rawValue, tz);
}

function coerceSheetDayKey_(rawValue, displayValue, tz) {
  const timezone = tz || Session.getScriptTimeZone();
  const shown = String(displayValue || "").trim();
  if (shown) {
    const fromShown = coerceDate_(shown, timezone);
    if (fromShown) return formatDate_(fromShown);
  }

  const fromRaw = coerceDate_(rawValue, timezone);
  if (!fromRaw) return null;
  return Utilities.formatDate(fromRaw, timezone, "yyyy-MM-dd");
}

function parseYmdLocal_(ymd) {
  const m = ymd.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  // Use noon instead of midnight to avoid timezone boundary shifts when
  // all-day dates are interpreted across script/calendar/spreadsheet zones.
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function formatDate_(d) {
  const x = new Date(d);
  const yyyy = x.getFullYear();
  const mm = String(x.getMonth() + 1).padStart(2, "0");
  const dd = String(x.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function firstLine_(s) {
  return String(s || "").split(/\r?\n/)[0];
}

function escapeRegex_(s) {
  return String(s || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function toast_(msg) {
  try {
    SpreadsheetApp.getActive().toast(msg, "Our206", 6);
  } catch (e) {
    // ignore in non-UI contexts
  }
}

function saveLog_(lines) {
  const msg = lines.join("\n");
  Logger.log(msg);
  PropertiesService.getScriptProperties().setProperty("LAST_RUN_LOG", msg);
}

function detectHeaderRow_(sheet) {
  const maxScan = Math.min(15, sheet.getLastRow() || 15);
  const lastCol = Math.max(sheet.getLastColumn() || 1, 10);
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
  const headerRow = detectHeaderRow_(sheet);
  const lastCol = sheet.getLastColumn() || 1;
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || "").trim());
  const lc = headers.map(h => h.toLowerCase());

  const existingIndex = lc.findIndex(h => h === CFG.uidHeader.toLowerCase());
  if (existingIndex !== -1) return;

  let insertAfterCol = lastCol;
  if (CFG.keepLastColumnHeader && CFG.keepLastColumnHeader.trim()) {
    const keep = CFG.keepLastColumnHeader.trim().toLowerCase();
    const keepIdx = lc.findIndex(h => h === keep);
    if (keepIdx !== -1) {
      insertAfterCol = keepIdx;
      if (insertAfterCol < 1) insertAfterCol = 1;
    }
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
  const headerRow = detectHeaderRow_(sheet);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || "").trim());
  const lc = headers.map(h => h.toLowerCase());

  function find(sub) {
    const s = sub.toLowerCase();
    const i = lc.findIndex(h => h.includes(s));
    if (i === -1) throw new Error(`Missing column whose header includes \"${sub}\". Headers: ${headers.join(", ")}`);
    return i;
  }

  const idx = {
    date: find(CFG.headerMatchers.date),
    artist: find(CFG.headerMatchers.artist),
    venue: find(CFG.headerMatchers.venue),
    rating: find(CFG.headerMatchers.rating),
    notes: find(CFG.headerMatchers.notes),
    ticket: find(CFG.headerMatchers.ticket)
  };

  const uidColIndex = lc.findIndex(h => h === CFG.uidHeader.toLowerCase());

  return { idx, uidColIndex: uidColIndex === -1 ? null : uidColIndex, headerRow };
}

function ensureColumnCount_(sheet, neededCols) {
  const have = sheet.getLastColumn();
  if (have >= neededCols) return;
  sheet.insertColumnsAfter(have, neededCols - have);
}

function mustGetSheet_(ss, name) {
  const s = ss.getSheetByName(name);
  if (!s) throw new Error(`Sheet \"${name}\" not found.`);
  return s;
}

function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function installOnEditTriggerIfMissing_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === "our206_onEdit");
  if (exists) return;

  ScriptApp.newTrigger("our206_onEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

function installDailyTriggerIfMissing_(handlerFunction, hourLocal) {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === handlerFunction);
  if (exists) return;

  ScriptApp.newTrigger(handlerFunction)
    .timeBased()
    .everyDays(1)
    .atHour(hourLocal)
    .create();
}

function clearTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
  });
}

function getCalendarId_() {
  const fromProps = PropertiesService.getScriptProperties().getProperty("OUR206_CALENDAR_ID");
  const id = String(fromProps || CFG.calendarId || "").trim();
  if (!id) {
    throw new Error("Missing calendar ID. Set Script Property OUR206_CALENDAR_ID.");
  }
  return id;
}

function purgeAllFutureEvents_our206_paced() {
  const SLEEP_BETWEEN_DELETES_MS = 750;
  const MAX_DELETES_PER_RUN = 80;
  const YEARS_AHEAD = 10;
  const STATE_KEY = "OUR206_PURGE_CURSOR";

  const cal = CalendarApp.getCalendarById(getCalendarId_());
  if (!cal) throw new Error(`Calendar not found: ${getCalendarId_()}`);

  const props = PropertiesService.getScriptProperties();
  let cursor = Number(props.getProperty(STATE_KEY) || "0");

  const today = startOfDay_(new Date());
  const end = new Date(today);
  end.setFullYear(end.getFullYear() + YEARS_AHEAD);

  const events = cal.getEvents(today, end);

  if (events.length === 0) {
    props.deleteProperty(STATE_KEY);
    SpreadsheetApp.getActive().toast("No future events found to purge.", "Our206", 8);
    return;
  }

  let deleted = 0;

  while (cursor < events.length && deleted < MAX_DELETES_PER_RUN) {
    const ev = events[cursor];

    let attempt = 0;
    while (true) {
      try {
        ev.deleteEvent();
        break;
      } catch (err) {
        const msg = String(err && err.message ? err.message : err);

        if (msg.includes("too many calendars or calendar events")) {
          attempt++;
          const backoff = Math.min(30000, 1000 * Math.pow(2, attempt));
          Logger.log(`Throttle hit. Backing off ${backoff}ms. Attempt ${attempt}.`);
          Utilities.sleep(backoff);
          continue;
        }

        throw err;
      }
    }

    deleted++;
    cursor++;
    Utilities.sleep(SLEEP_BETWEEN_DELETES_MS);
  }

  if (cursor >= events.length) {
    props.deleteProperty(STATE_KEY);
    SpreadsheetApp.getActive().toast(`Purge complete. Deleted ${events.length} event(s).`, "Our206", 10);
  } else {
    props.setProperty(STATE_KEY, String(cursor));
    SpreadsheetApp.getActive().toast(
      `Deleted ${deleted} this run (${cursor}/${events.length}). Run again to continue.`,
      "Our206",
      10
    );
  }
}

function importPastConcerts_our206() {
  const sheetName = CFG.pastSheetName;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const cal = CalendarApp.getCalendarById(getCalendarId_());
  if (!cal) throw new Error(`Calendar not found: ${getCalendarId_()}`);

  const headerRow = detectHeaderRow_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= headerRow) {
    SpreadsheetApp.getActive().toast("No past rows to import (sheet empty).", "Our206", 8);
    return;
  }

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const lc = headers.map(h => h.toLowerCase());

  const idx = {
    date: findHeaderIndex_(lc, CFG.headerMatchers.date),
    artist: findHeaderIndex_(lc, CFG.headerMatchers.artist),
    venue: findHeaderIndex_(lc, CFG.headerMatchers.venue),
    rating: optionalHeaderIndex_(lc, CFG.headerMatchers.rating),
    notes: optionalHeaderIndex_(lc, CFG.headerMatchers.notes),
    ticket: optionalHeaderIndex_(lc, CFG.headerMatchers.ticket)
  };

  const dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol);
  const values = dataRange.getValues();
  const displayValues = dataRange.getDisplayValues();

  const today = startOfDay_(new Date());
  const tz = ss.getSpreadsheetTimeZone();

  const startWindow = new Date(today);
  startWindow.setFullYear(startWindow.getFullYear() - 10);
  const endWindow = new Date(today);
  endWindow.setFullYear(endWindow.getFullYear() + 10);

  const existing = cal.getEvents(startWindow, endWindow);
  const existingUids = new Set();
  existing.forEach(ev => {
    const uid = extractUidFromDescription_(ev.getDescription());
    if (uid) existingUids.add(uid);
  });

  let created = 0;
  let skipped = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const dateCell = row[idx.date];
    const dateDisplay = displayValues[i][idx.date];
    const artist = String(row[idx.artist] || "").trim();
    const venue = String(row[idx.venue] || "").trim();
    if (!dateCell || !artist || !venue) continue;

    const dayKey = coerceSheetDayKey_(dateCell, dateDisplay, tz);
    const d = dayKey ? parseYmdLocal_(dayKey) : null;
    if (!d) continue;
    if (d >= today) continue;

    const uid = buildUid_(d, artist, venue);
    if (existingUids.has(uid)) {
      skipped++;
      continue;
    }

    const rating = safeCell_(row, idx.rating);
    const notes = safeCell_(row, idx.notes);
    const ticket = safeCell_(row, idx.ticket);

    const desc = buildDescription_(notes, rating, ticket);
    const finalDesc = attachUidToDescription_(uid, desc);

    cal.createAllDayEvent(artist, d, {
      location: buildLocation_(venue),
      description: finalDesc
    });

    created++;
    existingUids.add(uid);
  }

  SpreadsheetApp.getActive().toast(`Imported ${created} past event(s); skipped ${skipped} already-present`, "Our206", 10);
  Logger.log(`Past import complete: created=${created}, skipped=${skipped}`);
}

function findHeaderIndex_(lcHeaders, needle) {
  const n = needle.toLowerCase();
  const i = lcHeaders.findIndex(h => h.includes(n));
  if (i === -1) throw new Error(`Missing header containing \"${needle}\"`);
  return i;
}

function optionalHeaderIndex_(lcHeaders, needle) {
  const n = needle.toLowerCase();
  return lcHeaders.findIndex(h => h.includes(n));
}

function safeCell_(row, idx) {
  if (idx === -1 || idx == null) return "";
  return String(row[idx] || "").trim();
}

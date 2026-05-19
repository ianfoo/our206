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


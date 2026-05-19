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


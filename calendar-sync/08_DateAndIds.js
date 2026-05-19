function coerceTimestamp_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) return value;
  if (value == null || value === "") return null;
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
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


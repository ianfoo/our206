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


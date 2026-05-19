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


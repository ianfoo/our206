function setUpOur206() { return setupOur206(); }

function setupOur206() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureUidColumn_(getSheet_(ss, SHEET_KEYS.EVENTS));
  ensureUidColumn_(ensureSheetByKey_(ss, SHEET_KEYS.PAST_EVENTS));
  installOnEditTriggerIfMissing_();
  installDailyTriggerIfMissing_("our206_dailyMaintenance", 3);
  installFormSubmitTriggerIfMissing_();
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

function installFormSubmitTriggerIfMissing_() {
  const exists = ScriptApp.getProjectTriggers().some(
    t => t.getHandlerFunction() === "onFormSubmit"
  );

  if (!exists) {
    ScriptApp.newTrigger("onFormSubmit")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onFormSubmit()
      .create();
  }
}

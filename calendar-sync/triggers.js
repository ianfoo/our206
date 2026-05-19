function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = getSheet_(ss, SHEET_KEYS.EVENTS);

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
  if (!isSheet_(sheet, SHEET_KEYS.EVENTS)) return;

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


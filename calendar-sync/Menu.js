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

function saveLog_(lines) {
  const msg = lines.join("\n");
  Logger.log(msg);
  PropertiesService.getScriptProperties().setProperty("LAST_RUN_LOG", msg);
}

function toast_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, "Our206", 8); } catch (e) { Logger.log(msg); }
}

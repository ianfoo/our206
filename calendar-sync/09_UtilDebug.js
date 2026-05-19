function saveLog_(lines) {
  const msg = lines.join("\n");
  Logger.log(msg);
  PropertiesService.getScriptProperties().setProperty("LAST_RUN_LOG", msg);
}

function toast_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, "Our206", 8); } catch (e) { Logger.log(msg); }
}

// ---------------------- TRIGGER INSTALL ----------------------

function debugHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, SHEET_KEYS.CONCERTS);

  const lastCol = sheet.getLastColumn();
  const headerRow = 3;
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

  console.log("Sheet:", sheet.getName());
  console.log("Last col:", lastCol);
  headers.forEach((h, i) => console.log(`${i + 1}: [${h}]`));
}


/**
 * Silver & Strong: waitlist receiver.
 * Paste into Extensions > Apps Script of your Google Sheet, then deploy as a web app.
 * Each signup becomes one row. The header row is created automatically.
 */
var COLUMNS = [
  "submitted_at", "signing_up_for", "name", "parent_name", "phone", "age", "city", "area", "concern",
  "join_mode", "tech_comfort", "interests", "price_online", "price_band",
  "notes", "page"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rec = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS);
      sheet.setFrozenRows(1);
    }
    var row = COLUMNS.map(function (c) {
      var v = rec[c] == null ? "" : String(rec[c]);
      // Stop a cell starting with = + - @ from being read as a formula
      if (/^[=+\-@]/.test(v)) v = "'" + v;
      return v;
    });
    sheet.appendRow(row);
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error");
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("Silver & Strong waitlist is running.");
}

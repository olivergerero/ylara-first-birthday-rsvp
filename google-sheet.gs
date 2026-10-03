function doPost(e) {
  const sheet = getRsvpSheet();
  const data = JSON.parse(e.postData.contents);
  const names = (data.names || []).map(function (name) {
    return String(name).trim();
  }).filter(Boolean);
  const submittedAt = new Date();

  names.forEach(function (name) {
    sheet.appendRow([submittedAt, name]);
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, count: names.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRsvpSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("RSVP");
  if (!sheet) sheet = spreadsheet.insertSheet("RSVP");
  if (sheet.getLastColumn() > 2) {
    sheet.deleteColumns(3, sheet.getLastColumn() - 2);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Guest name"]);
  } else {
    sheet.getRange(1, 1, 1, 2).setValues([["Timestamp", "Guest name"]]);
  }
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  sheet.setFrozenRows(1);
  return sheet;
}

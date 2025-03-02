function changeFont() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  range.setFontFamily("Google Sans");
}

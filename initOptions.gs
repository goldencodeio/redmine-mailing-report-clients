var OPTIONS = {};

function initOptions() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  getOptionsData();

  var sheetName = 'Тексты писем';
  var existingSheet = _ss.getSheetByName(sheetName);
  if (existingSheet) _ss.deleteSheet(existingSheet);
  _ss.insertSheet(sheetName).setColumnWidth(1, 1000).activate();
}

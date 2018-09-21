function createMailingText() {
  initOptions();
  insertText();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.addMenu('GoldenCode Report', [
    {name: 'Создать текст для рассылки', functionName: 'createMailingText'}
  ]);
}

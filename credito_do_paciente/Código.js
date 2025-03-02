// arquivo: sidebar_js
// versão: 1.0
// autor: Juliano Ceconi

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Navegação')
    .addItem('Abrir Navegador', 'showSidebar')
    .addToUi();
    
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Navegador de Guias')
    .setWidth(250);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetNames() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(sheet => sheet.getName());
}

function navigateToSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.activate();
}

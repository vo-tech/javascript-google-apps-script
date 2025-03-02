// Versão: 1.0
// Autor: Juliano Ceconi

function capitalizeFirstLetterInSheet() {
    // Obter a planilha ativa
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Obter todas as células com valores
    var range = sheet.getDataRange();
    var values = range.getValues();
    
    // Percorrer todas as células e modificar os valores
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (typeof values[i][j] === 'string' && values[i][j] !== "") {
          values[i][j] = values[i][j].toLowerCase().replace(/^\w/, c => c.toUpperCase());
          values[i][j] = values[i][j].toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
        }
      }
    }
    
    // Definir os novos valores na planilha
    range.setValues(values);
  }
  
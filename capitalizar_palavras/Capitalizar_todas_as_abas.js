// Versão: 1.0
// Autor: Juliano Ceconi

function capitalizeFirstLetterInAllSheets() {
    // Obter o arquivo ativo com todas as abas
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Obter todas as planilhas (abas)
    var sheets = spreadsheet.getSheets();
    
    // Percorrer todas as planilhas
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      
      // Obter o intervalo de dados da planilha atual
      var range = sheet.getDataRange();
  
      // Definir a fonte de toda a planilha para 'Google Sans'
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily("Google Sans");
  
      // Obter os valores da planilha atual
      var values = range.getValues();
      
      // Verificar se há dados na planilha
      if (values.length === 0) {
        continue; // Se a planilha estiver vazia, pular para a próxima
      }
  
      // Percorrer todas as células e modificar os valores
      for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] === 'string' && values[i][j].trim() !== "") {
            // Quebrar o texto em palavras e capitalizar a primeira letra de cada palavra
            values[i][j] = values[i][j]
              .toLowerCase() // Converter todo o texto para minúsculas
              .replace(/(^|\s)([a-zà-úãõç])/g, function(match, p1, p2) {
                // Capitalizar a primeira letra, mesmo com caracteres acentuados
                return p1 + p2.toUpperCase();
              });
          }
        }
      }
      
      // Atualizar a planilha com os novos valores
      range.setValues(values);
    }
  }
  
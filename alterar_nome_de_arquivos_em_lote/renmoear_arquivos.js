function renmoearArquivos() {
    // ID da pasta no Drive
    var FOLDER_ID = "ocultado";
    
    // Nome da aba onde estão os dados (A: nome antigo, B: nome novo)
    var SHEET_NAME = "ocultado";
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
  
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("Não há dados suficientes na planilha.");
      return;
    }
  
    var range = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  
    // Acessar a pasta
    var folder = DriveApp.getFolderById(FOLDER_ID);
    
    for (var i = 0; i < range.length; i++) {
      var nomeAntigo = range[i][0];
      var nomeNovo = range[i][1];
      
      if (nomeAntigo && nomeNovo) {
        var files = folder.getFilesByName(nomeAntigo);
        if (files.hasNext()) {
          var file = files.next();
          file.setName(nomeNovo);
          Logger.log("Renomeado: " + nomeAntigo + " -> " + nomeNovo);
        } else {
          Logger.log("Arquivo não encontrado: " + nomeAntigo);
        }
      }
    }
  
    Logger.log("Processo concluído.");
  }
  
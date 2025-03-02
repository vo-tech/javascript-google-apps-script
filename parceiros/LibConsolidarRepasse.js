// arquivo: LibConsolidarRepasse.gs
// versão: 1.1
// autor: Juliano Ceconi

function launchRepasseDate() {
    var partnerSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var partnerDataRange = partnerSheet.getRange("A3:D" + partnerSheet.getLastRow());
    var partnerData = partnerDataRange.getValues();
  
    var caixaSpreadsheet = SpreadsheetApp.openById('ocultado');
    var guiasSheet = caixaSpreadsheet.getSheetByName('guias');
    var guiasDataRange = guiasSheet.getRange("A1:P" + guiasSheet.getLastRow());
    var guiasData = guiasDataRange.getValues();
  
    // Create a map of invoice numbers to their row indices in 'guias' sheet
    var invoiceMap = {};
    for (var i = 0; i < guiasData.length; i++) {
      invoiceMap[guiasData[i][0]] = i; // Assuming invoice numbers are in column A (index 0)
    }
  
    // Prepare an array to hold the updated repassage dates
    var updates = [];
  
    for (var i = 0; i < partnerData.length; i++) {
      var invoiceNumber = partnerData[i][0]; // Column A (index 0)
      var repasseDate = partnerData[i][3];  // Column D (index 3)
      if (invoiceNumber in invoiceMap) {
        var rowIndex = invoiceMap[invoiceNumber] + 1; // +1 for 1-based indexing
        // Store the update as [value, row, column]
        updates.push([repasseDate, rowIndex, 16]); // Column P is index 16 (A=0, B=1, ..., P=16)
      }
    }
  
    // Update the 'guias' sheet in bulk
    if (updates.length > 0) {
      for (var i = 0; i < updates.length; i++) {
        var value = updates[i][0];
        var row = updates[i][1];
        var col = updates[i][2];
        guiasSheet.getRange(row, col).setValue(value);
      }
      SpreadsheetApp.flush(); // Ensure all changes are applied
      SpreadsheetApp.getUi().alert(updates.length + " linhas atualizadas.");
    } else {
      SpreadsheetApp.getUi().alert("Nenhuma linha foi atualizada.");
    }
  }
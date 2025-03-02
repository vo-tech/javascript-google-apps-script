
function consolidateValues() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('Tabela'); // Nome da aba original
    const targetSheet = ss.getSheetByName('Consolidado') || ss.insertSheet('Consolidado');
    
    targetSheet.clear();
    const data = sourceSheet.getDataRange().getValues();
    
    let output = [];
    for (let i = 1; i < data.length; i++) {
      let name = data[i][0];
      for (let j = 1; j < data[i].length; j++) {
        let value = data[i][j];
        if (value !== 0 && value !== '') {
          output.push([name, value]);
        }
      }
    }
    
    targetSheet.getRange(1, 1, output.length, 2).setValues(output);
  }
// Versão 1.3
// Autor: Juliano Ceconi

function processarRelatorioComissao() {
    try {
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName("RelatorioNF");
      if (!sheet) throw new Error("A aba 'RelatorioNF' não foi encontrada.");
  
      var range = sheet.getDataRange();
      var values = range.getValues();
  
      // Limpeza e capitalização dos textos
      values = values.map(row => row.map(cell => (typeof cell === 'string') ? cleanAndTitleCase(cell) : cell));
      range.setValues(values);
  
      // Cabeçalhos e criação da nova aba
      var headers = ["Nº Guia", "Caixa", "C", "D", "E", "F", "G", "H", "I", "Procedimento", "Instituição", "L", "M", "N", "Nome Completo"];
      var newSheet = createOrClearSheet(spreadsheet, "RelatorioNF_edit", headers);
  
      var caixaMap = getCaixaMap();
      var normalizedMap = normalizeMap(caixaMap);
  
      // Processar os dados de uma vez
      var newData = values.slice(1).map(row => processRow(row, normalizedMap));
  
      newSheet.getRange(2, 1, newData.length, headers.length).setValues(newData);
      newSheet.getRange(2, 1, newData.length, headers.length).sort(1);
      newSheet.autoResizeColumns(1, newSheet.getLastColumn());
  
      // Backup em arquivo CSV apenas da nova aba
      backupAsCSV(newSheet);
  
      var notasFiscaisSpreadsheet = SpreadsheetApp.openById("ocultado");
      var triagemSheet = notasFiscaisSpreadsheet.getSheetByName("Triagem");
      if (!triagemSheet) throw new Error('A guia "Triagem" não foi encontrada.');
  
      // Transferir dados da aba original "RelatorioNF" para a aba "Triagem"
      transferirDadosTriagem(sheet, triagemSheet);
  
      Logger.log("Processamento concluído com sucesso!");
    } catch (error) {
      Logger.log("Erro: " + error.message);
      SpreadsheetApp.getUi().alert("Erro durante a execução: " + error.message);
    }
  }
  
  // Funções auxiliares e melhorias
  function cleanAndTitleCase(str) {
    return str.replace(/\s+/g, ' ').trim().replace(/\w\S*/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
  }
  
  function createOrClearSheet(spreadsheet, sheetName, headers) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) sheet.clear(); else sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(headers);
    return sheet;
  }
  
  function getCaixaMap() {
    return {
    // ocultado
    };
  }
  
  function normalizeMap(caixaMap) {
    var normalized = {};
    for (var key in caixaMap) {
      normalized[normalize(key)] = caixaMap[key];
    }
    return normalized;
  }
  
  function normalize(str) {
    return str.replace(/\s+/g, ' ').trim().toLowerCase();
  }
  
  function processRow(row, normalizedMap) {
    var caixaNome = row[14].trim();
    row[14] = normalizedMap[normalize(caixaNome)] || caixaNome;
    return [
      row[0], row[14], "", "", "", "", "", "", "", row[15], row[12], "", "", "", row[1]
    ];
  }
  
  function backupAsCSV(sheet) {
    var folder = DriveApp.getFolderById('ocultado');
    var today = new Date();
    var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd HH'h'mm");
    var fileName = formattedDate + " RelatorioComissaoNF";
  
    // Verificar se já existe um arquivo com o mesmo nome e adicionar um número sequencial
    var files = folder.getFilesByName(fileName + ".csv");
    var count = 1;
    while (files.hasNext()) {
      fileName = formattedDate + " RelatorioComissaoNF (" + count + ")";
      files = folder.getFilesByName(fileName + ".csv");
      count++;
    }
  
    var csvData = sheet.getDataRange().getValues().map(row => row.join(",")).join("\n");
    var file = folder.createFile(fileName + ".csv", csvData, MimeType.CSV);
    Logger.log("Backup criado: " + file.getName());
  }
  
  function transferirDadosTriagem(sourceSheet, targetSheet) {
    try {
      // Capturar os valores na coluna "C" da aba "Triagem" para comparação
      var triagemData = targetSheet.getRange(2, 3, targetSheet.getLastRow() - 1, 1).getValues();
      var triagemValues = triagemData.map(function(row) { return row[0]; });
  
      // Capturar dados da aba "RelatorioNF" (sem o cabeçalho)
      var relatorioData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  
      // Filtrar apenas as linhas cujos valores em "A" não foram encontrados na coluna "C" de "Triagem"
      var dataToPaste = relatorioData.filter(function(row) {
        return triagemValues.indexOf(row[0]) === -1;
      });
  
      // Identificar a última linha preenchida na aba "Triagem"
      var lastRowTriagem = targetSheet.getLastRow();
  
      if (dataToPaste.length > 0) {
        // Colar os dados filtrados a partir da coluna "C", deixando as colunas "A" e "B" em branco
        targetSheet.getRange(lastRowTriagem + 1, 3, dataToPaste.length, dataToPaste[0].length).setValues(dataToPaste);
        
        // Limpar a coluna P (16ª coluna) das novas linhas adicionadas
        var rangeP = targetSheet.getRange(lastRowTriagem + 1, 16, dataToPaste.length, 1);
        rangeP.clearContent();
      }
  
      Logger.log("Dados transferidos para a aba 'Triagem' com sucesso.");
    } catch (error) {
      Logger.log('Erro ao transferir dados para a aba "Triagem": ' + error.message);
      SpreadsheetApp.getUi().alert('Erro ao transferir dados para a aba "Triagem": ' + error.message);
    }
  }
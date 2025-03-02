// arquivo: GuiasAnalitico.gs
// versão 1.15
// autor: Juliano Ceconi

function formatarGuiasAnalitico() {
    try {
      // Abrir a planilha e selecionar a aba "GuiasAnalitico"
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("GuiasAnalitico");
      if (!sheet) throw new Error("A aba 'GuiasAnalitico' não foi encontrada.");
  
      // Transformar todo o texto em formato "Title" (Primeira letra maiúscula) e limpar e formatar os dados conforme as instruções
      var data = sheet.getDataRange().getValues().map(row => {
        var guia = parseInt(row[0], 10);
        if (guia >= 100 && !isNaN(guia)) { // Filtra linhas onde o valor de "Guia" é maior ou igual a 100 e não está vazio
          return row.map((cell, index) => {
            if (typeof cell === 'string') cell = toTitleCase(cell);
            if (index === 9 && !cell) cell = "Cancelado"; // Preencher "Forma De Pagamento" com "Cancelado" onde estiver em branco
            return cell;
          });
        }
        return null;
      }).filter(row => row !== null);
  
      // Ordenar por "Guia" (primeira coluna)
      data.sort((a, b) => a[0] - b[0]);
  
      // Selecionar as colunas de interesse para o relatório final
      var finalData = data.map(row => [
        row[0], // Guia
        row[2], // Data Emissão
        row[3], // Agenda
        row[9], // Forma De Pagamento (dados copiados de GuiasAnalitico - coluna J)
        row[6], // Valor Guia
        row[4], // Repasse
        row[5], // Comissão
        row[1]  // Paciente
      ]);
  
      // Aplicar as substituições na coluna "Forma De Pagamento" (índice 3) de finalData
      finalData = finalData.map(row => {
        row[3] = substitutePaymentTerms(row[3]);
        return row;
      });
  
      // Criar ou limpar a aba para o relatório formatado
      var finalSheet = createOrClearSheet(ss, "GuiasAnalitico_edit", ["Guia", "Data Emissão", "Agenda", "Forma De Pagamento", "Valor Guia", "Repasse", "Comissão", "Paciente"]);
  
      // Inserir os dados na nova aba e aplicar o preenchimento amarelo para as células com "Cancelado" na coluna "Forma De Pagamento"
      if (finalData.length > 0) {
        finalSheet.getRange(2, 1, finalData.length, finalData[0].length).setValues(finalData);
        var backgrounds = finalData.map(row => row[3] === "Cancelado" ? ["yellow"] : [null]);
        finalSheet.getRange(2, 4, finalData.length, 1).setBackgrounds(backgrounds);
        finalSheet.autoResizeColumns(1, finalData[0].length);
      }
  
      // Backup em arquivo CSV da aba formatada
      backupAsCSV(finalSheet);
  
      Logger.log("Relatório formatado com sucesso na aba 'GuiasAnalitico_edit'. As células com 'Cancelado' foram destacadas em amarelo e o backup foi criado.");
    } catch (e) {
      Logger.log("Erro: " + e.message);
      SpreadsheetApp.getUi().alert("Erro durante a execução: " + e.message);
    }
  }
  
  // Função para converter as expressões de forma parcial conforme mapeamento definido
  function substitutePaymentTerms(paymentString) {
    // Mapeamento dos termos a serem substituídos
    var substitutions = {
      "Cartão Credito Sem Taxa": "Crédito",
      "Cartão Crédito Com Taxa": "Crédito",
      "Cartão Crédito Sem Taxa": "Crédito",
      "Cartão Debito Sem Taxa": "Débito",
      "Credito Do Paciente": "Saldo Anterior"
    };
  
    // Iterar sobre cada termo do mapeamento e substituir todas as ocorrências
    for (var key in substitutions) {
      // Cria uma expressão regular com flag 'g' para substituir todas as ocorrências
      var regex = new RegExp(key, "g");
      paymentString = paymentString.replace(regex, substitutions[key]);
    }
    return paymentString;
  }
  
  // Função para converter uma string para o formato Title Case
  function toTitleCase(str) {
    return str.replace(/\w\S*/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
  }
  
  // Função para criar uma nova aba ou limpar a existente com os cabeçalhos fornecidos
  function createOrClearSheet(spreadsheet, sheetName, headers) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    sheet.appendRow(headers);
    return sheet;
  }
  
  // Função para realizar o backup dos dados como um arquivo CSV
  function backupAsCSV(sheet) {
    var folder = DriveApp.getFolderById('ocultado');
    var today = new Date();
    var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd HH'h'mm");
    var fileName = formattedDate + " GuiasAnalitico";
  
    // Verificar se já existe um arquivo com o mesmo nome e adicionar um número sequencial
    var files = folder.getFilesByName(fileName + ".csv");
    var count = 1;
    while (files.hasNext()) {
      fileName = formattedDate + " GuiasAnalitico (" + count + ")";
      files = folder.getFilesByName(fileName + ".csv");
      count++;
    }
  
    var csvData = sheet.getDataRange().getValues().map(row => row.join(",")).join("\n");
    var file = folder.createFile(fileName + ".csv", csvData, MimeType.CSV);
    Logger.log("Backup criado: " + file.getName());
  }
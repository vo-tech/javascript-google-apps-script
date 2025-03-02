// Versão: 1.0
// Autor: Juliano Ceconi

function copiarNovasLinhas(e) {
  Logger.log("Função copiarNovasLinhas iniciada");
  
  // Configurações personalizáveis
  var planilhaOrigem = "FormsDespesas";
  var colunaVerificacao = "A";
  var mensagemLog = "Script executado com sucesso.";

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOrigem = spreadsheet.getSheetByName(planilhaOrigem);
    var idPlanilhaDestino = "ocultado";
    var sheetDestino = SpreadsheetApp.openById(idPlanilhaDestino).getSheetByName("saidas");

    Logger.log("Planilha de origem: " + (sheetOrigem ? sheetOrigem.getName() : "não encontrada"));
    Logger.log("Planilha de destino: " + (sheetDestino ? sheetDestino.getName() : "não encontrada"));

    // Verifica se as planilhas foram encontradas
    if (!sheetOrigem || !sheetDestino) {
      throw new Error("Uma ou ambas as planilhas não foram encontradas.");
    }

    var ultimaLinhaOrigem = sheetOrigem.getLastRow();
    Logger.log("Última linha da origem: " + ultimaLinhaOrigem);

    var valorColunaVerificacao = sheetOrigem.getRange(colunaVerificacao + ultimaLinhaOrigem).getValue();
    Logger.log("Valor da coluna de verificação: " + valorColunaVerificacao);

    if (valorColunaVerificacao !== "") {
      // Lê os dados da última linha da planilha de origem
      var dadosOrigem = sheetOrigem.getRange(ultimaLinhaOrigem, 1, 1, sheetOrigem.getLastColumn()).getValues();
      Logger.log("Dados da origem: " + JSON.stringify(dadosOrigem));
      
      // Escreve os dados na última linha da planilha de destino
      var ultimaLinhaDestino = sheetDestino.getLastRow() + 1;
      Logger.log("Última linha do destino: " + ultimaLinhaDestino);
      
      sheetDestino.getRange(ultimaLinhaDestino, 1, 1, dadosOrigem[0].length).setValues(dadosOrigem);
      
      Logger.log(mensagemLog);
    } else {
      Logger.log("A coluna de verificação está vazia. Linha não copiada.");
    }
  } catch (error) {
    Logger.log("Erro ao copiar dados: " + error.message);
  }
  
  Logger.log("Função copiarNovasLinhas finalizada");
}

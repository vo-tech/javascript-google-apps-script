// Versão: 1.1
// Autor: Juliano Ceconi

function executarCodigosEmSequencia() {
    try {
      // Executa o código "Relatório NF GS"
      Logger.log("Executando 'Relatório Comissao NF Script'...");
      processarRelatorioComissao();
      
      Logger.log("'Relatório NF GS' concluído com sucesso.");
      
      // Executa o código "Guias Analitico GS"
      Logger.log("Executando 'Guias Analitico Script'...");
      formatarGuiasAnalitico();
      
      Logger.log("'Guias Analitico GS' concluído com sucesso.");
      
      // Executa o código "Para Planilha Caixa GS"
      Logger.log("Executando 'Para Planilha Caixa GS'...");
      integrarPlanilhas();
      
      Logger.log("'Para Planilha Caixa GS' concluído com sucesso.");
      
    } catch (e) {
      Logger.log("Erro durante a execução: " + e.message);
    }
  }
  
  // Código "Relatório NF GS"
  function processarRelatorioComissao() {
  
  }
  
  // Código "Guias Analitico GS"
  function formatarGuiasAnalitico() {
  
  }
  
  // Código "Para Planilha Caixa GS"
  function integrarPlanilhas() {
  
  }
  
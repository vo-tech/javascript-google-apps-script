// arquivo: atualizarPermissoes.gs
// Versão: 1.0
// Autor: Juliano Ceconi


// Concede acesso de editor para múltiplos emails em planilhas listadas
// @returns {void}
function concederAcessoEmMassa() {
    // Lista de emails para acesso
    const EMAILS = [
     // ocultado
    ];
    
    const RANGE_URLS = 'B2:B';
    const COLUNA_STATUS = 5;
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getActiveSheet();
      const urls = sheet.getRange(RANGE_URLS)
                       .getValues()
                       .filter(row => row[0] !== '');
      
      const cache = CacheService.getScriptCache();
      
      const resultados = urls.map((url, index) => {
        try {
          const cacheKey = `planilha_${url[0]}`;
          let status = cache.get(cacheKey);
          
          if (!status) {
            const idPlanilha = extrairIdUrl(url[0]);
            const planilha = SpreadsheetApp.openById(idPlanilha);
            
            EMAILS.forEach(email => {
              planilha.addEditor(email);
            });
            
            const dataHora = new Date().toLocaleString('pt-BR');
            status = `Acesso concedido em ${dataHora}`;
            cache.put(cacheKey, status, 21600);
          }
          
          sheet.getRange(index + 2, COLUNA_STATUS).setValue(status);
          return true;
          
        } catch (erro) {
          console.error(`Erro ao processar ${url[0]}: ${erro.message}`);
          sheet.getRange(index + 2, COLUNA_STATUS).setValue(`Erro: ${erro.message}`);
          return false;
        }
      });
      
      const sucessos = resultados.filter(r => r).length;
      const falhas = resultados.filter(r => !r).length;
      
      console.log(`Processamento concluído: ${sucessos} sucessos, ${falhas} falhas`);
      
    } catch (erro) {
      console.error('Erro geral:', erro.message);
      throw new Error(`Falha na execução: ${erro.message}`);
    }
  }
  
  // Extrai o ID da planilha a partir da URL
  // @param {string} url - URL da planilha Google
  // @returns {string} ID da planilha
  function extrairIdUrl(url) {
    const regex = /[-\w]{25,}/;
    const matches = url.match(regex);
    
    if (!matches) {
      throw new Error('URL inválida');
    }
    
    return matches[0];
  }
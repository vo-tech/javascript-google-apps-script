// arquivo: autofaturaSync.gs
// Versão: 1.9
// Autor: Juliano Ceconi

// Script para aplicar modelo 'autofatura' em múltiplas planilhas de parceiros

/**
 * Função principal para aplicar o modelo de autofatura
 */
function aplicarModeloAutofatura() {
    console.time('Tempo de Execução');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetParceiros = ss.getSheetByName('parceiros');
    const sheetModelo = ss.getSheetByName('autofatura');
    
    if (!sheetParceiros || !sheetModelo) {
      Browser.msgBox('Erro: Sheets "parceiros" ou "autofatura" não encontradas.');
      return;
    }
  
    const dadosParceiros = sheetParceiros.getDataRange().getValues();
    const numTotal = dadosParceiros.length - 1; // Excluindo o cabeçalho
  
    const intervalo = selecionarIntervaloLinhas(numTotal);
    if (!intervalo) return;
  
    // Converter números de linhas reais (planilha) para índices do array (subtraindo 1)
    const inicioArray = intervalo.inicio - 1;
    const fimArray = intervalo.fim - 1;
    
    const resultados = processarPlanilhas(dadosParceiros, inicioArray, fimArray, sheetModelo);
    atualizarResultados(sheetParceiros, resultados, intervalo.inicio);
  
    const numProcessadas = intervalo.fim - intervalo.inicio + 1;
    Logger.log(`Processamento concluído. ${numProcessadas} planilhas processadas.`);
    Browser.msgBox(`Processo concluído! ${numProcessadas} planilhas foram processadas.`);
    
    console.timeEnd('Tempo de Execução');
  }
  
  /**
   * Solicita ao usuário o número ou intervalo de linhas a processar, com validação aprimorada
   * @param {number} numTotal - Número total de linhas disponíveis (excluindo o cabeçalho)
   * @return {Object|null} Objeto com início e fim do intervalo ou null se cancelado/entrada inválida
   */
  function selecionarIntervaloLinhas(numTotal) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Selecionar linha(s) a processar',
      `Digite o número único ou o intervalo de linhas a processar (ex: 16 ou 2-10):\nTotal de linhas disponíveis (excluindo o cabeçalho): ${numTotal}`,
      ui.ButtonSet.OK_CANCEL
    );
  
    if (response.getSelectedButton() == ui.Button.CANCEL) {
      return null;
    }
    
    const entrada = response.getResponseText().trim();
    
    if (!entrada) {
      ui.alert('Entrada vazia. Por favor, insira um número ou intervalo válido.');
      return null;
    }
    
    let inicio, fim;
    
    if (entrada.indexOf('-') !== -1) {
      // Se for um intervalo no formato "início-fim"
      const partes = entrada.split('-').map(parte => parte.trim());
      if (partes.length !== 2) {
        ui.alert(`Formato inválido "${entrada}". Use o formato exato "início-fim", ex: 2-10.`);
        return null;
      }
      inicio = parseInt(partes[0]);
      fim = parseInt(partes[1]);
      
      if (isNaN(inicio) || isNaN(fim)) {
        ui.alert(`Valores numéricos inválidos. Recebido: "${entrada}".`);
        return null;
      }
      
    } else {
      // Se for apenas um número, seta início e fim iguais
      inicio = parseInt(entrada);
      if (isNaN(inicio)) {
        ui.alert(`Valor numérico inválido. Recebido: "${entrada}".`);
        return null;
      }
      fim = inicio;
    }
    
    // Verifica as condições: considerar a linha 1 como cabeçalho
    if (inicio < 2 || fim < 2) {
      ui.alert('As linhas devem ser maiores ou iguais a 2, pois a linha 1 é o cabeçalho.');
      return null;
    }
    if (fim > numTotal + 1) {
      ui.alert(`O valor máximo permitido é ${numTotal + 1} (incluindo o cabeçalho).`);
      return null;
    }
    if (inicio > fim) {
      ui.alert('O valor inicial não pode ser maior que o valor final.');
      return null;
    }
    
    return { inicio, fim };
  }
  
  /**
   * Extrai o ID da planilha a partir do link
   * @param {string} link - Link da planilha do Google Sheets
   * @return {string|null} ID da planilha ou null se não for possível extrair
   */
  function extrairIdDaplanilha(link) {
    if (!link) return null;
    const regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
    const match = link.match(regex);
    return match ? match[1] : null;
  }
  
  /**
   * Processa as planilhas dos parceiros
   * @param {Array} dadosParceiros - Dados dos parceiros
   * @param {number} inicio - Índice inicial do array (corresponde à linha da planilha - 1)
   * @param {number} fim - Índice final do array (corresponde à linha da planilha - 1)
   * @param {Sheet} sheetModelo - Sheet modelo a ser copiada
   * @return {Array} Resultados do processamento
   */
  function processarPlanilhas(dadosParceiros, inicio, fim, sheetModelo) {
    const resultados = [];
  
    for (let i = inicio; i <= fim; i++) {
      const [nomeParceiro, linkPlanilha] = dadosParceiros[i];
      
      if (!linkPlanilha) {
        resultados.push([new Date(), `ERRO: ${nomeParceiro} - Link da planilha não fornecido`]);
        Logger.log(`Erro: Link da planilha não fornecido para ${nomeParceiro}`);
        continue;
      }
  
      const idPlanilha = extrairIdDaplanilha(linkPlanilha);
      
      if (!idPlanilha) {
        resultados.push([new Date(), `ERRO: ${nomeParceiro} - ID da planilha não pôde ser extraído`]);
        Logger.log(`Erro ao extrair ID da planilha para ${nomeParceiro}`);
        continue;
      }
      
      try {
        const planilhaParceiro = SpreadsheetApp.openById(idPlanilha);
        let sheetAutofatura = planilhaParceiro.getSheetByName('autofatura');
        let tempA = null;
        let tempF = null;
        
        if (sheetAutofatura) {
          const lastRow = sheetAutofatura.getLastRow();
          if (lastRow >= 3) {
            tempA = sheetAutofatura.getRange("A3:A" + lastRow).getValues();
            tempF = sheetAutofatura.getRange("F3:F" + lastRow).getValues();
          }
          planilhaParceiro.deleteSheet(sheetAutofatura);
        }
        
        // Copia a sheet modelo inteira, preservando fórmulas e formatações
        sheetAutofatura = sheetModelo.copyTo(planilhaParceiro);
        sheetAutofatura.setName('autofatura');
        
        // Depois das alterações necessárias, colar os dados sem formatação (apenas valores)
        if (tempA && tempF) {
          const numRowsA = tempA.length;
          const numRowsF = tempF.length;
          // Presumindo que ambas colunas tenham mesma quantidade de linhas copiadas
          if(numRowsA > 0) {
            sheetAutofatura.getRange("A3:A" + (2 + numRowsA)).setValues(tempA);
          }
          if(numRowsF > 0) {
            sheetAutofatura.getRange("F3:F" + (2 + numRowsF)).setValues(tempF);
          }
        }
        
        resultados.push([new Date(), `Sucesso: ${nomeParceiro}`]);
        Logger.log(`Modelo aplicado com sucesso para: ${nomeParceiro}`);
      } catch (erro) {
        resultados.push([new Date(), `ERRO: ${nomeParceiro} - ${erro.toString()}`]);
        Logger.log(`Erro ao processar ${nomeParceiro}: ${erro.toString()}`);
      }
    }
  
    return resultados;
  }
  
  /**
   * Atualiza os resultados na planilha de parceiros
   * @param {Sheet} sheetParceiros - Sheet de parceiros
   * @param {Array} resultados - Resultados do processamento
   * @param {number} inicioLote - Número da linha inicial (exato, conforme o usuário informou)
   */
  function atualizarResultados(sheetParceiros, resultados, inicioLote) {
    // Atualiza a partir da linha exata informada pelo usuário
    const range = sheetParceiros.getRange(inicioLote, 3, resultados.length, 2);
    range.setValues(resultados);
  }
  
  /**
   * Adiciona menu personalizado à interface do usuário
   */
  function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('ocultado')
      .addItem('Aplicar modelo da sheet autofatura', 'aplicarModeloAutofatura')
      .addToUi();
  }
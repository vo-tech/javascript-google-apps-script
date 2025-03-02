// arquivo: converterParaFormatoTitulo.gs
// versão: 1.0
// autor: Juliano Ceconi

function formatCellsAsTitleCase() {
    // Obtém a planilha ativa
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // Define o intervalo de dados (todas as células com dados na planilha ativa)
    const range = sheet.getDataRange();
    // Obtém os valores do intervalo como uma matriz bidimensional
    const values = range.getValues();
    
    // Inicializa um array para armazenar logs de execução
    let log = [];
    
    // Mapeia os valores para aplicar a formatação de título em cada célula
    const formattedValues = values.map((row, rowIndex) => 
      row.map((cell, colIndex) => {
        // Verifica se o valor da célula é uma string
        if (typeof cell === 'string') {
          // Formata o texto para título
          const formattedCell = toTitleCase(cell);
          // Se houver alteração na célula, adiciona ao log
          if (cell !== formattedCell) {
            log.push(`Célula alterada: Linha ${rowIndex + 1}, Coluna ${colIndex + 1} | Original: '${cell}' | Novo: '${formattedCell}'`);
          }
          // Retorna a célula formatada
          return formattedCell;
        }
        // Se não for uma string, retorna o valor original
        return cell;
      })
    );
  
    // Aplica os valores formatados de volta ao intervalo
    range.setValues(formattedValues);
  
    // Registra os logs em uma nova planilha chamada 'Logs de Execução'
    const logSheet = getOrCreateLogSheet();
    log.forEach(entry => logSheet.appendRow([new Date(), entry]));
  
    // Exibe uma notificação para o usuário
    SpreadsheetApp.getUi().alert(`Formatação concluída! ${log.length} células foram alteradas.`);
  }
  
  // Função para criar ou obter a planilha de logs
  function getOrCreateLogSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Logs de Execução');
    // Cria a planilha de logs se ela não existir
    if (!logSheet) {
      logSheet = ss.insertSheet('Logs de Execução');
      logSheet.appendRow(['Data e Hora', 'Log de Execução']);
    }
    return logSheet;
  }
  
  // Função para formatar o texto para título, respeitando acentuação e caracteres especiais
  function toTitleCase(text) {
    // Divide o texto em palavras e capitaliza a primeira letra de cada palavra
    return text.toLowerCase().replace(/(?:^|\s|["'([{])+\S/g, function(word) {
      return word.toUpperCase();
    });
  }
  
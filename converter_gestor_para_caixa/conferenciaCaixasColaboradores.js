// arquivo: conferenciaCaixasColaboradores.gs
// versão: 1.2
// autor: Juliano Ceconi

function compararPlanilhasEPreencherData() {
    const caixaSheetId = 'ocultado'; // Planilha CAIXA
    const caixaSpreadsheet = SpreadsheetApp.openById(caixaSheetId);
    const caixaSheet = caixaSpreadsheet.getSheetByName('guias');
    const errosSheet = caixaSpreadsheet.getSheetByName('erros');
  
    const colaboradoresSheetIds = [
    // ocultado
    ];
  
    // Obter a data para conferência com base na célula C25 da aba "Scripts"
    let abaNome;
    const scriptsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Scripts');
    if (scriptsSheet) {
      abaNome = scriptsSheet.getRange('C25').getDisplayValue();
    }
    if (!abaNome) {
      const dataAtual = new Date();
      const dia = ('0' + dataAtual.getDate()).slice(-2);
      const mes = ('0' + (dataAtual.getMonth() + 1)).slice(-2);
      abaNome = `${dia}/${mes}`;
    }
  
    try {
      Logger.log('[INFO] Data selecionada para conferência: %s', abaNome);
  
      // Objeto para armazenar os dados dos colaboradores
      const colaboradorDataMap = {};
  
      // Ler os dados de cada colaborador e armazenar no objeto
      colaboradoresSheetIds.forEach(colaborador => {
        try {
          Logger.log('[INFO] Acessando a planilha do colaborador: %s...', colaborador.nome);
          const colaboradorSpreadsheet = SpreadsheetApp.openById(colaborador.id);
          const colaboradorSheet = colaboradorSpreadsheet.getSheetByName(abaNome);
          if (!colaboradorSheet) {
            Logger.log('[WARN] Aba "%s" não encontrada na planilha do colaborador %s.', abaNome, colaborador.nome);
            return;
          }
  
          Logger.log('[INFO] Lendo os dados da aba "%s" na planilha do colaborador %s...', abaNome, colaborador.nome);
          const lastRow = colaboradorSheet.getLastRow();
          // Expandindo range: de B a J (9 colunas a partir da coluna 2)
          const dataRange = colaboradorSheet.getRange(7, 2, lastRow - 6, 9);
          const dataValues = dataRange.getValues();
  
          // Mapeando os dados: índice 0 => Coluna B (guia), índice 3 => Coluna E (valor),
          // índice 7 => Coluna I (forma de pagamento)
          for (let index = 0; index < dataValues.length; index++) {
            const row = dataValues[index];
            const guia = row[0];
            const valor = row[3];
            const formaPagamento = row[7];
            if (/fim/i.test(guia)) {
              break;
            }
            if (guia && valor !== '') {
              colaboradorDataMap[guia] = {
                valor,
                formaPagamento, // Pode estar vazio
                colaborador: colaborador.nome,
                rowIndex: index + 7, // Linha real na planilha do colaborador
                sheetId: colaborador.id,
                sheetName: abaNome
              };
            }
          }
        } catch (error) {
          Logger.log('[ERROR] Erro ao processar a planilha do colaborador %s: %s', colaborador.nome, error.message);
        }
      });
  
      // Lendo os dados da planilha Caixa (sheet "guias")
      Logger.log('[INFO] Lendo os dados da planilha Caixa...');
      const caixaLastRow = caixaSheet.getLastRow();
      const caixaDataRange = caixaSheet.getRange(2, 1, caixaLastRow - 1, 11); // Colunas A a K
      const caixaDataValues = caixaDataRange.getValues();
  
      const atualizacoes = []; // Atualizações de data na planilha Caixa
      const erros = [];        // Armazena divergências encontradas
  
      // Mapear as guias na planilha Caixa: coluna A (guia) e valor na coluna J
      const caixaGuiasMap = {};
      caixaDataValues.forEach((row, index) => {
        const guia = row[0];
        if (guia) {
          caixaGuiasMap[guia] = {
            index: index + 2, // Linha real
            valor: row[9]     // Coluna J
          };
        }
      });
  
      // Processar cada guia encontrada nos dados dos colaboradores
      for (let guia in colaboradorDataMap) {
        const colaboradorInfo = colaboradorDataMap[guia];
        if (caixaGuiasMap.hasOwnProperty(guia)) {
          const caixaInfo = caixaGuiasMap[guia];
  
          // Atualiza a forma de pagamento, se houver conteúdo, na coluna O (índice 15)
          if (colaboradorInfo.formaPagamento && colaboradorInfo.formaPagamento.toString().trim() !== '') {
            caixaSheet.getRange(caixaInfo.index, 15).setValue(colaboradorInfo.formaPagamento);
            Logger.log('[INFO] Forma de pagamento atualizada para a guia %s: %s', guia, colaboradorInfo.formaPagamento);
          }
  
          // Comparar valor entre as planilhas
          if (colaboradorInfo.valor != caixaInfo.valor) {
            Logger.log('[WARN] Divergência para a guia %s: Caixa (%s) vs Colaborador (%s).', guia, caixaInfo.valor, colaboradorInfo.valor);
            erros.push([guia, caixaInfo.valor, colaboradorInfo.valor, abaNome, colaboradorInfo.colaborador]);
            // Marcar divergência na planilha do colaborador (faixa de B a J)
            const colaboradorSpreadsheet = SpreadsheetApp.openById(colaboradorInfo.sheetId);
            const colaboradorSheet = colaboradorSpreadsheet.getSheetByName(colaboradorInfo.sheetName);
            colaboradorSheet.getRange(colaboradorInfo.rowIndex, 2, 1, 9).setBackground('#FFEB3B'); // Amarelo
          } else {
            Logger.log('[INFO] Valores coincidem para a guia %s.', guia);
            // Marcar acerto na planilha do colaborador (faixa de B a J)
            const colaboradorSpreadsheet = SpreadsheetApp.openById(colaboradorInfo.sheetId);
            const colaboradorSheet = colaboradorSpreadsheet.getSheetByName(colaboradorInfo.sheetName);
            colaboradorSheet.getRange(colaboradorInfo.rowIndex, 2, 1, 9).setBackground('#D3EAFB'); // Azul
            atualizacoes.push({ index: caixaInfo.index, date: abaNome });
          }
        } else {
          Logger.log('[WARN] Guia %s não encontrada na planilha Caixa.', guia);
          erros.push([guia, '', colaboradorInfo.valor, abaNome, colaboradorInfo.colaborador]);
          const colaboradorSpreadsheet = SpreadsheetApp.openById(colaboradorInfo.sheetId);
          const colaboradorSheet = colaboradorSpreadsheet.getSheetByName(colaboradorInfo.sheetName);
          colaboradorSheet.getRange(colaboradorInfo.rowIndex, 2, 1, 9).setBackground('#FFA500'); // Laranja
        }
      }
  
      // Aplicar atualizações de data na planilha Caixa (coluna B)
      if (atualizacoes.length > 0) {
        atualizacoes.forEach(update => {
          caixaSheet.getRange(update.index, 2).setValue(update.date);
        });
        Logger.log('[INFO] Atualizações de data aplicadas na planilha Caixa.');
      }
  
      // Registrar divergências na planilha de erros
      if (erros.length > 0) {
        errosSheet.getRange(errosSheet.getLastRow() + 1, 1, erros.length, 5).setValues(erros);
        Logger.log('[INFO] Divergências registradas na planilha de erros.');
      }
  
      // Backup dos dados de erros em CSV no Google Drive
      backupDadosCSV(erros);
  
      Logger.log('[INFO] Resumo do processamento:');
      Logger.log('[INFO] - Total de guias processadas: %s', Object.keys(colaboradorDataMap).length);
      Logger.log('[INFO] - Divergências encontradas: %s', erros.length);
      Logger.log('[INFO] - Atualizações realizadas: %s', atualizacoes.length);
  
    } catch (error) {
      Logger.log('[ERROR] Erro geral: %s', error.message);
    }
  }
  
  
  function atualizarFormaPagamentoSeparado() {
    const caixaSheetId = 'ocultado'; // Planilha CAIXA
    const caixaSpreadsheet = SpreadsheetApp.openById(caixaSheetId);
    const caixaSheet = caixaSpreadsheet.getSheetByName('guias');
  
    const colaboradoresSheetIds = [
      // ocultado
    ];
  
    // Obter a aba de conferência a partir da célula C30 da aba "Scripts"
    let abaNome;
    const scriptsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Scripts');
    if (scriptsSheet) {
      abaNome = scriptsSheet.getRange('C30').getDisplayValue();
    }
    
    // Se a célula C30 estiver preenchida, usaremos a aba específica; caso contrário, iteraremos por todas as sheets.
    const usaAbaEspecifica = (abaNome && abaNome.trim() !== "");
    if (usaAbaEspecifica) {
      Logger.log('[INFO] Utilizando a aba especificada: %s', abaNome);
    } else {
      Logger.log('[INFO] Célula C30 está em branco. Iterando por todas as abas de cada planilha.');
    }
  
    // Objeto para armazenar os dados coletados dos colaboradores: chave é o número da guia
    const colaboradorDataMap = {};
  
    // Iterar sobre cada planilha de colaborador
    colaboradoresSheetIds.forEach(colaborador => {
      try {
        Logger.log('[INFO] Processando a planilha do colaborador: %s', colaborador.nome);
        const colaboradorSpreadsheet = SpreadsheetApp.openById(colaborador.id);
        let sheetsParaProcessar = [];
  
        if (usaAbaEspecifica) {
          const sheet = colaboradorSpreadsheet.getSheetByName(abaNome);
          if (!sheet) {
            Logger.log('[WARN] Aba "%s" não encontrada na planilha do colaborador %s.', abaNome, colaborador.nome);
          } else {
            sheetsParaProcessar.push(sheet);
          }
        } else {
          sheetsParaProcessar = colaboradorSpreadsheet.getSheets();
        }
  
        // Iterar por cada sheet definido para o colaborador
        sheetsParaProcessar.forEach(sheet => {
          try {
            const lastRow = sheet.getLastRow();
            if (lastRow < 7) {
              Logger.log('[WARN] Planilha %s do colaborador %s não possui linhas suficientes para processamento.', sheet.getName(), colaborador.nome);
              return;
            }
            const numRows = lastRow - 6;
            // Ler dados do intervalo de B a J (colunas 2 a 10), iniciando na linha 7
            const dataRange = sheet.getRange(7, 2, numRows, 9);
            const dataValues = dataRange.getValues();
  
            dataValues.forEach(row => {
              const guia = row[0];            // Coluna B: número da guia
              const formaPagamento = row[7];  // Coluna I: forma de pagamento
              if (typeof guia === 'string' && /fim/i.test(guia)) {
                // Se encontrar "fim", interromper o processamento dessa aba
                return;
              }
              if (guia && formaPagamento && formaPagamento.toString().trim() !== '') {
                // Caso haja duplicata, a última ocorrência sobrescreverá as anteriores
                colaboradorDataMap[guia] = {
                  formaPagamento: formaPagamento,
                  colaborador: colaborador.nome
                };
              }
            });
          } catch (error) {
            Logger.log('[ERROR] Erro ao processar a aba "%s" na planilha do colaborador %s: %s', sheet.getName(), colaborador.nome, error.message);
          }
        });
      } catch (error) {
        Logger.log('[ERROR] Erro ao acessar a planilha do colaborador %s: %s', colaborador.nome, error.message);
      }
    });
  
    // Ler os dados da planilha Caixa (sheet "guias")
    Logger.log('[INFO] Lendo os dados da planilha Caixa...');
    const caixaLastRow = caixaSheet.getLastRow();
    if (caixaLastRow < 2) {
      Logger.log('[WARN] A planilha "guias" está vazia ou sem dados.');
      return;
    }
    const numRowsCaixa = caixaLastRow - 1; // dados a partir da linha 2
  
    // Criar uma matriz temporária com os valores atuais da coluna O
    const colunaORange = caixaSheet.getRange(2, 15, numRowsCaixa, 1);
    const updateValues = colunaORange.getValues();
  
    // Mapear as guias na planilha Caixa (coluna A) para poder identificar as linhas correspondentes
    const caixaGuiasMap = {};
    const caixaDataRange = caixaSheet.getRange(2, 1, numRowsCaixa, 1);
    const caixaDataValues = caixaDataRange.getValues();
    caixaDataValues.forEach((row, index) => {
      const guia = row[0];
      if (guia) {
        caixaGuiasMap[guia] = index; // índice correspondente à linha na matriz updateValues
      }
    });
  
    // Acumular as atualizações: utilizar a estrutura temporária para atualizar a coluna O de uma única vez
    let updatesCount = 0;
    for (let guia in colaboradorDataMap) {
      if (caixaGuiasMap.hasOwnProperty(guia)) {
        const idx = caixaGuiasMap[guia];
        updateValues[idx][0] = colaboradorDataMap[guia].formaPagamento;
        Logger.log('[INFO] Atualizando guia %s com forma de pagamento: %s', guia, colaboradorDataMap[guia].formaPagamento);
        updatesCount++;
      } else {
        Logger.log('[WARN] Guia %s não encontrada na planilha Caixa.', guia);
      }
    }
  
    // Gravar todas as atualizações de uma única vez na coluna O
    if (updatesCount > 0) {
      colunaORange.setValues(updateValues);
      Logger.log('[INFO] Atualizações aplicadas em %s linhas.', updatesCount);
    } else {
      Logger.log('[INFO] Nenhuma atualização realizada.');
    }
  }
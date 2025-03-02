// Versão 1.12
// Autor: Juliano Ceconi

function integrarPlanilhas() {
    try {
      const competenciaMes = "02/25"; // Defina aqui o valor padrão para a coluna N (Competência)
      const ss = SpreadsheetApp.getActiveSpreadsheet();
  
      // Referências às planilhas dentro do mesmo arquivo
      const sheetGuias = ss.getSheetByName('GuiasAnalitico_edit');
      const sheetRelatorio = ss.getSheetByName('RelatorioNF_edit');
      let sheetNova = createOrClearSheet(ss, 'PlanilhaAposScript', [
        'Guia', 'Documentado em', 'Tipo', 'Filtro', 'Responsavel', 'Cidade', 'Data Emissão', 'Data Guia', 'Forma pgto', 'Valor recebido', 'Repasse', 'Comissão', 'Procedimento', 'Instituição', 'Tipo instituição', 'Data rep', 'Data NF', 'Competência', 'Nome'
      ]);
  
      // Obter os dados das planilhas
      const dadosGuias = sheetGuias.getDataRange().getValues();
      const dadosRelatorio = sheetRelatorio.getDataRange().getValues();
  
      // Mapear os dados da planilha RelatorioNF_edit para busca rápida
      const relatorioMap = new Map();
      for (let i = 1; i < dadosRelatorio.length; i++) {
        relatorioMap.set(dadosRelatorio[i][0], dadosRelatorio[i]);
      }
  
      // Array para armazenar todas as linhas que serão adicionadas de uma vez
      const novasLinhas = [];
  
      // Processar cada linha de GuiasAnalitico_edit e criar a nova linha
      for (let i = 1; i < dadosGuias.length; i++) {
        const guiaNumero = dadosGuias[i][0];
        const linhaRelatorio = relatorioMap.get(guiaNumero) || [];
  
        const linhaNova = [
          guiaNumero, // Coluna A: Guia Número
          '', // Coluna B: Documentado em
          'Entrada', // Coluna C: Tipo
          'Guia', // Coluna D: Filtro
          linhaRelatorio[1] || '', // Coluna E: Responsavel (Coluna B de RelatorioNF_edit)
          determinarCidade(guiaNumero), // Coluna F: Cidade baseada no último dígito do número da guia
          dadosGuias[i][1] || '', // Coluna G: Data Emissão
          dadosGuias[i][2] || '', // Coluna H: Data Guia
          dadosGuias[i][3] || '', // Coluna I: Forma pgto
          dadosGuias[i][4] || '', // Coluna J: Valor recebido
          dadosGuias[i][5] || '', // Coluna K: Repasse
          '', // Coluna L: Comissão
          linhaRelatorio[9] || '', // Coluna M: Procedimento
          linhaRelatorio[10] || '', // Coluna N: Instituição
          '', // Coluna O: Tipo instituição
          '', // Coluna P: Data rep
          '', // Coluna Q: Data NF
          competenciaMes, // Coluna R: Competência
          dadosGuias[i][7] || '' // Coluna S: Nome
        ];
  
        // Adicionar a nova linha ao array
        novasLinhas.push(linhaNova);
      }
  
      // Inserir os dados na nova aba
      if (novasLinhas.length > 0) {
        sheetNova.getRange(2, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
        sheetNova.autoResizeColumns(1, sheetNova.getLastColumn());
      }
  
      // Backup em arquivo CSV da aba processada
      backupAsCSV(sheetNova);
  
      // Transferir dados para a Planilha Caixa
      transferirParaPlanilhaCaixa(novasLinhas);
  
      Logger.log('Integração concluída com sucesso, backup criado em CSV e dados transferidos para a Planilha Caixa.');
    } catch (error) {
      Logger.log('Erro durante a integração das planilhas: ' + error.message);
      SpreadsheetApp.getUi().alert('Erro durante a execução: ' + error.message);
    }
  }
  
  // Funções auxiliares e melhorias
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
  
  function backupAsCSV(sheet) {
    var folder = DriveApp.getFolderById('ocultado');
    var today = new Date();
    var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd HH'h'mm");
    var fileName = formattedDate + " PlanilhaAposScript";
  
    // Verificar se já existe um arquivo com o mesmo nome e adicionar um número sequencial
    var files = folder.getFilesByName(fileName + ".csv");
    var count = 1;
    while (files.hasNext()) {
      fileName = formattedDate + " PlanilhaAposScript (" + count + ")";
      files = folder.getFilesByName(fileName + ".csv");
      count++;
    }
  
    var csvData = sheet.getDataRange().getValues().map(row => row.join(",")).join("\n");
    var file = folder.createFile(fileName + ".csv", csvData, MimeType.CSV);
    Logger.log("Backup criado: " + file.getName());
  }
  
  function determinarCidade(guiaNumero) {
    const ultimoDigito = guiaNumero.toString().slice(-1);
    switch (ultimoDigito) {
      case '1': return '1 Barreiras';
      case '2': return '2 Baianópolis';
      case '3': return '3 Santa Rita';
      default: return 'Erro';
    }
  }
  

  function transferirParaPlanilhaCaixa(novasLinhas) {
    try {
      const ssDestino = SpreadsheetApp.openById('ocultado'); // Substitua pelo ID correto da Planilha Caixa
      const sheetDestino = ssDestino.getSheetByName('guias');
      if (!sheetDestino) throw new Error("A aba 'guias' não foi encontrada na Planilha Caixa.");
  
      // Filtrar linhas que já não estejam presentes na Planilha Caixa
      const dadosExistentes = sheetDestino.getRange(2, 1, sheetDestino.getLastRow() - 1, 1).getValues().flat();
      const dadosParaInserir = novasLinhas.filter(linha => !dadosExistentes.includes(linha[0]));
  
      // Inserir os dados na Planilha Caixa
      if (dadosParaInserir.length > 0) {
        const ultimaLinha = sheetDestino.getLastRow();
        sheetDestino.getRange(ultimaLinha + 1, 1, dadosParaInserir.length, dadosParaInserir[0].length).setValues(dadosParaInserir);
      }
  
      Logger.log('Dados transferidos para a Planilha Caixa com sucesso.');
    } catch (error) {
      Logger.log('Erro ao transferir dados para a Planilha Caixa: ' + error.message);
      SpreadsheetApp.getUi().alert('Erro ao transferir dados para a Planilha Caixa: ' + error.message);
    }
  }

// arquivo: consolidar_contas_pagas.gs
// versão 1.8
// autor: Juliano Ceconi

const CONFIG = {
  SPREADSHEET_CONTAS_A_PAGAR_ID: 'ocultado',
  SHEET_CONTAS_A_PAGAR_NOME: 'contas_a_pagar',
  SPREADSHEET_SAIDAS_ID: 'ocultado',
  SHEET_SAIDAS_NOME: 'saidas',
  SHEET_RELATORIO_NOME: 'relatorio_contas_pagas',
  FOLDER_ID: 'ocultado',
  COR_LARANJA: '#FFA500',
  COR_CINZA: '#D3D3D3',
  TOLERANCIA_VALOR: 1.0 // tolerância de R$1,00 para considerar valores iguais
};

class ContasPagasTransferencia {
  constructor() {
    this.logger = new Logger();
    this.sheetContas = SpreadsheetApp.openById(CONFIG.SPREADSHEET_CONTAS_A_PAGAR_ID).getSheetByName(CONFIG.SHEET_CONTAS_A_PAGAR_NOME);
    this.sheetSaidas = SpreadsheetApp.openById(CONFIG.SPREADSHEET_SAIDAS_ID).getSheetByName(CONFIG.SHEET_SAIDAS_NOME);
    
    // Inicializa ou obtém a planilha de relatório
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_CONTAS_A_PAGAR_ID);
    let sheetRelatorio = ss.getSheetByName(CONFIG.SHEET_RELATORIO_NOME);
    if (!sheetRelatorio) {
      sheetRelatorio = ss.insertSheet(CONFIG.SHEET_RELATORIO_NOME);
      sheetRelatorio.appendRow(['TIMESTAMP', 'DESCRIÇÃO', 'DATA_PAGAMENTO', 'VALOR', 'STATUS', 'DRE']);
    }
    this.sheetRelatorio = sheetRelatorio;
    
    this.saidasMap = new Map();
    this.novasLinhasSaidas = [];
    this.relatorioLinhas = [];
    this.estatisticas = {
      processadas: 0,
      duplicadas: 0,
      novas: 0
    };
  }

  executar() {
    try {
      this.logger.log('Iniciando processamento de contas pagas...');
      this.logger.log(`Planilha contas: ${CONFIG.SPREADSHEET_CONTAS_A_PAGAR_ID}`);
      this.logger.log(`Planilha saídas: ${CONFIG.SPREADSHEET_SAIDAS_ID}`);
      
      this.carregarSaidasExistentes();
      this.processarContasAPagar();
      this.inserirNovasSaidas();
      this.gerarRelatorio();
      this.gerarCSV();
      
      this.logger.log(`Processamento concluído: ${this.estatisticas.processadas} contas processadas, ${this.estatisticas.novas} novas, ${this.estatisticas.duplicadas} duplicadas.`);
    } catch (error) {
      this.logger.error(`Erro durante o processamento: ${error.message}`, error.stack);
    }
  }

  carregarSaidasExistentes() {
    this.logger.log('Carregando saídas existentes...');
    const inicio = new Date();
    const dadosSaidas = this.sheetSaidas.getDataRange().getValues();
    
    dadosSaidas.slice(2).forEach((linha, index) => {
      // Cria várias versões da chave com diferentes tolerâncias
      const descricao = linha[7]; // coluna H
      const data = linha[1];      // coluna B
      const valor = linha[8];     // coluna I
      
      if (!descricao || !data) return;
      
      const valorNormalizado = this.normalizarValor(valor);
      const dataFormatada = this.formatarDataSemHora(data);
      
      // Adiciona à lista de saídas já existentes
      this.adicionarSaidaExistente(descricao, dataFormatada, valorNormalizado, index + 3);
    });
    
    const duracao = (new Date() - inicio) / 1000;
    this.logger.log(`Carregadas ${this.saidasMap.size} saídas existentes em ${duracao.toFixed(2)}s.`);
  }
  
  adicionarSaidaExistente(descricao, data, valor, linha) {
    // Cria chave padrão
    const chavePadrao = this.criarChaveSaida(descricao, data, valor);
    this.saidasMap.set(chavePadrao, linha);
    
    // Cria chaves com valores aproximados (para tolerar diferença de R$1)
    for (let diff = 0.01; diff <= CONFIG.TOLERANCIA_VALOR; diff += 0.01) {
      const valorMais = Math.round((valor + diff) * 100) / 100;
      const valorMenos = Math.round((valor - diff) * 100) / 100;
      
      const chaveMais = this.criarChaveSaida(descricao, data, valorMais);
      const chaveMenos = this.criarChaveSaida(descricao, data, valorMenos);
      
      this.saidasMap.set(chaveMais, linha);
      this.saidasMap.set(chaveMenos, linha);
    }
  }

  processarContasAPagar() {
    this.logger.log('Processando contas a pagar...');
    const inicio = new Date();
    const dadosContas = this.sheetContas.getDataRange().getValues();

    dadosContas.slice(2).forEach((linha, index) => {
      const pagoEm = linha[7];
      if (pagoEm && pagoEm.toString().trim() !== '') {
        this.estatisticas.processadas++;
        const resultado = this.processarContaPaga(linha, index + 3);
        if (resultado === 'duplicada') {
          this.estatisticas.duplicadas++;
        } else if (resultado === 'nova') {
          this.estatisticas.novas++;
        }
      }
    });
    
    const duracao = (new Date() - inicio) / 1000;
    this.logger.log(`Processamento concluído em ${duracao.toFixed(2)}s.`);
    this.logger.log(`Estatísticas: ${this.estatisticas.processadas} contas processadas.`);
    this.logger.log(`Encontradas ${this.estatisticas.duplicadas} duplicatas.`);
    this.logger.log(`Adicionadas ${this.estatisticas.novas} novas saídas.`);
  }

  processarContaPaga(linha, rowIndex) {
    const [motivo, competencia, valor, pagoEm, dre] = [linha[2], linha[3], linha[4], linha[7], linha[10]];
    
    // Validações básicas
    if (!motivo || !valor || !pagoEm) {
      this.logger.warn(`Conta na linha ${rowIndex} com dados incompletos: ${motivo}`);
      return 'ignorada';
    }
    
    const dadosConta = {
      descricao: motivo,
      competencia: competencia,
      valor: this.normalizarValor(valor),
      dataPagamento: this.formatarDataSemHora(pagoEm),
      dre: dre
    };

    const chaveComparacao = this.criarChaveSaida(dadosConta.descricao, dadosConta.dataPagamento, dadosConta.valor);
    this.logger.debug(`Verificando conta: ${chaveComparacao}`);

    if (this.saidasMap.has(chaveComparacao)) {
      const linhaExistente = this.saidasMap.get(chaveComparacao);
      this.logger.debug(`Encontrada duplicata na linha ${linhaExistente} da planilha de saídas`);
      this.marcarLinhaDuplicada(rowIndex, linhaExistente);
      return 'duplicada';
    } else {
      this.logger.debug(`Nova conta encontrada`);
      this.adicionarNovaSaida(dadosConta, rowIndex);
      
      // Adicionar todas as versões da chave ao mapa para evitar duplicatas futuras
      this.adicionarSaidaExistente(dadosConta.descricao, dadosConta.dataPagamento, dadosConta.valor, rowIndex);
      return 'nova';
    }
  }

  marcarLinhaDuplicada(rowIndex, linhaExistente) {
    this.sheetContas.getRange(rowIndex, 1, 1, this.sheetContas.getLastColumn()).setBackground(CONFIG.COR_LARANJA);
    const dadosConta = this.obterDadosLinha(rowIndex);
    this.adicionarAoRelatorio(dadosConta, `já estava lançada (linha ${linhaExistente})`);
  }

  adicionarNovaSaida(dadosConta, rowIndex) {
    this.novasLinhasSaidas.push(this.criarNovaSaida(dadosConta));
    this.sheetContas.getRange(rowIndex, 1, 1, this.sheetContas.getLastColumn()).setBackground(CONFIG.COR_CINZA);
    this.adicionarAoRelatorio(dadosConta, 'foi lançada');
  }

  criarNovaSaida({ dataPagamento, descricao, valor, competencia, dre }) {
    return [
      '', new Date(dataPagamento), '', '', '', '', '', 
      descricao, valor, competencia, 
      '', '', dre, '', ''
    ];
  }

  adicionarAoRelatorio(dados, status) {
    this.relatorioLinhas.push([
      new Date(), // timestamp atual
      dados.descricao,
      dados.dataPagamento,
      this.formatarValor(dados.valor),
      status,
      dados.dre
    ]);
  }

  inserirNovasSaidas() {
    if (this.novasLinhasSaidas.length > 0) {
      this.logger.log(`Inserindo ${this.novasLinhasSaidas.length} novas saídas...`);
      const ultimaLinha = this.sheetSaidas.getLastRow();
      this.sheetSaidas.insertRowsAfter(ultimaLinha, this.novasLinhasSaidas.length);
      this.sheetSaidas.getRange(ultimaLinha + 1, 1, this.novasLinhasSaidas.length, this.novasLinhasSaidas[0].length).setValues(this.novasLinhasSaidas);
      this.logger.log(`Inseridas ${this.novasLinhasSaidas.length} novas saídas.`);
    } else {
      this.logger.log('Nenhuma nova saída para inserir.');
    }
  }
  
  gerarRelatorio() {
    if (this.relatorioLinhas.length > 0) {
      this.logger.log(`Gerando relatório com ${this.relatorioLinhas.length} linhas...`);
      const ultimaLinha = this.sheetRelatorio.getLastRow();
      this.sheetRelatorio.getRange(ultimaLinha + 1, 1, this.relatorioLinhas.length, this.relatorioLinhas[0].length).setValues(this.relatorioLinhas);
      
      // Aplicar formatação
      const dataRange = this.sheetRelatorio.getRange(ultimaLinha + 1, 1, this.relatorioLinhas.length, 1);
      dataRange.setNumberFormat("dd/MM/yyyy HH:mm:ss");
      
      this.sheetRelatorio.getRange(ultimaLinha + this.relatorioLinhas.length + 2, 1).setValue(`Resumo do processamento em ${new Date().toLocaleString()}:`);
      this.sheetRelatorio.getRange(ultimaLinha + this.relatorioLinhas.length + 3, 1).setValue(`Total de contas processadas: ${this.estatisticas.processadas}`);
      this.sheetRelatorio.getRange(ultimaLinha + this.relatorioLinhas.length + 4, 1).setValue(`Total de duplicatas encontradas: ${this.estatisticas.duplicadas}`);
      this.sheetRelatorio.getRange(ultimaLinha + this.relatorioLinhas.length + 5, 1).setValue(`Total de novas saídas: ${this.estatisticas.novas}`);
      
      this.logger.log('Relatório gerado com sucesso.');
    }
  }

  gerarCSV() {
    this.logger.log('Gerando arquivo CSV...');
    
    const csvData = [['TIMESTAMP', 'DESCRIÇÃO', 'DATA_PAGAMENTO', 'VALOR', 'STATUS', 'DRE']];
    this.relatorioLinhas.forEach(linha => {
      csvData.push([
        Utilities.formatDate(linha[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
        linha[1], linha[2], linha[3], linha[4], linha[5]
      ]);
    });
    
    const csvContent = csvData.map(linha => 
      linha.map(campo => this.formatarCampoCSV(campo)).join(',')
    ).join('\r\n');

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const nomeArquivo = `${this.obterDataFormatada()}-despesas.csv`;
    folder.createFile(nomeArquivo, csvContent, MimeType.CSV);
    this.logger.log(`Relatório CSV gerado: ${nomeArquivo}`);
  }

  formatarCampoCSV(campo) {
    const valor = campo.toString();
    return valor.includes(',') || valor.includes('"') 
      ? `"${valor.replace(/"/g, '""')}"`
      : valor;
  }

  obterDataFormatada() {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH-mm');
  }

  criarChaveSaida(descricao, data, valor) {
    return `${this.normalizarTexto(descricao)}|${this.formatarDataSemHora(data)}|${this.normalizarValor(valor).toFixed(2)}`;
  }

  normalizarValor(valor) {
    if (typeof valor === 'string') {
      valor = valor.replace(/[^\d.,]/g, '').replace(',', '.');
    }
    return Math.round(parseFloat(valor) * 100) / 100 || 0;
  }

  normalizarTexto(texto) {
    if (!texto) return '';
    
    // Converte para string, lowercase e remove espaços extras
    let normalizado = texto.toString().toLowerCase().trim().replace(/\s+/g, ' ');
    
    // Remove acentos, cedilha e outros caracteres especiais
    normalizado = normalizado.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    
    // Substitui cedilha por 'c'
    normalizado = normalizado.replace(/ç/g, 'c');
    
    return normalizado;
  }

  formatarDataSemHora(data) {
    if (data instanceof Date) {
      return Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    const dataObj = new Date(data);
    return isNaN(dataObj.getTime()) ? '' : Utilities.formatDate(dataObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  formatarValor(valor) {
    return valor.toFixed(2).replace('.', ',');
  }

  obterDadosLinha(rowIndex) {
    const linha = this.sheetContas.getRange(rowIndex, 1, 1, this.sheetContas.getLastColumn()).getValues()[0];
    return {
      descricao: linha[2],
      dataPagamento: this.formatarDataSemHora(linha[7]),
      valor: this.normalizarValor(linha[4]),
      dre: linha[10]
    };
  }
}

class Logger {
  constructor() {
    this.logs = [];
    this.nivelLog = 'INFO'; // DEBUG, INFO, WARN, ERROR
    this.niveis = {
      'DEBUG': 0,
      'INFO': 1,
      'WARN': 2,
      'ERROR': 3
    };
    this.startTime = new Date();
  }

  debug(message) {
    if (this.niveis[this.nivelLog] <= this.niveis['DEBUG']) {
      this._log('DEBUG', message);
    }
  }

  log(message) {
    if (this.niveis[this.nivelLog] <= this.niveis['INFO']) {
      this._log('INFO', message);
    }
  }

  warn(message) {
    if (this.niveis[this.nivelLog] <= this.niveis['WARN']) {
      this._log('WARN', message);
    }
  }

  error(message, stack) {
    if (this.niveis[this.nivelLog] <= this.niveis['ERROR']) {
      this._log('ERROR', message);
      if (stack) {
        this._log('ERROR', `Stack: ${stack}`);
      }
    }
  }

  _log(level, message) {
    const timestamp = new Date();
    const elapsedMs = timestamp - this.startTime;
    const elapsedSec = Math.floor(elapsedMs / 1000);
    const formattedTime = `${elapsedSec}s ${elapsedMs % 1000}ms`;
    
    const logEntry = `[${level}] [${formattedTime}] ${timestamp.toISOString()}: ${message}`;
    console.log(logEntry);
    this.logs.push(logEntry);
  }

  getLogs() {
    return this.logs.join('\n');
  }
  
  saveToFile() {
    if (this.logs.length === 0) return;
    
    try {
      const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
      const now = new Date();
      const fileName = `log_${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.txt`;
      folder.createFile(fileName, this.getLogs(), MimeType.PLAIN_TEXT);
      console.log(`Log salvo em ${fileName}`);
    } catch (e) {
      console.error(`Erro ao salvar log: ${e.message}`);
    }
  }
}

function consolidarContasPagas() {
  const transferencia = new ContasPagasTransferencia();
  transferencia.executar();
  
  // Salvar logs em arquivo ao final da execução
  transferencia.logger.saveToFile();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Cria menu
  ui.createMenu('ocultado')
    .addItem('Nova Despesa', 'abrirFormulario')
    .addItem('Consolidar Contas Pagas', 'consolidarContasPagas')
    .addToUi();
} */
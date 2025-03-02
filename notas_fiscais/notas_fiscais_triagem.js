// arquivo: notas_fiscais_triagem.js
// Versão: 3.2
// Autor: Juliano Ceconi (Otimizado e Corrigido)

var CONFIG = {
  SHEET_NAME: 'Triagem',
  DUPLICADO_SHEET_NAME: 'duplicado_auto',
  ERROR_FOLDER_ID: 'ocultado',
  IMPORT_RANGE_ID: 'ocultado',
  DEFAULT_CITY: 'ocultado',
  DEFAULT_CEP: 'ocultado',
  BATCH_SIZE: 500, // Processa 500 linhas por vez
  MAX_RETRIES: 3
};

// Cache global para dados frequentemente acessados
var CACHE = {
  cepPorMunicipio: {},
  formasValidas: ['debito', 'credito', 'pix', 'cartao'],
  guiaSet: new Set()
};

// Sistema de log otimizado
var Logger = {
  logs: [],
  batchSize: 100,
  
  log: function(type, message, data) {
    data = data || {};
    this.logs.push({
      timestamp: new Date().toISOString(),
      type: type,
      message: message,
      data: JSON.stringify(data)
    });
    
    // Log no console do Apps Script (apenas mensagens importantes)
    if (type === 'error' || type === 'warning') {
      console.log(type + ': ' + message + ' - ' + JSON.stringify(data));
    }
    
    // Salva logs em lotes para reduzir chamadas de API
    if (this.logs.length >= this.batchSize) {
      this.saveErrors();
      this.logs = [];
    }
  },
  
  saveErrors: function() {
    if (this.logs.length === 0) return;
    
    try {
      var folder = DriveApp.getFolderById(CONFIG.ERROR_FOLDER_ID);
      var fileName = 'erros_triagem_' + new Date().toISOString().split('T')[0] + '.json';
      // Corrigido: Adicionado verificação de MimeType válido
      folder.createFile(fileName, JSON.stringify(this.logs), 'application/json');
    } catch (e) {
      console.error('Erro ao salvar logs: ' + e.message);
    }
  }
};

// Função principal para executar a triagem
function executarTriagem() {
  console.log('INICIANDO EXECUÇÃO DA TRIAGEM: ' + new Date().toISOString());
  var startTime = new Date();
  
  try {
    // Otimização 1: Carregar planilha uma única vez
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      console.error('Aba Triagem não encontrada!');
      return;
    }
    
    console.log('Carregando CEPs...');
    carregarCEPs(ss);
    
    // Otimização 2: Carregar todos os dados de uma vez
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    console.log('Total de linhas: ' + rows.length);
    
    var duplicados = [];
    var totalProcessado = 0;
    var batchUpdates = [];
    var batchFormattings = [];
    
    // Otimização 3: Uso de array temporário para armazenar atualizações
    for (var i = 0; i < rows.length; i += CONFIG.BATCH_SIZE) {
      var loteAtual = Math.floor(i / CONFIG.BATCH_SIZE) + 1;
      var totalLotes = Math.ceil(rows.length / CONFIG.BATCH_SIZE);
      
      console.log('Processando lote ' + loteAtual + ' de ' + totalLotes);
      
      var batch = rows.slice(i, Math.min(i + CONFIG.BATCH_SIZE, rows.length));
      var result = processarLote(batch, i + 2, headers.length);
      
      // Otimização 4: Consolidar atualizações
      batchUpdates = batchUpdates.concat(result.updates);
      batchFormattings = batchFormattings.concat(result.formattings);
      duplicados = duplicados.concat(result.duplicados);
      
      totalProcessado += batch.length;
      var porcentagem = Math.round((totalProcessado / rows.length) * 100);
      console.log('Progresso total: ' + porcentagem + '%');
      
      // Otimização 5: Flush periódico
      if (totalProcessado % 2000 === 0 || loteAtual === totalLotes) {
        // Aplica atualizações em lote
        aplicarAtualizacoesEmLote(sheet, batchUpdates);
        
        // Corrigido: Verificar se há formatações antes de aplicar
        if (batchFormattings.length > 0) {
          aplicarFormatacaoEmLote(sheet, batchFormattings);
        }
        
        batchUpdates = [];
        batchFormattings = [];
        
        SpreadsheetApp.flush();
        Utilities.sleep(500);
      }
    }
    
    // Aplica as atualizações restantes
    if (batchUpdates.length > 0) {
      aplicarAtualizacoesEmLote(sheet, batchUpdates);
    }
    
    // Corrigido: Verificar se há formatações antes de aplicar
    if (batchFormattings.length > 0) {
      aplicarFormatacaoEmLote(sheet, batchFormattings);
    }
    
    if (duplicados.length > 0) {
      console.log('Processando duplicados: ' + duplicados.length);
      inserirDuplicadosNaAba(ss, duplicados);
    }
    
    console.log('Preenchendo coluna Tipo Logradouro...');
    preencherColunaTipoLogradouro(sheet);
    
    var endTime = new Date();
    var tempoTotal = (endTime - startTime) / 1000;
    
    console.log('EXECUÇÃO CONCLUÍDA');
    console.log('Tempo total: ' + tempoTotal + ' segundos');
    console.log('Linhas processadas: ' + rows.length);
    console.log('Duplicados encontrados: ' + duplicados.length);
    console.log('Média de processamento: ' + Math.round(rows.length / tempoTotal) + ' linhas/segundo');

  } catch (error) {
    console.error('ERRO NA EXECUÇÃO: ' + error.message);
    console.error('Tempo decorrido: ' + ((new Date() - startTime) / 1000) + ' segundos');
    console.error('Stack: ' + error.stack);
    
    // Otimização 6: Salvamento de logs de erros
    Logger.log('error', 'Erro fatal na execução', {
      message: error.message,
      stack: error.stack
    });
    
    try {
      Logger.saveErrors();
    } catch (e) {
      console.error('Erro ao salvar logs: ' + e.message);
    }
    
    throw error;
  }
}

/**
 * Carrega todos os CEPs no início do processamento
 */
function carregarCEPs(ss) {
  var cepSheet = ss.getSheetByName('CEP');
  if (!cepSheet) return;
  
  // Otimização 7: Leitura única com getValues()
  var cepData = cepSheet.getDataRange().getValues();
  for (var i = 1; i < cepData.length; i++) {
    var municipio = cepData[i][0].toString().trim().toLowerCase();
    var cep = cepData[i][1].toString().replace(/\D/g, '');
    CACHE.cepPorMunicipio[municipio] = cep;
  }
}

/**
 * Processa um lote de linhas e retorna atualizações
 */
function processarLote(rows, startIndex, numCols) {
  var updates = [];
  var formattings = [];
  var duplicados = [];
  
  // Otimização 8: Processamento em massa
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i].slice(); // Clone para evitar modificação do original
    var rowIndex = startIndex + i;
    
    // Verifica duplicatas
    var numeroGuia = row[2];
    if (numeroGuia && CACHE.guiaSet.has(numeroGuia.toString())) {
      duplicados.push(rows[i]);
      continue;
    }
    
    if (numeroGuia) CACHE.guiaSet.add(numeroGuia.toString());
    
    // Processamento de linha
    try {
      // Corrigido: Garantir que os dados sejam strings antes de chamar trim()
      // Validações e formatações
      row[11] = toString(row[11]).trim() || CONFIG.DEFAULT_CITY; // Cidade
      row[6] = validarCEP(row[6], row[11]); // CEP
      row[4] = validarCPF(row[4]); // CPF
      row[5] = formatarTelefone(row[5]); // Telefone
      row[2] = formatarNumeroGuia(row[2]); // Número da Guia
      row[13] = validarEmail(row[13]); // Email
      
      // Processamento do procedimento
      if (row[14] && toString(row[14]).toLowerCase().indexOf('laboratório') !== -1) {
        row[17] = 'Agendamento para Exames Laboratoriais';
      } else if (row[17]) {
        var procString = toString(row[17]).trim();
        if (procString.toLowerCase().indexOf('agendamento para ') !== 0) {
          row[17] = 'Agendamento para ' + procString;
        }
      }
      
      // Adiciona à lista de atualizações
      updates.push({
        row: rowIndex,
        values: row
      });
      
      // Verifica condições para formatação
      if (row[4] === 'SEM CPF' || (row[15] !== null && row[15] !== undefined && row[15] < 0) || !validarFormaPagamento(row[0])) {
        formattings.push({
          row: rowIndex,
          color: '#cc4125'
        });
      } else if (row[14] && toString(row[14]).toLowerCase().indexOf('ocultado') !== -1) {
        formattings.push({
          row: rowIndex,
          color: 'yellow'
        });
      } else if (row[15] !== null && row[15] !== undefined && row[15] <= 0) {
        formattings.push({
          row: rowIndex,
          color: 'orange'
        });
      }
      
    } catch (error) {
      Logger.log('error', 'Erro ao processar linha', {
        rowIndex: rowIndex,
        error: error.message
      });
      // Mantém a linha original em caso de erro
      updates.push({
        row: rowIndex,
        values: rows[i]
      });
    }
  }
  
  return {
    updates: updates,
    formattings: formattings,
    duplicados: duplicados
  };
}

/**
 * Função auxiliar para converter com segurança para string
 * Evita erros com valores nulos ou undefined
 */
function toString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value);
}

/**
 * Aplica atualizações em lote na planilha
 */
function aplicarAtualizacoesEmLote(sheet, updates) {
  if (updates.length === 0) return;
  
  // Otimização 9: Agrupa atualizações por linhas consecutivas
  var currentRow = updates[0].row;
  var currentBatch = [updates[0].values];
  var batches = [];
  var rows = [currentRow];
  
  for (var i = 1; i < updates.length; i++) {
    if (updates[i].row === currentRow + 1) {
      // Linha consecutiva
      currentBatch.push(updates[i].values);
      currentRow = updates[i].row;
    } else {
      // Nova sequência
      batches.push(currentBatch);
      rows.push(updates[i].row);
      currentBatch = [updates[i].values];
      currentRow = updates[i].row;
    }
  }
  batches.push(currentBatch);
  
  // Atualiza em grupos
  for (var i = 0; i < batches.length; i++) {
    if (batches[i].length > 0) {
      var startRow = rows[i];
      var numRows = batches[i].length;
      var numCols = batches[i][0].length;
      
      // Corrigido: Verifica se o range é válido antes de atualizar
      if (numRows > 0 && numCols > 0) {
        sheet.getRange(startRow, 1, numRows, numCols).setValues(batches[i]);
        
        // Adiciona fórmulas para a coluna 16 (Valor)
        var formulas = [];
        for (var j = 0; j < numRows; j++) {
          formulas.push(['=XLOOKUP(C' + (startRow + j) + '; IMPORTRANGE("' + CONFIG.IMPORT_RANGE_ID + 
                         '"; "guias!A:A"); IMPORTRANGE("' + CONFIG.IMPORT_RANGE_ID + '"; "guias!L:L"))']);
        }
        sheet.getRange(startRow, 16, numRows, 1).setFormulas(formulas);
      }
    }
  }
}

/**
 * Aplica formatações em lote - CORRIGIDO
 */
function aplicarFormatacaoEmLote(sheet, formattings) {
  if (formattings.length === 0) return;
  
  // Otimização 10: Agrupar formatações por cor e aplicar em blocos
  var formatByColor = {};
  formattings.forEach(function(format) {
    if (!formatByColor[format.color]) {
      formatByColor[format.color] = [];
    }
    formatByColor[format.color].push(format.row);
  });
  
  // Aplicar formatações por grupo de cor
  for (var color in formatByColor) {
    if (formatByColor[color].length > 0) {
      // Corrigido: Evita uso de getRangeList que estava causando erro
      // Agrupa linhas consecutivas para minimizar o número de chamadas de API
      var rows = formatByColor[color].sort(function(a, b) { return a - b; });
      var ranges = [];
      var startRow = rows[0];
      var currentRow = startRow;
      var rowCount = 1;
      
      for (var i = 1; i < rows.length; i++) {
        if (rows[i] === currentRow + 1) {
          // Linha consecutiva
          rowCount++;
          currentRow = rows[i];
        } else {
          // Não consecutiva, salva o range atual e inicia um novo
          ranges.push({startRow: startRow, rowCount: rowCount});
          startRow = rows[i];
          currentRow = startRow;
          rowCount = 1;
        }
      }
      
      // Adiciona o último range
      ranges.push({startRow: startRow, rowCount: rowCount});
      
      // Aplica formatação para cada range
      var lastCol = sheet.getLastColumn();
      for (var i = 0; i < ranges.length; i++) {
        try {
          if (ranges[i].rowCount > 0 && lastCol > 0) {
            sheet.getRange(ranges[i].startRow, 1, ranges[i].rowCount, lastCol).setBackground(color);
          }
        } catch (e) {
          console.error('Erro ao aplicar formatação: ' + e.message);
          // Tenta aplicar linha por linha se ocorrer erro
          for (var j = 0; j < ranges[i].rowCount; j++) {
            try {
              sheet.getRange(ranges[i].startRow + j, 1, 1, lastCol).setBackground(color);
            } catch (e2) {
              console.error('Falha ao formatar linha ' + (ranges[i].startRow + j) + ': ' + e2.message);
            }
          }
        }
      }
    }
  }
}

/**
 * Valida o CPF usando o algoritmo oficial e mantém a formatação
 */
function validarCPF(cpf) {
  if (!cpf || toString(cpf).trim() === '') {
    return 'SEM CPF';
  }
  
  // Armazena o formato original para usar caso o CPF seja válido
  const cpfOriginal = toString(cpf).trim();
  const cpfLimpo = toString(cpf).replace(/\D/g, '');
  
  // Verifica se tem 11 dígitos
  if (cpfLimpo.length !== 11) {
    return 'SEM CPF';
  }
  
  // Verifica CPFs com dígitos repetidos (ex: 11111111111)
  if (/^(\d)\1{10}$/.test(cpfLimpo)) {
    return 'SEM CPF';
  }
  
  // Cálculo do primeiro dígito verificador
  let soma = 0;
  for (let i = 0; i < 9; i++) {
    soma += parseInt(cpfLimpo.charAt(i)) * (10 - i);
  }
  
  let resto = soma % 11;
  let digitoVerificador1 = resto < 2 ? 0 : 11 - resto;
  
  // Verifica o primeiro dígito
  if (digitoVerificador1 !== parseInt(cpfLimpo.charAt(9))) {
    return 'SEM CPF';
  }
  
  // Cálculo do segundo dígito verificador
  soma = 0;
  for (let i = 0; i < 10; i++) {
    soma += parseInt(cpfLimpo.charAt(i)) * (11 - i);
  }
  
  resto = soma % 11;
  let digitoVerificador2 = resto < 2 ? 0 : 11 - resto;
  
  // Verifica o segundo dígito
  if (digitoVerificador2 !== parseInt(cpfLimpo.charAt(10))) {
    return 'SEM CPF';
  }
  
  // Se passou por todas as validações, retorna o CPF no formato original
  // Isso manterá a formatação com pontos e traços como em 000.054.275-09
  return cpfOriginal;
}

/**
 * Valida o CEP e busca na tabela se necessário
 */
function validarCEP(cep, cidade) {
  try {
    if (!cep) {
      // Otimização 12: Uso do cache para CEPs
      const cidadeNormalizada = toString(cidade).trim().toLowerCase();
      return CACHE.cepPorMunicipio[cidadeNormalizada] || CONFIG.DEFAULT_CEP;
    }
    
    let cepLimpo = toString(cep).replace(/\D/g, '');
    
    // Validações básicas
    if (cepLimpo.length !== 8 || !/^[0-9]{8}$/.test(cepLimpo)) {
      const cidadeNormalizada = toString(cidade).trim().toLowerCase();
      return CACHE.cepPorMunicipio[cidadeNormalizada] || CONFIG.DEFAULT_CEP;
    }
    
    return cepLimpo;
  } catch (error) {
    return CONFIG.DEFAULT_CEP;
  }
}

/**
 * Valida e formata um email
 */
function validarEmail(email) {
  if (!email) return '';
  
  // Otimização 13: Validação simplificada de email
  const emailLimpo = toString(email).trim().toLowerCase();
  if (emailLimpo.indexOf('@') === -1 || emailLimpo.split('@')[1].indexOf('.') === -1) {
    return '';
  }
  
  return emailLimpo;
}

/**
 * Normaliza uma string removendo acentos, espaços e caracteres especiais
 */
function normalizarString(str) {
  if (!str) return '';
  
  // Otimização 14: Normalização simplificada
  return toString(str)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Verifica se a forma de pagamento é válida
 */
function validarFormaPagamento(formaPgto) {
  if (!formaPgto) return false;
  
  // Otimização 15: Simplificação da validação
  var pgtoNormalizado = normalizarString(formaPgto);
  
  for (var i = 0; i < CACHE.formasValidas.length; i++) {
    if (pgtoNormalizado.indexOf(CACHE.formasValidas[i]) !== -1) {
      return true;
    }
  }
  
  return false;
}

/**
 * Formata o número de telefone
 */
function formatarTelefone(telefone) {
  if (!telefone) return '';
  
  // Otimização 16: Simplificação da limpeza de telefone
  let telefoneLimpo = toString(telefone).replace(/\D/g, '');
  
  if (telefoneLimpo.startsWith('55') && telefoneLimpo.length > 2) {
    telefoneLimpo = telefoneLimpo.substring(2);
  }
  
  if (telefoneLimpo.length < 10 || telefoneLimpo.length > 11) {
    return toString(telefone);
  }
  
  return telefoneLimpo;
}

/**
 * Formata o número da guia
 */
function formatarNumeroGuia(numeroGuia) {
  if (!numeroGuia) return '';
  
  // Otimização 17: Simplificação do formato da guia
  const guiaLimpa = toString(numeroGuia).replace(/\D/g, '');
  return guiaLimpa.length >= 3 && guiaLimpa.length <= 7 ? guiaLimpa : toString(numeroGuia);
}

/**
 * Insere registros duplicados na aba 'duplicado_auto'
 */
function inserirDuplicadosNaAba(ss, duplicados) {
  if (duplicados.length === 0) return;

  // Otimização 18: Reutilizar a spreadsheet existente
  let duplicadoSheet = ss.getSheetByName(CONFIG.DUPLICADO_SHEET_NAME);
  if (!duplicadoSheet) {
    duplicadoSheet = ss.insertSheet(CONFIG.DUPLICADO_SHEET_NAME);
    var headers = ss.getSheetByName(CONFIG.SHEET_NAME).getRange(1, 1, 1, duplicados[0].length).getValues()[0];
    duplicadoSheet.appendRow(headers);
  }
  
  // Insere todos os duplicados de uma vez
  duplicadoSheet.getRange(duplicadoSheet.getLastRow() + 1, 1, duplicados.length, duplicados[0].length).setValues(duplicados);
}

/**
 * Preenche a coluna 'Tipo Logradouro' com 'Rua' nos campos vazios
 */
function preencherColunaTipoLogradouro(sheet) {
  // Otimização 19: Usar getValues e setValues para otimizar
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // Verifica se há dados além do cabeçalho
  
  const tipoLogradouroRange = sheet.getRange(1, 19, lastRow);
  const tipoLogradouroValues = tipoLogradouroRange.getValues();
  
  var alterado = false;
  for (let i = 1; i < tipoLogradouroValues.length; i++) {
    if (!tipoLogradouroValues[i][0]) {
      tipoLogradouroValues[i][0] = 'Rua';
      alterado = true;
    }
  }
  
  // Só atualiza se houver alterações
  if (alterado) {
    tipoLogradouroRange.setValues(tipoLogradouroValues);
  }
}

// Criação do menu na abertura da planilha
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('ocultado')
      .addItem('Executar Triagem', 'executarTriagem')
      .addToUi();
  } catch (e) {
    console.error('Erro ao criar menu: ' + e.message);
  }
}
// Versão 1.1
// Autor: Juliano Ceconi

function transferirContasPagas() {
  // IDs e nomes das planilhas e guias
  var SPREADSHEET_CONTAS_A_PAGAR_ID = 'ocultado';
  var SHEET_CONTAS_A_PAGAR_NOME = 'contas_a_pagar';

  var SPREADSHEET_SAIDAS_ID = 'ocultado';
  var SHEET_SAIDAS_NOME = 'saidas';

  // Pasta de destino para o CSV
  var FOLDER_ID = 'ocultado';

  // Abre as planilhas
  var ssContas = SpreadsheetApp.openById(SPREADSHEET_CONTAS_A_PAGAR_ID);
  var sheetContas = ssContas.getSheetByName(SHEET_CONTAS_A_PAGAR_NOME);

  var ssSaidas = SpreadsheetApp.openById(SPREADSHEET_SAIDAS_ID);
  var sheetSaidas = ssSaidas.getSheetByName(SHEET_SAIDAS_NOME);

  // Obter todos os dados das guias
  var rangeContas = sheetContas.getDataRange();
  var dadosContas = rangeContas.getValues();
  // Títulos da 'contas_a_pagar' (linha 2)
  // Indíces: OK(A=0), BENEFICIADO(B=1), MOTIVO(C=2), COMPETÊNCIA(D=3), VALOR(E=4), VENCIMENTO(F=5), VENCE EM(G=6), PAGO EM(H=7), CÓDIGO(I=8)
  
  // Títulos da 'saidas' (linha 2)
  // 'Cód.(A)	DATA PAGAMENTO(B)	RESPONSÁVEL(C)	CIDADE(D)	FORMA DE PAGAMENTO(E)	SETOR(F)	GRUPO DE DESPESA(G)	DESCRIÇÃO DA DESPESA(H)	VALOR(I)	COMPETÊNCIA(J)	Filtro(K)	DATA COMPROVANTE(L)	DRE(M)	FIXO(N)	FALSE(O)'

  // Obter todos os dados de 'saidas'
  var rangeSaidas = sheetSaidas.getDataRange();
  var dadosSaidas = rangeSaidas.getValues();

  // Criar um mapa para facilitar a comparação
  // Como chave, usaremos a combinação (DESCRIÇÃO + DATA + VALOR) já existentes em 'saidas'.
  // Mas precisamos padronizar: 
  // - DATA: considerar apenas a parte da data (sem hora)
  // - VALOR: arredondar para centavos
  var saidasMap = {};
  for (var i = 2; i < dadosSaidas.length; i++) {
    var linhaSaida = dadosSaidas[i];
    var descricaoSaida = (linhaSaida[7] || '').toString().trim(); // DESCRIÇÃO DA DESPESA(H)
    var dataPagamentoSaida = linhaSaida[1]; // DATA PAGAMENTO(B)
    var valorSaida = normalizarValor(linhaSaida[8]); // VALOR(I)
    var dataFormatadaSaida = formatarDataSemHora(dataPagamentoSaida);
    var chave = criarChave(descricaoSaida, dataFormatadaSaida, valorSaida);
    saidasMap[chave] = true;
  }

  // Vamos criar um array para armazenar as atualizações que serão realizadas em 'saidas'
  var novasLinhasSaidas = [];

  // Armazena informações para o CSV: [Descrição, Data Pagamento, Valor, Status ("já estava lançada" ou "foi lançada")]
  var relatorioCSV = [];
  relatorioCSV.push(['DESCRIÇÃO', 'DATA_PAGAMENTO', 'VALOR', 'STATUS']);

  // Cores
  var corLaranja = '#FFA500'; // duplicada
  var corCinza = '#D3D3D3'; // lançada
  // Vamos verificar as contas pagas
  // Começar a ler a partir da linha 3, já que linha 1 e 2 são títulos
  for (var j = 2; j < dadosContas.length; j++) {
    var linha = dadosContas[j];
    var pagoEm = linha[7]; // PAGO EM (H)
    if (pagoEm && !(pagoEm.toString().trim() === '')) {
      // Conta está paga, verificar se já está em 'saidas'
      var motivo = (linha[2] || '').toString().trim();       // MOTIVO (C)
      var competencia = (linha[3] || '').toString().trim(); // COMPETÊNCIA (D)
      var valor = normalizarValor(linha[4]);                 // VALOR (E)
      var dataPagamento = pagoEm;                            // PAGO EM(H)
      var descricao = motivo; // DESCRIÇÃO DA DESPESA será o motivo
      var dataFormatada = formatarDataSemHora(dataPagamento);

      // Cria chave para comparação
      var chaveComparacao = criarChave(descricao, dataFormatada, valor);

      if (saidasMap[chaveComparacao]) {
        // Já existe em 'saidas'
        // Pintar linha de laranja em 'contas_a_pagar'
        var rngLinha = sheetContas.getRange(j+1, 1, 1, dadosContas[0].length);
        rngLinha.setBackground(corLaranja);

        relatorioCSV.push([descricao, dataFormatada, valor.toString().replace('.',','), 'já estava lançada']);
      } else {
        // Não existe, lançar em 'saidas'
        // Mapeamento: 
        // contaspagar(C->H descr), (D->J comp), (E->I valor), (H->B data pag)
        // Precisamos inserir no final de 'saidas'
        // Campos obrigatórios em 'saidas':
        // Cód.(A) - deixar em branco
        // DATA PAGAMENTO(B) = pagoEm
        // RESPONSÁVEL(C) - deixar em branco
        // CIDADE(D) - deixar em branco
        // FORMA DE PAGAMENTO(E) - deixar em branco
        // SETOR(F) - deixar em branco
        // GRUPO DE DESPESA(G) - deixar em branco
        // DESCRIÇÃO DA DESPESA(H) = motivo
        // VALOR(I) = valor
        // COMPETÊNCIA(J) = competencia
        // Demais colunas podem ficar em branco ou conforme necessidade.

        novasLinhasSaidas.push([
          '',                 // Cód.
          new Date(dataFormatada), // DATA PAGAMENTO - Data normalizada
          '',                 // RESPONSÁVEL
          '',                 // CIDADE
          '',                 // FORMA DE PAGAMENTO
          '',                 // SETOR
          '',                 // GRUPO DE DESPESA
          descricao,          // DESCRIÇÃO DA DESPESA
          valor,              // VALOR
          competencia,        // COMPETÊNCIA
          '',                 // Filtro
          '',                 // DATA COMPROVANTE
          '',                 // DRE
          '',                 // FIXO
          ''                  // FALSE
        ]);

        // Pintar linha de cinza em 'contas_a_pagar'
        var rngLinha2 = sheetContas.getRange(j+1, 1, 1, dadosContas[0].length);
        rngLinha2.setBackground(corCinza);

        relatorioCSV.push([descricao, dataFormatada, valor.toString().replace('.',','), 'foi lançada']);
      }
    }
  }

  // Inserir novas linhas em 'saidas', se houver
  if (novasLinhasSaidas.length > 0) {
    var ultimaLinhaSaidas = sheetSaidas.getLastRow();
    sheetSaidas.insertRowsAfter(ultimaLinhaSaidas, novasLinhasSaidas.length);
    sheetSaidas.getRange(ultimaLinhaSaidas+1, 1, novasLinhasSaidas.length, novasLinhasSaidas[0].length).setValues(novasLinhasSaidas);
  }

  // Gerar CSV do relatório
  gerarCSV(relatorioCSV, FOLDER_ID);
}

/**
 * Normaliza o valor para comparação, arredondando para duas casas decimais.
 */
function normalizarValor(valor) {
  if (typeof valor === 'string') {
    valor = valor.replace(',', '.');
    valor = parseFloat(valor);
  }
  if (isNaN(valor)) {
    valor = 0;
  }
  return Math.round(valor * 100) / 100;
}

/**
 * Formata a data para ignorar hora, retornando apenas a parte da data (yyyy-MM-dd).
 */
function formatarDataSemHora(data) {
  if (!(data instanceof Date)) {
    // tenta converter
    data = new Date(data);
    if (isNaN(data.getTime())) {
      // Caso não seja uma data válida, retorna string vazia
      return '';
    }
  }
  var ano = data.getFullYear();
  var mes = ('0' + (data.getMonth()+1)).slice(-2);
  var dia = ('0' + data.getDate()).slice(-2);
  return ano + '-' + mes + '-' + dia;
}

/**
 * Cria uma chave única para comparação.
 */
function criarChave(descricao, data, valor) {
  // Remover espaços extras da descrição
  descricao = descricao.toLowerCase().trim();
  return descricao + '|' + data + '|' + valor;
}

/**
 * Gera um arquivo CSV na pasta especificada.
 */
function gerarCSV(dados, folderId) {
  var csvContent = [];
  for (var i = 0; i < dados.length; i++) {
    var linha = dados[i].map(function(campo) {
      // Se houver vírgula, aspas, etc., encapsular em aspas
      var val = campo.toString();
      if (val.indexOf(',') > -1 || val.indexOf('"') > -1) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      return val;
    }).join(',');
    csvContent.push(linha);
  }

  var csvString = csvContent.join('\r\n');

  var folder = DriveApp.getFolderById(folderId);
  var agora = new Date();
  var ano = agora.getFullYear();
  var mes = ('0' + (agora.getMonth()+1)).slice(-2);
  var dia = ('0' + agora.getDate()).slice(-2);
  var hora = ('0' + agora.getHours()).slice(-2);
  var minuto = ('0' + agora.getMinutes()).slice(-2);

  var nomeArquivo = ano + '-' + mes + '-' + dia + '-' + hora + '-' + minuto + '-despesas.csv';
  
  folder.createFile(nomeArquivo, csvString, MimeType.CSV);
}

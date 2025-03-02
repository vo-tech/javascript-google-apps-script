// arquivo: forms_despesas_js.gs
// versão: 1.2
// autor: Juliano Ceconi

function abrirFormulario() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('formulario')
      .setWidth(600)
      .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Formulário de Adição');
}

// Função para gerar automaticamente o código baseado no filtro e registrar logs no Apps Script
function gerarCodigo(filtro) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("saidas");
  var data = sheet.getDataRange().getValues();
  var ultimoCodigo = 0;

  // Verifica o último código gerado com base no filtro (coluna 0 é o código, coluna 10 é o filtro)
  for (var i = 1; i < data.length; i++) {
    var codigoAtual = data[i][0]; // Pega o valor da coluna de código (coluna 0)
    var filtroAtual = data[i][10]; // Pega o valor do filtro (coluna 10)

    // Verifica se o filtro bate
    if (filtroAtual === filtro) {
      var numeroCodigo = parseInt(codigoAtual.substring(1), 10);

      if (!isNaN(numeroCodigo)) {
        if (numeroCodigo > ultimoCodigo) {
          ultimoCodigo = numeroCodigo;
        }
      }
    }
  }

  // Registra o último código encontrado nos logs do Apps Script
  Logger.log("Último código encontrado para o filtro " + filtro + ": " + ultimoCodigo);

  // Gera o próximo código com base no último número encontrado
  var novoCodigo = filtro.charAt(0).toUpperCase() + ('0000' + (ultimoCodigo + 1)).slice(-4);

  // Adiciona o novo código gerado nos logs do Apps Script
  Logger.log("Novo código gerado: " + novoCodigo);

  return novoCodigo;
}

// Função para salvar os dados no Google Sheets
function salvarDados(dados) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("saidas");

  // Gera o código único automaticamente
  var codigo = gerarCodigo(dados.filtro);

  // Corrige o formato do valor para garantir 2 casas decimais
  var valorNumerico = parseFloat(dados.valor.replace(',', '.'));
  var valorCorrigido = valorNumerico.toFixed(2).replace('.', ',');

  // Corrige o formato da data para "d/m" sem zeros adicionais
  var dataPagamento = new Date(dados.dataPagamento);
  var dataFormatada = dataPagamento.getDate() + '/' + (dataPagamento.getMonth() + 1);

  // Adiciona os dados na próxima linha disponível
  sheet.appendRow([
    codigo,  // Insere o código gerado
    dataFormatada,  // Insere a data formatada
    dados.responsavel, 
    dados.cidade, 
    dados.formaPagamento, 
    dados.setor, 
    dados.grupoDespesa, 
    dados.descricaoDespesa, 
    valorCorrigido,  // Insere o valor corrigido
    dados.competencia, 
    dados.filtro
  ]);
}
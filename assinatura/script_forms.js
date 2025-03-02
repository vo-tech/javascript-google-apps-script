// Versão: 1.0
// Autor: Juliano Ceconi

// Função para abrir o formulário HTML 
function abrirFormulario() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('formulario')
        .setWidth(600) // Largura da janela do formulário
        .setHeight(800); // Altura da janela do formulário
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Cadastrar Assinante');
  }
  
  // Função para gerar automaticamente o código baseado na "Forma de Pagamento"
  function gerarCodigo(formaPagamento) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Assinantes");
    var data = sheet.getDataRange().getValues();
    var ultimoCodigo = 0;
  
    // Verifica o último código gerado com base na "Forma de Pagamento" (coluna 0 é o código, coluna 3 é "Forma de Pagamento")
    for (var i = 1; i < data.length; i++) {
      var codigoAtual = data[i][0]; // Pega o valor da coluna de código (coluna 0)
      var formaPagamentoAtual = data[i][3]; // Pega o valor da coluna "Forma de Pagamento" (coluna 3)
  
      // Verifica se a forma de pagamento bate
      if (formaPagamentoAtual === formaPagamento) {
        var numeroCodigo = parseInt(codigoAtual.substring(1), 10);
  
        if (!isNaN(numeroCodigo)) {
          if (numeroCodigo > ultimoCodigo) {
            ultimoCodigo = numeroCodigo;
          }
        }
      }
    }
  
    // Gera o próximo código com base no último número encontrado
    var prefixo = formaPagamento ? formaPagamento.charAt(0).toUpperCase() : 'X';
    var novoCodigo = prefixo + ('0000' + (ultimoCodigo + 1)).slice(-4);
  
    Logger.log("Novo código gerado: " + novoCodigo);
  
    return novoCodigo;
  }
  
  // Função para salvar os dados no Google Sheets
  function salvarDados(dados) {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ocultado");
  
      // Gera o código único automaticamente com base na "Forma de Pagamento"
      var codigo = gerarCodigo(dados.formaPagamento);
  
      // Encontra a primeira linha vazia logo após os dados existentes
      var ultimaLinha = sheet.getLastRow();
      var ultimaColuna = sheet.getLastColumn();
      var ultimaLinhaDados = ultimaLinha;
  
      // Confere se há linhas vazias no meio e as redefine
      for (var i = 2; i <= ultimaLinha; i++) {
        if (sheet.getRange(i, 1, 1, ultimaColuna).getValues()[0].every(cell => cell === "")) {
          ultimaLinhaDados = i - 1;
          break;
        }
      }
  
      var proximaLinha = ultimaLinhaDados + 1;
  
      // Adiciona os dados na próxima linha disponível
      sheet.getRange(proximaLinha, 1, 1, 10).setValues([[
        codigo,  // Insere o código gerado
        dados.titular || '',
        dados.status || 'Ativo',
        dados.formaPagamento || '',
        dados.valorTotal || '',
        dados.diaVencimento || '',
        dados.assinatura || '',
        dados.falta || '',
        dados.dataVenda || '',
        dados.vendedor || ''
      ]]);
  
      Logger.log('Dados salvos com sucesso: ' + JSON.stringify(dados));
    } catch (error) {
      Logger.log('Erro ao salvar os dados: ' + error.message);
      throw new Error('Erro ao salvar os dados.');
    }
  }
  
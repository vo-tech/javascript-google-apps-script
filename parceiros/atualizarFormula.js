// arquivo: atualizarFormula.gs
// Versão: 4.0
// Autor: Juliano Ceconi

function atualizarFormulasParceiros() {
    var planilhaAtual = SpreadsheetApp.getActiveSpreadsheet();
    var abaControle = planilhaAtual.getActiveSheet();
    
    var ultimaLinha = abaControle.getLastRow();
    if (ultimaLinha < 2) {
      Logger.log('A lista de parceiros está vazia.');
      return;
    }
    
    var dados = abaControle.getRange(2, 1, ultimaLinha - 1, 3).getValues();
    
    for (var i = 0; i < dados.length; i++) {
      var nomeParceiro = dados[i][0];
      var urlPlanilha = dados[i][1];
      var status = dados[i][2];
      
      var linhaAtual = i + 2;
      
      if (status) {
        Logger.log('Linha ' + linhaAtual + ': Já processado. Pulando.');
        continue;
      }
      
      if (!nomeParceiro || !urlPlanilha) {
        Logger.log('Linha ' + linhaAtual + ': Dados insuficientes. Pulando.');
        continue;
      }
      
      try {
        Logger.log('Nome do parceiro sendo procurado: [' + nomeParceiro + ']');
        var planilhaParceiro = SpreadsheetApp.openByUrl(urlPlanilha);
        Logger.log('Processando parceiro: ' + nomeParceiro);
        
        // Criar ou obter a aba de repasse
        var abaRepasse = criarOuObterAba(planilhaParceiro, "repasse");
        
        // Adicionar cabeçalho e formatar aba
        formatarAba(abaRepasse);
        
        // Inserir a fórmula exata com tratamento mais robusto
        var formulaExata = '=ARRAYFORMULA(QUERY({' +
          'TEXTO(IMPORTRANGE("ocultado", "guias!A2:A"), "@"),' +
          'IMPORTRANGE("ocultado", "guias!G2:G"),' +
          'TEXTO(IMPORTRANGE("ocultado", "guias!H2:H"), "d/m"),' +
          'TEXTO(IMPORTRANGE("ocultado", "guias!K2:K"), "@"),' +
          'TEXTO(IMPORTRANGE("ocultado", "guias!M2:M"), "@"),' +
          'TEXTO(IMPORTRANGE("ocultado", "guias!N2:N"), "@"),' +
          'TEXTO(IMPORTRANGE("ocultado", "guias!P2:P"), "d/m"),' +
          'TEXTO(IMPORTRANGE("ocultado", "guias!S2:S"), "@")' +
          '}, "SELECT * WHERE Col6 CONTAINS \'' + nomeParceiro + '\' ORDER BY Col2 ASC"))';
        
        // Log da fórmula para debug
        Logger.log('Fórmula a ser inserida: ' + formulaExata);
        
        // Inserir a fórmula
        abaRepasse.getRange('A2').setFormula(formulaExata);
        
        abaControle.getRange(linhaAtual, 3).setValue('OK');
        Logger.log('Linha ' + linhaAtual + ': Fórmula inserida.');
        
      } catch (e) {
        Logger.log('Linha ' + linhaAtual + ': Erro - ' + e.message);
        continue;
      }
    }
  }
  
  function criarOuObterAba(planilha, nomeAba) {
    var aba = planilha.getSheetByName(nomeAba);
    if (!aba) {
      aba = planilha.insertSheet(nomeAba);
    }
    return aba;
  }
  
  function formatarAba(aba) {
    // Adicionar cabeçalho atualizado
    var cabecalho = ['Guia', 'Emissão', 'Data Guia', 'Valor de repasse', 'Procedimento', 'Instituição', 'Data Repasse', 'Paciente'];
    aba.getRange('A1:H1').setValues([cabecalho]);
    
    // Formatar cabeçalho
    var cabecalhoRange = aba.getRange('A1:H1');
    cabecalhoRange.setBackground('#154734')
                  .setFontColor('white')
                  .setFontWeight('bold')
                  .setFontFamily('Google Sans');
  
    // Formatar corpo da tabela
    var corpoRange = aba.getRange('A2:H1000');
    corpoRange.setFontFamily('Google Sans')
              .setBorder(false, false, false, false, false, false)
              .setBackground(null);
  
    // Aplicar cores alternadas
    var regra = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([corpoRange])
      .whenFormulaSatisfied('=MOD(ROW(),2)=0')
      .setBackground('#E6F4EA')
      .build();
    var regras = aba.getConditionalFormatRules();
    regras.push(regra);
    aba.setConditionalFormatRules(regras);
  
    // Remover linhas de grade
    aba.setHiddenGridlines(true);
  }
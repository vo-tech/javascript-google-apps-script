function consolidarCodigosProcedimentos() {
    try {
      // Nome da planilha de destino onde os códigos únicos serão lançados
      const nomePlanilhaDestino = 'TODOS';
      
      // Acessa a planilha ativa
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Obter ou criar a planilha de destino
      const planilhaDestino = obterOuCriarPlanilha(ss, nomePlanilhaDestino);
      
      // Obter os códigos já existentes na planilha de destino para evitar duplicação
      const codigosExistentes = obterCodigosExistentes(planilhaDestino);
      
      // Percorre cada planilha e adiciona os novos códigos à planilha de destino
      const planilhas = ss.getSheets();
      planilhas.forEach(planilha => {
        if (planilha.getName() !== nomePlanilhaDestino) {
          adicionarNovosCodigos(planilha, planilhaDestino, codigosExistentes);
        }
      });
      
      Logger.log('Códigos consolidados com sucesso na planilha "TODOS CLIENTES".');
    } catch (error) {
      Logger.log('Erro ao consolidar códigos: ' + error.message);
    }
  }
  
  // Função para obter ou criar a planilha de destino
  function obterOuCriarPlanilha(ss, nomePlanilha) {
    let planilha = ss.getSheetByName(nomePlanilha);
    if (!planilha) {
      planilha = ss.insertSheet(nomePlanilha);
    }
    return planilha;
  }
  
  // Função para obter os códigos existentes na planilha de destino
  function obterCodigosExistentes(planilhaDestino) {
    const codigosExistentes = new Set();
    const ultimaLinha = planilhaDestino.getLastRow();
    if (ultimaLinha > 1) {
      const valoresExistentes = planilhaDestino.getRange('A2:A' + ultimaLinha).getValues();
      valoresExistentes.forEach(linha => {
        const codigo = linha[0].toString().trim();
        if (codigo) {
          codigosExistentes.add(codigo);
        }
      });
    }
    return codigosExistentes;
  }
  
  // Função para adicionar novos códigos à planilha de destino
  function adicionarNovosCodigos(planilhaOrigem, planilhaDestino, codigosExistentes) {
    const ultimaLinha = planilhaOrigem.getLastRow();
    const ultimaColuna = planilhaOrigem.getLastColumn();
    if (ultimaLinha > 1 && ultimaColuna > 0) {
      const dados = planilhaOrigem.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
      dados.forEach(linha => {
        const codigo = linha[0].toString().trim(); // Pega o valor da coluna A
        if (codigo && !codigosExistentes.has(codigo)) {
          codigosExistentes.add(codigo); // Adiciona o código ao conjunto de códigos existentes
          planilhaDestino.appendRow(linha); // Adiciona a linha inteira à planilha de destino
        }
      });
    }
  }
  
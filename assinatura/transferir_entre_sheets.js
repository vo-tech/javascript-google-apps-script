// Versão: 1.0
// Autor: Juliano Ceconi

function consolidarCodigos() {
  try {
    // Nomes das planilhas que serão verificadas
    const nomesPlanilhas = ['GRUPO BOLETO', 'GRUPO RECORRENTE', 'GRUPO CARTÃO DE CRÉDITO', 'GRUPO DINHEIRO'];
    
    // Nome da planilha de destino onde os códigos únicos serão lançados
    const nomePlanilhaDestino = 'TODOS CLIENTES';
    
    // Acessa a planilha ativa
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Obter ou criar a planilha de destino
    const planilhaDestino = obterOuCriarPlanilha(ss, nomePlanilhaDestino);
    
    // Obter os códigos já existentes na planilha de destino para evitar duplicação
    const codigosExistentes = obterCodigosExistentes(planilhaDestino);
    
    // Percorre cada planilha e adiciona os novos códigos à planilha de destino
    nomesPlanilhas.forEach(nome => {
      const planilha = ss.getSheetByName(nome);
      if (!planilha) {
        Logger.log(`Planilha ${nome} não encontrada.`);
        return;
      }
      adicionarNovosCodigos(planilha, planilhaDestino, codigosExistentes);
    });
    
    Logger.log('Códigos consolidados com sucesso na planilha "Consolidado".');
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
  const valoresExistentes = planilhaDestino.getRange('A2:A' + planilhaDestino.getLastRow()).getValues();
  valoresExistentes.forEach(linha => {
    const codigo = linha[0].toString().trim();
    if (codigo) {
      codigosExistentes.add(codigo);
    }
  });
  return codigosExistentes;
}

// Função para adicionar novos códigos à planilha de destino
function adicionarNovosCodigos(planilhaOrigem, planilhaDestino, codigosExistentes) {
  const dados = planilhaOrigem.getRange(2, 1, planilhaOrigem.getLastRow() - 1, planilhaOrigem.getLastColumn()).getValues();

  dados.forEach(linha => {
    const codigo = linha[0].toString().trim(); // Pega o valor da coluna A
    if (codigo && !codigosExistentes.has(codigo)) {
      codigosExistentes.add(codigo); // Adiciona o código ao conjunto de códigos existentes
      planilhaDestino.appendRow(linha); // Adiciona a linha inteira à planilha de destino
    }
  });
}
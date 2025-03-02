// arquivo: conferenciaContasMensal
// versão: 3.2.1
// autor: Juliano Ceconi

// data: 2024-12-05

// Este script cria uma correlação bidirecional entre as abas 'ocultado' e 'ocultado'

const CONFIG = {
  CONTAS_MES_SHEET: 'ocultado',
  ocultado_SHEET: 'ocultado',
  CONTAS_MES_COLUNA_COMPARACAO: 'A',
  ocultado_COLUNA_COMPARACAO: 'H',
  CONTAS_MES_COLUNA_VERIFICACAO: 'K',
  ocultado_COLUNA_ATUALIZACAO: 'O'
};

const VERDADEIRO = true;
const FALSO = false;

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (sheetName === CONFIG.CONTAS_MES_SHEET && 
      col === sheet.getRange(CONFIG.CONTAS_MES_COLUNA_VERIFICACAO + '1').getColumn()) {
    atualizarocultado(sheet, row, e.value);
  } else if (sheetName === CONFIG.ocultado_SHEET && 
             col === sheet.getRange(CONFIG.ocultado_COLUNA_ATUALIZACAO + '1').getColumn()) {
    atualizarContasMes(sheet, row, e.value);
  }
}

function atualizarocultado(contasMesSheet, row, valor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ocultadoSheet = ss.getSheetByName(CONFIG.ocultado_SHEET);
  if (!ocultadoSheet) {
    Logger.log("Aba 'ocultado' não encontrada");
    return;
  }

  const valorVerificacao = (valor === 'TRUE') ? VERDADEIRO : FALSO;
  const valorComparacao = contasMesSheet.getRange(row, contasMesSheet.getRange(CONFIG.CONTAS_MES_COLUNA_COMPARACAO + '1').getColumn()).getValue().toString().toLowerCase();

  // Otimização: Buscar apenas as linhas que correspondem ao valorComparacao
  const textFinder = ocultadoSheet.getRange(CONFIG.ocultado_COLUNA_COMPARACAO + ':' + CONFIG.ocultado_COLUNA_COMPARACAO).createTextFinder(valorComparacao);
  const foundRanges = textFinder.findAll();

  if (foundRanges.length === 0) {
    Logger.log(`Nenhuma correspondência encontrada para ${valorComparacao} em 'ocultado'.`);
    return;
  }

  let atualizacoes = [];
  foundRanges.forEach(range => {
    const row = range.getRow();
    const currentValue = ocultadoSheet.getRange(row, ocultadoSheet.getRange(CONFIG.ocultado_COLUNA_ATUALIZACAO + '1').getColumn()).getValue();
    if (currentValue !== valorVerificacao) {
      atualizacoes.push([row, valorVerificacao]);
    }
  });

  // Atualização em lote
  if (atualizacoes.length > 0) {
    const rangesToUpdate = ocultadoSheet.getRangeList(atualizacoes.map(update => `${CONFIG.ocultado_COLUNA_ATUALIZACAO}${update[0]}`));
    rangesToUpdate.setValue(valorVerificacao);
  }

  Logger.log(`Atualização concluída. ${atualizacoes.length} linha(s) atualizada(s) em 'ocultado'.`);
}

function atualizarContasMes(ocultadoSheet, row, valor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contasMesSheet = ss.getSheetByName(CONFIG.CONTAS_MES_SHEET);
  if (!contasMesSheet) {
    Logger.log("Aba 'ocultado' não encontrada");
    return;
  }

  const valorVerificacao = (valor === 'TRUE') ? VERDADEIRO : FALSO;
  const valorComparacao = ocultadoSheet.getRange(row, ocultadoSheet.getRange(CONFIG.ocultado_COLUNA_COMPARACAO + '1').getColumn()).getValue().toString().toLowerCase();

  // Verificar se todas as ocorrências em 'ocultado' estão como verdadeiro
  const textFinder = ocultadoSheet.getRange(CONFIG.ocultado_COLUNA_COMPARACAO + ':' + CONFIG.ocultado_COLUNA_COMPARACAO).createTextFinder(valorComparacao);
  const foundRanges = textFinder.findAll();

  const todasVerdadeiro = foundRanges.every(range => {
    const row = range.getRow();
    return ocultadoSheet.getRange(row, ocultadoSheet.getRange(CONFIG.ocultado_COLUNA_ATUALIZACAO + '1').getColumn()).getValue() === VERDADEIRO;
  });

  // Atualizar 'ocultado'
  const contasMesRange = contasMesSheet.getRange(CONFIG.CONTAS_MES_COLUNA_COMPARACAO + ':' + CONFIG.CONTAS_MES_COLUNA_COMPARACAO);
  const contasMesTextFinder = contasMesRange.createTextFinder(valorComparacao);
  const contasMesMatch = contasMesTextFinder.findNext();

  if (contasMesMatch) {
    const contasMesRow = contasMesMatch.getRow();
    const currentValue = contasMesSheet.getRange(contasMesRow, contasMesSheet.getRange(CONFIG.CONTAS_MES_COLUNA_VERIFICACAO + '1').getColumn()).getValue();
    if (currentValue !== todasVerdadeiro) {
      contasMesSheet.getRange(contasMesRow, contasMesSheet.getRange(CONFIG.CONTAS_MES_COLUNA_VERIFICACAO + '1').getColumn()).setValue(todasVerdadeiro);
      Logger.log(`'ocultado' atualizado para ${valorComparacao}: ${todasVerdadeiro}`);
    } else {
      Logger.log(`Nenhuma atualização necessária em 'ocultado' para ${valorComparacao}`);
    }
  } else {
    Logger.log(`Nenhuma correspondência encontrada para ${valorComparacao} em 'ocultado'.`);
  }
}
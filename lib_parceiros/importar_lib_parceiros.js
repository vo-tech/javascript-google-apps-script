// arquivo: Código.gs
// versão: 2.0
// autor: Juliano Ceconi

const BIBLIOTECA = ProjetoParceiros; // Ajuste para o apelido correto da sua biblioteca

/**
 * Cria os menus personalizados na planilha
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Cria menu de Faturamento
  ui.createMenu('Faturamento')
    .addItem('Gerar Fatura', 'chamarGeracaoFatura')
    .addToUi();

  // Cria menu de Consolidar Repasse
  ui.createMenu('Consolidar Repasse')
    .addItem('Executar Consolidação', 'executarConsolidacao')
    .addToUi();
}

/**
 * Função para gerar fatura através da biblioteca
 */
function chamarGeracaoFatura() {
  BIBLIOTECA.gerarFatura();
}

/**
 * Função para consolidar repasses através da biblioteca
 */
function executarConsolidacao() {
  try {
    BIBLIOTECA.launchRepasseDate();
    SpreadsheetApp.getUi().alert('Consolidação concluída com sucesso!');
  } catch (error) {
    console.error('Erro na consolidação:', error);
    SpreadsheetApp.getUi().alert(`Erro: ${error.message}\nVerifique os logs para detalhes.`);
  }
}
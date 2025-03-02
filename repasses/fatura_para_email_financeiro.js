// Script para enviar PDF da aba ativa por e-mail
// Versão: 1.1
// Autor: Juliano Ceconi
// Data: 2024-10-22
// 
// Função principal para enviar o PDF por e-mail

function enviarPDFporEmail() {
    console.log('Iniciando processo de envio de PDF');
  
    // Verifica se há dados na coluna A
    if (!verificarDadosColuna()) {
      console.log('Processo interrompido: sem dados na coluna A');
      return;
    }
  
    // Obter a planilha e aba ativa
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Gerar nome do arquivo
    const nomeArquivo = gerarNomeArquivo(sheet.getName());
    console.log('Nome do arquivo gerado: ' + nomeArquivo);
    
    // Gerar PDF
    console.log('Iniciando geração do PDF');
    const pdfBlob = gerarPDF(ss.getId(), sheet.getSheetId(), nomeArquivo);
    console.log('PDF gerado com sucesso');
    
    // Enviar e-mail
    console.log('Iniciando envio do e-mail');
    enviarEmail(pdfBlob, nomeArquivo);
    console.log('E-mail enviado com sucesso');
    
    // Mostrar pop-up de confirmação
    mostrarPopUp('E-mail enviado com sucesso!');
    
    // Desabilitar o botão temporariamente
    desabilitarBotao();
  
    console.log('Processo de envio de PDF concluído');
  }
  
  /**
   * Verifica se há dados na coluna A além do cabeçalho
   * @return {boolean} True se houver dados, False caso contrário
   */
  function verificarDadosColuna() {
    const dados = SpreadsheetApp.getActiveSheet().getRange('A2:A').getValues();
    const temDados = dados.some(row => row[0] !== '');
    console.log('Verificação de dados na coluna A: ' + (temDados ? 'Dados encontrados' : 'Nenhum dado encontrado'));
    if (!temDados) {
      mostrarPopUp('Nenhum dado encontrado na coluna A. O e-mail não será enviado.');
    }
    return temDados;
  }
  
  /**
   * Gera o nome do arquivo no formato especificado
   * @param {string} nomeAba - Nome da aba ativa
   * @return {string} Nome do arquivo formatado
   */
  function gerarNomeArquivo(nomeAba) {
    const agora = new Date();
    const dataFormatada = Utilities.formatDate(agora, 'GMT-3', "yyyy-MM-dd_HH'h'mm");
    return `${dataFormatada}_${nomeAba}`;
  }
  
  /**
   * Gera o PDF da aba ativa
   * @param {string} ssId - ID da planilha
   * @param {number} sheetId - ID da aba
   * @param {string} nomeArquivo - Nome do arquivo PDF
   * @return {Blob} Blob do PDF gerado
   */
  function gerarPDF(ssId, sheetId, nomeArquivo) {
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
                `format=pdf&` +
                `gid=${sheetId}&` +
                `size=A4&` +
                `portrait=false&` +
                `fitw=true&` +
                `gridlines=false&` +
                `printtitle=false&` +
                `sheetnames=false&` +
                `pagenum=CENTER&` +
                `fzr=false&` +
                `horizontal_alignment=CENTER&` +
                `vertical_alignment=TOP&` +
                `scale=4&` +
                `top_margin=0.50&` +
                `bottom_margin=0.50&` +
                `left_margin=0.50&` +
                `right_margin=0.50`;
  
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {'Authorization': 'Bearer ' + token}
    });
    
    return response.getBlob().setName(nomeArquivo + '.pdf');
  }
  
  /**
   * Envia o e-mail com o PDF anexado
   * @param {Blob} pdfBlob - Blob do PDF
   * @param {string} nomeArquivo - Nome do arquivo PDF
   */
  function enviarEmail(pdfBlob, nomeArquivo) {
    const destinatarios = ['ocultado'];
    const assunto = 'Pedido de Repasse Pendente';
    const corpo = 'Há um pedido de repasse pendente. Acesse a planilha em: ' +
                  'ocultado';
  
    MailApp.sendEmail({
      to: destinatarios.join(','),
      subject: assunto,
      body: corpo,
      attachments: [pdfBlob]
    });
    console.log('E-mail enviado para: ' + destinatarios.join(', '));
  }
  
  /**
   * Mostra um pop-up com a mensagem fornecida
   * @param {string} mensagem - Mensagem a ser exibida
   */
  function mostrarPopUp(mensagem) {
    SpreadsheetApp.getUi().alert(mensagem);
    console.log('Pop-up exibido: ' + mensagem);
  }
  
  /**
   * Desabilita o botão temporariamente
   */
  function desabilitarBotao() {
    console.log('Iniciando desabilitação temporária do botão');
    // Implementação depende de como o botão foi criado na planilha
    // Esta é uma implementação de exemplo
    const sheet = SpreadsheetApp.getActiveSheet();
    const botao = sheet.getDrawings()[0]; // Assume que o botão é o primeiro desenho na planilha
    
    if (botao) {
      botao.setDisabled(true);
      console.log('Botão desabilitado por 30 segundos');
      Utilities.sleep(30000); // Espera 30 segundos
      botao.setDisabled(false);
      console.log('Botão reabilitado');
    } else {
      console.log('Botão não encontrado');
    }
  }
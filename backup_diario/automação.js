// arquivo: automação.gs
// versão 2.1
// autor: Juliano Ceconi 

/***********************
 * CONFIGURAÇÕES GLOBAIS
 ***********************/
const CONFIG = {
    // IDs das pastas de origem (onde estão os arquivos Google Sheets)
    sourceFolderIds: ['ocultado', 'ocultado'],
  
    // ID da pasta de destino (onde serão salvos os PDFs gerados)
    destinationFolderId: 'ocultado',
  
    // E-mails que receberão o backup
    // emailAddresses: ['ocultado'],
    // Exemplo para enviar a vários destinatários:
    emailAddresses: ['ocultado', 'ocultado', 'ocultado'],
  
    // Fuso horário
    timeZone: 'America/Sao_Paulo'
  };
  
  /**
   * Função principal que orquestra todo o processo de backup
   */
  function mainBackupProcess() {
    const startTime = new Date();
    Logger.log('Iniciando processo de backup: ' + startTime);
  
    try {
      // 1. Converter as planilhas em PDFs e armazená-los na pasta de destino
      const pdfFiles = convertSheetsToPDF();
  
      // 2. Criar um arquivo ZIP contendo os PDFs
      const zipBlob = createZipFile(pdfFiles);
  
      // 3. Enviar o ZIP por e-mail
      sendEmail(zipBlob);
  
      // 4. Limpar (mover para lixeira) os PDFs temporários já que o ZIP foi gerado
      cleanupTempFiles(pdfFiles);
  
      const endTime = new Date();
      Logger.log('Processo de backup concluído com sucesso: ' + endTime);
      Logger.log('Tempo total de execução: ' + (endTime - startTime) / 1000 + ' segundos');
    } catch (error) {
      Logger.log('Erro durante o processo de backup: ' + error.toString());
    }
  }
  
  /**
   * Converte as planilhas para PDF e retorna uma lista de arquivos PDF
   * que foram efetivamente criados na pasta de destino
   *
   * @returns {Array<GoogleAppsScript.Drive.File>} Array de arquivos PDF gerados
   */
  function convertSheetsToPDF() {
    Logger.log('Iniciando conversão de planilhas para PDF');
    const pdfFiles = [];
  
    // Percorre cada pasta de origem
    CONFIG.sourceFolderIds.forEach(folderId => {
      const folder = DriveApp.getFolderById(folderId);
      
      // Obtem somente arquivos que sejam planilhas (Google Sheets)
      const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
      while (files.hasNext()) {
        const file = files.next();
  
        // 1. Gera o blob do PDF
        const pdfBlob = generatePDFFromSheet(file);
  
        // 2. Salva o PDF na pasta de destino
        const pdfFile = savePDFToDestination(pdfBlob, file.getName());
  
        // 3. Adiciona o arquivo PDF gerado ao array
        pdfFiles.push(pdfFile);
      }
    });
  
    Logger.log('Conversão para PDF concluída. Total de arquivos: ' + pdfFiles.length);
    return pdfFiles;
  }
  
  /**
   * Gera um PDF a partir de uma planilha com as especificações solicitadas,
   * sem modificar o arquivo original.
   *
   * @param {GoogleAppsScript.Drive.File} file - Arquivo de planilha do Google
   * @returns {GoogleAppsScript.Base.Blob} Blob contendo o PDF gerado
   */
  function generatePDFFromSheet(file) {
    const spreadsheet = SpreadsheetApp.open(file);
  
    // Monta a URL de exportação em PDF
    const url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?";
    const exportOptions = 
      "exportFormat=pdf&format=pdf" +
      "&size=A4" +
      "&portrait=false" +
      "&fitw=true" +
      "&fith=false" +
      "&top_margin=0.12" +
      "&bottom_margin=0.12" +
      "&left_margin=0.12" +
      "&right_margin=0.12" +
      "&sheetnames=true" +
      "&printtitle=false" +
      "&pagenumbers=true" +
      "&gridlines=false" +
      "&fzr=false" +
      "&scale=100" +
      `&headers.center=${encodeURIComponent(file.getName() + " - " + spreadsheet.getActiveSheet().getName())}` +
      `&headers.left=${encodeURIComponent(Utilities.formatDate(new Date(), CONFIG.timeZone, 'yyyy-MM-dd HH:mm:ss'))}` +
      `&headers.right=${encodeURIComponent('Página &P de &N')}`;
  
    try {
      // Faz a requisição de exportação
      const response = UrlFetchApp.fetch(url + exportOptions, {
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        },
        muteHttpExceptions: true
      });
  
      // Verifica se a resposta foi OK (código 200)
      if (response.getResponseCode() !== 200) {
        Logger.log('Resposta completa para ' + file.getName() + ': ' + response.getContentText());
        throw new Error(
          'Falha ao gerar PDF para ' + file.getName() +
          '. Código de resposta: ' + response.getResponseCode() +
          '. Conteúdo: ' + response.getContentText().substring(0, 200)
        );
      }
  
      // Retorna o blob do PDF gerado
      return response.getBlob().setName(file.getName() + ".pdf");
    } catch (error) {
      Logger.log('Erro detalhado ao processar ' + file.getName() + ': ' + error.toString());
      
      // Tenta uma abordagem alternativa de exportação (getAs PDF)
      return fallbackPDFGeneration(spreadsheet, file.getName());
    }
  }
  
  /**
   * Método alternativo para gerar PDF em caso de falha no método principal
   *
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - Planilha do Google
   * @param {string} fileName - Nome do arquivo
   * @returns {GoogleAppsScript.Base.Blob} Blob contendo o PDF gerado
   */
  function fallbackPDFGeneration(spreadsheet, fileName) {
    Logger.log('Tentando método alternativo de geração de PDF para ' + fileName);
    const blob = spreadsheet.getAs('application/pdf')
      .setName(fileName + ".pdf");
    return blob;
  }
  
  /**
   * Salva o PDF (Blob) na pasta de destino e retorna o arquivo criado
   *
   * @param {GoogleAppsScript.Base.Blob} pdfBlob - Blob do PDF gerado
   * @param {string} originalName - Nome original do arquivo de planilha
   * @returns {GoogleAppsScript.Drive.File} Arquivo PDF salvo no Drive
   */
  function savePDFToDestination(pdfBlob, originalName) {
    const destFolder = DriveApp.getFolderById(CONFIG.destinationFolderId);
    const timestamp = Utilities.formatDate(new Date(), CONFIG.timeZone, 'yyyy-MM-dd_HH:mm:ss');
    const newFileName = `${originalName}_${timestamp}.pdf`;
    return destFolder.createFile(pdfBlob).setName(newFileName);
  }
  
  /**
   * Cria um arquivo ZIP com todos os PDFs que foram salvos no Drive.
   * Observação: Utilities.zip() requer um array de BLOBS, não de FILES.
   *
   * @param {Array<GoogleAppsScript.Drive.File>} pdfFiles - Array de arquivos PDF
   * @returns {GoogleAppsScript.Base.Blob} Blob do arquivo ZIP
   */
  function createZipFile(pdfFiles) {
    Logger.log('Iniciando criação do arquivo ZIP');
  
    // Converte cada File em Blob
    const pdfBlobs = pdfFiles.map(file => file.getBlob());
  
    // Gera o blob zipado
    const zipBlob = Utilities.zip(pdfBlobs, 'Backup_' + new Date().toISOString().split('T')[0] + '.zip');
  
    Logger.log('Arquivo ZIP criado com sucesso');
    return zipBlob;
  }
  
  /**
   * Envia o e-mail com o arquivo ZIP anexado
   *
   * @param {GoogleAppsScript.Base.Blob} zipBlob - Blob do arquivo ZIP
   */
  function sendEmail(zipBlob) {
    Logger.log('Iniciando envio de e-mail');
    const subject = 'Backup diário de planilhas - ' + new Date().toISOString().split('T')[0];
    const body = 'Segue em anexo o backup diário das planilhas em formato PDF.';
  
    GmailApp.sendEmail(CONFIG.emailAddresses.join(','), subject, body, {
      attachments: [zipBlob],
      name: 'Backup Automático'
    });
  
    Logger.log('E-mail enviado com sucesso');
  }
  
  /**
   * Remove (move para a lixeira) os arquivos PDF temporários da pasta de destino
   *
   * @param {Array<GoogleAppsScript.Drive.File>} pdfFiles - Array de arquivos PDF temporários
   */
  function cleanupTempFiles(pdfFiles) {
    Logger.log('Iniciando limpeza de arquivos temporários');
    pdfFiles.forEach(file => file.setTrashed(true));
    Logger.log('Limpeza de arquivos temporários concluída');
  }
  
  /**
   * Configura o acionador para executar o script diariamente às 23:00
   */
  function setupTrigger() {
    ScriptApp.newTrigger('mainBackupProcess')
      .timeBased()
      .everyDays(1)
      .atHour(23)
      .create();
    
    Logger.log('Acionador configurado para execução diária às 23:00');
  }
  
// Versão: 1.0
// Autor: Juliano Ceconi

// Função que é acionada a cada edição na planilha
function onEdit(e) {
    try {
      if (!e || !e.range) throw new Error('Evento não definido ou inválido');
      const sheet = e.range.getSheet();
      const range = e.range;
      let oldValue = e.oldValue;
      let newValue = e.value;
      const timestamp = new Date();
  
      // Captura o e-mail do usuário que fez a edição
      let userEmail = e.user ? e.user.getEmail() : Session.getActiveUser().getEmail();
  
      // Se o e-mail não puder ser capturado, deixa o campo em branco
      if (!userEmail) userEmail = "";
  
      // Validações para evitar valores nulos
      if (oldValue === undefined) oldValue = "";
      if (newValue === undefined) newValue = "";
  
      // Formata a data e hora de forma concisa
      const formattedTimestamp = `${timestamp.getDate().toString().padStart(2, '0')}/${(timestamp.getMonth() + 1).toString().padStart(2, '0')}/${timestamp.getFullYear().toString().slice(-2)} ${timestamp.getHours().toString().padStart(2, '0')}:${timestamp.getMinutes().toString().padStart(2, '0')}`;
  
      // Cria uma mensagem de log com os detalhes da alteração em formato CSV
      const logMessage = `${formattedTimestamp};${userEmail};${sheet.getName()};${range.getA1Notation()};${oldValue};${newValue}\n`;
      const logLength = logMessage.length; // Tamanho do novo registro de log
  
      const cache = CacheService.getScriptCache();
      let logs = cache.get('logs') || '';
      let currentSize = parseInt(cache.get('logs_size')) || 0;
  
      // Atualiza os logs e o contador de tamanho
      logs += logMessage;
      currentSize += logLength;
  
      // Verifica se o tamanho acumulado está próximo do limite (50 KB)
      if (currentSize >= 50000) {
        saveLogsToTxt(logs); // Salva os logs em um arquivo .txt
        cache.remove('logs');
        cache.remove('logs_size');
      } else {
        // Atualiza a cache com os novos logs e tamanho
        cache.put('logs', logs, 21600); // Armazena os logs por até 6 horas
        cache.put('logs_size', currentSize.toString(), 21600); // Armazena o tamanho por até 6 horas
      }
    } catch (error) {
      Logger.log(`Erro na função onEdit: ${error.message}`);
    }
  }
  
  // Função que salva os logs acumulados em um arquivo .csv no Google Drive
  function saveLogsToTxt(logs) {
    try {
      const folderId = 'ocultado'; // ID da pasta onde deseja salvar o arquivo
      const timestamp = new Date();
      const formattedFileName = `${timestamp.getFullYear()}-${(timestamp.getMonth() + 1).toString().padStart(2, '0')}-${timestamp.getDate().toString().padStart(2, '0')} ${timestamp.getHours().toString().padStart(2, '0')}:${timestamp.getMinutes().toString().padStart(2, '0')} Logs Assinatura.csv`;
      const folder = DriveApp.getFolderById(folderId);
  
      // Adiciona o BOM (Byte Order Mark) ao início dos logs para forçar o encoding UTF-8
      const bom = '\uFEFF';
      const blob = Utilities.newBlob(bom + logs, 'text/csv', formattedFileName);
      folder.createFile(blob);
  
      Logger.log('Logs salvos com sucesso no arquivo .csv com BOM.');
    } catch (error) {
      Logger.log(`Erro ao salvar os logs no arquivo .csv: ${error.message}`);
    }
  }
  
  // Função chamada em intervalos de tempo para garantir que os logs não sejam perdidos
  function saveLogsFromCacheIfAny() {
    try {
      const cache = CacheService.getScriptCache();
      const logs = cache.get('logs');
      
      if (logs) {
        saveLogsToTxt(logs); // Salva os logs em um arquivo .csv
        cache.remove('logs');
        cache.remove('logs_size');
      }
    } catch (error) {
      Logger.log(`Erro na função saveLogsFromCacheIfAny: ${error.message}`);
    }
  }
  
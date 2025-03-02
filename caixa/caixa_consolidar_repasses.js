// arquivo: consolidar_repasses.gs
// versão: 1.0
// autor: Juliano Ceconi

let logMessages = []; // Array para acumular mensagens de log

function transferirDataRepasse() {
  const planilhaA = 'ocultado';
  const planilhaB = 'ocultado';
  const ocultadoSheetName = 'ocultado';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetA, sheetB, ocultadoSheet;
  let updatedRows = 0;
  let ocultadoData = [];

  try {
    // Obter referências das planilhas
    sheetA = ss.getSheetByName(planilhaA);
    sheetB = ss.getSheetByName(planilhaB);
    ocultadoSheet = ss.getSheetByName(ocultadoSheetName);

    // Verificar se todas as planilhas foram encontradas
    if (!sheetA || !sheetB || !ocultadoSheet) {
      throw new Error('Uma ou mais planilhas não foram encontradas.');
    }

    // Desbloquear a planilha antes de fazer alterações
    // unlockSheet(sheetA); // Descomente para desbloquear a planilha

    // Proteger intervalo de repasse
    const rangeToProtect = sheetA.getRange('K2:P' + sheetA.getLastRow());
    const protection = rangeToProtect.protect();
    protection.setDescription('Proteção temporária durante a execução do script');

    // Obter dados das planilhas
    const dataB = sheetB.getRange(2, 2, sheetB.getLastRow() - 1, 13).getValues();
    const dataA = sheetA.getRange(2, 1, sheetA.getLastRow() - 1, 16).getValues();

    // Obter o e-mail do usuário que executou o script
    const userEmail = Session.getActiveUser().getEmail();

    // Processar dados da planilha B
    dataB.forEach(function (rowB, i) {
      const guiaB = rowB[0];
      const dataRepasseB = rowB[11];
      const valorRepasseB = rowB[7]; // Coluna I (índice 7) de ocultado

      if (guiaB && (dataRepasseB || valorRepasseB)) {
        let found = false;

        // Processar dados da planilha A
        dataA.forEach(function (rowA, j) {
          const guiaA = rowA[0];
          const dataRepasseA = rowA[15]; // Coluna P (índice 15) de ocultado
          const valorRepasseA = rowA[10]; // Coluna K (índice 10) de ocultado

          if (guiaA === guiaB) {
            found = true;
            let updated = false;
            let updateMessage = [];

            // Verifica se a data de repasse em 'ocultado' está vazia ou diferente
            if (dataRepasseB && (!dataRepasseA || dataRepasseA.toString() !== dataRepasseB.toString())) {
              sheetA.getRange(j + 2, 16).setValue(Utilities.formatDate(new Date(dataRepasseB), Session.getScriptTimeZone(), 'd/M'));
              updateMessage.push("data de repasse");
              updated = true;
            }

            // Verifica se o valor de repasse está vazio ou é diferente
            if (valorRepasseB && (valorRepasseA === '' || valorRepasseA !== valorRepasseB)) {
              sheetA.getRange(j + 2, 11).setValue(valorRepasseB);
              updateMessage.push("valor de repasse");
              updated = true;
            }

            // Atualiza ocultado e log se houve alteração
            if (updated) {
              ocultadoData.push([new Date(), guiaB, dataRepasseB || '', valorRepasseB || '', userEmail]);
              logMessages.push(`Guia ${guiaB} - Sucesso (atualizada: ${updateMessage.join(", ")})`); // Acumula a mensagem
              updatedRows++;
            } else {
              logMessages.push(`Guia ${guiaB} - Sem alteração (já atualizada)`); // Acumula a mensagem
            }
          }
        });

        if (!found) {
          logMessages.push(`Guia ${guiaB} - Falha (não encontrada)`); // Acumula a mensagem
        }
      } else {
        logMessages.push(`Linha ${i + 2} - Guia ou informações de repasse vazias`); // Acumula a mensagem
      }
    });

    // Atualiza ocultado na planilha
    if (updatedRows > 0) {
      ocultadoSheet.getRange(ocultadoSheet.getLastRow() + 1, 1, ocultadoData.length, 5).setValues(ocultadoData);
    }

    // Realiza o backup imediatamente após a transferência bem-sucedida
    if (updatedRows > 0) {
      backupocultado(ocultadoData);
    }

    Logger.log(`Processo concluído: ${updatedRows} ocultado atualizadas com sucesso.`);
    
  } catch (error) {
    Logger.log(`Erro: ${error.message}`);
    logMessages.push(`Erro: ${error.message}`); // Acumula a mensagem de erro
  } finally {
    if (sheetA) {
      // Remove a proteção, independentemente de ter ocorrido um erro ou não
      const protections = sheetA.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (let i = 0; i < protections.length; i++) {
        if (protections[i].getDescription() === 'Proteção temporária durante a execução do script') {
          protections[i].remove();
        }
      }
      // Bloquear a planilha novamente, exceto a primeira linha
      // lockSheet(sheetA); // Descomente para bloquear a planilha novamente
    }

    // Escreve todas as mensagens de log no arquivo de texto ao final
    if (logMessages.length > 0) {
      logToFile(logMessages.join('\n')); // Escreve todas as mensagens de log de uma vez
    }
  }
}

function unlockSheet(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (let i = 0; i < protections.length; i++) {
    protections[i].remove();
  }
}

function lockSheet(sheet) {
  const protection = sheet.protect().setDescription('Protegido após execução do script');
  protection.setWarningOnly(true); // Permite que os editores vejam a planilha, mas não a editem
  protection.addEditor(Session.getEffectiveUser()); // Adiciona o usuário que executou o script como editor
  protection.removeEditors(protection.getEditors()); // Remove outros editores, se necessário

  // Permitir edição da primeira linha
  const range = sheet.getRange('1:1');
  protection.setUnprotectedRanges([range]);
}

function backupocultado(dadosocultado) {
  try {
    const ss = SpreadsheetApp.openById("ocultado");
    let ocultadoBackupSheet = ss.getSheetByName('ocultado');

    // Verifica se a planilha de backup existe, se não, cria-a
    if (!ocultadoBackupSheet) {
      Logger.log('Planilha "ocultado" não encontrada. Criando nova planilha.');
      ocultadoBackupSheet = ss.insertSheet('ocultado');
    }

    Logger.log('Iniciando backup do ocultado');

    // Backup do ocultado
    if (dadosocultado && dadosocultado.length > 0) {
      ocultadoBackupSheet.getRange(ocultadoBackupSheet.getLastRow() + 1, 1, dadosocultado.length, dadosocultado[0].length).setValues(dadosocultado);
      Logger.log(`Backup do ocultado concluído: ${dadosocultado.length} linhas adicionadas`);
    } else {
      Logger.log('Nenhum dado de ocultado para backup');
    }

  } catch (error) {
    Logger.log(`Erro durante o backup: ${error.message}`);
  }
}

// Função para transferir ocultado que atendem a condição
function transferirocultadoValidas() {
  const sheetA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ocultado');
  const sheetB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ocultado');
  const ocultadoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ocultado');
  let somatorioocultadoTransferidas = 0;
  let dataTransferencia = [];

  const dataA = sheetA.getRange(2, 1, sheetA.getLastRow() - 1, 16).getValues(); // Guia A e Valor P
  const dataB = sheetB.getRange(2, 2, sheetB.getLastRow() - 1, 13).getValues(); // Guia B e Valor Repasse M

  dataB.forEach(function (rowB, i) {
    const guiaB = rowB[0];
    const valorB = rowB[11]; // Valor de repasse em B

    if (guiaB && valorB) {
      dataA.forEach(function (rowA, j) {
        const guiaA = rowA[0];
        const valorA = rowA[15]; // Valor de repasse em A

        if (guiaA === guiaB && valorA === valorB) {
          somatorioocultadoTransferidas++;
          dataTransferencia.push([new Date(), guiaA, valorA]);
          sheetB.deleteRow(i + 2); // Apaga a linha transferida de B
        }
      });
    }
  });

  ocultadoSheet.getRange(ocultadoSheet.getLastRow() + 1, 1, dataTransferencia.length, 3).setValues(dataTransferencia); // Atualiza ocultado
  ocultadoSheet.getRange(ocultadoSheet.getLastRow() + 1, 1, 1, 1).setValue(`Total de ocultado transferidas: ${somatorioocultadoTransferidas}`);
}

function logToFile(logContent) {
  const folderId = 'ocultado'; // ID da sua pasta
  const fileName = 'logs.txt'; // Nome do arquivo de log
  const folder = DriveApp.getFolderById(folderId); // Obtém a pasta pelo ID
  let file;

  // Verifica se o arquivo já existe
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    file = files.next();
  } else {
    // Cria um novo arquivo se não existir
    file = folder.createFile(fileName, '', MimeType.PLAIN_TEXT);
  }

  // Adiciona o novo conteúdo de log ao arquivo
  const currentContent = file.getBlob().getDataAsString();
  const newContent = currentContent + logContent + '\n';
  file.setContent(newContent);
}
// arquivo: LibGeracaoFatura.gs
// versão: 2.1
// autor: Juliano Ceconi

function gerarFatura() {
    // Chamada para a função launchRepasseDate() para executar as tarefas preliminares
launchRepasseDate();

try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  Logger.log("Planilha ativa obtida: " + ss.getName());
  
  // Configurações de data e formatos
  var timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  var dataAtualAba = Utilities.formatDate(new Date(), timeZone, "dd-MM-yy");  // ex: 13-02-25
  var dataNomeArquivo = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd-HH-mm"); // ex: 2025-02-13-17-29
  var nomePlanilha = ss.getName().replace(/\s+/g, '-').toLowerCase();
  Logger.log("dataAtualAba: " + dataAtualAba + ", dataNomeArquivo: " + dataNomeArquivo);
  
  var abaOrigem = ss.getSheetByName('autofatura');
  if (!abaOrigem) {
    ui.alert("A aba 'autofatura' não foi encontrada.");
    Logger.log("Erro: aba 'autofatura' não encontrada.");
    return;
  }
  
  // Verificar se a aba com data atual já existe
  if (ss.getSheetByName(dataAtualAba)) {
    ui.alert("Já existe uma fatura fechada para a data " + dataAtualAba + ".");
    Logger.log("Fatura para a data " + dataAtualAba + " já existe.");
    return;
  }
  
  // Obter os valores exibidos para determinar o menor retângulo com dados visíveis
  var dadosDisplay = abaOrigem.getDataRange().getDisplayValues();
  var totalLinhas = dadosDisplay.length;
  var totalColunas = dadosDisplay[0].length;
  var ultimaLinhaComValor = 0;
  var ultimaColunaComValor = 0;
  
  for (var i = 0; i < totalLinhas; i++) {
    for (var j = 0; j < totalColunas; j++) {
      // Removendo espaços para garantir que células com apenas espaços sejam consideradas vazias
      if (dadosDisplay[i][j].toString().trim() !== "") {
        // Atualiza a última linha e coluna com conteúdo visível
        if ((i + 1) > ultimaLinhaComValor) {
          ultimaLinhaComValor = i + 1;
        }
        if ((j + 1) > ultimaColunaComValor) {
          ultimaColunaComValor = j + 1;
        }
      }
    }
  }
  
  // Se nenhum dado visível for encontrado, aborta o processo
  if (ultimaLinhaComValor === 0 || ultimaColunaComValor === 0) {
    ui.alert("Nenhum dado visível encontrado na aba 'autofatura'. O processo foi abortado.");
    Logger.log("Processo abortado: não há dados visíveis.");
    return;
  }
  
  Logger.log("Intervalo visível: Linhas = 1 até " + ultimaLinhaComValor + " e Colunas = 1 até " + ultimaColunaComValor);
  
  // Criar nova aba
  var novaAba = ss.insertSheet(dataAtualAba);
  Logger.log("Nova aba criada: " + dataAtualAba);
    
  // Redimensionar a nova aba para o tamanho exato do intervalo visível
  // Remover linhas extras
  var totalLinhasNovaAba = novaAba.getMaxRows();
  if (totalLinhasNovaAba > ultimaLinhaComValor) {
    novaAba.deleteRows(ultimaLinhaComValor + 1, totalLinhasNovaAba - ultimaLinhaComValor);
    Logger.log("Excluídas " + (totalLinhasNovaAba - ultimaLinhaComValor) + " linhas excedentes.");
  }
  // Remover colunas extras
  var totalColunasNovaAba = novaAba.getMaxColumns();
  if (totalColunasNovaAba > ultimaColunaComValor) {
    novaAba.deleteColumns(ultimaColunaComValor + 1, totalColunasNovaAba - ultimaColunaComValor);
    Logger.log("Excluídas " + (totalColunasNovaAba - ultimaColunaComValor) + " colunas excedentes.");
  }
  
  // Definir os intervalos da nova aba
  var intervaloVisivelOrigem = abaOrigem.getRange(1, 1, ultimaLinhaComValor, ultimaColunaComValor);
  var intervaloVisivelDestino = novaAba.getRange(1, 1, ultimaLinhaComValor, ultimaColunaComValor);
  
  // Copiar formatação e validações para o intervalo visível
  intervaloVisivelOrigem.copyTo(intervaloVisivelDestino, {formatOnly: true});
  intervaloVisivelOrigem.copyTo(intervaloVisivelDestino, {validationsOnly: true});
  
  // Copiar proteções que estejam integralmente dentro do intervalo visível
  var protecoesOrigem = abaOrigem.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protecoesOrigem.length; i++) {
    var protecaoOrigem = protecoesOrigem[i];
    var rng = protecaoOrigem.getRange();
    // Verifica se o intervalo da proteção está totalmente dentro do intervalo visível
    if (
      rng.getRow() >= 1 &&
      rng.getColumn() >= 1 &&
      (rng.getRow() + rng.getNumRows() - 1) <= ultimaLinhaComValor &&
      (rng.getColumn() + rng.getNumColumns() - 1) <= ultimaColunaComValor
    ) {
      var novaProtecaoRng = novaAba.getRange(rng.getRow(), rng.getColumn(), rng.getNumRows(), rng.getNumColumns());
      var novaProtecao = novaProtecaoRng.protect();
      novaProtecao.setDescription(protecaoOrigem.getDescription());
      novaProtecao.setWarningOnly(protecaoOrigem.isWarningOnly());
    }
  }
  Logger.log("Formatação, validações e proteções (dentro do intervalo visível) copiadas.");
  
  // Copiar os valores originais (incluindo fórmulas) para o intervalo visível
  var valores = intervaloVisivelOrigem.getValues();
  intervaloVisivelDestino.setValues(valores);
  
  // Ajustar dimensões: larguras das colunas e alturas das linhas para o intervalo visível
  for (var i = 1; i <= ultimaColunaComValor; i++) {
    novaAba.setColumnWidth(i, abaOrigem.getColumnWidth(i));
  }
  for (var i = 1; i <= ultimaLinhaComValor; i++) {
    novaAba.setRowHeight(i, abaOrigem.getRowHeight(i));
  }
  Logger.log("Ajuste de dimensões concluído para o intervalo visível.");
  
  // Gerar CSV a partir do intervalo visível (utilizando o timezone da planilha)
  var csv = gerarCSV(intervaloVisivelDestino, timeZone);
  var csvBlob = Utilities.newBlob(csv, MimeType.CSV, dataNomeArquivo + '-' + nomePlanilha + '.csv');
  Logger.log("Blob CSV criado: " + csvBlob.getName());
  
  // Gerar PDF da nova aba (já com o intervalo ajustado)
  var pdfBlob = gerarPDFDaAba(ss, novaAba, dataNomeArquivo, nomePlanilha);
  Logger.log("PDF gerado com sucesso.");
  
  // Obter o e-mail do parceiro
  var partnerEmail = getPartnerEmail(ss.getId());
  Logger.log("E-mail do parceiro obtido: " + partnerEmail);
  if (!partnerEmail) {
    ui.alert("E-mail do parceiro não encontrado para a planilha atual.");
    Logger.log("Erro: E-mail do parceiro não encontrado para a planilha ID " + ss.getId());
    return;
  }
  
  // Enviar o e-mail com o PDF (e opcionalmente com o CSV em anexo)
  enviarEmailComPDF(partnerEmail, pdfBlob, dataAtualAba, csvBlob);
  Logger.log("E-mail enviado para " + partnerEmail);
  
  ui.alert("Fatura gerada e e-mail enviado para " + partnerEmail);
  
} catch (e) {
  Logger.log("Erro na função gerarFatura: " + e.toString());
  SpreadsheetApp.getUi().alert("Erro ao gerar fatura: " + e.toString());
}
}

/**
* Gera o CSV a partir do range de dados.
* @param {Range} range - O range com os dados da nova aba.
* @param {string} timeZone - Timezone da planilha.
* @returns {string} - String CSV.
*/
function gerarCSV(range, timeZone) {
var dados = range.getValues();
return dados.map(function(linha) {
  return linha.map(function(celula) {
    if (typeof celula === 'string') {
      var valor = celula.replace(/"/g, '""');
      return '"' + valor + '"';
    } else if (celula instanceof Date) {
      return Utilities.formatDate(celula, timeZone, "yyyy-MM-dd HH:mm:ss");
    } else {
      return celula;
    }
  }).join(',');
}).join('\r\n');
}

/**
* Gera o PDF da aba especificada.
* @param {Spreadsheet} ss - A planilha ativa.
* @param {Sheet} aba - A aba que deseja exportar para PDF.
* @param {string} dataNomeArquivo - Nome base para o arquivo.
* @param {string} nomePlanilha - Nome da planilha.
* @returns {Blob} - O PDF convertido em blob.
*/
function gerarPDFDaAba(ss, aba, dataNomeArquivo, nomePlanilha) {
try {
  var ssId = ss.getId();
  var sheetId = aba.getSheetId();
  
  // Construir a URL de exportação para gerar somente a aba desejada em PDF
  var pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?';
  var exportOptions = [
    'exportFormat=pdf',
    'format=pdf',
    'gid=' + sheetId,
    'size=A4',
    'portrait=true',
    'fitw=true',
    'sheetnames=true',
    'printtitle=true',
    'pagenumbers=true',
    'gridlines=false',
    'fzr=false'
  ];
  pdfUrl += exportOptions.join('&');
  
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(pdfUrl, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  
  return response.getBlob().setName(dataNomeArquivo + '-' + nomePlanilha + '.pdf');
  
} catch (e) {
  Logger.log("Erro na função gerarPDFDaAba: " + e.toString());
  throw e;
}
}

/**
* Procura e retorna o e-mail do parceiro associado à planilha atual.
* Utiliza a planilha central "parceiros" com os cabeçalhos:
* "Parceiro", "Link", "autofatura1", "autofatura2", "Biblioteca", "e-mail", "id"
* 
* Assume:
* - A coluna "id" (índice 6) contém o ID da planilha do parceiro.
* - A coluna "e-mail" (índice 5) contém o e-mail do parceiro.
*
* @param {string} partnerSheetId - O ID da planilha individual do parceiro.
* @returns {string|null} - O e-mail do parceiro ou null se não encontrar.
*/
function getPartnerEmail(partnerSheetId) {
try {
  var masterSheetId = 'ocultado';
  var masterSS = SpreadsheetApp.openById(masterSheetId);
  var sheet = masterSS.getSheetByName('parceiros');
  if (!sheet) {
    Logger.log("Planilha 'parceiros' não encontrada no arquivo mestre.");
    return null;
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var idValue = data[i][6]; // coluna "id"
    if (idValue && idValue.toString().indexOf(partnerSheetId) !== -1) {
      return data[i][5]; // coluna "e-mail"
    }
  }
  Logger.log("Nenhum parceiro encontrado com o ID " + partnerSheetId);
  return null;
} catch (e) {
  Logger.log("Erro na função getPartnerEmail: " + e.toString());
  throw e;
}
}

/**
* Envia o e-mail com os PDFs (e opcionalmente o CSV) anexados.
* @param {string} partnerEmail - O e-mail do parceiro.
* @param {Blob} pdfBlob - O PDF da fatura.
* @param {string} dataAba - Data identificadora da fatura.
* @param {Blob} csvBlob - (Opcional) Blob do CSV da fatura.
*/
function enviarEmailComPDF(partnerEmail, pdfBlob, dataAba, csvBlob) {
try {
  var assunto = "Fatura referente a " + dataAba;
  var mensagem = "<p>Prezado parceiro,</p>" +
                 "<p>Segue em anexo a fatura referente à data " + dataAba + ".</p>" +
                 "<p>As notas fiscais referentes ao valor de repasse devem ser emitidas em nome dos respectivos pacientes atendidos.</p>" +
                 "<p>Pedimos por gentileza que nos coloquem como destinatários ao enviar as notas para os respectivos pacientes.</p>" +
                 "<p></p>" +
                 "<p>Atenciosamente,</p>" +
                 "<p>ocultado</p>";
  
  // Se desejar enviar também o CSV, inclua-o no array de attachments
  var attachments = [pdfBlob];
  // Para incluir o CSV, descomente a linha abaixo:
  // attachments.push(csvBlob);
  
  MailApp.sendEmail({
    to: partnerEmail,
    subject: assunto,
    htmlBody: mensagem,
    attachments: attachments,
    name: "Fatura ocultado"
  });
} catch (e) {
  Logger.log("Erro na função enviarEmailComPDF: " + e.toString());
  throw e;
}
}
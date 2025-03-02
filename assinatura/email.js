// Versão: 1.0
// Autor: Juliano Ceconi

function enviarEmails() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automsg');
    var data = sheet.getDataRange().getValues();
    
    var hoje = new Date();
    var emailEnviado = 0;
    var maxEmails = 100;  // Limite de e-mails para evitar exceder cotas diárias
    
    for (var i = 1; i < data.length; i++) {
      var linha = data[i];
      var nome = linha[1]; // Coluna B: 'nome'
      var email = linha[4];  // Coluna E: 'email'
      var dataAdesao = linha[5];  // Coluna F: 'Data de Adesão'
      var diaPagamento = linha[6];  // Coluna G: 'Dia de Pagamento'
      var email15DiasEnviado = linha[7];  // Coluna H: 'Último Vencimento Email 15 Dias'
      var emailNoDiaEnviado = linha[8];  // Coluna I: 'Último Vencimento Email No Dia'
      var valorParcela = linha[9];  // Coluna J: 'Valor Da Parcela'
      var parcelaPaga = linha[13];  // Coluna N: 'Parcela Atual Paga'
  
      // Verificação de valores obrigatórios
      if (!email || !valorParcela || !dataAdesao || !diaPagamento || !nome) {
        Logger.log("Dados incompletos para a linha " + (i + 1) + ". Nenhum e-mail enviado.");
        continue;
      }
  
      dataAdesao = new Date(dataAdesao);
      diaPagamento = Number(diaPagamento);
      if (isNaN(diaPagamento) || diaPagamento < 1 || diaPagamento > 31) {
        Logger.log("Dia de pagamento inválido para " + nome + ". Nenhum e-mail enviado.");
        continue;
      }
  
      // Verifica se o pagamento já foi realizado
      if (parcelaPaga && parcelaPaga.trim().toLowerCase() === "sim") {
        Logger.log("Pagamento já realizado para " + nome + ". Nenhum e-mail enviado.");
        continue;
      }
  
      // Calcula a próxima data de vencimento
      var nextDueDate = calculateNextDueDate(dataAdesao, diaPagamento);
  
      // Calcular a diferença entre a data de vencimento e hoje
      var diferencaDias = Math.ceil((nextDueDate - hoje) / (1000 * 60 * 60 * 24));
  
      try {
        // Verifica se deve enviar o e-mail de 15 dias antes
        if (diferencaDias === 15 && (!email15DiasEnviado || new Date(email15DiasEnviado) < nextDueDate) && emailEnviado < maxEmails) {
          if (isValidEmail(email)) {
            enviarEmail(
              email, 
              "ocultado", 
              "Prezado(a) " + nome + ",\n\nGostaríamos de informá-lo(a) que o pagamento da sua próxima parcela, no valor de " + valorParcela + ", está previsto para vencimento em 15 dias.\n\nEstamos à disposição para esclarecer qualquer dúvida ou fornecer mais informações.\n\nAtenciosamente,\nJuliano Ceconi\nContato: (77) 99827-2220"
            );
            sheet.getRange(i + 1, 8).setValue(nextDueDate);  // Atualiza 'Último Vencimento Email 15 Dias'
            emailEnviado++;
          }
        }
  
        // Verifica se deve enviar o e-mail no dia do vencimento
        if (diferencaDias === 0 && (!emailNoDiaEnviado || new Date(emailNoDiaEnviado) < nextDueDate) && emailEnviado < maxEmails) {
          if (isValidEmail(email)) {
            enviarEmail(
              email, 
              "ocultado", 
              "Prezado(a) " + nome + ",\n\nLembramos que o pagamento da sua parcela, no valor de " + valorParcela + ", vence hoje.\n\nPor favor, entre em contato caso precise de mais informações ou assistência.\n\nAtenciosamente,\nJuliano Ceconi\nContato: (77) 99827-2220"
            );
            sheet.getRange(i + 1, 9).setValue(nextDueDate);  // Atualiza 'Último Vencimento Email No Dia'
            emailEnviado++;
          }
        }
      } catch (e) {
        Logger.log("Erro ao enviar e-mail para " + nome + ": " + e.message);
      }
    }
  
    Logger.log("Envios concluídos: " + emailEnviado + " e-mails enviados.");
  }
  
  function calculateNextDueDate(dataAdesao, diaPagamento) {
    var hoje = new Date();
    var nextDueDate;
  
    // Ajusta para o dia de pagamento no mesmo mês ou no próximo
    if (dataAdesao.getDate() > diaPagamento) {
      nextDueDate = new Date(dataAdesao.getFullYear(), dataAdesao.getMonth() + 1, diaPagamento);
    } else {
      nextDueDate = new Date(dataAdesao.getFullYear(), dataAdesao.getMonth(), diaPagamento);
    }
  
    // Incrementa meses até que a próxima data de vencimento seja igual ou posterior a hoje
    while (nextDueDate < hoje) {
      nextDueDate = new Date(nextDueDate.getFullYear(), nextDueDate.getMonth() + 1, diaPagamento);
    }
  
    return nextDueDate;
  }
  
  function enviarEmail(destinatario, assunto, corpo) {
    MailApp.sendEmail(destinatario, assunto, corpo);
    Logger.log("E-mail enviado para: " + destinatario);
  }
  
  function isValidEmail(email) {
    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }
  
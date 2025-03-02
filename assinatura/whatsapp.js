// Versão: 1.0
// Autor: Juliano Ceconi

// Credenciais do Twilio
const ACCOUNT_SID = 'ocultado'; // Cole o Account SID do Twilio
const AUTH_TOKEN = 'ocultado';   // Cole o Auth Token do Twilio
const TWILIO_NUMBER = 'ocultado';  // Cole o número de WhatsApp do Twilio no formato +00000000000

// Função principal para enviar as mensagens
function sendWhatsAppMessages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  
  for (let i = 1; i < data.length; i++) { // Começa da linha 2 para ignorar o cabeçalho
    const titular = data[i][1]; // Nome do titular
    const numeroWpp = data[i][2]; // Número de WhatsApp
    const valorParcela = data[i][6]; // Valor da Parcela
    const diaVencimento = new Date(data[i][7]); // Dia do Vencimento
    
    // Verifica se o dia de vencimento é igual à data atual
    if (isSameDay(today, diaVencimento)) {
      const mensagem = `Olá ${titular}, lembrete de que a sua parcela no valor de R$ ${valorParcela} vence hoje. Por favor, efetue o pagamento o quanto antes.`;
      
      // Envia a mensagem via WhatsApp
      const resposta = sendWhatsAppMessage(numeroWpp, mensagem);
      Logger.log(`Mensagem enviada para ${numeroWpp}: ${resposta}`);
    }
  }
}

// Função auxiliar para comparar se duas datas são o mesmo dia
function isSameDay(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

// Função para enviar uma única mensagem via WhatsApp
function sendWhatsAppMessage(to, message) {
  const url = `https://api.twilio.com/2010-04-01/Accounts/${ACCOUNT_SID}/Messages.json`;
  
  const payload = {
    To: `whatsapp:${to}`,
    From: `whatsapp:${TWILIO_NUMBER}`,
    Body: message,
  };
  
  const options = {
    method: 'post',
    payload: payload,
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(ACCOUNT_SID + ':' + AUTH_TOKEN)
    },
  };
  
  const response = UrlFetchApp.fetch(url, options);
  return response.getContentText();
}
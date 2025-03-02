// arquivo: forms_codigo.gs
// versão: 1.1
// autor: Juliano Ceconi

function lancarContasNoSheets(beneficiado, motivo, dataInicial, numMeses, competenciaInicial, valor) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataParts = dataInicial.split("-");
    let anoInicial = parseInt(dataParts[0]);
    let mesInicial = parseInt(dataParts[1]) - 1; // Mês começa em 0 no JavaScript
    let diaVencimento = parseInt(dataParts[2]);

    for (let i = 0; i < numMeses; i++) {
      let dataVencimento = new Date(anoInicial, mesInicial + i, diaVencimento);
      
      // Ajuste para último dia do mês, se necessário
      if (dataVencimento.getDate() !== diaVencimento) {
        dataVencimento = new Date(anoInicial, mesInicial + i + 1, 0);
      }

      // Ajuste para dia útil anterior, se cair no fim de semana
      if (dataVencimento.getDay() === 0) {
        dataVencimento.setDate(dataVencimento.getDate() - 2); // Domingo, ajusta para sexta
      } else if (dataVencimento.getDay() === 6) {
        dataVencimento.setDate(dataVencimento.getDate() - 1); // Sábado, ajusta para sexta
      }

      let competencia = (competenciaInicial + i - 1) % 12 + 1;
      competencia = competencia < 10 ? '0' + competencia : competencia.toString();

      const linha = [
        "", // Coluna A em branco
        beneficiado, // Coluna B
        motivo, // Coluna C
        competencia, // Coluna D
        valor, // Coluna E
        Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy"), // Coluna F
        "", // Coluna G em branco
        "" // Coluna H em branco
      ];
      sheet.appendRow(linha);
    }

    SpreadsheetApp.getUi().alert("Contas lançadas com sucesso.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Falha ao lançar contas: " + e.message);
  }
}

function abrirFormulario() {
  const html = HtmlService.createHtmlOutputFromFile('formulario_lancamento_lote.html')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Lançar Contas a Pagar');
}

abrirFormulario

// Função para ser chamada no HTML
function lancarContas(beneficiado, motivo, dataInicial, numMeses, competenciaInicial, valor) {
  lancarContasNoSheets(beneficiado, motivo, dataInicial, numMeses, competenciaInicial, valor);
}
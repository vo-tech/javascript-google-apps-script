// arquivo: caixa.gs
// versão: 1.5
// autor: Juliano Ceconi

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ocultado')
    .addItem('Nova Guia', 'showForm')
    .addItem('Editar Guia', 'openEditForm')
    .addItem('Lançar Despesa', 'abrirFormulario')
    .addToUi();
}

function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('formularioGuia')
    .setWidth(500)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Nova Guia');
}

function openEditForm() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Editar Guia', 'Digite o número da Guia:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const guiaId = response.getResponseText();
    
    try {
      const html = HtmlService.createTemplateFromFile('editarGuia');
      html.guiaId = guiaId; // Passa o guiaId para o template
      
      const output = html.evaluate()
        .setWidth(600)
        .setHeight(800)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
      SpreadsheetApp.getUi().showModalDialog(output, `Editando Guia ${guiaId}`);
    } catch (error) {
      ui.alert('ERRO', error.message, ui.ButtonSet.OK);
    }
  }
}

function getGuiaById(guiaId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('guias');
  
  // Debug: Verificar dados brutos
  console.log('ID Recebido:', guiaId, 'Tipo:', typeof guiaId);
  
  const data = sheet.getDataRange().getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Busca otimizada
  const row = data.find(row => {
    const rowId = row[0]?.toString().trim().toLowerCase();
    const searchId = guiaId.toString().trim().toLowerCase();
    return rowId === searchId;
  });

  // Debug: Log de IDs existentes
  console.log('IDs na planilha:', data.slice(1).map(r => r[0]?.toString().trim()));

  if (!row) throw new Error(`Guia "${guiaId}" não encontrada. Verifique o ID.`);
  
  // Mapeamento com verificação
  return headers.reduce((obj, header, index) => {
    const key = header.normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, '');
    obj[key] = row[index] ?? '';
    return obj;
  }, {});
}

function processForm(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('guias');
    const expectedHeaders = ['Guia', 'Auditado em', 'TIPO', 'FILTRO', 'Responsável', 'Cidade', 'Emissão', 'Data Guia', 
                            'Forma de Recebimento', 'Valor Recebido', 'Valor de repasse', 'Valor comissão', 'procedimento', 
                            'Instituição', 'Tipo Instituição', 'Data Repasse', 'Data NF', 'Competência', 'Paciente', 
                            'Conferência', 'Cobertura', 'dif', 'Categoria'];

    // Validações
    if (!sheet) throw new Error('Planilha "guias" não encontrada!');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.every((h, i) => h === expectedHeaders[i])) {
      throw new Error('Estrutura da planilha alterada!');
    }

    if (!formData.Guia || !formData.Responsavel) {
      throw new Error('Preencha Guia e Responsável!');
    }

    // Construção da linha
    const rowData = headers.map(header => {
      const fieldId = header.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '');
      const value = formData[fieldId] || '';

      // Conversão de dados
      if (/Data|Emissão/.test(header)) {
        return value ? new Date(value + 'T00:00:00-03:00') : '';
      }

      if (/Valor|dif/.test(header)) {
        return value ? Number(value.toString().replace(',', '.')) : '';
      }

      return value;
    });

    sheet.appendRow(rowData);
    return true;

  } catch (error) {
    console.error(`[${new Date().toISOString()}] ERRO: ${error.message}`);
    throw error;
  }
}

function updateGuia(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('guias');
    const data = sheet.getDataRange().getValues();
    
    // Encontra a guia
    const rowIndex = data.findIndex(row => row[0] === formData.Guia);
    if (rowIndex === -1) throw new Error('Guia não encontrada!');

    // Atualização
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const updatedRow = headers.map((header, index) => {
      const fieldId = header.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '');
      return formData[fieldId] || data[rowIndex][index];
    });

    sheet.getRange(rowIndex + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
    return true;

  } catch (error) {
    console.error(`[${new Date().toISOString()}] ERRO: ${error.message}`);
    throw error;
  }
}
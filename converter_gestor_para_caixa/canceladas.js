// arquivo: canceladas.gs
// versão: 1.0
// autor: Juliano Ceconi

// Adiciona o menu customizado "ocultado" à planilha ativa
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('ocultado')
      .addItem('Processar Canceladas', 'processarCanceladas')
      .addToUi();
  }
  
  /**
   * Função principal para processar as guias canceladas.
   * Ela faz o seguinte:
   * - Abre a planilha "Canceladas" e varre as linhas a partir da terceira (com cabeçalho na 2).
   * - Ignora linhas onde a coluna "Lançado" (coluna H) já esteja preenchida.
   * - Para cada linha a ser processada, busca a guia correspondente na planilha "Caixa" > "guias".
   * - Se encontrada, atualiza as colunas J, K e P (para 0, 0 e "C", respectivamente).
   * - Se não encontrada, marca a linha na planilha "Canceladas" com preenchimento em amarelo.
   * - Ao final, atualiza a coluna "Lançado" com a letra "C" para guias processadas com sucesso e informa por alerta os resultados.
   */
  function processarCanceladas() {
    // IDs das planilhas
    var idPlanilhaCanceladas = 'ocultado';
    var idPlanilhaCaixa = 'ocultado';
  
    // Abre as planilhas
    var ssCanceladas = SpreadsheetApp.openById(idPlanilhaCanceladas);
    var sheetCanceladas = ssCanceladas.getSheetByName('Canceladas');
    
    var ssCaixa = SpreadsheetApp.openById(idPlanilhaCaixa);
    var sheetGuias = ssCaixa.getSheetByName('guias');
    
    // Obtém os dados da planilha 'Canceladas'
    // Cabeçalho está na linha 2, dados a partir da linha 3:
    var lastRowCanceladas = sheetCanceladas.getLastRow();
    if(lastRowCanceladas < 3){
      SpreadsheetApp.getUi().alert('Não há dados para processar na planilha Canceladas.');
      return;
    }
    
    // Considerando 8 colunas: A: Guia, B: Valor, C: Usuário, D: Emissão, E: Cancelamento, F: Paciente, G: Motivo, H: Lançado
    var rangeCanceladas = sheetCanceladas.getRange(3, 1, lastRowCanceladas - 2, 8);
    var dataCanceladas = rangeCanceladas.getValues();
    
    // Preparar arrays para atualizações na planilha Canceladas
    var updateLancado = [];
    var updateBackgrounds = [];
    
    // Armazenar qual índice de linha (local na matriz e planilha) terão que atualizar "Canceladas"
    // Valores de alerta
    var sucessoCount = 0;
    var naoEncontradoCount = 0;
    
    // Para evitar leituras repetidas da planilha "guias", vamos definir a faixa da coluna A (Guia) a partir da linha 2 até a última linha
    // Assim poderemos usar o método createTextFinder procurando na coluna A apenas
    var lastRowGuias = sheetGuias.getLastRow();
    var rangeGuiasColA = sheetGuias.getRange(2, 1, lastRowGuias - 1, 1);
    
    // Precaução: armazenar atualizações para a planilha "guias" em um array de objetos com a linha a ser atualizada
    var guiasParaAtualizar = [];
    
    // Percorre cada linha dos dados da planilha "Canceladas"
    for (var i = 0; i < dataCanceladas.length; i++) {
      var linha = dataCanceladas[i];
      var guiaCancelada = linha[0]; // Coluna A: Guia
      var lancado = linha[7];       // Coluna H: Lançado
      var linhaPlanilha = i + 3;     // A linha real na planilha Canceladas
    
      // Se "Lançado" já estiver preenchido, mantém os valores originais
      if (lancado !== "") {
        updateLancado.push([lancado]);
        updateBackgrounds.push([null]);
        continue;
      }
    
      // Procura a guia na planilha "guias" usando createTextFinder na coluna A
      var textFinder = rangeGuiasColA.createTextFinder(String(guiaCancelada));
      var cellEncontrada = textFinder.findNext();
      
      if (cellEncontrada) {
        // Guia encontrada, guarda a linha onde se encontra para atualizar as colunas J, K e P.
        var linhaGuia = cellEncontrada.getRow();
        guiasParaAtualizar.push({linha: linhaGuia});
    
        // Marca como processada na planilha Canceladas: "Lançado" recebe "C"
        updateLancado.push(["C"]);
        updateBackgrounds.push([null]); // Nenhuma alteração de cor, mantendo padrão
        sucessoCount++;
      } else {
        // Guia não encontrada: atualiza "Lançado" permanece em branco e marca a linha com fundo amarelo
        updateLancado.push([""]);
        updateBackgrounds.push([ "#FFFF00" ]);
        naoEncontradoCount++;
      }
    }
    
    // Atualiza a coluna "Lançado" na planilha Canceladas de uma única vez
    sheetCanceladas.getRange(3, 8, updateLancado.length, 1).setValues(updateLancado);
    
    // Atualiza a formatação (fundo amarelo) somente para as linhas não encontradas
    // Se a cor for definida como null, ignora.
    for (var j = 0; j < updateBackgrounds.length; j++) {
      if (updateBackgrounds[j][0] !== null) {
        sheetCanceladas.getRange(j + 3, 1, 1, sheetCanceladas.getLastColumn()).setBackground(updateBackgrounds[j][0]);
      }
    }
    
    // Atualiza as colunas na planilha "guias" para cada guia encontrada:
    // Em "guias", atualizar:
    // Coluna J (10) -> 0, Coluna K (11) -> 0, Coluna P (16) -> "C"
    guiasParaAtualizar.forEach(function(obj) {
      var linha = obj.linha;
      // Atualiza as 3 colunas de uma vez
      sheetGuias.getRange(linha, 10, 1, 1).setValue(0); // Coluna J
      sheetGuias.getRange(linha, 11, 1, 1).setValue(0); // Coluna K
      sheetGuias.getRange(linha, 16, 1, 1).setValue("C"); // Coluna P
    });
    
    // Exibe alerta com os resultados
    var mensagem = 'Processamento concluído.\n' +
                   'Guias atualizadas com sucesso: ' + sucessoCount + '\n' +
                   'Guias não encontradas (marcadas em amarelo): ' + naoEncontradoCount;
    SpreadsheetApp.getUi().alert(mensagem);
  }
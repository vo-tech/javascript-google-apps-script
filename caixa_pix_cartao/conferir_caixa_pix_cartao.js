// arquivo: conferir_guias_com_status_em_branco.gs
// versão: 2.5
// autor: Juliano Ceconi

function compareAndUpdateChunked() {
  Logger.log("===== [INÍCIO] Função compareAndUpdate (chunked B) =====");
  var startTime = new Date();
  
  try {
    Logger.log("Hora inicial de execução: " + startTime.toISOString());

    //-------------------------------------------------------------
    // CONFIGURAÇÕES E REFERÊNCIAS
    //-------------------------------------------------------------
    var planilhaAId = 'ocultado';
    var abaA = 'Respostas';
    var planilhaBId = 'ocultado';
    var abaB = 'guias';

    // Abas
    var ssA = SpreadsheetApp.openById(planilhaAId).getSheetByName(abaA);
    var ssB = SpreadsheetApp.openById(planilhaBId).getSheetByName(abaB);

    // Índices zero-based das colunas na Planilha A:
    // A=0(Data), C=2(Guia), F=5(Valor), J=9(Status)
    var COL_DATA_A   = 0;
    var COL_GUIA_A   = 2;
    var COL_VALOR_A  = 5;
    var COL_STATUS_A = 9;

    // Índices zero-based das colunas na Planilha B:
    // A=0(Guia), B=1(DataProc), J=9(Valor)
    var COL_GUIA_B       = 0; 
    var COL_DATA_PROC_B  = 1;
    var COL_VALOR_B      = 9; 

    //-------------------------------------------------------------
    // VERIFICAÇÃO DE TAMANHO
    //-------------------------------------------------------------
    var lastRowA = ssA.getLastRow();
    if (lastRowA < 2) {
      Logger.log("Planilha A sem dados além do cabeçalho. Encerrando.");
      return;
    }
    var lastRowB = ssB.getLastRow();
    if (lastRowB < 2) {
      Logger.log("Planilha B sem dados além do cabeçalho. Encerrando.");
      return;
    }

    Logger.log("Planilha A -> Última linha: " + lastRowA);
    Logger.log("Planilha B -> Última linha: " + lastRowB);

    //-------------------------------------------------------------
    // LEITURA DA PLANILHA A (colunas A até J)
    //-------------------------------------------------------------
    Logger.log("Lendo colunas A até J da Planilha A...");
    var rangeA = ssA.getRange(1, 1, lastRowA, 10).getValues();

    //-------------------------------------------------------------
    // CRIA DICIONÁRIO DE GUIAS A (SÓ ONDE STATUS ESTÁ VAZIO)
    //-------------------------------------------------------------
    var guiasA = {};
    var linhasVazias = 0;
    for (var r = 1; r < rangeA.length; r++) {
      var statusAtual = rangeA[r][COL_STATUS_A];
      if (statusAtual === '' || statusAtual === null || statusAtual === undefined) {
        linhasVazias++;
        var guia = rangeA[r][COL_GUIA_A];
        if (guia) {
          if (!guiasA[guia]) {
            guiasA[guia] = [];
          }
          guiasA[guia].push({
            rowA: r, // índice na matriz rangeA
            valorA: rangeA[r][COL_VALOR_A],
            dataA:  rangeA[r][COL_DATA_A]
          });
        }
      }
    }

    var totalGuiasParaProcessar = Object.keys(guiasA).length;
    Logger.log("Total de linhas com status vazio em A: " + linhasVazias);
    Logger.log("Guia(s) únicas em A com status vazio: " + totalGuiasParaProcessar);

    if (totalGuiasParaProcessar === 0) {
      Logger.log("Nenhuma guia com status vazio em A. Encerrando.");
      return;
    }

    //-------------------------------------------------------------
    // PROCESSAR PLANILHA B EM CHUNKS
    //-------------------------------------------------------------
    var chunkSizeB = 2000; 
    var totalProcessadas = 0;
    var totalDivergentes = 0;

    Logger.log("Iniciando loop em B com chunks de " + chunkSizeB + " linhas...");

    function formataDataDDMMYYYY(d) {
      if (!d) return '';
      var dt = new Date(d);
      if (isNaN(dt)) return '';
      var dia = dt.getDate();
      var mes = dt.getMonth() + 1;
      var ano = dt.getFullYear();
      return (dia < 10 ? '0'+dia : dia) + '/' + (mes<10?'0'+mes:mes) + '/' + ano;
    }

    for (var bStart = 1; bStart <= lastRowB; bStart += chunkSizeB) {
      var bEnd = Math.min(bStart + chunkSizeB - 1, lastRowB);
      Logger.log(" -- Lendo chunk B de " + bStart + " até " + bEnd);

      // Lemos 3 colunas separadamente
      var colGuiaRange = ssB.getRange(bStart, 1, (bEnd - bStart + 1), 1).getValues();
      var colDataProcRange = ssB.getRange(bStart, 2, (bEnd - bStart + 1), 1).getValues();
      var colValorRange = ssB.getRange(bStart, 10, (bEnd - bStart + 1), 1).getValues();

      // Percorrer esse bloco de linhas de B
      for (var i = 0; i < colGuiaRange.length; i++) {
        var guiaB = colGuiaRange[i][0];
        if (!guiaB) continue;

        // Verifica se esse guia está no dicionário guiasA
        if (guiasA.hasOwnProperty(guiaB)) {
          var valorB = colValorRange[i][0];
          
          // Para cada linha de A associada a esse guia...
          var entradasA = guiasA[guiaB];
          for (var k = 0; k < entradasA.length; k++) {
            var rowAIndex = entradasA[k].rowA;
            var valorA = entradasA[k].valorA;
            var dataA  = entradasA[k].dataA;

            // Compara valores
            if (Number(valorA) === Number(valorB)) {
              rangeA[rowAIndex][COL_STATUS_A] = 'OK';
              colDataProcRange[i][0] = formataDataDDMMYYYY(dataA);
            } else {
              rangeA[rowAIndex][COL_STATUS_A] = 'Divergente';
              colDataProcRange[i][0] = formataDataDDMMYYYY(dataA);
              totalDivergentes++;
            }
            totalProcessadas++;
          }
        }
      }

      // Atualiza coluna B (DataProc) da Planilha B
      Logger.log(" -- Atualizando col B (data proc) da Planilha B para linhas " + bStart + " a " + bEnd);
      ssB.getRange(bStart, 2, (bEnd - bStart + 1), 1).setValues(colDataProcRange);
    }

    //-------------------------------------------------------------
    // Atualiza Planilha A de uma só vez
    //-------------------------------------------------------------
    Logger.log("Atualizando Planilha A...");
    ssA.getRange(1, 1, rangeA.length, 10).setValues(rangeA);
    Logger.log("Planilha A atualizada.");

    //-------------------------------------------------------------
    // RELATÓRIO FINAL
    //-------------------------------------------------------------
    var msg = "Processamento concluído.\n" +
              "Total de guias processadas: " + totalProcessadas + "\n" +
              "Total de guias divergentes: " + totalDivergentes + "\n" +
              "Total de linhas com status vazio em A: " + linhasVazias;
    Logger.log(msg);
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Relatório", 10);

  } catch (err) {
    Logger.log("[ERRO FATAL] Ocorreu um erro geral: " + err);
  } finally {
    var endTime = new Date();
    var elapsed = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log("Tempo total de execução: " + elapsed + " segundos.");
    Logger.log("===== [FIM] Função compareAndUpdate (chunked B) =====");
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Cria menu de ocultado
  ui.createMenu('ocultado')
    .addItem('Consolidar na planilha Caixa', 'compareAndUpdateChunked')
    .addToUi();
}
// arquivo: criarPlanilhas.gs
// Versão: 1.0
// Autor: Juliano Ceconi


function criarPlanilhasParceiros() {
    var lock = LockService.getScriptLock();
    try {
      // Tenta adquirir o bloqueio por até 5 segundos
      lock.waitLock(5000);
    } catch (e) {
      Logger.log('Não foi possível adquirir o bloqueio. Outra instância do script está em execução.');
      return;
    }
  
    var TEMPO_LIMITE_EXECUCAO = 280000; // Tempo limite em milissegundos (aprox. 4,6 minutos)
    var TEMPO_INICIO = new Date().getTime();
  
    try {
      // Obter a planilha ativa e a aba atual
      var planilhaAtual = SpreadsheetApp.getActiveSpreadsheet();
      var abaAtual = planilhaAtual.getActiveSheet();
  
      // Obter a lista de parceiros da coluna A e verificar a coluna B
      var ultimaLinha = abaAtual.getLastRow();
      if (ultimaLinha < 2) {
        throw new Error('A lista de parceiros está vazia.');
      }
      var dados = abaAtual.getRange(2, 1, ultimaLinha - 1, 2).getValues(); // Colunas A e B
      var parceiros = []; // Lista de parceiros a processar
      for (var i = 0; i < dados.length; i++) {
        var nomeParceiro = dados[i][0];
        var urlPlanilha = dados[i][1];
        if (nomeParceiro && !urlPlanilha) {
          parceiros.push({ nome: nomeParceiro, linha: i + 2 });
        }
      }
  
      if (parceiros.length === 0) {
        Logger.log('Não há parceiros para processar.');
        return;
      }
  
      // Recuperar o índice do último parceiro processado (para pausa e retomada)
      var propriedades = PropertiesService.getDocumentProperties();
      var indiceInicio = parseInt(propriedades.getProperty('ultimoIndice')) || 0;
  
      // Definir IDs e nomes das planilhas
      var idPlanilhaMestre = 'ocultado';
      var nomeAbaMestre = 'guias';
  
      // Definir os nomes das colunas para o cabeçalho
      var nomesColunas = ['Guia', 'Emissão', 'Data Guia', 'Valor de repasse', 'Procedimento', 'Instituição', 'Data Repasse', 'Paciente'];
  
      // Iniciar o log
      var log = [];
      log.push('Início da execução: ' + new Date());
  
      // Iterar pelos parceiros a partir do último índice processado
      for (var i = indiceInicio; i < parceiros.length; i++) {
        var parceiro = parceiros[i].nome;
        var linhaParceiro = parceiros[i].linha;
  
        // Verificar se o tempo de execução está próximo do limite
        if (new Date().getTime() - TEMPO_INICIO > TEMPO_LIMITE_EXECUCAO) {
          log.push('Tempo limite próximo. Execução pausada no parceiro: ' + parceiro);
          propriedades.setProperty('ultimoIndice', i);
          break;
        }
  
        // Verificar se o nome do parceiro não está vazio
        if (!parceiro) {
          log.push('Nome do parceiro vazio na linha ' + linhaParceiro);
          continue;
        }
  
        // Nomear a nova planilha
        var nomeNovaPlanilha = 'Transparência - ' + parceiro + ' - MedPless';
  
        try {
          // Criar nova planilha
          var novaPlanilha = SpreadsheetApp.create(nomeNovaPlanilha);
          log.push('Planilha criada: ' + nomeNovaPlanilha);
  
          // Adicionar o cabeçalho
          var abaNova = novaPlanilha.getActiveSheet();
          abaNova.getRange(1, 1, 1, nomesColunas.length).setValues([nomesColunas]);
  
          // Inserir a fórmula QUERY na célula A2
          var formula = "=QUERY(IMPORTRANGE('" + idPlanilhaMestre + "', '" + nomeAbaMestre + "!A1:T'), " +
                        "'SELECT Col1, Col7, Col8, Col11, Col13, Col14, Col16, Col19 WHERE Col14 = '''" + 
                        parceiro + "'''', 1)";
          abaNova.getRange('A2').setFormula(formula);
  
          log.push('Fórmula inserida na planilha: ' + nomeNovaPlanilha);
  
          // Escrever o URL na coluna B da planilha atual
          abaAtual.getRange(linhaParceiro, 2).setValue(novaPlanilha.getUrl());
  
          // Atualizar o índice do último parceiro processado
          propriedades.setProperty('ultimoIndice', i + 1);
  
        } catch (erroCriacao) {
          log.push('Erro ao criar a planilha "' + nomeNovaPlanilha + '": ' + erroCriacao.message);
          continue;
        }
      }
  
      log.push('Execução concluída em ' + new Date());
  
      // Limpar o índice após conclusão
      if (i >= parceiros.length) {
        propriedades.deleteProperty('ultimoIndice');
      }
  
    } catch (e) {
      log.push('Erro: ' + e.message);
    } finally {
      registrarLog(log);
      lock.releaseLock();
    }
  }
  
  function registrarLog(log) {
    var propriedades = PropertiesService.getDocumentProperties();
    var arquivoLogId = propriedades.getProperty('arquivoLogId');
    var arquivoLog;
  
    if (arquivoLogId) {
      arquivoLog = DriveApp.getFileById(arquivoLogId);
    } else {
      arquivoLog = DriveApp.createFile('Log de Execução - ' + new Date().toISOString(), '');
      propriedades.setProperty('arquivoLogId', arquivoLog.getId());
    }
  
    var textoLog = log.join('\n') + '\n';
    var conteudoAtual = arquivoLog.getBlob().getDataAsString();
    arquivoLog.setContent(conteudoAtual + textoLog);
  }
  
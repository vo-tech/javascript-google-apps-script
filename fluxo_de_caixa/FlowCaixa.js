/**
 * FlowCaixa.gs
 *
 * Este módulo unificado engloba funções para:
 *  - Formatação das abas
 *  - Registro de lançamentos
 *  - Atualização de repasses
 *  - Geração de backup diário
 *  - Listagem de lançamentos
 *
 * As funções são expostas globalmente para serem chamadas via menu ou pela Sidebar.
 * Adota-se encapsulamento por IIFE para evitar conflitos no escopo global.
 */

(function (global) {
  // Cache para SpreadsheetApp
  let _spreadsheetCache = null;
  let _sheetsCache = {};
  
  /* CONFIGURAÇÃO =========================================================== */
  const CONFIG_FLUXO = {
    NOME_SHEET: "Fluxo de Caixa", // Nome da aba de fluxo de caixa
    CABECALHO_LINHAS: 1,          // Número de linhas do cabeçalho
    COLUNAS: [
      "Data", "Descrição", "Tipo", "Categoria", "Valor", 
      "Forma de Pagamento", "Conta", "Sede", "Observações"
    ]
  };
  
  const CONFIG_FORMATA = {
    // Configuração para formatação das abas (nome e largura de colunas)
    abas: [
      { nome: "Dashboard/Relatórios", larguraColunas: [120, 120, 100, 100, 100] },
      { nome: CONFIG_FLUXO.NOME_SHEET, larguraColunas: [100, 250, 100, 150, 100, 180, 100, 120, 200] },
      { nome: "Contas e CNPJs", larguraColunas: [180, 200, 120, 300] },
      { nome: "Categorias", larguraColunas: [150, 300] },
      { nome: "Configurações e Parâmetros", larguraColunas: [250, 250] },
      { nome: "Histórico de Lançamentos", larguraColunas: [100, 250, 100, 150, 100, 180, 100, 120, 200] }
    ]
  };
  
  /* FUNÇÕES DE FLUXO DE CAIXA ============================================== */
  
  /**
   * Função utilitária para acessar planilha com cache
   */
  function getSpreadsheet() {
    if (!_spreadsheetCache) {
      _spreadsheetCache = SpreadsheetApp.getActiveSpreadsheet();
    }
    return _spreadsheetCache;
  }
  
  /**
   * Função utilitária para acessar abas com cache
   */
  function getSheet(nomePlanilha) {
    if (!_sheetsCache[nomePlanilha]) {
      _sheetsCache[nomePlanilha] = getSpreadsheet().getSheetByName(nomePlanilha);
    }
    return _sheetsCache[nomePlanilha];
  }
  
  /**
   * Função otimizada para registrar múltiplos lançamentos
   */
  function registrarLancamentos(lancamentos) {
    try {
      const sheet = getSheet(CONFIG_FLUXO.NOME_SHEET);
      if (!sheet) throw new Error(`Aba '${CONFIG_FLUXO.NOME_SHEET}' não encontrada`);
      
      const dados = lancamentos.map(l => [
        l.data || "", l.descricao || "", l.tipo || "",
        l.categoria || "", l.valor || 0, l.formaPagamento || "",
        l.conta || "", l.sede || "", l.observacoes || ""
      ]);
      
      // Inserção em lote
      if (dados.length) {
        sheet.getRange(sheet.getLastRow() + 1, 1, dados.length, dados[0].length)
             .setValues(dados);
      }
      
      return true;
    } catch (e) {
      console.error(`Erro em registrarLancamentos: ${e.message}`);
      return false;
    }
  }
  
  /**
   * Versão otimizada para um único lançamento
   */
  function registrarLancamento(lancamento) {
    try {
      Logger.info('registrarLancamento', 'Iniciando registro de lançamento', lancamento);
      
      const resultado = registrarLancamentos([lancamento]);
      if (resultado) {
        Logger.info('registrarLancamento', 'Lançamento registrado com sucesso', lancamento);
        return { success: true, message: "Lançamento registrado com sucesso!" };
      } else {
        throw new Error("Falha ao registrar lançamento");
      }
    } catch (error) {
      Logger.error('registrarLancamento', 'Erro ao registrar lançamento', error);
      return { success: false, message: "Erro ao registrar lançamento: " + error.message };
    }
  }
  
  /**
   * Atualiza o valor de repasses com base no total de receitas,
   * usando o percentual informado (padrão: 83%).
   *
   * @param {number} percentual - Percentual de repasse.
   * @return {number} - Valor total a repassar.
   */
  function atualizarRepasses(percentual = 83) {
    try {
      // Adicionar validação do percentual
      if (typeof percentual !== 'number' || percentual < 0 || percentual > 100) {
        throw new Error('Percentual deve ser um número entre 0 e 100');
      }
      const valor = calcularRepasses(percentual);
      return {
        success: true,
        message: `Repasses calculados: R$ ${valor.toFixed(2)} (${percentual}%)`,
        valor: valor
      };
    } catch (error) {
      return { success: false, message: "Erro ao calcular repasses: " + error.message };
    }
  }
  
  /**
   * Gera um backup dos lançamentos e envia por email
   * @param {boolean} silencioso - Se verdadeiro, não exibe mensagens de log (usado no trigger)
   * @return {Object} - Objeto com status e mensagem
   */
  function backupDiario(silencioso = false) {
    try {
      console.log('Iniciando backup...'); // Log adicional
      
      const EMAIL_DESTINO = 'ocultado';
      const sheet = getSheet(CONFIG_FLUXO.NOME_SHEET);
      if (!sheet) throw new Error(`Aba '${CONFIG_FLUXO.NOME_SHEET}' não encontrada`);
      
      console.log('Sheet encontrada, gerando backup...'); // Log adicional
      
      // Obtém a planilha atual
      const ss = getSpreadsheet();
      
      // Cria uma cópia temporária para o backup
      const tempFile = DriveApp.createFile(
        ss.getBlob().setName(`Backup_Fluxo_${formatarDataHora()}.xlsx`)
      );
      
      // Prepara os dados para o email
      const dataHoje = Utilities.formatDate(
        new Date(), 
        Session.getScriptTimeZone(), 
        "dd/MM/yyyy 'às' HH:mm:ss"
      );
      
      const totalRegistros = sheet.getLastRow() - CONFIG_FLUXO.CABECALHO_LINHAS;
      
      // Monta o email
      const emailHtml = `
        <h2>Backup Automático - Fluxo de Caixa</h2>
        <p>Backup gerado em ${dataHoje}</p>
        <p>Total de registros: ${totalRegistros}</p>
        <br>
        <p><i>Este é um email automático, por favor não responda.</i></p>
      `;
      
      // Envia o email com o arquivo anexado
      MailApp.sendEmail({
        to: EMAIL_DESTINO,
        subject: `Backup Fluxo de Caixa - ${formatarData()}`,
        htmlBody: emailHtml,
        attachments: [tempFile.getAs(MimeType.MICROSOFT_EXCEL)],
        noReply: true
      });
      
      // Remove o arquivo temporário
      tempFile.setTrashed(true);
      
      console.log('Email enviado com sucesso'); // Log adicional
      
      return {
        success: true,
        message: `Backup enviado com sucesso para ${EMAIL_DESTINO}`
      };
      
    } catch (error) {
      console.error('Erro no backup:', error); // Log adicional
      return {
        success: false,
        message: `Erro ao gerar/enviar backup: ${error.message}`
      };
    }
  }
  
  /**
   * Funções auxiliares para formatação de data
   */
  function formatarData() {
    return Utilities.formatDate(
      new Date(), 
      Session.getScriptTimeZone(), 
      "dd-MM-yyyy"
    );
  }
  
  function formatarDataHora() {
    return Utilities.formatDate(
      new Date(), 
      Session.getScriptTimeZone(), 
      "yyyyMMdd_HHmmss"
    );
  }
  
  /**
   * Retorna uma matriz com todos os lançamentos da aba "Fluxo de Caixa",
   * inclusive o cabeçalho.
   *
   * @return {Array[]} - Matriz de dados da planilha.
   */
  function listarLancamentos() {
    try {
      const dados = getSheet(CONFIG_FLUXO.NOME_SHEET).getDataRange().getValues();
      return {
        success: true,
        message: `${dados.length - 1} lançamentos encontrados`,
        dados: dados
      };
    } catch (error) {
      return { success: false, message: "Erro ao listar lançamentos: " + error.message };
    }
  }
  
  /* FUNÇÕES DE FORMATAÇÃO =============================================== */
  
  /**
   * Formata as abas de acordo com a configuração definida em CONFIG_FORMATA.
   */
  function formatarAbas() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      CONFIG_FORMATA.abas.forEach(function (abaConfig) {
        const sheet = ss.getSheetByName(abaConfig.nome);
        if (sheet) {
          // Formata a linha de cabeçalho
          const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          headerRange.setBackground("#4CAF50")
            .setFontColor("#FFFFFF")
            .setFontWeight("bold")
            .setHorizontalAlignment("center");
          
          // Formata os dados (se houver)
          if (sheet.getLastRow() > 1) {
            const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
            dataRange.setFontFamily("Arial")
              .setFontSize(10)
              .setHorizontalAlignment("left");
          }
          
          // Ajusta a largura das colunas
          if (abaConfig.larguraColunas && abaConfig.larguraColunas.length) {
            abaConfig.larguraColunas.forEach(function (largura, index) {
              sheet.setColumnWidth(index + 1, largura);
            });
          }
          
          // Congela a linha de cabeçalho
          sheet.setFrozenRows(1);
          Logger.log("Aba formatada: " + abaConfig.nome);
        } else {
          Logger.log("Aba não encontrada: " + abaConfig.nome);
        }
      });
      return { success: true, message: "Todas as abas foram formatadas com sucesso!" };
    } catch (error) {
      Logger.log("Erro na formatação: " + error);
      return { success: false, message: "Erro ao formatar abas: " + error.message };
    }
  }
  
  /**
   * Inicializa o aplicativo: formata as abas e abre a Sidebar.
   */
  function inicializarApp() {
    try {
      formatarAbas();
      abrirSidebarUI();
      Logger.log("Aplicação inicializada com sucesso.");
    } catch (error) {
      Logger.log("Erro na inicialização: " + error);
      throw error;
    }
  }
  
  /* FUNÇÕES DE INTERFACE (UI) ============================================ */
  
  /**
   * Abre a Sidebar com uma interface moderna para o fluxo de caixa.
   */
  function abrirSidebarUI() {
    try {
      const htmlOutput = HtmlService.createTemplateFromFile("Sidebar")
        .evaluate()
        .setTitle("Painel de Controle - Fluxo de Caixa")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi().showSidebar(htmlOutput);
      Logger.log("Sidebar aberta.");
    } catch (error) {
      Logger.log("Erro ao abrir Sidebar: " + error);
      throw error;
    }
  }
  
  /* EXPOSTAÇÃO DAS FUNÇÕES =============================================== */
  global.registrarLancamento = registrarLancamento;
  global.atualizarRepasses = atualizarRepasses;
  global.backupDiario = backupDiario;
  global.listarLancamentos = listarLancamentos;
  global.formatarAbas = formatarAbas;
  global.inicializarApp = inicializarApp;
  global.configurarTriggerBackup = configurarTriggerBackup;
  global.registrarLogsCliente = registrarLogsCliente;
  
})(this);


/**
 * onOpen - Cria um menu personalizado no Google Sheets ao abrir a planilha.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Fluxo de Caixa")
    .addItem("Iniciar App", "inicializarApp")
    .addToUi();
}

/**
 * Configura o trigger diário para backup
 */
function configurarTriggerBackup() {
  try {
    // Remove triggers existentes para evitar duplicidade
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'backupDiario') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Cria novo trigger para 23:00
    ScriptApp.newTrigger('backupDiario')
      .timeBased()
      .atHour(23)
      .everyDays(1)
      .create();
    
    return true;
  } catch (error) {
    Logger.log("Erro ao configurar trigger: " + error.message);
    return false;
  }
}

// Adicionar função que está faltando
function calcularRepasses(percentual) {
  // Implementar lógica de cálculo
  const sheet = getSheet(CONFIG_FLUXO.NOME_SHEET);
  // ... lógica de cálculo ...
  return valorCalculado;
}

/* SISTEMA DE LOGGING ================================================== */
const Logger = {
  PASTA_LOGS: 'ocultado', // Substitua pelo ID da pasta onde deseja salvar os logs
  
  async log(tipo, funcao, mensagem, dados = null) {
    try {
      const timestamp = new Date().toISOString();
      const logEntry = {
        timestamp,
        tipo,
        funcao,
        mensagem,
        dados: dados ? JSON.stringify(dados) : null,
        usuario: Session.getActiveUser().getEmail()
      };

      // Formata a entrada do log
      const logLine = `[${timestamp}] ${tipo} | ${funcao} | ${mensagem} | ${logEntry.usuario}\n`;
      if (dados) logLine += `Dados: ${JSON.stringify(dados)}\n`;

      // Nome do arquivo de log baseado na data atual
      const hoje = new Date().toISOString().split('T')[0];
      const nomeArquivo = `log_fluxocaixa_${hoje}.txt`;

      // Tenta encontrar arquivo de log existente ou cria um novo
      let arquivo;
      const pasta = DriveApp.getFolderById(this.PASTA_LOGS);
      const arquivos = pasta.getFilesByName(nomeArquivo);

      if (arquivos.hasNext()) {
        arquivo = arquivos.next();
        const conteudoAtual = arquivo.getBlob().getDataAsString();
        arquivo.setContent(conteudoAtual + logLine);
      } else {
        arquivo = pasta.createFile(nomeArquivo, logLine);
      }

      return true;
    } catch (error) {
      console.error('Erro ao registrar log:', error);
      return false;
    }
  },

  info(funcao, mensagem, dados = null) {
    return this.log('INFO', funcao, mensagem, dados);
  },

  error(funcao, mensagem, erro) {
    return this.log('ERROR', funcao, mensagem, {
      error: erro.message,
      stack: erro.stack
    });
  },

  debug(funcao, mensagem, dados = null) {
    return this.log('DEBUG', funcao, mensagem, dados);
  }
};

function registrarLogsCliente(logsCliente) {
  try {
    logsCliente.forEach(log => {
      Logger.log(log.tipo, 'CLIENT-' + log.funcao, log.mensagem, log.dados);
    });
    return true;
  } catch (error) {
    console.error('Erro ao registrar logs do cliente:', error);
    return false;
  }
}

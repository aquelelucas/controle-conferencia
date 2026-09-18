const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';

const CHAVE_APP = 'KING-CONFERENCIA-2026';

function doGet(e) {
  const parametros = (e && e.parameter) ? e.parameter : {};
  const acao = String(parametros.acao || '').trim();
  const callback = String(parametros.callback || '').trim();

  if (!acao) {
    return responder_({
      sucesso: true,
      mensagem: 'API Link Separadores funcionando.'
    }, callback);
  }

  if (acao === 'separadores' || acao === 'conferentes') {
    const resultado = obterSeparadores_(e);

    if (acao === 'conferentes') {
      return responder_({
        sucesso: resultado.sucesso,
        conferentes: resultado.separadores || [],
        erro: resultado.erro || ''
      }, callback);
    }

    return responder_(resultado, callback);
  }

  if (acao === 'registrar') {
    return responder_(registrarPedido_(e), callback);
  }

  if (acao === 'quantidade') {
    return responder_(quantidadeHoje_(e), callback);
  }

  return responder_({
    sucesso: false,
    erro: 'Ação não reconhecida.'
  }, callback);
}

function registrarPedido_(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const parametros = (e && e.parameter) ? e.parameter : {};
    const chave = String(parametros.chave || '').trim();

    if (chave !== CHAVE_APP) {
      return { sucesso: false, erro: 'Chave inválida.' };
    }

    const pedido = String(parametros.pedido || '').trim();

    // Compatibilidade com versões antigas e novas do aplicativo.
    const separador = String(
      parametros.separador || parametros.conferente || ''
    ).trim();

    const conferente = String(
      parametros.conferente || parametros.separador || ''
    ).trim();

    if (!pedido) {
      return { sucesso: false, erro: 'Pedido não informado.' };
    }

    if (!separador) {
      return { sucesso: false, erro: 'Conferente não informado.' };
    }

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_LANCAMENTOS);

    if (!aba) {
      return {
        sucesso: false,
        erro: 'A aba "Lançamentos" não foi encontrada.'
      };
    }

    // Evita duplicidade causada por leituras repetidas da câmera
    // ou por dois aparelhos registrando o mesmo pedido ao mesmo tempo.
    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha >= 2) {
      const inicio = Math.max(2, ultimaLinha - 100);
      const quantidadeLinhas = ultimaLinha - inicio + 1;
      const registros = aba.getRange(inicio, 1, quantidadeLinhas, 5).getValues();

      const agora = new Date();
      const hoje = Utilities.formatDate(
        agora,
        Session.getScriptTimeZone(),
        'dd/MM/yyyy'
      );

      for (let i = registros.length - 1; i >= 0; i--) {
        const linha = registros[i];
        const data = String(linha[0] || '').trim();
        const pedidoRegistrado = String(linha[2] || '').trim();
        const separadorRegistrado = String(linha[3] || '').trim();

        if (
          data === hoje &&
          pedidoRegistrado === pedido &&
          separadorRegistrado === separador
        ) {
          return {
            sucesso: true,
            duplicado: true,
            pedido: pedido,
            separador: separador,
            conferente: conferente,
            mensagem: 'Pedido já registrado.'
          };
        }
      }
    }

    const agora = new Date();

    const data = Utilities.formatDate(
      agora,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

    const hora = Utilities.formatDate(
      agora,
      Session.getScriptTimeZone(),
      'HH:mm:ss'
    );

    const horaAtual = Number(
      Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH')
    );

    const minutoAtual = Number(
      Utilities.formatDate(agora, Session.getScriptTimeZone(), 'mm')
    );

    const minutosDoDia = (horaAtual * 60) + minutoAtual;

    // Até 12:00 = Manhã. Depois de 12:00 = Tarde.
    const turno = minutosDoDia <= 720 ? 'Manhã' : 'Tarde';

    // A planilha possui:
    // A Data | B Turno | C Pedido | D Separador | E Conferente
    const proximaLinha = aba.getLastRow() + 1;

    aba.getRange(proximaLinha, 1, 1, 5).setValues([[
      data,
      turno,
      pedido,
      separador,
      conferente
    ]]);

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      pedido: pedido,
      separador: separador,
      conferente: conferente,
      data: data,
      hora: hora,
      turno: turno
    };

  } catch (erro) {
    return {
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro desconhecido ao registrar o pedido.'
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function obterSeparadores_(e) {
  try {
    const parametros = (e && e.parameter) ? e.parameter : {};
    const chave = String(parametros.chave || '').trim();

    if (chave !== CHAVE_APP) {
      return {
        sucesso: false,
        erro: 'Chave inválida.'
      };
    }

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_CADASTRO);

    if (!aba) {
      return {
        sucesso: false,
        erro: 'A aba "Cadastro" não foi encontrada.'
      };
    }

    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      return {
        sucesso: true,
        separadores: []
      };
    }

    const nomes = aba
      .getRange(2, 1, ultimaLinha - 1, 1)
      .getValues()
      .flat()
      .map(function(nome) {
        return String(nome || '').trim();
      })
      .filter(function(nome) {
        return nome !== '';
      });

    return {
      sucesso: true,
      separadores: [...new Set(nomes)]
    };

  } catch (erro) {
    return {
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro ao carregar os separadores.'
    };
  }
}

function quantidadeHoje_(e) {
  try {
    const parametros = (e && e.parameter) ? e.parameter : {};
    const chave = String(parametros.chave || '').trim();

    if (chave !== CHAVE_APP) {
      return {
        sucesso: false,
        erro: 'Chave inválida.'
      };
    }

    const separador = String(
      parametros.separador || parametros.conferente || ''
    ).trim();

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_LANCAMENTOS);

    if (!aba || aba.getLastRow() < 2) {
      return {
        sucesso: true,
        quantidade: 0
      };
    }

    const dados = aba
      .getRange(2, 1, aba.getLastRow() - 1, 5)
      .getValues();

    const hoje = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

    let quantidade = 0;

    dados.forEach(function(linha) {
      const data = String(linha[0] || '').trim();
      const nomeSeparador = String(linha[3] || '').trim();

      if (
        data === hoje &&
        (!separador || nomeSeparador === separador)
      ) {
        quantidade++;
      }
    });

    return {
      sucesso: true,
      quantidade: quantidade
    };

  } catch (erro) {
    return {
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro ao carregar a quantidade.'
    };
  }
}

function responder_(objeto, callback) {
  const json = JSON.stringify(objeto);

  if (
    callback &&
    /^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)
  ) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
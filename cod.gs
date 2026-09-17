const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';

const CHAVE_APP = 'KING-CONFERENCIA-2026';

function doGet(e) {
  const acao = String((e && e.parameter && e.parameter.acao) || '').trim();
  const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
  let resultado;

  if (!acao) {
    resultado = {
      sucesso: true,
      mensagem: 'API Controle Profissional funcionando.'
    };
    return responder_(resultado, callback);
  }

  if (acao === 'separadores') {
    resultado = obterSeparadores_(e);
    return responder_(resultado, callback);
  }

  if (acao === 'registrar') {
    resultado = registrarPedido_(e);
    return responder_(resultado, callback);
  }

  if (acao === 'quantidade') {
    resultado = quantidadeHoje_(e);
    return responder_(resultado, callback);
  }

  resultado = {
    sucesso: false,
    erro: 'Ação não reconhecida.'
  };

  return responder_(resultado, callback);
}

function registrarPedido_(e) {
  try {
    const chave = String(e.parameter.chave || '').trim();

    if (chave !== CHAVE_APP) {
      return { sucesso: false, erro: 'Chave inválida.' };
    }

    const pedido = String(e.parameter.pedido || '').trim();
    const separador = String(e.parameter.separador || '').trim();

    if (!pedido) {
      return { sucesso: false, erro: 'Pedido não informado.' };
    }

    if (!separador) {
      return { sucesso: false, erro: 'Separador não informado.' };
    }

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_LANCAMENTOS);

    if (!aba) {
      return {
        sucesso: false,
        erro: 'A aba "Lançamentos" não foi encontrada.'
      };
    }

    const agora = new Date();
    const data = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const hora = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH:mm:ss');
    const horaAtual = Number(Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH'));
    const minutoAtual = Number(Utilities.formatDate(agora, Session.getScriptTimeZone(), 'mm'));
    const minutosDoDia = (horaAtual * 60) + minutoAtual;
    const turno = minutosDoDia <= 720 ? 'Manhã' : 'Tarde';
    const proximaLinha = aba.getLastRow() + 1;

    aba.getRange(proximaLinha, 1, 1, 4).setValues([[data, turno, pedido, separador]]);
    SpreadsheetApp.flush();

    return {
      sucesso: true,
      pedido: pedido,
      separador: separador,
      data: data,
      hora: hora,
      turno: turno
    };
  } catch (erro) {
    return {
      sucesso: false,
      erro: erro.message
    };
  }
}

function obterSeparadores_(e) {
  try {
    const chave = String(e.parameter.chave || '').trim();

    if (chave !== CHAVE_APP) {
      return { sucesso: false, erro: 'Chave inválida.' };
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
      return { sucesso: true, separadores: [] };
    }

    const nomes = aba.getRange(2, 1, ultimaLinha - 1, 1)
      .getValues()
      .flat()
      .map(function(nome) {
        return String(nome || '').trim();
      })
      .filter(function(nome) {
        return nome !== '';
      });

    const separadores = [...new Set(nomes)];

    return {
      sucesso: true,
      separadores: separadores
    };
  } catch (erro) {
    return {
      sucesso: false,
      erro: erro.message
    };
  }
}

function quantidadeHoje_(e) {
  try {
    const chave = String(e.parameter.chave || '').trim();

    if (chave !== CHAVE_APP) {
      return { sucesso: false, erro: 'Chave inválida.' };
    }

    const separador = String(e.parameter.separador || '').trim();
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(ABA_LANCAMENTOS);

    if (!aba || aba.getLastRow() < 2) {
      return { sucesso: true, quantidade: 0 };
    }

    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues();
    const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    let quantidade = 0;

    dados.forEach(function(linha) {
      const data = String(linha[0] || '').trim();
      const nomeSeparador = String(linha[3] || '').trim();

      if (data === hoje && (!separador || nomeSeparador === separador)) {
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
      erro: erro.message
    };
  }
}

function responder_(objeto, callback) {
  const json = JSON.stringify(objeto);

  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

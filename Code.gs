const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';
const CHAVE_APP = 'KING-CONFERENCIA-2026';

/*
 * ============================================================
 * LINK SEPARADORES
 * ============================================================
 *
 * Este projeto tem um objetivo único:
 *
 * 1. O separador escolhe o próprio nome.
 * 2. Entra na conferência.
 * 3. A câmera é aberta.
 * 4. O código de barras do pedido é bipado.
 * 5. O sistema registra uma NOVA LINHA em Lançamentos.
 *
 * O registro feito pelo Link Separadores é SOMENTE:
 * A = Data
 * B = Turno
 * C = Pedido
 * D = Separador
 *
 * Nenhuma outra coluna da tabela é alterada.
 * Este projeto NÃO é o projeto de conferência/auditoria.
 * Não existe confirmação SIM/NÃO neste fluxo.
 */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const acao = String(p.acao || '').trim();
  const callback = String(p.callback || '').trim();

  try {
    if (acao === '') {
      return responder_({
        sucesso: true,
        mensagem: 'API Link Separadores funcionando.',
        versao: '1.0'
      }, callback);
    }

    if (String(p.chave || '').trim() !== CHAVE_APP) {
      return responder_({
        sucesso: false,
        erro: 'Chave inválida.'
      }, callback);
    }

    if (acao === 'separadores') {
      return responder_(obterSeparadores_(), callback);
    }

    if (acao === 'registrar') {
      return responder_(registrarPedido_(p), callback);
    }

    if (acao === 'quantidade') {
      return responder_(quantidadeHoje_(p.separador), callback);
    }

    return responder_({
      sucesso: false,
      erro: 'Ação não reconhecida.'
    }, callback);

  } catch (erro) {
    return responder_({
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro interno no servidor.'
    }, callback);
  }
}

function obterSeparadores_() {
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

  // A lista de separadores fica na coluna A do Cadastro.
  const nomes = aba
    .getRange(2, 1, ultimaLinha - 1, 1)
    .getDisplayValues()
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
}

function registrarPedido_(p) {
  const pedido = String(p.pedido || '').trim();
  const separador = String(p.separador || '').trim();

  if (!pedido) {
    return {
      sucesso: false,
      erro: 'Pedido não informado.'
    };
  }

  if (!separador) {
    return {
      sucesso: false,
      erro: 'Separador não informado.'
    };
  }

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(ABA_LANCAMENTOS);

  if (!aba) {
    return {
      sucesso: false,
      erro: 'A aba "Lançamentos" não foi encontrada.'
    };
  }

  /*
   * Protege o append contra dois bipes ocorrendo exatamente
   * ao mesmo tempo em aparelhos diferentes.
   */
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    /*
     * O Link Separadores sempre cria uma NOVA LINHA.
     * Ele não procura pedido existente e não sobrescreve dados.
     */
    const agora = new Date();
    const data = new Date(
      agora.getFullYear(),
      agora.getMonth(),
      agora.getDate()
    );

    const hora = Number(
      Utilities.formatDate(
        agora,
        Session.getScriptTimeZone(),
        'HH'
      )
    );

    const turno = hora < 12 ? 'Manhã' : 'Tarde';
    const proximaLinha = aba.getLastRow() + 1;

    // SOMENTE A:D.
    aba.getRange(proximaLinha, 1, 1, 4).setValues([[
      data,
      turno,
      pedido,
      separador
    ]]);

    // Mantém A como data real, sem alterar as outras colunas.
    aba.getRange(proximaLinha, 1).setNumberFormat('dd/MM/yyyy');

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      data: Utilities.formatDate(
        data,
        Session.getScriptTimeZone(),
        'dd/MM/yyyy'
      ),
      turno: turno,
      pedido: pedido,
      separador: separador,
      linha: proximaLinha
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (erro) {}
  }
}

function quantidadeHoje_(separador) {
  const nome = String(separador || '').trim();

  if (!nome) {
    return {
      sucesso: true,
      quantidade: 0
    };
  }

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(ABA_LANCAMENTOS);

  if (!aba || aba.getLastRow() < 2) {
    return {
      sucesso: true,
      quantidade: 0
    };
  }

  /*
   * A contagem considera somente registros feitos por este
   * separador no dia atual, usando as colunas A e D.
   */
  const dados = aba
    .getRange(2, 1, aba.getLastRow() - 1, 4)
    .getDisplayValues();

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
      nomeSeparador.toLowerCase() === nome.toLowerCase()
    ) {
      quantidade++;
    }
  });

  return {
    sucesso: true,
    quantidade: quantidade
  };
}

function responder_(objeto, callback) {
  const json = JSON.stringify(objeto);

  if (
    callback &&
    /^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)
  ) {
    return ContentService
      .createTextOutput(
        callback + '(' + json + ');'
      )
      .setMimeType(
        ContentService.MimeType.JAVASCRIPT
      );
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(
      ContentService.MimeType.JSON
    );
}

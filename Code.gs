const ID_PLANILHA = '16l4PoccxeI_masCzuQh1vHfJV5z3A-yvz80A2yaubRk';
const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';
const CHAVE_APP = 'KING-CONFERENCIA-2026';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const acao = String(p.acao || '').trim();
  const callback = String(p.callback || '').trim();
  let resultado;

  try {
    if (acao === 'conferentes') resultado = obterConferentes_();
    else if (acao === 'consultarPedido') resultado = consultarPedido_(p.pedido);
    else if (acao === 'salvarConferencia') resultado = salvarConferencia_(p);
    else if (acao === 'quantidade') resultado = quantidadeSessao_(p.conferente);
    else resultado = { sucesso: true, mensagem: 'API Controle de Conferência funcionando.' };
  } catch (erro) {
    resultado = { sucesso: false, erro: erro.message };
  }

  return responder_(resultado, callback);
}

function validarChave_(p) {
  if (String(p.chave || '').trim() !== CHAVE_APP) {
    throw new Error('Chave inválida.');
  }
}

function planilha_() {
  return SpreadsheetApp.openById(ID_PLANILHA);
}

function obterConferentes_() {
  const aba = planilha_().getSheetByName(ABA_CADASTRO);
  if (!aba) throw new Error('A aba "Cadastro" não foi encontrada.');

  const ultima = aba.getLastRow();
  if (ultima < 2) return { sucesso: true, conferentes: [] };

  // No Cadastro, os nomes dos conferentes ficam na coluna A.
  // Mantemos a leitura a partir da linha 2 para ignorar o cabeçalho.
  const nomes = aba.getRange(2, 1, ultima - 1, 1).getDisplayValues()
    .flat()
    .map(v => String(v).trim())
    .filter(Boolean);

  return {
    sucesso: true,
    conferentes: [...new Set(nomes)]
  };
}

function localizarPedido_(pedido) {
  const valor = String(pedido || '').trim();
  if (!valor) throw new Error('Número do pedido não informado.');

  const aba = planilha_().getSheetByName(ABA_LANCAMENTOS);
  if (!aba) throw new Error('A aba "Lançamentos" não foi encontrada.');

  const ultima = aba.getLastRow();
  if (ultima < 2) throw new Error('Nenhum pedido foi lançado ainda.');

  const dados = aba.getRange(2, 1, ultima - 1, 14).getDisplayValues();

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const pedidoPlanilha = String(linha[2] || '').trim();

    if (pedidoPlanilha !== valor) continue;

    const conferente = String(linha[4] || '').trim();
    const resultado = String(linha[13] || '').trim();

    return {
      aba: aba,
      linha: i + 2,
      pedido: pedidoPlanilha,
      separador: String(linha[3] || '').trim(),
      conferenteAtual: conferente,
      resultadoAtual: resultado,
      conferido: !!conferente || !!resultado
    };
  }

  return null;
}

function consultarPedido_(pedido) {
  const registro = localizarPedido_(pedido);
  if (!registro) {
    throw new Error('Pedido ' + String(pedido || '').trim() + ' não encontrado na aba Lançamentos.');
  }

  if (registro.conferido) {
    return {
      sucesso: false,
      jaConferido: true,
      erro: 'O pedido ' + registro.pedido + ' já foi conferido.'
    };
  }

  if (!registro.separador) {
    return {
      sucesso: false,
      erro: 'O pedido ' + registro.pedido + ' foi encontrado, mas não possui separador informado na coluna D.'
    };
  }

  return {
    sucesso: true,
    pedido: registro.pedido,
    linha: registro.linha,
    separador: registro.separador
  };
}

function salvarConferencia_(p) {
  validarChave_(p);

  const conferente = String(p.conferente || '').trim();
  const pedido = String(p.pedido || '').trim();
  const correta = String(p.correta || '').trim();

  if (!conferente) throw new Error('Conferente não informado.');
  if (!pedido) throw new Error('Pedido não informado.');
  if (correta !== 'SIM' && correta !== 'NÃO') {
    throw new Error('Resultado da conferência inválido.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const registro = localizarPedido_(pedido);
    if (!registro) throw new Error('Pedido ' + pedido + ' não encontrado.');

    if (registro.conferido) {
      throw new Error('O pedido ' + registro.pedido + ' já foi conferido.');
    }

    const linha = registro.linha;
    const aba = registro.aba;

    // Nunca cria nova linha e não altera A:D.
    // E = Conferente, K = Erro Detectado?, N = resultado da conferência.
    aba.getRange(linha, 5).setValue(conferente);
    aba.getRange(linha, 11).setValue(correta === 'SIM' ? 'NÃO' : 'SIM');
    aba.getRange(linha, 14).setValue(correta);

    if (correta === 'NÃO') {
      aba.getRange(linha, 6, 1, 5).setValues([[
        String(p.sku || '').trim(),
        p.qtdSolicitada === '' ? '' : p.qtdSolicitada,
        p.qtdSeparada === '' ? '' : p.qtdSeparada,
        String(p.tipoErro || '').trim(),
        String(p.gravidade || '').trim()
      ]]);

      aba.getRange(linha, 12, 1, 2).setValues([[
        String(p.acaoTomada || '').trim(),
        String(p.observacao || '').trim()
      ]]);
    }

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      pedido: registro.pedido,
      separador: registro.separador,
      conferente: conferente,
      correta: correta
    };
  } finally {
    lock.releaseLock();
  }
}

function quantidadeSessao_(conferente) {
  const nome = String(conferente || '').trim();
  if (!nome) return { sucesso: true, quantidade: 0 };

  const aba = planilha_().getSheetByName(ABA_LANCAMENTOS);
  if (!aba || aba.getLastRow() < 2) {
    return { sucesso: true, quantidade: 0 };
  }

  const dados = aba.getRange(2, 5, aba.getLastRow() - 1, 10).getDisplayValues();
  let quantidade = 0;

  dados.forEach(linha => {
    const conferenteLinha = String(linha[0] || '').trim();
    const resultado = String(linha[9] || '').trim();

    if (
      conferenteLinha === nome &&
      (resultado === 'SIM' || resultado === 'NÃO')
    ) {
      quantidade++;
    }
  });

  return { sucesso: true, quantidade: quantidade };
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
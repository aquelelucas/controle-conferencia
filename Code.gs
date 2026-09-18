const ID_PLANILHA = '16l4PoccxeI_masCzuQh1vHfJV5z3A-yvz80A2yaubRk';
const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';
const CHAVE_APP = 'KING-CONFERENCIA-2026';

/**
 * API do Controle de Conferência.
 *
 * Fluxo:
 * - consulta o pedido já lançado em Lançamentos;
 * - lê o separador da coluna D;
 * - grava somente a conferência na mesma linha;
 * - não altera Data, Turno, Pedido ou Separador.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const acao = String(p.acao || '').trim();
  const callback = String(p.callback || '').trim();

  try {
    if (!acao) {
      return responder_({
        sucesso: true,
        mensagem: 'API Controle de Conferência funcionando.'
      }, callback);
    }

    validarChave_(p);

    let resultado;

    if (acao === 'conferentes') {
      resultado = obterConferentes_();
    } else if (acao === 'consultarPedido') {
      resultado = consultarPedido_(p.pedido);
    } else if (acao === 'salvarConferencia') {
      resultado = salvarConferencia_(p);
    } else if (acao === 'quantidade') {
      resultado = quantidadeConferencia_(p.conferente);
    } else {
      resultado = {
        sucesso: false,
        erro: 'Ação não reconhecida.'
      };
    }

    return responder_(resultado, callback);
  } catch (erro) {
    return responder_({
      sucesso: false,
      erro: erro && erro.message ? erro.message : String(erro)
    }, callback);
  }
}

function validarChave_(p) {
  if (String(p.chave || '').trim() !== CHAVE_APP) {
    throw new Error('Chave inválida.');
  }
}

function planilha_() {
  return SpreadsheetApp.openById(ID_PLANILHA);
}

function normalizarTexto_(valor) {
  return String(valor == null ? '' : valor).trim();
}

function normalizarCabecalho_(valor) {
  return normalizarTexto_(valor)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/**
 * Localiza a coluna dos conferentes de forma tolerante:
 * 1) procura um cabeçalho com "conferente";
 * 2) se não encontrar, usa a coluna D (layout esperado);
 * 3) se D estiver vazia, tenta a coluna A como compatibilidade.
 */
function descobrirColunaConferentes_(aba) {
  const ultimaColuna = Math.max(aba.getLastColumn(), 4);
  const cabecalhos = aba.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];

  for (let i = 0; i < cabecalhos.length; i++) {
    const cabecalho = normalizarCabecalho_(cabecalhos[i]);
    if (cabecalho.includes('conferente')) {
      return i + 1;
    }
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha >= 2) {
    const valoresD = aba.getRange(2, 4, ultimaLinha - 1, 1)
      .getDisplayValues()
      .flat()
      .map(normalizarTexto_)
      .filter(Boolean);

    if (valoresD.length) return 4;
  }

  return 1;
}

function lerNomesDaColuna_(aba, coluna) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  return aba.getRange(2, coluna, ultimaLinha - 1, 1)
    .getDisplayValues()
    .flat()
    .map(normalizarTexto_)
    .filter(Boolean)
    .filter((nome, indice, lista) => lista.indexOf(nome) === indice);
}

function obterConferentes_() {
  const aba = planilha_().getSheetByName(ABA_CADASTRO);

  if (!aba) {
    throw new Error('A aba "Cadastro" não foi encontrada.');
  }

  const coluna = descobrirColunaConferentes_(aba);
  const conferentes = lerNomesDaColuna_(aba, coluna);

  return {
    sucesso: true,
    conferentes: conferentes
  };
}

function conferenteCadastrado_(nome) {
  const alvo = normalizarTexto_(nome);
  if (!alvo) return false;

  const dados = obterConferentes_().conferentes;
  return dados.some(nomeCadastro => normalizarTexto_(nomeCadastro) === alvo);
}

function normalizarPedido_(valor) {
  const texto = normalizarTexto_(valor);

  if (!texto) return '';

  // Mantém pedidos alfanuméricos intactos.
  // Para códigos somente numéricos, permite equivalência de zeros à esquerda.
  if (/^\d+$/.test(texto)) {
    return texto.replace(/^0+/, '') || '0';
  }

  return texto.toUpperCase();
}

function localizarPedido_(pedido) {
  const valorBusca = normalizarTexto_(pedido);

  if (!valorBusca) {
    throw new Error('Número do pedido não informado.');
  }

  const aba = planilha_().getSheetByName(ABA_LANCAMENTOS);

  if (!aba) {
    throw new Error('A aba "Lançamentos" não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error('Nenhum pedido foi lançado ainda.');
  }

  const dados = aba.getRange(2, 1, ultimaLinha - 1, 14).getDisplayValues();
  const chaveBusca = normalizarPedido_(valorBusca);

  let primeiroConferido = null;

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const pedidoPlanilha = normalizarTexto_(linha[2]);

    if (!pedidoPlanilha) continue;
    if (normalizarPedido_(pedidoPlanilha) !== chaveBusca) continue;

    const registro = {
      aba: aba,
      linha: i + 2,
      pedido: pedidoPlanilha,
      separador: normalizarTexto_(linha[3]),
      conferenteAtual: normalizarTexto_(linha[4]),
      resultadoAtual: normalizarTexto_(linha[13])
    };

    registro.conferido = !!registro.conferenteAtual || !!registro.resultadoAtual;

    // Se houver mais de uma linha para o mesmo pedido, usa a primeira ainda
    // não conferida. Caso todas já estejam conferidas, retorna uma delas.
    if (!registro.conferido) {
      return registro;
    }

    if (!primeiroConferido) {
      primeiroConferido = registro;
    }
  }

  return primeiroConferido;
}

function consultarPedido_(pedido) {
  const registro = localizarPedido_(pedido);

  if (!registro) {
    throw new Error(
      'Pedido ' + normalizarTexto_(pedido) + ' não encontrado na aba Lançamentos.'
    );
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
      erro:
        'O pedido ' +
        registro.pedido +
        ' foi encontrado, mas não possui separador informado na coluna D.'
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
  const conferente = normalizarTexto_(p.conferente);
  const pedido = normalizarTexto_(p.pedido);
  const correta = normalizarTexto_(p.correta).toUpperCase();

  if (!conferente) throw new Error('Conferente não informado.');
  if (!pedido) throw new Error('Pedido não informado.');

  if (correta !== 'SIM' && correta !== 'NÃO') {
    throw new Error('Resultado da conferência inválido.');
  }

  if (!conferenteCadastrado_(conferente)) {
    throw new Error(
      'Conferente "' + conferente + '" não foi encontrado no Cadastro.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const registro = localizarPedido_(pedido);

    if (!registro) {
      throw new Error('Pedido ' + pedido + ' não encontrado.');
    }

    if (registro.conferido) {
      throw new Error('O pedido ' + registro.pedido + ' já foi conferido.');
    }

    const aba = registro.aba;
    const linha = registro.linha;

    // A:D permanecem intactas.
    aba.getRange(linha, 5).setValue(conferente); // E Conferente
    aba.getRange(linha, 11).setValue(correta === 'SIM' ? 'NÃO' : 'SIM'); // K Erro Detectado?
    aba.getRange(linha, 14).setValue(correta); // N Resultado

    if (correta === 'NÃO') {
      aba.getRange(linha, 6, 1, 5).setValues([[
        normalizarTexto_(p.sku),
        p.qtdSolicitada === '' ? '' : p.qtdSolicitada,
        p.qtdSeparada === '' ? '' : p.qtdSeparada,
        normalizarTexto_(p.tipoErro),
        normalizarTexto_(p.gravidade)
      ]]);

      aba.getRange(linha, 12, 1, 2).setValues([[
        normalizarTexto_(p.acaoTomada),
        normalizarTexto_(p.observacao)
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

function quantidadeConferencia_(conferente) {
  const nome = normalizarTexto_(conferente);

  if (!nome) {
    return {
      sucesso: true,
      quantidade: 0
    };
  }

  const aba = planilha_().getSheetByName(ABA_LANCAMENTOS);

  if (!aba || aba.getLastRow() < 2) {
    return {
      sucesso: true,
      quantidade: 0
    };
  }

  const dados = aba.getRange(2, 5, aba.getLastRow() - 1, 10).getDisplayValues();
  let quantidade = 0;

  dados.forEach(linha => {
    const conferenteLinha = normalizarTexto_(linha[0]);
    const resultado = normalizarTexto_(linha[9]);

    if (
      conferenteLinha === nome &&
      (resultado === 'SIM' || resultado === 'NÃO')
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
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

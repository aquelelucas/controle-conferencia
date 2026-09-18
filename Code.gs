const ID_PLANILHA = '16l4PoccxeI_masCzuQh1vHfJV5z3A-yvz80A2yaubRk';

const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';
const CHAVE_APP = 'KING-CONFERENCIA-2026';

/*
 * LINK SEPARADORES
 *
 * Fluxo:
 * 1. Os pedidos são pré-lançados na aba "Lançamentos".
 * 2. O pedido fica na coluna C e o separador na coluna D.
 * 3. O aplicativo lê o código de barras ou aceita o número digitado.
 * 4. O Apps Script localiza o pedido ainda não conferido.
 * 5. O conferente confirma "SIM" ou "NÃO".
 * 6. O script grava a conferência na mesma linha, sem alterar A:D.
 *
 * Estrutura de "Lançamentos":
 * A Data
 * B Turno
 * C Pedido
 * D Separador
 * E Conferente
 * F SKU/Produto
 * G Qtd. Solicitada
 * H Qtd. Separada
 * I Tipo de Erro
 * J Gravidade
 * K Erro Detectado?
 * L Ação Tomada
 * M Observação
 * N A separação está correta?
 */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const acao = String(p.acao || '').trim();
  const callback = String(p.callback || '').trim();

  try {
    validarChave_(p);

    let resultado;

    if (!acao) {
      resultado = {
        sucesso: true,
        mensagem: 'API Link Separadores funcionando.',
        versao: '2.1'
      };
    } else if (acao === 'conferentes' || acao === 'separadores') {
      resultado = obterConferentes_(acao);
    } else if (acao === 'consultarPedido') {
      resultado = consultarPedido_(p.pedido);
    } else if (acao === 'salvarConferencia') {
      resultado = salvarConferencia_(p);
    } else if (acao === 'registrar') {
      // Compatibilidade com versões antigas que registravam diretamente.
      resultado = salvarConferencia_(Object.assign({}, p, {
        correta: String(p.correta || 'SIM').trim().toUpperCase()
      }));
    } else if (acao === 'quantidade') {
      resultado = quantidadeConferencia_(p.conferente || p.separador);
    } else if (acao === 'diagnostico') {
      resultado = diagnostico_();
    } else {
      resultado = {
        sucesso: false,
        erro: 'Ação não reconhecida: ' + acao
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

function obterConferentes_(acao) {
  const aba = planilha_().getSheetByName(ABA_CADASTRO);

  if (!aba) {
    throw new Error('A aba "Cadastro" não foi encontrada.');
  }

  /*
   * Cadastro atual possui:
   * A = Separador
   * D = Conferente
   *
   * Usamos A como fonte principal, pois é a coluna que alimenta
   * os rankings. Também juntamos D para manter compatibilidade
   * com versões que preenchiam os nomes nela.
   */
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return acao === 'separadores'
      ? {sucesso: true, separadores: []}
      : {sucesso: true, conferentes: []};
  }

  const ultimaColuna = Math.max(4, aba.getLastColumn());
  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, ultimaColuna)
    .getDisplayValues();

  const nomes = [];

  dados.forEach(linha => {
    const nomeA = normalizarTexto_(linha[0]);
    const nomeD = normalizarTexto_(linha[3]);

    if (nomeA) nomes.push(nomeA);
    if (nomeD) nomes.push(nomeD);
  });

  const unicos = [...new Set(nomes)];

  return acao === 'separadores'
    ? {sucesso: true, separadores: unicos}
    : {sucesso: true, conferentes: unicos};
}

function conferenteCadastrado_(nome) {
  const alvo = normalizarTexto_(nome).toLowerCase();

  if (!alvo) return false;

  const lista = obterConferentes_('conferentes').conferentes;

  return lista.some(nomeCadastro =>
    normalizarTexto_(nomeCadastro).toLowerCase() === alvo
  );
}

function normalizarPedido_(valor) {
  const texto = normalizarTexto_(valor);

  if (!texto) return '';

  /*
   * Para códigos somente numéricos, ignora zeros à esquerda.
   * Para códigos alfanuméricos, preserva o conteúdo e normaliza
   * apenas a capitalização.
   */
  if (/^\d+$/.test(texto)) {
    return texto.replace(/^0+/, '') || '0';
  }

  return texto.toUpperCase();
}

function obterHojeTexto_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}

function localizarPedido_(pedido) {
  const busca = normalizarTexto_(pedido);

  if (!busca) {
    throw new Error('Número do pedido não informado.');
  }

  const aba = planilha_().getSheetByName(ABA_LANCAMENTOS);

  if (!aba) {
    throw new Error('A aba "Lançamentos" não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 14)
    .getDisplayValues();

  const chaveBusca = normalizarPedido_(busca);
  const hoje = obterHojeTexto_();

  let primeiroHojeConferido = null;
  let primeiroAntigo = null;

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];

    const data = normalizarTexto_(linha[0]);
    const pedidoPlanilha = normalizarTexto_(linha[2]);

    if (!pedidoPlanilha) continue;

    if (normalizarPedido_(pedidoPlanilha) !== chaveBusca) {
      continue;
    }

    const registro = {
      aba: aba,
      linha: i + 2,
      data: data,
      pedido: pedidoPlanilha,
      separador: normalizarTexto_(linha[3]),
      conferenteAtual: normalizarTexto_(linha[4]),
      resultadoAtual: normalizarTexto_(linha[13])
    };

    registro.conferido =
      !!registro.conferenteAtual ||
      registro.resultadoAtual === 'SIM' ||
      registro.resultadoAtual === 'NÃO';

    if (data === hoje && !registro.conferido) {
      return registro;
    }

    if (data === hoje && registro.conferido && !primeiroHojeConferido) {
      primeiroHojeConferido = registro;
    }

    if (data !== hoje && !registro.conferido && !primeiroAntigo) {
      primeiroAntigo = registro;
    }
  }

  return primeiroHojeConferido || primeiroAntigo || null;
}

function consultarPedido_(pedido) {
  const registro = localizarPedido_(pedido);

  if (!registro) {
    return {
      sucesso: false,
      erro: 'Pedido ' + normalizarTexto_(pedido) + ' não encontrado na aba Lançamentos.'
    };
  }

  const hoje = obterHojeTexto_();

  if (registro.data !== hoje) {
    return {
      sucesso: false,
      erro:
        'O pedido ' +
        registro.pedido +
        ' foi encontrado, mas a data registrada é ' +
        registro.data +
        '.'
    };
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

function normalizarQuantidade_(valor) {
  const texto = normalizarTexto_(valor);

  if (!texto) return '';

  const numero = Number(texto.replace(',', '.'));

  return Number.isFinite(numero) ? numero : texto;
}

function salvarConferencia_(p) {
  const conferente = normalizarTexto_(p.conferente);
  const pedido = normalizarTexto_(p.pedido);
  const correta = normalizarTexto_(p.correta).toUpperCase();

  if (!conferente) {
    throw new Error('Conferente não informado.');
  }

  if (!pedido) {
    throw new Error('Pedido não informado.');
  }

  if (correta !== 'SIM' && correta !== 'NÃO') {
    throw new Error('Resultado da conferência inválido.');
  }

  if (!conferenteCadastrado_(conferente)) {
    throw new Error(
      'Conferente "' +
      conferente +
      '" não foi encontrado no Cadastro.'
    );
  }

  const tipoErro = normalizarTexto_(p.tipoErro);
  const gravidade = normalizarTexto_(p.gravidade);
  const sku = normalizarTexto_(p.sku);
  const qtdSolicitada = normalizarQuantidade_(p.qtdSolicitada);
  const qtdSeparada = normalizarQuantidade_(p.qtdSeparada);
  const acaoTomada = normalizarTexto_(p.acaoTomada);
  const observacao = normalizarTexto_(p.observacao);

  if (correta === 'NÃO' && !tipoErro) {
    throw new Error('Informe o tipo de erro.');
  }

  if (correta === 'NÃO' && !gravidade) {
    throw new Error('Informe a gravidade do erro.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const registro = localizarPedido_(pedido);

    if (!registro) {
      throw new Error(
        'Pedido ' +
        pedido +
        ' não encontrado na aba Lançamentos.'
      );
    }

    if (registro.data !== obterHojeTexto_()) {
      throw new Error(
        'O pedido ' +
        registro.pedido +
        ' não pertence à data de hoje.'
      );
    }

    if (registro.conferido) {
      throw new Error(
        'O pedido ' +
        registro.pedido +
        ' já foi conferido.'
      );
    }

    const aba = registro.aba;
    const linha = registro.linha;

    /*
     * A:D são dados do pedido e não são alterados.
     * E:N recebem somente o resultado desta conferência.
     */
    const dadosConferencia = [
      conferente,                         // E
      correta === 'NÃO' ? sku : '',       // F
      correta === 'NÃO' ? qtdSolicitada : '', // G
      correta === 'NÃO' ? qtdSeparada : '',   // H
      correta === 'NÃO' ? tipoErro : '',  // I
      correta === 'NÃO' ? gravidade : '', // J
      correta === 'NÃO' ? 'Sim' : 'Não',  // K
      correta === 'NÃO' ? acaoTomada : '',// L
      correta === 'NÃO' ? observacao : '',// M
      correta                            // N
    ];

    aba.getRange(linha, 5, 1, 10).setValues([dadosConferencia]);

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      pedido: registro.pedido,
      linha: linha,
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

  const dados = aba
    .getRange(2, 1, aba.getLastRow() - 1, 14)
    .getDisplayValues();

  const hoje = obterHojeTexto_();
  let quantidade = 0;

  dados.forEach(linha => {
    const data = normalizarTexto_(linha[0]);
    const nomeLinha = normalizarTexto_(linha[4]);
    const resultado = normalizarTexto_(linha[13]);

    if (
      data === hoje &&
      nomeLinha === nome &&
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

function diagnostico_() {
  const planilha = planilha_();
  const cadastro = planilha.getSheetByName(ABA_CADASTRO);
  const lancamentos = planilha.getSheetByName(ABA_LANCAMENTOS);

  return {
    sucesso: true,
    versao: '2.1',
    planilha: planilha.getName(),
    cadastroExiste: !!cadastro,
    lancamentosExiste: !!lancamentos,
    linhasCadastro: cadastro ? cadastro.getLastRow() : 0,
    linhasLancamentos: lancamentos ? lancamentos.getLastRow() : 0,
    urlPlanilha: planilha.getUrl()
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

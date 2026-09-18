const ID_PLANILHA = '16l4PoccxeI_masCzuQh1vHfJV5z3A-yvz80A2yaubRk';

const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';
const CHAVE_APP = 'KING-CONFERENCIA-2026';

/*
 * ============================================================
 * LINK SEPARADORES
 * ============================================================
 *
 * Este arquivo pertence somente ao fluxo dos SEPARADORES.
 *
 * Fluxo:
 * 1. O separador escolhe o próprio nome na tela inicial.
 * 2. Entra na conferência.
 * 3. A câmera lê o código de barras do pedido.
 * 4. O aplicativo mostra o pedido e pede SIM ou NÃO.
 * 5. O Apps Script valida o pedido e o separador na aba
 *    Lançamentos.
 * 6. O resultado é gravado na MESMA LINHA do pedido.
 * 7. A câmera permanece aberta e fica pronta para o próximo pedido.
 *
 * Importante:
 * - A coluna D (Separador) é dado do lançamento e não é alterada.
 * - A coluna E (Conferente) pertence a outro fluxo e não é alterada.
 * - O Link Separadores não cadastra pedidos.
 * - O Link Separadores não preenche formulário de erro.
 *
 * Estrutura de Lançamentos:
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
    if (!acao) {
      return responder_({
        sucesso: true,
        mensagem: 'API Link Separadores funcionando.',
        versao: '3.0'
      }, callback);
    }

    validarChave_(p);

    let resultado;

    if (acao === 'separadores' || acao === 'conferentes') {
      resultado = obterSeparadores_();

      if (acao === 'conferentes') {
        // Compatibilidade temporária com versões antigas da interface.
        resultado = {
          sucesso: resultado.sucesso,
          conferentes: resultado.separadores || [],
          erro: resultado.erro || ''
        };
      }
    } else if (acao === 'registrar') {
      // Compatibilidade com a primeira versão do Link Separadores.
      resultado = registrarDireto_(p);
    } else if (acao === 'salvarConferencia') {
      resultado = salvarConferencia_(p);
    } else if (acao === 'quantidade') {
      resultado = quantidadeHoje_(p.separador || p.conferente);
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

function texto_(valor) {
  return String(valor == null ? '' : valor).trim();
}

function normalizarNome_(valor) {
  return texto_(valor).toLowerCase();
}

function normalizarPedido_(valor) {
  const texto = texto_(valor);

  if (!texto) return '';

  if (/^\d+$/.test(texto)) {
    return texto.replace(/^0+/, '') || '0';
  }

  return texto.toUpperCase();
}

function obterSeparadores_() {
  const aba = planilha_().getSheetByName(ABA_CADASTRO);

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

  // A coluna A é a lista oficial de separadores.
  const nomes = aba
    .getRange(2, 1, ultimaLinha - 1, 1)
    .getDisplayValues()
    .flat()
    .map(texto_)
    .filter(Boolean);

  return {
    sucesso: true,
    separadores: [...new Set(nomes)]
  };
}

function separadorCadastrado_(nome) {
  const alvo = normalizarNome_(nome);

  if (!alvo) return false;

  return obterSeparadores_()
    .separadores
    .some(nomeCadastro => normalizarNome_(nomeCadastro) === alvo);
}

function localizarPedido_(pedido, separador) {
  const chavePedido = normalizarPedido_(pedido);
  const chaveSeparador = normalizarNome_(separador);

  if (!chavePedido) return null;

  const aba = planilha_().getSheetByName(ABA_LANCAMENTOS);

  if (!aba) {
    throw new Error('A aba "Lançamentos" não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return null;

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 14)
    .getDisplayValues();

  let pedidoEncontradoOutroSeparador = null;
  let pedidoJaConferido = null;

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];

    const pedidoPlanilha = texto_(linha[2]);

    if (!pedidoPlanilha) continue;

    if (normalizarPedido_(pedidoPlanilha) !== chavePedido) {
      continue;
    }

    const registro = {
      aba: aba,
      linha: i + 2,
      data: texto_(linha[0]),
      turno: texto_(linha[1]),
      pedido: pedidoPlanilha,
      separador: texto_(linha[3]),
      resultado: texto_(linha[13])
    };

    const mesmoSeparador =
      normalizarNome_(registro.separador) === chaveSeparador;

    const conferido =
      registro.resultado === 'SIM' ||
      registro.resultado === 'NÃO';

    if (!mesmoSeparador) {
      if (!pedidoEncontradoOutroSeparador) {
        pedidoEncontradoOutroSeparador = registro;
      }
      continue;
    }

    if (!conferido) {
      return registro;
    }

    if (!pedidoJaConferido) {
      pedidoJaConferido = registro;
    }
  }

  return pedidoJaConferido || pedidoEncontradoOutroSeparador;
}

function consultarPedido_(p) {
  const separador = texto_(p.separador || p.conferente);
  const pedido = texto_(p.pedido);

  if (!separador) {
    return {
      sucesso: false,
      erro: 'Separador não informado.'
    };
  }

  if (!pedido) {
    return {
      sucesso: false,
      erro: 'Pedido não informado.'
    };
  }

  if (!separadorCadastrado_(separador)) {
    return {
      sucesso: false,
      erro: 'Separador "' + separador + '" não está cadastrado.'
    };
  }

  const registro = localizarPedido_(pedido, separador);

  if (!registro) {
    return {
      sucesso: false,
      erro:
        'Pedido ' +
        pedido +
        ' não encontrado para o separador ' +
        separador +
        '.'
    };
  }

  if (
    normalizarNome_(registro.separador) !==
    normalizarNome_(separador)
  ) {
    return {
      sucesso: false,
      erro:
        'O pedido ' +
        registro.pedido +
        ' pertence ao separador "' +
        registro.separador +
        '".'
    };
  }

  if (registro.resultado === 'SIM' || registro.resultado === 'NÃO') {
    return {
      sucesso: false,
      jaConferido: true,
      erro:
        'O pedido ' +
        registro.pedido +
        ' já foi conferido.'
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
  const separador = texto_(p.separador || p.conferente);
  const pedido = texto_(p.pedido);
  const correta = texto_(p.correta).toUpperCase();

  if (!separador) {
    throw new Error('Separador não informado.');
  }

  if (!pedido) {
    throw new Error('Pedido não informado.');
  }

  if (correta !== 'SIM' && correta !== 'NÃO') {
    throw new Error('Resposta da conferência inválida.');
  }

  if (!separadorCadastrado_(separador)) {
    throw new Error(
      'Separador "' +
      separador +
      '" não está cadastrado.'
    );
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const registro = localizarPedido_(pedido, separador);

    if (!registro) {
      throw new Error(
        'Pedido ' +
        pedido +
        ' não encontrado na aba Lançamentos.'
      );
    }

    if (
      normalizarNome_(registro.separador) !==
      normalizarNome_(separador)
    ) {
      throw new Error(
        'O pedido ' +
        registro.pedido +
        ' pertence ao separador "' +
        registro.separador +
        '".'
      );
    }

    if (registro.resultado === 'SIM' || registro.resultado === 'NÃO') {
      return {
        sucesso: true,
        duplicado: true,
        pedido: registro.pedido,
        mensagem: 'Pedido já conferido.'
      };
    }

    /*
     * Data e turno:
     * normalmente já vêm no lançamento.
     * Se estiverem vazios, o próprio Link Separadores completa
     * somente esses dois campos para não deixar o registro sem data.
     */
    const agora = new Date();

    if (!registro.data) {
      registro.aba.getRange(registro.linha, 1).setValue(
        inicioDoDia_(agora)
      );
      registro.aba.getRange(registro.linha, 1).setNumberFormat('dd/MM/yyyy');
    }

    if (!registro.turno) {
      registro.aba.getRange(registro.linha, 2).setValue(
        obterTurno_(agora)
      );
    }

    /*
     * Somente K e N pertencem diretamente à confirmação do separador.
     *
     * K = Erro Detectado?
     * N = A separação está correta?
     *
     * Não alteramos:
     * D = Separador
     * E = Conferente
     * F:J e L:M = detalhes de erro/auditoria
     */
    registro.aba.getRange(registro.linha, 11)
      .setValue(correta === 'SIM' ? 'Não' : 'Sim');

    registro.aba.getRange(registro.linha, 14)
      .setValue(correta);

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      duplicado: false,
      pedido: registro.pedido,
      separador: registro.separador,
      correta: correta
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (erro) {}
  }
}

function registrarDireto_(p) {
  /*
   * Compatibilidade com a versão antiga do frontend.
   * O novo Link Separadores usa salvarConferencia_ e exige
   * a confirmação SIM/NÃO antes de gravar.
   */
  return salvarConferencia_(Object.assign({}, p, {
    correta: texto_(p.correta) || 'SIM'
  }));
}

function quantidadeHoje_(separador) {
  const nome = texto_(separador);

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

  const hoje = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );

  let quantidade = 0;

  dados.forEach(linha => {
    const data = texto_(linha[0]);
    const separadorLinha = texto_(linha[3]);
    const resultado = texto_(linha[13]);

    if (
      data === hoje &&
      normalizarNome_(separadorLinha) === normalizarNome_(nome) &&
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
    versao: '3.0',
    planilha: planilha.getName(),
    cadastroExiste: !!cadastro,
    lancamentosExiste: !!lancamentos,
    linhasCadastro: cadastro ? cadastro.getLastRow() : 0,
    linhasLancamentos: lancamentos ? lancamentos.getLastRow() : 0,
    estruturaLançamentos: 'A:N',
    listaSeparadores: 'Cadastro!A:A'
  };
}

function inicioDoDia_(data) {
  const texto = Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'yyyy,MM,dd'
  );

  const partes = texto.split(',');

  return new Date(
    Number(partes[0]),
    Number(partes[1]) - 1,
    Number(partes[2])
  );
}

function obterTurno_(data) {
  const hora = Number(
    Utilities.formatDate(
      data,
      Session.getScriptTimeZone(),
      'HH'
    )
  );

  return hora < 12 ? 'Manhã' : 'Tarde';
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

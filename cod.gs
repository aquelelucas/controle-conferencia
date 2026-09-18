const ABA_LANCAMENTOS = 'Lançamentos';
const ABA_CADASTRO = 'Cadastro';

const CHAVE_APP = 'KING-CONFERENCIA-2026';

/*
 * Link Separadores
 * Backend único do aplicativo.
 *
 * Estrutura esperada na aba "Lançamentos":
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
  const parametros = (e && e.parameter) ? e.parameter : {};
  const acao = String(parametros.acao || '').trim();
  const callback = String(parametros.callback || '').trim();

  try {
    switch (acao) {
      case '':
        return responder_({
          sucesso: true,
          mensagem: 'API Link Separadores funcionando.',
          versao: '2.0'
        }, callback);

      case 'conferentes':
      case 'separadores':
        return responder_(obterConferentes_(e, acao), callback);

      case 'consultarPedido':
        return responder_(consultarPedido_(e), callback);

      case 'salvarConferencia':
        return responder_(salvarConferencia_(e), callback);

      case 'registrar':
        // Compatibilidade com versões anteriores do aplicativo.
        return responder_(salvarConferencia_(e), callback);

      case 'quantidade':
        return responder_(quantidadeHoje_(e), callback);

      case 'diagnostico':
        return responder_(diagnostico_(e), callback);

      default:
        return responder_({
          sucesso: false,
          erro: 'Ação não reconhecida: ' + acao
        }, callback);
    }
  } catch (erro) {
    return responder_({
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro interno no servidor.'
    }, callback);
  }
}

function validarChave_(e) {
  const parametros = (e && e.parameter) ? e.parameter : {};
  return String(parametros.chave || '').trim() === CHAVE_APP;
}

function obterConferentes_(e, acao) {
  if (!validarChave_(e)) {
    return {
      sucesso: false,
      erro: 'Chave inválida.'
    };
  }

  try {
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
        ...(acao === 'separadores'
          ? {separadores: []}
          : {conferentes: []})
      };
    }

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

    const unicos = [...new Set(nomes)];

    return {
      sucesso: true,
      ...(acao === 'separadores'
        ? {separadores: unicos}
        : {conferentes: unicos})
    };
  } catch (erro) {
    return {
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro ao carregar os nomes do Cadastro.'
    };
  }
}

/*
 * Não existe uma tabela mestre de pedidos no modelo atual.
 * Portanto, a consulta apenas valida o código recebido e o devolve
 * para a tela de confirmação. O registro real acontece em salvarConferencia.
 */
function consultarPedido_(e) {
  if (!validarChave_(e)) {
    return {
      sucesso: false,
      erro: 'Chave inválida.'
    };
  }

  const parametros = (e && e.parameter) ? e.parameter : {};
  const pedido = String(parametros.pedido || '').trim();
  const conferente = String(
    parametros.conferente || parametros.separador || ''
  ).trim();

  if (!pedido) {
    return {
      sucesso: false,
      erro: 'Pedido não informado.'
    };
  }

  if (!conferente) {
    return {
      sucesso: false,
      erro: 'Conferente não informado.'
    };
  }

  return {
    sucesso: true,
    pedido: pedido,
    separador: conferente,
    conferente: conferente
  };
}

function salvarConferencia_(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!validarChave_(e)) {
      return {
        sucesso: false,
        erro: 'Chave inválida.'
      };
    }

    // Evita duas gravações simultâneas para o mesmo ambiente.
    lock.waitLock(10000);

    const parametros = (e && e.parameter) ? e.parameter : {};

    const pedido = String(parametros.pedido || '').trim();

    // Aceita os dois nomes para manter compatibilidade.
    const conferente = String(
      parametros.conferente || parametros.separador || ''
    ).trim();

    const separador = String(
      parametros.separador || parametros.conferente || ''
    ).trim();

    if (!pedido) {
      return {
        sucesso: false,
        erro: 'Pedido não informado.'
      };
    }

    if (!conferente) {
      return {
        sucesso: false,
        erro: 'Conferente não informado.'
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

    const resultado = String(
      parametros.correta || ''
    ).trim().toUpperCase();

    // Versões antigas do app não enviavam "correta".
    // Nessa situação, a gravação direta é tratada como conferência correta.
    const correta = resultado === 'NÃO' ? 'NÃO' : 'SIM';

    const tipoErro = String(parametros.tipoErro || '').trim();
    const gravidade = String(parametros.gravidade || '').trim();
    const sku = String(parametros.sku || '').trim();
    const acaoTomada = String(parametros.acaoTomada || '').trim();
    const observacao = String(parametros.observacao || '').trim();

    const qtdSolicitada = normalizarQuantidade_(parametros.qtdSolicitada);
    const qtdSeparada = normalizarQuantidade_(parametros.qtdSeparada);

    if (correta === 'NÃO' && !tipoErro) {
      return {
        sucesso: false,
        erro: 'Informe o tipo de erro.'
      };
    }

    if (correta === 'NÃO' && !gravidade) {
      return {
        sucesso: false,
        erro: 'Informe a gravidade do erro.'
      };
    }

    const hoje = obterDataHoje_();
    const hojeTexto = Utilities.formatDate(
      hoje,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

    // Procura o mesmo pedido para o mesmo conferente no dia.
    // O objetivo é impedir duplicidade por repetição do barcode.
    if (aba.getLastRow() >= 2) {
      const inicio = Math.max(2, aba.getLastRow() - 500);
      const quantidadeLinhas = aba.getLastRow() - inicio + 1;

      const registros = aba
        .getRange(inicio, 1, quantidadeLinhas, 5)
        .getDisplayValues();

      for (let i = registros.length - 1; i >= 0; i--) {
        const linha = registros[i];
        const data = String(linha[0] || '').trim();
        const pedidoRegistrado = String(linha[2] || '').trim();
        const conferenteRegistrado = String(
          linha[4] || linha[3] || ''
        ).trim();

        if (
          data === hojeTexto &&
          pedidoRegistrado === pedido &&
          conferenteRegistrado === conferente
        ) {
          return {
            sucesso: true,
            duplicado: true,
            pedido: pedido,
            separador: separador,
            conferente: conferente,
            mensagem: 'Pedido já registrado para este conferente hoje.'
          };
        }
      }
    }

    const turno = obterTurno_();
    const proximaLinha = aba.getLastRow() + 1;

    /*
     * Grava exatamente as 14 colunas do modelo.
     * A data é um objeto Date real, não texto, para que TODAY(),
     * COUNTIFS e os rankings funcionem corretamente.
     */
    const linha = [
      hoje,
      turno,
      pedido,
      separador,
      conferente,
      correta === 'NÃO' ? sku : '',
      correta === 'NÃO' ? qtdSolicitada : '',
      correta === 'NÃO' ? qtdSeparada : '',
      correta === 'NÃO' ? tipoErro : '',
      correta === 'NÃO' ? gravidade : '',
      correta === 'NÃO' ? 'Sim' : 'Não',
      correta === 'NÃO' ? acaoTomada : '',
      correta === 'NÃO' ? observacao : '',
      correta
    ];

    aba.getRange(proximaLinha, 1, 1, 14).setValues([linha]);

    // Formato da data para manter a leitura da planilha consistente.
    aba.getRange(proximaLinha, 1).setNumberFormat('dd/MM/yyyy');

    // Reforça o formato das quantidades quando forem numéricas.
    if (correta === 'NÃO') {
      aba.getRange(proximaLinha, 7, 1, 2).setNumberFormat('0.##');
    }

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      duplicado: false,
      pedido: pedido,
      separador: separador,
      conferente: conferente,
      data: Utilities.formatDate(
        hoje,
        Session.getScriptTimeZone(),
        'dd/MM/yyyy'
      ),
      hora: Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'HH:mm:ss'
      ),
      turno: turno,
      correta: correta
    };

  } catch (erro) {
    return {
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro desconhecido ao salvar a conferência.'
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (erro) {}
  }
}

function quantidadeHoje_(e) {
  if (!validarChave_(e)) {
    return {
      sucesso: false,
      erro: 'Chave inválida.'
    };
  }

  try {
    const parametros = (e && e.parameter) ? e.parameter : {};

    const conferente = String(
      parametros.conferente || parametros.separador || ''
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
      .getDisplayValues();

    const hojeTexto = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

    let quantidade = 0;

    dados.forEach(function(linha) {
      const data = String(linha[0] || '').trim();
      const nome = String(
        linha[4] || linha[3] || ''
      ).trim();

      if (
        data === hojeTexto &&
        (!conferente || nome === conferente)
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
        : 'Erro ao calcular a quantidade.'
    };
  }
}

function diagnostico_(e) {
  if (!validarChave_(e)) {
    return {
      sucesso: false,
      erro: 'Chave inválida.'
    };
  }

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const cadastro = planilha.getSheetByName(ABA_CADASTRO);
    const lancamentos = planilha.getSheetByName(ABA_LANCAMENTOS);

    return {
      sucesso: true,
      versao: '2.0',
      planilha: planilha.getName(),
      cadastroExiste: !!cadastro,
      lancamentosExiste: !!lancamentos,
      totalLinhasCadastro: cadastro ? cadastro.getLastRow() : 0,
      totalLinhasLancamentos: lancamentos ? lancamentos.getLastRow() : 0,
      estruturaEsperada: 'Lançamentos A:N'
    };
  } catch (erro) {
    return {
      sucesso: false,
      erro: erro && erro.message
        ? erro.message
        : 'Erro no diagnóstico.'
    };
  }
}

function obterDataHoje_() {
  const agora = new Date();
  const partes = Utilities
    .formatDate(
      agora,
      Session.getScriptTimeZone(),
      'yyyy,MM,dd'
    )
    .split(',');

  const ano = Number(partes[0]);
  const mes = Number(partes[1]) - 1;
  const dia = Number(partes[2]);

  return new Date(ano, mes, dia);
}

function obterTurno_() {
  const hora = Number(
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'HH'
    )
  );

  return hora <= 12 ? 'Manhã' : 'Tarde';
}

function normalizarQuantidade_(valor) {
  const texto = String(valor || '').trim();

  if (!texto) {
    return '';
  }

  const normalizado = texto.replace(',', '.');
  const numero = Number(normalizado);

  return Number.isFinite(numero) ? numero : texto;
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
# Link Separadores

Aplicativo simples para registrar a separação de pedidos.

## Fluxo oficial

1. O separador escolhe o próprio nome.
2. Toca em **ENTRAR**.
3. A câmera abre.
4. O separador bipará o código de barras do pedido.
5. O sistema registra automaticamente o pedido.
6. A câmera continua aberta para o próximo pedido.
7. Para encerrar, o separador toca em **FINALIZAR**.

## Dados registrados

Cada bip cria um novo registro na aba **Lançamentos**, somente nas colunas A a D:

| Coluna | Informação |
|---|---|
| A | Data |
| B | Turno |
| C | Pedido |
| D | Separador |

A tabela e as demais colunas existentes na planilha não fazem parte deste fluxo e não devem ser alteradas pelo aplicativo.

## Cadastro

A lista de separadores é lida da aba **Cadastro**, coluna A, a partir da linha 2.

## Backend

O backend oficial é o Google Apps Script publicado como Web App. O arquivo oficial do projeto é `Code.gs`.

O `index.html` é somente a interface do Link Separadores.

## Regra para alterações futuras

Não misturar este projeto com o projeto de conferência/auditoria. O Link Separadores não pergunta se a separação está correta, não preenche conferente, não grava erros e não modifica outras colunas da planilha.

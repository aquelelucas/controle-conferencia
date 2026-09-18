# Link Separadores

Aplicação operacional para conferência da separação de pedidos.

## Fluxo oficial

1. O separador seleciona o próprio nome na entrada.
2. Toca em **ENTRAR**.
3. A câmera é aberta automaticamente.
4. O separador bipará o código de barras do pedido.
5. A tela mostra o número lido e pergunta se a separação está correta.
6. Ao responder **SIM**, o resultado é gravado e a câmera fica pronta imediatamente para o próximo pedido.
7. Ao responder **NÃO**, o resultado também é gravado e a câmera fica pronta para o próximo pedido.
8. O Link Separadores não é o sistema de lançamento de pedidos e não deve substituir o Separador da coluna D nem preencher a coluna E de Conferente.

## Planilha

A aba principal é **Lançamentos**:

- A: Data
- B: Turno
- C: Pedido
- D: Separador
- E: Conferente
- F: SKU/Produto
- G: Qtd. Solicitada
- H: Qtd. Separada
- I: Tipo de Erro
- J: Gravidade
- K: Erro Detectado?
- L: Ação Tomada
- M: Observação
- N: A separação está correta?

A aba **Cadastro** usa a coluna A como lista oficial de separadores.

## Responsabilidade do Link Separadores

O aplicativo atua somente sobre a conferência da separação. Data, turno, pedido e separador pertencem ao lançamento do pedido. Ao confirmar, o aplicativo registra apenas o resultado da conferência na mesma linha, preservando os dados operacionais já existentes.

## Backend

O backend oficial é o **Google Apps Script** publicado como Web App. O arquivo oficial no repositório é `Code.gs`.

O `index.html` é a interface do aplicativo hospedada no GitHub.

## Regra importante para alterações futuras

Não misturar o Link Separadores com o projeto de usuários/conferentes, rankings ou formulários de detalhamento de erros. Qualquer nova função deve preservar o fluxo operacional simples de leitura -> confirmação -> próximo pedido.

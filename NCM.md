# NCM do catálogo CONTT.s

Preenchido em 15/09/2026, nas 1188 variantes dos 314 produtos da loja.

**Este documento é para o contador conferir.** NCM é classificação fiscal, e a
decisão final é dele. O que está abaixo é a classificação padrão do setor de
moda fitness, com o texto da posição ao lado para a conferência ser rápida.
Três pontos estão genuinamente em aberto e estão marcados como tal.

## Onde o NCM ficou gravado

No campo nativo da Shopify `harmonizedSystemCode`, do item de estoque de cada
variante. No painel: **Produtos → (produto) → variante → Envio → Código do
Sistema Harmonizado (HS)**.

A loja **não tem nenhum metafield fiscal** (conferido: só existem `custom.cor`,
`reviews.rating`, `shopify.size` e `shopify.color-pattern`). Esse campo nativo é
de onde os integradores de NF-e (Bling, Tiny e afins) leem o NCM. O campo aceita
de 6 a 13 dígitos, então o NCM de 8 cabe inteiro — está gravado sem pontos
(`61046300`), que é o formato que os integradores esperam.

## A tabela

Todos são do **Capítulo 61** (vestuário de malha), exceto os tops. A premissa
em todos é **malha de fibras sintéticas** — poliamida/poliéster com elastano,
que é o tecido de moda fitness.

| Grupo na loja | Variantes | NCM | Posição |
|---|---:|---|---|
| Conjuntos (`CONJUNTOS` + `Conjunto`) | 500 | `6104.23.00` | Conjuntos, de fibras sintéticas, de malha, para senhoras/meninas |
| Leggings e shorts | 419 | `6104.63.00` | Calças, jardineiras, bermudas e shorts, de fibras sintéticas, de malha, para senhoras/meninas |
| Tops | 193 | `6212.10.00` | Sutiãs, mesmo de malha |
| Macacões, blusas e croppeds | 52 | `6114.30.00` | Outras roupas, de malha, de fibras sintéticas |
| Short-saia | 21 | `6104.53.00` | Saias e saias-calças, de fibras sintéticas, de malha, para senhoras/meninas |
| Jaqueta | 3 | `6110.30.00` | Pulôveres, cardigãs, coletes e artigos semelhantes, de fibras sintéticas, de malha |

**Fora da tabela:** "Embalagem para presente" (1 variante) ficou sem NCM de
propósito — não é peça de vestuário, e dar a ela um NCM de roupa seria errado.
Se ela precisar sair em nota, o contador diz qual usar.

## Os três pontos em aberto

### 1. Top: sutiã (`6212.10.00`) ou "outras roupas" (`6114.30.00`)?

Este é o que mais pesa, porque **`6212.10.00` costuma estar na lista de
substituição tributária de confecções em vários estados** — o regime muda,
não só o código.

Usei `6212.10.00` porque o top esportivo da loja tem bojo e função de
sustentação, e nesse caso ele funciona como sutiã. É o que a maior parte das
marcas de fitness usa. Mas há quem classifique top esportivo em `6114.30.00`.
**Confirmar com o contador, olhando o regime de ST do estado.**

A separação entre os 74 produtos de "TOP E CROPPED" foi feita pelo título:
- os 70 que começam com "Top" → `6212.10.00` (sutiã)
- os 4 que começam com "Blusa" ou "Cropped" → `6114.30.00` (peça de vestir, não sutiã)

Os 4 são: *Blusa amarração cropped manga curta*, *Blusa cropped manga longa*,
*Cropped gola alta electric rose*, *Cropped gola alta urban ink*.

### 2. Conjunto: `6104.23.00` depende da Nota 3(b) do Capítulo 61

A posição de "conjunto" exige peças do mesmo tecido, mesmo tamanho, vendidas
juntas e complementares — o que descreve exatamente os conjuntos legging+top
da loja, que são um SKU só.

O detalhe: se o top isolado é sutiã (posição 6212, fora do Capítulo 61), dá
para discutir se o conjunto ainda se enquadra como "conjunto" do 6104. Na
prática o setor usa `6104.23.00` normalmente. **Ponto 1 e ponto 2 andam juntos:
se o contador mudar o top, vale rever o conjunto.**

### 3. Composição do tecido

Toda a tabela assume fibras sintéticas. Se alguma peça for predominantemente
de **algodão**, a posição muda — legging de algodão, por exemplo, seria
`6104.62.00` e não `6104.63.00`.

As linhas **caneladas** (Legging canelada, Macacão canelado comum, Top milano
canelado) são as candidatas mais prováveis a ter algodão na composição. Vale
conferir a etiqueta dessas antes de fechar.

## Se o contador mudar alguma coisa

A tabela vive num lugar só: a função `classifica()` em `preenche-ncm.py`.
Mudando lá e rodando de novo, o catálogo inteiro é reclassificado — não precisa
mexer produto por produto.

```
python3 preenche-ncm.py catalogo.jsonl ncm.jsonl      # gera
python3 gera-lotes-ncm.py ncm.jsonl lotes/            # quebra em lotes
```

O `catalogo.jsonl` sai de uma operação em lote da Shopify (`bulkOperationRunQuery`)
com id, título, tipo e item de estoque de cada variante.

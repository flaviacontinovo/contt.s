# Planilha da Amazon — CONTT.s

Arquivo: `amazon-contts.xlsm`, gerado em 16/09/2026 e corrigido em 17/09/2026 a partir do modelo
`PANTS_ONE_PIECE_OUTFIT_SHORTS_BRA_APPAREL` que você baixou do Seller Central,
preenchido com o catálogo da Shopify.

**1353 linhas** na aba *Modelo*, da linha 7 em diante: **244 produtos pai** e
**1109 ofertas** (1055 variações + 54 produtos sem variação).

Só a aba *Modelo* foi tocada. As 712 validações de dados, a formatação
condicional, as abas ocultas com as listas suspensas e as linhas 1 a 6
(cabeçalhos e exemplo) continuam byte a byte como vieram — a Amazon recusa
arquivo que não seja exatamente o que ela gerou.

## Como os produtos foram classificados

| Shopify | Amazon | Linhas |
|---|---|---:|
| Conjuntos, jaquetas, blusas e croppeds | `APPAREL` | 500 |
| Leggings | `PANTS` | 390 |
| Tops (sutiã esportivo) | `BRA` | 256 |
| Shorts e short-saia | `SHORTS` | 159 |
| Macacões | `ONE_PIECE_OUTFIT` | 48 |

Conjunto legging+top foi para `APPAREL`, não `ONE_PIECE_OUTFIT`: são duas peças,
e a posição de peça única estaria errada.

## O que foi preenchido

Todos os 14 campos obrigatórios estão completos em todas as linhas. Além deles:
preço, estoque, imagens, cor, tamanho, marca, descrição, tópicos, palavras-chave,
departamento, gênero, faixa etária, cuidados, país de origem e condição.

- **Preço**: o da Shopify. Onde havia preço comparativo maior, ele entrou como
  preço sugerido (60 ofertas).
- **Estoque**: o da Shopify. Nenhuma oferta saiu com zero.
- **Imagens**: a da própria variante quando existe, senão a principal do produto.
  Até 5 imagens adicionais por linha.
- **Tópicos** (obrigatórios): saíram das frases da sua própria descrição, mais
  dois fatos conferidos (tamanhos e cores disponíveis). Nada foi inventado.

## Quatro coisas que dependem de você

### 1. 449 ofertas sem código de barras

660 ofertas têm EAN e foram como `EAN`. As outras **449 foram como "Isento de
GTIN"**, que é a opção correta da lista quando não há código — mas a Amazon só
aceita isenção se a marca CONTT.s tiver a **isenção de GTIN aprovada** no Seller
Central. Sem essa aprovação, essas 449 linhas voltam com erro.

Vale pedir a isenção antes de subir, ou subir primeiro só as 660 com EAN.

### 2. Composição do tecido

O campo é obrigatório. Onde a descrição declara a composição, ela foi extraída.
Onde não declara — **1307 das 1353 linhas** — entrou `Malha`, que é verdadeiro
para todas as peças mas genérico.

Não inventei percentuais. Se você me passar a composição por linha de produto
(ou ela estiver na etiqueta), eu troco de uma vez.

### 3. Peso e dimensões da embalagem: em branco

Deixei em branco de propósito. O peso na Shopify não está confiável: **520
variantes com 0,0 kg** e algumas com 0,01 kg — 10 gramas numa legging. Preencher
com isso daria frete errado.

São campos condicionalmente obrigatórios. Se a Amazon reclamar, me diga o peso e
as medidas médias por tipo de peça e eu preencho.

### 4. Tamanho dos conjuntos: dois eixos virando um

428 ofertas são conjuntos com **dois tamanhos independentes** (top e legging). A
Amazon aceita só um eixo de tamanho por tema de variação, então os dois viraram
um texto só: `Legging M / Top G`.

Funciona, mas fica estranho no filtro de tamanho da Amazon. A alternativa é
vender conjunto com tamanho único por peça. É decisão comercial sua.

Também ficaram em branco, porque dependem da configuração da sua conta:
**modelo de envio** e **tempo de processamento** (compromisso de prazo que é
seu, não meu).

## 16 produtos ficaram de fora

15 são **rascunhos na Shopify** — não estão publicados nem na sua loja, então não
fazia sentido publicá-los na Amazon. Os 10 produtos sem nenhuma imagem estão
todos entre esses rascunhos.

O 16º é a **"Embalagem para presente"**, que não é peça de vestuário.

Se quiser qualquer um deles na Amazon, é só publicar na Shopify (e pôr imagem) e
eu rodo de novo.


## Correção de 17/09/2026 — por que a Amazon pedia dado peça por peça

A primeira versão declarava tema de variação **"COR/TAMANHO" em 235 das 244
famílias onde a cor não varia**: um "Conjunto Aura Bronze" tem nove tamanhos e
uma cor só. A Amazon procurava um diferenciador de cor que não existia, não
achava, e passava a cobrar a cor item a item. Era erro meu, não campo faltando.

Agora o tema cita só o que de fato varia entre os filhos:

| tema | famílias |
|---|---:|
| `TAMANHO` | 235 |
| `TAMANHO/COR` | 5 |
| `COR/TAMANHO` | 2 |
| `COR` | 2 |

E quando a cor é constante na família, ela passou a ir **na linha do produto
pai** também, que é onde a Amazon espera encontrá-la nesse caso.

### Campos preenchidos nesta rodada

- **Marca**: `CONTT.s` em todas as linhas (antes vinha do fornecedor da Shopify,
  que trazia "CONTT.s FITNESS WEAR", "Allure®" e "J-winner" misturados).
- **Estilo**: lista fechada e diferente por tipo. Nenhuma opção descreve moda
  fitness direito, então foi a menos errada de cada uma: `Moderna` (calça),
  `Shorts híbridos` (short), `Moderno` (top), `Esportivo` (conjunto e macacão).
  Se preferir outra, é trocar uma linha do script.
- **Altura da cintura** (calça e short): saiu do próprio título — "cintura
  média" vira `Cintura média`, o resto fica `Cintura alta`, que é o padrão da loja.
- **Sistema de tamanho** (calça e short): `BR`, classe `Alfa`, e o valor é o
  tamanho da peça de baixo — nos conjuntos, o "M" de "Legging M / Top G".
- **País do tamanho**: `Brasil`. **Nome do modelo**: o título da peça.

### Embalagem: número meu, confira antes de confiar

Peso e medidas estavam em branco porque o dado da Shopify não presta (520
variantes em 0,0 kg). Como a Amazon cobra o campo, entrou uma estimativa por
tipo de peça:

| tipo | C × L × A | peso |
|---|---|---|
| Legging | 30 × 22 × 4 cm | 220 g |
| Short | 26 × 20 × 3 cm | 130 g |
| Top | 24 × 18 × 3 cm | 110 g |
| Macacão | 30 × 22 × 5 cm | 260 g |
| Conjunto, jaqueta, cropped | 32 × 24 × 6 cm | 330 g |

**Isso mexe no frete.** Pesa um envelope de cada tipo na balança e me diz os
valores reais — troco na tabela `EMBALAGEM` do script e regero em um minuto.

### Ainda em branco de propósito

**Modelo de Envio (BR)** é texto livre que precisa bater com o nome exato de um
modelo de envio cadastrado no seu Seller Central. Não tenho como adivinhar esse
nome; me diga qual é e eu preencho nas 1353 linhas.

## Refazer

```
python3 preenche-amazon.py shopify.jsonl <modelo-baixado>.xlsm amazon-contts.xlsm
```

`shopify.jsonl` sai de uma operação em lote da Shopify com título, descrição,
marca, tipo, imagens e as variantes com SKU, preço, estoque, código de barras,
opções e imagem.

O script confere antes de gravar: nenhum valor fora das listas fechadas, nenhuma
SKU repetida, nenhum título acima de 200 caracteres, nenhum filho sem pai. Se
algo não bate, ele para em vez de gerar arquivo quebrado.

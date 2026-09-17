# Planilha da Amazon — CONTT.s

Arquivo: `amazon-contts.xlsm`, gerado em 16/09/2026, corrigido e ampliado em 17/09/2026 a partir do modelo
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

### Embalagem

Caixa padrão da loja, a mesma para toda peça: **20 × 20 × 5 cm** (comprimento ×
largura × altura), em centímetros.

O **peso** continua estimativa minha, por tipo de peça, porque o da Shopify não
serve — 520 variantes estão com 0,0 kg:

| tipo | peso |
|---|---|
| Legging | 220 g |
| Short | 130 g |
| Top | 110 g |
| Macacão | 260 g |
| Conjunto, jaqueta, cropped | 330 g |

Peso mexe no frete. Se quiser exato, pesa uma peça de cada tipo e me diz — troco
na tabela `PESO` do script.

### Envio

**Canal de processamento**: `Logística do vendedor (Padrão)` nas 1109 ofertas —
envio por você, não por Logística da Amazon.

**Modelo de Envio**: em branco de propósito. A definição do próprio modelo diz
que *"a Amazon atribui um modelo padrão"*, então em branco significa usar o seu
modelo padrão — que é o que você quer. Só vale preencher se você criar modelos
adicionais e quiser escolher um deles peça por peça.


## Atributos de produto (17/09/2026)

**77 das 356 colunas preenchidas**, contra 56 na primeira versão. As respostas
da Flavia entraram assim:

| Ela respondeu | Onde foi |
|---|---|
| Poliamida + elastano | Material (`Mistura de nylon` + `Elastano`) e Tipo de tecido, que deixou de ser o genérico "Malha" |
| Levantamento de bumbum, à prova de agachamento, absorção de suor, respirável | Características especiais |
| Apoio alto | Nível de apoio dos 256 tops |
| Opaco | Opacidade |

**Compressão ficou de fora de propósito** — ela já tinha dito que os conjuntos
de R$ 189 não são de alta compressão, e a característica não foi marcada.

Derivado do catálogo, sem perguntar: estampa (poá → `Bolinhas`, onça → `Estampa
animal`, tie dye → `Tie dye`), quantidade de peças (conjunto = 2), tecelagem
(`knitted`), estilo e forma da perna, comprimento, temporada, tipo de esporte,
número de bolsos, manga, costas e gola.

Cada característica entra só onde faz sentido: "levantamento de bumbum" não vai
em top, "à prova de agachamento" não vai em cropped.

### As listas mudam de um tipo para outro

Foi a armadilha desta rodada. SHORTS escreve `Estampa de animal` onde PANTS
escreve `Estampa animal`; SHORTS não tem `Todas as estações`; MACACÃO não tem
`Absorção de suor`, só `Respirável`. Escrever o valor do tipo errado faz a
Amazon recusar a linha.

`extrai-listas-amazon.py` tira as listas do próprio modelo (116 a 135 colunas
por tipo) e o preenchedor resolve cada valor contra a lista certa, com
sinônimos de reserva. Hoje: **zero valores fora da lista** nas 1353 linhas.

### Os 31 que continuam vazios

Bateria, lítio, GHS, materiais perigosos, tamanho de shapewear, aro de sutiã,
nome de time e atleta. Não é teimosia: **o campo "Contém Bateria ou Célula" só
oferece `battery` e `cell`** — não existe opção "não tem". E responder "Não" em
*"A bateria é antiderramamento?"* afirma que existe uma bateria.

Nenhum deles é obrigatório para roupa, e nenhum bloqueia o upload.


## Fotos e compressão (17/09/2026, segunda rodada)

**89 colunas preenchidas.**

### As fotos estavam no lugar errado

O modelo tem duas famílias de coluna de imagem, e eu só tinha usado uma:

| colunas | o que são | antes | agora |
|---|---|---|---|
| `AC`–`AL` | imagens do **produto** — as que a Amazon mostra na página | 0 | 1353 |
| `HE`–`HJ` | imagens da **oferta** | 1353 | 1353 |

As do produto aceitam **8 fotos extras**, contra 5 da oferta, então produtos com
muitas fotos na Shopify agora aproveitam todas: 49 linhas subiram com 9 fotos.

`AL` é a imagem de amostra de cor, e só entra nas 106 linhas onde a variante tem
foto própria **e** a cor é o que varia na família — fora disso não é amostra de
nada.

### Compressão

A Amazon **não tem nível de compressão**: a característica é "tem" ou "não tem".
"Média compressão" não existe no vocabulário dela. Como a loja confirmou que as
peças têm compressão média, a característica `Compressão` entrou nas 390 leggings,
e os 159 shorts receberam `Shorts de Compressão` no campo de forma, que é o
equivalente deles.


## 116 colunas (terceira rodada de 17/09/2026)

Varredura em todas as 356 colunas. Entrou tudo que tem valor verdadeiro para o
catálogo:

- **Público-alvo** `Mulheres`, **estilo de vida** `Casual`, **formalidade** `Casual`
- **Tipo de ajuste** `Equipado`, **elasticidade** `Esticável`, **ajuste ao tamanho**
  `Veste de acordo com o tamanho`
- **Usos específicos** `Esportes`, `Yoga`, `Caminhada`
- **Tem bolsos** `Sim`/`Não`, tirado do título
- **Tema** `Animais` nas 16 de onça, zebra e animal print
- **Tipo de trama** `Canelado` nas 74 caneladas
- **Comprimento da manga** `Sem mangas` nas 102 regatas e nadador
- **Função do sutiã** `sports` e tipo `other` nos 256 tops
- **Comprimento do short** `not_covering_the_knee` nos 159 shorts
- **Negativas que são verdade**: pele de animal `Não`, feito à mão `Não`,
  restrição de exportação `Não`, envio global `Não`
- **Origem fiscal** `0` (nacional), coerente com país de origem Brasil

### O que sobra vazio, e por quê

**Campos de produto usado ou recondicionado** (grau de renovado, condição
cosmética, peças substituídas, tipo de fonte): só valem quando a condição não é
"Novo". As suas são todas novas.

**Campos de outra categoria**: dimensões de tela, década da moda, desgaste do
tecido (lavagem ácida, bigodado), tipo de feriado, evento da vida, enfeite,
forro de velo.

**Campos que variam peça a peça e eu não vejo as fotos**: tipo de alça, formato
e acolchoamento do bojo, estilo das costas e da gola onde o título não diz,
descrição do bolso, orientação do fecho. Dá para fazer depois, produto a
produto, se valer a pena.

**Campos fiscais que dependem do contador**: Código EX da TIPI, indicador da
origem do processo, atividade da empresa (fabricante, distribuidor ou
importador).

**Os 31 de bateria, lítio, GHS, shapewear, aro, time e atleta**: sem valor
verdadeiro para roupa, como já explicado acima.


## 136 colunas — rodada de perguntas

O que a Flavia respondeu, e onde entrou:

| Resposta | Campos |
|---|---|
| Bojo removível | Acolchoamento `Média`, cobertura `Cobertura total`, componente `Almofada` (256 tops) |
| Alça não ajustável | Tipo de alça |
| Fabricante | Atividade da empresa, e o campo Fabricante com `CONTT.s` |
| FPU 50+ | Fator de proteção UV e a característica `Proteção solar` |
| Só elástico, sem fecho | Tipo de fecho: `Fechamento Elástico` na legging, `Pull on` no short, `Sem fechamento` no top |
| Embalagem de presente sim | Pode ir como presente e embalagem disponível |
| 1 dia útil | Tempo de processamento nas 1109 ofertas |
| CSOSN 102, sem CEST | Situação tributária |
| flaviacarolineconti@gmail.com | Contato de conformidade |

### O NCM veio de graça

A planilha tem coluna de **Código NCM**, e o NCM das 1189 variantes já estava
gravado na Shopify desde 15/09. O preenchedor lê de lá em vez de redigitar — as
1353 linhas saem com o NCM, e se o contador mudar a classificação, muda num
lugar só e as duas coisas acompanham.

### Ainda esperando

- **Tabela de medidas** (cintura, comprimento da perna, abertura da barra): ela
  tem e vai passar. São campos de refinamento, não bloqueiam.
- **NCM 6104.23.00 dos conjuntos** e a dúvida do top ser sutiã (6212.10.00),
  ambos com o contador — ver NCM.md.

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

# Códigos de anúncios do Mercado Livre

## codigos-anuncios-mercado-livre.csv

328 códigos MLB extraídos da planilha **"Acompanhamento de Full - Mercado Livre |
Ecommerce Puro"** (Google Drive, compartilhada por `willian.pagane@ecommercepuro.com.br`).

O mesmo conteúdo foi publicado como planilha no Google Drive da Flávia:
*Códigos de Anúncios Mercado Livre*.

### Origem dos códigos

Os códigos vêm de duas abas que **não têm nenhum código em comum** entre si:

| Origem | Códigos | O que traz |
|---|---:|---|
| `items_list` | 220 | Só o código. A coluna SEARCHED DATE está vazia — fila ainda não consultada. |
| `prices` | 108 | Preço, desconto, campanha e visitas, todos com data 17/12/2025 17:59:02. |

A ausência total de interseção é o motivo de a coluna `ORIGEM` existir: não dá para
afirmar que os dois conjuntos descrevem o mesmo catálogo, então nada foi mesclado.

### Colunas

`CÓDIGO MLB` · `ORIGEM` · `PREÇO DE TABELA` · `PREÇO VIGENTE` · `DESCONTO` ·
`TIPO DE PROMOÇÃO` · `VISITAS` · `DATA DA CONSULTA`

`PREÇO DE TABELA` (REGULAR AMOUNT na origem) só é preenchido quando há promoção
ativa; `PREÇO VIGENTE` (AMOUNT) é o preço praticado em todos os casos.

### Link do anúncio

Não gravado como coluna — é derivável do código. Para gerar na planilha:

    ="https://produto.mercadolivre.com.br/MLB-"&MID(A2,4,20)&"-_JM"

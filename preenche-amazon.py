#!/usr/bin/env python3
"""
Preenche o modelo de inventario da Amazon (PANTS/ONE_PIECE_OUTFIT/SHORTS/BRA/
APPAREL, Amazon.com.br) com o catalogo da Shopify.

Entrada : shopify.jsonl  (export do bulkOperationRunQuery)
          o .xlsm original, baixado do Seller Central
Saida   : uma copia do .xlsm com a aba "Modelo" preenchida da linha 7 em diante

Por que mexer no XML na mao em vez de usar openpyxl: o arquivo traz 712
validacoes de dados, formatacao condicional, 8543 nomes definidos e as abas
ocultas com as listas suspensas. openpyxl perde parte disso na regravacao, e a
Amazon rejeita arquivo que nao seja o que ela gerou. Aqui so as linhas de dados
sao inseridas; o resto do zip e copiado byte a byte.

Linha 1 = configuracoes codificadas da Amazon (nao se toca).
Linha 4 = rotulos, linha 5 = nomes tecnicos, linha 6 = exemplo (fica).
Linha 7 em diante = dados.
"""
import json, re, html, zipfile, shutil, sys, collections, unicodedata

MODELO = 'xl/worksheets/sheet5.xml'
PRIMEIRA = 7

# ---------------------------------------------------------------- classificacao

def tipo_amazon(titulo, ptipo):
    """Tipo de produto da Amazon a partir do tipo da Shopify."""
    t = (ptipo or '').upper()
    tit = (titulo or '').lower()
    if t.startswith('CONJUNTO'):
        # conjunto legging+top e duas pecas: ONE_PIECE_OUTFIT seria errado
        return 'APPAREL'
    if t.startswith('SHORT'):          # inclui SHORT-SAIA
        return 'SHORTS'
    if t.startswith('LEGGING'):
        return 'PANTS'
    if t.startswith('MACAC'):
        return 'ONE_PIECE_OUTFIT'
    if t.startswith('JAQUETA'):
        return 'APPAREL'
    if t.startswith('TOP E CROPPED') or (not t and tit.startswith('top')):
        return 'APPAREL' if tit.startswith(('blusa', 'cropped')) else 'BRA'
    return None                        # Embalagem e qualquer tipo novo


# As 34 cores da loja -> as 20 do mapa de cores da Amazon. A ordem dos testes
# importa: "Verde Lime Green" tem verde e lime, e precisa cair em Verde.
def cor_amazon(nome):
    n = (nome or '').lower()
    if 'preto' in n or 'preta' in n:                                  return 'Preto'
    if 'off white' in n or 'offwhite' in n:                           return 'Branco Off'
    if 'branco' in n or 'branca' in n:                                return 'Branco'
    if 'cinza' in n or 'chumbo' in n or 'grafite' in n:               return 'Cinza'
    if 'bronze' in n:                                                 return 'Bronze'
    if 'marrom' in n or 'terracota' in n or 'café' in n or 'chocolate' in n: return 'Marrom'
    if 'nude' in n or 'bege' in n or 'capuccino' in n or 'manteiga' in n or 'creme' in n: return 'Bege'
    if 'vermelh' in n or 'cereja' in n or 'vinho' in n:               return 'Vermelho'
    if 'rosa' in n or 'pink' in n or 'framboesa' in n or 'rosé' in n: return 'Rosa'
    if 'lavanda' in n or 'lilás' in n or 'roxo' in n or 'uva' in n:   return 'Roxo'
    if 'turquesa' in n or 'tiffany' in n:                             return 'Turquesa'
    if 'azul' in n or 'marinho' in n or 'jeans' in n or 'petról' in n or 'denim' in n: return 'Azul'
    if 'verde' in n or 'matcha' in n or 'militar' in n or 'lime' in n or 'oliva' in n: return 'Verde'
    if 'laranja' in n or 'orange' in n or 'vulcano' in n or 'coral' in n: return 'Laranja'
    if 'amarel' in n or 'mostarda' in n or 'dourado' in n:            return 'Amarelo'
    if 'estampad' in n or 'tie dye' in n or 'animal print' in n or 'print' in n: return 'Multicolorido'
    return 'Multicolorido'   # so sobra estampa; nunca chuta uma cor lisa errada



# Estilo (AU): lista fechada e diferente por tipo. Nenhuma opcao descreve moda
# fitness direito, entao vai a menos errada de cada lista.
ESTILO = {
    'PANTS': 'Moderna', 'SHORTS': 'Shorts híbridos', 'BRA': 'Moderno',
    'APPAREL': 'Esportivo', 'ONE_PIECE_OUTFIT': 'Esportivo',
}

# Embalagem padrao da loja: 5 x 20 x 20 cm, a mesma para toda peca (informada
# pela Flavia em 17/09/2026). Vai como comprimento 20, largura 20, altura 5.
CAIXA = (20, 20, 5)

# O peso continua por tipo e continua ESTIMADO: o da Shopify nao serve (520
# variantes em 0,0 kg) e ela passou so as medidas da caixa. Em gramas.
PESO = {
    'PANTS': 220, 'SHORTS': 130, 'BRA': 110,
    'ONE_PIECE_OUTFIT': 260, 'APPAREL': 330,
}

MARCA = 'CONTT.s'


def altura_cintura(titulo):
    """Coluna FM, lista fechada. Sai do proprio titulo da peca."""
    t = (titulo or '').lower()
    if 'cintura m' in t:                      # media / média
        return 'Cintura média'
    if 'cintura baixa' in t:
        return 'Cintura baixa'
    return 'Cintura alta'   # o padrao da loja: quase tudo e cintura alta


def tamanho_de_baixo(tam):
    """
    Coluna FF. Nos conjuntos o tamanho e combinado ("Legging M / Top G");
    aqui interessa so a peca de baixo.
    """
    for parte in (tam or '').split(' / '):
        baixo = parte.lower()
        if baixo.startswith(('legging', 'short')):
            return parte.split()[-1]
    return tam



# ---------------------------------------------------------------- atributos da peca

# Respostas da Flavia em 17/09/2026. Compressao ficou de fora de proposito:
# ela ja tinha dito que os conjuntos de R$ 189 nao sao de alta compressao.
MATERIAL = ('Mistura de nylon', 'Elastano')   # poliamida + elastano
TECIDO_PADRAO = 'Poliamida com elastano'
APOIO_TOP = 'Alto'
OPACIDADE = 'Opaco'

# Estampa (BV), lista fechada. So o que a estampa realmente e.
ESTAMPAS = [
    (('poá', 'poa'),                                   'Bolinhas'),
    (('onça', 'onca', 'animal print', 'zebra'),        'Estampa animal'),
    (('tie dye',),                                     'Tie dye'),
    (('butterfly',),                                   'Insetos'),
    (('oasis', 'tropical', 'reflorescer', 'green love'), 'Plantas'),
    (('newspaper',),                                   'Impressão de letra'),
    (('gibi',),                                        'Desenho animado'),
    (('coração', 'coracao'),                           'Coração'),
    (('abstract', 'expressive', 'grafiatto', 'colmeia'), 'Geométrico'),
]


def estampa_de(titulo, cores, ptipo):
    t = (titulo or '').lower()
    for chaves, valor in ESTAMPAS:
        if any(k in t for k in chaves):
            return valor
    if 'ESTAMPADO' in (ptipo or '').upper() or {'Estampado', 'Tie Dye'} & set(cores):
        return 'Gráfico'      # estampada, mas sem familia reconhecida
    return 'Liso'


def caracteristicas(titulo, tipo):
    """
    Colunas DO a DS (cinco vagas). Cada vaga devolve uma lista de sinonimos,
    porque a mesma ideia tem nome diferente conforme o tipo: SHORTS diz
    "Resistente ao suor" onde PANTS diz "Absorcao de suor", e MACACAO nao tem
    nenhum dos dois. Quem resolve qual cabe e a funcao escolhe().

    So entra o que vale para aquela peca: "levantamento de bumbum" nao faz
    sentido num top, "a prova de agachamento" nao faz num cropped. E so o que a
    loja confirmou - compressao ficou de fora a pedido dela, e "secagem rapida"
    nao foi confirmada, entao a quinta vaga fica vazia.
    """
    t = (titulo or '').lower()
    vagas = []
    if 'empina bumbum' in t:
        vagas.append(['Levantamento de Bumbum', 'Bumbum Apertado'])
    if tipo in ('PANTS', 'SHORTS'):
        vagas.append(['À prova de agachamento', 'Elasticidade'])
    if tipo in ('PANTS', 'SHORTS'):
        # A Amazon nao tem nivel de compressao: ou a peca tem, ou nao tem.
        # A loja confirmou media compressao, entao a caracteristica entra.
        vagas.append(['Compressão'])
    vagas.append(['Absorção de suor', 'Resistente ao suor'])
    vagas.append(['Respirável'])
    return (vagas + [[]] * 5)[:5]


def manga_de(titulo):
    t = (titulo or '').lower()
    if 'manga longa' in t: return 'Manga longa'
    if 'manga curta' in t: return 'Manga curta'
    if 'regata' in t or 'nadador' in t or 'alças' in t or 'tiras' in t: return 'Sem manga'
    return ''


def costas_de(titulo):
    t = (titulo or '').lower()
    if 'nadador' in t:   return 'Costas Nadador'
    if 'amarração' in t: return 'Costas com amarração'
    if 'tiras' in t:     return 'Costas com tiras'
    return ''


def gola_de(titulo):
    return 'Gola alta' if 'gola alta' in (titulo or '').lower() else ''


def bolsos_de(titulo):
    t = (titulo or '').lower()
    return 1 if ('bolso' in t or 'bolsinho' in t) else 0



# ---------------------------------------------------------------- listas fechadas

# {tipo: {coluna: [opcoes]}}, extraido do proprio modelo por extrai-listas-amazon.py.
# Carregado em main(); vazio significa "nao sei", e ai nada e escrito nas colunas
# de lista, o que e melhor que escrever palpite.
LISTAS = {}


def escolhe(col, tipo, *candidatos):
    """
    Devolve o primeiro candidato que a lista daquele TIPO aceita.

    As listas mudam entre tipos, e nao so no conteudo: SHORTS diz
    "Estampa de animal" onde PANTS diz "Estampa animal", e SHORTS nao tem
    "Todas as estacoes". Escrever o valor do tipo errado faz a Amazon recusar
    a linha, entao aqui o valor e sempre conferido contra a lista certa.

    Coluna de texto livre aceita o primeiro candidato. Se nenhum candidato
    servir, devolve '' - campo vazio passa, campo invalido nao.
    """
    opcoes = LISTAS.get(tipo, {}).get(col)
    candidatos = [c for c in candidatos if c]
    if opcoes is None:                       # texto livre
        return candidatos[0] if candidatos else ''
    for c in candidatos:
        if c in opcoes:
            return c
        # a mesma ideia escrita diferente entre tipos ("Estampa de animal")
        alvo = c.lower().replace(' de ', ' ')
        for o in opcoes:
            if o.lower().replace(' de ', ' ') == alvo:
                return o
    return ''


# ---------------------------------------------------------------- texto

def texto_limpo(bruto):
    t = re.sub(r'<(br|/p|/div|/li)[^>]*>', ' ', bruto or '', flags=re.I)
    t = re.sub(r'<[^>]+>', ' ', t)
    t = html.unescape(t)
    t = ''.join(c for c in t if c == '\n' or unicodedata.category(c)[0] != 'C')
    return re.sub(r'\s+', ' ', t).strip()


def topicos(descricao, tamanhos, cores, tecido):
    """
    Os 5 topicos sao obrigatorios. Saem das frases da propria descricao da loja
    mais dois fatos conferidos (tamanhos e cores disponiveis) - nada inventado.
    """
    frases = [f.strip() for f in re.split(r'(?<=[.!?])\s+', descricao) if len(f.strip()) > 25]
    fora = [f[:500] for f in frases[:3]]
    if tamanhos:
        fora.append('Tamanhos disponíveis: ' + ', '.join(tamanhos))
    if cores:
        fora.append('Cor: ' + ', '.join(cores) if len(cores) == 1
                    else 'Cores disponíveis: ' + ', '.join(cores))
    if tecido:
        fora.append('Tecido: ' + tecido)
    return (fora + [''] * 5)[:5]


def tecido_de(descricao):
    """Composicao so quando a propria descricao declara. Senao, o generico."""
    m = re.search(r'(\d{1,3}\s*%[^.;]{0,60}(?:poliamida|poliéster|poliester|algodão|algodao|elastano)'
                  r'(?:[^.;]{0,60}\d{1,3}\s*%[^.;]{0,40})?)', descricao, re.I)
    if m:
        return re.sub(r'\s+', ' ', m.group(1)).strip(' ,.;')[:200]
    return TECIDO_PADRAO   # composicao confirmada pela loja; ver AMAZON.md


def palavras_chave(titulo, ptipo, cores):
    base = [p for p in re.split(r'\W+', (titulo or '').lower()) if len(p) > 3]
    extra = ['fitness', 'academia', 'moda fitness feminina', 'roupa de treino']
    vistas, fora = set(), []
    for p in base + extra + [c.lower() for c in cores]:
        if p not in vistas:
            vistas.add(p); fora.append(p)
    return '; '.join(fora)[:500]


# ---------------------------------------------------------------- escrita XML

def esc(v):
    return (str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


def celula(col, linha, valor):
    if valor is None or valor == '':
        return ''
    if isinstance(valor, (int, float)):
        return f'<c r="{col}{linha}"><v>{valor}</v></c>'
    return f'<c r="{col}{linha}" t="inlineStr"><is><t xml:space="preserve">{esc(valor)}</t></is></c>'


def num_col(letra):
    n = 0
    for c in letra: n = n * 26 + ord(c) - 64
    return n


# ---------------------------------------------------------------- montagem

def carrega_shopify(caminho):
    prods, filhos = {}, collections.defaultdict(list)
    for ln in open(caminho, encoding='utf-8'):
        o = json.loads(ln)
        if '__parentId' in o: filhos[o['__parentId']].append(o)
        else: prods[o['id']] = o
    for p in prods.values():
        fs = filhos.get(p['id'], [])
        p['_variantes'] = [f for f in fs if 'sku' in f]
        p['_imagens'] = [f['image']['url'] for f in fs
                         if 'sku' not in f and (f.get('image') or {}).get('url')]
    return prods


def sku_pai(produto, usados):
    """
    SKU do pai derivada das SKUs dos filhos, para ficar reconhecivel no relatorio
    da Amazon. Max 40 caracteres.

    O catalogo usa a mesma base para produtos diferentes, variando so o sufixo:
    CTTS-LEGCINALTC-P, CTTS-LEGCINALTC-P-01 e CTTS-LEGCINALTC-P-04 sao tres
    leggings distintas. Entao o prefixo comum sozinho nao separa: quando todos os
    filhos terminam no mesmo token, ele entra junto (CTTS-LEGCINALTC-04); quando
    nao ha token comum, entra o sufixo -PAI.
    """
    skus = [v['sku'] for v in produto['_variantes'] if v.get('sku')]
    if not skus:
        return ('CTTS-' + re.sub(r'[^A-Z0-9]+', '', (produto.get('handle') or '').upper()))[:40]

    pre = skus[0]
    for s in skus[1:]:
        while not s.startswith(pre):
            pre = pre[:-1]
    cand = pre.rstrip('-')
    # o prefixo comum costuma parar no meio de um token ("...-BRONZE-L"):
    # volta ao ultimo separador, desde que ainda sobre nome suficiente
    if '-' in cand and not all(k[len(cand):len(cand) + 1] in ('-', '') for k in skus):
        corte = cand.rsplit('-', 1)[0]
        if len(corte) >= 6:
            cand = corte

    if len(cand) < 6 or cand in usados:
        # token final comum a todos os filhos (o "-01"/"-04" que separa as linhas)
        finais = {k.rsplit('-', 1)[-1] for k in skus if '-' in k}
        base = cand if len(cand) >= 6 else 'CTTS-' + re.sub(
            r'[^A-Z0-9]+', '', (produto.get('handle') or '').upper())[:24]
        marca = finais.pop() if len(finais) == 1 else 'PAI'
        cand = f'{base}-{marca}'

    cand = cand[:40]
    base, n = cand, 2
    while cand in usados:
        sufixo = f'-{n}'
        cand = base[:40 - len(sufixo)] + sufixo
        n += 1
    usados.add(cand)
    return cand


def tamanho_da(variante):
    """
    Junta os eixos de tamanho num valor so. Os conjuntos tem dois (top e
    legging/short) e a Amazon so aceita um eixo de tamanho por tema de variacao.
    """
    partes = []
    for o in variante.get('selectedOptions', []):
        if o['name'].lower().startswith('tamanho'):
            rotulo = o['name'].replace('Tamanho do ', '').replace('Tamanho da ', '').replace('Tamanho', '').strip()
            partes.append(f'{rotulo.capitalize()} {o["value"]}' if rotulo else o['value'])
    return ' / '.join(partes)


def cor_da(variante):
    for o in variante.get('selectedOptions', []):
        if o['name'].lower() == 'cor':
            return o['value']
    return ''


def monta(prods, listas):
    linhas, pulados = [], []
    # Todas as SKUs reais entram antes de derivar qualquer SKU de pai: o prefixo
    # comum dos filhos de um produto pode ser exatamente a SKU de outro produto
    # de variante unica, e ai as duas linhas sairiam com a mesma SKU.
    usados = {v['sku'] for p in prods.values() for v in p['_variantes'] if v.get('sku')}
    ACAO = 'Criar ou substituir (atualização completa)'

    for p in sorted(prods.values(), key=lambda x: x['title']):
        tipo = tipo_amazon(p['title'], p.get('productType'))
        if tipo is None:
            pulados.append((p['title'], f"tipo \"{p.get('productType')}\" nao e vestuario")); continue
        if p['status'] != 'ACTIVE':
            pulados.append((p['title'], 'rascunho na Shopify')); continue
        if not p['_imagens']:
            pulados.append((p['title'], 'sem nenhuma imagem - a Amazon exige imagem principal')); continue
        vs = [v for v in p['_variantes'] if v.get('sku')]
        if not vs:
            pulados.append((p['title'], 'sem SKU')); continue

        desc = texto_limpo(p.get('descriptionHtml'))
        tecido = tecido_de(desc)
        tams = sorted({tamanho_da(v) for v in vs if tamanho_da(v)})
        cores = sorted({cor_da(v) for v in vs if cor_da(v)})
        tops = topicos(desc, tams, cores, tecido)
        marca = MARCA   # a loja pediu a mesma marca em tudo
        cuidado = 'Lavagem na máquina' if tipo == 'BRA' else 'Lavagem à máquina'
        img = p['_imagens']

        # O tema so pode citar o que de fato varia entre os filhos. Quase todo
        # produto da loja e de uma cor so com varios tamanhos: declarar
        # "COR/TAMANHO" ali faz a Amazon procurar um diferenciador de cor que
        # nao existe, e entao cobrar a cor peca por peca.
        varia_tam, varia_cor = len(tams) > 1, len(cores) > 1
        if varia_tam and varia_cor: tema = 'COR/TAMANHO' if tipo == 'APPAREL' else 'TAMANHO/COR'
        elif varia_tam:             tema = 'TAMANHO'
        elif varia_cor:             tema = 'COR'
        else:                       tema = ''
        # cor constante na familia e atributo do pai, nao eixo de variacao
        cor_fixa = cores[0] if len(cores) == 1 else ''

        # resolve cada vaga na lista do tipo e junta as que sobraram: sem isso a
        # primeira coluna pode ficar vazia com a segunda preenchida, que e o que
        # acontecia nos macacoes, cuja lista so aceita "Respiravel"
        carac = [escolhe('DO', tipo, *vaga) for vaga in caracteristicas(p['title'], tipo)]
        carac_res = ([c for c in carac if c] + [''] * 5)[:5]
        comp, larg, alt = CAIXA
        peso_g = PESO[tipo]
        comum = {
            'B': tipo, 'C': ACAO, 'G': p['title'][:200], 'I': marca,
            'AM': desc[:2000],
            'AN': tops[0], 'AO': tops[1], 'AP': tops[2], 'AQ': tops[3], 'AR': tops[4],
            'AS': palavras_chave(p['title'], p.get('productType'), cores),
            'AV': 'feminino', 'AW': 'Feminino', 'AX': 'Adulto',
            'AU': escolhe('AU', tipo, ESTILO[tipo]),
            'BE': escolhe('BE', tipo, MATERIAL[0], 'Nylon'),
            'BF': escolhe('BF', tipo, MATERIAL[1]),
            'BV': escolhe('BV', tipo, estampa_de(p['title'], cores, p.get('productType')),
                          'Geométrico', 'Liso'),
            'BI': 2 if tipo == 'APPAREL' and p['title'].lower().startswith('conjunto') else 1,
            'BZ': escolhe('BZ', tipo, OPACIDADE),
            'CU': escolhe('CU', tipo, 'Todas as estações', 'Verão'),
            'CF': escolhe('CF', tipo, 'Multi-esporte', 'Ioga', 'Caminhada'),
            'ML': escolhe('ML', tipo, 'knitted'),      # tudo na loja e malha
            'DO': carac_res[0], 'DP': carac_res[1], 'DQ': carac_res[2],
            'DR': carac_res[3], 'DS': carac_res[4],
            'ER': escolhe('ER', tipo, manga_de(p['title'])),
            'CK': escolhe('CK', tipo, costas_de(p['title'])),
            'CP': escolhe('CP', tipo, gola_de(p['title'])),
            'FY': bolsos_de(p['title']),
            'R': p['title'][:100],
            'BH': tecido, 'BU': escolhe('BU', tipo, cuidado),
            'DI': 'Brasil',
            'GX': 'Novo', 'JG': 'Brasil',
            'JH': 'Não', 'JI': 'Não',
            'LV': 'Não aplicável',
            # duas familias de foto no modelo: AC..AL sao as do PRODUTO, que
            # aparecem na pagina, e HE..HJ sao as da OFERTA. Antes so as da
            # oferta iam preenchidas, e a pagina ficava sem foto.
            'AC': img[0],
            'HE': img[0],
            # embalagem: estimada por tipo, ver EMBALAGEM
            'IW': comp, 'IX': 'Centímetros',
            'IY': larg, 'IZ': 'Centímetros',
            'JA': alt,  'JB': 'Centímetros',
            'JC': peso_g, 'JD': 'Gramas',
        }
        if tipo in ('PANTS', 'SHORTS'):
            comum['FM'] = altura_cintura(p['title'])   # altura da cintura
            comum['FD'] = 'BR'                          # sistema de tamanho
            comum['FE'] = 'Alfa'                        # classe (P/M/G, nao numerico)
            comum['DU'] = escolhe('DU', tipo,
                                  'Comprimento longo' if tipo == 'PANTS' else 'Comprimento curto',
                                  'Comprimento padrão')
        else:
            comum['DU'] = escolhe('DU', tipo, 'Comprimento padrão', 'Comprimento do tornozelo')
        if tipo == 'PANTS':
            comum['GD'] = escolhe('GD', tipo, 'Skinny')    # estilo da perna
            comum['GN'] = escolhe('GN', tipo, 'Legging', 'Compressão')
        if tipo == 'SHORTS':
            comum['GA'] = escolhe('GA', tipo, 'Shorts de Compressão')
        if tipo == 'BRA':
            comum['BO'] = escolhe('BO', tipo, APOIO_TOP)   # nivel de apoio
            comum['BS'] = escolhe('BS', tipo, 'Esportivo') # estilo do item
        # o produto aceita 8 fotos extras; a oferta, 5
        for i, u in enumerate(img[1:9]):
            comum['AD AE AF AG AH AI AJ AK'.split()[i]] = u
        for i, u in enumerate(img[1:6]):
            comum['HF HG HH HI HJ'.split()[i]] = u

        so_um = len(vs) == 1
        pai = None
        if not so_um:
            pai = sku_pai(p, usados)
            linha_pai = {**comum, 'A': pai, 'D': 'Produto Pai', 'F': tema,
                         'J': 'Isento de GTIN'}
            if cor_fixa:
                linha_pai['BN'] = cor_fixa
                linha_pai['BM'] = cor_amazon(cor_fixa)
            linhas.append(linha_pai)

        for v in vs:
            ean = (v.get('barcode') or '').strip()
            tam, cor = tamanho_da(v), cor_da(v)
            if tam.lower() in ('default title', ''): tam = ''
            preco = float(v['price']) if v.get('price') else None
            de = float(v['compareAtPrice']) if v.get('compareAtPrice') else None
            qtd = max(0, int(v.get('inventoryQuantity') or 0))
            imgv = ((v.get('image') or {}).get('url')) or img[0]

            linha = {**comum, 'A': v['sku'][:40], 'AC': imgv, 'HE': imgv,
                     # amostra de cor: so faz sentido quando a variante tem
                     # foto propria e a cor e o que varia na familia
                     'AL': imgv if (v.get('image') or {}).get('url') and len(cores) > 1 else None,
                     'J': 'EAN' if len(ean) == 13 and ean.isdigit() else 'Isento de GTIN',
                     'K': ean if len(ean) == 13 and ean.isdigit() else '',
                     'BM': cor_amazon(cor) if cor else '', 'BN': cor,
                     'DT': tam,
                     'FF': tamanho_de_baixo(tam) if tipo in ('PANTS','SHORTS') else None,
                     'GZ': de if de and de > (preco or 0) else None,
                     'IH': 'Logística do vendedor (Padrão)',
                     'II': qtd, 'IL': 'Desativado',
                     'IM': preco}
            if not so_um:
                linha.update({'D': 'Produto Filho', 'E': pai, 'F': tema})
            linhas.append(linha)

    return linhas, pulados


# ---------------------------------------------------------------- gravacao

def grava(origem, destino, linhas):
    zin = zipfile.ZipFile(origem)
    xml = zin.read(MODELO).decode('utf-8')

    # as linhas 1..6 (cabecalhos e exemplo) ficam como estao
    fim6 = xml.find('</row>', xml.find('<row', xml.find('r="6"') - 200)) + len('</row>')
    assert fim6 > len('</row>'), 'nao achei o fim da linha 6'

    partes = []
    for i, dados in enumerate(linhas):
        n = PRIMEIRA + i
        cels = ''.join(celula(c, n, dados[c])
                       for c in sorted(dados, key=num_col) if dados[c] not in (None, ''))
        partes.append(f'<row r="{n}" spans="1:356">{cels}</row>')
    ultima = PRIMEIRA + len(linhas) - 1

    novo = xml[:fim6] + ''.join(partes) + xml[fim6:]
    novo = re.sub(r'<dimension ref="A1:([A-Z]+)6"\s*/>',
                  lambda m: f'<dimension ref="A1:{m.group(1)}{ultima}"/>', novo, count=1)

    zout = zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        zout.writestr(it, novo.encode('utf-8') if it.filename == MODELO else zin.read(it.filename))
    zout.close(); zin.close()
    return ultima


def main(jsonl, xlsm_origem, xlsm_destino):
    global LISTAS
    import os
    caminho = os.path.join(os.path.dirname(os.path.abspath(jsonl)), 'listas-completas.json')
    if os.path.exists(caminho):
        LISTAS = json.load(open(caminho, encoding='utf-8'))
        print(f'listas fechadas carregadas: {len(LISTAS)} tipos')
    else:
        print('AVISO: listas-completas.json nao encontrado - colunas de lista ficarao vazias')
    prods = carrega_shopify(jsonl)
    listas = json.load(open('listas.json', encoding='utf-8')) if __import__('os').path.exists('listas.json') else {}
    linhas, pulados = monta(prods, listas)

    # nenhuma linha pode ir com valor fora das listas fechadas nem passar do limite
    FECHADAS = {'D': {'Produto Pai', 'Produto Filho'},
                'B': {'BRA', 'ONE_PIECE_OUTFIT', 'APPAREL', 'SHORTS', 'PANTS'},
                'AW': {'Feminino', 'Masculino', 'Unissex'},
                'AX': {'Adulto', 'Bebê', 'Criança'},
                'J': {'EAN', 'GTIN', 'UPC', 'ASIN', 'Isento de GTIN'},
                'BM': {'Amarelo','Azul','Bege','Branco','Branco Off','Bronze','Cinza','Claro',
                       'Laranja','Marrom','Metálico','Multicolorido','Ouro','Prata','Preto',
                       'Rosa','Roxo','Turquesa','Verde','Vermelho'}}
    for l in linhas:
        for col, ok in FECHADAS.items():
            if l.get(col) and l[col] not in ok:
                raise SystemExit(f'valor fora da lista em {col}: {l[col]!r} (SKU {l.get("A")})')
        assert len(l.get('G', '')) <= 200, f'titulo longo demais: {l.get("A")}'
        assert l.get('A'), 'linha sem SKU'
    skus = [l['A'] for l in linhas]
    assert len(skus) == len(set(skus)), 'SKU repetido entre as linhas'

    ultima = grava(xlsm_origem, xlsm_destino, linhas)

    pais = sum(1 for l in linhas if l.get('D') == 'Produto Pai')
    filhos = sum(1 for l in linhas if l.get('D') == 'Produto Filho')
    sozinhos = len(linhas) - pais - filhos
    print(f'{len(linhas)} linhas (7 a {ultima}): {pais} pais, {filhos} filhos, {sozinhos} sem variacao')
    print('tipos:', dict(collections.Counter(l['B'] for l in linhas)))
    print('ID do produto:', dict(collections.Counter(l['J'] for l in linhas if l.get('J'))))
    print(f'\n{len(pulados)} produtos de fora:')
    for t, m in pulados: print(f'  - {t}: {m}')


if __name__ == '__main__':
    main(*sys.argv[1:4])

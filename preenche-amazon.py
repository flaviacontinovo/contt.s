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

# Embalagem por tipo: comprimento x largura x altura em cm, e peso em gramas.
# ESTIMATIVA de saco/envelope de moda fitness - o peso da Shopify nao serve
# (520 variantes em 0,0 kg). Mexe no frete: conferir antes de confiar.
EMBALAGEM = {
    'PANTS':            (30, 22, 4, 220),
    'SHORTS':           (26, 20, 3, 130),
    'BRA':              (24, 18, 3, 110),
    'ONE_PIECE_OUTFIT': (30, 22, 5, 260),
    'APPAREL':          (32, 24, 6, 330),
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
    return 'Malha'   # verdadeiro para todas as pecas da loja; ver AMAZON.md


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

        comp, larg, alt, peso_g = EMBALAGEM[tipo]
        comum = {
            'B': tipo, 'C': ACAO, 'G': p['title'][:200], 'I': marca,
            'AM': desc[:2000],
            'AN': tops[0], 'AO': tops[1], 'AP': tops[2], 'AQ': tops[3], 'AR': tops[4],
            'AS': palavras_chave(p['title'], p.get('productType'), cores),
            'AV': 'feminino', 'AW': 'Feminino', 'AX': 'Adulto',
            'AU': ESTILO[tipo],
            'R': p['title'][:100],
            'BH': tecido, 'BU': cuidado,
            'DI': 'Brasil',
            'GX': 'Novo', 'JG': 'Brasil',
            'JH': 'Não', 'JI': 'Não',
            'LV': 'Não aplicável',
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
        for i, u in enumerate(img[1:6], start=1):
            comum['HF HG HH HI HJ'.split()[i - 1]] = u

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

            linha = {**comum, 'A': v['sku'][:40], 'HE': imgv,
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

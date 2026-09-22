# -*- coding: utf-8 -*-
"""Preenche o modelo de envio em massa da Shopee com os dados da Shopify.

Mesmo escopo do TikTok: a colecao "Conjuntos R$ 189" e os dois Macaquinhos
Essential.

O fiscal NAO e inventado. Origem, CSOSN e unidade de medida sao os mesmos que
a loja ja usa nas notas do Mercado Livre (arquivo Dados_Fiscais, linhas que ela
mesma preencheu). O NCM vem da Shopify. Tipo de operacao e CFOP vem da resposta
da loja em 23/09/2026: fabricante.

Cirurgia direta no XML: o openpyxl nem abre este arquivo (a Shopee gera um
sheetView com atributo invalido), e a aba Modelo tem 5005 validacoes de dados.
"""
import json, io, re, html, zipfile, os

MODELO  = 'shopee.xlsx'
DESTINO = 'shopee-contts.xlsx'
ABA     = 'xl/worksheets/sheet2.xml'
LINHA_1 = 7

MARCA = 'CONTT.s'
CAIXA = (20, 20, 5)                       # comprimento, largura, altura em cm
# peso bruto em kg. 0,35 do conjunto e o numero da propria loja, das notas do
# Mercado Livre. O do macaquinho e ESTIMADO entre a legging dela (0,25) e o
# conjunto (0,35), por ser peca unica que cobre tronco e short.
PESO = {'conjunto': 0.35, 'macaquinho': 0.30}

# listas fechadas da aba HiddenTax, copiadas ao pe da letra
ORIGEM   = '0 - Nacional, exceto as indicadas nos códigos 3, 4, 5 e 8'
CSOSN    = '102 - Tributada pelo Simples Nacional sem permissão de crédito'
UNIDADE  = 'UN (UNIDADE)'
CST_PIS  = '49 - Outras Operações de Saída'   # Simples Nacional
OPERACAO = '2 - Fabricante'                   # resposta da loja
CFOP_MESMO, CFOP_OUTRO = '5101', '6101'       # venda de producao do estabelecimento

CONJUNTOS = [
    'conjunto-legging-sculpt-top-vulcano',
    'conjunto-legging-sculpt-top-preto',
    'conjunto-shine-legging-top-off-white',
    'conjunto-shine-legging-top-matcha',
    'conjunto-legging-sculpt-top-rosa-retro',
    'conjunto-aura-legging-top-bronze',
    'conjunto-legging-aura-top-laranjaa',
    'conjunto-legging-aura-top-petroleo',
]
MACAQUINHOS = ['macaquinho-essential-vulcano', 'macaquinho-essential-preto']

# ---------------------------------------------------------------- Shopify
prods, midia, vars_, ncms = {}, {}, {}, {}
for linha in io.open('produtos.jsonl', encoding='utf-8'):
    o = json.loads(linha)
    pai = o.get('__parentId')
    if 'handle' in o:   prods[o['id']] = o
    elif 'sku' in o:    vars_.setdefault(pai, []).append(o)
    elif 'image' in o:  midia.setdefault(pai, []).append(o)
for linha in io.open('ncm.jsonl', encoding='utf-8'):
    o = json.loads(linha)
    if o.get('sku'):
        ncms[o['sku']] = ((o.get('inventoryItem') or {}).get('harmonizedSystemCode') or '')

por_handle = {p['handle']: p for p in prods.values()}
faltam = [h for h in CONJUNTOS + MACAQUINHOS if h not in por_handle]
assert not faltam, 'handle nao encontrado: %s' % faltam

def texto(h):
    t = h or ''
    t = re.sub(r'(?i)<br\s*/?>', '\n', t)
    t = re.sub(r'(?i)<li[^>]*>', '\n• ', t)
    t = re.sub(r'(?i)</(p|li|ul|div|h\d)>', '\n', t)
    t = re.sub(r'<[^>]+>', '', t)
    t = html.unescape(t)
    t = re.sub(r'[ \t]+', ' ', t)
    return re.sub(r'\n\s*\n\s*\n+', '\n\n', t).strip()

def opcao(v, nome):
    for o in v.get('selectedOptions') or []:
        if o['name'] == nome: return o['value']
    return ''

def sku_pai(skus):
    """maior prefixo comum, cortado no ultimo separador."""
    p = os.path.commonprefix(skus)
    p = p[:max(p.rfind('-'), p.rfind('_'))] if ('-' in p or '_' in p) else p
    return (p or skus[0])[:100]

linhas = []
for handle in CONJUNTOS + MACAQUINHOS:
    p = por_handle[handle]
    ehconj = handle in CONJUNTOS
    imgs = [m['image']['url'] for m in midia.get(p['id'], []) if m.get('image')][:9]
    vs = vars_.get(p['id'], [])
    assert imgs and vs, 'produto sem foto ou sem variante: %s' % handle
    pai = sku_pai([v['sku'] for v in vs])
    nome = ('%s %s' % (MARCA, p['title']))[:120]
    desc = texto(p.get('descriptionHtml'))
    assert 10 <= len(desc) <= 5000, 'descricao fora de 10..5000: %s' % handle
    ncm = next((ncms.get(v['sku'], '') for v in vs if ncms.get(v['sku'])), '')
    assert ncm, 'sem NCM: %s' % handle

    base = {'B': nome, 'C': desc, 'D': pai, 'E': pai,
            'R': imgs[0],
            'AA': PESO['conjunto' if ehconj else 'macaquinho'],
            'AB': CAIXA[0], 'AC': CAIXA[1], 'AD': CAIXA[2],
            'AE': 'Ligado',
            'AH': ncm, 'AI': CFOP_MESMO, 'AJ': CFOP_OUTRO,
            'AK': ORIGEM, 'AL': CSOSN, 'AN': UNIDADE,
            'AO': CST_PIS, 'AQ': OPERACAO, 'AV': 'No'}
    for i, u in enumerate(imgs[1:9]):
        base['S' if i == 0 else chr(ord('S') + i)] = u      # S..Z

    for v in vs:
        if ehconj:
            n1, v1 = 'Tamanho da legging', opcao(v, 'Tamanho da legging')
            n2, v2 = 'Tamanho do top',     opcao(v, 'Tamanho do top')
        else:
            n1, v1 = 'Tamanho', opcao(v, 'Tamanho')
            n2, v2 = 'Cor',     opcao(v, 'Cor')
        assert v1 and v2, 'variante sem variacao: %s' % v.get('sku')
        assert all(len(x) <= 30 for x in (n1, v1, n2, v2)), 'variacao acima de 30 caracteres'
        linhas.append(dict(base, F=n1, G=v1, I=n2, J=v2,
                           K=float(v['price']),
                           L=max(0, int(v.get('inventoryQuantity') or 0)),
                           M=(v.get('sku') or '')[:100]))

print('produtos: %d | linhas: %d' % (len(CONJUNTOS) + len(MACAQUINHOS), len(linhas)))
assert len({l['M'] for l in linhas}) == len(linhas), 'SKU repetida'

# ---------------------------------------------------------------- XML
def esc(v):
    return str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def col_idx(letras):
    n = 0
    for ch in letras: n = n * 26 + (ord(ch) - 64)
    return n

def celula(col, n, v):
    if isinstance(v, (int, float)):
        return '<c r="%s%d"><v>%s</v></c>' % (col, n, v)
    return ('<c r="%s%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (col, n, esc(v)))

with zipfile.ZipFile(MODELO) as z:
    xml = z.read(ABA).decode('utf-8')
corpo = re.search(r'<sheetData>(.*)</sheetData>', xml, re.S)

novas = []
for i, l in enumerate(linhas):
    n = LINHA_1 + i
    cs = ''.join(celula(c, n, l[c]) for c in sorted(l, key=col_idx) if l[c] != '')
    novas.append('<row r="%d" spans="1:52">%s</row>' % (n, cs))

xml = xml[:corpo.end(1)] + ''.join(novas) + xml[corpo.end(1):]
xml = re.sub(r'<dimension ref="[^"]*"',
             '<dimension ref="A1:AZ%d"' % (LINHA_1 - 1 + len(linhas)), xml, count=1)

tmp = DESTINO + '.tmp'
with zipfile.ZipFile(MODELO) as ent, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as sai:
    for item in ent.infolist():
        dados = ent.read(item.filename)
        if item.filename == ABA: dados = xml.encode('utf-8')
        sai.writestr(item, dados)
os.replace(tmp, DESTINO)
print('gravado:', DESTINO)

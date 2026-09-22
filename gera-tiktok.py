# -*- coding: utf-8 -*-
"""Preenche o modelo de envio em lote do TikTok Shop com os dados da Shopify.

Escopo pedido pela loja: os produtos da colecao "Conjuntos R$ 189" e os dois
Macaquinhos Essential.

Cirurgia direta no XML em vez de openpyxl: a aba Modelo tem 28 validacoes de
dados e 17 formatacoes condicionais. As linhas 1 a 6 (cabecalhos e a linha de
exemplo do TikTok) saem byte a byte iguais - as validacoes comecam em A7, que
e onde os dados entram.
"""
import json, io, re, html, zipfile, os

MODELO  = 'tiktok.xlsx'
DESTINO = 'tiktok-contts.xlsx'
ABA     = 'xl/worksheets/sheet1.xml'
LINHA_1 = 7                      # primeira linha de dados

MARCA = 'CONTT.s FITNESS WEAR (7652714581887026962)'
CAT_CONJUNTO   = 'Ternos e macacões femininos/Conjuntos'
CAT_MACAQUINHO = 'Ternos e macacões femininos/Macacões'
CAIXA = (20, 20, 5)              # comprimento, largura, altura em cm
PESO  = {'conjunto': 330, 'macaquinho': 260}   # gramas, ESTIMADO

# handles da colecao "Conjuntos R$ 189", conferidos na Shopify em 22/09/2026
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
prods, midia, vars_ = {}, {}, {}
for linha in io.open('produtos.jsonl', encoding='utf-8'):
    o = json.loads(linha)
    pai = o.get('__parentId')
    if 'handle' in o:
        prods[o['id']] = o
    elif 'sku' in o or 'selectedOptions' in o:
        vars_.setdefault(pai, []).append(o)
    elif 'image' in o:
        midia.setdefault(pai, []).append(o)

por_handle = {p['handle']: p for p in prods.values()}
faltam = [h for h in CONJUNTOS + MACAQUINHOS if h not in por_handle]
assert not faltam, 'handle nao encontrado na Shopify: %s' % faltam

def texto(h):
    """HTML da descricao -> texto limpo, que e o que o TikTok espera."""
    t = h or ''
    t = re.sub(r'(?i)<br\s*/?>', '\n', t)
    t = re.sub(r'(?i)<li[^>]*>', '\n• ', t)
    t = re.sub(r'(?i)</(p|li|ul|div|h\d)>', '\n', t)
    t = re.sub(r'<[^>]+>', '', t)
    t = html.unescape(t)
    t = re.sub(r'[ \t]+', ' ', t)
    t = re.sub(r'\n\s*\n\s*\n+', '\n\n', t)
    return t.strip()

def opcao(v, nome):
    for o in v.get('selectedOptions') or []:
        if o['name'] == nome: return o['value']
    return ''

linhas = []
for handle in CONJUNTOS + MACAQUINHOS:
    p = por_handle[handle]
    ehconj = handle in CONJUNTOS
    imgs = [m['image']['url'] for m in midia.get(p['id'], []) if m.get('image')][:9]
    assert imgs, 'produto sem foto: %s' % handle
    base = {
        'A': CAT_CONJUNTO if ehconj else CAT_MACAQUINHO,
        'B': MARCA,
        'C': p['title'],
        'D': texto(p.get('descriptionHtml')),
        'U': PESO['conjunto' if ehconj else 'macaquinho'],
        'V': CAIXA[0], 'W': CAIXA[1], 'X': CAIXA[2],
    }
    for i, u in enumerate(imgs):
        base[chr(ord('E') + i)] = u          # E..M

    vs = vars_.get(p['id'], [])
    assert vs, 'produto sem variante: %s' % handle
    for v in vs:
        if ehconj:
            # o TikTok aceita 2 dimensoes; a cor e unica no conjunto e ja esta
            # no nome, entao as duas dimensoes sao os dois tamanhos
            n1, v1 = 'Tamanho da legging', opcao(v, 'Tamanho da legging')
            n2, v2 = 'Tamanho do top',     opcao(v, 'Tamanho do top')
        else:
            n1, v1 = 'Tamanho', opcao(v, 'Tamanho')
            n2, v2 = 'Cor',     opcao(v, 'Cor')
        assert v1 and v2, 'variante sem valor de variacao: %s' % v.get('sku')
        assert len(n1) <= 20 and len(n2) <= 20, 'nome de variacao acima de 20 caracteres'
        linhas.append(dict(base, P=n1, Q=v1, S=n2, T=v2,
                           Z=float(v['price']),
                           AA=max(0, int(v.get('inventoryQuantity') or 0)),
                           AB=v.get('sku') or ''))

print('produtos: %d | linhas (variantes): %d' % (len(CONJUNTOS) + len(MACAQUINHOS), len(linhas)))
assert len({l['AB'] for l in linhas}) == len(linhas), 'SKU repetida'
assert all(len(l['C']) < 300 for l in linhas), 'nome de produto com 300+ caracteres'

# ---------------------------------------------------------------- XML
def esc(v):
    return str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def celula(col, n, v):
    if isinstance(v, (int, float)):
        return '<c r="%s%d"><v>%s</v></c>' % (col, n, v)
    return ('<c r="%s%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (col, n, esc(v)))

ORDEM = (['A', 'B', 'C', 'D'] + [chr(ord('E') + i) for i in range(9)]
         + ['P', 'Q', 'S', 'T', 'U', 'V', 'W', 'X', 'Z', 'AA', 'AB'])

with zipfile.ZipFile(MODELO) as z:
    xml = z.read(ABA).decode('utf-8')

corpo = re.search(r'<sheetData>(.*)</sheetData>', xml, re.S)
antigas = re.findall(r'<row r="(\d+)"[^>]*(?:/>|>.*?</row>)', corpo.group(1), re.S)
todas = re.findall(r'<row r="\d+"[^>]*(?:/>|>.*?</row>)', corpo.group(1), re.S)
mapa = dict(zip((int(n) for n in antigas), todas))

for i, l in enumerate(linhas):
    n = LINHA_1 + i
    cs = ''.join(celula(c, n, l[c]) for c in ORDEM if c in l and l[c] != '')
    # mantem as celulas vazias com estilo que o modelo ja trazia em O e AF
    cs = (cs.replace('<c r="P%d"' % n, '<c r="O%d" s="42"></c><c r="P%d"' % (n, n))
          + '<c r="AF%d" s="41"></c>' % n)
    mapa[n] = '<row r="%d" spans="1:32">%s</row>' % (n, cs)

novo = ''.join(mapa[k] for k in sorted(mapa))
xml = xml[:corpo.start(1)] + novo + xml[corpo.end(1):]
xml = xml.replace('<dimension ref="A3:AB6">',
                  '<dimension ref="A1:AF%d">' % (LINHA_1 - 1 + len(linhas)))

tmp = DESTINO + '.tmp'
with zipfile.ZipFile(MODELO) as ent, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as sai:
    for item in ent.infolist():
        dados = ent.read(item.filename)
        if item.filename == ABA:
            dados = xml.encode('utf-8')
        sai.writestr(item, dados)
os.replace(tmp, DESTINO)
print('gravado:', DESTINO)

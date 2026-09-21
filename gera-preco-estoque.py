# -*- coding: utf-8 -*-
"""Preenche o modelo de Preco e Quantidade da Amazon com o preco e o estoque
atuais da Shopify.

Cirurgia direta no XML em vez de openpyxl: a aba Modelo tem 32 validacoes de
dados e 16 formatacoes condicionais que o openpyxl perderia. As linhas 1 a 6
(cabecalhos e a linha de exemplo da Amazon) saem byte a byte iguais.
"""
import json, re, io, shutil, zipfile, os, sys

ORIGEM  = 'pq.xlsm'
DESTINO = 'amazon-preco-estoque.xlsm'
CANAL   = 'Logística do vendedor (Padrão)'   # envio por conta da loja
LEAD    = 1                                   # 1 dia util de processamento
SEMPRE  = 'Desativado'                        # estoque real, nao infinito

ofertas = json.load(open('ofertas.json'))          # as 1109 ofertas que foram pra Amazon
shopify = {}
for linha in io.open('variantes.jsonl', encoding='utf-8'):
    v = json.loads(linha)
    if v.get('sku'):
        shopify[v['sku']] = (float(v['price']), int(v['inventoryQuantity'] or 0))

linhas, sem_par, mud_preco, mud_qtd = [], [], 0, 0
for o in ofertas:
    sku = o['sku']
    if sku not in shopify:
        sem_par.append(sku)
        continue
    preco, qtd = shopify[sku]
    if abs(preco - float(o['preco'])) > 0.005: mud_preco += 1
    if qtd != int(o['qtd']):                   mud_qtd += 1
    linhas.append({'A': sku, 'B': CANAL, 'C': qtd, 'D': LEAD, 'F': SEMPRE, 'G': preco})

assert not sem_par, 'SKU da planilha sem par na Shopify: %s' % sem_par[:5]
assert len({l['A'] for l in linhas}) == len(linhas), 'SKU repetida'
print('linhas: %d | precos mudados desde 17/09: %d | estoques mudados: %d'
      % (len(linhas), mud_preco, mud_qtd))
print('estoque zerado: %d | total de pecas: %d'
      % (sum(1 for l in linhas if l['C'] == 0), sum(l['C'] for l in linhas)))

def esc(v):
    return (str(v).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;'))

def celula(col, i, v):
    if isinstance(v, (int, float)):
        return '<c r="%s%d"><v>%s</v></c>' % (col, i, repr(v) if isinstance(v,float) else v)
    return ('<c r="%s%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (col, i, esc(v)))

xml = io.open(os.path.join('pq','xl','worksheets','sheet5.xml'), encoding='utf-8').read()
novas = []
for n, l in enumerate(linhas, start=7):
    cs = ''.join(celula(c, n, l[c]) for c in ['A','B','C','D','F','G'] if c in l)
    novas.append('<row r="%d" spans="1:16">%s</row>' % (n, cs))
novas = ''.join(novas)

assert xml.count('</sheetData>') == 1
xml = xml.replace('</sheetData>', novas + '</sheetData>')
xml = xml.replace('<dimension ref="A1:P6" />',
                  '<dimension ref="A1:P%d" />' % (6 + len(linhas)))

shutil.copy(ORIGEM, DESTINO)
# regrava o zip trocando so a sheet5
tmp = DESTINO + '.tmp'
with zipfile.ZipFile(ORIGEM) as ent, zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as sai:
    for item in ent.infolist():
        dados = ent.read(item.filename)
        if item.filename == 'xl/worksheets/sheet5.xml':
            dados = xml.encode('utf-8')
        sai.writestr(item, dados)
os.replace(tmp, DESTINO)
print('gravado:', DESTINO)

#!/usr/bin/env python3
"""
Gera o arquivo de atualizacao em massa do NCM para o catalogo da CONTT.s.

Entrada : catalogo.jsonl  (export do bulkOperationRunQuery da Shopify)
Saida   : ncm.jsonl       (uma linha por variante, para o bulkOperationRunMutation)

O NCM vai no campo nativo `inventoryItem.harmonizedSystemCode`. A loja nao tem
metafield fiscal nenhum (conferido em 15/09/2026: so custom.cor, reviews.rating,
shopify.size e shopify.color-pattern), e esse campo nativo e de onde os
integradores de NF-e leem o NCM. Shopify aceita de 6 a 13 digitos, entao o NCM
de 8 cabe inteiro, sem pontos.

TODA a tabela abaixo pressupoe MALHA DE FIBRAS SINTETICAS (poliamida/poliester
com elastano), que e o tecido da moda fitness. Se alguma peca for
predominantemente de algodao, o NCM muda de posicao - ver NCM.md.
"""
import json, sys, collections

# NCM -> texto da posicao, para o relatorio e para a conferencia do contador
POSICAO = {
    '61046300': 'Calcas, jardineiras, bermudas e shorts, de fibras sinteticas, de malha, para senhoras/meninas',
    '61045300': 'Saias e saias-calcas, de fibras sinteticas, de malha, para senhoras/meninas',
    '61042300': 'Conjuntos, de fibras sinteticas, de malha, para senhoras/meninas',
    '62121000': 'Sutias, mesmo de malha',
    '61143000': 'Outras roupas, de malha, de fibras sinteticas',
    '61103000': 'Puloveres, cardigas, coletes e artigos semelhantes, de fibras sinteticas, de malha',
}


def classifica(titulo, tipo):
    """Devolve (ncm, motivo). ncm None = fora do escopo, nao escreve nada."""
    t = (tipo or '').upper()
    tit = (titulo or '').lower()

    # conjunto vem primeiro: o tipo e "CONJUNTOS" nos antigos e "Conjunto" nos novos
    if t.startswith('CONJUNTO'):
        return '61042300', 'conjunto (peca de cima + peca de baixo, mesmo tecido, vendidos juntos)'

    # short-saia antes de short, senao cai na regra de short
    if t.startswith('SHORT-SAIA') or t.startswith('SHORT SAIA'):
        return '61045300', 'saia-calca'

    if t.startswith('LEGGING') or t.startswith('SHORT'):
        return '61046300', 'calca/short'

    if t.startswith('TOP E CROPPED') or (not t and tit.startswith('top')):
        # "Top" = sutia esportivo (sustentacao). "Blusa"/"Cropped" = peca de vestir.
        if tit.startswith('blusa') or tit.startswith('cropped'):
            return '61143000', 'blusa/cropped - peca de vestir, nao sutia'
        return '62121000', 'top esportivo com funcao de sustentacao = sutia'

    if t.startswith('MACAC'):
        return '61143000', 'macacao/macaquinho - nao ha posicao propria, cai em "outras roupas"'

    if t.startswith('JAQUETA'):
        return '61103000', 'jaqueta de malha'

    return None, f'fora do escopo (tipo "{tipo}") - nao e vestuario ou tipo desconhecido'


def main(entrada, saida):
    prods, variantes = {}, []
    for ln in open(entrada, encoding='utf-8'):
        o = json.loads(ln)
        (variantes if '__parentId' in o else prods).setdefault(
            o['id'], o) if '__parentId' not in o else variantes.append(o)

    escritas, fora = [], []
    por_ncm = collections.Counter()
    for v in variantes:
        p = prods.get(v['__parentId'])
        if p is None:
            print(f'AVISO: variante {v["id"]} sem produto pai', file=sys.stderr)
            continue
        ncm, motivo = classifica(p['title'], p['productType'])
        if ncm is None:
            fora.append((p['title'], motivo))
            continue
        item = (v.get('inventoryItem') or {}).get('id')
        if not item:
            print(f'AVISO: variante {v["id"]} sem item de estoque', file=sys.stderr)
            continue
        escritas.append({'id': item, 'input': {'harmonizedSystemCode': ncm}})
        por_ncm[ncm] += 1

    # nenhum NCM pode sair com tamanho errado: a Shopify recusa fora de 6-13 digitos
    for l in escritas:
        c = l['input']['harmonizedSystemCode']
        assert c.isdigit() and len(c) == 8, f'NCM invalido: {c!r}'

    with open(saida, 'w', encoding='utf-8') as f:
        for l in escritas:
            f.write(json.dumps(l, ensure_ascii=False) + '\n')

    print(f'{len(escritas)} variantes recebem NCM, de {len(variantes)} no total\n')
    print(f'{"NCM":<12}{"variantes":>10}  posicao')
    for ncm, n in sorted(por_ncm.items(), key=lambda x: -x[1]):
        bonito = f'{ncm[:4]}.{ncm[4:6]}.{ncm[6:]}'
        print(f'{bonito:<12}{n:>10}  {POSICAO[ncm]}')
    if fora:
        print('\nfora do escopo (nada foi escrito):')
        for t, m in sorted(set(fora)):
            print(f'  {t} - {m}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])

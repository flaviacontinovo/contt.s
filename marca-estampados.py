#!/usr/bin/env python3
"""
Decide quais produtos da loja sao estampados e gera as marcacoes de tag.

Estampa e padrao aplicado no tecido. Jacquard, brocado e canelado sao TEXTURA e
ficam de fora de proposito.

O criterio nao e chute por titulo: cada nome de estampa e confirmado por um
sinal da propria loja.

  sinal 1  a variante tem cor "Estampado" ou "Tie Dye"
  sinal 2  o produto e do tipo "TOP E CROPPED > TOP ESTAMPADO"

Um nome de estampa entra na lista quando pelo menos um dos dois sinais o
confirma; dai ele vale para todas as pecas com aquele nome, porque a mesma
estampa sai em legging, short e top. Foi assim que "poa", "fake jeans", "gibi",
"urban ink" e "electric rose" entraram: a cor deles registra o tom (vermelho,
preto), mas ela tipou os tops como TOP ESTAMPADO.
"""
import json, re, sys, collections, unicodedata

def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s.lower())
                   if unicodedata.category(c) != 'Mn')

# candidatos a nome de estampa. Nenhum entra sem confirmacao de um dos sinais.
CANDIDATOS = [
    'animal print', 'zebra', 'onca', 'newspaper', 'abstract', 'butterfly', 'space',
    'tropical', 'poa', 'tie dye', 'fake jeans', 'destroyed', 'gibi', 'hollywood',
    'colmeia', 'oasis', 'expressive', 'urban ink', 'fantasy', 'africa', 'reflorescer',
    'green love', 'nuvem', 'college', 'ray', 'grafiatto', 'map', 'electric rose', 'nakay',
]
# textura, nao estampa - nunca entram, mesmo que algum sinal aponte
TEXTURA = ['jacquard', 'brocad', 'canelad', 'ribana']


def carrega(caminho):
    prods, filhos = {}, collections.defaultdict(list)
    for ln in open(caminho, encoding='utf-8'):
        o = json.loads(ln)
        if '__parentId' in o: filhos[o['__parentId']].append(o)
        else: prods[o['id']] = o
    for p in prods.values():
        p['_v'] = [f for f in filhos.get(p['id'], []) if 'sku' in f]
        p['_cores'] = {o['value'] for v in p['_v']
                       for o in v.get('selectedOptions', []) if o['name'] == 'Cor'}
        p['_t'] = sem_acento(p['title'])
    return prods


def decide(prods):
    ativos = [p for p in prods.values() if p['status'] == 'ACTIVE']
    por_cor  = {p['id'] for p in ativos if {'Estampado', 'Tie Dye'} & p['_cores']}
    por_tipo = {p['id'] for p in ativos
                if 'ESTAMPADO' in (p.get('productType') or '').upper()}
    confirmados = por_cor | por_tipo

    # um nome so entra se algum produto que o carrega estiver confirmado
    nomes, rejeitados = [], []
    for k in CANDIDATOS:
        com = [p for p in ativos if k in p['_t']]
        if com and any(p['id'] in confirmados for p in com):
            nomes.append(k)
        elif com:
            rejeitados.append((k, len(com)))

    estampados = set(confirmados)
    for k in nomes:
        estampados |= {p['id'] for p in ativos if k in p['_t']}

    # textura nunca entra so pela textura; se entrou por outro sinal, fica
    fora_textura = [p for p in ativos
                    if p['id'] not in estampados
                    and any(t in p['_t'] for t in TEXTURA)]

    return ativos, estampados, nomes, rejeitados, fora_textura, por_cor, por_tipo


def main(jsonl, saida=None):
    prods = carrega(jsonl)
    ativos, est, nomes, rej, tex, por_cor, por_tipo = decide(prods)

    print(f'{len(ativos)} produtos ativos → {len(est)} estampados, {len(ativos)-len(est)} lisos\n')
    print(f'nomes de estampa confirmados ({len(nomes)}):')
    print('  ' + ', '.join(sorted(nomes)))
    print(f'\nnomes descartados por falta de sinal ({len(rej)}):')
    print('  ' + ', '.join(f'{k} ({n})' for k, n in sorted(rej)))
    print(f'\ntextura (jacquard/brocado/canelado) que ficou de fora: {len(tex)} produtos')
    print(f'\nde onde vem cada um:')
    print(f'  só pela cor Estampado/Tie Dye : {len(por_cor - por_tipo)}')
    print(f'  só pelo tipo TOP ESTAMPADO    : {len(por_tipo - por_cor)}')
    print(f'  pelos dois                    : {len(por_cor & por_tipo)}')
    print(f'  pelo nome da estampa          : {len(est - por_cor - por_tipo)}')

    alvos = [{'id': p['id'], 'titulo': p['title'], 'tags': p.get('tags', [])}
             for p in ativos if p['id'] in est]
    if saida:
        json.dump(alvos, open(saida, 'w'), ensure_ascii=False, indent=1)
        print(f'\n{len(alvos)} produtos gravados em {saida}')
    return alvos


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)

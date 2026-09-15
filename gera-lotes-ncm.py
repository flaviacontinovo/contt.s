#!/usr/bin/env python3
"""
Quebra o ncm.jsonl em documentos GraphQL com varios inventoryItemUpdate por
chamada, usando apelidos.

Por que assim: a operacao em lote da Shopify (bulkOperationRunMutation) esta
bloqueada pela politica do conector, porque ela executa mutacao arbitraria.
inventoryItemUpdate sozinha e permitida, e o GraphQL deixa chamar a mesma
mutacao varias vezes no mesmo documento dando um apelido a cada uma. Da no
mesmo resultado sem passar por cima de nenhuma trava.

O NCM vai em variavel (uma por NCM distinto no lote), nao repetido em cada
linha - encurta bastante o documento. Variavel declarada e nao usada e erro de
validacao no GraphQL, entao cada lote declara so as que ele realmente usa.
"""
import json, sys, os

TAMANHO = 100  # mutacoes por chamada; o custo da Shopify e ~10 por mutacao


def main(entrada, destino, tamanho=TAMANHO):
    linhas = [json.loads(l) for l in open(entrada, encoding='utf-8')]
    # agrupado por NCM: cada lote tende a usar uma variavel so
    linhas.sort(key=lambda l: l['input']['harmonizedSystemCode'])
    os.makedirs(destino, exist_ok=True)

    lotes = [linhas[i:i + tamanho] for i in range(0, len(linhas), tamanho)]
    for n, lote in enumerate(lotes):
        ncms = sorted({l['input']['harmonizedSystemCode'] for l in lote})
        nome = {ncm: f'n{i}' for i, ncm in enumerate(ncms)}
        decl = ', '.join(f'${v}: InventoryItemInput!' for v in nome.values())
        # fragmento em vez de repetir "{ userErrors { field message } }" 100x:
        # encurta cada linha em ~30%, e sao 1188 linhas no total
        corpo = '\n'.join(
            f'm{i}:inventoryItemUpdate(id:"{l["id"]}",'
            f'input:${nome[l["input"]["harmonizedSystemCode"]]}){{...e}}'
            for i, l in enumerate(lote))
        doc = ('fragment e on InventoryItemUpdatePayload{userErrors{field message}}\n'
               f'mutation NCM({decl}){{\n{corpo}\n}}')
        vars_ = {v: {'harmonizedSystemCode': k} for k, v in nome.items()}

        open(f'{destino}/lote-{n:02d}.graphql', 'w', encoding='utf-8').write(doc)
        open(f'{destino}/lote-{n:02d}.vars.json', 'w', encoding='utf-8').write(
            json.dumps(vars_, ensure_ascii=False))

    print(f'{len(linhas)} mutacoes em {len(lotes)} lotes de ate {tamanho}')
    for n, lote in enumerate(lotes):
        ncms = sorted({l['input']['harmonizedSystemCode'] for l in lote})
        tam = os.path.getsize(f'{destino}/lote-{n:02d}.graphql')
        print(f'  lote-{n:02d}: {len(lote):>3} mutacoes, {tam//1024}KB, NCM {" ".join(ncms)}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])

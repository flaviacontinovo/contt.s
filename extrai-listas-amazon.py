#!/usr/bin/env python3
"""
Extrai TODAS as listas fechadas do modelo da Amazon, por tipo de produto.

As listas mudam de um tipo para outro, e nao so no conteudo: SHORTS escreve
"Estampa de animal" onde PANTS escreve "Estampa animal", e SHORTS nao tem
"Todas as estacoes". Escrever o valor de um tipo na linha de outro faz a Amazon
recusar. Por isso o preenchedor resolve cada valor contra a lista certa em vez
de assumir que sao iguais.

Saida: listas-completas.json  ->  {tipo: {coluna: [opcoes]}}
"""
import zipfile, re, html, json, sys
sys.path.insert(0, __import__('os').path.dirname(__file__) or '.')


def num_col(l):
    n = 0
    for c in l: n = n * 26 + ord(c) - 64
    return n


def letra_col(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


def celulas(z, aba):
    xml = z.read(aba).decode('utf-8')
    ss = []
    try:
        sx = z.read('xl/sharedStrings.xml').decode('utf-8')
        for si in re.findall(r'<si>(.*?)</si>', sx, re.S):
            ss.append(html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))))
    except KeyError:
        pass
    fora = {}
    for m in re.finditer(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
        n, corpo, i = int(m.group(1)), m.group(2), 0
        fora[n] = {}
        while True:
            ini = corpo.find('<c ', i)
            if ini < 0: break
            ft = corpo.find('>', ini)
            auto = corpo[ft - 1] == '/'
            fim = ft + 1 if auto else corpo.find('</c>', ft) + 4
            cel = corpo[ini:fim]
            ref = re.search(r'r="([A-Z]+)\d+"', cel)
            if ref:
                t = re.search(r' t="([^"]+)"', cel)
                v = re.search(r'<v>(.*?)</v>', cel, re.S)
                if t and t.group(1) == 'inlineStr':
                    val = html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', cel, re.S)))
                elif v:
                    bruto = html.unescape(v.group(1))
                    val = ss[int(bruto)] if (t and t.group(1) == 's' and bruto.isdigit()
                                             and int(bruto) < len(ss)) else bruto
                else:
                    val = ''
                if val: fora[n][ref.group(1)] = val
            i = fim
    return fora


def main(xlsm, saida):
    z = zipfile.ZipFile(xlsm)
    modelo = 'xl/worksheets/sheet5.xml'
    xml = z.read(modelo).decode('utf-8')

    # coluna -> formula da validacao, so na faixa de dados (linha 7 em diante)
    val = {}
    for m in re.finditer(r'<dataValidation ([^>]*?)(?:/>|>(.*?)</dataValidation>)', xml, re.S):
        at, corpo = m.group(1), m.group(2) or ''
        sq = re.search(r'sqref="([A-Z]+)7:[A-Z]+1048576"', at)
        f1 = re.search(r'<formula1>(.*?)</formula1>', corpo, re.S)
        if sq and f1: val[sq.group(1)] = html.unescape(f1.group(1)).strip()

    wb = z.read('xl/workbook.xml').decode('utf-8')
    defn = dict(re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]+)</definedName>', wb))
    DL = celulas(z, 'xl/worksheets/sheet9.xml')   # aba oculta Dropdown Lists

    def faixa(nome):
        ref = defn.get(nome)
        if not ref: return None
        m = re.search(r"!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", ref)
        if not m: return None
        return [v for r in range(int(m.group(2)), int(m.group(4)) + 1)
                for cn in range(num_col(m.group(1)), num_col(m.group(3)) + 1)
                if (v := DL.get(r, {}).get(letra_col(cn), ''))]

    tipos = faixa(val.get('B', '').lstrip('=').strip()) or []
    fora = {}
    for t in tipos:
        fora[t] = {}
        for col, f in val.items():
            if f.startswith('INDIRECT'):
                s = re.search(r'&"([^"]+)"\)?$', f)
                o = faixa(t.replace('-', '_').replace(' ', '') + s.group(1)) if s else None
            else:
                o = faixa(f.lstrip('=').strip())
            if o: fora[t][col] = o

    json.dump(fora, open(saida, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'{len(tipos)} tipos: ' + ', '.join(tipos))
    for t in tipos:
        print(f'  {t:<18}{len(fora[t])} colunas com lista fechada')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])

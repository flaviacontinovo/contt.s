# -*- coding: utf-8 -*-
"""
Preenche os dados fiscais dos 3 conjuntos Sculpt no Mercado Livre.

Diferente da planilha fiscal anterior, esta nao traz nenhuma linha ja
preenchida para servir de referencia -- sao so os 3 anuncios novos. Entao os
valores vem do padrao que a propria loja ja usa nos outros conjuntos, que foi
lido das linhas que a Flavia mesma tinha preenchido:

    tipo de origem   Nacional
    origem           0 - Nacional, exceto as indicadas nos codigos 3 a 5
    unidade          UN
    CSOSN de venda   102 (tributada pelo Simples Nacional sem credito)
    NCM              61042300, conjunto de malha de fibra sintetica feminino
    peso             0,300 kg liquido e 0,350 kg bruto

Cada valor escrito nas colunas de lista e conferido contra a lista da aba
"list-data" antes de gravar, para nao repetir o erro de mandar algo fora do
que a categoria aceita.

A descricao para NF-e repete o titulo do anuncio, como nas linhas que ela
preencheu. Como cada codigo de produto aparece uma vez so aqui, nao ha o
conflito de "configuracoes diferentes para o mesmo codigo" da rodada passada.
"""
import html
import re
import zipfile

ORIGINAL = "/root/.claude/uploads/94e285ac-89f9-5c02-a1e9-19cb80af0e4a/137961be-Dados_Fiscais-2026_09_15.xlsx"
SAIDA = "/home/user/contt.s/Dados_Fiscais-2026_09_15.xlsx"
ABA = "xl/worksheets/sheet4.xml"
LINHA_CABECALHO = 2

VALORES = {
    "Tipo de origem": "Nacional",
    "Origem": "0 - Nacional, exceto as indicadas nos códigos 3 a 5",
    "NCM": "61042300",
    "Unidade de medida comercial": "UN",
    "CSOSN de venda": "102",
    "Peso líquido": 0.3,
    "Peso bruto": 0.35,
}
NUMERICAS = {"Peso líquido", "Peso bruto"}
# colunas de lista e onde conferir o valor permitido
LISTAS = {"Tipo de origem": None,          # lista literal na validacao
          "Origem": ("list-data", 1),
          "CSOSN de venda": ("list-data", 2)}


def separa_celulas(corpo):
    saida, i = [], 0
    while True:
        ini = corpo.find("<c ", i)
        if ini < 0:
            return saida
        ft = corpo.find(">", ini)
        auto = corpo[ft - 1] == "/"
        fim = ft + 1 if auto else corpo.find("</c>", ft) + 4
        cel = corpo[ini:fim]
        ref = re.search(r'r="([A-Z]+)\d+"', cel)
        if ref:
            saida.append((ref.group(1), cel))
        i = fim


def col_num(l):
    n = 0
    for ch in l:
        n = n * 26 + (ord(ch) - 64)
    return n


def estilo(cel):
    m = re.search(r's="(\d+)"', cel)
    return m.group(1) if m else "0"


def idx(cel):
    if 't="s"' not in cel:
        return None
    m = re.search(r"<v>(\d+)</v>", cel)
    return int(m.group(1)) if m else None


def limpo(t):
    return html.unescape(re.sub(r"<[^>]+>", "", t)).strip()


def main():
    zin = zipfile.ZipFile(ORIGINAL)
    sheet = zin.read(ABA).decode("utf-8")
    shared = zin.read("xl/sharedStrings.xml").decode("utf-8")
    textos = [limpo(t) for t in re.findall(r"<si>(.*?)</si>", shared, re.S)]

    # ---- colunas pelo cabecalho
    m = re.search(r'<row r="{}"(?: [^>]*)?>(.*?)</row>'.format(LINHA_CABECALHO), sheet, re.S)
    colunas = {}
    for col, cel in separa_celulas(m.group(1)):
        i = idx(cel)
        if i is not None:
            colunas[textos[i]] = col
    faltando = [k for k in VALORES if k not in colunas]
    assert not faltando, "coluna ausente: {}".format(faltando)

    # ---- confere os valores de lista contra a aba list-data
    ld = zin.read("xl/worksheets/sheet3.xml").decode("utf-8")
    permitidos = {}
    for r in (1, 2):
        mm = re.search(r'<row r="{}"(?: [^>]*)?>(.*?)</row>'.format(r), ld, re.S)
        if mm:
            permitidos[r] = [textos[idx(c)] for _, c in separa_celulas(mm.group(1)) if idx(c) is not None]
    for nome, onde in LISTAS.items():
        if onde and onde[1] in permitidos:
            v = VALORES[nome]
            # A Origem grava o rotulo inteiro; o CSOSN grava so o codigo ("102"),
            # que e como as linhas que o Mercado Livre ja aceitou estao gravadas.
            casa = [x for x in permitidos[onde[1]] if x == v or x.startswith(v + " -")]
            assert casa, "{} = {!r} fora da lista da categoria".format(nome, v)
            print("  ok: {} = {!r} -> {!r}".format(nome, v, casa[0][:52]))

    # ---- preenche as linhas de produto
    novos, celulas_novas, mapa = [], 0, {}

    def idx_texto(t):
        if t in mapa:
            return mapa[t]
        i = len(textos) + len(novos)
        mapa[t] = i
        novos.append(t)
        return i

    col_anuncio = colunas["Código do anúncio"]
    col_desc_anuncio = colunas["Descrição do anúncio"]
    col_nfe = colunas["Descrição do produto para NF-e"]

    r, feitas = LINHA_CABECALHO + 2, 0
    while True:
        m = re.search(r'<row r="{}"(?: [^>]*)?>(.*?)</row>'.format(r), sheet, re.S)
        if not m:
            break
        cels = dict(separa_celulas(m.group(1)))
        if col_anuncio not in cels or idx(cels[col_anuncio]) is None:
            break

        alvo = dict(VALORES)
        # a descricao para NF-e reaproveita o proprio titulo do anuncio
        cels[col_nfe] = '<c r="{}{}" s="{}" t="s"><v>{}</v></c>'.format(
            col_nfe, r, estilo(cels.get(col_nfe, '<c s="0"/>')), idx(cels[col_desc_anuncio]))

        for nome, val in alvo.items():
            col = colunas[nome]
            s = estilo(cels.get(col, '<c s="0"/>'))
            if nome in NUMERICAS:
                cels[col] = '<c r="{}{}" s="{}" t="n"><v>{}</v></c>'.format(col, r, s, val)
            else:
                if col not in cels:
                    celulas_novas += 1
                cels[col] = '<c r="{}{}" s="{}" t="s"><v>{}</v></c>'.format(
                    col, r, s, idx_texto(val))

        abertura = re.search(r'<row r="{}"(?: [^>]*)?>'.format(r), sheet).group(0)
        corpo = "".join(cels[c] for c in sorted(cels, key=col_num))
        sheet = sheet.replace(m.group(0), abertura + corpo + "</row>", 1)
        feitas += 1
        r += 1

    if novos:
        bloco = "".join('<si><t xml:space="preserve">{}</t></si>'.format(
            t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")) for t in novos)
        shared = shared.replace("</sst>", bloco + "</sst>")
        c = re.search(r'<sst count="(\d+)" uniqueCount="(\d+)"', shared)
        shared = shared.replace(c.group(0), '<sst count="{}" uniqueCount="{}"'.format(
            int(c.group(1)) + celulas_novas, int(c.group(2)) + len(novos)))

    with zipfile.ZipFile(SAIDA, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == ABA:
                zout.writestr(item, sheet.encode("utf-8"))
            elif item.filename == "xl/sharedStrings.xml":
                zout.writestr(item, shared.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    zin.close()
    print("\nlinhas preenchidas: {} ({} a {})".format(feitas, LINHA_CABECALHO + 2, r - 1))
    print("arquivo: {}".format(SAIDA))


if __name__ == "__main__":
    main()

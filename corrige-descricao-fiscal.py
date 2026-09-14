# -*- coding: utf-8 -*-
"""
Corrige a descricao para NF-e das linhas que o Mercado Livre recusou.

O erro foi "Descricao do produto para NF-e: Configuracoes diferentes
encontradas para o mesmo codigo de produto".

Causa: seis SKUs de legging estao em dois anuncios ao mesmo tempo, com titulos
diferentes -- CTTS-LEGCINALTJ-M-BRAN aparece como "Branco Brocado M" no
MLB5006004501 e como "Branco Trabalhado M" no MLB7521788338. Como a descricao
para NF-e vinha do titulo do anuncio, o mesmo codigo de produto ficou com duas
descricoes, e a ficha fiscal e por codigo de produto, nao por anuncio.

Correcao: quando um SKU tem descricoes diferentes, todas as linhas dele passam
a usar as palavras comuns a essas descricoes, na ordem em que aparecem. Assim a
descricao continua verdadeira para os dois anuncios e vira uma so.
"""
import re
import zipfile
from collections import OrderedDict, defaultdict

ORIGINAL = "/root/.claude/uploads/94e285ac-89f9-5c02-a1e9-19cb80af0e4a/10534f9c-ErrosDados_Fiscais-2026_09_14.xlsx"
SAIDA = "/home/user/contt.s/ErrosDados_Fiscais-2026_09_14.xlsx"
ABA = "xl/worksheets/sheet4.xml"
PRIMEIRA = 4


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


def indice(cel):
    if 't="s"' not in cel:
        return None
    m = re.search(r"<v>(\d+)</v>", cel)
    return int(m.group(1)) if m else None


def estilo(cel):
    m = re.search(r's="(\d+)"', cel)
    return m.group(1) if m else "0"


def descricao_comum(descricoes):
    """Palavras que aparecem em todas as descricoes, na ordem da primeira."""
    listas = [d.split() for d in descricoes]
    base = listas[0]
    comuns = [p for p in base if all(p in outra for outra in listas[1:])]
    return " ".join(comuns)


def main():
    zin = zipfile.ZipFile(ORIGINAL)
    sheet = zin.read(ABA).decode("utf-8")
    shared = zin.read("xl/sharedStrings.xml").decode("utf-8")
    textos = re.findall(r"<si>(?:<t[^>]*>)?(.*?)(?:</t>)?</si>", shared, re.S)

    # ---- levanta as linhas e agrupa por codigo de produto
    linhas = OrderedDict()
    r = PRIMEIRA
    while True:
        m = re.search(r'<row r="%d"(?: [^>]*)?>(.*?)</row>' % r, sheet, re.S)
        if not m:
            break
        cels = dict(separa_celulas(m.group(1)))
        if "B" not in cels or indice(cels["B"]) is None:
            break
        linhas[r] = cels
        r += 1

    por_sku = defaultdict(list)
    for r, cels in linhas.items():
        por_sku[textos[indice(cels["H"])]].append(r)

    # ---- resolve os conflitos
    novos = []
    celulas_novas = 0

    def idx_texto(t):
        if t in textos:
            return textos.index(t)
        if t in novos:
            return len(textos) + novos.index(t)
        novos.append(t)
        return len(textos) + len(novos) - 1

    ajustes = {}
    for sku, rs in sorted(por_sku.items()):
        descs = {textos[indice(linhas[r]["R"])] for r in rs}
        if len(descs) == 1:
            continue
        comum = descricao_comum(sorted(descs))
        assert comum, "sem palavras em comum para o SKU {}".format(sku)
        ajustes[sku] = (sorted(descs), comum)
        for rr in rs:
            linhas[rr]["_nova_desc"] = comum

    print("SKUs em mais de um anuncio, com descricao unificada:")
    for sku, (antigas, nova) in ajustes.items():
        print("  {}".format(sku))
        for a in antigas:
            print("      antes: {}".format(a))
        print("      agora: {}".format(nova))

    # ---- reescreve so a coluna R das linhas afetadas
    trocadas = 0
    for r, cels in linhas.items():
        nova = cels.pop("_nova_desc", None)
        if not nova:
            continue
        i = idx_texto(nova)
        if i >= len(textos):
            celulas_novas += 1
        cels["R"] = '<c r="R{}" s="{}" t="s"><v>{}</v></c>'.format(r, estilo(cels["R"]), i)
        corpo = "".join(cels[c] for c in sorted(cels, key=col_num))
        m = re.search(r'<row r="%d"(?: [^>]*)?>(.*?)</row>' % r, sheet, re.S)
        abertura = re.search(r'<row r="%d"(?: [^>]*)?>' % r, sheet).group(0)
        sheet = sheet.replace(m.group(0), abertura + corpo + "</row>", 1)
        trocadas += 1

    # ---- limpa a coluna de erros das linhas corrigidas
    for r in linhas:
        m = re.search(r'<row r="%d"(?: [^>]*)?>(.*?)</row>' % r, sheet, re.S)
        corpo = m.group(1)
        cels = dict(separa_celulas(corpo))
        if "A" in cels and indice(cels["A"]) is not None:
            cels["A"] = '<c r="A{}" s="{}"/>'.format(r, estilo(cels["A"]))
            abertura = re.search(r'<row r="%d"(?: [^>]*)?>' % r, sheet).group(0)
            novo_corpo = "".join(cels[c] for c in sorted(cels, key=col_num))
            sheet = sheet.replace(m.group(0), abertura + novo_corpo + "</row>", 1)

    if novos:
        bloco = "".join('<si><t xml:space="preserve">{}</t></si>'.format(t) for t in novos)
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
    print("\nlinhas no arquivo: {} | descricoes trocadas: {}".format(len(linhas), trocadas))
    print("arquivo: {}".format(SAIDA))


if __name__ == "__main__":
    main()

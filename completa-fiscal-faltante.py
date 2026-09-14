# -*- coding: utf-8 -*-
"""
Completa as linhas de dados fiscais que ficaram em branco.

A regra e a mais segura possivel: cada linha vazia copia a ficha de outra
linha JA PREENCHIDA do mesmo anuncio. Sao o mesmo produto, mudando so o
tamanho, entao NCM, origem, CSOSN, unidade, peso e descricao de NF-e sao
necessariamente iguais. Nada e deduzido de fora do proprio arquivo.

Se alguma linha vazia nao tiver irma preenchida no mesmo anuncio, o script
para e avisa, em vez de chutar.
"""
import re
import sys
import zipfile

ABA = "xl/worksheets/sheet4.xml"
PRIMEIRA = 4
CAMPOS = ["K", "L", "M", "N", "O", "R", "S", "T", "U", "W"]


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


def conteudo(cel):
    """Tipo e valor da celula: ('s', indice) | ('n', numero) | None se vazia."""
    if "<v>" not in cel:
        return None
    v = re.search(r"<v>([^<]*)</v>", cel).group(1)
    return ("s", int(v)) if 't="s"' in cel else ("n", v)


def estilo(cel):
    m = re.search(r's="(\d+)"', cel)
    return m.group(1) if m else "0"


def main(origem, saida):
    zin = zipfile.ZipFile(origem)
    sheet = zin.read(ABA).decode("utf-8")
    textos = re.findall(r"<si>(?:<t[^>]*>)?(.*?)(?:</t>)?</si>",
                        zin.read("xl/sharedStrings.xml").decode("utf-8"), re.S)

    linhas, r = {}, PRIMEIRA
    while True:
        m = re.search(r'<row r="%d"(?: [^>]*)?>(.*?)</row>' % r, sheet, re.S)
        if not m:
            break
        cels = dict(separa_celulas(m.group(1)))
        if "B" not in cels or conteudo(cels["B"]) is None:
            break
        linhas[r] = cels
        r += 1

    completas = [r for r, c in linhas.items() if conteudo(c.get("M", "")) is not None]
    vazias = [r for r in linhas if r not in completas]
    print("linhas no arquivo: {} | ja completas: {} | a completar: {}".format(
        len(linhas), len(completas), vazias or "nenhuma"))

    if not vazias:
        print("nada a fazer.")
        return

    trocadas = 0
    for r in vazias:
        anuncio = conteudo(linhas[r]["B"])[1]
        irmas = [x for x in completas if conteudo(linhas[x]["B"])[1] == anuncio]
        if not irmas:
            sys.exit("linha {} ({}) nao tem irma preenchida no mesmo anuncio; "
                     "nao vou chutar a ficha fiscal".format(r, textos[conteudo(linhas[r]['H'])[1]]))
        fonte = linhas[irmas[0]]
        print("  L{} {} <- copia de L{} ({})".format(
            r, textos[conteudo(linhas[r]["H"])[1]], irmas[0],
            textos[conteudo(fonte["R"])[1]]))

        for col in CAMPOS:
            val = conteudo(fonte.get(col, ""))
            s = estilo(linhas[r].get(col, '<c s="0"/>'))
            if val is None:
                linhas[r][col] = '<c r="{}{}" s="{}"/>'.format(col, r, s)
            elif val[0] == "s":
                linhas[r][col] = '<c r="{}{}" s="{}" t="s"><v>{}</v></c>'.format(col, r, s, val[1])
            else:
                linhas[r][col] = '<c r="{}{}" s="{}" t="n"><v>{}</v></c>'.format(col, r, s, val[1])

        if "A" in linhas[r]:
            linhas[r]["A"] = '<c r="A{}" s="{}"/>'.format(r, estilo(linhas[r]["A"]))

        m = re.search(r'<row r="%d"(?: [^>]*)?>(.*?)</row>' % r, sheet, re.S)
        abertura = re.search(r'<row r="%d"(?: [^>]*)?>' % r, sheet).group(0)
        corpo = "".join(linhas[r][c] for c in sorted(linhas[r], key=col_num))
        sheet = sheet.replace(m.group(0), abertura + corpo + "</row>", 1)
        trocadas += 1

    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, sheet.encode("utf-8") if item.filename == ABA
                          else zin.read(item.filename))
    zin.close()
    print("\nlinhas completadas: {}\narquivo: {}".format(trocadas, saida))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

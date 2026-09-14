# -*- coding: utf-8 -*-
"""
Preenche os dados fiscais dos anuncios da CONTT.s no Mercado Livre.

O metodo aqui nao e inventar classificacao fiscal: a planilha ja vinha com 4
linhas preenchidas por ela (as leggings brocadas pretas), e essas linhas sao o
padrao que a empresa ja usa. Todas as outras 27 linhas recebem o mesmo padrao.

Duas coisas mudam, e mudam por motivo:

  NCM   as linhas prontas usam 61046300, que e calca de fibra sintetica de
        malha, feminina -- certo para legging avulsa. Os conjuntos nao sao
        calca: sao legging + top do mesmo tecido vendidos juntos, o que na NCM
        e "conjunto", posicao 61042300. As leggings novas seguem com 61046300.

  PESO  as linhas prontas declaram 0,200 kg liquido e 0,250 kg bruto para uma
        legging. O conjunto leva legging mais top, entao vai 0,300 e 0,350 --
        derivado do numero dela, nao de estimativa solta.

Como na planilha de anuncio, a edicao e feita direto no XML para nao perder
validacoes, estilos, imagens nem as abas ocultas.
"""
import re
import zipfile

ORIGINAL = "/root/.claude/uploads/94e285ac-89f9-5c02-a1e9-19cb80af0e4a/90e54c39-Dados_Fiscais-2026_09_14.xlsx"
SAIDA = "/home/user/contt.s/Dados_Fiscais-2026_09_14.xlsx"
ABA = "xl/worksheets/sheet4.xml"

PRIMEIRA, ULTIMA = 4, 34
LINHA_MODELO = 16          # linha que ela mesma preencheu, usada como referencia

NCM_CONJUNTO = "61042300"
NCM_LEGGING = "61046300"
PESO_CONJUNTO = (0.3, 0.35)
PESO_LEGGING = (0.2, 0.25)


def separa_celulas(corpo):
    """(coluna, xml) de cada <c> da linha, aguentando celula vazia e CDATA."""
    saida, i = [], 0
    while True:
        ini = corpo.find("<c ", i)
        if ini < 0:
            return saida
        fim_tag = corpo.find(">", ini)
        auto = corpo[fim_tag - 1] == "/"
        fim = fim_tag + 1 if auto else corpo.find("</c>", fim_tag) + 4
        cel = corpo[ini:fim]
        ref = re.search(r'r="([A-Z]+)\d+"', cel)
        if ref:
            saida.append((ref.group(1), cel))
        i = fim


def col_num(letras):
    n = 0
    for ch in letras:
        n = n * 26 + (ord(ch) - 64)
    return n


def linha_xml(sheet, r):
    m = re.search(r'<row r="%d"(?: [^>]*)?>(.*?)</row>' % r, sheet, re.S)
    assert m, "linha %d nao encontrada" % r
    return m


def estilo_de(cel):
    m = re.search(r's="(\d+)"', cel)
    return m.group(1) if m else "0"


def indice_de(cel):
    """Indice do sharedStrings da celula, ou None se estiver vazia."""
    if 't="s"' not in cel:
        return None
    m = re.search(r"<v>(\d+)</v>", cel)
    return int(m.group(1)) if m else None


def main():
    zin = zipfile.ZipFile(ORIGINAL)
    sheet = zin.read(ABA).decode("utf-8")
    shared = zin.read("xl/sharedStrings.xml").decode("utf-8")
    textos = re.findall(r"<si>(?:<t[^>]*>)?(.*?)(?:</t>)?</si>", shared, re.S)

    # ---- le da linha modelo os valores que a empresa ja usa
    modelo = dict(separa_celulas(linha_xml(sheet, LINHA_MODELO).group(1)))
    padrao = {c: indice_de(modelo[c]) for c in ("K", "L", "S", "W")}
    assert None not in padrao.values(), "linha modelo sem os campos de referencia"
    idx_ncm_legging = indice_de(modelo["M"])
    assert textos[idx_ncm_legging] == NCM_LEGGING, \
        "NCM da linha modelo mudou: {}".format(textos[idx_ncm_legging])
    print("padrao lido da linha {}:".format(LINHA_MODELO))
    for c in ("K", "L", "S", "W"):
        print("   {} = {}".format(c, textos[padrao[c]]))

    # ---- o NCM de conjunto pode nao existir ainda no arquivo
    novos = []
    if NCM_CONJUNTO in textos:
        idx_ncm_conjunto = textos.index(NCM_CONJUNTO)
    else:
        idx_ncm_conjunto = len(textos)
        novos.append(NCM_CONJUNTO)

    preenchidas, puladas = 0, []
    texto_novo_em_celulas = 0

    for r in range(PRIMEIRA, ULTIMA + 1):
        m = linha_xml(sheet, r)
        cels = dict(separa_celulas(m.group(1)))
        if "B" not in cels or indice_de(cels["B"]) is None:
            continue                                  # linha sem anuncio
        if indice_de(cels.get("M", "")) is not None:
            puladas.append(r)                          # ela ja preencheu
            continue

        desc = textos[indice_de(cels["C"])]
        e_conjunto = desc.lower().startswith("conjunto")
        ncm = idx_ncm_conjunto if e_conjunto else idx_ncm_legging
        liquido, bruto = PESO_CONJUNTO if e_conjunto else PESO_LEGGING

        valores = {
            "K": ("s", padrao["K"]),
            "L": ("s", padrao["L"]),
            "M": ("s", ncm),
            # a descricao para a NF-e repete o titulo do anuncio, como nas
            # linhas que ela preencheu: reaproveita o mesmo texto da coluna C
            "R": ("s", indice_de(cels["C"])),
            "S": ("s", padrao["S"]),
            "T": ("n", liquido),
            "U": ("n", bruto),
            "W": ("s", padrao["W"]),
        }

        for col, (tipo, val) in valores.items():
            s = estilo_de(cels[col])
            if tipo == "n":
                cels[col] = '<c r="{}{}" s="{}" t="n"><v>{}</v></c>'.format(col, r, s, val)
            else:
                if val >= len(textos):
                    texto_novo_em_celulas += 1
                cels[col] = '<c r="{}{}" s="{}" t="s"><v>{}</v></c>'.format(col, r, s, val)

        corpo = "".join(cels[c] for c in sorted(cels, key=col_num))
        abertura = re.search(r'<row r="%d"(?: [^>]*)?>' % r, sheet).group(0)
        sheet = sheet.replace(m.group(0), abertura + corpo + "</row>", 1)
        preenchidas += 1

    if novos:
        bloco = "".join('<si><t xml:space="preserve">{}</t></si>'.format(t) for t in novos)
        shared = shared.replace("</sst>", bloco + "</sst>")
        c = re.search(r'<sst count="(\d+)" uniqueCount="(\d+)"', shared)
        shared = shared.replace(c.group(0), '<sst count="{}" uniqueCount="{}"'.format(
            int(c.group(1)) + texto_novo_em_celulas, int(c.group(2)) + len(novos)))

    with zipfile.ZipFile(SAIDA, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == ABA:
                zout.writestr(item, sheet.encode("utf-8"))
            elif item.filename == "xl/sharedStrings.xml":
                zout.writestr(item, shared.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    zin.close()

    print("\nlinhas preenchidas: {}".format(preenchidas))
    print("linhas que ela ja tinha preenchido, intocadas: {}".format(puladas))
    print("arquivo: {}".format(SAIDA))


if __name__ == "__main__":
    main()

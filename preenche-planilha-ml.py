# -*- coding: utf-8 -*-
"""
Preenche a planilha de anuncio em massa do Mercado Livre (categoria Conjuntos)
com os 5 conjuntos das linhas Shape e Aura da CONTT.s.

O arquivo do Mercado Livre e carimbado por dentro: tem uma aba veryHidden
("extra info") com as listas de valores, 22 validacoes, formulas de conferencia
e imagens. Por isso o arquivo NAO e reescrito por uma biblioteca de planilha,
que perderia parte disso: a edicao e feita direto no XML, trocando so os valores
das celulas das linhas 7 a 21 e acrescentando os textos novos ao sharedStrings.
Tudo o mais do pacote e copiado byte a byte.

Colunas obrigatorias, segundo a propria formula de erro da planilha (coluna AA):
C, D, E, H, J, K, M, O, P, Q, U, V e W.
"""
import re
import shutil
import zipfile

ORIGINAL = "/root/.claude/uploads/94e285ac-89f9-5c02-a1e9-19cb80af0e4a/f8356a79-Anunciar-09-14-19_23_19.xlsx"
SAIDA = "/home/user/contt.s/Anunciar-09-14-19_23_19.xlsx"
ABA = "xl/worksheets/sheet3.xml"
PRIMEIRA_LINHA = 7

CDN = "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/"

DESC_SHAPE = (
    "O Conjunto {cor} foi desenvolvido para valorizar suas curvas com conforto, "
    "segurança e muito estilo. Confeccionado em poliamida de alta qualidade, possui "
    "toque macio, excelente elasticidade e ZERO TRANSPARÊNCIA, proporcionando mais "
    "segurança durante os treinos.\n\n"
    "O top possui bojo, oferecendo melhor sustentação e valorizando o busto. A legging "
    "conta com modelagem estratégica que valoriza e realça o bumbum, desenhando as "
    "curvas e proporcionando um efeito ainda mais bonito ao corpo.\n\n"
    "Um conjunto perfeito para quem busca unir performance, conforto e uma modelagem "
    "que valoriza a silhueta.\n\n"
    "COMPOSIÇÃO\n"
    "Poliamida com elastano. Legging de cintura alta e top com bojo.\n\n"
    "COMO ESCOLHER O TAMANHO\n"
    "O conjunto vai com a legging e o top no mesmo tamanho. Se você precisa de tamanhos "
    "diferentes entre as duas peças, nos chame antes de comprar que separamos. A tabela "
    "de medidas está nas fotos do anúncio."
)

DESC_AURA = (
    "O Conjunto Aura une conforto, sofisticação e uma modelagem desenvolvida para "
    "valorizar suas curvas. Confeccionado em poliamida premium, possui toque "
    "extremamente macio, confortável e agradável ao corpo, além de oferecer ZERO "
    "TRANSPARÊNCIA, garantindo mais segurança durante os treinos.\n\n"
    "O top acompanha bojo removível, proporcionando melhor sustentação, conforto e "
    "valorizando o busto. Sua modelagem se ajusta perfeitamente ao corpo para "
    "acompanhar seus movimentos.\n\n"
    "A legging possui cintura alta, ajudando a disfarçar a região da barriga e modelar "
    "a silhueta. Além disso, sua modelagem estratégica valoriza e realça o bumbum, "
    "destacando as curvas de forma bonita e natural.\n\n"
    "COMPOSIÇÃO\n"
    "Poliamida com elastano. Legging de cintura alta e top com bojo removível.\n\n"
    "COMO ESCOLHER O TAMANHO\n"
    "O conjunto vai com a legging e o top no mesmo tamanho. Se você precisa de tamanhos "
    "diferentes entre as duas peças, nos chame antes de comprar que separamos. A tabela "
    "de medidas está nas fotos do anúncio."
)

# As 9 combinacoes de legging x top que existem na loja. No Mercado Livre o
# tamanho e um so por variacao, entao cada tamanho aqui vale o MENOR estoque
# entre as combinacoes que o usam: so da para entregar um P se houver legging P
# e top P sobrando.
COMBOS = ["P/P", "M/P", "G/P", "P/M", "P/G", "M/M", "M/G", "G/M", "G/G"]

PRODUTOS = [
    {
        "linha": "Shape", "cor": "Matcha", "desc": DESC_SHAPE.format(cor="Matcha"),
        "sku": "CTS-MATCHA",
        "fotos": ["1_79b30e4a-aeb6-41b7-8fed-2fdb2171fd34.png",
                  "3_1732dfaa-7367-4ad9-aa42-523be0187a90.png",
                  "2_55ecdae7-11f5-444e-8209-e77864ad95a4.png",
                  "2CE0018D-4C58-4BD1-AF17-2ACB97C88580.png"],
        "est": {"P/P": 13, "M/P": 13, "G/P": 13, "P/M": 13, "P/G": 13,
                "M/M": 13, "M/G": 12, "G/M": 13, "G/G": 13},
    },
    {
        "linha": "Shape", "cor": "Off White", "desc": DESC_SHAPE.format(cor="Off White"),
        "sku": "CTS-OFFWHITE",
        "fotos": ["6_da2222ad-a81d-42cd-9b4e-ed2f00b32c3c.png",
                  "4_070ccd73-5eb7-4560-b6ba-d54a369ae2d0.png",
                  "5_aab050bd-6eac-458c-a8d1-400b3e1894bd.png",
                  "2CF040B5-41D0-40B2-8023-04AE68142EC1.png"],
        "est": {"P/P": 12, "M/P": 12, "G/P": 12, "P/M": 12, "P/G": 12,
                "M/M": 12, "M/G": 11, "G/M": 12, "G/G": 12},
    },
    {
        "linha": "Aura", "cor": "Petróleo", "desc": DESC_AURA, "sku": "CTS-AURA",
        "fotos": ["11_e3552932-1780-4696-8c2c-93416e4fc612.png",
                  "12_e35d7013-5d49-41c7-98d1-509bd0ddcc28.png",
                  "13_898ab4d8-a1ce-4fe0-afaa-83496462a715.png",
                  "14_48939258-2086-4b46-add5-675e4955adb9.png"],
        "est": {c: 14 for c in COMBOS},
    },
    {
        "linha": "Aura", "cor": "Solar Orange", "desc": DESC_AURA, "sku": "CTS-AURA-SOLAR",
        "fotos": ["4_51508998-643b-41ed-a3d1-688345ebf40f.png",
                  "1_862fde68-29ab-423b-85cb-6622cfdcacd3.png",
                  "2_93fcfeef-e45b-4c53-8ff7-879e437f8c20.png",
                  "3_a752fa1f-6f85-4757-89f3-05006ef13d7c.png",
                  "5_a63365ea-a917-43a2-9c75-cf321e6444f9.png"],
        "est": {c: 14 for c in COMBOS},
    },
    {
        "linha": "Aura", "cor": "Bronze", "desc": DESC_AURA, "sku": "CTS-AURA-BRONZE",
        "fotos": ["6_8a345ae5-9d03-479b-8104-ad3d564facf3.png",
                  "9_84752be2-f1ee-47e4-8788-919b83d40b52.png",
                  "7_26b7815b-1529-4c74-b681-6dd89e56e98a.png",
                  "10_0653cd3c-7a50-4762-974f-5a55b71b1827.png",
                  "8_7684cb2f-6b22-4982-a402-0ae092619f47.png"],
        "est": {c: 13 for c in COMBOS},
    },
]


def estoque_do_tamanho(p, tam):
    usados = [c for c in COMBOS if tam in c.split("/")]
    return min(p["est"][c] for c in usados)


def titulo(p):
    t = "Conjunto Fitness Legging Cintura Alta + Top " + p["cor"]
    assert len(t) <= 60, "titulo com {} caracteres: {}".format(len(t), t)
    return t


def monta_linhas():
    """Uma linha por variacao. Titulo, preco, descricao e fotos se repetem em
    todas as linhas do mesmo anuncio, como a propria planilha manda."""
    linhas = []
    for p in PRODUTOS:
        t = titulo(p)
        fotos = "\n".join(CDN + f for f in p["fotos"])
        modelo = "Conjunto {} {}".format(p["linha"], p["cor"])
        for tam in ["P", "M", "G"]:
            linhas.append({
                "A": t,
                "C": "Novo",
                "D": "O produto não tem código cadastrado",
                "E": p["cor"],
                "F": "Lisa",
                "G": tam,
                "H": fotos,
                "I": "{}-{}".format(p["sku"], tam),
                "J": estoque_do_tamanho(p, tam),
                "K": 189,
                "L": p["desc"],
                "M": "Clássico",
                "O": "Mercado Envios",
                "P": "Por conta do comprador",
                "Q": "Não aceito",
                "R": "Garantia do vendedor",
                "S": 30,
                "T": "dias",
                "U": "Contt.s",
                "V": modelo,
                "W": "Feminino",
                "X": 2,
                "Y": "Poliamida",
                "Z": "Não",
            })
    return linhas


# estilo de cada coluna, copiado das celulas que ja vinham na planilha
ESTILO = {"A": "0", "C": "0", "D": "0", "E": "0", "F": "0", "G": "0", "H": "0",
          "I": "0", "J": "38", "K": "38", "L": "0", "M": "0", "O": "0", "P": "0",
          "Q": "0", "R": "45", "S": "38", "T": "0", "U": "0", "V": "0", "W": "45",
          "X": "38", "Y": "0", "Z": "45"}

# colunas calculadas pela propria planilha; ficam exatamente como vieram
FORMULAS = {"B", "N", "AA", "AB", "AC"}


def separa_celulas(corpo):
    """Devolve (coluna, xml) de cada <c> da linha.

    Nao da para recortar com expressao regular simples: as formulas de
    conferencia da planilha tem <v/> e CDATA dentro, e um .*? ate o primeiro
    "/>" corta a celula no meio. Aqui o fim de cada celula e achado na mao:
    se a tag de abertura fecha em "/>", a celula acabou ali; senao vale o
    proximo "</c>", que nao aninha.
    """
    saida = []
    i = 0
    while True:
        ini = corpo.find("<c ", i)
        if ini < 0:
            break
        fim_tag = corpo.find(">", ini)
        auto = corpo[fim_tag - 1] == "/"
        fim = fim_tag + 1 if auto else corpo.find("</c>", fim_tag) + 4
        cel = corpo[ini:fim]
        ref = re.search(r'r="([A-Z]+)\d+"', cel)
        if ref:
            saida.append((ref.group(1), cel))
        i = fim
    return saida


def col_num(letras):
    n = 0
    for ch in letras:
        n = n * 26 + (ord(ch) - 64)
    return n


def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def main():
    zin = zipfile.ZipFile(ORIGINAL)
    sheet = zin.read(ABA).decode("utf-8")
    shared = zin.read("xl/sharedStrings.xml").decode("utf-8")

    # ---- textos novos vao para o fim do sharedStrings, sem mexer nos antigos
    existentes = re.findall(r"<si>(.*?)</si>", shared, re.S)
    indice = {}
    novos = []

    def idx_texto(txt):
        if txt in indice:
            return indice[txt]
        i = len(existentes) + len(novos)
        indice[txt] = i
        novos.append(txt)
        return i

    linhas = monta_linhas()
    assert len(linhas) == 15, "esperava 15 variacoes, montei {}".format(len(linhas))

    # ---- reescreve so as linhas 7..21
    trocadas = 0
    celulas_texto_novas = 0

    for n, dados in enumerate(linhas):
        r = PRIMEIRA_LINHA + n
        m = re.search(r'<row r="{}"(?: [^>]*)?>(.*?)</row>'.format(r), sheet, re.S)
        assert m, "linha {} nao encontrada".format(r)
        abertura = re.search(r'<row r="{}"(?: [^>]*)?>'.format(r), sheet).group(0)

        antigas = dict(separa_celulas(m.group(1)))

        saida = {}
        for col, cel in antigas.items():
            if col in FORMULAS:
                saida[col] = cel          # formula da planilha: intocada

        for col, val in dados.items():
            s = ESTILO[col]
            if isinstance(val, (int, float)):
                saida[col] = '<c r="{}{}" t="n" s="{}"><v>{}</v></c>'.format(col, r, s, val)
            else:
                if col not in antigas:
                    celulas_texto_novas += 1
                saida[col] = '<c r="{}{}" t="s" s="{}"><v>{}</v></c>'.format(
                    col, r, s, idx_texto(esc(val)))

        corpo = "".join(saida[c] for c in sorted(saida, key=col_num))
        sheet = sheet.replace(m.group(0), abertura + corpo + "</row>", 1)
        trocadas += 1

    # ---- fecha o sharedStrings
    if novos:
        bloco = "".join('<si><t xml:space="preserve">{}</t></si>'.format(t) for t in novos)
        shared = shared.replace("</sst>", bloco + "</sst>")
        cont = re.search(r'<sst count="(\d+)" uniqueCount="(\d+)"', shared)
        shared = shared.replace(
            cont.group(0),
            '<sst count="{}" uniqueCount="{}"'.format(
                int(cont.group(1)) + celulas_texto_novas,
                int(cont.group(2)) + len(novos)))

    # ---- regrava o pacote, copiando tudo o que nao foi tocado
    with zipfile.ZipFile(SAIDA, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == ABA:
                zout.writestr(item, sheet.encode("utf-8"))
            elif item.filename == "xl/sharedStrings.xml":
                zout.writestr(item, shared.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    zin.close()

    print("linhas preenchidas: {} (7 a {})".format(trocadas, PRIMEIRA_LINHA + len(linhas) - 1))
    print("textos novos no sharedStrings: {}".format(len(novos)))
    print("arquivo: {}".format(SAIDA))


if __name__ == "__main__":
    main()

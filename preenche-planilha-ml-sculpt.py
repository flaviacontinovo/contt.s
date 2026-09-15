# -*- coding: utf-8 -*-
"""
Preenche a planilha de anuncio em massa do Mercado Livre com os 3 conjuntos
Sculpt da loja: Preto, Vulcano e Rosa Retro.

Mesmo metodo da planilha anterior: edicao direta no XML do pacote, mexendo so
nas linhas de dados e nos textos novos do sharedStrings, para nao perder as
validacoes, os estilos, as imagens nem a aba oculta "extra info".

Duas licoes da primeira rodada estao aplicadas aqui:

  FOTOS    separadas por VIRGULA. A aba "Ajuda" da planilha manda assim; com
           quebra de linha o Mercado Livre le tudo como uma URL so e recusa
           com "Uma das URLs inseridas e muito longa".

  TAMANHO  um por variacao. A lista fechada da categoria (aba oculta, linha 6)
           aceita P, M, G, PP, GG e afins -- nao aceita P/G. Entao vai o
           conjunto inteiro no mesmo tamanho, e o estoque de cada tamanho e o
           MENOR entre as combinacoes de legging x top que o usam.

A descricao sai sem emoji e sem HTML, que e o que a coluna exige.
"""
import html
import re
import unicodedata
import zipfile

ORIGINAL = "/root/.claude/uploads/94e285ac-89f9-5c02-a1e9-19cb80af0e4a/a18c30e4-Anunciar-09-15-13_45_48.xlsx"
SAIDA = "/home/user/contt.s/Anunciar-09-15-13_45_48.xlsx"
ABA = "xl/worksheets/sheet3.xml"
LINHA_CABECALHO = 3

CDN = "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/"

NOTA_TAMANHO = (
    "\n\nCOMO ESCOLHER O TAMANHO\n"
    "O conjunto vai com a legging e o top no mesmo tamanho. Se voce precisa de "
    "tamanhos diferentes entre as duas pecas, nos chame antes de comprar que "
    "separamos."
)

DESC_PRETO = """O Conjunto Sculpt Preto foi desenvolvido para acompanhar seus movimentos com conforto e seguranca, unindo uma estetica sofisticada a uma modelagem que valoriza a silhueta.

O top possui alcas cruzadas nas costas, proporcionando um visual moderno e feminino, enquanto a legging apresenta cintura alta e modelagem anatomica, envolvendo o corpo e realcando as curvas. O tecido com textura canelada acrescenta um toque sofisticado e deixa a producao ainda mais elegante.

Perfeito para treinos, academia, caminhadas ou para compor um look fitness casual, o Sculpt e aquele conjunto que une performance, conforto e estilo em uma unica producao."""

DESC_VULCANO = """O Conjunto Legging Sculpt + Top Vulcano combina duas pecas pensadas para proporcionar um visual marcante e sofisticado durante os treinos.

A Legging Sculpt possui cintura alta e modelagem anatomica que acompanha o corpo, valorizando a silhueta e proporcionando um ajuste confortavel. Seu tecido texturizado acrescenta um toque sofisticado ao look, enquanto os recortes traseiros ajudam a destacar e valorizar as curvas.

O Top Vulcano completa a producao com seu design de costas cruzadas, trazendo modernidade, feminilidade e seguranca para os movimentos.

DESTAQUES
- Legging de cintura alta
- Modelagem que valoriza as curvas
- Tecido com textura sofisticada
- Top com costas cruzadas
- Excelente ajuste ao corpo
- Conforto para os movimentos
- Ideal para academia, treino e lifestyle"""

DESC_ROSA = DESC_VULCANO.replace("Top Vulcano", "Top Rosa Retro")

PRODUTOS = [
    {
        "cor": "Preto", "sku": "CON-SCULPT-PRE", "estoque": 11, "desc": DESC_PRETO,
        "fotos": ["1_62684014-12f1-4d50-ba0b-60cad8f8b4ca.png",
                  "2_08baa615-3ae1-47c5-8689-fd86aae71140.png",
                  "4_aef518dd-7de3-4815-a69d-03d2e7541007.png",
                  "3_af1ddaa5-1173-4566-bf2d-840cf4db9a98.png",
                  "5_e2b14fd7-9213-4e47-ab31-541ff7e4d5ed.png"],
    },
    {
        "cor": "Vulcano", "sku": "CON-SCULPT-VUL", "estoque": 10, "desc": DESC_VULCANO,
        "fotos": ["1_157719d5-9056-471c-8c4f-1f955a49d646.png",
                  "2_ebe5e5e0-9a75-4cef-97c4-8ab987a72ea4.png",
                  "3_8a12d03d-474b-49dc-b212-1a26db330c40.png",
                  "4_28b8f22d-8c41-4eee-a06f-78753a48298b.png"],
    },
    {
        "cor": "Rosa Retro", "sku": "CON-SCULPT-ROSA", "estoque": 10, "desc": DESC_ROSA,
        "fotos": ["1_baded739-c0c6-49fb-a8fe-92e2d9f7681c.png",
                  "2_d29679ce-70f3-4e07-893a-82482d22b0bd.png",
                  "4_96aa75fd-ddfa-4dba-9bf0-b7f9b1aadda9.png",
                  "3_0ebe3b9d-a422-4ac6-86dc-1a86a3c0e12d.png"],
    },
]

TAMANHOS = ["P", "M", "G"]

NUMERICAS = {"Estoque", "Preço [R$]", "Tempo de garantia", "Quantidade de peças"}
ESTILO_TEXTO, ESTILO_NUM = "0", "38"
ESTILO_ESPECIAL = {"Tipo de garantia": "45", "Gênero": "45", "Materiais reciclados": "45"}
FORMULAS_POR_NOME = {"Quantidade de caracteres", "Tarifa de venda",
                     "Resumo de erros", "BUYBOX_FORMULA", "HIDDEN_PICTURES"}


def titulo(cor):
    t = "Conjunto Fitness Sculpt Legging Cintura Alta Top " + cor
    assert len(t) <= 60, "titulo com {} caracteres: {}".format(len(t), t)
    return t


def sem_emoji(t):
    return "".join(c for c in t if unicodedata.category(c) != "So")


def monta_linhas():
    linhas = []
    for p in PRODUTOS:
        t = titulo(p["cor"])
        fotos = ",".join(CDN + f for f in p["fotos"])
        desc = sem_emoji(p["desc"] + NOTA_TAMANHO)
        assert "<" not in desc and "http" not in desc, "descricao com HTML ou link"
        for tam in TAMANHOS:
            linhas.append({
                "Título": t,
                "Condição": "Novo",
                "Código universal de produto": "O produto não tem código cadastrado",
                "Varia por: Nome comercial da cor": p["cor"],
                "Varia por: Desenho do tecido": "Lisa",
                "Varia por: Tamanho": tam,
                "Fotos": fotos,
                "SKU": "{}-{}".format(p["sku"], tam),
                "Estoque": p["estoque"],
                "Preço [R$]": 189,
                "Descrição": desc,
                "Tipo de anúncio": "Clássico",
                "Forma de envio": "Mercado Envios",
                "Custo de envio": "Por conta do comprador",
                "Retirar pessoalmente": "Não aceito",
                "Tipo de garantia": "Garantia do vendedor",
                "Tempo de garantia": 30,
                "Unidade de Tempo de garantia": "dias",
                "Marca": "Contt.s",
                "Modelo": "Conjunto Sculpt " + p["cor"],
                "Gênero": "Feminino",
                "Quantidade de peças": 2,
                "Material principal": "Poliamida",
                "Materiais reciclados": "Não",
            })
    return linhas


def primeira_linha_de_dados(sheet, textos, colunas):
    """Acha a primeira linha de dados em vez de supor que e a 7.

    Nao da para fixar o numero: nesta planilha a linha 7 e uma faixa de
    celulas mescladas com um aviso ("Revise as condicoes de venda...") e os
    dados so comecam na 8, enquanto na planilha do dia anterior comecavam na
    7. A marca segura e a coluna Condicao ja vir com "Novo" preenchido, que e
    o padrao que o Mercado Livre repete em toda linha de dados.
    """
    col_cond = colunas["Condição"]
    for r in range(LINHA_CABECALHO + 1, LINHA_CABECALHO + 20):
        m = re.search(r'<row r="{}"(?: [^>]*)?>(.*?)</row>'.format(r), sheet, re.S)
        if not m:
            continue
        for col, cel in separa_celulas(m.group(1)):
            if col != col_cond or 't="s"' not in cel:
                continue
            v = re.search(r"<v>(\d+)</v>", cel)
            if v and html.unescape(re.sub(r"<[^>]+>", "", textos[int(v.group(1))])).strip() == "Novo":
                return r
    raise AssertionError("nao achei a primeira linha de dados")


def separa_celulas(corpo):
    """(coluna, xml) de cada <c>. Recorte na mao porque as formulas de
    conferencia da planilha tem <v/> e CDATA dentro."""
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


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def main():
    zin = zipfile.ZipFile(ORIGINAL)
    sheet = zin.read(ABA).decode("utf-8")
    shared = zin.read("xl/sharedStrings.xml").decode("utf-8")
    textos = re.findall(r"<si>(?:<t[^>]*>)?(.*?)(?:</t>)?</si>", shared, re.S)

    # ---- liga nome de coluna a letra, pelo cabecalho
    m = re.search(r'<row r="{}"(?: [^>]*)?>(.*?)</row>'.format(LINHA_CABECALHO), sheet, re.S)
    colunas = {}
    for col, cel in separa_celulas(m.group(1)):
        if 't="s"' not in cel:
            continue
        v = re.search(r"<v>(\d+)</v>", cel)
        if not v:
            continue
        nome = re.sub(r"<[^>]+>", "", textos[int(v.group(1))])
        nome = html.unescape(nome).strip()
        if nome.startswith("Título"):
            nome = "Título"
        colunas[nome] = col

    primeira = primeira_linha_de_dados(sheet, textos, colunas)
    linhas = monta_linhas()
    assert len(linhas) == 9, "esperava 9 linhas, montei {}".format(len(linhas))
    faltando = [k for k in linhas[0] if k not in colunas]
    assert not faltando, "coluna ausente no cabecalho: {}".format(faltando)

    indice, novos, celulas_texto_novas = {}, [], 0

    def idx_texto(t):
        if t in indice:
            return indice[t]
        i = len(textos) + len(novos)
        indice[t] = i
        novos.append(t)
        return i

    formulas = {colunas[n] for n in FORMULAS_POR_NOME if n in colunas}

    for n, dados in enumerate(linhas):
        r = primeira + n
        m = re.search(r'<row r="{}"(?: [^>]*)?>(.*?)</row>'.format(r), sheet, re.S)
        assert m, "linha {} nao encontrada".format(r)
        abertura = re.search(r'<row r="{}"(?: [^>]*)?>'.format(r), sheet).group(0)
        antigas = dict(separa_celulas(m.group(1)))
        saida = {c: cel for c, cel in antigas.items() if c in formulas}

        for nome, val in dados.items():
            col = colunas[nome]
            if nome in NUMERICAS:
                saida[col] = '<c r="{}{}" t="n" s="{}"><v>{}</v></c>'.format(
                    col, r, ESTILO_NUM, val)
            else:
                estilo = ESTILO_ESPECIAL.get(nome, ESTILO_TEXTO)
                if col not in antigas:
                    celulas_texto_novas += 1
                saida[col] = '<c r="{}{}" t="s" s="{}"><v>{}</v></c>'.format(
                    col, r, estilo, idx_texto(esc(val)))

        corpo = "".join(saida[c] for c in sorted(saida, key=col_num))
        sheet = sheet.replace(m.group(0), abertura + corpo + "</row>", 1)

    if novos:
        bloco = "".join('<si><t xml:space="preserve">{}</t></si>'.format(t) for t in novos)
        shared = shared.replace("</sst>", bloco + "</sst>")
        c = re.search(r'<sst count="(\d+)" uniqueCount="(\d+)"', shared)
        shared = shared.replace(c.group(0), '<sst count="{}" uniqueCount="{}"'.format(
            int(c.group(1)) + celulas_texto_novas, int(c.group(2)) + len(novos)))

    with zipfile.ZipFile(SAIDA, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == ABA:
                zout.writestr(item, sheet.encode("utf-8"))
            elif item.filename == "xl/sharedStrings.xml":
                zout.writestr(item, shared.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    zin.close()

    print("colunas mapeadas: {}".format(len(colunas)))
    print("primeira linha de dados detectada: {}".format(primeira))
    print("linhas preenchidas: {} ({} a {})".format(len(linhas), primeira, primeira + len(linhas) - 1))
    for p in PRODUTOS:
        print("   {:<12} titulo {} caracteres | {} por tamanho | {} fotos".format(
            p["cor"], len(titulo(p["cor"])), p["estoque"], len(p["fotos"])))
    print("arquivo: {}".format(SAIDA))


if __name__ == "__main__":
    main()

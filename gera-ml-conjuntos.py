# -*- coding: utf-8 -*-
"""
Monta a planilha de apoio para a publicacao em massa dos conjuntos Shape e Aura
no Mercado Livre.

Os dados de produto, SKU, preco e estoque vem da Shopify (loja CONTT.s FITNESS,
leitura de 14/09/2026). Nada aqui e inventado: o que nao existe na loja fica
como celula amarela para a Flavia preencher.

Nao usa formula de proposito. A planilha existe para ser copiada para dentro do
template do Mercado Livre, e formula colada la quebra o arquivo deles.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONTE = "Arial"
AZUL = "1F3864"
CINZA = "F2F2F2"
AMARELO = "FFF2CC"
BORDA = Border(*[Side(style="thin", color="D9D9D9")] * 4)

# ---------------------------------------------------------------- dados

DESC_SHAPE = (
    "O Conjunto {cor} foi desenvolvido para valorizar suas curvas com conforto, "
    "seguranca e muito estilo. Confeccionado em poliamida de alta qualidade, possui "
    "toque macio, excelente elasticidade e ZERO TRANSPARENCIA, proporcionando mais "
    "seguranca durante os treinos.\n\n"
    "O top possui bojo, oferecendo melhor sustentacao e valorizando o busto. A legging "
    "conta com modelagem estrategica que valoriza e realca o bumbum, desenhando as "
    "curvas e proporcionando um efeito ainda mais bonito ao corpo.\n\n"
    "Um conjunto perfeito para quem busca unir performance, conforto e uma modelagem "
    "que valoriza a silhueta."
)

DESC_AURA = (
    "O Conjunto Aura une conforto, sofisticacao e uma modelagem desenvolvida para "
    "valorizar suas curvas. Confeccionado em poliamida premium, possui toque "
    "extremamente macio, confortavel e agradavel ao corpo, alem de oferecer ZERO "
    "TRANSPARENCIA, garantindo mais seguranca durante os treinos.\n\n"
    "O top acompanha bojo removivel, proporcionando melhor sustentacao, conforto e "
    "valorizando o busto. Sua modelagem se ajusta perfeitamente ao corpo para "
    "acompanhar seus movimentos.\n\n"
    "A legging possui cintura alta, ajudando a disfarcar a regiao da barriga e modelar "
    "a silhueta. Alem disso, sua modelagem estrategica valoriza e realca o bumbum, "
    "destacando as curvas de forma bonita e natural.\n\n"
    "Conforto, seguranca e uma modelagem que valoriza o seu corpo."
)

NOTA_TAMANHO = (
    "\n\nCOMO ESCOLHER O TAMANHO\n"
    "Cada conjunto sai com a legging e o top no tamanho que voce escolher, e eles podem "
    "ser diferentes. Na opcao de tamanho, o primeiro valor e o da LEGGING e o segundo e "
    "o do TOP. Exemplo: M/G = legging M com top G. Se voce usa o mesmo tamanho nas duas "
    "pecas, escolha P/P, M/M ou G/G. A tabela de medidas esta nas fotos do anuncio."
)

PRODUTOS = [
    {
        "linha": "Shape", "cor": "Matcha", "cor_ml": "Verde",
        "handle": "conjunto-shine-legging-top-matcha",
        "sku_base": "CTS-MATCHA", "desc": DESC_SHAPE.format(cor="Matcha"),
        "bojo": "Bojo fixo",
        "fotos": [
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/1_79b30e4a-aeb6-41b7-8fed-2fdb2171fd34.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/3_1732dfaa-7367-4ad9-aa42-523be0187a90.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/2_55ecdae7-11f5-444e-8209-e77864ad95a4.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/2CE0018D-4C58-4BD1-AF17-2ACB97C88580.png",
        ],
        "estoque": {"P/P": 13, "M/P": 13, "G/P": 13, "P/M": 13, "P/G": 13,
                    "M/M": 13, "M/G": 12, "G/M": 13, "G/G": 13},
    },
    {
        "linha": "Shape", "cor": "Off White", "cor_ml": "Branco",
        "handle": "conjunto-shine-legging-top-off-white",
        "sku_base": "CTS-OFFWHITE", "desc": DESC_SHAPE.format(cor="Off White"),
        "bojo": "Bojo fixo",
        "fotos": [
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/6_da2222ad-a81d-42cd-9b4e-ed2f00b32c3c.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/4_070ccd73-5eb7-4560-b6ba-d54a369ae2d0.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/5_aab050bd-6eac-458c-a8d1-400b3e1894bd.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/2CF040B5-41D0-40B2-8023-04AE68142EC1.png",
        ],
        "estoque": {"P/P": 12, "M/P": 12, "G/P": 12, "P/M": 12, "P/G": 12,
                    "M/M": 12, "M/G": 11, "G/M": 12, "G/G": 12},
    },
    {
        "linha": "Aura", "cor": "Petroleo", "cor_ml": "Azul petroleo",
        "handle": "conjunto-legging-aura-top-petroleo",
        "sku_base": "CTS-AURA", "desc": DESC_AURA, "bojo": "Bojo removivel",
        "fotos": [
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/11_e3552932-1780-4696-8c2c-93416e4fc612.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/12_e35d7013-5d49-41c7-98d1-509bd0ddcc28.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/13_898ab4d8-a1ce-4fe0-afaa-83496462a715.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/14_48939258-2086-4b46-add5-675e4955adb9.png",
        ],
        "estoque": {k: 14 for k in ["P/P", "M/P", "G/P", "P/M", "P/G", "M/M", "M/G", "G/M", "G/G"]},
    },
    {
        "linha": "Aura", "cor": "Solar Orange", "cor_ml": "Laranja",
        "handle": "conjunto-legging-aura-top-laranjaa",
        "sku_base": "CTS-AURA-SOLAR", "desc": DESC_AURA, "bojo": "Bojo removivel",
        "fotos": [
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/4_51508998-643b-41ed-a3d1-688345ebf40f.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/1_862fde68-29ab-423b-85cb-6622cfdcacd3.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/2_93fcfeef-e45b-4c53-8ff7-879e437f8c20.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/3_a752fa1f-6f85-4757-89f3-05006ef13d7c.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/5_a63365ea-a917-43a2-9c75-cf321e6444f9.png",
        ],
        "estoque": {k: 14 for k in ["P/P", "M/P", "G/P", "P/M", "P/G", "M/M", "M/G", "G/M", "G/G"]},
    },
    {
        "linha": "Aura", "cor": "Bronze", "cor_ml": "Bronze",
        "handle": "conjunto-aura-legging-top-bronze",
        "sku_base": "CTS-AURA-BRONZE", "desc": DESC_AURA, "bojo": "Bojo removivel",
        "fotos": [
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/6_8a345ae5-9d03-479b-8104-ad3d564facf3.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/9_84752be2-f1ee-47e4-8788-919b83d40b52.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/7_26b7815b-1529-4c74-b681-6dd89e56e98a.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/10_0653cd3c-7a50-4762-974f-5a55b71b1827.png",
            "https://cdn.shopify.com/s/files/1/0986/2430/7486/files/8_7684cb2f-6b22-4982-a402-0ae092619f47.png",
        ],
        "estoque": {k: 13 for k in ["P/P", "M/P", "G/P", "P/M", "P/G", "M/M", "M/G", "G/M", "G/G"]},
    },
]

# ordem das 9 combinacoes, na mesma sequencia que a loja usa
COMBOS = ["P/P", "M/P", "G/P", "P/M", "P/G", "M/M", "M/G", "G/M", "G/G"]
SUFIXO = {"P": "P", "M": "M", "G": "G"}


def sku_de(base, combo):
    leg, top = combo.split("/")
    return "{}-L{}-T{}".format(base, SUFIXO[leg], SUFIXO[top])


def titulo_de(p):
    t = "Conjunto Fitness Legging Cintura Alta + Top {}".format(p["cor"])
    assert len(t) <= 60, "titulo passou de 60 caracteres: {} ({})".format(t, len(t))
    return t


# ---------------------------------------------------------------- planilha

wb = Workbook()

COLUNAS = [
    ("Anuncio", 10, "ref"),
    ("Titulo do anuncio (max. 60)", 46, "ml"),
    ("Variacao: Cor", 15, "ml"),
    ("Variacao: Tamanho", 17, "ml"),
    ("Tam. legging", 12, "ref"),
    ("Tam. top", 10, "ref"),
    ("SKU", 24, "ml"),
    ("Preco (R$)", 11, "ml"),
    ("Estoque ML", 11, "preencher"),
    ("Estoque no site hoje", 19, "ref"),
    ("Condicao", 10, "ml"),
    ("Codigo universal (EAN)", 21, "preencher"),
    ("Marca", 12, "ml"),
    ("Modelo", 18, "ml"),
    ("Genero", 10, "ml"),
    ("Tipo de roupa", 14, "ml"),
    ("Composicao", 22, "ml"),
    ("Tipo de bojo", 16, "ml"),
    ("E esportivo", 12, "ml"),
    ("Tipo de anuncio", 15, "preencher"),
    ("Forma de envio", 16, "preencher"),
    ("Garantia", 22, "ml"),
    ("Peso da embalagem (g)", 20, "preencher"),
    ("Altura emb. (cm)", 15, "preencher"),
    ("Largura emb. (cm)", 16, "preencher"),
    ("Comprimento emb. (cm)", 20, "preencher"),
    ("Fotos (URLs separadas por virgula)", 60, "ml"),
    ("Descricao", 80, "ml"),
]

ws = wb.active
ws.title = "Anuncios ML"

for i, (nome, larg, _) in enumerate(COLUNAS, start=1):
    c = ws.cell(row=1, column=i, value=nome)
    c.font = Font(name=FONTE, bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", fgColor=AZUL)
    c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.column_dimensions[get_column_letter(i)].width = larg
ws.row_dimensions[1].height = 34
ws.freeze_panes = "C2"

linha = 2
for n, p in enumerate(PRODUTOS, start=1):
    titulo = titulo_de(p)
    fotos = ", ".join(p["fotos"])
    desc = p["desc"] + NOTA_TAMANHO
    modelo = "Conjunto {} {}".format(p["linha"], p["cor"])
    for combo in COMBOS:
        leg, top = combo.split("/")
        valores = [
            "#{}".format(n), titulo, p["cor_ml"], combo, leg, top,
            sku_de(p["sku_base"], combo), 189.00, p["estoque"][combo],
            p["estoque"][combo], "Novo", "", "Contt.s", modelo, "Feminino",
            "Conjunto", "Poliamida com elastano", p["bojo"], "Sim", "", "",
            "Garantia do vendedor: 30 dias", "", "", "", "", fotos, desc,
        ]
        for i, v in enumerate(valores, start=1):
            c = ws.cell(row=linha, column=i, value=v)
            c.font = Font(name=FONTE, size=10)
            c.alignment = Alignment(vertical="top", wrap_text=(i in (2, 27, 28)))
            c.border = BORDA
            tipo = COLUNAS[i - 1][2]
            if tipo == "preencher":
                c.fill = PatternFill("solid", fgColor=AMARELO)
            elif tipo == "ref":
                c.fill = PatternFill("solid", fgColor=CINZA)
        ws.cell(row=linha, column=8).number_format = '#,##0.00'
        ws.row_dimensions[linha].height = 15
        linha += 1

ultima = linha - 1

# ---------------------------------------------------------------- como usar

wi = wb.create_sheet("Como usar")
wi.column_dimensions["A"].width = 4
wi.column_dimensions["B"].width = 112

def bloco(titulo, paragrafos):
    global lin
    c = wi.cell(row=lin, column=2, value=titulo)
    c.font = Font(name=FONTE, bold=True, size=12, color=AZUL)
    lin += 1
    for t in paragrafos:
        c = wi.cell(row=lin, column=2, value=t)
        c.font = Font(name=FONTE, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        wi.row_dimensions[lin].height = 15 * (1 + len(t) // 105)
        lin += 1
    lin += 1

lin = 2
c = wi.cell(row=lin, column=2, value="Conjuntos Shape e Aura no Mercado Livre")
c.font = Font(name=FONTE, bold=True, size=15, color=AZUL)
lin += 1
c = wi.cell(row=lin, column=2, value="5 anuncios, 9 tamanhos cada, 45 linhas. Dados da loja em 14/09/2026.")
c.font = Font(name=FONTE, size=10, italic=True, color="7F7F7F")
lin += 3

bloco("As cores das celulas", [
    "BRANCO - pronto, veio da loja. Nao precisa mexer.",
    "AMARELO - falta preencher. Sao dados que nao existem na Shopify: peso e medidas da caixa, "
    "codigo de barras, tipo de anuncio e forma de envio.",
    "CINZA - so para sua conferencia. Nao vai para o Mercado Livre.",
])

bloco("Os dois tamanhos no mesmo anuncio", [
    "O Mercado Livre so aceita UM campo de tamanho por variacao, e o conjunto tem dois: o da "
    "legging e o do top. Por isso a coluna Tamanho vem no formato P/G, onde o primeiro e o da "
    "legging e o segundo e o do top.",
    "A explicacao foi escrita no fim de toda descricao, para a cliente nao errar na compra.",
    "Motivo de manter as 9 combinacoes: o unico pedido de conjunto que a loja teve ate hoje "
    "(pedido #1084, Marcia) foi justamente tamanho misturado, legging M com top G. Publicar so "
    "P/P, M/M e G/G deixaria de fora exatamente o que ja vendeu.",
    "Se o Mercado Livre recusar esses valores na sua categoria, o plano B e publicar tres "
    "tamanhos (P, M e G, conjunto inteiro no mesmo tamanho) e tratar a mistura por mensagem.",
])

bloco("Estoque - cuidado aqui", [
    "A coluna Estoque ML veio com o numero real que esta na loja hoje. Se voce subir esse numero "
    "cheio, a mesma peca fica anunciada em dois lugares ao mesmo tempo e os dois podem vender.",
    "Duas saidas: separar uma parte do estoque so para o Mercado Livre (por exemplo, 5 de cada "
    "tamanho) e baixar o resto na Shopify, ou usar um integrador que sincroniza os dois.",
    "Enquanto nao houver integracao, quem vender no Mercado Livre precisa ser descontado na "
    "Shopify na mao, no mesmo dia.",
])

bloco("Preco - decisao sua", [
    "Coloquei R$ 189,00, o mesmo preco do site.",
    "O Mercado Livre cobra comissao por venda e, acima de um valor minimo, o frete gratis sai do "
    "seu bolso. Nos R$ 189,00 isso pesa. Vale rodar a conta na sua propria calculadora do Mercado "
    "Livre antes de publicar, e decidir se o preco de la sobe.",
    "Nao consegui consultar as taxas atuais daqui: o acesso a api.mercadolibre.com esta bloqueado "
    "pela politica da organizacao. Por isso nao chutei nenhum valor.",
])

bloco("Fotos", [
    "As URLs sao as fotos que ja estao no ar no site. O Mercado Livre baixa as imagens direto "
    "desses enderecos, entao nao precisa subir arquivo nenhum.",
    "A ultima foto de cada conjunto e a tabela de medidas. Deixe ela no anuncio, ajuda a reduzir "
    "troca por tamanho errado.",
])

bloco("Como jogar isso na planilha do Mercado Livre", [
    "1. No Mercado Livre: Anuncios > Publicar em massa. Escolha a categoria de conjunto fitness "
    "feminino e baixe a planilha modelo deles.",
    "2. A planilha modelo tem os nomes de coluna proprios do Mercado Livre e muda conforme a "
    "categoria. Copie daqui coluna por coluna, olhando o nome de la.",
    "3. Preencha o que esta em amarelo.",
    "4. Suba a planilha e confira o relatorio de erros antes de publicar.",
    "Nao consegui gerar o arquivo ja no formato exato deles porque o modelo so sai logado na sua "
    "conta e e especifico da categoria que voce escolher.",
])

wb.move_sheet("Como usar", offset=-1)

wb.save("/home/user/contt.s/mercadolivre-conjuntos.xlsx")
print("ok - {} linhas de variacao".format(ultima - 1))

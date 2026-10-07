"""Avisa no WhatsApp quando alguma variante dos Conjuntos R$ 189 fica sem estoque.

Roda no GitHub Actions (.github/workflows/alerta-estoque.yml). Só manda mensagem
quando a situação muda: uma variante que acabou de zerar (ou baixou do limite) ou
que voltou a ter estoque. O estado da última verificação fica em ESTADO_ARQUIVO.
"""
import json
import os
import sys
import urllib.parse
import urllib.request

LOJA = os.environ.get("SHOPIFY_LOJA", "0qyay7-vt.myshopify.com")
TOKEN = os.environ["SHOPIFY_TOKEN"]
TAG = os.environ.get("TAG_PRODUTOS", "Conjunto 189")
LIMITE = int(os.environ.get("LIMITE_ESTOQUE", "0"))
WHATSAPP_NUMERO = os.environ["WHATSAPP_NUMERO"]
WHATSAPP_APIKEY = os.environ["WHATSAPP_APIKEY"]
ESTADO_ARQUIVO = os.environ.get("ESTADO_ARQUIVO", "estado-estoque.json")

CONSULTA = """
query($q: String!, $after: String) {
  products(first: 50, query: $q, after: $after) {
    pageInfo { hasNextPage endCursor }
    edges { node {
      title status
      variants(first: 100) { edges { node { id title sku inventoryQuantity } } }
    } }
  }
}
"""


def shopify(variaveis):
    req = urllib.request.Request(
        f"https://{LOJA}/admin/api/2025-10/graphql.json",
        data=json.dumps({"query": CONSULTA, "variables": variaveis}).encode(),
        headers={"Content-Type": "application/json", "X-Shopify-Access-Token": TOKEN},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        resp = json.load(r)
    if resp.get("errors"):
        sys.exit(f"Erro da Shopify: {resp['errors']}")
    return resp["data"]["products"]


def variantes():
    after = None
    while True:
        pagina = shopify({"q": f"tag:'{TAG}' AND status:active", "after": after})
        for p in pagina["edges"]:
            for v in p["node"]["variants"]["edges"]:
                v = v["node"]
                yield v["id"], f"{p['node']['title']} ({v['title']})", v["inventoryQuantity"] or 0
        if not pagina["pageInfo"]["hasNextPage"]:
            return
        after = pagina["pageInfo"]["endCursor"]


def whatsapp(texto):
    url = "https://api.callmebot.com/whatsapp.php?" + urllib.parse.urlencode(
        {"phone": WHATSAPP_NUMERO, "text": texto, "apikey": WHATSAPP_APIKEY}
    )
    with urllib.request.urlopen(url, timeout=30) as r:
        r.read()


def main():
    try:
        with open(ESTADO_ARQUIVO) as f:
            avisados = set(json.load(f))
    except (FileNotFoundError, ValueError):
        avisados = set()

    baixos = {vid: (nome, qtd) for vid, nome, qtd in variantes() if qtd <= LIMITE}
    novos = [baixos[v] for v in baixos if v not in avisados]
    repostos = len(avisados - baixos.keys())

    if novos:
        titulo = "acabou o estoque" if LIMITE == 0 else f"estoque em {LIMITE} ou menos"
        linhas = [f"- {nome}: {qtd}" for nome, qtd in sorted(novos)]
        whatsapp(f"CONTT.s - {titulo} nos Conjuntos R$ 189:\n" + "\n".join(linhas))
    if repostos:
        whatsapp(f"CONTT.s - {repostos} variante(s) dos Conjuntos R$ 189 voltaram a ter estoque.")

    with open(ESTADO_ARQUIVO, "w") as f:
        json.dump(sorted(baixos), f)
    print(f"{len(baixos)} variante(s) no limite; {len(novos)} aviso(s) novo(s); {repostos} reposta(s).")


if __name__ == "__main__":
    main()

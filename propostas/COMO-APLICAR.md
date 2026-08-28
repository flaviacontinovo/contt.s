# Aviso temporário de envio — congresso (28/08 → 10/09)

Alteração pronta, **ainda não enviada** ao repositório do tema
(`flaviacontinovo/contts-shopify-theme`). São três passos.

## 1. Novo arquivo

`snippets/aviso-envio.liquid` — cópia em `propostas/aviso-envio.liquid`.

## 2. Chamada na página de produto

Em `sections/main-product.liquid`, imediatamente acima de `<div class="freight">`
(ou seja, depois do botão de WhatsApp e antes do "frete e prazo"):

```liquid
      {% render 'aviso-envio' %}
```

## 3. Estilo

Em `assets/site.css`, na seção `produto.html`, logo antes de `.freight{margin-bottom:26px;}`:

```css
  /* aviso temporário de envio (congresso) — remover junto com snippets/aviso-envio.liquid */
  .aviso-envio{display:flex;align-items:flex-start;gap:10px;background:#faf3e2;border:1px solid #e3cf9f;border-radius:9px;padding:11px 13px;margin:0 0 18px;}
  .aviso-envio__ic{flex:0 0 auto;color:#8a6d2e;margin-top:1px;}
  .aviso-envio__ic svg{width:18px;height:18px;display:block;}
  .aviso-envio__txt{display:flex;flex-direction:column;gap:3px;min-width:0;}
  .aviso-envio__txt>b{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;color:#8a6d2e;}
  .aviso-envio__txt small{font-size:12px;line-height:1.45;color:#8a6d2e;opacity:.9;}
  .aviso-envio__txt small b{font-weight:700;opacity:1;}
```

As cores são as mesmas que o tema já usa no aviso de retirada
(`.freight__opt--pickup`), para não destoar.

## Prorrogar

Trocar as duas datas no topo de `snippets/aviso-envio.liquid`. Só lá.

## Remover depois do congresso

1. Apagar a linha `{% render 'aviso-envio' %}` de `sections/main-product.liquid`.
2. Apagar `snippets/aviso-envio.liquid`.
3. Apagar o bloco `.aviso-envio` do `assets/site.css`.

## Atenção ao publicar

O repositório do tema é conectado à Shopify: o que entra na branch `main`
**vai ao ar na loja automaticamente**. Uma branch separada não publica nada.

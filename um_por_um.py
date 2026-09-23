# -*- coding: utf-8 -*-
"""URL 1:1 da Shopify: recorte quadrado pelo centro, conferido contra o
que a propria Shopify devolve no campo image(transform:).
"""

def um_por_um(url, lado=1200):
    base, _, query = url.partition('?')
    raiz, ponto, ext = base.rpartition('.')
    assert ponto and ext, 'URL sem extensao: %s' % url
    novo = '%s_%dx%d_crop_center.%s' % (raiz, lado, lado, ext)
    return novo + ('?' + query if query else '')

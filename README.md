# contt.s

Site institucional da empresa — página estática, sem dependências ou etapa de build.

## Estrutura

- `index.html` — página única do site (HTML, CSS embutido e favicon inline)

## Como visualizar

Abra o arquivo diretamente no navegador:

```bash
open index.html      # macOS
xdg-open index.html  # Linux
```

Ou sirva localmente, se preferir testar por HTTP:

```bash
python3 -m http.server 8000
# acesse http://localhost:8000
```

## Como editar

Todo o conteúdo e o estilo estão em `index.html`. O texto visível fica dentro
de `<main>`; as cores ficam nas variáveis CSS em `:root` (com variação
automática para modo escuro via `prefers-color-scheme`).

## Status

A página está com aviso de **"em construção"**. O conteúdo definitivo
(descrição da empresa, serviços, contato) ainda precisa ser adicionado.

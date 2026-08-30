# Dashboard estático — Controle de Calibração de Instrumentos

Este é um snapshot estático gerado a partir da planilha
"Controle_Calibracao_Extreme_Wind_2026_FINAL.xlsx" (Dropbox → 3 - SUPRIMENTOS
/ 01 - ALMOXARIFADO / 04 - ETIQUETAS / ETIQUETAS CALIBRAÇÃO). Isso é um
snapshot, não um painel ao vivo — os dados só atualizam quando alguém roda
`generate_data.py` de novo e faz commit/push do `data.json` resultante.

Se você precisa de um painel que busca a planilha do Dropbox automaticamente
toda vez que é aberto, esse é o painel ao vivo do Cowork (fora do escopo
deste repositório).

## Arquivos

- `index.html` — o dashboard em si (HTML/CSS/JS puro, usa Chart.js e Grid.js
  via CDN). Lê `data.json` via `fetch()`; se isso falhar (ex.: aberto direto
  como `file://` sem servidor), usa o bloco `embedded-data` já embutido no
  próprio HTML como fallback.
- `data.json` — os dados processados (gerado por `generate_data.py`).
- `generate_data.py` — script que lê `raw_text.txt` e gera `data.json` (além
  de embutir os dados dentro do `index.html`).
- `raw_text.txt` — o texto bruto colado da planilha, usado como entrada do
  script. Precisa ser atualizado manualmente a cada rodada.

## Como atualizar

1. Copie o texto bruto atualizado da aba "Controle de Calibração" da
   planilha para `raw_text.txt`.
2. Rode `python3 generate_data.py`.
3. Isso atualiza `data.json` e embute os dados no `index.html`.
4. Faça commit/push de `data.json` (e `index.html`, se mudou) no
   repositório do GitHub.

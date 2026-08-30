#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera data.json para o dashboard estático (index.html) a partir do texto
exportado da planilha "Controle_Calibracao_Extreme_Wind_2026_FINAL.xlsx".

Como atualizar o dashboard:
1. Abra a planilha no Dropbox (ou peça pro Claude buscar o texto via
   conector do Dropbox) e copie o conteúdo bruto de texto (aba "Controle de
   Calibração") para um arquivo chamado raw_text.txt, na mesma pasta deste
   script.
2. Rode:  python3 generate_data.py
3. Isso gera/atualiza data.json.
4. Faça commit + push de data.json no repositório do GitHub. O index.html
   lê data.json automaticamente na próxima visita/deploy do GitHub Pages.
"""
import json
import re
import sys
import datetime

RAW_TEXT_PATH = "raw_text.txt"
OUTPUT_PATH = "data.json"
INDEX_HTML_PATH = "index.html"


def us_short_date_to_br(s):
    # Linhas "curtas" (sem Proposta/Arquivo de Origem/Observações) costumam
    # vir com datas em formato americano M/D/AA sem zero à esquerda
    # (ex.: "2/19/25" só pode ser 19/fev/2025, já que dia 19 não existe como
    # mês). O restante da planilha usa DD/MM/AAAA. Convertemos aqui só a
    # EXIBIÇÃO — o "Dias para Vencer"/Status já vêm calculados corretamente
    # pela própria planilha a partir da data real.
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", s.strip())
    if not m:
        return s
    mm, dd, yy = m.groups()
    if len(yy) == 2:
        yy = "20" + yy
    return f"{dd.zfill(2)}/{mm.zfill(2)}/{yy}"


def status_class(s):
    if s == "OK":
        return "OK"
    if s.startswith("VENCE EM"):
        return "VENCE"
    if s.startswith("VENCIDO"):
        return "VENCIDO"
    return "SEM"


def to_int(v):
    try:
        return int(v)
    except Exception:
        return None


def main():
    try:
        with open(RAW_TEXT_PATH, "r", encoding="utf-8") as f:
            raw = f.read()
    except FileNotFoundError:
        print(f"ERRO: não encontrei {RAW_TEXT_PATH}. Cole o texto da planilha nesse arquivo antes de rodar o script.")
        sys.exit(1)

    lines = raw.split("\n")
    rows = []
    in_table = False
    for line in lines:
        if line.startswith("\tCódigo\tStatus"):
            in_table = True
            continue
        if in_table:
            if not line.startswith("\t"):
                in_table = False
                continue
            parts = line.split("\t")
            # "Dias para Vencer" é sempre o ÚLTIMO campo da linha (não um
            # índice fixo), pois algumas linhas não têm Proposta/Arquivo de
            # Origem/Observações — usar índice fixo descarta essas linhas
            # silenciosamente ou lê o campo errado.
            if len(parts) < 6:
                continue
            is_short_row = len(parts) <= 7
            data_cal = parts[3].strip()
            prox_cal = parts[4].strip()
            if is_short_row:
                data_cal = us_short_date_to_br(data_cal)
                prox_cal = us_short_date_to_br(prox_cal)
            rows.append({
                "codigo": parts[1].strip(),
                "status": parts[2].strip(),
                "dataCalibracao": data_cal,
                "proximaCalibracao": prox_cal,
                "diasRaw": parts[-1].strip(),
            })

    if not rows:
        print("ERRO: nenhuma linha de instrumento foi reconhecida em raw_text.txt. Verifique o conteúdo colado.")
        sys.exit(1)

    counts = {"OK": 0, "VENCE": 0, "VENCIDO": 0, "SEM": 0}
    for r in rows:
        counts[status_class(r["status"])] += 1
    due_soon = sum(
        1 for r in rows
        if (to_int(r["diasRaw"]) is not None) and 0 <= to_int(r["diasRaw"]) <= 90
    )

    now = datetime.datetime.now()
    generated_at = now.strftime("%d/%m/%Y às %H:%M")

    data = {
        "generated_at": generated_at,
        "total": len(rows),
        "counts": counts,
        "due_soon": due_soon,
        "rows": rows,
    }

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # Também embute os dados dentro do index.html (entre os marcadores
    # EMBEDDED_DATA_START/END). Isso serve de fallback para quando o arquivo
    # é aberto direto no navegador (file://), onde fetch("./data.json")
    # é bloqueado por segurança. Quando hospedado via HTTP normal, o
    # index.html sempre tenta o data.json primeiro (mais atual) e só usa
    # esse embutido se o fetch falhar.
    try:
        with open(INDEX_HTML_PATH, "r", encoding="utf-8") as f:
            html = f.read()
        # Escapa "</" para não fechar a tag <script> prematuramente se algum
        # texto (ex.: observação) contiver algo como "</script".
        embedded_json = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
        new_html, n = re.subn(
            r'(<script id="embedded-data" type="application/json">).*?(</script>)',
            lambda m: m.group(1) + embedded_json + m.group(2),
            html,
            flags=re.S,
        )
        if n == 0:
            print(f"AVISO: não encontrei o bloco embedded-data em {INDEX_HTML_PATH}; "
                  f"pulei essa etapa (data.json foi gerado normalmente).")
        else:
            with open(INDEX_HTML_PATH, "w", encoding="utf-8") as f:
                f.write(new_html)
    except FileNotFoundError:
        print(f"AVISO: {INDEX_HTML_PATH} não encontrado nesta pasta; pulei a etapa de embutir dados.")

    print(f"OK: {OUTPUT_PATH} gerado com {len(rows)} instrumentos "
          f"(VENCIDO={counts['VENCIDO']}, VENCE EM 90 DIAS={counts['VENCE']}, "
          f"OK={counts['OK']}, SEM DADOS={counts['SEM']}).")


if __name__ == "__main__":
    main()

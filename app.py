def parse_distribuicao(arquivo_excel, ano):
    arquivo_excel.seek(0)
    df = pd.read_excel(arquivo_excel, header=None, dtype=object)
    etapas = {}
    etapa_num = 0
    etapa_ativa = None

    for idx, row in df.iterrows():
        cells = [normalizar_texto(c) for c in row.tolist()]
        if not any(cells):
            continue

        # marcador explícito de etapa
        etapa_explicita = None
        for c in cells:
            m = re.search(r"\b([123])\s*[ªº°]?\s*etapa\b", c, re.IGNORECASE)
            if m:
                etapa_explicita = int(m.group(1))
                break
        if etapa_explicita is not None:
            etapa_num = etapa_explicita

        if _eh_linha_total(cells):
            # captura os totais da linha (colunas C, D, E, F, G)
            totais = {}
            for dia, col in [("2ª Feira", 2), ("3ª Feira", 3),
                             ("4ª Feira", 4), ("5ª Feira", 5),
                             ("6ª Feira", 6)]:
                if col < len(row):
                    v = row.iloc[col]
                    if pd.notna(v):
                        totais[dia] = int(v)
            etapas[f"{etapa_num}ª Etapa"]["_totais"] = totais
            etapa_num += 1
            continue

        if etapa_ativa is None:
            if _eh_cabecalho(cells):
                etapa_num = 1
            else:
                continue

        etapa_ativa = f"{etapa_num}ª Etapa"
        etapas.setdefault(etapa_ativa, {d: [] for d in DIAS_SEMANA})
        etapas[etapa_ativa].setdefault("_totais", {})

        # ... (restante do código de datas permanece igual)
        # mas NÃO usaremos as listas de datas para contagem, apenas para exibição

# ============================================================
# PARSER DO CALENDÁRIO ESCOLAR
# ============================================================
def parse_calendario(arquivo_pdf, ano=2026):
    nao_letivos = set()
    with pdfplumber.open(arquivo_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ''
            for linha in texto.split('\n'):
                linha = linha.strip().lower()
                # Detecta feriados, recesso, planejamento
                if any(palavra in linha for palavra in ['feriado','recesso','planejamento']):
                    # Extrai datas da linha
                    datas = _extrair_datas_celula(linha, None, ano)
                    nao_letivos.update(datas)
    return nao_letivos


# ============================================================
# CONFERÊNCIA COM CALENDÁRIO
# ============================================================
def conferir(etapas, etapa, aulas_por_dia, df_diario, nao_letivos):
    dias_esperados = []
    for dia_sem in DIAS_SEMANA:
        qtd = int(aulas_por_dia.get(dia_sem, 0))
        if qtd <= 0:
            continue
        for dt in etapas[etapa][dia_sem]:
            # Ignora dias não letivos
            if dt in nao_letivos:
                continue
            dias_esperados.append({'data': dt, 'dia_semana': dia_sem, 'aulas_esperadas': qtd})

    df_esp = pd.DataFrame(dias_esperados)
    if df_esp.empty:
        return None

    df_esp['data'] = pd.to_datetime(df_esp['data']).dt.date
    df_diario = df_diario.copy()
    df_diario['data'] = pd.to_datetime(df_diario['data']).dt.date

    total_esperado = int(df_esp['aulas_esperadas'].sum())
    total_lancado = int(df_diario['aulas'].sum())

    faltantes = df_esp[~df_esp['data'].isin(df_diario['data'])].copy()
    extras = df_diario[~df_diario['data'].isin(df_esp['data'])].copy()
    merged = df_esp.merge(df_diario[['data','aulas']], on='data', how='outer', indicator=True)
    divergentes = merged[(merged['_merge']=='both') & (merged['aulas_esperadas']!=merged['aulas'])].copy()

    return {
        'total_esperado': total_esperado,
        'total_lancado': total_lancado,
        'faltantes': faltantes.sort_values('data'),
        'extras': extras.sort_values('data'),
        'divergentes': divergentes.sort_values('data'),
    }

import streamlit as st
import pandas as pd
import pdfplumber
import re
from datetime import datetime, date

# ============================================================
# CONFIGURAÇÕES
# ============================================================

st.set_page_config(
    page_title="Conferência de Aulas — CTPM",
    page_icon="📚",
    layout="wide"
)

MESES = {
    'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4,
    'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
    'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12
}
MESES_ABREV = {
    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
}
DIAS_SEMANA = ['2ª Feira', '3ª Feira', '4ª Feira', '5ª Feira', '6ª Feira', 'Sábado']

# ============================================================
# EXTRAÇÃO DE DATAS (PLANILHA)
# ============================================================

def _extrair_datas_celula(valor, mes_num=None, ano=2026):
    if valor is None:
        return []
    if isinstance(valor, datetime):
        return [valor.date()]
    if isinstance(valor, date):
        return [valor]
    texto = str(valor).strip()
    if not texto or texto.upper() in ('NAN', 'NONE', 'NAT'):
        return []
    numeros = re.findall(r'\d+', texto)
    datas = []
    for n in numeros:
        try:
            datas.append(date(ano, mes_num, int(n)))
        except ValueError:
            continue
    return datas

# ============================================================
# PARSER — DISTRIBUIÇÃO DE DIAS LETIVOS
# ============================================================

def parse_distribuicao(arquivo_excel, ano=2026):
    df = pd.read_excel(arquivo_excel, header=None, dtype=object)
    etapas = {}
    etapa_atual = None
    col_dias = {}

    for _, row in df.iterrows():
        cells = [str(c).strip() if pd.notna(c) else '' for c in row]
        if not any(cells):
            continue

        # Detecta cabeçalho de dias da semana
        if any('Feira' in c for c in cells):
            novo = {}
            for i, c in enumerate(cells):
                m = re.match(r'^(\d)\s*[ºª]?\s*Feira', c, re.IGNORECASE)
                if m and m.group(1) in '23456':
                    novo[f'{m.group(1)}ª Feira'] = i
            if novo:
                col_dias = novo
            continue

        # Detecta etapa
        etapa_encontrada = None
        for c in cells:
            m = re.match(r'^\s*(\d)\s*[ªA]?\s*ETAPA\s*$', c, re.IGNORECASE)
            if m:
                etapa_encontrada = f"{m.group(1)}ª ETAPA"
                break

        if etapa_encontrada:
            etapa_atual = etapa_encontrada
            if etapa_atual not in etapas:
                etapas[etapa_atual] = {d: [] for d in DIAS_SEMANA}
            continue

        if etapa_atual is None:
            continue

        # Detecta mês
        mes_encontrado = None
        for c in cells:
            if c in MESES:
                mes_encontrado = c
                break

        if mes_encontrado:
            mes_num = MESES[mes_encontrado]
            for dia_sem, col_idx in col_dias.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if pd.notna(val) and str(val).strip() != '':
                        datas = _extrair_datas_celula(val, mes_num, ano)
                        etapas[etapa_atual][dia_sem].extend(datas)
            continue

    for et in etapas:
        for d in etapas[et]:
            etapas[et][d] = sorted(set(etapas[et][d]))
    return etapas

# ============================================================
# PARSER DO DIÁRIO
# ============================================================

def parse_diario_pdf(arquivo_pdf, ano_padrao=2026):
    registros = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for page in pdf.pages:
            for tabela in page.extract_tables():
                if not tabela or len(tabela) < 2:
                    continue
                for linha in tabela[1:]:
                    if not linha:
                        continue
                    try:
                        data_str = str(linha[0]).strip()
                        aulas = int(str(linha[1]).strip())
                        dia, mes = map(int, data_str.split('/'))
                        dt = date(ano_padrao, mes, dia)
                        registros.append({'data': dt, 'aulas': aulas})
                    except:
                        continue
    return pd.DataFrame(registros)

# ============================================================
# CONFERÊNCIA
# ============================================================

def conferir(etapas, etapa, aulas_por_dia, df_diario):
    dias_esperados = []
    for dia_sem in DIAS_SEMANA:
        qtd = int(aulas_por_dia.get(dia_sem, 0))
        if qtd <= 0:
            continue
        for dt in etapas[etapa][dia_sem]:
            dias_esperados.append({'data': dt, 'dia_semana': dia_sem, 'aulas_esperadas': qtd})

    df_esp = pd.DataFrame(dias_esperados)
    df_esp['data'] = pd.to_datetime(df_esp['data']).dt.date
    df_diario['data'] = pd.to_datetime(df_diario['data']).dt.date

    total_esperado = int(df_esp['aulas_esperadas'].sum())
    total_lancado = int(df_diario['aulas'].sum())

    faltantes = df_esp[~df_esp['data'].isin(df_diario['data'])]
    extras = df_diario[~df_diario['data'].isin(df_esp['data'])]

    divergentes = df_esp.merge(df_diario, on='data')
    divergentes = divergentes[divergentes['aulas_esperadas'] != divergentes['aulas']]

    return {
        'total_esperado': total_esperado,
        'total_lancado': total_lancado,
        'faltantes': faltantes,
        'extras': extras,
        'divergentes': divergentes,
    }

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
# EXTRAÇÃO DE DATAS
# ============================================================
def _extrair_datas_celula(valor, mes_num=None, ano=2026):
    """Extrai lista de datas de uma célula."""
    if valor is None:
        return []
    if isinstance(valor, datetime):
        return [valor.date()]
    if isinstance(valor, date):
        return [valor]

    texto = str(valor).strip()
    if not texto or texto.upper() in ('NAN', 'NONE', 'NAT'):
        return []

    # 1) ISO: "2026-03-07 00:00:00"
    if re.match(r'^\d{4}-\d{2}-\d{2}', texto):
        try:
            return [datetime.strptime(texto[:10], '%Y-%m-%d').date()]
        except ValueError:
            pass

    # 2) dd/mm/aaaa ou dd-mm-aaaa
    m = re.match(r'^(\d{1,2})[/\.\-](\d{1,2})[/\.\-](\d{2,4})', texto)
    if m:
        dia, mes, ano_str = int(m.group(1)), int(m.group(2)), m.group(3)
        ano_conv = int(ano_str) if len(ano_str) == 4 else 2000 + int(ano_str)
        try:
            return [date(ano_conv, mes, dia)]
        except ValueError:
            return []

    # 3) dd/mmm (abreviado): "07/mar", "30/mai", "26/set"
    m = re.match(r'^(\d{1,2})\s*[/\-\.]\s*([a-zA-ZçÇ]{3})', texto)
    if m:
        dia = int(m.group(1))
        mes_ab = m.group(2).lower()[:3]
        if mes_ab in MESES_ABREV:
            try:
                return [date(ano, MESES_ABREV[mes_ab], dia)]
            except ValueError:
                return []

    # 4) Dias soltos: "9.23", "04, 11,25"
    if mes_num is None:
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
    """Lê a planilha e retorna {etapa: {dia_semana: [datas]}}."""
    df = pd.read_excel(arquivo_excel, header=None, dtype=object)

    etapas = {}
    etapa_atual = None
    col_dias = {}

    for _, row in df.iterrows():
        cells = [str(c).strip() if pd.notna(c) else '' for c in row]
        if not any(cells):
            continue

        # ---- Cabeçalho de dias da semana ----
        if any('Feira' in c for c in cells):
            novo = {}
            for i, c in enumerate(cells):
                m = re.match(r'^(\d)\s*[ºª]?\s*Feira', c, re.IGNORECASE)
                if m:
                    num = m.group(1)
                    if num in '23456':
                        novo[f'{num}ª Feira'] = i
            if novo:
                col_dias = novo
            continue

        # ---- Etapa: "1ª ETAPA", "2ª Etapa" etc ----
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

        # ---- Mês ----
        mes_encontrado = None
        for c in cells:
            if c in MESES:
                mes_encontrado = c
                break

        if mes_encontrado:
            mes_num = MESES[mes_encontrado]
            # Se por algum motivo col_dias ainda não foi detectado, usa padrão C..G
            if not col_dias:
                col_dias = {'2ª Feira': 2, '3ª Feira': 3, '4ª Feira': 4,
                            '5ª Feira': 5, '6ª Feira': 6}
            for dia_sem, col_idx in col_dias.items():
                if col_idx < len(row):
                    val = row[col_idx]
                    if pd.notna(val) and str(val).strip() != '':
                        datas = _extrair_datas_celula(val, mes_num, ano)
                        etapas[etapa_atual][dia_sem].extend(datas)
            continue

        # ---- Sábado ----
        for i, c in enumerate(cells):
            if c in ('Sábado', 'Sabado'):
                for j in range(i + 1, len(row)):
                    val = row[j]
                    if pd.notna(val) and str(val).strip() != '':
                        datas = _extrair_datas_celula(val, None, ano)
                        etapas[etapa_atual]['Sábado'].extend(datas)
                break

    # Remove duplicatas e ordena
    for et in etapas:
        for d in etapas[et]:
            etapas[et][d] = sorted(set(etapas[et][d]))

    return etapas


# ============================================================
# PARSER — DIÁRIO EM PDF
# ============================================================
def parse_diario_pdf(arquivo_pdf, ano_padrao=2026):
    registros = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for page in pdf.pages:
            for tabela in page.extract_tables():
                if not tabela or len(tabela) < 2:
                    continue

                col_dia = col_aulas = col_conteudo = None
                for header_row in tabela[:3]:
                    for i, c in enumerate(header_row):
                        if c is None:
                            continue
                        h = str(c).strip().upper()
                        if 'DIA/MÊS' in h or 'DIA/MES' in h or h == 'MÊS':
                            col_dia = i
                        elif h.startswith('AULAS') or h == 'AULAS':
                            col_aulas = i
                        elif 'SÍNTESE' in h or 'SINTESE' in h:
                            col_conteudo = i

                if col_dia is None or col_aulas is None:
                    continue

                for linha in tabela[1:]:
                    if not linha or len(linha) <= max(col_dia, col_aulas):
                        continue
                    v_dia, v_aulas = linha[col_dia], linha[col_aulas]
                    if v_dia is None or v_aulas is None:
                        continue
                    texto_dia = str(v_dia).strip()
                    m = re.match(r'^(\d{1,2})[/\.\-](\d{1,2})(?:[/\.\-](\d{2,4}))?', texto_dia)
                    if not m:
                        continue
                    dia, mes = int(m.group(1)), int(m.group(2))
                    ano_str = m.group(3)
                    ano = int(ano_str) if ano_str and len(ano_str) == 4 else (
                        2000 + int(ano_str) if ano_str else ano_padrao
                    )
                    try:
                        dt = date(ano, mes, dia)
                    except ValueError:
                        continue
                    digitos = re.sub(r'\D', '', str(v_aulas))
                    if not digitos:
                        continue
                    conteudo = ''
                    if col_conteudo is not None and col_conteudo < len(linha):
                        conteudo = str(linha[col_conteudo] or '').replace('\n', ' ').strip()
                    registros.append({'data': dt, 'aulas': int(digitos), 'conteudo': conteudo})

    df = pd.DataFrame(registros)
    if not df.empty:
        df = df.groupby('data', as_index=False).agg({'aulas': 'max', 'conteudo': 'first'})
    return df


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
    if df_esp.empty:
        return None

    df_esp['data'] = pd.to_datetime(df_esp['data']).dt.date
    if df_diario.empty:
        df_diario = pd.DataFrame(columns=['data', 'aulas', 'conteudo'])
    else:
        df_diario = df_diario.copy()
        df_diario['data'] = pd.to_datetime(df_diario['data']).dt.date

    total_esperado = int(df_esp['aulas_esperadas'].sum())
    total_lancado = int(df_diario['aulas'].sum()) if not df_diario.empty else 0

    faltantes = df_esp[~df_esp['data'].isin(df_diario['data'])].copy()
    extras = df_diario[~df_diario['data'].isin(df_esp['data'])].copy()

    merged = df_esp.merge(df_diario[['data', 'aulas']], on='data', how='outer',
                          indicator=True, suffixes=('_esp', '_lanc'))
    divergentes = merged[(merged['_merge'] == 'both') &
                         (merged['aulas_esperadas'] != merged['aulas'])].copy()

    pasta1 = []
    for dia_sem in DIAS_SEMANA:
        total_dias = len(etapas[etapa][dia_sem])
        qtd = int(aulas_por_dia.get(dia_sem, 0))
        pasta1.append({
            'Dias da semana': dia_sem,
            'Nº de aulas no dia': qtd,
            'Total de dias da semana na etapa': total_dias,
            'Total de aulas': total_dias * qtd
        })
    df_pasta1 = pd.DataFrame(pasta1)

    return {
        'df_pasta1': df_pasta1,
        'total_esperado': total_esperado,
        'total_lancado': total_lancado,
        'faltantes': faltantes.sort_values('data'),
        'extras': extras.sort_values('data'),
        'divergentes': divergentes,
    }


# ============================================================
# INTERFACE
# ============================================================
st.title("📚 Sistema de Conferência de Aulas — CTPM Lavras")
st.caption("Faça o upload dos arquivos, informe as aulas por dia e clique em RODAR.")

with st.sidebar:
    st.header("📂 Arquivos")
    arquivo_dist = st.file_uploader("1) Distribuição de Dias Letivos (.xlsx)", type=['xlsx', 'xls'])
    arquivo_diario = st.file_uploader("2) Diário Escolar (.pdf)", type=['pdf'])
    ano_letivo = st.number_input("Ano letivo", min_value=2020, max_value=2100, value=2026, step=1)

if not arquivo_dist or not arquivo_diario:
    st.info("⬅️ Faça o upload dos dois arquivos (planilha + diário) na barra lateral para começar.")
    st.stop()

# Lê planilha
try:
    etapas = parse_distribuicao(arquivo_dist, ano=ano_letivo)
except Exception as e:
    st.error(f"Erro ao ler a planilha: {e}")
    st.stop()

if not etapas:
    st.error("⚠️ Não foi possível encontrar as etapas no arquivo de distribuição.")
    with st.expander("🔍 Ver as 30 primeiras linhas da planilha (diagnóstico)", expanded=True):
        df_debug = pd.read_excel(arquivo_dist, header=None)
        st.dataframe(df_debug.head(30))
    st.stop()

# Lê diário
try:
    df_diario = parse_diario_pdf(arquivo_diario, ano_padrao=ano_letivo)
except Exception as e:
    st.error(f"Erro ao ler o diário: {e}")
    st.stop()

st.success(f"✔ Planilha carregada — etapas encontradas: {', '.join(etapas.keys())}")
st.success(f"✔ Diário lido — {len(df_diario)} lançamento(s) encontrado(s)")

# Debug: mostra o que foi extraído da planilha
with st.expander("🔍 Ver o que foi extraído da planilha (conferência)"):
    for et_nome, et_dados in etapas.items():
        st.markdown(f"**{et_nome}**")
        linhas = []
        for dia_sem, datas in et_dados.items():
            linhas.append({
                'Dia': dia_sem,
                'Qtd': len(datas),
                'Datas': ', '.join(d.strftime('%d/%m') for d in datas) if datas else '—'
            })
        st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

# Seleção de etapa
etapa = st.selectbox("📌 Etapa", list(etapas.keys()))

st.subheader("⚙️ Aulas por dia da semana")
st.caption("Ajuste conforme a turma/disciplina. Depois clique em RODAR CONFERÊNCIA.")

default_aulas = {'2ª Feira': 2, '3ª Feira': 1, '4ª Feira': 1,
                 '5ª Feira': 1, '6ª Feira': 1, 'Sábado': 0}

aulas_por_dia = {}
colunas = st.columns(len(DIAS_SEMANA))
for i, dia_sem in enumerate(DIAS_SEMANA):
    with colunas[i]:
        aulas_por_dia[dia_sem] = st.number_input(
            dia_sem, min_value=0, max_value=10,
            value=default_aulas[dia_sem], step=1, key=f"aulas_{dia_sem}"
        )

# --------- BOTÃO PRINCIPAL ---------
if st.button("▶️ RODAR CONFERÊNCIA", type="primary", use_container_width=True):
    resultado = conferir(etapas, etapa, aulas_por_dia, df_diario)
    if resultado is None:
        st.warning("Nenhum dia letivo encontrado para essa etapa com as aulas informadas.")
        st.stop()

    st.markdown("---")
    st.subheader("📊 Tabela de referência (estilo Pasta1)")
    df_p1 = resultado['df_pasta1'].copy()
    total_row = pd.DataFrame([{
        'Dias da semana': 'TOTAL', 'Nº de aulas no dia': '',
        'Total de dias da semana na etapa': df_p1['Total de dias da semana na etapa'].sum(),
        'Total de aulas': df_p1['Total de aulas'].sum()
    }])
    st.dataframe(pd.concat([df_p1, total_row], ignore_index=True),
                 use_container_width=True, hide_index=True)

    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total esperado", resultado['total_esperado'])
    c2.metric("Total lançado", resultado['total_lancado'])
    dif = resultado['total_lancado'] - resultado['total_esperado']
    c3.metric("Diferença", dif, delta=dif,
              delta_color="inverse" if dif < 0 else "normal")

    st.markdown("---")
    st.subheader("❗ Dias letivos sem lançamento")
    if resultado['faltantes'].empty:
        st.success("Nenhum dia faltante. 🎉")
    else:
        st.warning(f"{len(resultado['faltantes'])} dia(s) sem lançamento:")
        df_f = resultado['faltantes'].copy()
        df_f['data'] = pd.to_datetime(df_f['data']).dt.strftime('%d/%m/%Y')
        st.dataframe(df_f, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("⚠️ Lançamentos fora dos dias esperados")
    if resultado['extras'].empty:
        st.info("Nenhum lançamento fora do esperado.")
    else:
        df_e = resultado['extras'].copy()
        df_e['data'] = pd.to_datetime(df_e['data']).dt.strftime('%d/%m/%Y')
        st.dataframe(df_e, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("🔀 Divergências de quantidade")
    if resultado['divergentes'].empty:
        st.info("Nenhuma divergência de quantidade.")
    else:
        df_d = resultado['divergentes'].copy()
        df_d['data'] = pd.to_datetime(df_d['data']).dt.strftime('%d/%m/%Y')
        df_d = df_d[['data', 'aulas_esperadas', 'aulas']]
        df_d.columns = ['Data', 'Aulas esperadas', 'Aulas lançadas']
        st.dataframe(df_d, use_container_width=True, hide_index=True)

    if not resultado['faltantes'].empty:
        csv = resultado['faltantes'].copy()
        csv['data'] = pd.to_datetime(csv['data']).dt.strftime('%d/%m/%Y')
        st.download_button(
            "⬇️ Baixar dias faltantes (CSV)",
            csv.to_csv(index=False, sep=';').encode('utf-8-sig'),
            file_name=f"dias_faltantes_{etapa.replace(' ', '_')}.csv",
            mime="text/csv"
        )

import streamlit as st
import pandas as pd
import pdfplumber
import openpyxl
import re
from datetime import datetime, date, timedelta
from io import BytesIO

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

DIAS_SEMANA = ['2ª Feira', '3ª Feira', '4ª Feira', '5ª Feira', '6ª Feira', 'Sábado']

# Mapeamento weekday Python (0=segunda) -> nome
WEEKDAY_MAP = {
    0: '2ª Feira',
    1: '3ª Feira',
    2: '4ª Feira',
    3: '5ª Feira',
    4: '6ª Feira',
    5: 'Sábado',
}


# ============================================================
# PARSER — PLANILHA DE DISTRIBUIÇÃO DE DIAS LETIVOS
# ============================================================
def _extrair_datas_celula(valor, mes_num=None, ano=2026):
    """Extrai uma lista de datas a partir do conteúdo de uma célula."""
    if valor is None:
        return []

    # Já é datetime/date (openpyxl devolve isso para datas reais)
    if isinstance(valor, datetime):
        return [valor.date()]
    if isinstance(valor, date):
        return [valor]

    texto = str(valor).strip()
    if not texto or texto.upper() in ('NAN', 'NONE'):
        return []

    # Data completa: "2026-03-07 00:00:00"
    if re.match(r'^\d{4}-\d{2}-\d{2}', texto):
        try:
            return [datetime.strptime(texto[:10], '%Y-%m-%d').date()]
        except ValueError:
            pass

    # Data completa: "07/03/2026" ou "07-03-2026"
    m = re.match(r'^(\d{1,2})[/\.\-](\d{1,2})[/\.\-](\d{2,4})', texto)
    if m:
        dia, mes, ano_str = int(m.group(1)), int(m.group(2)), m.group(3)
        ano_conv = int(ano_str) if len(ano_str) == 4 else 2000 + int(ano_str)
        try:
            return [date(ano_conv, mes, dia)]
        except ValueError:
            return []

    # Dias soltos: "9.23", "04, 11,25", "1, 2 3 4 5 6 7"
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


def parse_distribuicao(arquivo_excel, ano=2026):
    """
    Lê a planilha de distribuição de dias letivos.
    Retorna: { '1ª ETAPA': { '2ª Feira': [date, ...], ... }, ... }
    """
    wb = openpyxl.load_workbook(arquivo_excel, data_only=True)
    ws = wb.active

    etapas = {}
    etapa_atual = None
    mes_atual = None

    for row in ws.iter_rows(values_only=True):
        if row is None or all(c is None for c in row):
            continue

        primeira = str(row[0]).strip() if row[0] is not None else ''

        # Detecta o início de uma etapa
        if primeira in ('1ª ETAPA', '2ª ETAPA', '3ª ETAPA'):
            etapa_atual = primeira
            etapas[etapa_atual] = {d: [] for d in DIAS_SEMANA}
            mes_atual = None
            continue

        if etapa_atual is None:
            continue

        # Ignora linhas de cabeçalho / total
        if primeira.lower() in ('total', 'conselho de classe:', '2º feira'):
            continue
        if primeira.startswith('Conselho'):
            continue

        # Linha de mês
        if primeira in MESES:
            mes_atual = MESES[primeira]
            for idx, dia_sem in enumerate(DIAS_SEMANA[:5]):  # 2ª a 6ª
                col_idx = idx + 1
                if col_idx < len(row):
                    datas = _extrair_datas_celula(row[col_idx], mes_atual, ano)
                    etapas[etapa_atual][dia_sem].extend(datas)
            continue

        # Linha de sábado
        if primeira == 'Sábado':
            for idx, valor in enumerate(row[1:], start=1):
                datas = _extrair_datas_celula(valor, None, ano)
                etapas[etapa_atual]['Sábado'].extend(datas)
            continue

    # Remove duplicatas e ordena
    for et in etapas:
        for d in etapas[et]:
            etapas[et][d] = sorted(set(etapas[et][d]))

    return etapas


# ============================================================
# PARSER — DIÁRIO ESCOLAR (PDF)
# ============================================================
def parse_diario_pdf(arquivo_pdf, ano_padrao=2026):
    """
    Extrai do PDF do diário as datas e quantidades de aulas lançadas.
    Foca na tabela do 'REGISTRO COMPLEMENTAR / SÍNTESE DOS CONTEÚDOS LECIONADOS'.
    """
    registros = []

    with pdfplumber.open(arquivo_pdf) as pdf:
        for page in pdf.pages:
            tabelas = page.extract_tables()
            for tabela in tabelas:
                if not tabela or len(tabela) < 2:
                    continue

                # Descobre índices das colunas de interesse pelo cabeçalho
                col_dia = None
                col_aulas = None
                col_conteudo = None

                # Varre as 2 primeiras linhas (cabeçalhos podem ter quebra)
                for header_row in tabela[:2]:
                    for i, celula in enumerate(header_row):
                        if celula is None:
                            continue
                        h = str(celula).strip().upper()
                        if 'DIA/MÊS' in h or 'DIA/MES' in h:
                            col_dia = i
                        elif h == 'AULAS' or h.startswith('AULAS'):
                            col_aulas = i
                        elif 'SÍNTESE' in h or 'SINTESE' in h:
                            col_conteudo = i

                if col_dia is None or col_aulas is None:
                    continue

                for linha in tabela[1:]:
                    if not linha or len(linha) <= max(col_dia, col_aulas):
                        continue

                    valor_dia = linha[col_dia]
                    valor_aulas = linha[col_aulas]

                    if valor_dia is None or valor_aulas is None:
                        continue

                    texto_dia = str(valor_dia).strip()
                    texto_aulas = str(valor_aulas).strip()

                    # Aceita "18/05", "18.05", "18/05/2026"
                    m = re.match(r'^(\d{1,2})[/\.\-](\d{1,2})(?:[/\.\-](\d{2,4}))?', texto_dia)
                    if not m:
                        continue

                    dia = int(m.group(1))
                    mes = int(m.group(2))
                    ano_str = m.group(3)
                    ano = int(ano_str) if ano_str and len(ano_str) == 4 else (
                        2000 + int(ano_str) if ano_str else ano_padrao
                    )

                    try:
                        dt = date(ano, mes, dia)
                    except ValueError:
                        continue

                    # Só dígitos do campo aulas
                    digitos = re.sub(r'\D', '', texto_aulas)
                    if not digitos:
                        continue
                    aulas = int(digitos)

                    conteudo = ''
                    if col_conteudo is not None and col_conteudo < len(linha):
                        conteudo = str(linha[col_conteudo] or '').replace('\n', ' ').strip()

                    registros.append({
                        'data': dt,
                        'aulas': aulas,
                        'conteudo': conteudo
                    })

    df = pd.DataFrame(registros)
    if not df.empty:
        # Se a mesma data aparecer em várias linhas (uma por aluno), agregamos
        df = df.groupby('data', as_index=False).agg({
            'aulas': 'max',           # evita duplicação
            'conteudo': 'first'
        })
    return df


# ============================================================
# CONFERÊNCIA
# ============================================================
def conferir(etapas, etapa, aulas_por_dia, df_diario):
    """Confronta o esperado x o lançado na etapa escolhida."""

    dias_esperados = []
    for dia_sem in DIAS_SEMANA:
        qtd = int(aulas_por_dia.get(dia_sem, 0))
        if qtd <= 0:
            continue
        for dt in etapas[etapa][dia_sem]:
            dias_esperados.append({
                'data': dt,
                'dia_semana': dia_sem,
                'aulas_esperadas': qtd
            })

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

    # Faltantes
    faltantes = df_esp[~df_esp['data'].isin(df_diario['data'])].copy()

    # Extras (lançados fora dos dias esperados)
    extras = df_diario[~df_diario['data'].isin(df_esp['data'])].copy()

    # Divergências
    merged = df_esp.merge(
        df_diario[['data', 'aulas']],
        on='data',
        how='outer',
        indicator=True,
        suffixes=('_esp', '_lanc')
    )
    divergentes = merged[
        (merged['_merge'] == 'both') &
        (merged['aulas_esperadas'] != merged['aulas'])
    ].copy()

    # Tabela estilo Pasta1
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
        'df_esp': df_esp.sort_values('data'),
        'df_diario': df_diario.sort_values('data'),
    }


# ============================================================
# INTERFACE
# ============================================================
st.title("📚 Sistema de Conferência de Aulas — CTPM Lavras")
st.caption("Faça o upload dos arquivos, informe as aulas por dia da semana e verifique o que falta lançar no diário.")

# --- Uploads ---
with st.sidebar:
    st.header("📂 Upload de arquivos")

    arquivo_dist = st.file_uploader(
        "1) Distribuição de Dias Letivos (.xlsx)",
        type=['xlsx', 'xls'],
        help="Planilha 'DISTRIBUIÇÃO DIAS LETIVOS - CALENDÁRIO UNIDADE LAVRAS.xlsx'"
    )

    arquivo_diario = st.file_uploader(
        "2) Diário Escolar (.pdf)",
        type=['pdf'],
        help="Ex.: 'Diário_Disciplina_Turma_12201_017.pdf'"
    )

    arquivo_calendario = st.file_uploader(
        "3) Calendário Escolar (.pdf) — opcional",
        type=['pdf']
    )

    ano_letivo = st.number_input("Ano letivo", min_value=2020, max_value=2100, value=2026, step=1)

# --- Processamento ---
if arquivo_dist and arquivo_diario:
    try:
        etapas = parse_distribuicao(arquivo_dist, ano=ano_letivo)
        df_diario = parse_diario_pdf(arquivo_diario, ano_padrao=ano_letivo)

        if not etapas:
            st.error("Não foi possível encontrar as etapas no arquivo de distribuição.")
            st.stop()

        st.success(f"✔ Distribuição carregada: {len(etapas)} etapa(s) encontrada(s).")
        st.success(f"✔ Diário lido: {len(df_diario)} lançamento(s) extraído(s).")

        etapa = st.selectbox("📌 Selecione a etapa para conferência", list(etapas.keys()))

        st.subheader("⚙️ Aulas por dia da semana")
        st.caption("Informe quantas aulas a turma/disciplina tem em cada dia da semana (nesta etapa).")

        default_aulas = {
            '2ª Feira': 2,
            '3ª Feira': 1,
            '4ª Feira': 1,
            '5ª Feira': 1,
            '6ª Feira': 1,
            'Sábado': 0,
        }

        aulas_por_dia = {}
        colunas = st.columns(len(DIAS_SEMANA))
        for i, dia_sem in enumerate(DIAS_SEMANA):
            with colunas[i]:
                aulas_por_dia[dia_sem] = st.number_input(
                    dia_sem,
                    min_value=0,
                    max_value=10,
                    value=default_aulas[dia_sem],
                    step=1,
                    key=f"aulas_{dia_sem}"
                )

        if st.button("🔍 Conferir", type="primary", use_container_width=True):
            resultado = conferir(etapas, etapa, aulas_por_dia, df_diario)

            if resultado is None:
                st.warning("Nenhum dia letivo encontrado para essa etapa.")
                st.stop()

            st.markdown("---")
            st.subheader("📊 Tabela de referência (estilo Pasta1)")

            df_p1 = resultado['df_pasta1'].copy()
            total_row = pd.DataFrame([{
                'Dias da semana': 'TOTAL',
                'Nº de aulas no dia': '',
                'Total de dias da semana na etapa': df_p1['Total de dias da semana na etapa'].sum(),
                'Total de aulas': df_p1['Total de aulas'].sum()
            }])
            st.dataframe(
                pd.concat([df_p1, total_row], ignore_index=True),
                use_container_width=True,
                hide_index=True
            )

            # ---- Métricas ----
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            col1.metric("Total esperado", resultado['total_esperado'])
            col2.metric("Total lançado", resultado['total_lancado'])
            diferenca = resultado['total_lancado'] - resultado['total_esperado']
            col3.metric("Diferença", diferenca, delta=diferenca, delta_color="inverse" if diferenca < 0 else "normal")

            # ---- Dias faltantes ----
            st.markdown("---")
            st.subheader("❗ Dias letivos sem lançamento")

            if resultado['faltantes'].empty:
                st.success("Nenhum dia faltante. 🎉")
            else:
                st.warning(f"Foram encontrados **{len(resultado['faltantes'])}** dia(s) sem lançamento.")
                df_falt = resultado['faltantes'].copy()
                df_falt['data'] = pd.to_datetime(df_falt['data']).dt.strftime('%d/%m/%Y')
                st.dataframe(df_falt, use_container_width=True, hide_index=True)

            # ---- Extras ----
            st.markdown("---")
            st.subheader("⚠️ Lançamentos fora dos dias letivos esperados")
            if resultado['extras'].empty:
                st.info("Nenhum lançamento fora do esperado.")
            else:
                df_ex = resultado['extras'].copy()
                df_ex['data'] = pd.to_datetime(df_ex['data']).dt.strftime('%d/%m/%Y')
                st.dataframe(df_ex, use_container_width=True, hide_index=True)

            # ---- Divergências ----
            st.markdown("---")
            st.subheader("🔀 Divergências de quantidade de aulas")
            if resultado['divergentes'].empty:
                st.info("Nenhuma divergência de quantidade.")
            else:
                df_div = resultado['divergentes'].copy()
                df_div['data'] = pd.to_datetime(df_div['data']).dt.strftime('%d/%m/%Y')
                df_div = df_div[['data', 'aulas_esperadas', 'aulas']]
                df_div.columns = ['Data', 'Aulas esperadas', 'Aulas lançadas']
                st.dataframe(df_div, use_container_width=True, hide_index=True)

            # ---- Download CSV ----
            st.markdown("---")
            csv = resultado['faltantes'].copy()
            if not csv.empty:
                csv['data'] = pd.to_datetime(csv['data']).dt.strftime('%d/%m/%Y')
                csv_bytes = csv.to_csv(index=False, sep=';').encode('utf-8-sig')
                st.download_button(
                    "⬇️ Baixar dias faltantes (CSV)",
                    csv_bytes,
                    file_name=f"dias_faltantes_{etapa.replace(' ', '_')}.csv",
                    mime="text/csv"
                )

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
        with st.expander("Detalhes técnicos"):
            st.exception(e)

else:
    st.info("⬅️ Faça o upload dos arquivos na barra lateral para começar.")
import streamlit as st
import pandas as pd
import pdfplumber
import re
from datetime import datetime, date

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(
    page_title="Conferência de Aulas | CTPM Lavras",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

MESES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
}

DIAS_SEMANA = [
    "2ª Feira",
    "3ª Feira",
    "4ª Feira",
    "5ª Feira",
    "6ª Feira",
    "Sábado",
]

COLUNAS_DIAS = {
    "2ª Feira": 2,
    "3ª Feira": 3,
    "4ª Feira": 4,
    "5ª Feira": 5,
    "6ª Feira": 6,
}

PADRAO_AULAS = {
    "2ª Feira": 2,
    "3ª Feira": 1,
    "4ª Feira": 1,
    "5ª Feira": 1,
    "6ª Feira": 1,
    "Sábado": 0,
}


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def normalizar_texto(valor):
    if valor is None or pd.isna(valor):
        return ""
    return re.sub(r"\s+", " ", str(valor).strip())


def nome_mes_para_numero(valor):
    texto = normalizar_texto(valor).lower()
    return MESES.get(texto)


def extrair_numeros_dias(valor):
    """
    Converte células como:
      01,08,15,22,29
      9.23
      10.24
      04, 11,25
      2026-03-07
    em uma lista de números de dia.
    """
    if valor is None or pd.isna(valor):
        return []

    if isinstance(valor, (datetime, pd.Timestamp)):
        return [int(valor.day)]

    if isinstance(valor, date):
        return [int(valor.day)]

    texto = normalizar_texto(valor)
    if not texto:
        return []

    # Datas completas
    m = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", texto)
    if m:
        return [int(m.group(3))]

    # Células da planilha usam vírgula ou ponto para separar dias.
    numeros = re.findall(r"\d{1,2}", texto)

    resultado = []
    for n in numeros:
        dia = int(n)
        if 1 <= dia <= 31:
            resultado.append(dia)

    return resultado


def datas_da_celula(valor, mes_num, ano):
    resultado = []

    if valor is None or pd.isna(valor):
        return resultado

    # Se for uma data real do Excel
    if isinstance(valor, (datetime, pd.Timestamp)):
        try:
            return [date(valor.year, valor.month, valor.day)]
        except ValueError:
            return []

    if isinstance(valor, date):
        return [valor]

    for dia in extrair_numeros_dias(valor):
        try:
            resultado.append(date(ano, mes_num, dia))
        except ValueError:
            pass

    return resultado


# ============================================================
# LEITURA DA DISTRIBUIÇÃO DE DIAS LETIVOS
# ============================================================
def parse_distribuicao(arquivo_excel, ano):
    """
    Lê a planilha de distribuição e cria:

    {
        "1ª Etapa": {
            "2ª Feira": [datas...],
            ...
        },
        "2ª Etapa": {...},
        "3ª Etapa": {...}
    }

    A planilha enviada possui as colunas de segunda a sexta
    e linhas específicas para sábado.
    """
    arquivo_excel.seek(0)
    df = pd.read_excel(arquivo_excel, header=None, dtype=object)

    etapas = {}
    etapa_atual = None

    for _, row in df.iterrows():
        cells = [normalizar_texto(c) for c in row.tolist()]

        if not any(cells):
            continue

        # --------------------------------------------------------
        # Detecta etapa
        # --------------------------------------------------------
        etapa_encontrada = None

        for c in cells:
            m = re.search(
                r"\b([123])\s*[ªº]?\s*etapa\b",
                c,
                flags=re.IGNORECASE,
            )
            if m:
                etapa_encontrada = f"{m.group(1)}ª Etapa"
                break

        if etapa_encontrada:
            etapa_atual = etapa_encontrada
            etapas.setdefault(
                etapa_atual,
                {dia: [] for dia in DIAS_SEMANA}
            )

        if etapa_atual is None:
            continue

        # --------------------------------------------------------
        # Detecta mês
        # --------------------------------------------------------
        mes_num = None

        for c in cells:
            mes = nome_mes_para_numero(c)
            if mes:
                mes_num = mes
                break

        # --------------------------------------------------------
        # Linhas normais de segunda a sexta
        # --------------------------------------------------------
        if mes_num:
            for dia_semana, coluna in COLUNAS_DIAS.items():
                if coluna >= len(row):
                    continue

                valor = row.iloc[coluna]

                if pd.isna(valor) or normalizar_texto(valor) == "":
                    continue

                etapas[etapa_atual][dia_semana].extend(
                    datas_da_celula(valor, mes_num, ano)
                )

            continue

        # --------------------------------------------------------
        # Linha especial de sábado
        # Exemplo:
        # Sábado | 2026-03-07 | ...
        # --------------------------------------------------------
        if len(cells) > 1 and cells[1].lower() in ("sábado", "sabado"):
            for valor in row.iloc[2:]:
                if pd.isna(valor):
                    continue

                if isinstance(valor, (datetime, pd.Timestamp)):
                    etapas[etapa_atual]["Sábado"].append(valor.date())
                elif isinstance(valor, date):
                    etapas[etapa_atual]["Sábado"].append(valor)

    # Remove duplicidades e ordena
    for etapa in etapas:
        for dia in etapas[etapa]:
            etapas[etapa][dia] = sorted(
                set(etapas[etapa][dia])
            )

    return etapas


# ============================================================
# LEITURA DO DIÁRIO PDF
# ============================================================
def extrair_etapa_diario(texto):
    """
    Tenta descobrir a etapa no cabeçalho do diário.
    Exemplo: PERÍODO DE REFERÊNCIA 2 ETAPA
    """
    m = re.search(
        r"PER[IÍ]ODO\s+DE\s+REFER[EÊ]NCIA\s+([123])\s*ETAPA",
        texto,
        flags=re.IGNORECASE,
    )

    if m:
        return f"{m.group(1)}ª Etapa"

    return None


def extrair_cabecalho_diario(texto):
    dados = {
        "etapa": extrair_etapa_diario(texto),
        "ano": None,
        "unidade": None,
        "turma": None,
        "disciplina": None,
        "professor": None,
    }

    m = re.search(r"\b(20\d{2})\b", texto)
    if m:
        dados["ano"] = int(m.group(1))

    m = re.search(r"UNIDADE\s+(\d+)", texto, re.IGNORECASE)
    if m:
        dados["unidade"] = m.group(1)

    m = re.search(
        r"TURMA\s+([A-Z0-9]+)",
        texto,
        re.IGNORECASE,
    )
    if m:
        dados["turma"] = m.group(1)

    m = re.search(
        r"DISCIPLINA\s+(.+?)(?:\s+PROFESSOR|\s+MESES|\n)",
        texto,
        re.IGNORECASE,
    )
    if m:
        dados["disciplina"] = normalizar_texto(m.group(1))

    m = re.search(
        r"PROFESSOR\s+(.+?)(?:\n|$)",
        texto,
        re.IGNORECASE,
    )
    if m:
        dados["professor"] = normalizar_texto(m.group(1))

    return dados


def parse_diario_pdf(arquivo_pdf, ano_padrao):
    """
    O PDF do diário não necessariamente coloca a data no início
    da linha. Em algumas páginas ela aparece depois da matrícula
    do aluno.

    Por isso usamos re.search(), e não re.match().

    Exemplo real:
      001 ... 18/05 2 Correção de atividades...

    O resultado final é agrupado por data.
    """
    registros = []
    texto_completo = ""

    arquivo_pdf.seek(0)

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            texto_completo += "\n" + texto

            for linha in texto.splitlines():
                linha = linha.strip()

                if not linha:
                    continue

                # Data + número de aulas + conteúdo.
                # Procura em qualquer posição da linha.
                padrao = re.compile(
                    r"(?<!\d)"
                    r"(\d{1,2})/(\d{1,2})"
                    r"(?:/(\d{2,4}))?"
                    r"\s+"
                    r"(\d{1,2})"
                    r"\s+"
                    r"(.+?)"
                    r"\s*$"
                )

                m = padrao.search(linha)

                if not m:
                    continue

                dia = int(m.group(1))
                mes = int(m.group(2))
                ano_str = m.group(3)
                aulas = int(m.group(4))
                conteudo = normalizar_texto(m.group(5))

                if ano_str:
                    ano = int(ano_str)
                    if len(ano_str) == 2:
                        ano += 2000
                else:
                    ano = ano_padrao

                try:
                    dt = date(ano, mes, dia)
                except ValueError:
                    continue

                registros.append(
                    {
                        "data": dt,
                        "aulas": aulas,
                        "conteudo": conteudo,
                    }
                )

    cabecalho = extrair_cabecalho_diario(texto_completo)

    if not registros:
        df = pd.DataFrame(
            columns=["data", "aulas", "conteudo"]
        )
        return df, cabecalho, 0

    df = pd.DataFrame(registros)

    # ------------------------------------------------------------
    # Se a mesma data aparecer mais de uma vez, soma as aulas.
    # Não usamos drop_duplicates mantendo a maior quantidade,
    # pois isso poderia apagar informação.
    # ------------------------------------------------------------
    df = (
        df.groupby("data", as_index=False)
        .agg(
            aulas=("aulas", "sum"),
            conteudo=("conteudo", lambda x: " | ".join(
                dict.fromkeys(
                    str(v).strip()
                    for v in x
                    if str(v).strip()
                )
            )),
        )
        .sort_values("data")
        .reset_index(drop=True)
    )

    return df, cabecalho, len(registros)


# ============================================================
# CONFERÊNCIA
# ============================================================
def montar_dias_esperados(etapas, etapa, aulas_por_dia):
    registros = []

    if etapa not in etapas:
        return pd.DataFrame()

    for dia_semana in DIAS_SEMANA:
        qtd_aulas = int(aulas_por_dia.get(dia_semana, 0))

        if qtd_aulas <= 0:
            continue

        for dt in etapas[etapa].get(dia_semana, []):
            registros.append(
                {
                    "data": dt,
                    "dia_semana": dia_semana,
                    "aulas_esperadas": qtd_aulas,
                }
            )

    return pd.DataFrame(registros)


def conferir(etapas, etapa, aulas_por_dia, df_diario):
    """
    Compara somente a etapa selecionada.

    Importante:
    - Dia previsto e sem lançamento = FALTANTE
    - Dia previsto e quantidade diferente = DIVERGÊNCIA
    - Lançamento dentro do período, mas em dia não previsto =
      FORA DOS DIAS ESPERADOS
    """
    df_esperado = montar_dias_esperados(
        etapas,
        etapa,
        aulas_por_dia,
    )

    if df_esperado.empty:
        return None

    df_esperado["data"] = pd.to_datetime(
        df_esperado["data"]
    ).dt.date

    df_diario = df_diario.copy()

    if df_diario.empty:
        df_diario = pd.DataFrame(
            columns=["data", "aulas", "conteudo"]
        )
    else:
        df_diario["data"] = pd.to_datetime(
            df_diario["data"]
        ).dt.date

    # ------------------------------------------------------------
    # Considera no diário somente o intervalo da etapa.
    # Isso impede que lançamentos de outra etapa sejam somados.
    # ------------------------------------------------------------
    inicio = df_esperado["data"].min()
    fim = df_esperado["data"].max()

    diario_etapa = df_diario[
        (df_diario["data"] >= inicio)
        & (df_diario["data"] <= fim)
    ].copy()

    datas_esperadas = set(df_esperado["data"])
    datas_lancadas = set(diario_etapa["data"])

    total_esperado = int(
        df_esperado["aulas_esperadas"].sum()
    )

    total_lancado = int(
        diario_etapa["aulas"].sum()
    ) if not diario_etapa.empty else 0

    # ------------------------------------------------------------
    # FALTANTES
    # ------------------------------------------------------------
    faltantes = df_esperado[
        ~df_esperado["data"].isin(datas_lancadas)
    ].copy()

    # ------------------------------------------------------------
    # LANÇAMENTOS FORA DOS DIAS ESPERADOS
    # ------------------------------------------------------------
    extras = diario_etapa[
        ~diario_etapa["data"].isin(datas_esperadas)
    ].copy()

    # ------------------------------------------------------------
    # DIVERGÊNCIA DE QUANTIDADE
    # ------------------------------------------------------------
    comparacao = df_esperado.merge(
        diario_etapa[["data", "aulas", "conteudo"]],
        on="data",
        how="left",
        suffixes=("_esperada", "_lancada"),
    )

    comparacao["aulas"] = comparacao["aulas"].fillna(0).astype(int)

    divergentes = comparacao[
        (
            comparacao["aulas"] > 0
        )
        & (
            comparacao["aulas_esperadas"]
            != comparacao["aulas"]
        )
    ].copy()

    # ------------------------------------------------------------
    # RESUMO POR MÊS
    # ------------------------------------------------------------
    resumo = df_esperado.copy()
    resumo["mês"] = pd.to_datetime(
        resumo["data"]
    ).dt.strftime("%m/%Y")

    resumo = (
        resumo.groupby("mês", as_index=False)
        .agg(
            dias_previstos=("data", "count"),
            aulas_previstas=("aulas_esperadas", "sum"),
        )
    )

    lanc_mes = diario_etapa.copy()

    if not lanc_mes.empty:
        lanc_mes["mês"] = pd.to_datetime(
            lanc_mes["data"]
        ).dt.strftime("%m/%Y")

        lanc_mes = (
            lanc_mes.groupby("mês", as_index=False)
            .agg(aulas_lancadas=("aulas", "sum"))
        )

        resumo = resumo.merge(
            lanc_mes,
            on="mês",
            how="left",
        )

    resumo["aulas_lancadas"] = (
        resumo["aulas_lancadas"]
        .fillna(0)
        .astype(int)
    )

    resumo["diferença"] = (
        resumo["aulas_lancadas"]
        - resumo["aulas_previstas"]
    )

    return {
        "total_esperado": total_esperado,
        "total_lancado": total_lancado,
        "faltantes": faltantes.sort_values("data"),
        "extras": extras.sort_values("data"),
        "divergentes": divergentes.sort_values("data"),
        "resumo": resumo,
        "diario_etapa": diario_etapa,
        "esperado": df_esperado,
    }


# ============================================================
# INTERFACE
# ============================================================
st.title("📚 Conferência de Aulas")
st.caption(
    "CTPM Lavras • Conferência por dia letivo, etapa e quantidade de aulas"
)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("📂 1. Arquivos")

    arquivo_dist = st.file_uploader(
        "Distribuição de dias letivos (.xlsx)",
        type=["xlsx", "xls"],
        help="Arquivo que contém os dias letivos por etapa e dia da semana.",
    )

    arquivo_diario = st.file_uploader(
        "Diário escolar (.pdf)",
        type=["pdf"],
        help="Diário da disciplina/turma que será conferido.",
    )


    ano_letivo = st.number_input(
        "Ano letivo",
        min_value=2020,
        max_value=2100,
        value=2026,
        step=1,
    )

if not arquivo_dist or not arquivo_diario:
    st.info(
        "👈 Envie os dois arquivos: distribuição de dias letivos e "
        "diário escolar."
    )

    st.markdown(
        """
        ### Como o sistema funciona

        **1. Distribuição de dias letivos**
        - informa quantos dias existem em cada dia da semana;
        - separa 1ª, 2ª e 3ª etapas;
        - define as datas que podem receber aulas.

        **2. Diário escolar**
        - lê as datas lançadas pelo professor;
        - lê o número de aulas lançado em cada data;
        - identifica o total efetivamente lançado.

        **3. Calendário escolar**
        - fica associado à conferência como documento de referência do ano.

        **4. Conferência**
        - calcula as aulas previstas;
        - calcula as aulas lançadas;
        - encontra dias previstos sem lançamento;
        - encontra quantidade de aulas divergente;
        - encontra lançamentos em dias não previstos.
        """
    )

    st.stop()


# ============================================================
# CARREGA ARQUIVOS
# ============================================================
try:
    etapas = parse_distribuicao(
        arquivo_dist,
        ano=int(ano_letivo),
    )
except Exception as e:
    st.error(f"Erro ao ler a distribuição: {e}")
    st.stop()

if not etapas:
    st.error(
        "Não foi possível identificar as etapas na planilha."
    )
    st.stop()

try:
    (
        df_diario,
        info_diario,
        quantidade_registros,
    ) = parse_diario_pdf(
        arquivo_diario,
        ano_padrao=int(ano_letivo),
    )
except Exception as e:
    st.error(f"Erro ao ler o diário PDF: {e}")
    st.stop()


# ============================================================
# CABEÇALHO / INFORMAÇÕES
# ============================================================
st.success(
    f"✅ Distribuição carregada: {len(etapas)} etapa(s)"
)

st.success(
    f"✅ Diário processado: {len(df_diario)} data(s) "
    f"identificada(s)"
)

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "Ano",
    info_diario.get("ano") or int(ano_letivo),
)

c2.metric(
    "Etapa do diário",
    info_diario.get("etapa") or "Não identificada",
)

c3.metric(
    "Turma",
    info_diario.get("turma") or "Não identificada",
)

c4.metric(
    "Disciplina",
    info_diario.get("disciplina") or "Não identificada",
)

if info_diario.get("professor"):
    st.caption(
        f"👨‍🏫 Professor(a): **{info_diario['professor']}**"
    )

# ============================================================
# ETAPA
# ============================================================
st.markdown("---")
st.header("⚙️ 2. Configuração da conferência")

etapas_disponiveis = list(etapas.keys())

etapa_padrao = info_diario.get("etapa")

if etapa_padrao in etapas_disponiveis:
    indice = etapas_disponiveis.index(etapa_padrao)
else:
    indice = 0

etapa = st.selectbox(
    "Etapa que será conferida",
    etapas_disponiveis,
    index=indice,
)

# ============================================================
# TABELA DE AULAS DA DISCIPLINA - MODELO PASTA 1
# ============================================================
st.subheader("📝 Distribuição de aulas da disciplina")

st.caption(
    "O campo **Nº de aulas no dia** é preenchido por você. "
    "A quantidade de dias letivos é puxada automaticamente da "
    "planilha de Distribuição de Dias Letivos."
)

chave_editor = f"config_aulas_{etapa}_{ano_letivo}"

if chave_editor not in st.session_state:
    linhas = []
    for dia in DIAS_SEMANA:
        qtd_dias = len(etapas[etapa].get(dia, []))
        linhas.append(
            {
                "Dias da semana": dia,
                "Nº de aulas no dia": PADRAO_AULAS[dia],
                "Total de dias da semana na etapa": qtd_dias,
                "Total de aulas": PADRAO_AULAS[dia] * qtd_dias,
            }
        )
    st.session_state[chave_editor] = pd.DataFrame(linhas)

df_config = st.data_editor(
    st.session_state[chave_editor],
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    key=f"editor_{chave_editor}",
    column_config={
        "Dias da semana": st.column_config.TextColumn(
            "Dias da semana", disabled=True, width="medium"
        ),
        "Nº de aulas no dia": st.column_config.NumberColumn(
            "Nº de aulas no dia",
            min_value=0,
            max_value=20,
            step=1,
            width="small",
            help="✏️ Informe quantas aulas dessa disciplina são dadas nesse dia.",
        ),
        "Total de dias da semana na etapa": st.column_config.NumberColumn(
            "Total de dias da semana na etapa",
            disabled=True,
            width="large",
            help="🔒 Vem automaticamente da planilha de distribuição.",
        ),
        "Total de aulas": st.column_config.NumberColumn(
            "Total de aulas",
            disabled=True,
            width="small",
            help="🔒 Nº de aulas no dia × total de dias da semana.",
        ),
    },
)

df_config["Nº de aulas no dia"] = (
    pd.to_numeric(df_config["Nº de aulas no dia"], errors="coerce")
    .fillna(0).astype(int)
)
df_config["Total de dias da semana na etapa"] = (
    pd.to_numeric(
        df_config["Total de dias da semana na etapa"], errors="coerce"
    ).fillna(0).astype(int)
)
df_config["Total de aulas"] = (
    df_config["Nº de aulas no dia"]
    * df_config["Total de dias da semana na etapa"]
)

st.session_state[chave_editor] = df_config

total_dias = int(df_config["Total de dias da semana na etapa"].sum())
total_aulas = int(df_config["Total de aulas"].sum())

df_total = pd.DataFrame([{
    "Dias da semana": "TOTAL",
    "Nº de aulas no dia": "",
    "Total de dias da semana na etapa": total_dias,
    "Total de aulas": total_aulas,
}])

st.markdown("### 📊 Resumo da etapa")
m1, m2 = st.columns(2)
m1.metric("Total de dias letivos da etapa", total_dias)
m2.metric("Total de aulas da disciplina", total_aulas)

st.dataframe(
    pd.concat([df_config, df_total], ignore_index=True),
    use_container_width=True,
    hide_index=True,
)

aulas_por_dia = dict(
    zip(df_config["Dias da semana"], df_config["Nº de aulas no dia"])
)

# ============================================================
# LISTA DE DATAS ESPERADAS
# ============================================================
with st.expander("📅 Ver todos os dias letivos considerados"):
    df_datas = montar_dias_esperados(
        etapas,
        etapa,
        aulas_por_dia,
    )

    if df_datas.empty:
        st.warning("Nenhuma data encontrada.")
    else:
        df_datas = df_datas.copy()
        df_datas["Data"] = pd.to_datetime(
            df_datas["data"]
        ).dt.strftime("%d/%m/%Y")

        df_datas = df_datas[
            [
                "Data",
                "dia_semana",
                "aulas_esperadas",
            ]
        ]

        df_datas.columns = [
            "Data",
            "Dia da semana",
            "Aulas esperadas",
        ]

        st.dataframe(
            df_datas,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# BOTÃO DE CONFERÊNCIA
# ============================================================
st.markdown("---")
st.header("🔎 3. Conferência")

if st.button(
    "▶️ RODAR CONFERÊNCIA",
    type="primary",
    use_container_width=True,
):
    resultado = conferir(
        etapas,
        etapa,
        aulas_por_dia,
        df_diario,
    )

    if resultado is None:
        st.warning(
            "Não existem dias letivos configurados para essa etapa."
        )
        st.stop()

    st.session_state["ultimo_resultado"] = resultado


# ============================================================
# MOSTRA RESULTADO
# ============================================================
resultado = st.session_state.get("ultimo_resultado")

if resultado is not None:
    st.markdown("---")

    esperado = resultado["total_esperado"]
    lancado = resultado["total_lancado"]
    diferenca = lancado - esperado

    r1, r2, r3, r4 = st.columns(4)

    r1.metric(
        "📚 Aulas previstas",
        esperado,
    )

    r2.metric(
        "✏️ Aulas lançadas",
        lancado,
    )

    r3.metric(
        "📌 Diferença",
        diferenca,
    )

    r4.metric(
        "❗ Dias faltantes",
        len(resultado["faltantes"]),
    )

    # --------------------------------------------------------
    # STATUS PRINCIPAL
    # --------------------------------------------------------
    if diferenca == 0 and resultado["faltantes"].empty and resultado["divergentes"].empty:
        st.success(
            "🎉 Conferência fechada: a quantidade lançada "
            "corresponde às aulas previstas e não há dias "
            "faltantes ou divergências."
        )
    else:
        st.warning(
            "⚠️ A conferência encontrou diferenças. "
            "Veja os detalhes abaixo."
        )

    # --------------------------------------------------------
    # FALTANTES
    # --------------------------------------------------------
    st.subheader("🚨 Dias previstos sem lançamento")

    faltantes = resultado["faltantes"].copy()

    if faltantes.empty:
        st.success(
            "Nenhum dia letivo previsto está sem lançamento."
        )
    else:
        faltantes["Data"] = pd.to_datetime(
            faltantes["data"]
        ).dt.strftime("%d/%m/%Y")

        faltantes = faltantes[
            [
                "Data",
                "dia_semana",
                "aulas_esperadas",
            ]
        ]

        faltantes.columns = [
            "Data",
            "Dia da semana",
            "Aulas esperadas",
        ]

        st.error(
            f"🔴 Foram encontrados "
            f"**{len(faltantes)} dia(s)** sem lançamento."
        )

        st.dataframe(
            faltantes,
            use_container_width=True,
            hide_index=True,
        )

        # Destaque das datas em formato simples
        datas_faltantes = ", ".join(
            faltantes["Data"].tolist()
        )

        st.info(
            f"📌 **Data(s) a verificar:** {datas_faltantes}"
        )

    # --------------------------------------------------------
    # DIVERGÊNCIAS
    # --------------------------------------------------------
    st.subheader("🔀 Divergência no número de aulas")

    divergentes = resultado["divergentes"].copy()

    if divergentes.empty:
        st.success(
            "Nenhum lançamento possui quantidade diferente "
            "do esperado."
        )
    else:
        divergentes["Data"] = pd.to_datetime(
            divergentes["data"]
        ).dt.strftime("%d/%m/%Y")

        divergentes = divergentes[
            [
                "Data",
                "aulas_esperadas",
                "aulas",
                "conteudo",
            ]
        ]

        divergentes.columns = [
            "Data",
            "Aulas esperadas",
            "Aulas lançadas",
            "Conteúdo",
        ]

        st.dataframe(
            divergentes,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # EXTRAS
    # --------------------------------------------------------
    st.subheader("📍 Lançamentos em dias não previstos")

    extras = resultado["extras"].copy()

    if extras.empty:
        st.success(
            "Nenhum lançamento foi encontrado fora dos "
            "dias previstos pela distribuição."
        )
    else:
        extras["Data"] = pd.to_datetime(
            extras["data"]
        ).dt.strftime("%d/%m/%Y")

        extras = extras[
            [
                "Data",
                "aulas",
                "conteudo",
            ]
        ]

        extras.columns = [
            "Data",
            "Aulas lançadas",
            "Conteúdo",
        ]

        st.dataframe(
            extras,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # RESUMO POR MÊS
    # --------------------------------------------------------
    st.subheader("📊 Resumo por mês")

    resumo = resultado["resumo"].copy()

    resumo.columns = [
        "Mês",
        "Dias previstos",
        "Aulas previstas",
        "Aulas lançadas",
        "Diferença",
    ]

    st.dataframe(
        resumo,
        use_container_width=True,
        hide_index=True,
    )

    # --------------------------------------------------------
    # DIÁRIO DA ETAPA
    # --------------------------------------------------------
    with st.expander("📖 Ver lançamentos considerados no diário"):
        diario = resultado["diario_etapa"].copy()

        if diario.empty:
            st.info(
                "Nenhum lançamento do diário foi encontrado "
                "dentro do período da etapa."
            )
        else:
            diario["Data"] = pd.to_datetime(
                diario["data"]
            ).dt.strftime("%d/%m/%Y")

            diario = diario[
                [
                    "Data",
                    "aulas",
                    "conteudo",
                ]
            ]

            diario.columns = [
                "Data",
                "Aulas",
                "Conteúdo",
            ]

            st.dataframe(
                diario,
                use_container_width=True,
                hide_index=True,
            )

    # --------------------------------------------------------
    # EXPORTAÇÃO
    # --------------------------------------------------------
    st.subheader("⬇️ Exportar resultado")

    linhas_exportacao = []

    for _, row in resultado["esperado"].iterrows():
        dt = row["data"]

        lanc = resultado["diario_etapa"][
            resultado["diario_etapa"]["data"] == dt
        ]

        if lanc.empty:
            aulas_lancadas = 0
            conteudo = ""
        else:
            aulas_lancadas = int(
                lanc["aulas"].sum()
            )
            conteudo = " | ".join(
                lanc["conteudo"].astype(str)
            )

        if aulas_lancadas == 0:
            status = "FALTANTE"
        elif aulas_lancadas != int(row["aulas_esperadas"]):
            status = "DIVERGENTE"
        else:
            status = "OK"

        linhas_exportacao.append(
            {
                "Data": dt.strftime("%d/%m/%Y"),
                "Dia da semana": row["dia_semana"],
                "Aulas esperadas": int(
                    row["aulas_esperadas"]
                ),
                "Aulas lançadas": aulas_lancadas,
                "Diferença": (
                    aulas_lancadas
                    - int(row["aulas_esperadas"])
                ),
                "Status": status,
                "Conteúdo": conteudo,
            }
        )

    df_exportacao = pd.DataFrame(
        linhas_exportacao
    )

    csv = df_exportacao.to_csv(
        index=False,
        sep=";",
        encoding="utf-8-sig",
    ).encode("utf-8-sig")

    st.download_button(
        "📥 Baixar conferência completa (CSV)",
        data=csv,
        file_name=(
            f"conferencia_{etapa.replace(' ', '_')}_"
            f"{ano_letivo}.csv"
        ),
        mime="text/csv",
        use_container_width=True,
    )

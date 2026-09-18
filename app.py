import streamlit as st
import pandas as pd
import pdfplumber
import re
import hashlib
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
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
}

DIAS_SEMANA = ["2ª Feira", "3ª Feira", "4ª Feira", "5ª Feira", "6ª Feira", "Sábado"]

COLUNAS_DIAS = {
    "2ª Feira": 2, "3ª Feira": 3, "4ª Feira": 4,
    "5ª Feira": 5, "6ª Feira": 6,
}

PADRAO_AULAS = {
    "2ª Feira": 2, "3ª Feira": 1, "4ª Feira": 1,
    "5ª Feira": 1, "6ª Feira": 1, "Sábado": 0,
}

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def normalizar_texto(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", " ", str(valor).strip())


def nome_mes_para_numero(valor):
    texto = normalizar_texto(valor).lower()
    return MESES.get(texto)


def extrair_numeros_dias(valor):
    """Converte '9.23', '04, 11,25', '01,08,15' em [9,23], [4,11,25], ..."""
    if valor is None or pd.isna(valor):
        return []
    if isinstance(valor, (datetime, pd.Timestamp)):
        return [int(valor.day)]
    if isinstance(valor, date):
        return [int(valor.day)]

    texto = normalizar_texto(valor)
    if not texto:
        return []

    m = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", texto)
    if m:
        return [int(m.group(3))]

    return [int(n) for n in re.findall(r"\d{1,2}", texto)
            if 1 <= int(n) <= 31]


def datas_da_celula(valor, mes_num, ano):
    resultado = []
    if valor is None or pd.isna(valor):
        return resultado
    if isinstance(valor, (datetime, pd.Timestamp)):
        return [valor.date()]
    if isinstance(valor, date):
        return [valor]
    for dia in extrair_numeros_dias(valor):
        try:
            resultado.append(date(ano, mes_num, dia))
        except ValueError:
            pass
    return resultado


def _eh_cabecalho(cells):
    """Detecta linha de cabeçalho tipo '2º Feira | 3º Feira | ...'."""
    padrao = re.compile(r"[2-6][ªº°]?\s*feira", re.IGNORECASE)
    return any(padrao.search(c) for c in cells if c)


def _eh_linha_total(cells):
    """Total exato na coluna B."""
    return len(cells) >= 2 and cells[1].strip().lower() == "total"


def _eh_sabado(cells):
    return len(cells) >= 2 and cells[1].strip().lower() in ("sábado", "sabado")


def _file_id(f):
    if f is None:
        return "none"
    return hashlib.md5(
        f"{getattr(f, 'name', '')}_{getattr(f, 'size', 0)}".encode()
    ).hexdigest()[:8]


# ============================================================
# LEITURA DA DISTRIBUIÇÃO DE DIAS LETIVOS  (CORRIGIDO)
# ============================================================
def parse_distribuicao(arquivo_excel, ano):
    """
    Lê a planilha de distribuição.

    CORREÇÃO PRINCIPAL:
    - A seção de cada etapa é delimitada pela linha 'Total'.
    - Sempre que encontramos um 'Total', incrementamos o contador.
    - Se existir marcador explícito ('1ª etapa', '2ª etapa', '3ª etapa'),
      ele sobrescreve o contador.
    - Assim a 3ª etapa (cujo rótulo está em célula mesclada e só aparece
      na linha de Novembro) é detectada corretamente: Setembro e Outubro
      ficam na 3ª etapa, e não na 2ª.
    """
    arquivo_excel.seek(0)
    df = pd.read_excel(arquivo_excel, header=None, dtype=object)

    etapas = {}
    etapa_num = 0
    etapa_ativa = None

    for _, row in df.iterrows():
        cells = [normalizar_texto(c) for c in row.tolist()]
        if not any(cells):
            continue

        # --- marcador explícito de etapa ---
        etapa_explicita = None
        for c in cells:
            m = re.search(r"\b([123])\s*[ªº°]?\s*etapa\b", c, re.IGNORECASE)
            if m:
                etapa_explicita = int(m.group(1))
                break

        if etapa_explicita is not None:
            etapa_num = etapa_explicita
            etapa_ativa = f"{etapa_explicita}ª Etapa"
            etapas.setdefault(etapa_ativa, {d: [] for d in DIAS_SEMANA})

        # --- linha "Total" encerra a seção e avança a etapa ---
        if _eh_linha_total(cells):
            etapa_num += 1
            continue

        # --- primeira seção sem marcador explícito ---
        if etapa_ativa is None:
            if _eh_cabecalho(cells):
                etapa_num = 1
                etapa_ativa = "1ª Etapa"
                etapas.setdefault(etapa_ativa, {d: [] for d in DIAS_SEMANA})
            else:
                continue

        # sincroniza nome da etapa ativa com o contador atualizado
        etapa_ativa = f"{etapa_num}ª Etapa"
        etapas.setdefault(etapa_ativa, {d: [] for d in DIAS_SEMANA})

        # cabeçalho sem mês: só garante que a etapa existe
        if _eh_cabecalho(cells) and not any(
            nome_mes_para_numero(c) for c in cells
        ):
            continue

        # --- linha de Sábado ---
        if _eh_sabado(cells):
            for valor in row.iloc[2:]:
                if valor is None or pd.isna(valor):
                    continue
                if isinstance(valor, (datetime, pd.Timestamp)):
                    etapas[etapa_ativa]["Sábado"].append(valor.date())
                elif isinstance(valor, date):
                    etapas[etapa_ativa]["Sábado"].append(valor)
            continue

        # --- linhas com mês ---
        mes_num = None
        for c in cells:
            mes = nome_mes_para_numero(c)
            if mes:
                mes_num = mes
                break
        if mes_num is None:
            continue

        for dia_semana, coluna in COLUNAS_DIAS.items():
            if coluna >= len(row):
                continue
            valor = row.iloc[coluna]
            if valor is None or pd.isna(valor):
                continue
            if normalizar_texto(valor) == "":
                continue
            etapas[etapa_ativa][dia_semana].extend(
                datas_da_celula(valor, mes_num, ano)
            )

    # remove duplicatas
    for etapa in etapas:
        for dia in etapas[etapa]:
            etapas[etapa][dia] = sorted(set(etapas[etapa][dia]))

    return etapas


# ============================================================
# LEITURA DO DIÁRIO PDF
# ============================================================
def extrair_etapa_diario(texto):
    m = re.search(
        r"PER[IÍ]ODO\s+DE\s+REFER[EÊ]NCIA\s+([123])\s*ETAPA",
        texto, flags=re.IGNORECASE,
    )
    return f"{m.group(1)}ª Etapa" if m else None


def extrair_cabecalho_diario(texto):
    dados = {
        "etapa": extrair_etapa_diario(texto),
        "ano": None, "unidade": None, "turma": None,
        "disciplina": None, "professor": None,
    }
    m = re.search(r"\b(20\d{2})\b", texto)
    if m:
        dados["ano"] = int(m.group(1))
    m = re.search(r"UNIDADE\s+(\d+)", texto, re.IGNORECASE)
    if m:
        dados["unidade"] = m.group(1)
    m = re.search(r"TURMA\s+([A-Z0-9]+)", texto, re.IGNORECASE)
    if m:
        dados["turma"] = m.group(1)
    m = re.search(r"DISCIPLINA\s+(.+?)(?:\s+PROFESSOR|\s+MESES|\n)",
                  texto, re.IGNORECASE)
    if m:
        dados["disciplina"] = normalizar_texto(m.group(1))
    m = re.search(r"PROFESSOR\s+(.+?)(?:\n|$)", texto, re.IGNORECASE)
    if m:
        dados["professor"] = normalizar_texto(m.group(1))
    return dados


def parse_diario_pdf(arquivo_pdf, ano_padrao):
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

                # DD/MM [opcional /AA ou /AAAA] <aulas> <conteúdo>
                padrao = re.compile(
                    r"(?<!\d)(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?"
                    r"\s+(\d{1,2})\s+(.+?)\s*$"
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

                registros.append({"data": dt, "aulas": aulas, "conteudo": conteudo})

    cabecalho = extrair_cabecalho_diario(texto_completo)

    if not registros:
        return (pd.DataFrame(columns=["data", "aulas", "conteudo"]),
                cabecalho, 0)

    df = pd.DataFrame(registros)
    df = (
        df.groupby("data", as_index=False)
          .agg(
              aulas=("aulas", "sum"),
              conteudo=("conteudo", lambda x: " | ".join(
                  dict.fromkeys(str(v).strip() for v in x if str(v).strip())
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
        qtd = int(aulas_por_dia.get(dia_semana, 0))
        if qtd <= 0:
            continue
        for dt in etapas[etapa].get(dia_semana, []):
            registros.append({
                "data": dt,
                "dia_semana": dia_semana,
                "aulas_esperadas": qtd,
            })
    return pd.DataFrame(registros)


def conferir(etapas, etapa, aulas_por_dia, df_diario):
    df_esperado = montar_dias_esperados(etapas, etapa, aulas_por_dia)
    if df_esperado.empty:
        return None

    df_esperado["data"] = pd.to_datetime(df_esperado["data"]).dt.date

    if df_diario.empty:
        df_diario = pd.DataFrame(columns=["data", "aulas", "conteudo"])
    else:
        df_diario = df_diario.copy()
        df_diario["data"] = pd.to_datetime(df_diario["data"]).dt.date

    inicio = df_esperado["data"].min()
    fim = df_esperado["data"].max()
    diario_etapa = df_diario[
        (df_diario["data"] >= inicio) & (df_diario["data"] <= fim)
    ].copy()

    datas_esperadas = set(df_esperado["data"])
    datas_lancadas = set(diario_etapa["data"])

    total_esperado = int(df_esperado["aulas_esperadas"].sum())
    total_lancado = int(diario_etapa["aulas"].sum()) if not diario_etapa.empty else 0

    faltantes = df_esperado[~df_esperado["data"].isin(datas_lancadas)].copy()
    extras = diario_etapa[~diario_etapa["data"].isin(datas_esperadas)].copy()

    comparacao = df_esperado.merge(
        diario_etapa[["data", "aulas", "conteudo"]],
        on="data", how="left", suffixes=("_esperada", "_lancada"),
    )
    comparacao["aulas"] = comparacao["aulas"].fillna(0).astype(int)
    divergentes = comparacao[
        (comparacao["aulas"] > 0)
        & (comparacao["aulas_esperadas"] != comparacao["aulas"])
    ].copy()

    # Resumo por mês
    resumo = df_esperado.copy()
    resumo["mês"] = pd.to_datetime(resumo["data"]).dt.strftime("%m/%Y")
    resumo = resumo.groupby("mês", as_index=False).agg(
        dias_previstos=("data", "count"),
        aulas_previstas=("aulas_esperadas", "sum"),
    )

    if not diario_etapa.empty:
        l = diario_etapa.copy()
        l["mês"] = pd.to_datetime(l["data"]).dt.strftime("%m/%Y")
        l = l.groupby("mês", as_index=False).agg(aulas_lancadas=("aulas", "sum"))
        resumo = resumo.merge(l, on="mês", how="left")

    resumo["aulas_lancadas"] = resumo["aulas_lancadas"].fillna(0).astype(int)
    resumo["diferença"] = resumo["aulas_lancadas"] - resumo["aulas_previstas"]

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
st.caption("CTPM Lavras • Conferência por dia letivo, etapa e quantidade de aulas")

with st.sidebar:
    st.header("📂 1. Arquivos")
    arquivo_dist = st.file_uploader(
        "Distribuição de dias letivos (.xlsx)",
        type=["xlsx", "xls"],
    )
    arquivo_diario = st.file_uploader("Diário escolar (.pdf)", type=["pdf"])
    ano_letivo = st.number_input("Ano letivo", 2020, 2100, 2026, 1)

if not arquivo_dist or not arquivo_diario:
    st.info("👈 Envie os dois arquivos: distribuição de dias letivos e diário escolar.")
    st.stop()

# ------------------------------------------------------------
# Carrega arquivos
# ------------------------------------------------------------
try:
    etapas = parse_distribuicao(arquivo_dist, ano=int(ano_letivo))
except Exception as e:
    st.error(f"Erro ao ler a distribuição: {e}")
    st.stop()

if not etapas:
    st.error("Não foi possível identificar as etapas na planilha.")
    st.stop()

try:
    df_diario, info_diario, quantidade_registros = parse_diario_pdf(
        arquivo_diario, ano_padrao=int(ano_letivo)
    )
except Exception as e:
    st.error(f"Erro ao ler o diário PDF: {e}")
    st.stop()

st.success(f"✅ Distribuição carregada: {len(etapas)} etapa(s)")
st.success(f"✅ Diário processado: {len(df_diario)} data(s) identificada(s)")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Ano", info_diario.get("ano") or int(ano_letivo))
c2.metric("Etapa do diário", info_diario.get("etapa") or "Não identificada")
c3.metric("Turma", info_diario.get("turma") or "Não identificada")
c4.metric("Disciplina", info_diario.get("disciplina") or "Não identificada")
if info_diario.get("professor"):
    st.caption(f"👨‍🏫 Professor(a): **{info_diario['professor']}**")

# ------------------------------------------------------------
# Etapa
# ------------------------------------------------------------
st.markdown("---")
st.header("⚙️ 2. Configuração da conferência")

etapas_disponiveis = list(etapas.keys())
etapa_padrao = info_diario.get("etapa")
indice = etapas_disponiveis.index(etapa_padrao) if etapa_padrao in etapas_disponiveis else 0
etapa = st.selectbox("Etapa que será conferida", etapas_disponiveis, index=indice)

# ------------------------------------------------------------
# Tabela de aulas (modelo Pasta1)
# ------------------------------------------------------------
st.subheader("📝 Distribuição de aulas da disciplina")
st.caption(
    "Informe o **Nº de aulas no dia**. A quantidade de dias letivos é "
    "puxada automaticamente da planilha de Distribuição de Dias Letivos."
)

fid = _file_id(arquivo_dist)
chave_editor = f"config_{etapa}_{ano_letivo}_{fid}"

if chave_editor not in st.session_state:
    linhas = []
    for dia in DIAS_SEMANA:
        qtd_dias = len(etapas[etapa].get(dia, []))
        linhas.append({
            "Dias da semana": dia,
            "Nº de aulas no dia": PADRAO_AULAS[dia],
            "Total de dias da semana na etapa": qtd_dias,
            "Total de aulas": PADRAO_AULAS[dia] * qtd_dias,
        })
    st.session_state[chave_editor] = pd.DataFrame(linhas)

df_config = st.data_editor(
    st.session_state[chave_editor],
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    key=f"editor_{chave_editor}",
    column_config={
        "Dias da semana": st.column_config.TextColumn(
            "Dias da semana", disabled=True, width="medium"),
        "Nº de aulas no dia": st.column_config.NumberColumn(
            "Nº de aulas no dia", min_value=0, max_value=20, step=1,
            width="small", help="✏️ Quantas aulas desta disciplina nesse dia."),
        "Total de dias da semana na etapa": st.column_config.NumberColumn(
            "Total de dias da semana na etapa", disabled=True, width="large",
            help="🔒 Vem automaticamente da planilha de distribuição."),
        "Total de aulas": st.column_config.NumberColumn(
            "Total de aulas", disabled=True, width="small",
            help="🔒 Nº de aulas no dia × total de dias da semana."),
    },
)

# recalcula
df_config["Nº de aulas no dia"] = (
    pd.to_numeric(df_config["Nº de aulas no dia"], errors="coerce")
      .fillna(0).astype(int))
df_config["Total de dias da semana na etapa"] = (
    pd.to_numeric(df_config["Total de dias da semana na etapa"], errors="coerce")
      .fillna(0).astype(int))
df_config["Total de aulas"] = (
    df_config["Nº de aulas no dia"] * df_config["Total de dias da semana na etapa"])
st.session_state[chave_editor] = df_config

# totais (só dias efetivamente usados contam como "letivos")
total_dias = int(df_config.loc[df_config["Nº de aulas no dia"] > 0,
                               "Total de dias da semana na etapa"].sum())
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

st.dataframe(pd.concat([df_config, df_total], ignore_index=True),
             use_container_width=True, hide_index=True)

aulas_por_dia = dict(zip(df_config["Dias da semana"], df_config["Nº de aulas no dia"]))

# ------------------------------------------------------------
# Lista de datas esperadas
# ------------------------------------------------------------
with st.expander("📅 Ver todos os dias letivos considerados"):
    df_datas = montar_dias_esperados(etapas, etapa, aulas_por_dia)
    if df_datas.empty:
        st.warning("Nenhuma data encontrada.")
    else:
        df_datas = df_datas.copy()
        df_datas["Data"] = pd.to_datetime(df_datas["data"]).dt.strftime("%d/%m/%Y")
        df_datas = df_datas[["Data", "dia_semana", "aulas_esperadas"]]
        df_datas.columns = ["Data", "Dia da semana", "Aulas esperadas"]
        st.dataframe(df_datas, use_container_width=True, hide_index=True)

# ------------------------------------------------------------
# Conferência
# ------------------------------------------------------------
st.markdown("---")
st.header("🔎 3. Conferência")

if st.button("▶️ RODAR CONFERÊNCIA", type="primary", use_container_width=True):
    resultado = conferir(etapas, etapa, aulas_por_dia, df_diario)
    if resultado is None:
        st.warning("Não existem dias letivos configurados para essa etapa.")
        st.stop()
    st.session_state["ultimo_resultado"] = resultado

resultado = st.session_state.get("ultimo_resultado")

if resultado is not None:
    st.markdown("---")
    esperado = resultado["total_esperado"]
    lancado = resultado["total_lancado"]
    diferenca = lancado - esperado

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("📚 Aulas previstas", esperado)
    r2.metric("✏️ Aulas lançadas", lancado)
    r3.metric("📌 Diferença", diferenca)
    r4.metric("❗ Dias faltantes", len(resultado["faltantes"]))

    if diferenca == 0 and resultado["faltantes"].empty and resultado["divergentes"].empty:
        st.success("🎉 Conferência fechada: quantidade lançada corresponde às aulas previstas.")
    else:
        st.warning("⚠️ A conferência encontrou diferenças. Veja os detalhes abaixo.")

    st.subheader("🚨 Dias previstos sem lançamento")
    faltantes = resultado["faltantes"].copy()
    if faltantes.empty:
        st.success("Nenhum dia letivo previsto está sem lançamento.")
    else:
        faltantes["Data"] = pd.to_datetime(faltantes["data"]).dt.strftime("%d/%m/%Y")
        faltantes = faltantes[["Data", "dia_semana", "aulas_esperadas"]]
        faltantes.columns = ["Data", "Dia da semana", "Aulas esperadas"]
        st.error(f"🔴 **{len(faltantes)} dia(s)** sem lançamento.")
        st.dataframe(faltantes, use_container_width=True, hide_index=True)

    st.subheader("🔀 Divergência no número de aulas")
    divergentes = resultado["divergentes"].copy()
    if divergentes.empty:
        st.success("Nenhum lançamento com quantidade diferente do esperado.")
    else:
        divergentes["Data"] = pd.to_datetime(divergentes["data"]).dt.strftime("%d/%m/%Y")
        divergentes = divergentes[["Data", "aulas_esperadas", "aulas", "conteudo"]]
        divergentes.columns = ["Data", "Aulas esperadas", "Aulas lançadas", "Conteúdo"]
        st.dataframe(divergentes, use_container_width=True, hide_index=True)

    st.subheader("📍 Lançamentos em dias não previstos")
    extras = resultado["extras"].copy()
    if extras.empty:
        st.success("Nenhum lançamento fora dos dias previstos.")
    else:
        extras["Data"] = pd.to_datetime(extras["data"]).dt.strftime("%d/%m/%Y")
        extras = extras[["Data", "aulas", "conteudo"]]
        extras.columns = ["Data", "Aulas lançadas", "Conteúdo"]
        st.dataframe(extras, use_container_width=True, hide_index=True)

    st.subheader("📊 Resumo por mês")
    resumo = resultado["resumo"].copy()
    resumo.columns = ["Mês", "Dias previstos", "Aulas previstas",
                      "Aulas lançadas", "Diferença"]
    st.dataframe(resumo, use_container_width=True, hide_index=True)

    with st.expander("📖 Ver lançamentos considerados no diário"):
        diario = resultado["diario_etapa"].copy()
        if diario.empty:
            st.info("Nenhum lançamento dentro do período da etapa.")
        else:
            diario["Data"] = pd.to_datetime(diario["data"]).dt.strftime("%d/%m/%Y")
            diario = diario[["Data", "aulas", "conteudo"]]
            diario.columns = ["Data", "Aulas", "Conteúdo"]
            st.dataframe(diario, use_container_width=True, hide_index=True)

    # exportação
    st.subheader("⬇️ Exportar resultado")
    linhas_exportacao = []
    for _, row in resultado["esperado"].iterrows():
        dt = row["data"]
        lanc = resultado["diario_etapa"][resultado["diario_etapa"]["data"] == dt]
        if lanc.empty:
            aulas_lancadas, conteudo = 0, ""
        else:
            aulas_lancadas = int(lanc["aulas"].sum())
            conteudo = " | ".join(lanc["conteudo"].astype(str))
        if aulas_lancadas == 0:
            status = "FALTANTE"
        elif aulas_lancadas != int(row["aulas_esperadas"]):
            status = "DIVERGENTE"
        else:
            status = "OK"
        linhas_exportacao.append({
            "Data": dt.strftime("%d/%m/%Y"),
            "Dia da semana": row["dia_semana"],
            "Aulas esperadas": int(row["aulas_esperadas"]),
            "Aulas lançadas": aulas_lancadas,
            "Diferença": aulas_lancadas - int(row["aulas_esperadas"]),
            "Status": status,
            "Conteúdo": conteudo,
        })
    df_exp = pd.DataFrame(linhas_exportacao)
    csv = df_exp.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        "📥 Baixar conferência completa (CSV)",
        data=csv,
        file_name=f"conferencia_{etapa.replace(' ', '_')}_{ano_letivo}.csv",
        mime="text/csv",
        use_container_width=True,
    )

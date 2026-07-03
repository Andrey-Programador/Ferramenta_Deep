import io
import os
import tempfile
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# CONFIGURAÇÃO GERAL
# =========================================================

st.set_page_config(
    page_title="Sistema de Automações",
    layout="wide"
)

st.title("Sistema Completo de Automações")
st.caption("Filtros, relatórios, exportações e acompanhamento diário de produtividade.")


# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def normalize_text(texto):
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join([c for c in texto if not unicodedata.combining(c)])
    texto = " ".join(texto.split())
    return texto


def agora_sao_paulo():
    return datetime.now(ZoneInfo("America/Sao_Paulo"))


def converter_data_brasil(serie):
    return pd.to_datetime(
        serie,
        errors="coerce",
        dayfirst=True
    ).dt.date


def formatar_data_brasil(valor):
    if pd.isna(valor):
        return ""

    try:
        return pd.to_datetime(valor, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return str(valor)


def encontrar_coluna(df, palavras):
    for col in df.columns:
        col_norm = normalize_text(col)
        if all(normalize_text(p) in col_norm for p in palavras):
            return col
    return None


def safe_sheet_name(nome):
    nome = str(nome)

    for c in ['\\', '/', '*', '[', ']', ':', '?']:
        nome = nome.replace(c, "")

    return nome[:30] if nome else "ABA"


def read_excel_any(uploaded_file, dtype=None):
    uploaded_file.seek(0)
    ext = os.path.splitext(uploaded_file.name)[1].lower()

    if ext == ".xls":
        return pd.read_excel(uploaded_file, dtype=dtype, engine="xlrd")

    return pd.read_excel(uploaded_file, dtype=dtype, engine="openpyxl")


def excel_bytes_from_wb(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def excel_value(valor):
    if isinstance(valor, tuple):
        if len(valor) == 1:
            return valor[0]
        return " - ".join(str(v) for v in valor)

    if pd.isna(valor):
        return ""

    return valor


def aplicar_bordas_e_larguras(ws, max_width=55):
    from openpyxl.styles import Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue

            if cell.value is not None:
                cell.border = borda

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0

        for cell in col_cells:
            if isinstance(cell, MergedCell):
                continue

            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, max_width)


# =========================================================
# CLASSIFICAÇÕES
# =========================================================

def classificar_periodo(hora):
    try:
        hora = int(hora)
    except Exception:
        return "FORA DO PERÍODO"

    if 7 <= hora <= 12:
        return "MANHÃ"

    if 13 <= hora <= 18:
        return "TARDE"

    return "FORA DO PERÍODO"


def classificar_faixa_horario_excel(hora):
    try:
        hora = int(hora)
    except Exception:
        return "FORA DO PERÍODO"

    if 7 <= hora <= 12:
        return "MANHÃ"

    if 13 <= hora <= 18:
        return "TARDE"

    return "FORA DO PERÍODO"


def classificar_tipo_atendimento(valor):
    v = normalize_text(valor)

    if "ADESAO" in v:
        return "ADESÕES"

    if "MORADOR AUSENTE" in v or "AUSENTE" in v:
        return "AUSENTES"

    if "RECUSA" in v:
        return "RECUSAS"

    if "AGENDAMENTO" in v or "AGEND" in v:
        return "AGENDAMENTOS"

    if "VAGO" in v:
        return "IMOVEIS VAGOS"

    if "DEMOLIDO" in v:
        return "DEMOLIDO"

    if "ABANDONADO" in v:
        return "ABANDONADO"

    return "OUTROS"


# =========================================================
# FILTRO DE ADESÕES
# =========================================================

def filtrar_adesoes_realizadas(df, col_tipo):
    if not col_tipo:
        return df.iloc[0:0].copy()

    tipo_norm = df[col_tipo].astype(str).apply(normalize_text)

    df_filtrado = df[
        tipo_norm.str.contains("ADESAO REALIZADA", na=False)
    ].copy()

    df_filtrado = df_filtrado.reset_index(drop=True)

    return df_filtrado


def gerar_mensagem_resultado_adesoes(df_adesoes):
    if df_adesoes is None or df_adesoes.empty:
        return "Resultado sobre adesões:\n\nNenhuma adesão encontrada no arquivo enviado."

    total_adesoes = len(df_adesoes)

    col_data = encontrar_coluna(df_adesoes, ["DATA", "REGISTRO"])
    col_asro = encontrar_coluna(df_adesoes, ["ASRO"])
    col_novo = encontrar_coluna(df_adesoes, ["NOVO"])

    data_texto = "Na data analisada"

    if col_data:
        datas = pd.to_datetime(
            df_adesoes[col_data],
            errors="coerce",
            dayfirst=True
        )

        datas_validas = datas.dropna()

        if not datas_validas.empty:
            data_min = datas_validas.min().strftime("%d/%m/%Y")
            data_max = datas_validas.max().strftime("%d/%m/%Y")

            if data_min == data_max:
                data_texto = f"No dia {data_min}"
            else:
                data_texto = f"No período de {data_min} a {data_max}"

    novos_clientes = 0

    if col_novo:
        novos_clientes = (
            df_adesoes[col_novo]
            .astype(str)
            .str.strip()
            .str.upper()
            .eq("SIM")
            .sum()
        )

    mensagem = "Resultado sobre adesões:\n\n"
    mensagem += (
        f"{data_texto} teve o total de {total_adesoes} adesões, "
        f"sendo {novos_clientes} novos clientes.\n\n"
    )

    if col_asro:
        for asro, dados_asro in sorted(df_adesoes.groupby(col_asro), key=lambda x: str(x[0])):
            total_asro = len(dados_asro)

            novos_asro = 0

            if col_novo:
                novos_asro = (
                    dados_asro[col_novo]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    .eq("SIM")
                    .sum()
                )

            mensagem += (
                f"{asro}: {total_asro} adesões, "
                f"sendo {novos_asro} novos clientes.\n"
            )

    return mensagem


def processar_filtro_adesoes(uploaded_file):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
    from openpyxl.worksheet.table import Table, TableStyleInfo

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.astype(str).str.strip()

    col_tipo = encontrar_coluna(df, ["TIPO", "ATENDIMENTO"])

    if not col_tipo:
        raise ValueError("Coluna 'Tipo de atendimento' não encontrada.")

    df_adesoes = filtrar_adesoes_realizadas(df, col_tipo)

    df_final = pd.DataFrame()

    for coluna in df_adesoes.columns:
        nome = str(coluna).strip().lower()

        if "link backoffice" in nome:
            df_final["Link Backoffice"] = df_adesoes[coluna]

        elif "code deep" in nome:
            df_final["Code Deep"] = df_adesoes[coluna]

        elif "data do registro" in nome:
            df_final["Data do registro"] = df_adesoes[coluna]

        elif nome == "asro":
            df_final["ASRO"] = df_adesoes[coluna]

        elif nome == "nome completo:":
            df_final["Cliente"] = df_adesoes[coluna]

        elif nome in ["é novo cliente?", "e novo cliente?"]:
            df_final["É novo cliente?"] = df_adesoes[coluna]

        elif nome in ["situação backoffice", "situacao backoffice"]:
            df_final["Backoffice"] = df_adesoes[coluna]

        elif nome == "tipo de atendimento":
            df_final["Tipo de atendimento"] = df_adesoes[coluna]

    colunas_finais = [
        "Link Backoffice",
        "Code Deep",
        "Data do registro",
        "ASRO",
        "É novo cliente?",
        "Cliente",
        "Backoffice",
        "Tipo de atendimento"
    ]

    for col in colunas_finais:
        if col not in df_final.columns:
            df_final[col] = ""

    df_final = df_final[colunas_finais].reset_index(drop=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "RELATORIO"

    header_fill = PatternFill(start_color="4B0082", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for col_id, col_name in enumerate(df_final.columns, start=1):
        cell = ws.cell(row=1, column=col_id, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = borda

    for row_idx, row in enumerate(df_final.itertuples(index=False), start=2):
        for col_id, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_id, value=excel_value(value))
            cell.border = borda

    if len(df_final) > 0:
        ultima_linha = len(df_final) + 1

        tabela = Table(
            displayName="TabelaAdesoes",
            ref=f"A1:H{ultima_linha}"
        )

        estilo = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False
        )

        tabela.tableStyleInfo = estilo
        ws.add_table(tabela)

    aplicar_bordas_e_larguras(ws)

    mensagem = gerar_mensagem_resultado_adesoes(df_adesoes)

    return (
        excel_bytes_from_wb(wb),
        f"Relatorio_Adesoes_{agora_sao_paulo().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mensagem
    )


# =========================================================
# TERMO DE DOAÇÃO
# =========================================================

def processar_termo_doacao(uploaded_file, logo_file=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.drawing.image import Image

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.astype(str).str.strip().str.upper()

    def col_por_palavras(palavras):
        return next((c for c in df.columns if all(p in c for p in palavras)), None)

    col_code = col_por_palavras(["CODE"])
    col_nome = col_por_palavras(["NOME", "COMPLETO"])
    col_novo = col_por_palavras(["NOVO"])
    col_asro = col_por_palavras(["ASRO"])
    col_end = col_por_palavras(["ENDERE"])
    col_comp = col_por_palavras(["COMPLEMENTO"])

    if not col_code or not col_nome or not col_novo:
        raise ValueError("Colunas obrigatórias não encontradas: CODE, NOME COMPLETO e NOVO.")

    df[col_novo] = df[col_novo].astype(str).str.upper().str.strip()
    df = df[df[col_novo] == "SIM"]

    df_saida = pd.DataFrame()
    df_saida["CODE DEEP"] = df[col_code]
    df_saida["NOME COMPLETO"] = df[col_nome]
    df_saida["ASRO"] = df[col_asro] if col_asro else ""
    df_saida["ENDEREÇO"] = df[col_end] if col_end else ""
    df_saida["COMPLEMENTO"] = df[col_comp] if col_comp else ""

    df_saida = df_saida.sort_values(by="NOME COMPLETO")

    wb = Workbook()
    wb.remove(wb.active)

    periodo = agora_sao_paulo().strftime("%d-%m-%Y")

    ranking = (
        df_saida
        .groupby("ASRO")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    ws_rank = wb.create_sheet("RANKING")
    ws_rank.append(["ASRO", "TOTAL"])

    for _, r in ranking.iterrows():
        ws_rank.append([r["ASRO"], r["TOTAL"]])

    aplicar_bordas_e_larguras(ws_rank)

    temp_logo_path = None

    if logo_file is not None:
        suffix = os.path.splitext(logo_file.name)[1] or ".png"

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(logo_file.read())
            temp_logo_path = tmp.name

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for asro, dados in sorted(df_saida.groupby("ASRO"), key=lambda x: str(x[0])):
        ws = wb.create_sheet(title=safe_sheet_name(f"{asro} - {periodo}"))

        logo_path = None

        if temp_logo_path and os.path.exists(temp_logo_path):
            logo_path = temp_logo_path
        elif os.path.exists("logo.png"):
            logo_path = "logo.png"

        if logo_path:
            try:
                img = Image(logo_path)
                img.width = 120
                img.height = 40
                ws.add_image(img, "A1")
            except Exception:
                pass

        ws.merge_cells("A1:E1")
        ws["A1"] = "TERMO DE DOAÇÃO DE PADRÃO"
        ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="0DB39E", fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center")

        ws.merge_cells("A2:E2")
        ws["A2"] = f"Período: {periodo}"

        headers = ["CODE DEEP", "ASRO", "NOME COMPLETO", "ENDEREÇO", "COMPLEMENTO"]

        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="0DB39E", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = borda

        for i, row in enumerate(dados.itertuples(index=False), start=5):
            valores = [row[0], row[2], row[1], row[3], row[4]]

            for col, val in enumerate(valores, start=1):
                ws.cell(row=i, column=col, value=excel_value(val)).border = borda

        aplicar_bordas_e_larguras(ws)

    if temp_logo_path and os.path.exists(temp_logo_path):
        try:
            os.remove(temp_logo_path)
        except Exception:
            pass

    return (
        excel_bytes_from_wb(wb),
        f"resultado_{agora_sao_paulo().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )


# =========================================================
# RELATÓRIO DE AGENTES
# =========================================================

def processar_relatorio_agentes(uploaded_file):
    from openpyxl import Workbook

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.astype(str).str.strip()

    col_asro = encontrar_coluna(df, ["ASRO"])
    col_agente = encontrar_coluna(df, ["NOME", "AGENTE"])
    col_tipo = encontrar_coluna(df, ["TIPO", "ATENDIMENTO"])
    col_data = encontrar_coluna(df, ["DATA"])

    faltando = []

    if not col_asro:
        faltando.append("ASRO")

    if not col_agente:
        faltando.append("Nome do agente")

    if not col_tipo:
        faltando.append("Tipo de atendimento")

    if not col_data:
        faltando.append("Data")

    if faltando:
        raise ValueError("Colunas obrigatórias não encontradas: " + ", ".join(faltando))

    df[col_data] = converter_data_brasil(df[col_data])
    df["TIPO_CLASS"] = df[col_tipo].apply(classificar_tipo_atendimento)

    wb = Workbook()
    wb.remove(wb.active)

    for asro, dados_asro in df.groupby(col_asro):
        ws = wb.create_sheet(safe_sheet_name(asro))

        tabela = (
            dados_asro
            .groupby([col_data, col_agente, "TIPO_CLASS"])
            .size()
            .reset_index(name="TOTAL")
        )

        ws.append(["DATA", "AGENTE", "TIPO", "TOTAL"])

        for _, r in tabela.iterrows():
            ws.append([
                formatar_data_brasil(r[col_data]),
                r[col_agente],
                r["TIPO_CLASS"],
                int(r["TOTAL"])
            ])

        aplicar_bordas_e_larguras(ws)

    return (
        excel_bytes_from_wb(wb),
        f"relatorio_agentes_{agora_sao_paulo().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )


# =========================================================
# RELATÓRIO ENVIO / VISITAS / ADESÕES
# =========================================================

def preparar_base_ultima_visita(df, col_code, col_tipo):
    df = df.copy()

    df["IMOVEL"] = (
        df[col_code]
        .astype(str)
        .str.split("-")
        .str[0]
        .str.strip()
    )

    ordem = (
        df[col_code]
        .astype(str)
        .str.split("-")
        .str[1]
    )

    df["ORDEM_VISITA"] = pd.to_numeric(ordem, errors="coerce").fillna(0).astype(int)
    df["TIPO_CLASS"] = df[col_tipo].apply(classificar_tipo_atendimento)

    df = df.sort_values(by=["IMOVEL", "ORDEM_VISITA"])

    df_ultima_visita = df.drop_duplicates(
        subset=["IMOVEL"],
        keep="last"
    ).copy()

    return df, df_ultima_visita


def calcular_metricas_envio(base_visitas, base_ultima_visita, col_data):
    visitas_totais = len(base_visitas)
    imoveis_visitados = base_visitas["IMOVEL"].nunique()
    dias = base_visitas[col_data].dropna().nunique()
    media = round(visitas_totais / dias) if dias else 0

    adesoes = int((base_ultima_visita["TIPO_CLASS"] == "ADESÕES").sum())
    ausentes = int((base_ultima_visita["TIPO_CLASS"] == "AUSENTES").sum())
    recusas = int((base_ultima_visita["TIPO_CLASS"] == "RECUSAS").sum())
    agendamentos = int((base_ultima_visita["TIPO_CLASS"] == "AGENDAMENTOS").sum())

    def pct(valor):
        return valor / imoveis_visitados if imoveis_visitados else 0

    return {
        "Visitas totais": visitas_totais,
        "Imóveis visitados": imoveis_visitados,
        "Média visitas diárias": media,
        "Dias trabalhados": dias,
        "Moradores ausentes": ausentes,
        "% Moradores ausentes": pct(ausentes),
        "Adesões": adesoes,
        "% Adesões": pct(adesoes),
        "Recusas": recusas,
        "% Recusas": pct(recusas),
        "Agendamentos": agendamentos,
        "% Agendamentos": pct(agendamentos),
    }


def processar_relatorio_envio_visitas_adesoes(uploaded_file):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.astype(str).str.strip()

    col_asro = encontrar_coluna(df, ["ASRO"])
    col_tipo = encontrar_coluna(df, ["TIPO", "ATENDIMENTO"])
    col_agente = encontrar_coluna(df, ["NOME", "AGENTE"])
    col_data = encontrar_coluna(df, ["DATA", "REGISTRO"]) or encontrar_coluna(df, ["DATA"])
    col_code = encontrar_coluna(df, ["CODE"])

    faltando = []

    if not col_asro:
        faltando.append("ASRO")

    if not col_tipo:
        faltando.append("Tipo de atendimento")

    if not col_agente:
        faltando.append("Nome do agente")

    if not col_data:
        faltando.append("Data do registro")

    if not col_code:
        faltando.append("Code Deep")

    if faltando:
        raise ValueError("Colunas obrigatórias não encontradas: " + ", ".join(faltando))

    df[col_data] = converter_data_brasil(df[col_data])

    df_base, df_ultima = preparar_base_ultima_visita(df, col_code, col_tipo)

    wb = Workbook()
    wb.remove(wb.active)

    verde = PatternFill(start_color="006400", fill_type="solid")
    verde_medio = PatternFill(start_color="0DB39E", fill_type="solid")
    branco = Font(color="FFFFFF", bold=True)

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    def criar_resumo(ws, titulo, base_visitas, base_ultima_visita):
        m = calcular_metricas_envio(base_visitas, base_ultima_visita, col_data)

        ws.append([titulo])
        ws.append([])
        ws.append(["Indicador", "Valor", "Percentual"])

        linhas = [
            ("Visitas totais", m["Visitas totais"], ""),
            ("Imóveis visitados", m["Imóveis visitados"], ""),
            ("Média visitas diárias", m["Média visitas diárias"], ""),
            ("Dias trabalhados", m["Dias trabalhados"], ""),
            ("Moradores ausentes", m["Moradores ausentes"], m["% Moradores ausentes"]),
            ("Adesões", m["Adesões"], m["% Adesões"]),
            ("Recusas", m["Recusas"], m["% Recusas"]),
            ("Agendamentos", m["Agendamentos"], m["% Agendamentos"]),
        ]

        for item in linhas:
            ws.append(list(item))

        ws.merge_cells("A1:C1")
        ws["A1"].fill = verde
        ws["A1"].font = branco
        ws["A1"].alignment = Alignment(horizontal="center")

        for cell in ws[3]:
            cell.fill = verde_medio
            cell.font = branco
            cell.alignment = Alignment(horizontal="center")

        for row in range(8, 12):
            ws.cell(row=row, column=3).number_format = "0.00%"

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.border = borda

        inicio = 14

        dados_grafico = pd.DataFrame([
            ["Adesões", m["Adesões"]],
            ["Moradores ausentes", m["Moradores ausentes"]],
            ["Recusas", m["Recusas"]],
            ["Agendamentos", m["Agendamentos"]],
        ], columns=["Indicador", "Total"])

        dados_grafico = dados_grafico.sort_values("Total", ascending=False)

        ws.cell(row=inicio, column=1, value="Indicador")
        ws.cell(row=inicio, column=2, value="Total")

        for idx, linha in enumerate(dados_grafico.itertuples(index=False), start=inicio + 1):
            ws.cell(row=idx, column=1, value=linha[0])
            ws.cell(row=idx, column=2, value=int(linha[1]))

        chart = BarChart()
        chart.title = titulo
        chart.y_axis.title = "Quantidade"
        chart.x_axis.title = "Indicadores"

        data_ref = Reference(ws, min_col=2, min_row=inicio, max_row=inicio + 4)
        cats_ref = Reference(ws, min_col=1, min_row=inicio + 1, max_row=inicio + 4)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True

        ws.add_chart(chart, "E3")

        aplicar_bordas_e_larguras(ws)

    ws_geral = wb.create_sheet("GERAL")

    criar_resumo(
        ws_geral,
        "RELATÓRIO DE ADESÕES, VISITAS E ÁREAS CORRELATAS",
        df_base,
        df_ultima
    )

    for asro, dados_asro in sorted(df_base.groupby(col_asro), key=lambda x: str(x[0])):
        ws = wb.create_sheet(safe_sheet_name(asro))

        imoveis_asro = dados_asro["IMOVEL"].unique()
        ultima_asro = df_ultima[df_ultima["IMOVEL"].isin(imoveis_asro)].copy()

        criar_resumo(
            ws,
            f"RELATÓRIO - {asro}",
            dados_asro,
            ultima_asro
        )

        linha = 22

        headers = [
            "Agente",
            "Visitas totais",
            "Imóveis visitados",
            "Adesões",
            "% Adesões",
            "Ausentes",
            "% Ausentes",
            "Recusas",
            "% Recusas",
            "Agendamentos",
            "% Agendamentos"
        ]

        for c, h in enumerate(headers, start=1):
            ws.cell(row=linha, column=c, value=h)

        linha += 1

        for agente, dados_agente in dados_asro.groupby(col_agente):
            imoveis_agente = dados_agente["IMOVEL"].unique()
            ultima_agente = ultima_asro[ultima_asro["IMOVEL"].isin(imoveis_agente)].copy()

            ma = calcular_metricas_envio(dados_agente, ultima_agente, col_data)

            valores = [
                excel_value(agente),
                ma["Visitas totais"],
                ma["Imóveis visitados"],
                ma["Adesões"],
                ma["% Adesões"],
                ma["Moradores ausentes"],
                ma["% Moradores ausentes"],
                ma["Recusas"],
                ma["% Recusas"],
                ma["Agendamentos"],
                ma["% Agendamentos"]
            ]

            for c, v in enumerate(valores, start=1):
                cell = ws.cell(row=linha, column=c, value=excel_value(v))

                if c in [5, 7, 9, 11]:
                    cell.number_format = "0.00%"

            linha += 1

        aplicar_bordas_e_larguras(ws)

    return (
        excel_bytes_from_wb(wb),
        f"Relatorio_Adesoes_Visitas_Areas_{agora_sao_paulo().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )


# =========================================================
# ACOMPANHAMENTO DIÁRIO
# =========================================================

def preparar_acompanhamento(arquivos):
    bases = []

    for arquivo in arquivos:
        df = read_excel_any(arquivo)
        df.columns = df.columns.astype(str).str.strip()
        df["ARQUIVO_ORIGEM"] = arquivo.name
        bases.append(df)

    if not bases:
        raise ValueError("Nenhum arquivo enviado.")

    df = pd.concat(bases, ignore_index=True)

    col_asro = encontrar_coluna(df, ["ASRO"])
    col_agente = encontrar_coluna(df, ["NOME", "AGENTE"])
    col_data = encontrar_coluna(df, ["DATA", "REGISTRO"]) or encontrar_coluna(df, ["DATA"])
    col_horario = encontrar_coluna(df, ["HORARIO", "REGISTRO"]) or encontrar_coluna(df, ["HORA", "REGISTRO"])
    col_tipo = encontrar_coluna(df, ["TIPO", "ATENDIMENTO"])

    faltando = []

    if not col_asro:
        faltando.append("ASRO")

    if not col_agente:
        faltando.append("Nome do agente")

    if not col_data:
        faltando.append("Data do registro")

    if not col_horario:
        faltando.append("Horário de registro")

    if not col_tipo:
        faltando.append("Tipo de atendimento")

    if faltando:
        raise ValueError("Colunas obrigatórias não encontradas: " + ", ".join(faltando))

    df["DATA_REGISTRO_TRATADA"] = converter_data_brasil(df[col_data])

    horario_str = df[col_horario].astype(str).str.strip()

    hora_dt = pd.to_datetime(
        horario_str,
        errors="coerce",
        dayfirst=True
    ).dt.hour

    hora_num = pd.to_numeric(
        horario_str.str.extract(r"(\d{1,2})", expand=False),
        errors="coerce"
    )

    df["HORA_EXTRAIDA"] = hora_dt.fillna(hora_num).astype("Int64")

    df["PERIODO"] = df["HORA_EXTRAIDA"].apply(classificar_periodo)

    df["ASRO"] = df[col_asro].astype(str).str.strip()
    df["AGENTE"] = df[col_agente].astype(str).str.strip()
    df["TIPO_ORIGINAL"] = df[col_tipo].astype(str).str.strip()
    df["TIPO_CLASS"] = df[col_tipo].apply(classificar_tipo_atendimento)

    cols = {
        "ASRO": col_asro,
        "Agente": col_agente,
        "Data": col_data,
        "Horário": col_horario,
        "Tipo de atendimento": col_tipo
    }

    return df, cols


def resumo_agente_acomp(base, incluir_asro=True):
    colunas = [
        "Agente",
        "Visitas totais",
        "Imóveis visitados",
        "Adesões",
        "Ausentes",
        "Recusas",
        "Agendamentos",
        "Imoveis vagos"
    ]

    if incluir_asro:
        colunas = ["ASRO"] + colunas

    if base.empty:
        return pd.DataFrame(columns=colunas)

    grupos = ["ASRO", "AGENTE"] if incluir_asro else "AGENTE"

    linhas = []

    for chave, dados in base.groupby(grupos):
        if incluir_asro:
            asro, agente = chave
        else:
            asro = None
            agente = excel_value(chave)

        item = {
            "Agente": excel_value(agente),
            "Visitas totais": len(dados),
            "Imóveis visitados": len(dados),
            "Adesões": int((dados["TIPO_CLASS"] == "ADESÕES").sum()),
            "Ausentes": int((dados["TIPO_CLASS"] == "AUSENTES").sum()),
            "Recusas": int((dados["TIPO_CLASS"] == "RECUSAS").sum()),
            "Agendamentos": int((dados["TIPO_CLASS"] == "AGENDAMENTOS").sum()),
            "Imoveis vagos": int(
                dados["TIPO_ORIGINAL"]
                .astype(str)
                .apply(lambda x: "VAGO" in normalize_text(x))
                .sum()
            ),
        }

        if incluir_asro:
            item = {
                "ASRO": excel_value(asro),
                **item
            }

        linhas.append(item)

    return pd.DataFrame(linhas).sort_values("Visitas totais", ascending=False)


def relatorio_final_simplificado(df_base):
    if df_base.empty:
        return pd.DataFrame(
            columns=[
                "ASRO",
                "Agente",
                "Período",
                "Horários",
                "Visitas",
                "Adesões",
                "Ausentes",
                "Recusas",
                "Agendamentos",
                "Imoveis vagos",
                "Principal atendimento"
            ]
        )

    linhas = []

    for (asro, agente, periodo), dados in df_base.groupby(["ASRO", "AGENTE", "PERIODO"]):
        tipos = dados["TIPO_ORIGINAL"].value_counts()
        principal = tipos.index[0] if len(tipos) else ""

        horarios = ", ".join(
            str(int(h))
            for h in sorted(dados["HORA_EXTRAIDA"].dropna().unique())
        )

        linhas.append({
            "ASRO": asro,
            "Agente": agente,
            "Período": periodo,
            "Horários": horarios,
            "Visitas": len(dados),
            "Adesões": int((dados["TIPO_CLASS"] == "ADESÕES").sum()),
            "Ausentes": int((dados["TIPO_CLASS"] == "AUSENTES").sum()),
            "Recusas": int((dados["TIPO_CLASS"] == "RECUSAS").sum()),
            "Agendamentos": int((dados["TIPO_CLASS"] == "AGENDAMENTOS").sum()),
            "Imoveis vagos": int(
                dados["TIPO_ORIGINAL"]
                .astype(str)
                .apply(lambda x: "VAGO" in normalize_text(x))
                .sum()
            ),
            "Principal atendimento": principal,
        })

    return pd.DataFrame(linhas).sort_values(
        ["ASRO", "Período", "Visitas"],
        ascending=[True, True, False]
    )


def montar_resumo_geral_acompanhamento(df_base, hora_extracao=None):
    resumo_asro = (
        df_base
        .groupby("ASRO")
        .agg(
            Visitas=("AGENTE", "size"),
            Agentes=("AGENTE", "nunique")
        )
        .reset_index()
    )

    resumo_asro["Média por agente"] = resumo_asro.apply(
        lambda r: round(r["Visitas"] / r["Agentes"]) if r["Agentes"] else 0,
        axis=1
    )

    resumo_asro = resumo_asro[
        ["ASRO", "Visitas", "Média por agente"]
    ].sort_values("Visitas", ascending=False)

    ranking = (
        df_base
        .groupby(["AGENTE", "ASRO"])
        .size()
        .reset_index(name="Visitas")
        .sort_values("Visitas", ascending=False)
        .reset_index(drop=True)
    )

    ranking.insert(0, "RANKING", ranking.index + 1)

    datas_formatadas = []

    for data in sorted(df_base["DATA_REGISTRO_TRATADA"].dropna().unique()):
        datas_formatadas.append(formatar_data_brasil(data))

    if hora_extracao is None:
        hora_extracao = agora_sao_paulo().strftime("%H:%M:%S")

    return {
        "datas": ", ".join(datas_formatadas),
        "extracao": hora_extracao,
        "asros_atuando": df_base["ASRO"].nunique(),
        "total_registros": len(df_base),
        "total_agentes": df_base["AGENTE"].nunique(),
        "resumo_asro": resumo_asro,
        "ranking": ranking,
    }


def montar_resumo_agente_por_faixa(dados_agente):
    if dados_agente.empty:
        return pd.DataFrame(
            columns=[
                "Período",
                "Visitas",
                "Adesões",
                "Ausentes",
                "Recusas",
                "Agendamentos",
                "Imoveis vagos"
            ]
        )

    dados_agente = dados_agente.copy()
    dados_agente["FAIXA_EXCEL"] = dados_agente["HORA_EXTRAIDA"].apply(classificar_faixa_horario_excel)

    linhas = []

    for faixa in ["MANHÃ", "TARDE", "FORA DO PERÍODO"]:
        bloco = dados_agente[dados_agente["FAIXA_EXCEL"] == faixa]

        if bloco.empty:
            continue

        linhas.append({
            "Período": faixa,
            "Visitas": len(bloco),
            "Adesões": int((bloco["TIPO_CLASS"] == "ADESÕES").sum()),
            "Ausentes": int((bloco["TIPO_CLASS"] == "AUSENTES").sum()),
            "Recusas": int((bloco["TIPO_CLASS"] == "RECUSAS").sum()),
            "Agendamentos": int((bloco["TIPO_CLASS"] == "AGENDAMENTOS").sum()),
            "Imoveis vagos": int(
                bloco["TIPO_ORIGINAL"]
                .astype(str)
                .apply(lambda x: "VAGO" in normalize_text(x))
                .sum()
            ),
        })

    linhas.append({
        "Período": "TOTAL",
        "Visitas": len(dados_agente),
        "Adesões": int((dados_agente["TIPO_CLASS"] == "ADESÕES").sum()),
        "Ausentes": int((dados_agente["TIPO_CLASS"] == "AUSENTES").sum()),
        "Recusas": int((dados_agente["TIPO_CLASS"] == "RECUSAS").sum()),
        "Agendamentos": int((dados_agente["TIPO_CLASS"] == "AGENDAMENTOS").sum()),
        "Imoveis vagos": int(
            dados_agente["TIPO_ORIGINAL"]
            .astype(str)
            .apply(lambda x: "VAGO" in normalize_text(x))
            .sum()
        ),
    })

    return pd.DataFrame(linhas)


def gerar_excel_acompanhamento(df_base):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    hora_extracao = agora_sao_paulo().strftime("%H:%M:%S")

    wb = Workbook()
    wb.remove(wb.active)

    verde = PatternFill(start_color="00B050", fill_type="solid")
    verde_escuro = PatternFill(start_color="006400", fill_type="solid")
    cinza = PatternFill(start_color="D9D9D9", fill_type="solid")
    cinza_claro = PatternFill(start_color="F2F2F2", fill_type="solid")

    branco = Font(color="FFFFFF", bold=True)
    fonte_titulo = Font(color="1F2937", bold=True, size=14)
    fonte_negrito = Font(bold=True)

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    def estilizar_range(ws):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.border = borda
                    cell.alignment = Alignment(vertical="center")

        for col_idx, col_cells in enumerate(ws.columns, start=1):
            max_len = 0

            for cell in col_cells:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 55)

    def escrever_df(ws, df, row, col, header_fill=cinza):
        for j, nome_col in enumerate(df.columns, start=col):
            cell = ws.cell(row=row, column=j, value=str(nome_col))
            cell.fill = header_fill
            cell.font = fonte_negrito
            cell.border = borda
            cell.alignment = Alignment(horizontal="center")

        for i, linha in enumerate(df.itertuples(index=False), start=row + 1):
            for j, valor in enumerate(linha, start=col):
                cell = ws.cell(row=i, column=j, value=excel_value(valor))
                cell.border = borda
                cell.fill = cinza_claro if i % 2 == 0 else PatternFill(fill_type=None)

        return row + len(df) + 2

    geral = montar_resumo_geral_acompanhamento(
        df_base,
        hora_extracao=hora_extracao
    )

    ws = wb.create_sheet("GERAL")

    ws.merge_cells("A1:E1")
    ws["A1"] = "ACOMPANHAMENTO DIÁRIO - PRODUTIVIDADE DOS AGENTES"
    ws["A1"].fill = verde
    ws["A1"].font = fonte_titulo
    ws["A1"].alignment = Alignment(horizontal="center")

    labels = [
        "Data(s) filtrada(s)",
        "Horário da extração",
        "ASROs atuando",
        "Total de registros",
        "Total de agentes"
    ]

    valores = [
        geral["datas"],
        hora_extracao,
        geral["asros_atuando"],
        geral["total_registros"],
        geral["total_agentes"]
    ]

    for idx, label in enumerate(labels, start=3):
        ws.cell(row=idx, column=1, value=label).font = fonte_negrito

        cell_valor = ws.cell(row=idx, column=2, value=valores[idx - 3])

        if label == "Horário da extração":
            cell_valor.number_format = "@"

    resumo_start = 10

    ws.cell(row=resumo_start, column=1, value="ASRO").fill = cinza
    ws.cell(row=resumo_start, column=2, value="Visitas").fill = cinza
    ws.cell(row=resumo_start, column=3, value="Média por agente").fill = cinza

    for c in range(1, 4):
        ws.cell(row=resumo_start, column=c).font = fonte_negrito
        ws.cell(row=resumo_start, column=c).alignment = Alignment(horizontal="center")

    linha = resumo_start + 1

    for _, r in geral["resumo_asro"].iterrows():
        ws.cell(row=linha, column=1, value=r["ASRO"])
        ws.cell(row=linha, column=2, value=int(r["Visitas"]))
        ws.cell(row=linha, column=3, value=int(r["Média por agente"]))
        linha += 1

    ranking_start = 10

    for c, label in zip(
        range(5, 9),
        ["RANKING", "AGENTE", "ASRO", "Visitas"]
    ):
        ws.cell(row=ranking_start, column=c, value=label).fill = cinza
        ws.cell(row=ranking_start, column=c).font = fonte_negrito
        ws.cell(row=ranking_start, column=c).alignment = Alignment(horizontal="center")

    linha_rank = ranking_start + 1

    for _, r in geral["ranking"].iterrows():
        ws.cell(row=linha_rank, column=5, value=int(r["RANKING"]))
        ws.cell(row=linha_rank, column=6, value=r["AGENTE"])
        ws.cell(row=linha_rank, column=7, value=r["ASRO"])
        ws.cell(row=linha_rank, column=8, value=int(r["Visitas"]))
        linha_rank += 1

    estilizar_range(ws)

    for asro, dados_asro in df_base.groupby("ASRO"):
        ws_asro = wb.create_sheet(safe_sheet_name(asro))

        ws_asro.merge_cells("A1:G1")
        ws_asro["A1"] = f"ACOMPANHAMENTO DIÁRIO - ASRO {asro}"
        ws_asro["A1"].fill = verde
        ws_asro["A1"].font = fonte_titulo
        ws_asro["A1"].alignment = Alignment(horizontal="center")

        ws_asro.cell(row=3, column=1, value="Total de registros").font = fonte_negrito
        ws_asro.cell(row=3, column=2, value="Total de agentes").font = fonte_negrito
        ws_asro.cell(row=4, column=1, value=len(dados_asro))
        ws_asro.cell(row=4, column=2, value=dados_asro["AGENTE"].nunique())

        ws_asro.cell(
            row=6,
            column=1,
            value="* MANHÃ - 07H ÀS 12H | TARDE - 13H ÀS 18H | FORA DO PERÍODO"
        )
        ws_asro.cell(row=6, column=1).font = fonte_negrito

        linha = 8

        for agente, dados_agente in sorted(
            dados_asro.groupby("AGENTE"),
            key=lambda x: str(x[0])
        ):
            ws_asro.merge_cells(
                start_row=linha,
                start_column=1,
                end_row=linha,
                end_column=7
            )

            cell_agente = ws_asro.cell(row=linha, column=1, value=str(agente))
            cell_agente.fill = verde_escuro
            cell_agente.font = branco
            cell_agente.alignment = Alignment(horizontal="left")

            linha += 1

            tabela_agente = montar_resumo_agente_por_faixa(dados_agente)
            linha = escrever_df(ws_asro, tabela_agente, linha, 1, header_fill=cinza)
            linha += 1

        estilizar_range(ws_asro)

    ws_base = wb.create_sheet("BASE TRATADA")
    ws_base["A1"] = "BASE TRATADA"
    ws_base["A1"].fill = verde
    ws_base["A1"].font = fonte_titulo

    base_export = df_base[
        [
            "ARQUIVO_ORIGEM",
            "DATA_REGISTRO_TRATADA",
            "HORA_EXTRAIDA",
            "PERIODO",
            "ASRO",
            "AGENTE",
            "TIPO_ORIGINAL",
            "TIPO_CLASS"
        ]
    ].copy()

    escrever_df(ws_base, base_export, 3, 1, header_fill=cinza)
    estilizar_range(ws_base)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return (
        output.getvalue(),
        f"Acompanhamento_Diario_{agora_sao_paulo().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )


# =========================================================
# INTERFACE EM ABAS
# =========================================================

tab_filtro, tab_termo, tab_agentes, tab_envio, tab_acompanhamento = st.tabs([
    "Filtro de Adesões",
    "Termo de Doação",
    "Relatório de Agentes",
    "Relatório Envio / Visitas / Adesões",
    "Acompanhamento diário",
])


# =========================================================
# ABA FILTRO DE ADESÕES
# =========================================================

with tab_filtro:
    st.header("Filtro de Adesões")

    st.info(
        """
        **Objetivo da rotina:**  
        Diariamente, os dados anexados no sistema no dia anterior são extraídos do arquivo bruto.

        O sistema filtra apenas as adesões realizadas, organiza as informações principais do cliente
        e gera um Excel padronizado para que o time de conferência possa validar os cadastros.

        Também é gerada automaticamente uma mensagem resumo para envio ao coordenador Caio,
        contendo o total de adesões, novos clientes e separação por ASRO.
        """
    )

    arquivo = st.file_uploader(
        "Selecione o Excel bruto",
        type=["xlsx", "xls"],
        key="filtro_adesoes"
    )

    if arquivo and st.button("Executar Filtro de Adesões", key="btn_filtro"):
        try:
            data, nome, mensagem = processar_filtro_adesoes(arquivo)

            st.success("Relatório gerado com sucesso!")

            st.download_button(
                "Baixar Excel",
                data=data,
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.subheader("Mensagem automática para envio")

            st.text_area(
                "Copie a mensagem abaixo para enviar ao Caio:",
                value=mensagem,
                height=220
            )

            st.download_button(
                "Baixar mensagem em TXT",
                data=mensagem.encode("utf-8"),
                file_name=f"Mensagem_Adesoes_{agora_sao_paulo().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )

        except Exception as e:
            st.error(f"Erro: {e}")


# =========================================================
# ABA TERMO DE DOAÇÃO
# =========================================================

with tab_termo:
    st.header("Termo de Doação")

    st.info(
        """
        **Objetivo da rotina:**  
        Semanalmente, o sistema gera o arquivo de Termo de Doação em formato padronizado e profissional.

        A rotina considera os registros de novos clientes da semana anterior,
        geralmente no período de terça-feira passada até a segunda-feira atual.

        O arquivo gerado deve ser enviado ao Caio.
        """
    )

    arquivo = st.file_uploader(
        "Selecione o Excel bruto",
        type=["xlsx"],
        key="termo_doacao"
    )

    logo = st.file_uploader(
        "Logo opcional",
        type=["png", "jpg", "jpeg"],
        key="termo_logo"
    )

    if arquivo and st.button("Gerar Termo de Doação", key="btn_termo"):
        try:
            data, nome = processar_termo_doacao(arquivo, logo)

            st.success("Termo gerado com sucesso!")

            st.download_button(
                "Baixar Excel",
                data=data,
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {e}")


# =========================================================
# ABA RELATÓRIO DE AGENTES
# =========================================================

with tab_agentes:
    st.header("Relatório Semanal dos Agentes")

    st.info(
        """
        **Objetivo da rotina:**  
        Sistema utilizado para extrair o arquivo bruto e gerar um Excel com a produtividade dos agentes.

        O relatório organiza a produção por ASRO, agente e tipo de atendimento.
        """
    )

    arquivo = st.file_uploader(
        "Selecione o Excel bruto",
        type=["xlsx", "xls"],
        key="rel_agentes"
    )

    if arquivo and st.button("Gerar Relatório de Agentes", key="btn_agentes"):
        try:
            data, nome = processar_relatorio_agentes(arquivo)

            st.success("Relatório gerado com sucesso!")

            st.download_button(
                "Baixar Excel",
                data=data,
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {e}")


# =========================================================
# ABA RELATÓRIO ENVIO / VISITAS / ADESÕES
# =========================================================

with tab_envio:
    st.header("Relatório Envio / Visitas / Adesões")

    st.info(
        """
        **Regra aplicada:**  
        - Visitas totais = total de linhas da planilha.  
        - Imóveis visitados = contagem distinta de Code Deep sem o número após o traço.  
        - Adesões, ausentes, recusas e agendamentos são calculados pela última visita do imóvel.  
        - Percentuais usam como base o total de imóveis visitados.
        """
    )

    arquivo = st.file_uploader(
        "Selecione o Excel bruto",
        type=["xlsx", "xls"],
        key="rel_envio"
    )

    if arquivo and st.button("Gerar Relatório Envio / Visitas / Adesões", key="btn_envio"):
        try:
            data, nome = processar_relatorio_envio_visitas_adesoes(arquivo)

            st.success("Relatório gerado com sucesso!")

            st.download_button(
                "Baixar Excel",
                data=data,
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {e}")


# =========================================================
# ABA ACOMPANHAMENTO DIÁRIO
# =========================================================

with tab_acompanhamento:
    st.header("Acompanhamento diário")

    st.info(
        """
        **Objetivo da rotina:**  
        Sistema usado para acompanhar diariamente a produtividade total da operação em campo.

        O acompanhamento considera todos os registros do dia selecionado,
        independentemente do horário de sincronização.

        **Períodos usados para análise visual:**  
        - MANHÃ: 07h às 12h  
        - TARDE: 13h às 18h  
        - FORA DO PERÍODO: demais horários  

        **Observação de sincronização:**  
        Os horários obrigatórios para sincronizarem os dados são:

        - 10h  
        - 12h antes do almoço  
        - 15h  
        - 18h final do dia, antes do celular bloquear  

        Esses horários são referência operacional, mas o sistema não exclui registros fora desses horários.
        """
    )

    arquivos = st.file_uploader(
        "Enviar arquivo bruto ou vários arquivos",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="upload_acompanhamento"
    )

    if arquivos:
        try:
            df_original, cols = preparar_acompanhamento(arquivos)

            st.success(
                f"Arquivos carregados: {len(arquivos)} | "
                f"Total de registros: {len(df_original)}"
            )

            st.caption(
                "Colunas detectadas: "
                + " | ".join([f"{k}: {v}" for k, v in cols.items()])
            )

            st.subheader("Filtros")

            f1, f2 = st.columns(2)

            asros = sorted(
                df_original["ASRO"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            datas = sorted(
                df_original["DATA_REGISTRO_TRATADA"]
                .dropna()
                .unique()
                .tolist()
            )

            with f1:
                asros_sel = st.multiselect(
                    "ASRO",
                    asros,
                    default=asros
                )

            with f2:
                datas_sel = st.multiselect(
                    "Data",
                    datas,
                    default=datas
                )

            df_dash = df_original[
                df_original["ASRO"].isin(asros_sel)
                & df_original["DATA_REGISTRO_TRATADA"].isin(datas_sel)
            ].copy()

            if df_dash.empty:
                st.warning("Nenhum dado encontrado com os filtros selecionados.")

            else:
                total_visitas = len(df_dash)
                agentes_unicos = df_dash["AGENTE"].nunique()

                total_manha = int((df_dash["PERIODO"] == "MANHÃ").sum())
                total_tarde = int((df_dash["PERIODO"] == "TARDE").sum())
                total_fora = int((df_dash["PERIODO"] == "FORA DO PERÍODO").sum())

                k1, k2, k3, k4, k5 = st.columns(5)

                k1.metric("Total de registros", total_visitas)
                k2.metric("Manhã", total_manha)
                k3.metric("Tarde", total_tarde)
                k4.metric("Fora do período", total_fora)
                k5.metric(
                    "Média/agente",
                    round(total_visitas / agentes_unicos) if agentes_unicos else 0
                )

                st.subheader("Painel")

                col_pizza, col_top = st.columns([1.25, 2])

                with col_pizza:
                    st.markdown("**Distribuição por período**")

                    periodo = (
                        df_dash
                        .groupby("PERIODO")
                        .size()
                        .reset_index(name="TOTAL")
                    )

                    fig_pie = px.pie(
                        periodo,
                        names="PERIODO",
                        values="TOTAL",
                        hole=0.35,
                        color_discrete_sequence=[
                            "#38bdf8",
                            "#22c55e",
                            "#94a3b8"
                        ]
                    )

                    fig_pie.update_traces(
                        textposition="inside",
                        textinfo="percent+label+value"
                    )

                    fig_pie.update_layout(
                        height=430,
                        margin=dict(l=10, r=10, t=40, b=10),
                        showlegend=True
                    )

                    st.plotly_chart(fig_pie, use_container_width=True)

                with col_top:
                    st.markdown("**Top agentes por visitas**")

                    top_agentes = resumo_agente_acomp(
                        df_dash,
                        incluir_asro=True
                    ).head(15)

                    fig_bar = px.bar(
                        top_agentes,
                        x="Visitas totais",
                        y="Agente",
                        color="ASRO",
                        orientation="h",
                        text="Visitas totais"
                    )

                    fig_bar.update_layout(
                        height=430,
                        yaxis={"categoryorder": "total ascending"}
                    )

                    st.plotly_chart(fig_bar, use_container_width=True)

                st.subheader("Divisão por período, ASRO e agente")

                for periodo_nome in ["MANHÃ", "TARDE", "FORA DO PERÍODO"]:
                    dados_periodo = df_dash[df_dash["PERIODO"] == periodo_nome]

                    st.markdown(f"### {periodo_nome}")

                    if dados_periodo.empty:
                        st.info("Sem registros nesse período.")
                        continue

                    for asro_nome in sorted(
                        dados_periodo["ASRO"]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    ):
                        dados_asro_periodo = dados_periodo[
                            dados_periodo["ASRO"] == asro_nome
                        ]

                        with st.expander(
                            f"ASRO {asro_nome} - {periodo_nome}",
                            expanded=True
                        ):
                            heat = (
                                dados_asro_periodo
                                .groupby(["AGENTE", "HORA_EXTRAIDA"])
                                .size()
                                .reset_index(name="VISITAS")
                            )

                            pivot = heat.pivot_table(
                                index="AGENTE",
                                columns="HORA_EXTRAIDA",
                                values="VISITAS",
                                fill_value=0
                            )

                            if not pivot.empty:
                                pivot = pivot.reindex(sorted(pivot.columns), axis=1)

                                fig_heat = px.imshow(
                                    pivot,
                                    text_auto=True,
                                    aspect="auto",
                                    color_continuous_scale=[
                                        "#f7fbff",
                                        "#c6dbef",
                                        "#6baed6",
                                        "#2171b5",
                                        "#08306b"
                                    ],
                                    labels=dict(
                                        x="Horário",
                                        y="Agente",
                                        color="Visitas"
                                    )
                                )

                                fig_heat.update_layout(
                                    height=max(280, 28 * len(pivot.index)),
                                    margin=dict(l=20, r=20, t=30, b=20)
                                )

                                st.plotly_chart(fig_heat, use_container_width=True)

                            st.dataframe(
                                resumo_agente_acomp(
                                    dados_asro_periodo,
                                    incluir_asro=False
                                ),
                                use_container_width=True
                            )

                st.subheader("Relatório final detalhado")

                st.dataframe(
                    relatorio_final_simplificado(df_dash),
                    use_container_width=True
                )

                st.subheader("Gerar arquivo Excel")

                if st.button("Gerar relatório final em Excel", key="btn_excel_acompanhamento"):
                    excel_data, nome_arquivo = gerar_excel_acompanhamento(df_dash)

                    st.success("Arquivo Excel gerado com sucesso!")

                    st.download_button(
                        "Baixar Excel do acompanhamento diário",
                        data=excel_data,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Erro ao gerar acompanhamento diário: {e}")

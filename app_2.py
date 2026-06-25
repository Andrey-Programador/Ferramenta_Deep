import io
import os
import tempfile
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(page_title="Sistema de Automações", layout="wide")

st.markdown(
    """
    <style>
        .stApp {
            background-color: 	#000000
            ;
        }
        h1, h2, h3 {
            color: #064e3b;
        }
        div[data-baseweb="tab-list"] button {
            font-size: 15px !important;
            font-weight: 700 !important;
        }
        .block-container {
            padding-top: 1.5rem;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("Sistema Completo de Automações")
st.caption("Interface web em Streamlit mantendo a lógica original dos seus scripts Python.")


# =========================
# FUNÇÕES AUXILIARES
# =========================
def excel_bytes_from_wb(workbook):
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def excel_bytes_from_df(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def normalize_text(texto):
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join([c for c in texto if not unicodedata.combining(c)])
    texto = " ".join(texto.split())
    return texto


def safe_sheet_name(nome):
    nome = str(nome)
    for c in ['\\', '/', '*', '[', ']', ':', '?']:
        nome = nome.replace(c, "")
    nome = nome[:30]
    return nome if nome else "ABA"


def read_excel_any(uploaded_file, dtype=None):
    uploaded_file.seek(0)
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext == ".xls":
        return pd.read_excel(uploaded_file, dtype=dtype, engine="xlrd")
    return pd.read_excel(uploaded_file, dtype=dtype, engine="openpyxl")


# =========================
# 1) FILTRO DE ADESÕES
# =========================
def processar_filtro_adesoes(uploaded_file):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Border, Side, Font, Alignment

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.str.strip()

    if "Tipo de atendimento" in df.columns:
        df = df[
            df["Tipo de atendimento"].astype(str).str.strip().str.upper()
            == "ADESÃO REALIZADA (MORADOR PRESENTE)"
        ]

    df_final = pd.DataFrame()

    for coluna in df.columns:
        nome = str(coluna).strip().lower()
        if "link backoffice" in nome:
            df_final["Link Backoffice"] = df[coluna]
        elif "code deep" in nome:
            df_final["Code Deep"] = df[coluna]
        elif "data do registro" in nome:
            df_final["Data do registro"] = df[coluna]
        elif nome == "asro":
            df_final["ASRO"] = df[coluna]
        elif nome == "nome completo:":
            df_final["Cliente"] = df[coluna]
        elif nome == "é novo cliente?" or nome == "e novo cliente?":
            df_final["É novo cliente?"] = df[coluna]
        elif nome == "situação backoffice" or nome == "situacao backoffice":
            df_final["Backoffice"] = df[coluna]
        elif nome == "tipo de atendimento":
            df_final["Tipo de atendimento"] = df[coluna]

    colunas_finais = [
        "Link Backoffice",
        "Code Deep",
        "Data do registro",
        "ASRO",
        "É novo cliente?",
        "Cliente",
        "Backoffice",
        "Tipo de atendimento",
    ]

    for coluna in colunas_finais:
        if coluna not in df_final.columns:
            df_final[coluna] = ""

    df_final = df_final[colunas_finais]

    wb = Workbook()
    ws = wb.active
    ws.title = "RELATORIO"

    header_fill = PatternFill(start_color="4B0082", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_id, col_name in enumerate(df_final.columns, start=1):
        cell = ws.cell(row=1, column=col_id)
        cell.value = col_name
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = borda

    for i, row in df_final.iterrows():
        for col_id, value in enumerate(row, start=1):
            cell = ws.cell(row=i + 2, column=col_id)
            cell.value = value
            cell.border = borda

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 5

    nome_arquivo = f"Relatorio_Adesoes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return excel_bytes_from_wb(wb), nome_arquivo


# =========================
# 2) TERMO DE DOAÇÃO
# =========================
def processar_termo_doacao(uploaded_file, logo_file=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.drawing.image import Image

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()

    def encontrar_coluna(palavras):
        colunas = [c for c in df.columns if all(p in c for p in palavras)]
        return colunas[0] if colunas else None

    col_code = encontrar_coluna(["CODE"])
    col_nome = encontrar_coluna(["NOME", "COMPLETO"])
    col_novo = encontrar_coluna(["NOVO"])
    col_asro = encontrar_coluna(["ASRO"])
    col_end = encontrar_coluna(["ENDERE"])
    col_comp = encontrar_coluna(["COMPLEMENTO"])

    if not col_code or not col_nome or not col_novo:
        raise ValueError("Colunas obrigatórias não encontradas")

    df[col_novo] = df[col_novo].astype(str).str.upper().str.strip()
    df = df[df[col_novo] == "SIM"]

    df_saida = pd.DataFrame()
    df_saida["CODE DEEP"] = df[col_code]
    df_saida["NOME COMPLETO"] = df[col_nome]
    df_saida["ASRO"] = df[col_asro] if col_asro else ""
    df_saida["ENDEREÇO"] = df[col_end] if col_end else ""

    if col_comp:
        col_temp = df.loc[:, col_comp]
        if isinstance(col_temp, pd.DataFrame):
            df_saida["COMPLEMENTO"] = col_temp.iloc[:, 0]
        else:
            df_saida["COMPLEMENTO"] = col_temp
    else:
        df_saida["COMPLEMENTO"] = ""

    periodo = datetime.now().strftime("%d-%m-%Y")
    df_saida = df_saida.sort_values(by="NOME COMPLETO")

    wb = Workbook()
    wb.remove(wb.active)

    ranking = df_saida.groupby("ASRO").size().reset_index(name="TOTAL")
    ranking = ranking.sort_values(by="TOTAL", ascending=False)

    ws_rank = wb.create_sheet("RANKING")
    ws_rank.append(["ASRO", "TOTAL"])
    for _, r in ranking.iterrows():
        ws_rank.append([r["ASRO"], r["TOTAL"]])

    grupos = df_saida.groupby("ASRO") if "ASRO" in df_saida.columns else [("GERAL", df_saida)]
    grupos = sorted(grupos, key=lambda x: str(x[0]))

    temp_logo_path = None
    if logo_file is not None:
        suffix = os.path.splitext(logo_file.name)[1] or ".png"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(logo_file.read())
            temp_logo_path = tmp.name

    def limpar_nome_aba(nome):
        for c in ['\\', '/', '*', '[', ']', ':', '?']:
            nome = nome.replace(c, "")
        return nome[:30]

    for nome_aba, dados in grupos:
        nome_final = limpar_nome_aba(f"{nome_aba} - {periodo}")
        ws = wb.create_sheet(title=nome_final)

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
            cell = ws.cell(row=4, column=col)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="0DB39E", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")

        borda = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for i, row in enumerate(dados.itertuples(index=False), start=5):
            valores = [row[0], row[2], row[1], row[3], row[4]]
            for col, val in enumerate(valores, start=1):
                cell = ws.cell(row=i, column=col)
                cell.value = val
                cell.border = borda

        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 50
        ws.column_dimensions["D"].width = 60
        ws.column_dimensions["E"].width = 25

    nome_arquivo = f"resultado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    data = excel_bytes_from_wb(wb)

    if temp_logo_path and os.path.exists(temp_logo_path):
        try:
            os.remove(temp_logo_path)
        except Exception:
            pass

    return data, nome_arquivo


# =========================
# 3) RELATÓRIO DE AGENTES
# =========================
def processar_relatorio_agentes(uploaded_file):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
    from openpyxl.chart import BarChart, Reference

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()

    def col(p):
        return next((c for c in df.columns if p in c), None)

    col_agente = col("NOME DO AGENTE")
    col_asro = col("ASRO")
    col_tipo = col("TIPO DE ATENDIMENTO")
    col_data = col("DATA")

    faltando = [nome for nome, valor in {
        "NOME DO AGENTE": col_agente,
        "ASRO": col_asro,
        "TIPO DE ATENDIMENTO": col_tipo,
        "DATA": col_data,
    }.items() if valor is None]

    if faltando:
        raise ValueError(f"Colunas obrigatórias não encontradas: {', '.join(faltando)}")

    df = df[[col_agente, col_asro, col_tipo, col_data]].copy()
    df[col_data] = pd.to_datetime(df[col_data], errors="coerce").dt.date

    def classificar(valor):
        v = str(valor).upper()
        if "ADESÃO" in v or "ADESAO" in v:
            return "ADESAO"
        if "DEMOLIDO" in v:
            return "DEMOLIDO"
        if "AUSENTE" in v:
            return "AUSENTE"
        if "RECUSA" in v:
            return "RECUSA"
        return None

    df["TIPO_CLASS"] = df[col_tipo].apply(classificar)

    wb = Workbook()
    wb.remove(wb.active)

    cinza = PatternFill(start_color="D9D9D9", fill_type="solid")
    cinza_total = PatternFill(start_color="BFBFBF", fill_type="solid")
    verde = PatternFill(start_color="006400", fill_type="solid")

    borda = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for asro, dados_asro in df.groupby(col_asro):
        ws = wb.create_sheet(title=safe_sheet_name(str(asro)))

        ws.merge_cells("A1:G1")
        ws["A1"] = "RELATÓRIO AGENTES"
        ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
        ws["A1"].fill = verde
        ws["A1"].alignment = Alignment(horizontal="center")

        ws.merge_cells("A2:G2")

        linha = 3
        headers = ["DATAS", "AGENTES", "ADESÃO", "DEMOLIDO", "AUSENTE", "RECUSA", "TOTAL GERAL"]
        for col_id, h in enumerate(headers, start=1):
            cell = ws.cell(row=linha, column=col_id)
            cell.value = h
            cell.border = borda

        linha += 1

        for data_item, dados_data in dados_asro.groupby(col_data):
            inicio_bloco = linha
            for agente, dados_agente in dados_data.groupby(col_agente):
                cont = dados_agente["TIPO_CLASS"].value_counts()

                adesao = cont.get("ADESAO", 0)
                demolido = cont.get("DEMOLIDO", 0)
                ausente = cont.get("AUSENTE", 0)
                recusa = cont.get("RECUSA", 0)

                ws.cell(row=linha, column=2).value = agente
                ws.cell(row=linha, column=3).value = adesao
                ws.cell(row=linha, column=4).value = demolido
                ws.cell(row=linha, column=5).value = ausente
                ws.cell(row=linha, column=6).value = recusa
                ws.cell(row=linha, column=7).value = f"=C{linha}+D{linha}+E{linha}+F{linha}"

                for col_id in range(1, 8):
                    cell = ws.cell(row=linha, column=col_id)
                    cell.fill = cinza
                    cell.border = borda

                linha += 1

            if linha - 1 >= inicio_bloco:
                ws.merge_cells(start_row=inicio_bloco, end_row=linha - 1, start_column=1, end_column=1)
                ws.cell(row=inicio_bloco, column=1).value = str(data_item)

        ws.cell(row=linha, column=1).value = "TOTAL"
        ws.cell(row=linha, column=2).value = f"=COUNTA(B4:B{linha-1})"

        for col_id in range(3, 8):
            letra = chr(64 + col_id)
            ws.cell(row=linha, column=col_id).value = f"=SUM({letra}4:{letra}{linha-1})"

        for col_id in range(1, 8):
            cell = ws.cell(row=linha, column=col_id)
            cell.fill = cinza_total
            cell.border = borda

        grafico_inicio = linha + 2
        ws.cell(row=grafico_inicio, column=2, value="TIPO")
        ws.cell(row=grafico_inicio, column=3, value="TOTAL")

        tipos = ["ADESAO", "DEMOLIDO", "AUSENTE", "RECUSA"]
        linha_chart = grafico_inicio + 1
        for tipo in tipos:
            valor = dados_asro[dados_asro["TIPO_CLASS"] == tipo].shape[0]
            ws.cell(row=linha_chart, column=2).value = tipo
            ws.cell(row=linha_chart, column=3).value = valor
            linha_chart += 1

        chart = BarChart()
        chart.title = "Resumo por Atendimento"

        data_ref = Reference(ws, min_col=3, min_row=grafico_inicio, max_row=linha_chart - 1)
        cats_ref = Reference(ws, min_col=2, min_row=grafico_inicio + 1, max_row=linha_chart - 1)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws.add_chart(chart, f"I{grafico_inicio}")

        ws.column_dimensions["A"].width = 15
        ws.column_dimensions["B"].width = 35
        ws.column_dimensions["C"].width = 15
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["E"].width = 20
        ws.column_dimensions["F"].width = 20
        ws.column_dimensions["G"].width = 18

    ws_dash = wb.create_sheet("DASHBOARD")
    ws_dash["A1"] = "DASHBOARD GERAL"
    ws_dash["A1"].font = Font(size=16, bold=True)

    total_adesao = len(df[df["TIPO_CLASS"] == "ADESAO"])
    total_demolido = len(df[df["TIPO_CLASS"] == "DEMOLIDO"])
    total_ausente = len(df[df["TIPO_CLASS"] == "AUSENTE"])
    total_recusa = len(df[df["TIPO_CLASS"] == "RECUSA"])

    ws_dash["A3"] = "TIPO"
    ws_dash["B3"] = "TOTAL"
    ws_dash.append(["ADESÃO", total_adesao])
    ws_dash.append(["DEMOLIDO", total_demolido])
    ws_dash.append(["AUSENTE", total_ausente])
    ws_dash.append(["RECUSA", total_recusa])

    chart = BarChart()
    chart.title = "Resumo Geral de Atendimentos"

    data_ref = Reference(ws_dash, min_col=2, min_row=3, max_row=7)
    cats_ref = Reference(ws_dash, min_col=1, min_row=4, max_row=7)

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.style = 10
    ws_dash.add_chart(chart, "D5")

    nome_arquivo = f"relatorio_agentes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return excel_bytes_from_wb(wb), nome_arquivo


# =========================
# 4) RELATÓRIO ENVIO / VISITAS / ADESÕES
# =========================
def processar_relatorio_envio_visitas_adesoes(uploaded_file):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.chart import BarChart, Reference

    def encontrar_coluna(df_local, palavras):
        colunas_normalizadas = {}
        for col in df_local.columns:
            col_norm = normalize_text(col)
            colunas_normalizadas[col] = col_norm
        for col_original, col_norm in colunas_normalizadas.items():
            if all(normalize_text(p) in col_norm for p in palavras):
                return col_original
        return None

    def classificar_tipo_atendimento(valor):
        v = normalize_text(valor)
        if "ADESAO REALIZADA" in v and "MORADOR PRESENTE" in v:
            return "ADESÕES"
        if "AGENDAMENTO" in v:
            return "AGENDAMENTOS"
        if "MORADOR AUSENTE" in v:
            return "MORADORES AUSENTES"
        if "RECUSA" in v:
            return "RECUSAS"
        return None

    def calcular_metricas(df_base, col_data, col_agente, col_tipo):
        visitas_totais = df_base[col_data].notna().sum()
        dias_trabalhados = df_base[col_data].dropna().nunique()
        media_visitas_diarias = visitas_totais / dias_trabalhados if dias_trabalhados > 0 else 0

        df_validos = df_base[df_base["TIPO_CLASS"].notna()].copy()
        imoveis_visitados = len(df_validos)

        adesoes = (df_validos["TIPO_CLASS"] == "ADESÕES").sum()
        agendamentos = (df_validos["TIPO_CLASS"] == "AGENDAMENTOS").sum()
        moradores_ausentes = (df_validos["TIPO_CLASS"] == "MORADORES AUSENTES").sum()
        recusas = (df_validos["TIPO_CLASS"] == "RECUSAS").sum()

        def percentual(valor):
            return valor / imoveis_visitados if imoveis_visitados > 0 else 0

        return {
            "Visitas totais": visitas_totais,
            "Imóveis visitados": imoveis_visitados,
            "Média visitas diárias": round(media_visitas_diarias, 2),
            "Dias trabalhados": dias_trabalhados,
            "Moradores ausentes": moradores_ausentes,
            "% Moradores ausentes": percentual(moradores_ausentes),
            "Adesões": adesoes,
            "% Adesões": percentual(adesoes),
            "Recusas": recusas,
            "% Recusas": percentual(recusas),
            "Agendamentos": agendamentos,
            "% Agendamentos": percentual(agendamentos),
        }

    def aplicar_layout_resumo(ws):
        verde_escuro = PatternFill(start_color="006400", fill_type="solid")
        verde_medio = PatternFill(start_color="0DB39E", fill_type="solid")
        branco = Font(color="FFFFFF", bold=True)
        negrito = Font(bold=True)

        borda = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        ws.merge_cells("A1:D1")
        ws["A1"] = "RELATÓRIO DE ADESÕES, VISITAS E ÁREAS CORRELATAS"
        ws["A1"].fill = verde_escuro
        ws["A1"].font = Font(color="FFFFFF", bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        for cell in ws[3]:
            cell.fill = verde_medio
            cell.font = branco
            cell.alignment = Alignment(horizontal="center")
            cell.border = borda

        for row in ws.iter_rows():
            for cell in row:
                if cell.row > 3 and cell.value is not None:
                    cell.border = borda
                    if cell.column in [1, 3]:
                        cell.font = negrito

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 25
        ws.column_dimensions["D"].width = 18

    def aplicar_layout_tabela(ws):
        verde_escuro = PatternFill(start_color="006400", fill_type="solid")
        verde_medio = PatternFill(start_color="0DB39E", fill_type="solid")
        cinza = PatternFill(start_color="D9D9D9", fill_type="solid")
        branco = Font(color="FFFFFF", bold=True)

        borda = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.border = borda
                    if cell.row == 1:
                        cell.fill = verde_escuro
                        cell.font = Font(color="FFFFFF", bold=True, size=14)
                        cell.alignment = Alignment(horizontal="center")
                    elif cell.row == 3:
                        cell.fill = verde_medio
                        cell.font = branco
                        cell.alignment = Alignment(horizontal="center")
                    elif cell.row > 3:
                        cell.fill = cinza

        larguras = {"A": 25, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18}
        for col, width in larguras.items():
            ws.column_dimensions[col].width = width

    def criar_grafico_indicadores(ws, linha_inicio, titulo, posicao):
        chart = BarChart()
        chart.title = titulo
        chart.y_axis.title = "Quantidade"
        chart.x_axis.title = "Indicadores"

        data_ref = Reference(ws, min_col=2, min_row=linha_inicio, max_row=linha_inicio + 4)
        cats_ref = Reference(ws, min_col=1, min_row=linha_inicio + 1, max_row=linha_inicio + 4)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.height = 8
        chart.width = 14
        ws.add_chart(chart, posicao)

    df = read_excel_any(uploaded_file)
    df.columns = df.columns.str.strip()

    col_asro = encontrar_coluna(df, ["ASRO"])
    col_tipo = encontrar_coluna(df, ["TIPO", "ATENDIMENTO"])
    col_agente = encontrar_coluna(df, ["NOME", "AGENTE"])
    col_data = encontrar_coluna(df, ["DATA", "REGISTRO"])

    faltando = []
    if not col_asro:
        faltando.append("ASRO")
    if not col_tipo:
        faltando.append("Tipo de atendimento")
    if not col_agente:
        faltando.append("Nome do agente")
    if not col_data:
        faltando.append("Data do registro")
    if faltando:
        raise ValueError("Colunas obrigatórias não encontradas: " + ", ".join(faltando))

    df[col_data] = pd.to_datetime(df[col_data], errors="coerce").dt.date
    df["TIPO_CLASS"] = df[col_tipo].apply(classificar_tipo_atendimento)

    wb = Workbook()
    wb.remove(wb.active)

    ws_geral = wb.create_sheet("GERAL")
    metricas_geral = calcular_metricas(df, col_data, col_agente, col_tipo)

    ws_geral.append(["RELATÓRIO DE ADESÕES, VISITAS E ÁREAS CORRELATAS"])
    ws_geral.append([])
    ws_geral.append(["Indicador", "Valor", "Indicador Percentual", "Percentual"])

    ws_geral.append(["Visitas totais", metricas_geral["Visitas totais"], "", ""])
    ws_geral.append(["Imóveis visitados", metricas_geral["Imóveis visitados"], "", ""])
    ws_geral.append(["Média visitas diárias", metricas_geral["Média visitas diárias"], "", ""])
    ws_geral.append(["Dias trabalhados", metricas_geral["Dias trabalhados"], "", ""])
    ws_geral.append(["Moradores ausentes", metricas_geral["Moradores ausentes"], "% Moradores ausentes", metricas_geral["% Moradores ausentes"]])
    ws_geral.append(["Adesões", metricas_geral["Adesões"], "% Adesões", metricas_geral["% Adesões"]])
    ws_geral.append(["Recusas", metricas_geral["Recusas"], "% Recusas", metricas_geral["% Recusas"]])
    ws_geral.append(["Agendamentos", metricas_geral["Agendamentos"], "% Agendamentos", metricas_geral["% Agendamentos"]])

    for row in range(8, 12):
        ws_geral.cell(row=row, column=4).number_format = "0.00%"

    aplicar_layout_resumo(ws_geral)

    grafico_linha = 14
    ws_geral.cell(row=grafico_linha, column=1, value="Indicador")
    ws_geral.cell(row=grafico_linha, column=2, value="Total")

    dados_grafico = [
        ("Adesões", metricas_geral["Adesões"]),
        ("Moradores ausentes", metricas_geral["Moradores ausentes"]),
        ("Recusas", metricas_geral["Recusas"]),
        ("Agendamentos", metricas_geral["Agendamentos"]),
    ]

    linha = grafico_linha + 1
    for indicador, valor in dados_grafico:
        ws_geral.cell(row=linha, column=1, value=indicador)
        ws_geral.cell(row=linha, column=2, value=valor)
        linha += 1

    criar_grafico_indicadores(ws_geral, grafico_linha, "Resumo Geral", "F3")

    linha_asro = 22
    ws_geral.cell(row=linha_asro, column=1, value="ASRO")
    ws_geral.cell(row=linha_asro, column=2, value="Visitas totais")
    ws_geral.cell(row=linha_asro, column=3, value="Imóveis visitados")
    ws_geral.cell(row=linha_asro, column=4, value="Adesões")
    ws_geral.cell(row=linha_asro, column=5, value="% Adesões")
    ws_geral.cell(row=linha_asro, column=6, value="Ausentes")
    ws_geral.cell(row=linha_asro, column=7, value="% Ausentes")
    ws_geral.cell(row=linha_asro, column=8, value="Recusas")
    ws_geral.cell(row=linha_asro, column=9, value="% Recusas")

    linha_asro += 1
    grupos = sorted(df.groupby(col_asro), key=lambda x: str(x[0]))

    for asro, dados_asro in grupos:
        m = calcular_metricas(dados_asro, col_data, col_agente, col_tipo)
        ws_geral.cell(row=linha_asro, column=1, value=str(asro))
        ws_geral.cell(row=linha_asro, column=2, value=m["Visitas totais"])
        ws_geral.cell(row=linha_asro, column=3, value=m["Imóveis visitados"])
        ws_geral.cell(row=linha_asro, column=4, value=m["Adesões"])
        ws_geral.cell(row=linha_asro, column=5, value=m["% Adesões"])
        ws_geral.cell(row=linha_asro, column=6, value=m["Moradores ausentes"])
        ws_geral.cell(row=linha_asro, column=7, value=m["% Moradores ausentes"])
        ws_geral.cell(row=linha_asro, column=8, value=m["Recusas"])
        ws_geral.cell(row=linha_asro, column=9, value=m["% Recusas"])

        ws_geral.cell(row=linha_asro, column=5).number_format = "0.00%"
        ws_geral.cell(row=linha_asro, column=7).number_format = "0.00%"
        ws_geral.cell(row=linha_asro, column=9).number_format = "0.00%"
        linha_asro += 1

    for asro, dados_asro in grupos:
        ws = wb.create_sheet(safe_sheet_name(str(asro)))
        metricas = calcular_metricas(dados_asro, col_data, col_agente, col_tipo)

        ws.append([f"RELATÓRIO - {asro}"])
        ws.append([])
        ws.append(["Indicador", "Valor", "Indicador Percentual", "Percentual"])

        ws.append(["Visitas totais", metricas["Visitas totais"], "", ""])
        ws.append(["Imóveis visitados", metricas["Imóveis visitados"], "", ""])
        ws.append(["Média visitas diárias", metricas["Média visitas diárias"], "", ""])
        ws.append(["Dias trabalhados", metricas["Dias trabalhados"], "", ""])
        ws.append(["Moradores ausentes", metricas["Moradores ausentes"], "% Moradores ausentes", metricas["% Moradores ausentes"]])
        ws.append(["Adesões", metricas["Adesões"], "% Adesões", metricas["% Adesões"]])
        ws.append(["Recusas", metricas["Recusas"], "% Recusas", metricas["% Recusas"]])
        ws.append(["Agendamentos", metricas["Agendamentos"], "% Agendamentos", metricas["% Agendamentos"]])

        for row in range(8, 12):
            ws.cell(row=row, column=4).number_format = "0.00%"

        aplicar_layout_resumo(ws)

        grafico_linha = 14
        ws.cell(row=grafico_linha, column=1, value="Indicador")
        ws.cell(row=grafico_linha, column=2, value="Total")

        dados_grafico_asro = [
            ("Adesões", metricas["Adesões"]),
            ("Moradores ausentes", metricas["Moradores ausentes"]),
            ("Recusas", metricas["Recusas"]),
            ("Agendamentos", metricas["Agendamentos"]),
        ]

        linha = grafico_linha + 1
        for indicador, valor in dados_grafico_asro:
            ws.cell(row=linha, column=1, value=indicador)
            ws.cell(row=linha, column=2, value=valor)
            linha += 1

        criar_grafico_indicadores(ws, grafico_linha, f"Resumo {asro}", "F3")

        linha_agente = 22
        ws.cell(row=linha_agente, column=1, value="Agente")
        ws.cell(row=linha_agente, column=2, value="Visitas totais")
        ws.cell(row=linha_agente, column=3, value="Imóveis visitados")
        ws.cell(row=linha_agente, column=4, value="Adesões")
        ws.cell(row=linha_agente, column=5, value="Ausentes")
        ws.cell(row=linha_agente, column=6, value="Recusas")
        ws.cell(row=linha_agente, column=7, value="Agendamentos")

        linha_agente += 1
        for agente, dados_agente in sorted(dados_asro.groupby(col_agente), key=lambda x: str(x[0])):
            m_agente = calcular_metricas(dados_agente, col_data, col_agente, col_tipo)
            ws.cell(row=linha_agente, column=1, value=str(agente))
            ws.cell(row=linha_agente, column=2, value=m_agente["Visitas totais"])
            ws.cell(row=linha_agente, column=3, value=m_agente["Imóveis visitados"])
            ws.cell(row=linha_agente, column=4, value=m_agente["Adesões"])
            ws.cell(row=linha_agente, column=5, value=m_agente["Moradores ausentes"])
            ws.cell(row=linha_agente, column=6, value=m_agente["Recusas"])
            ws.cell(row=linha_agente, column=7, value=m_agente["Agendamentos"])
            linha_agente += 1

        aplicar_layout_tabela(ws)

    nome_arquivo = f"Relatorio_Adesoes_Visitas_Areas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return excel_bytes_from_wb(wb), nome_arquivo


# =========================
# 5) SANDBOX / RELATÓRIO DIÁRIO
# =========================
def processar_sandbox(uploaded_file):
    df = read_excel_any(uploaded_file, dtype=str)

    col_data = None
    col_tipo = None
    for col in df.columns:
        nome = str(col).lower()
        if "data" in nome:
            col_data = col
        if "atendimento" in nome:
            col_tipo = col

    if col_data:
        df[col_data] = pd.to_datetime(df[col_data], errors="coerce")

    if col_tipo:
        df["Tipo_Tratado"] = df[col_tipo].astype(str).str.upper()
    else:
        df["Tipo_Tratado"] = ""

    def classificar(x):
        if "ADES" in x:
            return "Adesão"
        elif "RECUSA" in x:
            return "Recusa"
        elif "AUSENTE" in x:
            return "Ausente"
        elif "ABANDONADO" in x:
            return "Abandonado"
        elif "AGEND" in x:
            return "Agendamento"
        return "Outros"

    df["Status"] = df["Tipo_Tratado"].apply(classificar)
    return df, col_data


# =========================
# 6) PROGRAMA EXPORTAÇÃO
# =========================
def processar_programa_exportacao(uploaded_file):
    df = read_excel_any(uploaded_file)

    cols_norm = {str(c).strip().lower(): c for c in df.columns}
    required = ["link", "codigo", "data", "cliente"]
    faltantes = [c for c in required if c not in cols_norm]
    if faltantes:
        raise ValueError(f"Colunas obrigatórias não encontradas: {', '.join(faltantes)}")

    df_final = pd.DataFrame()
    df_final["Link Backoffice"] = df[cols_norm["link"]]
    df_final["Code Deep"] = df[cols_norm["codigo"]]
    df_final["Data do registro"] = df[cols_norm["data"]]
    df_final["ASRO"] = "PREENCHER"
    df_final["É novo cliente?"] = df[cols_norm["cliente"]]
    df_final["Backoffice"] = "OK"
    df_final["Motivos"] = ""
    df_final["Analise"] = ""

    nome_arquivo = "relatorio_final.xlsx"
    return excel_bytes_from_df(df_final), nome_arquivo


# =========================
# ABAS
# =========================
tab_filtro, tab_termo, tab_agentes, tab_envio, tab_sandbox, tab_exportacao = st.tabs([
    "Filtro de Adesões",
    "Termo de Doação",
    "Relatório de Agentes",
    "Relatório Envio / Visitas / Adesões",
    "Sandbox / Relatório Diário",
    "Programa Exportação",
])

with tab_filtro:
    st.header("Filtro de Adesões")
    st.write(
        """
        **Para que serve:**  
        Filtra exatamente os registros de **ADESÃO REALIZADA (MORADOR PRESENTE)**  
        e gera um Excel final formatado, com:
        - cabeçalho roxo
        - bordas
        - largura ajustada automaticamente
        - mesmas colunas do seu modelo original
        """
    )

    arquivo = st.file_uploader("Selecione o Excel bruto", type=["xlsx", "xls"], key="filtro_adesoes")
    if arquivo and st.button("Executar Filtro de Adesões", key="btn_filtro"):
        try:
            data, nome = processar_filtro_adesoes(arquivo)
            st.success("Relatório gerado com sucesso!")
            st.download_button("Baixar Excel", data=data, file_name=nome, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erro: {e}")

with tab_termo:
    st.header("Termo de Doação")
    st.write(
        """
        **Para que serve:**  
        Gera o Termo de Doação de Padrão mantendo a lógica original:
        - filtra somente clientes marcados como SIM
        - cria a aba RANKING
        - gera abas por ASRO
        - escreve período
        - aplica cabeçalho e largura formatada
        - pode incluir logo opcional
        """
    )

    arquivo = st.file_uploader("Selecione o Excel bruto", type=["xlsx"], key="termo_doacao")
    logo = st.file_uploader("Logo opcional (PNG/JPG)", type=["png", "jpg", "jpeg"], key="termo_logo")
    if arquivo and st.button("Gerar Termo de Doação", key="btn_termo"):
        try:
            data, nome = processar_termo_doacao(arquivo, logo_file=logo)
            st.success("Termo gerado com sucesso!")
            st.download_button("Baixar Excel", data=data, file_name=nome, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erro: {e}")

with tab_agentes:
    st.header("Relatório de Agentes")
    st.write(
        """
        **Para que serve:**  
        Gera relatório completo por ASRO e por agente, com:
        - totais por tipo de atendimento
        - fórmulas no Excel
        - abas por ASRO
        - gráfico por aba
        - dashboard geral no workbook
        """
    )

    arquivo = st.file_uploader("Selecione o Excel bruto", type=["xlsx"], key="rel_agentes")
    if arquivo and st.button("Gerar Relatório de Agentes", key="btn_agentes"):
        try:
            data, nome = processar_relatorio_agentes(arquivo)
            st.success("Relatório gerado com sucesso!")
            st.download_button("Baixar Excel", data=data, file_name=nome, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erro: {e}")

with tab_envio:
    st.header("Relatório Envio / Visitas / Adesões")
    st.write(
        """
        **Para que serve:**  
        Gera um dashboard Excel completo com:
        - visitas totais
        - imóveis visitados
        - adesões
        - recusas
        - moradores ausentes
        - agendamentos
        - gráficos
        - resumo geral
        - resumo por ASRO
        - abas por ASRO
        - tabela por agente
        """
    )

    arquivo = st.file_uploader("Selecione o Excel bruto", type=["xlsx", "xls"], key="rel_envio")
    if arquivo and st.button("Gerar Relatório Envio / Visitas / Adesões", key="btn_envio"):
        try:
            data, nome = processar_relatorio_envio_visitas_adesoes(arquivo)
            st.success("Relatório gerado com sucesso!")
            st.download_button("Baixar Excel", data=data, file_name=nome, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erro: {e}")

with tab_sandbox:
    st.header("Sandbox / Relatório Diário")
    st.write(
        """
        **Para que serve:**  
        Exibe análise diária dos atendimentos:
        - detecta automaticamente coluna de data e atendimento
        - classifica os status
        - mostra KPIs
        - mostra gráfico diário
        - mostra tabela consolidada
        - permite baixar o Excel final
        """
    )

    arquivo = st.file_uploader("Selecione o Excel", type=["xlsx"], key="sandbox")
    if arquivo:
        try:
            import plotly.express as px

            df, col_data = processar_sandbox(arquivo)
            total = len(df)
            adesao = (df["Status"] == "Adesão").sum()
            recusa = (df["Status"] == "Recusa").sum()

            c1, c2, c3 = st.columns(3)
            c1.metric("Visitas", total)
            c2.metric("Adesões", adesao)
            c3.metric("Recusas", recusa)

            st.subheader("Volume diário de visitas")
            if col_data:
                volume = df.groupby(df[col_data].dt.date).size().reset_index(name="Visitas")
                volume.columns = ["Data", "Visitas"]
                volume = volume.sort_values("Data")

                fig = px.line(volume, x="Data", y="Visitas", markers=True, text="Visitas")
                fig.update_traces(textposition="top center")
                fig.update_layout(
                    plot_bgcolor="#f3f7f5",
                    paper_bgcolor="#f3f7f5",
                    font=dict(color="#0b6b57"),
                    xaxis_title="Data",
                    yaxis_title="Quantidade"
                )
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("Relatório por Data")
            if col_data:
                df["Adesão"] = df["Tipo_Tratado"].str.contains("ADES", na=False).astype(int)
                df["Recusa"] = df["Tipo_Tratado"].str.contains("RECUSA", na=False).astype(int)
                tabela = df.groupby(df[col_data].dt.date).agg({"Adesão": "sum", "Recusa": "sum"}).reset_index()
                tabela.columns = ["Data", "Adesão", "Recusa"]
                tabela = tabela.sort_values("Data")

                total_row = {"Data": "Total", "Adesão": tabela["Adesão"].sum(), "Recusa": tabela["Recusa"].sum()}
                tabela = pd.concat([tabela, pd.DataFrame([total_row])], ignore_index=True)
                st.dataframe(tabela, use_container_width=True)

            st.download_button(
                "Baixar Excel",
                data=excel_bytes_from_df(df),
                file_name="Relatorio_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erro: {e}")

with tab_exportacao:
    st.header("Programa Exportação")
    st.write(
        """
        **Para que serve:**  
        Lê a base operacional e monta o relatório final padronizado com:
        - Link Backoffice
        - Code Deep
        - Data do registro
        - ASRO
        - É novo cliente?
        - Backoffice
        - Motivos
        - Analise
        """
    )

    arquivo = st.file_uploader("Selecione o Excel de entrada", type=["xlsx"], key="exportacao")
    if arquivo and st.button("Executar Exportação", key="btn_exportacao"):
        try:
            data, nome = processar_programa_exportacao(arquivo)
            st.success("Relatório gerado com sucesso!")
            st.download_button("Baixar Excel", data=data, file_name=nome, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erro: {e}")

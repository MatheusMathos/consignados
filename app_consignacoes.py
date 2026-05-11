import pandas as pd
import streamlit as st
import os
from io import BytesIO
import unicodedata

# ========================
# CONFIG
# ========================
st.set_page_config(page_title="Dashboard Consignações", layout="wide")

FILE_PATH = "Consignacoes_Acumulado.xlsx"

# ========================
# ESTILO
# ========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,100..900;1,100..900&display=swap" rel="stylesheet');

/* Base */
html, body, [class*="css"] {
    font-family: 'Roboto', sans-serif;
    background-color: #EBE5DE;
    color: #1C1C1C;
}
.stApp { background-color: #EBE5DE; }

/* Padding geral */
.block-container {
    padding: 2.5rem 3.5rem 3rem 3.5rem !important;
    max-width: 1380px !important;
}

/* Título h1 */
h1 {
    font-family: 'Roboto', sans-serif !important;
    font-size: 2.6rem !important;
    font-weight: 700 !important;
    color: #1C1C1C !important;
    letter-spacing: -0.5px !important;
    line-height: 1.1 !important;
    margin-bottom: 0.1rem !important;
}

/* Subtítulos h2/h3 */
h2, h3 {
    font-family: 'Roboto', sans-serif !important;
    font-size: 1.2rem !important;
    font-weight: 600 !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    color: #737270 !important;
    margin-top: 2.2rem !important;
    margin-bottom: 0.8rem !important;
    padding-bottom: 0.6rem !important;
    border-bottom: 1px solid #E8E4DE !important;
}

.kpi-card {
    background: #FFFFFF;
    border: 1px solid #EAE6E0;
    border-top: 3px solid #C9B99A;
    border-radius: 12px;
    padding: 1.2rem 1rem 1rem 1rem;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    transition: box-shadow 0.2s ease;
    flex: 1 1 0;
    min-width: 0;
    overflow: hidden;
    word-break: break-word;
}
.kpi-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.08); }
.kpi-label {
    font-family: 'Roboto', sans-serif;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    color: #737270;
    margin-bottom: 0.5rem;
    line-height: 1.4;
}
.kpi-value {
    font-family: 'Roboto', sans-serif;
    font-size: clamp(1rem, 2vw, 1.6rem);
    font-weight: 600;
    color: #1C1C1C;
    line-height: 1.3;
    overflow: hidden;
    text-overflow: ellipsis;
}
.kpi-row {
    display: flex;
    gap: 0.75rem;
    margin-bottom: 1.8rem;
}

/* Dataframes */
[data-testid="stDataFrame"] {
    border: 1px solid #EAE6E0 !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    box-shadow: 0 1px 6px rgba(0,0,0,0.03) !important;
}

/* Tabs */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 2px solid #EAE6E0 !important;
    gap: 0 !important;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    font-family: 'Roboto', sans-serif !important;
    font-size: 0.72rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.8px !important;
    text-transform: uppercase !important;
    color: #737270 !important;
    padding: 0.55rem 1.4rem !important;
    border: none !important;
    background: transparent !important;
}
[data-testid="stTabs"] [aria-selected="true"] {
    color: #1C1C1C !important;
    border-bottom: 2px solid #C9B99A !important;
}

/* Botões download */
.stDownloadButton > button {
    background: #FFFFFF !important;
    color: #1C1C1C !important;
    border: 1px solid #C9B99A !important;
    border-radius: 6px !important;
    font-family: 'Roboto', sans-serif !important;
    font-size: 0.68rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.3px !important;
    padding: 0.28rem 0.7rem !important;
    transition: all 0.18s ease !important;
}
.stDownloadButton > button:hover {
    background: #C9B99A !important;
    color: #FFFFFF !important;
    border-color: #C9B99A !important;
}

/* Multiselect */
.stMultiSelect [data-baseweb="select"] > div {
    background: #FFFFFF !important;
    border: 1.5px solid #DDD8D0 !important;
    border-radius: 8px !important;
}
.stMultiSelect [data-baseweb="tag"] {
    background: #F0EBE3 !important;
    color: #5C4F3D !important;
    border-radius: 5px !important;
    border: 1px solid #DDD8D0 !important;
}
.stMultiSelect label, .stSelectbox label {
    font-size: 0.62rem !important;
    letter-spacing: 1.4px !important;
    text-transform: uppercase !important;
    color: #737270 !important;
    font-weight: 600 !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #FFFFFF;
    border: 1.5px dashed #DDD8D0;
    border-radius: 10px;
    padding: 1rem;
}

/* Expander */
[data-testid="stExpander"] {
    border: 1px solid #EAE6E0 !important;
    border-radius: 10px !important;
    background: #FFFFFF !important;
}

/* Divisores */
hr { border: none; border-top: 1px solid #EAE6E0; margin: 1.8rem 0; }

/* Texto forte */
strong { color: #5C4F3D; }

/* Remove marca Streamlit */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
header    { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ========================
# FUNÇÕES
# ========================
def converter_valor(valor):
    if pd.isna(valor):
        return 0.0
    
    valor = str(valor).strip()

    if ',' in valor:
        valor = valor.replace('.', '').replace(',', '.')
    
    try:
        return float(valor)
    except:
        return 0.0


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    
    texto = str(texto).strip().upper()

    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

    return texto


def load_data():
    if os.path.exists(FILE_PATH):
        df = pd.read_excel(FILE_PATH)
        return tratar_df(df)

    uploaded_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        return tratar_df(df)

    st.error("Arquivo não encontrado. Coloque o Excel no projeto ou faça upload.")
    st.stop()


def tratar_df(df):
    df['Data Emissão'] = pd.to_datetime(df['Data Emissão'], errors='coerce')
    df['Data do Pagamento/Previsão'] = pd.to_datetime(df['Data do Pagamento/Previsão'], errors='coerce')

    df['Total da Nota'] = df['Total da Nota'].apply(converter_valor)

    df['Anotações'] = df['Anotações'].apply(normalizar_texto)
    df['Pareado'] = df['Pareado'].apply(normalizar_texto)
    df['Espécie'] = df['Espécie'].apply(normalizar_texto)

    df['Permanencia'] = (df['Data do Pagamento/Previsão'] - df['Data Emissão']).dt.days
    df['Permanencia'] = df['Permanencia'].abs()

    return df


def format_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_int(valor):
    return f"{valor:,}".replace(",", ".")


def to_excel(df):
    output = BytesIO()
    df.to_excel(output, index=False)
    return output.getvalue()


def to_excel_duas_abas(df_resumo, df_detalhe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='Diferença por Loja', index=False)
        df_detalhe.to_excel(writer, sheet_name='Detalhamento NFs', index=False)
    return output.getvalue()


# ========================
# APP
# ========================
try:
    df = load_data()

    st.markdown("""
    <div style="margin-bottom: 0.2rem;">
        <span style="font-family:'Roboto',sans-serif; font-size:2.6rem; font-weight:700; color:#1C1C1C; letter-spacing:-0.5px; line-height:1.1;">
            👠 Consignações
        </span>
    </div>
    <p style="font-family:'Roboto',sans-serif; font-size:0.72rem; font-weight:500; letter-spacing:2px; text-transform:uppercase; color:#737270; margin-top:0.3rem; margin-bottom:2rem;">
        Acompanhamento &nbsp;·&nbsp; Análise &nbsp;·&nbsp; Controle
    </p>
    <hr style="border:none; border-top:1px solid #EAE6E0; margin-bottom:1.8rem;">
    """, unsafe_allow_html=True)

    # ========================
    # FILTROS
    # ========================
    df['Ano'] = df['Data Emissão'].dt.year
    df['Mes'] = df['Data Emissão'].dt.month

    col1, col2 = st.columns(2)

    with col1:
        anos_disponiveis = sorted(df['Ano'].dropna().unique())
        anos = st.multiselect("Ano", anos_disponiveis, default=anos_disponiveis)

    df_ano_filtrado = df[df['Ano'].isin(anos)]
    meses_disponiveis = sorted(df_ano_filtrado['Mes'].dropna().unique())

    with col2:
        meses = st.multiselect("Mês", meses_disponiveis, default=meses_disponiveis)

    df_filtrado = df[
        (df['Ano'].isin(anos)) &
        (df['Mes'].isin(meses))
    ]

    # ========================
    # KPIs
    # ========================
    total_notas = len(df_filtrado)
    total_geral = df_filtrado['Total da Nota'].sum()

    df_ok = df_filtrado[df_filtrado['Anotações'] == 'PROCESSO OK']
    valor_ok = df_ok['Total da Nota'].sum()

    df_divergencia = df_filtrado[
        (df_filtrado['Pareado'] == 'NAO PAREADO') &
        (df_filtrado['Anotações'] != 'PROCESSO OK')
    ]
    valor_divergencia = df_divergencia['Total da Nota'].sum()

    entradas = df_filtrado[df_filtrado['Espécie'] == 'ENTRADA']['Total da Nota'].sum()
    saidas = df_filtrado[df_filtrado['Espécie'] == 'SAIDA']['Total da Nota'].sum()
    diferenca_geral = entradas - saidas

    kpi_html = (
        '<div class="kpi-row">'
        f'<div class="kpi-card"><div class="kpi-label">Total de Registros</div><div class="kpi-value">{format_int(total_notas)}</div></div>'
        f'<div class="kpi-card"><div class="kpi-label">Total Geral</div><div class="kpi-value">{format_brl(total_geral)}</div></div>'
        f'<div class="kpi-card"><div class="kpi-label">Processo OK</div><div class="kpi-value">{format_brl(valor_ok)}</div></div>'
        f'<div class="kpi-card"><div class="kpi-label">Divergências</div><div class="kpi-value">{format_brl(valor_divergencia)}</div></div>'
        f'<div class="kpi-card"><div class="kpi-label">Diferença Geral</div><div class="kpi-value">{format_brl(diferenca_geral)}</div></div>'
        '</div>'
    )
    st.markdown(kpi_html, unsafe_allow_html=True)

    # ========================
    # MATRIZ
    # ========================
    st.subheader("📊 Diferença Entradas vs Saídas")

    tab_geral, tab_loja = st.tabs(["Geral", "Por Loja"])

    with tab_geral:
        tabela = df_filtrado.groupby(['Ano', 'Mes', 'Espécie'])['Total da Nota'].sum().reset_index()

        pivot = tabela.pivot_table(
            index=['Ano', 'Mes'],
            columns='Espécie',
            values='Total da Nota',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot['Diferença'] = pivot.get('ENTRADA', 0) - pivot.get('SAIDA', 0)

        total_row = pd.DataFrame({
            'Ano': ['TOTAL'],
            'Mes': [''],
            'ENTRADA': [pivot['ENTRADA'].sum() if 'ENTRADA' in pivot else 0],
            'SAIDA': [pivot['SAIDA'].sum() if 'SAIDA' in pivot else 0],
            'Diferença': [pivot['Diferença'].sum()]
        })

        pivot_final = pd.concat([pivot, total_row], ignore_index=True)

        st.dataframe(
            pivot_final.style.format({col: "R$ {:,.2f}" for col in pivot_final.columns if col not in ['Ano', 'Mes']}),
            use_container_width=True
        )

    with tab_loja:
        tabela_loja = df_filtrado.groupby(['Loja', 'Espécie'])['Total da Nota'].sum().reset_index()

        pivot_loja = tabela_loja.pivot_table(
            index=['Loja'],
            columns='Espécie',
            values='Total da Nota',
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        pivot_loja.columns.name = None

        if 'ENTRADA' not in pivot_loja.columns:
            pivot_loja['ENTRADA'] = 0
        if 'SAIDA' not in pivot_loja.columns:
            pivot_loja['SAIDA'] = 0

        pivot_loja['Diferença'] = pivot_loja['ENTRADA'] - pivot_loja['SAIDA']
        pivot_loja = pivot_loja.sort_values('Diferença')

        total_loja_row = pd.DataFrame({
            'Loja': ['TOTAL'],
            'ENTRADA': [pivot_loja['ENTRADA'].sum()],
            'SAIDA': [pivot_loja['SAIDA'].sum()],
            'Diferença': [pivot_loja['Diferença'].sum()]
        })

        pivot_loja_final = pd.concat([pivot_loja, total_loja_row], ignore_index=True)

        st.dataframe(
            pivot_loja_final.style.format({col: "R$ {:,.2f}" for col in pivot_loja_final.columns if col not in ['Loja']}),
            use_container_width=True
        )

        df_nfs_loja = df_filtrado[[
            'Loja', 'NF', 'Espécie', 'Data Emissão', 'Nome da Cliente',
            'Nome da Consultora', 'Total da Nota', 'Data do Pagamento/Previsão',
            'NF de Retorno/Saída', 'Anotações', 'Pareado'
        ]].sort_values(['Loja', 'Data Emissão'])

        st.download_button(
            "📥 Baixar Diferença por Loja",
            data=to_excel_duas_abas(pivot_loja_final, df_nfs_loja),
            file_name="diferenca_por_loja.xlsx"
        )

    # ========================
    # 🔥 RANKING DE ERROS
    # ========================
    st.subheader("⚠️ Ranking de Erros por Loja (R$)")

    df_erros = df_filtrado[
        (df_filtrado['Pareado'] == 'NAO PAREADO') &
        (df_filtrado['Anotações'] != 'PROCESSO OK')
    ].copy()

    df_erros['Numero NF'] = df_erros['NF']
    df_erros['Erro'] = df_erros['Anotações']
    df_erros['Data NF'] = df_erros['Data Emissão']

    ranking_erros = df_erros.groupby('Loja')['Total da Nota'].sum().sort_values(ascending=False).reset_index()

    st.dataframe(
        ranking_erros.style.format({"Total da Nota": "R$ {:,.2f}"}),
        use_container_width=True
    )

    st.download_button(
        "📥 Baixar Ranking de Erros (Detalhado)",
        data=to_excel(df_erros[
            ['Loja', 'Numero NF', 'Erro', 'Data NF',
             'Nome da Cliente', 'Nome da Consultora', 'Total da Nota']
        ]),
        file_name="ranking_erros_detalhado.xlsx"
    )

    # ========================
    # PERMANÊNCIA
    # ========================
    st.subheader("⏳ Tempo de Permanência")

    df_perm = df_filtrado[
        (df_filtrado['Permanencia'] <= 500) &
        (df_filtrado['Permanencia'] >= 1)
    ].copy()

    st.write(f"Média: **{df_perm['Permanencia'].mean():.1f} dias**")

    df_perm_view = df_perm[
        ['NF', 'Loja', 'Nome da Cliente', 'Nome da Consultora',
         'Data Emissão', 'Data do Pagamento/Previsão', 'Permanencia', 'Total da Nota']
    ].dropna()

    df_perm_view = df_perm_view.sort_values(by='Permanencia', ascending=False)

    st.dataframe(df_perm_view.head(20))

    st.download_button(
        "📥 Baixar Permanência (Detalhado)",
        data=to_excel(df_perm_view),
        file_name="permanencia_detalhado.xlsx"
    )

    # ========================
    # RANK CLIENTES
    # ========================
    st.subheader("👩🏻‍🦰 Ranking de Clientes (Valor x Permanência)")

    ranking_cliente = df_perm_view.groupby(['Nome da Cliente', 'Loja']).agg(
        Total_Valor=('Total da Nota', 'sum'),
        Media_Permanencia=('Permanencia', 'mean'),
        Qtd_NF=('NF', 'count')
    ).reset_index().sort_values(by='Total_Valor', ascending=False)

    st.dataframe(
        ranking_cliente.style.format({
            "Total_Valor": "R$ {:,.2f}",
            "Media_Permanencia": "{:.1f}"
        }),
        use_container_width=True
    )

    st.download_button(
        "📥 Baixar Ranking Clientes (Detalhado)",
        data=to_excel(ranking_cliente),
        file_name="ranking_clientes_detalhado.xlsx"
    )

except Exception as e:
    st.error(f"Erro: {e}")
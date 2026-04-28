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


def to_excel(df):
    output = BytesIO()
    df.to_excel(output, index=False)
    return output.getvalue()


# ========================
# APP
# ========================
try:
    df = load_data()

    st.title("👠📝 Acompanhamento Consignações")

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

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de Registros", total_notas)
    c2.metric("Total Geral", format_brl(total_geral))
    c3.metric("Processo OK", format_brl(valor_ok))
    c4.metric("Divergências", format_brl(valor_divergencia))
    c5.metric("Diferença Geral", format_brl(diferenca_geral))

    # ========================
    # MATRIZ
    # ========================
    st.subheader("📊 Diferença Entradas vs Saídas")

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

    # 🔥 NOVO: EXPORTAR DIFERENÇA POR LOJA
    diff_loja = df_filtrado.groupby(['Ano', 'Mes', 'Loja', 'Espécie'])['Total da Nota'].sum().reset_index()

    diff_loja_pivot = diff_loja.pivot_table(
        index=['Ano', 'Mes', 'Loja'],
        columns='Espécie',
        values='Total da Nota',
        fill_value=0
    ).reset_index()

    diff_loja_pivot['Diferença'] = diff_loja_pivot.get('ENTRADA', 0) - diff_loja_pivot.get('SAIDA', 0)

    st.download_button(
        "📥 Baixar Diferença por Loja (Entradas vs Saídas)",
        data=to_excel(diff_loja_pivot),
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
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import seaborn as sns
import io

# --- Configurações da Página e Cores ---
st.set_page_config(layout="wide", page_title="Análise de Pareto")

cores = {
    'azul': '#134883',
    'amarelo': '#F8AC2E',
    'cinza': '#6E7274',
    'branco': '#FFFFFF'
}

# --- Interface Principal ---
st.title("📊 Gerador de Análise de Pareto")
st.write("Faça o upload do seu arquivo Excel para gerar a tabela e o gráfico de Pareto.")
st.info("O arquivo precisa conter uma aba chamada **'Entrada'** com as colunas: `Fornecedor`, `Não Conformes` e `Entregues`.", icon="💡")

# Componente de upload
uploaded_file = st.file_uploader(
    "Escolha o arquivo Excel",
    type=['xlsx', 'xls']
)

if uploaded_file is not None:
    if st.button("Gerar Análise", type="primary"):
        with st.spinner("Analisando os dados..."):
            try:
                # Todo o código a partir daqui deve ter este nível de indentação
                df = pd.read_excel(uploaded_file, sheet_name="Entrada")
                df = df[['Fornecedor', 'Não Conformes', 'Entregues']].dropna()
                df['Não Conformes'] = pd.to_numeric(df['Não Conformes'], errors='coerce')
                df['Entregues'] = pd.to_numeric(df['Entregues'], errors='coerce')
                df = df.dropna()

                if df.empty:
                    st.error("Não foram encontrados dados válidos. Verifique as colunas do seu arquivo.")
                else:
                    # Cálculos com a escala de 0-100
                    df['Taxa NC (%)'] = (df['Não Conformes'] / df['Entregues']) * 100
                    df['% Individual'] = (df['Não Conformes'] / df['Não Conformes'].sum()) * 100
                    df['% Acumulada'] = df['% Individual'].cumsum()

                    st.success("Análise gerada com sucesso!")

                    # Exibir Tabela
                    st.subheader("Tabela de Dados da Análise")
                    st.dataframe(df.style.format({
                        'Taxa NC (%)': '{:.2f}%',
                        'Não Conformes': '{:,.0f}',
                        'Entregues': '{:,.0f}',
                        '% Individual': '{:.2f}%',
                        '% Acumulada': '{:.2f}%'
                    }))

                    # --- Gerar Gráfico ---
                    st.subheader("Gráfico de Pareto")
                    fig, ax = plt.subplots(figsize=(10, 6))
                    sns.set_style("whitegrid")
                    ax.bar(df['Fornecedor'], df['% Individual'], color=cores['azul'])

                    # Laço para adicionar os rótulos de dados
                    for index, row in df.iterrows():
                        valor_nao_conforme = int(row['Não Conformes'])
                        altura_barra = row['% Individual']
                        ax.text(
                            index,
                            altura_barra + 1,
                            f'{valor_nao_conforme:,.0f}',
                            ha='center',
                            va='bottom',
                            fontsize=9,
                            color=cores['cinza'],
                            fontweight='bold'
                        )

                    ax.set_ylabel('% Individual', color=cores['azul'])
                    ax.tick_params(axis='y', labelcolor=cores['azul'])
                    plt.setp(ax.get_xticklabels(), rotation=45, ha="right")

                    ax2 = ax.twinx()
                    ax2.plot(df['Fornecedor'], df['% Acumulada'], color=cores['amarelo'], marker='o', linewidth=2.5)
                    ax2.set_ylabel('% Acumulada', color=cores['amarelo'])
                    ax2.tick_params(axis='y', labelcolor=cores['amarelo'])
                    ax2.set_ylim(0, 105)

                    fig.tight_layout()
                    st.pyplot(fig)

                    # Gerar Download
                    st.subheader("Download do Relatório")
                    img_data = io.BytesIO()
                    fig.savefig(img_data, format='png', bbox_inches='tight')
                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Analise_Pareto')
                        worksheet = writer.sheets['Analise_Pareto']
                        worksheet.set_column('A:A', 20)
                        worksheet.set_column('B:C', 15)
                        worksheet.set_column('D:F', 15)
                        worksheet.insert_image('H2', 'pareto_graph.png', {'image_data': img_data})

                    st.download_button(
                        label="📥 Baixar Relatório Completo (Tabela + Gráfico)",
                        data=output_excel.getvalue(),
                        file_name="relatorio_pareto.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
                st.warning("Por favor, verifique se o arquivo tem uma aba chamada 'Entrada' e as colunas corretas.")

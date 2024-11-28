import streamlit as st
import pandas as pd
import os

# Carregar o arquivo Excel usando caminho relativo (arquivo no mesmo repositório)
caminho_excel = "Controle Pendência AS BUILT (DUTOS) 2022 2 (1).xlsx"  # Substitua pelo nome real da sua planilha

# Carregar o arquivo Excel
df = pd.read_excel(caminho_excel, sheet_name="CAMINHOS_POÇOS", skiprows=2)

# Título do aplicativo
st.title("Abrir Pastas dos Poços")

# Caixa de busca com sugestão dinâmica
termo_busca = st.text_input("Buscar poço ou projeto:")

# Filtrar as opções dinamicamente com base no termo digitado
if termo_busca:
    df_filtrado = df[df['POÇO'].str.contains(termo_busca, case=False, na=False)]
else:
    df_filtrado = df

# Exibir a tabela com todos os poços e projetos
st.write("Poços e Projetos disponíveis:")
st.dataframe(df_filtrado[["PROJETO", "POÇO"]], use_container_width=True)

# Criar um selectbox com os poços filtrados
poco_selecionado = st.selectbox(
    "Escolha um poço para abrir a pasta:",
    df_filtrado["POÇO"].unique()
)

# Abrir pasta ao clicar no botão
if st.button("Abrir Pasta"):
    # Obter o caminho correspondente ao poço selecionado
    caminho_pasta = df.loc[df['POÇO'] == poco_selecionado, 'CAMINHO'].values[0]
    if os.path.exists(caminho_pasta):
        os.startfile(caminho_pasta)  # Abre a pasta no Windows Explorer
        st.success(f"Pasta aberta: {caminho_pasta}")
    else:
        st.error(f"Caminho não encontrado: {caminho_pasta}")

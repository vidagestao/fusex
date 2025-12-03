import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime

# --- Configuração da Página ---
st.set_page_config(page_title="Gestão Corpore", layout="wide")

st.title("🏥 Gestão Corpore - Financeiro")
st.markdown("---")

# --- Conexão com o Google Sheets ---
# O Streamlit busca automaticamente as credenciais que colamos nos "Secrets"
conn = st.connection("gsheets", type=GSheetsConnection)

# Lendo os dados da planilha (Aba padrão 'Página1' ou 'Sheet1')
# ttl=5 garante que ele atualize os dados a cada 5 segundos se houver mudança
data = conn.read(ttl=5)

# --- Barra Lateral para Inserir Novos Dados ---
st.sidebar.header("📝 Novo Lançamento")

with st.sidebar.form(key="form_lancamento"):
    data_lancamento = st.date_input("Data", datetime.today())
    descricao = st.text_input("Descrição (Ex: Material de escritório)")
    valor = st.number_input("Valor (R$)", min_value=0.0, step=0.01)
    tipo = st.selectbox("Tipo", ["Despesa", "Receita"])
    
    submit_button = st.form_submit_button(label="💾 Salvar Lançamento")

# --- Lógica de Salvar ---
if submit_button:
    # 1. Cria uma nova linha com os dados
    novo_dado = pd.DataFrame([
        {
            "Data": data_lancamento.strftime("%Y-%m-%d"),
            "Descricao": descricao,
            "Valor": valor,
            "Tipo": tipo
        }
    ])
    
    # 2. Junta com os dados que já existiam
    # Se a planilha estiver vazia, usa apenas o novo dado
    if data.empty:
        updated_df = novo_dado
    else:
        updated_df = pd.concat([data, novo_dado], ignore_index=True)
    
    # 3. Envia de volta para o Google Sheets
    conn.update(data=updated_df)
    
    st.success("Lançamento salvo com sucesso! Atualize a página para ver na tabela.")
    st.balloons()

# --- Exibição dos Dados (Dashboard) ---
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📊 Extrato de Lançamentos")
    # Mostra a tabela. Use_container_width faz ela ocupar o espaço todo
    st.dataframe(data, use_container_width=True)

with col2:
    st.subheader("📈 Resumo Rápido")
    if not data.empty:
        # Pequeno cálculo para mostrar totais se houver dados
        # Certifique-se que a coluna 'Tipo' e 'Valor' existam na planilha
        try:
            total_receitas = data[data["Tipo"] == "Receita"]["Valor"].sum()
            total_despesas = data[data["Tipo"] == "Despesa"]["Valor"].sum()
            saldo = total_receitas - total_despesas
            
            st.metric("Total Receitas", f"R$ {total_receitas:,.2f}")
            st.metric("Total Despesas", f"R$ {total_despesas:,.2f}")
            st.metric("Saldo Atual", f"R$ {saldo:,.2f}", delta=saldo)
        except Exception as e:
            st.info("Adicione colunas 'Valor' e 'Tipo' na planilha para ver os cálculos.")
    else:
        st.warning("A planilha ainda está vazia.")

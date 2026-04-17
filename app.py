import streamlit as st
import os
import pandas as pd
from datetime import datetime
import main  # Importa o seu main.py

# Configuração da Página
st.set_page_config(page_title="Reposição RDC", layout="wide", page_icon="📊")

st.title("📊 Automação de Análise RDC")
st.markdown("---")

# --- SIDEBAR (ENTRADA DE DADOS) ---
st.sidebar.header("📅 Parâmetros de Data")

with st.sidebar:
    # --- SEÇÃO ENTREGA ---
    st.subheader("Entrega")
    d_ent_ini = st.date_input("Entrega de:", datetime(2026, 4, 30))
    d_ent_fim = st.date_input("Entrega até:", datetime(2026, 5, 11))
    
    # CÁLCULO: Prazo Máximo (T15)
    t15_valor = abs((d_ent_fim - d_ent_ini).days) + 1
    st.info(f"🚚 **Prazo máx. entrega:** {t15_valor} dias")
    
    st.divider()
    
    # --- SEÇÃO VENDA ---
    st.subheader("Venda")
    d_venda_ini = st.date_input("Venda de:", datetime(2026, 2, 13))
    d_venda_fim = st.date_input("Venda até:", datetime(2026, 4, 16))
    
    # CÁLCULO: Período de Venda (T18)
    t18_valor = abs((d_venda_fim - d_venda_ini).days) + 1
    st.warning(f"📈 **Período de venda:** {t18_valor} dias")
    
    st.divider()
    
    # Outros parâmetros
    lojas_alvo = st.text_input("Lojas Alvo:", "161, 318, 328, 473, 533, 567, 582, 610, 611")
    fat_minimo = st.number_input("Fat. Mínimo Padrão (R$):", value=3500.00, step=100.0)

# --- PAINEL PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Arquivos Pendentes")
    if os.path.exists("RDCs_Originais"):
        arquivos = [f for f in os.listdir("RDCs_Originais") if f.endswith(".xlsx")]
        if arquivos:
            for arq in arquivos:
                st.write(f"✅ {arq}")
        else:
            st.warning("Nenhum arquivo RDC encontrado.")
    else:
        st.error("Pasta 'RDCs_Originais' não encontrada!")

with col2:
    st.subheader("⚙️ Execução")
    
    # EXIBIÇÃO DE RESUMO PARA O CHEFE
    st.write(f"**Resumo da Operação:**")
    st.write(f"- Analisando vendas de **{t18_valor} dias**.")
    st.write(f"- Garantindo estoque para **{t15_valor} dias**.")
    
    if st.button("🚀 INICIAR PROCESSAMENTO", use_container_width=True):
        with st.spinner("Conectando ao banco e processando..."):
            try:
                main.rodar_automacao_v2(
                    venda_ini = datetime.combine(d_venda_ini, datetime.min.time()),
                    venda_fim = datetime.combine(d_venda_fim, datetime.max.time()),
                    ent_ini   = datetime.combine(d_ent_ini, datetime.min.time()),
                    ent_fim   = datetime.combine(d_ent_fim, datetime.max.time()),
                    lojas     = lojas_alvo,
                    fat_min   = fat_minimo
                )
                st.success("Análises geradas com sucesso!")
                st.balloons()
                st.rerun() 
            except Exception as e:
                st.error(f"Erro: {e}")

# --- LISTAGEM DE RESULTADOS ---
st.markdown("---")
st.subheader("📑 Resultados na Pasta de Saída")

if os.path.exists("Analises"):
    analisados = [f for f in os.listdir("Analises") if f.endswith(".xlsx")]
    if analisados:
        for f in analisados:
            caminho_arquivo = os.path.join("Analises", f)
            with open(caminho_arquivo, "rb") as file:
                st.download_button(label=f"📥 Baixar {f}", data=file, file_name=f)
    else:
        st.write("Ainda não há análises prontas.")

st.caption("Desenvolvido por John Arllon - TI")
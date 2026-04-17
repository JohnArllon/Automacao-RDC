import streamlit as st
import os
import pandas as pd
from datetime import datetime
import time # Importado para a simulação do spinner

# Configuração da Página
st.set_page_config(page_title="Reposição RDC", layout="wide", page_icon="📊")

st.title("📊 Automação de Análise RDC")
st.markdown("---")

# --- SIDEBAR (ENTRADA DE DADOS) ---
st.sidebar.header("📅 Parâmetros de Data")

with st.sidebar:
    # Datas de Entrega
    st.subheader("Entrega")
    d_ent_ini = st.date_input("Entrega de:", datetime(2026, 4, 30))
    d_ent_fim = st.date_input("Entrega até:", datetime(2026, 5, 11))
    t15_valor = abs((d_ent_fim - d_ent_ini).days) + 1
    st.info(f"Prazo máx. (T15): **{t15_valor} dias**")
    
    st.divider()
    
    # Datas de Venda
    st.subheader("Venda")
    d_venda_ini = st.date_input("Venda de:", datetime(2026, 2, 13))
    d_venda_fim = st.date_input("Venda até:", datetime(2026, 4, 16))
    t18_valor = abs((d_venda_fim - d_venda_ini).days) + 1
    st.info(f"Período (T18): **{t18_valor} dias**")
    
    st.divider()
    
    # Outros parâmetros
    lojas_alvo = st.text_input("Lojas Alvo:", "161, 318, 328, 473, 533, 567, 582, 610, 611")
    fat_minimo = st.number_input("Fat. Mínimo Padrão (R$):", value=3500.00, step=100.0)

# --- PAINEL PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Arquivos Pendentes")
    # Lógica para listar arquivos na pasta RDCs_Originais
    if os.path.exists("RDCs_Originais"):
        arquivos = [f for f in os.listdir("RDCs_Originais") if f.endswith(".xlsx")]
        if arquivos:
            for arq in arquivos:
                st.write(f"✅ {arq}")
        else:
            st.warning("Nenhum arquivo RDC encontrado na pasta de entrada.")
    else:
        # NOVIDADE AQUI: Caso a pasta não exista, o Streamlit avisa e oferece criar
        if st.button("Criar pastas automaticamente"):
            os.makedirs("RDCs_Originais")
            os.makedirs("Analises")
            st.rerun()
        st.error("Pastas de trabalho não encontradas!")

with col2:
    st.subheader("⚙️ Execução")
    if st.button("🚀 INICIAR PROCESSAMENTO", use_container_width=True):
        
        # INTEGRAÇÃO FUTURA: Aqui chamaremos a função do main.py
        # Ex: main.processar(d_venda_ini, d_venda_fim, d_ent_ini, d_ent_fim, lojas_alvo, fat_minimo)
        
        with st.spinner("Conectando ao banco e gerando planilhas..."):
            time.sleep(2) # Simulação de carregamento
            
            # NOVIDADE AQUI: Captura o caminho onde o arquivo foi salvo
            caminho_final = os.path.abspath("Analises")
            
            st.success("Análises geradas com sucesso!")
            
            # NOVIDADE AQUI: Mostra o caminho exato para o chefe não se perder
            st.info(f"📂 **Local de salvamento:**\n\n`{caminho_final}`")
            
            st.balloons()

# NOVIDADE AQUI: Seção para listar os arquivos que já foram analisados
st.markdown("---")
st.subheader("📑 Resultados na Pasta de Saída")
if os.path.exists("Analises"):
    analisados = [f for f in os.listdir("Analises") if f.endswith(".xlsx")]
    if analisados:
        for f in analisados:
            # Botão de download opcional para cada arquivo
            with open(os.path.join("Analises", f), "rb") as file:
                st.download_button(label=f"📥 Baixar {f}", data=file, file_name=f)
    else:
        st.write("Nenhum arquivo processado ainda.")

st.markdown("---")
st.caption("Desenvolvido por John Arllon - TI")
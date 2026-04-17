import streamlit as st
import os
import pandas as pd
from datetime import datetime
import main  # <--- IMPORTANTE: Importa o seu main.py alterado

# Configuração da Página
st.set_page_config(page_title="Reposição RDC", layout="wide", page_icon="📊")

st.title("📊 Automação de Análise RDC")
st.markdown("---")

# --- SIDEBAR (ENTRADA DE DADOS) ---
st.sidebar.header("📅 Parâmetros de Data")

with st.sidebar:
    st.subheader("Entrega")
    d_ent_ini = st.date_input("Entrega de:", datetime(2026, 4, 30))
    d_ent_fim = st.date_input("Entrega até:", datetime(2026, 5, 11))
    
    st.divider()
    
    st.subheader("Venda")
    d_venda_ini = st.date_input("Venda de:", datetime(2026, 2, 13))
    d_venda_fim = st.date_input("Venda até:", datetime(2026, 4, 16))
    
    st.divider()
    
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
            st.warning("Nenhum arquivo RDC encontrado em 'RDCs_Originais'.")
    else:
        st.error("Pasta 'RDCs_Originais' não encontrada!")

with col2:
    st.subheader("⚙️ Execução")
    if st.button("🚀 INICIAR PROCESSAMENTO", use_container_width=True):
        
        # --- ALTERAÇÃO AQUI: INTEGRAÇÃO REAL COM O MAIN.PY ---
        with st.spinner("Conectando ao banco e processando... Isso pode levar alguns segundos."):
            try:
                # Chamamos a função do seu main.py passando os valores da tela
                # Convertemos date para datetime para evitar erros de cálculo no main
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
                
                # Força o Streamlit a atualizar a lista de arquivos abaixo
                st.rerun() 
                
            except Exception as e:
                st.error(f"Erro crítico no processamento: {e}")

# --- LISTAGEM DE RESULTADOS PARA DOWNLOAD ---
st.markdown("---")
st.subheader("📑 Resultados na Pasta de Saída")

if os.path.exists("Analises"):
    analisados = [f for f in os.listdir("Analises") if f.endswith(".xlsx")]
    if analisados:
        # Criamos colunas para os botões de download ficarem organizados
        for f in analisados:
            caminho_arquivo = os.path.join("Analises", f)
            with open(caminho_arquivo, "rb") as file:
                st.download_button(
                    label=f"📥 Baixar {f}",
                    data=file,
                    file_name=f,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.write("Ainda não há análises prontas para download.")

st.markdown("---")
st.caption("Desenvolvido por John Arllon - TI")
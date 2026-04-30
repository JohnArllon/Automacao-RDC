import streamlit as st
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv
load_dotenv()
import main
import importlib
importlib.reload(main)

# ==============================================================================
#                         CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Reposição RDC",
    layout="wide",
    page_icon="📊"
)

st.title("📊 Automação de Análise RDC")
st.markdown("---")

# Instruções de uso
with st.expander("📖 Como usar o sistema — clique para ver o passo a passo"):
    st.markdown("""
    **Siga a ordem abaixo para não errar:**

    **1️⃣ Baixar os RDCs da Intranet**
    - Acesse a Intranet e baixe os arquivos RDC desejados
    - Cole todos na pasta **`RDCs_Originais`**

    **2️⃣ Gerar a Análise**
    - Configure as datas e lojas na barra lateral
    - Em **Arquivos Selecionados**, escolha os RDCs que deseja analisar
    - Clique em **🚀 INICIAR PROCESSAMENTO**
    - Baixe a análise gerada na seção **Resultados na Pasta de Saída**

    **3️⃣ Preencher o Q1**
    - Abra o arquivo de análise baixado
    - Preencha a coluna **Q1** com as quantidades desejadas
    - Salve e feche o arquivo

    **4️⃣ Preencher o RDC**
    - Em **Preencher RDC com Q1**, selecione a análise preenchida
    - Selecione o RDC Original correspondente
    - Clique em **📤 PREENCHER RDC**
    - O RDC preenchido será salvo na pasta **`Prontos_Intranet`**

    **5️⃣ Subir na Intranet**
    - Pegue o arquivo da pasta **`Prontos_Intranet`**
    - Suba de volta na Intranet
    """)
    
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_RDC   = os.path.join(BASE_DIR, "RDCs_Originais")
PASTA_SAIDA = os.path.join(BASE_DIR, "Analises")
PASTA_PRONTOS = os.path.join(BASE_DIR, "Prontos_Intranet")

# ==============================================================================
#                         SIDEBAR (ENTRADA DE DADOS)
# ==============================================================================
st.sidebar.header("📅 Parâmetros de Data")

with st.sidebar:
    st.subheader("Entrega")
    d_ent_ini = st.date_input("Entrega de:", datetime(2026, 4, 28))
    d_ent_fim = st.date_input("Entrega até:", datetime(2026, 5, 11))

    hoje = datetime.now().date()
    st.write(f"Hoje é: {hoje.strftime('%d/%m/%Y')}")

    if d_ent_fim:
        dias_restantes = (d_ent_fim - hoje).days
        t15_valor = dias_restantes + 1
        st.info(f"🚚 **Prazo máx. entrega: {t15_valor} dias**")
    else:
        t15_valor = 0

    st.divider()

    st.subheader("Venda")

    # 1. Calculamos as datas dinamicamente
    hoje_dt = datetime.now().date()
    ontem_dt = hoje_dt - timedelta(days=1)
    
    # 2. Para resultar em exatamente 63 dias (considerando o +1 da fórmula)
    # Nós retrocedemos 62 dias a partir de ontem.
    # Exemplo: Se hoje é 30/04, ontem é 29/04. 
    # (29/04 - 62 dias) até 29/04 = 63 dias totais.
    venda_inicio_padrao = ontem_dt - timedelta(days=62)

    d_venda_ini = st.date_input("Venda de:", venda_inicio_padrao)
    d_venda_fim = st.date_input("Venda até:", ontem_dt)

    # 3. A conta que o Streamlit mostra (mantendo sua lógica original)
    t18_valor = abs((d_venda_fim - d_venda_ini).days) + 1
    st.warning(f"📈 **Período de venda:** {t18_valor} dias")

    st.divider()

    lojas_alvo = st.text_input("Lojas Alvo:", "161, 318, 328, 473, 533, 567, 582, 610, 611")

# ==============================================================================
#                         PAINEL PRINCIPAL — 3 COLUNAS
# ==============================================================================
arquivos_selecionados = []
col1, col2, col3 = st.columns(3)

# --- COLUNA 1: Seleção de RDCs ---
with col1:
    st.subheader("📂 Arquivos Selecionados")
    if os.path.exists(PASTA_RDC):
        arquivos_disponiveis = [f for f in os.listdir(PASTA_RDC) if f.endswith(".xlsx") and f.startswith("RDC")]
        if arquivos_disponiveis:
            arquivos_selecionados = st.multiselect(
                "Selecione os RDCs para processar:",
                options=arquivos_disponiveis,
            )
        else:
            st.warning("Nenhum arquivo RDC encontrado na pasta 'RDCs_Originais'.")
    else:
        st.error("Pasta 'RDCs_Originais' não encontrada!")

# --- COLUNA 2: Executar Análise ---
with col2:
    st.subheader("⚙️ Gerar Análise")
    st.write("**Resumo da Operação:**")
    st.write(f"- Analisando vendas de **{t18_valor} dias**.")
    st.write(f"- Garantindo estoque para **{t15_valor} dias**.")

    if st.button("🚀 INICIAR PROCESSAMENTO", use_container_width=True):
        if not arquivos_selecionados:
            st.warning("⚠️ Selecione ao menos um arquivo RDC antes de processar.")
        else:
            with st.spinner("Conectando ao banco e processando as abas..."):
                try:
                    main.rodar_automacao_v2(
                        venda_ini = datetime.combine(d_venda_ini, datetime.min.time()),
                        venda_fim = datetime.combine(d_venda_fim, datetime.max.time()),
                        ent_ini   = datetime.combine(d_ent_ini, datetime.min.time()),
                        ent_fim   = datetime.combine(d_ent_fim, datetime.max.time()),
                        lojas     = lojas_alvo,
                        prazo_t15 = t15_valor,
                        arquivos  = arquivos_selecionados
                    )
                    st.success("Análises geradas com sucesso!")
                    st.balloons()
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro crítico no processamento: {e}")

# --- COLUNA 3: Preencher RDC com Q1 ---
with col3:
    st.subheader("📝 Preencher RDC com Q1")

    analise_selecionada = ""
    rdc_selecionado = ""

    if os.path.exists(PASTA_SAIDA):
        analises_disponiveis = [f for f in os.listdir(PASTA_SAIDA) if f.startswith("Analise_") and f.endswith(".xlsx")]
        analise_selecionada = st.selectbox(
            "Selecione a Análise preenchida:",
            options=[""] + analises_disponiveis,
            format_func=lambda x: "— Selecione —" if x == "" else x
        )
    else:
        st.warning("Pasta de análises não encontrada.")

    if os.path.exists(PASTA_RDC):
        rdcs_disponiveis = [f for f in os.listdir(PASTA_RDC) if f.startswith("RDC") and f.endswith(".xlsx")]
        rdc_selecionado = st.selectbox(
            "Selecione o RDC Original:",
            options=[""] + rdcs_disponiveis,
            format_func=lambda x: "— Selecione —" if x == "" else x
        )
    else:
        st.warning("Pasta RDCs_Originais não encontrada.")

    if st.button("📤 PREENCHER RDC", use_container_width=True):
        if not analise_selecionada or not rdc_selecionado:
            st.warning("⚠️ Selecione a análise e o RDC antes de continuar.")
        else:
            with st.spinner("Preenchendo o RDC com os valores de Q1..."):
                try:
                    caminho_analise = os.path.join(PASTA_SAIDA, analise_selecionada)
                    caminho_rdc     = os.path.join(PASTA_RDC, rdc_selecionado)
                    sucesso = main.preencher_rdc_com_q1(caminho_analise, caminho_rdc)
                    if sucesso:
                        st.success("RDC preenchido com sucesso!")
                        st.balloons()
                        st.rerun()
                    else:
                        st.error("Erro ao preencher o RDC. Verifique o terminal.")
                except Exception as e:
                    st.error(f"Erro crítico: {e}")

# ==============================================================================
#                         ÁREA DE DOWNLOADS
# ==============================================================================
st.markdown("---")
st.subheader("📑 Resultados na Pasta de Saída")

if os.path.exists(PASTA_SAIDA):
    analisados = [f for f in os.listdir(PASTA_SAIDA) if f.endswith(".xlsx")]
    if analisados:
        for f in analisados:
            caminho_arquivo = os.path.join(PASTA_SAIDA, f)
            with open(caminho_arquivo, "rb") as file:
                st.download_button(
                    label=f"📥 Baixar {f}",
                    data=file,
                    file_name=f,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("Aguardando processamento de arquivos para gerar downloads.")

st.markdown("---")
st.caption("Desenvolvido por John Arllon - TI")
import streamlit as st
import os
import tempfile

from vr_module import extrair_vr_completo
from vt_module import extrair_vt_completo
from folha_module import extrair_folha_completo

# ==========================================
# DICIONÁRIO DE PLANO DE CONTAS POR PROJETO
# ==========================================
# Aqui você pode adicionar quantos projetos quiser!
MAPA_CONTAS = {
    "CENTRO POP": {
        "VR": "5.1.5. Vale Refeição",
        "VT": "5.1.4. Vale Transporte"
    },
    "NAVES": {
        "VR": "5.1.14.1. Vale Alimentação",
        "VT": "5.1.14.2. Vale Transporte"
    },
    "JOGOS ESCOLARES": {
        "VR": "5.1.24. Vale Refeição",
        "VT": "5.1.23. Vale Transporte"
    },
    "ESCOLA DE ESPORTES": {
        "VR": "5.1.24. Vale Refeição",
        "VT": "5.1.23. Vale Transporte"
    },
    "ESPORTE ATIVO 1": {
        "VR": "5.1.24. Vale Refeição",
        "VT": "5.1.23. Vale Transporte"
    },
    "ESPORTE ATIVO 2": {
        "VR": "5.1.24. Vale Refeição",
        "VT": "5.1.23. Vale Transporte"
    },
    "FAVELA COM DIGNIDADE": {
        "VR": "5.1.11. Alimentação",
        "VT": "5.1.10. Vale Transporte"
    }
}
# ==========================================

st.set_page_config(page_title="Extrator ERP - DP", page_icon="📊", layout="centered")

st.title("Processador de Importação - ERP 📊")
st.markdown("Faça o upload dos relatórios em PDF do Departamento Pessoal para gerar a planilha consolidada.")

# Seleção do Projeto
projeto_selecionado = st.selectbox(
    "Selecione o Projeto / Plano de Contas:",
    list(MAPA_CONTAS.keys())
)

tipo_extracao = st.radio(
    "Selecione o tipo de relatório que deseja processar:",
    [
        "Vale Refeição / Alimentação (VR/VA)", 
        "Vale Transporte (VT)",
        "Folha de Pagamento (Nasajon)"
    ]
)

arquivos_upados = st.file_uploader(
    "Selecione um ou mais relatórios em PDF", 
    type="pdf", 
    accept_multiple_files=True
)

if st.button("Processar Arquivos"):
    if not arquivos_upados:
        st.warning("Por favor, faça o upload de pelo menos um arquivo PDF.")
    else:
        with st.spinner("Lendo PDFs e mapeando plano de contas..."):
            
            with tempfile.TemporaryDirectory() as temp_dir:
                caminhos_pdfs = []
                
                for arquivo in arquivos_upados:
                    caminho_temp = os.path.join(temp_dir, arquivo.name)
                    with open(caminho_temp, "wb") as f:
                        f.write(arquivo.getbuffer())
                    caminhos_pdfs.append(caminho_temp)
                
                caminho_excel = os.path.join(temp_dir, "importacao_erp_consolidada.xlsx")
                
                try:
                    # ROTEAMENTO COM INJEÇÃO DA CONTA CONTÁBIL
                    if "Refeição" in tipo_extracao:
                        # Busca no dicionário a conta de VR do projeto selecionado
                        conta_vr = MAPA_CONTAS[projeto_selecionado]["VR"]
                        extrair_vr_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel, conta_contabil=conta_vr)
                        
                    elif "Transporte" in tipo_extracao:
                        # Busca no dicionário a conta de VT do projeto selecionado
                        conta_vt = MAPA_CONTAS[projeto_selecionado]["VT"]
                        extrair_vt_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel, conta_contabil=conta_vt)
                        
                    elif "Folha" in tipo_extracao:
                        # Folha não precisa de conta mapeada, então passamos normalmente
                        extrair_folha_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel)
                    
                    
                    if os.path.exists(caminho_excel):
                        with open(caminho_excel, "rb") as f:
                            bytes_excel = f.read()
                        
                        st.success(f"✅ Relatórios processados com sucesso para o projeto: **{projeto_selecionado}**!")
                        
                        st.download_button(
                            label="📥 Baixar Planilha para o ERP (.xlsx)",
                            data=bytes_excel,
                            file_name=f"importacao_{projeto_selecionado.replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("Erro: A planilha não foi gerada. Verifique se você enviou os PDFs corretos.")
                
                except Exception as e:
                    st.error(f"Ocorreu um erro inesperado durante o processamento: {e}")
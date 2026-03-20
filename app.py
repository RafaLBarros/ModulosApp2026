import streamlit as st
import os
import tempfile

# Importa as funções dos três módulos
from vr_module import extrair_vr_completo
from vt_module import extrair_vt_completo
from folha_module import extrair_folha_completo

# Configuração da página
st.set_page_config(page_title="Extrator ERP - DP", page_icon="📊", layout="centered")

st.title("Processador de Importação - ERP 📊")
st.markdown("Faça o upload dos relatórios em PDF do Departamento Pessoal para gerar a planilha consolidada de importação.")

# Seleção do tipo de módulo (Agora com a Folha de Pagamento)
tipo_extracao = st.radio(
    "Selecione o tipo de relatório que deseja processar:",
    [
        "Vale Refeição / Alimentação (VR/VA)", 
        "Vale Transporte (VT)",
        "Folha de Pagamento (Nasajon)"
    ]
)

# Upload de múltiplos arquivos
arquivos_upados = st.file_uploader(
    "Selecione um ou mais relatórios em PDF", 
    type="pdf", 
    accept_multiple_files=True
)

# Botão de ação
if st.button("Processar Arquivos"):
    if not arquivos_upados:
        st.warning("Por favor, faça o upload de pelo menos um arquivo PDF.")
    else:
        # Spinner de carregamento visual
        with st.spinner("Lendo PDFs e gerando a planilha mágica..."):
            
            # Diretório temporário para não ocupar espaço no servidor
            with tempfile.TemporaryDirectory() as temp_dir:
                caminhos_pdfs = []
                
                # Salva os arquivos upados no diretório temporário
                for arquivo in arquivos_upados:
                    caminho_temp = os.path.join(temp_dir, arquivo.name)
                    with open(caminho_temp, "wb") as f:
                        f.write(arquivo.getbuffer())
                    caminhos_pdfs.append(caminho_temp)
                
                # Destino do Excel gerado
                caminho_excel = os.path.join(temp_dir, "importacao_erp_consolidada.xlsx")
                
                try:
                    # =======================================================
                    # ROTEADOR DA INTERFACE WEB
                    # =======================================================
                    if "Refeição" in tipo_extracao:
                        extrair_vr_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel)
                    elif "Transporte" in tipo_extracao:
                        extrair_vt_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel)
                    elif "Folha" in tipo_extracao:
                        extrair_folha_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel)
                    
                    # =======================================================
                    
                    # Verifica se os módulos realmente conseguiram gerar o Excel
                    if os.path.exists(caminho_excel):
                        with open(caminho_excel, "rb") as f:
                            bytes_excel = f.read()
                        
                        st.success("✅ Relatórios processados e consolidados com sucesso!")
                        
                        # Botão de download
                        st.download_button(
                            label="📥 Baixar Planilha para o ERP (.xlsx)",
                            data=bytes_excel,
                            file_name="importacao_erp_consolidada.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("Erro: A planilha não foi gerada. Verifique se você enviou os PDFs corretos para a opção selecionada acima.")
                
                except Exception as e:
                    st.error(f"Ocorreu um erro inesperado durante o processamento: {e}")
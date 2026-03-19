import streamlit as st
import os
import tempfile

# Importa as funções dos seus módulos
from vr_module import extrair_vr_completo
from vt_module import extrair_vt_completo

# Configuração da página
st.set_page_config(page_title="Extrator ERP - VR/VT", page_icon="📊", layout="centered")

st.title("Processador de Faturas 📊")
st.markdown("Faça o upload dos relatórios em PDF para gerar a planilha consolidada de importação.")

# Seleção do tipo de módulo
tipo_extracao = st.radio(
    "Selecione o tipo de extração:",
    ["Vale Refeição / Alimentação (VR/VA)", "Vale Transporte (VT)"]
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
        # Mostra um spinner de carregamento enquanto processa
        with st.spinner("Lendo PDFs e gerando planilha..."):
            
            # Cria uma pasta temporária para salvar os PDFs e gerar o Excel
            # Isso é fundamental para não poluir o servidor quando estiver hospedado
            with tempfile.TemporaryDirectory() as temp_dir:
                caminhos_pdfs = []
                
                # Salva os arquivos upados pelo usuário na pasta temporária
                for arquivo in arquivos_upados:
                    caminho_temp = os.path.join(temp_dir, arquivo.name)
                    with open(caminho_temp, "wb") as f:
                        f.write(arquivo.getbuffer())
                    caminhos_pdfs.append(caminho_temp)
                
                # Define onde o Excel será salvo
                caminho_excel = os.path.join(temp_dir, "importacao_erp_consolidada.xlsx")
                
                try:
                    # Roteamento: chama a função correta baseada na escolha do usuário
                    if "Refeição" in tipo_extracao:
                        extrair_vr_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel)
                    else:
                        extrair_vt_completo(caminhos_pdfs, caminho_saida_excel=caminho_excel)
                    
                    # Verifica se o Excel foi realmente criado pelos módulos
                    if os.path.exists(caminho_excel):
                        # Lê o arquivo Excel gerado para a memória
                        with open(caminho_excel, "rb") as f:
                            bytes_excel = f.read()
                        
                        st.success("✅ Arquivos processados com sucesso!")
                        
                        # Cria o botão de download para o usuário
                        st.download_button(
                            label="📥 Baixar Planilha Consolidada (.xlsx)",
                            data=bytes_excel,
                            file_name="importacao_erp_consolidada.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("Erro: O arquivo Excel não foi gerado. Verifique se os PDFs são válidos.")
                
                except Exception as e:
                    st.error(f"Ocorreu um erro inesperado durante o processamento: {e}")
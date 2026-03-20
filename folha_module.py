import pdfplumber
import pandas as pd
import re
import os

def extrair_folha_completo(lista_caminhos_pdf, caminho_saida_excel=None):
    if isinstance(lista_caminhos_pdf, str):
        lista_caminhos_pdf = [lista_caminhos_pdf]
        
    if not lista_caminhos_pdf: return

    # Se não houver caminho de saída, salva na pasta do primeiro PDF
    if not caminho_saida_excel:
        base = os.path.splitext(os.path.basename(lista_caminhos_pdf[0]))[0]
        diretorio = os.path.dirname(lista_caminhos_pdf[0])
        caminho_saida_excel = os.path.join(diretorio, base + "_consolidado.xlsx")

    # Impede substituição gerando nome com sufixo numérico
    if os.path.exists(caminho_saida_excel):
        base_no_ext = os.path.splitext(caminho_saida_excel)[0]
        contador = 1
        novo = f"{base_no_ext}_{contador}.xlsx"
        while os.path.exists(novo):
            contador += 1
            novo = f"{base_no_ext}_{contador}.xlsx"
        caminho_saida_excel = novo
        
    todos_dados_consolidados = []
    
    padrao_cpf = re.compile(r'\d{3}\.\d{3}\.\d{3}-\d{2}')
    padrao_valor = re.compile(r'((?:\d{1,3}(?:\.\d{3})*|\d+)[.,]\d{2})')
    
    for caminho_pdf in lista_caminhos_pdf:
        print(f"Processando Folha de Pagamento: {os.path.basename(caminho_pdf)}...")
        dados_pdf_atual = []
        
        with pdfplumber.open(caminho_pdf) as pdf:
            # =======================================================
            # VALIDAÇÃO EXPLÍCITA (TRAVA DE SEGURANÇA)
            # =======================================================
            texto_pag1 = pdf.pages[0].extract_text().lower()
            
            # Verifica se a palavra 'listagem' está no texto (ex: "Listagem de Salários Líquidos")
            if "listagem" not in texto_pag1:
                print(f"[ERRO] O arquivo '{os.path.basename(caminho_pdf)}' não é um relatório de Folha válido. Arquivo ignorado.")
                continue
            
            for page in pdf.pages:
                texto = page.extract_text(layout=True)
                if not texto: continue
                
                for linha in texto.split('\n'):
                    # Limpa espaços excessivos gerados pelo layout=True
                    linha_limpa = re.sub(r'\s+', ' ', linha).strip()
                    
                    match_cpf = padrao_cpf.search(linha_limpa)
                    
                    if match_cpf:
                        parte_antes_cpf = linha_limpa[:match_cpf.start()].strip()
                        nome = re.sub(r'^\d+\s*[-]?\s*', '', parte_antes_cpf).strip()
                        
                        parte_depois_cpf = linha_limpa[match_cpf.end():].strip()
                        matches_valores = padrao_valor.findall(parte_depois_cpf)
                        
                        if matches_valores:
                            valor_str = matches_valores[-1]
                            
                            if nome and valor_str:
                                dados_pdf_atual.append({
                                    'ITENS DO DIÁRIO / PRODUTO': nome,
                                    'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_str
                                })

        if dados_pdf_atual:
            for i, linha in enumerate(dados_pdf_atual):
                linha['ITENS DO DIÁRIO / QUANTIDADE'] = 1
                linha['PARCEIRO'] = 'FOLHA DE PAGAMENTO' if i == 0 else ""
                
            todos_dados_consolidados.extend(dados_pdf_atual)

    if not todos_dados_consolidados:
        print("Erro: Nenhum dado encontrado. Certifique-se de enviar PDFs válidos de Folha de Pagamento.")
        return

    df = pd.DataFrame(todos_dados_consolidados)
    
    def converter_para_float(valor):
        v = str(valor).replace('R$', '').strip()
        if ',' in v and '.' in v:
            if v.rfind(',') > v.rfind('.'):
                v = v.replace('.', '').replace(',', '.')
            else:
                v = v.replace(',', '')
        elif ',' in v:
            v = v.replace(',', '.')
        return float(v)

    df['ITENS DO DIÁRIO / PREÇO UNITÁRIO'] = df['ITENS DO DIÁRIO / PREÇO UNITÁRIO'].apply(converter_para_float)

    ordem_colunas = [
        'PARCEIRO',
        'ITENS DO DIÁRIO / PRODUTO',
        'ITENS DO DIÁRIO / QUANTIDADE',
        'ITENS DO DIÁRIO / PREÇO UNITÁRIO'
    ]
    df_final = df[ordem_colunas]
    df_final.to_excel(caminho_saida_excel, index=False, engine='openpyxl')
    print(f"\n[OK] Planilha de Folha de Pagamento gerada com {len(df_final)} registros!\nSalva em: {caminho_saida_excel}")

#extrair_folha_completo("C:/Users/WINDOWS/Downloads/CIAP (1).PDF")
#extrair_folha_completo("C:/Users/WINDOWS/Downloads/REL VR FEV CIAP 20260126015703.pdf")
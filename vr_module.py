import pdfplumber
import pandas as pd
import re
import os

def extrair_vr_completo(lista_caminhos_pdf, caminho_saida_excel=None, conta_contabil='5.1.5. Vale Refeição'):
    if isinstance(lista_caminhos_pdf, str):
        lista_caminhos_pdf = [lista_caminhos_pdf]
        
    if not lista_caminhos_pdf: return

    if not caminho_saida_excel:
        base = os.path.splitext(os.path.basename(lista_caminhos_pdf[0]))[0]
        diretorio = os.path.dirname(lista_caminhos_pdf[0])
        caminho_saida_excel = os.path.join(diretorio, base + "_consolidado.xlsx")

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
    padrao_valor_caixa = re.compile(r'((?:\d{1,3}(?:\.\d{3})*|\d+)[.,]\d{2})\s*(.*?)$')
    padrao_valor_sodexo = re.compile(r'R\$\s*([\d.,]+)')
    
    for caminho_pdf in lista_caminhos_pdf:
        print(f"Processando: {os.path.basename(caminho_pdf)}...")
        dados_pdf_atual = []
        
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_pag1 = pdf.pages[0].extract_text().lower()
            if "sodexo" in texto_pag1 or "pluxee" in texto_pag1:
                tipo_layout = "sodexo"
                fornecedor = "PLUXEE"
            elif "caixa" in texto_pag1 or ("auxílio alimentação" in texto_pag1 and "relatório de detalhes do pedido" in texto_pag1):
                tipo_layout = "caixa"
                fornecedor = "Caixa Econômica Federal"
            else:
                # Se não for nenhum dos dois, ignora o arquivo e pula para o próximo
                print(f"[ERRO] O arquivo '{os.path.basename(caminho_pdf)}' não é um relatório Sodexo/Pluxee nem CAIXA válido. Arquivo ignorado.")
                continue
                
            for page in pdf.pages:
                texto = page.extract_text(layout=True)
                if not texto: continue
                
                for linha in texto.split('\n'):
                    # Mantemos a linha original intacta (com os múltiplos espaços das colunas)
                    linha_original = linha.strip()
                    # Criamos uma linha limpa (sem múltiplos espaços) para facilitar a busca do Regex
                    linha_limpa = re.sub(r'\s+', ' ', linha_original).strip()
                    
                    match_cpf = padrao_cpf.search(linha_limpa)
                    match_cpf_orig = padrao_cpf.search(linha_original)
                    
                    if match_cpf and match_cpf_orig:
                        # Corta tudo que vem DEPOIS do CPF
                        parte_antes_cpf_limpa = linha_limpa[:match_cpf.start()].strip()
                        parte_antes_cpf_orig = linha_original[:match_cpf_orig.start()].strip()
                        
                        nome = ""
                        valor_str = ""
                        
                        if tipo_layout == "caixa":
                            match_valor = padrao_valor_caixa.search(parte_antes_cpf_limpa)
                            if match_valor:
                                valor_str = match_valor.group(1)
                                nome = parte_antes_cpf_limpa[:match_valor.start()].strip()
                                
                        elif tipo_layout == "sodexo":
                            parte_depois_cpf = linha_limpa[match_cpf.end():].strip()
                            match_valor = padrao_valor_sodexo.search(parte_depois_cpf)
                            
                            if match_valor:
                                valor_str = match_valor.group(1)
                                
                                # O SEGREDO: Cortamos as colunas visuais através dos espaços duplos
                                colunas_visuais = re.split(r'\s{2,}', parte_antes_cpf_orig)
                                
                                if len(colunas_visuais) > 0:
                                    # A última coluna logo antes do CPF será sempre o Nome do Colaborador,
                                    # não importa se há departamento ou matrícula lá atrás.
                                    nome_bruto = colunas_visuais[-1].strip()
                                    
                                    # Limpeza final de segurança (remove números residuais caso existam)
                                    nome = re.sub(r'^[\d\s]+', '', nome_bruto).strip()
                                    
                                    # Exceção de segurança para o DEP. grudado:
                                    nome = nome.replace("ECOS CONJ CENTRAL", "").strip()
                                
                        if nome and valor_str:
                            dados_pdf_atual.append({
                                'ITENS DO DIÁRIO / PRODUTO': nome,
                                'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_str
                            })

        if dados_pdf_atual:
            for i, linha in enumerate(dados_pdf_atual):
                linha['ITENS DO DIÁRIO / QUANTIDADE'] = 1
                linha['ITENS DO DIÁRIO / CONTA'] = conta_contabil
                linha['PARCEIRO'] = fornecedor if i == 0 else ""
                
            todos_dados_consolidados.extend(dados_pdf_atual)

    if not todos_dados_consolidados:
        print("Erro: Nenhum dado encontrado nos arquivos fornecidos.")
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
        'ITENS DO DIÁRIO / CONTA',
        'ITENS DO DIÁRIO / QUANTIDADE',
        'ITENS DO DIÁRIO / PREÇO UNITÁRIO'
    ]
    df_final = df[ordem_colunas]
    df_final.to_excel(caminho_saida_excel, index=False, engine='openpyxl')
    print(f"\n[OK] Planilha consolidada de VR gerada com {len(df_final)} registros!\nSalva em: {caminho_saida_excel}")
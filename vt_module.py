import pdfplumber
import pandas as pd
import re
import os

def extrair_vt_completo(lista_caminhos_pdf, caminho_saida_excel=None, conta_contabil='5.1.4. Vale Transporte'):
    if isinstance(lista_caminhos_pdf, str):
        lista_caminhos_pdf = [lista_caminhos_pdf]
        
    if not lista_caminhos_pdf: return

    if not caminho_saida_excel:
        base = os.path.splitext(os.path.basename(lista_caminhos_pdf[0]))[0]
        diretorio = os.path.dirname(lista_caminhos_pdf[0])
        caminho_saida_excel = os.path.join(diretorio, base + "_consolidado.xlsx")

    # Tratamento para não sobrescrever arquivo existente
    if os.path.exists(caminho_saida_excel):
        base_no_ext = os.path.splitext(caminho_saida_excel)[0]
        contador = 1
        novo = f"{base_no_ext}_{contador}.xlsx"
        while os.path.exists(novo):
            contador += 1
            novo = f"{base_no_ext}_{contador}.xlsx"
        caminho_saida_excel = novo
        
    todos_dados_consolidados = []
    
    def converter_para_float(valor):
        v = str(valor).replace('R$', '').strip()
        if ',' in v and '.' in v:
            if v.rfind(',') > v.rfind('.'):
                v = v.replace('.', '').replace(',', '.')
            else:
                v = v.replace(',', '')
        elif ',' in v:
            v = v.replace(',', '.')
        try:
            return float(v)
        except ValueError:
            return 0.0
    
    padrao_cpf = re.compile(r'\d{3}\.\d{3}\.\d{3}-\d{2}')
    padrao_valor_mais = re.compile(r'(\d+[.,]\d{2})$')
    padrao_valor_jae = re.compile(r'R\$\s*([\d.,]+)', re.IGNORECASE)
    
    for caminho_pdf in lista_caminhos_pdf:
        print(f"Processando: {os.path.basename(caminho_pdf)}...")
        dados_pdf_atual = []
        valor_tarifa = None
        
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_pag1 = pdf.pages[0].extract_text().lower()
            
            if "relatório resumido do pedido" in texto_pag1:
                tipo_layout = "pluxee"
                fornecedor = "Sodexo"
            elif "pedido loja" in texto_pag1 or "jae" in texto_pag1 or "cbd bilhete" in texto_pag1:
                tipo_layout = "jae"
                fornecedor = "CBD BILHETE DIGITAL S/A"
            elif "relatório de resumo do pedido" in texto_pag1 or "mais.mobi" in texto_pag1:
                tipo_layout = "mais_mobi"
                fornecedor = "MAIS.MOBI"
            else:
                print(f"[ERRO] O arquivo '{os.path.basename(caminho_pdf)}' não é um relatório reconhecido.")
                continue 
                
            # ==========================================
            # MOTOR EXCLUSIVO PLUXEE
            # ==========================================
            if tipo_layout == "pluxee":
                nome_atual = ""
                valor_atual = ""
                achou_cpf = False
                
                for page in pdf.pages:
                    texto = page.extract_text()
                    if not texto: continue
                    
                    linhas = texto.split('\n')
                    for linha in linhas:
                        linha = linha.strip()
                        if not linha: continue
                        
                        match_cpf = re.search(r'(\d{11})([A-ZÀ-Ÿ\s]+?)(?=\s+(?:CIAP|Riocard|Semove|Card|Bilhete|Único|Tarifa|Variável|Verificar|\[|$))', linha)
                        
                        if match_cpf:
                            # Salva o funcionário anterior se o total não foi encontrado (fallback)
                            if achou_cpf and nome_atual and valor_atual:
                                dados_pdf_atual.append({
                                    'ITENS DO DIÁRIO / PRODUTO': nome_atual.strip(), 
                                    'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_atual
                                })
                            
                            cpf = match_cpf.group(1)
                            nome_atual = match_cpf.group(2).strip()
                            
                            # Extrai os valores iniciais (ex: 224,00) como segurança
                            valores = re.findall(r'\b\d+[.,]\d{2}\b', linha)
                            if valores:
                                valores_floats = [converter_para_float(v) for v in valores]
                                valor_vt = max(valores_floats) 
                                valor_atual = f"{valor_vt:.2f}".replace('.', ',')
                            else:
                                valor_atual = "0,00"
                                
                            achou_cpf = True
                            
                        elif achou_cpf:
                            # ========================================================
                            # TRAVA DE SEGURANÇA: ENCONTROU O TOTAL DO FUNCIONÁRIO
                            # ========================================================
                            if re.match(r'^\d+[.,]\d{2}$', linha):
                                valor_atual = linha.strip()
                                
                                # Salva imediatamente o funcionário finalizado
                                dados_pdf_atual.append({
                                    'ITENS DO DIÁRIO / PRODUTO': nome_atual.strip(), 
                                    'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_atual
                                })
                                
                                # Fecha o bloco! Isso impede que o "Total Geral" sobrescreva a Yasmin
                                achou_cpf = False 
                                nome_atual = ""
                                valor_atual = ""
                                continue
                            
                            # Continua reunindo as partes do nome cortado
                            pedacos_nome = []
                            for palavra in linha.split():
                                palavra_limpa = re.sub(r'[^a-zA-ZÀ-ÿ]', '', palavra)
                                if palavra_limpa.isupper() or palavra_limpa.upper() in ['E', 'DE', 'DA', 'DO', 'DAS', 'DOS']:
                                    pedacos_nome.append(palavra_limpa)
                                else:
                                    break
                            
                            if pedacos_nome:
                                nome_atual += " " + " ".join(pedacos_nome)

                # Salva o último funcionário caso o PDF acabe abruptamente sem a linha de total
                if achou_cpf and nome_atual and valor_atual:
                    dados_pdf_atual.append({
                        'ITENS DO DIÁRIO / PRODUTO': nome_atual.strip(), 
                        'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_atual
                    })
                    
            # ==========================================
            # MOTORES ORIGINAIS (MAIS.MOBI / JAE)
            # ==========================================
            else:
                for num_pagina, page in enumerate(pdf.pages):
                    if tipo_layout == "mais_mobi" and num_pagina == 0 and not valor_tarifa:
                        texto_layout = page.extract_text(layout=True)
                        if texto_layout:
                            linhas_layout = texto_layout.split('\n')
                            for i, linha in enumerate(linhas_layout):
                                idx = linha.lower().find('tarifa de entrega')
                                if idx != -1:
                                    if i + 1 < len(linhas_layout):
                                        inicio_corte = max(0, idx - 5)
                                        trecho_abaixo = linhas_layout[i + 1][inicio_corte:]
                                        match_abaixo = re.search(r'([\d.,]+)', trecho_abaixo)
                                        if match_abaixo:
                                            valor_tarifa = match_abaixo.group(1)
                                    break
                    
                    texto = page.extract_text()
                    if not texto: continue
                    
                    for linha in texto.split('\n'):
                        linha = linha.strip()
                        match_cpf = padrao_cpf.search(linha)
                        if match_cpf:
                            if tipo_layout == "mais_mobi":
                                parte_antes_cpf = linha[:match_cpf.start()].strip()
                                nome = re.sub(r'^\d+\s+', '', parte_antes_cpf).strip()
                                parte_depois_cpf = linha[match_cpf.end():].strip()
                                match_valor = padrao_valor_mais.search(parte_depois_cpf)
                                if match_valor:
                                    valor_str = match_valor.group(1)
                                    if nome and valor_str:
                                        dados_pdf_atual.append({'ITENS DO DIÁRIO / PRODUTO': nome, 'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_str})

                            elif tipo_layout == "jae":
                                parte_depois_cpf = linha[match_cpf.end():].strip()
                                match_valor = padrao_valor_jae.search(parte_depois_cpf)
                                if match_valor:
                                    valor_str = match_valor.group(1)
                                    nome = parte_depois_cpf[:match_valor.start()].strip()
                                    if nome and valor_str:
                                        dados_pdf_atual.append({'ITENS DO DIÁRIO / PRODUTO': nome, 'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_str})

        # Adiciona a Tarifa (Mais.Mobi) se existir
        if valor_tarifa and valor_tarifa not in ['0,00', '0.00']:
            dados_pdf_atual.append({
                'ITENS DO DIÁRIO / PRODUTO': 'TARIFA DE ENTREGA',
                'ITENS DO DIÁRIO / PREÇO UNITÁRIO': valor_tarifa
            })

        # Finaliza e formata
        if dados_pdf_atual:
            for i, linha in enumerate(dados_pdf_atual):
                linha['ITENS DO DIÁRIO / QUANTIDADE'] = 1
                linha['ITENS DO DIÁRIO / CONTA'] = conta_contabil
                linha['PARCEIRO'] = fornecedor if i == 0 else ""
                
            todos_dados_consolidados.extend(dados_pdf_atual)

    if not todos_dados_consolidados:
        print("Erro: Nenhum dado de funcionário encontrado nos arquivos fornecidos.")
        return

    df = pd.DataFrame(todos_dados_consolidados)
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
    print(f"\n[OK] Planilha consolidada de VT gerada com {len(df_final)} registros!\nSalva em: {caminho_saida_excel}")
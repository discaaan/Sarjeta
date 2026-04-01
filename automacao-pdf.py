import pdfplumber
import pandas as pd
import re
import os


pasta_pdfs = "C:/Users/Z0058B3H/Siemens Energy/Junior, Edson Roberto Caceres - Project - Vapor phase drying - Block Assembly/1 - Relatório para extratificação de dados/VP1"
caminho_excel = "C:/Users/Z0058B3H/Downloads/Dados_Transformador.xlsx"

padrao_linha = r"(\d{4,7}[/-]\d{1,2})\s+(.*?)\s+(\d+(?:,\d+)?)\s+(\d+[Aa]?)\s+(\d+)\s+(\d+(?:,\d+)?)(?:\s+(\d+(?:,\d+)?))?(?=\s|$|DATA|DURAÇÃO|ESTUFA|SECAGEM)"
padrao_inicial = r"DATA INICIAL:\s*(\d{2}/\d{2}/\d{2}\s*\d{2}:\d{2})"
padrao_final = r"DATA FINAL:\s*(\d{2}/\d{2}/\d{2}\s*\d{2}:\d{2})"
padrao_duracao = r"DURAÇÃO\s*\[horas\]:\s*(\d+)"

def processar_pdf(caminho_completo_pdf, nome_arquivo):
    dados_extraidos = []
    
    try:
        with pdfplumber.open(caminho_completo_pdf) as pdf:
            primeira_pag = pdf.pages[0]
            texto_limpo = primeira_pag.extract_text()
            
        if not texto_limpo:
            return dados_extraidos, f"{nome_arquivo} -> Sem texto."

       
        busca_inicial = re.search(padrao_inicial, texto_limpo)
        busca_final = re.search(padrao_final, texto_limpo)
        data_ini = busca_inicial.group(1) if busca_inicial else ""
        data_fin = busca_final.group(1) if busca_final else ""

    
        busca_duracao = re.search(padrao_duracao, texto_limpo)
        duracao = busca_duracao.group(1) if busca_duracao else ""

       
        matches = list(re.finditer(padrao_linha, texto_limpo))
        
        if not matches:
            return dados_extraidos, f"{nome_arquivo} -> Não encontrou o padrão no texto."



        massa_anterior = "Verificar no PDF"


        for match in matches:
            pa = match.group(1)
            texto_meio = match.group(2).strip()
            potencia = match.group(3)
            tensao = match.group(4)
            nr_abaix = match.group(5)
            mat_isol = match.group(6)
            
      
            if match.group(7):
                massa_total = match.group(7)
                massa_anterior = massa_total 
            else:
                massa_total = massa_anterior 
           
            partes_texto = texto_meio.split()
            if len(partes_texto) > 1:
                carregamento = partes_texto[-1]
                cliente = " ".join(partes_texto[:-1])
            else:
                carregamento = texto_meio
                cliente = "Não Identificado"

            dados_extraidos.append([
                pa, cliente, carregamento, potencia, tensao, 
                nr_abaix, mat_isol, massa_total, data_ini, data_fin, duracao, nome_arquivo
            ])
            
        return dados_extraidos, None

    except Exception as e:
        return dados_extraidos, f"{nome_arquivo} -> Erro inesperado: {e}"

lista_de_dados_geral = []
arquivos_com_erro = []

for nome_arquivo in os.listdir(pasta_pdfs):
    if nome_arquivo.lower().endswith(".pdf"):
        caminho_completo_pdf = os.path.join(pasta_pdfs, nome_arquivo)
        
        dados_do_arquivo, erro = processar_pdf(caminho_completo_pdf, nome_arquivo)
        
        if dados_do_arquivo:
            lista_de_dados_geral.extend(dados_do_arquivo)
            
        if erro:
            arquivos_com_erro.append(erro)

colunas = [
    "P.A.", "CLIENTE", "CARREGAMENTO", "POT. [MVA]", 
    "TENS. [kV]", "Nr. Abaix.", "MAT. ISOL [kg]", "MASSA TOTAL [Ton]",
    "DATA INICIAL", "DATA FINAL", "DURAÇÃO [horas]", "NOME_DO_ARQUIVO"
]

if lista_de_dados_geral:
    tabela_consolidada = pd.DataFrame(lista_de_dados_geral, columns=colunas)
    tabela_consolidada.to_excel(caminho_excel, index=False)
    print(f"\nProcesso concluído! Sucesso total.")
else:
    print("\nNenhum dado extraído.")

if arquivos_com_erro:
    print("\nArquivos com problemas:")
    for erro in arquivos_com_erro:
        print(erro)

print("\n" + "="*50)
print("RELATÓRIO DE LEITURA:")
print(f"Linhas extraídas com sucesso: {len(lista_de_dados_geral)}")
print(f"Arquivos com falha/ignorados: {len(arquivos_com_erro)}")
print("="*50)
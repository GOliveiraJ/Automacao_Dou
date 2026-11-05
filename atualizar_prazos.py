import os
import re
import pdfplumber
import openpyxl

# --- CONFIGURAÇÕES ---
ARQUIVO_EXCEL = "(COPIA DE SEGURANÇA) CONTROLE DOS PRAZOS DE ANÁLISE_CCTAB_2025.xlsx"
PASTA_PDFS = "arquivos_dou"
NOME_COLUNA_EXPEDIENTE = "EXPEDIENTE"
NOME_COLUNA_DATA_PUB = "DATA DE PUBLICAÇÃO"

# Lista exata das abas das pessoas monitoradas (com e sem acento por segurança)
ABAS_ALVO = ["MÁRCIA", "PATRÍCIA", "THAÍLA", "CLEICI", "MARCIA", "PATRICIA", "THAILA", "CLEICI"]

# Mapa de meses para conversão
MESES = {
    'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04',
    'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08',
    'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'
}

def limpar_expediente(texto):
    """Remove espaços e caracteres estranhos para garantir o match"""
    if texto is None:
        return ""
    return str(texto).replace(" ", "").strip()

def extrair_dados_pdf(caminho_pdf):
    texto_completo = ""
    data_encontrada = None
    expedientes = set()

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"
    except Exception as e:
        print(f"Erro ao ler PDF: {e}")
        return None, []

    # 1. BUSCAR DATA NO CABEÇALHO DO PDF
    match_data = re.search(r'Brasília.*?(\d{1,2})\s+de\s+([a-zA-Zç]+)\s+de\s+(\d{4})', texto_completo, re.IGNORECASE)
    if match_data:
        dia, mes_extenso, ano = match_data.groups()
        mes_num = MESES.get(mes_extenso.lower())
        if mes_num:
            data_encontrada = f"{int(dia):02d}/{mes_num}/{ano}"

    # 2. BUSCAR EXPEDIENTES
    matches = re.findall(r'(\d{7}\s*/\s*\d{2}\s*-\s*\d)', texto_completo)
    for m in matches:
        expedientes.add(limpar_expediente(m))

    return data_encontrada, expedientes

def main():
    print("=" * 70)
    print(" ATUALIZADOR DE PRAZOS (EXCEL) ".center(70))
    print("=" * 70 + "\n")

    if not os.path.exists(PASTA_PDFS):
        print(f"ERRO: Pasta '{PASTA_PDFS}' não encontrada.")
        return

    # 1. Mapear dados dos PDFs
    print("1. Lendo os arquivos PDF e extraindo as datas de publicação...")
    mapa_expediente_data = {}
    arquivos_pdf = [f for f in os.listdir(PASTA_PDFS) if f.lower().endswith('.pdf')]
    
    if not arquivos_pdf:
        print("Nenhum PDF encontrado na pasta.")
        return

    for arq in arquivos_pdf:
        caminho = os.path.join(PASTA_PDFS, arq)
        data, lista_exps = extrair_dados_pdf(caminho)
        if data:
            for exp in lista_exps:
                mapa_expediente_data[exp] = data
    
    if not mapa_expediente_data:
        print("Nenhum dado extraído dos PDFs. Encerrando.")
        return

    # 2. Atualizar o Excel mantendo fórmulas
    print(f"\n2. Atualizando o arquivo Excel: {ARQUIVO_EXCEL}")
    print("-" * 70)
    
    try:
        wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
        total_atualizados = 0

        for sheet in wb.worksheets:
            # ---> FILTRO PRINCIPAL: Processa apenas as abas das 4 responsáveis <---
            if sheet.title.strip().upper() not in ABAS_ALVO:
                continue
                
            col_exp_index = None
            col_data_index = None
            
            # Localizar índices das colunas no cabeçalho
            for cell in sheet[1]: 
                if cell.value and str(cell.value).strip() == NOME_COLUNA_EXPEDIENTE:
                    col_exp_index = cell.column
                elif cell.value and str(cell.value).strip() == NOME_COLUNA_DATA_PUB:
                    col_data_index = cell.column
            
            if col_exp_index and col_data_index:
                # Itera sobre as linhas
                for row in sheet.iter_rows(min_row=2):
                    cell_exp = row[col_exp_index - 1]
                    cell_data = row[col_data_index - 1]
                    
                    valor_exp = limpar_expediente(cell_exp.value)
                    
                    if valor_exp in mapa_expediente_data:
                        # Só preenche se a célula da data estiver vazia
                        if not cell_data.value:
                            nova_data = mapa_expediente_data[valor_exp]
                            cell_data.value = nova_data
                            
                            # ---> MENSAGEM PERSONALIZADA NA TELA <---
                            linha_atual = cell_exp.row
                            nome_aba = sheet.title
                            print(f"> o expediente {valor_exp} linha {linha_atual} teve sua data de publicação atualizada para {nova_data} na aba {nome_aba}")
                            
                            total_atualizados += 1

        print("-" * 70)
        
        if total_atualizados > 0:
            wb.save("CONTROLE_ATUALIZADO_COM_FORMULAS.xlsx")
            print(f"SUCESSO! {total_atualizados} linhas foram atualizadas.")
            print("Arquivo salvo como: CONTROLE_ATUALIZADO_COM_FORMULAS.xlsx")
        else:
            print("Nenhum expediente novo precisava de atualização na planilha.")

    except Exception as e:
        print(f"Erro crítico no Excel: {e}")

if __name__ == "__main__":
    main()
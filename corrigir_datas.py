import openpyxl

ARQUIVO_EXCEL = "(COPIA DE SEGURANÇA) CONTROLE DOS PRAZOS DE ANÁLISE_CCTAB_2025.xlsx"
NOME_COLUNA_ALVO = "DATA FINAL PARA PUBLICAÇÃO"

def corrigir_formulas_1900_definitivo():
    print(f"--- A corrigir falha matemática das Exigências Pendentes ---")
    wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
    
    for sheet in wb.worksheets:
        if sheet.title == "VALIDAÇÃO":
            continue
            
        col_alvo_idx = None
        for cell in sheet[1]:
            if cell.value and NOME_COLUNA_ALVO in str(cell.value).strip():
                col_alvo_idx = cell.column
                break
        
        if col_alvo_idx:
            for row in range(2, sheet.max_row + 1):
                # Fórmula corrigida utilizando MAX(0, ...) para evitar números negativos
                nova_formula = (
                    f'=IF(OR(A{row}="", F{row}="", F{row}=0), "", '
                    f'IF(A{row}="REGISTRO", '
                    f'F{row}+90 + MAX(0, J{row}-I{row}) + MAX(0, L{row}-K{row}) + MAX(0, N{row}-M{row}) + MAX(0, P{row}-O{row}) + MAX(0, R{row}-Q{row}), '
                    f'F{row}+150 + MAX(0, J{row}-I{row}) + MAX(0, L{row}-K{row}) + MAX(0, N{row}-M{row}) + MAX(0, P{row}-O{row}) + MAX(0, R{row}-Q{row})))'
                )
                sheet.cell(row=row, column=col_alvo_idx).value = nova_formula
                
    wb.save("CONTROLE_CORRIGIDO_1900.xlsx")
    print("Concluído! O ficheiro CONTROLE_CORRIGIDO_1900.xlsx foi gerado com sucesso.")

if __name__ == "__main__":
    corrigir_formulas_1900_definitivo()
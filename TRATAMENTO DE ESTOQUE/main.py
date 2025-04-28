import pandas as pd
import os
from datetime import datetime
import math

# Definir os nomes dos meses por extenso
meses = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
]

# Obter a data atual
hoje = datetime.now()
mes_extenso = meses[hoje.month - 1]

# Carregar a planilha CSV
df = pd.read_csv("TRATAMENTO DE ESTOQUE/products_export_1.csv")

# Selecionar as colunas relevantes
df_filtered = df[["Handle", "Title", "Vendor", "Variant Inventory Qty", "Cost per item", "Variant Price"]].copy()

# Renomear colunas
df_filtered.rename(columns={
    "Title": "Produto",
    "Vendor": "Marca",
    "Variant Inventory Qty": "Quantidade",
    "Cost per item": "Custo",
    "Variant Price": "PVD"
}, inplace=True)

# Excluir produtos cujo nome é "teste" (ignorando letras maiúsculas/minúsculas)
df_filtered = df_filtered[df_filtered["Produto"].str.lower() != "teste"]

# Converter valores para numérico e tratar NaN
df_filtered["Quantidade"] = pd.to_numeric(df_filtered["Quantidade"], errors="coerce").fillna(0)
df_filtered["PVD"] = pd.to_numeric(df_filtered["PVD"], errors="coerce").fillna(0)
df_filtered["Custo"] = pd.to_numeric(df_filtered["Custo"], errors="coerce").fillna(0)

# Converter todas as marcas para maiúsculas
df_filtered["Marca"] = df_filtered["Marca"].str.upper()

# Agrupar pelo campo "Handle" para consolidar todas as variações do mesmo produto
df_simplificado = df_filtered.groupby("Handle", as_index=False).agg({
    "Produto": "first",
    "Marca": "first",
    "Quantidade": "sum",
    "Custo": "first",
    "PVD": "first"
})

# Criar as colunas 'Total PVD' e 'Total Custo' (valores atuais; serão substituídos por fórmulas)
df_simplificado["Total PVD"] = df_simplificado["Quantidade"] * df_simplificado["PVD"]
df_simplificado["Total Custo"] = df_simplificado["Quantidade"] * df_simplificado["Custo"]

# Criar a coluna 'MKUP' (PVD / Custo) com 3 casas decimais (valor atual; depois sobrescrito)
df_simplificado["MKUP"] = df_simplificado.apply(
    lambda row: round(row["PVD"] / row["Custo"], 3) if row["Custo"] != 0 else 0, axis=1
)

# Criar a coluna 'LUCRO' (Total PVD - Total Custo)
df_simplificado["LUCRO"] = df_simplificado["Total PVD"] - df_simplificado["Total Custo"]

# Filtrar para excluir produtos com Quantidade menor ou igual a 0
df_simplificado = df_simplificado[df_simplificado["Quantidade"] > 0].copy()

# Remover a coluna "Handle" da Tabela Principal
df_simplificado.drop(columns=["Handle"], inplace=True)

# ------------------------
# Tabela 1: Média de MKUP por Marca
df_mkup = df_simplificado.groupby("Marca", as_index=False).agg({"MKUP": "mean"})
df_mkup.rename(columns={"MKUP": "Média MKUP"}, inplace=True)
df_mkup["Média MKUP"] = df_mkup["Média MKUP"].round(3)

# ------------------------
# Tabela 2: Top 15 Produtos com Maior Total Custo
df_top = df_simplificado.sort_values("Total Custo", ascending=False).head(15)

# ------------------------
# Configurar pasta de destino e nome do arquivo
output_folder = os.path.join(".", "PLANILHAS DE ESTOQUE")
os.makedirs(output_folder, exist_ok=True)
nome_arquivo = f"PLANILHA_ESTOQUE_URB_LAB_{hoje.day} de {mes_extenso}.xlsx"
output_excel = os.path.join(output_folder, nome_arquivo)

# Salvar os resultados em um único arquivo Excel com múltiplas tabelas
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    # Tabela Principal na aba "Estoque"
    df_simplificado.to_excel(writer, sheet_name='Estoque', index=False, startrow=0, startcol=0)
    
    workbook  = writer.book
    worksheet = writer.sheets['Estoque']
    
    # Formatação para os cabeçalhos
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    
    # Formato de moeda (Real)
    currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
    
    # Formato para números com 3 casas decimais (MKUP)
    num_format = workbook.add_format({'num_format': '0.000'})
    
    # Obter dimensões da Tabela Principal
    n_rows, n_cols = df_simplificado.shape
    
    # Aplicar formatação aos cabeçalhos da Tabela Principal e ajustar colunas
    for col_num, value in enumerate(df_simplificado.columns.values):
        worksheet.write(0, col_num, value, header_format)
        column_len = df_simplificado[value].astype(str).map(len).max()
        column_len = max(column_len, len(value)) + 2
        worksheet.set_column(col_num, col_num, column_len)
    
    # Aplicar formato de moeda às colunas monetárias da Tabela Principal
    for col in ["Custo", "PVD"]:
        col_idx = df_simplificado.columns.get_loc(col)
        worksheet.set_column(col_idx, col_idx, 12, currency_format)
    
    # Aplicar formato padrão para a coluna Quantidade
    col_idx = df_simplificado.columns.get_loc("Quantidade")
    worksheet.set_column(col_idx, col_idx, 10)
    
    # Adicionar filtros à Tabela Principal
    worksheet.autofilter(0, 0, n_rows, n_cols - 1)
    
    # Congelar o cabeçalho (primeira linha) para que fique sempre visível
    worksheet.freeze_panes(1, 0)
    
    # ------------------------
    # Atualizar células para fórmulas dinâmicas (os dados abaixo serão recalculados se PVD ou outros forem alterados)
    #
    # Considerando a ordem das colunas:
    # 0: Produto, 1: Marca, 2: Quantidade, 3: Custo, 4: PVD,
    # 5: Total PVD, 6: Total Custo, 7: MKUP, 8: LUCRO
    #
    # Lembre que a primeira linha (linha 0) é o cabeçalho; os dados começam na linha 1.
    # Em A1 notation (Excel), as linhas são 1-indexadas; portanto, a primeira linha de dados é a linha 2.
    for i in range(n_rows):
        excel_row = i + 2  # Ajusta para a numeração do Excel (cabeçalho na linha 1)
        # Total PVD = Quantidade (coluna C) * PVD (coluna E)
        worksheet.write_formula(i+1, 5, f"=C{excel_row}*E{excel_row}", currency_format)
        # Total Custo = Quantidade (coluna C) * Custo (coluna D)
        worksheet.write_formula(i+1, 6, f"=C{excel_row}*D{excel_row}", currency_format)
        # MKUP = IF(Custo ≠ 0, PVD / Custo, 0)  => =IF(D{row}<>0,E{row}/D{row},0)
        worksheet.write_formula(i+1, 7, f"=IF(D{excel_row}<>0,E{excel_row}/D{excel_row},0)", num_format)
        # LUCRO = Total PVD - Total Custo  => =F{row} - G{row}
        worksheet.write_formula(i+1, 8, f"=F{excel_row}-G{excel_row}", currency_format)
    
    # ------------------------
    # Adicionar totais com SUBTOTAL na Tabela Principal
    subtotal_row = n_rows + 1  # Linha onde vamos escrever os totais
    worksheet.write(subtotal_row, 0, "TOTAIS (SUBTOTAL)", header_format)

    # Mapear colunas para subtotal (as fórmulas ignoram linhas ocultas)
    subtotal_cols = {
        "Quantidade": "Quantidade",
        "Total PVD": "Total PVD",
        "Total Custo": "Total Custo",
        "LUCRO": "LUCRO"
    }

    for col_name in subtotal_cols:
        col_idx = df_simplificado.columns.get_loc(col_name)
        col_letter = chr(ord('A') + col_idx)  # converte índice para letra (até Z)
        formula = f'=SUBTOTAL(109,{col_letter}2:{col_letter}{n_rows + 1})'
        cell_format = currency_format if "Total" in col_name or col_name == "LUCRO" else None
        worksheet.write_formula(subtotal_row, col_idx, formula, cell_format)

    # ------------------------
    # Tabela 1: Média de MKUP por Marca, à direita da Tabela Principal
    start_col_2 = n_cols + 2  # Espaçamento de 2 colunas
    df_mkup.to_excel(writer, sheet_name='Estoque', index=False, startrow=0, startcol=start_col_2)
    
    # Aplicar formatação aos cabeçalhos da Tabela 1
    for col_num, value in enumerate(df_mkup.columns.values):
        worksheet.write(0, start_col_2 + col_num, value, header_format)
        column_len = df_mkup[value].astype(str).map(len).max()
        column_len = max(column_len, len(value)) + 2
        worksheet.set_column(start_col_2 + col_num, start_col_2 + col_num, column_len)
    
    # ------------------------
    # Tabela 2: Top 15 Produtos com Maior Total Custo, abaixo da Tabela Principal
    start_row_3 = subtotal_row + 3  # Começa após a linha de subtotais
    df_top.to_excel(writer, sheet_name='Estoque', index=False, startrow=start_row_3, startcol=0)
    
    # Aplicar formatação aos cabeçalhos da Tabela 2
    for col_num, value in enumerate(df_top.columns.values):
        worksheet.write(start_row_3, col_num, value, header_format)
        column_len = df_top[value].astype(str).map(len).max()
        column_len = max(column_len, len(value)) + 2
        worksheet.set_column(col_num, col_num, column_len)
    
    # Aplicar formato de moeda às colunas monetárias da Tabela 2
    for col in ["Custo", "PVD", "Total PVD", "Total Custo", "LUCRO"]:
        col_idx = df_top.columns.get_loc(col)
        worksheet.set_column(col_idx, col_idx, 12, currency_format)

print(f"Arquivo salvo em: {output_excel} 🚀")

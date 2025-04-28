import pandas as pd
import os
from datetime import datetime
import math

# List of month names
meses = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
]

# Get current date information
hoje = datetime.now()
mes_extenso = meses[hoje.month - 1]

# Load CSV file
df = pd.read_csv("TRATAMENTO DE ESTOQUE/products_export_1.csv")

# Select relevant columns
df_filtered = df[["Handle", "Title", "Vendor", "Variant Inventory Qty", "Cost per item", "Variant Price"]].copy()

# Rename columns for clarity
df_filtered.rename(columns={
    "Title": "Produto",
    "Vendor": "Marca",
    "Variant Inventory Qty": "Quantidade",
    "Cost per item": "Custo",
    "Variant Price": "PVD"
}, inplace=True)

# Remove products named "teste" (case-insensitive)
df_filtered = df_filtered[df_filtered["Produto"].str.lower() != "teste"]

# Convert columns to numeric and handle missing values
df_filtered["Quantidade"] = pd.to_numeric(df_filtered["Quantidade"], errors="coerce").fillna(0)
df_filtered["PVD"] = pd.to_numeric(df_filtered["PVD"], errors="coerce").fillna(0)
df_filtered["Custo"] = pd.to_numeric(df_filtered["Custo"], errors="coerce").fillna(0)

# Standardize brand names to uppercase
df_filtered["Marca"] = df_filtered["Marca"].str.upper()

# Aggregate data by product Handle
df_simplificado = df_filtered.groupby("Handle", as_index=False).agg({
    "Produto": "first",
    "Marca": "first",
    "Quantidade": "sum",
    "Custo": "first",
    "PVD": "first"
})

# Calculate Total PVD and Total Cost
df_simplificado["Total PVD"] = df_simplificado["Quantidade"] * df_simplificado["PVD"]
df_simplificado["Total Custo"] = df_simplificado["Quantidade"] * df_simplificado["Custo"]

# Calculate Markup (MKUP)
df_simplificado["MKUP"] = df_simplificado.apply(
    lambda row: round(row["PVD"] / row["Custo"], 3) if row["Custo"] != 0 else 0, axis=1
)

# Calculate Profit (LUCRO)
df_simplificado["LUCRO"] = df_simplificado["Total PVD"] - df_simplificado["Total Custo"]

# Filter out products with zero or negative stock
df_simplificado = df_simplificado[df_simplificado["Quantidade"] > 0].copy()

# Drop Handle column
df_simplificado.drop(columns=["Handle"], inplace=True)

# ------------------------
# Table 1: Average Markup per Brand
df_mkup = df_simplificado.groupby("Marca", as_index=False).agg({"MKUP": "mean"})
df_mkup.rename(columns={"MKUP": "Média MKUP"}, inplace=True)
df_mkup["Média MKUP"] = df_mkup["Média MKUP"].round(3)

# ------------------------
# Table 2: Top 15 products by Total Cost
df_top = df_simplificado.sort_values("Total Custo", ascending=False).head(15)

# ------------------------
# Define output folder and file name
output_folder = os.path.join(".", "PLANILHAS DE ESTOQUE")
os.makedirs(output_folder, exist_ok=True)
nome_arquivo = f"PLANILHA_ESTOQUE_URB_LAB_{hoje.day} de {mes_extenso}.xlsx"
output_excel = os.path.join(output_folder, nome_arquivo)

# ------------------------
# Save results to an Excel file with multiple tables
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    # Write main stock table
    df_simplificado.to_excel(writer, sheet_name='Estoque', index=False, startrow=0, startcol=0)
    
    workbook  = writer.book
    worksheet = writer.sheets['Estoque']
    
    # Format headers
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    
    # Format for currency (R$)
    currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
    
    # Format for numbers with 3 decimals
    num_format = workbook.add_format({'num_format': '0.000'})
    
    n_rows, n_cols = df_simplificado.shape
    
    # Apply formatting to headers and adjust column widths
    for col_num, value in enumerate(df_simplificado.columns.values):
        worksheet.write(0, col_num, value, header_format)
        column_len = df_simplificado[value].astype(str).map(len).max()
        column_len = max(column_len, len(value)) + 2
        worksheet.set_column(col_num, col_num, column_len)
    
    # Apply currency formatting to cost columns
    for col in ["Custo", "PVD"]:
        col_idx = df_simplificado.columns.get_loc(col)
        worksheet.set_column(col_idx, col_idx, 12, currency_format)
    
    # Adjust quantity column width
    col_idx = df_simplificado.columns.get_loc("Quantidade")
    worksheet.set_column(col_idx, col_idx, 10)
    
    # Add autofilter and freeze header
    worksheet.autofilter(0, 0, n_rows, n_cols - 1)
    worksheet.freeze_panes(1, 0)
    
    # ------------------------
    # Add dynamic formulas for Total PVD, Total Cost, MKUP, and LUCRO
    for i in range(n_rows):
        excel_row = i + 2
        worksheet.write_formula(i+1, 5, f"=C{excel_row}*E{excel_row}", currency_format)
        worksheet.write_formula(i+1, 6, f"=C{excel_row}*D{excel_row}", currency_format)
        worksheet.write_formula(i+1, 7, f"=IF(D{excel_row}<>0,E{excel_row}/D{excel_row},0)", num_format)
        worksheet.write_formula(i+1, 8, f"=F{excel_row}-G{excel_row}", currency_format)
    
    # ------------------------
    # Add SUBTOTAL formulas at the bottom
    subtotal_row = n_rows + 1
    worksheet.write(subtotal_row, 0, "TOTAIS (SUBTOTAL)", header_format)

    subtotal_cols = {
        "Quantidade": "Quantidade",
        "Total PVD": "Total PVD",
        "Total Custo": "Total Custo",
        "LUCRO": "LUCRO"
    }

    for col_name in subtotal_cols:
        col_idx = df_simplificado.columns.get_loc(col_name)
        col_letter = chr(ord('A') + col_idx)
        formula = f'=SUBTOTAL(109,{col_letter}2:{col_letter}{n_rows + 1})'
        cell_format = currency_format if "Total" in col_name or col_name == "LUCRO" else None
        worksheet.write_formula(subtotal_row, col_idx, formula, cell_format)

    # ------------------------
    # Write Table 1: Average MKUP by brand
    start_col_2 = n_cols + 2
    df_mkup.to_excel(writer, sheet_name='Estoque', index=False, startrow=0, startcol=start_col_2)

    for col_num, value in enumerate(df_mkup.columns.values):
        worksheet.write(0, start_col_2 + col_num, value, header_format)
        column_len = df_mkup[value].astype(str).map(len).max()
        column_len = max(column_len, len(value)) + 2
        worksheet.set_column(start_col_2 + col_num, start_col_2 + col_num, column_len)

    # ------------------------
    # Write Table 2: Top 15 products with highest Total Cost
    start_row_3 = subtotal_row + 3
    df_top.to_excel(writer, sheet_name='Estoque', index=False, startrow=start_row_3, startcol=0)

    for col_num, value in enumerate(df_top.columns.values):
        worksheet.write(start_row_3, col_num, value, header_format)
        column_len = df_top[value].astype(str).map(len).max()
        column_len = max(column_len, len(value)) + 2
        worksheet.set_column(col_num, col_num, column_len)

    for col in ["Custo", "PVD", "Total PVD", "Total Custo", "LUCRO"]:
        col_idx = df_top.columns.get_loc(col)
        worksheet.set_column(col_idx, col_idx, 12, currency_format)

print(f"Arquivo salvo em: {output_excel} 🚀")

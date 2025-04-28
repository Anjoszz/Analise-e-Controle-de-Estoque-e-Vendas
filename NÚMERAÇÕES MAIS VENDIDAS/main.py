import pandas as pd
import numpy as np
import xlsxwriter

# ============================
# Step 1: Load CSV and prepare data
# ============================
file_path = r"C:\Users\55119\Desktop\SISTEMAS DE TOMADA DE DECISÃO\NÚMERAÇÕES MAIS VENDIDAS\NÚMERAÇÕES MAIS VENDIDAS/NÚMERAÇÕES MAIS VENDIDAS/Total de vendas por produto - 2025-03-01 - 2025-04-25.csv"
df = pd.read_csv(file_path)

# Strip extra spaces from column names
df.columns = df.columns.str.strip()

# Validate expected columns
colunas_esperadas = ['Mês', 'Fornecedor do produto', 'Título da variante do produto', 'Itens líquidos vendidos']
for col in colunas_esperadas:
    if col not in df.columns:
        raise ValueError(f"Coluna '{col}' não encontrada no CSV. Verifique o nome exato.")

# Select and rename columns
df_reduced = df[['Mês', 'Fornecedor do produto', 'Título da variante do produto', 'Itens líquidos vendidos']]
df_reduced.columns = ['Mes', 'Marca', 'Numeracao', 'QtdVendida']

# Drop rows with missing numeration
df_reduced = df_reduced.dropna(subset=['Numeracao'])

# ============================
# Step 2: Group data
# ============================
df_grouped = df_reduced.groupby(['Mes', 'Marca', 'Numeracao'], as_index=False).agg({'QtdVendida': 'sum'})

# ============================
# Step 3A: Top 5 by Brand (per month)
# ============================
def extrai_top5(sub_df):
    return sub_df.sort_values('QtdVendida', ascending=False).head(5)

lista_top5_marca = []
meses = sorted(df_grouped['Mes'].unique())
for mes in meses:
    df_mes = df_grouped[df_grouped['Mes'] == mes]
    marcas = df_mes['Marca'].unique()
    for marca in marcas:
        df_marca = df_mes[df_mes['Marca'] == marca]
        top5 = extrai_top5(df_marca)
        lista_top5_marca.append(top5)

df_top5_marca = pd.concat(lista_top5_marca, ignore_index=True)

# ============================
# Step 3B: Top 5 by Month (all brands)
# ============================
df_top5_mes = df_reduced.groupby(['Mes', 'Numeracao'], as_index=False).agg({'QtdVendida': 'sum'})
df_top5_mes = df_top5_mes.groupby('Mes', group_keys=False).apply(lambda x: x.sort_values('QtdVendida', ascending=False).head(5)).reset_index(drop=True)

# ============================
# Step 4: Sales report by numeration
# ============================
df_final = df_reduced.copy()
df_final['Numeracao'] = pd.to_numeric(df_final['Numeracao'], errors='coerce')
df_final = df_final.dropna(subset=['Numeracao'])

relatorio_numeracoes = df_final.groupby('Numeracao', as_index=False).agg({'QtdVendida': 'sum'})
relatorio_numeracoes = relatorio_numeracoes.sort_values('QtdVendida', ascending=False)

total_vendido = relatorio_numeracoes['QtdVendida'].sum()
relatorio_numeracoes['Percentual'] = (relatorio_numeracoes['QtdVendida'] / total_vendido) * 100

# ============================
# Step 5: Gender classification and relative percentages
# ============================
def classify_gender(num):
    if 34 <= num <= 39:
        return 'Feminino'
    elif 40 <= num <= 44:
        return 'Masculino'
    else:
        return 'Outros'

relatorio_numeracoes['Genero'] = relatorio_numeracoes['Numeracao'].apply(classify_gender)

# Split by gender
comparacao_feminino = relatorio_numeracoes[relatorio_numeracoes['Genero'] == 'Feminino'].copy()
comparacao_masculino = relatorio_numeracoes[relatorio_numeracoes['Genero'] == 'Masculino'].copy()

# Calculate relative percentages within each gender
total_feminino = comparacao_feminino['QtdVendida'].sum()
total_masculino = comparacao_masculino['QtdVendida'].sum()

comparacao_feminino['Percentual'] = (comparacao_feminino['QtdVendida'] / total_feminino) * 100
comparacao_masculino['Percentual'] = (comparacao_masculino['QtdVendida'] / total_masculino) * 100

# Overall gender summary
df_genero = relatorio_numeracoes[relatorio_numeracoes['Genero'].isin(['Feminino', 'Masculino'])]
resumo_genero = df_genero.groupby('Genero', as_index=False).agg({'QtdVendida': 'sum'})
resumo_genero['Percentual'] = (resumo_genero['QtdVendida'] / total_vendido) * 100

# ============================
# Step 6: Export results to Excel
# ============================
output_file = "Relatorio_Vendas_Numeracoes.xlsx"
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

# --- Sheet: Relatorio_Final
relatorio_numeracoes.to_excel(writer, sheet_name='Relatorio_Final', index=False)
worksheet_rf = writer.sheets['Relatorio_Final']

# Add sales by numeration chart
chart1 = workbook.add_chart({'type': 'column'})
n_rows = relatorio_numeracoes.shape[0]
chart1.add_series({
    'name': 'QtdVendida',
    'categories': ['Relatorio_Final', 1, 0, n_rows, 0],
    'values': ['Relatorio_Final', 1, 1, n_rows, 1],
})
chart1.set_title({'name': 'Vendas por Numeração'})
chart1.set_x_axis({'name': 'Numeração'})
chart1.set_y_axis({'name': 'Quantidade Vendida'})
worksheet_rf.insert_chart('F2', chart1, {'x_offset': 25, 'y_offset': 10})

# --- Sheet: Resumo_Genero
resumo_genero.to_excel(writer, sheet_name='Resumo_Genero', index=False)
worksheet_rg = writer.sheets['Resumo_Genero']

# Add gender distribution pie chart
chart2 = workbook.add_chart({'type': 'pie'})
n_rows_rg = resumo_genero.shape[0]
chart2.add_series({
    'name': 'Resumo por Gênero',
    'categories': ['Resumo_Genero', 1, 0, n_rows_rg, 0],
    'values': ['Resumo_Genero', 1, 1, n_rows_rg, 1],
})
chart2.set_title({'name': 'Percentual por Gênero'})
worksheet_rg.insert_chart('E2', chart2, {'x_offset': 25, 'y_offset': 10})

# --- Sheet: Top5_Marca
df_top5_marca.to_excel(writer, sheet_name='Top5_Marca', index=False)
worksheet_tm = writer.sheets['Top5_Marca']

# Add chart for the first month's Top5 by Brand
if not df_top5_marca.empty:
    mes_inicial = df_top5_marca['Mes'].iloc[0]
    df_primeiro_mes = df_top5_marca[df_top5_marca['Mes'] == mes_inicial]
    n_rows_tm = df_primeiro_mes.shape[0]

    if n_rows_tm > 0:
        chart3 = workbook.add_chart({'type': 'column'})
        chart3.add_series({
            'name': f'Top5 por Marca - {mes_inicial}',
            'categories': ['Top5_Marca', 1, 2, n_rows_tm, 2],
            'values': ['Top5_Marca', 1, 3, n_rows_tm, 3],
        })
        chart3.set_title({'name': f'Top5 Numerações por Marca - {mes_inicial}'})
        chart3.set_x_axis({'name': 'Numeração'})
        chart3.set_y_axis({'name': 'Quantidade Vendida'})
        worksheet_tm.insert_chart('H2', chart3, {'x_offset': 25, 'y_offset': 10})

# --- Sheet: Top5_Mes
df_top5_mes.to_excel(writer, sheet_name='Top5_Mes', index=False)
worksheet_tmo = writer.sheets['Top5_Mes']

# Add chart for Top5 numerations by month
n_rows_tmo = df_top5_mes.shape[0]
chart4 = workbook.add_chart({'type': 'column'})
chart4.add_series({
    'name': 'Top5 por Mês',
    'categories': ['Top5_Mes', 1, 1, n_rows_tmo, 1],
    'values': ['Top5_Mes', 1, 2, n_rows_tmo, 2],
})
chart4.set_title({'name': 'Top5 Numerações por Mês'})
chart4.set_x_axis({'name': 'Numeração'})
chart4.set_y_axis({'name': 'Quantidade Vendida'})
worksheet_tmo.insert_chart('E2', chart4, {'x_offset': 25, 'y_offset': 10})

# --- Sheet: Comparacao_Feminino
comparacao_feminino.to_excel(writer, sheet_name='Comparacao_Feminino', index=False)

# --- Sheet: Comparacao_Masculino
comparacao_masculino.to_excel(writer, sheet_name='Comparacao_Masculino', index=False)

# Finalize Excel file
writer.close()

print(f"✅ Arquivo '{output_file}' gerado com sucesso!")

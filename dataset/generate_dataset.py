# script per scaricare da Internet alcuni indici di tutti i paesi del mondo dal 2004 al 2024
import wbgapi as wb  # libreria ufficiale della banca mondiale per le API
import pandas as pd  # libreria per manipolazione dati
import openpyxl      # libreria per gestire file excel e formattazione
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

print("avvio...")

# Dizionario con gli indicatori ordinati esattamente per rispecchiare le colonne desiderate
indicatori = {
    'NY.GDP.PCAP.KD': 'GDP Per Capita',
    'IT.NET.USER.ZS': 'Internet Users (%)',
    'IT.CEL.SETS.P2': 'Mobile Subscriptions (per 100 people)',
    'IT.NET.BBND.P2': 'Fixed Broadband Subscriptions (per 100 people)',
    'SP.URB.TOTL.IN.ZS': 'Urban Population (%)',
    'EG.ELC.ACCS.ZS': 'Electricity Access (%)'
}

intervallo_anni = range(2004, 2025)

# Scarico i dati grezzi (skipAggs evita aggregazioni globali o continentali)
print("scaricamento dati dalla world bank in corso...")
df_raw = wb.data.DataFrame(list(indicatori.keys()), time=intervallo_anni, skipAggs=True).reset_index()

# Sostituisce i codici delle serie con i nomi leggibili
df_raw['series'] = df_raw['series'].map(indicatori)

# Appiattisce la tabella (melt) e pulisce gli anni (rimuove 'YR')
df_melted = df_raw.melt(id_vars=['economy', 'series'], var_name='Year', value_name='Value')
df_melted['Year_Clean'] = df_melted['Year'].str.replace('YR', '', regex=False).astype(int)

# Pivot per ottenere la struttura tabellare: righe = (Paese, Anno), colonne = Indicatori
df_pivot = df_melted.pivot_table(
    index=['economy', 'Year_Clean'],
    columns='series',
    values='Value',
    dropna=False
).reset_index()

df_pivot.columns.name = None
df_pivot = df_pivot.rename(columns={'economy': 'Paese', 'Year_Clean': 'Anno'})

# Riordina le colonne esattamente nell'ordine stabilito
cols_order = ['Paese', 'Anno'] + list(indicatori.values())
df_pivot = df_pivot[cols_order]

# Ordinamento pulito per Paese e Anno
df_pivot = df_pivot.sort_values(['Paese', 'Anno']).reset_index(drop=True)

print("creazione e formattazione del file excel in corso...")
wb_out = openpyxl.Workbook()
ws = wb_out.active
ws.title = "Dati_Strutturati"

# Stili grafici
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
country_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
data_font = Font(name="Calibri", size=10, color="000000")
group_alt_fill = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

thin_side = Side(style='thin', color='E0E0E0')
border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# Scrittura Intestazione (Riga 1)
for col_idx, col_name in enumerate(cols_order, start=1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

ws.row_dimensions[1].height = 30
ws.freeze_panes = "B2"  # Blocca la colonna A (Paese) e la riga 1 (Intestazione)

# Scrittura righe 
prev_country = None
country_idx = 0

for row_idx, row in df_pivot.iterrows():
    excel_row = row_idx + 2  # Riga 1 è l'intestazione
    curr_country = row['Paese']
    is_first_row = (curr_country != prev_country)
    
    if is_first_row:
        country_idx += 1
        prev_country = curr_country
        
    row_fill = group_alt_fill if (country_idx % 2 == 0) else white_fill
    
    # 1. Colonna Paese (scritto solo se è la prima riga del paese)
    cell_paese = ws.cell(row=excel_row, column=1, value=curr_country if is_first_row else "")
    cell_paese.font = country_font
    cell_paese.fill = row_fill
    cell_paese.border = border
    cell_paese.alignment = Alignment(horizontal="left", vertical="center")
    
    # 2. Colonna Anno
    cell_anno = ws.cell(row=excel_row, column=2, value=int(row['Anno']))
    cell_anno.font = data_font
    cell_anno.fill = row_fill
    cell_anno.border = border
    cell_anno.alignment = Alignment(horizontal="center", vertical="center")
    cell_anno.number_format = '0'
    
    # 3. Colonne Indicatori
    for col_i, ind_name in enumerate(indicatori.values(), start=3):
        val = row[ind_name]
        cell_val = ws.cell(row=excel_row, column=col_i)
        if pd.notna(val):
            cell_val.value = round(float(val), 2)
            cell_val.number_format = '#,##0.00'
        cell_val.font = data_font
        cell_val.fill = row_fill
        cell_val.border = border
        cell_val.alignment = Alignment(horizontal="right", vertical="center")

# Larghezza colonne ottimizzata
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 10
for col_i in range(3, len(indicatori) + 3):
    ws.column_dimensions[get_column_letter(col_i)].width = 22

wb_out.save("dataset_r.xlsx")
print("file creato con successo")
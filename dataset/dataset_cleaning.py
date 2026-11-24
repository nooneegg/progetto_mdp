#libreria con centinaia di funzioni per gestire tabelle, file excel ecc...
import pandas as pd
#oggetto matematico che si occupa di fare la normalizzazione usando lo z-score scaling, come funziona:
#dato un valore x fa -> (x-media)/sigma -> in questo modo tutti i valori assumono una media pari a 0,
#i valori si aggireranno quindi tutti intorno allo 0 e bla bla bla ok ci siamo capiti
from sklearn.preprocessing import StandardScaler
#librerie per gestire grafica in excel
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

print('avvio...')

#grazie alla libreria pandas creiamo un dataframe, una tabella dati indentica al foglio excel che carichiamo 
df = pd.read_excel('dataset_r.xlsx')

#per questioni di leggibilità il nome del paese è scritto solo nel primo anno,
#per usare il groupby dobbiamo far sapere a python che anche gli anni successivi fino al 2024 fanno parte di quel determinato paese,
#.ffill() guarda sopra e copia lo stesso valore finchè non trova una cella già riempita (paese dopo)
df['Paese'] = df['Paese'].ffill()

#definiamo lista indicatori
indicatori = [
    'GDP Per Capita', 
    'Internet Users (%)', 
    'Mobile Subscriptions (per 100 people)',
    'Fixed Broadband Subscriptions (per 100 people)', 
    'Urban Population (%)', 
    'Electricity Access (%)'
]
#salviamo una copia del dataframe originale prima dell'interpolzione
#questo ci serve per scoprire quali celle erano vuote prima dell'esecuzione di questo script
df_originale_na = df[indicatori].isna()

#questo comando riempie i buchi di dati
#df.groupby('Paese') divide i dati per paese 
#[indicatori] divide i dati per ogni indicatore, creando quindi istanze (paese, indicatore)
#.trasform gestisce i dati in verticale applicando le seguenti funzioni:
#.interpolate(method='linear') riempie i buchi in mezzo creando una linea retta tra i valori e prendendo la media, es: 30, NaN, 50 -> 30, 40, 50
#limit=3 riempie i dati solo se il buco è minore di 4 anni, altrimenti lo lascia vuoto per evitare over imputation
#limit_direction='both' permette di riempire buchi in ambo le direzioni, senza il both l'interpolazione riempirebbe solo in avanti
#i buchi interni (valore, NaN, valore) vengono riempiti con il metodo lineare classico detto prima
#i buchi esterni (valore, ..., valore, NaN) vengono riempiti incollando l'ultimo/il primo valore valido
#questo principalmente perchè molti indicatori (5 su 6) presentano dei tetti di valori ammissibili: una percentuale non può avere come ultimo valore 102.4
#quindi al posto di fare una stima dei valori ricopio e basta
print("interpolazione lineare in corso...")
df[indicatori] = df.groupby('Paese')[indicatori].transform(
    lambda x: x.interpolate(method='linear', limit=3, limit_direction='both')
)

#elimino una riga solo se nessun indicatore presenta un dato
righe_valide = df[indicatori].notna().any(axis=1)

#se invece una riga non ha nessun daato per nessun indice viene elimiata
df = df[righe_valide].reset_index(drop=True)
#viene eliminata anche nel dataframe iniziale (quello che serve per comparare prima vs dopo e colorare le celle giuste)
df_originale_na = df_originale_na[righe_valide].reset_index(drop=True)

#calcolo media e sigma SOLO dalla prima riga fino all'indice 3632 (che corrisponde alla riga excel 3634) e normalizzo tutto, questo perchè
#media e sigma devono essere calcolati solo sui paesi del training perchè senno influenzerebbero il testing
#se dovessi fare media e sigma su tutto, quando la rete neurale stima la % internet users di un paese
#usa media e sigma di tutti i paesi, al posto di solo quelli del training, in questo modo il testing verrebbe "falsificato"
print("normalizzazione in corso...")
df_scaled = df.copy()
scaler = StandardScaler()
scaler.fit(df.loc[:3632, indicatori])
df_scaled[indicatori] = scaler.transform(df[indicatori])

wb_out = openpyxl.Workbook()
ws = wb_out.active
ws.title = "Dati_Strutturati"

#formattazione (gemini)
print("formattazione file excel in corso...")
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
country_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
data_font = Font(name="Calibri", size=10, color="000000")
group_alt_fill = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
fill_giallo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
thin_side = Side(style='thin', color='E0E0E0')
border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
cols_order = ['Paese', 'Anno'] + indicatori
for col_idx, col_name in enumerate(cols_order, start=1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border
ws.row_dimensions[1].height = 30
ws.freeze_panes = "B2" 
prev_country = None
country_idx = 0
for row_idx, row in df_scaled.iterrows():
    excel_row = row_idx + 2 
    curr_country = row['Paese']
    is_first_row = (curr_country != prev_country)
    if is_first_row:
        country_idx += 1
        prev_country = curr_country
    base_fill = group_alt_fill if (country_idx % 2 == 0) else white_fill
    cell_paese = ws.cell(row=excel_row, column=1, value=curr_country if is_first_row else "")
    cell_paese.font = country_font
    cell_paese.fill = base_fill
    cell_paese.border = border
    cell_paese.alignment = Alignment(horizontal="left", vertical="center")
    cell_anno = ws.cell(row=excel_row, column=2, value=int(row['Anno']))
    cell_anno.font = data_font
    cell_anno.fill = base_fill
    cell_anno.border = border
    cell_anno.alignment = Alignment(horizontal="center", vertical="center")
    cell_anno.number_format = '0'
    for col_i, ind_name in enumerate(indicatori, start=3):
        val = row[ind_name]
        cell_val = ws.cell(row=excel_row, column=col_i)
        if pd.notna(val):
            cell_val.value = round(float(val), 6)  # Precisione decimale ideale per lo z-score
            cell_val.number_format = '0.000000'
        cell_val.font = data_font
        cell_val.border = border
        cell_val.alignment = Alignment(horizontal="right", vertical="center")
#verifica: se la cella originale era un NaN ed è stata riempita, coloriamo di giallo
        was_imputed = df_originale_na.loc[row_idx, ind_name]
        if was_imputed:
            cell_val.fill = fill_giallo
        else:
            cell_val.fill = base_fill
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 10
for col_i in range(3, len(indicatori) + 3):
    ws.column_dimensions[get_column_letter(col_i)].width = 22

#salvataggio del file
wb_out.save("dataset_c.xlsx")
print("file creato con successo")

#stampo a terminale media e deviazione standard di ogni indice
print("="*50)
print("PARAMETRI DI NORMALIZZAZIONE:")
print("="*50)
for i, ind in enumerate(indicatori):
    print(f"- {ind:<45} | Media: {scaler.mean_[i]:12.6f} | Dev. Std: {scaler.scale_[i]:12.6f}")
print("\n")

#nonostante la pulizia del dataset rimangono molti buchi, alcune celle non possono essere riempite per mancanza di dati sufficienti a fare stime
#altri sono più complessi da gestire, come congo o guinea dove ci sono 6 anni vuoti e poi il primo valore è 0.0, 0.1, 0.2 e sempre crescendo
#questo fa capire che i 10 anni prima non sono da eliminare, sono semplicemente tutti 0, anche se dai dati risultato vuoti
#essendo più di 4500 righe per analizzare e valutare quali buchi potrebbero essere riempiti con una semplice analisi logica e razionale ci siamo affidati a gemini
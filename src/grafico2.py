from pathlib import Path

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# cartella in cui si trova questo script (src)
cartella_src = Path(__file__).resolve().parent

# cartella principale del progetto
cartella_progetto = cartella_src.parent

# percorso del dataset nella cartella data
file_excel = cartella_progetto / "data" / "dataset_r.xlsx"
# si usa la libreria pandas per aprire il file e dargli un nome
df = pd.read_excel(
    file_excel, 
    sheet_name="Dati_Strutturati")

# i dati vengono sistemati 
df = df.dropna(subset=["Anno"]).copy()   # elimina righe che non hanno un anno
df["Paese"] = df["Paese"].ffill()       # ffill riempie le celle vuote della colonna paese 
df["Anno"] = df["Anno"].astype(int)     # controlla che l'anno sia un numero intero (per evitare errori)

# si calcola per ogni anno la mediana 
colonna = "Internet Users (%)"
mediana = df.groupby(       #groupby raggruppa per anno
    "Anno", as_index=False
    )[colonna].median()       #poi .median() calcola il valore centrale

#si prende l'italia per confrontarla alla mediana mondiale
italia = df[df["Paese"] == "Italy"]


# paesi scelti per il confronto, sono stati scelti casi "significativi" ovvero alcuni paesi poveri e altri ricchi
paesi = ["Germany", "United States", "Brazil", "India", "Nigeria", "Italy"]
# si cambia il nome in inglese (proveniente dal dataset) in italiano
nomi = {"Germany": "Germania", "United States": "Stati Uniti",  
        "Brazil": "Brasile", "India": "India", "Nigeria": "Nigeria",
        "Italy": "Italia"}
# si assegnano colori ai vari paesi in modo da distinguerli sul grafico
colori = {"Germany": "#4179A5", "United States": "#8465A8",
          "Brazil": "#4F9562", "India": "#D19A45", "Nigeria": "#7D8895",
          "Italy": "#D62728"}

# si disegnano i 2 grafici
sns.set_theme(style="whitegrid")
fig, assi = plt.subplots(
    1, 2, figsize=(15, 6), sharey=True #sharey=True da la stessa scala verticale ad entrambi i grafici per semplificare i confronti
)

# questa è la funzione che disegna il grafico usando i valori delle colonne anno e la % internet user come riga (mondo)
sns.lineplot(data=mediana, x="Anno", y=colonna, ax=assi[0],
             label="Mediana dei Paesi", color="#4179A5", linewidth=2.5)
# si ripete per l'italia
sns.lineplot(data=italia, x="Anno", y=colonna, ax=assi[0],
             label="Italia", color=colori["Italy"], linewidth=3.2)
#titolo
assi[0].set_title("Panoramica mondiale: mediana e Italia")


# secondo grafico => si cicla su tutti i paesi all interno del dizionario paesi e si fa un lineplot
for paese in paesi:
    dati_paese = df[df["Paese"] == paese]
    sns.lineplot(data=dati_paese, x="Anno", y=colonna, ax=assi[1],
                 label=nomi[paese], color=colori[paese],
                 linewidth=3.2 if paese == "Italy" else 1.9) #se il paese è l'italia la linea è più spessa per metterla in risalto
# titolo 2
assi[1].set_title("Confronto tra Paesi")

# si aggiungono le legende al frafico
for ax in assi:
    ax.set_xlim(2004, 2024)
    ax.set_ylim(0, 100)
    ax.set_xticks(range(2004, 2025, 4))
    ax.set_xlabel("Anno")
    ax.set_ylabel("Popolazione che utilizza Internet (%)")
    ax.legend(loc="upper left", frameon=True, fontsize=9)

# varie label
fig.suptitle("Diffusione di Internet nel mondo (2004-2024)", fontsize=16, fontweight="bold")
fig.text(0.5, 0.01,
         "Fonte: dataset del gruppo, indicatori World Bank/ITU. Mediana non ponderata dei Paesi con dati disponibili.",
         ha="center", fontsize=9, color="#555555")
fig.tight_layout(rect=(0, 0.04, 1, 0.94))

#si salva il grafico
fig.savefig(cartella_src / "grafico_02.png", dpi=500, bbox_inches="tight")
plt.show()

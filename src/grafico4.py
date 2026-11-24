from pathlib import Path

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# si apre il dataset originale, nella stessa cartella dello script
# cartella in cui si trova lo script (src)
cartella = Path(__file__).resolve().parent

# dataset nella cartella dataset, allo stesso livello di src
file_excel = cartella.parent / "dataset" / "dataset_r.xlsx"

# lettura del file Excel
df = pd.read_excel(file_excel, sheet_name="Dati_Strutturati")

# si sistemano il nome del paese e il tipo dell'anno, come negli altri grafici
df = df.dropna(subset=["Anno"]).copy()
df["Paese"] = df["Paese"].ffill()
df["Anno"] = df["Anno"].astype(int)

# si scelgono le 6 variabili numeriche che vogliamo confrontare
variabili = {
    "Internet Users (%)": "Internet (%)",
    "GDP Per Capita": "PIL pro capite",
    "Mobile Subscriptions (per 100 people)": "Abbonamenti mobili",
    "Fixed Broadband Subscriptions (per 100 people)": "Banda larga fissa",
    "Urban Population (%)": "Popolazione urbana (%)",
    "Electricity Access (%)": "Accesso elettricita' (%)",
}

# si seleziona solo il 2024, per confrontare i paesi nello stesso anno
# dropna elimina le righe con almeno un valore mancante tra le 6 variabili
anno = 2024
dati = df[df["Anno"] == anno].copy()
dati = dati.dropna(subset=list(variabili))

# si rinominano solo le colonne usate nel grafico, senza cambiare l'Excel
dati = dati[list(variabili)].rename(columns=variabili)

# corr() calcola la correlazione di Pearson tra tutte le coppie di variabili
correlazioni = dati.corr(method="pearson")
print(f"Anno: {anno} | Paesi con dati completi: {len(dati)}")
print(correlazioni.round(2))

# si costruisce la heatmap: il colore indica il segno e l'intensità
# della correlazione
# si mette il numero nella casella per indicare il valore esatto
sns.set_theme(style="white", font_scale=1)
fig, ax = plt.subplots(figsize=(11.5, 8.7))

sns.heatmap(
    correlazioni,
    ax=ax,
    annot=True,             # mostra il coefficiente dentro ogni casella
    fmt=".2f",              # arrotonda i numeri a due cifre decimali
    cmap="RdBu_r",          # blu: negativo, rosso: positivo
    vmin=-1, vmax=1,        # la scala dei colori va sempre da -1 a +1
    center=0,               # lo zero e' il centro della scala
    square=True,            # caselle quadrate
    linewidths=0.8,        # linee di separazione tra caselle
    linecolor="white",
    annot_kws={"size": 11},
    cbar_kws={"label": "Correlazione di Pearson (r)", "shrink": 0.85},
)

# titoli, etichette e fonte
ax.set_title(
    f"Correlazioni tra le variabili del divario digitale - {anno}",
    fontsize=15, fontweight="bold", pad=20
)
ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
ax.set_yticklabels(ax.get_yticklabels(), rotation=0)
ax.tick_params(axis="both", length=0)

fig.text(
    0.5, 0.025,
    f"Fonte: dataset del gruppo (World Bank/ITU). "
    f"Correlazione di Pearson su {len(dati)} Paesi e territori con dati completi.",
    ha="center", fontsize=9, color="#555555"
)
fig.tight_layout(rect=(0, 0.055, 1, 1))

# si esporta la figura come immagine ad alta risoluzione per PowerPoint
fig.savefig(cartella / "grafico_4.png", dpi=500, bbox_inches="tight")
plt.show()

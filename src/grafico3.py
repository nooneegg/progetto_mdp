from pathlib import Path

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# 1. Si legge il dataset
# cartella in cui si trova lo script (src)
cartella = Path(__file__).resolve().parent

# percorso del dataset nella cartella dataset
file_excel = cartella.parent / "dataset" / "dataset_r.xlsx"

# lettura del file Excel
df = pd.read_excel(file_excel, sheet_name="Dati_Strutturati")

# 2. i dati vengono riempiti in maniera simile a quelli del primo grafico 
df["Paese"] = df["Paese"].ffill() #riempie le celle vuote
df = df.dropna(subset=["Anno"]).copy() #elimina le righe che non hanno un anno
df["Anno"] = df["Anno"].astype(int) #controlla che l'anno sia un intero


x = "GDP Per Capita"
y = "Internet Users (%)"

# i dati vengono selezionati solo nel 2024
dati = df[df["Anno"] == 2024].copy()
# Per lo scatter plot servono entrambe le variabili, si usa di nuovo dropna per cancellare i dati che non hanno pil o %internet
dati = dati.dropna(subset=[x, y])
# Usiamo una scala logaritmica per il PIL: non ammette zero o valori negativi.
dati = dati[dati[x] > 0]

# 3. paesi scelti per fare un confronto di alcuni specifici, italia in rosso
paesi_confronto = {
    "Italy": "Italia",
    "Germany": "Germania",
    "United States": "Stati Uniti",
    "Brazil": "Brasile",
    "India": "India",
    "Nigeria": "Nigeria",
}
altri = dati[~dati["Paese"].isin(paesi_confronto)] #tutti gli altri paesi
selezionati = dati[dati["Paese"].isin(paesi_confronto)]
italia = dati[dati["Paese"] == "Italy"]
confronto_senza_italia = selezionati[selezionati["Paese"] != "Italy"]

# 4. Si crea il grafico: tutti i Paesi, con alcuni casi riconoscibili.
sns.set_theme(style="whitegrid")
fig, ax = plt.subplots(figsize=(12, 7))

sns.scatterplot( #creazione scatterplot con paesi generici
    data=altri, x=x, y=y, ax=ax, #si associa x e y a gdp per capita e %internet
    color="#427DA8", s=58, alpha=0.58,
    edgecolor="white", linewidth=0.4,
    label="Altri Paesi e territori",
)

# si crea la linea di tendenza con i dati 
sns.regplot(
    data=dati, x=x, y=y, ax=ax,
    logx=True, scatter=False, ci=None,
    line_kws={"color": "#D49A38", "linewidth": 2.2,
               "linestyle": "--", "label": "Tendenza complessiva"},
)

# creazione scatterplot con paesi selezionati
sns.scatterplot(
    data=confronto_senza_italia, x=x, y=y, ax=ax,
    color="#2E5266", s=88, edgecolor="white", linewidth=1,
    label="Paesi selezionati", zorder=4,
)
#  creazione scatterplot con italia
sns.scatterplot(
    data=italia, x=x, y=y, ax=ax,
    color="#D62728", s=185, marker="*", edgecolor="white",
    linewidth=0.9, label="Italia", zorder=6,
)

# Etichette dei sei Paesi selezionati per facilitarne il confronto
posizioni = {
    "Italy": (9, -17),
    "Germany": (8, 10),
    "United States": (8, -14),
    "Brazil": (8, 10),
    "India": (9, 8),
    "Nigeria": (9, -14),
}
# si impostano le etichette ciclando sui paesi
for _, riga in selezionati.iterrows():
    paese = riga["Paese"]
    colore = "#D62728" if paese == "Italy" else "#263B4A"
    ax.annotate(
        paesi_confronto[paese],
        (riga[x], riga[y]),
        xytext=posizioni[paese], textcoords="offset points",
        fontsize=10, color=colore,
        fontweight="bold" if paese == "Italy" else "normal",
    )

# 5. Assi, titolo, fonte ed esportazione.
ax.set_xscale("log")  # Mostra bene sia PIL piccoli sia PIL molto grandi
ax.set_xlim(200, 350000)
ax.set_ylim(0, 103)
ax.set_xlabel("PIL pro capite (scala logaritmica)", fontsize=12)
ax.set_ylabel("Popolazione che utilizza Internet (%)", fontsize=12)
ax.set_title("PIL pro capite e utilizzo di Internet nel 2024", fontsize=16, weight="bold", pad=16)
ax.plot([], [], color="#D49A38", linewidth=2.2, linestyle="--", label="Tendenza complessiva")
ax.legend(loc="lower right", frameon=True)
ax.grid(True, alpha=0.28)
fig.text(
    0.5, 0.015,
    f"Fonte: dataset originale del gruppo (World Bank/ITU). "
    f"{len(dati)} Paesi/territori con entrambi i dati disponibili. "
    "La linea indica associazione, non causalità.",
    ha="center", fontsize=9, color="#52606D",
)
fig.tight_layout(rect=(0, 0.045, 1, 1))
fig.savefig(cartella / "grafico_3.png", dpi=300, bbox_inches="tight")
plt.show()

print(f"Osservazioni nel grafico (2024): {len(dati)}")
if not italia.empty:
    print(f"Italia: PIL pro capite = {italia.iloc[0][x]:.2f}; "
          f"Internet = {italia.iloc[0][y]:.2f}%")

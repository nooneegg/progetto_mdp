from pathlib import Path

import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.patheffects as effetti
import pandas as pd
import pycountry
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable
from matplotlib.patches import Patch

# cartella in cui si trova lo script (src)
cartella = Path(__file__).resolve().parent

# dataset Excel nella cartella dataset
file_excel = cartella.parent / "dataset" / "dataset_r.xlsx"

# mappa GeoJSON nella cartella dataset
file_mappa = cartella.parent / "dataset" / "ne_110m_admin_0_countries.geojson"

if not file_excel.exists():
    raise FileNotFoundError(f"Dataset mancante: {file_excel}")
if not file_mappa.exists():
    raise FileNotFoundError(
        f"Mappa mancante: {file_mappa}\n"
        "Scarica il file GeoJSON fornito insieme allo script."
    )

# lettura dati non normalizzati 
anno = 2024
variabile = "Internet Users (%)"
df = pd.read_excel(file_excel, sheet_name="Dati_Strutturati")
df = df.dropna(subset=["Anno"]).copy()
df["Paese"] = df["Paese"].ffill()
df["Anno"] = df["Anno"].astype(int)
dati = df.loc[df["Anno"] == anno, ["Paese", variabile]].copy()

# Converte i nomi della Banca Mondiale in codici paese ISO a 3 lettere.
# Alcune denominazioni World Bank non sono riconosciute automaticamente.
nomi_speciali = {
    "Bahamas, The": "BHS",
    "Congo, Dem. Rep.": "COD",
    "Congo, Rep.": "COG",
    "Cote d'Ivoire": "CIV",
    "Egypt, Arab Rep.": "EGY",
    "Gambia, The": "GMB",
    "Hong Kong SAR, China": "HKG",
    "Iran, Islamic Rep.": "IRN",
    "Korea, Rep.": "KOR",
    "Korea, Dem. People's Rep.": "PRK",
    "Lao PDR": "LAO",
    "Macao SAR, China": "MAC",
    "Micronesia, Fed. Sts.": "FSM",
    "Russian Federation": "RUS",
    "St. Kitts and Nevis": "KNA",
    "St. Lucia": "LCA",
    "St. Vincent and the Grenadines": "VCT",
    "Turkiye": "TUR",
    "Venezuela, RB": "VEN",
    "West Bank and Gaza": "PSE",
    "Kyrgyz Republic": "KGZ",
    "Slovak Republic": "SVK",
    "Eswatini": "SWZ",
    "Kosovo": "XKX",
}

# funzione che scambia i nomi dei paesi con il dizionario sopra se il nome originale non è presente nella libreria per il plot del grafico
def codice_iso(paese):
    if paese in nomi_speciali:
        return nomi_speciali[paese]
    try:
        return pycountry.countries.lookup(paese).alpha_3
    except LookupError:
        return None


dati["ISO_A3"] = dati["Paese"].apply(codice_iso)

# Lettura dei confini dalla mappa geografica, è come una tabella che contiene una colonna con le coordinate geografiche dei paesi
mondo = gpd.read_file(file_mappa)

# correzione alcuni paesi
mondo.loc[mondo["NAME"] == "France", "ISO_A3"] = "FRA"
mondo.loc[mondo["NAME"] == "Norway", "ISO_A3"] = "NOR"
mondo.loc[mondo["NAME"] == "Kosovo", "ISO_A3"] = "XKX"

# si uniscono con un left join le tabelle del nostro dataset con % internet e la tabella mondo con le coordinate dei paesi
mondo = mondo.merge(dati[["ISO_A3", variabile]], on="ISO_A3", how="left")

#colore = percentuale, grigio = dato non disponibile, si da il colore alla mappa
plt.rcParams["font.family"] = "DejaVu Sans"
fig = plt.figure(figsize=(16, 9), facecolor="white")
ax = fig.add_axes([0.035, 0.22, 0.93, 0.62])
ax.set_facecolor("#F3F7FA")

scala = Normalize(vmin=0, vmax=100)
colormap = plt.get_cmap("YlGnBu")
#plot del grafico
mondo.plot(
    column=variabile,
    ax=ax, #indica l'aera in cui si disegna il grafico
    cmap=colormap, # usa la palette di colosi scelta
    norm=scala, 
    linewidth=0.45, #imposta lo spessore dei confini
    edgecolor="white",
    missing_kwds={"color": "#C9D1D8", "edgecolor": "white"}, #gestione dati mancanti con colore grigio
)


# Ritaglia l'Antartide per rendere la mappa leggibile nelle slide.
ax.set_xlim(-169, 194)
ax.set_ylim(-60, 86)
ax.set_aspect("auto")
ax.set_axis_off()

# titolo, sottotitolo e barra cromatica 
fig.text(0.055, 0.95, "DIVARIO DIGITALE NEL MONDO", fontsize=22,
         fontweight="bold", color="#152F41", ha="left", va="top")
fig.text(0.055, 0.907, "Quota di popolazione che utilizza Internet, per Paese",
         fontsize=13.5, color="#486376", ha="left", va="top")
fig.text(0.935, 0.940, str(anno), fontsize=25,
         fontweight="bold", color="#152F41", ha="right", va="top")

barra_ax = fig.add_axes([0.22, 0.155, 0.58, 0.025])
mappa_colori = ScalarMappable(norm=scala, cmap=colormap)
mappa_colori.set_array([])
barra = fig.colorbar(mappa_colori, cax=barra_ax, orientation="horizontal")
barra.set_ticks([0, 20, 40, 60, 80, 100])
barra.ax.tick_params(labelsize=11, length=0, pad=7)
barra.outline.set_visible(False)
fig.text(0.22, 0.191, "Popolazione che utilizza Internet (%)", color="#334C5B",
         fontsize=11, fontweight="normal", ha="left")

# legenda per i paesi senza dati
fig.legend(
    handles=[
        Patch(facecolor="#C9D1D8", edgecolor="none", label="Dato assente"),
    ],
    loc="lower right", bbox_to_anchor=(0.96, 0.12),
    frameon=False, ncol=1, fontsize=10,
)

n_valori = int(dati[variabile].notna().sum())
n_mappa = int(mondo[variabile].notna().sum())
fig.text(
    0.055, 0.075,
    f"Fonte: dataset del gruppo (World Bank/ITU); confini: Natural Earth 1:110m. "
    f"Valori presenti: {n_valori}/217; territori visibili e associati: {n_mappa}.",
    fontsize=9.5, color="#526575", ha="left"
)
fig.text(
    0.055, 0.046,
    "I microstati possono non comparire alla scala della mappa. "
    "Un Paese senza colore non equivale a uno con utilizzo Internet pari a zero.",
    fontsize=9.1, color="#62727F", ha="left"
)

#crea un PNG ad alta risoluzione 
output = cartella / "grafico_1.png"
fig.savefig(output, dpi=300, facecolor="white", bbox_inches="tight")
print(f"\nImmagine salvata: {output}")
print(f"Paesi con dato valido nel file: {n_valori}")
print(f"Paesi/territori con dato visualizzati nella mappa: {n_mappa}")
non_associati = dati.loc[dati[variabile].notna() & ~dati["ISO_A3"].isin(mondo["ISO_A3"]), "Paese"].tolist()
print(f"Non presenti sulla base cartografica semplificata: {non_associati}\n")
plt.show()

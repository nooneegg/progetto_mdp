#questo script serve addestra una rete neurale artificiale per stimare la % di Internet Users a partire da 5 indici dei paesi e gli anni  
from pathlib import Path
import pandas as pd
#sklearn è la libreria per l'apprendimento automatico, MLPRegressor è il motore della rete
#MLP significa multy layer perceptron, è un tipo di rete neurale composta da strati di neuroni
#strato di input, riceve i dati iniziali
#strato nascosto, ogni neurone prende tutti gli input e fa dei calcoli matematici su di essi e tira fuori un risultato
#strato di output, è un singolo output con il risulato in % di Internet Users stimata per il paese x nell'anno y
#Regressor indica invece che il modello usa la regressione, cioe prevede un numero continuo e non una categoria come "mammifero" o "jazz" 
from sklearn.neural_network import MLPRegressor
#queste sono metriche che fanno capire se la rete ha lavorato bene o no
#MAE --> mean absolute error, calcola di quanto si sbaglia la rete mediamente rispetto al valore reale (basso = meglio)
#R2  --> guarda quanto i dati si distaccano dalla media globale e gli da un voto (alto = meglio)
# - R2 > 0 --> i dati si distaccano dalla media globale e si avvicinano al target
# - R2 = 0 --> i dati sono uguali alla media globale (es: se la media di Internet Users è 5 allora ogni valore di Internet Users stimato per ogni paese e per ogni anno sarà uguale a 5)
# - R2 < 0 --> i dati si distaccano dalla media globale e si allontanano dal target
from sklearn.metrics import mean_absolute_error, r2_score

#creo dataframe caricandoci tutto il dataset 
# cartella in cui si trova lo script (src)
cartella = Path(__file__).resolve().parent

# percorso del dataset nella cartella dataset
file_excel = cartella.parent / "dataset" / "dataset_c.xlsx"

# lettura del file Excel
df = pd.read_excel(file_excel, sheet_name="Dati_Strutturati", engine="openpyxl")

#riempio le celle vuote dei paesi, cosi script capisce subito a che paese corrisponde una certa riga altrimenti avremmo dovuto fare più controlli
df["Paese"] = df["Paese"].ffill().str.strip()
df["Anno"] = df["Anno"].astype(int) #rendo anno un intero per poterlo usare dopo come indice

#seleziono feature (input) e target (output), paese NON viene dato alla rete come input. Anno si perchè abbiamo visto che migliora la precisione del modello
feature = [
    "GDP Per Capita",
    "Mobile Subscriptions (per 100 people)",
    "Fixed Broadband Subscriptions (per 100 people)",
    "Urban Population (%)",
    "Electricity Access (%)",
    "Anno",
]
target = "Internet Users (%)"

#forza tutte le colonne ad essere numeriche
colonne_numeriche = feature + [target]
df[colonne_numeriche] = df[colonne_numeriche].apply(pd.to_numeric, errors="raise")

#training termina con l'ultima riga dello Zimbabwe.
#testing inizia con la prima riga di Aruba.
#una volta diviso il dataset abbiamo spostato tutti i paesi di testing in fondo subito dopo quelli di training
#per quello nel nuovo dataset abbiamo Zimbabwe e dopo Aruba e altri paesi
righe_zimbabwe = df.index[df["Paese"] == "Zimbabwe"]
righe_aruba = df.index[df["Paese"] == "Aruba"]
fine_training = righe_zimbabwe[-1]
inizio_testing = righe_aruba[0]
training = df.iloc[:fine_training + 1].copy()
testing = df.iloc[inizio_testing:].copy()
n_training = len(training)
n_testing = len(testing)

#scarto le righe che hanno anche solo un valore mancante, questo perchè MLPRegressor non sa apprendere da una riga con anche solo un dato mancante e andrebbe in errore
#(non si riempiono altri buchi e non si applica una nuova normalizzazione)
training = training.dropna(subset=colonne_numeriche).copy()
testing = testing.dropna(subset=colonne_numeriche).copy()

#abbiamo trasformato tutto in z-score cosicche la rete potesse avere tutti i dati sulla stessa scala in modo da non compromettere il training
#poichè la rete usa un singolo linguaggio per parlare (che ora è quello dello z-score) se vogliamo passargli pure l'anno per addestrarla dobbiamo normalizzarlo anch'esso con z-score
training["Anno_scalato"] = (training["Anno"] - training["Anno"].mean()) / training["Anno"].std()
testing["Anno_scalato"] = (testing["Anno"] - training["Anno"].mean()) / training["Anno"].std()

#stampo le righe di training/testing totali e quelle che ho escluso per mancanza di dati 
print(f"Training: {len(training)} righe; escluse {n_training - len(training)} incomplete.")
print(f"Testing: {len(testing)} righe; escluse {n_testing - len(testing)} incomplete.")

colonne_rete = feature[:-1] + ["Anno_scalato"]
X_train = training[colonne_rete] #contiene i valori delle 6 colonne di input che fanno parte del training       --> serve ad allenare la rete
X_test = testing[colonne_rete]   #contiene i valori delle 6 colonne di input che fanno parte del testing        --> serve a dare input alla rete per il testing
y_train = training[target]       #contiene i valori della colonna Internet Users che fanno parte del training   --> serve per vedere quanto la previsione si distacca dalla realtà

#creazione della rete neurale coi seguenti parametri
rete = MLPRegressor(
    hidden_layer_sizes=(12,12,12),  # 3 strati nascosti con 12 neuroni ciascuno
    activation="tanh",              # funzione applicata dai neuroni nascosti
    solver="adam",                  # algoritmo che aggiorna i pesi della rete
    alpha=0.5,                      # penalizzazione dei pesi troppo grandi per ridurre overfitting (rischio che la rete impari a memoria )
    learning_rate_init=0.001,       # dimensione iniziale dei passi di apprendimento
    max_iter=20000,                 # massimo numero di passaggi sul training
    tol=0.0001,                     # miglioramento minimo richiesto
    n_iter_no_change=100,           # arresto se il training non migliora abbastanza
    early_stopping=False,           # nessuna suddivisione automatica aggiuntiva
    random_state=42,                # fissa la casualita' dell'addestramento
)

#allena la rete usando i dati del training
print("Addestramento della rete...")
rete.fit(X_train, y_train)

#faccio le previsioni usando i dati del testing
print("Previsioni sui dati di testing...")
previsioni = rete.predict(X_test)

#calcolo le metriche di valutazione (confrontando target reali e previsioni)
mae = mean_absolute_error(testing[target], previsioni)
r2 = r2_score(testing[target], previsioni)
#rappresentazione in % del mae, 31.51... è la deviazione standard del % di Internet Users, cosi si torna da z-score a normale
perc = mae*31.510324
#stampo mae e R² nel terminale
print(f"MAE: {mae:.4f} | {perc:.2f} %")
print(f"R²: {r2:.4f}")

#creo la tabella con internet target, previsto e il rispettivo errore, tutto espresso in z-score
risultati = testing[["Paese", "Anno", target]].copy().rename(columns={target: "Internet_target"})
risultati["Internet_previsto"] = previsioni
risultati["Errore"] = (risultati["Internet_previsto"] - risultati["Internet_target"]).abs()

#salvo i risultati del test in un nuovo excel
file_risultati = "risultati_test.xlsx"
risultati.to_excel(
    file_risultati,
    sheet_name="Risultati_test",
    index=False,
    engine="openpyxl",
    freeze_panes=(1, 2),
)
print(f"Risultati salvati in: {file_risultati}")
# Creazione file word revisione

# Scopo
Il codice deve creare un file word precompiato, modificando solo i dati necessari al suo completamento come elencati già nel documento.
# Funzionalità
Il codice utilizza già dei dizionari per allocare le voci importanti. Le variabili context\_generalita\_relazione e context rappresentano un contesto generale che l'utente deve compilare. Non modificare queste voci.
Vi sono altri dizionari salvati nelle variabili context\_tabella\_<nome> con <nome> = \[dpi, mansioni, heg\]. Modifica il codice nel modo seguente:
## Tabella dpi
Il codice legge il file excel indicato nella variabile excel\_data. Prende ogni valore di DPI indicato nella tabella del foglio "DPI art.191" a riga 2, 4, 6, 8, .... ecc fino a che non si trova riga vuota alla prossima.
Esegue la seguente mappatura: (riporto di seguito il dizionario python dove il valore è la colonna corrispondente del foglio excel ad una data riga)

```python
context_tabella_dpi = {
    "tabella_dpi":[
        {"codice_DPI":"A",
        "descrizione": "B",
        "marca":"C",
        "modello":"D",
        "snr":"E",
        "H":"F",
        "L":"G",
        "M":"H",
        "note":""}
    ],
}
```

Scrivi il codice che implementa la logica di iterazione di questi valori.

## Tabella mansioni
Il codice legge il file excel indicato nella variabile excel\_data. Itera su tutti i fogli che si chiamano "Scheda n" con n=1,2,3,4,...
Per ogni foglio, estrae il nome del foglio (esempio "Scheda 1") ed estrae il valore della cella F8.
Assegna alle seguenti voci del dizionario i seguenti valori:
ID: Nome foglio
Mansione: Valore cella F8
N\_lavoratori: default a 1.

## Tabella HEG
Il codice legge il file excel indicato nella variabile excel\_data. Legge il foglio chiamato "Riepilogo". Itera sul numero di righe non nulle della tabella del foglio a partire da riga 2 (riga 1 è l'header con i titoli delle colonne) ed assegna per ogni riga i seguenti valori:
chiave: valore, con valore che indica la colonna rispettiva da prendere e chiave è la chiave del dizionario python del codice.

```plain
context_tabella_heg = {
    "tabella_HEG":[
        {"gruppo_HEG": "A", # 
        "numero_scheda":"B", # Numero della scheda del gruppo omogeneo (es: 1)
        "codice_HEG": "Default M01, M02, M03,..." ,#codice del gruppo omogeneo (es: M01)
        "parametro_riferimento": "C, #default
        "lex8h": "D",
        "incertezza": "E",
        "lexmax":"F",
        "peakmax":"G",
        "classe_rischio": "J", #es: alta, media o bassa
        "esposizione_vibrazioni": "K", #WBV, HAV o NO
        "esposizione_ototossici": "L", #si no
        "rumori_impulsivi":"M", #si no
        }
    ],
}
```
# =============================================================================
# config.py – Variabili di configurazione per revisione_vr.py
# Modifica qui tutti i parametri prima di eseguire lo script.
# =============================================================================


# -----------------------------------------------------------------------------
# Controllo flusso
# -----------------------------------------------------------------------------

# Se True, salta la fase di copia dati (orchestrazione_copia_incolla) e
# parte direttamente dal punto dopo la pausa (ricarica → nascondi → esporta PDF).
# Utile quando il file vr-excel è già stato modificato manualmente e
# si vuole solo rieseguire la parte di export.
SKIP_TO_POST_PAUSA = False
SAVE_MEASURE_PDF = False
SAVE_HEG_PDF = False
# -----------------------------------------------------------------------------
# Percorsi (path)
# -----------------------------------------------------------------------------

# Directory che contiene i file data-excel (mis*.xlsx)
PATH_DATA_EXCEL = "/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/SIT - RSM/rev2/Rumore/misure/data"

# Path completo del file vr-excel di partenza (template)
PATH_VR_EXCEL = '/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/SIT - RSM/rev2/Rumore/SIT_SA_CILINDRI_2026_rev2.xlsx'

# Directory di output: il file vr-excel modificato e i PDF verranno salvati qui
PATH_OUTPUT = "/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/SIT - RSM/rev2/Rumore" + "/output"

# -----------------------------------------------------------------------------
# Parametri revisione
# -----------------------------------------------------------------------------

# Numero di revisione: aggiunto come pedice agli ID (es. "D1" → "D1_2")
REVISIONE_NUMERO = 2

# Data delle misure (stringa libera, esempio: "21-lug-25")
DATA_MISURE = "26-feb-26"
STRATEGIA = "Compito"

# -----------------------------------------------------------------------------
# Nomi file output
# -----------------------------------------------------------------------------

# Nome del file excel valutazione rischio salvato nella directory di output
NOME_VR_EXCEL_OUT = "vr_out.xlsx"

# Nome del PDF del foglio Tab-Mis
NOME_PDF_TABELLA_MISURE = "Tabella_Misure.pdf"


# -----------------------------------------------------------------------------
# EXCEL_TAB-MIS – Colonne del data-excel e del foglio Tab-Mis nel vr-excel
# -----------------------------------------------------------------------------

# Colonna del data-excel che contiene l'ID lettera del lavoratore (es. "A", "B", ...)
DE_COL_LETTER_ID = "B"

# Colonna del data-excel che contiene il numero di traccia (nTrack)
DE_COL_NTRACK = "C"

# Colonna del data-excel con il livello equivalente ponderato A (LeqA_eq)
DE_COL_F = "F"

# Colonna del data-excel con il livello equivalente ponderato C (LeqC_eq)
DE_COL_I = "I"

# Colonna del data-excel con il picco massimo ponderato C (PeakC_max)
DE_COL_J = "J"

# Riga di intestazione nel data-excel (i dati iniziano dalla riga successiva)
DE_HEADER_ROW = 1

# -----------------------------------------------------------------------------
# EXCEL_TAB-MIS – Ricerca colonne del foglio Tab-Mis per stringa di intestazione
# Il codice legge la riga VR_HEADER_ROW e trova le colonne cercando queste stringhe.
# Aggiornare le stringhe se le intestazioni nel file Excel dovessero cambiare.
# -----------------------------------------------------------------------------

# Riga di intestazione del foglio Tab-Mis (dove ci sono i nomi delle colonne)
VR_HEADER_ROW = 3

# Stringa da cercare nell'intestazione per trovare la colonna dell'ID lavoratore
VR_ID_STR = "ID"

# Stringa da cercare per trovare le colonne con valore fisso 300 (durata massima traccia)
# Tutte le colonne il cui header contiene questa stringa riceveranno il valore 300
VR_300_STR = "Sec."

# Stringhe da cercare per trovare le colonne delle formule da copiare dalla riga template
# Un elemento per ogni colonna formula (nell'ordine in cui compaiono nel foglio)
VR_FORMULA_STRS = [
    "LAeq,T ",     # da verificare con l'intestazione reale del foglio Tab-Mis
    "LCeq,T",     # da verificare con l'intestazione reale del foglio Tab-Mis
    "Lpicco,C", # da verificare con l'intestazione reale del foglio Tab-Mis
    "LCEQ,TP",
    "LAEQ,Tp"
]

# Stringhe da cercare per trovare le colonne di misurazione per ogni traccia.
# Ogni stringa può comparire più volte nell'intestazione (una per ogni traccia);
# le occorrenze vengono abbinate in ordine per costruire la mappa nTrack → colonne.
VR_LEQA_STR  = "LAeq,T"   # colonne LeqA_eq  (es. R, X, AB, AF, AJ, AN)
VR_LEQC_STR  = "LCeq,T"   # colonne LeqC_eq  (es. S, Y, AC, AG, AK, AO)
VR_PEAK_STR  = "Lpicco,C"  # colonne PeakC_max (es. T, Z, AD, AH, AL, AP)



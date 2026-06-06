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

# Colonna di partenza per la copia delle formule dalla riga precedente
CP_FROM_CLM = "AQ"
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

# Colonna del foglio Tab-Mis in cui scrivere l'ID del lavoratore (con pedice revisione)
VR_COL_ID = "B"

# Colonna usata per individuare l'ultima riga occupata nel foglio Tab-Mis
VR_COL_LASTROW_SEARCH = "B"

# Colonne del foglio Tab-Mis in cui inserire il valore fisso 300 (durata massima traccia)
VR_COLS_300 = ["Q", "W", "AA", "AE", "AI", "AM"]

# Colonne del foglio Tab-Mis che contengono le formule da copiare dalla riga template
VR_FORMULA_COLS = ["M", "N", "O", "P"]

# Colonna di partenza (lettera) per la copia delle celle dalla riga precedente nel foglio Tab-Mis
VR_CP_FROM_COL = "AQ"

# Mappa nTrack → (colonna LeqA_eq, colonna LeqC_eq, colonna PeakC_max) nel foglio Tab-Mis
VR_NTRACK_MAP = {
    1: ("R",  "S",  "T"),
    2: ("X",  "Y",  "Z"),
    3: ("AB", "AC", "AD"),
    4: ("AF", "AG", "AH"),
    5: ("AJ", "AK", "AL"),
    6: ("AN", "AO", "AP"),
}



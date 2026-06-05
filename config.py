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



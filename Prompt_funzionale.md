# Codice revisione rischi - prompt

# Descrizione
Il codice deve servire per automatizzare il flusso di lavoro tra i dati di un file ecxel ed un altro foglio ecxel.
Chiamiamo gli ecxel dei dati, **data-ecxel**, mentre chiamiamo l'ecxel su cui vogliamo copiare i dati **vr-ecxel**. Il vr-ecxel è il documento ecxel su cui poi il codice deve eseguire delle operazioni per poter completare il processo descritto nelle sezioni seguenti.
Infine il codice esporta in formato pdf le specifiche aree di lavoro del vc-ecxel in una cartella definita dall'utente

# Generalità
Il codice deve essere scritto in linguaggio python, utilizzando la libreria openpuxl per interagire con i file ecxel.
Per l'esportazione in pdf, consigliami quale delle opzioni è migliore:
1. Usando LibreOffice tramite CLI (sempre implementata nel codice python)
2. Usare **xlsxwriter + reportlab**
Inserisci nello script le variabili globali all'inizio del documento, quali per esempio:
**PATH\_DATA\_ECXEL** = stringa, contiene il path dei documenti data-ecxel
**PATH\_VR\_ECXEL** = stringa, contiene il path del documento vr-ecxel
**PATH\_OUTPUT** = stringa, contiene il path dei file della directory di lavoro e di dove verranno salvati i dati
**REVISIONE\_NUMERO** = intero, serve per etichettare meglio le ID delle misure, come descritto di seguito
**DATA\_MISURE** = datetime, formato <gg><mese><anno> (esempio: 11-lug-21)

In tutte le funzioni, ovunque serve il riferimento alle colonne o alle righe del documento ecxel, impostale come variabili pre-configurate all'interno della funzione specifica, in modo che sia più facile la modifica se necessario.

Imposta log per debug e stampa i punti chiave delle funzioni in modo da tenere traccia di quello che sta succedendo.
# Funzionalità del codice
Il codice deve avere le seguenti sezioni:
## Funzione data-ecxel2vr-ecxel
La funzione riceve gli input:
*   path con nome del documento ecxel data-ecxel
*   il path del vr-ecxel
*   l'intero REVISIONE\_NUMERO
*   DATA\_MISURE
*   nome\_foglio = stringa, default="Tab-Mis"

In tutta la seguente descrizione le righe e le colonne del documento vr-ecxel sono riferite al foglio chiamato nome\_foglio.
Copia le righe uniche della colonna B del documento data-ecxel (chiamata letter\_ID alla riga 1) nella colonna B del documento vr-ecxel sotto l'ultima riga occupata (se serve deve creare nuove righe).
Esempio di struttura di un documento data-ecxel:

| fileID | letter\_ID | nTrack | LeqA\_min | LeqA\_max | LeqA\_eq | LeqC\_min | LeqC\_max | LeqC\_eq | PeakC\_max | PeakC\_eq | LASeq\_T | LAIeq\_T | durata | inizio | fine |
| ---| ---| ---| ---| ---| ---| ---| ---| ---| ---| ---| ---| ---| ---| ---| --- |
| 20260213\_092322\_094202.csv | D1 | 1 | 43,2 | 69,1 | 51,5 | 63,3 | 71,4 | 66,7 | 92,8 | 76,6 | 51,5 | 61,9 | 00:03:06 | 09:23:25 | 09:26:31 |
| 20260213\_092322\_094202.csv | D1 | 2 | 43,5 | 67,8 | 54,9 | 62,5 | 76,7 | 66,4 | 99,2 | 77,5 | 54,9 | 64,6 | 00:03:05 | 09:26:32 | 09:29:37 |
| 20260213\_092322\_094202.csv | D1 | 3 | 47,1 | 74,8 | 59,5 | 62,2 | 83 | 67,7 | 103 | 78,8 | 59,5 | 68,8 | 00:03:04 | 09:29:39 | 09:32:43 |
| 20260213\_092322\_094202.csv | D1 | 4 | 43,5 | 65,2 | 53,8 | 60,8 | 72,7 | 64,8 | 92,8 | 75,9 | 53,9 | 62,4 | 00:03:03 | 09:32:46 | 09:35:49 |
| 20260213\_092322\_094202.csv | D1 | 5 | 45,6 | 72,1 | 59,7 | 61,5 | 79,2 | 67,1 | 100,5 | 78,7 | 59,6 | 68,4 | 00:03:02 | 09:35:53 | 09:38:55 |
| 20260213\_092322\_094202.csv | D1 | 6 | 44,3 | 71 | 57,2 | 61,9 | 73,7 | 65,6 | 92,4 | 77,8 | 57,2 | 66,1 | 00:03:01 | 09:39:00 | 09:42:01 |
| 20260213\_094347\_100241.csv | D2 | 1 | 46,6 | 77,9 | 65,5 | 65,2 | 80,6 | 69,8 | 95,5 | 82,4 | 65,6 | 72,4 | 00:03:08 | 09:43:52 | 09:47:00 |
| 20260213\_094347\_100241.csv | D2 | 2 | 47,5 | 85,5 | 69,4 | 64,7 | 86,5 | 71,8 | 107 | 83,1 | 69,3 | 78,1 | 00:03:07 | 09:47:01 | 09:50:08 |
| 20260213\_094347\_100241.csv | D2 | 3 | 46,4 | 79,6 | 64,8 | 64,8 | 78,7 | 68,9 | 102,4 | 80,3 | 64,8 | 75,1 | 00:03:06 | 09:50:10 | 09:53:16 |
| 20260213\_094347\_100241.csv | D2 | 4 | 47,9 | 82,6 | 67,8 | 65,7 | 82,8 | 70,8 | 103,1 | 84,4 | 67,8 | 76,4 | 00:03:05 | 09:53:19 | 09:56:24 |

Per copia di righe uniche si intende che nella colonna B si devono prendere solo i valori univoci del letter\_ID (in questo caso per esempio solo "D1" e "D2").

La colonna B del documento data-ecxel contiene stringhe in formato "F1", "D4", ecc. Prima di incollare i valori in vr-ecxel, inserisci il pedice "\_<numero\_revisione>" ala stringa di partenza (esempio se numero\_revisione = 2: "F3" → "F3\_2")

Inserisce il valore intero 300 nelle colonne Q,W,AA,AE,AI,AM delle righe appena copiate nel documento vr-ecxel.

La colonna C (chiamata nTrack in riga 1) del documento data-ecxel contiene valori numerici interi come 1,2,3,4,5,6. A questi corrisponde lo stesso valore in letter\_ID (colonna B).
Per ogni numero che compare per lo stesso letter\_ID, occorre copiare la riga corrispondente al numero di nTrack e alle colonne F,I e J (numeri reali con massimo 1 decimale) rispettivamente nelle colonne R,S e T (oppure X,Y,Z, oppure AB,AC,AD oppure AF,AG,AH oppure AJ,Ak,AL oppure AN,AO,AP) del documento vr-ecxel, in base al numero che compare in nTrack.
Esempio:
se nTrack = 1 copio i valori nelle colonne R,S e T; se nTrack = 2 copio i valori in X,Y,Z; se nTrack = 3 copio i valori in AB,AC,AD e così via.

Per ogni valore copiato ed inserito nelle righe specifiche, copia le formule presente nelle colonne M,N,O,P nelle righe aggiunte con i valori nuovi di ID in modo che i riferimenti di riga siano aggiornati.

## Funzione orchestrazione\_copia\_incolla
La funzione prende in ingresso
*   il path dei documenti data-ecxel
*   il path del documento vr-ecxel

Per ogni elemento che inizia per "mis" e finisce con "xlsx", utilizza la funzione **_data-ecxel2vr-ecxel_** per copiare i valori selezionati secondo quando specificato sopra.

## Funzione nascondi\_righe\_tab\_mis
La funzione prende in ingresso
*   nome\_foglio = stringa, default = "Tab-Mis"
*   il path del documento vr-ecxel = default = **PATH\_VR\_ECXEL**

La funzione itera su tutti i fogli del documento chiamati "Scheda 1", "Scheda 2", ... ecc e salva in una variabile lista i valori della colonna E ( da riga 11 in avanti fino all'ultima riga prima delle celle nascoste). Fatto questo, seleziona il foglio nome\_foglio e nasconde tutte le righe tali per cui i valori di colonna B siano presenti nella lista.

## Funzione select\_area\_stampa
La funzione prende in ingresso il documento xlsx caricato.
Seleziona per il foglio Tab-Mis l'area di stampa che va da A1 a BC<ultima\_riga\_popolata>
Seleziona per i fogli "Scheda 1", "Scheda 2", ... ecc l'area da A1 a Q66.

## Funzione esporta\_pdf\_misure
Input:
output\_name = stringa, default = Tabella\_Misure.pdf
La funzione esporta in pdf il foglio Tab-Mis secondo l'area di stampa selezionata precedentemente dalla funzione select\_area\_stampa. Scegli quale metodologia usare e riportarmi la motivazione prima di procedere.
L'output deve chiamarsi output\_name.
## Funzione esporta\_pdf\_schede
Input:
*   **PATH\_OUTPUT**
La funzione esporta in pdf le schede HEG chiamate "Scheda 1", "Scheda 2",.... fino alla fine delle schede disponibili del file excel caricato. Controlla che l'area di stampa sia la stessa salvata dopo l'esecuzione della funzione select\_area\_stampa (riapplica la funzione in caso).
Salva i pdf con la forma <nome\_scheda>\_DPI\_<id\_DPI>.pdf dove:
*   <id\_DPI> è il valore stringa che si trova nella cella E62 di ogni foglio Scheda.
La funzione esegue questo in modo iterativo, per ogni valore di DPI che è presente nella colonna A del foglio "DPI art.191" del file excel, righe 2:2:n (ossia riga 2, 4, 6, ecc) fino a che non si trova valore nullo.
L'esempio della tabella riportata in questo foglio è mostrato di seguito

| ID | Tipologia | Marca | Modello | SNR | H | M | L | dm |
| ---| ---| ---| ---| ---| ---| ---| ---| --- |
| A1 | Inserti espandibili | HONEYWELL | LASER LITE | 35 | 34 | 32 | 31 |  |
|  |  | Coeffciente correttivo β = | 0,5 | 18 | 17 | 16 | 16 | 11 |
| A3 | Cuffie auricolari | PELTOR | H510A | 27 | 32 | 25 | 15 |  |
|  |  | Coeffciente correttivo β = | 0,75 | 20 | 24 | 19 | 11 | 6 |
| A4 |  |  |  |  |  |  |  |  |
|  |  | Coeffciente correttivo β = | 0,65 | 0 | 0 | 0 | 0 | \-5 |

Si utilizzano i valori ID di questa tabella (chiamiamola tabella dpi) e in maniera iterativa si sostituiscono ad ogni foglio chiamato "Scheda n" (n = 1, 2, 3,...) nella cella E62.
Esempio operativo:
*   si parte dal valore A1 e si sostituisce nei fogli Scheda 1, Scheda 2, ...,
*   si esportano in pdf tutti i fogli
*   si cambia valore e si imposta A2 (se presente) e lo si sostituisce nei fogli
*   si esporta in pdf tutti i fogli.
*   Si procede fino a che si hanno ID dei DPI.

## Funzione orchestratore
La funzione governa tutte le funzioni in modo da creare un flusso di lavoro principale.
Riceve in input:
*   PATH\_OUTPUT
*   nome\_vr\_ecxel\_out = stringa, default a "vr\_out.xlsx"
*   tutti gli input necessari per chiamare le altre funzioni

La funzione deve eseguire i seguenti step:
1. Chiamata orchestrazione\_copia\_incolla per inserire i dati necessari dai documento data-ecxel al documento vr-ecxel.
2. Salvataggio di vr-ecxel nella directory **PATH\_OUTPUT** con nome **nome\_vr\_ecxel\_out**
3. **Attesa**. Il codice si mette in pausa e stampa la stringa: "In attesa di modifica del documento {**nome\_vr\_ecxel\_out**} Modifica le schede HEG e i DPI del documento".

L'utente a questo punto modificherà il documento **nome\_vr\_ecxel\_out** e il codice ripartirà una volta premuto invio nel terminale.

4. Ricarica il documento xlsx **nome\_vr\_ecxel\_out**
5. Nasconde le righe in Tab-Mis con la funzione nascondi\_righe\_tab\_mis
6. Seleziona area di stampa con la funzione select\_area\_stampa
7. Salva il documento **nome\_vr\_ecxel\_out**
8. Esporta in pdf il foglio "Tab-Mis" chiamando la funzione esporta\_pdf\_misure e salvalo nella directory **PATH\_OUTPUT/PDFs**
9. Esporta in pdf i fogli "Scheda 1", "Scheda 2", ... chiamando la funzione esporta\_pdf\_schede e salvali nella directory **PATH\_OUTPUT/PDFs**

### Punti critici
Pausa del codice: prevedi un modo per poter scegliere se eseguire il codice saltando tutta la prima parte fino all'attesa (quindi ripartire dalla sezione dopo l'attesa) o se eseguire il codice per intero, per esempio con una flag True o False in una parte di codice

## Punti fondamentali
Metti tutte le variabili globali e quelle variabili che devono essere modificate dall'utente in un file .py con una descrizione sintetica di quello che rappresentano, in modo che possano essere modificate da li. Aggiungi anche la flag se partire dall'inizio o dopo la pausa del codice.
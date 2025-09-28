Certamente. Ecco il file `README.md` aggiornato con le istruzioni complete per l'uso del programma, tenendo conto di tutte le modifiche apportate.

-----

# Analisi Semantica e Classificazione di Professioni

Questo script Python è uno strumento di analisi dati progettato per elaborare un elenco di professioni da un file Excel. Esegue diverse operazioni, tra cui la pulizia dei dati, la classificazione delle professioni (gestionali vs. operative), l'analisi di correlazione semantica con parole chiave predefinite e la generazione di report e grafici di visualizzazione interattivi.

## Funzionalità Principali

  * **Caricamento e Pulizia Dati**: Legge i dati da un file Excel (`FABBISOGNI_all.xlsx`) e normalizza i nomi delle professioni.
  * **Analisi Statistica**: Identifica e visualizza le 10 professioni con più entrate programmate e le 10 più difficili da reperire.
  * **Classificazione Gerarchica**: Classifica ogni professione come `Gestionale` o `Operativa` in base a termini manageriali.
  * **Correlazione Semantica**: Utilizza un modello linguistico (`sentence-transformer`) per trovare correlazioni tra le professioni e una lista di competenze e keyword.
  * **Generazione di Report**:
      * Salva tutte le analisi in un file Excel (`Results.xlsx`) organizzato in più fogli.
      * Crea una cartella (`GRAPH/`) per salvare i grafici generati.
  * **Visualizzazione Dati**:
      * Crea grafici a barre statici (`.png`) per le analisi "Top 10".
      * Genera grafici a bolle interattivi (`.html`) per analizzare le relazioni tra le variabili.
      * Mostra automaticamente i grafici interattivi nel browser durante l'esecuzione.

## Dipendenze

Per eseguire lo script, è necessario avere installato Python 3 e le seguenti librerie:

  * `pandas`
  * `openpyxl`
  * `matplotlib`
  * `plotly`
  * `sentence-transformers`
  * `torch`
  * `transformers`

## Installazione

1.  Assicurati di avere Python 3 installato sul tuo sistema.

2.  Apri un terminale o un prompt dei comandi.

3.  Installa tutte le librerie necessarie con un unico comando:

    ```bash
    pip install pandas openpyxl matplotlib plotly sentence-transformers torch transformers
    ```

## Istruzioni per l'Uso

### 1\. Preparazione del File di Input

Prima di lanciare lo script, è fondamentale preparare il file di dati di input.

  * **Posizione**: Il file deve trovarsi nella stessa cartella dello script Python.
  * **Nome del file**: Deve essere nominato esattamente `FABBISOGNI_all.xlsx`.
  * **Nome del foglio**: Il foglio di lavoro contenente i dati deve chiamarsi `  PROFESSIONE LVL4 CP2011 `. **Attenzione:** il nome nel codice inizia con uno spazio.
  * **Struttura delle colonne**: Il foglio deve contenere 4 colonne nel seguente ordine esatto:
    1.  Nome della professione
    2.  ID CP2011 LVL4 (codice identificativo)
    3.  Numero di entrate programmate
    4.  Numero di entrate di difficile reperimento

### 2\. Esecuzione dello Script

1.  Apri un terminale o un prompt dei comandi nella cartella dove hai salvato lo script e il file Excel.

2.  Esegui lo script con il comando:

    ```bash
    python nome_del_tuo_script.py
    ```

3.  Lo script ti chiederà di inserire un valore per la soglia di similarità semantica (`score_threshold`). Questo valore (da 0 a 1) determina la "vicinanza" semantica richiesta tra una professione e una keyword.

    ```
    Inserisci il valore di score_threshold (premi Invio per usare il valore predefinito 0.6):
    ```

      * Puoi inserire un valore decimale (es. `0.55`) e premere Invio.
      * Se premi `Invio` senza inserire nulla, verrà usato il valore predefinito **0.6**.

### 3\. Visualizzazione

Durante l'esecuzione, lo script aprirà automaticamente ogni grafico a bolle interattivo in una nuova scheda del tuo browser predefinito.

## Descrizione dell'Output

Al termine dell'esecuzione, troverai i seguenti file e cartelle:

1.  **File Excel (`Results.xlsx`)**: Un report completo contenente i seguenti fogli:

      * `Top 10 Entrate`: Le 10 professioni più richieste.
      * `Top 10 Difficili`: Le 10 professioni più difficili da trovare.
      * `Professioni Gestionali`: L'elenco completo delle professioni classificate come gestionali.
      * `Professioni Operative`: L'elenco completo delle professioni classificate come operative.
      * `Correlazione Gestionali-Keyword`: Professioni gestionali con le relative keyword correlate.
      * `Correlazione Operative-Keyword`: Professioni operative con le relative keyword correlate.

2.  **Cartella `GRAPH/`**: Una cartella creata automaticamente per contenere tutti i grafici:

      * **Grafici a barre statici**:
          * `top10_entrate.png`
          * `top10_difficili.png`
      * **Grafici a bolle interattivi**:
          * `analisi_tutti_livelli_[timestamp].html`
          * `professioni_gestionali_[timestamp].html`
          * `correlazione_gestionali_keyword_[timestamp].html`
          * `correlazione_operative_keyword_[timestamp].html`

    *(Nota: `[timestamp]` rappresenta la data e l'ora di creazione del file, es. `20250928_042830`)*
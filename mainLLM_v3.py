import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import os
import re
from sentence_transformers import SentenceTransformer, util
from datetime import datetime

# ---
# Caricamento del modello di embedding per l'analisi semantica
# ---
print("Caricamento del modello di embedding per l'analisi semantica...")
model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
print("Modello caricato con successo.")

# ---
# Struttura dati e funzioni per la classificazione e l'analisi semantica
# ---

# Lista con le coppie singolare/plurale di termini manageriali
managerial_terms = [
    ("direttore", "direttori"),
    ("dirigente", "dirigenti"),
    ("imprenditore", "imprenditori"),
    ("responsabile", "responsabili"),
    ("manager", "managers"),
    ("coordinatore", "coordinatori"),
    ("amministratore", "amministratori"),
    ("capo", "capi"),
    ("gestore", "gestori"),
    ("direzione", "direzioni"),
    ("gestione", "gestioni"),
    ("coordinamento", "coordinamenti")
]

# Costruzione dell'espressione regolare che cerca le coppie esatte singolare/plurale
managerial_regex_list = [t[0] for t in managerial_terms] + [t[1] for t in managerial_terms]
managerial_regex = re.compile(r'\b(' + '|'.join(managerial_regex_list) + r')\b', re.IGNORECASE)


def clean_profession_name(name):
    """Pulisce il nome della professione da caratteri speciali e spazi extra."""
    return re.sub(r'[^\w\s]', '', str(name)).strip()


def classify_profession(profession_name, regex):
    """
    Classifica una professione come 'Gestionale' o 'Operativa'
    in base al suo nome, usando espressioni regolari.
    """
    if regex.search(profession_name):
        return 'Gestionale'
    return 'Operativa'


# Lista di keyword fornite per la correlazione
keywords = [
    "Risk Management", "Eccellenza operativa", "Metodologie agili", "Leadership", "Ricerca e Sviluppo",
    "Testing", "Verifica e validazione", "Management", "Project Management", "Programming", "Python",
    "Project delivery", "Coaching", "Key Performance Indicator", "Trasformazione digitale", "Direttore",
    "Lavoro di squadra", "Operations", "Commissioning", "Ingegneria dei processi", "Strategy",
    "Analisi dei dati", "Functional test", "Negoziazione", "Troubleshooting", "Ingegnere elettronico",
    "Produzione industriale", "Dirigente", "analisi dati", "intelligenza artificiale", "statistica",
    "Comunicazione", "Machinery", "Mentoring", "Team Management", "Gestione dei dati",
    "Training", "Conduzione training", "Automotive", "Electrification", "SW Engineering", "Data Analyst"
]

# Calcolo degli embedding per le keyword
keyword_embeddings = model.encode(keywords, convert_to_tensor=True)


def find_related_keywords(profession_name, keywords, keyword_embeddings, model, score_threshold):
    """
    Trova corrispondenze semantiche tra una professione e le keyword
    utilizzando vettori di embedding. Se la soglia non trova corrispondenze,
    ne tenta una più bassa.
    """
    found_matches = []

    # Matching semantico tramite similarità del coseno
    profession_embedding = model.encode(profession_name, convert_to_tensor=True)
    cosine_scores = util.cos_sim(profession_embedding, keyword_embeddings)[0]

    # Prova con la soglia specificata dall'utente
    for i, score in enumerate(cosine_scores):
        if score > score_threshold:
            if keywords[i] not in found_matches:
                found_matches.append(keywords[i])

    # Se non sono state trovate corrispondenze, prova con una soglia più bassa
    if not found_matches:
        for i, score in enumerate(cosine_scores):
            if score > 0.4:
                if keywords[i] not in found_matches:
                    found_matches.append(keywords[i])

    return found_matches


# ---
# Inizio del programma principale
# ---

# Configurazione file e cartelle
file_excel_input = 'FABBISOGNI_all.xlsx'
file_excel_output = 'Results.xlsx'
nome_foglio = ' PROFESSIONE LVL4 CP2011'
output_folder = 'GRAPH' # Cartella di output per i grafici

# Crea la cartella se non esiste
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Chiede all'utente di inserire il valore di score_threshold
score_input = input("Inserisci il valore di score_threshold (premi Invio per usare il valore predefinito 0.6): ")
score_threshold = 0.6
if score_input:
    try:
        score_threshold = float(score_input)
    except ValueError:
        print("Valore non valido. Verrà usato il valore predefinito di 0.6.")

# Caricamento e lettura del file Excel
try:
    df_professioni = pd.read_excel(file_excel_input, sheet_name=nome_foglio)
except FileNotFoundError:
    print(f"Errore: il file '{file_excel_input}' non è stato trovato.")
    exit()
except ValueError as e:
    print(f"Errore: il foglio '{nome_foglio}' non è stato trovato in '{file_excel_input}'.")
    exit()

df_professioni.columns = [
    'PROFESSIONE',
    'ID_CP2011_LVL4',
    'ENTRATE_PROGRAMMATE',
    'DI_CUI_DI_DIFFICILE_REPERIMENTO'
]

writer = pd.ExcelWriter(file_excel_output, engine='openpyxl')

# Applica la pulizia del nome della professione
df_professioni['PROFESSIONE'] = df_professioni['PROFESSIONE'].apply(clean_profession_name)

# ---
# Analisi Statistiche e Grafici a Barre
# ---

# Analisi 1: Top 10 per entrate programmate
df_top10_entrate = df_professioni.sort_values(by='ENTRATE_PROGRAMMATE', ascending=False).head(10)
print("Top 10 Professioni per Entrate Programmate:")
print(df_top10_entrate[['PROFESSIONE', 'ENTRATE_PROGRAMMATE']].to_string(index=False))
df_top10_entrate.to_excel(writer, sheet_name='Top 10 Entrate', index=False)
plt.figure(figsize=(12, 8))
plt.bar(df_top10_entrate['PROFESSIONE'], df_top10_entrate['ENTRATE_PROGRAMMATE'], color='skyblue')
plt.title('Top 10 Professioni per Entrate Programmate')
plt.xlabel('Professione')
plt.ylabel('Entrate Programmate')
plt.xticks(rotation=90, ha='right', fontsize=10)
plt.tight_layout()
plt.savefig(os.path.join(output_folder, 'top10_entrate.png'))
plt.close()

# Analisi 2: Top 10 per difficoltà di reperimento
df_top10_difficili = df_professioni.sort_values(by='DI_CUI_DI_DIFFICILE_REPERIMENTO', ascending=False).head(10)
print("\nTop 10 Professioni per Difficoltà di Reperimento:")
print(df_top10_difficili[['PROFESSIONE', 'DI_CUI_DI_DIFFICILE_REPERIMENTO']].to_string(index=False))
df_top10_difficili.to_excel(writer, sheet_name='Top 10 Difficili', index=False)
plt.figure(figsize=(12, 8))
plt.bar(df_top10_difficili['PROFESSIONE'], df_top10_difficili['DI_CUI_DI_DIFFICILE_REPERIMENTO'], color='tomato')
plt.title('Top 10 Professioni per Difficoltà di Reperimento')
plt.xlabel('Professione')
plt.ylabel('Entrate di Difficile Reperimento')
plt.xticks(rotation=90, ha='right', fontsize=10)
plt.tight_layout()
plt.savefig(os.path.join(output_folder, 'top10_difficili.png'))
plt.close()

# ---
# Classificazione e Correlazione Semantica
# ---
print("\nAvvio dell'analisi di classificazione e correlazione semantica...")

# 1. Classificazione
df_professioni['LIVELLO_PROFESSIONE'] = df_professioni['PROFESSIONE'].apply(
    lambda x: classify_profession(x, managerial_regex)
)
print("Classificazione completata.")

# 2. Correlazione Semantica
df_professioni['KEYWORD_CORRELATE'] = df_professioni['PROFESSIONE'].apply(
    lambda x: find_related_keywords(x, keywords, keyword_embeddings, model, score_threshold)
)
print("Correlazione semantica completata.")

# Filtra, ordina e salva le professioni gestionali
df_gestionali = df_professioni[df_professioni['LIVELLO_PROFESSIONE'] == 'Gestionale'].copy()
df_gestionali_sorted = df_gestionali.sort_values(by='DI_CUI_DI_DIFFICILE_REPERIMENTO', ascending=False)
df_gestionali_sorted.to_excel(writer, sheet_name='Professioni Gestionali', index=False)

# Filtra, ordina e salva le professioni operative
df_operative = df_professioni[df_professioni['LIVELLO_PROFESSIONE'] == 'Operativa'].copy()
df_operative_sorted = df_operative.sort_values(by='DI_CUI_DI_DIFFICILE_REPERIMENTO', ascending=False)
df_operative_sorted.to_excel(writer, sheet_name='Professioni Operative', index=False)

# Filtra e salva i dati di correlazione per le professioni gestionali
df_gest_keywords = df_gestionali_sorted[df_gestionali_sorted['KEYWORD_CORRELATE'].apply(len) > 0].copy()
df_gest_keywords.to_excel(writer, sheet_name='Correlazione Gestionali-Keyword', index=False)

# Filtra e salva i dati di correlazione per le professioni operative
df_op_keywords = df_operative_sorted[df_operative_sorted['KEYWORD_CORRELATE'].apply(len) > 0].copy()
df_op_keywords.to_excel(writer, sheet_name='Correlazione Operative-Keyword', index=False)

# ---
# Generazione Grafici Interattivi a Bolle
# ---
print("\nGenerazione grafici a bolle interattivi...")

# Crea un timestamp per rendere unici i nomi dei file
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# Grafico 1: Analisi completa (tutti i livelli)
fig_all_levels = px.scatter(
    df_professioni,
    x='ENTRATE_PROGRAMMATE',
    y='DI_CUI_DI_DIFFICILE_REPERIMENTO',
    size='ENTRATE_PROGRAMMATE',
    color='LIVELLO_PROFESSIONE',
    hover_name='PROFESSIONE',
    hover_data={'LIVELLO_PROFESSIONE': True, 'ENTRATE_PROGRAMMATE': True, 'DI_CUI_DI_DIFFICILE_REPERIMENTO': True},
    title='Analisi di Tutte le Professioni: Entrate vs. Difficoltà di Reperimento',
    labels={'ENTRATE_PROGRAMMATE': 'Entrate Programmate',
            'DI_CUI_DI_DIFFICILE_REPERIMENTO': 'Entrate di Difficile Reperimento'},
    size_max=60
)
fig_all_levels.show() # RIPRISTINATO
fig_all_levels.write_html(os.path.join(output_folder, f"analisi_tutti_livelli_{timestamp}.html"))

# Grafico 2: Solo per le professioni gestionali
fig_gestionali = px.scatter(
    df_gestionali_sorted,
    x='ENTRATE_PROGRAMMATE',
    y='DI_CUI_DI_DIFFICILE_REPERIMENTO',
    size='ENTRATE_PROGRAMMATE',
    color='DI_CUI_DI_DIFFICILE_REPERIMENTO',
    hover_name='PROFESSIONE',
    hover_data={'ENTRATE_PROGRAMMATE': True, 'DI_CUI_DI_DIFFICILE_REPERIMENTO': True},
    title='Professioni di Livello Gestionale: Entrate vs. Difficoltà di Reperimento',
    labels={'ENTRATE_PROGRAMMATE': 'Entrate Programmate',
            'DI_CUI_DI_DIFFICILE_REPERIMENTO': 'Entrate di Difficile Reperimento'},
    size_max=60
)
fig_gestionali.show() # RIPRISTINATO
fig_gestionali.write_html(os.path.join(output_folder, f"professioni_gestionali_{timestamp}.html"))

# Grafico 3: Correlazione tra Professioni Gestionali e Keyword
fig_gest_keywords = px.scatter(
    df_gest_keywords,
    x='ENTRATE_PROGRAMMATE',
    y='DI_CUI_DI_DIFFICILE_REPERIMENTO',
    size='ENTRATE_PROGRAMMATE',
    color='DI_CUI_DI_DIFFICILE_REPERIMENTO',
    hover_name='PROFESSIONE',
    hover_data={'PROFESSIONE': True, 'KEYWORD_CORRELATE': True, 'ENTRATE_PROGRAMMATE': True, 'DI_CUI_DI_DIFFICILE_REPERIMENTO': True},
    title='Correlazione: Professioni Gestionali vs. Keyword (Analisi Semantica)',
    labels={'ENTRATE_PROGRAMMATE': 'Entrate Programmate',
            'DI_CUI_DI_DIFFICILE_REPERIMENTO': 'Entrate di Difficile Reperimento'},
    size_max=60
)
fig_gest_keywords.show() # RIPRISTINATO
fig_gest_keywords.write_html(os.path.join(output_folder, f"correlazione_gestionali_keyword_{timestamp}.html"))

# Grafico 4: Correlazione tra Professioni Operative e Keyword
fig_op_keywords = px.scatter(
    df_op_keywords,
    x='ENTRATE_PROGRAMMATE',
    y='DI_CUI_DI_DIFFICILE_REPERIMENTO',
    size='ENTRATE_PROGRAMMATE',
    color='DI_CUI_DI_DIFFICILE_REPERIMENTO',
    hover_name='PROFESSIONE',
    hover_data={'PROFESSIONE': True, 'KEYWORD_CORRELATE': True, 'ENTRATE_PROGRAMMATE': True, 'DI_CUI_DI_DIFFICILE_REPERIMENTO': True},
    title='Correlazione: Professioni Operative vs. Keyword (Analisi Semantica)',
    labels={'ENTRATE_PROGRAMMATE': 'Entrate Programmate',
            'DI_CUI_DI_DIFFICILE_REPERIMENTO': 'Entrate di Difficile Reperimento'},
    size_max=60
)
fig_op_keywords.show() # RIPRISTINATO
fig_op_keywords.write_html(os.path.join(output_folder, f"correlazione_operative_keyword_{timestamp}.html"))

# ---
# Salvataggio Finale
# ---
writer.close()
print(f"\nAnalisi completata. I risultati sono stati salvati in '{file_excel_output}' e nella cartella '{output_folder}'.")
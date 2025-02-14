import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Strumento di Confronto Excel")
st.write("Seleziona il tipo di confronto che desideri effettuare:")

# Menu a tendina per la selezione del confronto
confronto_tipo = st.selectbox(
    "Tipo di confronto",
    ("Confronto Maxi e Market", "Confronto Maxi e Maxi", "Confronto Market e Market")
)

# Carica i file in base al tipo di confronto selezionato
if confronto_tipo == "Confronto Maxi e Market":
    st.write("Carica i file per il confronto tra Maxi e Market.")
    file_maxi = st.file_uploader("Carica il file MAXIPIU", type=["xlsx"])
    file_market = st.file_uploader("Carica il file MARKET", type=["xlsx"])
elif confronto_tipo == "Confronto Maxi e Maxi":
    st.write("Carica il file per il confronto tra Maxi e Maxi.")
    file_maxi1 = st.file_uploader("Carica il primo file MAXIPIU", type=["xlsx"])
    file_maxi2 = st.file_uploader("Carica il secondo file MAXIPIU", type=["xlsx"])
elif confronto_tipo == "Confronto Market e Market":
    st.write("Carica il file per il confronto tra Market e Market.")
    file_market1 = st.file_uploader("Carica il primo file MARKET", type=["xlsx"])
    file_market2 = st.file_uploader("Carica il secondo file MARKET", type=["xlsx"])

# Funzione per eseguire il confronto
def esegui_confronto(df1, df2, nome1, nome2):
    # Rinomina le colonne con riferimenti stile Excel
    df1.columns = ['A', 'B', 'C', 'D', 'E']
    df2.columns = ['A', 'B', 'C', 'D', 'E']
    
    # Unisci i due file sulla colonna 'A' (che rappresenta il 'codice'), mantenendo tutti i prodotti
    comparison_df = pd.merge(df1, df2, on="A", how="outer", suffixes=(f'_{nome1}', f'_{nome2}'))
    
    # Identifica le differenze
    def identify_differences(row):
        differences = []
        if row[f'C_{nome1}'] != row[f'C_{nome2}']:  # Colonna C per Net
            differences.append('Net')
        if row[f'D_{nome1}'] != row[f'D_{nome2}']:  # Colonna D per Cessione
            differences.append('Cessione')
        if row[f'E_{nome1}'] != row[f'E_{nome2}']:  # Colonna E per Uscita
            differences.append('Uscita')
        return ", ".join(differences)

    # Applica la funzione di confronto
    comparison_df['Confronto'] = comparison_df.apply(identify_differences, axis=1)
    
    # Seleziona tutte le righe, anche quelle senza differenze
    result_df = comparison_df[['A', f'B_{nome1}', f'C_{nome1}', f'C_{nome2}', f'D_{nome1}', f'D_{nome2}', f'E_{nome1}', f'E_{nome2}', 'Confronto']]
    
    # Rinomina le colonne per maggiore chiarezza
    result_df.columns = ['Codice', 'Descrizione', f'Net ({nome1})', f'Net ({nome2})', f'Cessione ({nome1})', f'Cessione ({nome2})', f'Uscita ({nome1})', f'Uscita ({nome2})', 'Confronto']
    
    # Mostra i risultati nel browser
    st.write("Risultati del confronto:")
    st.dataframe(result_df)
    
    # Esportazione in Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
    output.seek(0)

    # Bottone per scaricare il file
    st.write("Scarica i risultati in formato Excel:")
    st.download_button(
        label="Scarica Excel",
        data=output,
        file_name="risultati_confronto.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Esegui il confronto in base alla selezione
if confronto_tipo == "Confronto Maxi e Market" and file_maxi and file_market:
    # Carica e confronta i file Maxi e Market
    df1 = pd.read_excel(file_maxi, header=0)
    df2 = pd.read_excel(file_market, header=0)
    esegui_confronto(df1, df2, "Maxi", "Market")

elif confronto_tipo == "Confronto Maxi e Maxi" and file_maxi1 and file_maxi2:
    # Carica e confronta i due file Maxi
    df1 = pd.read_excel(file_maxi1, header=0)
    df2 = pd.read_excel(file_maxi2, header=0)
    esegui_confronto(df1, df2, "Maxi1", "Maxi2")

elif confronto_tipo == "Confronto Market e Market" and file_market1 and file_market2:
    # Carica e confronta i due file Market
    df1 = pd.read_excel(file_market1, header=0)
    df2 = pd.read_excel(file_market2, header=0)
    esegui_confronto(df1, df2, "Market1", "Market2")

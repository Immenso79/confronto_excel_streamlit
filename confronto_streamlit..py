import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Strumento di Confronto Excel")
st.write("Carica i due file Excel (MAXIPIU e MARKET) per confrontare i dati.")

# Carica i file Excel
file_maxi = st.file_uploader("Carica il file MAXIPIU", type=["xlsx"])
file_market = st.file_uploader("Carica il file MARKET", type=["xlsx"])

if file_maxi and file_market:
    # Leggi i file Excel caricati
    try:
        maxi_df = pd.read_excel(file_maxi, header=0)
        market_df = pd.read_excel(file_market, header=0)
        
        # Rinomina le colonne con riferimenti stile Excel
        maxi_df.columns = ['A', 'B', 'C', 'D', 'E']
        market_df.columns = ['A', 'B', 'C', 'D', 'E']
        
        # Unisci i due file sulla colonna 'A' (che rappresenta il 'codice')
        comparison_df = pd.merge(maxi_df, market_df, on="A", suffixes=('_maxi', '_market'))
        
        # Identifica le differenze
        def identify_differences(row):
            differences = []
            if row['C_maxi'] != row['C_market']:  # Colonna C per Net
                differences.append('Net')
            if row['D_maxi'] != row['D_market']:  # Colonna D per Cessione
                differences.append('Cessione')
            if row['E_maxi'] != row['E_market']:  # Colonna E per Uscita
                differences.append('Uscita')
            return ", ".join(differences)
        
        comparison_df['Confronto'] = comparison_df.apply(identify_differences, axis=1)
        
        # Filtra solo le righe con differenze
        differences_df = comparison_df[comparison_df['Confronto'] != ""]
        
        # Seleziona le colonne necessarie per l'output
        result_df = differences_df[['A', 'B_maxi', 'C_maxi', 'C_market', 'D_maxi', 'D_market', 'E_maxi', 'E_market', 'Confronto']]
        
        # Rinomina le colonne per maggiore chiarezza
        result_df.columns = ['Codice', 'Descrizione', 'Net (Maxi)', 'Net (Market)', 'Cessione (Maxi)', 'Cessione (Market)', 'Uscita (Maxi)', 'Uscita (Market)', 'Confronto']
        
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
        
    except Exception as e:
        st.write("Errore durante la lettura dei file Excel:", e)

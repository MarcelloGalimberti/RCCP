# env neuralprophet

# 07/10/25
# per il deployment forzare python 3.11 in streamlit
#   RCCP 2026 Motore
# aggiornamento con i nuovi file

# import necessary libraries
import pandas as pd
import numpy as np
import warnings
#import matplotlib.pyplot as plt
import plotly.express as px
warnings.filterwarnings('ignore')
from io import BytesIO
import streamlit as st
import plotly.graph_objects as go
import calendar
import numpy as np
import plotly.colors

####### Impaginazione

st.set_page_config(layout="wide")

url_immagine = 'https://github.com/MarcelloGalimberti/Ducati_RCCP/blob/main/Ducati_red_logo-4.PNG?raw=true'

col_1, col_2 = st.columns([1, 4])
with col_1:
    st.image(url_immagine, width=200)

with col_2:
    st.title('RCCP | Motore')

st.header('Caricamento dati', divider='red')

####### Caricamento dati
# Caricamento del file CdC veicolo
st.header('Caricamento CdC Motore', divider='red')
uploaded_cdc_motore = st.file_uploader(
    "Carica file CdC motore 20xx.xlsx)",
)
if not uploaded_cdc_motore:
    st.stop()

df_famiglia_linea_mese_motore = pd.read_excel(uploaded_cdc_motore)#, parse_dates=True, index_col=0)
df_melt_motore = df_famiglia_linea_mese_motore.melt(id_vars='FAMIGLIA')
df_melt_motore.columns=['famiglia','anno-mese','linea']

# Unione dei DataFrame per veicolo e motore
#df_melt = pd.concat([df_melt_veicolo, df_melt_motore], ignore_index=True)
df_melt = df_melt_motore.copy()

# Rinomina le colonne per chiarezza
df_melt.columns = ['Famiglia', 'Anno-Mese', 'Linea']
# Anno-Mese in formato datetime anno - mese - giorno
df_melt['Anno-Mese'] = pd.to_datetime(df_melt['Anno-Mese'], format='%d/%m/%Y')
df_melt.rename(columns={'Anno-Mese': 'Data', 'Linea':'CDC'}, inplace=True)

# Filtra le linee di interesse per veicolo
# CdC Motore: 585, 590, 595
cdc_veicolo = [559,560,571,572,573,581,586,591,592]
cdc_motore = [585, 590, 595]    

# Filtra il DataFrame per le linee di interesse
df_melt = df_melt[df_melt['CDC'].isin(cdc_motore)]

# Caricamento del file abbinamento famiglia - modello
# nella verisione precedente: Modello Famiglia, nella versione attuale: Modello, Famiglia Veicolo,Famiglia Motore

st.header('Caricamento abbinamento famiglia | modello', divider='red')
uploaded_famiglia_modello = st.file_uploader(
    "Carica file abbinamento_modello_famiglia.xlsx)",
)
if not uploaded_famiglia_modello:
    st.stop()

# Caricamento del file di abbinamento modello-famiglia
df_modello_famiglia = pd.read_excel(uploaded_famiglia_modello)#, parse_dates=True, index_col=0)

# Caricamento del file PPP 2026
st.header('Caricamento PPP', divider='red')
uploaded_file = st.file_uploader(
    "Carica PPP 20xx motore",
)
if not uploaded_file:
    st.stop()


xls_PPP = pd.ExcelFile(uploaded_file)
fogli_PPP = xls_PPP.sheet_names # PR74 OP 2026 MaGa

fogli_menu_PPP = ['--- seleziona un foglio ---'] + fogli_PPP
foglio_selezionato_PPP = st.selectbox("Scegli il foglio da processare:", fogli_menu_PPP)

if foglio_selezionato_PPP == '--- seleziona un foglio ---':
    st.warning("Seleziona un foglio per continuare.")
    st.stop()



# Caricamento del file Calendario PPP 2026
st.header('Caricamento Calendario 20xx', divider='red')
uploaded_calendario = st.file_uploader(
    "Carica Calendario 20xx (Calendario_RCCP_20xx_turni.xlsx)",
)
if not uploaded_calendario:
    st.stop()

xls_calendario = pd.ExcelFile(uploaded_calendario) #
fogli_calendario = xls_calendario.sheet_names # DB risorse mensile

fogli_menu_calendario = ['--- seleziona un foglio ---'] + fogli_calendario
foglio_selezionato_calendario = st.selectbox("Scegli il foglio da processare:", fogli_menu_calendario)

if foglio_selezionato_calendario == '--- seleziona un foglio ---':
    st.warning("Seleziona un foglio per continuare.")
    st.stop()



# Caricamento del file PPP 2026, specificando il nome del foglio e le righe da saltare
df_PPP = pd.read_excel(uploaded_file, sheet_name=foglio_selezionato_PPP, skiprows=[1, 2, 3, 4, 5, 6, 7, 8, 9], header=0)

df_anno_PPP = pd.read_excel(uploaded_file, sheet_name=foglio_selezionato_PPP, nrows=1, header=None)


# prende da df_annp_PPP la seconda parola della prima riga e prima colonna
anno_PPP = str(df_anno_PPP.iat[0, 0]).split()[1]
#st.write('Anno PPP selezionato:', anno_PPP)

# Elimina le colonne non necessarie
df_PPP = df_PPP.drop(columns=['I°sem', 'II°sem', 'TOT'])
# Prendo il nome della prima colonna di ciascun DataFrame
colonna_chiave_ppp = df_PPP.columns[0]
colonna_chiave_modello = df_modello_famiglia.columns[0]
# Filtro df_PPP mantenendo solo i valori presenti in df_modello_famiglia
df_PPP = df_PPP[df_PPP[colonna_chiave_ppp].isin(df_modello_famiglia[colonna_chiave_modello])].fillna(0)
# Genera 12 date mensili a partire da gennaio 2026
date_nuove = pd.date_range(start=f"{anno_PPP}-01-01", periods=12, freq='MS').strftime('%d/%m/%Y')
# Costruisce la lista completa: prima colonna invariata + nuove date

df_PPP.columns = [df_PPP.columns[0]] + list(date_nuove)

# Rinomina la prima colonna in 'Modello'
df_PPP.rename(columns={df_PPP.columns[0]: 'Modello'}, inplace=True)


# Crea un dizionario Modello → Famiglia
mappa_famiglia = df_modello_famiglia.set_index('Modello')['Famiglia Motore']
# Aggiunge la colonna Famiglia a df_PPP usando la mappatura
df_PPP['Famiglia Motore'] = df_PPP['Modello'].map(mappa_famiglia)


# 1. Determina le colonne delle date (escludendo 'Modello' e 'Famiglia')
colonne_mensili = df_PPP.columns[1:-1]  # da colonna 1 a 12
# 2. Raggruppa per Famiglia e somma i valori mensili
df_sommato_per_famiglia = df_PPP.groupby('Famiglia Motore')[colonne_mensili].sum()
# Elimina le righe dove tutti i valori sono zero
df_sommato_per_famiglia = df_sommato_per_famiglia[(df_sommato_per_famiglia != 0).any(axis=1)]
#st.dataframe(df_sommato_per_famiglia, use_container_width=True)
df_unpivot = df_sommato_per_famiglia.reset_index().melt(
    id_vars='Famiglia Motore',
    var_name='Data',
    value_name='Qty'
)

df_unpivot['Data'] = pd.to_datetime(df_unpivot['Data'], format='%d/%m/%Y')
# st.write('PPP processato: Quantità per Famiglia e Data:')
# st.dataframe(df_unpivot, use_container_width=True)


# volumi per mese
df_volumi_per_mese = df_unpivot.groupby('Data')['Qty'].sum().reset_index()

# Grafico
# Crea la colonna per l'asse x con il formato "mese-anno" (in italiano)
mesi_it = [
    '','gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
    'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'
]
df_volumi_per_mese['Mese-Anno'] = df_volumi_per_mese['Data'].apply(
    lambda x: f"{mesi_it[x.month]}-{x.year}"
)

# Ordina per data
df_volumi_per_mese = df_volumi_per_mese.sort_values('Data')

st.header(f'Dati RCCP {anno_PPP}  motore', divider='red')
st.write('Volume totale:', df_volumi_per_mese['Qty'].sum())

# Grafico
fig = px.bar(
    df_volumi_per_mese,
    x='Mese-Anno',
    y='Qty',
    text='Qty',  # Mostra il valore sopra la barra
    title='Volumi mensili',
    labels={'Qty': 'Qty', 'Mese-Anno': 'Mese-Anno'},
    color_discrete_sequence=['#E32431']  # <-- Colore personalizzato
)

fig.update_traces(
    texttemplate='%{text:.0f}',  # Nessun decimale
    textposition='outside',      # Testo sopra la barra
    textfont_size=20             # Font grande per i valori sulle barre
)

fig.update_layout(
    xaxis_tickangle=45,
    yaxis_title="Qty",
    xaxis_title="Mese-Anno",
    font=dict(
        family="Arial, sans-serif",
        size=18,         # Font di base per tick e legenda
        color="black"
    ),
    title=dict(
        text='Volumi mensili',
        font=dict(size=26, family="Arial", color="black"),
        x=0.5,
        xanchor='center'
    ),
    xaxis=dict(
        title_font=dict(size=22),
        tickfont=dict(size=18)
    ),
    yaxis=dict(
        title_font=dict(size=22),
        tickfont=dict(size=18)
    ),
    legend=dict(
        font=dict(size=18),
        title_font=dict(size=20)
    ),
    height=700
)

st.plotly_chart(fig, use_container_width=True)

df_TC = pd.read_excel(uploaded_calendario, sheet_name=foglio_selezionato_calendario, parse_dates=True)
df_TC.rename(columns={'FAMIGLIA MOTORE': 'Famiglia Motore', 'MESE': 'Data'}, inplace=True)

# st.write('df_TC:')
# st.dataframe(df_TC, use_container_width=True)

# Merge con volumi
df_RCCP = pd.merge(
    df_TC,
    df_unpivot,
    left_on=['Famiglia Motore', 'Data'],
    right_on=['Famiglia Motore', 'Data'],
    how='left'
)


df_melt = df_melt.merge(df_modello_famiglia[['Famiglia Veicolo', 'Famiglia Motore']], left_on='Famiglia', right_on='Famiglia Veicolo', how='left')
df_melt.drop(columns=['Famiglia Veicolo'], inplace=True)
df_melt.drop_duplicates(inplace=True)
df_melt.reset_index(drop=True, inplace=True)


# Merge linee veicolo e motore con Qty
df_linee = pd.merge(
    df_melt,
    df_unpivot,
    left_on=['Famiglia Motore', 'Data'],
    right_on=['Famiglia Motore', 'Data'],
    how='left'
)


df_linee.drop(columns=['Famiglia'], inplace=True)
df_linee.drop_duplicates(inplace=True)
df_linee.reset_index(drop=True, inplace=True)


# Crea DataFrame cdc_motore con concatenazione delle famiglie motore
cdc_motore = (
    df_linee
    .dropna(subset=['Famiglia Motore'])
    .groupby(['Data', 'CDC'])['Famiglia Motore']
    .apply(lambda x: ', '.join(sorted(set(map(str, x)))))
    .reset_index()
    .rename(columns={'Famiglia Motore': 'Famiglie'})
)


df_aggregato = df_linee.groupby(['Data', 'CDC'])['Qty'].sum().reset_index()


# Filtra df_RCCP per mantenere solo le linee presenti in df_linee
df_RCCP_mask = df_RCCP['CDC'].isin(df_aggregato['CDC'])


#Fai un merge per estrarre i Qty corretti da df_linee
df_merge = df_RCCP[df_RCCP_mask].merge(
    df_aggregato[['Data', 'CDC', 'Qty']],
    on=['Data', 'CDC'],
    how='left',
    suffixes=('', '_from_linee')
)


# Fai un merge per unire i dati da df_merge su Data e CDC
df_temp = df_RCCP.merge(
    df_merge[['Data', 'CDC', 'Qty_from_linee']],
    on=['Data', 'CDC'],
    how='left'
)

# Sostituisci i valori di Qty solo dove Qty_from_line non è NaN
df_temp['Qty'] = df_temp['Qty_from_linee'].combine_first(df_temp['Qty'])

# Rimuovi la colonna temporanea
df_RCCP = df_temp.drop(columns='Qty_from_linee')


#########
# Calcolo del workload
df_RCCP['Workload'] = df_RCCP['Qty'] * df_RCCP['T.C.']/60/df_RCCP['OEE']

# st.write('df_RCCP con Workload:')
# st.dataframe(df_RCCP, use_container_width=True)


# Calcolo della saturazione
# Tabella pivot
pivot_df = pd.pivot_table(
    df_RCCP,
    index=['Data',  'RISORSA PRIMARIA','TURNO'], # eliminato 'REPARTO', 'CDC',
    values=["MOLTEPLICITA'", 'ore/mese', 'Workload'],
    aggfunc={
        "MOLTEPLICITA'": 'mean',
        'ore/mese': 'mean',
        'Workload': 'sum'
    }
).reset_index()

pivot_df = pivot_df.round(2)

# 1. Calcola la Capacity = MOLTEPLICITA' × ore/mese
pivot_df['Capacity'] = pivot_df["MOLTEPLICITA'"] * pivot_df['ore/mese']

# 2. Calcola la Saturazione = Workload / Capacity
# Usa np.where per evitare divisioni per zero

pivot_df['Saturazione'] = np.where(
    pivot_df['Capacity'] > 0,
    pivot_df['Workload'] / pivot_df['Capacity'],
    np.nan
)

# 3. Formatta la Saturazione come percentuale con un decimale
pivot_df['Saturazione'] = (pivot_df['Saturazione'] * 100).round(1).astype(str) + '%'

#pivot_df = pivot_df[pivot_df['CDC'].isin(cdc_motore)]

st.subheader('Tabella Saturazione per Data e Risorsa Primaria:', divider='red')
st.dataframe(pivot_df, use_container_width=True)    

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Foglio1')
    return output.getvalue()

# Crea il bottone per scaricare file Saturazione
saturazione_file = to_excel_bytes(pivot_df)
st.download_button(
    label="📥 Scarica file Excel saturazione",
    data=saturazione_file,
    file_name='df_saturazione.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)


#########

# Lavora su una copia del dataframe
df_heatmap = pivot_df.copy()

# Assicurati che 'Data' sia datetime
df_heatmap['Data'] = pd.to_datetime(df_heatmap['Data'])

# Crea una colonna con solo mese-anno (es: '2025-07')
df_heatmap['Mese-Anno'] = df_heatmap['Data'].dt.strftime('%Y-%m')

# Opzionale: se Saturazione contiene '%' o testo, converti in float
def parse_saturazione(v):
    if pd.isnull(v) or str(v).strip() == '':
        return 0
    return float(str(v).replace('%','').replace(',','.'))

df_heatmap['Saturazione_num'] = df_heatmap['Saturazione'].apply(parse_saturazione)

# Crea tabella pivot per heatmap
heatmap_data = df_heatmap.pivot_table(
    index='Mese-Anno',
    columns='RISORSA PRIMARIA',
    values='Saturazione_num',
    aggfunc='mean'  # oppure 'max' se preferisci
)

# Costruisci la heatmap con Plotly
fig = px.imshow(
    heatmap_data,
    text_auto=True,
    color_continuous_scale='RdYlGn_r',  # o altro schema colore
    aspect='auto',
    labels=dict(x="Risorsa primaria", y="Mese-Anno", color="Saturazione (%)"),
    title="Heatmap saturazione [%] per mese e risorsa primaria"
)

fig.update_layout(
    font=dict(size=16),
    height=900,
    xaxis_title="Risorsa primaria",
    yaxis_title="Mese-Anno"
)

st.plotly_chart(fig, use_container_width=True)


#########
# Step 0
# Sostituisci i valori None in df_RCCP['Famiglia Motore'] con i valori di cdc_motore['Famiglie']
df_RCCP = df_RCCP.merge(
    cdc_motore[['Data', 'CDC', 'Famiglie']], 
    on=['Data', 'CDC'], 
    how='left', 
    suffixes=('', '_cdc')
)

# Sostituisci i valori nella colonna 'Famiglia Motore' solo se CDC è in [585, 590, 595]
mask_cdc_motore = df_RCCP['CDC'].isin([585, 590, 595])

df_RCCP.loc[mask_cdc_motore, 'Famiglia Motore'] = df_RCCP.loc[mask_cdc_motore, 'Famiglie']

# Rimuovi la colonna temporanea
df_RCCP = df_RCCP.drop(columns='Famiglie')

# st.write('df_RCCP dopo sostituzione valori None per CDC motore [585, 590, 595]:')
# st.dataframe(df_RCCP, use_container_width=True)


# Step 1: Calcola Saturazione riga per riga
df_RCCP['Capacity'] = df_RCCP["MOLTEPLICITA'"] * df_RCCP['ore/mese']

# Evita divisioni per 0
df_RCCP['Saturazione'] = np.where(
    df_RCCP['Capacity'] > 0,
    df_RCCP['Workload'] / df_RCCP['Capacity'],
    np.nan
)




# Step 2: Crea pivot con Famiglia come colonne
pivot_famiglia = pd.pivot_table(
    df_RCCP,
    index=['Data','RISORSA PRIMARIA','TURNO'],  # eliminato 'REPARTO', 'CDC',
    columns='Famiglia Motore',
    values='Saturazione',
    aggfunc='mean'  # se ci sono più righe per combinazione
).reset_index()

# st.write('pivot_famiglia prima della formattazione:')
# st.dataframe(pivot_famiglia, use_container_width=True)


# Seleziona solo le colonne delle Famiglie (dopo reset_index)
fam_cols = pivot_famiglia.columns.difference(['Data', 'RISORSA PRIMARIA','TURNO']) 


# Moltiplica per 100 e formatta
pivot_famiglia[fam_cols] = pivot_famiglia[fam_cols] \
    .multiply(100) \
    .round(1) \
    .astype(str) + '%'



# 3. Funzione robusta per convertire le stringhe percentuali in float
def clean_pct(val):
    if isinstance(val, str):
        v = val.strip().replace('%','').replace(',','.')
        if v.lower() in ('nan',''):
            return np.nan
        try:
            return float(v)
        except Exception:
            return np.nan
    if pd.isna(val):
        return np.nan
    return val

# 4. Applica la funzione su ogni valore delle colonne Famiglia
df_numeric = pivot_famiglia[fam_cols].applymap(clean_pct)



# 5. Per ogni riga: se TUTTE le colonne famiglia sono nan, la saturazione sarà vuota, altrimenti la somma formattata
saturazione_valori = df_numeric.sum(axis=1, skipna=True)
tutti_nan = df_numeric.isna().all(axis=1)

# 6. Colonna finale
pivot_famiglia['Saturazione'] = [
    "" if is_nan else f"{val:.1f}%"
    for is_nan, val in zip(tutti_nan, saturazione_valori)
]

# 7. (Opzionale) Rimuovi visualmente 'nan%' dalle colonne famiglia
pivot_famiglia[fam_cols] = pivot_famiglia[fam_cols].replace('nan%', '', regex=False)

st.subheader('Tabella di Saturazione per Famiglia:', divider='red')
st.dataframe(pivot_famiglia, use_container_width=True)


# Crea il bottone per scaricare file Saturazione per Famiglia
saturazione_famiglia_file = to_excel_bytes(pivot_famiglia)
st.download_button(
    label="📥 Scarica file Excel saturazione per famiglia",
    data=saturazione_famiglia_file,
    file_name='df_saturazione_famiglia.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

############# Grafici 1

st.subheader('Analisi saturazione mensile:', divider='red')

# Funzione per il titolo data
def mese_anno(dt):
    # Se dt è un Timestamp, estrai mese e anno
    if isinstance(dt, pd.Timestamp):
        return f"{calendar.month_name[dt.month]} {dt.year}"
    # Se è una stringa, tenta conversione
    try:
        dt = pd.to_datetime(dt)
        return f"{calendar.month_name[dt.month]} {dt.year}"
    except Exception:
        return str(dt)


# Prepara la colonna x_label
pivot_famiglia['x_label'] = (
    pivot_famiglia['RISORSA PRIMARIA'].astype(str) + ' | ' +
    pivot_famiglia['TURNO'].astype(str)
)


# Trova le colonne delle famiglie
indici = ['Data', 'RISORSA PRIMARIA', 'Saturazione', 'x_label','TURNO']  # aggiunto TURNO
fam_cols = [c for c in pivot_famiglia.columns if c not in indici]


# Definisci una palette e mappa le famiglie ai colori
color_palette = plotly.colors.qualitative.Dark24 # Puoi scegliere anche altre, es: D3, Set3, Pastel Dark24 Plotly
fam_color_map = {fam: color_palette[i % len(color_palette)] for i, fam in enumerate(fam_cols)}


for dt in pivot_famiglia['Data'].unique():
    df_dt = pivot_famiglia[pivot_famiglia['Data'] == dt]
    fig = go.Figure()
    for fam in fam_cols:
        y_val = df_dt[fam].replace('nan%', '', regex=False).replace('', np.nan)
        y_val = y_val.apply(lambda v: float(str(v).replace('%','').replace(',','.')) if pd.notnull(v) and str(v).strip() != '' else 0)
        y_val = y_val.where(y_val > 0, 0)
        if y_val.sum() > 0:
            fig.add_bar(
                x=df_dt['x_label'],
                y=y_val,
                name=fam,
                marker_color=fam_color_map[fam]  # <-- qui assegni il colore
            )
    
    # 2. Calcola la somma totale delle barre impilate per ciascuna x_label
    stacked_sum = (
        df_dt[fam_cols]
        .replace('nan%', '', regex=False)
        .replace('', np.nan)
        .applymap(lambda v: float(str(v).replace('%','').replace(',','.')) if pd.notnull(v) and str(v).strip() != '' else 0)
        .sum(axis=1)
        .values
    )

    # 3. Aggiungi scatter con etichette testo sopra le barre
    fig.add_trace(
        go.Scatter(
            x=df_dt['x_label'],
            y=stacked_sum,
            text=[f"{v:.1f}%" if v > 0 else "" for v in stacked_sum],  # Puoi scegliere il formato
            mode="text",
            textposition="top center",
            showlegend=False,
            hoverinfo="skip"  # Così non si aggiunge una voce in leggenda
        )
    )

    
    # Aggiungi linea orizzontale al 100%
    fig.add_shape(
        type="line",
        x0=-0.5,
        y0=100,
        x1=len(df_dt['x_label']) - 0.5,
        y1=100,
        line=dict(color="red", width=2, dash="dash"),
    )

    fig.update_layout(
    barmode='stack',
    title=dict(
        text=f"Saturazione per risorsa e famiglia<br><sup>{mese_anno(dt)}</sup>",
        font=dict(size=24, family="Arial", color="black"),
        x=0.5,  # centra il titolo
        xanchor='center'
    ),
    xaxis_title="Risorsa primaria | Turni",
    yaxis_title="Saturazione (%)",
    legend_title="Famiglia",
    font=dict(
        family="Arial, sans-serif",
        size=18,      # font di base per tick e legenda
        color="black"
    ),
    legend=dict(
        font=dict(size=14), #18
        title_font=dict(size=20)
    ),
    yaxis=dict(
        ticksuffix='%',
        range=[0, max(120, df_dt[fam_cols].replace('nan%', '', regex=False)
                                 .replace('', np.nan)
                                 .applymap(lambda v: float(str(v).replace('%','').replace(',','.')) if pd.notnull(v) and str(v).strip() != '' else 0)
                                 .max().max() + 10)],
        title_font=dict(size=20),
        tickfont=dict(size=18)
    ),
    xaxis=dict(
        title_font=dict(size=20),
        tickfont=dict(size=18)
    ),
    xaxis_tickangle=45,
    height=900
)

    st.plotly_chart(fig, use_container_width=True)

st.stop()



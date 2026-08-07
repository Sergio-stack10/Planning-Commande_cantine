import streamlit as st
import pandas as pd
import numpy as np
import datetime
import io

# --- NETTOYAGE DU CACHE ---
st.cache_data.clear()

st.set_page_config(page_title="Planning & Commandes cantine", layout="wide")

# --- INJECTION CSS POUR LA CHARTRE GRAPHIQUE ---
custom_css = """
<style>
    /* 1. Réduire les espaces et la police */
    .stApp, .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    html, body, .stApp {
        font-size: 14px;
    }
    
    /* 2. Titre principal */
    h1 {
        color: #003D5B !important;
        font-weight: 600;
        padding-bottom: 10px;
        border-bottom: 3px solid #25E2CC;
    }
    
    /* 3. Menu latéral (Midnight) */
    section[data-testid="stSidebar"] {
        background-color: #002032;
    }
    section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] header {
        color: #FFFFFF !important;
    }
    
    /* 4. Onglets transformés en KPI Cards */
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #FFFFFF;
        color: #2A2B2C;
        border: 2px solid #F2F2F2;
        border-radius: 8px;
        padding: 12px 20px;
        /* Coin gauche tronqué */
        clip-path: polygon(15px 0%, 100% 0%, 100% 100%, 15px 100%, 0% 50%);
        font-weight: bold;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        border-color: #25E2CC;
        background-color: #E9FCFA;
    }
    /* Onglet sélectionné (Vert jade) */
    .stTabs [aria-selected="true"] {
        background-color: #007380 !important;
        color: #FFFFFF !important;
        border: 2px solid #007380 !important;
        box-shadow: 0 4px 12px rgba(0, 115, 128, 0.3);
    }
    .stTabs [data-baseweb="tab-highlight"] {
        background-color: transparent !important;
    }
    .stTabs [data-baseweb="tab-border-bottom"] {
        display: none;
    }
    
    /* 5. Boutons d'action (Forme Pilule / Formée) */
    div.stButton > button {
        background-color: #003D5B; /* Bleu roi */
        color: #FFFFFF;
        border: 2px solid #003D5B;
        padding: 10px 25px;
        border-radius: 25px; /* Forme pilule */
        font-weight: bold;
        box-shadow: 0 4px 8px rgba(0, 61, 91, 0.2);
        transition: all 0.3s ease;
    }
    div.stButton > button:hover {
        background-color: #FBCA18; /* Jaune soleil au survol */
        color: #002032; /* Texte Midnight */
        border-color: #FBCA18;
        transform: translateY(-2px); /* Léger soulèvement */
    }
    
    /* 6. Boutons de téléchargement (Turquoise) */
    .stDownloadButton > button {
        background-color: #25E2CC !important;
        color: #002032 !important;
        border: 2px solid #25E2CC !important;
        border-radius: 8px !important;
        font-weight: bold;
    }
    .stDownloadButton > button:hover {
        background-color: #007380 !important;
        color: #FFFFFF !important;
        border-color: #007380 !important;
    }
    
    /* 7. Alertes */
    .stAlert [data-testid="stAlertContent"] {
        border-left: 5px solid #25E2CC;
    }
    
    /* 8. Chiffres clés (st.metric) */
    [data-testid="stMetricValue"] {
        color: #007380;
        font-weight: bold;
    }
    /* 9. Signature fixée en bas de page */
    .footer-fix {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #FFFFFF; /* Fond blanc pour bien la voir */
        color: #2A2B2C;
        text-align: center;
        font-size: 12px;
        padding: 8px 0;
        z-index: 999;
        border-top: 1px solid #F2F2F2; /* Petite ligne de séparation */
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

st.title("📊 Planning & Commandes cantine")

st.sidebar.header("1. Importation des fichiers")
files_planning = st.sidebar.file_uploader("Fichiers Planning (1 ou 2)", type=['xlsx', 'xls', 'xlsb'], accept_multiple_files=True)
file_commande = st.sidebar.file_uploader("Fichier Commandes", type=['xlsx'])

st.sidebar.header("2. Paramètres d'absentéisme")
taux_absenteisme = st.sidebar.slider("Estimation de l'absentéisme (%)", 0, 30, 5)

jours = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche']

# --- INITIALISATION DE LA MEMOIRE (SESSION STATE) ---
if 'planning_data' not in st.session_state: st.session_state.planning_data = None
if 'commande_data' not in st.session_state: st.session_state.commande_data = None
if 'show_p1' not in st.session_state: st.session_state.show_p1 = False
if 'show_p2' not in st.session_state: st.session_state.show_p2 = False
if 'show_p3' not in st.session_state: st.session_state.show_p3 = False
if 'show_p4_menus' not in st.session_state: st.session_state.show_p4_menus = False
if 'show_p5' not in st.session_state: st.session_state.show_p5 = False
if 'page3_df' not in st.session_state: st.session_state.page3_df = None
if 'menus_df' not in st.session_state: st.session_state.menus_df = None
if 'anomalies_df' not in st.session_state: st.session_state.anomalies_df = None

# --- FONCTIONS UTILITAIRES ---

@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    processed_data = output.getvalue()
    return processed_data

def is_planned(val):
    if pd.isna(val) or isinstance(val, bool): return False
    if isinstance(val, (int, float, np.number)): return val > 0
    if isinstance(val, (datetime.time, datetime.datetime, pd.Timestamp)):
        t = val.time() if isinstance(val, (datetime.datetime, pd.Timestamp)) else val
        return t != datetime.time(0, 0, 0)
    val_str = str(val).strip()
    if val_str in ['', '*', 'nan', 'None', '0', '0:00', '00:00', '0:00:00', '00:00:00']: return False
    try:
        dt = pd.to_datetime(val_str, errors='coerce')
        if not pd.isna(dt): return dt.time() != datetime.time(0, 0, 0)
    except: pass
    try: return float(val_str) > 0
    except: pass
    if any(c.isalpha() for c in val_str): return False
    return False

def get_time_obj(val):
    if pd.isna(val) or str(val).strip() in ['', '*', 'nan']: return None
    if isinstance(val, datetime.time): return val
    if isinstance(val, (datetime.datetime, pd.Timestamp)): return val.time()
    if isinstance(val, (int, float, np.number)) and not isinstance(val, bool):
        if 0 < val < 1: 
            total_seconds = int(val * 86400)
            h = total_seconds // 3600
            m = (total_seconds % 3600) // 60
            s = total_seconds % 60
            return datetime.time(h, m, s)
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if not pd.isna(dt): return dt.time()
        except: pass
    val_str = str(val).strip()
    if val_str in ['0', '0:00', '00:00', '0:00:00', '00:00:00']: return None
    try:
        dt = pd.to_datetime(val_str, errors='coerce')
        if not pd.isna(dt): return dt.time()
    except: pass
    return None

def format_time_display(val):
    t = get_time_obj(val)
    if t: return t.strftime('%H:%M')
    return str(val).strip() if not pd.isna(val) and str(val).strip() not in ['nan'] else ""

def get_planning_status(de, a):
    if is_planned(de): return "Planifié"
    val_str = str(de).upper() if not pd.isna(de) else ""
    if 'CONGE' in val_str or 'CONGÉ' in val_str or 'MATERNITÉ' in val_str or 'MALADIE' in val_str: return "CONGE"
    if 'OFF' in val_str or 'LIBRE' in val_str or 'REPOS' in val_str: return "LIBRE"
    if 'DISPO' in val_str: return "DISPONIBILITE"
    if val_str in ['', '*', 'NAN', '0:00', '0:00:00']: return ""
    return val_str

def is_absence_command(cmd_val):
    if pd.isna(cmd_val): return False
    return 'JE NE SERAI PAS' in str(cmd_val).strip().upper()

# --- FONCTIONS DE TRAITEMENT ---

def parse_planning(files):
    all_planning = []
    for file in files:
        engine = 'pyxlsb' if file.name.endswith('.xlsb') else None
        xls = pd.ExcelFile(file, engine=engine)
        df = None
        if "Tout (WFO+WFH)" in xls.sheet_names:
            df = pd.read_excel(file, sheet_name="Tout (WFO+WFH)", header=None, skiprows=3, engine=engine)
            cols = [3, 4, 5, 6, 7, 10, 11, 12, 15, 16, 19, 20, 23, 24, 27, 28, 31, 32, 35, 36]
            new_cols = ['TRANSPORT', 'WORKDAY ID', 'Paid ID', 'Nom', 'Projet', 'Statut', 'Lundi_DE', 'Lundi_A', 'Mardi_DE', 'Mardi_A', 'Mercredi_DE', 'Mercredi_A', 'Jeudi_DE', 'Jeudi_A', 'Vendredi_DE', 'Vendredi_A', 'Samedi_DE', 'Samedi_A', 'Dimanche_DE', 'Dimanche_A']
            df = df.iloc[:, cols]
            df.columns = new_cols
        elif "TMM" in xls.sheet_names:
            df = pd.read_excel(file, sheet_name="TMM", header=None, skiprows=2, engine=engine)
            cols = [0, 4, 2, 5, 8, 10, 11, 12, 17, 18, 23, 24, 29, 30, 35, 36, 41, 42, 47, 48]
            new_cols = ['TRANSPORT', 'WORKDAY ID', 'Paid ID', 'Nom', 'Projet', 'Statut', 'Lundi_DE', 'Lundi_A', 'Mardi_DE', 'Mardi_A', 'Mercredi_DE', 'Mercredi_A', 'Jeudi_DE', 'Jeudi_A', 'Vendredi_DE', 'Vendredi_A', 'Samedi_DE', 'Samedi_A', 'Dimanche_DE', 'Dimanche_A']
            df = df.iloc[:, cols]
            df.columns = new_cols
        else: continue
        df['Paid ID'] = df['Paid ID'].astype(str).str.replace(" ", "").str.upper()
        df = df[df['Paid ID'].str.contains(r'\d', na=False)]
        for j in jours:
            df[f'{j}_Flag'] = df[f'{j}_DE'].apply(lambda x: 1 if is_planned(x) else 0)
        all_planning.append(df)
    if all_planning: return pd.concat(all_planning, ignore_index=True).drop_duplicates(subset=['Paid ID'])
    return pd.DataFrame()

def parse_commande(file):
    df = pd.read_excel(file)
    df = df.rename(columns={df.columns[0]: 'Paid ID'})
    if len(df.columns) >= 9:
        df = df.iloc[:, [0] + list(range(2, 9))]
        df.columns = ['Paid ID'] + jours
    else:
        df = df.iloc[:, [0] + list(range(1, 8))]
        df.columns = ['Paid ID'] + jours
    df['Paid ID'] = df['Paid ID'].astype(str).str.replace(" ", "").str.upper()
    df = df[df['Paid ID'].str.contains(r'[A-Z]-?\d', na=False)]
    return df

# --- AFFICHAGE DES ONGLETS (TOUJOURS VISIBLES) ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📄 1. Regroupement Planning", 
    "📈 2. Effectifs & Prévisions", 
    "⚠️ 3. Confrontation planning & commande", 
    "🍽️ 4. Commandes par menu",
    "❌ 5. Anomalies"
])

# --- PAGE 1 : REGROUPEMENT ---
with tab1:
    st.header("Regroupement des plannings")
    if st.button("🚀 Lancer l'import et le regroupement", key="btn_p1"):
        if files_planning and file_commande:
            with st.spinner("Traitement des fichiers en cours..."):
                st.session_state.planning_data = parse_planning(files_planning)
                st.session_state.commande_data = parse_commande(file_commande)
                st.session_state.show_p1 = True
                st.session_state.show_p2 = False
                st.session_state.show_p3 = False
                st.session_state.show_p4_menus = False
                st.session_state.show_p5 = False
            st.success("Données chargées avec succès !")
        else:
            st.error("Veuillez importer les fichiers dans le menu de gauche.")
            
    if st.session_state.show_p1 and st.session_state.planning_data is not None:
        st.markdown("---")
        
        display_planning = st.session_state.planning_data.copy()
        for j in jours:
            for suffix in ['_DE', '_A']:
                col = f'{j}{suffix}'
                display_planning[col] = display_planning[col].apply(format_time_display)
        
        # --- FILTRES (MULTISELECT) ---
        col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
        with col_f1:
            opts_trans = sorted(display_planning['TRANSPORT'].astype(str).unique().tolist())
            sel_trans = st.multiselect("Transport", opts_trans, default=[], key="f1_trans")
        with col_f2:
            opts_paid = sorted(display_planning['Paid ID'].astype(str).unique().tolist())
            sel_paid = st.multiselect("Paid ID", opts_paid, default=[], key="f1_paid")
        with col_f3:
            opts_nom = sorted(display_planning['Nom'].astype(str).unique().tolist())
            sel_nom = st.multiselect("Nom", opts_nom, default=[], key="f1_nom")
        with col_f4:
            opts_projet = sorted(display_planning['Projet'].astype(str).unique().tolist())
            sel_projet = st.multiselect("Projet", opts_projet, default=[], key="f1_projet")
        with col_f5:
            opts_statut = sorted(display_planning['Statut'].astype(str).unique().tolist())
            sel_statut = st.multiselect("Statut", opts_statut, default=[], key="f1_statut")
            
        # Application des filtres (si vide, aucun filtre appliqué)
        df_filtered = display_planning.copy()
        if sel_trans:
            df_filtered = df_filtered[df_filtered['TRANSPORT'].astype(str).isin(sel_trans)]
        if sel_paid:
            df_filtered = df_filtered[df_filtered['Paid ID'].astype(str).isin(sel_paid)]
        if sel_nom:
            df_filtered = df_filtered[df_filtered['Nom'].astype(str).isin(sel_nom)]
        if sel_projet:
            df_filtered = df_filtered[df_filtered['Projet'].astype(str).isin(sel_projet)]
        if sel_statut:
            df_filtered = df_filtered[df_filtered['Statut'].astype(str).isin(sel_statut)]
        
        cols_to_show = ['TRANSPORT', 'WORKDAY ID', 'Paid ID', 'Nom', 'Projet', 'Statut']
        for j in jours: cols_to_show += [f'{j}_DE', f'{j}_A', f'{j}_Flag']
        
        st.markdown("---")
        excel_data = to_excel(df_filtered[cols_to_show])
        st.download_button("📥 Télécharger le planning regroupé (Filtré)", data=excel_data, file_name="planning_regroupé.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_filtered[cols_to_show], use_container_width=True, height=600)

# --- PAGE 2 : EFFECTIFS ---
with tab2:
    st.header("Nombre de planifiés par projet et par jour")
    if st.button("📈 Calculer les effectifs", key="btn_p2"):
        if st.session_state.planning_data is not None:
            st.session_state.show_p2 = True
        else:
            st.error("Veuillez d'abord charger les données sur la Page 1.")
            
    if st.session_state.show_p2 and st.session_state.planning_data is not None:
        st.markdown("---")
        
        # --- FILTRE PROJET MULTIPLE ---
        opts_projet_p2 = sorted(st.session_state.planning_data['Projet'].astype(str).unique().tolist())
        sel_projet_p2 = st.multiselect("Filtrer par Projet", opts_projet_p2, default=[], key="f2_projet")
        
        # Filtrage des données avant le pivot (si vide, aucun filtre appliqué)
        planning_calc = st.session_state.planning_data.copy()
        if sel_projet_p2:
            planning_calc = planning_calc[planning_calc['Projet'].astype(str).isin(sel_projet_p2)]
        
        st.markdown("---")
        pivot_df = planning_calc.pivot_table(index='Projet', values=[f'{j}_Flag' for j in jours], aggfunc='sum', fill_value=0)
        pivot_df = pivot_df[[f'{j}_Flag' for j in jours]]
        pivot_df.columns = jours
        pivot_df.loc['Total Théorique'] = pivot_df.sum()
        pivot_df.loc[f'Total Estimé (-{taux_absenteisme}%)'] = (pivot_df.loc['Total Théorique'] * (1 - taux_absenteisme / 100)).round(0)
        
        excel_data = to_excel(pivot_df.reset_index())
        st.download_button("📥 Télécharger les effectifs (Filtré)", data=excel_data, file_name="effectifs.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(pivot_df.style.format("{:.0f}"), use_container_width=True)
        
        st.markdown("---")
        st.subheader(f"📦 Récapitulatif pour la commande (Estimé -{taux_absenteisme}%)")
        cols = st.columns(7)
        for i, j in enumerate(jours):
            with cols[i]:
                val = int(pivot_df.loc[f'Total Estimé (-{taux_absenteisme}%)', j])
                st.metric(label=j, value=f"{val} pax")

# --- PAGE 3 : CONFRONTATION ---
with tab3:
    st.header("Confrontation Planning & Commandes")
    if st.button("⚠️ Générer la confrontation", key="btn_p3"):
        if st.session_state.planning_data is not None and st.session_state.commande_data is not None:
            with st.spinner("Génération de la confrontation..."):
                merged = pd.merge(st.session_state.planning_data, st.session_state.commande_data, on='Paid ID', how='outer')
                display_rows = []
                for _, row in merged.iterrows():
                    has_planning = not pd.isna(row.get('Nom', np.nan))
                    display_row = {
                        'Paid ID': row['Paid ID'],
                        'WORKDAY ID': row['WORKDAY ID'] if has_planning else "",
                        'Nom': row['Nom'] if has_planning else "",
                        'Projet': row['Projet'] if has_planning else "",
                        'Statut': row['Statut'] if has_planning else ""
                    }
                    for j in jours:
                        de_col = f'{j}_DE'
                        a_col = f'{j}_A'
                        cmd_col = j
                        if has_planning and de_col in row:
                            planning_str = get_planning_status(row[de_col], row[a_col])
                        else:
                            planning_str = "hors planning" if cmd_col in row and not pd.isna(row[cmd_col]) and str(row[cmd_col]).strip() not in ['*', ''] else ""
                        commande_str = row[cmd_col] if cmd_col in row and not pd.isna(row[cmd_col]) else ""
                        if str(commande_str).strip() in ['*', '']: commande_str = ""
                        display_row[f'{j} - Planning'] = planning_str
                        display_row[f'{j} - Commande'] = commande_str
                    display_rows.append(display_row)
                st.session_state.page3_df = pd.DataFrame(display_rows)
                st.session_state.show_p3 = True
        else:
            st.error("Veuillez d'abord charger les données sur la Page 1.")
            
    if st.session_state.show_p3 and st.session_state.page3_df is not None:
        st.markdown("---")
        df_p3 = st.session_state.page3_df.copy()
        
        # --- FILTRES (MULTISELECT - SANS TRANSPORT) ---
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1:
            opts_paid = sorted(df_p3['Paid ID'].astype(str).unique().tolist())
            sel_paid = st.multiselect("Paid ID", opts_paid, default=[], key="f3_paid")
        with col_f2:
            opts_nom = sorted(df_p3['Nom'].astype(str).unique().tolist())
            sel_nom = st.multiselect("Nom", opts_nom, default=[], key="f3_nom")
        with col_f3:
            opts_projet = sorted(df_p3['Projet'].astype(str).unique().tolist())
            sel_projet = st.multiselect("Projet", opts_projet, default=[], key="f3_projet")
        with col_f4:
            opts_statut = sorted(df_p3['Statut'].astype(str).unique().tolist())
            sel_statut = st.multiselect("Statut", opts_statut, default=[], key="f3_statut")
            
        # Application des filtres (si vide, aucun filtre appliqué)
        df_filtered_p3 = df_p3.copy()
        if sel_paid:
            df_filtered_p3 = df_filtered_p3[df_filtered_p3['Paid ID'].astype(str).isin(sel_paid)]
        if sel_nom:
            df_filtered_p3 = df_filtered_p3[df_filtered_p3['Nom'].astype(str).isin(sel_nom)]
        if sel_projet:
            df_filtered_p3 = df_filtered_p3[df_filtered_p3['Projet'].astype(str).isin(sel_projet)]
        if sel_statut:
            df_filtered_p3 = df_filtered_p3[df_filtered_p3['Statut'].astype(str).isin(sel_statut)]
        
        st.markdown("---")
        excel_data = to_excel(df_filtered_p3)
        st.download_button("📥 Télécharger la confrontation (Filtré)", data=excel_data, file_name="confrontation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_filtered_p3, use_container_width=True, height=600)

# --- PAGE 4 : COMMANDES PAR MENU ---
with tab4:
    st.header("Nombre de commandes par menu et par jour")
    if st.button("🍽️ Calculer les commandes par menu", key="btn_p4_menus"):
        if st.session_state.commande_data is not None:
            with st.spinner("Calcul des menus en cours..."):
                cmd_data = st.session_state.commande_data.copy()
                cmd_melted = cmd_data.melt(id_vars=['Paid ID'], value_vars=jours, var_name='Jour', value_name='Menu')
                cmd_melted = cmd_melted.dropna(subset=['Menu'])
                cmd_melted['Menu'] = cmd_melted['Menu'].astype(str).str.strip()
                cmd_melted = cmd_melted[~cmd_melted['Menu'].str.upper().isin(['', '*', 'NAN', 'NONE', 'JE NE SERAI PAS PRÉSENT', 'JE NE SERAI PAS PRESENT'])]
                
                if not cmd_melted.empty:
                    pivot_menus = cmd_melted.pivot_table(index='Menu', columns='Jour', values='Paid ID', aggfunc='count', fill_value=0)
                    pivot_menus = pivot_menus.reindex(columns=jours, fill_value=0)
                    pivot_menus['Total Semaine'] = pivot_menus.sum(axis=1)
                    pivot_menus.loc['Total Commandes'] = pivot_menus.sum(axis=0)
                    st.session_state.menus_df = pivot_menus
                else:
                    st.session_state.menus_df = pd.DataFrame()
                st.session_state.show_p4_menus = True
        else:
            st.error("Veuillez d'abord charger les données sur la Page 1.")
            
    if st.session_state.show_p4_menus and st.session_state.menus_df is not None:
        st.markdown("---")
        excel_data = to_excel(st.session_state.menus_df.reset_index())
        st.download_button("📥 Télécharger les commandes par menu (Excel)", data=excel_data, file_name="menus_commandes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(st.session_state.menus_df.style.format("{:.0f}"), use_container_width=True, height=700)

# --- PAGE 5 : ANOMALIES ---
with tab5:
    st.header("Liste des anomalies (Planification vs Commande)")
    if st.button("❌ Extraire les anomalies", key="btn_p5"):
        if st.session_state.page3_df is not None:
            with st.spinner("Extraction des anomalies..."):
                anomalies = []
                for _, row in st.session_state.page3_df.iterrows():
                    for j in jours:
                        plan_col = f'{j} - Planning'
                        cmd_col = f'{j} - Commande'
                        if plan_col in row and cmd_col in row:
                            plan_val = str(row[plan_col]).strip()
                            cmd_val = str(row[cmd_col]).strip()
                            is_absence = is_absence_command(cmd_val)
                            
                            if plan_val == "Planifié" and (cmd_val == "" or is_absence):
                                anomalies.append({
                                    'Paid ID': row['Paid ID'], 'Nom': row['Nom'], 'Projet': row['Projet'],
                                    'Jour': j, "Type d'anomalie": "Planifié sans commande", 
                                    'Statut Planning': plan_val, 'Commande': cmd_val if cmd_val else "Aucune"
                                })
                            elif plan_val != "Planifié" and cmd_val != "" and not is_absence:
                                statut = plan_val if plan_val != "" else "Hors planning"
                                anomalies.append({
                                    'Paid ID': row['Paid ID'], 'Nom': row['Nom'], 'Projet': row['Projet'],
                                    'Jour': j, "Type d'anomalie": "Non planifié avec commande", 
                                    'Statut Planning': statut, 'Commande': cmd_val
                                })
                                
                st.session_state.anomalies_df = pd.DataFrame(anomalies)
                st.session_state.show_p5 = True
        else:
            st.error("Veuillez d'abord générer la confrontation sur la Page 3.")
            
    if st.session_state.show_p5:
        st.markdown("---")
        
        if st.session_state.anomalies_df.empty:
            st.success("✅ Aucune anomalie : tous les planifiés ont une commande et inversement.")
        else:
            df_anom = st.session_state.anomalies_df.copy()
            df_anom = df_anom.sort_values(by=["Type d'anomalie", "Jour", "Nom"])
            
            # --- FILTRES (MULTISELECT) ---
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                opts_paid = sorted(df_anom['Paid ID'].astype(str).unique().tolist())
                sel_paid = st.multiselect("Paid ID", opts_paid, default=[], key="f5_paid")
            with col_f2:
                opts_nom = sorted(df_anom['Nom'].astype(str).unique().tolist())
                sel_nom = st.multiselect("Nom", opts_nom, default=[], key="f5_nom")
            with col_f3:
                opts_projet = sorted(df_anom['Projet'].astype(str).unique().tolist())
                sel_projet = st.multiselect("Projet", opts_projet, default=[], key="f5_projet")
                
            # Application des filtres (si vide, aucun filtre appliqué)
            df_filtered_anom = df_anom.copy()
            if sel_paid:
                df_filtered_anom = df_filtered_anom[df_filtered_anom['Paid ID'].astype(str).isin(sel_paid)]
            if sel_nom:
                df_filtered_anom = df_filtered_anom[df_filtered_anom['Nom'].astype(str).isin(sel_nom)]
            if sel_projet:
                df_filtered_anom = df_filtered_anom[df_filtered_anom['Projet'].astype(str).isin(sel_projet)]
            
            st.markdown("---")
            excel_data = to_excel(df_filtered_anom)
            st.download_button("📥 Télécharger les anomalies (Filtré)", data=excel_data, file_name="anomalies.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(df_filtered_anom, use_container_width=True, height=600)
# --- SIGNATURE FIXEE EN BAS ---
st.markdown(
    "<div class='footer-fix'>"
    "Développé avec ❤️ par <b>[Ravo / Sergio]</b> • Concentrix Tamatave"
    "</div>", 
    unsafe_allow_html=True
)

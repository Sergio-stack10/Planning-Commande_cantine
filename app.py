import streamlit as st
import pandas as pd
import numpy as np
import datetime
import io
import os
import pickle

# --- NETTOYAGE DU CACHE ---
st.cache_data.clear()

st.set_page_config(page_title="LogiPlan", layout="wide")

# --- INJECTION CSS POUR LA CHARTRE GRAPHIQUE ---
custom_css = """
<style>
    .stApp, .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    html, body, .stApp { font-size: 14px; }
    h1 { color: #25E2CC !important; font-weight: 600; padding-bottom: 10px; border-bottom: 2px solid #003D5B; }
    section[data-testid="stSidebar"] { background-color: #002032; width: 260px !important; }
    section[data-testid="stSidebar"] > div:first-child { width: 260px !important; padding-top: 20px; }
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"], 
    section[data-testid="stSidebar"] label { font-size: 13px !important; color: #747474 !important; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #A8F3EB !important; font-size: 15px !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 15px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #FFFFFF; color: #2A2B2C; border: 2px solid #F2F2F2; border-radius: 8px;
        padding: 12px 20px; clip-path: polygon(15px 0%, 100% 0%, 100% 100%, 15px 100%, 0% 50%);
        font-weight: bold; box-shadow: 0 4px 6px rgba(0,0,0,0.05); transition: all 0.3s ease;
    }
    .stTabs [data-baseweb="tab"]:hover { border-color: #25E2CC; background-color: #E9FCFA; }
    .stTabs [aria-selected="true"] { background-color: #003D5B !important; color: #FFFFFF !important; box-shadow: 0 4px 12px rgba(0, 115, 128, 0.3); }
    .stTabs [data-baseweb="tab-highlight"] { background-color: transparent !important; }
    .stTabs [data-baseweb="tab-border-bottom"] { display: none; }
    div.stButton > button {
        background-color: #003D5B; color: #FFFFFF; border: 2px solid #003D5B; padding: 10px 25px;
        border-radius: 25px; font-weight: bold; box-shadow: 0 4px 8px rgba(0, 61, 91, 0.2); transition: all 0.3s ease;
    }
    div.stButton > button:hover { background-color: #FBCA18; color: #002032; border-color: #FBCA18; transform: translateY(-2px); }
    .stDownloadButton > button { background-color: #25E2CC !important; color: #002032 !important; border: 2px solid #25E2CC !important; border-radius: 8px !important; font-weight: bold; }
    .stDownloadButton > button:hover { background-color: #007380 !important; color: #FFFFFF !important; border-color: #007380 !important; }
    .stAlert [data-testid="stAlertContent"] { border-left: 5px solid #25E2CC; }
    [data-testid="stMetricValue"] { color: #007380; font-weight: bold; }
    .footer-fix {
        position: fixed !important; left: 0 !important; bottom: 0 !important; width: 100% !important;
        background-color: #002032 !important; color: #FFFFFF !important; text-align: left !important;
        font-size: 10px !important; padding: 5px 15px !important; z-index: 999999 !important; border-top: 1px solid #F2F2F2 !important;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

st.title("📊 LogiPlan")

st.sidebar.header("1. Importation des fichiers")
files_planning = st.sidebar.file_uploader("Fichiers Planning (Obligatoire)", type=['xlsx', 'xls', 'xlsb'], accept_multiple_files=True)
file_commande = st.sidebar.file_uploader("Fichier Commandes (Optionnel)", type=['xlsx'])

st.sidebar.header("2. Paramètres d'absentéisme")
taux_absenteisme = st.sidebar.slider("Estimation de l'absentéisme (%)", 0, 30, 5)

jours = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche']

# --- SYSTÈME D'HISTORIQUE PERSISTANT ---
HISTORY_FILE = "planning_history.pkl"

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "rb") as f:
                return pickle.load(f)
        except:
            return {}
    return {}

def save_history():
    try:
        with open(HISTORY_FILE, "wb") as f:
            pickle.dump({'plannings': st.session_state.history_plannings, 'commandes': st.session_state.history_commandes, 'calculs': st.session_state.history_calculs}, f)
    except Exception as e:
        st.error(f"Erreur lors de la sauvegarde : {e}")

if 'history_plannings' not in st.session_state:
    loaded = load_history()
    st.session_state.history_plannings = loaded.get('plannings', {})
    st.session_state.history_commandes = loaded.get('commandes', {})
    st.session_state.history_calculs = loaded.get('calculs', {})
    
if 'current_week' not in st.session_state:
    st.session_state.current_week = None

# Menu déroulant pour l'historique
st.sidebar.markdown("---")
st.sidebar.header("3. Historique des Semaines")
available_weeks = list(st.session_state.history_plannings.keys())
if available_weeks:
    available_weeks.sort()
    st.session_state.current_week = st.sidebar.selectbox("Semaine à afficher", available_weeks, key="week_selector")
    if st.sidebar.button("🗑️ Supprimer cette semaine"):
        del st.session_state.history_plannings[st.session_state.current_week]
        if st.session_state.current_week in st.session_state.history_commandes:
            del st.session_state.history_commandes[st.session_state.current_week]
        if st.session_state.current_week in st.session_state.history_calculs:
            del st.session_state.history_calculs[st.session_state.current_week]
        save_history()
        available_weeks = list(st.session_state.history_plannings.keys())
        st.session_state.current_week = available_weeks[0] if available_weeks else None
        st.rerun()
else:
    st.session_state.current_week = None
    st.sidebar.info("Aucune semaine chargée. Importez un fichier planning.")

# Fonctions utilitaires pour récupérer les données de la semaine sélectionnée
def get_current_planning():
    return st.session_state.history_plannings.get(st.session_state.current_week)

def get_current_commande():
    return st.session_state.history_commandes.get(st.session_state.current_week)

def get_calc(key):
    if st.session_state.current_week and st.session_state.current_week in st.session_state.history_calculs:
        return st.session_state.history_calculs[st.session_state.current_week].get(key)
    return None

def set_calc(key, df):
    if st.session_state.current_week:
        if st.session_state.current_week not in st.session_state.history_calculs:
            st.session_state.history_calculs[st.session_state.current_week] = {}
        st.session_state.history_calculs[st.session_state.current_week][key] = df

# --- FONCTIONS UTILITAIRES ---

@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

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

def get_pause_start(val):
    if pd.isna(val) or str(val).strip() in ['', '*', 'nan', 'None', '0', '0:00', '00:00', '0:00:00', '00:00:00']: return None
    if isinstance(val, datetime.time): return val
    if isinstance(val, (datetime.datetime, pd.Timestamp)): return val.time()
    val_str = str(val).strip()
    if '-' in val_str: val_str = val_str.split('-')[0].strip()
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

def calculate_slots(de, a, pause_start):
    if not de or not a: return []
    slots = []
    de_h = de.hour
    a_h = a.hour
    if a.minute > 0 or a.second > 0: a_h += 1
    if a <= de:
        for h in range(de_h, 24): slots.append((0, h))
        for h in range(0, a_h): slots.append((1, h))
    else:
        for h in range(de_h, a_h): slots.append((0, h))
    if pause_start:
        pause_h = pause_start.hour
        slots = [s for s in slots if s[1] != pause_h]
    else:
        fallback_h = (de_h + 4) % 24
        slots = [s for s in slots if s[1] != fallback_h]
    return slots

# --- FONCTIONS DE TRAITEMENT ---

def get_week_number(file, engine):
    try:
        xls = pd.ExcelFile(file, engine=engine)
        for sheet in xls.sheet_names:
            df_head = pd.read_excel(file, sheet_name=sheet, nrows=5, header=None, engine=engine)
            for i in range(min(5, len(df_head))):
                for j in range(min(15, len(df_head.columns))):
                    val = df_head.iloc[i, j]
                    if pd.notna(val):
                        dt = pd.to_datetime(val, errors='coerce')
                        if pd.isna(dt):
                            dt = pd.to_datetime(str(val), errors='coerce')
                        if not pd.isna(dt):
                            return f"S{dt.isocalendar().week:02d}"
    except:
        pass
    return None

def parse_planning(files, jours):
    all_planning = []
    for file in files:
        engine = 'pyxlsb' if file.name.endswith('.xlsb') else None
        xls = pd.ExcelFile(file, engine=engine)
        df = None
        if "Tout (WFO+WFH)" in xls.sheet_names:
            df = pd.read_excel(file, sheet_name="Tout (WFO+WFH)", header=None, skiprows=3, engine=engine)
            cols = [3, 4, 5, 6, 7, 10, 11, 12, 13, 15, 16, 17, 19, 20, 21, 23, 24, 25, 27, 28, 29, 31, 32, 33, 35, 36, 37]
            new_cols = ['TRANSPORT', 'WORKDAY ID', 'Paid ID', 'Nom', 'Projet', 'Statut', 
                        'Lundi_DE', 'Lundi_A', 'Lundi_Pause', 'Mardi_DE', 'Mardi_A', 'Mardi_Pause', 
                        'Mercredi_DE', 'Mercredi_A', 'Mercredi_Pause', 'Jeudi_DE', 'Jeudi_A', 'Jeudi_Pause', 
                        'Vendredi_DE', 'Vendredi_A', 'Vendredi_Pause', 'Samedi_DE', 'Samedi_A', 'Samedi_Pause', 
                        'Dimanche_DE', 'Dimanche_A', 'Dimanche_Pause']
            df = df.iloc[:, cols]
            df.columns = new_cols
            
        elif "TMM" in xls.sheet_names:
            df_head = pd.read_excel(file, sheet_name="TMM", header=None, nrows=10, engine=engine)
            header_row_idx = None
            trans_col_idx = 0
            for i in range(len(df_head)):
                row = df_head.iloc[i].astype(str).str.strip().tolist()
                if "Transport" in row:
                    header_row_idx = i
                    trans_col_idx = row.index("Transport")
                    break
            if header_row_idx is not None:
                df = pd.read_excel(file, sheet_name="TMM", header=None, skiprows=header_row_idx + 1, engine=engine)
                offset = trans_col_idx
                cols = [0 + offset, 4 + offset, 2 + offset, 5 + offset, 8 + offset, 10 + offset, 11 + offset, 12 + offset, 13 + offset, 17 + offset, 18 + offset, 19 + offset, 23 + offset, 24 + offset, 25 + offset, 29 + offset, 30 + offset, 31 + offset, 35 + offset, 36 + offset, 37 + offset, 41 + offset, 42 + offset, 43 + offset, 47 + offset, 48 + offset, 49 + offset]
                new_cols = ['TRANSPORT', 'WORKDAY ID', 'Paid ID', 'Nom', 'Projet', 'Statut', 
                            'Lundi_DE', 'Lundi_A', 'Lundi_Pause', 'Mardi_DE', 'Mardi_A', 'Mardi_Pause', 
                            'Mercredi_DE', 'Mercredi_A', 'Mercredi_Pause', 'Jeudi_DE', 'Jeudi_A', 'Jeudi_Pause', 
                            'Vendredi_DE', 'Vendredi_A', 'Vendredi_Pause', 'Samedi_DE', 'Samedi_A', 'Samedi_Pause', 
                            'Dimanche_DE', 'Dimanche_A', 'Dimanche_Pause']
                df = df.iloc[:, cols]
                df.columns = new_cols
            else:
                continue
        else: 
            continue
            
        df['WORKDAY ID'] = df['WORKDAY ID'].astype(str).str.replace(" ", "").str.replace(".0", "").str.upper()
        df['Paid ID'] = df['Paid ID'].astype(str).str.replace(" ", "").str.upper()
        df = df[df['WORKDAY ID'].str.contains(r'[A-Z0-9]', na=False)]
        df = df[~df['WORKDAY ID'].isin(['NAN', 'NONE', '*', ''])]
        for j in jours:
            df[f'{j}_Flag'] = df[f'{j}_DE'].apply(lambda x: 1 if is_planned(x) else 0)
        all_planning.append(df)
        
    if all_planning: 
        return pd.concat(all_planning, ignore_index=True).drop_duplicates(subset=['WORKDAY ID'])
    return pd.DataFrame()

def parse_commande(file, jours):
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

# --- AFFICHAGE DES ONGLETS ---
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📄 1. Regroupement Planning", 
    "📈 2. Effectifs & Prévisions", 
    "🚌 3. Planifiés par Shift", 
    "🕒 4. Planifiés par créneau", 
    "⚠️ 5. Confrontation planning & commande", 
    "🍽️ 6. Commandes par menu",
    "❌ 7. Anomalies"
])

current_planning = get_current_planning()

# --- PAGE 1 : REGROUPEMENT ---
with tab1:
    st.header("Regroupement des plannings")
    
    # Détection automatique du numéro de semaine pour pré-remplir le champ texte
    default_week_name = ""
    if files_planning:
        for f in files_planning:
            engine = 'pyxlsb' if f.name.endswith('.xlsb') else None
            wk = get_week_number(f, engine)
            if wk:
                default_week_name = wk
                break
                
    # Champ de saisie pour nommer la semaine
    week_name_input = st.text_input("Nom de la semaine à enregistrer", value=default_week_name, placeholder="Ex: S33, Semaine 34, etc.")
    
    if st.button("🚀 Lancer l'import et le regroupement", key="btn_p1"):
        if files_planning:
            # On utilise le nom saisi, ou on retombe sur la détection auto si vide
            week_num = week_name_input.strip() if week_name_input else default_week_name
            if not week_num:
                week_num = f"S{datetime.datetime.now().isocalendar().week:02d}"
                
            with st.spinner("Traitement des fichiers en cours..."):
                planning_df = parse_planning(files_planning, jours)
                st.session_state.history_plannings[week_num] = planning_df
                
                if file_commande:
                    cmd_df = parse_commande(file_commande, jours)
                    st.session_state.history_commandes[week_num] = cmd_df
                    
                save_history()
                st.session_state.current_week = week_num
            st.success(f"Semaine {week_num} chargée et sauvegardée avec succès !")
            st.rerun()
        else:
            st.error("Veuillez importer au moins un fichier de Planning dans le menu de gauche.")
            
    if st.session_state.current_week and current_planning is not None:
        st.markdown("---")
        display_planning = current_planning.copy()
        for j in jours:
            for suffix in ['_DE', '_A', '_Pause']:
                col = f'{j}{suffix}'
                if col in display_planning.columns:
                    display_planning[col] = display_planning[col].apply(format_time_display)
        
        col_f1, col_f2, col_f3, col_f4, col_f5, col_f6 = st.columns(6)
        with col_f1: sel_trans = st.multiselect("Transport", sorted(display_planning['TRANSPORT'].astype(str).unique().tolist()), key="f1_trans")
        with col_f2: sel_workday = st.multiselect("Workday ID", sorted(display_planning['WORKDAY ID'].astype(str).unique().tolist()), key="f1_workday")
        with col_f3: sel_paid = st.multiselect("Paid ID", sorted(display_planning['Paid ID'].astype(str).unique().tolist()), key="f1_paid")
        with col_f4: sel_nom = st.multiselect("Nom", sorted(display_planning['Nom'].astype(str).unique().tolist()), key="f1_nom")
        with col_f5: sel_projet = st.multiselect("Projet", sorted(display_planning['Projet'].astype(str).unique().tolist()), key="f1_projet")
        with col_f6: sel_statut = st.multiselect("Statut", sorted(display_planning['Statut'].astype(str).unique().tolist()), key="f1_statut")
            
        df_filtered = display_planning.copy()
        if sel_trans: df_filtered = df_filtered[df_filtered['TRANSPORT'].astype(str).isin(sel_trans)]
        if sel_workday: df_filtered = df_filtered[df_filtered['WORKDAY ID'].astype(str).isin(sel_workday)]
        if sel_paid: df_filtered = df_filtered[df_filtered['Paid ID'].astype(str).isin(sel_paid)]
        if sel_nom: df_filtered = df_filtered[df_filtered['Nom'].astype(str).isin(sel_nom)]
        if sel_projet: df_filtered = df_filtered[df_filtered['Projet'].astype(str).isin(sel_projet)]
        if sel_statut: df_filtered = df_filtered[df_filtered['Statut'].astype(str).isin(sel_statut)]
        
        cols_to_show = ['TRANSPORT', 'WORKDAY ID', 'Paid ID', 'Nom', 'Projet', 'Statut']
        for j in jours: cols_to_show += [f'{j}_DE', f'{j}_A', f'{j}_Pause', f'{j}_Flag']
        
        st.markdown("---")
        excel_data = to_excel(df_filtered[cols_to_show])
        st.download_button("📥 Télécharger le planning regroupé (Filtré)", data=excel_data, file_name="planning_regroupé.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_filtered[cols_to_show], use_container_width=True, height=600)
    elif not st.session_state.current_week:
        st.info("Veuillez importer un fichier pour commencer.")

# --- PAGE 2 : EFFECTIFS ---
with tab2:
    st.header("Nombre de planifiés par projet et par jour")
    if current_planning is not None:
        opts_projet_p2 = sorted(current_planning['Projet'].astype(str).unique().tolist())
        sel_projet_p2 = st.multiselect("Filtrer par Projet", opts_projet_p2, default=[], key="f2_projet")
        
        planning_calc = current_planning.copy()
        if sel_projet_p2: planning_calc = planning_calc[planning_calc['Projet'].astype(str).isin(sel_projet_p2)]
        
        st.markdown("---")
        pivot_df = planning_calc.pivot_table(index='Projet', values=[f'{j}_Flag' for j in jours], aggfunc='sum', fill_value=0)
        pivot_df = pivot_df[[f'{j}_Flag' for j in jours]]
        pivot_df.columns = jours
        pivot_df.loc['Total Théorique'] = pivot_df.sum()
        pivot_df.loc[f'Total Estimé (-{taux_absenteisme}%)'] = (pivot_df.loc['Total Théorique'] * (1 - taux_absenteisme / 100)).round(0)

        st.markdown("---")
        st.subheader("✍️ Saisie manuelle (Prestataires hors planning)")
        cols_in = st.columns(7)
        prestataires_vals = []
        for i, j in enumerate(jours):
            with cols_in[i]:
                val = st.number_input(f"{j}", min_value=0, step=1, value=0, key=f"prest_{j}")
                prestataires_vals.append(val)
                
        pivot_df.loc['Prestataires (Hors Planning)'] = prestataires_vals
        pivot_df.loc['Total à commander'] = pivot_df.loc[f'Total Estimé (-{taux_absenteisme}%)'] + pivot_df.loc['Prestataires (Hors Planning)']
        
        st.markdown("---")
        st.download_button("📥 Télécharger les effectifs (Filtré)", data=to_excel(pivot_df.reset_index()), file_name="effectifs.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(pivot_df.style.format("{:.0f}"), use_container_width=True)
        
        st.markdown("---")
        st.subheader(f"📦 Récapitulatif final à commander (Estimé + Prestataires)")
        cols = st.columns(7)
        for i, j in enumerate(jours):
            with cols[i]:
                val = int(pivot_df.loc['Total à commander', j])
                st.metric(label=j, value=f"{val} pax")
    else:
        st.warning("Aucune donnée disponible. Importez un planning.")

# --- PAGE 3 : PLANIFIES PAR SHIFT ---
with tab3:
    st.header("Planifiés par Shift (Début de journée)")
    if current_planning is not None:
        col_f1, col_f2 = st.columns(2)
        with col_f1: sel_trans_p3 = st.multiselect("Transport (OUI/NON)", sorted(current_planning['TRANSPORT'].astype(str).unique().tolist()), key="f3_trans")
        with col_f2: sel_projet_p3 = st.multiselect("Filtrer par Projet", sorted(current_planning['Projet'].astype(str).unique().tolist()), key="f3_projet")
        
        planning_shifts = current_planning.copy()
        if sel_trans_p3: planning_shifts = planning_shifts[planning_shifts['TRANSPORT'].astype(str).isin(sel_trans_p3)]
        if sel_projet_p3: planning_shifts = planning_shifts[planning_shifts['Projet'].astype(str).isin(sel_projet_p3)]
            
        st.markdown("---")
        with st.spinner("Calcul des shifts en cours..."):
            shift_rows = []
            for _, row in planning_shifts.iterrows():
                for j in jours:
                    de_col = f'{j}_DE'
                    if de_col in row:
                        de = get_time_obj(row[de_col])
                        if de:
                            shift_rows.append({
                                'Workday ID': row['WORKDAY ID'], 'Nom': row['Nom'], 'Projet': row['Projet'], 'Transport': row['TRANSPORT'],
                                'Jour': j, 'Shift (Début)': de.strftime('%H:%M')
                            })
            df_shifts = pd.DataFrame(shift_rows)
            if not df_shifts.empty:
                pivot_shifts = df_shifts.pivot_table(index='Shift (Début)', columns='Jour', values='Nom', aggfunc='count', fill_value=0)
                pivot_shifts = pivot_shifts.reindex(columns=jours, fill_value=0)
                pivot_shifts['Total Semaine'] = pivot_shifts.sum(axis=1)
                pivot_shifts.loc['Total'] = pivot_shifts.sum()
                st.markdown("#### 📊 Nombre de personnes par Heure de Début")
                st.download_button("📥 Télécharger le résumé par Shift", data=to_excel(pivot_shifts.reset_index()), file_name="resume_shifts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(pivot_shifts.style.format("{:.0f}"), use_container_width=True)
                
                st.markdown("---")
                with st.expander("👁️ Voir la liste détaillée des personnes par shift (Cliquez pour dérouler)"):
                    df_detail = df_shifts.sort_values(by=['Jour', 'Shift (Début)', 'Nom'])
                    st.download_button("📥 Télécharger la liste détaillée", data=to_excel(df_detail), file_name="detail_shifts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.dataframe(df_detail, use_container_width=True, height=500)
            else:
                st.warning("Aucun shift trouvé dans les données filtrées.")
    else:
        st.warning("Aucune donnée disponible. Importez un planning.")

# --- PAGE 4 : PLANIFIES PAR CRENEAU ---
with tab4:
    st.header("Nombre de planifiés par créneau horaire")
    if current_planning is not None:
        col_f1, col_f2 = st.columns(2)
        with col_f1: sel_workday_p4 = st.multiselect("Workday ID", sorted(current_planning['WORKDAY ID'].astype(str).unique().tolist()), key="f4_workday")
        with col_f2: sel_projet_p4 = st.multiselect("Filtrer par Projet", sorted(current_planning['Projet'].astype(str).unique().tolist()), key="f4_projet")
        
        planning_slots = current_planning.copy()
        if sel_workday_p4: planning_slots = planning_slots[planning_slots['WORKDAY ID'].astype(str).isin(sel_workday_p4)]
        if sel_projet_p4: planning_slots = planning_slots[planning_slots['Projet'].astype(str).isin(sel_projet_p4)]
            
        st.markdown("---")
        with st.spinner("Calcul des créneaux horaires en cours..."):
            hours = [f"{h:02d}:00" for h in range(24)]
            pivot_slots = pd.DataFrame(0, index=hours, columns=jours)
            for _, row in planning_slots.iterrows():
                for day_idx, j in enumerate(jours):
                    de_col = f'{j}_DE'; a_col = f'{j}_A'; pause_col = f'{j}_Pause'
                    if de_col in row and a_col in row:
                        de = get_time_obj(row[de_col]); a = get_time_obj(row[a_col]); pause = get_pause_start(row[pause_col]) if pause_col in row else None
                        slots = calculate_slots(de, a, pause)
                        for offset, hour in slots:
                            target_day = (day_idx + offset) % 7
                            pivot_slots.loc[f"{hour:02d}:00", jours[target_day]] += 1
            pivot_slots['Total Jour'] = pivot_slots.sum(axis=1)
            pivot_slots.loc['Total par Créneau'] = pivot_slots.sum(axis=0)
            
        st.download_button("📥 Télécharger les créneaux (Filtré)", data=to_excel(pivot_slots.reset_index()), file_name="creneaux_horaires.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(pivot_slots.style.format("{:.0f}"), use_container_width=True, height=700)
    else:
        st.warning("Aucune donnée disponible. Importez un planning.")

# --- PAGE 5 : CONFRONTATION ---
with tab5:
    st.header("↔️ Confrontation Planning & Commandes")
    current_commande = get_current_commande()
    
    if st.button("↔️ Générer la confrontation", key="btn_p5_conf"):
        if current_planning is not None:
            if current_commande is None and file_commande is not None:
                with st.spinner("Traitement du fichier Commandes..."):
                    current_commande = parse_commande(file_commande, jours)
                    st.session_state.history_commandes[st.session_state.current_week] = current_commande
                    save_history()
            if current_commande is not None:
                with st.spinner("Génération de la confrontation..."):
                    merged = pd.merge(current_planning, current_commande, on='Paid ID', how='outer')
                    display_rows = []
                    for _, row in merged.iterrows():
                        has_planning = not pd.isna(row.get('Nom', np.nan))
                        display_row = {
                            'Workday ID': row['WORKDAY ID'] if has_planning else "", 'Paid ID': row['Paid ID'],
                            'Nom': row['Nom'] if has_planning else "", 'Projet': row['Projet'] if has_planning else "",
                            'Statut': row['Statut'] if has_planning else ""
                        }
                        for j in jours:
                            de_col = f'{j}_DE'; a_col = f'{j}_A'; cmd_col = j
                            if has_planning and de_col in row:
                                planning_str = get_planning_status(row[de_col], row[a_col])
                            else:
                                planning_str = "hors planning" if cmd_col in row and not pd.isna(row[cmd_col]) and str(row[cmd_col]).strip() not in ['*', ''] else ""
                            commande_str = row[cmd_col] if cmd_col in row and not pd.isna(row[cmd_col]) else ""
                            if str(commande_str).strip() in ['*', '']: commande_str = ""
                            display_row[f'{j} - Planning'] = planning_str
                            display_row[f'{j} - Commande'] = commande_str
                        display_rows.append(display_row)
                    conf_df = pd.DataFrame(display_rows)
                    set_calc('conf', conf_df)
            else:
                st.error("Veuillez importer le fichier Commandes dans le menu de gauche.")
        else:
            st.error("Veuillez d'abord charger les données sur la Page 1.")

    conf_df = get_calc('conf')
    if conf_df is not None:
        st.markdown("---")
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1: sel_paid = st.multiselect("Paid ID", sorted(conf_df['Paid ID'].astype(str).unique().tolist()), key="f5_paid")
        with col_f2: sel_nom = st.multiselect("Nom", sorted(conf_df['Nom'].astype(str).unique().tolist()), key="f5_nom")
        with col_f3: sel_projet = st.multiselect("Projet", sorted(conf_df['Projet'].astype(str).unique().tolist()), key="f5_projet")
        with col_f4: sel_statut = st.multiselect("Statut", sorted(conf_df['Statut'].astype(str).unique().tolist()), key="f5_statut")
            
        df_filtered_p5 = conf_df.copy()
        if sel_paid: df_filtered_p5 = df_filtered_p5[df_filtered_p5['Paid ID'].astype(str).isin(sel_paid)]
        if sel_nom: df_filtered_p5 = df_filtered_p5[df_filtered_p5['Nom'].astype(str).isin(sel_nom)]
        if sel_projet: df_filtered_p5 = df_filtered_p5[df_filtered_p5['Projet'].astype(str).isin(sel_projet)]
        if sel_statut: df_filtered_p5 = df_filtered_p5[df_filtered_p5['Statut'].astype(str).isin(sel_statut)]
        
        st.markdown("---")
        st.download_button("📥 Télécharger la confrontation (Filtré)", data=to_excel(df_filtered_p5), file_name="confrontation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(df_filtered_p5, use_container_width=True, height=600)
    elif current_planning is None:
        st.warning("Aucune donnée disponible.")

# --- PAGE 6 : COMMANDES PAR MENU ---
with tab6:
    st.header("Nombre de commandes par menu et par jour")
    current_commande = get_current_commande()
    
    if st.button("🍽️ Calculer les commandes par menu", key="btn_p6_menus"):
        if current_commande is None and file_commande is not None:
            current_commande = parse_commande(file_commande, jours)
            st.session_state.history_commandes[st.session_state.current_week] = current_commande
            save_history()
        if current_commande is not None:
            with st.spinner("Calcul des menus en cours..."):
                cmd_melted = current_commande.melt(id_vars=['Paid ID'], value_vars=jours, var_name='Jour', value_name='Menu')
                cmd_melted = cmd_melted.dropna(subset=['Menu'])
                cmd_melted['Menu'] = cmd_melted['Menu'].astype(str).str.strip()
                cmd_melted = cmd_melted[~cmd_melted['Menu'].str.upper().isin(['', '*', 'NAN', 'NONE', 'JE NE SERAI PAS PRÉSENT', 'JE NE SERAI PAS PRESENT'])]
                if not cmd_melted.empty:
                    pivot_menus = cmd_melted.pivot_table(index='Menu', columns='Jour', values='Paid ID', aggfunc='count', fill_value=0)
                    pivot_menus = pivot_menus.reindex(columns=jours, fill_value=0)
                    pivot_menus['Total Semaine'] = pivot_menus.sum(axis=1)
                    pivot_menus.loc['Total Commandes'] = pivot_menus.sum(axis=0)
                    set_calc('menus', pivot_menus)
                else:
                    set_calc('menus', pd.DataFrame())
        else:
            st.error("Veuillez importer le fichier Commandes.")
            
    menus_df = get_calc('menus')
    if menus_df is not None:
        st.markdown("---")
        st.download_button("📥 Télécharger les commandes par menu (Excel)", data=to_excel(menus_df.reset_index()), file_name="menus_commandes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(menus_df.style.format("{:.0f}"), use_container_width=True, height=700)
    elif current_planning is None:
        st.warning("Aucune donnée disponible.")

# --- PAGE 7 : ANOMALIES ---
with tab7:
    st.header("⚠️ Liste des anomalies (Planification vs Commande)")
    conf_df = get_calc('conf')
    if st.button("⚠️ Extraire les anomalies", key="btn_p7_anom"):
        if conf_df is not None:
            with st.spinner("Extraction des anomalies..."):
                anomalies = []
                for _, row in conf_df.iterrows():
                    for j in jours:
                        plan_col = f'{j} - Planning'; cmd_col = f'{j} - Commande'
                        if plan_col in row and cmd_col in row:
                            plan_val = str(row[plan_col]).strip(); cmd_val = str(row[cmd_col]).strip()
                            is_absence = is_absence_command(cmd_val)
                            if plan_val == "Planifié" and (cmd_val == "" or is_absence):
                                anomalies.append({'Paid ID': row['Paid ID'], 'Nom': row['Nom'], 'Projet': row['Projet'], 'Jour': j, "Type d'anomalie": "Planifié sans commande", 'Statut Planning': plan_val, 'Commande': cmd_val if cmd_val else "Aucune"})
                            elif plan_val != "Planifié" and cmd_val != "" and not is_absence:
                                statut = plan_val if plan_val != "" else "Hors planning"
                                anomalies.append({'Paid ID': row['Paid ID'], 'Nom': row['Nom'], 'Projet': row['Projet'], 'Jour': j, "Type d'anomalie": "Non planifié avec commande", 'Statut Planning': statut, 'Commande': cmd_val})
                anom_df = pd.DataFrame(anomalies)
                set_calc('anom', anom_df)
        else:
            st.error("Veuillez d'abord générer la confrontation sur la Page 5.")
            
    anom_df = get_calc('anom')
    if anom_df is not None:
        st.markdown("---")
        if anom_df.empty:
            st.success("✅ Aucune anomalie.")
        else:
            df_anom = anom_df.copy()
            df_anom = df_anom.sort_values(by=["Type d'anomalie", "Jour", "Nom"])
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1: sel_paid = st.multiselect("Paid ID", sorted(df_anom['Paid ID'].astype(str).unique().tolist()), key="f7_paid")
            with col_f2: sel_nom = st.multiselect("Nom", sorted(df_anom['Nom'].astype(str).unique().tolist()), key="f7_nom")
            with col_f3: sel_projet = st.multiselect("Projet", sorted(df_anom['Projet'].astype(str).unique().tolist()), key="f7_projet")
            df_filtered_anom = df_anom.copy()
            if sel_paid: df_filtered_anom = df_filtered_anom[df_filtered_anom['Paid ID'].astype(str).isin(sel_paid)]
            if sel_nom: df_filtered_anom = df_filtered_anom[df_filtered_anom['Nom'].astype(str).isin(sel_nom)]
            if sel_projet: df_filtered_anom = df_filtered_anom[df_filtered_anom['Projet'].astype(str).isin(sel_projet)]
            st.markdown("---")
            st.download_button("📥 Télécharger les anomalies (Filtré)", data=to_excel(df_filtered_anom), file_name="anomalies.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(df_filtered_anom, use_container_width=True, height=600)
    elif current_planning is None:
        st.warning("Aucune donnée disponible.")

# --- SIGNATURE FIXEE EN BAS ---
st.markdown(
    "<div class='footer-fix'>Powered By <span style='color: #25E2CC; font-weight: 700; letter-spacing: 1px;'>RAVO SERGIO</span></div>", 
    unsafe_allow_html=True
)
import streamlit as st
import pandas as pd
import numpy as np
import datetime
import io
import os
import pickle
import re            # ═══ [MODIF A] requis par la page 6 ═══
import unicodedata   # ═══ [MODIF A] ═══

# --- NETTOYAGE DU CACHE ---
st.cache_data.clear()

st.set_page_config(page_title="LogiPlan", layout="wide")

# --- INJECTION CSS POUR LA CHARTRE GRAPHIQUE ---
custom_css = """
<style>
    html, body, .stApp {
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif !important;
    }
    .stApp, .block-container { 
        padding-top: 4rem !important; 
        padding-bottom: 2rem !important; 
    }
    h1 { 
        color: #25E2CC !important; 
        font-weight: 700 !important; 
        padding-bottom: 15px !important; 
        border-bottom: 3px solid #003D5B !important; 
        margin-bottom: 30px !important;
    }
    h2, h3 { color: #003D5B !important; font-weight: 600 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px !important; }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent !important;
        color: gray !important;
        border: 1px solid transparent !important;
        border-radius: 30px !important; 
        padding: 10px 25px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: none !important;
        clip-path: none !important; 
    }
    .stTabs [data-baseweb="tab"]:hover { 
        background-color: rgba(37, 226, 204, 0.1) !important; 
        color: #25E2CC !important;
    }
    .stTabs [aria-selected="true"] { 
        background-color: #003D5B !important; 
        color: #FFFFFF !important; 
        box-shadow: 0 4px 12px rgba(0, 61, 91, 0.3) !important;
    }
    .stTabs [data-baseweb="tab-highlight"] { background-color: transparent !important; }
    .stTabs [data-baseweb="tab-border-bottom"] { display: none !important; }
    div.stButton > button, .stDownloadButton > button {
        border-radius: 30px !important; font-weight: 600 !important; transition: all 0.3s ease !important;
        border: none !important;
    }
    div.stButton > button {
        background-color: #003D5B !important; color: #FFFFFF !important;
    }
    div.stButton > button:hover { 
        background-color: #25E2CC !important; color: #002032 !important; 
        transform: translateY(-3px);
        box-shadow: 0 8px 15px rgba(37, 226, 204, 0.3) !important; 
    }
    .stDownloadButton > button { 
        background-color: #25E2CC !important; color: #002032 !important;
    }
    .stDownloadButton > button:hover { 
        background-color: #007380 !important; color: #FFFFFF !important; transform: translateY(-3px);
    }
    [data-testid="stMetric"] {
        background-color: rgba(128, 128, 128, 0.05);
        border: 1px solid rgba(128, 128, 128, 0.1);
        border-radius: 12px;
        padding: 15px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    [data-testid="stMetricLabel"] { color: gray !important; font-size: 12px !important; text-transform: uppercase; letter-spacing: 1px; }
    [data-testid="stMetricValue"] { color: #003D5B !important; font-weight: 700 !important; font-size: 24px !important; }
    .footer-fix {
        position: fixed !important; left: 0 !important; bottom: 0 !important; width: 100% !important;
        background-color: #001a26 !important; color: #A8F3EB !important; text-align: center !important;
        font-size: 12px !important; padding: 10px !important; z-index: 999999 !important; 
        border-top: 2px solid #25E2CC !important;
    }
    .stDeployButton { display: none !important; }
    div[class*="_link_"], div[class*="_profilePreview_"], img[data-testid="appCreatorAvatar"], [data-testid="stLogo"] { display: none !important; }
    [data-testid="stHeaderActionElements"] a[href*="github.com"], [data-testid="stHeaderActionElements"] a[href*="streamlit.io"] { display: none !important; }
    [data-testid="stSidebarCollapseButton"] {
        opacity: 1 !important; background-color: #25E2CC !important; border: none !important; border-radius: 20px !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2) !important;
    }
    [data-testid="stSidebarCollapseButton"] svg { color: #002032 !important; fill: #002032 !important; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

st.title("📊 LogiPlan")

st.sidebar.header("1. Importation des fichiers")
files_planning = st.sidebar.file_uploader("Fichiers Planning (Obligatoire)", type=['xlsx', 'xls', 'xlsb'], accept_multiple_files=True)
file_commande = st.sidebar.file_uploader("Fichier Commandes (Optionnel)", type=['xlsx'])
file_reference = st.sidebar.file_uploader("Fichier Liste Actif (Optionnel)", type=['xlsx'])

st.sidebar.header("2. Paramètres d'absentéisme")
taux_absenteisme = st.sidebar.slider("Estimation de l'absentéisme (%)", 0, 30, 5)

jours = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche']

# ═══ [MODIF B] Validation des données : une semaine n'est exploitable que si
# c'est un DataFrame contenant les colonnes minimales du planning. ═══
COLONNES_OBLIGATOIRES = ['TRANSPORT', 'WORKDAY ID', 'Paid ID', 'Nom', 'Projet', 'Statut']

def planning_valide(df):
    return (isinstance(df, pd.DataFrame)
            and not df.empty
            and all(c in df.columns for c in COLONNES_OBLIGATOIRES))
# ═══ [FIN MODIF B] ═══

# --- SYSTÈME D'HISTORIQUE PERSISTANT ---
HISTORY_FILE = "planning_history_v2.pkl"   # ═══ [MODIF C] nouveau nom : l'ancien fichier corrompu n'est plus jamais lu ═══

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

st.sidebar.caption("LogiPlan v5 — Recap")   # ═══ [MODIF D] témoin de déploiement : doit apparaître en bas de la sidebar ═══

def get_current_planning():
    # ═══ [MODIF E] ne retourne la semaine que si les données sont exploitables ═══
    v = st.session_state.history_plannings.get(st.session_state.current_week)
    return v if planning_valide(v) else None
    # ═══ [FIN MODIF E] ═══

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

# ═══ [MODIF F-1] HELPERS DE LA PAGE 6 UNIQUEMENT (aucune autre page ne les utilise) ═══
ENTITES = ["PRESTA", "SUPPORT + SAI", "PROD / PLANIFIÉ PROD", "AUTRE / IGNORÉ"]
ENTITES_MAIN = ENTITES[:3]
LBL_SANS_CHOIX = "SANS CHOIX"
ENTITY_COLORS = {"PRESTA": "#4472C4", "SUPPORT + SAI": "#1F9AA8",
                 "PROD / PLANIFIÉ PROD": "#548235", "AUTRE / IGNORÉ": "#7F7F7F"}

# Règle 1 & 7 : les noms de jours / en-têtes ne sont jamais des menus
VALEURS_NON_MENU = {"LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE",
                    "SHIFT", "WKD", "MENU", "CHOIX", "NOMS", "NOM", "PRENOMS", "PRÉNOMS", "PRENOM",
                    "PRÉNOM", "PROJETS", "PROJET", "DEPARTEMENT", "DÉPARTEMENT", "DEPT", "CHECK",
                    "UNIQUE", "CODE", "MATRICULE", "VOTRE MATRICULE", "", "*", "NAN", "NONE", "0"}

def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn")

def get_prefix(mat):
    m = str(mat).strip().upper()
    return m[:2] if len(m) >= 2 else m

def is_absence_label(val):
    if pd.isna(val): return False
    s = strip_accents(str(val)).upper()
    return ("NE SERAI PAS" in s) or (s.strip() in ("ABSENT", "ABSENTE", "ABSENCE"))

def clean_menu_label(val):
    if pd.isna(val): return None
    s = " ".join(str(val).split())
    if s.upper() in VALEURS_NON_MENU: return None
    if is_absence_label(s): return None
    return s

def normalize_entity(raw):
    if raw is None or pd.isna(raw): return None
    s = strip_accents(str(raw)).upper().strip()
    if not s: return None
    if "PRESTA" in s: return "PRESTA"
    if "SUPPORT" in s or "SAI" in s: return "SUPPORT + SAI"
    if "PROD" in s: return "PROD / PLANIFIÉ PROD"
    return "AUTRE / IGNORÉ"

def build_entity_seed(cmd_df):
    """Table Préfixe -> Entité, proposée d'après la colonne « Departement » du fichier commande
    (équivalent de vos colonnes P:Q du Recap)."""
    seed = {}
    try:
        if cmd_df is not None and not cmd_df.empty:
            dep_col = next((c for c in cmd_df.columns
                            if "DEPARTEMENT" in strip_accents(str(c)).upper()
                            or strip_accents(str(c)).upper() in ("DEPT", "ENTITE", "SERVICE")), None)
            if dep_col:
                tmp = cmd_df[["Paid ID", dep_col]].dropna(subset=[dep_col]).copy()
                tmp["prefix"] = tmp["Paid ID"].astype(str).apply(get_prefix)
                tmp["ent"] = tmp[dep_col].apply(normalize_entity)
                tmp = tmp[tmp["ent"].notna()]
                if not tmp.empty:
                    seed = tmp.groupby("prefix")["ent"].agg(lambda x: x.mode().iloc[0]).to_dict()
    except Exception:
        pass
    seed["SA"] = "SUPPORT + SAI"   # SI(GAUCHE(matricule;2)="SA";"SUPPORT";...)
    return seed

def entity_badge_html(ent):
    c = ENTITY_COLORS.get(ent, "#7F7F7F")
    return (f"<div style='writing-mode:vertical-rl;transform:rotate(180deg);background:{c};color:#fff;"
            f"font-weight:700;font-size:13px;letter-spacing:1px;border-radius:10px;padding:16px 6px;"
            f"min-height:140px;display:flex;align-items:center;justify-content:center;'>{ent}</div>")

def style_recap_table(df):
    def hl(row):
        c = str(row["Choix"])
        if c == "TOTAL":
            return ["font-weight:800;background-color:rgba(0,61,91,0.12);"] * len(row)
        if c.startswith("dont «"):
            return ["font-style:italic;color:#8a8a8a;"] * len(row)
        if c == LBL_SANS_CHOIX:
            return ["font-weight:600;"] * len(row)
        return [""] * len(row)
    f_pct = lambda v: "" if pd.isna(v) else ("%.1f" % v).replace(".", ",") + " %"
    f_int = lambda v: "" if (pd.isna(v) or not isinstance(v, (int, float, np.integer, np.floating))) else f"{int(round(float(v)))}"
    return df.style.apply(hl, axis=1).format({"Nombres": f_int, "Pourcentage": f_pct, "A preparer": f_int})

def derive_week_dates(week_key):
    try:
        m = re.search(r"S(\d{1,2})", str(week_key).upper())
        if not m: return {}
        wk = int(m.group(1))
        if not 1 <= wk <= 53: return {}
        year = datetime.datetime.now().year
        for y in (year, year - 1, year + 1):
            try:
                monday = datetime.date.fromisocalendar(y, wk, 1)
                return {jours[i]: monday + datetime.timedelta(days=i) for i in range(7)}
            except ValueError:
                continue
    except Exception:
        pass
    return {}

def compute_recap_menus(planning_df, cmd_df, mapping, jours, taux_by_entity, taux_default):
    """Logique feuille Recap : entité par préfixe matricule ; Nombres = commandes par menu
    + SANS CHOIX (planifiés sans commande + « Je ne serai pas présent ») ;
    Pourcentage = part du total entité/jour ; A preparer = Nombres × (1 − taux)."""
    melted = pd.DataFrame()
    if cmd_df is not None and not cmd_df.empty:
        day_cols = [j for j in jours if j in cmd_df.columns]
        melted = cmd_df.melt(id_vars=["Paid ID"], value_vars=day_cols, var_name="Jour", value_name="Brut")
        melted = melted[melted["Brut"].notna()].copy()
        melted["Brut"] = melted["Brut"].astype(str).str.strip()
        melted = melted[melted["Brut"] != ""]
        melted["Entite"] = melted["Paid ID"].astype(str).apply(lambda x: mapping.get(get_prefix(x), "PROD / PLANIFIÉ PROD"))
        melted["Absence"] = melted["Brut"].apply(is_absence_label)
        melted["Menu"] = melted["Brut"].apply(clean_menu_label)

    planned_ids = {(j, e): set() for j in jours for e in ENTITES}
    if planning_df is not None and not planning_df.empty:
        p = planning_df.copy()
        p["Paid ID"] = p["Paid ID"].astype(str)
        p = p[~p["Paid ID"].isin(["", "NAN", "NONE", "*"])]
        p["Entite"] = p["Paid ID"].apply(lambda x: mapping.get(get_prefix(x), "PROD / PLANIFIÉ PROD"))
        for j in jours:
            if f"{j}_Flag" in p.columns:
                for ent, g in p[p[f"{j}_Flag"] == 1].groupby("Entite"):
                    planned_ids[(j, ent)] |= set(g["Paid ID"])

    menus_cnt, abs_cnt, ordered, day_menu_totals = {}, {}, {}, {}
    if not melted.empty:
        for (j, ent, mn), g in melted[melted["Menu"].notna()].groupby(["Jour", "Entite", "Menu"]):
            menus_cnt[(j, ent, mn)] = len(g)
            day_menu_totals.setdefault(j, {})
            day_menu_totals[j][mn] = day_menu_totals[j].get(mn, 0) + len(g)
        for (j, ent), g in melted[melted["Absence"]].groupby(["Jour", "Entite"]):
            abs_cnt[(j, ent)] = len(g)
        valid_resp = melted[melted["Menu"].notna() | melted["Absence"]]
        for (j, ent), g in valid_resp.groupby(["Jour", "Entite"]):
            ordered[(j, ent)] = set(g["Paid ID"].astype(str))

    recap, day_totals = {}, {}
    for j in jours:
        menu_list = [m for m, _ in sorted(day_menu_totals.get(j, {}).items(), key=lambda kv: (-kv[1], kv[0]))]
        for ent in ENTITES:
            facteur = 1.0 - (taux_by_entity.get(ent, taux_default) / 100.0)
            pids, ords = planned_ids.get((j, ent), set()), ordered.get((j, ent), set())
            abs_n = abs_cnt.get((j, ent), 0)
            sans_choix = len(pids - ords) + abs_n
            total_n = sum(menus_cnt.get((j, ent, m), 0) for m in menu_list) + sans_choix
            rows = []
            for m in menu_list:
                n = menus_cnt.get((j, ent, m), 0)
                rows.append({"Choix": m, "Nombres": n,
                             "Pourcentage": (n / total_n * 100) if total_n else 0.0,
                             "A preparer": int(n * facteur + 0.5)})
            if abs_n:
                rows.append({"Choix": "dont « Je ne serai pas présent » (inclus dans SANS CHOIX)",
                             "Nombres": abs_n,
                             "Pourcentage": (abs_n / total_n * 100) if total_n else 0.0, "A preparer": ""})
            rows.append({"Choix": LBL_SANS_CHOIX, "Nombres": sans_choix,
                         "Pourcentage": (sans_choix / total_n * 100) if total_n else 0.0,
                         "A preparer": int(sans_choix * facteur + 0.5)})
            total_prep = sum(r["A preparer"] for r in rows if isinstance(r["A preparer"], (int, np.integer)))
            rows.append({"Choix": "TOTAL", "Nombres": total_n,
                         "Pourcentage": 100.0 if total_n else 0.0, "A preparer": total_prep})
            recap[(j, ent)] = {"df": pd.DataFrame(rows), "planned_n": len(pids), "reponses": len(ords),
                               "abs_n": abs_n, "sans_choix": sans_choix, "total_n": total_n,
                               "total_prep": total_prep}
        day_totals[j] = sum(recap[(j, e)]["total_prep"] for e in ENTITES_MAIN)
    return recap, day_totals
# ═══ [FIN MODIF F-1] ═══

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
    # ═══ [MODIF F-2] version robuste : détecte automatiquement la ligne d'en-tête
    # (le modèle peut avoir « Votre matricule » seul sur la 1re ligne, puis
    # « Exemple: W00123 | Shift | Lundi | ... » sur la 2e) et exclut la ligne d'exemple. ═══
    raw = pd.read_excel(file, header=None, nrows=6)
    header_row = 0
    for i in range(len(raw)):
        vals = [str(v).strip().upper() for v in raw.iloc[i].tolist()]
        if any("LUNDI" in v for v in vals):
            header_row = i
            break
    df = pd.read_excel(file, header=header_row)
    df = df.rename(columns={df.columns[0]: 'Paid ID'})
    day_idx = list(range(2, 9)) if len(df.columns) >= 9 else list(range(1, 8))
    # Colonnes complémentaires détectées par en-tête (Noms, Projets, Departement...)
    extras, used = [], set()
    for idx, c in enumerate(df.columns):
        if idx == 0 or idx in day_idx:
            continue
        cu = strip_accents(str(c)).upper().strip()
        target = None
        if cu in ("NOMS", "NOM", "PRENOMS", "PRÉNOMS", "PRENOM", "PRÉNOM"): target = "CMD_Nom"
        elif cu in ("PROJETS", "PROJET"): target = "CMD_Projet"
        elif "DEPARTEMENT" in cu or cu in ("DEPT", "ENTITE", "ENTITÉ", "SERVICE"): target = "CMD_Departement"
        if target and target not in used:
            extras.append((idx, target)); used.add(target)
    out = df.iloc[:, [0] + day_idx + [i for i, _ in extras]].copy()
    out.columns = ["Paid ID"] + jours + [t for _, t in extras]
    out["Paid ID"] = out["Paid ID"].astype(str).str.replace(" ", "").str.upper()
    out = out[out["Paid ID"].str.contains(r'[A-Z]-?\d', na=False)]
    # Règle 1 : la ligne d'exemple « Exemple: W00123 » du modèle n'est jamais comptée
    out = out[~out["Paid ID"].str.contains("EXEMPLE|VOTRE|MATRICULE", na=False)]
    return out
    # ═══ [FIN MODIF F-2] ═══

def parse_reference(file):
    """Lit le fichier Liste Actif et retourne un mapping Workday ID -> Paid ID"""
    try:
        df = pd.read_excel(file)
        cols_cleaned = []
        for c in df.columns:
            c_str = str(c).upper()
            c_str = "".join(c_str.split())
            cols_cleaned.append(c_str)
        df.columns = cols_cleaned
    except Exception as e:
        return None

    wd_col = None
    pd_col = None
    for c in df.columns:
        if 'WORKDAY' in c or 'EMPLOYEEID' in c: wd_col = c
        if 'PAYROLLID' in c or 'PAIDID' in c or 'MATRICULEPAIE' in c: pd_col = c

    if wd_col is None or pd_col is None:
        return {'error': True, 'columns': list(df.columns)}

    df = df[[wd_col, pd_col]].copy()
    df[wd_col] = df[wd_col].astype(str).str.replace(" ", "").str.replace(".0", "").str.upper()
    df[pd_col] = df[pd_col].astype(str).str.replace(" ", "").str.replace(".0", "").str.upper()
    df = df.dropna(subset=[wd_col])
    df = df[df[wd_col].str.contains(r'[A-Z0-9]', na=False)]
    df = df[~df[wd_col].isin(['NAN', 'NONE', '*', ''])]
    df = df.drop_duplicates(subset=[wd_col])
    return df.rename(columns={wd_col: 'WORKDAY ID', pd_col: 'REF_PAID_ID'})

# --- GESTION DU FICHIER DE RÉFÉRENCE ---
if 'reference_data' not in st.session_state:
    st.session_state.reference_data = None

if file_reference is not None:
    ref_parsed = parse_reference(file_reference)
    if isinstance(ref_parsed, dict) and ref_parsed.get('error'):
        st.session_state.reference_data = None
        st.session_state.ref_error = ref_parsed.get('columns', [])
    elif ref_parsed is not None:
        st.session_state.reference_data = ref_parsed
        st.session_state.ref_error = None

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

# --- PAGE 1 : REGROUPEMENT ---   (INCHANGÉ, sauf validation avant sauvegarde)
with tab1:
    st.header("Regroupement des plannings")
    
    default_week_name = ""
    if files_planning:
        for f in files_planning:
            engine = 'pyxlsb' if f.name.endswith('.xlsb') else None
            wk = get_week_number(f, engine)
            if wk:
                default_week_name = wk
                break
                
    week_name_input = st.text_input("Nom de la semaine à enregistrer", value=default_week_name, placeholder="Ex: S33, Semaine 34, etc.")
    
    if st.button("🚀 Lancer l'import et le regroupement", key="btn_p1"):
        if files_planning:
            week_num = week_name_input.strip() if week_name_input else default_week_name
            if not week_num:
                week_num = f"S{datetime.datetime.now().isocalendar().week:02d}"
            with st.spinner("Traitement des fichiers en cours..."):
                planning_df = parse_planning(files_planning, jours)
                # ═══ [MODIF B-suite] on ne sauvegarde que si le planning est exploitable ═══
                if planning_valide(planning_df):
                    st.session_state.history_plannings[week_num] = planning_df
                    if file_commande:
                        try:
                            st.session_state.history_commandes[week_num] = parse_commande(file_commande, jours)
                        except Exception as e:
                            st.warning(f"Fichier Commandes ignoré ({e}) — l'import du planning est conservé.")
                    save_history()
                    st.session_state.current_week = week_num
                    st.success(f"Semaine {week_num} chargée et sauvegardée avec succès !")
                    st.rerun()
                else:
                    st.error("❌ Aucun planning exploitable reconnu (feuille « Tout (WFO+WFH) » ou « TMM » attendue). Rien n'a été enregistré.")
                # ═══ [FIN MODIF B-suite] ═══
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

# --- PAGE 2 : EFFECTIFS ---   (STRICTEMENT IDENTIQUE À L'ORIGINAL)
with tab2:
    st.header("Nombre de planifiés par projet et par jour")
    if current_planning is not None:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            opts_projet_p2 = sorted(current_planning['Projet'].astype(str).unique().tolist())
            sel_projet_p2 = st.multiselect("Filtrer par Projet", opts_projet_p2, default=[], key="f2_projet")
        with col_f2:
            opts_statut_p2 = sorted(current_planning['Statut'].astype(str).unique().tolist())
            sel_statut_p2 = st.multiselect("Filtrer par Statut", opts_statut_p2, default=[], key="f2_statut")
        
        planning_calc = current_planning.copy()
        if sel_projet_p2: planning_calc = planning_calc[planning_calc['Projet'].astype(str).isin(sel_projet_p2)]
        if sel_statut_p2: planning_calc = planning_calc[planning_calc['Statut'].astype(str).isin(sel_statut_p2)]
        
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

# --- PAGE 3 : PLANIFIES PAR SHIFT ---   (STRICTEMENT IDENTIQUE À L'ORIGINAL)
with tab3:
    st.header("Planifiés par Shift (Début de journée)")
    if current_planning is not None:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1: sel_trans_p3 = st.multiselect("Transport (OUI/NON)", sorted(current_planning['TRANSPORT'].astype(str).unique().tolist()), key="f3_trans")
        with col_f2: sel_projet_p3 = st.multiselect("Filtrer par Projet", sorted(current_planning['Projet'].astype(str).unique().tolist()), key="f3_projet")
        with col_f3: sel_statut_p3 = st.multiselect("Filtrer par Statut", sorted(current_planning['Statut'].astype(str).unique().tolist()), key="f3_statut")
        
        planning_shifts = current_planning.copy()
        if sel_trans_p3: planning_shifts = planning_shifts[planning_shifts['TRANSPORT'].astype(str).isin(sel_trans_p3)]
        if sel_projet_p3: planning_shifts = planning_shifts[planning_shifts['Projet'].astype(str).isin(sel_projet_p3)]
        if sel_statut_p3: planning_shifts = planning_shifts[planning_shifts['Statut'].astype(str).isin(sel_statut_p3)]
            
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

# --- PAGE 4 : PLANIFIES PAR CRENEAU ---   (STRICTEMENT IDENTIQUE À L'ORIGINAL)
with tab4:
    st.header("Prevpoz et staffing par créneaux horaires")
    if current_planning is not None:
        col_f1, col_f2 = st.columns(2)
        with col_f1: sel_projet_p4 = st.multiselect("Filtrer par Projet", sorted(current_planning['Projet'].astype(str).unique().tolist()), key="f4_projet")
        with col_f2: sel_statut_p4 = st.multiselect("Filtrer par Statut", sorted(current_planning['Statut'].astype(str).unique().tolist()), key="f4_statut")
        
        planning_slots = current_planning.copy()
        if sel_projet_p4: planning_slots = planning_slots[planning_slots['Projet'].astype(str).isin(sel_projet_p4)]
        if sel_statut_p4: planning_slots = planning_slots[planning_slots['Statut'].astype(str).isin(sel_statut_p4)]
            
        st.markdown("---")
        with st.spinner("Calcul des créneaux horaires en cours..."):
            hours = [f"{h:02d}:00" for h in range(24)]
            pivot_slots = pd.DataFrame(0, index=hours, columns=jours)
            project_hourly = {}
            for _, row in planning_slots.iterrows():
                projet = row['Projet']
                if projet not in project_hourly:
                    project_hourly[projet] = {j: {h: 0 for h in range(24)} for j in jours}
                for day_idx, j in enumerate(jours):
                    de_col = f'{j}_DE'; a_col = f'{j}_A'; pause_col = f'{j}_Pause'
                    if de_col in row and a_col in row:
                        de = get_time_obj(row[de_col]); a = get_time_obj(row[a_col]); pause = get_pause_start(row[pause_col]) if pause_col in row else None
                        slots = calculate_slots(de, a, pause)
                        for offset, hour in slots:
                            target_day = (day_idx + offset) % 7
                            target_day_name = jours[target_day]
                            pivot_slots.loc[f"{hour:02d}:00", target_day_name] += 1
                            project_hourly[projet][target_day_name][hour] += 1
                            
            pivot_slots['Total Jour'] = pivot_slots.sum(axis=1)
            pivot_slots.loc['Total par Créneau'] = pivot_slots.sum(axis=0)
            
            peak_data = []
            for proj, day_data in project_hourly.items():
                row_data = {'Projet': proj}
                for j in jours:
                    row_data[j] = max(day_data[j].values()) if day_data[j].values() else 0
                peak_data.append(row_data)
            df_peaks = pd.DataFrame(peak_data).set_index('Projet')
            global_peaks = df_peaks[jours].sum().to_dict()
            df_peaks.loc['Pic Global (Tous Projets)'] = global_peaks
            
        st.markdown("#### 📊 Prevpoz et staffing par créneaux horaires")
        st.write("Ce tableau indique le nombre maximum de personnes présentes simultanément (en overlapping de shifts).")
        st.dataframe(df_peaks.style.format("{:.0f}"), use_container_width=True)
        excel_peaks = to_excel(df_peaks.reset_index())
        st.download_button("📥 Télécharger le Pic de présence", data=excel_peaks, file_name="pic_presence.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown("---")
        with st.expander("🕒 Voir le détail complet par créneau horaire"):
            excel_data = to_excel(pivot_slots.reset_index())
            st.download_button("📥 Télécharger les créneaux détaillés", data=excel_data, file_name="creneaux_horaires.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(pivot_slots.style.format("{:.0f}"), use_container_width=True, height=700)
    else:
        st.warning("Aucune donnée disponible. Importez un planning.")

# --- PAGE 5 : CONFRONTATION ---   (STRICTEMENT IDENTIQUE À L'ORIGINAL)
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

# ═══ [MODIF F-3] PAGE 6 : COMMANDES PAR MENU — SEULE PAGE REFAITE (feuille « Recap ») ═══
with tab6:
    st.header("Nombre de commandes par menu et par jour")
    st.caption("Logique « Recap » : entité = préfixe du matricule (« SA » → SUPPORT + SAI, sinon table de correspondance) · "
               "SANS CHOIX = planifiés sans commande + « Je ne serai pas présent » · A preparer = Nombres × (1 − absentéisme).")
    current_commande = get_current_commande()

    # Re-lecture du fichier Commandes si la semaine n'en contient pas encore
    if (current_commande is None or (isinstance(current_commande, pd.DataFrame) and current_commande.empty)) \
            and file_commande is not None and st.session_state.current_week:
        try:
            current_commande = parse_commande(file_commande, jours)
            st.session_state.history_commandes[st.session_state.current_week] = current_commande
            save_history()
        except Exception as e:
            st.error(f"Fichier Commandes illisible : {e}")
            current_commande = None

    if current_commande is None or (isinstance(current_commande, pd.DataFrame) and current_commande.empty):
        st.warning("Aucune commande disponible. Importez le fichier Commandes dans le menu de gauche, puis rechargez la semaine sur la Page 1.")
    else:
        # 1) Table Préfixe -> Entité (équivalent RECHERCHEX de votre Recap)
        ids = set(current_commande["Paid ID"].astype(str))
        if current_planning is not None:
            ids |= set(current_planning["Paid ID"].astype(str))
        prefixes = sorted({get_prefix(i) for i in ids if i and i.strip() and i.upper() not in ("NAN", "NONE")})

        if "entity_mapping_by_week" not in st.session_state:
            st.session_state.entity_mapping_by_week = {}
        if st.session_state.current_week not in st.session_state.entity_mapping_by_week:
            st.session_state.entity_mapping_by_week[st.session_state.current_week] = {}
        emap = st.session_state.entity_mapping_by_week[st.session_state.current_week]

        seed = build_entity_seed(current_commande)
        for p in prefixes:
            if p not in emap:
                emap[p] = seed.get(p, "PROD / PLANIFIÉ PROD")

        with st.expander("🗂️ Correspondance Préfixe matricule → Entité (équivalent RECHERCHEX)", expanded=len(prefixes) <= 12):
            st.caption("Règle : 2 premiers caractères = « SA » → SUPPORT + SAI ; sinon recherche du préfixe ci-dessous. "
                       "Défaut proposé depuis la colonne « Departement » du fichier commande, sinon PROD / PLANIFIÉ PROD.")
            filtre = st.text_input("🔍 Filtrer les préfixes affichés", "", key="filtre_prefixes").strip().upper()
            shown = [p for p in prefixes if filtre in p] if filtre else prefixes
            if len(shown) > 40:
                st.info(f"{len(shown)} préfixes : seuls les 40 premiers sont affichés, affinez le filtre.")
                shown = shown[:40]
            mcols = st.columns(4)
            for i, p in enumerate(shown):
                with mcols[i % 4]:
                    try:
                        cur_idx = ENTITES.index(emap.get(p, "PROD / PLANIFIÉ PROD"))
                    except ValueError:
                        cur_idx = 2
                    emap[p] = st.selectbox(f"« {p}… »", ENTITES, index=cur_idx, key=f"map_{st.session_state.current_week}_{p}")

        # 2) Taux d'absentéisme par entité (vos cases 20% / 7% / 14% du Recap)
        with st.expander("📉 Taux d'absentéisme appliqué au « A preparer » (par entité)"):
            taux_by_entity = {}
            tcols = st.columns(3)
            for i, e in enumerate(ENTITES_MAIN):
                with tcols[i]:
                    taux_by_entity[e] = st.number_input(e, min_value=0.0, max_value=50.0, step=0.5,
                                                        value=float(taux_absenteisme), key=f"taux_ent_{e}")

        # 3) « Nombres de presta prévu » (case du Recap)
        st.markdown("##### 👷 Prestataires prévus (ajoutés au total « à commander »)")
        pcols = st.columns(7)
        presta_prevus = {}
        for i, j in enumerate(jours):
            with pcols[i]:
                presta_prevus[j] = st.number_input(j, min_value=0, step=1, value=0, key=f"presta_prevu_{j}")

        # 4) Calcul
        recap, day_totals = compute_recap_menus(current_planning, current_commande, emap, jours, taux_by_entity, taux_absenteisme)

        # 5) Synthèse de la semaine
        summary_df = pd.DataFrame([{"Jour": j, **{e: recap[(j, e)]["total_prep"] for e in ENTITES_MAIN},
                                    "Presta. prévus": int(presta_prevus[j]),
                                    "À commander": day_totals[j] + int(presta_prevus[j])} for j in jours])
        total_row = {"Jour": "TOTAL SEMAINE"}
        for e in ENTITES_MAIN:
            total_row[e] = int(sum(recap[(j, e)]["total_prep"] for j in jours))
        total_row["Presta. prévus"] = int(sum(presta_prevus.values()))
        total_row["À commander"] = int(sum(day_totals.values()) + sum(presta_prevus.values()))
        summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)
        st.markdown("#### 📈 Synthèse de la semaine (repas à préparer)")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        # 6) Export Excel
        export_rows = [{"Jour": j, "Entité": ent, "Choix": r["Choix"], "Nombres": r["Nombres"],
                        "Pourcentage (%)": round(float(r["Pourcentage"]), 1), "A preparer": r["A preparer"]}
                       for j in jours for ent in ENTITES for _, r in recap[(j, ent)]["df"].iterrows()]
        st.download_button("📥 Télécharger le récapitulatif complet (Excel)", data=to_excel(pd.DataFrame(export_rows)),
                           file_name="recap_commandes_menus.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_p6")

        # 7) Détail par jour : blocs empilés PRESTA / SUPPORT+SAI / PROD, comme la feuille Recap
        week_dates = derive_week_dates(st.session_state.current_week)
        mois_fr = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
                   "septembre", "octobre", "novembre", "décembre"]
        if current_planning is None:
            st.info("ℹ️ Planning non chargé : SANS CHOIX ne contient que les absences déclarées.")
        st.markdown("---")
        st.subheader("🗓️ Détail par jour")

        for j in jours:
            d = week_dates.get(j)
            date_txt = f" {d.day:02d} {mois_fr[d.month - 1]} {d.year}" if d else ""
            total_day = day_totals[j] + int(presta_prevus[j])
            ent_line = " &nbsp;•&nbsp; ".join(
                f"<b style='color:{ENTITY_COLORS[e]}'>{e} : {recap[(j, e)]['total_prep']}</b>" for e in ENTITES_MAIN)
            st.markdown(
                f"""
                <div style="background:#003D5B;color:#FFFFFF;padding:10px 18px;border-radius:12px 12px 0 0;font-weight:700;font-size:17px;">
                    📅 {j}{date_txt}
                    <span style="float:right;font-weight:400;font-size:13px;opacity:.85;">Semaine {st.session_state.current_week or ''}</span>
                </div>
                <div style="border:2px solid #003D5B;border-top:none;border-radius:0 0 12px 12px;padding:10px 18px;margin-bottom:12px;">
                    <span style="font-size:17px;font-weight:800;color:#003D5B;">À commander : {total_day} repas</span>
                    <span style="color:#888888;font-size:13px;">&nbsp;(dont {int(presta_prevus[j])} prestataire(s) prévu(s))</span><br>
                    <span style="font-size:13px;color:#444444;">À préparer → {ent_line}</span>
                </div>
                """, unsafe_allow_html=True)
            for ent in ENTITES_MAIN:
                blk = recap[(j, ent)]
                left, right = st.columns([0.08, 0.92])
                with left:
                    st.markdown(entity_badge_html(ent), unsafe_allow_html=True)
                with right:
                    st.caption(f"Planifiés : {blk['planned_n']} · Réponses : {blk['reponses']} · "
                               f"Absences déclarées : {blk['abs_n']} · SANS CHOIX : {blk['sans_choix']}")
                    st.dataframe(style_recap_table(blk["df"]), use_container_width=True, hide_index=True)
            blk_a = recap.get((j, "AUTRE / IGNORÉ"))
            if blk_a and (blk_a["total_n"] > 0 or blk_a["planned_n"] > 0):
                st.warning(f"⚠️ {j} : {blk_a['total_n']} commande(s) / {blk_a['planned_n']} planifié(s) avec un préfixe non classé. Ajustez la correspondance ci-dessus.")
            st.markdown("")
# ═══ [FIN MODIF F-3] ═══

# --- PAGE 7 : ANOMALIES ---   (STRICTEMENT IDENTIQUE À L'ORIGINAL)
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
            st.success("✅ Aucune anomalie commande.")
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
            st.download_button("📥 Télécharger les anomalies commande (Filtré)", data=to_excel(df_filtered_anom), file_name="anomalies_commande.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(df_filtered_anom, use_container_width=True, height=600)
    elif current_planning is None:
        st.warning("Aucune donnée disponible.")

    # --- VÉRIFICATION DES MATRICULES (EN 2e PLAN) ---   (STRICTEMENT IDENTIQUE À L'ORIGINAL)
    st.markdown("---")
    with st.expander("🆔 Vérification des Matricules Paie (Workday vs Liste Actif)"):
        if current_planning is None:
            st.warning("Veuillez d'abord charger une semaine de planning sur la Page 1 pour comparer les matricules.")
        elif st.session_state.get('ref_error'):
            st.error("Le fichier Liste Actif a été importé, mais les colonnes 'Employee ID' et 'Previous Payroll ID' n'ont pas été trouvées.")
            st.write("**Colonnes détectées dans votre fichier par l'outil :**")
            st.write(st.session_state.ref_error)
            st.info("Astuce : Vérifiez que la première ligne de votre fichier Excel contient bien ces en-têtes exacts. S'il y a une ligne vide au-dessus, supprimez-la.")
        elif st.session_state.reference_data is None:
            if file_reference is None:
                st.info("Veuillez importer le fichier 'Liste Actif' dans le menu de gauche pour activer cette vérification.")
        else:
            ref_df = st.session_state.reference_data
            check_df = pd.merge(current_planning[['WORKDAY ID', 'Paid ID', 'Nom', 'Projet']], 
                                ref_df[['WORKDAY ID', 'REF_PAID_ID']], 
                                on='WORKDAY ID', how='left')
            mismatch_df = check_df[(check_df['REF_PAID_ID'].notna()) & 
                                   (check_df['Paid ID'].astype(str) != check_df['REF_PAID_ID'].astype(str))]
            not_found_df = check_df[check_df['REF_PAID_ID'].isna()]
            if not mismatch_df.empty:
                st.warning(f"⚠️ {len(mismatch_df)} collaborateurs ont un matricule paie différent dans le planning par rapport à la Liste Actif.")
                st.dataframe(mismatch_df[['Nom', 'Projet', 'WORKDAY ID', 'Paid ID', 'REF_PAID_ID']], use_container_width=True)
                excel_mismatch = to_excel(mismatch_df[['Nom', 'Projet', 'WORKDAY ID', 'Paid ID', 'REF_PAID_ID']])
                st.download_button("📥 Télécharger les matricules erronés", data=excel_mismatch, file_name="matricules_errones.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.success("✅ Tous les matricules paie présents dans la Liste Actif correspondent au planning.")
            if not not_found_df.empty:
                st.markdown("---")
                st.info(f"ℹ️ {len(not_found_df)} collaborateurs du planning sont introuvables dans la Liste Actif.")
                with st.expander("Voir les collaborateurs introuvables"):
                    st.dataframe(not_found_df[['Nom', 'Projet', 'WORKDAY ID', 'Paid ID']], use_container_width=True)

# --- SIGNATURE FIXEE EN BAS ---
st.markdown(
    "<div class='footer-fix'>Powered by <span style='color: #25E2CC; font-weight: 700; letter-spacing: 1px;'>RAVO SERGIO</span></div>", 
    unsafe_allow_html=True
)

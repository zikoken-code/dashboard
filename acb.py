# =============================================================================
# DASHBOARD DE MAINTENANCE INDUSTRIELLE
# Auteur : Dounje Zakaria
# Description : Dashboard interactif Streamlit pour le suivi de la maintenance
#               à partir du fichier Excel "données.xlsx"
# Pages :
#   - Page 1 : KPIs, jauges, graphiques d'évolution et répartition par type
#   - Page 2 : Pareto machines + AMDEC (composants + criticité)
#   - Page 3 : Plan de maintenance préventive
# =============================================================================

# ── IMPORTS ──────────────────────────────────────────────────────────────────
import os                              # lecture des variables d'environnement
from pathlib import Path               # chemin relatif au script
import streamlit as st                 # framework pour l'interface web
import pandas as pd                    # manipulation des données
import plotly.graph_objects as go      # graphiques avancés
import plotly.express as px            # graphiques rapides
from plotly.subplots import make_subplots  # graphiques à double axe
import warnings
warnings.filterwarnings("ignore")


# =============================================================================
# CONFIGURATION DE LA PAGE (doit être la PREMIÈRE commande Streamlit)
# =============================================================================
st.set_page_config(
    page_title="Dashboard Maintenance Industrielle",
    page_icon="⚙️",
    layout="wide",                    # utilise toute la largeur de l'écran
    initial_sidebar_state="expanded"  # la barre latérale est ouverte au départ
)


# =============================================================================
# STYLES CSS PERSONNALISÉS (thème sombre inspiré GitHub)
# =============================================================================
st.markdown("""
<style>
    /* Fond général de l'application */
    .stApp { background-color: #0d1117; }
    .main .block-container { padding-top: 1rem; padding-bottom: 1rem; }

    /* Barre latérale (sidebar) */
    [data-testid="stSidebar"] { background-color: #161b22; border-right: 1px solid #30363d; }
    [data-testid="stSidebar"] * { color: #e6edf3 !important; }

    /* Carte KPI standard */
    .kpi-card {
        background: linear-gradient(135deg, #1c2128, #21262d);
        border: 1px solid #30363d;
        border-radius: 12px;
        padding: 18px 20px;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        transition: transform 0.2s;
    }
    .kpi-card:hover { transform: translateY(-2px); border-color: #58a6ff; }

    /* Carte KPI en alerte (rouge) */
    .alert-card {
        background: linear-gradient(135deg, #2d1b1b, #3d2020);
        border: 1px solid #f85149;
        border-radius: 12px;
        padding: 18px 20px;
        text-align: center;
    }
    .alert-card .kpi-value { color: #f85149; }

    /* Carte KPI positive (vert) */
    .good-card {
        background: linear-gradient(135deg, #1b2d1b, #203d20);
        border: 1px solid #3fb950;
        border-radius: 12px;
        padding: 18px 20px;
        text-align: center;
    }
    .good-card .kpi-value { color: #3fb950; }

    /* Textes à l'intérieur des cartes KPI */
    .kpi-title  { font-size: 0.78rem; color: #8b949e; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; }
    .kpi-value  { font-size: 2rem;    font-weight: 700; color: #e6edf3; line-height: 1.2; }
    .kpi-unit   { font-size: 0.75rem; color: #8b949e; margin-top: 4px; }

    /* Titres de section */
    .section-header {
        font-size: 1.1rem; font-weight: 600; color: #58a6ff;
        padding: 8px 0; border-bottom: 1px solid #30363d; margin-bottom: 16px;
    }

    /* Titres principaux */
    h1 { color: #e6edf3 !important; }
    h2, h3 { color: #c9d1d9 !important; }

    /* Onglets */
    .stTabs [data-baseweb="tab-list"] { background-color: #161b22; border-radius: 8px; padding: 4px; }
    .stTabs [data-baseweb="tab"] { color: #8b949e !important; border-radius: 6px; }
    .stTabs [aria-selected="true"] { background-color: #21262d !important; color: #58a6ff !important; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# AUTHENTIFICATION
# CORRECTION SÉCURITÉ : les mots de passe sont lus depuis des variables
# d'environnement. Si elles ne sont pas définies, des valeurs par défaut
# sont utilisées (à ne jamais laisser en production).
# Pour définir les variables : export MDPASS_ADMIN="votre_mdp" etc.
# =============================================================================
USERS = {
    "admin":     os.environ.get("MDPASS_ADMIN",     "maintenance2026"),
    "zakaria":   os.environ.get("MDPASS_ZAKARIA",   "dounje123"),
    "encadrant": os.environ.get("MDPASS_ENCADRANT", "fst123"),
}


def check_login(username, password):
    """
    Vérifie si le nom d'utilisateur et le mot de passe sont corrects.
    Retourne True si les identifiants sont valides, False sinon.
    """
    return USERS.get(username) == password


def show_login_page():
    """
    Affiche la page de connexion avec un formulaire centré.
    """
    col_left, col_center, col_right = st.columns([1, 2, 1])

    with col_center:
        st.markdown("<br><br>", unsafe_allow_html=True)

        st.markdown("""
        <div style="text-align: center; margin-bottom: 30px;">
            <div style="font-size: 3rem;">⚙️</div>
            <div style="font-size: 1.6rem; font-weight: 700; color: #e6edf3; margin-top: 10px;">
                Dashboard Maintenance
            </div>
            <div style="font-size: 0.9rem; color: #8b949e; margin-top: 6px;">
                Connectez-vous pour accéder au tableau de bord
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div style="background: linear-gradient(135deg, #1c2128, #21262d);
                    border: 1px solid #30363d; border-radius: 12px; padding: 30px;">
        """, unsafe_allow_html=True)

        username = st.text_input("👤 Nom d'utilisateur", placeholder="Entrez votre nom...")
        password = st.text_input("🔒 Mot de passe", type="password", placeholder="Entrez votre mot de passe...")

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🔑 Se connecter", use_container_width=True):
            if username == "":
                st.error("⚠️ Veuillez entrer un nom d'utilisateur.")
            elif password == "":
                st.error("⚠️ Veuillez entrer un mot de passe.")
            elif check_login(username, password):
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.rerun()
            else:
                st.error("❌ Nom d'utilisateur ou mot de passe incorrect.")

        st.markdown("</div>", unsafe_allow_html=True)


# ── VÉRIFICATION DE L'ÉTAT DE CONNEXION ──────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    show_login_page()
    st.stop()

# ── BOUTON DE DÉCONNEXION ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"👤 Connecté : **{st.session_state.get('username', '')}**")
    if st.button("🚪 Se déconnecter"):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()
    st.markdown("---")


# =============================================================================
# CONSTANTES : COULEURS ET NOMS DES MOIS
# =============================================================================

MONTH_NAMES = {
    1:"Janv", 2:"Févr", 3:"Mars", 4:"Avr",  5:"Mai",  6:"Juin",
    7:"Juil", 8:"Août", 9:"Sept", 10:"Oct", 11:"Nov", 12:"Déc"
}

COLORS = {
    "primary": "#58a6ff",
    "warning": "#f0883e",
    "danger":  "#f85149",
    "success": "#3fb950",
    "purple":  "#a371f7",
    "bg":      "#0d1117",
    "card":    "#21262d",
    "border":  "#30363d",
    "text":    "#e6edf3",
    "subtext": "#8b949e",
}

# Couleurs pour le graphique de criticité AMDEC
CRITICITE_COLORS = {
    "Critique":   COLORS["danger"],
    "Très élevé": COLORS["warning"],
    "Élevé":      "#d29922",
    "Moyen":      COLORS["primary"],
    "Faible":     COLORS["success"],
}

# CORRECTION : dictionnaire CRIT_STYLE manquant — utilisé dans carte_operation()
# Définit la couleur et le fond du badge de criticité sur chaque carte
CRIT_STYLE = {
    "Critique":   {"couleur": "#f85149", "bg": "#f8514920"},
    "Très élevé": {"couleur": "#f0883e", "bg": "#f0883e20"},
    "Élevé":      {"couleur": "#d29922", "bg": "#d2992220"},
    "Moyen":      {"couleur": "#58a6ff", "bg": "#58a6ff20"},
    "Faible":     {"couleur": "#3fb950", "bg": "#3fb95020"},
}

# Paramètres de périodicité pour la page 3
PERIODICITES = {
    "Q": {"label": "Quotidien",      "couleur": "#f85149", "bg": "#f8514910", "icone": "🔴"},
    "H": {"label": "Hebdomadaire",   "couleur": "#f0883e", "bg": "#f0883e10", "icone": "🟠"},
    "M": {"label": "Mensuel",        "couleur": "#58a6ff", "bg": "#58a6ff10", "icone": "🔵"},
    "T": {"label": "Trimestriel",    "couleur": "#3fb950", "bg": "#3fb95010", "icone": "🟢"},
    "A": {"label": "Annuel",         "couleur": "#a371f7", "bg": "#a371f710", "icone": "🟣"},
}

# Paramètres par défaut des graphiques Plotly (fond sombre)
PLOT_LAYOUT = dict(
    paper_bgcolor="#161b22",
    plot_bgcolor="#0d1117",
    font=dict(color="#e6edf3", family="Inter, sans-serif"),
    margin=dict(t=40, b=40, l=40, r=20),
    xaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
    yaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
)


# =============================================================================
# FONCTIONS UTILITAIRES GRAPHIQUES
# =============================================================================

def style_fig(fig, **kwargs):
    """Applique le thème sombre à un graphique Plotly."""
    layout = dict(PLOT_LAYOUT)
    for key in ("xaxis", "yaxis"):
        if key in kwargs and key in layout:
            layout[key] = {**layout[key], **kwargs.pop(key)}
    fig.update_layout(**layout, **kwargs)
    return fig


def gauge_fig(value, title, max_val, unit, thresholds):
    """
    Crée une jauge (gauge) Plotly colorée selon des seuils.

    Paramètres :
        value      : valeur à afficher
        title      : titre de la jauge
        max_val    : valeur maximale de la jauge
        unit       : unité affichée (ex : "%", " h", " min")
        thresholds : tuple (seuil_bas, seuil_haut)
    """
    if value >= thresholds[1]:
        color = COLORS["success"]
    elif value >= thresholds[0]:
        color = COLORS["warning"]
    else:
        color = COLORS["danger"]

    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        title={"text": title, "font": {"size": 14, "color": "#e6edf3"}},
        number={"suffix": unit, "font": {"size": 28, "color": color}},
        gauge={
            "axis": {"range": [0, max_val], "tickcolor": "#8b949e"},
            "bar":  {"color": color, "thickness": 0.3},
            "bgcolor":     "#21262d",
            "bordercolor": "#30363d",
            "steps": [
                {"range": [0,              thresholds[0]], "color": "#2d1b1b"},
                {"range": [thresholds[0],  thresholds[1]], "color": "#2d2a1b"},
                {"range": [thresholds[1],  max_val],       "color": "#1b2d1b"},
            ],
            "threshold": {
                "line": {"color": color, "width": 3},
                "thickness": 0.8,
                "value": value
            },
        }
    ))
    fig.update_layout(
        paper_bgcolor="#161b22",
        height=220,
        margin=dict(t=30, b=0, l=20, r=20),
        font=dict(color="#e6edf3")
    )
    return fig


# =============================================================================
# FONCTIONS UTILITAIRES PAGE 3 : PLAN DE MAINTENANCE
# CORRECTION ORDRE : définies ici (avant tout appel) plutôt qu'après tab2
# =============================================================================

def preparer_plan(plan_df):
    """
    Prépare le DataFrame du plan de maintenance :
    - Renomme les colonnes selon la structure réelle du fichier Excel
    - Remplit les cellules fusionnées (ffill)
    - Supprime les lignes sans opération
    Retourne le DataFrame nettoyé.
    """
    pm = plan_df.copy()

    pm.columns = [
        "Composant", "Machines", "Opération",
        "Q", "H", "M", "T", "A",
        "IC", "Criticité", "Observations", "Responsable"
    ]

    pm["Composant"] = pm["Composant"].ffill()
    pm["Machines"]  = pm["Machines"].ffill()
    pm["IC"]        = pm["IC"].ffill()
    pm["Criticité"] = pm["Criticité"].ffill()

    pm["Composant"] = pm["Composant"].astype(str).str.strip()

    pm = pm.dropna(subset=["Opération"])
    pm = pm[pm["Opération"].astype(str).str.strip() != ""]

    pm["IC"] = pd.to_numeric(pm["IC"], errors="coerce").fillna(0).astype(int)

    return pm.reset_index(drop=True)


def filtrer_par_periode(pm, cle):
    """
    Retourne les opérations qui ont un 'X' dans la colonne de la périodicité donnée.
    Paramètre cle : 'Q', 'H', 'M', 'T' ou 'A'
    """
    return pm[
        pm[cle].astype(str).str.strip().str.upper() == "X"
    ].reset_index(drop=True)


def carte_operation(op, machine, composant, criticite, responsable, ipr, observations, couleur_periode):
    """
    Génère le HTML d'une carte pour une opération de maintenance.
    La barre colorée à gauche correspond à la périodicité.
    """
    # CORRECTION : CRIT_STYLE est maintenant défini dans les constantes
    crit = CRIT_STYLE.get(criticite, {"couleur": "#8b949e", "bg": "#8b949e20"})

    machine_affiche = machine[:60] + "..." if len(str(machine)) > 60 else machine

    obs_html = ""
    if observations and str(observations).strip() not in ["", "nan"]:
        obs_affiche = str(observations)[:80] + "..." if len(str(observations)) > 80 else observations
        obs_html = f"""
        <div style="font-size:0.72rem; color:#8b949e; margin-top:6px;
                    border-top:1px solid #30363d; padding-top:5px;
                    font-style:italic; line-height:1.3;">
            💡 {obs_affiche}
        </div>"""

    return f"""
    <div style="
        background: #1c2128;
        border: 1px solid #30363d;
        border-left: 4px solid {couleur_periode};
        border-radius: 10px;
        padding: 12px 14px;
        margin-bottom: 10px;
        min-height: 140px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    ">
        <div>
            <div style="font-size:0.68rem; color:#a371f7; font-weight:600;
                        text-transform:uppercase; letter-spacing:0.5px; margin-bottom:2px;">
                {composant}
            </div>
            <div style="font-size:0.72rem; color:#8b949e; margin-bottom:6px;
                        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;"
                 title="{machine}">
                🏭 {machine_affiche}
            </div>
            <div style="font-size:0.85rem; color:#e6edf3; font-weight:500;
                        line-height:1.4; margin-bottom:8px;">
                {op}
            </div>
        </div>
        <div>
            {obs_html}
            <div style="display:flex; justify-content:space-between;
                        align-items:center; flex-wrap:wrap; gap:4px; margin-top:8px;">
                <span style="
                    background:{crit['bg']};
                    color:{crit['couleur']};
                    border:1px solid {crit['couleur']}44;
                    font-size:0.68rem; padding:2px 8px;
                    border-radius:10px; font-weight:600;
                ">{criticite}</span>
                <span style="font-size:0.68rem; color:#8b949e;">
                    IPR: <b style="color:#e6edf3">{ipr}</b>
                    &nbsp;|&nbsp; 👤 {responsable}
                </span>
            </div>
        </div>
    </div>
    """


def afficher_section_periode(cle, pm_filtre):
    """
    Affiche un bloc complet pour une périodicité :
    - En-tête avec icône, nom, nombre d'opérations
    - Grille de 4 cartes par ligne
    Utilise st.components.v1.html() pour le rendu des cartes afin d'éviter
    que Streamlit n'affiche le code HTML brut sous les cartes.
    """
    import streamlit.components.v1 as components

    config = PERIODICITES[cle]
    ops    = filtrer_par_periode(pm_filtre, cle)

    if ops.empty:
        return

    nb_ops      = len(ops)
    nb_critique = len(ops[ops["Criticité"] == "Critique"])

    # ── En-tête de section ────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="
        display:flex; align-items:center; gap:12px;
        background:{config['bg']};
        border:1px solid {config['couleur']}44;
        border-radius:10px; padding:12px 20px; margin-bottom:14px;
    ">
        <span style="font-size:1.4rem">{config['icone']}</span>
        <div style="flex:1">
            <div style="font-size:1rem;font-weight:700;color:{config['couleur']}">{config['label']}</div>
            <div style="font-size:0.78rem;color:#8b949e;margin-top:2px">
                {nb_ops} opération(s)
                {f'— dont <b style="color:#f85149">{nb_critique} critique(s)</b>' if nb_critique > 0 else ''}
            </div>
        </div>
        <div style="background:{config['couleur']}22;color:{config['couleur']};
                    border:1px solid {config['couleur']}44;border-radius:20px;
                    padding:4px 14px;font-size:0.8rem;font-weight:600;">{nb_ops}</div>
    </div>
    """, unsafe_allow_html=True)

    # ── Construction du HTML complet de la grille ─────────────────────────────
    # On génère une grille CSS à 4 colonnes dans un seul bloc HTML injecté via
    # components.html() — cela contourne la sanitisation de st.markdown().
    cards_html = ""
    for _, row in ops.iterrows():
        op           = str(row.get("Opération",    "—"))
        machine      = str(row.get("Machines",     "—"))
        composant    = str(row.get("Composant",    "—"))
        criticite    = str(row.get("Criticité",    "—"))
        responsable  = str(row.get("Responsable",  "—"))
        observations = str(row.get("Observations", ""))
        ipr          = int(row.get("IC", 0))
        cards_html += carte_operation(
            op, machine, composant, criticite,
            responsable, ipr, observations,
            config["couleur"]
        )

    # Hauteur estimée : ~170px par carte, 4 par ligne
    nb_lignes      = (nb_ops + 3) // 4
    hauteur_iframe = nb_lignes * 185 + 20

    grille_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ margin:0; padding:0; background:transparent; }}
        .grille {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
        }}
    </style>
    </head>
    <body>
        <div class="grille">
            {cards_html}
        </div>
    </body>
    </html>
    """

    components.html(grille_html, height=hauteur_iframe, scrolling=False)
    st.markdown("<br>", unsafe_allow_html=True)


# =============================================================================
# CHARGEMENT DES DONNÉES
# CORRECTION : chemin relatif au script via pathlib (évite les erreurs si
# l'application est lancée depuis un répertoire différent)
# =============================================================================
BASE_DIR = Path(__file__).parent


@st.cache_data
def load_data():
    """
    Lit le fichier 'données.xlsx' et retourne cinq DataFrames :
        - données    : données de pannes (feuille "Données")
        - planning   : temps de fonctionnement planifié par machine
        - amdec_m    : scores AMDEC par machine (colonnes gauches)
        - amdec_c    : analyse AMDEC par composant (colonnes droites)
        - plan_maint : plan de maintenance préventive
    """
    fichier = BASE_DIR / "données.xlsx"

    # ── Feuille "Données" ─────────────────────────────────────────────────────
    données = pd.read_excel(fichier, sheet_name="Données")
    données = données.rename(columns={"Durée(min)": "Durée_min"})
    données["Date"]   = pd.to_datetime(données["Date"], errors="coerce")
    données["Année"]  = données["Date"].dt.year
    données["Mois"]   = données["Mois"].astype(int)
    données["Machine"] = données["Machine"].str.strip()

    # ── Planning (extrait de la feuille Données) ──────────────────────────────
    planning = (
        données[["Machines", "Temps de fonctionnement planifiée par semaine (min)"]]
        .dropna(subset=["Machines"])
        .drop_duplicates(subset=["Machines"])
        .rename(columns={
            "Machines": "Machine_plan",
            "Temps de fonctionnement planifiée par semaine (min)": "Temps_hebdo"
        })
        .reset_index(drop=True)
    )

    # ── Feuille "AMDEC" ───────────────────────────────────────────────────────
    amdec_raw = pd.read_excel(fichier, sheet_name="AMDEC")

    amdec_m = (
        amdec_raw.iloc[:, :7]
        .rename(columns={
            amdec_raw.columns[0]: "Machine",
            amdec_raw.columns[1]: "Nb_Pannes",
            amdec_raw.columns[2]: "Gravité",
            amdec_raw.columns[3]: "Fréquence",
            amdec_raw.columns[4]: "Détectabilité",
            amdec_raw.columns[5]: "Score_total",
            amdec_raw.columns[6]: "Criticité",
        })
        .dropna(subset=["Machine"])
    )
    amdec_m["Machine"] = amdec_m["Machine"].astype(str).str.strip()

    cols_comp = amdec_raw.columns[9:20]
    amdec_c = (
        amdec_raw[list(cols_comp)]
        .rename(columns={
            cols_comp[0]:  "Composant",
            cols_comp[1]:  "Fonctionnement",
            cols_comp[2]:  "Mode_Défaillance",
            cols_comp[3]:  "Effets",
            cols_comp[4]:  "Causes",
            cols_comp[5]:  "Détection",
            cols_comp[6]:  "F",
            cols_comp[7]:  "G",
            cols_comp[8]:  "D",
            cols_comp[9]:  "IPR",
            cols_comp[10]: "Criticité_comp",
        })
        .dropna(subset=["Composant"])
    )
    amdec_c["Composant"] = amdec_c["Composant"].astype(str).str.strip()

    # ── Feuille "plan de maintenance" ─────────────────────────────────────────
    plan_maint = pd.read_excel(fichier, sheet_name="plan de maintenance", skiprows=1)
    plan_maint = plan_maint.rename(columns={
        plan_maint.columns[0]:  "Composant",
        plan_maint.columns[1]:  "Machines",
        plan_maint.columns[2]:  "Opération",
        plan_maint.columns[3]:  "Q",
        plan_maint.columns[4]:  "H",
        plan_maint.columns[5]:  "M",
        plan_maint.columns[6]:  "T",
        plan_maint.columns[7]:  "A",
        plan_maint.columns[8]:  "IC",
        plan_maint.columns[9]:  "Criticité",
        plan_maint.columns[10]: "Observations",
        plan_maint.columns[11]: "Responsable",
    })
    plan_maint["Composant"] = plan_maint["Composant"].ffill()
    plan_maint["Machines"]  = plan_maint["Machines"].ffill()

    return données, planning, amdec_m, amdec_c, plan_maint


# ── Chargement effectif ───────────────────────────────────────────────────────
try:
    données_raw, planning_données, amdec_machines, amdec_composants, plan_maintenance = load_data()
except Exception as e:
    st.error(f"❌ Impossible de lire le fichier 'données.xlsx'. Erreur : {e}")
    st.stop()


# =============================================================================
# BARRE LATÉRALE : FILTRES INTERACTIFS
# =============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Maintenance Dashboard")
    st.markdown("---")
    st.markdown("### 🎛️ Filtres")

    all_years = sorted(données_raw["Année"].dropna().unique())
    sel_years = st.multiselect("Année", all_years, default=all_years)

    month_name_to_num = {v: k for k, v in MONTH_NAMES.items()}
    sel_month_names = st.multiselect(
        "Mois", list(MONTH_NAMES.values()), default=list(MONTH_NAMES.values())
    )
    sel_months = [month_name_to_num[m] for m in sel_month_names]

    all_machines = sorted(données_raw["Machine"].dropna().unique())
    sel_machines = st.multiselect("Machine", all_machines, default=all_machines)

    if st.button("🔄 Réinitialiser les filtres"):
        st.rerun()

    st.markdown("---")
    st.markdown("### 📊 Info données")

    # CORRECTION : période calculée dynamiquement à partir des données réelles
    date_min = données_raw["Date"].min()
    date_max = données_raw["Date"].max()
    periode_str = (
        f"{MONTH_NAMES.get(date_min.month, '')} {date_min.year} – "
        f"{MONTH_NAMES.get(date_max.month, '')} {date_max.year}"
        if pd.notna(date_min) and pd.notna(date_max) else "N/A"
    )

    st.info(
        f"**{len(données_raw)} pannes** chargées\n\n"
        f"Machines : {len(all_machines)}\n\n"
        f"Période : {periode_str}"
    )
    st.markdown("---")
    st.caption("Maintenance Industrielle")
    st.caption("Dounje Zakaria")


# =============================================================================
# FILTRAGE DES DONNÉES
# =============================================================================
données = données_raw[
    données_raw["Année"].isin(sel_years) &
    données_raw["Mois"].isin(sel_months) &
    données_raw["Machine"].isin(sel_machines)
].copy()

if données.empty:
    st.warning("⚠️ Aucune donnée ne correspond aux filtres sélectionnés.")
    st.stop()


# =============================================================================
# CALCUL DES KPIs PRINCIPAUX
# =============================================================================
def calc_kpis(df_filtered, planning, sel_machines, sel_months, sel_years):
    """
    Calcul des KPIs.
    - Le temps planifié est en minutes dans le fichier → converti en heures (÷ 60)
    - Le MTBF et MTTR sont en heures
    - La disponibilité = MTBF / (MTBF + MTTR) × 100
    """
    nb_pannes   = len(df_filtered)
    temps_arret = df_filtered["Durée_min"].sum()

    mttr     = (temps_arret / nb_pannes / 60) if nb_pannes > 0 else 0
    mttr_min = mttr * 60

    if "Sem." in df_filtered.columns:
        nb_semaines = df_filtered[["Année", "Sem."]].drop_duplicates().shape[0]
    else:
        nb_semaines = len(sel_months) * len(sel_years) * 4.33

    total_planifie_h = 0
    total_reel_h     = 0

    for _, row in planning.iterrows():
        machine     = str(row["Machine_plan"]).strip().lower()
        temps_hebdo = row["Temps_hebdo"]

        planifie_h = (temps_hebdo * nb_semaines) / 60
        arret_h    = df_filtered[
            df_filtered["Machine"].str.strip().str.lower() == machine
        ]["Durée_min"].sum() / 60
        reel_h = max(planifie_h - arret_h, 0)

        total_planifie_h += planifie_h
        total_reel_h     += reel_h

    mtbf = total_reel_h / nb_pannes if nb_pannes > 0 else 0

    # Calcul de disponibilité inchangé (tel que demandé)
    mttt          = mtbf + mttr
    disponibilite = (mtbf / mttt * 100) if total_planifie_h > 0 else 0

    return {
        "nb_pannes":      nb_pannes,
        "temps_arret":    temps_arret,
        "mttr":           mttr,
        "mttr_min":       mttr_min,
        "mtbf":           mtbf,
        "disponibilite":  disponibilite,
        "total_planifie": total_planifie_h,
        "total_reel":     total_reel_h,
    }


kpis = calc_kpis(données, planning_données, sel_machines, sel_months, sel_years)


# =============================================================================
# STRUCTURE EN ONGLETS (3 pages)
# =============================================================================
tab1, tab2, tab3 = st.tabs([
    "📊 Dashboard KPI",
    "📈 Pareto & AMDEC",
    "🔧 Plan de Maintenance"
])


# =============================================================================
# PAGE 1 : DASHBOARD KPI
# AMÉLIORATION : contenu extrait dans la fonction afficher_tab1()
# =============================================================================
def afficher_tab1(données, kpis, sel_machines):
    """Affiche le contenu de l'onglet Dashboard KPI."""
    st.markdown("## 📊 Tableau de bord — Indicateurs de Performance")

    # ── LIGNE 1 : 5 KPIs principaux ──────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)

    dispo = kpis["disponibilite"]
    if dispo < 80:
        dispo_class = "alert-card"
    elif dispo >= 90:
        dispo_class = "good-card"
    else:
        dispo_class = "kpi-card"

    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🔢 Nombre de Pannes</div>
            <div class="kpi-value">{kpis['nb_pannes']:,}</div>
            <div class="kpi-unit">interventions</div>
        </div>""", unsafe_allow_html=True)

    with c2:
        hrs  = int(kpis['temps_arret'] // 60)
        mins = int(kpis['temps_arret'] % 60)
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">⏱️ Temps d'Arrêt Total</div>
            <div class="kpi-value">{hrs:,}h {mins:02d}m</div>
            <div class="kpi-unit">{kpis['temps_arret']:,.0f} minutes</div>
        </div>""", unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🔧 MTTR</div>
            <div class="kpi-value">{kpis['mttr']:.2f}</div>
            <div class="kpi-unit">heures / panne</div>
        </div>""", unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">📅 MTBF</div>
            <div class="kpi-value">{kpis['mtbf']:.2f}</div>
            <div class="kpi-unit">heures entre pannes</div>
        </div>""", unsafe_allow_html=True)

    with c5:
        dispo_label = "⚠️ Sous objectif 80%" if dispo < 80 else "✓ Objectif atteint"
        st.markdown(f"""
        <div class="{dispo_class}">
            <div class="kpi-title">✅ Disponibilité</div>
            <div class="kpi-value">{dispo:.1f}%</div>
            <div class="kpi-unit">{dispo_label}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── LIGNE 2 : 5 KPIs secondaires ─────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)

    type_counts     = données["Type de panne"].value_counts()
    top_type        = type_counts.index[0]   if len(type_counts) > 0 else "N/A"
    top_machine     = données["Machine"].value_counts().index[0] if len(données) > 0 else "N/A"
    nb_top_machine  = données["Machine"].value_counts().iloc[0]  if len(données) > 0 else 0
    moy_par_machine = kpis["nb_pannes"] / len(sel_machines) if sel_machines else 0
    taux_meca       = (type_counts.get("Mécanique", 0) / kpis["nb_pannes"] * 100) if kpis["nb_pannes"] > 0 else 0

    # CORRECTION : planifie_h supprimée (était divisée par 60 une 2e fois, erronée)

    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🏭 Machine + Défaillante</div>
            <div class="kpi-value" style="font-size:1.2rem">{top_machine}</div>
            <div class="kpi-unit">{nb_top_machine} pannes</div>
        </div>""", unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">⚙️ Type Dominant</div>
            <div class="kpi-value" style="font-size:1.2rem">{top_type}</div>
            <div class="kpi-unit">{type_counts.iloc[0] if len(type_counts) > 0 else 0} cas</div>
        </div>""", unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">📊 Moy. Pannes/Machine</div>
            <div class="kpi-value">{moy_par_machine:.1f}</div>
            <div class="kpi-unit">pannes par machine</div>
        </div>""", unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🔩 Taux Mécanique</div>
            <div class="kpi-value">{taux_meca:.1f}%</div>
            <div class="kpi-unit">des pannes</div>
        </div>""", unsafe_allow_html=True)

    with c5:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">📅 Temps Planifié</div>
            <div class="kpi-value">{kpis['total_planifie']:,.0f}h</div>
            <div class="kpi-unit">total période filtrée</div>
        </div>""", unsafe_allow_html=True)

    # ── JAUGES ────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📡 Jauges de Performance</div>', unsafe_allow_html=True)
    g1, g2, g3 = st.columns(3)

    with g1:
        st.plotly_chart(
            gauge_fig(dispo, "Disponibilité (%)", 100, "%", (70, 85)),
            use_container_width=True
        )

    with g2:
        fig_mttr = go.Figure(go.Indicator(
            mode="gauge+number",
            value=kpis["mttr_min"],
            title={"text": "MTTR (min)", "font": {"size": 14, "color": "#e6edf3"}},
            number={"suffix": " min", "font": {"size": 28, "color": COLORS["warning"]}},
            gauge={
                "axis": {"range": [0, 500], "tickcolor": "#8b949e"},
                "bar":  {"color": COLORS["warning"], "thickness": 0.3},
                "bgcolor": "#21262d", "bordercolor": "#30363d",
                "steps": [
                    {"range": [0,   120], "color": "#1b2d1b"},
                    {"range": [120, 300], "color": "#2d2a1b"},
                    {"range": [300, 500], "color": "#2d1b1b"},
                ],
            }
        ))
        fig_mttr.update_layout(
            paper_bgcolor="#161b22", height=220,
            margin=dict(t=30, b=0, l=20, r=20), font=dict(color="#e6edf3")
        )
        st.plotly_chart(fig_mttr, use_container_width=True)

    with g3:
        fig_mtbf = go.Figure(go.Indicator(
            mode="gauge+number",
            value=round(kpis["mtbf"], 1),
            title={"text": "MTBF (heures)", "font": {"size": 14, "color": "#e6edf3"}},
            number={"suffix": " h", "font": {"size": 28, "color": COLORS["primary"]}},
            gauge={
                "axis": {"range": [0, 500]},
                "steps": [
                    {"range": [0,   50],  "color": "#2d1b1b"},
                    {"range": [50,  150], "color": "#2d2a1b"},
                    {"range": [150, 500], "color": "#1b2d1b"},
                ],
            }
        ))
        fig_mtbf.update_layout(
            paper_bgcolor="#161b22", height=220,
            margin=dict(t=30, b=0, l=20, r=20), font=dict(color="#e6edf3")
        )
        st.plotly_chart(fig_mtbf, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── GRAPHIQUES D'ÉVOLUTION ────────────────────────────────────────────────
    st.markdown('<div class="section-header">📈 Analyse Temporelle & Équipements</div>', unsafe_allow_html=True)
    ch1, ch2 = st.columns(2)

    with ch1:
        données["Mois_Année"] = (
            données["Année"].astype(str) + "-" +
            données["Mois"].map(lambda x: f"{x:02d}")
        )
        monthly = (
            données.groupby("Mois_Année")
            .agg(nb_pannes=("Durée_min", "count"), temps_arret=("Durée_min", "sum"))
            .reset_index()
            .sort_values("Mois_Année")
        )
        monthly["label"] = monthly["Mois_Année"].apply(
            lambda x: MONTH_NAMES.get(int(x.split("-")[1]), x) + " " + x.split("-")[0]
        )
        fig_monthly = make_subplots(specs=[[{"secondary_y": True}]])
        fig_monthly.add_trace(
            go.Bar(
                x=monthly["label"], y=monthly["nb_pannes"],
                name="Nb Pannes", marker_color=COLORS["primary"], opacity=0.85
            ),
            secondary_y=False
        )
        fig_monthly.add_trace(
            go.Scatter(
                x=monthly["label"], y=monthly["temps_arret"],
                name="Temps Arrêt (min)",
                line=dict(color=COLORS["warning"], width=2),
                mode="lines+markers", marker=dict(size=5)
            ),
            secondary_y=True
        )
        fig_monthly.update_layout(
            title="Évolution Mensuelle des Pannes & Temps d'Arrêt",
            paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
            font=dict(color="#e6edf3"), height=340,
            legend=dict(bgcolor="#21262d", bordercolor="#30363d"),
            margin=dict(t=40, b=40, l=40, r=40),
            xaxis=dict(gridcolor="#21262d", tickangle=-35),
        )
        fig_monthly.update_yaxes(title_text="Nb pannes",         gridcolor="#21262d", secondary_y=False)
        fig_monthly.update_yaxes(title_text="Temps arrêt (min)", gridcolor="#21262d", secondary_y=True)
        st.plotly_chart(fig_monthly, use_container_width=True)

    with ch2:
        machine_stats = (
            données.groupby("Machine")
            .agg(nb_pannes=("Durée_min", "count"), temps_arret=("Durée_min", "sum"))
            .reset_index()
            .sort_values("temps_arret", ascending=True)
        )
        fig_machines = go.Figure(go.Bar(
            y=machine_stats["Machine"],
            x=machine_stats["temps_arret"],
            orientation="h",
            marker=dict(
                color=machine_stats["temps_arret"],
                colorscale=[[0, "#1c2128"], [0.5, "#1c4f7c"], [1, "#58a6ff"]],
                showscale=False,
            ),
            text=machine_stats["temps_arret"].apply(lambda x: f"{x:,} min"),
            textposition="outside",
            textfont=dict(color="#e6edf3", size=11),
        ))
        style_fig(
            fig_machines,
            title="Temps d'Arrêt Total par Machine",
            height=340,
            xaxis_title="Minutes d'arrêt",
            xaxis=dict(gridcolor="#21262d"),
        )
        st.plotly_chart(fig_machines, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── DONUT + TOP COMPOSANTS ─────────────────────────────────────────────────
    st.markdown('<div class="section-header">🍩 Répartition & Composants</div>', unsafe_allow_html=True)
    cd1, cd2 = st.columns(2)

    with cd1:
        type_data = données["Type de panne"].value_counts().reset_index()
        type_data.columns = ["Type", "Count"]
        fig_donut = go.Figure(go.Pie(
            labels=type_data["Type"],
            values=type_data["Count"],
            hole=0.55,
            marker=dict(colors=[
                COLORS["primary"], COLORS["warning"],
                COLORS["success"], COLORS["purple"]
            ]),
            textinfo="label+percent",
            textfont=dict(size=11, color="#e6edf3"),
        ))
        fig_donut.update_layout(
            title="Répartition par Type de Panne",
            paper_bgcolor="#161b22", height=320,
            font=dict(color="#e6edf3"),
            margin=dict(t=40, b=20, l=20, r=20),
            showlegend=False,
            annotations=[dict(
                text=f"<b>{kpis['nb_pannes']}</b><br>pannes",
                x=0.5, y=0.5, font_size=16, showarrow=False, font_color="#e6edf3"
            )]
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    with cd2:
        comp_stats = (
            données.groupby("Composant")
            .agg(nb_pannes=("Durée_min", "count"), temps_arret=("Durée_min", "sum"))
            .reset_index()
            .sort_values("nb_pannes", ascending=False)
            .head(10)
            .sort_values("nb_pannes", ascending=True)
        )
        fig_comp = go.Figure(go.Bar(
            y=comp_stats["Composant"],
            x=comp_stats["nb_pannes"],
            orientation="h",
            marker_color=COLORS["purple"],
            opacity=0.85,
            text=comp_stats["nb_pannes"],
            textposition="outside",
            textfont=dict(color="#e6edf3", size=11),
        ))
        style_fig(
            fig_comp,
            title="Top 10 Composants Défaillants",
            height=320,
            xaxis_title="Nombre de pannes",
            xaxis=dict(gridcolor="#21262d"),
        )
        st.plotly_chart(fig_comp, use_container_width=True)


# =============================================================================
# PAGE 2 : PARETO & AMDEC
# AMÉLIORATION : contenu extrait dans la fonction afficher_tab2()
# =============================================================================
def afficher_tab2(données, amdec_machines, amdec_composants):
    """Affiche le contenu de l'onglet Pareto & AMDEC."""
    st.markdown("## 📈 Analyse Pareto & AMDEC")

    # ── PARETO 1 : TEMPS D'ARRÊT PAR MACHINE ─────────────────────────────────
    st.markdown('<div class="section-header">⏱️ Pareto — Temps d\'Arrêt par Machine</div>', unsafe_allow_html=True)

    pareto_temps = (
        données.groupby("Machine")["Durée_min"].sum()
        .reset_index()
        .rename(columns={"Durée_min": "Temps_Arret"})
        .sort_values("Temps_Arret", ascending=False)
        .reset_index(drop=True)
    )
    pareto_temps["Cumul_%"] = (
        pareto_temps["Temps_Arret"].cumsum()
        / pareto_temps["Temps_Arret"].sum() * 100
    )

    fig_p1 = make_subplots(specs=[[{"secondary_y": True}]])
    fig_p1.add_trace(
        go.Bar(
            x=pareto_temps["Machine"],
            y=pareto_temps["Temps_Arret"],
            name="Temps d'arrêt (min)",
            marker=dict(color=[
                COLORS["danger"] if c <= 80 else COLORS["subtext"]
                for c in pareto_temps["Cumul_%"]
            ]),
            opacity=0.9,
            text=pareto_temps["Temps_Arret"].apply(lambda x: f"{x:,.0f}"),
            textposition="outside",
            textfont=dict(size=11, color="#e6edf3"),
        ),
        secondary_y=False
    )
    fig_p1.add_trace(
        go.Scatter(
            x=pareto_temps["Machine"],
            y=pareto_temps["Cumul_%"],
            name="Cumul %",
            mode="lines+markers",
            line=dict(color=COLORS["warning"], width=2.5),
            marker=dict(size=7, color=COLORS["warning"]),
        ),
        secondary_y=True
    )
    fig_p1.add_hline(
        y=80, line_dash="dash", line_color=COLORS["success"],
        annotation_text="Seuil 80%",
        annotation_font_color=COLORS["success"],
        secondary_y=True
    )
    fig_p1.update_layout(
        title="Pareto — Temps d'arrêt total par machine",
        paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
        font=dict(color="#e6edf3"), height=400,
        legend=dict(bgcolor="#21262d", bordercolor="#30363d"),
        margin=dict(t=50, b=50, l=50, r=50),
        xaxis=dict(gridcolor="#21262d", tickangle=-20),
    )
    fig_p1.update_yaxes(title_text="Temps d'arrêt (min)", gridcolor="#21262d", secondary_y=False)
    fig_p1.update_yaxes(title_text="Cumul (%)", range=[0, 105], gridcolor="#21262d", secondary_y=True)
    st.plotly_chart(fig_p1, use_container_width=True)

    # ── PARETO 2 : NOMBRE DE PANNES PAR MACHINE ───────────────────────────────
    st.markdown('<div class="section-header">🔢 Pareto — Nombre de Pannes par Machine</div>', unsafe_allow_html=True)

    pareto_nb = (
        données.groupby("Machine").size()
        .reset_index(name="Nb_Pannes")
        .sort_values("Nb_Pannes", ascending=False)
        .reset_index(drop=True)
    )
    pareto_nb["Cumul_%"] = (
        pareto_nb["Nb_Pannes"].cumsum()
        / pareto_nb["Nb_Pannes"].sum() * 100
    )

    fig_p2 = make_subplots(specs=[[{"secondary_y": True}]])
    fig_p2.add_trace(
        go.Bar(
            x=pareto_nb["Machine"],
            y=pareto_nb["Nb_Pannes"],
            name="Nb pannes",
            marker=dict(color=[
                COLORS["primary"] if c <= 80 else COLORS["subtext"]
                for c in pareto_nb["Cumul_%"]
            ]),
            opacity=0.9,
            text=pareto_nb["Nb_Pannes"],
            textposition="outside",
            textfont=dict(size=11, color="#e6edf3"),
        ),
        secondary_y=False
    )
    fig_p2.add_trace(
        go.Scatter(
            x=pareto_nb["Machine"],
            y=pareto_nb["Cumul_%"],
            name="Cumul %",
            mode="lines+markers",
            line=dict(color=COLORS["warning"], width=2.5),
            marker=dict(size=7, color=COLORS["warning"]),
        ),
        secondary_y=True
    )
    fig_p2.add_hline(
        y=80, line_dash="dash", line_color=COLORS["success"],
        annotation_text="Seuil 80%",
        annotation_font_color=COLORS["success"],
        secondary_y=True
    )
    fig_p2.update_layout(
        title="Pareto — Nombre de pannes par machine",
        paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
        font=dict(color="#e6edf3"), height=400,
        legend=dict(bgcolor="#21262d", bordercolor="#30363d"),
        margin=dict(t=50, b=50, l=50, r=50),
        xaxis=dict(gridcolor="#21262d", tickangle=-20),
    )
    fig_p2.update_yaxes(title_text="Nombre de pannes", gridcolor="#21262d", secondary_y=False)
    fig_p2.update_yaxes(title_text="Cumul (%)", range=[0, 105], gridcolor="#21262d", secondary_y=True)
    st.plotly_chart(fig_p2, use_container_width=True)

    # ── AMDEC : CRITICITÉ MACHINES + COMPOSANTS ───────────────────────────────
    st.markdown('<div class="section-header">⚠️ Analyse AMDEC — Criticité Machines & Composants</div>', unsafe_allow_html=True)

    a1, a2 = st.columns(2)

    with a1:
        amdec_sorted = amdec_machines.sort_values("Score_total", ascending=True)
        bar_colors   = [CRITICITE_COLORS.get(c, COLORS["subtext"]) for c in amdec_sorted["Criticité"]]

        fig_amdec_m = go.Figure(go.Bar(
            x=amdec_sorted["Score_total"],
            y=amdec_sorted["Machine"],
            orientation="h",
            marker_color=bar_colors,
            text=amdec_sorted.apply(
                lambda r: f"{r['Score_total']} — {r['Criticité']}", axis=1
            ),
            textposition="outside",
            textfont=dict(size=10, color="#e6edf3"),
        ))
        style_fig(
            fig_amdec_m,
            title="Score Criticité AMDEC par Machine",
            height=350,
            xaxis_title="Score (G × F × D)",
            xaxis=dict(range=[0, 130], gridcolor="#21262d"),
        )
        for _, row in amdec_sorted.iterrows():
            fig_amdec_m.add_annotation(
                x=0, y=row["Machine"],
                text=f"  {int(row['Nb_Pannes'])} pannes",
                showarrow=False,
                xanchor="left",
                font=dict(size=9, color="#8b949e")
            )
        st.plotly_chart(fig_amdec_m, use_container_width=True)

    with a2:
        comp_sorted  = amdec_composants.sort_values("IPR", ascending=True)
        bar_colors_c = [CRITICITE_COLORS.get(c, COLORS["subtext"]) for c in comp_sorted["Criticité_comp"]]

        fig_amdec_c = go.Figure(go.Bar(
            x=comp_sorted["IPR"],
            y=comp_sorted["Composant"],
            orientation="h",
            marker_color=bar_colors_c,
            text=comp_sorted.apply(
                lambda r: f"IPR={int(r['IPR'])} — {r['Criticité_comp']}", axis=1
            ),
            textposition="outside",
            textfont=dict(size=9, color="#e6edf3"),
        ))
        style_fig(
            fig_amdec_c,
            title="IPR par Composant (F × G × D)",
            height=350,
            xaxis_title="IPR",
            xaxis=dict(range=[0, 130], gridcolor="#21262d"),
        )
        st.plotly_chart(fig_amdec_c, use_container_width=True)

    # ── TABLEAU AMDEC COMPOSANTS ──────────────────────────────────────────────
    st.markdown('<div class="section-header">📋 Détail AMDEC — Composants</div>', unsafe_allow_html=True)

    données_amdec_display = amdec_composants[[
        "Composant", "Mode_Défaillance", "Effets", "Causes",
        "F", "G", "D", "IPR", "Criticité_comp"
    ]].rename(columns={
        "Mode_Défaillance": "Mode Défaillance",
        "Criticité_comp":   "Criticité",
    })

    def color_criticite(val):
        """Colore les cellules de la colonne Criticité."""
        colors_map = {
            "Critique":   "background-color: #3d1515; color: #f85149",
            "Très élevé": "background-color: #3d2e15; color: #f0883e",
            "Élevé":      "background-color: #2e2a15; color: #d29922",
            "Moyen":      "background-color: #152030; color: #58a6ff",
            "Faible":     "background-color: #152315; color: #3fb950",
        }
        return colors_map.get(val, "")

    st.dataframe(
        données_amdec_display.style.map(color_criticite, subset=["Criticité"]),
        use_container_width=True,
        height=400
    )


# =============================================================================
# PAGE 3 : PLAN DE MAINTENANCE PRÉVENTIVE
# AMÉLIORATION : contenu extrait dans la fonction afficher_tab3()
# CORRECTION : st.stop() remplacé par structure conditionnelle (else)
# =============================================================================
def afficher_tab3(plan_maintenance):
    """Affiche le contenu de l'onglet Plan de Maintenance Préventive."""
    st.markdown("## 🔧 Plan de Maintenance Préventive")

    pm = preparer_plan(plan_maintenance)

    # ── STATISTIQUES GLOBALES ─────────────────────────────────────────────────
    st.markdown('<div class="section-header">📊 Vue d\'ensemble du plan</div>', unsafe_allow_html=True)

    k1, k2, k3, k4, k5 = st.columns(5)

    nb_total      = len(pm)
    nb_composants = pm["Composant"].nunique()
    nb_critique   = len(pm[pm["Criticité"] == "Critique"])
    nb_quotidien  = len(filtrer_par_periode(pm, "Q"))
    nb_annuel     = len(filtrer_par_periode(pm, "A"))

    with k1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">📋 Opérations totales</div>
            <div class="kpi-value">{nb_total}</div>
            <div class="kpi-unit">dans le plan</div>
        </div>""", unsafe_allow_html=True)

    with k2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🔩 Composants couverts</div>
            <div class="kpi-value">{nb_composants}</div>
            <div class="kpi-unit">types de composants</div>
        </div>""", unsafe_allow_html=True)

    with k3:
        st.markdown(f"""
        <div class="alert-card">
            <div class="kpi-title">🚨 Opérations critiques</div>
            <div class="kpi-value">{nb_critique}</div>
            <div class="kpi-unit">priorité maximale</div>
        </div>""", unsafe_allow_html=True)

    with k4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🔴 Vérifications/jour</div>
            <div class="kpi-value">{nb_quotidien}</div>
            <div class="kpi-unit">opérations quotidiennes</div>
        </div>""", unsafe_allow_html=True)

    with k5:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🟣 Révisions annuelles</div>
            <div class="kpi-value">{nb_annuel}</div>
            <div class="kpi-unit">opérations annuelles</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── FILTRES ────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">🎛️ Filtres</div>', unsafe_allow_html=True)

    f1, f2, f3 = st.columns(3)

    with f1:
        options_periode = ["Toutes"] + [v["label"] for v in PERIODICITES.values()]
        sel_periode = st.radio("📅 Périodicité", options=options_periode, horizontal=True)

    with f2:
        options_crit = ["Toutes"] + sorted(pm["Criticité"].dropna().unique().tolist())
        sel_crit_pm  = st.selectbox("⚠️ Criticité", options_crit)

    with f3:
        options_comp = ["Tous"] + sorted(pm["Composant"].dropna().unique().tolist())
        sel_comp_pm  = st.selectbox("🔩 Composant", options_comp)

    st.markdown("<br>", unsafe_allow_html=True)

    # Appliquer les filtres
    pm_filtre = pm.copy()
    if sel_crit_pm != "Toutes":
        pm_filtre = pm_filtre[pm_filtre["Criticité"] == sel_crit_pm]
    if sel_comp_pm != "Tous":
        pm_filtre = pm_filtre[pm_filtre["Composant"] == sel_comp_pm]

    # CORRECTION : st.stop() remplacé par un bloc else pour ne pas bloquer
    # les autres onglets du dashboard
    if pm_filtre.empty:
        st.warning("⚠️ Aucune opération ne correspond aux filtres sélectionnés.")
    else:
        # ── AFFICHAGE DES CARTES PAR PÉRIODICITÉ ──────────────────────────────
        st.markdown('<div class="section-header">🗂️ Opérations de maintenance</div>', unsafe_allow_html=True)

        cle_choisie = None
        for k, v in PERIODICITES.items():
            if v["label"] == sel_periode:
                cle_choisie = k
                break

        if sel_periode == "Toutes":
            for cle in PERIODICITES.keys():
                afficher_section_periode(cle, pm_filtre)
        else:
            afficher_section_periode(cle_choisie, pm_filtre)
            ops_periode = filtrer_par_periode(pm_filtre, cle_choisie)
            if ops_periode.empty:
                st.info(f"ℹ️ Aucune opération **{sel_periode}** pour les filtres sélectionnés.")

        # ── LÉGENDE ───────────────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""
        <div style="
            background: linear-gradient(135deg, #1c2128, #21262d);
            border: 1px solid #30363d;
            border-radius: 12px;
            padding: 16px 20px;
        ">
            <div style="color:#58a6ff; font-weight:600; margin-bottom:10px; font-size:0.9rem;">
                📌 Légende
            </div>
            <div style="display:flex; flex-wrap:wrap; gap:20px; font-size:0.8rem; color:#8b949e;">
                <span><b style="color:#f85149">🔴 Q</b> = Quotidien</span>
                <span><b style="color:#f0883e">🟠 H</b> = Hebdomadaire</span>
                <span><b style="color:#58a6ff">🔵 M</b> = Mensuel</span>
                <span><b style="color:#3fb950">🟢 T</b> = Trimestriel</span>
                <span><b style="color:#a371f7">🟣 A</b> = Annuel</span>
                <span style="border-left:1px solid #30363d; padding-left:20px;">
                    <b style="color:#e6edf3">IPR</b> = Indice Priorité Risque (F × G × D)
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)


# =============================================================================
# APPEL DES FONCTIONS DANS LES ONGLETS
# =============================================================================
with tab1:
    afficher_tab1(données, kpis, sel_machines)

with tab2:
    afficher_tab2(données, amdec_machines, amdec_composants)

with tab3:
    afficher_tab3(plan_maintenance)

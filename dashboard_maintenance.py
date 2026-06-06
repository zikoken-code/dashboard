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
import streamlit as st        # framework pour l'interface web
import pandas as pd           # manipulation des données
import plotly.graph_objects as go   # graphiques avancés
import plotly.express as px         # graphiques rapides
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
# AUTHENTIFICATION — À PLACER JUSTE APRÈS LES IMPORTS, AVANT TOUT LE RESTE
# =============================================================================

# Dictionnaire des utilisateurs autorisés : { "nom_utilisateur": "mot_de_passe" }
# Tu peux ajouter autant d'utilisateurs que tu veux
USERS = {
    "admin":   "maintenance2026",
    "zakaria": "dounje123",
    "encadrant": "fst123",
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
    # Centrer le formulaire avec des colonnes vides à gauche et droite
    col_left, col_center, col_right = st.columns([1, 2, 1])

    with col_center:
        st.markdown("<br><br>", unsafe_allow_html=True)

        # Logo / titre
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

        # Carte de connexion
        st.markdown("""
        <div style="background: linear-gradient(135deg, #1c2128, #21262d);
                    border: 1px solid #30363d; border-radius: 12px; padding: 30px;">
        """, unsafe_allow_html=True)

        # Champs du formulaire
        username = st.text_input("👤 Nom d'utilisateur", placeholder="Entrez votre nom...")
        password = st.text_input("🔒 Mot de passe", type="password", placeholder="Entrez votre mot de passe...")

        st.markdown("<br>", unsafe_allow_html=True)

        # Bouton de connexion
        if st.button("🔑 Se connecter", use_container_width=True):
            if username == "":
                st.error("⚠️ Veuillez entrer un nom d'utilisateur.")
            elif password == "":
                st.error("⚠️ Veuillez entrer un mot de passe.")
            elif check_login(username, password):
                # Connexion réussie : on sauvegarde l'état dans la session
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.rerun()  # recharger la page pour afficher le dashboard
            else:
                st.error("❌ Nom d'utilisateur ou mot de passe incorrect.")

        st.markdown("</div>", unsafe_allow_html=True)

# ── VÉRIFICATION DE L'ÉTAT DE CONNEXION ──────────────────────────────────────
# st.session_state persiste entre les interactions de l'utilisateur
# Si "authenticated" n'est pas dans la session, l'utilisateur n'est pas connecté

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    show_login_page()
    st.stop()  # ← IMPORTANT : arrête l'exécution du reste du code (le dashboard)

# ── BOUTON DE DÉCONNEXION (affiché dans la sidebar quand connecté) ────────────
# À placer dans le bloc "with st.sidebar:" existant, en haut
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

# Correspondance numéro de mois → abréviation française
MONTH_NAMES = {
    1:"Janv", 2:"Févr", 3:"Mars", 4:"Avr",  5:"Mai",  6:"Juin",
    7:"Juil", 8:"Août", 9:"Sept", 10:"Oct", 11:"Nov", 12:"Déc"
}

# Palette de couleurs pour les graphiques
COLORS = {
    "primary": "#58a6ff",   # bleu clair – graphiques principaux
    "warning": "#f0883e",   # orange    – avertissements
    "danger":  "#f85149",   # rouge     – alertes / Pareto 80%
    "success": "#3fb950",   # vert      – bon état
    "purple":  "#a371f7",   # violet    – 4e type de panne
    "bg":      "#0d1117",   # fond général
    "card":    "#21262d",   # fond des cartes
    "border":  "#30363d",   # bordures
    "text":    "#e6edf3",   # texte principal
    "subtext": "#8b949e",   # texte secondaire
}

# Couleurs pour le graphique de criticité AMDEC
CRITICITE_COLORS = {
    "Critique":   COLORS["danger"],
    "Très élevé": COLORS["warning"],
    "Élevé":      "#d29922",
    "Moyen":      COLORS["primary"],
    "Faible":     COLORS["success"],
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
# FONCTIONS UTILITAIRES
# =============================================================================

def style_fig(fig, **kwargs):
    """
    Applique le thème sombre à un graphique Plotly.
    Les kwargs permettent de surcharger ou d'ajouter des paramètres de layout.
    """
    layout = dict(PLOT_LAYOUT)
    # Fusion des axes si fournis en kwargs (pour éviter un remplacement brutal)
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
                     - en dessous de seuil_bas → rouge (danger)
                     - entre les deux          → orange (warning)
                     - au-dessus de seuil_haut → vert (success)
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
                {"range": [0,              thresholds[0]], "color": "#2d1b1b"},  # zone rouge
                {"range": [thresholds[0],  thresholds[1]], "color": "#2d2a1b"},  # zone orange
                {"range": [thresholds[1],  max_val],       "color": "#1b2d1b"},  # zone verte
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
# CHARGEMENT DES DONNÉES (mis en cache pour ne lire le fichier qu'une seule fois)
# =============================================================================

@st.cache_data
def load_data():
    """
    Lit le fichier 'données.xlsx' et retourne trois DataFrames :
        - df         : données de pannes (feuille "Données")
        - planning   : temps de fonctionnement planifié par machine
        - amdec_m    : scores AMDEC par machine (colonnes gauches)
        - amdec_c    : analyse AMDEC par composant (colonnes droites)
        - plan_maint : plan de maintenance préventive (feuille "plan de maintenance")
    """
    # ── Lecture feuille "Données" ─────────────────────────────────────────────
    df = pd.read_excel("données.xlsx", sheet_name="Données")

    # Renommer la colonne avec espaces pour faciliter l'accès
    df = df.rename(columns={"Durée(min)": "Durée_min"})

    # Convertir la colonne Date en datetime si ce n'est pas déjà fait
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # Extraire l'année à partir de la date
    df["Année"] = df["Date"].dt.year

    # S'assurer que Mois est bien un entier
    df["Mois"] = df["Mois"].astype(int)

    # Nettoyer les noms de machine (supprimer espaces inutiles)
    df["Machine"] = df["Machine"].str.strip()

    # ── Extraction du tableau de planification ────────────────────────────────
    # Le planning se trouve dans les colonnes "Machines" et
    # "Temps de fonctionnement planifiée par semaine (min)" du même onglet
    planning = (
        df[["Machines", "Temps de fonctionnement planifiée par semaine (min)"]]
        .dropna(subset=["Machines"])  # garder les lignes non vides
        .drop_duplicates(subset=["Machines"])  # une ligne par machine
        .rename(columns={
            "Machines": "Machine_plan",
            "Temps de fonctionnement planifiée par semaine (min)": "Temps_hebdo"
        })
        .reset_index(drop=True)
    )

    # ── Lecture feuille "AMDEC" ───────────────────────────────────────────────
    amdec_raw = pd.read_excel("données.xlsx", sheet_name="AMDEC")

    # Partie GAUCHE : scores globaux par machine (colonnes 0 à 6)
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

    # Partie DROITE : analyse détaillée par composant (colonnes 9 à 19)
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

    # ── Lecture feuille "plan de maintenance" ─────────────────────────────────
    plan_maint = pd.read_excel(
        "données.xlsx",
        sheet_name="plan de maintenance",
        skiprows=1   # sauter la ligne de titre fusionnée
    )
    # Renommer les colonnes de périodicité
    plan_maint = plan_maint.rename(columns={
        plan_maint.columns[0]:  "Composant",
        plan_maint.columns[1]:  "Machines",
        plan_maint.columns[2]:  "Opération",
        plan_maint.columns[3]:  "Q",   # Quotidien
        plan_maint.columns[4]:  "H",   # Hebdomadaire
        plan_maint.columns[5]:  "M",   # Mensuel
        plan_maint.columns[6]:  "T",   # Trimestriel
        plan_maint.columns[7]:  "A",   # Annuel
        plan_maint.columns[8]:  "IC",  # Indice de criticité
        plan_maint.columns[9]:  "Criticité",
        plan_maint.columns[10]: "Observations",
        plan_maint.columns[11]: "Responsable",
    })
    # Remplir les cellules fusionnées (composant et machine)
    plan_maint["Composant"] = plan_maint["Composant"].ffill()
    plan_maint["Machines"]  = plan_maint["Machines"].ffill()

    return df, planning, amdec_m, amdec_c, plan_maint


# ── Chargement effectif des données ──────────────────────────────────────────
try:
    df_raw, planning_df, amdec_machines, amdec_composants, plan_maintenance = load_data()
except Exception as e:
    st.error(f"❌ Impossible de lire le fichier 'données.xlsx'. Erreur : {e}")
    st.stop()  # arrêter l'application si le fichier est introuvable


# =============================================================================
# BARRE LATÉRALE : FILTRES INTERACTIFS
# =============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Maintenance Dashboard")
    st.markdown("---")
    st.markdown("### 🎛️ Filtres")

    # ── Filtre par année ──────────────────────────────────────────────────────
    all_years = sorted(df_raw["Année"].dropna().unique())
    sel_years = st.multiselect("Année", all_years, default=all_years)

    # ── Filtre par mois ───────────────────────────────────────────────────────
    # On affiche les noms (Janv, Févr…) et on reconvertit en numéros
    month_name_to_num = {v: k for k, v in MONTH_NAMES.items()}
    sel_month_names = st.multiselect(
        "Mois", list(MONTH_NAMES.values()), default=list(MONTH_NAMES.values())
    )
    sel_months = [month_name_to_num[m] for m in sel_month_names]

    # ── Filtre par machine ────────────────────────────────────────────────────
    all_machines = sorted(df_raw["Machine"].dropna().unique())
    sel_machines = st.multiselect("Machine", all_machines, default=all_machines)

    # ── Bouton de réinitialisation ────────────────────────────────────────────
    if st.button("🔄 Réinitialiser les filtres"):
        st.rerun()

    st.markdown("---")
    st.markdown("### 📊 Info données")
    st.info(
        f"**{len(df_raw)} pannes** chargées\n\n"
        f"Machines : {len(all_machines)}\n\n"
        f"Période : Jan 2024 – Juil 2025"
    )
    st.markdown("---")
    st.caption("Maintenance Industrielle")
    st.caption("Dounje Zakaria")


# =============================================================================
# FILTRAGE DES DONNÉES selon les sélections de la sidebar
# =============================================================================
# On ne garde que les lignes correspondant aux filtres choisis
df = df_raw[
    df_raw["Année"].isin(sel_years) &
    df_raw["Mois"].isin(sel_months) &
    df_raw["Machine"].isin(sel_machines)
].copy()

# Afficher un avertissement si les filtres donnent un résultat vide
if df.empty:
    st.warning("⚠️ Aucune donnée ne correspond aux filtres sélectionnés.")
    st.stop()


# =============================================================================
# CALCUL DES KPIs PRINCIPAUX
# =============================================================================
def calc_kpis(df_filtered, planning, sel_machines, sel_months, sel_years):
    """
    Calcule les indicateurs de performance de maintenance.
    Correction : le temps planifié est calculé en semaines réelles
    (nb_semaines = nombre de semaines distinctes dans les données filtrées)
    """
    nb_pannes    = len(df_filtered)
    temps_arret  = df_filtered["Durée_min"].sum()
    mttr         = temps_arret / nb_pannes if nb_pannes > 0 else 0

    # ── Calcul du nombre de semaines réelles dans la période filtrée ──────────
    # On utilise la colonne "Sem." du fichier Excel qui contient le numéro de semaine
    # nb_semaines = nombre de semaines distinctes × nombre d'années sélectionnées
    if "Sem." in df_filtered.columns:
        # Compter les paires (Année, Semaine) uniques dans les données filtrées
        nb_semaines = df_filtered[["Année", "Sem."]].drop_duplicates().shape[0]
    else:
        # Fallback : estimation 4.33 semaines par mois
        nb_semaines = len(sel_months) * len(sel_years) * 4

    total_planifie = 0
    total_reel     = 0

    for _, row in planning.iterrows():
        machine       = row["Machine_plan"]
        temps_hebdo   = row["Temps_hebdo"]

        # Temps planifié total = temps hebdo × nombre de semaines réelles
        planifie_total = temps_hebdo * nb_semaines

        # Temps d'arrêt de cette machine sur la période filtrée
        arret_machine = df_filtered[
            df_filtered["Machine"] == machine
        ]["Durée_min"].sum()

        # Temps de fonctionnement réel (jamais négatif)
        reel = max(planifie_total - arret_machine, 0)

        total_planifie += planifie_total
        total_reel     += reel

    mtbf         = total_reel / nb_pannes if nb_pannes > 0 else 0
    disponibilite = (total_reel / total_planifie * 100) if total_planifie > 0 else 0

    return {
        "nb_pannes":      nb_pannes,
        "temps_arret":    temps_arret,
        "mttr":           mttr,
        "mtbf":           mtbf,
        "disponibilite":  disponibilite,
        "total_planifie": total_planifie,
        "total_reel":     total_reel,
    }


kpis = calc_kpis(df, planning_df, sel_machines, sel_months, sel_years)


# =============================================================================
# STRUCTURE EN ONGLETS (3 pages)
# =============================================================================
tab1, tab2, tab3 = st.tabs([
    "📊 Dashboard KPI",
    "📈 Pareto & AMDEC",
    "🔧 Plan de Maintenance"
])


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  PAGE 1 : DASHBOARD KPI                                                  ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
with tab1:
    st.markdown("## 📊 Tableau de bord — Indicateurs de Performance")

    # ── LIGNE 1 : 5 KPIs principaux ──────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)

    # Disponibilité : couleur de la carte selon le seuil
    dispo = kpis["disponibilite"]
    if dispo < 80:
        dispo_class = "alert-card"   # rouge si < 80%
    elif dispo >= 90:
        dispo_class = "good-card"    # vert si ≥ 90%
    else:
        dispo_class = "kpi-card"     # neutre sinon

    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🔢 Nombre de Pannes</div>
            <div class="kpi-value">{kpis['nb_pannes']:,}</div>
            <div class="kpi-unit">interventions</div>
        </div>""", unsafe_allow_html=True)

    with c2:
        # Conversion minutes → heures + minutes pour l'affichage
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
            <div class="kpi-value">{kpis['mttr']:.1f}</div>
            <div class="kpi-unit">min / panne</div>
        </div>""", unsafe_allow_html=True)

    with c4:
        mtbf_h = kpis['mtbf'] / 60  # conversion min → heures
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">📅 MTBF</div>
            <div class="kpi-value">{mtbf_h:.1f}</div>
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

    # Calculs pour les KPIs secondaires
    type_counts     = df["Type de panne"].value_counts()
    top_type        = type_counts.index[0]    if len(type_counts) > 0 else "N/A"
    top_machine     = df["Machine"].value_counts().index[0] if len(df) > 0 else "N/A"
    nb_top_machine  = df["Machine"].value_counts().iloc[0]  if len(df) > 0 else 0
    moy_par_machine = kpis["nb_pannes"] / len(sel_machines) if sel_machines else 0
    taux_meca       = (type_counts.get("Mécanique", 0) / kpis["nb_pannes"] * 100) if kpis["nb_pannes"] > 0 else 0
    planifie_h      = kpis["total_planifie"] / 60

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
            <div class="kpi-value">{planifie_h:,.0f}h</div>
            <div class="kpi-unit">total période filtrée</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── JAUGES : Disponibilité, MTTR, MTBF ───────────────────────────────────
    st.markdown('<div class="section-header">📡 Jauges de Performance</div>', unsafe_allow_html=True)
    g1, g2, g3 = st.columns(3)

    with g1:
        # Disponibilité : seuils à 70% et 85%
        st.plotly_chart(
            gauge_fig(dispo, "Disponibilité (%)", 100, "%", (70, 85)),
            use_container_width=True
        )

    with g2:
        # MTTR : plus c'est bas, mieux c'est → les zones sont inversées
        fig_mttr = go.Figure(go.Indicator(
            mode="gauge+number",
            value=kpis["mttr"],
            title={"text": "MTTR (min)", "font": {"size": 14, "color": "#e6edf3"}},
            number={"suffix": " min", "font": {"size": 28, "color": COLORS["warning"]}},
            gauge={
                "axis": {"range": [0, 500], "tickcolor": "#8b949e"},
                "bar":  {"color": COLORS["warning"], "thickness": 0.3},
                "bgcolor": "#21262d", "bordercolor": "#30363d",
                "steps": [
                    {"range": [0,   120], "color": "#1b2d1b"},  # bon (vert)
                    {"range": [120, 300], "color": "#2d2a1b"},  # moyen (orange)
                    {"range": [300, 500], "color": "#2d1b1b"},  # mauvais (rouge)
                ],
            }
        ))
        fig_mttr.update_layout(
            paper_bgcolor="#161b22", height=220,
            margin=dict(t=30, b=0, l=20, r=20), font=dict(color="#e6edf3")
        )
        st.plotly_chart(fig_mttr, use_container_width=True)

    with g3:
        # MTBF en heures : plus c'est haut, mieux c'est
        mtbf_h_val = round(kpis["mtbf"] / 60, 1)
        fig_mtbf = go.Figure(go.Indicator(
            mode="gauge+number",
            value=mtbf_h_val,
            title={"text": "MTBF (heures)", "font": {"size": 14, "color": "#e6edf3"}},
            number={"suffix": " h", "font": {"size": 28, "color": COLORS["primary"]}},
            gauge={
                "axis": {"range": [0, 200], "tickcolor": "#8b949e"},
                "bar":  {"color": COLORS["primary"], "thickness": 0.3},
                "bgcolor": "#21262d", "bordercolor": "#30363d",
                "steps": [
                    {"range": [0,   50],  "color": "#2d1b1b"},
                    {"range": [50,  100], "color": "#2d2a1b"},
                    {"range": [100, 200], "color": "#1b2d1b"},
                ],
            }
        ))
        fig_mtbf.update_layout(
            paper_bgcolor="#161b22", height=220,
            margin=dict(t=30, b=0, l=20, r=20), font=dict(color="#e6edf3")
        )
        st.plotly_chart(fig_mtbf, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── GRAPHIQUES D'ÉVOLUTION ET TEMPS D'ARRÊT PAR MACHINE ──────────────────
    st.markdown('<div class="section-header">📈 Analyse Temporelle & Équipements</div>', unsafe_allow_html=True)
    ch1, ch2 = st.columns(2)

    with ch1:
        # Évolution mensuelle : barres pour nb pannes + courbe pour temps d'arrêt
        df["Mois_Année"] = (
            df["Année"].astype(str) + "-" +
            df["Mois"].map(lambda x: f"{x:02d}")
        )
        monthly = (
            df.groupby("Mois_Année")
            .agg(nb_pannes=("Durée_min", "count"), temps_arret=("Durée_min", "sum"))
            .reset_index()
            .sort_values("Mois_Année")
        )
        # Créer un label lisible "Janv 2024", "Févr 2024", etc.
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
        fig_monthly.update_yaxes(title_text="Nb pannes",       gridcolor="#21262d", secondary_y=False)
        fig_monthly.update_yaxes(title_text="Temps arrêt (min)", gridcolor="#21262d", secondary_y=True)
        st.plotly_chart(fig_monthly, use_container_width=True)

    with ch2:
        # Barres horizontales : temps d'arrêt total par machine (tri décroissant)
        machine_stats = (
            df.groupby("Machine")
            .agg(nb_pannes=("Durée_min", "count"), temps_arret=("Durée_min", "sum"))
            .reset_index()
            .sort_values("temps_arret", ascending=True)  # ascending pour bar horizontal
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

    # ── GRAPHIQUES DONUT + BARRES COMPOSANTS ──────────────────────────────────
    st.markdown('<div class="section-header">🍩 Répartition & Composants</div>', unsafe_allow_html=True)
    cd1, cd2 = st.columns(2)

    with cd1:
        # Donut : répartition par type de panne
        type_data = df["Type de panne"].value_counts().reset_index()
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
        # Barres : top 10 composants les plus défaillants
        comp_stats = (
            df.groupby("Composant")
            .agg(nb_pannes=("Durée_min", "count"), temps_arret=("Durée_min", "sum"))
            .reset_index()
            .sort_values("nb_pannes", ascending=False)
            .head(10)
            .sort_values("nb_pannes", ascending=True)  # pour barre horizontale
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


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  PAGE 2 : PARETO & AMDEC                                                 ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
with tab2:
    st.markdown("## 📈 Analyse Pareto & AMDEC")

    # ── PARETO 1 : TEMPS D'ARRÊT PAR MACHINE ─────────────────────────────────
    st.markdown('<div class="section-header">⏱️ Pareto — Temps d\'Arrêt par Machine</div>', unsafe_allow_html=True)

    # Calcul du Pareto : trier par temps décroissant, puis cumuler le %
    pareto_temps = (
        df.groupby("Machine")["Durée_min"].sum()
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
            # Barres rouges pour les machines dans les 80% du total
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
    # Ligne de référence à 80% (règle de Pareto)
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
        df.groupby("Machine").size()
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

    # ── AMDEC : CRITICITÉ DES MACHINES + COMPOSANTS ───────────────────────────
    st.markdown('<div class="section-header">⚠️ Analyse AMDEC — Criticité Machines & Composants</div>', unsafe_allow_html=True)

    a1, a2 = st.columns(2)

    with a1:
        # Barres horizontales : score de criticité par machine (AMDEC machines)
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
        # Annotation du nombre de pannes sur chaque barre
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
        # Barres horizontales : IPR (score) par composant (AMDEC composants)
        comp_sorted = amdec_composants.sort_values("IPR", ascending=True)
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

    # ── TABLEAU AMDEC COMPOSANTS (détail) ─────────────────────────────────────
    st.markdown('<div class="section-header">📋 Détail AMDEC — Composants</div>', unsafe_allow_html=True)

    # Sélection des colonnes utiles pour l'affichage
    df_amdec_display = amdec_composants[[
        "Composant", "Mode_Défaillance", "Effets", "Causes",
        "F", "G", "D", "IPR", "Criticité_comp"
    ]].rename(columns={
        "Mode_Défaillance": "Mode Défaillance",
        "Criticité_comp":   "Criticité",
    })

    # Colorier le fond des lignes selon la criticité
    def color_criticite(val):
        """Fonction pour colorer les cellules de la colonne Criticité."""
        colors_map = {
            "Critique":   "background-color: #3d1515; color: #f85149",
            "Très élevé": "background-color: #3d2e15; color: #f0883e",
            "Élevé":      "background-color: #2e2a15; color: #d29922",
            "Moyen":      "background-color: #152030; color: #58a6ff",
            "Faible":     "background-color: #152315; color: #3fb950",
        }
        return colors_map.get(val, "")

    st.dataframe(
        df_amdec_display.style.map(color_criticite, subset=["Criticité"]),
        use_container_width=True,
        height=400
    )


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  PAGE 3 : PLAN DE MAINTENANCE PRÉVENTIVE                                 ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
with tab3:
    st.markdown("## 🔧 Plan de Maintenance Préventive")

    # ── FILTRES DU PLAN DE MAINTENANCE ───────────────────────────────────────
    col_f1, col_f2, col_f3 = st.columns(3)

    with col_f1:
        # Filtre par composant
        comp_list = ["Tous"] + sorted(
            plan_maintenance["Composant"].dropna().unique().tolist()
        )
        sel_comp = st.selectbox("🔩 Composant", comp_list)

    with col_f2:
        # Filtre par criticité
        crit_list = ["Toutes"] + sorted(
            plan_maintenance["Criticité"].dropna().unique().tolist()
        )
        sel_crit = st.selectbox("⚠️ Criticité", crit_list)

    with col_f3:
        # Filtre par périodicité
        periodicite_options = {
            "Toutes":        None,
            "Quotidien":     "Q",
            "Hebdomadaire":  "H",
            "Mensuel":       "M",
            "Trimestriel":   "T",
            "Annuel":        "A",
        }
        sel_period = st.selectbox("📅 Périodicité", list(periodicite_options.keys()))

    # Appliquer les filtres sur le plan de maintenance
    pm = plan_maintenance.copy()
    if sel_comp != "Tous":
        pm = pm[pm["Composant"] == sel_comp]
    if sel_crit != "Toutes":
        pm = pm[pm["Criticité"] == sel_crit]
    if sel_period != "Toutes":
        col_period = periodicite_options[sel_period]
        pm = pm[pm[col_period].notna() & (pm[col_period] == "X")]

    st.markdown("<br>", unsafe_allow_html=True)

    # ── STATISTIQUES RAPIDES DU PLAN ──────────────────────────────────────────
    stat1, stat2, stat3, stat4 = st.columns(4)

    nb_operations  = len(pm)
    nb_composants  = pm["Composant"].nunique()
    nb_critiques   = len(pm[pm["Criticité"] == "Critique"])
    nb_quotidiennes = len(pm[pm["Q"].notna() & (pm["Q"] == "X")])

    with stat1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">📋 Opérations</div>
            <div class="kpi-value">{nb_operations}</div>
            <div class="kpi-unit">au total</div>
        </div>""", unsafe_allow_html=True)

    with stat2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">🔩 Composants</div>
            <div class="kpi-value">{nb_composants}</div>
            <div class="kpi-unit">couverts</div>
        </div>""", unsafe_allow_html=True)

    with stat3:
        st.markdown(f"""
        <div class="alert-card">
            <div class="kpi-title">🚨 Critiques</div>
            <div class="kpi-value">{nb_critiques}</div>
            <div class="kpi-unit">opérations prioritaires</div>
        </div>""", unsafe_allow_html=True)

    with stat4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">📅 Quotidiennes</div>
            <div class="kpi-value">{nb_quotidiennes}</div>
            <div class="kpi-unit">vérifications / jour</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── GRAPHIQUE : RÉPARTITION DES OPÉRATIONS PAR PÉRIODICITÉ ───────────────
    st.markdown('<div class="section-header">📊 Répartition des Opérations</div>', unsafe_allow_html=True)

    gc1, gc2 = st.columns(2)

    with gc1:
        # Compter le nombre d'opérations pour chaque périodicité
        periods = {"Quotidien": "Q", "Hebdomadaire": "H", "Mensuel": "M",
                   "Trimestriel": "T", "Annuel": "A"}
        period_counts = {
            name: len(pm[pm[col].notna() & (pm[col] == "X")])
            for name, col in periods.items()
        }
        fig_periods = go.Figure(go.Bar(
            x=list(period_counts.keys()),
            y=list(period_counts.values()),
            marker_color=[
                COLORS["danger"], COLORS["warning"], COLORS["primary"],
                COLORS["purple"], COLORS["success"]
            ],
            text=list(period_counts.values()),
            textposition="outside",
            textfont=dict(color="#e6edf3"),
        ))
        style_fig(
            fig_periods,
            title="Nb Opérations par Périodicité",
            height=280,
            yaxis_title="Nombre d'opérations",
        )
        st.plotly_chart(fig_periods, use_container_width=True)

    with gc2:
        # Donut : répartition par criticité dans le plan de maintenance
        crit_counts = pm["Criticité"].value_counts().reset_index()
        crit_counts.columns = ["Criticité", "Count"]
        fig_crit = go.Figure(go.Pie(
            labels=crit_counts["Criticité"],
            values=crit_counts["Count"],
            hole=0.5,
            marker=dict(colors=[
                CRITICITE_COLORS.get(c, COLORS["subtext"])
                for c in crit_counts["Criticité"]
            ]),
            textinfo="label+percent",
            textfont=dict(size=11, color="#e6edf3"),
        ))
        fig_crit.update_layout(
            title="Répartition par Criticité",
            paper_bgcolor="#161b22", height=280,
            font=dict(color="#e6edf3"),
            margin=dict(t=40, b=20, l=20, r=20),
            showlegend=False,
        )
        st.plotly_chart(fig_crit, use_container_width=True)

    # ── TABLEAU DU PLAN DE MAINTENANCE ───────────────────────────────────────
    st.markdown('<div class="section-header">📋 Tableau du Plan de Maintenance</div>', unsafe_allow_html=True)

    # Construire la colonne "Périodicité" lisible à partir des colonnes Q/H/M/T/A
    def get_period_label(row):
        """Retourne une chaîne comme 'Q / M' si les colonnes Q et M contiennent 'X'."""
        labels = []
        for lbl, col in [("Q", "Q"), ("H", "H"), ("M", "M"), ("T", "T"), ("A", "A")]:
            if row.get(col) == "X":
                labels.append(lbl)
        return " / ".join(labels) if labels else "—"

    pm_display = pm.copy()
    pm_display["Périodicité"] = pm_display.apply(get_period_label, axis=1)

    # Colonnes à afficher dans le tableau final
    cols_display = ["Composant", "Machines", "Opération", "Périodicité",
                    "IC", "Criticité", "Responsable", "Observations"]
    pm_display = pm_display[[c for c in cols_display if c in pm_display.columns]]
    pm_display = pm_display.dropna(subset=["Opération"])

    # Appliquer la couleur de criticité
    def style_pm(val):
        """Colore la cellule selon la valeur de criticité."""
        colors_map = {
            "Critique":   "background-color: #3d1515; color: #f85149; font-weight: bold",
            "Très élevé": "background-color: #3d2e15; color: #f0883e",
            "Élevé":      "background-color: #2e2a15; color: #d29922",
            "Moyen":      "background-color: #152030; color: #58a6ff",
            "Faible":     "background-color: #152315; color: #3fb950",
        }
        return colors_map.get(val, "")

    if "Criticité" in pm_display.columns:
        st.dataframe(
            pm_display.style.map(style_pm, subset=["Criticité"]),
            use_container_width=True,
            height=500
        )
    else:
        st.dataframe(pm_display, use_container_width=True, height=500)

    # ── NOTE D'UTILISATION ────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1c2128, #21262d);
                border: 1px solid #30363d; border-radius: 12px; padding: 16px;">
        <div style="color: #58a6ff; font-weight: 600; margin-bottom: 8px;">
            📌 Légende des périodicités
        </div>
        <div style="color: #8b949e; font-size: 0.88rem; line-height: 2;">
            <b style="color:#f85149">Q</b> = Quotidien &nbsp;|&nbsp;
            <b style="color:#f0883e">H</b> = Hebdomadaire &nbsp;|&nbsp;
            <b style="color:#58a6ff">M</b> = Mensuel &nbsp;|&nbsp;
            <b style="color:#a371f7">T</b> = Trimestriel &nbsp;|&nbsp;
            <b style="color:#3fb950">A</b> = Annuel &nbsp;|&nbsp;
            <b style="color:#e6edf3">IC</b> = Indice de Criticité (IPR)
        </div>
    </div>
    """, unsafe_allow_html=True)

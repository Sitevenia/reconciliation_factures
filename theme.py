"""
theme.py — Design system partagé
Applique un style cohérent (couleurs, polices, boutons, widgets)
identique à la page d'accueil sur toutes les pages de l'application.
"""
import streamlit as st

# ── Palette ───────────────────────────────────────────────────────────────────
COLOR_DARK   = "#1a3a5c"   # bleu marine foncé
COLOR_MID    = "#2d6094"   # bleu moyen
COLOR_LIGHT  = "#4a90d9"   # bleu clair / survol
COLOR_BG     = "#f0f4f8"   # fond général
COLOR_TEXT_S = "#e8f4fd"   # texte clair sidebar
COLOR_LABEL  = "#c8ddf0"   # labels sidebar
COLOR_MUTED  = "#6b7c93"   # texte secondaire
COLOR_BORDER = "#dde6f0"   # bordures légères

# ── CSS global ────────────────────────────────────────────────────────────────
_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* ── Typographie globale ──────────────────────────── */
html, body, [class*="css"], .stMarkdown, .stText, p, div, label {{
    font-family: 'Inter', sans-serif !important;
}}
span:not([class*="material"]):not([class*="notranslate"]) {{
    font-family: 'Inter', sans-serif !important;
}}
span[class*="material"],
span.notranslate {{
    font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
}}

/* ── Fond général ─────────────────────────────────── */
[data-testid="stAppViewContainer"] > .main {{
    background-color: {COLOR_BG};
}}

/* ── Masquer chrome Streamlit ─────────────────────── */
#MainMenu                        {{ visibility: hidden; }}
footer                           {{ visibility: hidden; }}
[data-testid="stToolbar"]        {{ visibility: hidden; }}
[data-testid="stDecoration"]     {{ visibility: hidden; }}

/* ── Sidebar ──────────────────────────────────────── */
[data-testid="stSidebar"] {{
    background-color: {COLOR_DARK} !important;
}}
[data-testid="stSidebar"] * {{
    color: {COLOR_TEXT_S} !important;
    font-family: 'Inter', sans-serif !important;
}}
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stSelectbox select,
[data-testid="stSidebar"] .stNumberInput input {{
    background-color: #2d5a8a !important;
    color: {COLOR_TEXT_S} !important;
    border-color: #3d7ab0 !important;
    border-radius: 6px !important;
}}
[data-testid="stSidebar"] label {{
    color: {COLOR_LABEL} !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
}}
[data-testid="stSidebar"] .stMarkdown p {{
    color: {COLOR_LABEL} !important;
}}
[data-testid="stSidebarNav"] a span {{
    color: {COLOR_TEXT_S} !important;
    font-weight: 500;
}}
[data-testid="stSidebarNav"] a:hover {{
    background-color: rgba(255,255,255,0.08) !important;
    border-radius: 8px;
}}
[data-testid="stSidebarNav"] a[aria-selected="true"] {{
    background-color: rgba(74,144,217,0.25) !important;
    border-radius: 8px;
}}
[data-testid="stSidebar"] .stButton > button {{
    background-color: {COLOR_MID} !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    width: 100%;
}}
[data-testid="stSidebar"] .stButton > button:hover {{
    background-color: {COLOR_LIGHT} !important;
}}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] span,
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small,
[data-testid="stSidebar"] [data-testid="stFileUploader"] label,
[data-testid="stSidebar"] [data-testid="stFileUploader"] span:not(button span),
[data-testid="stSidebar"] [data-testid="stFileUploader"] small {{
    color: {COLOR_DARK} !important;
}}

/* ── Titres ───────────────────────────────────────── */
h1, h2, h3 {{
    font-family: 'Inter', sans-serif !important;
    color: {COLOR_DARK} !important;
    font-weight: 700 !important;
}}
h1 {{ font-size: 1.8rem !important; }}
h2 {{ font-size: 1.3rem !important; }}
h3 {{ font-size: 1.05rem !important; }}

/* ── Boutons principaux ───────────────────────────── */
.stButton > button {{
    background-color: {COLOR_DARK} !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.875rem !important;
    padding: 0.45rem 1.1rem !important;
    transition: background-color 0.15s ease, box-shadow 0.15s ease !important;
    box-shadow: 0 1px 3px rgba(26,58,92,0.25) !important;
}}
.stButton > button:hover {{
    background-color: {COLOR_MID} !important;
    box-shadow: 0 3px 10px rgba(45,96,148,0.35) !important;
}}
.stButton > button:active {{
    background-color: #122840 !important;
}}
.stButton > button[kind="primary"],
[data-testid="baseButton-primary"] {{
    background-color: {COLOR_MID} !important;
}}
.stButton > button[kind="primary"]:hover,
[data-testid="baseButton-primary"]:hover {{
    background-color: {COLOR_DARK} !important;
}}

/* ── Bouton de téléchargement ─────────────────────── */
[data-testid="stDownloadButton"] > button {{
    background-color: {COLOR_DARK} !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.875rem !important;
    box-shadow: 0 1px 3px rgba(26,58,92,0.25) !important;
    transition: background-color 0.15s ease !important;
}}
[data-testid="stDownloadButton"] > button:hover {{
    background-color: {COLOR_MID} !important;
}}

/* ── Métriques ────────────────────────────────────── */
[data-testid="stMetric"] {{
    background: white;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    border-left: 4px solid {COLOR_MID};
}}
[data-testid="stMetricLabel"] {{
    color: {COLOR_MUTED} !important;
    font-size: 0.8rem !important;
    font-weight: 500 !important;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}}
[data-testid="stMetricValue"] {{
    color: {COLOR_DARK} !important;
    font-weight: 700 !important;
    font-size: 1.6rem !important;
}}

/* ── Expanders ────────────────────────────────────── */
[data-testid="stExpander"] {{
    background: white;
    border-radius: 10px !important;
    border: 1px solid {COLOR_BORDER} !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    margin-bottom: 0.6rem;
}}
[data-testid="stExpander"] summary {{
    color: {COLOR_DARK} !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
}}
[data-testid="stExpander"] summary:hover {{
    color: {COLOR_MID} !important;
}}

/* ── Tabs ─────────────────────────────────────────── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {{
    border-bottom: 2px solid {COLOR_BORDER};
    gap: 0.2rem;
    background: transparent;
}}
[data-testid="stTabs"] [data-baseweb="tab"] {{
    color: {COLOR_MUTED} !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    border-radius: 6px 6px 0 0;
    padding: 0.5rem 1rem !important;
    background: transparent !important;
    border: none !important;
}}
[data-testid="stTabs"] [aria-selected="true"] {{
    color: {COLOR_DARK} !important;
    font-weight: 700 !important;
    border-bottom: 3px solid {COLOR_MID} !important;
    background: transparent !important;
}}

/* ── Inputs / Selectbox ───────────────────────────── */
[data-testid="stSelectbox"] > div > div,
[data-testid="stMultiSelect"] > div > div {{
    border-color: {COLOR_BORDER} !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
}}
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stTextArea"] textarea {{
    border-color: {COLOR_BORDER} !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
}}
[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {{
    border-color: {COLOR_MID} !important;
    box-shadow: 0 0 0 2px rgba(45,96,148,0.15) !important;
}}

/* ── Checkbox / Radio ─────────────────────────────── */
[data-testid="stCheckbox"] label,
[data-testid="stRadio"] label {{
    font-family: 'Inter', sans-serif !important;
    color: {COLOR_DARK} !important;
}}

/* ── Dataframes ───────────────────────────────────── */
[data-testid="stDataFrame"] thead tr th {{
    background-color: {COLOR_DARK} !important;
    color: white !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
}}
[data-testid="stDataFrame"] tbody tr:hover td {{
    background-color: #e8f0f8 !important;
}}

/* ── Zone de dépôt de fichier ─────────────────────── */
[data-testid="stFileUploader"] {{
    background: white;
    border-radius: 10px;
    border: 2px dashed {COLOR_BORDER} !important;
    padding: 0.5rem;
    transition: border-color 0.2s ease;
}}
[data-testid="stFileUploader"]:hover {{
    border-color: {COLOR_MID} !important;
}}
[data-testid="stFileUploaderDropzone"] {{
    background: #f7fafd !important;
    border-radius: 8px !important;
    border: none !important;
}}
[data-testid="stFileUploaderDropzone"] > div > span {{
    color: {COLOR_DARK} !important;
    font-weight: 500 !important;
    font-family: 'Inter', sans-serif !important;
}}
[data-testid="stFileUploaderDropzone"] > div > small {{
    color: {COLOR_MUTED} !important;
    font-family: 'Inter', sans-serif !important;
}}
[data-testid="stFileUploaderDropzoneButton"],
[data-testid="stFileUploader"] button {{
    background-color: {COLOR_DARK} !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    padding: 0.35rem 0.9rem !important;
}}
[data-testid="stFileUploaderDropzoneButton"]:hover,
[data-testid="stFileUploader"] button:hover {{
    background-color: {COLOR_MID} !important;
}}

/* ── Alertes info / success / warning / error ─────── */
[data-testid="stAlert"] {{
    border-radius: 8px !important;
    border-left-width: 4px !important;
    border-left-style: solid !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 0.9rem !important;
    padding: 0.75rem 1rem !important;
}}
[data-testid="stAlert"][data-baseweb="notification"][kind="info"],
div[data-testid="stAlert"] > div[class*="info"] {{
    background-color: #e8f0fb !important;
    border-left-color: {COLOR_MID} !important;
    color: {COLOR_DARK} !important;
}}
[data-testid="stAlert"][kind="success"] {{
    background-color: #e8f5ee !important;
    border-left-color: #2e7d52 !important;
    color: #1a4a33 !important;
}}
[data-testid="stAlert"][kind="warning"] {{
    background-color: #fff8e6 !important;
    border-left-color: #c8860a !important;
    color: #7a5200 !important;
}}
[data-testid="stAlert"][kind="error"] {{
    background-color: #fdecea !important;
    border-left-color: #c0392b !important;
    color: #7a1a14 !important;
}}

/* ── Barre de progression ─────────────────────────── */
[data-testid="stProgressBar"] > div > div {{
    background-color: {COLOR_MID} !important;
    border-radius: 4px !important;
}}
[data-testid="stProgressBar"] > div {{
    background-color: {COLOR_BORDER} !important;
    border-radius: 4px !important;
    height: 8px !important;
}}

/* ── Slider ───────────────────────────────────────── */
[data-testid="stSlider"] [role="slider"] {{
    background-color: {COLOR_MID} !important;
    border: 2px solid white !important;
    box-shadow: 0 0 0 2px {COLOR_MID} !important;
}}
[data-testid="stSlider"] [data-baseweb="slider"] div[class*="Track"] {{
    background-color: {COLOR_BORDER} !important;
}}
[data-testid="stSlider"] [data-baseweb="slider"] div[class*="Track"]:first-child {{
    background-color: {COLOR_MID} !important;
}}
[data-testid="stSlider"] label {{
    color: {COLOR_DARK} !important;
    font-weight: 500 !important;
}}

/* ── Date input ───────────────────────────────────── */
[data-testid="stDateInput"] input {{
    border-color: {COLOR_BORDER} !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
    color: {COLOR_DARK} !important;
}}
[data-testid="stDateInput"] input:focus {{
    border-color: {COLOR_MID} !important;
    box-shadow: 0 0 0 2px rgba(45,96,148,0.15) !important;
}}

/* ── Modales / Dialogues ──────────────────────────── */
[data-testid="stModal"],
div[role="dialog"] {{
    border-radius: 14px !important;
    box-shadow: 0 8px 40px rgba(26,58,92,0.22) !important;
    font-family: 'Inter', sans-serif !important;
    border: 1px solid {COLOR_BORDER} !important;
}}
div[role="dialog"] header,
[data-testid="stModal"] header {{
    background: linear-gradient(135deg, {COLOR_DARK} 0%, {COLOR_MID} 100%) !important;
    color: white !important;
    border-radius: 14px 14px 0 0 !important;
    padding: 1rem 1.4rem !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
}}
div[role="dialog"] [data-testid="stMarkdown"] p {{
    font-family: 'Inter', sans-serif !important;
    color: {COLOR_DARK} !important;
}}

/* ── Popover ──────────────────────────────────────── */
[data-testid="stPopover"] [data-baseweb="popover"] {{
    border-radius: 10px !important;
    border: 1px solid {COLOR_BORDER} !important;
    box-shadow: 0 4px 18px rgba(26,58,92,0.15) !important;
    font-family: 'Inter', sans-serif !important;
}}

/* ── Spinner / Status ─────────────────────────────── */
[data-testid="stSpinner"] > div {{
    border-top-color: {COLOR_MID} !important;
}}
[data-testid="stStatusWidget"] {{
    border-radius: 8px !important;
    border: 1px solid {COLOR_BORDER} !important;
    font-family: 'Inter', sans-serif !important;
}}

/* ── Toast notifications ──────────────────────────── */
[data-testid="stToast"] {{
    background-color: {COLOR_DARK} !important;
    color: white !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 500 !important;
    box-shadow: 0 4px 16px rgba(26,58,92,0.35) !important;
}}

/* ── Conteneur principal (padding) ───────────────── */
.block-container {{
    padding-top: 3rem !important;
    padding-bottom: 2rem !important;
}}

/* ── Bandeau en-tête de page ──────────────────────── */
.page-header {{
    background: linear-gradient(135deg, {COLOR_DARK} 0%, {COLOR_MID} 100%);
    color: white;
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 18px rgba(26, 58, 92, 0.35);
}}
.page-header h1 {{
    font-size: 1.7rem !important;
    font-weight: 700 !important;
    margin: 0 0 0.3rem 0 !important;
    color: white !important;
}}
.page-header p {{
    font-size: 0.9rem;
    color: #a8c8e8;
    margin: 0;
}}

/* ── Titre de section ─────────────────────────────── */
.section-title {{
    color: {COLOR_DARK};
    font-size: 1.1rem;
    font-weight: 600;
    border-bottom: 2px solid {COLOR_MID};
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 1rem 0;
    letter-spacing: 0.02em;
    font-family: 'Inter', sans-serif !important;
}}

/* ── Cartes génériques ────────────────────────────── */
.card {{
    background: white;
    border-radius: 10px;
    padding: 1.2rem 1.4rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    border-left: 4px solid {COLOR_MID};
    margin-bottom: 0.8rem;
}}

/* ── Bandeau accueil (app-header) ─────────────────── */
.app-header {{
    background: linear-gradient(135deg, {COLOR_DARK} 0%, {COLOR_MID} 100%);
    color: white;
    padding: 2rem 2.5rem;
    border-radius: 12px;
    margin-bottom: 2rem;
    box-shadow: 0 4px 18px rgba(26, 58, 92, 0.35);
}}
.app-header h1 {{
    font-size: 2rem !important;
    font-weight: 700 !important;
    margin: 0 0 0.4rem 0 !important;
    color: white !important;
}}
.app-header p {{
    font-size: 1rem;
    color: #a8c8e8;
    margin: 0;
}}

/* ── Cartes outils (accueil) ──────────────────────── */
.tool-card {{
    background: white;
    border-radius: 10px;
    padding: 1.4rem 1.2rem 1rem 1.2rem;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.07);
    border-left: 4px solid {COLOR_MID};
    transition: box-shadow 0.2s ease, transform 0.2s ease;
    height: 100%;
    margin-bottom: 0.6rem;
}}
.tool-card:hover {{
    box-shadow: 0 6px 22px rgba(45, 96, 148, 0.22);
    transform: translateY(-2px);
    border-left-color: {COLOR_LIGHT};
}}
.tool-card .tc-icon {{ font-size: 1.8rem; line-height: 1; margin-bottom: 0.5rem; }}
.tool-card .tc-title {{
    color: {COLOR_DARK};
    font-size: 1rem;
    font-weight: 700;
    margin: 0 0 0.35rem 0;
    font-family: 'Inter', sans-serif !important;
}}
.tool-card .tc-desc {{
    color: {COLOR_MUTED};
    font-size: 0.82rem;
    line-height: 1.45;
    margin: 0;
}}

/* ── Liens page (st.page_link) ────────────────────── */
[data-testid="stPageLink"] a,
[data-testid="stPageLink-NavLink"] {{
    display: inline-block;
    width: auto;
    background-color: {COLOR_MID} !important;
    color: white !important;
    border: none;
    border-radius: 6px;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600;
    font-size: 0.85rem;
    padding: 0.4rem 1rem;
    margin-top: 0.4rem;
    text-align: center;
    text-decoration: none !important;
    white-space: nowrap !important;
    transition: background-color 0.15s ease;
}}
[data-testid="stPageLink"] a:hover,
[data-testid="stPageLink-NavLink"]:hover {{
    background-color: {COLOR_DARK} !important;
    color: white !important;
}}

/* ── Texte de l'uploader (zone principale) ────────── */
[data-testid="stFileUploaderDropzone"] span {{
    color: {COLOR_DARK} !important;
    font-weight: 500 !important;
}}
[data-testid="stFileUploaderDropzone"] small {{
    color: {COLOR_MUTED} !important;
}}
[data-testid="stFileUploader"] label {{
    color: {COLOR_DARK} !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
}}

/* ── Bouton retour accueil ────────────────────────── */
a.back-btn {{
    display: inline-block;
    background-color: {COLOR_DARK};
    color: white !important;
    text-decoration: none !important;
    font-family: 'Inter', sans-serif;
    font-weight: 700;
    font-size: 1rem;
    padding: 0.7rem 1.6rem;
    border-radius: 8px;
    white-space: nowrap;
    box-shadow: 0 2px 8px rgba(26,58,92,0.30);
    transition: background-color 0.15s ease, box-shadow 0.15s ease;
    margin-bottom: 1.2rem;
    margin-top: 0;
}}
a.back-btn:hover {{
    background-color: {COLOR_MID} !important;
    color: white !important;
    box-shadow: 0 4px 14px rgba(45,96,148,0.40);
}}
</style>
"""


def apply_theme() -> None:
    """Injecte le CSS global du design system dans la page courante."""
    st.markdown(_CSS, unsafe_allow_html=True)


def back_button(label: str = "← Revenir à l'accueil", target: str = "/") -> None:
    """Bouton retour HTML — texte toujours sur une seule ligne."""
    st.markdown(
        f'<a href="{target}" class="back-btn" target="_self">{label}</a>',
        unsafe_allow_html=True,
    )


def page_header(title: str, subtitle: str = "") -> None:
    """Affiche le bandeau d'en-tête standardisé d'une page."""
    apply_theme()
    sub = f"<p>{subtitle}</p>" if subtitle else ""
    st.markdown(
        f'<div class="page-header"><h1>{title}</h1>{sub}</div>',
        unsafe_allow_html=True,
    )

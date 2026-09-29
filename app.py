

import re
import sqlite3
import json
from datetime import datetime
from urllib.parse import quote_plus

import pandas as pd
import requests
import streamlit as st
from docx import Document
from pypdf import PdfReader
from io import BytesIO

st.set_page_config(
    page_title="Buscador Laboral Paulina",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================================
# ESTILO EJECUTIVO
# ==========================================================
st.markdown("""
<style>
    :root{
        --navy:#123B63;
        --navy2:#0D2F4F;
        --blue:#1D6FA5;
        --blue-soft:#EAF3FA;
        --gold:#D69222;
        --gold-soft:#FFF2CF;
        --green:#138A5B;
        --green-soft:#DDF5E8;
        --text:#17324D;
        --muted:#6D7D8D;
        --line:#E1E8EF;
        --bg:#F7F9FC;
        --card:#FFFFFF;
    }

    /* Barra superior de Streamlit: elimina el negro */
    [data-testid="stHeader"] {
        background: #F7F9FC !important;
        border-bottom: 1px solid #E1E8EF;
    }

    [data-testid="stToolbar"] {
        background: transparent !important;
    }

    [data-testid="stDecoration"] {
        display: none;
    }

    header[data-testid="stHeader"] {
        color: #123B63 !important;
    }

    .stApp {
        background: var(--bg);
        color: var(--text);
    }

    .block-container {
        max-width: 1400px;
        padding-top: 1.35rem;
        padding-bottom: 3rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }

    [data-testid="stSidebar"] {
        background: #F4F7FA;
        border-right: 1px solid var(--line);
    }

    /* Mejora visibilidad de textos en barra lateral */
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div {
        color: #394B5A;
    }

    [data-testid="stSidebar"] .stCheckbox label,
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stSlider label {
        color: #394B5A !important;
        font-weight: 600;
    }

    [data-testid="stSidebar"] [data-baseweb="select"] *,
    [data-testid="stSidebar"] [role="slider"] {
        color: #394B5A !important;
    }

    [data-testid="stSidebar"] .st-bq,
    [data-testid="stSidebar"] .st-br,
    [data-testid="stSidebar"] .st-bs,
    [data-testid="stSidebar"] .st-bt {
        color: #394B5A !important;
    }

    h1,h2,h3,h4 {
        color: var(--navy2) !important;
        letter-spacing: -0.02em;
    }

    .brand {
        display:flex;
        align-items:center;
        gap:12px;
        margin-bottom:18px;
    }

    .brand-icon {
        width:42px;height:42px;
        border-radius:11px;
        background:linear-gradient(135deg,var(--gold),#B97A16);
        display:flex;
        align-items:center;
        justify-content:center;
        color:white;
        font-size:22px;
        box-shadow:0 5px 12px rgba(214,146,34,.22);
    }

    .brand-title {
        font-size:1.25rem;
        font-weight:800;
        color:var(--navy2);
        line-height:1.1;
    }

    .nav-active {
        background:linear-gradient(90deg,var(--navy),#175482);
        color:white;
        border-radius:10px;
        padding:11px 14px;
        font-weight:700;
        margin-bottom:7px;
    }

    .nav-item {
        color:var(--text);
        padding:10px 14px;
        margin-bottom:3px;
        border-radius:10px;
    }

    .sidebar-section {
        font-size:.9rem;
        font-weight:800;
        color:var(--navy2);
        margin-top:20px;
        margin-bottom:8px;
    }

    .hero {
        border:1px solid #E8DCC8;
        background:linear-gradient(90deg,#FFF9F0 0%,#FFF5E7 52%,#EFE5D5 100%);
        border-radius:18px;
        padding:22px 24px;
        box-shadow:0 5px 18px rgba(28,53,79,.06);
        margin-bottom:16px;
    }

    .hero h2 {
        margin:0 0 6px 0;
        font-size:1.25rem;
    }

    .hero p {
        margin:0;
        color:#4E6276;
    }

    div.stButton > button[kind="primary"] {
        background:linear-gradient(90deg,var(--navy),#155B89);
        color:white;
        border:0;
        border-radius:10px;
        min-height:46px;
        font-weight:800;
        box-shadow:0 5px 12px rgba(18,59,99,.16);
    }

    div.stButton > button[kind="primary"]:hover {
        background:linear-gradient(90deg,var(--navy2),var(--navy));
        color:white;
    }

    .metric-card {
        background:#FFFFFF;
        border:1px solid #E1E8EF;
        border-radius:15px;
        padding:15px 17px;
        min-height:98px;
        box-shadow:0 4px 14px rgba(28,53,79,.05);
        display:flex;
        flex-direction:column;
        justify-content:center;
    }

    .metric-label {
        color:#6D7D8D;
        font-size:.82rem;
        font-weight:700;
        margin-bottom:7px;
    }

    .metric-value {
        color:#0D2F4F;
        font-size:1.62rem;
        line-height:1.05;
        font-weight:800;
        white-space:nowrap;
    }

    .metric-value.salary {
        font-size:1.42rem;
        letter-spacing:-0.03em;
    }

    .metric-sub {
        color:#138A5B;
        font-size:.76rem;
        margin-top:5px;
        font-weight:700;
    }

    div[data-baseweb="tab-list"] {
        gap:8px;
        border-bottom:1px solid var(--line);
    }

    button[data-baseweb="tab"] {
        background:transparent;
        color:#52677C;
        border-radius:0;
        font-weight:700;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color:var(--navy);
        border-bottom:3px solid var(--navy);
    }

    [data-testid="stVerticalBlockBorderWrapper"] {
        background:var(--card);
        border:1px solid var(--line) !important;
        border-radius:15px !important;
        box-shadow:0 4px 14px rgba(28,53,79,.045);
    }

    [data-testid="stLinkButton"] a {
        background:linear-gradient(90deg,var(--navy),#155B89);
        color:white !important;
        border:0;
        border-radius:9px;
        font-weight:800;
    }

    .badge-green {
        display:inline-block;
        background:var(--green-soft);
        color:var(--green);
        border-radius:8px;
        padding:4px 9px;
        font-weight:800;
        font-size:.82rem;
    }

    .badge-gold {
        display:inline-block;
        background:#FFF2CF;
        color:#875E08;
        border-radius:8px;
        padding:3px 8px;
        font-weight:800;
        font-size:.78rem;
    }

    .badge-blue {
        display:inline-block;
        background:var(--blue-soft);
        color:var(--navy);
        border-radius:999px;
        padding:4px 9px;
        font-size:.79rem;
        margin-right:5px;
        margin-top:4px;
    }

    .muted {
        color:var(--muted);
        font-size:.9rem;
    }

    .job-title {
        font-size:1.12rem;
        font-weight:800;
        color:var(--navy2);
        margin-bottom:2px;
    }

    .job-meta {
        color:#4F6376;
        font-size:.91rem;
        margin-bottom:6px;
    }

    .salary-line {
        color:#314A61;
        font-size:.9rem;
        margin-bottom:6px;
    }
</style>
""", unsafe_allow_html=True)

DEFAULT_PROFILE = {
    "roles": [
        "Ingeniera de Procesos","Analista Senior de Procesos","Mejora Continua",
        "Excelencia Operacional","Experiencia Cliente","Customer Experience",
        "Business Process Analyst","Analista de Operaciones","PMO",
        "Control de Gestión","Business Intelligence","KPIs"
    ],
    "skills": [
        "Power BI","Excel","SQL","Bizagi","BPMN","Lean","Kaizen",
        "AS IS","TO BE","KPIs","automatización","mejora continua",
        "customer journey","gestión de procesos","reportes ejecutivos",
        "coordinación transversal","stakeholders"
    ]
}

DEFAULT_SEARCH_QUERIES = [
    "analista senior procesos",
    "ingeniero procesos mejora continua",
    "excelencia operacional",
    "control de gestion power bi",
    "business process analyst",
]

SKILL_CATALOG = [
    "Power BI","Excel","SQL","Bizagi","BPMN","BPM","Lean","Kaizen",
    "AS IS","TO BE","KPIs","automatización","RPA","mejora continua",
    "customer journey","experiencia cliente","gestión de procesos",
    "reportes ejecutivos","stakeholders","gestión del cambio",
    "Design Thinking","Agile","MS Project","Salesforce","Visio",
    "Python","Tableau","SAP","control de gestión","PMO"
]

ROLE_RULES = [
    (["proceso","bpmn","as is","to be","mejora continua"], [
        "analista senior procesos",
        "ingeniero procesos mejora continua",
        "excelencia operacional"
    ]),
    (["control de gestión","control de gestion","kpi","indicadores"], [
        "analista control de gestion power bi"
    ]),
    (["power bi","business intelligence","bi ","sql"], [
        "analista business intelligence power bi"
    ]),
    (["experiencia cliente","customer experience","customer journey","nps","csat"], [
        "analista experiencia cliente procesos"
    ]),
    (["pmo","proyecto","project"], [
        "pmo proyectos mejora continua"
    ]),
    (["operaciones","operacional","operational"], [
        "analista operaciones mejora continua"
    ]),
    (["transformación digital","transformacion digital","automatización","automatizacion","rpa"], [
        "transformacion digital procesos automatizacion"
    ]),
]

def extract_cv_text(uploaded_file):
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()

    if name.endswith(".docx"):
        doc = Document(BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

    if name.endswith(".pdf"):
        reader = PdfReader(BytesIO(data))
        return "\n".join((page.extract_text() or "") for page in reader.pages)

    if name.endswith(".txt"):
        return data.decode("utf-8", errors="ignore")

    return ""

def profile_from_cv(cv_text):
    low = cv_text.lower()

    skills = []
    for skill in SKILL_CATALOG:
        if skill.lower() in low and skill not in skills:
            skills.append(skill)

    queries = []
    for triggers, suggested in ROLE_RULES:
        if any(t in low for t in triggers):
            for q in suggested:
                if q not in queries:
                    queries.append(q)

    if not queries:
        queries = DEFAULT_SEARCH_QUERIES.copy()

    # Mantener la búsqueda acotada para no consumir demasiadas consultas API.
    queries = queries[:7]

    # Los cargos sirven también para el cálculo de compatibilidad.
    roles = [q.title() for q in queries]

    return {
        "roles": roles,
        "skills": skills if skills else DEFAULT_PROFILE["skills"].copy()
    }, queries

def extra_skills():
    raw = st.session_state.get("extra_skills_editor", "")
    parts = re.split(r"[,;\n]+", raw)
    return [p.strip() for p in parts if p.strip()]

def active_profile():
    base = st.session_state.get("active_profile", DEFAULT_PROFILE)
    profile = {
        "roles": list(base.get("roles", [])),
        "skills": list(base.get("skills", []))
    }

    # Las habilidades escritas manualmente impactan el % de ajuste.
    for skill in extra_skills():
        if skill.lower() not in [x.lower() for x in profile["skills"]]:
            profile["skills"].append(skill)

    return profile

def selected_work_modes():
    modes = []
    if st.session_state.get("filter_hybrid", True):
        modes.append("Híbrido")
    if st.session_state.get("filter_remote", True):
        modes.append("Remoto")
    if st.session_state.get("filter_onsite", False):
        modes.append("Presencial / no informado")
    return modes

def active_queries():
    raw = st.session_state.get("query_editor", "")
    base_queries = [q.strip() for q in raw.splitlines() if q.strip()]
    if not base_queries:
        base_queries = DEFAULT_SEARCH_QUERIES.copy()

    queries = list(base_queries)

    # Si el usuario agrega nuevas habilidades, se agregan búsquedas nuevas.
    # Se limita la cantidad para no consumir excesivamente la API.
    skills = extra_skills()[:4]
    for skill in skills:
        skill_low = skill.lower()

        # Búsqueda especializada por habilidad.
        q = f"{skill} procesos"
        if q.lower() not in [x.lower() for x in queries]:
            queries.append(q)

        # Para habilidades BI / datos, agregar una consulta más específica.
        if any(k in skill_low for k in ["power bi", "sql", "tableau", "python", "data", "datos"]):
            q2 = f"analista {skill}"
            if q2.lower() not in [x.lower() for x in queries]:
                queries.append(q2)

    # La modalidad seleccionada también influye en la búsqueda.
    modes = selected_work_modes()
    mode_suffix = ""
    if modes == ["Remoto"]:
        mode_suffix = " remoto"
    elif modes == ["Híbrido"]:
        mode_suffix = " híbrido"
    elif modes == ["Presencial / no informado"]:
        mode_suffix = " presencial"

    if mode_suffix:
        queries = [
            q if mode_suffix.strip().lower() in q.lower()
            else f"{q}{mode_suffix}"
            for q in queries
        ]

    # Quitar duplicados preservando orden.
    unique = []
    seen = set()
    for q in queries:
        key = q.lower().strip()
        if key and key not in seen:
            seen.add(key)
            unique.append(q.strip())

    return unique[:12]


TARGET_PLUS = 2_500_000

# Referencias salariales Chile.
# Se usan SOLO cuando el aviso no publica una renta.
# Las cifras son referencias de mercado y no una promesa de sueldo.
MARKET_BENCHMARKS = {
    "procesos": {
        "avg": 1_163_306,
        "low": 685_565,
        "high": 1_973_965,
        "label": "Analista de procesos · Santiago",
        "source": "Indeed Chile",
        "url": "https://cl.indeed.com/career/analista-de-procesos/salaries/Santiago-de-Chile--Regi%C3%B3n-Metropolitana"
    },
    "control_gestion": {
        "avg": 906_005,
        "low": 508_749,
        "high": 1_613_458,
        "label": "Analista de control de gestión · Santiago",
        "source": "Indeed Chile",
        "url": "https://cl.indeed.com/career/analista-de-control-de-gesti%C3%B3n/salaries/Santiago-de-Chile--Regi%C3%B3n-Metropolitana"
    },
    "bi": {
        "avg": 1_316_462,
        "low": 902_055,
        "high": 1_921_250,
        "label": "Analista Business Intelligence · Santiago",
        "source": "Indeed Chile",
        "url": "https://cl.indeed.com/career/analista-business-intelligence/salaries/Santiago-de-Chile--Regi%C3%B3n-Metropolitana"
    },
    "operaciones": {
        "avg": 718_827,
        "low": 423_724,
        "high": 1_422_362,
        "label": "Analista de operaciones · Región Metropolitana",
        "source": "Indeed Chile",
        "url": "https://cl.indeed.com/career/analista-de-operaciones/salaries/Regi%C3%B3n-Metropolitana"
    },
}

def money_clp(value):
    if value is None:
        return "—"
    return "$" + f"{int(value):,}".replace(",", ".")

def safe_money_html(value):
    """
    Evita que Streamlit interprete montos con $...$ como LaTeX.
    Convierte $ a entidad HTML y normaliza el separador del rango.
    """
    if value is None:
        return "—"
    txt = str(value)
    txt = txt.replace("$", "&#36;")
    txt = txt.replace(" - ", " – ")
    return txt

def round_50k(value):
    return int(round(value / 50_000) * 50_000)

def market_key_for_job(title, description=""):
    txt = norm(f"{title} {description}")

    if any(k in txt for k in [
        "business intelligence", "power bi", "bi analyst",
        "inteligencia de negocios"
    ]):
        return "bi"

    if any(k in txt for k in [
        "control de gestión", "control de gestion",
        "controller", "gestión comercial", "gestion comercial"
    ]):
        return "control_gestion"

    if any(k in txt for k in [
        "operaciones", "operational", "operations analyst"
    ]):
        return "operaciones"

    if any(k in txt for k in [
        "proceso", "process", "mejora continua",
        "excelencia operacional", "bpmn", "transformación"
    ]):
        return "procesos"

    return None

def salary_bounds(text):
    """Devuelve (mínimo, máximo) mensual CLP cuando puede interpretarlo."""
    if not text:
        return None, None

    s = str(text).lower().replace("\xa0", " ")

    # Detecta si la publicación está expresada por año.
    annual = any(x in s for x in ["por año", "al año", "/año", "anual"])

    # 1,5 millones / 2.5 millones
    million_vals = []
    for a, b in re.findall(r"(\d{1,2})[\.,](\d)\s*(?:mill[oó]n|millones|m)\b", s):
        million_vals.append(int(float(f"{a}.{b}") * 1_000_000))

    for a in re.findall(r"(\d{1,2})\s*(?:mill[oó]n|millones)\b", s):
        million_vals.append(int(a) * 1_000_000)

    # $1.500.000 / 1500000
    raw_vals = []
    for raw in re.findall(r"\$?\s*(\d[\d\.\,]{5,})", s):
        digits = re.sub(r"\D", "", raw)
        if digits.isdigit():
            v = int(digits)
            if 500_000 <= v <= 100_000_000:
                raw_vals.append(v)

    vals = million_vals + raw_vals

    if not vals:
        return None, None

    # Quitar duplicados y ordenar.
    vals = sorted(set(vals))

    low = vals[0]
    high = vals[-1]

    # Convertir valores anuales a mensuales.
    if annual:
        low = int(low / 12)
        high = int(high / 12)

    return low, high

def suggested_salary(title, description, published_text, user_min, user_max):
    """
    Genera una pretensión orientativa.
    Prioridad:
    1. Renta publicada en el aviso.
    2. Referencia de mercado Chile.
    3. Rango objetivo definido por la usuaria.
    """
    pub_low, pub_high = salary_bounds(published_text)

    if pub_low is not None:
        # Si hay banda publicada, sugerimos posicionarse entre el 70% y 90%
        # del recorrido de la banda, sin superar el máximo informado.
        if pub_high and pub_high > pub_low:
            low = pub_low + (pub_high - pub_low) * 0.65
            high = pub_low + (pub_high - pub_low) * 0.90
        else:
            low = pub_low * 0.95
            high = pub_low * 1.05

        low = round_50k(low)
        high = round_50k(high)

        if user_min:
            low = max(low, user_min)
        if user_max:
            high = min(high, user_max)

        if high < low:
            # El aviso está bajo el objetivo configurado.
            return (
                f"{money_clp(pub_low)} – {money_clp(pub_high or pub_low)}",
                "Renta del aviso bajo tu rango objetivo",
                "Aviso publicado"
            )

        return (
            f"{money_clp(low)} – {money_clp(high)}",
            "Basada en la banda salarial publicada",
            "Aviso publicado"
        )

    key = market_key_for_job(title, description)

    if key and key in MARKET_BENCHMARKS:
        ref = MARKET_BENCHMARKS[key]

        # Para un perfil con experiencia, se utiliza la zona media-alta
        # de la referencia, respetando el piso definido en la app.
        low = max(user_min or 0, round_50k(max(ref["avg"] * 1.10, ref["low"])))
        high_market = round_50k(ref["high"] * 0.95)

        if "senior" in norm(title) or "especialista" in norm(title) or "ingeniero" in norm(title):
            low = max(low, round_50k(ref["avg"] * 1.20))
            high_market = max(high_market, round_50k(ref["avg"] * 1.45))

        high = high_market
        if user_max:
            high = min(high, user_max)

        if high < low:
            high = low

        return (
            f"{money_clp(low)} – {money_clp(high)}",
            f"Mercado de {ref['label']}",
            ref["source"]
        )

    # Último recurso: usar el rango objetivo configurado.
    fallback_high = user_max if user_max else max(user_min + 500_000, 2_500_000)
    return (
        f"{money_clp(user_min)} – {money_clp(fallback_high)}",
        "Estimación basada en tu rango objetivo; sin referencia salarial específica para este cargo",
        "Perfil de búsqueda"
    )

@st.cache_resource
def get_db():
    conn = sqlite3.connect("jobs.db", check_same_thread=False)

    # Tabla histórica de ofertas. Conserva resultados entre búsquedas.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS tracked_jobs(
            job_key TEXT PRIMARY KEY,
            payload_json TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Nueva',
            first_seen TEXT NOT NULL,
            last_seen TEXT NOT NULL,
            saved_at TEXT,
            applied_at TEXT,
            discarded_at TEXT,
            updated_at TEXT NOT NULL
        )
    """)

    # Se mantiene para compatibilidad con versiones anteriores.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS job_status(
            job_key TEXT PRIMARY KEY,
            status TEXT,
            updated_at TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS saved_searches(
            search_key TEXT PRIMARY KEY,
            cargo TEXT,
            fuente TEXT,
            ubicacion TEXT,
            enlace TEXT,
            pretension TEXT,
            referencia TEXT,
            saved_at TEXT
        )
    """)

    conn.commit()
    return conn

DB = get_db()

def now_iso():
    return datetime.now().isoformat(timespec="seconds")

def get_status(key):
    row = DB.execute(
        "SELECT status FROM tracked_jobs WHERE job_key=?",
        (key,)
    ).fetchone()

    if row:
        return row[0]

    # Compatibilidad con registros de versiones anteriores.
    old = DB.execute(
        "SELECT status FROM job_status WHERE job_key=?",
        (key,)
    ).fetchone()

    return old[0] if old else "Nueva"

def get_job_dates(key):
    row = DB.execute(
        """
        SELECT first_seen, saved_at, applied_at, discarded_at, updated_at
        FROM tracked_jobs
        WHERE job_key=?
        """,
        (key,)
    ).fetchone()

    if not row:
        return {
            "first_seen": None,
            "saved_at": None,
            "applied_at": None,
            "discarded_at": None,
            "updated_at": None,
        }

    return {
        "first_seen": row[0],
        "saved_at": row[1],
        "applied_at": row[2],
        "discarded_at": row[3],
        "updated_at": row[4],
    }

def set_status(key, status):
    stamp = now_iso()

    row = DB.execute(
        "SELECT status, saved_at, applied_at, discarded_at FROM tracked_jobs WHERE job_key=?",
        (key,)
    ).fetchone()

    if not row:
        # Si la oferta todavía no fue persistida, solo conserva estado legacy.
        DB.execute("""
            INSERT INTO job_status(job_key,status,updated_at)
            VALUES(?,?,?)
            ON CONFLICT(job_key)
            DO UPDATE SET status=excluded.status, updated_at=excluded.updated_at
        """, (key,status,stamp))
        DB.commit()
        return

    old_status, saved_at, applied_at, discarded_at = row

    if status == "Guardada" and not saved_at:
        saved_at = stamp

    if status == "Postulada" and not applied_at:
        applied_at = stamp

    if status == "Descartada" and not discarded_at:
        discarded_at = stamp

    # Si vuelve a Nueva, no borramos fechas históricas.
    DB.execute("""
        UPDATE tracked_jobs
        SET status=?,
            saved_at=?,
            applied_at=?,
            discarded_at=?,
            updated_at=?
        WHERE job_key=?
    """, (
        status,
        saved_at,
        applied_at,
        discarded_at,
        stamp,
        key
    ))

    DB.execute("""
        INSERT INTO job_status(job_key,status,updated_at)
        VALUES(?,?,?)
        ON CONFLICT(job_key)
        DO UPDATE SET status=excluded.status, updated_at=excluded.updated_at
    """, (key,status,stamp))

    DB.commit()

def persist_jobs(df):
    """Acumula ofertas reales encontradas sin eliminar resultados anteriores."""
    if df is None or df.empty:
        return

    stamp = now_iso()

    for _, row in df.iterrows():
        if bool(row.get("_fallback", False)):
            continue

        key = row["_key"]

        payload = {
            "Cargo": row.get("Cargo", ""),
            "Empresa": row.get("Empresa", ""),
            "Ubicación": row.get("Ubicación", ""),
            "Fuente": row.get("Fuente", ""),
            "Score": int(row.get("Score", 0) or 0),
            "Renta": row.get("Renta", ""),
            "RentaValor": (
                None if pd.isna(row.get("RentaValor"))
                else int(row.get("RentaValor"))
            ),
            "Pretension": row.get("Pretension", ""),
            "PretensionNota": row.get("PretensionNota", ""),
            "PretensionFuente": row.get("PretensionFuente", ""),
            "ReferenciaURL": row.get("ReferenciaURL", ""),
            "Modalidad": row.get("Modalidad", ""),
            "Descripción": row.get("Descripción", ""),
            "Skills": row.get("Skills", []) or [],
            "Enlace": row.get("Enlace", ""),
        }

        existing = DB.execute(
            "SELECT status, first_seen FROM tracked_jobs WHERE job_key=?",
            (key,)
        ).fetchone()

        if existing:
            DB.execute("""
                UPDATE tracked_jobs
                SET payload_json=?, last_seen=?, updated_at=?
                WHERE job_key=?
            """, (
                json.dumps(payload, ensure_ascii=False),
                stamp,
                stamp,
                key
            ))
        else:
            legacy_status = get_status(key)
            DB.execute("""
                INSERT INTO tracked_jobs(
                    job_key, payload_json, status,
                    first_seen, last_seen, updated_at
                )
                VALUES(?,?,?,?,?,?)
            """, (
                key,
                json.dumps(payload, ensure_ascii=False),
                legacy_status,
                stamp,
                stamp,
                stamp
            ))

    DB.commit()

def load_tracked_jobs(status=None):
    if status:
        rows = DB.execute("""
            SELECT job_key, payload_json, status,
                   first_seen, last_seen, saved_at,
                   applied_at, discarded_at, updated_at
            FROM tracked_jobs
            WHERE status=?
            ORDER BY updated_at DESC
        """, (status,)).fetchall()
    else:
        rows = DB.execute("""
            SELECT job_key, payload_json, status,
                   first_seen, last_seen, saved_at,
                   applied_at, discarded_at, updated_at
            FROM tracked_jobs
            ORDER BY last_seen DESC
        """).fetchall()

    data = []

    for row in rows:
        payload = json.loads(row[1])

        # ----------------------------------------------------------
        # Migración/enriquecimiento para ofertas guardadas por
        # versiones anteriores de la app.
        # ----------------------------------------------------------
        cargo = payload.get("Cargo", "")
        descripcion = payload.get("Descripción", "") or ""

        renta = payload.get("Renta", "") or ""
        renta_valor = payload.get("RentaValor", None)

        # Si antes no quedó guardada una renta legible, intentamos
        # recuperarla desde descripción/renta histórica.
        if not renta or renta == "—":
            low, high = salary_bounds(descripcion)
            if low is not None:
                if high and high != low:
                    renta = f"{money_clp(low)} – {money_clp(high)}"
                    renta_valor = high
                else:
                    renta = money_clp(low)
                    renta_valor = low
            else:
                renta = "Sin renta publicada"

        # Asegura que haya un valor numérico utilizable.
        if renta_valor is None and renta != "Sin renta publicada":
            low, high = salary_bounds(renta)
            renta_valor = high or low

        # Recalcula pretensión si el registro antiguo no la tenía.
        pretension = payload.get("Pretension", "") or ""
        pretension_note = payload.get("PretensionNota", "") or ""
        pretension_source = payload.get("PretensionFuente", "") or ""

        if not pretension:
            current_min = salary_min if "salary_min" in globals() else 1_200_000
            current_max = salary_max if "salary_max" in globals() else 2_500_000

            pretension, pretension_note, pretension_source = suggested_salary(
                cargo,
                descripcion,
                renta if renta != "Sin renta publicada" else "",
                current_min,
                current_max
            )

        market_key = market_key_for_job(cargo, descripcion)
        ref_url = payload.get("ReferenciaURL", "") or ""
        if not ref_url and market_key in MARKET_BENCHMARKS:
            ref_url = MARKET_BENCHMARKS[market_key]["url"]

        payload.update({
            "Renta": renta,
            "RentaValor": renta_valor,
            "Pretension": pretension,
            "PretensionNota": pretension_note,
            "PretensionFuente": pretension_source,
            "ReferenciaURL": ref_url,
            "_key": row[0],
            "Estado": row[2],
            "FechaEncontrada": row[3],
            "UltimaVezVista": row[4],
            "FechaGuardada": row[5],
            "FechaPostulacion": row[6],
            "FechaDescarte": row[7],
            "FechaActualizacion": row[8],
            "_fallback": False,
        })

        data.append(payload)

    if not data:
        return pd.DataFrame()

    return pd.DataFrame(data)

def tracked_counts():
    total = DB.execute(
        "SELECT COUNT(*) FROM tracked_jobs"
    ).fetchone()[0]

    counts = {"Resultados": total}

    for status, label in [
        ("Guardada", "Guardadas"),
        ("Postulada", "Postuladas"),
        ("Descartada", "Descartadas"),
    ]:
        counts[label] = DB.execute(
            "SELECT COUNT(*) FROM tracked_jobs WHERE status=?",
            (status,)
        ).fetchone()[0]

    return counts

def persist_fallback_as_job(row, status):
    key = f"fallback|{row.get('Cargo','')}|{row.get('Empresa','')}|{row.get('Fuente','')}"
    stamp = now_iso()

    payload = {
        "Cargo": row.get("Cargo",""),
        "Empresa": row.get("Empresa",""),
        "Ubicación": row.get("Ubicación",""),
        "Fuente": row.get("Fuente",""),
        "Score": int(row.get("Score", 0) or 0),
        "Renta": row.get("Renta","Sin renta publicada"),
        "RentaValor": (
            None if pd.isna(row.get("RentaValor"))
            else int(row.get("RentaValor"))
        ),
        "Pretension": row.get("Pretension",""),
        "PretensionNota": row.get("PretensionNota",""),
        "PretensionFuente": row.get("PretensionFuente",""),
        "ReferenciaURL": row.get("ReferenciaURL",""),
        "Modalidad": row.get("Modalidad",""),
        "Descripción": row.get("Descripción",""),
        "Skills": row.get("Skills",[]) or [],
        "Enlace": row.get("Enlace",""),
    }

    existing = DB.execute(
        "SELECT first_seen, saved_at, applied_at, discarded_at FROM tracked_jobs WHERE job_key=?",
        (key,)
    ).fetchone()

    if existing:
        first_seen, saved_at, applied_at, discarded_at = existing
    else:
        first_seen = stamp
        saved_at = applied_at = discarded_at = None

    if status == "Guardada" and not saved_at:
        saved_at = stamp
    if status == "Postulada" and not applied_at:
        applied_at = stamp
    if status == "Descartada" and not discarded_at:
        discarded_at = stamp

    DB.execute("""
        INSERT INTO tracked_jobs(
            job_key, payload_json, status,
            first_seen, last_seen, saved_at,
            applied_at, discarded_at, updated_at
        )
        VALUES(?,?,?,?,?,?,?,?,?)
        ON CONFLICT(job_key)
        DO UPDATE SET
            payload_json=excluded.payload_json,
            status=excluded.status,
            last_seen=excluded.last_seen,
            saved_at=COALESCE(tracked_jobs.saved_at, excluded.saved_at),
            applied_at=COALESCE(tracked_jobs.applied_at, excluded.applied_at),
            discarded_at=COALESCE(tracked_jobs.discarded_at, excluded.discarded_at),
            updated_at=excluded.updated_at
    """, (
        key,
        json.dumps(payload, ensure_ascii=False),
        status,
        first_seen,
        stamp,
        saved_at,
        applied_at,
        discarded_at,
        stamp
    ))
    DB.commit()

def save_fallback_search(row):
    key = f"{row.get('Cargo','')}|{row.get('Empresa','')}|{row.get('Fuente','')}"
    note = (row.get("PretensionNota", "") or "").replace("Referencia:", "").strip()
    source = row.get("PretensionFuente", "") or ""
    referencia = f"{note} — {source}".strip(" —")

    DB.execute("""
        INSERT INTO saved_searches(
            search_key, cargo, fuente, ubicacion,
            enlace, pretension, referencia, saved_at
        )
        VALUES(?,?,?,?,?,?,?,?)
        ON CONFLICT(search_key)
        DO UPDATE SET
            cargo=excluded.cargo,
            fuente=excluded.fuente,
            ubicacion=excluded.ubicacion,
            enlace=excluded.enlace,
            pretension=excluded.pretension,
            referencia=excluded.referencia,
            saved_at=excluded.saved_at
    """, (
        key,
        row.get("Cargo",""),
        row.get("Fuente",""),
        row.get("Ubicación",""),
        row.get("Enlace",""),
        row.get("Pretension",""),
        referencia,
        now_iso()
    ))
    DB.commit()

def load_saved_searches():
    rows = DB.execute("""
        SELECT search_key, cargo, fuente, ubicacion,
               enlace, pretension, referencia, saved_at
        FROM saved_searches
        ORDER BY saved_at DESC
    """).fetchall()

    return [
        {
            "search_key": r[0],
            "cargo": r[1],
            "fuente": r[2],
            "ubicacion": r[3],
            "enlace": r[4],
            "pretension": r[5],
            "referencia": r[6],
            "saved_at": r[7],
        }
        for r in rows
    ]

def delete_saved_search(search_key):
    DB.execute("DELETE FROM saved_searches WHERE search_key=?", (search_key,))
    DB.commit()

def applications_excel_bytes(df):
    """Genera Excel de postulaciones para descarga."""
    if df is None or df.empty:
        return None

    export = pd.DataFrame({
        "Fecha postulación": df["FechaPostulacion"].fillna(""),
        "Cargo": df["Cargo"].fillna(""),
        "Empresa": df["Empresa"].fillna(""),
        "Ubicación": df["Ubicación"].fillna(""),
        "Modalidad": df["Modalidad"].fillna(""),
        "% ajuste CV": df["Score"].fillna(0),
        "Renta publicada": df["Renta"].fillna(""),
        "Pretensión sugerida": df["Pretension"].fillna(""),
        "Detalle pretensión": df["PretensionNota"].fillna(""),
        "Fuente salarial": df["PretensionFuente"].fillna(""),
        "Fuente oferta": df["Fuente"].fillna(""),
        "Estado": df["Estado"].fillna(""),
        "Fecha encontrada": df["FechaEncontrada"].fillna(""),
        "Última actualización": df["FechaActualizacion"].fillna(""),
        "Enlace": df["Enlace"].fillna(""),
    })

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export.to_excel(
            writer,
            sheet_name="Postulaciones",
            index=False
        )

        workbook = writer.book
        sheet = writer.sheets["Postulaciones"]

        header_fmt = workbook.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#123B63",
            "border": 0,
            "align": "center",
            "valign": "vcenter",
        })

        date_fmt = workbook.add_format({
            "num_format": "dd-mm-yyyy hh:mm",
        })

        pct_fmt = workbook.add_format({
            "num_format": '0"%"',
            "align": "center",
        })

        for col_num, value in enumerate(export.columns.values):
            sheet.write(0, col_num, value, header_fmt)

        widths = {
            0: 20, 1: 34, 2: 26, 3: 25, 4: 20,
            5: 13, 6: 22, 7: 24, 8: 35, 9: 22, 10: 18,
            11: 16, 12: 20, 13: 20, 14: 55
        }

        for col, width in widths.items():
            sheet.set_column(col, col, width)

        sheet.freeze_panes(1, 0)
        sheet.autofilter(0, 0, len(export), len(export.columns)-1)

        # Formato condicional sobre ajuste.
        sheet.conditional_format(
            1, 5, len(export), 5,
            {
                "type": "3_color_scale",
                "min_color": "#F8D7DA",
                "mid_color": "#FFF3CD",
                "max_color": "#D1E7DD",
            }
        )

        sheet.set_column(5, 5, 13, pct_fmt)

    output.seek(0)
    return output.getvalue()

def norm(t):
    return re.sub(r"\s+"," ",(t or "").lower()).strip()

def score_job(title,company,desc):
    text = norm(f"{title} {company} {desc}")
    score = 35
    profile = active_profile()
    score += min(sum(norm(s) in text for s in profile["skills"]) * 4,36)

    hits = 0
    for role in profile["roles"]:
        words = [w for w in norm(role).split() if len(w)>3]
        if words and sum(w in text for w in words) >= max(1,len(words)//2):
            hits += 1

    score += min(hits*6,24)

    if any(x in text for x in ["senior","sr.","especialista","ingeniero","ingeniera"]):
        score += 4

    if any(x in text for x in ["práctica","practica","intern","vendedor","promotor","operario"]):
        score -= 30

    return max(0,min(score,100))

def parse_salary(text):
    if not text:
        return None

    s = str(text).lower().replace("\xa0"," ")

    m = re.search(r"(\d{1,2})[\.,](\d)\s*(?:mill[oó]n|millones|m)\b",s)
    if m:
        return int(float(f"{m.group(1)}.{m.group(2)}")*1_000_000)

    m = re.search(r"(\d{1,2})\s*(?:mill[oó]n|millones)\b",s)
    if m:
        return int(m.group(1))*1_000_000

    vals = []
    for raw in re.findall(r"\$?\s*(\d[\d\.\,]{5,})",s):
        d = re.sub(r"\D","",raw)
        if d.isdigit():
            v = int(d)
            if 500_000 <= v <= 20_000_000:
                vals.append(v)

    return max(vals) if vals else None

def salary_data(job):
    ext = job.get("detected_extensions") or {}

    if isinstance(ext,dict) and ext.get("salary"):
        txt = str(ext["salary"])
        low, high = salary_bounds(txt)
        val = high or low
        return txt, val

    text = " ".join(job.get("extensions",[]) or []) + " " + job.get("description","")
    low, high = salary_bounds(text)

    if low is not None:
        if high and high != low:
            return f"{money_clp(low)} – {money_clp(high)}", high
        return money_clp(low), low

    return "Sin renta publicada",None

def detect_work_mode(job):
    text = norm(
        f"{job.get('title','')} {job.get('location','')} "
        f"{job.get('description','')} {' '.join(job.get('extensions',[]) or [])}"
    )

    remote_terms = [
        "remoto", "remote", "work from home", "home office",
        "teletrabajo", "telecommute", "100% remoto", "100% remote"
    ]
    hybrid_terms = [
        "híbrido", "hibrido", "hybrid", "modalidad mixta",
        "semipresencial", "home office parcial"
    ]

    has_remote = any(term in text for term in remote_terms)
    has_hybrid = any(term in text for term in hybrid_terms)

    # Prioriza híbrido cuando el aviso explícitamente mezcla oficina y remoto.
    if has_hybrid:
        return "Híbrido"
    if has_remote:
        return "Remoto"

    return "Presencial / no informado"

def extract_skills(desc):
    text = norm(desc)
    mapping = [
        ("bpmn","BPM/BPMN"),("power bi","Power BI"),("sql","SQL"),
        ("mejora continua","Mejora Continua"),("automatiz","Automatización"),
        ("stakeholder","Stakeholders"),("kpi","KPIs"),("as is","AS IS / TO BE"),
        ("to be","AS IS / TO BE"),("lean","Lean"),("kaizen","Kaizen")
    ]

    labels = []
    for key,label in mapping:
        if key in text and label not in labels:
            labels.append(label)

    return labels[:6]

def serp_key():
    try:
        return st.secrets.get("SERPAPI_KEY","")
    except Exception:
        return ""

def search_live():
    key = serp_key()

    if not key or key == "TU_CLAVE_AQUI":
        return [],["No se encontró una SERPAPI_KEY válida."]

    jobs,errors = [],[]

    for q in active_queries():
        try:
            r = requests.get(
                "https://serpapi.com/search.json",
                params={
                    "engine":"google_jobs",
                    "q":f"{q} Santiago Chile",
                    "hl":"es",
                    "gl":"cl",
                    "api_key":key
                },
                timeout=25
            )

            if r.status_code != 200:
                try:
                    detail = r.json().get("error",r.text[:160])
                except Exception:
                    detail = r.text[:160]

                errors.append(f"{q}: HTTP {r.status_code} — {detail}")
                continue

            data = r.json()

            if data.get("error"):
                errors.append(f"{q}: {data['error']}")
                continue

            jobs.extend(data.get("jobs_results",[]) or [])

        except Exception as e:
            errors.append(f"{q}: {type(e).__name__}")

    unique,keys = [],set()

    for j in jobs:
        k = (j.get("title",""),j.get("company_name",""),j.get("location",""))
        if k not in keys:
            keys.add(k)
            unique.append(j)

    return unique,errors

def fallback_links():
    rows = []

    for q in active_queries():
        enc = quote_plus(q)
        google_indeed = quote_plus(f'site:cl.indeed.com "{q}" Santiago sueldo')
        google_chiletrabajos = quote_plus(f'site:chiletrabajos.cl/trabajo "{q}" Santiago salario')
        google_laborum = quote_plus(f'site:laborum.cl "{q}" Santiago sueldo')

        rows += [
            {
                "title":q.title(),
                "company_name":"Indeed Chile",
                "location":"Santiago / Chile",
                "via":"Indeed",
                "description":"Búsqueda prioritaria en un portal que muestra renta en parte de sus avisos y referencias salariales por cargo.",
                "job_apply_link":f"https://cl.indeed.com/jobs?q={enc}&l=Santiago%2C+Regi%C3%B3n+Metropolitana",
                "_fallback":True
            },
            {
                "title":q.title(),
                "company_name":"Chiletrabajos",
                "location":"Santiago / Chile",
                "via":"Chiletrabajos",
                "description":"Búsqueda en Chiletrabajos, donde varios avisos publican un campo de salario.",
                "job_apply_link":f"https://www.google.com/search?q={google_chiletrabajos}",
                "_fallback":True
            },
            {
                "title":q.title(),
                "company_name":"Laborum Chile",
                "location":"Santiago / Chile",
                "via":"Laborum",
                "description":"Búsqueda en Laborum y referencia de renta pretendida/salarios para Chile.",
                "job_apply_link":f"https://www.google.com/search?q={google_laborum}",
                "_fallback":True
            },
            {
                "title":q.title(),
                "company_name":"Computrabajo",
                "location":"Santiago / Chile",
                "via":"Computrabajo",
                "description":"Búsqueda complementaria de ofertas.",
                "job_apply_link":f"https://cl.computrabajo.com/trabajo-de-{q.replace(' ','-')}",
                "_fallback":True
            },
            {
                "title":q.title(),
                "company_name":"LinkedIn Jobs",
                "location":"Santiago / Chile",
                "via":"LinkedIn",
                "description":"Búsqueda complementaria; muchos avisos no publican renta.",
                "job_apply_link":f"https://www.linkedin.com/jobs/search/?keywords={enc}&location=Santiago%2C%20Chile",
                "_fallback":True
            },
        ]

    return rows

def build_df(jobs):
    rows = []

    for j in jobs:
        title = j.get("title","Sin título")
        company = j.get("company_name","Empresa no informada")
        desc = j.get("description","")
        loc = j.get("location","No informado")

        link = j.get("job_apply_link") or ""
        if not link and j.get("apply_options"):
            link = j["apply_options"][0].get("link","")

        score = score_job(title,company,desc)
        sal_txt,sal_val = salary_data(j)
        key = f"{title}|{company}|{loc}"

        # Usa el rango salarial que esté actualmente seleccionado en el sidebar.
        current_min = salary_min if "salary_min" in globals() else 1_200_000
        current_max = salary_max if "salary_max" in globals() else 2_500_000

        pretension, pretension_note, pretension_source = suggested_salary(
            title,
            desc,
            sal_txt if sal_txt != "Sin renta publicada" else "",
            current_min,
            current_max
        )

        market_key = market_key_for_job(title, desc)
        market_url = (
            MARKET_BENCHMARKS[market_key]["url"]
            if market_key in MARKET_BENCHMARKS
            else ""
        )

        rows.append({
            "Cargo":title,
            "Empresa":company,
            "Ubicación":loc,
            "Fuente":j.get("via",""),
            "Score":score,
            "Renta":sal_txt,
            "RentaValor":sal_val,
            "Pretension":pretension,
            "PretensionNota":pretension_note,
            "PretensionFuente":pretension_source,
            "ReferenciaURL":market_url,
            "Modalidad":detect_work_mode(j),
            "Descripción":desc,
            "Skills":extract_skills(desc),
            "Enlace":link,
            "Estado":get_status(key),
            "_key":key,
            "_fallback":bool(j.get("_fallback",False))
        })

    return pd.DataFrame(rows)

# ==========================================================
# SIDEBAR
# ==========================================================
with st.sidebar:
    counts = tracked_counts()

    st.markdown("""
    <div class="brand">
        <div class="brand-icon">💼</div>
        <div class="brand-title">Buscador Laboral<br>Paulina Vergara</div>
    </div>
    """, unsafe_allow_html=True)

    if "page" not in st.session_state:
        st.session_state["page"] = "Inicio"

    saved_search_count = len(load_saved_searches())

    nav_items = [
        ("🏠", "Inicio", None),
        ("🔎", "Resultados", counts["Resultados"]),
        ("⭐", "Guardadas", counts["Guardadas"] + saved_search_count),
        ("✅", "Postuladas", counts["Postuladas"]),
        ("🗑️", "Descartadas", counts["Descartadas"]),
    ]

    for icon, label, count in nav_items:
        button_label = f"{icon}  {label}"
        if count is not None:
            button_label += f"  ({count})"

        if st.button(
            button_label,
            key=f"nav_{label}",
            use_container_width=True,
            type="primary" if st.session_state["page"] == label else "secondary"
        ):
            st.session_state["page"] = label
            st.rerun()

    st.markdown('<div class="sidebar-section">CV de búsqueda</div>', unsafe_allow_html=True)


    uploaded_cv = st.file_uploader(
        "Subir nuevo CV",
        type=["pdf","docx","txt"],
        help="El CV se utiliza para detectar habilidades y ajustar las búsquedas."
    )

    if uploaded_cv is not None:
        if st.button("📄 Usar este CV", use_container_width=True):
            try:
                cv_text = extract_cv_text(uploaded_cv)
                if len(cv_text.strip()) < 80:
                    st.error("No pude extraer suficiente texto del CV.")
                else:
                    new_profile, new_queries = profile_from_cv(cv_text)
                    st.session_state["active_profile"] = new_profile
                    st.session_state["active_cv_name"] = uploaded_cv.name
                    st.session_state["query_editor"] = "\n".join(new_queries)
                    # Borra resultados anteriores para no mezclar perfiles.
                    st.session_state.pop("jobs", None)
                    st.session_state.pop("mode", None)
                    st.session_state["found_count"] = 0
                    st.success("CV aplicado a la búsqueda.")
                    st.rerun()
            except Exception as e:
                st.error(f"No pude leer el CV: {type(e).__name__}")

    current_cv = st.session_state.get("active_cv_name", "CV inicial de Paulina")
    st.caption(f"CV activo: **{current_cv}**")

    if st.button("↩️ Volver al CV inicial", use_container_width=True):
        st.session_state["active_profile"] = DEFAULT_PROFILE.copy()
        st.session_state["active_cv_name"] = "CV inicial de Paulina"
        st.session_state["query_editor"] = "\n".join(DEFAULT_SEARCH_QUERIES)
        st.session_state["extra_skills_editor"] = ""
        st.session_state.pop("jobs", None)
        st.session_state.pop("mode", None)
        st.session_state["found_count"] = 0
        st.rerun()

    # Inicializa los cargos objetivo la primera vez.
    if "query_editor" not in st.session_state:
        st.session_state["query_editor"] = "\n".join(DEFAULT_SEARCH_QUERIES)

    st.text_area(
        "Cargos / búsquedas objetivo",
        key="query_editor",
        height=135,
        help="Una búsqueda por línea. Puedes editar estas frases manualmente."
    )

    with st.expander("Ver habilidades detectadas"):
        st.write(" · ".join(active_profile()["skills"]))

    if "extra_skills_editor" not in st.session_state:
        st.session_state["extra_skills_editor"] = ""

    st.text_area(
        "Habilidades adicionales",
        key="extra_skills_editor",
        height=85,
        placeholder="Ej: Tableau, Python, Power Automate",
        help=(
            "Sepáralas por coma o por línea. Estas habilidades aumentan el "
            "alcance de la búsqueda y también influyen en el % de ajuste."
        )
    )

    if st.session_state.get("extra_skills_editor", "").strip():
        st.caption(
            "🔎 La próxima búsqueda incluirá: "
            + " · ".join(extra_skills())
        )

    st.markdown('<div class="sidebar-section">Filtros de búsqueda</div>',unsafe_allow_html=True)

    min_score = st.slider("Compatibilidad mínima",0,100,65,5)

    salary_min_label = st.selectbox(
        "Renta mínima",
        ["$1,2M","$1,5M","$1,8M","$2,0M"],
        index=0
    )

    salary_min_map = {
        "$1,2M":1_200_000,
        "$1,5M":1_500_000,
        "$1,8M":1_800_000,
        "$2,0M":2_000_000
    }
    salary_min = salary_min_map[salary_min_label]

    salary_max_label = st.selectbox(
        "Renta máxima",
        ["$2,0M","$2,5M+","Sin límite"],
        index=1
    )

    salary_max_map = {
        "$2,0M":2_000_000,
        "$2,5M+":2_500_000,
        "Sin límite":None
    }
    salary_max = salary_max_map[salary_max_label]

    only_new = st.checkbox("Solo nuevas ofertas",False)
    show_low = st.checkbox("Mostrar ajuste bajo",False)
    show_below = st.checkbox("Mostrar rentas < $1,2M",False)
    only_salary = st.checkbox("Solo con renta publicada",False)

    prefer_salary = st.checkbox(
        "Priorizar ofertas con renta",
        value=True,
        help="Ordena primero las ofertas donde el portal informa una renta."
    )

    st.markdown('<div class="sidebar-section">Fuentes salariales Chile</div>', unsafe_allow_html=True)
    st.caption(
        "La app prioriza referencias de Indeed Chile y Chiletrabajos cuando "
        "la renta está disponible. Laborum se usa como referencia de mercado."
    )

    st.markdown('<div class="sidebar-section">Tipo de trabajo</div>',unsafe_allow_html=True)

    hybrid = st.checkbox("Híbrido", value=True, key="filter_hybrid")
    remote = st.checkbox("Remoto", value=True, key="filter_remote")
    onsite = st.checkbox("Presencial", value=False, key="filter_onsite")

    if not any([hybrid, remote, onsite]):
        st.warning("Selecciona al menos una modalidad para ver resultados.")

    st.markdown('<div class="sidebar-section">Ordenar resultados</div>',unsafe_allow_html=True)
    sort_option = st.selectbox(
        "Orden",
        [
            "Mayor % de ajuste",
            "Menor % de ajuste",
            "Renta publicada: mayor a menor",
            "Empresa: A → Z"
        ],
        index=0,
        label_visibility="collapsed"
    )

# ==========================================================
# HEADER
# ==========================================================
st.markdown("## 💼 Buscador Laboral · Paulina Vergara")
st.caption("Procesos · Mejora Continua · Operaciones · CX · PMO · Control de Gestión · BI")

st.markdown("""
<div class="hero">
    <h2>Encuentra oportunidades que se ajusten a tu perfil</h2>
    <p>La búsqueda usa tu CV activo, tus cargos objetivo, habilidades adicionales y modalidad seleccionada.</p>
</div>
""",unsafe_allow_html=True)

profile_now = active_profile()
skills_preview = profile_now["skills"][:10]
if skills_preview:
    st.caption("🎯 Perfil activo: " + " · ".join(skills_preview))

if st.button("🔎 BUSCAR OFERTAS DE HOY",type="primary",use_container_width=True):
    with st.spinner("Buscando oportunidades compatibles..."):
        live,errors = search_live()

    if live:
        st.session_state["jobs"] = live
        st.session_state["mode"] = "live"
        st.session_state["errors"] = errors
        st.session_state["found_count"] = len(live)

        live_df_to_save = build_df(live)
        persist_jobs(live_df_to_save)
        st.session_state["page"] = "Resultados"
    else:
        fallback = fallback_links()
        st.session_state["jobs"] = fallback
        st.session_state["mode"] = "fallback"
        st.session_state["errors"] = errors
        st.session_state["found_count"] = len(fallback)
        st.session_state["page"] = "Resultados"

    st.rerun()

jobs = st.session_state.get("jobs",[])
mode = st.session_state.get("mode")
errors = st.session_state.get("errors",[])

all_df = build_df(jobs) if jobs else pd.DataFrame()
preview_df = all_df.copy()

if mode == "live" and not all_df.empty:
    if not show_low:
        preview_df = preview_df[preview_df["Score"] >= min_score]

    if not show_below:
        preview_df = preview_df[
            preview_df["RentaValor"].isna() |
            (preview_df["RentaValor"] >= salary_min)
        ]

    if salary_max is not None:
        preview_df = preview_df[
            preview_df["RentaValor"].isna() |
            (preview_df["RentaValor"] <= salary_max) |
            (preview_df["RentaValor"] >= TARGET_PLUS)
        ]

    if only_salary:
        preview_df = preview_df[preview_df["RentaValor"].notna()]

    modes = []
    if hybrid:
        modes.append("Híbrido")
    if remote:
        modes.append("Remoto")
    if onsite:
        modes.append("Presencial / no informado")

    # El filtro es estricto: si marcas solo Remoto, verás solo Remoto.
    if modes:
        preview_df = preview_df[preview_df["Modalidad"].isin(modes)]
    else:
        preview_df = preview_df.iloc[0:0]

    if only_new:
        preview_df = preview_df[preview_df["Estado"]=="Nueva"]

# ==========================================================
# MÉTRICAS PERSONALIZADAS (sin cortes de texto)
# ==========================================================
if mode == "fallback":
    first_label = "Búsquedas disponibles"
    first_value = len(all_df)
    first_sub = "accesos directos activos"
else:
    first_label = "Ofertas encontradas"
    first_value = len(all_df) if mode == "live" else 0
    first_sub = "ofertas únicas"

visible_count = len(preview_df) if mode == "live" else len(all_df) if mode == "fallback" else 0
high_count = int((preview_df["Score"]>=80).sum()) if mode=="live" and not preview_df.empty else 0
salary_count = int(preview_df["RentaValor"].notna().sum()) if mode=="live" and not preview_df.empty else 0

c1,c2,c3,c4,c5 = st.columns(5)

with c1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{first_label}</div>
        <div class="metric-value">{first_value}</div>
        <div class="metric-sub">{first_sub}</div>
    </div>
    """,unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Visibles con filtros</div>
        <div class="metric-value">{visible_count}</div>
        <div class="metric-sub">resultados mostrados</div>
    </div>
    """,unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Ajuste alto</div>
        <div class="metric-value">{high_count}</div>
        <div class="metric-sub">80% o más</div>
    </div>
    """,unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Con renta publicada</div>
        <div class="metric-value">{salary_count}</div>
        <div class="metric-sub">sueldo visible</div>
    </div>
    """,unsafe_allow_html=True)

with c5:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Renta objetivo</div>
        <div class="metric-value salary">{salary_min_label} – {salary_max_label}</div>
        <div class="metric-sub">CLP mensuales</div>
    </div>
    """,unsafe_allow_html=True)

if mode == "live":
    st.success(f"✅ Se encontraron {len(all_df)} ofertas únicas. Mostrando {len(preview_df)} según tus filtros. Orden: {sort_option}.")

    if errors:
        with st.expander(f"⚠️ {len(errors)} consulta(s) tuvieron problemas"):
            for e in errors:
                st.write("•",e)

elif mode == "fallback":
    st.warning(
        f"⚠️ La búsqueda integrada no respondió. "
        f"Hay {len(all_df)} búsquedas directas disponibles como respaldo."
    )

    if errors:
        with st.expander("Ver diagnóstico"):
            for e in errors:
                st.write("•",e)

# ==========================================================
# VISTAS / GESTIÓN DE POSTULACIONES
# ==========================================================
page = st.session_state.get("page", "Inicio")
counts = tracked_counts()

def apply_current_filters(df):
    if df is None or df.empty:
        return df

    filtered = df.copy()

    if not show_low and "Score" in filtered.columns:
        filtered = filtered[filtered["Score"] >= min_score]

    if not show_below and "RentaValor" in filtered.columns:
        filtered = filtered[
            filtered["RentaValor"].isna() |
            (filtered["RentaValor"] >= salary_min)
        ]

    if salary_max is not None and "RentaValor" in filtered.columns:
        filtered = filtered[
            filtered["RentaValor"].isna() |
            (filtered["RentaValor"] <= salary_max) |
            (filtered["RentaValor"] >= TARGET_PLUS)
        ]

    if only_salary and "RentaValor" in filtered.columns:
        filtered = filtered[filtered["RentaValor"].notna()]

    modes = []
    if hybrid:
        modes.append("Híbrido")
    if remote:
        modes.append("Remoto")
    if onsite:
        modes.append("Presencial / no informado")

    if modes and "Modalidad" in filtered.columns:
        filtered = filtered[filtered["Modalidad"].isin(modes)]
    elif not modes:
        filtered = filtered.iloc[0:0]

    if only_new and "Estado" in filtered.columns:
        filtered = filtered[filtered["Estado"] == "Nueva"]

    return filtered

def sort_jobs(df):
    if df is None or df.empty:
        return df

    out = df.copy()
    out["_TieneRenta"] = out["RentaValor"].notna().astype(int)

    if sort_option == "Mayor % de ajuste":
        if prefer_salary:
            return out.sort_values(
                ["_TieneRenta","Score","Empresa"],
                ascending=[False,False,True],
                na_position="last"
            )
        return out.sort_values(
            ["Score","Empresa"],
            ascending=[False,True],
            na_position="last"
        )

    if sort_option == "Menor % de ajuste":
        return out.sort_values(
            ["Score","Empresa"],
            ascending=[True,True],
            na_position="last"
        )

    if sort_option == "Renta publicada: mayor a menor":
        return out.sort_values(
            ["RentaValor","Score"],
            ascending=[False,False],
            na_position="last"
        )

    return out.sort_values(
        ["Empresa","Score"],
        ascending=[True,False],
        na_position="last"
    )

def action_button(label, key, job_key, new_status):
    if st.button(label, key=key, use_container_width=True):
        set_status(job_key, new_status)
        st.rerun()

def render_job_cards(data, context="Resultados"):
    if data is None or data.empty:
        st.info("No hay cargos en esta sección.")
        return

    for _, r in data.iterrows():
        with st.container(border=True):
            left, right = st.columns([5.2, 1.35])

            with left:
                score = int(r.get("Score", 0) or 0)
                cargo = r.get("Cargo", "")
                empresa = r.get("Empresa", "")
                ubicacion = r.get("Ubicación", "")
                modalidad = r.get("Modalidad", "")

                # Título + ajuste
                st.markdown(
                    f"### {cargo} — {empresa}"
                )

                ajuste_icon = "🟢" if score >= 80 else ("🟠" if score >= 65 else "⚪")
                st.markdown(
                    f"{ajuste_icon} **{score}% de ajuste**"
                )

                # Renta publicada
                renta = r.get("Renta", "") or "Sin renta publicada"
                renta_html = safe_money_html(renta)
                st.markdown(
                    f"<div style='font-size:1rem;margin:4px 0;'>"
                    f"💵 <b>Renta publicada:</b> {renta_html}</div>",
                    unsafe_allow_html=True
                )

                # Pretensión sugerida
                pretension = r.get("Pretension", "") or "No calculada"
                pretension_html = safe_money_html(pretension)
                st.markdown(
                    f"<div style='font-size:1rem;margin:4px 0;'>"
                    f"🎯 <b>Pretensión sugerida:</b> {pretension_html}</div>",
                    unsafe_allow_html=True
                )

                # Referencia y criterio salarial
                pret_note = r.get("PretensionNota", "") or ""
                pret_source = r.get("PretensionFuente", "") or ""

                if pret_note or pret_source:
                    if pret_source == "Aviso publicado":
                        st.markdown(
                            f"**Referencia salarial:** banda informada en el aviso"
                        )
                        st.caption(
                            "Criterio: la pretensión se posiciona dentro de la banda publicada, "
                            "considerando tu rango objetivo."
                        )
                    elif pret_source and pret_source != "Perfil de búsqueda":
                        clean_note = pret_note.replace("Referencia:", "").strip()
                        st.markdown(
                            f"**Referencia salarial:** {clean_note} — {pret_source}"
                        )
                        st.caption(
                            "Criterio: estimación orientativa usando referencia de mercado chilena, "
                            "seniority del cargo y tu rango objetivo."
                        )
                    else:
                        clean_note = pret_note.replace("Referencia:", "").strip()
                        st.markdown(
                            f"**Referencia salarial:** {clean_note or 'Sin referencia específica del cargo'}"
                        )
                        st.caption(
                            "Criterio: estimación basada principalmente en el rango de renta "
                            "configurado en tu búsqueda."
                        )

                # Ubicación / modalidad / fuente
                fuente = r.get("Fuente", "") or "No informada"
                st.caption(
                    f"📍 {ubicacion} · 🏢 {modalidad} · 🌐 Fuente: {fuente}"
                )

                # Fechas
                if r.get("FechaEncontrada"):
                    fecha = str(r.get("FechaEncontrada")).replace("T", " ")
                    st.caption(f"🔎 Encontrada: {fecha}")

                if r.get("FechaPostulacion"):
                    fecha = str(r.get("FechaPostulacion")).replace("T", " ")
                    st.markdown(
                        f"📅 **Postulada:** {fecha}"
                    )

                if r.get("FechaGuardada") and context == "Guardadas":
                    fecha = str(r.get("FechaGuardada")).replace("T", " ")
                    st.markdown(
                        f"⭐ **Guardada:** {fecha}"
                    )

                if r.get("FechaDescarte") and context == "Descartadas":
                    fecha = str(r.get("FechaDescarte")).replace("T", " ")
                    st.markdown(
                        f"🗑️ **Descartada:** {fecha}"
                    )

                # Descripción
                desc = r.get("Descripción", "") or ""
                if desc:
                    st.markdown(
                        f"<div class='muted' style='margin-top:8px;'>"
                        f"{desc[:450]}{'...' if len(desc)>450 else ''}"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                # Habilidades detectadas
                skills = r.get("Skills", []) or []
                if isinstance(skills, str):
                    skills = [s.strip() for s in skills.split(",") if s.strip()]

                if skills:
                    chips = "".join(
                        f"<span class='badge-blue'>{s}</span>"
                        for s in skills
                    )
                    st.markdown(
                        f"<div style='margin-top:8px;'>{chips}</div>",
                        unsafe_allow_html=True
                    )

                # Link a referencia salarial cuando existe
                ref_url = r.get("ReferenciaURL", "") or ""
                if ref_url and renta == "Sin renta publicada":
                    st.markdown(
                        f"[Ver referencia salarial]({ref_url})"
                    )

            with right:
                link = r.get("Enlace", "")
                key = r["_key"]

                estado_actual = r.get("Estado", "Nueva")
                estado_icono = {
                    "Nueva": "🆕",
                    "Guardada": "⭐",
                    "Postulada": "✅",
                    "Descartada": "🗑️",
                }.get(estado_actual, "•")

                st.caption(f"{estado_icono} Estado: {estado_actual}")

                if link:
                    st.link_button(
                        "🔗 Ver oferta",
                        link,
                        use_container_width=True
                    )

                if context == "Resultados":
                    if r.get("Estado") != "Guardada":
                        action_button(
                            "⭐ Guardar",
                            f"save_{abs(hash(key))}",
                            key,
                            "Guardada"
                        )

                    if r.get("Estado") != "Postulada":
                        action_button(
                            "✅ Postular",
                            f"apply_{abs(hash(key))}",
                            key,
                            "Postulada"
                        )

                    if r.get("Estado") != "Descartada":
                        action_button(
                            "🗑️ Descartar",
                            f"discard_{abs(hash(key))}",
                            key,
                            "Descartada"
                        )

                elif context == "Guardadas":
                    action_button(
                        "✅ Marcar postulada",
                        f"saved_apply_{abs(hash(key))}",
                        key,
                        "Postulada"
                    )

                    action_button(
                        "🗑️ Descartar",
                        f"saved_discard_{abs(hash(key))}",
                        key,
                        "Descartada"
                    )

                    action_button(
                        "↩️ Quitar guardado",
                        f"saved_new_{abs(hash(key))}",
                        key,
                        "Nueva"
                    )

                elif context == "Postuladas":
                    action_button(
                        "🗑️ Descartar",
                        f"post_discard_{abs(hash(key))}",
                        key,
                        "Descartada"
                    )

                    action_button(
                        "⭐ Guardar",
                        f"post_save_{abs(hash(key))}",
                        key,
                        "Guardada"
                    )

                elif context == "Descartadas":
                    action_button(
                        "↩️ Recuperar",
                        f"discard_restore_{abs(hash(key))}",
                        key,
                        "Nueva"
                    )

                    action_button(
                        "✅ Postular",
                        f"discard_apply_{abs(hash(key))}",
                        key,
                        "Postulada"
                    )

# ---------- INICIO ----------
if page == "Inicio":
    st.markdown("### Resumen de tu búsqueda laboral")

    a, b, c, d = st.columns(4)
    a.metric("Ofertas acumuladas", counts["Resultados"])
    b.metric("Guardadas", counts["Guardadas"])
    c.metric("Postuladas", counts["Postuladas"])
    d.metric("Descartadas", counts["Descartadas"])

    st.info(
        "Las búsquedas nuevas se acumulan. Una oferta ya encontrada no se elimina "
        "cuando vuelves a ejecutar la búsqueda."
    )

# ---------- RESULTADOS ----------
elif page == "Resultados":
    st.markdown(f"### 🔎 Resultados acumulados ({counts['Resultados']})")

    tracked = load_tracked_jobs()
    visible = apply_current_filters(tracked)
    visible = sort_jobs(visible)

    st.caption(
        f"Mostrando {len(visible)} de {len(tracked)} ofertas acumuladas según tus filtros."
    )

    render_job_cards(visible, "Resultados")

    # Si falló la búsqueda integrada, conservamos accesos directos de respaldo.
    # Estas tarjetas también muestran pretensión y referencia salarial.
    if mode == "fallback" and jobs:
        st.markdown("### 🌐 Búsquedas directas de respaldo")
        st.caption(
            "Estas son búsquedas por portal, no ofertas individuales. "
            "La pretensión mostrada es una referencia orientativa para el tipo de cargo."
        )

        fallback_df = build_df(jobs)

        for _, r in fallback_df.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([5.2, 1.2])

                with c1:
                    score = int(r.get("Score", 0) or 0)
                    icon = "🟢" if score >= 80 else ("🟠" if score >= 65 else "⚪")

                    st.markdown(
                        f"### {r.get('Cargo','')} — {r.get('Empresa','')}"
                    )

                    st.markdown(
                        f"{icon} **{score}% de ajuste**"
                    )

                    st.markdown(
                        "<div style='font-size:1rem;margin:4px 0;'>"
                        "💵 <b>Renta publicada:</b> Sin renta publicada</div>",
                        unsafe_allow_html=True
                    )

                    pretension = r.get("Pretension", "") or "No calculada"
                    pretension_html = safe_money_html(pretension)
                    st.markdown(
                        f"<div style='font-size:1rem;margin:4px 0;'>"
                        f"🎯 <b>Pretensión sugerida:</b> {pretension_html}</div>",
                        unsafe_allow_html=True
                    )

                    note = (r.get("PretensionNota", "") or "").replace("Referencia:", "").strip()
                    source = r.get("PretensionFuente", "") or ""

                    if source == "Aviso publicado":
                        st.markdown(
                            "**Referencia salarial:** banda informada en el aviso"
                        )
                        st.caption(
                            "Criterio: estimación basada en la banda salarial publicada "
                            "y en tu rango objetivo."
                        )
                    elif source and source != "Perfil de búsqueda":
                        ref_text = note if note else "Mercado chileno para cargos comparables"
                        st.markdown(
                            f"**Referencia salarial:** {ref_text} — {source}"
                        )
                        st.caption(
                            "Criterio: estimación orientativa según mercado chileno, "
                            "seniority del cargo y rango objetivo configurado."
                        )
                    else:
                        st.markdown(
                            f"**Referencia salarial:** "
                            f"{note or 'Sin referencia específica para este cargo'}"
                        )
                        st.caption(
                            "Criterio: estimación basada principalmente en el rango "
                            "de renta configurado en tu búsqueda."
                        )

                    st.caption(
                        f"📍 {r.get('Ubicación','')} · "
                        f"🌐 Fuente: {r.get('Fuente','')}"
                    )

                    descripcion = r.get("Descripción", "") or ""
                    if descripcion:
                        st.markdown(
                            f"<div class='muted'>{descripcion}</div>",
                            unsafe_allow_html=True
                        )

                    ref_url = r.get("ReferenciaURL", "") or ""
                    if ref_url:
                        st.markdown(
                            f"[Ver referencia salarial]({ref_url})"
                        )

                with c2:
                    if r.get("Enlace"):
                        st.link_button(
                            "🔎 Abrir búsqueda",
                            r["Enlace"],
                            use_container_width=True
                        )

                    base_key = abs(hash(
                        str(r.get("Cargo","")) +
                        str(r.get("Empresa","")) +
                        str(r.get("Fuente",""))
                    ))

                    if st.button(
                        "⭐ Guardar",
                        key=f"save_fallback_{base_key}",
                        use_container_width=True
                    ):
                        persist_fallback_as_job(r, "Guardada")
                        st.success("Guardada en la sección Guardadas.")
                        st.rerun()

                    if st.button(
                        "✅ Postular",
                        key=f"apply_fallback_{base_key}",
                        use_container_width=True
                    ):
                        persist_fallback_as_job(r, "Postulada")
                        st.success("Registrada en Postuladas con fecha y hora.")
                        st.rerun()

                    if st.button(
                        "🗑️ Descartar",
                        key=f"discard_fallback_{base_key}",
                        use_container_width=True
                    ):
                        persist_fallback_as_job(r, "Descartada")
                        st.success("Movida a Descartadas.")
                        st.rerun()

# ---------- GUARDADAS ----------
elif page == "Guardadas":
    df_guardadas = load_tracked_jobs("Guardada")
    saved_searches = load_saved_searches()
    total_saved = len(df_guardadas) + len(saved_searches)

    st.markdown(f"### ⭐ Guardadas ({total_saved})")
    st.caption(
        "Aquí quedan las ofertas y búsquedas que quieres revisar más adelante."
    )

    if not df_guardadas.empty:
        st.markdown("#### Ofertas guardadas")
        render_job_cards(sort_jobs(df_guardadas), "Guardadas")

    if saved_searches:
        st.markdown("#### Búsquedas guardadas")
        for item in saved_searches:
            with st.container(border=True):
                c1, c2 = st.columns([5,1.3])
                with c1:
                    st.markdown(f"**{item['cargo']} — {item['fuente']}**")
                    st.caption(f"📍 {item['ubicacion']} · ⭐ Guardada: {item['saved_at'].replace('T',' ')}")
                    if item["pretension"]:
                        st.markdown(f"🎯 **Pretensión orientativa:** {item['pretension']}")
                    if item["referencia"]:
                        st.markdown(f"**Referencia salarial:** {item['referencia']}")
                with c2:
                    if item["enlace"]:
                        st.link_button("🔎 Abrir búsqueda", item["enlace"], use_container_width=True)
                    if st.button(
                        "🗑️ Quitar",
                        key=f"del_search_{abs(hash(item['search_key']))}",
                        use_container_width=True
                    ):
                        delete_saved_search(item["search_key"])
                        st.rerun()

    if df_guardadas.empty and not saved_searches:
        st.info("No hay elementos guardados.")

# ---------- POSTULADAS ----------
elif page == "Postuladas":
    df_postuladas = load_tracked_jobs("Postulada")
    st.markdown(f"### ✅ Postuladas ({len(df_postuladas)})")
    st.caption(
        "La fecha se registra automáticamente cuando marcas una oferta como Postulada."
    )

    if not df_postuladas.empty:
        excel_bytes = applications_excel_bytes(df_postuladas)

        if excel_bytes:
            st.download_button(
                "📊 Descargar Excel de postulaciones",
                data=excel_bytes,
                file_name=f"Postulaciones_Paulina_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False
            )

    render_job_cards(sort_jobs(df_postuladas), "Postuladas")

# ---------- DESCARTADAS ----------
elif page == "Descartadas":
    df_descartadas = load_tracked_jobs("Descartada")
    st.markdown(f"### 🗑️ Descartadas ({len(df_descartadas)})")
    st.caption(
        "Puedes mover aquí las postulaciones cuando recibas una respuesta negativa "
        "o cuando decidas no continuar. También puedes recuperarlas."
    )
    render_job_cards(sort_jobs(df_descartadas), "Descartadas")

st.markdown("---")
st.caption(
    "La compatibilidad es orientativa. La renta solo se muestra cuando aparece en el aviso original."
)



import re
import sqlite3
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

@st.cache_resource
def get_db():
    conn = sqlite3.connect("jobs.db", check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS job_status(
            job_key TEXT PRIMARY KEY,
            status TEXT,
            updated_at TEXT
        )
    """)
    conn.commit()
    return conn

DB = get_db()

def get_status(key):
    row = DB.execute("SELECT status FROM job_status WHERE job_key=?", (key,)).fetchone()
    return row[0] if row else "Nueva"

def set_status(key,status):
    DB.execute("""
        INSERT INTO job_status(job_key,status,updated_at)
        VALUES(?,?,?)
        ON CONFLICT(job_key)
        DO UPDATE SET status=excluded.status,updated_at=excluded.updated_at
    """,(key,status,datetime.now().isoformat()))
    DB.commit()

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
        return txt,parse_salary(txt)

    text = " ".join(job.get("extensions",[]) or []) + " " + job.get("description","")
    v = parse_salary(text)

    if v:
        return f"${v:,.0f} CLP".replace(",","."),v

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
        slug = q.replace(" ","-")

        rows += [
            {
                "title":q.title(),
                "company_name":"LinkedIn Jobs",
                "location":"Santiago / Chile",
                "via":"LinkedIn",
                "description":"Búsqueda directa en LinkedIn.",
                "job_apply_link":f"https://www.linkedin.com/jobs/search/?keywords={enc}&location=Santiago%2C%20Chile",
                "_fallback":True
            },
            {
                "title":q.title(),
                "company_name":"Computrabajo",
                "location":"Santiago / Chile",
                "via":"Computrabajo",
                "description":"Búsqueda directa en Computrabajo.",
                "job_apply_link":f"https://cl.computrabajo.com/trabajo-de-{slug}",
                "_fallback":True
            },
            {
                "title":q.title(),
                "company_name":"Trabajando.com",
                "location":"Santiago / Chile",
                "via":"Trabajando",
                "description":"Búsqueda directa en Trabajando.com.",
                "job_apply_link":f"https://www.trabajando.cl/trabajo-empleo/?q={enc}",
                "_fallback":True
            }
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

        rows.append({
            "Cargo":title,
            "Empresa":company,
            "Ubicación":loc,
            "Fuente":j.get("via",""),
            "Score":score,
            "Renta":sal_txt,
            "RentaValor":sal_val,
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
    st.markdown("""
    <div class="brand">
        <div class="brand-icon">💼</div>
        <div class="brand-title">Buscador Laboral<br>Paulina Vergara</div>
    </div>
    <div class="nav-active">🏠 &nbsp; Inicio</div>
    <div class="nav-item">🔎 &nbsp; Resultados</div>
    <div class="nav-item">⭐ &nbsp; Guardadas</div>
    <div class="nav-item">✅ &nbsp; Postuladas</div>
    <div class="nav-item">🗑️ &nbsp; Descartadas</div>
    <div class="nav-item">🕘 &nbsp; Historial</div>
    """,unsafe_allow_html=True)


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
    else:
        fallback = fallback_links()
        st.session_state["jobs"] = fallback
        st.session_state["mode"] = "fallback"
        st.session_state["errors"] = errors
        st.session_state["found_count"] = len(fallback)

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
# RESULTADOS
# ==========================================================
if jobs:
    if mode == "live":
        if sort_option == "Mayor % de ajuste":
            df = preview_df.sort_values(
                ["Score","Empresa"],
                ascending=[False,True],
                na_position="last"
            )
        elif sort_option == "Menor % de ajuste":
            df = preview_df.sort_values(
                ["Score","Empresa"],
                ascending=[True,True],
                na_position="last"
            )
        elif sort_option == "Renta publicada: mayor a menor":
            df = preview_df.sort_values(
                ["RentaValor","Score"],
                ascending=[False,False],
                na_position="last"
            )
        else:
            df = preview_df.sort_values(
                ["Empresa","Score"],
                ascending=[True,False],
                na_position="last"
            )
    else:
        df = all_df

    tabs = st.tabs([
        f"Resultados ({len(df)})",
        f"Guardadas ({int((df['Estado']=='Guardada').sum()) if not df.empty else 0})",
        f"Postuladas ({int((df['Estado']=='Postulada').sum()) if not df.empty else 0})",
        f"Descartadas ({int((df['Estado']=='Descartada').sum()) if not df.empty else 0})"
    ])

    def render(data,status=None):
        view = data if not status else data[data["Estado"]==status]

        if view.empty:
            st.info("No hay elementos en esta sección.")
            return

        for _,r in view.iterrows():
            with st.container(border=True):
                left,right = st.columns([5,1])

                with left:
                    st.markdown(
                        f"<div class='job-title'>{r['Cargo']} "
                        f"<span class='badge-green'>{r['Score']}% de ajuste</span></div>",
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        f"<div class='job-meta'><b>{r['Empresa']}</b> · "
                        f"{r['Ubicación']} · {r['Modalidad']}</div>",
                        unsafe_allow_html=True
                    )

                    if r["_fallback"]:
                        st.markdown(
                            f"<div class='salary-line'>🌐 Fuente: {r['Fuente']}</div>",
                            unsafe_allow_html=True
                        )
                    else:
                        salary_tag = (
                            "2,5M+"
                            if r["RentaValor"] is not None and r["RentaValor"] >= TARGET_PLUS
                            else r["Renta"]
                        )

                        st.markdown(
                            f"<div class='salary-line'>💵 {salary_tag} &nbsp;&nbsp; "
                            f"🏢 {r['Modalidad']}</div>",
                            unsafe_allow_html=True
                        )

                    if r["Descripción"]:
                        t = r["Descripción"]
                        st.markdown(
                            f"<div class='muted'>{t[:360]}{'...' if len(t)>360 else ''}</div>",
                            unsafe_allow_html=True
                        )

                    if r["Skills"]:
                        chips = "".join(
                            f"<span class='badge-blue'>{s}</span>"
                            for s in r["Skills"]
                        )
                        st.markdown(chips,unsafe_allow_html=True)

                with right:
                    if r["Enlace"]:
                        st.link_button(
                            "Ver oferta" if not r["_fallback"] else "Abrir búsqueda",
                            r["Enlace"],
                            use_container_width=True
                        )

                    if not r["_fallback"]:
                        values = ["Nueva","Guardada","Postulada","Descartada"]
                        selected = st.selectbox(
                            "Estado",
                            values,
                            index=values.index(r["Estado"]),
                            key=f"state_{abs(hash(r['_key']))}"
                        )

                        if selected != r["Estado"]:
                            set_status(r["_key"],selected)
                            st.rerun()

    with tabs[0]:
        render(df)

    with tabs[1]:
        render(df,"Guardada")

    with tabs[2]:
        render(df,"Postulada")

    with tabs[3]:
        render(df,"Descartada")

else:
    st.markdown("""
    ### Cómo usarla
    1. Presiona **BUSCAR OFERTAS DE HOY**.
    2. Usa los filtros laterales para ajustar compatibilidad, renta y modalidad.
    3. Revisa primero los resultados con mayor compatibilidad.
    4. Guarda, postula o descarta cada oportunidad.
    """)

st.markdown("---")
st.caption(
    "La compatibilidad es orientativa. La renta solo se muestra cuando aparece en el aviso original."
)

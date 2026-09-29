

import re
import sqlite3
from datetime import datetime
from urllib.parse import quote_plus

import pandas as pd
import requests
import streamlit as st

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
        --red:#B34B4B;
        --text:#17324D;
        --muted:#6D7D8D;
        --line:#E1E8EF;
        --bg:#F7F9FC;
        --card:#FFFFFF;
    }

    .stApp {
        background: var(--bg);
        color: var(--text);
    }

    .block-container {
        max-width: 1400px;
        padding-top: 1.2rem;
        padding-bottom: 3rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }

    [data-testid="stSidebar"] {
        background: #F4F7FA;
        border-right: 1px solid var(--line);
    }

    [data-testid="stSidebar"] .block-container {
        padding-top: 1.4rem;
    }

    h1,h2,h3,h4 {
        color: var(--navy2) !important;
        letter-spacing: -0.02em;
    }

    .brand {
        display:flex;
        align-items:center;
        gap:12px;
        margin-bottom: 18px;
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
        font-size: 1.25rem;
        font-weight: 800;
        color: var(--navy2);
        line-height:1.1;
    }

    .nav-active {
        background: linear-gradient(90deg,var(--navy),#175482);
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
        background:
            linear-gradient(90deg,rgba(255,249,240,.97) 0%,rgba(255,246,232,.90) 45%,rgba(238,225,204,.72) 100%),
            url("https://images.unsplash.com/photo-1518005020951-eccb494ad742?auto=format&fit=crop&w=1800&q=75");
        background-size:cover;
        background-position:center;
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
        margin:0 0 14px 0;
        color:#4E6276;
    }

    div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg,var(--navy),#155B89);
        color:white;
        border:0;
        border-radius:10px;
        min-height:46px;
        font-weight:800;
        box-shadow:0 5px 12px rgba(18,59,99,.16);
    }

    div.stButton > button[kind="primary"]:hover {
        background: linear-gradient(90deg,var(--navy2),var(--navy));
        color:white;
    }

    [data-testid="stMetric"] {
        background: var(--card);
        border:1px solid var(--line);
        border-radius:15px;
        padding:15px 18px;
        box-shadow:0 4px 14px rgba(28,53,79,.05);
    }

    [data-testid="stMetricLabel"] {
        color:var(--muted);
        font-weight:700;
    }

    [data-testid="stMetricValue"] {
        color:var(--navy2);
        font-weight:800;
    }

    div[data-baseweb="tab-list"] {
        gap:8px;
        border-bottom:1px solid var(--line);
    }

    button[data-baseweb="tab"] {
        background:transparent;
        color:#52677C;
        border-radius:0;
        padding-left:8px;
        padding-right:8px;
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

    .badge-gold {
        display:inline-block;
        background:var(--gold-soft);
        color:#875E08;
        border-radius:8px;
        padding:4px 9px;
        font-weight:800;
        font-size:.8rem;
    }

    .muted {
        color:var(--muted);
        font-size:.9rem;
    }

    .section-title {
        font-size:1.02rem;
        font-weight:800;
        color:var(--navy2);
        margin:15px 0 8px 0;
    }

    div[data-testid="stSelectbox"] label,
    div[data-testid="stSlider"] label,
    div[data-testid="stCheckbox"] label {
        color:var(--text);
        font-weight:600;
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

# ==========================================================
# PERFIL
# ==========================================================
PROFILE = {
    "roles": [
        "Ingeniera de Procesos",
        "Analista Senior de Procesos",
        "Mejora Continua",
        "Excelencia Operacional",
        "Experiencia Cliente",
        "Customer Experience",
        "Business Process Analyst",
        "Analista de Operaciones",
        "PMO",
        "Control de Gestión",
        "Business Intelligence",
        "KPIs",
    ],
    "skills": [
        "Power BI","Excel","SQL","Bizagi","BPMN","Lean","Kaizen",
        "AS IS","TO BE","KPIs","automatización","mejora continua",
        "customer journey","gestión de procesos","reportes ejecutivos",
        "coordinación transversal","stakeholders"
    ],
}

SEARCH_QUERIES = [
    "analista senior procesos",
    "ingeniero procesos mejora continua",
    "excelencia operacional",
    "control de gestion power bi",
    "business process analyst",
]

TARGET_MIN = 1_200_000
TARGET_PLUS = 2_500_000

# ==========================================================
# BASE DE DATOS LOCAL
# ==========================================================
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

def set_status(key, status):
    DB.execute("""
        INSERT INTO job_status(job_key,status,updated_at)
        VALUES(?,?,?)
        ON CONFLICT(job_key)
        DO UPDATE SET status=excluded.status,updated_at=excluded.updated_at
    """, (key,status,datetime.now().isoformat()))
    DB.commit()

# ==========================================================
# LÓGICA DE COMPATIBILIDAD
# ==========================================================
def norm(t):
    return re.sub(r"\s+"," ",(t or "").lower()).strip()

def score_job(title, company, desc):
    text = norm(f"{title} {company} {desc}")
    score = 35

    score += min(sum(norm(s) in text for s in PROFILE["skills"]) * 4, 36)

    role_hits = 0
    for role in PROFILE["roles"]:
        words = [w for w in norm(role).split() if len(w) > 3]
        if words and sum(w in text for w in words) >= max(1, len(words)//2):
            role_hits += 1

    score += min(role_hits * 6, 24)

    if any(x in text for x in ["senior","sr.","especialista","ingeniero","ingeniera"]):
        score += 4

    if any(x in text for x in ["práctica","practica","intern","vendedor","promotor","operario"]):
        score -= 30

    return max(0,min(score,100))

def compatibility(score):
    if score >= 80:
        return "Alto"
    if score >= 65:
        return "Medio"
    return "Bajo"

def parse_salary(text):
    if not text:
        return None

    s = str(text).lower().replace("\xa0"," ")

    m = re.search(r"(\d{1,2})[\.,](\d)\s*(?:mill[oó]n|millones|m)\b", s)
    if m:
        return int(float(f"{m.group(1)}.{m.group(2)}") * 1_000_000)

    m = re.search(r"(\d{1,2})\s*(?:mill[oó]n|millones)\b", s)
    if m:
        return int(m.group(1)) * 1_000_000

    vals = []
    for raw in re.findall(r"\$?\s*(\d[\d\.\,]{5,})", s):
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
        return txt, parse_salary(txt)

    text = " ".join(job.get("extensions",[]) or []) + " " + job.get("description","")
    v = parse_salary(text)

    if v:
        return f"${v:,.0f} CLP".replace(",","."), v

    return "Sin renta publicada", None

def detect_work_mode(job):
    text = norm(
        f"{job.get('title','')} {job.get('location','')} "
        f"{job.get('description','')} {' '.join(job.get('extensions',[]) or [])}"
    )
    if "remoto" in text or "remote" in text:
        return "Remoto"
    if "híbrido" in text or "hibrido" in text or "hybrid" in text:
        return "Híbrido"
    return "Presencial / no informado"

def extract_skills(desc):
    text = norm(desc)
    labels = []
    mapping = [
        ("bpmn","BPM/BPMN"),
        ("power bi","Power BI"),
        ("sql","SQL"),
        ("mejora continua","Mejora Continua"),
        ("automatiz","Automatización"),
        ("stakeholder","Stakeholders"),
        ("kpi","KPIs"),
        ("as is","AS IS / TO BE"),
        ("to be","AS IS / TO BE"),
        ("lean","Lean"),
        ("kaizen","Kaizen"),
    ]
    for key,label in mapping:
        if key in text and label not in labels:
            labels.append(label)
    return labels[:6]

# ==========================================================
# BÚSQUEDA
# ==========================================================
def serp_key():
    try:
        return st.secrets.get("SERPAPI_KEY","")
    except Exception:
        return ""

def search_live():
    key = serp_key()

    if not key or key == "TU_CLAVE_AQUI":
        return [], ["No se encontró una SERPAPI_KEY válida."]

    jobs, errors = [], []

    for q in SEARCH_QUERIES:
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

    unique, keys = [], set()

    for j in jobs:
        k = (
            j.get("title",""),
            j.get("company_name",""),
            j.get("location","")
        )
        if k not in keys:
            keys.add(k)
            unique.append(j)

    return unique, errors

def fallback_links():
    rows = []

    for q in SEARCH_QUERIES:
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
            "Ajuste":compatibility(score),
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
    """, unsafe_allow_html=True)

    st.markdown('<div class="sidebar-section">Filtros de búsqueda</div>', unsafe_allow_html=True)

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
        "$2,0M":2_000_000,
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
        "Sin límite":None,
    }
    salary_max = salary_max_map[salary_max_label]

    only_new = st.checkbox("Solo nuevas ofertas",False)
    show_low = st.checkbox("Mostrar ajuste bajo",False)
    show_below = st.checkbox("Mostrar rentas < $1,2M",False)
    only_salary = st.checkbox("Solo con renta publicada",False)

    st.markdown('<div class="sidebar-section">Tipo de trabajo</div>', unsafe_allow_html=True)
    hybrid = st.checkbox("Híbrido",True)
    remote = st.checkbox("Remoto",True)
    onsite = st.checkbox("Presencial",False)

    st.markdown('<div class="sidebar-section">Ubicación</div>', unsafe_allow_html=True)
    st.selectbox(
        "Zona",
        ["Santiago (híbrido / remoto)","Santiago","Remoto Chile"],
        index=0,
        label_visibility="collapsed"
    )

# ==========================================================
# HEADER
# ==========================================================
head1, head2 = st.columns([4,1])
with head1:
    st.markdown("## 💼 Buscador Laboral · Paulina Vergara")
    st.caption("Procesos · Mejora Continua · Operaciones · CX · PMO · Control de Gestión · BI")

with head2:
    st.markdown(
        f"<div style='text-align:right;color:#425B72;font-weight:700;padding-top:8px;'>"
        f"📅 {datetime.now().strftime('%d de %B de %Y')}</div>",
        unsafe_allow_html=True
    )

st.markdown("""
<div class="hero">
    <h2>Encuentra oportunidades que se ajusten a tu perfil</h2>
    <p>Cargos en procesos, mejora continua, operaciones, experiencia cliente, PMO y control de gestión.</p>
</div>
""", unsafe_allow_html=True)

if st.button("🔎 BUSCAR OFERTAS DE HOY",type="primary",use_container_width=True):
    with st.spinner("Buscando oportunidades compatibles..."):
        live, errors = search_live()

    if live:
        st.session_state["jobs"] = live
        st.session_state["mode"] = "live"
        st.session_state["errors"] = errors
        st.session_state["found_count"] = len(live)
    else:
        st.session_state["jobs"] = fallback_links()
        st.session_state["mode"] = "fallback"
        st.session_state["errors"] = errors
        st.session_state["found_count"] = 0

    st.rerun()

jobs = st.session_state.get("jobs",[])
mode = st.session_state.get("mode")
errors = st.session_state.get("errors",[])

# ==========================================================
# MÉTRICAS
# ==========================================================
if jobs:
    all_df = build_df(jobs)
else:
    all_df = pd.DataFrame()

if mode == "live" and not all_df.empty:
    preview_df = all_df.copy()

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

    work_modes = []
    if hybrid: work_modes.append("Híbrido")
    if remote: work_modes.append("Remoto")
    if onsite: work_modes.append("Presencial / no informado")

    if work_modes:
        preview_df = preview_df[preview_df["Modalidad"].isin(work_modes)]

    if only_new:
        preview_df = preview_df[preview_df["Estado"] == "Nueva"]

else:
    preview_df = all_df.copy()

m1,m2,m3,m4,m5 = st.columns(5)
m1.metric("Ofertas encontradas", st.session_state.get("found_count",0))
m2.metric("Visibles con filtros", len(preview_df) if mode=="live" else 0)
m3.metric("Ajuste alto", int((preview_df["Score"]>=80).sum()) if mode=="live" and not preview_df.empty else 0)
m4.metric("Con renta publicada", int(preview_df["RentaValor"].notna().sum()) if mode=="live" and not preview_df.empty else 0)
m5.metric("Renta objetivo", f"{salary_min_label} – {salary_max_label}")

if mode == "live":
    st.success(f"✅ Se encontraron {len(all_df)} ofertas únicas. Mostrando {len(preview_df)} según tus filtros.")

    if errors:
        with st.expander(f"⚠️ {len(errors)} consulta(s) tuvieron problemas"):
            for e in errors:
                st.write("•",e)

elif mode == "fallback":
    st.warning("⚠️ No hubo resultados integrados. Se muestran búsquedas directas de respaldo.")
    if errors:
        with st.expander("Ver diagnóstico"):
            for e in errors:
                st.write("•",e)

# ==========================================================
# RESULTADOS
# ==========================================================
if jobs:
    df = preview_df.sort_values(["Score","Empresa"],ascending=[False,True]) if mode=="live" else all_df

    tabs = st.tabs([
        f"Resultados ({len(df) if mode=='live' else 0})",
        f"Guardadas ({int((df['Estado']=='Guardada').sum()) if not df.empty else 0})",
        f"Postuladas ({int((df['Estado']=='Postulada').sum()) if not df.empty else 0})",
        f"Descartadas ({int((df['Estado']=='Descartada').sum()) if not df.empty else 0})",
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
                        salary_tag = "2,5M+" if (
                            r["RentaValor"] is not None and r["RentaValor"] >= TARGET_PLUS
                        ) else r["Renta"]

                        st.markdown(
                            f"<div class='salary-line'>💵 {salary_tag} &nbsp;&nbsp; "
                            f"💼 Tiempo indefinido / revisar aviso &nbsp;&nbsp; "
                            f"🏢 {r['Modalidad']}</div>",
                            unsafe_allow_html=True
                        )

                    if r["Descripción"]:
                        text = r["Descripción"]
                        st.markdown(
                            f"<div class='muted'>{text[:360]}{'...' if len(text)>360 else ''}</div>",
                            unsafe_allow_html=True
                        )

                    if r["Skills"]:
                        chips = "".join(
                            [f"<span class='badge-blue'>{s}</span>" for s in r["Skills"]]
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
    2. La aplicación buscará oportunidades relacionadas con tu perfil.
    3. Usa los filtros de la izquierda para ajustar compatibilidad, sueldo y modalidad.
    4. Revisa primero las ofertas de mayor compatibilidad.
    5. Marca las vacantes como **Guardada**, **Postulada** o **Descartada**.
    """)

st.markdown("---")
st.caption(
    "La compatibilidad es orientativa. La renta solo se muestra cuando aparece en la publicación. "
    "Revisa siempre las condiciones en el aviso original."
)

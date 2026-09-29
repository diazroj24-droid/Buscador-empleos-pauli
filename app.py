
import re
import sqlite3
from datetime import datetime
from urllib.parse import quote_plus

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Buscador Laboral Paulina", page_icon="🔎", layout="wide")

PROFILE = {
    "roles": [
        "Ingeniera de Procesos", "Analista Senior de Procesos", "Mejora Continua",
        "Excelencia Operacional", "Experiencia Cliente", "Customer Experience",
        "Business Process Analyst", "Analista de Operaciones", "PMO",
        "Control de Gestión", "Business Intelligence", "KPIs",
    ],
    "skills": [
        "Power BI", "Excel", "SQL", "Bizagi", "BPMN", "Lean", "Kaizen",
        "AS IS", "TO BE", "KPIs", "automatización", "mejora continua",
        "customer journey", "gestión de procesos", "reportes ejecutivos",
        "coordinación transversal"
    ],
}

SEARCH_QUERIES = [
    "analista senior procesos",
    "ingeniero procesos mejora continua",
    "excelencia operacional",
    "control de gestion power bi",
    "business process analyst",
]

DB_PATH = "jobs.db"

@st.cache_resource
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS job_status (
            job_key TEXT PRIMARY KEY,
            status TEXT,
            updated_at TEXT
        )
    """)
    conn.commit()
    return conn

CONN = get_conn()

def get_status(job_key):
    row = CONN.execute("SELECT status FROM job_status WHERE job_key=?", (job_key,)).fetchone()
    return row[0] if row else "Nueva"

def set_status(job_key, status):
    CONN.execute("""
        INSERT INTO job_status(job_key,status,updated_at)
        VALUES (?,?,?)
        ON CONFLICT(job_key)
        DO UPDATE SET status=excluded.status, updated_at=excluded.updated_at
    """, (job_key, status, datetime.now().isoformat()))
    CONN.commit()

def normalize(txt):
    txt = (txt or "").lower()
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()

def score_job(title, company, description):
    text = normalize(f"{title} {company} {description}")
    score = 35
    skill_hits = sum(normalize(skill) in text for skill in PROFILE["skills"])
    role_hits = 0
    for role in PROFILE["roles"]:
        words = [w for w in normalize(role).split() if len(w) > 3]
        if words and sum(w in text for w in words) >= max(1, len(words)//2):
            role_hits += 1
    score += min(skill_hits * 4, 36)
    score += min(role_hits * 6, 24)
    if any(x in text for x in ["senior", "sr.", "especialista", "ingeniero", "ingeniera"]):
        score += 4
    if any(x in text for x in ["práctica", "practica", "intern", "vendedor", "ventas terreno", "promotor", "operario"]):
        score -= 30
    return max(0, min(score, 100))

def level(score):
    if score >= 80:
        return "🟢 Alto"
    if score >= 65:
        return "🟡 Medio"
    return "⚪ Bajo"

def extract_salary(job):
    ext = job.get("detected_extensions") or {}
    if isinstance(ext, dict) and ext.get("salary"):
        return ext["salary"]
    return "No publicada"

def get_serpapi_key():
    try:
        return st.secrets.get("SERPAPI_KEY", "")
    except Exception:
        return ""

def search_serpapi():
    api_key = get_serpapi_key()
    if not api_key or api_key == "TU_CLAVE_AQUI":
        return [], ["No se encontró una SERPAPI_KEY válida en Streamlit Secrets."]

    all_jobs, errors = [], []
    for q in SEARCH_QUERIES:
        params = {
            "engine": "google_jobs",
            "q": f"{q} Santiago Chile",
            "hl": "es",
            "gl": "cl",
            "api_key": api_key,
        }
        try:
            r = requests.get("https://serpapi.com/search.json", params=params, timeout=25)
            if r.status_code != 200:
                try:
                    detail = r.json().get("error", r.text[:180])
                except Exception:
                    detail = r.text[:180]
                errors.append(f"{q}: HTTP {r.status_code} — {detail}")
                continue
            data = r.json()
            if data.get("error"):
                errors.append(f"{q}: {data['error']}")
                continue
            all_jobs.extend(data.get("jobs_results", []) or [])
        except requests.Timeout:
            errors.append(f"{q}: tiempo de espera agotado.")
        except Exception as e:
            errors.append(f"{q}: {type(e).__name__}.")

    seen, unique = set(), []
    for job in all_jobs:
        key = (job.get("title",""), job.get("company_name",""), job.get("location",""))
        if key not in seen:
            seen.add(key)
            unique.append(job)
    return unique, errors

def search_links():
    items = []
    for query in SEARCH_QUERIES:
        encoded = quote_plus(query)
        slug = query.replace(" ", "-")
        items.extend([
            {
                "title": query.title(), "company_name": "LinkedIn Jobs",
                "location": "Santiago / Chile", "via": "LinkedIn",
                "description": "Búsqueda directa para revisar las ofertas más recientes en LinkedIn.",
                "job_apply_link": f"https://www.linkedin.com/jobs/search/?keywords={encoded}&location=Santiago%2C%20Chile",
                "_fallback": True
            },
            {
                "title": query.title(), "company_name": "Computrabajo",
                "location": "Santiago / Chile", "via": "Computrabajo",
                "description": "Búsqueda directa para revisar las ofertas más recientes en Computrabajo.",
                "job_apply_link": f"https://cl.computrabajo.com/trabajo-de-{slug}",
                "_fallback": True
            },
            {
                "title": query.title(), "company_name": "Trabajando.com",
                "location": "Santiago / Chile", "via": "Trabajando",
                "description": "Búsqueda directa para revisar las ofertas más recientes en Trabajando.com.",
                "job_apply_link": f"https://www.trabajando.cl/trabajo-empleo/?q={encoded}",
                "_fallback": True
            },
        ])
    return items

def build_df(jobs):
    rows = []
    for job in jobs:
        title = job.get("title", "Sin título")
        company = job.get("company_name", "Empresa no informada")
        desc = job.get("description", "")
        location = job.get("location", "No informado")
        source = job.get("via", "")
        link = job.get("job_apply_link") or ""
        if not link:
            opts = job.get("apply_options") or []
            if opts:
                link = opts[0].get("link", "")
        score = score_job(title, company, desc)
        key = f"{title}|{company}|{location}"
        rows.append({
            "Ajuste": level(score), "Score": score, "Cargo": title,
            "Empresa": company, "Ubicación": location, "Fuente": source,
            "Renta": extract_salary(job), "Descripción": desc,
            "Enlace": link, "Estado": get_status(key), "_key": key,
            "_fallback": bool(job.get("_fallback", False)),
        })
    return pd.DataFrame(rows)

st.title("🔎 Buscador Laboral – Paulina Vergara")
st.caption("Procesos · Mejora Continua · Operaciones · CX · PMO · Control de Gestión · BI")

with st.sidebar:
    st.header("Filtros")
    min_score = st.slider("Compatibilidad mínima", 0, 100, 65, 5)
    only_new = st.checkbox("Solo nuevas", value=False)
    show_low = st.checkbox("Mostrar ajuste bajo", value=False)
    st.markdown("---")
    st.write("**Renta objetivo**")
    st.write("$1.500.000 – $2.000.000 CLP")
    st.write("**Ubicación**")
    st.write("Santiago / híbrido / remoto")
    st.markdown("---")
    st.caption("Si una publicación no informa renta, aparecerá como 'No publicada'.")

c1, c2, c3 = st.columns(3)
c1.metric("Perfil", "4+ años")
c2.metric("Renta objetivo", "$1,5M – $2,5M")
c3.metric("Fecha", datetime.now().strftime("%d-%m-%Y"))

if st.button("🔍 BUSCAR OFERTAS DE HOY", type="primary", use_container_width=True):
    with st.spinner("Buscando ofertas compatibles..."):
        live_jobs, errors = search_serpapi()
    if live_jobs:
        st.session_state["jobs"] = live_jobs
        st.session_state["mode"] = "live"
        st.session_state["errors"] = errors
    else:
        st.session_state["jobs"] = search_links()
        st.session_state["mode"] = "fallback"
        st.session_state["errors"] = errors

jobs = st.session_state.get("jobs", [])
mode = st.session_state.get("mode")
errors = st.session_state.get("errors", [])

if mode == "live":
    st.success(f"✅ Búsqueda completada. Se obtuvieron {len(jobs)} resultados.")
    if errors:
        with st.expander(f"⚠️ {len(errors)} consulta(s) no respondieron; las demás sí funcionaron"):
            for err in errors:
                st.write("•", err)

elif mode == "fallback":
    st.warning("⚠️ La búsqueda integrada no devolvió resultados. Te muestro búsquedas directas de respaldo.")
    if errors:
        with st.expander("Ver diagnóstico de la búsqueda"):
            for err in errors:
                st.write("•", err)
        st.info("Si el diagnóstico menciona límite mensual, créditos o autenticación, revisa tu cuenta/clave de SerpAPI.")

if jobs:
    df = build_df(jobs)
    if not show_low and mode == "live":
        df = df[df["Score"] >= min_score]
    if only_new:
        df = df[df["Estado"] == "Nueva"]
    df = df.sort_values(["Score", "Empresa"], ascending=[False, True])

    if mode == "fallback":
        st.subheader("🌐 Búsquedas directas")
        st.caption("Abre cada portal con los cargos ya configurados.")

    tabs = st.tabs(["🆕 Resultados", "⭐ Guardadas", "✅ Postuladas", "🗑️ Descartadas"])

    def render_jobs(dataframe, filter_status=None):
        data = dataframe if not filter_status else dataframe[dataframe["Estado"] == filter_status]
        if data.empty:
            st.write("No hay elementos en esta sección.")
            return
        for _, row in data.iterrows():
            with st.container(border=True):
                left, right = st.columns([4, 1])
                with left:
                    st.subheader(row["Cargo"])
                    st.write(f"**{row['Empresa']}** · {row['Ubicación']}")
                    if row["_fallback"]:
                        st.write(f"🌐 **Fuente:** {row['Fuente']}")
                    else:
                        st.write(f"{row['Ajuste']} · **Compatibilidad: {row['Score']}%** · **Renta:** {row['Renta']}")
                    if row["Descripción"]:
                        txt = row["Descripción"]
                        st.caption(txt[:600] + ("..." if len(txt) > 600 else ""))
                with right:
                    if row["Enlace"]:
                        label = "Abrir búsqueda" if row["_fallback"] else "Abrir oferta"
                        st.link_button(label, row["Enlace"], use_container_width=True)
                    if not row["_fallback"]:
                        statuses = ["Nueva", "Guardada", "Postulada", "Descartada"]
                        selected = st.selectbox(
                            "Estado", statuses,
                            index=statuses.index(row["Estado"]),
                            key=f"status_{abs(hash(row['_key']))}"
                        )
                        if selected != row["Estado"]:
                            set_status(row["_key"], selected)
                            st.rerun()

    with tabs[0]:
        render_jobs(df)
    with tabs[1]:
        render_jobs(df, "Guardada")
    with tabs[2]:
        render_jobs(df, "Postulada")
    with tabs[3]:
        render_jobs(df, "Descartada")
else:
    st.markdown("""
### Cómo usarla
1. Presiona **BUSCAR OFERTAS DE HOY**.
2. La app intentará traer ofertas integradas.
3. Si el servicio externo falla, mostrará búsquedas directas.
4. Revisa primero las oportunidades de mayor compatibilidad.
5. Abre siempre el aviso original antes de postular.
""")

st.markdown("---")
st.caption("La compatibilidad es orientativa. Revisa requisitos, renta y condiciones en el aviso original.")


import os
import re
import sqlite3
from datetime import datetime
from urllib.parse import quote_plus

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Buscador Laboral Paulina",
    page_icon="🔎",
    layout="wide"
)

# -----------------------------
# Perfil precargado desde CV
# -----------------------------
PROFILE = {
    "name": "Paulina Vergara",
    "location": "Santiago, Chile",
    "target_salary_min": 1500000,
    "target_salary_max": 2000000,
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
        "BI",
        "KPIs",
    ],
    "skills": [
        "Power BI", "Excel", "SQL", "Bizagi", "BPMN",
        "Lean", "Kaizen", "AS IS", "TO BE", "KPIs",
        "automatización", "mejora continua", "customer journey",
        "gestión de procesos", "reportes ejecutivos",
        "coordinación transversal"
    ],
}

SEARCH_QUERIES = [
    "analista senior procesos",
    "ingeniero procesos",
    "mejora continua",
    "excelencia operacional",
    "business process analyst",
    "analista operaciones power bi sql",
    "control de gestion power bi",
    "pmo procesos",
    "experiencia cliente procesos",
]

DB_PATH = "jobs.db"

def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS job_status (
            job_key TEXT PRIMARY KEY,
            status TEXT,
            updated_at TEXT
        )
    """)
    return conn

CONN = db()

def get_status(job_key):
    row = CONN.execute(
        "SELECT status FROM job_status WHERE job_key=?",
        (job_key,)
    ).fetchone()
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

    skill_hits = 0
    for skill in PROFILE["skills"]:
        if normalize(skill) in text:
            skill_hits += 1

    role_hits = 0
    for role in PROFILE["roles"]:
        role_words = [w for w in normalize(role).split() if len(w) > 3]
        if role_words and sum(w in text for w in role_words) >= max(1, len(role_words)//2):
            role_hits += 1

    score += min(skill_hits * 4, 36)
    score += min(role_hits * 6, 24)

    if any(x in text for x in ["senior", "sr.", "especialista", "ingeniero"]):
        score += 4
    if any(x in text for x in ["práctica", "practica", "intern", "vendedor", "ventas terreno"]):
        score -= 25

    return max(0, min(score, 100))

def level(score):
    if score >= 80:
        return "🟢 Alto"
    if score >= 65:
        return "🟡 Medio"
    return "⚪ Bajo"

def parse_salary(text):
    if not text:
        return None
    clean = text.lower().replace(".", "").replace(",", "")
    nums = re.findall(r"\b(\d{6,8})\b", clean)
    vals = [int(n) for n in nums]
    vals = [v for v in vals if 500000 <= v <= 10000000]
    if not vals:
        return None
    return sum(vals) // len(vals)

def salary_label(raw_salary):
    est = parse_salary(raw_salary)
    if est:
        return f"${est:,.0f} CLP".replace(",", ".")
    return "No publicada"

def fallback_links():
    rows = []
    queries = SEARCH_QUERIES[:6]
    platforms = [
        ("LinkedIn", "https://www.linkedin.com/jobs/search/?keywords={q}&location=Santiago%2C%20Chile"),
        ("Computrabajo", "https://cl.computrabajo.com/trabajo-de-{q}"),
        ("Trabajando", "https://www.trabajando.cl/trabajo-empleo/?q={q}"),
    ]
    for q in queries:
        for platform, template in platforms:
            slug = quote_plus(q).replace("+", "%20")
            if platform == "Computrabajo":
                slug = q.replace(" ", "-")
            rows.append({
                "title": q.title(),
                "company_name": platform,
                "location": "Santiago / Chile",
                "via": platform,
                "description": "Búsqueda directa guardada. Abre el enlace para revisar avisos nuevos.",
                "job_apply_link": template.format(q=slug),
                "detected_extensions": {},
            })
    return rows

def search_serpapi():
    api_key = st.secrets.get("SERPAPI_KEY", "")
    if not api_key:
        return None, "No hay SERPAPI_KEY configurada."

    all_jobs = []
    errors = []
    for q in SEARCH_QUERIES:
        params = {
            "engine": "google_jobs",
            "q": f"{q} Santiago Chile",
            "hl": "es",
            "gl": "cl",
            "api_key": api_key,
        }
        try:
            r = requests.get("https://serpapi.com/search.json", params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
            all_jobs.extend(data.get("jobs_results", []))
        except Exception as e:
            errors.append(str(e))

    # dedupe
    seen = set()
    unique = []
    for j in all_jobs:
        key = (j.get("title",""), j.get("company_name",""), j.get("location",""))
        if key not in seen:
            seen.add(key)
            unique.append(j)

    msg = None
    if errors:
        msg = "Algunas consultas no respondieron correctamente."
    return unique, msg

def build_df(jobs):
    rows = []
    for j in jobs:
        title = j.get("title", "Sin título")
        company = j.get("company_name", "Empresa no informada")
        desc = j.get("description", "")
        loc = j.get("location", "No informado")
        via = j.get("via", "")
        apply_link = j.get("job_apply_link") or ""
        if not apply_link:
            opts = j.get("apply_options") or []
            if opts:
                apply_link = opts[0].get("link", "")

        ext = j.get("detected_extensions") or {}
        raw_salary = ""
        if isinstance(ext, dict):
            raw_salary = ext.get("salary", "") or ""
        score = score_job(title, company, desc)
        key = f"{title}|{company}|{loc}"

        rows.append({
            "Ajuste": level(score),
            "Score": score,
            "Cargo": title,
            "Empresa": company,
            "Ubicación": loc,
            "Fuente": via,
            "Renta": salary_label(raw_salary),
            "Descripción": desc,
            "Enlace": apply_link,
            "Estado": get_status(key),
            "_key": key,
        })
    return pd.DataFrame(rows)

st.title("🔎 Buscador Laboral – Paulina Vergara")
st.caption("Procesos · Mejora Continua · Operaciones · CX · PMO · Control de Gestión · BI")

with st.sidebar:
    st.header("Filtros")
    min_score = st.slider("Compatibilidad mínima", 0, 100, 65, 5)
    only_new = st.checkbox("Solo nuevas", value=False)
    include_low = st.checkbox("Mostrar ajuste bajo", value=False)
    st.markdown("---")
    st.write("**Renta objetivo**")
    st.write("$1.500.000 – $2.000.000 CLP")
    st.write("**Ubicación**")
    st.write("Santiago / híbrido / remoto")
    st.markdown("---")
    st.info("Cuando una oferta no publica sueldo, la app la muestra como 'No publicada' para evitar inventar una cifra.")

col1, col2, col3 = st.columns(3)
col1.metric("Perfil", "4+ años")
col2.metric("Renta objetivo", "$1,5M – $2,0M")
col3.metric("Última búsqueda", datetime.now().strftime("%d-%m-%Y"))

if st.button("🔍 BUSCAR OFERTAS DE HOY", type="primary", use_container_width=True):
    jobs, msg = search_serpapi()
    if jobs is None:
        st.session_state["jobs"] = fallback_links()
        st.session_state["search_mode"] = "links"
        st.warning(
            "La app aún no tiene una clave SerpAPI configurada. "
            "Mostrando búsquedas directas guardadas en LinkedIn, Computrabajo y Trabajando."
        )
    else:
        st.session_state["jobs"] = jobs
        st.session_state["search_mode"] = "live"
        if msg:
            st.warning(msg)

jobs = st.session_state.get("jobs", [])

if jobs:
    df = build_df(jobs)

    if not include_low:
        df = df[df["Score"] >= min_score]
    else:
        df = df[df["Score"] >= 0]

    if only_new:
        df = df[df["Estado"] == "Nueva"]

    df = df.sort_values(["Score", "Empresa"], ascending=[False, True])

    if st.session_state.get("search_mode") == "live":
        st.success(f"Se encontraron {len(df)} oportunidades después de aplicar filtros.")
    else:
        st.info("Modo enlaces: abre las búsquedas guardadas. Para resultados integrados dentro de la app, configura SerpAPI.")

    tabs = st.tabs(["🆕 Resultados", "⭐ Guardadas", "✅ Postuladas", "🗑️ Descartadas"])

    def render_jobs(dataframe, filter_status=None):
        data = dataframe
        if filter_status:
            data = dataframe[dataframe["Estado"] == filter_status]

        if data.empty:
            st.write("No hay ofertas en esta sección.")
            return

        for _, row in data.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.subheader(row["Cargo"])
                    st.write(f"**{row['Empresa']}** · {row['Ubicación']}")
                    st.write(f"{row['Ajuste']} · **Compatibilidad: {row['Score']}%** · Renta: {row['Renta']}")
                    if row["Descripción"]:
                        st.caption(row["Descripción"][:500] + ("..." if len(row["Descripción"]) > 500 else ""))
                with c2:
                    if row["Enlace"]:
                        st.link_button("Abrir oferta", row["Enlace"], use_container_width=True)

                    status = st.selectbox(
                        "Estado",
                        ["Nueva", "Guardada", "Postulada", "Descartada"],
                        index=["Nueva", "Guardada", "Postulada", "Descartada"].index(row["Estado"]),
                        key=f"status_{abs(hash(row['_key']))}"
                    )
                    if status != row["Estado"]:
                        set_status(row["_key"], status)
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
    2. Revisa primero las ofertas con ajuste **Alto**.
    3. Abre el aviso original.
    4. Marca cada resultado como **Guardada**, **Postulada** o **Descartada**.
    5. Vuelve al día siguiente y repite la búsqueda.
    """)

st.markdown("---")
st.caption(
    "La compatibilidad es orientativa y se calcula comparando palabras clave del aviso "
    "con las competencias del CV. Siempre revisa los requisitos completos antes de postular."
)

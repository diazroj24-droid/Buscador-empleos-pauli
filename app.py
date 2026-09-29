
import re, sqlite3
from datetime import datetime
from urllib.parse import quote_plus
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Buscador Laboral Paulina", page_icon="🔎", layout="wide")

st.markdown("""
<style>
.stApp {background: linear-gradient(180deg,#F8F3EC 0%,#F3ECE3 100%); color:#3C2F2F;}
.block-container {max-width:1320px; padding-top:2rem;}
[data-testid="stSidebar"] {background:#EFE4D5; border-right:1px solid #DDCBB7;}
[data-testid="stMetric"] {background:#FFFDF9; border:1px solid #E4D5C5; padding:14px 16px; border-radius:14px; box-shadow:0 4px 14px rgba(70,45,35,.06);}
[data-testid="stMetricValue"] {color:#7A3E3E; font-weight:700;}
div.stButton>button[kind="primary"] {background:linear-gradient(90deg,#B66A4D,#7A3E3E); color:white; border:0; border-radius:12px; height:48px; font-weight:700;}
[data-testid="stLinkButton"] a {background:#7A3E3E; color:white!important; border:0; border-radius:10px;}
[data-testid="stVerticalBlockBorderWrapper"] {background:#FFFDF9; border:1px solid #E4D5C5!important; border-radius:14px!important;}
button[data-baseweb="tab"] {background:#EEE1D2; border-radius:10px 10px 0 0; color:#5C4941;}
button[data-baseweb="tab"][aria-selected="true"] {background:#B66A4D; color:white;}
.banner {background:linear-gradient(135deg,#7A3E3E,#A95F48 62%,#C89B5A); color:white; border-radius:18px; padding:22px 26px; margin-bottom:18px;}
.banner h2 {color:white!important; margin:0 0 6px 0;}
.banner p {margin:0; opacity:.94;}
</style>
""", unsafe_allow_html=True)

PROFILE = {
    "roles":["Ingeniera de Procesos","Analista Senior de Procesos","Mejora Continua","Excelencia Operacional","Experiencia Cliente","Customer Experience","Business Process Analyst","Analista de Operaciones","PMO","Control de Gestión","Business Intelligence","KPIs"],
    "skills":["Power BI","Excel","SQL","Bizagi","BPMN","Lean","Kaizen","AS IS","TO BE","KPIs","automatización","mejora continua","customer journey","gestión de procesos","reportes ejecutivos","coordinación transversal"]
}
SEARCH_QUERIES=["analista senior procesos","ingeniero procesos mejora continua","excelencia operacional","control de gestion power bi","business process analyst"]
TARGET_MIN=1_200_000
TARGET_PLUS=2_500_000

@st.cache_resource
def conn():
    c=sqlite3.connect("jobs.db",check_same_thread=False)
    c.execute("CREATE TABLE IF NOT EXISTS job_status (job_key TEXT PRIMARY KEY,status TEXT,updated_at TEXT)")
    c.commit()
    return c
DB=conn()

def get_status(k):
    r=DB.execute("SELECT status FROM job_status WHERE job_key=?",(k,)).fetchone()
    return r[0] if r else "Nueva"

def set_status(k,s):
    DB.execute("""INSERT INTO job_status(job_key,status,updated_at) VALUES(?,?,?)
    ON CONFLICT(job_key) DO UPDATE SET status=excluded.status,updated_at=excluded.updated_at""",(k,s,datetime.now().isoformat()))
    DB.commit()

def norm(t): return re.sub(r"\s+"," ",(t or "").lower()).strip()

def score_job(title,company,desc):
    text=norm(f"{title} {company} {desc}")
    score=35+min(sum(norm(x) in text for x in PROFILE["skills"])*4,36)
    for role in PROFILE["roles"]:
        words=[w for w in norm(role).split() if len(w)>3]
        if words and sum(w in text for w in words)>=max(1,len(words)//2): score+=6
    if any(x in text for x in ["senior","sr.","especialista","ingeniero","ingeniera"]): score+=4
    if any(x in text for x in ["práctica","practica","intern","vendedor","ventas terreno","promotor","operario"]): score-=30
    return max(0,min(score,100))

def level(s):
    return "🟢 Alto" if s>=80 else ("🟠 Medio" if s>=65 else "⚪ Bajo")

def parse_salary(text):
    if not text: return None
    s=str(text).lower().replace("\xa0"," ")
    m=re.search(r"(\d{1,2})[\.,](\d)\s*(?:mill[oó]n|millones|m)\b",s)
    if m: return int(float(f"{m.group(1)}.{m.group(2)}")*1_000_000)
    m=re.search(r"(\d{1,2})\s*(?:mill[oó]n|millones)\b",s)
    if m: return int(m.group(1))*1_000_000
    vals=[]
    for raw in re.findall(r"\$?\s*(\d[\d\.\,]{5,})",s):
        d=re.sub(r"\D","",raw)
        if d.isdigit() and 500_000<=int(d)<=20_000_000: vals.append(int(d))
    return max(vals) if vals else None

def salary_data(job):
    ext=job.get("detected_extensions") or {}
    if isinstance(ext,dict) and ext.get("salary"):
        txt=str(ext["salary"]); return txt,parse_salary(txt)
    text=" ".join(job.get("extensions",[]) or [])+" "+job.get("description","")
    v=parse_salary(text)
    return ((f"${v:,.0f} CLP".replace(",","."),v) if v else ("No publicada",None))

def serp_key():
    try: return st.secrets.get("SERPAPI_KEY","")
    except: return ""

def search_live():
    key=serp_key()
    if not key or key=="TU_CLAVE_AQUI": return [],["No se encontró una SERPAPI_KEY válida."]
    jobs,errors=[],[]
    for q in SEARCH_QUERIES:
        try:
            r=requests.get("https://serpapi.com/search.json",params={"engine":"google_jobs","q":f"{q} Santiago Chile","hl":"es","gl":"cl","api_key":key},timeout=25)
            if r.status_code!=200:
                try: detail=r.json().get("error",r.text[:160])
                except: detail=r.text[:160]
                errors.append(f"{q}: HTTP {r.status_code} — {detail}"); continue
            data=r.json()
            if data.get("error"): errors.append(f"{q}: {data['error']}"); continue
            jobs.extend(data.get("jobs_results",[]) or [])
        except Exception as e: errors.append(f"{q}: {type(e).__name__}")
    seen=[]; unique=[]; keys=set()
    for j in jobs:
        k=(j.get("title",""),j.get("company_name",""),j.get("location",""))
        if k not in keys: keys.add(k); unique.append(j)
    return unique,errors

def fallback_links():
    rows=[]
    for q in SEARCH_QUERIES:
        enc=quote_plus(q); slug=q.replace(" ","-")
        rows += [
            {"title":q.title(),"company_name":"LinkedIn Jobs","location":"Santiago / Chile","via":"LinkedIn","description":"Búsqueda directa en LinkedIn.","job_apply_link":f"https://www.linkedin.com/jobs/search/?keywords={enc}&location=Santiago%2C%20Chile","_fallback":True},
            {"title":q.title(),"company_name":"Computrabajo","location":"Santiago / Chile","via":"Computrabajo","description":"Búsqueda directa en Computrabajo.","job_apply_link":f"https://cl.computrabajo.com/trabajo-de-{slug}","_fallback":True},
            {"title":q.title(),"company_name":"Trabajando.com","location":"Santiago / Chile","via":"Trabajando","description":"Búsqueda directa en Trabajando.com.","job_apply_link":f"https://www.trabajando.cl/trabajo-empleo/?q={enc}","_fallback":True}
        ]
    return rows

def build_df(jobs):
    rows=[]
    for j in jobs:
        title=j.get("title","Sin título"); company=j.get("company_name","Empresa no informada")
        desc=j.get("description",""); loc=j.get("location","No informado")
        link=j.get("job_apply_link") or ""
        if not link and j.get("apply_options"): link=j["apply_options"][0].get("link","")
        s=score_job(title,company,desc); sal_txt,sal_val=salary_data(j); key=f"{title}|{company}|{loc}"
        rows.append({"Ajuste":level(s),"Score":s,"Cargo":title,"Empresa":company,"Ubicación":loc,"Fuente":j.get("via",""),
                     "Renta":sal_txt,"RentaValor":sal_val,"Descripción":desc,"Enlace":link,"Estado":get_status(key),"_key":key,
                     "_fallback":bool(j.get("_fallback",False))})
    return pd.DataFrame(rows)

st.markdown('<div class="banner"><h2>🔎 Buscador Laboral · Paulina Vergara</h2><p>Procesos · Mejora Continua · Operaciones · CX · PMO · Control de Gestión · BI</p></div>',unsafe_allow_html=True)

with st.sidebar:
    st.header("Filtros")
    min_score=st.slider("Compatibilidad mínima",0,100,65,5)
    only_new=st.checkbox("Solo nuevas",False)
    show_low=st.checkbox("Mostrar ajuste bajo",False)
    show_below=st.checkbox("Mostrar rentas bajo $1,2M",False,help="Las ofertas sin renta publicada se mantienen visibles.")
    st.markdown("---")
    st.write("**Renta objetivo**")
    st.write("**$1,2M – $2,5M+ CLP**")
    st.write("**Ubicación**")
    st.write("Santiago / híbrido / remoto")
    st.caption("No se inventan sueldos. Si el aviso no publica renta, seguirá visible.")

top1,top2,top3,top4=st.columns(4)
top1.metric("Perfil","4+ años")
top2.metric("Renta objetivo","$1,2M – $2,5M+")
top3.metric("Fecha",datetime.now().strftime("%d-%m-%Y"))
top4.metric("Ofertas encontradas",st.session_state.get("found_count",0))

if st.button("🔍 BUSCAR OFERTAS DE HOY",type="primary",use_container_width=True):
    with st.spinner("Buscando oportunidades compatibles..."):
        live,errors=search_live()
    if live:
        st.session_state["jobs"]=live; st.session_state["mode"]="live"; st.session_state["found_count"]=len(live)
    else:
        st.session_state["jobs"]=fallback_links(); st.session_state["mode"]="fallback"; st.session_state["found_count"]=0
    st.session_state["errors"]=errors
    st.rerun()

jobs=st.session_state.get("jobs",[])
mode=st.session_state.get("mode")
errors=st.session_state.get("errors",[])

if mode=="live":
    st.success(f"✅ Búsqueda completada: {st.session_state.get('found_count',len(jobs))} ofertas únicas encontradas.")
    if errors:
        with st.expander(f"⚠️ {len(errors)} consulta(s) con problemas"):
            for e in errors: st.write("•",e)
elif mode=="fallback":
    st.warning("⚠️ No hubo resultados integrados. Se muestran búsquedas directas de respaldo.")
    if errors:
        with st.expander("Ver diagnóstico"):
            for e in errors: st.write("•",e)

if jobs:
    all_df=build_df(jobs); df=all_df.copy()
    if mode=="live":
        if not show_low: df=df[df["Score"]>=min_score]
        if not show_below: df=df[df["RentaValor"].isna() | (df["RentaValor"]>=TARGET_MIN)]
    if only_new: df=df[df["Estado"]=="Nueva"]
    df=df.sort_values(["Score","Empresa"],ascending=[False,True])

    if mode=="live":
        a,b,c,d=st.columns(4)
        a.metric("Resultados totales",len(all_df))
        b.metric("Visibles con filtros",len(df))
        c.metric("Ajuste alto",int((df["Score"]>=80).sum()) if len(df) else 0)
        d.metric("Con renta publicada",int(df["RentaValor"].notna().sum()) if len(df) else 0)
        st.caption(f"Mostrando {len(df)} de {len(all_df)} ofertas únicas. Las ofertas sin renta publicada permanecen visibles.")

    if mode=="fallback":
        st.subheader("🌐 Búsquedas directas")

    tabs=st.tabs(["🆕 Resultados","⭐ Guardadas","✅ Postuladas","🗑️ Descartadas"])

    def render(data,status=None):
        view=data if not status else data[data["Estado"]==status]
        if view.empty:
            st.info("No hay elementos en esta sección."); return
        for _,r in view.iterrows():
            with st.container(border=True):
                left,right=st.columns([4,1])
                with left:
                    st.subheader(r["Cargo"])
                    st.write(f"**{r['Empresa']}** · {r['Ubicación']}")
                    if r["_fallback"]:
                        st.write(f"🌐 **Fuente:** {r['Fuente']}")
                    else:
                        rango="2,5M+" if (r["RentaValor"] is not None and r["RentaValor"]>=TARGET_PLUS) else ("En rango" if (r["RentaValor"] is not None and r["RentaValor"]>=TARGET_MIN) else "Sin renta publicada")
                        st.write(f"{r['Ajuste']} · **Compatibilidad: {r['Score']}%** · **Renta:** {r['Renta']} · **Rango:** {rango}")
                    if r["Descripción"]:
                        t=r["Descripción"]; st.caption(t[:650]+("..." if len(t)>650 else ""))
                with right:
                    if r["Enlace"]:
                        st.link_button("Abrir búsqueda" if r["_fallback"] else "Abrir oferta",r["Enlace"],use_container_width=True)
                    if not r["_fallback"]:
                        vals=["Nueva","Guardada","Postulada","Descartada"]
                        sel=st.selectbox("Estado",vals,index=vals.index(r["Estado"]),key=f"s_{abs(hash(r['_key']))}")
                        if sel!=r["Estado"]: set_status(r["_key"],sel); st.rerun()

    with tabs[0]: render(df)
    with tabs[1]: render(df,"Guardada")
    with tabs[2]: render(df,"Postulada")
    with tabs[3]: render(df,"Descartada")
else:
    st.markdown("""
### Cómo usarla
1. Presiona **BUSCAR OFERTAS DE HOY**.
2. La app elimina duplicados y muestra la cantidad real encontrada.
3. Por defecto oculta rentas publicadas bajo **$1,2M**.
4. Mantiene visibles las ofertas sin renta publicada.
5. Las oportunidades de **$2,5M o más** se marcan como **2,5M+**.
""")

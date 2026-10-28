import html
import json
import secrets
from datetime import datetime

import bcrypt
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core import (
    get_db_path, get_db_connection, init_db, extract_text_from_pdf_bytes,
    parse_requisicion_text, excel_to_requisicion_text, build_candidate_record,
    generate_terna_pdf, generate_excel_historico, send_selection_email,
    save_convocatoria_db, get_convocatorias_db, save_convocatoria_results, load_convocatoria,
    analyze_cv_for_tables, generate_candidate_interview_questions
)

st.set_page_config(page_title="TalentValidator Pro | Auditoría Clínica de Talento",
                   page_icon="🩺", layout="wide", initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# Estilos: paleta pastel del sector salud y auditoría clínica
# ---------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap');
:root{
  --mint:#B2DFDB; --mint-soft:#E0F2F1; --teal:#00897B; --teal-dark:#00695C;
  --sky:#B3E5FC; --sky-soft:#E1F5FE; --lav:#D1C4E9; --lav-soft:#EDE7F6;
  --peach:#FFE0B2; --peach-soft:#FFF3E0; --rose:#F8BBD0; --rose-soft:#FCE4EC; --ink:#1F3B3A;
}
html, body, [class*="css"], .stApp { font-family:'Nunito','Segoe UI',sans-serif; }
.stApp { background: radial-gradient(circle at 0% 0%, #EAF7F4 0%, #F7FBFB 45%, #F3F6FD 100%); color:var(--ink); }
@keyframes fadeUp { from {opacity:0; transform:translateY(8px);} to {opacity:1; transform:none;} }
@keyframes pulse { 0%,100% {box-shadow:0 0 0 0 rgba(0,137,123,.35);} 50% {box-shadow:0 0 0 6px rgba(0,137,123,0);} }

/* Encabezado */
.hero { background:linear-gradient(120deg,#D7F2EE 0%,#E3F2FD 55%,#EDE7F6 100%); border:1px solid var(--mint);
  border-radius:20px; padding:18px 24px; margin-bottom:14px; box-shadow:0 6px 22px rgba(0,105,92,.08); animation:fadeUp .5s ease; }
.hero-row { display:flex; justify-content:space-between; align-items:center; gap:14px; flex-wrap:wrap; }
.hero h1 { font-size:1.55rem; font-weight:800; color:var(--teal-dark); margin:0; }
.hero p { margin:2px 0 0 0; color:#3E6B66; font-size:.92rem; }
.conv-pill { background:#fff; border:1.5px solid var(--mint); border-radius:14px; padding:6px 14px; text-align:right; }
.conv-pill small { color:#6B8A87; font-weight:700; letter-spacing:.04em; text-transform:uppercase; font-size:.7rem; }
.conv-pill b { color:var(--teal); font-size:1rem; display:block; }

/* Stepper del proceso */
.stepper { display:flex; gap:8px; margin-top:14px; flex-wrap:wrap; }
.step { flex:1; min-width:130px; display:flex; align-items:center; gap:8px; background:rgba(255,255,255,.75);
  border:1px solid #DDEDEA; border-radius:12px; padding:7px 10px; font-weight:700; font-size:.85rem; color:#6B8A87; }
.step .dot { width:24px; height:24px; border-radius:50%; display:flex; align-items:center; justify-content:center;
  background:#ECEFF1; color:#78909C; font-size:.8rem; }
.step.done { color:var(--teal-dark); border-color:var(--mint); background:#fff; }
.step.done .dot { background:var(--teal); color:#fff; }
.step.current .dot { background:var(--sky); color:#01579B; animation:pulse 1.8s infinite; }

/* Pestañas tipo píldora */
.stTabs [data-baseweb="tab-list"] { gap:6px; background:#fff; padding:6px; border-radius:16px; border:1px solid #DDEDEA;
  box-shadow:0 2px 10px rgba(0,105,92,.05); }
.stTabs [data-baseweb="tab"] { border-radius:12px; padding:8px 16px; font-weight:700; color:#4F6F6B; transition:all .2s; }
.stTabs [data-baseweb="tab"]:hover { background:var(--mint-soft); color:var(--teal-dark); }
.stTabs [aria-selected="true"] { background:linear-gradient(135deg,var(--mint),var(--sky)) !important; color:#004D40 !important; }
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display:none; }

/* Tarjetas (st.container con key card_*) */
[class*="st-key-card"] { background:#fff; border:1px solid #DDEDEA; border-radius:18px; padding:18px 18px 10px 18px;
  box-shadow:0 4px 16px rgba(0,105,92,.06); animation:fadeUp .45s ease; transition:box-shadow .2s, transform .2s; }
[class*="st-key-card"]:hover { box-shadow:0 8px 24px rgba(0,105,92,.10); }
[class*="st-key-cardlav"] { background:linear-gradient(180deg,#fff 0%,#F7F3FD 100%); border-color:var(--lav); }
[class*="st-key-cardsky"] { background:linear-gradient(180deg,#fff 0%,#F2FAFE 100%); border-color:var(--sky); }
[class*="st-key-cardq"] { border-left:5px solid var(--mint); }
.card-title { font-weight:800; color:var(--teal-dark); font-size:1.05rem; margin-bottom:6px; }
.card-sub { color:#5F7F7B; font-size:.86rem; margin-bottom:8px; }
.section-title { font-weight:800; font-size:1.25rem; color:var(--teal-dark); margin:4px 0 10px 0; }

/* KPI */
.kpi { border-radius:16px; padding:14px 16px; display:flex; align-items:center; gap:12px; border:1px solid transparent;
  animation:fadeUp .5s ease; transition:transform .2s; }
.kpi:hover { transform:translateY(-3px); }
.kpi .ic { font-size:1.6rem; width:46px; height:46px; border-radius:14px; display:flex; align-items:center; justify-content:center; background:rgba(255,255,255,.7); }
.kpi .n { font-size:1.6rem; font-weight:800; line-height:1; }
.kpi .l { font-size:.8rem; font-weight:700; opacity:.8; }
.kpi.mint { background:var(--mint-soft); border-color:var(--mint); color:var(--teal-dark); }
.kpi.sky { background:var(--sky-soft); border-color:var(--sky); color:#01579B; }
.kpi.rose { background:var(--rose-soft); border-color:var(--rose); color:#AD1457; }
.kpi.lav { background:var(--lav-soft); border-color:var(--lav); color:#4527A0; }
.kpi.peach { background:var(--peach-soft); border-color:var(--peach); color:#E65100; }

/* Badges / chips */
.badge { display:inline-block; padding:4px 11px; border-radius:999px; font-weight:700; font-size:.78rem; border:1px solid; margin:2px 3px 2px 0; }
.badge-ok { background:#E8F5E9; color:#2E7D32; border-color:#C8E6C9; }
.badge-ko { background:var(--rose-soft); color:#AD1457; border-color:var(--rose); }
.badge-lav { background:var(--lav-soft); color:#4527A0; border-color:var(--lav); }
.badge-sky { background:var(--sky-soft); color:#0277BD; border-color:var(--sky); }
.badge-peach { background:var(--peach-soft); color:#E65100; border-color:var(--peach); }

/* Barras de progreso */
.bar { height:9px; background:#ECEFF1; border-radius:99px; overflow:hidden; margin:3px 0 8px 0; }
.bar-fill { height:100%; border-radius:99px; transition:width .6s ease; }
.score-big { font-size:2rem; font-weight:800; color:var(--teal); line-height:1.1; }
.alert-soft { border-radius:12px; padding:9px 13px; font-size:.86rem; margin:6px 0; border:1px solid; }
.alert-rose { background:var(--rose-soft); border-color:var(--rose); color:#AD1457; }
.alert-mint { background:var(--mint-soft); border-color:var(--mint); color:var(--teal-dark); }
.alert-peach { background:var(--peach-soft); border-color:var(--peach); color:#BF360C; }

/* Podio de la terna */
.podium { border-radius:18px; padding:16px; background:#fff; border:1px solid #DDEDEA; text-align:center;
  box-shadow:0 4px 16px rgba(0,105,92,.07); animation:fadeUp .5s ease; }
.podium .medal { font-size:2rem; }
.podium h3 { margin:2px 0; color:var(--ink); font-size:1.1rem; }

/* Botones */
div.stButton > button, div.stDownloadButton > button, div.stFormSubmitButton > button {
  border-radius:12px; font-weight:700; transition:all .2s ease; }
div.stButton > button[kind="primary"], div.stFormSubmitButton > button[kind="primary"] {
  background:linear-gradient(135deg,#26A69A,#00897B); border:none; color:#fff; }
div.stButton > button[kind="primary"]:hover, div.stFormSubmitButton > button[kind="primary"]:hover {
  transform:translateY(-1px); box-shadow:0 6px 16px rgba(0,137,123,.28); }
div.stButton > button[kind="secondary"]:hover { border-color:var(--teal); color:var(--teal); }

/* Sidebar */
section[data-testid="stSidebar"] { background:linear-gradient(180deg,#E6F5F2 0%,#F1F7FC 100%); border-right:1px solid #D5ECE8; }
.side-brand { background:#fff; border:1px solid var(--mint); border-radius:16px; padding:12px 14px; margin-bottom:10px; }
.side-brand b { color:var(--teal-dark); font-size:1.08rem; }
.side-brand span { color:#5F7F7B; font-size:.78rem; display:block; }
div[data-testid="stDataFrame"] { border-radius:12px; overflow:hidden; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Configuración e inicialización
# ---------------------------------------------------------------------------
DB_PATH = get_db_path(st.secrets)
try:
    ADMIN_DEFAULT_PASS = st.secrets["security"]["superadmin_default_pass"]
except Exception:
    ADMIN_DEFAULT_PASS = "AdminPass2026!"
init_db(DB_PATH, ADMIN_DEFAULT_PASS)

OPCIONES_ENTREVISTA = {"❌ No cumple · 0": 0, "◐ Parcial · 10": 10, "✅ Cumple · 20": 20}
COLORES_TERNA = [("rgba(0,137,123,.30)", "#00897B"), ("rgba(2,136,209,.28)", "#0288D1"),
                 ("rgba(149,117,205,.30)", "#7E57C2")]
MEDALLAS = ["🥇", "🥈", "🥉"]


def nueva_vacante() -> dict:
    return {
        "codigo_vacante": "CONV-" + datetime.now().strftime("%Y%m%d-%H%M"),
        "nombre_puesto": "Coordinador Médico / Auditor de Calidad",
        "meses_experiencia": 24,
        "requiere_tarjeta": True,
        "skills_obligatorias": ["Resolución 3100", "Auditoría médica", "Indicadores de gestión", "Seguridad del paciente"],
    }


defaults = {"user": None, "vacante_activa": nueva_vacante(), "evaluaciones": [], "entrevistas": {},
            "entrevistas_detalle": {}, "lote": 0, "rev": 0, "archivada": False, "req_file_id": None, "flash": None}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def reset_proceso(vacante: dict, evaluaciones=None, entrevistas=None, detalle=None, archivada=False):
    st.session_state.vacante_activa = vacante
    st.session_state.evaluaciones = evaluaciones or []
    st.session_state.entrevistas = entrevistas or {}
    st.session_state.entrevistas_detalle = detalle or {}
    st.session_state.archivada = archivada
    st.session_state.req_file_id = None
    st.session_state.lote += 1
    st.session_state.rev += 1


def flash(msg: str, icon: str = "✅"):
    st.session_state.flash = (msg, icon)


def esc(x) -> str:
    return html.escape(str(x if x is not None else ""))


def badge(text, kind="sky") -> str:
    return f"<span class='badge badge-{kind}'>{esc(text)}</span>"


def barra(valor: float, maximo: float, color: str = "#4DB6AC") -> str:
    pct = 0 if not maximo else max(0.0, min(100.0, valor / maximo * 100))
    return f"<div class='bar'><div class='bar-fill' style='width:{pct:.0f}%;background:{color}'></div></div>"


def kpi(icon: str, num, label: str, tone: str) -> str:
    return (f"<div class='kpi {tone}'><div class='ic'>{icon}</div>"
            f"<div><div class='n'>{num}</div><div class='l'>{esc(label)}</div></div></div>")


def section(title: str):
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)


def card_header(title: str, sub: str = ""):
    st.markdown(f"<div class='card-title'>{title}</div>" + (f"<div class='card-sub'>{sub}</div>" if sub else ""),
                unsafe_allow_html=True)


def smtp_config():
    try:
        cfg = dict(st.secrets.get("smtp", {}))
    except Exception:
        return None
    if not cfg.get("host") or not cfg.get("password") or "xxxx" in str(cfg.get("password")):
        return None
    return cfg


# ---------------------------------------------------------------------------
# Login
# ---------------------------------------------------------------------------
if not st.session_state.user:
    st.markdown("""
        <div class='hero' style='text-align:center; max-width:720px; margin:40px auto 18px auto;'>
            <div style='font-size:2.6rem;'>🩺</div>
            <h1>TalentValidator Pro</h1>
            <p>Selección determinista de talento y auditoría clínica de hojas de vida para el sector salud</p>
        </div>
    """, unsafe_allow_html=True)
    _, col2, _ = st.columns([1, 1.1, 1])
    with col2:
        with st.form("login_form"):
            usr = st.text_input("👤 Usuario o correo institucional")
            pwd = st.text_input("🔒 Contraseña", type="password")
            if st.form_submit_button("Ingresar a la plataforma", type="primary", width="stretch"):
                conn = get_db_connection(DB_PATH)
                u = conn.execute("SELECT * FROM usuarios WHERE username = ?", (usr,)).fetchone()
                conn.close()
                if u and bcrypt.checkpw(pwd.encode('utf-8'), u['password_hash']):
                    if u['activo'] == 1:
                        st.session_state.user = dict(u)
                        flash(f"Bienvenido/a, {u['username']}", "👋")
                        st.rerun()
                    else:
                        st.error("Suscripción suspendida o inhabilitada por el administrador.")
                else:
                    st.error("Credenciales inválidas.")
    st.stop()

if st.session_state.flash:
    msg, icon = st.session_state.flash
    st.toast(msg, icon=icon)
    st.session_state.flash = None

vac = st.session_state.vacante_activa
conv_list = get_convocatorias_db(DB_PATH)
conv_guardada = any(c["codigo"] == vac["codigo_vacante"] for c in conv_list)

# ---------------------------------------------------------------------------
# Barra lateral: gestión de convocatorias
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("<div class='side-brand'><b>🩺 TalentValidator Pro</b><span>Auditoría clínica de talento humano</span></div>",
                unsafe_allow_html=True)
    st.markdown(f"{badge(st.session_state.user['username'], 'sky')} {badge(st.session_state.user['rol'], 'lav')}",
                unsafe_allow_html=True)
    if st.session_state.user.get("api_key"):
        with st.expander("🔑 API Key"):
            st.code(st.session_state.user["api_key"], language=None)

    st.divider()
    st.markdown("##### 📁 Convocatorias")
    if st.button("➕ Nueva convocatoria", width="stretch", type="primary"):
        nv = nueva_vacante()
        reset_proceso(nv)
        flash(f"Convocatoria {nv['codigo_vacante']} iniciada", "🆕")
        st.rerun()

    if conv_list:
        opciones_c = {f"{'🔒' if c['estado'] == 'CERRADA' else '🟢'} {c['codigo']} · {c['nombre_puesto']}": c["codigo"]
                      for c in conv_list}
        conv_elegida = st.selectbox("Abrir convocatoria guardada", list(opciones_c.keys()), index=None,
                                    placeholder="Seleccione una convocatoria…")
        if conv_elegida and st.button("📂 Cargar convocatoria", width="stretch"):
            data = load_convocatoria(opciones_c[conv_elegida], DB_PATH)
            if data:
                reset_proceso(data["vacante"], data["evaluaciones"], data["entrevistas"], data["entrevistas_detalle"],
                              archivada=bool(data["evaluaciones"]))
                flash(f"{data['vacante']['codigo_vacante']} cargada con {len(data['evaluaciones'])} aspirantes", "📂")
                st.rerun()
    else:
        st.caption("Aún no hay convocatorias guardadas.")

    st.divider()
    if st.button("🚪 Cerrar sesión", width="stretch"):
        st.session_state.clear()
        st.rerun()

# ---------------------------------------------------------------------------
# Encabezado con progreso del proceso
# ---------------------------------------------------------------------------
pasos = [("Convocatoria", conv_guardada), ("Criba de CV", bool(st.session_state.evaluaciones)),
         ("Entrevista", bool(st.session_state.entrevistas)), ("Terna y cierre", st.session_state.archivada)]
primer_pendiente = next((i for i, (_, ok) in enumerate(pasos) if not ok), None)
steps_html = "".join(
    f"<div class='step {'done' if ok else ('current' if i == primer_pendiente else '')}'>"
    f"<span class='dot'>{'✓' if ok else i + 1}</span>{name}</div>"
    for i, (name, ok) in enumerate(pasos))
st.markdown(f"""
    <div class='hero'>
      <div class='hero-row'>
        <div><h1>🏥 Selección y Auditoría Clínica de Talento</h1>
        <p>Verificación de ReTHUS, criba multicriterio, entrevista personalizada y analítica de competencias.</p></div>
        <div class='conv-pill'><small>Convocatoria activa</small><b>{esc(vac['codigo_vacante'])}</b>
        <span style='font-size:.8rem;color:#5F7F7B;'>{esc(vac['nombre_puesto'])}</span></div>
      </div>
      <div class='stepper'>{steps_html}</div>
    </div>
""", unsafe_allow_html=True)

tabs_labels = ["📋 Convocatoria", "🔬 Criba de hojas de vida", "🗣️ Entrevista técnica",
               "🏆 Terna y estadísticas", "🔐 Mi cuenta"]
es_admin = st.session_state.user['rol'] == 'SuperAdmin'
if es_admin:
    tabs_labels.append("👥 Administración")
tabs = st.tabs(tabs_labels)

# ---------------------------------------------------------------------------
# Pestaña 1: Convocatoria y requisición
# ---------------------------------------------------------------------------
with tabs[0]:
    section("📋 Parametrización y codificación de la convocatoria")
    col_up, col_form = st.columns([1, 1.25], gap="large")
    with col_up:
        with st.container(key="cardsky_req"):
            card_header("📤 Carga automática de requisición",
                        "Suba el perfil del cargo en PDF o Excel. Se reconocen líneas como "
                        "<i>Cargo:</i>, <i>Experiencia:</i>, <i>Tarjeta profesional:</i> y <i>Competencias:</i>.")
            req_file = st.file_uploader("Documento de requisición", type=["pdf", "xlsx"],
                                        key=f"req_{st.session_state.lote}", label_visibility="collapsed")
            if req_file is not None:
                file_id = (req_file.name, req_file.size)
                if file_id != st.session_state.req_file_id:
                    try:
                        if req_file.name.lower().endswith(".pdf"):
                            parsed = parse_requisicion_text(extract_text_from_pdf_bytes(req_file.read()))
                        else:
                            df_r = pd.read_excel(req_file)
                            parsed = parse_requisicion_text(excel_to_requisicion_text(df_r))
                            if parsed["nombre_puesto"] == "Puesto Requerido" and not df_r.empty:
                                parsed["nombre_puesto"] = str(df_r.iloc[0, 0])
                        parsed["codigo_vacante"] = vac["codigo_vacante"]
                        if not parsed["skills_obligatorias"]:
                            parsed["skills_obligatorias"] = vac["skills_obligatorias"]
                        st.session_state.vacante_activa = parsed
                        st.session_state.req_file_id = file_id
                        flash("Requisición procesada: revise y guarde los parámetros", "📄")
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo leer el documento: {e}")
                else:
                    st.markdown("<div class='alert-soft alert-mint'>✅ Requisición aplicada al formulario.</div>",
                                unsafe_allow_html=True)
        with st.container(key="card_resumen_conv"):
            card_header("🧾 Resumen vigente")
            st.markdown(
                f"**Cargo:** {esc(vac['nombre_puesto'])}<br>**Experiencia mínima:** {vac['meses_experiencia']} meses<br>"
                f"**Tarjeta / ReTHUS:** {'Obligatoria' if vac['requiere_tarjeta'] else 'Opcional'}<br>"
                f"**Estado:** {badge('Guardada', 'ok') if conv_guardada else badge('Sin guardar', 'peach')}",
                unsafe_allow_html=True)
            st.markdown(" ".join(badge(s, "sky") for s in vac["skills_obligatorias"]) or "Sin competencias definidas",
                        unsafe_allow_html=True)

    with col_form:
        with st.container(key="card_form_conv"):
            card_header("⚙️ Identificación y requisitos excluyentes")
            with st.form("vacante_form"):
                v_cod = st.text_input("Código único de la convocatoria", value=vac["codigo_vacante"])
                v_nom = st.text_input("Denominación del cargo", value=vac["nombre_puesto"])
                c1, c2 = st.columns(2)
                v_exp = c1.number_input("Experiencia mínima (meses)", min_value=0, max_value=600,
                                        value=int(vac["meses_experiencia"]), step=6)
                c2.write("")
                v_tar = c2.toggle("Exigir Tarjeta Profesional / ReTHUS", value=bool(vac["requiere_tarjeta"]))
                v_sk = st.text_area("Competencias clínicas y normativas obligatorias (separadas por coma)",
                                    value=", ".join(vac["skills_obligatorias"]), height=110)
                if st.form_submit_button("💾 Guardar convocatoria", type="primary", width="stretch"):
                    codigo = v_cod.strip()
                    if not codigo or not v_nom.strip():
                        st.error("El código y el nombre del cargo son obligatorios.")
                    else:
                        sk_list = [s.strip() for s in v_sk.split(",") if s.strip()]
                        st.session_state.vacante_activa = {
                            "codigo_vacante": codigo, "nombre_puesto": v_nom.strip(),
                            "meses_experiencia": int(v_exp), "requiere_tarjeta": v_tar, "skills_obligatorias": sk_list}
                        save_convocatoria_db(codigo, v_nom.strip(), int(v_exp), v_tar, sk_list, DB_PATH)
                        flash(f"Convocatoria {codigo} guardada", "💾")
                        st.rerun()

# ---------------------------------------------------------------------------
# Pestaña 2: Criba de hojas de vida
# ---------------------------------------------------------------------------
with tabs[1]:
    section("🔬 Criba asistida de hojas de vida y auditoría manual")
    with st.container(key="cardsky_upload"):
        cu, ca = st.columns([2.2, 1], vertical_alignment="bottom")
        with cu:
            cv_files = st.file_uploader(f"📂 Hojas de vida en PDF para {vac['codigo_vacante']} (máximo 10)",
                                        type=["pdf"], accept_multiple_files=True, key=f"cv_{st.session_state.lote}")
        with ca:
            analizar_btn = st.button("🚀 Analizar y calificar", type="primary", width="stretch",
                                     disabled=not cv_files)
        if cv_files and len(cv_files) > 10:
            st.error("Límite superado: cargue como máximo 10 archivos PDF por lote.")
        elif cv_files and analizar_btn:
            res_list, usados, fallidos = [], set(), []
            with st.status("Analizando hojas de vida…", expanded=False) as status:
                for f in cv_files:
                    try:
                        txt = extract_text_from_pdf_bytes(f.read())
                    except Exception:
                        fallidos.append(f.name)
                        continue
                    base = f.name.rsplit(".", 1)[0].replace("_", " ").strip().title()
                    nombre, n = base, 2
                    while nombre in usados:
                        nombre, n = f"{base} ({n})", n + 1
                    usados.add(nombre)
                    res_list.append(build_candidate_record(nombre, txt, vac))
                status.update(label=f"{len(res_list)} hojas de vida analizadas", state="complete")
            res_list.sort(key=lambda x: (x["knockout"] == "APROBADO", x["score_cv"]), reverse=True)
            st.session_state.evaluaciones = res_list
            st.session_state.entrevistas, st.session_state.entrevistas_detalle = {}, {}
            st.session_state.archivada = False
            st.session_state.rev += 1
            if fallidos:
                st.warning(f"No se pudieron leer: {', '.join(fallidos)}")
            flash(f"Análisis completado: {len(res_list)} aspirantes", "🔬")
            st.rerun()

    evaluaciones = st.session_state.evaluaciones
    if not evaluaciones:
        st.info("Cargue hojas de vida en PDF y pulse **Analizar y calificar** para iniciar la criba.")
    else:
        total = len(evaluaciones)
        aprob = sum(1 for r in evaluaciones if r["knockout"] == "APROBADO")
        audit = sum(1 for r in evaluaciones if r.get("validado_manualmente"))
        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(kpi("👥", total, "Postulados", "sky"), unsafe_allow_html=True)
        k2.markdown(kpi("✅", aprob, "Aprobados fase 1", "mint"), unsafe_allow_html=True)
        k3.markdown(kpi("⛔", total - aprob, "Descartados", "rose"), unsafe_allow_html=True)
        k4.markdown(kpi("🩺", audit, "Auditados manualmente", "lav"), unsafe_allow_html=True)
        st.write("")

        df_ev = pd.DataFrame([{
            "Aspirante": r["nombre"],
            "Fase 1": "✅ APROBADO" if r["knockout"] == "APROBADO" else "⛔ DESCARTADO",
            "Validación": "🩺 Manual" if r.get("validado_manualmente") else "🤖 Automática",
            "Puntaje CV": r["score_cv"], "Trayectoria": r["m2"], "Skills": r["m3"],
            "Logros": r["m4"], "Formato": r["m5"],
            "Observación": r["motivo_descarte"] or "Cumple requisitos",
        } for r in evaluaciones])
        st.dataframe(df_ev, width="stretch", hide_index=True, column_config={
            "Puntaje CV": st.column_config.ProgressColumn("Puntaje CV", min_value=0, max_value=100, format="%.1f"),
            "Trayectoria": st.column_config.NumberColumn("Trayectoria /35", format="%.1f"),
            "Skills": st.column_config.NumberColumn("Skills /30", format="%.1f"),
            "Logros": st.column_config.NumberColumn("Logros /25", format="%.1f"),
            "Formato": st.column_config.NumberColumn("Formato /10", format="%.1f"),
        })

        st.write("")
        with st.container(key="cardlav_inspector"):
            card_header("🩺 Inspección estructurada y validación personal",
                        "Revise la información de la hoja de vida organizada en tablas y corrija lo que la lectura "
                        "automática no haya detectado.")
            nombres = [r["nombre"] for r in evaluaciones]
            sel = st.selectbox("Aspirante a inspeccionar", nombres, key="sel_insp")
            cand = next(r for r in evaluaciones if r["nombre"] == sel)
            texto = cand.get("texto_cv", "")
            tablas = analyze_cv_for_tables(texto, vac, cand)

            col_t, col_v = st.columns([1.45, 1], gap="large")
            with col_t:
                if not texto.strip():
                    st.markdown("<div class='alert-soft alert-peach'>⚠️ No se extrajo texto del PDF (posible "
                                "documento escaneado). Valide manualmente con el documento original.</div>",
                                unsafe_allow_html=True)
                st.markdown("**📋 Tabla 1 · Datos generales y requisitos clave**")
                st.dataframe(tablas["datos_generales"], width="stretch", hide_index=True)
                st.markdown("**🎯 Tabla 2 · Competencias exigidas en la convocatoria**")
                st.dataframe(tablas["skills"], width="stretch", hide_index=True)
                st.markdown("**⚖️ Tabla 3 · Calificación por módulos (CV sobre 100)**")
                st.dataframe(tablas["modulos"], width="stretch", hide_index=True, column_config={
                    "Puntaje": st.column_config.NumberColumn(format="%.1f"),
                    "Máximo": st.column_config.NumberColumn(format="%.0f")})
                st.markdown("**📈 Tabla 4 · Evidencias de logros e indicadores**")
                st.dataframe(tablas["logros"], width="stretch", hide_index=True)
                with st.expander("📄 Texto original extraído del documento"):
                    termino = st.text_input("🔍 Buscar en el texto", key=f"buscar_{sel}")
                    if termino:
                        st.caption(f"{texto.lower().count(termino.lower())} coincidencia(s) de «{termino}».")
                    st.text_area("Texto extraído", value=texto, height=260, disabled=True,
                                 label_visibility="collapsed", key=f"txt_{st.session_state.rev}_{sel}")

            with col_v:
                estado_b = badge(cand["knockout"], "ok" if cand["knockout"] == "APROBADO" else "ko")
                valid_b = badge("Validado manualmente", "lav") if cand.get("validado_manualmente") \
                    else badge("Evaluación automática", "sky")
                st.markdown(f"{estado_b} {valid_b}", unsafe_allow_html=True)
                if cand.get("motivo_descarte"):
                    st.markdown(f"<div class='alert-soft alert-rose'><b>Causal:</b> {esc(cand['motivo_descarte'])}</div>",
                                unsafe_allow_html=True)

                k = f"{st.session_state.rev}_{sel}"
                st.markdown("**1 · Compuerta de requisitos excluyentes**")
                val_tarjeta = st.checkbox("Certifico Tarjeta Profesional / ReTHUS vigente",
                                          value=bool(cand.get("tarjeta_detectada")), key=f"tar_{k}",
                                          help="Márquelo si verificó el registro en ReTHUS o en el documento físico.")
                meses_def = cand.get("meses_acreditados")
                if meses_def is None:
                    meses_def = cand.get("meses_detectados", 0)
                val_exp = st.number_input("Meses de experiencia acreditados", min_value=0, max_value=600,
                                          value=int(meses_def), step=6, key=f"exp_{k}")

                st.markdown("**2 · Calificación por módulos**")
                vm2 = st.slider("Trayectoria y liderazgo (máx. 35)", 0.0, 35.0, float(cand.get("m2", 0.0)), 1.0, key=f"m2_{k}")
                vm3 = st.slider("Hard skills y normativa (máx. 30)", 0.0, 30.0, float(cand.get("m3", 0.0)), 1.0, key=f"m3_{k}")
                vm4 = st.slider("Logros y gestión (máx. 25)", 0.0, 25.0, float(cand.get("m4", 0.0)), 1.0, key=f"m4_{k}")
                vm5 = st.slider("Calidad formal (máx. 10)", 0.0, 10.0, float(cand.get("m5", 0.0)), 1.0, key=f"m5_{k}")
                score_manual = round(vm2 + vm3 + vm4 + vm5, 1)
                st.markdown(f"<div class='alert-soft alert-mint' style='text-align:center;'>Puntaje CV recalculado"
                            f"<div class='score-big'>{score_manual} / 100</div>{barra(score_manual, 100)}</div>",
                            unsafe_allow_html=True)

                st.markdown("**3 · Dictamen del evaluador**")
                decision = st.radio("Dictamen", ["✅ APROBADO", "⛔ DESCARTADO"], horizontal=True,
                                    index=0 if cand["knockout"] == "APROBADO" else 1, key=f"dec_{k}",
                                    label_visibility="collapsed")
                aprobado = decision.endswith("APROBADO")
                motivo = st.text_input("Motivo del descarte", disabled=aprobado, key=f"mot_{k}",
                                       value=cand.get("motivo_descarte") or "No cumple los requisitos mínimos exigidos.")
                notas = st.text_area("Observaciones de auditoría", value=cand.get("notas_auditor", ""), key=f"not_{k}",
                                     placeholder="Ej.: Registro verificado en ReTHUS. Experiencia de 3 años en IPS acreditada.")

                if st.button("💾 Guardar validación", type="primary", width="stretch", key=f"save_{k}"):
                    errores = []
                    if aprobado and vac["requiere_tarjeta"] and not val_tarjeta:
                        errores.append("La convocatoria exige Tarjeta Profesional / ReTHUS: certifíquela o descarte al aspirante.")
                    if aprobado and val_exp < int(vac["meses_experiencia"]):
                        errores.append(f"La experiencia acreditada ({val_exp} meses) es inferior a la exigida "
                                       f"({vac['meses_experiencia']} meses).")
                    if not aprobado and not motivo.strip():
                        errores.append("Indique el motivo del descarte.")
                    for e in errores:
                        st.error(e)
                    if not errores:
                        fortalezas = [f for f in cand.get("fortalezas", []) if f != "Validación personalizada por el evaluador."]
                        if aprobado:
                            fortalezas.append("Validación personalizada por el evaluador.")
                        cand.update({
                            "knockout": "APROBADO" if aprobado else "DESCARTADO",
                            "score_cv": score_manual if aprobado else 0.0,
                            "m2": vm2 if aprobado else 0.0, "m3": vm3 if aprobado else 0.0,
                            "m4": vm4 if aprobado else 0.0, "m5": vm5 if aprobado else 0.0,
                            "motivo_descarte": None if aprobado else motivo.strip(),
                            "fortalezas": fortalezas if aprobado else [],
                            "debilidades": ([d for d in cand.get("debilidades", []) if d != cand.get("motivo_descarte")]
                                            if aprobado else [motivo.strip()]),
                            "validado_manualmente": True, "tarjeta_detectada": val_tarjeta,
                            "meses_acreditados": int(val_exp), "notas_auditor": notas.strip(),
                        })
                        if not aprobado:
                            st.session_state.entrevistas.pop(sel, None)
                            st.session_state.entrevistas_detalle.pop(sel, None)
                        evaluaciones.sort(key=lambda x: (x["knockout"] == "APROBADO", x["score_cv"]), reverse=True)
                        st.session_state.archivada = False
                        flash(f"Validación de {sel} guardada", "🩺")
                        st.rerun()

# ---------------------------------------------------------------------------
# Pestaña 3: Entrevista técnica personalizada
# ---------------------------------------------------------------------------
with tabs[2]:
    section("🗣️ Entrevista técnica personalizada por aspirante")
    aprobados = [r for r in st.session_state.evaluaciones if r["knockout"] == "APROBADO"]
    if not aprobados:
        st.info("No hay aspirantes aprobados en la fase 1. Complete la criba en la pestaña de hojas de vida.")
    else:
        entrevistados = sum(1 for r in aprobados if r["nombre"] in st.session_state.entrevistas_detalle)
        st.markdown(f"<div class='alert-soft alert-mint'>🗂️ Entrevistas registradas: <b>{entrevistados} de "
                    f"{len(aprobados)}</b> aspirantes aprobados.</div>{barra(entrevistados, len(aprobados))}",
                    unsafe_allow_html=True)

        etiquetas = {r["nombre"]: f"{'✅' if r['nombre'] in st.session_state.entrevistas_detalle else '⏳'} {r['nombre']}"
                     for r in aprobados}
        cand_sel = st.selectbox("Aspirante a entrevistar", list(etiquetas.keys()), format_func=etiquetas.get,
                                key="sel_entrevista")
        cand_obj = next(r for r in aprobados if r["nombre"] == cand_sel)
        preguntas = generate_candidate_interview_questions(cand_obj, vac)
        previas = st.session_state.entrevistas_detalle.get(cand_sel, [])

        st.markdown(
            f"{badge('Score CV: ' + str(cand_obj['score_cv']), 'ok')} {badge(vac['nombre_puesto'], 'sky')} "
            f"{badge('Preguntas adaptadas a sus fortalezas y brechas', 'lav')}", unsafe_allow_html=True)
        st.write("")

        etiquetas_op = list(OPCIONES_ENTREVISTA.keys())
        registro, pts_total, respondidas = [], 0, 0
        for i, p in enumerate(preguntas):
            prev = previas[i] if i < len(previas) else {}
            idx_prev = None
            if "puntaje" in prev:
                idx_prev = min(range(len(etiquetas_op)),
                               key=lambda j: abs(OPCIONES_ENTREVISTA[etiquetas_op[j]] - prev["puntaje"]))
            kq = f"{st.session_state.rev}_{cand_sel}_{i}"
            with st.container(key=f"cardq_{kq}"):
                st.markdown(f"<div class='card-title'>📌 Pregunta {i + 1} · {esc(p['tema'])}</div>"
                            f"<div style='font-weight:600;margin-bottom:4px;'>{esc(p['pregunta'])}</div>"
                            f"<div class='card-sub'>🎯 Objetivo: {esc(p['foco'])}</div>", unsafe_allow_html=True)
                cq, co = st.columns([1, 1.9], gap="medium")
                with cq:
                    resp = st.radio(f"Calificación P{i + 1}", etiquetas_op, index=idx_prev, key=f"cal_{kq}")
                with co:
                    obs = st.text_area(f"Observación sobre la respuesta (P{i + 1})", value=prev.get("observacion", ""),
                                       key=f"obs_{kq}", height=100,
                                       placeholder="Resuma lo que respondió el aspirante, argumentos y evidencias…")
            puntaje = OPCIONES_ENTREVISTA[resp] if resp else 0
            respondidas += 1 if resp else 0
            pts_total += puntaje
            registro.append({"pregunta_num": i + 1, "tema": p["tema"], "enunciado": p["pregunta"],
                             "puntaje": puntaje, "observacion": obs.strip()})

        cs, cb = st.columns([1.3, 1], vertical_alignment="center")
        cs.markdown(f"<div class='alert-soft alert-mint'>Puntaje de la entrevista "
                    f"<div class='score-big'>{pts_total} / 100</div>{barra(pts_total, 100, '#64B5F6')}"
                    f"<span style='font-size:.8rem;'>{respondidas} de {len(preguntas)} preguntas calificadas</span></div>",
                    unsafe_allow_html=True)
        if cb.button("💾 Guardar entrevista y observaciones", type="primary", width="stretch",
                     key=f"save_ent_{st.session_state.rev}_{cand_sel}"):
            if respondidas < len(preguntas):
                st.error("Califique todas las preguntas antes de guardar.")
            else:
                st.session_state.entrevistas[cand_sel] = float(pts_total)
                st.session_state.entrevistas_detalle[cand_sel] = registro
                st.session_state.archivada = False
                flash(f"Entrevista de {cand_sel} guardada ({pts_total} pts)", "🗣️")
                st.rerun()

# ---------------------------------------------------------------------------
# Pestaña 4: Terna, analítica y cierre
# ---------------------------------------------------------------------------
with tabs[3]:
    section("🏆 Terna final y analítica de competencias")
    resumen = []
    for r in st.session_state.evaluaciones:
        s_ent = float(st.session_state.entrevistas.get(r["nombre"], 0.0))
        s_fin = round(r["score_cv"] * 0.6 + s_ent * 0.4, 1) if r["knockout"] == "APROBADO" else 0.0
        resumen.append({**r, "score_entrevista": s_ent, "score_final": s_fin,
                        "entrevista_detalle": st.session_state.entrevistas_detalle.get(r["nombre"], [])})
    resumen.sort(key=lambda x: (x["knockout"] == "APROBADO", x["score_final"]), reverse=True)
    aprob_res = [c for c in resumen if c["knockout"] == "APROBADO"]
    top_3 = aprob_res[:3]

    if not top_3:
        st.info("Aún no hay aspirantes aprobados para conformar la terna.")
    else:
        sin_entrevista = [c["nombre"] for c in top_3 if c["nombre"] not in st.session_state.entrevistas_detalle]
        if sin_entrevista:
            st.markdown(f"<div class='alert-soft alert-peach'>⏳ Pendiente de entrevista: {esc(', '.join(sin_entrevista))}. "
                        f"Su puntaje de entrevista cuenta como 0.</div>", unsafe_allow_html=True)

        cols_t = st.columns(len(top_3))
        for idx, c in enumerate(top_3):
            with cols_t[idx]:
                st.markdown(f"""
                    <div class='podium'>
                      <div class='medal'>{MEDALLAS[idx]}</div>
                      <h3>{esc(c['nombre'])}</h3>
                      <div class='score-big'>{c['score_final']}</div>
                      <div style='font-size:.8rem;color:#5F7F7B;'>puntaje consolidado / 100</div>
                      {barra(c['score_final'], 100, COLORES_TERNA[idx][1])}
                      <div style='font-size:.82rem;text-align:left;'>📄 CV {c['score_cv']} {barra(c['score_cv'], 100, '#80CBC4')}
                      🗣️ Entrevista {c['score_entrevista']} {barra(c['score_entrevista'], 100, '#90CAF9')}</div>
                    </div>""", unsafe_allow_html=True)
                with st.expander("Fortalezas y brechas"):
                    for f in c["fortalezas"]:
                        st.markdown(f"🟢 {f}")
                    for d in c["debilidades"]:
                        st.markdown(f"🟠 {d}")
                    if not c["fortalezas"] and not c["debilidades"]:
                        st.caption("Sin observaciones registradas.")

        st.write("")
        section("📊 Gráficas de habilidades para la toma de decisiones")
        g1, g2 = st.columns([1.1, 0.9], gap="large")
        with g1:
            with st.container(key="card_radar"):
                cats = ["Liderazgo", "Hard skills", "Logros", "Calidad formal", "Entrevista"]
                fig_radar = go.Figure()
                for idx, c in enumerate(top_3):
                    vals = [c["m2"] / 35 * 100, c["m3"] / 30 * 100, c["m4"] / 25 * 100, c["m5"] / 10 * 100,
                            c["score_entrevista"]]
                    fill, line = COLORES_TERNA[idx]
                    fig_radar.add_trace(go.Scatterpolar(r=vals + vals[:1], theta=cats + cats[:1], fill="toself",
                                                        name=f"{MEDALLAS[idx]} {c['nombre']}", fillcolor=fill,
                                                        line=dict(color=line, width=2)))
                fig_radar.update_layout(
                    title=dict(text="Perfil de competencias (% del máximo)", font=dict(size=14, color="#00695C")),
                    polar=dict(bgcolor="rgba(224,242,241,.35)", radialaxis=dict(range=[0, 100], tickfont=dict(size=9))),
                    height=380, margin=dict(l=40, r=40, t=50, b=20), paper_bgcolor="rgba(0,0,0,0)",
                    legend=dict(orientation="h", y=-0.12), font=dict(family="Nunito"))
                st.plotly_chart(fig_radar, width="stretch")
        with g2:
            with st.container(key="card_stack"):
                nombres_t = [f"{MEDALLAS[i]} {c['nombre']}" for i, c in enumerate(top_3)]
                fig_bar = go.Figure()
                fig_bar.add_trace(go.Bar(name="CV (60%)", y=nombres_t, x=[round(c["score_cv"] * .6, 1) for c in top_3],
                                         orientation="h", marker_color="#80CBC4"))
                fig_bar.add_trace(go.Bar(name="Entrevista (40%)", y=nombres_t,
                                         x=[round(c["score_entrevista"] * .4, 1) for c in top_3],
                                         orientation="h", marker_color="#90CAF9"))
                fig_bar.update_layout(
                    barmode="stack", title=dict(text="Composición del puntaje final", font=dict(size=14, color="#00695C")),
                    xaxis=dict(range=[0, 100], title="Puntos"), yaxis=dict(autorange="reversed"),
                    height=380, margin=dict(l=10, r=10, t=50, b=30), paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)", legend=dict(orientation="h", y=-0.2), font=dict(family="Nunito"))
                st.plotly_chart(fig_bar, width="stretch")

        g3, g4 = st.columns([1.1, 0.9], gap="large")
        with g3:
            with st.container(key="card_heat"):
                mat = pd.DataFrame(
                    [[c["m2"] / 35 * 100, c["m3"] / 30 * 100, c["m4"] / 25 * 100, c["m5"] / 10 * 100,
                      c["score_entrevista"]] for c in aprob_res],
                    index=[c["nombre"] for c in aprob_res], columns=cats).round(0)
                fig_heat = px.imshow(mat, text_auto=True, aspect="auto", zmin=0, zmax=100,
                                     color_continuous_scale=["#FCE4EC", "#FFF3E0", "#E0F2F1", "#4DB6AC"])
                fig_heat.update_layout(title=dict(text="Mapa de calor de habilidades · todos los aprobados",
                                                  font=dict(size=14, color="#00695C")),
                                       height=max(260, 70 + 42 * len(aprob_res)), margin=dict(l=10, r=10, t=50, b=10),
                                       paper_bgcolor="rgba(0,0,0,0)", coloraxis_showscale=False, font=dict(family="Nunito"))
                st.plotly_chart(fig_heat, width="stretch")
        with g4:
            with st.container(key="card_skills"):
                card_header("🎯 Cobertura de competencias obligatorias")
                skills_vac = vac.get("skills_obligatorias", [])
                if skills_vac:
                    matriz = [{"Competencia": s, **{c["nombre"]: ("✅" if s.lower() in c.get("texto_cv", "").lower() else "⚠️")
                                                    for c in top_3}} for s in skills_vac]
                    st.dataframe(pd.DataFrame(matriz), width="stretch", hide_index=True)
                    cobertura = [sum(s.lower() in c.get("texto_cv", "").lower() for s in skills_vac) / len(skills_vac) * 100
                                 for c in top_3]
                    for idx, c in enumerate(top_3):
                        st.markdown(f"<span style='font-size:.85rem;'>{MEDALLAS[idx]} {esc(c['nombre'])} · "
                                    f"<b>{cobertura[idx]:.0f}%</b></span>{barra(cobertura[idx], 100, COLORES_TERNA[idx][1])}",
                                    unsafe_allow_html=True)
                else:
                    st.caption("La convocatoria no define competencias obligatorias.")

        st.write("")
        with st.container(key="cardlav_cierre"):
            card_header("✍️ Decisión final y cierre de la convocatoria",
                        "Al confirmar, los resultados se guardan bajo el código de la convocatoria. "
                        "Volver a confirmar reemplaza el registro anterior (sin duplicados).")
            cl, cr = st.columns([1.5, 1], gap="large")
            with cl:
                elegido = st.selectbox("Aspirante seleccionado", [c["nombre"] for c in top_3], key="sel_elegido")
                correo_el = st.text_input("Correo del aspirante (opcional, para enviar el acta)", key="correo_elegido")
            pdf_bytes = generate_terna_pdf(vac, top_3, elegido, st.session_state.user["username"])
            with cr:
                st.download_button(f"📥 Acta de terna ({vac['codigo_vacante']})", data=pdf_bytes,
                                   file_name=f"Acta_Terna_{vac['codigo_vacante']}.pdf", mime="application/pdf",
                                   width="stretch")
                if st.button("✅ Confirmar y guardar resultados", type="primary", width="stretch", key="btn_cierre"):
                    n = save_convocatoria_results(vac, resumen, elegido, correo_el.strip() or None, DB_PATH)
                    st.session_state.archivada = True
                    msg = f"{vac['codigo_vacante']} cerrada: {n} aspirantes guardados"
                    if correo_el.strip():
                        cfg = smtp_config()
                        if cfg:
                            res_mail = send_selection_email(cfg["host"], int(cfg.get("port", 587)), cfg["sender_email"],
                                                            cfg["password"], correo_el.strip(), elegido,
                                                            vac["nombre_puesto"], vac["codigo_vacante"], pdf_bytes)
                            msg += " · " + res_mail["message"]
                        else:
                            msg += " · Correo no enviado (SMTP sin configurar)"
                    flash(msg, "🏁")
                    st.rerun()

    st.write("")
    section("🗃️ Histórico de convocatorias")
    conn = get_db_connection(DB_PATH)
    df_raw = pd.read_sql_query("""
        SELECT vacante_codigo AS "Convocatoria", nombre AS "Aspirante", email AS "Correo",
               estado_knockout AS "Fase 1", score_cv AS "Score CV", score_entrevista AS "Score entrevista",
               score_final AS "Score consolidado", CASE WHEN seleccionado = 1 THEN 'SÍ' ELSE 'NO' END AS "Seleccionado",
               fecha_evaluacion AS "Fecha"
        FROM candidatos_historico ORDER BY id DESC
    """, conn)
    conn.close()

    if df_raw.empty:
        st.caption("Todavía no hay resultados archivados.")
    else:
        f1, f2 = st.columns(2)
        v_sel = f1.selectbox("Convocatoria", ["TODAS"] + sorted(df_raw["Convocatoria"].unique().tolist()), key="hist_conv")
        e_sel = f2.selectbox("Estado", ["TODOS", "SÓLO SELECCIONADOS", "NO SELECCIONADOS"], key="hist_estado")
        df_f = df_raw.copy()
        if v_sel != "TODAS":
            df_f = df_f[df_f["Convocatoria"] == v_sel]
        if e_sel == "SÓLO SELECCIONADOS":
            df_f = df_f[df_f["Seleccionado"] == "SÍ"]
        elif e_sel == "NO SELECCIONADOS":
            df_f = df_f[df_f["Seleccionado"] == "NO"]

        h1, h2 = st.columns(2, gap="large")
        with h1:
            fig_fun = go.Figure(go.Funnel(
                y=["Postulados", "Superan fase 1", "Seleccionados"],
                x=[len(df_f), int((df_f["Fase 1"] == "APROBADO").sum()), int((df_f["Seleccionado"] == "SÍ").sum())],
                marker={"color": ["#B3E5FC", "#80CBC4", "#26A69A"]}))
            fig_fun.update_layout(title_text="Embudo de selección", height=280, margin=dict(l=10, r=10, t=40, b=10),
                                  paper_bgcolor="rgba(0,0,0,0)", font=dict(family="Nunito"))
            st.plotly_chart(fig_fun, width="stretch")
        with h2:
            fig_hist = px.histogram(df_f, x="Score consolidado", color="Fase 1", nbins=10,
                                    color_discrete_map={"APROBADO": "#4DB6AC", "DESCARTADO": "#F48FB1"})
            fig_hist.update_layout(title_text="Distribución de puntajes", height=280, margin=dict(l=10, r=10, t=40, b=10),
                                   paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Nunito"))
            st.plotly_chart(fig_hist, width="stretch")

        st.dataframe(df_f, width="stretch", hide_index=True, column_config={
            "Score consolidado": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.1f")})
        st.download_button("📊 Descargar histórico (.xlsx)", data=generate_excel_historico(df_f),
                           file_name=f"Historico_{v_sel}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------------------
# Pestaña 5: Cuenta
# ---------------------------------------------------------------------------
with tabs[4]:
    section("🔐 Seguridad de la cuenta")
    _, cc, _ = st.columns([0.2, 1, 0.2])
    with cc, st.container(key="card_pass"):
        card_header("Cambiar contraseña")
        with st.form("cambio_pass", clear_on_submit=True):
            p_old = st.text_input("Contraseña actual", type="password")
            p_new = st.text_input("Nueva contraseña (mínimo 8 caracteres)", type="password")
            p_conf = st.text_input("Confirmar nueva contraseña", type="password")
            if st.form_submit_button("Actualizar contraseña", type="primary"):
                if p_new != p_conf or len(p_new) < 8:
                    st.error("Las contraseñas deben coincidir y tener al menos 8 caracteres.")
                else:
                    conn = get_db_connection(DB_PATH)
                    u_row = conn.execute("SELECT password_hash FROM usuarios WHERE id = ?",
                                         (st.session_state.user["id"],)).fetchone()
                    if u_row and bcrypt.checkpw(p_old.encode('utf-8'), u_row["password_hash"]):
                        conn.execute("UPDATE usuarios SET password_hash = ? WHERE id = ?",
                                     (bcrypt.hashpw(p_new.encode('utf-8'), bcrypt.gensalt()), st.session_state.user["id"]))
                        conn.commit()
                        st.success("Contraseña actualizada.")
                    else:
                        st.error("La contraseña actual no es correcta.")
                    conn.close()

# ---------------------------------------------------------------------------
# Pestaña 6: Administración (SuperAdmin)
# ---------------------------------------------------------------------------
if es_admin:
    with tabs[5]:
        section("👥 Evaluadores y licenciamiento")
        c_alta, c_lista = st.columns([1, 1.4], gap="large")
        with c_alta, st.container(key="cardsky_alta"):
            card_header("➕ Registrar evaluador")
            with st.form("alta_user", clear_on_submit=True):
                nu_user = st.text_input("Usuario o correo")
                nu_pass = st.text_input("Contraseña provisional (mínimo 6 caracteres)", type="password")
                nu_plan = st.selectbox("Plan", ["starter", "pro", "enterprise"])
                if st.form_submit_button("Registrar evaluador", type="primary", width="stretch"):
                    if not nu_user.strip() or len(nu_pass) < 6:
                        st.error("Indique un usuario y una contraseña de al menos 6 caracteres.")
                    else:
                        conn = get_db_connection(DB_PATH)
                        try:
                            conn.execute("""
                                INSERT INTO usuarios (username, password_hash, rol, activo, plan, api_key, fecha_creacion)
                                VALUES (?, ?, 'Evaluador', 1, ?, ?, ?)
                            """, (nu_user.strip(), bcrypt.hashpw(nu_pass.encode('utf-8'), bcrypt.gensalt()), nu_plan,
                                  f"key_live_{secrets.token_hex(12)}", datetime.now().isoformat()))
                            conn.commit()
                            flash(f"Evaluador {nu_user.strip()} registrado", "👤")
                            st.rerun()
                        except Exception:
                            st.error("El usuario ya se encuentra registrado.")
                        finally:
                            conn.close()
        with c_lista, st.container(key="card_lista"):
            card_header("📇 Evaluadores registrados")
            conn = get_db_connection(DB_PATH)
            users = conn.execute("SELECT id, username, plan, activo FROM usuarios WHERE rol != 'SuperAdmin'").fetchall()
            conn.close()
            if not users:
                st.caption("No hay evaluadores registrados.")
            for u in users:
                ca, cb, cc2 = st.columns([2, 1, 1], vertical_alignment="center")
                ca.markdown(f"**{esc(u['username'])}** {badge(u['plan'], 'sky')}", unsafe_allow_html=True)
                cb.markdown(badge("ACTIVO", "ok") if u["activo"] else badge("SUSPENDIDO", "ko"), unsafe_allow_html=True)
                if cc2.button("Suspender" if u["activo"] else "Habilitar", key=f"t_{u['id']}", width="stretch"):
                    conn = get_db_connection(DB_PATH)
                    conn.execute("UPDATE usuarios SET activo = ? WHERE id = ?", (0 if u["activo"] else 1, u["id"]))
                    conn.commit()
                    conn.close()
                    st.rerun()

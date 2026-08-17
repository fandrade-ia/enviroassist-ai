"""
EnviroAssist AI — Interfaz web
Módulo Fauna (BDIEET + cuadrícula UTM) | Módulo Flora (Anthos)
Normativa: LESRPE/CEEA (nacional) + Catálogo autonómico (en memoria)
"""
import sys, io, tempfile
from pathlib import Path
from datetime import datetime
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

import streamlit as st
from dotenv import load_dotenv
load_dotenv()

from modulo_fauna import AnalizadorFauna
from modulo_flora import AnalizadorFlora, _es_anthos
from normativa_autonomica import (
    CCAA_LISTA, cargar_lesrpe, extraer_texto_pdf,
    procesar_excel_normativa, procesar_pdf_normativa, procesar_docx_normativa,
    cruzar_con_normativa
)
from exportar_word import markdown_a_docx_bytes
from exportar_excel import generar_excel
from exportar_excel import generar_excel

# ── Página ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="EnviroAssist AI", page_icon="🌿",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:#f9fafb}
.block-container{padding-top:1.5rem}

/* ── Sidebar general ─────────────────────────────────────────────── */
[data-testid="stSidebar"]{background:#f4f6f8;border-right:1px solid #e5e9ee}
[data-testid="stSidebar"] .block-container{padding-top:1.75rem}
[data-testid="stSidebar"] h1{font-size:1.55rem;font-weight:800;color:#0f172a;letter-spacing:-.01em;margin-bottom:.15rem}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"]{color:#6b7280;font-size:.92rem}
[data-testid="stSidebar"] h3{font-size:1.05rem;font-weight:700;color:#0f172a;margin-top:.25rem}
[data-testid="stSidebar"] hr{margin:1.1rem 0;border-color:#e5e9ee}
[data-testid="stSidebar"] label{font-size:.88rem;font-weight:500;color:#374151}

/* ── Inputs y selects estilo "card" ──────────────────────────────── */
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] > div{
    background:#ffffff;border:1px solid #e5e7eb;border-radius:10px;
    box-shadow:0 1px 2px rgba(16,24,40,.04);min-height:44px
}
[data-testid="stSidebar"] .stTextInput input:focus{
    border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,.15)
}
.stTextInput input, .stSelectbox div[data-baseweb="select"] > div{
    border-radius:10px !important
}

/* ── Tabs ─────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"]{gap:4px;border-bottom:1px solid #e5e7eb}
.stTabs [data-baseweb="tab"]{font-weight:600;font-size:1rem;padding:10px 22px;color:#374151}
.stTabs [aria-selected="true"]{color:#dc2626 !important}
.stTabs [data-baseweb="tab-highlight"]{background-color:#dc2626 !important;height:3px}

[data-testid="stToolbar"]{display:none !important}
#MainMenu{visibility:hidden}
footer{visibility:hidden}
header{visibility:hidden}
.alerta-EP{background:#ffedd5;border-left:4px solid #ea580c;padding:8px 12px;border-radius:6px;margin:3px 0}
.alerta-V{background:#fef9c3;border-left:4px solid #ca8a04;padding:8px 12px;border-radius:6px;margin:3px 0}
.alerta-OTRO{background:#dcfce7;border-left:4px solid #16a34a;padding:8px 12px;border-radius:6px;margin:3px 0}
.info-box{background:#eff6ff;border-left:4px solid #3b82f6;padding:12px 14px;border-radius:10px;margin:10px 0;font-size:.88em;line-height:1.5;box-shadow:0 1px 2px rgba(16,24,40,.04)}
.ok-box{background:#f0fdf4;border-left:4px solid #16a34a;padding:8px 12px;border-radius:6px;margin:4px 0}
.warn-box{background:#fff7ed;border-left:4px solid #f97316;padding:8px 12px;border-radius:6px;margin:4px 0}
</style>""", unsafe_allow_html=True)

# ── Session state init ─────────────────────────────────────────────────────────
if "normativas" not in st.session_state:
    st.session_state["normativas"] = {}  # El usuario añade las CCAA

# Cargar LESRPE
if "lesrpe" not in st.session_state:
    st.session_state["lesrpe"] = cargar_lesrpe()

LESRPE = st.session_state["lesrpe"]

def get_normativa(ccaa):
    return st.session_state["normativas"].get(ccaa, {})

def tiene_normativa(ccaa):
    return ccaa in st.session_state["normativas"]

def detectar_protegidas(nombres: list, ccaa: str) -> list:
    normativa = get_normativa(ccaa)
    alertas = []
    for especie in nombres:
        especie = str(especie).strip()
        partes = especie.split()
        if len(partes) < 2:
            continue
        clave2 = partes[0] + " " + partes[1]
        # LESRPE
        info_l = LESRPE.get(clave2) or LESRPE.get(especie)
        if info_l is None:
            for k in LESRPE:
                if k.lower().startswith(clave2.lower()):
                    info_l = LESRPE[k]; break
        # Autonómica
        info_a = cruzar_con_normativa(especie, normativa)
        if info_l or info_a:
            cat_l = info_l.get("categoria","—") if info_l else "—"
            nivel_l = {"En peligro de extinción":1,"Vulnerable":2,"LESRPE":3}.get(cat_l,99)
            nivel_a = info_a["nivel"] if info_a else 99
            alertas.append({
                "especie": especie,
                "nombre_comun": (info_a or info_l or {}).get("nombre_comun",""),
                "cat_lesrpe": cat_l,
                "cat_auto_cod": info_a["cat_cod"] if info_a else "—",
                "cat_auto_nombre": info_a["cat_nombre"] if info_a else "—",
                "color_cod": info_a["color_cod"] if info_a else "—",
                "emoji_auto": info_a["emoji"] if info_a else "",
                "nivel": min(nivel_l, nivel_a),
                "decreto_auto": info_a["decreto"] if info_a else "—",
            })
    alertas.sort(key=lambda x: x["nivel"])
    return alertas

def generar_informe_md(titulo, cuads, fuente, stats, alertas, listado_df, redaccion, ccaa):
    n_prot = len(alertas)
    tabla = f"| Especie | Nombre común | Grupo | LESRPE/CEEA | {ccaa} | Normativa |\n"
    tabla += "|---|---|---|---|---|---|\n"
    for a in alertas:
        tabla += (f"| *{a.get('especie','?')}* | {a.get('nombre_comun','')} | "
                  f"{a.get('grupo','—')} | {a.get('cat_lesrpe','—')} | "
                  f"{a.get('emoji_auto','')} {a.get('cat_auto_nombre','—')} | "
                  f"{a.get('decreto_auto','—')} |\n")
    tabla_bruta = listado_df.to_markdown(index=False) if listado_df is not None else ""
    return f"""# {titulo}
*Fuente: {fuente} | Cuadrículas: {', '.join(cuads) if cuads else '—'}*
*EnviroAssist AI · {datetime.now().strftime('%d/%m/%Y %H:%M')}*
*Base normativa: LESRPE/CEEA (RD 139/2011, MITECO junio 2025) + {ccaa}*

---

## 1. Resumen

| Métrica | Valor |
|---|---|
| Cuadrículas | {', '.join(cuads) if cuads else '—'} |
| Total especies/taxones | {stats.get('n_especies', len(alertas))} |
| Con estatus de protección | {n_prot} |

---

## 2. Especies con estatus de protección

{tabla if alertas else '_No se detectaron especies con estatus de protección._'}

**Fuentes normativas:**
- Nacional: RD 139/2011 — LESRPE/CEEA (MITECO, junio 2025)
- Autonómica: {get_normativa(ccaa).get('decreto', 'No disponible')}

---

## 3. Redacción técnica para el EsIA

{redaccion}

---

## 4. Listado completo

{tabla_bruta}

---
*⚠️ Borrador generado con IA. Revisar y validar antes de presentación oficial.*
"""

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("# 🌿 EnviroAssist AI")
    st.caption("Asistente IA para Inventarios de Biodiversidad")
    st.divider()
    st.markdown("### Datos del proyecto")
    proyecto  = st.text_input("Nombre del proyecto", "")
    ubicacion = st.text_input("Ubicación", "")
    ccaa_sel  = st.selectbox("Comunidad autónoma", ["Selecciona CCAA..."] + CCAA_LISTA, index=0, key="ccaa_sel")
    tipo      = st.selectbox("Tipo de proyecto", [
        "Selecciona una opción","Planta solar fotovoltaica","Parque eólico","Línea eléctrica",
        "Infraestructura viaria","Urbanización","Industria","Otro"])

    st.divider()
    disp = list(st.session_state["normativas"].keys())
    if disp:
        st.markdown("**CCAA cargadas:** " + " · ".join(f"✅{c}" for c in disp))
    st.divider()
    st.markdown(
        "<div class='info-box'><b>LESRPE/CEEA:</b> RD 139/2011<br>MITECO junio 2025</div>",
        unsafe_allow_html=True
    )
    st.caption("v1.0 · TFM Máster IA · 2025")

# Sincroniza el selector de CCAA de la pestaña "Normativa autonómica" con el
# del sidebar: si el usuario cambia la CCAA a la izquierda, el desplegable
# de la pestaña Normativa se actualiza automáticamente.
if ccaa_sel != "Selecciona CCAA..." and st.session_state.get("_ccaa_sel_prev") != ccaa_sel:
    st.session_state["ccaa_norm"] = ccaa_sel
st.session_state["_ccaa_sel_prev"] = ccaa_sel


# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_norm, tab_fauna, tab_flora = st.tabs(["⚙️ Normativa autonómica","🦅 Fauna — BDIEET/IEET","🌿 Flora — Anthos/IEET"])

# Notificar si viene del botón del sidebar
if st.session_state.get("ir_a_normativa"):
    st.session_state["ir_a_normativa"] = False
    st.toast("👆 Haz clic en la pestaña ⚙️ Normativa autonómica", icon="⚙️")


# ══════════════════════════════════════════════════════════════════════════════
# TAB FAUNA
# ══════════════════════════════════════════════════════════════════════════════
with tab_fauna:
    st.markdown(f"## 🦅 Inventario Faunístico — {ccaa_sel}")
    if tiene_normativa(ccaa_sel):
        n = get_normativa(ccaa_sel)
        st.markdown(f"<div class='ok-box'>✅ Normativa: {n.get('decreto','')}</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='warn-box'>➕ Añade normativa autonómica en la pestaña ⚙️</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([2,1])
    with col1:
        bdieet_file = st.file_uploader("📂 BDIEET (CSV o Excel)", type=["csv","xlsx","xls"], key="bdieet")
    with col2:
        cuads_input = st.text_input("📍 Cuadrícula(s) UTM 10x10", "", help="Varias separadas por coma. Ej: 30TVK15, 30TVK25")
        st.caption("🗺️ [Buscar en IGN Iberpix](https://www.ign.es/iberpix/visor/)")

    if bdieet_file:
        suf = Path(bdieet_file.name).suffix or ".csv"
        with tempfile.NamedTemporaryFile(suffix=suf, delete=False) as t:
            t.write(bdieet_file.read())
            st.session_state["ruta_bdieet"] = t.name
        st.success(f"✅ {bdieet_file.name}")

    if "ruta_bdieet" in st.session_state:
        cuads = [c.strip().upper() for c in cuads_input.split(",") if c.strip()]
        if cuads and st.button("🔬 Analizar fauna", type="primary"):
            af = AnalizadorFauna()
            af.ccaa = ccaa_sel
            with st.spinner("Cargando BDIEET..."):
                af.cargar_bdieet(Path(st.session_state["ruta_bdieet"]))
            with st.spinner(f"Filtrando {cuads}..."):
                af.filtrar_cuadriculas(cuads)

            if af.df is None or len(af.df) == 0:
                st.error(f"❌ Sin registros para {cuads}. Verifica el código UTM.")
            else:
                nombres = af.df["nombre_cientifico"].tolist() if "nombre_cientifico" in af.df.columns else []
                grupo_map = dict(zip(af.df.get("nombre_cientifico",[]), af.df.get("grupo",[]))) if "grupo" in af.df.columns else {}
                alertas = detectar_protegidas(nombres, ccaa_sel)
                for a in alertas:
                    a["grupo"] = grupo_map.get(a["especie"], "—")
                af.calcular_estadisticas(alertas)
                stats = af.estadisticas

                c1,c2,c3,c4 = st.columns(4)
                c1.metric("Total especies", stats.get("n_especies",0))
                c2.metric("LESRPE/CEEA", len([a for a in alertas if a["cat_lesrpe"]!="—"]))
                c3.metric(f"CCAA {ccaa_sel}", len([a for a in alertas if a["cat_auto_cod"]!="—"]))
                c4.metric("Total protegidas", len(alertas))

                if stats.get("por_grupo"):
                    st.markdown("**📊 Por grupo taxonómico:**")
                    df_g = pd.DataFrame(list(stats["por_grupo"].items()), columns=["Grupo","Nº especies"]).sort_values("Nº especies",ascending=False)
                    st.bar_chart(df_g.set_index("Grupo"))

                if alertas:
                    st.markdown("**🚨 Especies con estatus de protección:**")
                    st.dataframe(pd.DataFrame([{
                        "Especie":a["especie"],"Nombre común":a["nombre_comun"],
                        "Grupo":a["grupo"],"LESRPE/CEEA":a["cat_lesrpe"],
                        f"{ccaa_sel}":f"{a['emoji_auto']} {a['cat_auto_nombre']}",
                        "Normativa autonómica":a["decreto_auto"]
                    } for a in alertas]), use_container_width=True)
                    for cod,cls,titulo in [("EP","alerta-EP","🟠 En peligro de extinción"),
                                           ("V","alerta-V","🟡 Vulnerable"),
                                           ("OTRO","alerta-OTRO","🟢 Otro")]:
                        spp = [a for a in alertas if a["color_cod"]==cod]
                        if spp:
                            st.markdown(f"**{titulo} ({len(spp)} spp.):**")
                            for a in spp:
                                st.markdown(f'<div class="{cls}"><i>{a["especie"]}</i> — {a["nombre_comun"] or "—"} | {ccaa_sel}: {a["cat_auto_nombre"]} | LESRPE: {a["cat_lesrpe"]}</div>', unsafe_allow_html=True)
                else:
                    st.success("✅ Sin especies protegidas detectadas.")

                with st.expander(f"📂 Listado completo ({stats.get('n_especies',0)} especies)"):
                    cols_ok = [c for c in ["nombre_cientifico","grupo"] if c in af.df.columns]
                    st.dataframe(af.df[cols_ok], use_container_width=True)

                with st.spinner("Generando redacción técnica con phi-4..."):
                    redaccion = af.generar_redaccion(proyecto, alertas)
                cols_show = [c for c in ["nombre_cientifico","grupo"] if c in af.df.columns]
                informe_md = generar_informe_md(
                    f"Inventario Faunístico — {proyecto}", cuads,
                    f"IEET/BDIEET (MITECO)", stats, alertas,
                    af.df[cols_show].head(100), redaccion, ccaa_sel
                )
                st.success("✅ Informe generado")
                d1,d2,d3 = st.columns(3)
                with d1:
                    st.download_button("⬇️ .md", data=informe_md, file_name="informe_fauna.md", mime="text/markdown")
                with d2:
                    st.download_button("📄 .docx",
                        data=markdown_a_docx_bytes(informe_md,{"nombre":proyecto,"ubicacion":ubicacion,"ccaa":ccaa_sel,"tipo":tipo,"superficie":"—"}),
                        file_name="informe_fauna.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with d3:
                    excel_bytes = generar_excel(
                        proyecto=proyecto or "Proyecto",
                        ubicacion=ubicacion or "—",
                        ccaa=ccaa_sel,
                        tipo=tipo or "—",
                        cuadriculas=cuads,
                        alertas_fauna=alertas,
                        alertas_flora=[],
                        listado_fauna=af.df[cols_show] if cols_show else af.df,
                        listado_flora=[],
                        stats_fauna=stats,
                        decreto_auto=get_normativa(ccaa_sel).get("decreto","—"),
                    )
                    st.download_button("📊 .xlsx",
                        data=excel_bytes,
                        file_name="informe_fauna.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("👆 Sube el BDIEET para comenzar.\n\n📥 [Descarga en MITECO](https://www.miteco.gob.es/es/biodiversidad/temas/inventarios-nacionales/inventario-especies-terrestres/inventario-nacional-de-biodiversidad/bdn-ieet-default.html)")


# ══════════════════════════════════════════════════════════════════════════════
# TAB FLORA
# ══════════════════════════════════════════════════════════════════════════════
with tab_flora:
    st.markdown(f"## 🌿 Inventario Florístico — {ccaa_sel}")
    if tiene_normativa(ccaa_sel):
        n = get_normativa(ccaa_sel)
        st.markdown(f"<div class='ok-box'>✅ Normativa: {n.get('decreto','')}</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='warn-box'>➕ Añade normativa autonómica en la pestaña ⚙️</div>", unsafe_allow_html=True)

    col1,col2 = st.columns([2,1])
    with col1:
        flora_file = st.file_uploader("📂 Flora (Anthos, CSV o Excel)", type=["csv","txt","xlsx","xls"], key="flora")
    with col2:
        cuad_flora = st.text_input("📍 Cuadrícula UTM 10x10", "", key="cuad_fl")
        st.caption("🌿 [Descarga en anthos.es](http://www.anthos.es)")

    if flora_file:
        suf = Path(flora_file.name).suffix or ".csv"
        with tempfile.NamedTemporaryFile(suffix=suf, delete=False) as t:
            t.write(flora_file.read())
            st.session_state["ruta_flora"] = t.name
        fmt = "Anthos" if _es_anthos(Path(st.session_state["ruta_flora"])) else "CSV/Excel"
        st.success(f"✅ {flora_file.name} — {fmt}")

    if "ruta_flora" in st.session_state and st.button("🔬 Analizar flora", type="primary"):
        afl = AnalizadorFlora()
        afl.ccaa = ccaa_sel
        ruta_fl = Path(st.session_state["ruta_flora"])
        with st.spinner("Cargando flora..."):
            if _es_anthos(ruta_fl):
                afl.cargar_anthos(ruta_fl, cuad_flora)
            else:
                afl.cargar_csv_flora(ruta_fl, cuad_flora)

        if not afl.taxones:
            st.error("❌ No se encontraron taxones.")
        else:
            alertas_fl = detectar_protegidas(afl.taxones, ccaa_sel)
            c1,c2,c3 = st.columns(3)
            c1.metric("Total taxones", len(afl.taxones))
            c2.metric("LESRPE/CEEA", len([a for a in alertas_fl if a["cat_lesrpe"]!="—"]))
            c3.metric(f"CCAA {ccaa_sel}", len([a for a in alertas_fl if a["cat_auto_cod"]!="—"]))

            if alertas_fl:
                st.markdown("**🌿 Taxones protegidos:**")
                st.dataframe(pd.DataFrame([{
                    "Taxón":a["especie"],"Nombre común":a["nombre_comun"],
                    "LESRPE/CEEA":a["cat_lesrpe"],
                    f"{ccaa_sel}":f"{a['emoji_auto']} {a['cat_auto_nombre']}",
                    "Normativa":a["decreto_auto"]
                } for a in alertas_fl]), use_container_width=True)
                for cod,cls,titulo in [("EP","alerta-EP","🟠 En peligro de extinción"),
                                       ("V","alerta-V","🟡 Vulnerable"),
                                       ("OTRO","alerta-OTRO","🟢 Otro")]:
                    spp = [a for a in alertas_fl if a["color_cod"]==cod]
                    if spp:
                        st.markdown(f"**{titulo} ({len(spp)} spp.):**")
                        for a in spp:
                            st.markdown(f'<div class="{cls}"><i>{a["especie"]}</i> — {a["nombre_comun"] or "—"} | {ccaa_sel}: {a["cat_auto_nombre"]} | LESRPE: {a["cat_lesrpe"]}</div>', unsafe_allow_html=True)
            else:
                st.success("✅ Sin taxones protegidos detectados.")

            with st.expander(f"📂 Listado completo ({len(afl.taxones)} taxones)"):
                st.dataframe(pd.DataFrame({"Taxón":afl.taxones}), use_container_width=True)

            with st.spinner("Generando redacción técnica con phi-4..."):
                redaccion_fl = afl.generar_redaccion(proyecto, alertas_fl)
            informe_md_fl = generar_informe_md(
                f"Inventario Florístico — {proyecto}", [cuad_flora],
                afl.fuente, {"n_especies":len(afl.taxones)},
                alertas_fl, pd.DataFrame({"Taxón":afl.taxones[:100]}),
                redaccion_fl, ccaa_sel
            )
            st.success("✅ Informe generado")
            d1,d2,d3 = st.columns(3)
            with d1:
                st.download_button("⬇️ .md", data=informe_md_fl, file_name="informe_flora.md", mime="text/markdown")
            with d2:
                st.download_button("📄 .docx",
                    data=markdown_a_docx_bytes(informe_md_fl,{"nombre":proyecto,"ubicacion":ubicacion,"ccaa":ccaa_sel,"tipo":tipo,"superficie":"—"}),
                    file_name="informe_flora.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with d3:
                excel_bytes_fl = generar_excel(
                    proyecto=proyecto or "Proyecto",
                    ubicacion=ubicacion or "—",
                    ccaa=ccaa_sel,
                    tipo=tipo or "—",
                    cuadriculas=[cuad_flora],
                    alertas_fauna=[],
                    alertas_flora=alertas_fl,
                    listado_fauna=pd.DataFrame(),
                    listado_flora=afl.taxones,
                    stats_fauna={"por_grupo":{"Flora":len(afl.taxones)}},
                    decreto_auto=get_normativa(ccaa_sel).get("decreto","—"),
                )
                st.download_button("📊 .xlsx",
                    data=excel_bytes_fl,
                    file_name="informe_flora.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif "ruta_flora" not in st.session_state:
        st.info("👆 Sube el listado de flora para comenzar.\n\n📥 [Descarga en anthos.es](http://www.anthos.es)")


# ══════════════════════════════════════════════════════════════════════════════
# TAB NORMATIVA
# ══════════════════════════════════════════════════════════════════════════════
with tab_norm:
    st.markdown("## ⚙️ Gestión de Normativa Autonómica")
    st.markdown(
        "<div class='info-box'>Añade el catálogo de especies amenazadas de cualquier CCAA. "
        "Se aplica automáticamente en los módulos de Fauna y Flora.<br>"
        "<b>Formatos:</b> Excel (.xlsx/.xls), CSV (.csv), Word (.docx), PDF (.pdf) — puedes subir varios archivos a la vez.</div>",
        unsafe_allow_html=True
    )

    col_sel, col_up = st.columns([1, 2])
    with col_sel:
        ccaa_norm = st.selectbox("Comunidad autónoma", ["Selecciona CCAA..."] + CCAA_LISTA, key="ccaa_norm")
        if ccaa_norm == "Selecciona CCAA...":
            st.markdown("<div class='info-box'>👈 Elige una CCAA para ver o cargar su normativa.</div>", unsafe_allow_html=True)
        elif tiene_normativa(ccaa_norm):
            n = get_normativa(ccaa_norm)
            st.markdown(f"<div class='ok-box'>✅ Cargada: {len(n.get('especies',{}))} especies</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='warn-box'>⚠️ Sin normativa todavía</div>", unsafe_allow_html=True)
    with col_up:
        norm_files = st.file_uploader(
            "Sube uno o varios catálogos (Excel, CSV, Word o PDF)",
            type=["xlsx","xls","csv","pdf","docx"],
            accept_multiple_files=True,
            key="norm_up",
            disabled=(ccaa_norm == "Selecciona CCAA...")
        )

    if norm_files and ccaa_norm == "Selecciona CCAA...":
        st.warning("⚠️ Selecciona primero una comunidad autónoma.")
    elif norm_files:
        st.info(f"📂 {len(norm_files)} archivo(s) cargado(s): {', '.join(f.name for f in norm_files)}")
        if st.button("⚡ Procesar y añadir a la sesión", type="primary"):
            todas_especies = {}
            decretos = []
            for norm_file in norm_files:
                archivo_bytes = norm_file.read()
                ext = norm_file.name.lower().rsplit(".", 1)[-1] if "." in norm_file.name else ""
                if ext == "pdf":
                    with st.spinner(f"Leyendo {norm_file.name} (respetando el orden de columnas)..."):
                        try:
                            texto = extraer_texto_pdf(archivo_bytes)
                        except Exception as e:
                            texto = ""
                            st.warning(f"Error leyendo {norm_file.name}: {e}")
                        data = procesar_pdf_normativa(texto, ccaa_norm) if texto else {}
                elif ext == "docx":
                    with st.spinner(f"Procesando {norm_file.name}..."):
                        try:
                            data = procesar_docx_normativa(archivo_bytes, ccaa_norm, norm_file.name)
                        except Exception as e:
                            data = {}
                            st.warning(f"Error leyendo {norm_file.name}: {e}")
                else:
                    with st.spinner(f"Procesando {norm_file.name}..."):
                        data = procesar_excel_normativa(archivo_bytes, ccaa_norm, norm_file.name)
                if data.get("especies"):
                    todas_especies.update(data["especies"])
                    decretos.append(data.get("decreto", norm_file.name))

            if todas_especies:
                decreto_final = " | ".join(decretos)
                from normativa_autonomica import CATEGORIAS_DEFAULT
                st.session_state["normativas"][ccaa_norm] = {
                    "ccaa": ccaa_norm,
                    "decreto": decreto_final,
                    "categorias": CATEGORIAS_DEFAULT,
                    "especies": todas_especies
                }
                n_esp = len(todas_especies)
                st.success(f"✅ **{ccaa_norm}** — {n_esp} especies añadidas")
                muestra = list(todas_especies.items())[:8]
                st.dataframe(pd.DataFrame([{"Especie":k,"Cat.":v.get("cat",""),"Grupo":v.get("grupo","")} for k,v in muestra]), use_container_width=True)
                if n_esp > 8:
                    st.caption(f"... y {n_esp-8} especies más")
                st.rerun()
            else:
                st.error("❌ No se extrajeron especies. Revisa el formato de los archivos.")

    st.divider()
    st.markdown("### CCAA con normativa en sesión")
    normativas_sess = st.session_state["normativas"]
    if normativas_sess:
        cols_n = st.columns(min(4, len(normativas_sess)))
        for i,(ccaa_n,n) in enumerate(normativas_sess.items()):
            with cols_n[i % 4]:
                st.markdown(
                    f"<div class='ok-box'><b>{ccaa_n}</b><br>{len(n.get('especies',{}))} especies</div>",
                    unsafe_allow_html=True
                )
                if st.button(f"🗑️ Quitar", key=f"del_{ccaa_n}"):
                    del st.session_state["normativas"][ccaa_n]
                    st.rerun()
    else:
        st.info("No hay normativas cargadas todavía. Sube el catálogo de tu CCAA arriba.")

    st.divider()
    st.markdown("### Catálogos autonómicos")
    for ccaa_r in CCAA_LISTA:
        icon = "✅" if tiene_normativa(ccaa_r) else "📋"
        st.markdown(f"**{icon} {ccaa_r}**")

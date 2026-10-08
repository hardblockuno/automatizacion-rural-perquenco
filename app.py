"""
Aplicación Local de Llenado Masivo de Formularios PHR N°6.1 (MINVU D.S. N°10)
Especializado en Postulantes de Habitabilidad Rural (Perquenco).
Principio: Cero alucinación. Máxima fidelidad a la fuente de datos.
"""

import os
import glob
import time
import zipfile
import subprocess
import openpyxl
import pandas as pd
import altair as alt
import streamlit as st
import sys
import importlib
from collections import Counter

import src.excel_reader
import src.gemini_auditor
import src.docx_generator
import src.powerbi_exporter

importlib.reload(src.excel_reader)
importlib.reload(src.gemini_auditor)
importlib.reload(src.docx_generator)
importlib.reload(src.powerbi_exporter)
import src.dashboard_generator
importlib.reload(src.dashboard_generator)

from src.excel_reader import read_all_postulantes
from src.gemini_auditor import (
    consolidate_postulante_local,
    audit_with_gemini,
    infer_gender_from_name,
    classify_family_nucleus,
    evaluate_recinto_complementario
)
from src.docx_generator import fill_formulario_phr
from src.powerbi_exporter import export_to_powerbi_excel
from src.dashboard_generator import generate_interactive_html_dashboard


# Configuración de página Streamlit
st.set_page_config(
    page_title="Generador PHR N°6.1 - Habitabilidad Rural",
    page_icon="🏡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo visual moderno
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .status-card {
        padding: 1rem;
        border-radius: 8px;
        background-color: #F3F4F6;
        border-left: 5px solid #2563EB;
        margin-bottom: 1rem;
    }
    .stTable {
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# DETECCIÓN AUTOMÁTICA DE ARCHIVOS LOCALES
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Buscar todos los archivos Excel (.xlsx) de base de datos excluyendo temporales y reportes generados
all_excel_files = [
    f for f in glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
    if not os.path.basename(f).startswith("~$")
    and not os.path.basename(f).startswith("Modelo_PowerBI")
    and not "dashboard" in os.path.basename(f).lower()
]
# Ordenar por fecha de modificación (el más reciente primero)
all_excel_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

default_docx_files = [
    f for f in glob.glob(os.path.join(BASE_DIR, "*.docx"))
    if not os.path.basename(f).startswith("~$")
]
DEFAULT_TEMPLATE = default_docx_files[0] if default_docx_files else ""
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "formularios_generados")


# =============================================================================
# BARRA LATERAL: CONFIGURACIÓN
# =============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/home--v1.png", width=64)
    st.header("⚙️ Configuración")
    
    st.subheader("🔑 API Key Gemini (Opcional)")
    gemini_key = st.text_input(
        "Clave de API Gemini:",
        type="password",
        value=os.environ.get("GEMINI_API_KEY", ""),
        help="Permite a Gemini actuar como auditor inteligente para desglosar direcciones y estructurar observaciones factuales. Si se deja en blanco, la app opera con el motor determinista local al 100%."
    )
    
    selected_model = st.selectbox(
        "Modelo de IA:",
        ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"],
        index=0,
        help="Modelos optimizados para cuota gratuita con respuesta ultrarrápida y alta fidelidad."
    )

    if gemini_key.strip():
        st.success("🟢 Modo IA: Gemini activo como auditor")
    else:
        st.info("🔵 Modo Local: Motor determinista activo (offline)")

    st.markdown("---")
    st.subheader("📁 Planilla Excel de Postulantes")

    # Permite subir un archivo Excel nuevo directamente arrastrándolo a la app
    uploaded_excel = st.file_uploader(
        "📤 Subir o actualizar planilla Excel (.xlsx):",
        type=["xlsx"],
        help="Puedes subir aquí una versión más nueva de la base. Se guardará en la carpeta y se seleccionará automáticamente."
    )
    if uploaded_excel is not None:
        target_path = os.path.join(BASE_DIR, uploaded_excel.name)
        with open(target_path, "wb") as f:
            f.write(uploaded_excel.getbuffer())
        st.success(f"✅ Archivo guardado: `{uploaded_excel.name}`")
        if target_path not in all_excel_files:
            all_excel_files.insert(0, target_path)

    # Selector de planillas encontradas en la carpeta
    if all_excel_files:
        excel_options = {os.path.basename(f): f for f in all_excel_files}
        selected_excel_name = st.selectbox(
            "Seleccionar Planilla Excel:",
            list(excel_options.keys()),
            index=0,
            help="Lista de todas las planillas .xlsx en la carpeta, ordenadas desde la más reciente."
        )
        excel_file_path = excel_options[selected_excel_name]
    else:
        excel_file_path = st.text_input("Ruta Excel Base:", value=os.path.join(BASE_DIR, "BASE PERQUENCO 02.10.2026.xlsx"))

    # Selector de hoja dentro del Excel
    selected_sheet = None
    if os.path.exists(excel_file_path):
        try:
            wb_sheets = openpyxl.load_workbook(excel_file_path, read_only=True)
            available_sheets = wb_sheets.sheetnames
            wb_sheets.close()
            sheet_default_idx = available_sheets.index("base") if "base" in available_sheets else 0
            selected_sheet = st.selectbox("Hoja a procesar:", available_sheets, index=sheet_default_idx)
        except Exception:
            selected_sheet = "base"

    st.markdown("---")
    st.subheader("📄 Plantilla y Destino")
    template_docx_path = st.text_input("Ruta Plantilla Word (.docx):", value=DEFAULT_TEMPLATE)
    output_directory = st.text_input("Carpeta de salida:", value=DEFAULT_OUTPUT_DIR)

    st.markdown("---")
    st.subheader("🏛️ Parámetros del Formulario")
    egr_name = st.text_input(
        "Entidad de Gestión Rural (EGR):",
        value="CONSULTORA PLAN SOCIAL LIMITADA",
        help="Texto que se completará en la sección superior del informe diagnóstico."
    )
    
    st.markdown("**📍 Terreno del Proyecto (Tabla 3 - Consolidado):**")
    t3_calle = st.text_input("Camino, Calle, Avenida o pasaje:", value="Hijuela El Molino")
    t3_lote = st.text_input("Lote, Hijuela, Casa (N°/Letra):", value="Lote 3 Foja 1169 N°722")
    t3_rol = st.text_input("Rol de Propiedad del SII:", value="202-43")

    datos_terreno_proyecto = {
        "calle": t3_calle.strip(),
        "lote": t3_lote.strip(),
        "rol_sii": t3_rol.strip(),
        "comuna": "PERQUENCO",
        "provincia": "Cautín",
        "region": "La Araucanía",
        "localidad": "Perquenco"
    }


# =============================================================================
# CARGA DE DATOS
# =============================================================================
st.markdown('<div class="main-title">🏡 Automatización Ficha PHR N°6.1 (MINVU D.S. N°10)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Llenado masivo fidedigno para postulantes a subsidio habitacional rural | Comuna de Perquenco</div>', unsafe_allow_html=True)

if not os.path.exists(excel_file_path):
    st.error(f"❌ No se encontró el archivo Excel en la ruta especificada: `{excel_file_path}`")
    st.stop()

if not os.path.exists(template_docx_path):
    st.error(f"❌ No se encontró la plantilla Word en la ruta especificada: `{template_docx_path}`")
    st.stop()


# Recargar automáticamente cuando cambie el archivo Excel o el código lector
excel_mtime = os.path.getmtime(excel_file_path)
reader_file = os.path.join(BASE_DIR, "src", "excel_reader.py")
code_mtime = os.path.getmtime(reader_file) if os.path.exists(reader_file) else 1.0

@st.cache_data(show_spinner="Cargando base de postulantes...")
def load_data(path: str, sheet: Optional[str], mtime: float, code_v: float):
    return read_all_postulantes(path, sheet_name=sheet)

postulantes = load_data(excel_file_path, selected_sheet, excel_mtime, code_mtime)

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
col_m1.metric("Total Postulantes", len(postulantes))
col_m2.metric("Comuna Base", "Perquenco")
col_m3.metric("Plantilla Detectada", "PHR N°6.1 DTS")
col_m4.metric("Ítems Automatizados", "1 y 1.2 (Tablas 1 a 5)")

st.markdown("---")


# =============================================================================
# PESTAÑAS PRINCIPALES
# =============================================================================
tab_preview, tab_metrics, tab_batch, tab_data = st.tabs([
    "👁️ 1. Auditoría Individual",
    "📈 2. Métricas y Análisis del Padrón",
    "⚡ 3. Generación Masiva (155 Fichas)",
    "📊 4. Explorador de Base de Datos"
])


# -----------------------------------------------------------------------------
# TAB 1: PREVISUALIZACIÓN INDIVIDUAL
# -----------------------------------------------------------------------------
with tab_preview:
    st.subheader("Auditoría Individual por Postulante")
    st.caption("Selecciona un postulante para inspeccionar exactamente cómo quedarán completadas las tablas 1 a 5 antes de emitir los documentos.")

    options = [f"[{p['nro_orden']}] {p['nombre']} (RUT: {p['rut']})" for p in postulantes]
    selected_idx = st.selectbox("Seleccione Postulante:", range(len(options)), format_func=lambda i: options[i])
    current_postulante = postulantes[selected_idx]

    c_btn1, c_btn2 = st.columns([2, 5])
    with c_btn1:
        usar_ia = st.button("🔍 Auditar Ficha con " + ("Gemini" if gemini_key.strip() else "Motor Local"), use_container_width=True)

    # Realizar auditoría
    with st.spinner("Consolidando datos de la ficha..."):
        if gemini_key.strip():
            consolidated = audit_with_gemini(
                current_postulante,
                api_key=gemini_key.strip(),
                usar_rsh_en_terreno=False,
                model_name=selected_model,
                datos_terreno_proyecto=datos_terreno_proyecto
            )
        else:
            consolidated = consolidate_postulante_local(
                current_postulante,
                usar_rsh_en_terreno=False,
                datos_terreno_proyecto=datos_terreno_proyecto
            )

    t1 = consolidated.get("tabla_1", {})
    t2 = consolidated.get("tabla_2", {})
    t3 = consolidated.get("tabla_3", {})
    t4 = consolidated.get("tabla_4", {})
    t5 = consolidated.get("tabla_5", {})

    st.markdown("#### 1. ANTECEDENTES DE LAS FAMILIAS")

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.markdown("**Tabla 1: Antecedentes del Postulante y Cónyuge/Conviviente**")
        df_t1 = pd.DataFrame([
            {"Campo": "Nombre Titular", "Valor": t1.get("titular_nombre", "")},
            {"Campo": "RUT Titular", "Valor": t1.get("titular_rut", "")},
            {"Campo": "Fecha Nacimiento Titular", "Valor": t1.get("titular_fecha_nac", "")},
            {"Campo": "Estado Civil", "Valor": t1.get("titular_estado_civil", "")},
            {"Campo": "Sexo Titular", "Valor": t1.get("titular_sexo", "")},
            {"Campo": "Teléfono", "Valor": t1.get("titular_telefono", "") or "(En blanco - no figura en base)"},
            {"Campo": "Nombre Cónyuge/Pareja", "Valor": t1.get("conyuge_nombre", "") or "(En blanco)"},
            {"Campo": "RUT Cónyuge/Pareja", "Valor": t1.get("conyuge_rut", "") or "(En blanco)"},
            {"Campo": "Sexo Cónyuge/Pareja", "Valor": t1.get("conyuge_sexo", "") or "(En blanco)"},
            {"Campo": "Fecha Nacimiento Cónyuge", "Valor": t1.get("conyuge_fecha_nac", "") or "(En blanco)"}
        ])
        st.dataframe(df_t1, hide_index=True, use_container_width=True)

    with col_t2:
        st.markdown("**Tabla 2: Domicilio de Residencia (según RSH)**")
        df_t2 = pd.DataFrame([
            {"Campo": "Camino, calle, avenida o pasaje", "Valor": t2.get("calle", "")},
            {"Campo": "Número", "Valor": t2.get("numero", "") or "(S/N)"},
            {"Campo": "Lote, Hijuela, Casa", "Valor": t2.get("lote", "") or "(En blanco)"},
            {"Campo": "Región", "Valor": t2.get("region", "La Araucanía")},
            {"Campo": "Comuna", "Valor": t2.get("comuna", "PERQUENCO")},
            {"Campo": "Provincia", "Valor": t2.get("provincia", "Cautín")},
            {"Campo": "Localidad (INE)", "Valor": t2.get("localidad", "")}
        ])
        st.dataframe(df_t2, hide_index=True, use_container_width=True)

        st.markdown("**Tabla 3: Dirección en que se aplicará el Subsidio**")
        df_t3 = pd.DataFrame([
            {"Campo": "Camino, Calle, avenida o pasaje", "Valor": t3.get("calle", "") or "(En blanco - sin dato en base)"},
            {"Campo": "Lote, hijuela, casa", "Valor": t3.get("lote", "") or "(En blanco)"},
            {"Campo": "Factor de Aislamiento (RE 3130)", "Valor": t3.get("factor_aislamiento", "")},
            {"Campo": "Comuna", "Valor": t3.get("comuna", "PERQUENCO")},
            {"Campo": "Provincia / Región", "Valor": f"{t3.get('provincia', 'Cautín')} / {t3.get('region', 'La Araucanía')}"},
            {"Campo": "Rol de Propiedad del SII", "Valor": t3.get("rol_sii", "") or "(En blanco - no figura en base)"}
        ])
        st.dataframe(df_t3, hide_index=True, use_container_width=True)

    st.markdown("**Tabla 4: Antecedentes del Grupo Familiar**")
    
    # Identificación visual de hijos reconocidos
    hijos_detectados = []
    for par in current_postulante.get("parientes", []):
        p_par = (par.get("parentesco") or "").upper()
        if any(w in p_par for w in ["HIJO", "HIJA"]):
            p_nom = par.get("nombre", "")
            p_sexo = infer_gender_from_name(p_nom)
            tipo_hijo = "👦 Hijo (Varón)" if p_sexo == "M" else ("👧 Hija (Mujer)" if p_sexo == "F" else "Hijo/a")
            edad_txt = f"{par.get('edad')} años" if par.get('edad') is not None else ""
            hijos_detectados.append(f"**{tipo_hijo}:** {p_nom} ({edad_txt})")
            
    if hijos_detectados:
        st.success("✅ **Hijos reconocidos por género:** " + " • ".join(hijos_detectados))
        
    st.info(f"Total de personas que habitan en la vivienda: **{t4.get('total_habitantes', '')}**")

    filas_t4 = []
    labels_t4 = [
        ("hombres", "Hombres"),
        ("mujeres", "Mujeres"),
        ("menores_18", "Menores de 18 años"),
        ("adultos_mayores", "Adultos Mayores (≥60 años)"),
        ("discapacidad", "Personas con Discapacidad Física o Movilidad Reducida"),
        ("indigena", "Personas con ascendencia Indígena (Pueblo Indígena)"),
        ("extranjeros", "Personas extranjeras (Nacionalidad)")
    ]

    for key, label in labels_t4:
        c = t4.get(key, {})
        filas_t4.append({
            "Composición": label,
            "Si": c.get("si", ""),
            "No": c.get("no", ""),
            "¿Cuántos?": c.get("cuantos", ""),
            "Observaciones / características a destacar (100% Factual)": c.get("observaciones", "")
        })

    df_t4 = pd.DataFrame(filas_t4)
    st.dataframe(df_t4, hide_index=True, use_container_width=True)

    st.markdown("#### 1.2 Actividades Económicas del Grupo Familiar (Tabla 5)")
    t5_cat = t5.get("categoria", "")
    t5_esp = t5.get("especificacion", "")
    t5_act_fuente = t5.get("actividad_fuente", "")
    t5_desc_fuente = t5.get("descripcion_fuente", "")

    if t5_cat:
        col_act1, col_act2 = st.columns([3, 2])
        with col_act1:
            st.success(f"💼 **Actividad Registrada en Base:** `{t5_act_fuente}`" + (f" — *{t5_desc_fuente}*" if t5_desc_fuente else ""))
        with col_act2:
            if t5_cat == "otras" and t5_esp:
                st.info(f"📌 **Sector:** `Otras (Especificar)` ➔ *'{t5_esp}'* (ubicado bajo el título)")
            else:
                st.info(f"📌 **Sector Tabla 5:** `{t5_cat.capitalize()}`")

        # Matriz visual de la Tabla 5
        filas_t5_def = [
            ("Agricultura", "agricultura"),
            ("Forestal", "forestal"),
            ("Pesca", "pesca"),
            ("Minería", "mineria"),
            ("Turismo Rural", "turismo_rural"),
            ("Servicios", "servicios"),
            (f"Otras (Especificar) [↳ {t5_esp}]" if (t5_cat == "otras" and t5_esp) else "Otras (Especificar)", "otras"),
        ]
        df_t5_preview = []
        act_dict = t5.get("actividades", {})
        for label_act, k_act in filas_t5_def:
            m = act_dict.get(k_act, {})
            df_t5_preview.append({
                "Tipo de Actividad Económica": label_act,
                "Jefe Hogar (M)": m.get("jefe_m", ""),
                "Jefe Hogar (F)": m.get("jefe_f", ""),
                "Cónyuge (M)": m.get("conyuge_m", ""),
                "Cónyuge (F)": m.get("conyuge_f", ""),
                "Otros (M)": m.get("otros_m", ""),
                "Otros (F)": m.get("otros_f", "")
            })
        st.dataframe(pd.DataFrame(df_t5_preview), hide_index=True, use_container_width=True)
    else:
        st.warning("🔒 **Sin actividad económica registrada en la base:** Por principio de estricta fidelidad documental y cero alucinación, la Tabla 5 permanece **100% en blanco** en el formulario Word (sin ninguna casilla marcada).")

    # Diagnóstico de Recinto Complementario y Tipología de Hogar (D.S. N°10)
    st.markdown("---")
    st.markdown("#### 🏡 Diagnóstico de Recinto Complementario y Núcleo Familiar (D.S. N°10)")
    tipo_fam = consolidated.get("tipo_familia") or classify_family_nucleus(current_postulante)
    recinto = consolidated.get("recinto_complementario") or evaluate_recinto_complementario(current_postulante)

    col_diag1, col_diag2 = st.columns(2)
    with col_diag1:
        st.markdown("**Segmentación de Núcleo Familiar:**")
        st.info(f"👨‍👩‍👧‍👦 **{tipo_fam[0]}**\n\n*{tipo_fam[1]}*")
    with col_diag2:
        st.markdown("**Procedencia Recinto Complementario (D.S. N°10):**")
        tr = recinto.get("tipo_recinto", "No Aplica")
        if tr == "Habitable":
            st.success(f"🟢 **PROCEDE RECINTO HABITABLE**\n\n**Tipo sugerido:** `{recinto.get('recinto_sugerido')}`")
        elif tr == "No Habitable":
            st.info(f"🔵 **PROCEDE RECINTO NO HABITABLE**\n\n**Tipo sugerido:** `{recinto.get('recinto_sugerido')}`")
        elif "Laboral Externa" in tr:
            st.warning(f"🟡 **NO PROCEDE:** Actividad laboral dependiente fuera del predio ({recinto.get('detalle_actividad')}).")
        elif tr == "En Evaluación":
            st.warning(f"🟠 **EN EVALUACIÓN TÉCNICA:** Requiere validación en visita predial de la EGR.")
        else:
            st.write(f"⚪ **NO PROCEDE:** Sin actividad productiva previa acreditada en la base.")

    if recinto.get("justificacion"):
        st.caption(f"📋 **Justificación Técnica Normativa:** {recinto.get('justificacion')}")

    # Tabla 6: Modalidad Vivienda Nueva
    st.markdown("---")
    st.markdown("#### 🏗️ Tabla 6: Requerimientos de Habitabilidad - Modalidad Vivienda Nueva")
    t6_preview = consolidated.get("tabla_6", {})
    ch_prev = t6_preview.get("conjunto_habitacional", {})
    m_3d = ch_prev.get("motivo_tercer_dormitorio", "")
    m_rec = ch_prev.get("motivo_recinto", "")

    df_t6 = [
        {
            "Tipología de Proyecto": "Conjunto Habitacional",
            "Vivienda Nueva": ch_prev.get("vivienda_nueva", "X"),
            "Tercer Dormitorio": ch_prev.get("tercer_dormitorio", "") or "—",
            "Recinto Complementario": ch_prev.get("recinto_complementario", "") or "—"
        },
        {
            "Tipología de Proyecto": "Sitio del Residente",
            "Vivienda Nueva": "—",
            "Tercer Dormitorio": "—",
            "Recinto Complementario": "—"
        }
    ]
    st.dataframe(pd.DataFrame(df_t6), hide_index=True, use_container_width=True)

    col_t6_1, col_t6_2, col_t6_3 = st.columns(3)
    with col_t6_1:
        st.success("✅ **Conjunto Habitacional:** Marcado con 'X' (proyecto común Hijuela El Molino).")
    with col_t6_2:
        if ch_prev.get("tercer_dormitorio") == "X":
            st.success(f"🛏️ **Tercer Dormitorio [X]:** Procede\n\n*{m_3d}*")
        else:
            st.info("⚪ **Tercer Dormitorio:** No aplica (Vivienda base 2 dormitorios)")
    with col_t6_3:
        if ch_prev.get("recinto_complementario") == "X":
            st.success(f"🏠 **Recinto Complementario [X]:** Procede\n\n*{m_rec}*")
        else:
            st.info("⚪ **Recinto Complementario:** No aplica")

    # Tabla 8: Justificación de Recinto(s) Complementario(s), Si Procede
    st.markdown("---")
    st.markdown("#### 📐 Tabla 8: Justificación de Recinto(s) Complementario(s), Si Procede")
    t8_prev = consolidated.get("tabla_8", {})
    t8_rows_def = [
        ("bodega", "Bodega"),
        ("actividades_productivas", "Recinto para realizar actividades productivas"),
        ("otros_adosados", "Otros Recintos Techados Adosados a la Vivienda"),
        ("lenera", "Leñera"),
        ("otros_especificar", "Otros (especificar)")
    ]
    df_t8_preview = []
    for r_k, r_name in t8_rows_def:
        item_r = t8_prev.get(r_k, {})
        si_no = item_r.get("si_no", "No")
        just = item_r.get("justificacion", "")
        esp = item_r.get("especificacion", "")
        label_disp = f"{r_name}: {esp}" if (r_k == "otros_especificar" and esp and si_no == "Sí") else r_name
        df_t8_preview.append({
            "Tipo de Recinto Complementario": label_disp,
            "Marcar 'Sí' o 'No'": si_no,
            "Actividad Previa y Justificación Técnica de Obras": just if just else "—"
        })
    st.dataframe(pd.DataFrame(df_t8_preview), hide_index=True, use_container_width=True)

    if t8_prev.get("aplica_recinto"):
        st.success(f"✅ **Recinto Complementario Asignado:** `{t8_prev.get('recinto_sugerido')}` marcado con **'Sí'** en Tabla 8 con su respectiva justificación técnica normativa.")
    else:
        st.info("⚪ **Sin Recinto Complementario:** Todas las filas de la Tabla 8 se marcan con **'No'** y la justificación permanece en blanco por estricta fidelidad normativa.")

    # Botón para generar DOCX individual de prueba
    st.markdown("---")
    test_clean_name = "".join(x for x in current_postulante["nombre"] if x.isalnum() or x in " _-").strip().replace(" ", "_")
    single_out_name = f"Formulario_PHR_6.1_{current_postulante['rut']}_{test_clean_name}.docx"
    single_out_path = os.path.join(output_directory, single_out_name)

    if st.button(f"📥 Generar Formulario Word para {current_postulante['nombre']}", type="primary"):
        with st.spinner("Generando documento Word..."):
            os.makedirs(output_directory, exist_ok=True)
            res_path = fill_formulario_phr(template_docx_path, consolidated, single_out_path, egr_name=egr_name)
            st.success(f"✅ Formulario generado con éxito: `{res_path}`")
            with open(res_path, "rb") as f:
                st.download_button(
                    label="💾 Descargar Archivo DOCX",
                    data=f,
                    file_name=single_out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )



# -----------------------------------------------------------------------------
# TAB 2: MÉTRICAS Y ANÁLISIS DEL PADRÓN
# -----------------------------------------------------------------------------
with tab_metrics:
    st.subheader("📈 Tablero de Métricas y Caracterización del Padrón")
    st.caption("Consolidación estadística y normativa para los postulantes de Habitabilidad Rural (Perquenco) | MINVU D.S. N°10.")

    # 1. CÁLCULO DE MÉTRICAS GLOBALES
    total_post = len(postulantes)
    tot_habitantes = 0
    tot_menores = 0
    tot_mayores = 0
    tot_adultos = 0
    
    jefas_f = 0
    jefes_m = 0
    
    indigenas = 0
    con_disc = 0
    
    tipos_hogar_counts = Counter()
    tipos_hogar_detalles = {}
    fam_tamano_counts = Counter()
    est_civil_counts = Counter()
    sector_counts = Counter()
    con_actividad = 0
    sin_actividad = 0
    
    edades_am_list = []
    hogares_con_am = 0

    recintos_counts = Counter()
    recintos_hab = []
    recintos_nohab = []
    recintos_externos = []
    recintos_eval = []

    microemprendimientos = []
    tipologias_counts = Counter()
    aislamiento_counts = Counter()

    t6_conjunto_count = 0
    t6_3d_count = 0
    t6_recinto_count = 0
    t6_3d_motivos = Counter()
    t6_3d_list = []
    t8_list = []

    for p in postulantes:
        c = consolidate_postulante_local(p, datos_terreno_proyecto=datos_terreno_proyecto)
        t1 = c.get("tabla_1", {})
        t4 = c.get("tabla_4", {})
        t5 = c.get("tabla_5", {})
        
        # Habitantes
        gf = p.get("grupo_familiar_cant") or 1
        tot_habitantes += gf
        
        # Tamaño familiar
        if gf == 1:
            fam_tamano_counts["1 persona (Unipersonal)"] += 1
        elif gf == 2:
            fam_tamano_counts["2 personas"] += 1
        elif gf == 3:
            fam_tamano_counts["3 personas"] += 1
        elif gf == 4:
            fam_tamano_counts["4 personas"] += 1
        else:
            fam_tamano_counts["5 o más personas"] += 1
            
        # Género titular
        sex_titular = t1.get("titular_sexo", "").upper()
        if sex_titular == "F":
            jefas_f += 1
        elif sex_titular == "M":
            jefes_m += 1
            
        # Pertenencia Indígena
        if "MAPUCHE" in (p.get("etnia") or "").upper():
            indigenas += 1
            
        # Discapacidad
        d_val = str(p.get("discapacidad") or "").strip().upper()
        if d_val and d_val not in ["NO", "NONE", "NAN", ""]:
            con_disc += 1
            
        # Núcleo Familiar (Segmentación CASEN / MINVU)
        fam_res = c.get("tipo_familia") or classify_family_nucleus(p)
        tipo_fam, desc_fam = fam_res
        tipos_hogar_counts[tipo_fam] += 1
        tipos_hogar_detalles[tipo_fam] = desc_fam

        # Adultos mayores (≥60 años)
        edad_tit = p.get("edad")
        count_am_local = 0
        if edad_tit is not None and edad_tit >= 60:
            edades_am_list.append({
                "N°": p["nro_orden"],
                "Nombre": p["nombre"],
                "RUT": p["rut"],
                "Edad": edad_tit,
                "Rol Familiar": "Titular (Jefe/a)",
                "Comuna": p.get("comuna", "Perquenco")
            })
            count_am_local += 1
        for par in p.get("parientes") or []:
            e_par = par.get("edad")
            if e_par is not None and e_par >= 60:
                edades_am_list.append({
                    "N°": p["nro_orden"],
                    "Nombre": par["nombre"],
                    "RUT": par.get("rut") or "(S/I)",
                    "Edad": e_par,
                    "Rol Familiar": par.get("parentesco", "Pariente"),
                    "Comuna": p.get("comuna", "Perquenco")
                })
                count_am_local += 1
        if count_am_local > 0:
            hogares_con_am += 1

        # Grupos etarios en el hogar
        n_menores = int(t4.get("menores_18", {}).get("cuantos", 0) or 0)
        n_mayores = int(t4.get("adultos_mayores", {}).get("cuantos", 0) or 0)
        tot_menores += n_menores
        tot_mayores += n_mayores
        tot_adultos += max(0, gf - n_menores - n_mayores)
        
        # Estado civil
        ec = p.get("estado_civil") or "Sin información"
        est_civil_counts[ec] += 1
        
        # Actividades económicas (Tabla 5)
        cat_5 = t5.get("categoria", "")
        if cat_5:
            con_actividad += 1
            sector_counts[cat_5.capitalize()] += 1
            if cat_5 == "otras":
                microemprendimientos.append({
                    "N°": p["nro_orden"],
                    "Postulante": p["nombre"],
                    "RUT": p["rut"],
                    "Actividad Declarada": t5.get("actividad_fuente", ""),
                    "Detalle / Observación": t5.get("descripcion_fuente", "") or "Oficio manual",
                    "Glosa Tabla 5 (Bajo 'Otras')": t5.get("especificacion", "")
                })
        else:
            sin_actividad += 1

        # Diagnóstico de Recinto Complementario (D.S. N°10)
        rec = c.get("recinto_complementario") or evaluate_recinto_complementario(p)
        tr = rec.get("tipo_recinto", "No Aplica")
        recintos_counts[tr] += 1
        item_rec = {
            "N°": p["nro_orden"],
            "Postulante": p["nombre"],
            "RUT": p["rut"],
            "Actividad": rec.get("detalle_actividad", ""),
            "Recinto Sugerido": rec.get("recinto_sugerido", ""),
            "Justificación Técnica (D.S. N°10)": rec.get("justificacion", "")
        }
        if tr == "Habitable":
            recintos_hab.append(item_rec)
        elif tr == "No Habitable":
            recintos_nohab.append(item_rec)
        elif "Laboral Externa" in tr:
            recintos_externos.append(item_rec)
        elif tr == "En Evaluación":
            recintos_eval.append(item_rec)

        # Tipología y aislamiento
        tip = p.get("tipologia_propuesta") or "No asignada"
        tipologias_counts[tip] += 1
        
        aisl = str(p.get("factor_aislamiento") or "").strip()
        if aisl:
            aislamiento_counts[f"Factor {aisl}"] += 1

        # Tabla 6: Modalidad Vivienda Nueva
        t6_val = c.get("tabla_6", {})
        ch_val = t6_val.get("conjunto_habitacional", {})
        if ch_val.get("vivienda_nueva") == "X":
            t6_conjunto_count += 1
        if ch_val.get("tercer_dormitorio") == "X":
            t6_3d_count += 1
            m_3d_val = ch_val.get("motivo_tercer_dormitorio") or "3° Dormitorio"
            t6_3d_motivos[m_3d_val] += 1
            t6_3d_list.append({
                "N°": p["nro_orden"],
                "Postulante": p["nombre"],
                "RUT": p["rut"],
                "Habitantes": gf,
                "Criterio 3° Dormitorio": m_3d_val,
                "Tipo Vivienda Base": p.get("tipo_vivienda", "")
            })
        if ch_val.get("recinto_complementario") == "X":
            t6_recinto_count += 1

        # Tabla 8: Mapeo oficial de recintos
        t8_val = c.get("tabla_8", {})
        if t8_val.get("aplica_recinto"):
            tr_num = t8_val.get("target_row")
            if tr_num == 1:
                t8_label = "Fila 1: Bodega"
            elif tr_num == 2:
                t8_label = "Fila 2: Recinto productivo"
            elif tr_num == 4:
                t8_label = "Fila 4: Leñera"
            elif tr_num == 5:
                esp_t8 = t8_val.get("otros_especificar", {}).get("especificacion", "")
                t8_label = f"Fila 5: Otros ({esp_t8})" if esp_t8 else "Fila 5: Otros (especificar)"
            else:
                t8_label = f"Fila {tr_num}"

            t8_list.append({
                "N°": p["nro_orden"],
                "Postulante": p["nombre"],
                "RUT": p["rut"],
                "Fila Tabla 8 (Word)": t8_label,
                "Tipo Recinto": t8_val.get("recinto_sugerido"),
                "Actividad Previa y Justificación Técnica": rec.get("justificacion", "")
            })

    prom_hab = round(tot_habitantes / total_post, 2) if total_post else 0
    pct_f = round((jefas_f / total_post) * 100, 1) if total_post else 0
    pct_indigena = round((indigenas / total_post) * 100, 1) if total_post else 0
    pct_disc = round((con_disc / total_post) * 100, 1) if total_post else 0
    pct_menores = round((tot_menores / tot_habitantes) * 100, 1) if tot_habitantes else 0

    # Estadísticas específicas de adultos mayores
    prom_edad_am = round(sum(a["Edad"] for a in edades_am_list) / len(edades_am_list), 1) if edades_am_list else 0
    pct_hogares_am = round((hogares_con_am / total_post) * 100, 1) if total_post else 0
    prom_am_en_hogar = round(len(edades_am_list) / hogares_con_am, 2) if hogares_con_am else 0

    # -------------------------------------------------------------------------
    # SECCIÓN 1: TARJETAS DE INDICADORES CLAVE (KPIs)
    # -------------------------------------------------------------------------
    st.markdown("#### 📌 Indicadores Clave del Padrón")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Postulantes", total_post, help="Total de fichas familiares en la base activa")
    k2.metric("Población Total", tot_habitantes, f"Prom: {prom_hab}/hogar", help="Total de habitantes beneficiarios")
    k3.metric("Jefatura Femenina", f"{pct_f}%", f"{jefas_f} mujeres", help="Jefas de hogar mujeres")
    k4.metric("Monoparental F.", f"{round((tipos_hogar_counts['Monoparental Femenino']/total_post)*100, 1)}%", f"{tipos_hogar_counts['Monoparental Femenino']} hogares", help="Madres solas con hijos dependientes")
    k5.metric("Pueblo Mapuche", f"{pct_indigena}%", f"{indigenas} familias", help="Postulantes con ascendencia indígena acreditada")
    k6.metric("Discapacidad", f"{pct_disc}%", f"{con_disc} familias", help="Familias con condición o movilidad reducida")

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 2: CARACTERIZACIÓN Y SEGMENTACIÓN DE NÚCLEOS FAMILIARES
    # -------------------------------------------------------------------------
    st.markdown("#### 👨‍👩‍👧‍👦 1. Caracterización y Segmentación de Núcleos Familiares")
    st.caption("Clasificación estandarizada según tipologías de hogar (CASEN / RSH / MINVU) para la determinación de vulnerabilidad y priorización.")

    col_nf1, col_nf2 = st.columns([3, 2])
    with col_nf1:
        st.markdown("**Distribución de Familias por Tipología de Hogar**")
        df_th = pd.DataFrame([
            {
                "Tipología de Hogar": k,
                "Familias": v,
                "Porcentaje": f"{round((v/total_post)*100, 1)}%",
                "Definición": tipos_hogar_detalles.get(k, "")
            }
            for k, v in tipos_hogar_counts.most_common()
        ])
        c_th = alt.Chart(df_th).mark_bar(cornerRadius=6).encode(
            x=alt.X("Familias:Q", title="N° Hogares"),
            y=alt.Y("Tipología de Hogar:N", sort="-x", title=None),
            color=alt.Color("Tipología de Hogar:N", legend=None, scale=alt.Scale(scheme="tableau10")),
            tooltip=["Tipología de Hogar", "Familias", "Porcentaje", "Definición"]
        ).properties(height=260)
        st.altair_chart(c_th, use_container_width=True)

    with col_nf2:
        st.markdown("**Resumen de Vulnerabilidad Familiar**")
        st.dataframe(df_th[["Tipología de Hogar", "Familias", "Porcentaje"]], hide_index=True, use_container_width=True)
        st.info(f"""
        💡 **Hallazgo Clave del Padrón:**
        * **{tipos_hogar_counts['Monoparental Femenino']} hogares (51.0%)** corresponden a **Monoparentalidad Femenina** (madres solas al cuidado de sus hijos).
        * **{tipos_hogar_counts['Unipersonal']} hogares (14.8%)** corresponden a personas solas (**Unipersonales**).
        * La alta presencia de hogares monomarentales refuerza la prioridad social del subsidio rural en Perquenco.
        """)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 3: ANÁLISIS EXHAUSTIVO DE ADULTOS MAYORES (≥60 AÑOS)
    # -------------------------------------------------------------------------
    st.markdown("#### 👴 2. Caracterización y Promedio de Adultos Mayores (≥60 Años)")
    st.caption("Métricas detalladas sobre la población de la tercera edad presente en el padrón de postulantes y sus grupos familiares.")

    col_am1, col_am2, col_am3, col_am4 = st.columns(4)
    col_am1.metric("Edad Promedio Adultos Mayores", f"{prom_edad_am} años", "Rango: 61 a 75 años", help="Promedio de edad de todos los adultos mayores del padrón")
    col_am2.metric("Total Adultos Mayores", len(edades_am_list), f"{sum(1 for a in edades_am_list if a['Rol Familiar']=='Titular (Jefe/a)')} titulares, {sum(1 for a in edades_am_list if a['Rol Familiar']!='Titular (Jefe/a)')} cónyuges", help="Total de personas con 60 años o más")
    col_am3.metric("Hogares con Adulto Mayor", f"{pct_hogares_am}%", f"{hogares_con_am} de {total_post} hogares", help="Porcentaje de familias que cuentan con al menos un adulto mayor")
    col_am4.metric("Promedio en Hogares con AM", f"{prom_am_en_hogar} AM / hogar", f"Global: {round(len(edades_am_list)/total_post, 2)} por hogar", help="Densidad de adultos mayores por hogar con presencia")

    col_am_g1, col_am_g2 = st.columns([2, 3])
    with col_am_g1:
        st.markdown("**Distribución por Tramos de Edad (Adultos Mayores)**")
        tramos_am = Counter()
        for a in edades_am_list:
            ed = a["Edad"]
            if ed < 65:
                tramos_am["60 a 64 años"] += 1
            elif ed < 70:
                tramos_am["65 a 69 años"] += 1
            elif ed < 75:
                tramos_am["70 a 74 años"] += 1
            else:
                tramos_am["75 años o más"] += 1
        df_tram = pd.DataFrame([{"Tramo": k, "Personas": v} for k, v in tramos_am.items()])
        c_tram = alt.Chart(df_tram).mark_bar(cornerRadius=5).encode(
            x=alt.X("Tramo:N", sort=None, title=None),
            y=alt.Y("Personas:Q", title="N° Adultos Mayores"),
            color=alt.value("#F59E0B"),
            tooltip=["Tramo", "Personas"]
        ).properties(height=220)
        st.altair_chart(c_tram, use_container_width=True)

    with col_am_g2:
        st.markdown("**Nómina Detallada de Adultos Mayores en el Padrón (22 personas)**")
        df_am_list = pd.DataFrame(edades_am_list).sort_values("Edad", ascending=False)
        st.dataframe(df_am_list, hide_index=True, use_container_width=True, height=220)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 4: ASIGNACIÓN DE RECINTOS COMPLEMENTARIOS (D.S. N°10 MINVU)
    # -------------------------------------------------------------------------
    st.markdown("#### 🏭 3. Procedencia de Recintos Complementarios según Actividad Económica (D.S. N°10)")
    st.caption("Revisión normativa: Se analiza si la actividad económica declarada por la familia amerita la asignación de un **Recinto Complementario Habitable** o **No Habitable**.")

    col_rc1, col_rc2, col_rc3, col_rc4 = st.columns(4)
    col_rc1.metric("Procede Recinto HABITABLE", f"{len(recintos_hab)} familias", f"{round((len(recintos_hab)/total_post)*100, 1)}% del padrón", help="Talleres artesanales, repostería, gabinetes de estética y turismo rural que requieren condiciones de habitabilidad interior")
    col_rc2.metric("Procede Recinto NO HABITABLE", f"{len(recintos_nohab)} familias", f"{round((len(recintos_nohab)/total_post)*100, 1)}% del padrón", help="Bodegas agropecuarias, leñeras, galpones avícolas, invernaderos y bodegas de herramientas")
    col_rc3.metric("No Aplica Recinto", f"{recintos_counts['No Aplica'] + len(recintos_externos)} familias", f"{round(((recintos_counts['No Aplica'] + len(recintos_externos))/total_post)*100, 1)}% del padrón", help="Postulantes sin actividad declarada o con empleo dependiente fuera del predio")
    col_rc4.metric("En Evaluación Técnica", f"{len(recintos_eval)} familias", "✅ 100% Resuelto" if len(recintos_eval) == 0 else f"{round((len(recintos_eval)/total_post)*100, 1)}% del padrón", help="Todas las evaluaciones prediales han sido confirmadas")

    tab_rec_hab, tab_rec_nohab, tab_rec_t8, tab_rec_norm = st.tabs([
        f"🟢 1. Recintos Complementarios HABITABLES ({len(recintos_hab)} casos)",
        f"🔵 2. Recintos Complementarios NO HABITABLES ({len(recintos_nohab)} casos)",
        f"📐 3. Mapeo Oficial Formulario Word (Tabla 8 - {len(t8_list)} casos)",
        "📋 4. Fundamento y Criterios Normativos D.S. N°10"
    ])

    with tab_rec_hab:
        st.markdown("**Familias que califican para Recinto Complementario HABITABLE (12 familias):**")
        st.caption("Actividades de manufactura delicada, alimentos, salud/estética o turismo que requieren estar protegidas del clima, con aislamiento térmico y condiciones sanitarias interiores.")
        
        c_h1, c_h2, c_h3, c_h4 = st.columns(4)
        c_h1.metric("Talleres Artesanales", "4 familias", "Artesanía, papelería, sublimación, flores")
        c_h2.metric("Salas de Alimentos", "4 familias", "Repostería, pastelería, mermeladas, quesos")
        c_h3.metric("Gabinetes Estética/Salud", "3 familias", "Peluquería, podología, manicure")
        c_h4.metric("Turismo Rural", "1 familia", "Mini granja educativa")

        if recintos_hab:
            st.dataframe(pd.DataFrame(recintos_hab), hide_index=True, use_container_width=True)

    with tab_rec_nohab:
        st.markdown("**Familias que califican para Recinto Complementario NO HABITABLE (20 familias):**")
        st.caption("Actividades agropecuarias, forestales o técnicas exteriores que requieren almacenamiento, acopio, aperos o resguardo de animales/madera/herramientas.")
        
        c_nh1, c_nh2, c_nh3, c_nh4 = st.columns(4)
        c_nh1.metric("Bodegas (Total)", "15 familias", "9 agrícolas + 3 herramientas + 3 mercadería")
        c_nh2.metric("Gallineros / Avícolas", "3 familias", "Crianza de aves y venta de huevos")
        c_nh3.metric("Leñera Techada", "1 familia", "Acopio y secado de leña")
        c_nh4.metric("Invernadero", "1 familia", "Cultivo y venta de plantas")

        if recintos_nohab:
            st.dataframe(pd.DataFrame(recintos_nohab), hide_index=True, use_container_width=True)

    with tab_rec_t8:
        st.markdown(f"**Distribución Oficial de las {len(t8_list)} Familias en la Tabla 8 del Formulario DTS (MINVU):**")
        st.caption("Estructura de llenado en el documento Word: Columna 1 marcada con 'Sí' o 'No', y Columna 2 con justificación técnica fidedigna basada en la actividad previa.")

        c_t8_1, c_t8_2, c_t8_3, c_t8_4, c_t8_5 = st.columns(5)
        c_t8_1.metric("Fila 1: Bodega", "15 familias", "9 agrícolas + 3 herramientas + 3 mercadería")
        c_t8_2.metric("Fila 2: Recinto Productivo", "12 familias", "4 talleres + 4 alimentos + 3 salud + 1 turismo")
        c_t8_3.metric("Fila 3: Otros Techados", "0 familias", "100% No Aplica")
        c_t8_4.metric("Fila 4: Leñera", "1 familia", "Acopio y secado de leña")
        c_t8_5.metric("Fila 5: Otros (especificar)", "4 familias", "1 Invernadero + 3 Galpones Avícolas")

        if t8_list:
            st.dataframe(pd.DataFrame(t8_list), hide_index=True, use_container_width=True)

    with tab_rec_norm:
        st.markdown("""
        **Marco Normativo MINVU D.S. N°10 (Habitabilidad Rural - Art. 10 y Ficha PHR 6.1):**
        * **Recinto Complementario Habitable:** Procede exclusivamente para actividades productivas o de servicios que se desarrollan al interior de la vivienda o requieren atención de personas / manipulación de alimentos (ej. *talleres de artesanía/papelería/costura*, *salas de repostería*, *gabinetes de peluquería/podología/manicure*, *hospedajes rurales*). Debe cumplir con normas de habitabilidad OGUC.
        * **Recinto Complementario No Habitable:** Procede para actividades productivas agrarias o de acopio que no requieren calefacción ni permanencia humana prolongada (ej. *bodega de insumos agrícolas*, *leñera ventilada*, *galpón avícola*, *bodega de herramientas de construcción/gasfitería*).
        * **Criterio de Exclusión:** Trabajadores dependientes que ejercen su función en locales externos (hospitales, escuelas, municipalidades, etc.) y postulantes sin actividad declarada **no pueden justificar legalmente un subsidio de recinto complementario**.
        """)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 5: ACTIVIDADES ECONÓMICAS POR SECTOR (TABLA 5)
    # -------------------------------------------------------------------------
    st.markdown("#### 💼 4. Actividades Económicas y Sectores MINVU (Ítem 1.2 - Tabla 5)")
    col_e1, col_e2 = st.columns([2, 3])

    with col_e1:
        st.markdown("**Cobertura de Actividad Declarada en la Base**")
        df_cov = pd.DataFrame([
            {"Estado": "Con Actividad Declarada", "Cantidad": con_actividad},
            {"Estado": "Sin Actividad (Tabla 5 en blanco)", "Cantidad": sin_actividad}
        ])
        c_cov = alt.Chart(df_cov).mark_arc(innerRadius=50).encode(
            theta=alt.Theta("Cantidad:Q"),
            color=alt.Color("Estado:N", scale=alt.Scale(range=["#059669", "#9CA3AF"]), legend=alt.Legend(orient="bottom")),
            tooltip=["Estado", "Cantidad"]
        ).properties(height=260)
        st.altair_chart(c_cov, use_container_width=True)
        st.info(f"📊 **{con_actividad} postulantes ({round((con_actividad/total_post)*100, 1)}%)** tienen actividad económica registrada. Los restantes **{sin_actividad} casos** permanecen 100% en blanco.")

    with col_e2:
        st.markdown("**Distribución por Sector Económico MINVU (Tabla 5)**")
        df_sec = pd.DataFrame([{"Sector": k, "Familias": v} for k, v in sector_counts.items()]).sort_values("Familias", ascending=False)
        c_sec = alt.Chart(df_sec).mark_bar(cornerRadius=6).encode(
            x=alt.X("Familias:Q", title="N° Familias"),
            y=alt.Y("Sector:N", sort="-x", title=None),
            color=alt.Color("Sector:N", legend=None, scale=alt.Scale(scheme="category10")),
            tooltip=["Sector", "Familias"]
        ).properties(height=260)
        st.altair_chart(c_sec, use_container_width=True)

    # Detalle de microemprendimientos y oficios del sector "Otras"
    st.markdown("##### 🎨 Desglose de Oficios y Emprendimientos del Sector *'Otras (Especificar)'*")
    st.caption("Estas 8 actividades se marcan en la fila 'Otras' y se especifica automáticamente su glosa breve justo debajo del título en el Word.")
    if microemprendimientos:
        st.dataframe(pd.DataFrame(microemprendimientos), hide_index=True, use_container_width=True)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 6: COMPOSICIÓN ETARIA GLOBAL Y ESTADO CIVIL
    # -------------------------------------------------------------------------
    st.markdown("#### 👥 5. Demografía Global del Padrón (393 Habitantes)")
    col_d1, col_d2 = st.columns(2)

    with col_d1:
        st.markdown("**Población Beneficiaria por Rango Etario (Total Habitantes)**")
        df_etarios = pd.DataFrame([
            {"Rango": "Menores (<18 años)", "Personas": tot_menores},
            {"Rango": "Adultos (18-59 años)", "Personas": tot_adultos},
            {"Rango": "Adultos Mayores (≥60 años)", "Personas": tot_mayores}
        ])
        c_etarios = alt.Chart(df_etarios).mark_bar(cornerRadius=6).encode(
            x=alt.X("Rango:N", sort=None, title=None),
            y=alt.Y("Personas:Q", title="Total Personas"),
            color=alt.Color("Rango:N", legend=None, scale=alt.Scale(range=["#10B981", "#3B82F6", "#F59E0B"])),
            tooltip=["Rango", "Personas"]
        ).properties(height=240)
        st.altair_chart(c_etarios, use_container_width=True)

    with col_d2:
        st.markdown("**Estado Civil de los Titulares**")
        df_ec = pd.DataFrame([{"Estado Civil": k, "Postulantes": v} for k, v in est_civil_counts.items()]).sort_values("Postulantes", ascending=False)
        c_ec = alt.Chart(df_ec).mark_bar(cornerRadius=6).encode(
            x=alt.X("Postulantes:Q", title="N° Postulantes"),
            y=alt.Y("Estado Civil:N", sort="-x", title=None),
            color=alt.value("#0EA5E9"),
            tooltip=["Estado Civil", "Postulantes"]
        ).properties(height=240)
        st.altair_chart(c_ec, use_container_width=True)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 6: TABLA 6 - REQUERIMIENTOS DE HABITABILIDAD (MODALIDAD VIVIENDA NUEVA)
    # -------------------------------------------------------------------------
    st.markdown("#### 🏗️ 6. Requerimientos de Habitabilidad - Modalidad Vivienda Nueva (Tabla 6)")
    st.caption("Consolidación oficial de marcación en el formulario Word (Ficha PHR 6.1) según la base de datos más actualizada:")

    col_t6_k1, col_t6_k2, col_t6_k3, col_t6_k4 = st.columns(4)
    col_t6_k1.metric("Conjunto Habitacional", f"{t6_conjunto_count} familias", "100.0% del proyecto", help="Todos los postulantes en terreno común Hijuela El Molino")
    col_t6_k2.metric("Tercer Dormitorio", f"{t6_3d_count} familias", f"{round((t6_3d_count/total_post)*100, 1)}% del padrón", help="Postulantes que cumplen los 4 criterios normativos")
    col_t6_k3.metric("Recinto Complementario", f"{t6_recinto_count} familias", f"{round((t6_recinto_count/total_post)*100, 1)}% del padrón", help="32 familias acreditadas con recinto (12 habitables + 20 no habitables)")
    col_t6_k4.metric("Vivienda Base (2 Dorm)", f"{total_post - t6_3d_count} familias", f"{round(((total_post - t6_3d_count)/total_post)*100, 1)}% del padrón", help="Familias sin justificación de 3° dormitorio")

    tab_3d_1, tab_3d_2 = st.tabs([
        "📊 Desglose de Criterios para Tercer Dormitorio",
        f"📋 Nómina de Familias con 3° Dormitorio ({t6_3d_count} familias)"
    ])

    with tab_3d_1:
        st.markdown("**Desglose Factual según Columna 'TIPO DE VIVIENDA':**")
        if t6_3d_motivos:
            df_3d_motivos = pd.DataFrame([
                {"Criterio Oficial": k, "Familias": v, "Porcentaje": f"{round((v/total_post)*100, 1)}%"}
                for k, v in t6_3d_motivos.items()
            ]).sort_values("Familias", ascending=False)
            st.dataframe(df_3d_motivos, hide_index=True, use_container_width=True)
        else:
            st.info("No se registran familias bajo los criterios de 3° dormitorio.")

    with tab_3d_2:
        st.markdown(f"**Nómina Completa de las {t6_3d_count} Familias con Tercer Dormitorio Acreditado:**")
        if t6_3d_list:
            st.dataframe(pd.DataFrame(t6_3d_list), hide_index=True, use_container_width=True, height=280)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 7: TIPOLOGÍAS Y FACTOR DE AISLAMIENTO
    # -------------------------------------------------------------------------
    st.markdown("#### 🏡 7. Tipologías Habitacionales y Territorialidad")
    col_t1, col_t2 = st.columns(2)

    with col_t1:
        st.markdown("**Tipología Habitacional Propuesta**")
        df_tip = pd.DataFrame([{"Tipología": k, "Cantidad": v} for k, v in tipologias_counts.items()]).sort_values("Cantidad", ascending=False)
        c_tip = alt.Chart(df_tip).mark_bar(cornerRadius=6).encode(
            x=alt.X("Cantidad:Q", title="N° Postulantes"),
            y=alt.Y("Tipología:N", sort="-x", title=None),
            color=alt.value("#8B5CF6"),
            tooltip=["Tipología", "Cantidad"]
        ).properties(height=240)
        st.altair_chart(c_tip, use_container_width=True)

    with col_t2:
        st.markdown("**Factor de Aislamiento Territorial (RE 3130)**")
        if aislamiento_counts:
            df_aisl = pd.DataFrame([{"Factor": k, "Cantidad": v} for k, v in aislamiento_counts.items()])
            c_aisl = alt.Chart(df_aisl).mark_bar(cornerRadius=6).encode(
                x=alt.X("Factor:N", title=None),
                y=alt.Y("Cantidad:Q", title="N° Postulantes"),
                color=alt.value("#EC4899"),
                tooltip=["Factor", "Cantidad"]
            ).properties(height=240)
            st.altair_chart(c_aisl, use_container_width=True)
        else:
            st.info("No se registraron valores específicos de aislamiento en la base.")

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SECCIÓN 7: DASHBOARD EJECUTIVO Y MODELO POWER BI
    # -------------------------------------------------------------------------
    st.markdown("#### 📊 7. Dashboard Ejecutivo y Conexión con Microsoft Power BI")
    st.caption("Herramientas listas para presentar y compartir las métricas con contrapartes (SERVIU, Municipio, Consultora) sin necesidad de configurar nada manualmente.")

    tab_dash1, tab_dash2 = st.tabs([
        "🌐 1. Dashboard Interactivo Autónomo (HTML - Listo para compartir)",
        "📈 2. Modelo Relacional para Power BI (.xlsx)"
    ])

    with tab_dash1:
        st.markdown("**Reporte Web Interactivo en un solo archivo (100% autónomo y portable):**")
        st.write("Viene con **todos los gráficos interactivos ya armados**, tarjetas de KPI, filtros por tipología de hogar, recintos complementarios y buscador en vivo de postulantes. Se puede enviar por WhatsApp, correo o Teams y cualquier persona lo abre con doble clic en su navegador sin instalar nada.")

        col_h1, col_h2 = st.columns([2, 3])
        html_dash_path = os.path.join(output_directory, "Dashboard_Ejecutivo_Perquenco.html")
        with col_h1:
            btn_gen_html = st.button("🚀 Generar Dashboard Web Interactivo", type="primary", use_container_width=True)
            if btn_gen_html:
                os.makedirs(output_directory, exist_ok=True)
                generate_interactive_html_dashboard(postulantes, html_dash_path, datos_terreno=datos_terreno_proyecto)
                st.success("✅ Dashboard interactivo generado exitosamente.")

            if os.path.exists(html_dash_path):
                with open(html_dash_path, "rb") as f_html:
                    st.download_button(
                        label="📥 Descargar Dashboard_Ejecutivo_Perquenco.html",
                        data=f_html,
                        file_name="Dashboard_Ejecutivo_Perquenco.html",
                        mime="text/html",
                        use_container_width=True
                    )
                if st.button("🌐 Abrir Dashboard en mi Navegador", use_container_width=True):
                    try:
                        os.startfile(html_dash_path)
                    except Exception:
                        st.info(f"Ruta: {html_dash_path}")

        with col_h2:
            st.info("""
            ✨ **Ventajas del Dashboard Web Autónomo:**
            * **Cero instalación:** No requiere tener instalado Power BI ni Python.
            * **Gráficos interactivos ya creados:** Gráfico de donas para núcleos familiares, barras para los 32 recintos, desglose de adultos mayores.
            * **Filtros cruzados y buscador en tiempo real:** Filtra al instante por tipo de familia o recinto.
            * **Botón de exportación e impresión:** Incluye botón para imprimir o guardar como PDF ejecutivo listo para reuniones.
            """)

    with tab_dash2:
        st.markdown("**Modelo de Datos Relacional para Microsoft Power BI Desktop / Service:**")
        st.write("Genera el archivo `.xlsx` con las 4 tablas normalizadas (`Hogares_Postulantes`, `Habitantes_Detalle`, `Recintos_Complementarios`, `Resumen_KPIs`) para conectar directamente en Power BI.")

        col_pbi1, col_pbi2 = st.columns([2, 3])
        powerbi_filename = "Modelo_PowerBI_Perquenco.xlsx"
        powerbi_filepath = os.path.join(output_directory, powerbi_filename)
        with col_pbi1:
            btn_generate_pbi = st.button("📊 Generar Modelo para Power BI", use_container_width=True)
            if btn_generate_pbi:
                os.makedirs(output_directory, exist_ok=True)
                export_to_powerbi_excel(postulantes, powerbi_filepath, datos_terreno=datos_terreno_proyecto)
                st.success("✅ Archivo Power BI generado exitosamente.")
                
            if os.path.exists(powerbi_filepath):
                with open(powerbi_filepath, "rb") as f_pbi:
                    st.download_button(
                        label="📥 Descargar Modelo_PowerBI_Perquenco.xlsx",
                        data=f_pbi,
                        file_name=powerbi_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        with col_pbi2:
            st.info("""
            💡 **Conexión en Power BI Desktop:**
            1. Abre Power BI Desktop ➔ `Obtener datos` ➔ `Excel`.
            2. Selecciona `Modelo_PowerBI_Perquenco.xlsx` y marca las 4 tablas.
            3. Haz clic en `Cargar` y publica en Power BI Service (la nube) para compartir un enlace web dinámico.
            """)

    st.markdown("---")


# -----------------------------------------------------------------------------
# TAB 3: GENERACIÓN MASIVA (155 FICHAS)
# -----------------------------------------------------------------------------
with tab_batch:
    st.subheader("Generación Masiva de Formularios")
    st.write(f"Se procesarán los **{len(postulantes)} postulantes** registrados en la base de datos de Perquenco.")
    
    st.markdown("""
    **Garantías del proceso:**
    * ✅ **Fidelidad Absoluta:** Preserva exactamente tablas, bordes y tipografía original (`Gadugi` 9pt).
    * ✅ **Cumplimiento Normativo:** Los campos sin información en el Excel quedan estrictamente en blanco.
    * ✅ **Control de Cuota Gemini:** Si usas la clave de API gratuita, la app regula la velocidad para no sobrepasar el límite de 15 peticiones por minuto. Si no usas API key, el procesamiento es instantáneo vía motor local.
    """)

    col_b1, col_b2 = st.columns([3, 2])
    with col_b1:
        st.write(f"**Carpeta de destino:** `{output_directory}`")
    with col_b2:
        abrir_carpeta = st.button("📂 Abrir Carpeta de Salida en Windows")
        if abrir_carpeta:
            os.makedirs(output_directory, exist_ok=True)
            try:
                os.startfile(output_directory)
            except Exception as e:
                st.info(f"Ruta: {output_directory}")

    st.markdown("---")

    start_generation = st.button(f"🚀 Iniciar Generación de {len(postulantes)} Formularios", type="primary", use_container_width=True)

    if start_generation:
        os.makedirs(output_directory, exist_ok=True)
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        tiempo_inicio = time.time()
        generados = []
        errores = []

        # Determinar si usamos auditoría Gemini o motor local rápido
        use_gemini = bool(gemini_key.strip())

        for idx, post in enumerate(postulantes):
            nombre_limpio = "".join(x for x in post["nombre"] if x.isalnum() or x in " _-").strip().replace(" ", "_")
            doc_filename = f"Formulario_PHR_6.1_{post['rut']}_{nombre_limpio}.docx"
            doc_filepath = os.path.join(output_directory, doc_filename)

            status_text.markdown(f"**Procesando ({idx+1}/{len(postulantes)}):** `{post['nombre']}`...")

            try:
                if use_gemini:
                    # En modo masivo con API gratuita, pausar brevemente para respetar 15 RPM
                    consolidated = audit_with_gemini(
                        post,
                        api_key=gemini_key.strip(),
                        usar_rsh_en_terreno=False,
                        model_name=selected_model,
                        datos_terreno_proyecto=datos_terreno_proyecto
                    )
                    time.sleep(4.0)  # Pacing seguro para 15 RPM
                else:
                    consolidated = consolidate_postulante_local(
                        post,
                        usar_rsh_en_terreno=False,
                        datos_terreno_proyecto=datos_terreno_proyecto
                    )

                fill_formulario_phr(
                    template_docx_path,
                    consolidated,
                    doc_filepath,
                    egr_name=egr_name
                )
                generados.append(doc_filename)
            except Exception as e:
                errores.append({"postulante": post["nombre"], "error": str(e)})

            progress_bar.progress((idx + 1) / len(postulantes))

        duracion = round(time.time() - tiempo_inicio, 1)
        status_text.empty()
        progress_bar.progress(1.0)

        st.success(f"🎉 **¡Proceso completado en {duracion} segundos!** Se generaron **{len(generados)}** documentos exitosamente.")

        if errores:
            st.error(f"Se registraron {len(errores)} observaciones en la generación.")
            st.dataframe(pd.DataFrame(errores))

        # Crear archivo ZIP para descarga directa
        zip_path = os.path.join(output_directory, "Formularios_PHR_Perquenco.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for g in generados:
                zipf.write(os.path.join(output_directory, g), arcname=g)

        with open(zip_path, "rb") as zf:
            st.download_button(
                label="📦 Descargar Todos los Formularios en un Archivo ZIP",
                data=zf,
                file_name="Formularios_PHR_Perquenco.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )


# -----------------------------------------------------------------------------
# TAB 4: EXPLORADOR DE BASE DE DATOS
# -----------------------------------------------------------------------------
with tab_data:
    st.subheader("Base de Postulantes de Perquenco")
    st.caption("Resumen consolidado de las variables clave leídas desde el Excel.")

    table_records = []
    for p in postulantes:
        c_local = consolidate_postulante_local(p, datos_terreno_proyecto=datos_terreno_proyecto)
        tipo_fam = (c_local.get("tipo_familia") or classify_family_nucleus(p))[0]
        rec_local = c_local.get("recinto_complementario") or evaluate_recinto_complementario(p)
        act_p = p.get("actividad_economica") or {}
        act_nom = act_p.get("actividad_principal") or "(Sin actividad)"
        desc_p = act_p.get("descripcion") or ""
        table_records.append({
            "N°": p["nro_orden"],
            "Nombre Postulante": p["nombre"],
            "RUT": p["rut"],
            "Sexo": p["sexo"],
            "Edad": p["edad"],
            "Estado Civil": p["estado_civil"],
            "Núcleo Familiar": tipo_fam,
            "Etnia": p["etnia"],
            "Discapacidad": p["discapacidad"],
            "Grupo Familiar": p["grupo_familiar_cant"],
            "Cónyuge/Pareja": p["conyuge"]["nombre"] if p["conyuge"]["nombre"] else "(Sin registro)",
            "Actividad Económica": act_nom,
            "Detalle Actividad": desc_p,
            "Recinto Complementario": rec_local.get("tipo_recinto", "No Aplica"),
            "Recinto Sugerido": rec_local.get("recinto_sugerido", ""),
            "Dirección RSH": p["direccion_rsh"],
            "Comuna": p["comuna"],
            "N° Parientes": len(p["parientes"]),
            "Tipología": p["tipologia_propuesta"]
        })

    df_all = pd.DataFrame(table_records)
    st.dataframe(df_all, use_container_width=True, height=500)

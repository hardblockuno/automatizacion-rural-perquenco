"""
Módulo para completar el formulario Word oficial (Formulario PHR N°6.1 DTS)
preservando la tipografía (Gadugi), bordes, anchos de columna y estilos originales.
Aplica estrictamente los datos consolidados en los ítems 1 y 1.2.
"""

import os
import re
import docx
from docx.shared import Pt
from typing import Dict, Any, Optional


def set_cell_value(
    cell,
    text: str,
    font_name: str = "Gadugi",
    font_size_pt: float = 9.0,
    bold: bool = False,
    center: bool = False
) -> None:
    """Escribe un valor en una celda preservando el formato y estilo."""
    while len(cell.paragraphs) > 1:
        p_extra = cell.paragraphs[-1]._p
        p_extra.getparent().remove(p_extra)
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    p.text = str(text) if text is not None else ""
    if center:
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    if p.runs:
        run = p.runs[0]
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.bold = bold


def fill_formulario_phr(
    template_path: str,
    data: Dict[str, Any],
    output_path: str,
    egr_name: Optional[str] = "CONSULTORA PLAN SOCIAL LIMITADA",
    profesionales_firmantes: Optional[Dict[str, Any]] = None
) -> str:
    """
    Rellena los ítems 1 (Tablas 1, 2, 3, 4) y 1.2 (Tabla 5) en el documento Word.
    Guarda el documento en output_path y retorna la ruta del archivo generado.
    """
    doc = docx.Document(template_path)

    # Actualizar Nombre de EGR en el encabezado del documento
    final_egr = (egr_name or "CONSULTORA PLAN SOCIAL LIMITADA").strip()
    for p in doc.paragraphs[:15]:
        if "IDENTIFICACIÓN DE LA ENTIDAD DE GESTIÓN RURAL" in p.text.upper():
            p.text = f"IDENTIFICACIÓN DE LA ENTIDAD DE GESTIÓN RURAL RESPONSABLE DEL PRESENTE DIAGNÓSTICO: {final_egr}"
            if p.runs:
                p.runs[0].font.name = "Gadugi"
                p.runs[0].font.size = Pt(9)
            break

    t1_data = data.get("tabla_1", {})
    t2_data = data.get("tabla_2", {})
    t3_data = data.get("tabla_3", {})
    t4_data = data.get("tabla_4", {})
    t5_data = data.get("tabla_5", {})

    # =========================================================================
    # TABLA 1: Antecedentes de él o la Postulante y su Cónyuge o Conviviente
    # =========================================================================
    t1 = doc.tables[1]
    # Fila 1: Titular Nombre y RUT
    set_cell_value(t1.cell(1, 1), t1_data.get("titular_nombre", ""))
    set_cell_value(t1.cell(1, 3), t1_data.get("titular_rut", ""))
    # Fila 2: Fecha Nacimiento y Estado Civil
    set_cell_value(t1.cell(2, 1), t1_data.get("titular_fecha_nac", ""))
    set_cell_value(t1.cell(2, 3), t1_data.get("titular_estado_civil", ""))
    # Fila 3: Sexo y Teléfono
    set_cell_value(t1.cell(3, 1), t1_data.get("titular_sexo", ""))
    set_cell_value(t1.cell(3, 3), t1_data.get("titular_telefono", ""))
    # Fila 4: Cónyuge/Conviviente Nombre y RUT
    set_cell_value(t1.cell(4, 1), t1_data.get("conyuge_nombre", ""))
    set_cell_value(t1.cell(4, 3), t1_data.get("conyuge_rut", ""))
    # Fila 5: Cónyuge Sexo y Fecha Nacimiento
    set_cell_value(t1.cell(5, 1), t1_data.get("conyuge_sexo", ""))
    set_cell_value(t1.cell(5, 3), t1_data.get("conyuge_fecha_nac", ""))

    # =========================================================================
    # TABLA 2: Domicilio en que reside (según RSH)
    # =========================================================================
    t2 = doc.tables[2]
    # Fila 1: Camino, calle, avenida o pasaje (celda combinada cols 1-3)
    set_cell_value(t2.cell(1, 1), t2_data.get("calle", ""))
    # Fila 2: Número y Lote/Hijuela
    set_cell_value(t2.cell(2, 1), t2_data.get("numero", ""))
    set_cell_value(t2.cell(2, 3), t2_data.get("lote", ""))
    # Fila 3: Región y Comuna
    set_cell_value(t2.cell(3, 1), t2_data.get("region", "La Araucanía"))
    set_cell_value(t2.cell(3, 3), t2_data.get("comuna", "PERQUENCO"))
    # Fila 4: Provincia y Localidad (INE)
    set_cell_value(t2.cell(4, 1), t2_data.get("provincia", "Cautín"))
    set_cell_value(t2.cell(4, 3), t2_data.get("localidad", ""))

    # =========================================================================
    # TABLA 3: Dirección en que se aplicará el Subsidio
    # =========================================================================
    t3 = doc.tables[3]
    # Fila 1: Camino, calle, avenida o pasaje | Lote, hijuela, casa
    set_cell_value(t3.cell(1, 1), t3_data.get("calle", ""))
    set_cell_value(t3.cell(1, 3), t3_data.get("lote", ""))
    # Fila 2: Localidad (INE) | Factor de Aislamiento
    set_cell_value(t3.cell(2, 1), t3_data.get("localidad", ""))
    set_cell_value(t3.cell(2, 3), t3_data.get("factor_aislamiento", ""))
    # Fila 3: Comuna | Provincia
    set_cell_value(t3.cell(3, 1), t3_data.get("comuna", "PERQUENCO"))
    set_cell_value(t3.cell(3, 3), t3_data.get("provincia", "Cautín"))
    # Fila 4: Región | Rol de Propiedad del SII
    set_cell_value(t3.cell(4, 1), t3_data.get("region", "La Araucanía"))
    set_cell_value(t3.cell(4, 3), t3_data.get("rol_sii", ""))

    # =========================================================================
    # TABLA 4: Antecedentes del Grupo Familiar
    # =========================================================================
    t4 = doc.tables[4]
    # Fila 1: ¿Cuántas personas habitan actualmente en la Vivienda?
    tot_hab = t4_data.get("total_habitantes", "")
    set_cell_value(t4.cell(1, 1), f"N° {tot_hab}" if tot_hab else "", bold=True)

    # Mapeo de filas en Tabla 4
    cat_rows = {
        "hombres": 3,
        "mujeres": 4,
        "menores_18": 5,
        "adultos_mayores": 6,
        "discapacidad": 7,
        "indigena": 8,
        "extranjeros": 9
    }

    for cat_key, r_idx in cat_rows.items():
        cat_info = t4_data.get(cat_key, {})
        set_cell_value(t4.cell(r_idx, 1), cat_info.get("si", ""), center=True)
        set_cell_value(t4.cell(r_idx, 2), cat_info.get("no", ""), center=True)
        set_cell_value(t4.cell(r_idx, 3), cat_info.get("cuantos", ""), center=True)
        set_cell_value(t4.cell(r_idx, 4), cat_info.get("observaciones", ""))

    # =========================================================================
    # TABLA 5: 1.2 Actividades Económicas del Grupo Familiar
    # =========================================================================
    t5 = doc.tables[5]
    # Filas: 2=Agricultura, 3=Forestal, 4=Pesca, 5=Minería, 6=Turismo Rural, 7=Servicios, 8=Otras
    # Columnas: 1=Jefe M, 2=Jefe F, 3=Conyuge M, 4=Conyuge F, 5=Otros M, 6=Otros F
    categoria = t5_data.get("categoria", "")
    especificacion = t5_data.get("especificacion", "")
    actividades = t5_data.get("actividades", {})

    # Fila 8 Celda 0: 'Otras (Especificar)'
    cell_8_0 = t5.cell(8, 0)
    if categoria == "otras" and especificacion:
        set_cell_value(cell_8_0, f"Otras (Especificar: {especificacion})", font_name="Gadugi", font_size_pt=8.5, bold=True, center=False)
    else:
        set_cell_value(cell_8_0, "Otras (Especificar)", font_name="Gadugi", font_size_pt=9.0, bold=True, center=False)

    act_rows = {
        "agricultura": 2,
        "forestal": 3,
        "pesca": 4,
        "mineria": 5,
        "turismo_rural": 6,
        "servicios": 7,
        "otras": 8
    }

    for act_key, r_idx in act_rows.items():
        if act_key in actividades:
            marks = actividades[act_key]  # Dict con keys: 'jefe_m', 'jefe_f', 'conyuge_m', etc.
            set_cell_value(t5.cell(r_idx, 1), marks.get("jefe_m", ""), center=True)
            set_cell_value(t5.cell(r_idx, 2), marks.get("jefe_f", ""), center=True)
            set_cell_value(t5.cell(r_idx, 3), marks.get("conyuge_m", ""), center=True)
            set_cell_value(t5.cell(r_idx, 4), marks.get("conyuge_f", ""), center=True)
            set_cell_value(t5.cell(r_idx, 5), marks.get("otros_m", ""), center=True)
            set_cell_value(t5.cell(r_idx, 6), marks.get("otros_f", ""), center=True)
        else:
            # Asegurar que permanezcan en blanco
            for c_idx in range(1, 7):
                set_cell_value(t5.cell(r_idx, c_idx), "", center=True)

    # =========================================================================
    # TABLA 6: REQUERIMIENTOS DE HABITABILIDAD - MODALIDAD VIVIENDA NUEVA
    # =========================================================================
    if len(doc.tables) > 6:
        t6 = doc.tables[6]
        t6_data = data.get("tabla_6", {})
        ch_data = t6_data.get("conjunto_habitacional", {})
        sr_data = t6_data.get("sitio_residente", {})

        # Fila 1: Conjunto Habitacional
        # Columna 1: Vivienda Nueva (100% de los casos en este proyecto colectivo)
        set_cell_value(t6.cell(1, 1), ch_data.get("vivienda_nueva", "X"), font_size_pt=10.0, bold=True, center=True)
        # Columna 2: Tercer Dormitorio (solo personas bajo criterios específicos)
        set_cell_value(t6.cell(1, 2), ch_data.get("tercer_dormitorio", ""), font_size_pt=10.0, bold=True, center=True)
        # Columna 3: Recinto Complementario (solo los 32 casos que efectivamente lo requieren)
        set_cell_value(t6.cell(1, 3), ch_data.get("recinto_complementario", ""), font_size_pt=10.0, bold=True, center=True)

        # Fila 2: Sitio del Residente (permanece en blanco en este proyecto)
        set_cell_value(t6.cell(2, 1), sr_data.get("vivienda_nueva", ""), font_size_pt=10.0, bold=True, center=True)
        set_cell_value(t6.cell(2, 2), sr_data.get("tercer_dormitorio", ""), font_size_pt=10.0, bold=True, center=True)
        set_cell_value(t6.cell(2, 3), sr_data.get("recinto_complementario", ""), font_size_pt=10.0, bold=True, center=True)

    # =========================================================================
    # TABLA 8: JUSTIFICACIÓN DE RECINTO(S) COMPLEMENTARIO(S), SI PROCEDE
    # =========================================================================
    if len(doc.tables) > 8:
        t8 = doc.tables[8]
        t8_data = data.get("tabla_8")
        if not t8_data:
            from src.gemini_auditor import build_tabla_8
            t8_data = build_tabla_8({}, data.get("recinto_complementario", {}))

        # Mapeo oficial de filas en Tabla 8 de la plantilla DTS:
        # Fila 1: Bodega
        # Fila 2: Recinto para realizar actividades productivas
        # Fila 3: Otros Recintos Techados Adosados a la Vivienda
        # Fila 4: Leñera
        # Fila 5: Otros (especificar)
        row_keys = [
            (1, "bodega"),
            (2, "actividades_productivas"),
            (3, "otros_adosados"),
            (4, "lenera"),
            (5, "otros_especificar")
        ]

        for r_idx, r_key in row_keys:
            r_info = t8_data.get(r_key, {})
            val_si_no = r_info.get("si_no", "No") if r_info else "No"
            just_text = r_info.get("justificacion", "") if r_info else ""

            # Columna 1: Marcar 'Sí' o 'No' (centrado, Gadugi 9.5pt, negrita si 'Sí')
            set_cell_value(t8.cell(r_idx, 1), val_si_no, font_size_pt=9.5, bold=(val_si_no == "Sí"), center=True)

            # Columna 2: Explicar actividad previa y justificación técnica (Gadugi 8.5pt)
            set_cell_value(t8.cell(r_idx, 2), just_text, font_size_pt=8.5, bold=False, center=False)

            # Fila 5: Columna 0 especificar si aplica
            if r_idx == 5:
                esp = r_info.get("especificacion", "")
                if esp and val_si_no == "Sí":
                    set_cell_value(t8.cell(5, 0), f"Otros (especificar: {esp})", font_size_pt=9.0, bold=False, center=False)
                else:
                    set_cell_value(t8.cell(5, 0), "Otros (especificar)", font_size_pt=9.0, bold=False, center=False)

    # =========================================================================
    # TABLA 9: REQUERIMIENTOS DE HABITABILIDAD ASOCIADOS AL EQUIPAMIENTO COMUNITARIO
    # =========================================================================
    if len(doc.tables) > 9:
        t9 = doc.tables[9]
        t9_data = data.get("tabla_9")
        if not t9_data:
            from src.gemini_auditor import build_tabla_9
            t9_data = build_tabla_9()

        c_info = t9_data.get("construccion", {})
        obras_val = c_info.get("obras", "1")
        desc_val = c_info.get("descripcion", "Recinto de acopio y de apoyo de producción agrícola sustentable")

        # Fila 1: Mejoramiento del Equipamiento Comunitario Existente (en blanco)
        set_cell_value(t9.cell(1, 1), "", center=True)
        set_cell_value(t9.cell(1, 2), "", center=False)

        # Fila 2: Ampliación del Equipamiento Comunitario Existente (en blanco)
        set_cell_value(t9.cell(2, 1), "", center=True)
        set_cell_value(t9.cell(2, 2), "", center=False)

        # Fila 3: Construcción de Equipamiento Comunitario
        set_cell_value(t9.cell(3, 1), obras_val, font_size_pt=9.5, bold=True, center=True)
        set_cell_value(t9.cell(3, 2), desc_val, font_size_pt=9.0, bold=False, center=False)

    # =========================================================================
    # TABLA 15: ACCESO A SERVICIOS BÁSICOS EN TERRENOS ERIAZOS (D.S. N°10)
    # =========================================================================
    if len(doc.tables) > 15:
        t15 = doc.tables[15]
        t15_data = data.get("tabla_15")
        if not t15_data:
            from src.gemini_auditor import build_tabla_15
            t15_data = build_tabla_15()

        # Mapeo de filas en Tabla 15 de la plantilla Word:
        # Agua Potable (Filas 2 a 8)
        agua_rows = t15_data.get("agua_potable", [])
        for idx, item in enumerate(agua_rows):
            r_idx = 2 + idx
            if r_idx < len(t15.rows):
                set_cell_value(t15.cell(r_idx, 1), item.get("si", ""), font_size_pt=9.5, bold=True, center=True)
                set_cell_value(t15.cell(r_idx, 2), item.get("no", ""), font_size_pt=9.5, bold=True, center=True)
                set_cell_value(t15.cell(r_idx, 3), item.get("observaciones", ""), font_size_pt=8.5, bold=False, center=False)

        # Alcantarillado (Filas 10 a 14)
        alc_rows = t15_data.get("alcantarillado", [])
        for idx, item in enumerate(alc_rows):
            r_idx = 10 + idx
            if r_idx < len(t15.rows):
                set_cell_value(t15.cell(r_idx, 1), item.get("si", ""), font_size_pt=9.5, bold=True, center=True)
                set_cell_value(t15.cell(r_idx, 2), item.get("no", ""), font_size_pt=9.5, bold=True, center=True)
                obs_txt = item.get("observaciones", "")
                set_cell_value(t15.cell(r_idx, 3), obs_txt, font_size_pt=8.5, bold=False, center=False)

        # Electricidad (Filas 16 a 21)
        elec_rows = t15_data.get("electricidad", [])
        for idx, item in enumerate(elec_rows):
            r_idx = 16 + idx
            if r_idx < len(t15.rows):
                set_cell_value(t15.cell(r_idx, 1), item.get("si", ""), font_size_pt=9.5, bold=True, center=True)
                set_cell_value(t15.cell(r_idx, 2), item.get("no", ""), font_size_pt=9.5, bold=True, center=True)
                set_cell_value(t15.cell(r_idx, 3), item.get("observaciones", ""), font_size_pt=8.5, bold=False, center=False)

    # =========================================================================
    # TABLAS 17, 18, 19: PROFESIONALES SUSCRIBIENTES Y POSTULANTE
    # =========================================================================
    prof_data = profesionales_firmantes or data.get("firmantes") or {}
    tec = prof_data.get("tecnico", {})
    soc = prof_data.get("social", {})
    post = prof_data.get("postulante", {})

    egr_tec = tec.get("egr", "CONSULTORA PLAN SOCIAL LIMITADA")
    nom_tec = tec.get("nombre", "Vanessa Schneider Martínez")
    rut_tec = tec.get("rut", "13.730.440-6")
    prof_tec = tec.get("profesion", "Arquitecta")

    egr_soc = soc.get("egr", "CONSULTORA PLAN SOCIAL")
    nom_soc = soc.get("nombre", "Erica Reyes Peña")
    rut_soc = soc.get("rut", "20.402.240-2")
    prof_soc = soc.get("profesion", "TRABAJADORA SOCIAL")

    t1_data = data.get("tabla_1", {})
    nom_post = post.get("nombre") or t1_data.get("titular_nombre", "")
    rut_post = post.get("rut") or t1_data.get("titular_rut", "")

    # Búsqueda dinámica de tablas por texto o por índice de respaldo (17, 18, 19)
    t17, t18, t19 = None, None, None
    for tbl in doc.tables:
        if len(tbl.rows) >= 5:
            r1_txt = tbl.cell(1, 0).text.upper() if len(tbl.rows) > 1 and len(tbl.columns) > 0 else ""
            if "PROFESIONAL" in r1_txt and ("TÉCNICA" in r1_txt or "TECNICA" in r1_txt):
                t17 = tbl
            elif "PROFESIONAL" in r1_txt and "SOCIAL" in r1_txt:
                t18 = tbl
            elif "POSTULANTE" in r1_txt and len(tbl.rows) >= 4:
                t19 = tbl

    if t17 is None and len(doc.tables) > 17:
        t17 = doc.tables[17]
    if t18 is None and len(doc.tables) > 18:
        t18 = doc.tables[18]
    if t19 is None and len(doc.tables) > 19:
        t19 = doc.tables[19]

    if t17:
        set_cell_value(t17.cell(1, 2), egr_tec, font_size_pt=9.0, bold=False, center=False)
        set_cell_value(t17.cell(2, 2), nom_tec, font_size_pt=9.0, bold=False, center=False)
        set_cell_value(t17.cell(3, 2), rut_tec, font_size_pt=9.0, bold=False, center=False)
        set_cell_value(t17.cell(4, 2), prof_tec, font_size_pt=9.0, bold=False, center=False)

    if t18:
        set_cell_value(t18.cell(1, 2), egr_soc, font_size_pt=9.0, bold=False, center=False)
        set_cell_value(t18.cell(2, 2), nom_soc, font_size_pt=9.0, bold=False, center=False)
        set_cell_value(t18.cell(3, 2), rut_soc, font_size_pt=9.0, bold=False, center=False)
        set_cell_value(t18.cell(4, 2), prof_soc, font_size_pt=9.0, bold=False, center=False)

    if t19:
        set_cell_value(t19.cell(2, 2), nom_post, font_size_pt=9.0, bold=False, center=False)
        set_cell_value(t19.cell(3, 2), rut_post, font_size_pt=9.0, bold=False, center=False)

    # Crear carpeta destino si no existe
    dir_name = os.path.dirname(output_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
    doc.save(output_path)
    return output_path

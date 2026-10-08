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
    egr_name: Optional[str] = "CONSULTORA PLAN SOCIAL LIMITADA"
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
    # Limpiar párrafos adicionales previos si existieran
    while len(cell_8_0.paragraphs) > 1:
        p_extra = cell_8_0.paragraphs[-1]._p
        p_extra.getparent().remove(p_extra)
    cell_8_0.paragraphs[0].text = "Otras (Especificar)"
    if cell_8_0.paragraphs[0].runs:
        cell_8_0.paragraphs[0].runs[0].font.name = "Gadugi"
        cell_8_0.paragraphs[0].runs[0].font.size = Pt(9)
        cell_8_0.paragraphs[0].runs[0].font.bold = True

    # Si la actividad corresponde a 'otras' y trae especificación, agregarla justo debajo
    if categoria == "otras" and especificacion:
        p_esp = cell_8_0.add_paragraph()
        p_esp.paragraph_format.space_before = Pt(1)
        p_esp.paragraph_format.space_after = Pt(1)
        r_esp = p_esp.add_run(especificacion)
        r_esp.font.name = "Gadugi"
        r_esp.font.size = Pt(8.0)
        r_esp.font.italic = True

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

    # Crear carpeta destino si no existe
    dir_name = os.path.dirname(output_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
    doc.save(output_path)
    return output_path

"""
Módulo para exportar la base de datos y métricas del proyecto
en un formato optimizado para Microsoft Power BI.
Genera tablas limpias, normalizadas y listas para dashboards interactivos.
"""

import os
import re
import pandas as pd
from typing import List, Dict, Any, Optional

from src.gemini_auditor import (
    consolidate_postulante_local,
    infer_gender_from_name,
    classify_family_nucleus,
    evaluate_recinto_complementario
)


def get_age_group(age: Optional[int]) -> str:
    if age is None or age < 0:
        return "Sin Información"
    if age < 18:
        return "Menor de 18"
    if age < 30:
        return "18 a 29 años"
    if age < 45:
        return "30 a 44 años"
    if age < 60:
        return "45 a 59 años"
    if age < 70:
        return "60 a 69 años (Adulto Mayor)"
    return "70 años o más (Adulto Mayor)"


def build_powerbi_dataset(postulantes: List[Dict[str, Any]], datos_terreno: Optional[Dict[str, str]] = None) -> Dict[str, pd.DataFrame]:
    """
    Construye las tablas del modelo relacional para Power BI:
    1. Hogares_Postulantes (Tabla de hechos a nivel de hogar / titular)
    2. Habitantes_Detalle (Tabla de dimensión de personas / integrantes)
    3. Recintos_Complementarios (Detalle de los 32 casos con recinto)
    4. Resumen_Metricas (KPIs consolidados para tarjetas)
    """
    hogares_rows = []
    habitantes_rows = []
    recintos_rows = []
    criterios_sociales_rows = []

    for p in postulantes:
        c = consolidate_postulante_local(p, datos_terreno_proyecto=datos_terreno)
        t1 = c.get("tabla_1", {})
        t4 = c.get("tabla_4", {})
        t5 = c.get("tabla_5", {})
        tipo_fam, desc_fam = c.get("tipo_familia", ("Sin clasificar", ""))
        rec = c.get("recinto_complementario", {})

        id_post = p.get("nro_orden")
        rut_titular = p.get("rut", "")
        nom_titular = p.get("nombre", "")
        sexo_titular = t1.get("titular_sexo", "")
        sexo_tit_desc = "Femenino" if sexo_titular == "F" else ("Masculino" if sexo_titular == "M" else "Sin Información")
        edad_tit = p.get("edad")
        tramo_tit = get_age_group(edad_tit)

        cant_hab = p.get("grupo_familiar_cant") or 1
        n_menores = int(t4.get("menores_18", {}).get("cuantos", 0) or 0)
        n_mayores = int(t4.get("adultos_mayores", {}).get("cuantos", 0) or 0)
        
        # Pertenencia y vulnerabilidad
        es_mapuche = "Sí" if "MAPUCHE" in (p.get("etnia") or "").upper() else "No"
        disc_val = str(p.get("discapacidad") or "").strip()
        tiene_disc = "Sí" if disc_val and disc_val.upper() not in ["NO", "NONE", "NAN", ""] else "No"

        # Actividad económica (Tabla 5)
        es_social = t5.get("sin_actividad_declarada", False)
        glosa_soc = t5.get("criterio_social", "")
        if es_social:
            tiene_act = "No (Declarada)"
            es_crit_social = "Sí"
            act_nombre = f"Sin actividad declarada ({glosa_soc})"
            sector_t5 = f"Otras (Especificar: {glosa_soc})"
            desc_act = f"Criterio normativo social ({glosa_soc})"
            glosa_t5 = glosa_soc
        else:
            tiene_act = "Sí" if t5.get("categoria") else "No"
            es_crit_social = "No"
            act_nombre = t5.get("actividad_fuente", "") or "(Sin actividad)"
            esp_val = t5.get("especificacion", "")
            if t5.get("categoria") == "otras" and esp_val:
                sector_t5 = f"Otras (Especificar: {esp_val})"
            else:
                sector_t5 = (t5.get("categoria", "") or "Sin actividad").capitalize()
            desc_act = t5.get("descripcion_fuente", "")
            glosa_t5 = esp_val

        # Recinto complementario
        tipo_recinto = rec.get("tipo_recinto", "No Aplica")
        procede_recinto = "Sí" if tipo_recinto in ["Habitable", "No Habitable"] else "No"
        subtipo_recinto = rec.get("recinto_sugerido", "No Aplica")
        justif_recinto = rec.get("justificacion", "")

        t8 = c.get("tabla_8", {})
        fila_t8 = "No Aplica"
        if t8.get("aplica_recinto"):
            tr = t8.get("target_row")
            if tr == 1:
                fila_t8 = "Fila 1: Bodega"
            elif tr == 2:
                fila_t8 = "Fila 2: Recinto para realizar actividades productivas"
            elif tr == 3:
                fila_t8 = "Fila 3: Otros Recintos Techados Adosados"
            elif tr == 4:
                fila_t8 = "Fila 4: Leñera"
            elif tr == 5:
                esp_t8 = t8.get("otros_especificar", {}).get("especificacion", "")
                fila_t8 = f"Fila 5: Otros (especificar: {esp_t8})" if esp_t8 else "Fila 5: Otros (especificar)"

        # Terreno
        t3 = c.get("tabla_3", {})

        # 1. Registro Hogares
        hogares_rows.append({
            "ID_Postulante": id_post,
            "Nombre_Titular": nom_titular,
            "RUT_Titular": rut_titular,
            "Sexo_Titular": sexo_tit_desc,
            "Edad_Titular": edad_tit,
            "Tramo_Etario_Titular": tramo_tit,
            "Estado_Civil": p.get("estado_civil", "Sin información"),
            "Cantidad_Habitantes": cant_hab,
            "Tipologia_Hogar": tipo_fam,
            "Descripcion_Tipologia": desc_fam,
            "Cantidad_Menores_18": n_menores,
            "Cantidad_Adultos_Mayores": n_mayores,
            "Tiene_Adulto_Mayor": "Sí" if n_mayores > 0 else "No",
            "Pueblo_Originario_Mapuche": es_mapuche,
            "Tiene_Discapacidad": tiene_disc,
            "Detalle_Discapacidad": t4.get("discapacidad", {}).get("observaciones", "") if tiene_disc == "Sí" else "",
            "Direccion_RSH": p.get("direccion_rsh", ""),
            "Comuna": p.get("comuna", "Perquenco"),
            "Factor_Aislamiento": p.get("factor_aislamiento", "1.2"),
            "Tiene_Actividad_Economica": tiene_act,
            "Es_Criterio_Social_Tabla5": es_crit_social,
            "Criterio_Social_Asignado": glosa_soc if es_social else "No Aplica",
            "Actividad_Declarada": act_nombre,
            "Sector_MINVU_Tabla5": sector_t5,
            "Glosa_Especificada_Tabla5": glosa_t5,
            "Detalle_Actividad": desc_act,
            "Procede_Recinto": procede_recinto,
            "Categoria_Recinto": tipo_recinto,
            "Subtipo_Recinto": subtipo_recinto,
            "Fila_Formulario_Tabla8": fila_t8,
            "Justificacion_Normativa_DS10": justif_recinto,
            "Modalidad_Proyecto": "Conjunto Habitacional",
            "Requiere_3er_Dormitorio": "Sí" if c.get("tabla_6", {}).get("conjunto_habitacional", {}).get("tercer_dormitorio") == "X" else "No",
            "Criterio_3er_Dormitorio": c.get("tabla_6", {}).get("conjunto_habitacional", {}).get("motivo_tercer_dormitorio") or "Vivienda base 2 dormitorios",
            "Tipo_Vivienda_Base": p.get("tipo_vivienda", ""),
            "Terreno_Proyecto_Calle": t3.get("calle", ""),
            "Terreno_Proyecto_Lote": t3.get("lote", ""),
            "Terreno_Proyecto_RolSII": t3.get("rol_sii", "")
        })

        # 2. Habitantes (Detalle de cada miembro)
        # Titular
        habitantes_rows.append({
            "ID_Postulante": id_post,
            "RUT_Familia": rut_titular,
            "Nombre_Integrante": nom_titular,
            "RUT_Integrante": rut_titular,
            "Rol_Familiar": "Titular (Jefe/a)",
            "Sexo": sexo_tit_desc,
            "Edad": edad_tit,
            "Tramo_Etario": tramo_tit,
            "Es_Adulto_Mayor": "Sí" if (edad_tit and edad_tit >= 60) else "No",
            "Es_Menor_18": "Sí" if (edad_tit and edad_tit < 18) else "No"
        })

        # Cónyuge
        cony = p.get("conyuge", {})
        if cony.get("nombre"):
            c_edad = cony.get("edad")
            c_sexo_raw = cony.get("sexo", "").upper()
            c_sexo = "Femenino" if c_sexo_raw == "F" else ("Masculino" if c_sexo_raw == "M" else "Sin información")
            habitantes_rows.append({
                "ID_Postulante": id_post,
                "RUT_Familia": rut_titular,
                "Nombre_Integrante": cony.get("nombre"),
                "RUT_Integrante": cony.get("rut") or "(S/I)",
                "Rol_Familiar": "Cónyuge / Conviviente",
                "Sexo": c_sexo,
                "Edad": c_edad,
                "Tramo_Etario": get_age_group(c_edad),
                "Es_Adulto_Mayor": "Sí" if (c_edad and c_edad >= 60) else "No",
                "Es_Menor_18": "Sí" if (c_edad and c_edad < 18) else "No"
            })

        # Parientes
        for par in p.get("parientes", []):
            par_nom = par.get("nombre", "")
            par_edad = par.get("edad")
            par_rel = par.get("parentesco", "Pariente").capitalize()
            p_sexo_inf = infer_gender_from_name(par_nom)
            par_sexo = "Femenino" if p_sexo_inf == "F" else ("Masculino" if p_sexo_inf == "M" else "Sin información")
            habitantes_rows.append({
                "ID_Postulante": id_post,
                "RUT_Familia": rut_titular,
                "Nombre_Integrante": par_nom,
                "RUT_Integrante": par.get("rut") or "(S/I)",
                "Rol_Familiar": par_rel,
                "Sexo": par_sexo,
                "Edad": par_edad,
                "Tramo_Etario": get_age_group(par_edad),
                "Es_Adulto_Mayor": "Sí" if (par_edad and par_edad >= 60) else "No",
                "Es_Menor_18": "Sí" if (par_edad and par_edad < 18) else "No"
            })

        # 3. Recintos complementarios (Solo los 32 casos que aplican)
        if procede_recinto == "Sí":
            recintos_rows.append({
                "ID_Postulante": id_post,
                "Nombre_Postulante": nom_titular,
                "RUT_Postulante": rut_titular,
                "Fila_Formulario_Tabla8": fila_t8,
                "Categoria_Recinto": tipo_recinto,
                "Subtipo_Recinto": subtipo_recinto,
                "Actividad_Productiva": rec.get("detalle_actividad", act_nombre),
                "Justificacion_Normativa": justif_recinto
            })

        # 4. Criterios sociales Tabla 5 (Los 94 casos sin actividad declarada)
        if es_social:
            criterios_sociales_rows.append({
                "ID_Postulante": id_post,
                "Nombre_Postulante": nom_titular,
                "RUT_Postulante": rut_titular,
                "Sexo_Titular": sexo_tit_desc,
                "Edad_Titular": edad_tit,
                "Criterio_Social_Asignado": glosa_soc,
                "Fila_Formulario_Tabla5": f"Otras (Especificar: {glosa_soc})",
                "Casilla_Marcada": "Jefe de Hogar (F)" if sexo_titular == "F" else "Jefe de Hogar (M)",
                "Fundamento_Normativo": f"Titular {sexo_tit_desc.lower()} con edad {edad_tit} años sin actividad declarada en base"
            })

    df_hogares = pd.DataFrame(hogares_rows)
    df_habitantes = pd.DataFrame(habitantes_rows)
    df_recintos = pd.DataFrame(recintos_rows)
    df_criterios_sociales = pd.DataFrame(criterios_sociales_rows)

    # 5. Tabla de KPIs Resumen
    tot_post = len(hogares_rows)
    tot_hab = int(df_hogares["Cantidad_Habitantes"].sum())
    tot_muj = sum(1 for h in hogares_rows if h["Sexo_Titular"] == "Femenino")
    tot_mono = sum(1 for h in hogares_rows if h["Tipologia_Hogar"] == "Monoparental Femenino")
    tot_am = sum(1 for hab in habitantes_rows if hab["Es_Adulto_Mayor"] == "Sí")
    tot_rec_hab = sum(1 for h in hogares_rows if h["Categoria_Recinto"] == "Habitable")
    tot_rec_nohab = sum(1 for h in hogares_rows if h["Categoria_Recinto"] == "No Habitable")

    kpi_rows = [
        {"Indicador": "Total Postulantes (Hogares)", "Valor": tot_post, "Categoria": "Demografía"},
        {"Indicador": "Población Beneficiaria Total (Habitantes)", "Valor": tot_hab, "Categoria": "Demografía"},
        {"Indicador": "Promedio Habitantes por Hogar", "Valor": round(tot_hab / tot_post, 2), "Categoria": "Demografía"},
        {"Indicador": "Jefatura de Hogar Femenina", "Valor": tot_muj, "Categoria": "Vulnerabilidad"},
        {"Indicador": "% Jefatura Femenina", "Valor": f"{round((tot_muj/tot_post)*100, 1)}%", "Categoria": "Vulnerabilidad"},
        {"Indicador": "Hogares Monoparentales Femeninos", "Valor": tot_mono, "Categoria": "Vulnerabilidad"},
        {"Indicador": "% Monoparental Femenino", "Valor": f"{round((tot_mono/tot_post)*100, 1)}%", "Categoria": "Vulnerabilidad"},
        {"Indicador": "Total Adultos Mayores (≥60 años)", "Valor": tot_am, "Categoria": "Tercera Edad"},
        {"Indicador": "Actividad Económica Declarada en Base", "Valor": sum(1 for h in hogares_rows if h["Es_Criterio_Social_Tabla5"] == "No"), "Categoria": "Tabla 5 MINVU"},
        {"Indicador": "Criterio Social Normativo Asignado", "Valor": len(criterios_sociales_rows), "Categoria": "Tabla 5 MINVU"},
        {"Indicador": "Criterio Social: Dueña de Casa (<60)", "Valor": sum(1 for c in criterios_sociales_rows if c["Criterio_Social_Asignado"] == "Dueña de casa"), "Categoria": "Tabla 5 MINVU"},
        {"Indicador": "Criterio Social: Jubilada (≥60)", "Valor": sum(1 for c in criterios_sociales_rows if c["Criterio_Social_Asignado"] == "Jubilada"), "Categoria": "Tabla 5 MINVU"},
        {"Indicador": "Criterio Social: Cesante (<65)", "Valor": sum(1 for c in criterios_sociales_rows if c["Criterio_Social_Asignado"] == "Cesante"), "Categoria": "Tabla 5 MINVU"},
        {"Indicador": "Criterio Social: Jubilado (≥65)", "Valor": sum(1 for c in criterios_sociales_rows if c["Criterio_Social_Asignado"] == "Jubilado"), "Categoria": "Tabla 5 MINVU"},
        {"Indicador": "Total Recintos Complementarios", "Valor": tot_rec_hab + tot_rec_nohab, "Categoria": "Infraestructura"},
        {"Indicador": "Recintos Complementarios HABITABLES", "Valor": tot_rec_hab, "Categoria": "Infraestructura"},
        {"Indicador": "Recintos Complementarios NO HABITABLES", "Valor": tot_rec_nohab, "Categoria": "Infraestructura"},
        {"Indicador": "Tabla 8: Fila 1 - Bodega", "Valor": sum(1 for r in recintos_rows if "Fila 1" in r["Fila_Formulario_Tabla8"]), "Categoria": "Tabla 8 MINVU"},
        {"Indicador": "Tabla 8: Fila 2 - Recinto Productivo", "Valor": sum(1 for r in recintos_rows if "Fila 2" in r["Fila_Formulario_Tabla8"]), "Categoria": "Tabla 8 MINVU"},
        {"Indicador": "Tabla 8: Fila 4 - Leñera", "Valor": sum(1 for r in recintos_rows if "Fila 4" in r["Fila_Formulario_Tabla8"]), "Categoria": "Tabla 8 MINVU"},
        {"Indicador": "Tabla 8: Fila 5 - Otros (Invernadero/Galpón)", "Valor": sum(1 for r in recintos_rows if "Fila 5" in r["Fila_Formulario_Tabla8"]), "Categoria": "Tabla 8 MINVU"},
        {"Indicador": "No Aplica Recinto Complementario", "Valor": tot_post - (tot_rec_hab + tot_rec_nohab), "Categoria": "Infraestructura"},
    ]
    df_kpi = pd.DataFrame(kpi_rows)

    return {
        "Hogares_Postulantes": df_hogares,
        "Habitantes_Detalle": df_habitantes,
        "Recintos_Complementarios": df_recintos,
        "Criterios_Sociales_Tabla5": df_criterios_sociales,
        "Resumen_KPIs": df_kpi
    }


def export_to_powerbi_excel(
    postulantes: List[Dict[str, Any]],
    output_path: str,
    datos_terreno: Optional[Dict[str, str]] = None
) -> str:
    """
    Genera un archivo Excel con múltiples hojas estructuradas y tipadas para Power BI.
    """
    datasets = build_powerbi_dataset(postulantes, datos_terreno)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in datasets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output_path

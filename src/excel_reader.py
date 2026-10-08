"""
Módulo para leer y normalizar la información de los postulantes
desde cualquier archivo Excel (.xlsx).
Diseñado para admitir planillas nuevas o actualizadas con mapeo dinámico de columnas.
Principio: Fidelidad estricta a la fuente. No inventar datos.
"""

import openpyxl
import datetime
import unicodedata
import re
from typing import List, Dict, Any, Optional


def norm_str(s: Any) -> str:
    """Normaliza un texto quitando tildes, espacios y caracteres especiales."""
    if not s:
        return ""
    text = str(s).strip()
    norm = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    return re.sub(r'[^a-zA-Z0-9]', '', norm).lower()


def format_rut(rut_val: Any, dv_val: Any = None) -> str:
    """Formatea un RUT al estándar chileno XX.XXX.XXX-X si es posible."""
    if not rut_val:
        return ""
    rut_str = str(rut_val).strip()
    if not rut_str or rut_str.lower() in ["none", "nan", "s/i", "0"]:
        return ""
    
    # Si viene con DV separado
    if dv_val is not None and str(dv_val).strip() and str(dv_val).strip().lower() not in ["none", "nan"]:
        dv = str(dv_val).strip().upper()
        clean_num = ''.join(filter(str.isdigit, rut_str))
        if clean_num:
            rev = clean_num[::-1]
            parts = [rev[i:i+3] for i in range(0, len(rev), 3)]
            formatted_num = '.'.join(parts)[::-1]
            return f"{formatted_num}-{dv}"
    
    # Si ya viene con guión
    if "-" in rut_str:
        num, dv = rut_str.rsplit("-", 1)
        num_clean = ''.join(filter(str.isdigit, num))
        dv_clean = dv.strip().upper()
        if num_clean:
            rev = num_clean[::-1]
            parts = [rev[i:i+3] for i in range(0, len(rev), 3)]
            formatted_num = '.'.join(parts)[::-1]
            return f"{formatted_num}-{dv_clean}"
        return rut_str.strip()
    
    return rut_str.strip()


def format_date(dt_val: Any) -> str:
    """Formatea una fecha como DD/MM/AAAA."""
    if not dt_val:
        return ""
    if isinstance(dt_val, (datetime.datetime, datetime.date)):
        return dt_val.strftime("%d/%m/%Y")
    val_str = str(dt_val).strip()
    if val_str.lower() in ["none", "nan", "00:00:00"]:
        return ""
    # Si viene como YYYY-MM-DD
    if "-" in val_str and len(val_str) >= 10:
        parts = val_str[:10].split("-")
        if len(parts) == 3 and len(parts[0]) == 4:
            return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return val_str


def build_column_mapping(sheet) -> Dict[str, int]:
    """
    Construye un mapa de columnas por nombre normalizado.
    Permite encontrar columnas aunque cambien de posición en planillas nuevas.
    """
    mapping = {}
    for c in range(1, sheet.max_column + 1):
        header_val = sheet.cell(1, c).value
        if header_val is not None:
            key = norm_str(header_val)
            if key and key not in mapping:
                mapping[key] = c
    return mapping


def get_col(mapping: Dict[str, int], possible_keys: List[str], default_col: int) -> int:
    """Busca la columna por alias normalizados, con fallback a default_col."""
    for k in possible_keys:
        norm_k = norm_str(k)
        if norm_k in mapping:
            return mapping[norm_k]
    return default_col


def parse_parientes_dynamic(sheet, row_idx: int, col_map: Dict[str, int]) -> List[Dict[str, Any]]:
    """Extrae los parientes usando mapeo dinámico de encabezados o índices estándar."""
    parientes = []
    for p_i in range(1, 7):
        c_nom = get_col(col_map, [f"PARIENTE {p_i}", f"PARIENTE{p_i}"], 60 + (p_i - 1) * 7)
        c_rut = get_col(col_map, [f"RUT {p_i}", f"RUT{p_i}"], c_nom + 1)
        c_ec = get_col(col_map, [f"ESTADO CIVIL {p_i}", f"ESTADO CIVIL"], c_nom + 2)
        c_fn = get_col(col_map, [f"FEC NAC {p_i}", f"FEC NAC{p_i}", f"FECHA NACIMIENTO {p_i}"], c_nom + 3)
        c_ed = get_col(col_map, [f"EDAD {p_i}", f"EDAD{p_i}"], c_nom + 4)
        c_par = get_col(col_map, [f"PARENTEZCO {p_i}", f"PARENTESCO {p_i}", f"PARENTEZCO", f"PARENTESCO"], c_nom + 5)
        c_disc = get_col(col_map, [f"DISCAPACIDAD {p_i}", f"DISCAPACIDAD{p_i}"], c_nom + 6)

        nombre = sheet.cell(row_idx, c_nom).value
        if not nombre or not str(nombre).strip():
            continue
        
        nombre_str = str(nombre).strip()
        rut_val = sheet.cell(row_idx, c_rut).value
        ecivil_val = sheet.cell(row_idx, c_ec).value
        fec_nac_val = sheet.cell(row_idx, c_fn).value
        edad_val = sheet.cell(row_idx, c_ed).value
        parentesco_val = sheet.cell(row_idx, c_par).value
        disc_val = sheet.cell(row_idx, c_disc).value

        edad_int = None
        if edad_val is not None:
            try:
                edad_int = int(float(str(edad_val).strip()))
            except (ValueError, TypeError):
                edad_int = None

        parientes.append({
            "indice": p_i,
            "nombre": nombre_str,
            "rut": format_rut(rut_val),
            "estado_civil": str(ecivil_val).strip() if ecivil_val else "",
            "fecha_nacimiento": format_date(fec_nac_val),
            "edad": edad_int,
            "parentesco": str(parentesco_val).strip().upper() if parentesco_val else "",
            "discapacidad": str(disc_val).strip() if disc_val else "NO"
        })
    return parientes


def find_best_sheet(wb: openpyxl.Workbook, preferred_name: Optional[str] = None) -> openpyxl.worksheet.worksheet.Worksheet:
    """Encuentra la hoja adecuada (por nombre preferido o buscando la que tenga 'NOMBRE' y 'RUT')."""
    if preferred_name and preferred_name in wb.sheetnames:
        return wb[preferred_name]
    
    if "base" in wb.sheetnames:
        return wb["base"]
    
    for sname in wb.sheetnames:
        sheet = wb[sname]
        headers = [norm_str(sheet.cell(1, c).value) for c in range(1, min(15, sheet.max_column + 1))]
        if "nombre" in headers and "rut" in headers:
            return sheet
            
    return wb.active


def load_actividades_economicas(wb: openpyxl.Workbook) -> Dict[str, Dict[str, Any]]:
    """
    Carga la hoja de 'Actividad Económica' si existe en el libro Excel.
    Crea un diccionario indexado por el número limpio de RUT (sin puntos ni guión)
    y también por nombre normalizado (para casos con errores tipográficos en el RUT).
    Principio: Fidelidad estricta. Si no hay actividad, queda en blanco.
    """
    ws_act = None
    for name in wb.sheetnames:
        norm_name = norm_str(name)
        if "actividad" in norm_name or "economica" in norm_name or "actividadeconomica" in norm_name:
            ws_act = wb[name]
            break
            
    if not ws_act:
        return {}

    col_map = build_column_mapping(ws_act)
    c_nom = get_col(col_map, ["NOMBRE", "POSTULANTE"], 1)
    c_rut = get_col(col_map, ["RUT"], 2)
    c_act = get_col(col_map, ["ACTIVIDAD ECONOMICA PRINCIPAL", "ACTIVIDAD ECONOMICA", "ACTIVIDAD"], 3)
    c_cert = get_col(col_map, ["CERTIFICADO/INICIO ACTIVIDADES", "CERTIFICADO"], 4)
    c_desc = get_col(col_map, ["DESCRIPCION DE LA ACTIVIDAD", "DESCRIPCION"], 5)
    c_jefe = get_col(col_map, ["JEFE DE HOGAR", "JEFE"], 6)
    c_cony = get_col(col_map, ["CONYUGE/PAREJA", "CONYUGE"], 7)
    c_otros = get_col(col_map, ["OTROS"], 8)

    records = {}

    def clean_rut_key(v: Any) -> str:
        if not v:
            return ""
        s = str(v).strip()
        if "-" in s:
            s = s.split("-")[0]
        return ''.join(filter(str.isdigit, s))

    for r in range(2, ws_act.max_row + 1):
        nom_val = ws_act.cell(r, c_nom).value
        rut_val = ws_act.cell(r, c_rut).value
        act_val = ws_act.cell(r, c_act).value
        desc_val = ws_act.cell(r, c_desc).value
        jefe_val = ws_act.cell(r, c_jefe).value
        cony_val = ws_act.cell(r, c_cony).value
        otros_val = ws_act.cell(r, c_otros).value

        if not nom_val:
            continue

        act_clean = str(act_val).strip() if act_val and str(act_val).strip().lower() not in ["none", "nan"] else ""
        desc_clean = str(desc_val).strip() if desc_val and str(desc_val).strip().lower() not in ["none", "nan"] else ""
        jefe_clean = str(jefe_val).strip().upper() if jefe_val and str(jefe_val).strip().lower() not in ["none", "nan"] else ""
        cony_clean = str(cony_val).strip().upper() if cony_val and str(cony_val).strip().lower() not in ["none", "nan"] else ""
        otros_clean = str(otros_val).strip().upper() if otros_val and str(otros_val).strip().lower() not in ["none", "nan"] else ""

        # Solo se considera que tiene actividad si viene una actividad o descripción real
        tiene_act = bool(act_clean or desc_clean)

        entry = {
            "actividad_principal": act_clean,
            "descripcion": desc_clean,
            "jefe_hogar": jefe_clean,
            "conyuge_pareja": cony_clean,
            "otros": otros_clean,
            "tiene_actividad": tiene_act
        }

        rut_k = clean_rut_key(rut_val)
        if rut_k:
            records[f"rut_{rut_k}"] = entry
        
        name_k = norm_str(nom_val)
        if name_k:
            records[f"nom_{name_k}"] = entry

    return records


def read_all_postulantes(excel_path: str, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Lee todos los postulantes válidos de la planilla Excel con mapeo dinámico.
    Soporta planillas actualizadas con columnas reordenadas o adicionales.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = find_best_sheet(wb, sheet_name)
    col_map = build_column_mapping(sheet)
    act_map = load_actividades_economicas(wb)

    # Resolución dinámica de columnas principales
    col_orden = get_col(col_map, ["N°", "N° ", "N", "ORDEN"], 1)
    col_nombre = get_col(col_map, ["NOMBRE", "POSTULANTE", "NOMBRE COMPLETO"], 2)
    col_rut = get_col(col_map, ["RUT", "RUT POSTULANTE"], 3)
    col_dv = get_col(col_map, ["DV", "DIGITO"], 4)
    col_fono = get_col(col_map, ["FONO", "TELEFONO", "CELULAR"], 5)
    col_dir_rsh = get_col(col_map, ["DIRECCION RSH", "DIRECCION", "DOMICILIO"], 6)
    col_rol_sii = get_col(col_map, ["ROL SII", "ROL", "ROL PROPIEDAD"], 7)
    col_dir_terreno = get_col(col_map, ["DIRECCION TERRENO", "TERRENO"], 8)
    col_etnia = get_col(col_map, ["ETNIA", "PUEBLO INDIGENA"], 14)
    col_sexo = get_col(col_map, ["SEXO", "GENERO"], 15)
    col_ecivil = get_col(col_map, ["ESTADO CIVIL", "ESTADO_CIVIL"], 16)
    col_c_nombre = get_col(col_map, ["NOMBRE CONYUGE", "CONYUGE"], 17)
    col_c_rut = get_col(col_map, ["RUT CONYUGE"], 18)
    col_c_fec_nac = get_col(col_map, ["FECHA DE NACIMIENTO CONYUGE", "FEC NAC CONYUGE"], 19)
    col_c_sexo = get_col(col_map, ["SEXO CONYUGE"], 20)
    col_pais = get_col(col_map, ["PAIS", "NACIONALIDAD"], 21)
    col_fec_nac = get_col(col_map, ["FECHA DE NACIMIENTO", "FEC NAC"], 22)
    col_edad = get_col(col_map, ["EDAD HOY", "EDAD"], 23)
    col_disc = get_col(col_map, ["DISCAPACIDAD"], 24)
    col_comuna = get_col(col_map, ["COMUNA"], 36)
    col_gf_cant = get_col(col_map, ["GRUPO FAMILIAR", "TOTAL HABITANTES"], 37)
    col_tipo_viv = get_col(col_map, ["TIPO VIVIENDA", "TIPO DE VIVIENDA"], 48)
    col_factor = get_col(col_map, ["FACTOR AISLAMIENTO", "AISLAMIENTO"], 59)
    col_tipologia = get_col(col_map, ["TIPOLOGÍA PROPUESTA", "TIPOLOGIA"], 103)
    col_fundamento = get_col(col_map, ["FUNDAMENTO / OBSERVACIÓN", "FUNDAMENTO", "OBSERVACION"], 104)

    postulantes = []
    
    for r in range(2, sheet.max_row + 1):
        nombre_val = sheet.cell(r, col_nombre).value
        if not nombre_val or not str(nombre_val).strip():
            continue
        
        nro_orden = sheet.cell(r, col_orden).value
        rut_num = sheet.cell(r, col_rut).value
        dv_val = sheet.cell(r, col_dv).value
        fono = sheet.cell(r, col_fono).value
        dir_rsh = sheet.cell(r, col_dir_rsh).value
        rol_sii = sheet.cell(r, col_rol_sii).value
        dir_terreno = sheet.cell(r, col_dir_terreno).value
        etnia = sheet.cell(r, col_etnia).value
        sexo = sheet.cell(r, col_sexo).value
        estado_civil = sheet.cell(r, col_ecivil).value
        raw_c_nom = sheet.cell(r, col_c_nombre).value
        raw_c_rut = sheet.cell(r, col_c_rut).value
        raw_c_fn = sheet.cell(r, col_c_fec_nac).value
        raw_c_sex = sheet.cell(r, col_c_sexo).value
        pais = sheet.cell(r, col_pais).value
        fecha_nac = sheet.cell(r, col_fec_nac).value
        edad = sheet.cell(r, col_edad).value
        discapacidad = sheet.cell(r, col_disc).value
        comuna = sheet.cell(r, col_comuna).value
        grupo_fam_cant = sheet.cell(r, col_gf_cant).value
        raw_tipo_viv = sheet.cell(r, col_tipo_viv).value
        factor_aislamiento = sheet.cell(r, col_factor).value
        tipologia = sheet.cell(r, col_tipologia).value
        fundamento = sheet.cell(r, col_fundamento).value

        # Normalizar edad del titular
        edad_titular = None
        if edad is not None:
            try:
                edad_titular = int(float(str(edad).strip()))
            except (ValueError, TypeError):
                edad_titular = None

        # Normalizar cantidad de grupo familiar
        gf_num = None
        if grupo_fam_cant is not None:
            try:
                gf_num = int(float(str(grupo_fam_cant).strip()))
            except (ValueError, TypeError):
                gf_num = None

        parientes = parse_parientes_dynamic(sheet, r, col_map)

        if gf_num is None or gf_num <= 0:
            gf_num = 1 + len(parientes)

        # Consolidar cónyuge / pareja
        conyuge_nombre_final = str(raw_c_nom).strip() if raw_c_nom else ""
        conyuge_rut_final = format_rut(raw_c_rut)
        conyuge_sexo_final = str(raw_c_sex).strip().upper() if raw_c_sex else ""
        conyuge_fn_final = format_date(raw_c_fn)

        if not conyuge_nombre_final:
            for p in parientes:
                if any(w in p["parentesco"] for w in ["CONYUGE", "CÓNYUGE", "CNYUGE", "PAREJA", "CONVIVIENTE"]):
                    conyuge_nombre_final = p["nombre"]
                    if not conyuge_rut_final:
                        conyuge_rut_final = p["rut"]
                    if not conyuge_fn_final:
                        conyuge_fn_final = p["fecha_nacimiento"]
                    break

        def clean_s(val: Any) -> str:
            if val is None:
                return ""
            s = str(val).strip()
            return "" if s.lower() in ["none", "nan"] else s

        # Buscar actividad económica fidedigna en la hoja especializada
        rut_clean_digits = ''.join(filter(str.isdigit, str(rut_num or "")))
        act_info = act_map.get(f"rut_{rut_clean_digits}") or act_map.get(f"nom_{norm_str(nombre_val)}") or {
            "actividad_principal": "",
            "descripcion": "",
            "jefe_hogar": "",
            "conyuge_pareja": "",
            "otros": "",
            "tiene_actividad": False
        }

        postulantes.append({
            "fila_excel": r,
            "nro_orden": int(nro_orden) if nro_orden and str(nro_orden).isdigit() else r - 1,
            "nombre": clean_s(nombre_val),
            "rut": format_rut(rut_num, dv_val),
            "telefono": clean_s(fono),
            "fecha_nacimiento": format_date(fecha_nac),
            "edad": edad_titular,
            "sexo": clean_s(sexo).upper(),
            "estado_civil": clean_s(estado_civil).upper(),
            "etnia": clean_s(etnia).upper(),
            "pais": clean_s(pais).upper() if pais else "CHILE",
            "discapacidad": clean_s(discapacidad),
            "direccion_rsh": clean_s(dir_rsh),
            "direccion_terreno": clean_s(dir_terreno),
            "rol_sii": clean_s(rol_sii),
            "comuna": clean_s(comuna) or "PERQUENCO",
            "provincia": "Cautín",
            "region": "La Araucanía",
            "factor_aislamiento": str(factor_aislamiento).strip() if factor_aislamiento is not None else "",
            "grupo_familiar_cant": gf_num,
            "conyuge": {
                "nombre": conyuge_nombre_final,
                "rut": conyuge_rut_final,
                "sexo": conyuge_sexo_final,
                "fecha_nacimiento": conyuge_fn_final
            },
            "parientes": parientes,
            "actividad_economica": act_info,
            "tipo_vivienda": clean_s(raw_tipo_viv),
            "tipologia_propuesta": clean_s(tipologia),
            "fundamento": clean_s(fundamento)
        })

    return postulantes

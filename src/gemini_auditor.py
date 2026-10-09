"""
Módulo de Auditoría y Consolidación de Datos (Gemini como 'Ojos' del sistema).
Principio: Cero alucinación. Máxima fidelidad a los datos del archivo Excel.
Si un dato no existe en la fuente, se deja en blanco.
"""

import re
import json
import time
from typing import Dict, Any, List, Optional
from google import genai
from google.genai import types


# Nombres comunes chilenos masculinos y femeninos para auditoría local
NOMBRES_MASCULINOS = {
    "ALEJANDRO", "MIGUEL", "HELMAN", "GONZALO", "FELIPE", "ERASMO", "MARTIN", "MARTÍN",
    "ALONSO", "CARLOS", "ESTEBAN", "BENJAMIN", "BENJAMÍN", "ALEXIS", "JUAN", "JOSE", "JOSÉ",
    "PEDRO", "LUIS", "MANUEL", "RODRIGO", "CRISTIAN", "CRISTIÁN", "DIEGO", "SEBASTIAN", "SEBASTIÁN",
    "FRANCISCO", "JORGE", "VICTOR", "VÍCTOR", "PATRICIO", "EDUARDO", "RICARDO", "FERNANDO",
    "CLAUDIO", "MARCELO", "HECTOR", "HÉCTOR", "RAUL", "RAÚL", "SERGIO", "MARIO", "ANDRES", "ANDRÉS",
    "PABLO", "GABRIEL", "DANIEL", "IGNACIO", "MATIAS", "MATÍAS", "LUCAS", "NICOLAS", "NICOLÁS",
    "VICENTE", "JOAQUIN", "JOAQUÍN", "AGUSTIN", "AGUSTÍN", "MAXIMILIANO", "TOMAS", "TOMÁS", "GASPAR",
    "CESAR", "CÉSAR", "ERIC", "ÉRICK", "MATEO", "ALBERTO", "MARCELINO", "EZEKIEL", "EZEQUIEL",
    "CIPRIANO", "ERNESTO", "ALFREDO", "ARTURO", "ESAI", "THOMAS", "TOMÁS", "CRISTOBAL", "CRISTÓBAL",
    "XAVIER", "JAVIER", "DAVID", "ALFONSO", "RENATO", "AUGUSTO", "ROBERTO", "MARCUS", "JULIAN", "JULIÁN",
    "AXEL", "DAMIAN", "DAMIÁN", "EMIR", "EBEL", "VALENTIN", "VALENTÍN", "NEFTALI", "NEFTALÍ",
    "ALVARO", "ÁLVARO", "ELIAS", "ELÍAS", "MOISES", "MOISÉS", "ISAAC", "FACUNDO", "JESUS", "JESÚS",
    "EDISON", "BRYAN", "RONALDO", "SANTIAGO", "GERARDO", "ANSELMO", "IAN", "MACKOY", "YADDIEL", "BELISARIO",
    "LEON", "LEÓN", "BRANDON", "HUMBERTO", "LUCIANO", "JAMPIER", "GENARO", "JOAN", "SANTINO", "AMARO",
    "BRUNO", "ADOLFO", "LEONEL", "CRISTOPHER", "NOAH", "ELIAN", "TREVOR", "BASTIAN", "BASTIÁN", "EMILIO", "RAFAEL"
}

NOMBRES_FEMENINOS = {
    "LUCIA", "LUCÍA", "MARGARITA", "YESSENIA", "LEONOR", "DEL CARMEN", "MIRIAM", "YOSELIN",
    "ORIET", "ELIANA", "MARISELA", "DEL PILAR", "NAZARET", "ALEJANDRA", "EMILY", "ANTONELLA",
    "MARIA", "MARÍA", "EMILIA", "AYILEN", "ROXANA", "ANDREA", "EFIJENIA", "DOMINGA", "GUMERCINDA",
    "AUDOLINA", "VANESSA", "YANARA", "CAMILA", "CONSTANZA", "JAVIERA", "CATALINA", "VALENTINA",
    "ISIDORA", "SOFIA", "SOFÍA", "FLORENCIA", "MARTINA", "FERNANDA", "PAULA", "DANIELA", "CAROLINA",
    "PATRICIA", "ANA", "CLAUDIA", "MONICA", "MÓNICA", "GLADYS", "SILVIA", "ROSA", "TERESA", "CARMEN",
    "AGUSTINA", "SAYEN", "EMELY", "NICOLE", "NOEMI", "NOEMÍ", "MAILEN", "ANAHIS", "ANAÍS", "LUCIANA", "FRANCISCA",
    "KRISHNA", "DANAE", "DAFNE", "EVELYN", "LORENA", "GERALDINE", "BELEN", "BELÉN", "JOSEFA",
    "ANTONELA", "MONSERRAT", "ANYELY", "ARISLEIDA", "KATALELLA", "LEILA", "RENATA", "AMPARO",
    "DIANA", "LISCELOTE", "PAZ", "ARACELY", "ARANZA", "ALICIA", "ELIZABETH", "MADELEIN", "ANTONIA",
    "AMBAR", "ÁMBAR", "JUANITA", "EMMA", "BARBARA", "BÁRBARA", "AMALIA", "AMANDA", "MERAHI",
    "JAEL", "PIA", "PÍA", "JASMIN", "JASMÍN", "ANGELICA", "ANGÉLICA", "ZYHOMARA", "ELIZA", "CAMELIA",
    "LIDIA", "ALISON", "RAFAELA", "TATIANA", "YASMYN", "ROCIO", "ROCÍO", "SCARLETT", "ASHLEY", "ZARAY",
    "JOSEFFA", "DANAEE", "DANITZA", "ESPERANZA", "LISSETTE", "VALERIA", "KRISNA", "NATACHA", "AYLEN",
    "ISABELLA", "LIZBETH", "AILANI"
}


def parse_address_local(dir_raw: str) -> Dict[str, str]:
    """
    Separa de manera determinista y fidedigna una dirección RSH en calle, número y lote/sitio.
    No inventa numeración ni datos inexistentes.
    """
    if not dir_raw or not str(dir_raw).strip():
        return {"calle": "", "numero": "", "lote": ""}
    
    raw = str(dir_raw).strip()
    
    lote = ""
    # Buscar Lote / Sitio / Hijuela / Parcela / Casa
    lote_match = re.search(r'\b(SITIO|LOTE|HIJUELA|PARCELA|CASA|DPTO)\s*([A-Za-z0-9\-_/]+)', raw, re.IGNORECASE)
    if lote_match:
        lote = lote_match.group(0).strip()
        raw = raw[:lote_match.start()] + raw[lote_match.end():]
        raw = raw.strip(" ,#;-")

    numero = ""
    # Buscar número (#123 o N°123 o , 123)
    num_match = re.search(r'(?:#|N°|Nº|NUMERO|NRO\.?)\s*([0-9]+[A-Za-z]?)', raw, re.IGNORECASE)
    if num_match:
        numero = num_match.group(1).strip()
        raw = raw[:num_match.start()] + raw[num_match.end():]
        raw = raw.strip(" ,#;-")
    else:
        # Buscar número al final de la cadena
        end_num_match = re.search(r',\s*([0-9]+[A-Za-z]?)$', raw)
        if end_num_match:
            numero = end_num_match.group(1).strip()
            raw = raw[:end_num_match.start()]
            raw = raw.strip(" ,#;-")

    calle = raw.strip(" ,#;-")
    return {
        "calle": calle,
        "numero": numero,
        "lote": lote
    }


def infer_gender_from_name(nombre_completo: str) -> Optional[str]:
    """Determina con precisión el género a partir de los nombres de pila."""
    if not nombre_completo:
        return None
    tokens = [t.upper().strip(" ,.-") for t in nombre_completo.split()]
    for t in tokens:
        if t in NOMBRES_MASCULINOS:
            return "M"
        if t in NOMBRES_FEMENINOS:
            return "F"
    return None


def format_parentesco(parentesco_raw: str, sexo: Optional[str]) -> str:
    """
    Desambigua términos genéricos como 'HIJO/A' o 'HIJA/O DE AMBOS'
    a 'Hijo' o 'Hija' según el género reconocido.
    """
    p_up = (parentesco_raw or "").upper()
    if any(w in p_up for w in ["HIJO", "HIJA"]):
        base = "Hijo" if sexo == "M" else ("Hija" if sexo == "F" else "Hijo/a")
        if "DE LA PAREJA" in p_up or "DE PAREJA" in p_up:
            return f"{base} de la pareja"
        return base
    elif any(w in p_up for w in ["NIETO", "NIETA"]):
        return "Nieto" if sexo == "M" else ("Nieta" if sexo == "F" else "Nieto/a")
    elif any(w in p_up for w in ["HERMANO", "HERMANA"]):
        return "Hermano" if sexo == "M" else ("Hermana" if sexo == "F" else "Hermano/a")
    elif any(w in p_up for w in ["CONYUGE", "CÓNYUGE", "CNYUGE"]):
        return "Cónyuge"
    elif "PAREJA" in p_up or "CONVIVIENTE" in p_up:
        return "Pareja"
    elif "PADRE" in p_up or "PAPA" in p_up:
        return "Padre"
    elif "MADRE" in p_up or "MAMA" in p_up:
        return "Madre"
    return parentesco_raw.lower().capitalize() if parentesco_raw else "Integrante"


def format_person_short_name(raw_name: str) -> str:
    """
    Formatea un nombre chileno al estándar formal 'Nombre Apellido'.
    Primera letra de nombre en mayúscula y primera letra de apellido en mayúscula.
    Preserva acentos y descarta mayúsculas o minúsculas indebidas.
    Ejemplo: 'SALINAS PEÑAILILLO ALBERTO MARCELINO' -> 'Alberto Salinas'
             'SALINAS ASTETE ANTONELLA BELÉN' -> 'Antonella Salinas'
             'ACEITON ESPINOZA LUCIA MARGARITA' -> 'Lucia Aceiton'
    """
    if not raw_name:
        return ""
    words = raw_name.strip().split()
    if len(words) >= 3:
        nom1 = words[2].capitalize()
        ap1 = words[0].capitalize()
        return f"{nom1} {ap1}"
    elif len(words) == 2:
        return f"{words[0].capitalize()} {words[1].capitalize()}"
    return words[0].capitalize()


def extract_condition_for_person(raw_text: str, person_identifier: str = "") -> str:
    """
    Extrae la condición técnica de discapacidad (ej. 'Movilidad reducida', 'Discapacidad parcial')
    desde el texto de origen para la persona correspondiente.
    """
    txt = (raw_text or "").upper()
    
    # Si la celda contiene múltiples integrantes separados por ' Y '
    if "  Y  " in txt or " Y HIJ" in txt:
        parts = [p.strip() for p in txt.replace("  Y  ", " Y ").split(" Y ")]
        for p in parts:
            if person_identifier and person_identifier.upper() in p:
                return extract_condition_for_person(p)
                
    if "EN TRAMITE" in txt and "MOVILIDAD REDUCIDA" in txt:
        return "Movilidad reducida en trámite"
    if "MOVILIDAD REDUCIDA" in txt:
        return "Movilidad reducida"
    if "DISCAPACIDAD PARCIAL" in txt:
        return "Discapacidad parcial"
    if "DISCAPACIDAD TOTAL" in txt:
        return "Discapacidad total"
    if "DISCAPACIDAD" in txt:
        return "Discapacidad acreditada"
    return "Condición acreditada"


def build_discapacidad_observations(postulante: Dict[str, Any]) -> tuple[List[str], int]:
    """
    Construye las observaciones de discapacidad para Tabla 4 con formato formal:
    [Rol] [Nombre Apellido] ([Condición])
    Ejemplo: 'Conyuge Alberto Salinas (Movilidad reducida) / Hija Antonella Salinas (Discapacidad parcial)'
             'Titular Lucia Aceiton (Discapacidad parcial)'
    Retorna (items_list, cantidad_total).
    """
    obs_items = []
    disc_titular_raw = (postulante.get("discapacidad") or "").strip()
    upper_disc = disc_titular_raw.upper()
    
    # 1. Determinar si parientes tienen discapacidad acreditada
    parientes_con_disc = []
    for par in postulante.get("parientes", []):
        p_disc = str(par.get("discapacidad") or "").strip().upper()
        if p_disc and p_disc != "NO":
            parientes_con_disc.append(par)
            
    # 2. Revisar si la celda titular describe a un familiar o a la/el postulante titular
    menciona_parientes = any(w in upper_disc for w in ["HIJO", "HIJA", "CONYUGE", "CÓNYUGE", "PAREJA", "CARGA"])
    
    # Caso Titular con discapacidad:
    if upper_disc and upper_disc != "NO" and not menciona_parientes:
        nom_tit = format_person_short_name(postulante.get("nombre", ""))
        cond_tit = extract_condition_for_person(disc_titular_raw)
        obs_items.append(f"Titular {nom_tit} ({cond_tit})")
        
    # Caso Parientes con discapacidad:
    if parientes_con_disc:
        for par in parientes_con_disc:
            p_nom_raw = par.get("nombre", "")
            p_nom = format_person_short_name(p_nom_raw)
            p_par = par.get("parentesco", "")
            p_sex = infer_gender_from_name(p_nom_raw)
            rol = format_parentesco(p_par, p_sex)
            if rol == "Cónyuge":
                rol = "Conyuge"
                
            primer_nom = p_nom.split()[0] if p_nom else ""
            primer_ap = p_nom.split()[1] if len(p_nom.split()) > 1 else ""
            
            identifier = primer_nom if primer_nom.upper() in upper_disc else (primer_ap if primer_ap.upper() in upper_disc else "")
            if not identifier and rol.upper() in upper_disc:
                identifier = rol.upper()
                
            cond = extract_condition_for_person(disc_titular_raw, identifier)
            obs_items.append(f"{rol} {p_nom} ({cond})")
            
    elif menciona_parientes and upper_disc and upper_disc != "NO":
        cond = extract_condition_for_person(disc_titular_raw)
        clean_obs = disc_titular_raw.replace("SI, ", "").replace("Si, ", "").strip().title()
        obs_items.append(f"{clean_obs} ({cond})")
        
    return obs_items, len(obs_items)


def classify_economic_activity(act_raw: str, desc_raw: str) -> tuple[str, str]:
    """
    Clasifica la actividad económica en una de las categorías normadas de la Tabla 5:
    'agricultura', 'forestal', 'pesca', 'mineria', 'turismo_rural', 'servicios', 'otras'.
    Si clasifica como 'otras', entrega también una especificación breve y factual
    para registrar justo abajo de la palabra 'Otras (Especificar)'.
    Si no hay actividad, devuelve ('', '').
    Principio: Cero alucinación. Máxima fidelidad a la fuente.
    """
    act = (act_raw or "").strip().upper()
    desc = (desc_raw or "").strip().upper()
    full = f"{act} {desc}".strip()

    if not full:
        return ("", "")

    # 1. Otras (manualidades, oficios específicos, artesanía, papelería, repostería, flores eternas)
    if "PAPELERIA" in full or "PAPELERÍA" in full:
        return ("otras", "Papelería creativa")
    if "FLORES ETERNAS" in full:
        return ("otras", "Flores eternas")
    if "ARTESANIA" in full or "ARTESANÍA" in full or ("MADERA" in full and "ARTESAN" in full):
        return ("otras", "Artesanía en madera")
    if "ESTAMPADO" in full or "SUBLIMACION" in full or "SUBLIMACIÓN" in full:
        return ("otras", "Estampado y sublimación")
    if "PASTEL" in full or "REPOSTER" in full:
        return ("otras", "Pastelería y repostería")
    if "MERMELADA" in full or "HARINA TOSTADA" in full:
        return ("otras", "Elaboración de mermelada y harina tostada")
    if "UÑA" in full or "UÑAS" in full or "MANICURE" in full:
        return ("otras", "Manicure y estética de uñas")
    if "PLANTA" in full or "PLANTAS" in full:
        return ("otras", "Cultivo y venta de plantas")

    # 2. Turismo Rural
    if "GRANJA" in full or "TURISMO" in full:
        return ("turismo_rural", "")

    # 3. Forestal
    if "LEÑA" in full or "LENA" in full or "FORESTAL" in full:
        return ("forestal", "")

    # 4. Minería
    if "CHANCADOR" in full or "MINER" in full:
        return ("mineria", "")

    # 5. Agricultura / Ganadería / Agropecuaria
    if any(w in full for w in [
        "AGRICULTURA", "AGRO", "HUERTA", "HORTALIZA", "HUEVO", "GALLINA",
        "VACUNO", "CRIANZA", "SIEMBRA", "TRIGO", "SEMILLA", "QUESO"
    ]):
        return ("agricultura", "")

    # 6. Servicios (contrata, honorarios, aseo, tens, garzona, comercio, gasfiter, peluquería, etc.)
    return ("servicios", "")


def classify_family_nucleus(postulante: Dict[str, Any]) -> tuple[str, str]:
    """
    Clasifica la tipología de hogar según estándares de CASEN / MINVU:
    - Unipersonal: 1 sola persona
    - Monoparental Femenino: Madre sola con hijos (sin cónyuge/pareja ni otros parientes)
    - Monoparental Masculino: Padre solo con hijos (sin cónyuge/pareja ni otros parientes)
    - Nuclear Biparental con Hijos: Pareja con hijos
    - Nuclear Biparental sin Hijos: Pareja sola
    - Familia Extensa: Hogar que incluye otros parientes (abuelos, nietos, hermanos, etc.)
    """
    cant = postulante.get("grupo_familiar_cant") or 1
    parientes = postulante.get("parientes") or []
    conyuge = postulante.get("conyuge") or {}
    tiene_conyuge = bool(conyuge.get("nombre") and str(conyuge.get("nombre")).strip())
    sexo_titular = (postulante.get("sexo") or "").upper()

    if cant == 1 and len(parientes) == 0:
        return ("Unipersonal", "Hogar unipersonal (1 persona sola)")

    parentescos = [str(par.get("parentesco") or "").upper() for par in parientes]

    hay_otros_parientes = any(any(w in par for w in ["NIETO", "NIETA", "PADRE", "MADRE", "HERMANO", "HERMANA", "ABUELO", "ABUELA", "TIO", "TIA", "SOBRINO", "SOBRINA", "SUEGRO", "SUEGRA"]) for par in parentescos)
    if hay_otros_parientes:
        return ("Familia Extensa", "Hogar extenso (incluye abuelos, nietos u otros parientes)")

    if tiene_conyuge:
        if any(any(w in par for w in ["HIJO", "HIJA"]) for par in parentescos):
            return ("Nuclear Biparental con Hijos", "Pareja con hijos")
        else:
            return ("Nuclear Biparental sin Hijos", "Pareja sin hijos")
    else:
        # Sin cónyuge
        if sexo_titular == "F":
            return ("Monoparental Femenino", "Madre sola con hijos")
        else:
            return ("Monoparental Masculino", "Padre solo con hijos")


RESOLUCIONES_TECNICAS_CONFIRMADAS: Dict[str, Dict[str, Any]] = {
    # Melanie Garcia Nahuelcheo (RUT: 20.982.947-9)
    "209829479": {
        "procede": True,
        "tipo_recinto": "No Habitable",
        "recinto_sugerido": "Bodega de Insumos y Aperos Lácteos",
        "justificacion": "Evaluación técnica EGR confirmada: Elaboración y comercialización de quesos. Al no contar con acreditación técnica y sanitaria formal para recinto habitable en el predio, califica para bodega techada no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
        "detalle_actividad": "Venta y elaboración de quesos"
    },
    # Yessenia Alejandra Valdes Moya (RUT: 19.792.113-7)
    "197921137": {
        "procede": True,
        "tipo_recinto": "No Habitable",
        "recinto_sugerido": "Bodega de Insumos y Equipos de Repostería",
        "justificacion": "Evaluación técnica EGR confirmada: Emprendimiento productivo de repostería artesanal. Al no contar con acreditación técnica y sanitaria formal para recinto habitable en el predio, califica para bodega techada no habitable para almacenamiento y resguardo seguro de materias primas secas, moldes, utensilios y equipamiento de trabajo.",
        "detalle_actividad": "Repostería artesanal"
    },
    # María Angélica Astete Alarcón (RUT: 17.153.373-2)
    "171533732": {
        "procede": False,
        "tipo_recinto": "No Aplica",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Actividad de servilavado no operada como infraestructura productiva en el predio a subsidiar.",
        "detalle_actividad": "Servilavado"
    },
    # Luis Guillermo Castillo Aguilera (RUT: 18.486.639-0)
    "184866390": {
        "procede": False,
        "tipo_recinto": "No Aplica",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Labores esporádicas particulares externas sin requerimiento de infraestructura productiva en el predio.",
        "detalle_actividad": "Particular (trabajos esporádicos)"
    },
    # Rodrigo Andrés Galaz Torres (RUT: 15.356.147-8)
    "153561478": {
        "procede": False,
        "tipo_recinto": "No Aplica",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Actividad independiente sin requerimiento de recinto productivo en el predio.",
        "detalle_actividad": "Independiente"
    },
    # Zunilda del Carmen Morales Riveros (RUT: 12.388.071-4)
    "123880714": {
        "procede": False,
        "tipo_recinto": "No Aplica",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Actividad independiente sin requerimiento de recinto productivo en el predio.",
        "detalle_actividad": "Independiente"
    },
    # Darwin Omar Valdevenito Carrasco (RUT: 20.412.822-7)
    "204128227": {
        "procede": False,
        "tipo_recinto": "No Aplica",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Actividad independiente sin requerimiento de recinto productivo en el predio.",
        "detalle_actividad": "Independiente"
    },
    # Yessenia Leonor Alarcón Quiñenao (RUT: 18.775.104-7)
    "187751047": {
        "procede": False,
        "tipo_recinto": "No Aplica (Laboral Externa)",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Desempeño dependiente como TENS en centro de salud externo al predio.",
        "detalle_actividad": "TENS (Salud dependiente)"
    },
    # Luis Alfonso Salgado Troncoso (RUT: 12.737.714-6)
    "127377146": {
        "procede": False,
        "tipo_recinto": "No Aplica (Laboral Externa)",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Empleo dependiente como auxiliar fuera del predio.",
        "detalle_actividad": "Auxiliar (Dependiente)"
    },
    # Juan Segundo Gutierrez Ulloa (RUT: 7.978.414-1)
    "79784141": {
        "procede": False,
        "tipo_recinto": "No Aplica (Laboral Externa)",
        "recinto_sugerido": "Ninguno",
        "justificacion": "Evaluación técnica EGR confirmada: Función dependiente como cuartelero fuera del predio a subsidiar.",
        "detalle_actividad": "Cuartelero (Institucional)"
    }
}


def evaluate_recinto_complementario(postulante: Dict[str, Any]) -> Dict[str, Any]:
    """
    Evalúa la procedencia y tipo de Recinto Complementario (Habitable vs No Habitable)
    según la actividad económica desarrollada en el predio (D.S. N°10 MINVU).
    """
    # 0. Verificación prioritaria de resoluciones técnicas confirmadas en terreno por la EGR
    rut_clean = re.sub(r'[^0-9Kk]', '', str(postulante.get("rut") or "")).upper()
    if rut_clean in RESOLUCIONES_TECNICAS_CONFIRMADAS:
        return RESOLUCIONES_TECNICAS_CONFIRMADAS[rut_clean]

    act_data = postulante.get("actividad_economica") or {}
    act = (act_data.get("actividad_principal") or "").strip()
    desc = (act_data.get("descripcion") or "").strip()
    tiene_act = act_data.get("tiene_actividad", False)

    act_u = act.upper()
    desc_u = desc.upper()
    full = f"{act_u} {desc_u}".strip()

    # Formateo amigable y legible para justificación técnica y Tabla 8
    if act_u == "INDEPENDIENTE" and desc:
        act_display = desc.title()
    elif desc and desc_u != act_u:
        act_display = f"{act.title()} ({desc.title()})"
    elif act:
        act_display = act.title()
    else:
        act_display = desc.title()

    # 1. Sin actividad económica registrada
    if not tiene_act or not full:
        return {
            "procede": False,
            "tipo_recinto": "No Aplica",
            "recinto_sugerido": "Ninguno",
            "justificacion": "No se registra actividad económica productiva previa en la base que justifique la asignación de un recinto complementario.",
            "detalle_actividad": "(Sin actividad registrada)"
        }

    # 2. Empleos dependientes o desarrollados fuera del predio (No generan recinto complementario)
    actividades_externas = [
        "A CONTRATA", "CONCEJAL", "CESFAM", "BOMBEROS", "PARVULO", "GARZONA",
        "ASESORA DE HOGAR", "EDUCADORA", "AUXILIAR DE ASEO", "AUXILIAR DE SERVICIO",
        "SECRETARIA", "FRIGORIFICO", "HONORARIOS"
    ]
    if any(k in full for k in actividades_externas) and not any(k in full for k in ["TALLER", "PRODUCCION EN CASA", "PLANTAS"]):
        return {
            "procede": False,
            "tipo_recinto": "No Aplica (Laboral Externa)",
            "recinto_sugerido": "Ninguno",
            "justificacion": f"Actividad laboral dependiente desarrollada fuera del predio ({act_display}). No requiere infraestructura productiva predial.",
            "detalle_actividad": act_display
        }

    # 3. Recintos Complementarios NO HABITABLES derivados de actividades de manufactura, alimentos o servicios sin acreditación habitable formal
    if any(k in full for k in ["PAPELERIA", "PAPELERÍA", "FLORES ETERNAS", "ARTESANIA", "ARTESANÍA", "ESTAMPADO", "SUBLIMACION"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos y Materiales de Manufactura",
            "justificacion": f"Actividad artesanal y de manufactura ({act_display}). Al no contar con acreditación técnica formal para recinto habitable, califica para bodega techada no habitable destinada al almacenamiento y resguardo seguro de materias primas, herramientas, maquinarias e insumos de trabajo.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["PASTEL", "REPOSTER", "MERMELADA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos y Almacenamiento de Alimentos",
            "justificacion": f"Elaboración de alimentos y conservas ({act_display}). Al no contar con acreditación técnica y sanitaria formal para recinto habitable, califica para bodega techada no habitable para almacenamiento y resguardo seguro de insumos secos, frascos, herramientas y equipamiento de trabajo.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["UÑA", "UÑAS", "MANICURE", "PELUQUERIA", "PODOLOGA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos y Equipamiento de Servicios",
            "justificacion": f"Prestación de servicios personales y estética ({act_display}). Al no contar con acreditación técnica formal para recinto de atención habitable en el predio, califica para bodega techada no habitable destinada al almacenamiento y resguardo seguro de insumos, instrumental y equipamiento de trabajo.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["TURISMO", "GRANJA EDUCATIVA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos / Resguardo de Granja",
            "justificacion": f"Actividad de granja y recreación rural ({act_display}). Al no contar con acreditación técnica formal para recinto de atención habitable, califica para bodega techada no habitable destinada al resguardo seguro de alimentos para animales, insumos y herramientas de granja.",
            "detalle_actividad": act_display
        }

    # 4. Recintos Complementarios NO HABITABLES (almacenamiento, acopio, animales, herramientas)
    if any(k in full for k in ["LEÑA", "LENA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Leñera Techada / Cobertizo de Acopio",
            "justificacion": f"Comercialización y acopio de leña ({act_display}). Requiere leñera techada y ventilada para acopio, secado y protección de humedad previo a la distribución.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["AGRICULTURA", "HUERTA", "HORTALIZA", "SIEMBRA", "TRIGO", "SEMILLA", "VACUNO"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos y Aperos Agrícolas",
            "justificacion": f"Explotación agropecuaria y cultivo ({act_display}). Requiere bodega techada para almacenamiento seguro de herramientas de labranza, fertilizantes, semillas y cosechas.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["HUEVO", "GALLINA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega / Galpón Avícola",
            "justificacion": f"Producción avícola y venta de huevos ({act_display}). Requiere recinto productivo no habitable para resguardo de alimentos, cajones y aves.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["PLANTA", "PLANTAS"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Invernadero / Bodega de Plantas",
            "justificacion": f"Venta y cultivo de plantas ({act_display}). Requiere recinto complementario tipo invernadero para resguardo de especies vegetales frente a heladas y almacenamiento de sustratos.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["GASFITER", "CONSTRUCCION", "CHANCADOR"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Herramientas y Equipos",
            "justificacion": f"Oficio técnico independiente ({act_display}). Requiere espacio techado para almacenaje seguro de herramientas, repuestos y maquinarias de trabajo.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["CATALOGO", "ROPA", "COMERCIANTE"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Mercadería",
            "justificacion": f"Venta y comercio de productos ({act_display}). Requiere espacio seco y seguro para almacenamiento de mercadería e inventario comercial.",
            "detalle_actividad": act_display
        }

    return {
        "procede": False,
        "tipo_recinto": "En Evaluación",
        "recinto_sugerido": "A determinar según visita técnica",
        "justificacion": f"Actividad registrada ({act_display}). Requiere inspección en terreno para validar procedencia de recinto.",
        "detalle_actividad": act_display
    }


def build_tabla_8(
    postulante: Dict[str, Any],
    recinto_info: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Construye la estructura de datos oficial para la Tabla 8 del Formulario PHR N°6.1 DTS:
    'JUSTIFICACIÓN DE RECINTO(S) COMPLEMENTARIO(S), SI PROCEDE'

    Filas en la plantilla Word:
    1: Bodega
    2: Recinto para realizar actividades productivas
    3: Otros Recintos Techados Adosados a la Vivienda
    4: Leñera
    5: Otros (especificar)

    Reglas de marcación:
    - Columna 1: Marcar 'Sí' o 'No' según corresponda (solo 'Sí' en la fila asignada si califica).
    - Columna 2: Explicar la actividad previa y justificación técnica (solo si tiene 'Sí').
    - Si no procede recinto complementario, las 5 filas se marcan con 'No' y Columna 2 vacía.
    """
    if recinto_info is None:
        recinto_info = evaluate_recinto_complementario(postulante)

    procede = bool(recinto_info.get("procede"))
    sug = recinto_info.get("recinto_sugerido", "")
    tipo = recinto_info.get("tipo_recinto", "")
    just = recinto_info.get("justificacion", "")

    target_row = None
    esp_label = ""

    if procede:
        if "Leñera" in sug or "Leera" in sug:
            target_row = 4
        elif "Invernadero" in sug:
            target_row = 5
            esp_label = "Invernadero"
        elif "Avícola" in sug or "Avcola" in sug or "Gallin" in sug:
            target_row = 5
            esp_label = "Galpón / Gallinero Avícola"
        elif tipo == "Habitable":
            target_row = 2
        else:
            target_row = 1

    return {
        "bodega": {
            "si_no": "Sí" if target_row == 1 else "No",
            "justificacion": just if target_row == 1 else ""
        },
        "actividades_productivas": {
            "si_no": "Sí" if target_row == 2 else "No",
            "justificacion": just if target_row == 2 else ""
        },
        "otros_adosados": {
            "si_no": "No",
            "justificacion": ""
        },
        "lenera": {
            "si_no": "Sí" if target_row == 4 else "No",
            "justificacion": just if target_row == 4 else ""
        },
        "otros_especificar": {
            "si_no": "Sí" if target_row == 5 else "No",
            "especificacion": esp_label if target_row == 5 else "",
            "justificacion": just if target_row == 5 else ""
        },
        "aplica_recinto": procede,
        "target_row": target_row,
        "recinto_sugerido": sug,
        "tipo_recinto": tipo
    }


def evaluate_tercer_dormitorio(postulante: Dict[str, Any]) -> tuple:
    """
    Determina si a la familia le corresponde marcar 'Tercer Dormitorio' en la Tabla 6.
    Criterios oficiales basados en la columna 'TIPO DE VIVIENDA' de la base más actualizada:
    1. 3° dormitorio con ahorro (52 familias)
    2. Movilidad reducida / 3° dormitorio (3 familias)
    3. Grupo Familiar sin ahorro (7 familias)
    4. Mov reducida conyuge + hija grup fam sin ahorro (1 familia)
    Retorna (aplica: bool, motivo: str).
    """
    tv_raw = str(postulante.get("tipo_vivienda") or "").upper()
    if not tv_raw:
        tv_raw = str(postulante.get("tipologia_propuesta") or "").upper()
    if not tv_raw:
        return False, ""

    # Normalizar caracteres ordinales, grados y espacios
    s = tv_raw.replace('º', ' ').replace('°', ' ').replace('\xba', ' ').replace('\xb0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()

    if "3" in s and "DORMITORIO" in s and "AHORRO" in s and "SIN AHORRO" not in s:
        return True, "3° dormitorio con ahorro"
    if "MOVILIDAD REDUCIDA" in s and "3" in s and "DORMITORIO" in s:
        return True, "Movilidad reducida / 3° dormitorio"
    if "MOV REDUCIDA CONYUGE" in s and "SIN AHORRO" in s:
        return True, "Mov reducida conyuge + hija grup fam sin ahorro"
    if "SIN AHORRO" in s:
        return True, "Grupo Familiar sin ahorro"

    return False, ""


TERRENO_PROYECTO_CONSOLIDADO: Dict[str, str] = {
    "calle": "Hijuela El Molino",
    "lote": "Lote 3 Foja 1169 N°722",
    "rol_sii": "202-43",
    "comuna": "PERQUENCO",
    "provincia": "Cautín",
    "region": "La Araucanía",
    "localidad": "Perquenco"
}


def consolidate_postulante_local(
    postulante: Dict[str, Any],
    usar_rsh_en_terreno: bool = False,
    datos_terreno_proyecto: Optional[Dict[str, str]] = None
) -> Dict[str, Any]:
    """
    Consolida la ficha del postulante de forma 100% determinista basada en la fuente.
    Garantiza que no se inventa ningún dato y que los campos sin respaldo quedan en blanco.
    Para la Tabla 3, utiliza el terreno consolidado del proyecto común.
    """
    # 1. Dirección RSH
    addr_rsh = parse_address_local(postulante.get("direccion_rsh", ""))
    
    # 2. Dirección Terreno / Subsidio (Terreno Consolidado del Proyecto)
    terr = datos_terreno_proyecto or TERRENO_PROYECTO_CONSOLIDADO

    # 3. Clasificación de integrantes y género
    sexo_titular = postulante.get("sexo", "").upper()
    if not sexo_titular:
        sexo_titular = infer_gender_from_name(postulante.get("nombre", "")) or ""

    edad_titular = postulante.get("edad")

    hombres_desc = []
    mujeres_desc = []
    menores_desc = []
    mayores_desc = []

    # Titular
    if sexo_titular == "M":
        desc = f"Jefe de hogar ({edad_titular} años)" if edad_titular else "Jefe de hogar"
        hombres_desc.append(desc)
    elif sexo_titular == "F":
        desc = f"Jefa de hogar ({edad_titular} años)" if edad_titular else "Jefa de hogar"
        mujeres_desc.append(desc)

    if edad_titular is not None:
        if edad_titular < 18:
            menores_desc.append(f"Titular ({edad_titular} años)")
        elif edad_titular >= 60:
            desc_m = f"Titular ({edad_titular} años)"
            mayores_desc.append(desc_m)

    # Parientes
    for p in postulante.get("parientes", []):
        p_nom = p.get("nombre", "")
        p_edad = p.get("edad")
        p_par = p.get("parentesco", "")

        # Determinar sexo
        p_sexo = infer_gender_from_name(p_nom)
        if not p_sexo:
            if "HIJA" in p_par or "MADRE" in p_par or "HERMANA" in p_par:
                p_sexo = "F"
            elif "HIJO" in p_par or "PADRE" in p_par or "HERMANO" in p_par:
                p_sexo = "M"

        # Descripción de parentesco exacto según género reconocido (Hijo vs Hija)
        label_par = format_parentesco(p_par, p_sexo)
        if p_edad is not None:
            edad_str = "1 año" if p_edad == 1 else f"{p_edad} años"
        else:
            edad_str = ""
        desc_p = f"{label_par} ({edad_str})" if edad_str else label_par

        if p_sexo == "M":
            hombres_desc.append(desc_p)
        elif p_sexo == "F":
            mujeres_desc.append(desc_p)

        if p_edad is not None:
            if p_edad < 18:
                menores_desc.append(desc_p)
            elif p_edad >= 60:
                mayores_desc.append(desc_p)

    cant_hombres = len(hombres_desc)
    cant_mujeres = len(mujeres_desc)
    cant_menores = len(menores_desc)
    cant_mayores = len(mayores_desc)

    # Discapacidad formal según criterios técnicos y rol acreditado
    disc_desc, cant_disc = build_discapacidad_observations(postulante)

    # Indígena
    etnia_raw = postulante.get("etnia", "").strip().upper()
    es_indigena = "MAPUCHE" in etnia_raw
    indigena_obs = "Pueblo Mapuche" if es_indigena else ""
    cant_indigena = postulante.get("grupo_familiar_cant", 1) if es_indigena else 0

    # Extranjeros
    pais_raw = postulante.get("pais", "").strip().upper()
    es_extranjero = pais_raw != "" and pais_raw != "CHILE"
    extranjero_obs = pais_raw if es_extranjero else ""
    cant_extranjeros = 1 if es_extranjero else 0

    # Cónyuge consolidado
    conyuge = postulante.get("conyuge", {})

    # 4. Actividad Económica (Tabla 5)
    act_data = postulante.get("actividad_economica") or {}
    act_raw = act_data.get("actividad_principal", "")
    desc_raw = act_data.get("descripcion", "")
    tiene_act = act_data.get("tiene_actividad", False)

    # Actualización factual según resolución técnica de terreno EGR
    rut_clean_post = re.sub(r'[^0-9Kk]', '', str(postulante.get("rut") or "")).upper()
    if rut_clean_post == "197921137":
        act_raw = "REPOSTERIA"
        desc_raw = "Repostería artesanal"
        tiene_act = True
    elif rut_clean_post == "209829479":
        act_raw = "VENTA DE QUESOS"
        desc_raw = "Elaboración y venta de quesos"
        tiene_act = True

    tabla_5 = {
        "categoria": "",
        "especificacion": "",
        "actividad_fuente": act_raw,
        "descripcion_fuente": desc_raw,
        "sin_actividad_declarada": False,
        "criterio_social": None,
        "actividades": {}
    }

    if tiene_act and (act_raw or desc_raw):
        cat, esp = classify_economic_activity(act_raw, desc_raw)
        if cat:
            jefe_val = act_data.get("jefe_hogar", "")
            cony_val = act_data.get("conyuge_pareja", "")
            otros_val = act_data.get("otros", "")

            jefe_m = ""
            jefe_f = ""
            cony_m = ""
            cony_f = ""
            otros_m = ""
            otros_f = ""

            # Si la columna Jefe de Hogar está marcada, o si no se especificó rol (actividad propia del titular)
            if jefe_val or (not cony_val and not otros_val):
                # Usar el género factual del titular para no crear contradicción documental
                if sexo_titular == "M":
                    jefe_m = "X"
                elif sexo_titular == "F":
                    jefe_f = "X"
                else:
                    if jefe_val == "M":
                        jefe_m = "X"
                    else:
                        jefe_f = "X"

            # Si la columna Cónyuge está marcada
            if cony_val:
                cony_sexo = conyuge.get("sexo", "").upper()
                if cony_sexo == "M":
                    cony_m = "X"
                elif cony_sexo == "F":
                    cony_f = "X"
                elif cony_val == "M":
                    cony_m = "X"
                elif cony_val == "F":
                    cony_f = "X"

            # Si la columna Otros está marcada
            if otros_val == "M":
                otros_m = "X"
            elif otros_val == "F":
                otros_f = "X"

            tabla_5 = {
                "categoria": cat,
                "especificacion": esp,
                "actividad_fuente": act_raw,
                "descripcion_fuente": desc_raw,
                "sin_actividad_declarada": False,
                "criterio_social": None,
                "actividades": {
                    cat: {
                        "jefe_m": jefe_m,
                        "jefe_f": jefe_f,
                        "conyuge_m": cony_m,
                        "conyuge_f": cony_f,
                        "otros_m": otros_m,
                        "otros_f": otros_f
                    }
                }
            }

    # Si no tiene actividad económica declarada o no aplica justificación:
    # Criterio normativo social según género y edad del titular:
    # - Mujeres < 60 años: Dueña de casa (Jefe Hogar F = X)
    # - Mujeres >= 60 años: Jubilada (Jefe Hogar F = X)
    # - Hombres < 65 años: Cesante (Jefe Hogar M = X)
    # - Hombres >= 65 años: Jubilado (Jefe Hogar M = X)
    # Se explicita en la fila 'Otras (Especificar)' de la Tabla 5.
    if not tabla_5.get("categoria"):
        if sexo_titular == "F":
            glosa = "Dueña de casa" if (edad_titular is None or edad_titular < 60) else "Jubilada"
            jefe_m, jefe_f = "", "X"
        else:
            glosa = "Cesante" if (edad_titular is None or edad_titular < 65) else "Jubilado"
            jefe_m, jefe_f = "X", ""

        tabla_5 = {
            "categoria": "otras",
            "especificacion": glosa,
            "actividad_fuente": "",
            "descripcion_fuente": f"Criterio normativo social ({glosa})",
            "sin_actividad_declarada": True,
            "criterio_social": glosa,
            "actividades": {
                "otras": {
                    "jefe_m": jefe_m,
                    "jefe_f": jefe_f,
                    "conyuge_m": "",
                    "conyuge_f": "",
                    "otros_m": "",
                    "otros_f": ""
                }
            }
        }

    return {
        "tabla_1": {
            "titular_nombre": postulante.get("nombre", ""),
            "titular_rut": postulante.get("rut", ""),
            "titular_fecha_nac": postulante.get("fecha_nacimiento", ""),
            "titular_estado_civil": postulante.get("estado_civil", ""),
            "titular_sexo": sexo_titular,
            "titular_telefono": postulante.get("telefono", ""),  # Vacío en Excel
            "conyuge_nombre": conyuge.get("nombre", ""),
            "conyuge_rut": conyuge.get("rut", ""),
            "conyuge_sexo": conyuge.get("sexo", ""),
            "conyuge_fecha_nac": conyuge.get("fecha_nacimiento", "")
        },
        "tabla_2": {
            "calle": addr_rsh.get("calle", ""),
            "numero": addr_rsh.get("numero", ""),
            "lote": addr_rsh.get("lote", ""),
            "region": "La Araucanía",
            "comuna": postulante.get("comuna", "PERQUENCO"),
            "provincia": "Cautín",
            "localidad": "Perquenco" if addr_rsh.get("calle") else ""
        },
        "tabla_3": {
            "calle": terr.get("calle", "Hijuela El Molino"),
            "lote": terr.get("lote", "Lote 3 Foja 1169 N°722"),
            "localidad": terr.get("localidad", "Perquenco"),
            "factor_aislamiento": postulante.get("factor_aislamiento", ""),
            "comuna": terr.get("comuna", postulante.get("comuna", "PERQUENCO")),
            "provincia": terr.get("provincia", "Cautín"),
            "region": terr.get("region", "La Araucanía"),
            "rol_sii": terr.get("rol_sii", "202-43")
        },
        "tabla_4": {
            "total_habitantes": str(postulante.get("grupo_familiar_cant", "")),
            "hombres": {
                "si": "X" if cant_hombres > 0 else "",
                "no": "X" if cant_hombres == 0 else "",
                "cuantos": str(cant_hombres) if cant_hombres > 0 else "0",
                "observaciones": ", ".join(hombres_desc)
            },
            "mujeres": {
                "si": "X" if cant_mujeres > 0 else "",
                "no": "X" if cant_mujeres == 0 else "",
                "cuantos": str(cant_mujeres) if cant_mujeres > 0 else "0",
                "observaciones": ", ".join(mujeres_desc)
            },
            "menores_18": {
                "si": "X" if cant_menores > 0 else "",
                "no": "X" if cant_menores == 0 else "",
                "cuantos": str(cant_menores) if cant_menores > 0 else "0",
                "observaciones": ", ".join(menores_desc)
            },
            "adultos_mayores": {
                "si": "X" if cant_mayores > 0 else "",
                "no": "X" if cant_mayores == 0 else "",
                "cuantos": str(cant_mayores) if cant_mayores > 0 else "0",
                "observaciones": ", ".join(mayores_desc)
            },
            "discapacidad": {
                "si": "X" if cant_disc > 0 else "",
                "no": "X" if cant_disc == 0 else "",
                "cuantos": str(cant_disc) if cant_disc > 0 else "0",
                "observaciones": " / ".join(disc_desc)
            },
            "indigena": {
                "si": "X" if es_indigena else "",
                "no": "X" if not es_indigena else "",
                "cuantos": str(cant_indigena) if es_indigena else "0",
                "observaciones": indigena_obs
            },
            "extranjeros": {
                "si": "X" if es_extranjero else "",
                "no": "X" if not es_extranjero else "",
                "cuantos": str(cant_extranjeros) if es_extranjero else "0",
                "observaciones": extranjero_obs
            }
        },
        "tabla_5": tabla_5,
        "tabla_6": {
            "conjunto_habitacional": {
                "vivienda_nueva": "X",
                "tercer_dormitorio": "X" if evaluate_tercer_dormitorio(postulante)[0] else "",
                "recinto_complementario": "X" if evaluate_recinto_complementario(postulante).get("tipo_recinto") in ["Habitable", "No Habitable"] else "",
                "motivo_tercer_dormitorio": evaluate_tercer_dormitorio(postulante)[1],
                "motivo_recinto": evaluate_recinto_complementario(postulante).get("recinto_sugerido", "") if evaluate_recinto_complementario(postulante).get("tipo_recinto") in ["Habitable", "No Habitable"] else ""
            },
            "sitio_residente": {
                "vivienda_nueva": "",
                "tercer_dormitorio": "",
                "recinto_complementario": ""
            }
        },
        "tabla_8": build_tabla_8(postulante, evaluate_recinto_complementario(postulante)),
        "tercer_dormitorio": {
            "aplica": evaluate_tercer_dormitorio(postulante)[0],
            "motivo": evaluate_tercer_dormitorio(postulante)[1]
        },
        "tipo_familia": classify_family_nucleus(postulante),
        "recinto_complementario": evaluate_recinto_complementario(postulante)
    }


def audit_with_gemini(
    postulante: Dict[str, Any],
    api_key: str,
    usar_rsh_en_terreno: bool = False,
    model_name: str = "gemini-2.5-flash"
) -> Dict[str, Any]:
    """
    Utiliza Gemini como auditor ('ojos' del sistema) para refinar la extracción de datos
    complejos (direcciones no estructuradas, notas de discapacidad, desambiguación de nombres)
    con estricta prohibición de alucinar o inventar información.
    """
    # Si no hay API key, usar directamente el motor determinista local
    if not api_key or not str(api_key).strip():
        return consolidate_postulante_local(postulante, usar_rsh_en_terreno)

    # Base consolidada previa
    base_data = consolidate_postulante_local(postulante, usar_rsh_en_terreno)

    prompt = f"""
Eres el auditor estricto de datos ("los ojos del sistema") para el llenado oficial del formulario MINVU D.S. N°10 Habitabilidad Rural.
TU MISIÓN: Refinar y consolidar la información con fidelidad absoluta a la fuente provista.
REGLA INQUEBRANTABLE:
- JAMÁS inventes datos que no estén en la fuente.
- Si un campo no viene en el archivo o no se deduce con certeza absoluta, debe ser una cadena vacía ("").
- En Actividades Económicas (Tabla 5), preserva fielmente los datos fidedignos sin crear actividades ficticias.

DATOS CRUDOS DEL POSTULANTE:
- Nombre Titular: {postulante.get('nombre')}
- RUT: {postulante.get('rut')}
- Fecha Nacimiento: {postulante.get('fecha_nacimiento')} | Edad: {postulante.get('edad')}
- Sexo Titular: {postulante.get('sexo')}
- Estado Civil: {postulante.get('estado_civil')}
- Teléfono: {postulante.get('telefono')}
- Dirección RSH cruda: {postulante.get('direccion_rsh')}
- Dirección Terreno cruda: {postulante.get('direccion_terreno')}
- Comuna: {postulante.get('comuna')}
- Etnia: {postulante.get('etnia')}
- País: {postulante.get('pais')}
- Nota Discapacidad: {postulante.get('discapacidad')}
- Cantidad Grupo Familiar: {postulante.get('grupo_familiar_cant')}
- Cónyuge/Pareja: {postulante.get('conyuge')}
- Parientes: {json.dumps(postulante.get('parientes', []), ensure_ascii=False)}

TAREAS DE AUDITORÍA:
1. Revisa la separación de la dirección RSH en 'calle', 'numero', 'lote' sin inventar números ni agregar prefijos innecesarios.
2. Identifica con total precisión si cada hijo o pariente es Masculino ('Hijo') o Femenino ('Hija') analizando sus nombres de pila chilenos.
3. En la Tabla 4 debes clasificar y describir cada hijo/a como 'Hijo' en la fila de Hombres y 'Hija' en la fila de Mujeres (nunca como 'Hijo/a').
4. En la fila 'Menores de 18 años', desglosa explícitamente los hijos identificando su género y edad (ej. '1 hijo (11 años), 1 hija (16 años)').
5. Extrae la condición factual exacta de discapacidad si existe una nota, o deja en blanco si no hay.
6. Genera observaciones breves y 100% factuales (edad y rol) para la Tabla 4.
7. Devuelve la respuesta EXACTA en el siguiente esquema JSON (idéntico a la estructura requerida):

Responde ÚNICAMENTE con el objeto JSON que complete los campos de tabla_1, tabla_2, tabla_3 y tabla_4.
"""

    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.0
            )
        )
        if response.text:
            gemini_json = json.loads(response.text)
            # Mezclar verificando que conserve la estructura base
            if "tabla_1" in gemini_json and "tabla_4" in gemini_json:
                gemini_json["tabla_3"] = base_data["tabla_3"]  # Preservar el terreno consolidado del proyecto
                gemini_json["tabla_5"] = base_data["tabla_5"]  # Preservar la clasificación fidedigna de Tabla 5
                gemini_json["tabla_6"] = base_data["tabla_6"]  # Preservar la marcación de Modalidad Vivienda Nueva
                gemini_json["tabla_8"] = base_data["tabla_8"]  # Preservar la marcación y justificación técnica de Tabla 8
                gemini_json["tercer_dormitorio"] = base_data.get("tercer_dormitorio")
                gemini_json["tipo_familia"] = base_data.get("tipo_familia")
                gemini_json["recinto_complementario"] = base_data.get("recinto_complementario")
                return gemini_json
    except Exception as e:
        # Si ocurre cualquier error en API o cuota, caemos de manera segura al consolidado local
        print(f"Aviso: Fallback a consolidación local debido a: {e}")

    return base_data

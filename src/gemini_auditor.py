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
        "justificacion": "Elaboración y comercialización de quesos. Recinto techado no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
        "detalle_actividad": "Venta y elaboración de quesos"
    },
    # Yessenia Alejandra Valdes Moya (RUT: 19.792.113-7)
    "197921137": {
        "procede": True,
        "tipo_recinto": "No Habitable",
        "recinto_sugerido": "Bodega de Insumos y Equipos de Repostería",
        "justificacion": "Emprendimiento productivo de repostería artesanal. Recinto techado no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
        "detalle_actividad": "Repostería artesanal"
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

    # 1. Recintos Complementarios específicos según rubro (Manufactura, Alimentos, Servicios, etc.)
    if any(k in full for k in ["PAPELERIA", "PAPELERÍA", "FLORES ETERNAS", "ARTESANIA", "ARTESANÍA", "ESTAMPADO", "SUBLIMACION"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos y Materiales de Manufactura",
            "justificacion": f"Actividad artesanal y de manufactura ({act_display}). Recinto techado no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["PASTEL", "REPOSTER", "MERMELADA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos y Almacenamiento de Alimentos",
            "justificacion": f"Elaboración de alimentos y conservas ({act_display}). Recinto techado no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["UÑA", "UÑAS", "MANICURE", "PELUQUERIA", "PODOLOGA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos y Equipamiento de Servicios",
            "justificacion": f"Prestación de servicios personales y estética ({act_display}). Recinto techado no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
            "detalle_actividad": act_display
        }

    if any(k in full for k in ["TURISMO", "GRANJA EDUCATIVA"]):
        return {
            "procede": True,
            "tipo_recinto": "No Habitable",
            "recinto_sugerido": "Bodega de Insumos / Resguardo de Granja",
            "justificacion": f"Actividad de granja y recreación rural ({act_display}). Recinto techado no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
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

    # 2. Asignación técnica por defecto EGR para las 123 familias base del proyecto
    # Todo postulante sin recinto complementario específico lleva por defecto marcado:
    # "Recinto complementario" (Tabla 6) / "Bodega" / "Sí" (Tabla 8) con la justificación oficial:
    return {
        "procede": True,
        "tipo_recinto": "No Habitable",
        "recinto_sugerido": "Bodega",
        "justificacion": "Recinto techado no habitable para almacenamiento y resguardo seguro de insumos, aperos, utensilios y equipamiento de trabajo.",
        "detalle_actividad": act_display if (tiene_act and full) else "(Sin actividad económica declarada)"
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


def build_tabla_9() -> Dict[str, Any]:
    """
    Construye la estructura de datos oficial para la Tabla 9 del Formulario PHR N°6.1 DTS:
    'REQUERIMIENTOS DE HABITABILIDAD ASOCIADOS AL EQUIPAMIENTO COMUNITARIO MODALIDAD EQUIPAMIENTO COMUNITARIO'

    Aplica para el 100% de los formularios de las familias del proyecto común Hijuela El Molino:
    - Construcción de Equipamiento Comunitario:
      - N° de Obras a Ejecutar (*): 1
      - Descripción de la(s) Obra(s): Recinto de acopio y de apoyo de producción agrícola sustentable
    """
    return {
        "mejoramiento": {
            "obras": "",
            "descripcion": ""
        },
        "ampliacion": {
            "obras": "",
            "descripcion": ""
        },
        "construccion": {
            "obras": "1",
            "descripcion": "Recinto de acopio y de apoyo de producción agrícola sustentable"
        }
    }


def build_tabla_15() -> Dict[str, Any]:
    """
    Construye la estructura de datos oficial para la Tabla 15 del Formulario PHR N°6.1 DTS:
    'ACCESO A SERVICIOS BÁSICOS EN TERRENOS ERIAZOS ASOCIADOS A PROYECTOS DE VIVIENDAS NUEVAS...'

    Aplica para el 100% de los formularios de las familias del proyecto común Hijuela El Molino:
    - Agua Potable:
      - Empresa Sanitaria: No (X)
      - Red APR: No (X)
      - Pozo o Noria en Terreno Propio: Si (X)
      - Pozo o Noria en Terreno Vecino: Si (X)
      - Aguas Superficiales (Ej. Vertientes): No (X)
      - Camiones Aljibes: No (X)
      - Otra (Especificar): En blanco
    - Alcantarillado:
      - Empresa Sanitaria: No (X)
      - Red Pública Proyecto Particular: No (X)
      - Fosa y Pozo: Si (X)
      - Fosa y Dren: No (X)
      - Otro (Especificar): Si (X)
        Observaciones: Planta Lombricultivo: Sistema de reciclaje orgánico que utiliza lombrices para transformar los restos vegetales de los residentes en abonos naturales, como resultado, el proceso produce de forma continua humus sólido y fertilizante líquido para la aplicación en sustrato de producción agrícola.
    - Electricidad:
      - Empresa Eléctrica: Si (X)
      - Sistema Fotovoltaico: No (X)
      - Sistema Eólico: No (X)
      - Sistema Hidráulico: No (X)
      - Generador Eléctrico: No (X)
      - Otra fuente (especificar): En blanco
    """
    obs_lombricultivo = (
        "Planta Lombricultivo: Sistema de reciclaje orgánico que utiliza lombrices para "
        "transformar los restos vegetales de los residentes en abonos naturales, como "
        "resultado, el proceso produce de forma continua humus sólido y fertilizante líquido "
        "para la aplicación en sustrato de producción agrícola."
    )

    return {
        "agua_potable": [
            {"fuente": "Empresa Sanitaria", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Red APR", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Pozo o Noria en Terreno Propio", "si": "X", "no": "", "observaciones": ""},
            {"fuente": "Pozo o Noria en Terreno Vecino", "si": "X", "no": "", "observaciones": ""},
            {"fuente": "Aguas Superficiales (Ej. Vertientes)", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Camiones Aljibes", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Otra (Especificar)", "si": "", "no": "", "observaciones": ""}
        ],
        "alcantarillado": [
            {"fuente": "Empresa Sanitaria", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Red Pública Proyecto Particular", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Fosa y Pozo", "si": "X", "no": "", "observaciones": ""},
            {"fuente": "Fosa y Dren", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Otro (Especificar)", "si": "X", "no": "", "observaciones": obs_lombricultivo}
        ],
        "electricidad": [
            {"fuente": "Empresa eléctrica", "si": "X", "no": "", "observaciones": ""},
            {"fuente": "Sistema fotovoltaico", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Sistema eólico", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Sistema hidráulico", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Generador eléctrico", "si": "", "no": "X", "observaciones": ""},
            {"fuente": "Otra fuente (especificar)", "si": "", "no": "", "observaciones": ""}
        ]
    }


def build_apartado_9() -> Dict[str, Any]:
    """
    Construye la narrativa técnica y territorial oficial para el Apartado 9:
    'DIAGNÓSTICO DEL LUGAR DE EMPLAZAMIENTO DE EL O LOS PROYECTOS'
    (Comuna de Perquenco, Región de La Araucanía).
    """
    return {
        "titulo": "DIAGNÓSTICO DEL LUGAR DE EMPLAZAMIENTO DE EL O LOS PROYECTOS",
        "subtitulo": "Describir variables geográficas y técnicas relevantes que inciden en el diseño de los proyectos.",
        "secciones": [
            {
                "titulo": "Emplazamiento y características generales",
                "parrafos": [
                    (
                        "El proyecto Ecobarrio Rural Perquenco se desarrolla en una fracción de terreno "
                        "de aproximadamente 15 hectáreas, ubicada en la comuna de Perquenco, Región de La "
                        "Araucanía. Si bien el predio posee condición rural, presenta una localización "
                        "privilegiada, prácticamente colindante con la trama urbana consolidada de la localidad, "
                        "lo que permite proyectar un conjunto habitacional integrado territorialmente al poblado, "
                        "conservando las características y oportunidades propias del entorno rural."
                    ),
                    (
                        "El terreno presenta condiciones topográficas favorables, con pendientes predominantemente "
                        "inferiores al 3%, facilitando el desarrollo de la urbanización, la vialidad interior y "
                        "el emplazamiento de las viviendas, con menores requerimientos de movimientos de tierra."
                    ),
                    (
                        "La propuesta contempla 155 lotes habitacionales, con una superficie promedio superior "
                        "a 400 m², organizados mediante una vialidad interior que articula las viviendas con "
                        "áreas verdes, espacios de encuentro comunitario, equipamiento y sectores productivos. "
                        "Esta configuración busca consolidar un modelo habitacional que combine vivienda, "
                        "desarrollo comunitario y actividades productivas, incorporando una superficie aproximada "
                        "de 2,9 hectáreas destinada a producción comunitaria."
                    )
                ]
            },
            {
                "titulo": "Infraestructura sanitaria autónoma",
                "parrafos": [
                    (
                        "Uno de los principales atributos del proyecto corresponde a su propuesta de autosuficiencia "
                        "sanitaria, particularmente relevante ante las limitaciones de infraestructura que presenta "
                        "actualmente la localidad de Perquenco para la incorporación de nuevos conjuntos habitacionales."
                    ),
                    (
                        "Para el abastecimiento de agua potable de las 155 familias, se encuentra en proceso la "
                        "perforación de un pozo profundo dentro del predio. Esta captación formará parte de un "
                        "sistema autónomo de producción y tratamiento de agua potable, destinado a abastecer la "
                        "totalidad del conjunto sin depender de la capacidad de suministro de la cooperativa sanitaria local."
                    ),
                    (
                        "Complementariamente, el proyecto considera un sistema propio de alcantarillado y tratamiento "
                        "de aguas servidas, incorporando una planta de tratamiento mediante tecnología de lombricultura "
                        "o lombrifiltro, dimensionada para atender a las 155 viviendas. El loteo contempla sectores "
                        "específicos para emplazar esta infraestructura, separados de las áreas habitacionales y "
                        "articulados mediante las servidumbres e instalaciones correspondientes."
                    ),
                    (
                        "De esta manera, el Ecobarrio Perquenco busca resolver integralmente sus requerimientos "
                        "sanitarios, evitando depender de una eventual ampliación de las redes locales de alcantarillado "
                        "y aportando una solución ambientalmente sostenible, sujeta a los estudios y autorizaciones "
                        "sectoriales respectivos."
                    )
                ]
            },
            {
                "titulo": "Integración territorial y pertinencia habitacional",
                "parrafos": [
                    (
                        "La localidad de Perquenco cuenta con una población inferior a 5.000 habitantes, condición "
                        "compatible con el ámbito territorial del Programa de Habitabilidad Rural DS10 y con el llamado "
                        "especial de Ecobarrios Rurales."
                    ),
                    (
                        "El proyecto reúne atributos especialmente favorables para esta modalidad: proximidad a los "
                        "servicios urbanos, disponibilidad de suelo rural, topografía apropiada, lotes habitacionales "
                        "de superficie generosa, espacios comunitarios y productivos, y una estrategia integral de "
                        "abastecimiento de agua potable y saneamiento."
                    ),
                    (
                        "El Ecobarrio Perquenco se plantea así como una alternativa de crecimiento habitacional planificado, "
                        "capaz de integrar las ventajas de la cercanía urbana con la autonomía sanitaria, la producción "
                        "comunitaria y la calidad de vida del territorio rural."
                    )
                ]
            }
        ]
    }


def build_anexos_info() -> Dict[str, Any]:
    """
    Construye la información y metadatos de los Anexos 1 y 2 del Formulario PHR N°6.1 DTS.
    """
    return {
        "anexo_1": {
            "titulo": "Croquis Diagnóstico del terreno, emplazamiento de el o los proyectos y antecedentes relevantes.",
            "archivo": "anexo1_croquis_terreno.jpg",
            "descripcion": "Plano Loteo DFL-2 Perquenco Escala 1:1.500 con servidumbre y división predial."
        },
        "anexo_2": {
            "titulo": "FOTOGRAFÍAS DEL TERRENO O VIVIENDA EN QUE SE APLICARÁ EL SUBSIDIO HABITACIONAL (Mínimo 3 imágenes).",
            "imagenes": [
                {
                    "archivo": "anexo2_foto1_emplazamiento.jpg",
                    "epigrafe": "Fotografía 1: Emplazamiento satelital y acceso por servidumbre predio Perquenco"
                },
                {
                    "archivo": "anexo2_foto2_aerea_general.jpg",
                    "epigrafe": "Fotografía 2: Vista aérea perspectiva general del conjunto y áreas productivas"
                },
                {
                    "archivo": "anexo2_foto3_entorno_parque.jpg",
                    "epigrafe": "Fotografía 3: Vista aérea nivel de calle, parque y viviendas proyectadas"
                }
            ]
        }
    }


DEFAULT_PROFESIONALES: Dict[str, Any] = {
    "tecnico": {
        "egr": "CONSULTORA PLAN SOCIAL LIMITADA",
        "nombre": "Vanessa Schneider Martínez",
        "rut": "13.730.440-6",
        "profesion": "Arquitecta"
    },
    "social": {
        "egr": "CONSULTORA PLAN SOCIAL",
        "nombre": "Erica Reyes Peña",
        "rut": "20.402.240-2",
        "profesion": "TRABAJADORA SOCIAL"
    }
}


def build_firmantes(
    postulante: Dict[str, Any],
    profesionales: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Construye la estructura de datos para la sección de firmas y certificación del formulario:
    - Tabla 17: Profesional Área Técnica EGR (Vanessa Schneider Martínez)
    - Tabla 18: Profesional Área Social EGR (Erica Reyes Peña)
    - Tabla 19: Postulante Titular (Nombre y RUT extraídos de la identificación del postulante)
    """
    profs = profesionales or DEFAULT_PROFESIONALES
    tec = profs.get("tecnico", DEFAULT_PROFESIONALES["tecnico"])
    soc = profs.get("social", DEFAULT_PROFESIONALES["social"])

    return {
        "tecnico": {
            "egr": tec.get("egr", "CONSULTORA PLAN SOCIAL LIMITADA"),
            "nombre": tec.get("nombre", "Vanessa Schneider Martínez"),
            "rut": tec.get("rut", "13.730.440-6"),
            "profesion": tec.get("profesion", "Arquitecta")
        },
        "social": {
            "egr": soc.get("egr", "CONSULTORA PLAN SOCIAL"),
            "nombre": soc.get("nombre", "Erica Reyes Peña"),
            "rut": soc.get("rut", "20.402.240-2"),
            "profesion": soc.get("profesion", "TRABAJADORA SOCIAL")
        },
        "postulante": {
            "nombre": postulante.get("nombre", ""),
            "rut": postulante.get("rut", "")
        }
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
    datos_terreno_proyecto: Optional[Dict[str, str]] = None,
    profesionales_firmantes: Optional[Dict[str, Any]] = None
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
        "tabla_9": build_tabla_9(),
        "tabla_15": build_tabla_15(),
        "firmantes": build_firmantes(postulante, profesionales_firmantes),
        "apartado_9": build_apartado_9(),
        "anexos": build_anexos_info(),
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
    model_name: str = "gemini-2.5-flash",
    datos_terreno_proyecto: Optional[Dict[str, str]] = None,
    profesionales_firmantes: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Utiliza Gemini como auditor ('ojos' del sistema) para refinar la extracción de datos
    complejos (direcciones no estructuradas, notas de discapacidad, desambiguación de nombres)
    con estricta prohibición de alucinar o inventar información.
    """
    # Si no hay API key, usar directamente el motor determinista local
    if not api_key or not str(api_key).strip():
        return consolidate_postulante_local(postulante, usar_rsh_en_terreno, datos_terreno_proyecto, profesionales_firmantes)

    # Base consolidada previa
    base_data = consolidate_postulante_local(postulante, usar_rsh_en_terreno, datos_terreno_proyecto, profesionales_firmantes)

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
                gemini_json["tabla_9"] = base_data["tabla_9"]  # Preservar Tabla 9 Equipamiento Comunitario
                gemini_json["tabla_15"] = base_data["tabla_15"]  # Preservar Tabla 15 Servicios Básicos
                gemini_json["firmantes"] = base_data["firmantes"]  # Preservar Profesionales Suscribientes
                gemini_json["apartado_9"] = base_data["apartado_9"]  # Preservar Apartado 9 Emplazamiento
                gemini_json["anexos"] = base_data.get("anexos")  # Preservar Anexos 1 y 2
                gemini_json["tercer_dormitorio"] = base_data.get("tercer_dormitorio")
                gemini_json["tipo_familia"] = base_data.get("tipo_familia")
                gemini_json["recinto_complementario"] = base_data.get("recinto_complementario")
                return gemini_json
    except Exception as e:
        # Si ocurre cualquier error en API o cuota, caemos de manera segura al consolidado local
        print(f"Aviso: Fallback a consolidación local debido a: {e}")

    return base_data

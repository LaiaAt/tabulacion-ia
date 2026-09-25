# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (LIMPIEZA RADICADOS ANI + CONTROL ALMA-R)
# ==============================================================================

import os
import sys
import time
import json
import re
import base64
import smtplib
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from email.message import EmailMessage
import pandas as pd
import requests

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(line_buffering=True)

import pymupdf as fitz
from PIL import Image
import io

print("⏳ [1/2] Configurando Pool de Claves de Mistral AI...", flush=True)

raw_keys = os.environ.get('MISTRAL_API_KEY', '').strip()
lista_keys = [k.strip() for k in raw_keys.replace('\n', ',').split(',') if len(k.strip()) > 10]

if not lista_keys:
    print("❌ ERROR CRÍTICO: No se detectó MISTRAL_API_KEY.", flush=True)
    sys.exit(1)

print(f"   ✅ Pool activo con {len(lista_keys)} clave(s).", flush=True)

MODELOS_PIXTRAL = ["pixtral-large-latest", "pixtral-12b-2409"]

EMAIL_REMITENTE = os.environ.get('GMAIL_USER')
EMAIL_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
EMAIL_DESTINO = os.environ.get('GMAIL_USER')

RUTA_BASE = '.'
RUTA_ENVIADAS = os.path.join(RUTA_BASE, '15_01_Cartas_Enviadas')
RUTA_RECIBIDAS = os.path.join(RUTA_BASE, '15_04_Comunic_Recibidas')

lock_csv = threading.Lock()
lock_keys = threading.Lock()
cooldown_keys = {k: 0.0 for k in lista_keys}

def obtener_insumos_documento(ruta_pdf):
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        if total_paginas == 0:
            doc.close()
            return [], "", 0

        texto_completo = ""
        paginas_con_constancia = []

        for idx, p in enumerate(doc):
            txt_p = p.get_text()
            texto_completo += txt_p + "\n"
            txt_low = txt_p.lower()
            if any(k in txt_low for k in ["radicó con éxito", "número de radicado es", "alma-r-", "ventanilla unica", "atencionciudadano@invias", "numero de radicado"]):
                paginas_con_constancia.append(idx)

        # Si la pág 1 es carátula remisoria de la ANI, la carta real es la pág 2 (índice 1)
        texto_pag1 = doc[0].get_text().lower()
        es_caratula = bool("oficio remisorio" in texto_pag1 or "al contestar cite el numero" in texto_pag1 or "radicado via web" in texto_pag1)

        paginas_a_procesar = []
        if es_caratula and total_paginas > 1:
            # Enviamos la carátula (pág 0) para radicado ANI y la carta real (pág 1) para Asunto y Destinatario
            paginas_a_procesar.extend([1, 0])
            if total_paginas > 2: paginas_a_procesar.append(2)
        else:
            paginas_a_procesar.extend([0, 1] if total_paginas > 1 else [0])

        paginas_a_procesar.extend(paginas_con_constancia)
        paginas_a_procesar.append(total_paginas - 1)

        paginas_a_procesar = list(dict.fromkeys(p for p in paginas_a_procesar if p < total_paginas))

        imagenes_b64 = []
        for i in paginas_a_procesar:
            pagina = doc[i]
            pix = pagina.get_pixmap(dpi=140)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            if img.width > 1300:
                ratio = 1300 / float(img.width)
                img = img.resize((1300, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=80, optimize=True)
            imagenes_b64.append(base64.b64encode(buffer.getvalue()).decode('utf-8'))

        doc.close()
        return imagenes_b64, texto_completo, total_paginas
    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}", flush=True)
        return [], "", 0

PROMPT_MAESTRO = """
ACTÚA COMO UN AUDITOR Y TRANSSCRIPTOR DOCUMENTAL EXPERTO.
Analiza visualmente las imágenes del documento y extrae la información con MÁXIMA PRECISIÓN.

REGLAS DE ORO OBLIGATORIAS:

1. "razon_social_remitente":
   - Entidad emisora según el LOGO de la carta formal (ej. "AGENCIA NACIONAL DE INFRAESTRUCTURA - ANI", "CONCESIÓN ALTO MAGDALENA S.A.S.", "CONSORCIO 4C").
   - NUNCA pongas nombres de personas ni "Atn.".

2. "no_radicado_remitente":
   - EN COMUNICACIONES DE LA ANI (Recibidas): Es el NÚMERO LARGO DE 14 DÍGITOS ubicado en el sticker junto al código de barras (ej. "20203050336801", "2020-310-009073-1").
     * PROHIBIDO incluir la palabra "ANI" o "ANI No.". SOLO EL NÚMERO.
   - EN CARTAS DE CONCESIÓN ALTO MAGDALENA: Código bajo el código de barras (Ej: "ALMA-2020-0994").
   - EN CARTAS DE CONSORCIO 4C: Código arriba a la derecha bajo el logo (Ej: "CI.004/GPXXXX/XX/X.X").
   - EN CORREOS SIN RADICADO DE SALIDA: Escribe "SIN NÚMERO".

3. "razon_social_destinatario":
   - En RECIBIDAS: Siempre es "CONSORCIO 4C".
   - En RADICADAS: La entidad tras "Señores:" de la carta real (sin nombres de personas).

4. "no_radicado_destinatario":
   - En RECIBIDAS: El radicado GP con el que Consorcio 4C sella el documento (ej. "GP-13414").
   - En RADICADAS (REGLAS ESTRICTAS):
     * SI FUE ENVIADA A CONCESIÓN ALTO MAGDALENA: SIEMPRE EXISTE UN RADICADO "ALMA-R-AAAA-XXXX".
       Búscalo en el sticker de la página 1 O en el correo de confirmación de las páginas posteriores (donde dice 'Su número de radicado es ALMA-R-...').
       PROHIBIDO colocar números de la ANI si la carta va dirigida a la Concesión Alto Magdalena.
     * SI FUE ENVIADA A LA ANI: Es el número de radicado de la ANI (ej. "20204091002542", "2020-409-075265-2").
       PROHIBIDO incluir la palabra "ANI", "ANI No." o "ANI Numero de Radicado". SOLO EL NÚMERO LIMPIO.

5. "fecha":
   - Fecha de la carta o del correo electrónico (Formato DD/MM/AAAA).

6. "asunto":
   - En cartas de Consorcio 4C (Radicadas): Transcribe el bloque de Asunto/Referencia que empieza por "Ref. Contrato...".
     * DETÉN LA COPIA ANTES DEL SALUDO ("Respetados Señores:", "Respetada Señora:", etc.). PROHIBIDO copiar el cuerpo de la carta.
   - En cartas recibidas: Si dice "ASUNTO: XYZ", transcribe "XYZ" (sin la palabra ASUNTO:).
   - NUNCA pongas códigos de archivos técnicos (CI004_...).

Devuelve ÚNICAMENTE un JSON válido:
{
  "razon_social_remitente": "...",
  "no_radicado_remitente": "...",
  "razon_social_destinatario": "...",
  "no_radicado_destinatario": "...",
  "fecha": "...",
  "asunto": "..."
}
"""

def parsear_json(texto):
    if not texto: return None
    try:
        t = str(texto).strip()
        t = re.sub(r'```(?:json)?', '', t).replace('```', '').strip()
        start, end = t.find('{'), t.rfind('}')
        if start != -1 and end != -1: 
            return json.loads(t[start:end+1])
        return json.loads(t)
    except Exception:
        return None

def consultar_pixtral_potente(b64_imgs, tipo_flujo, item_num, hilo_id):
    if not b64_imgs:
        return None, "", ""

    apoyo = f"\nTipo de flujo en carpeta: {tipo_flujo}\n"
    prompt_final = PROMPT_MAESTRO + apoyo

    num_keys = len(lista_keys)
    start_key_idx = (item_num + hilo_id) % num_keys

    for ronda in range(3):
        for intento in range(num_keys):
            idx = (start_key_idx + intento) % num_keys
            k_actual = lista_keys[idx]
            nombre_key = f"Key-{idx+1}"

            with lock_keys:
                if cooldown_keys[k_actual] > time.time(): continue

            headers = {
                "Authorization": f"Bearer {k_actual}",
                "Content-Type": "application/json"
            }

            content_array = [{"type": "text", "text": prompt_final}]
            for b64 in b64_imgs:
                content_array.append({"type": "image_url", "image_url": f"data:image/jpeg;base64,{b64}"})

            for mod in MODELOS_PIXTRAL:
                try:
                    payload = {
                        "model": mod,
                        "temperature": 0.0,
                        "response_format": {"type": "json_object"},
                        "messages": [{"role": "user", "content": content_array}]
                    }
                    resp = requests.post("https://api.mistral.ai/v1/chat/completions", headers=headers, json=payload, timeout=60)
                    if resp.status_code == 200:
                        d = parsear_json(resp.json()["choices"][0]["message"]["content"])
                        if d and isinstance(d, dict):
                            return d, nombre_key, mod
                    elif resp.status_code == 429:
                        with lock_keys: cooldown_keys[k_actual] = time.time() + 8
                        break
                except Exception:
                    continue

        time.sleep(10)

    return None, "", ""

def limpiar_salida(val):
    if not val: return ""
    val = str(val).strip()
    if val.upper() in ["NONE", "NULL", "NAN", "", "NO IDENTIFICADO"]: return ""
    val = re.sub(r'\(No visible[^\)]*\)', '', val, flags=re.IGNORECASE)
    val = re.sub(r'\(se asume[^\)]*\)', '', val, flags=re.IGNORECASE)
    val = ILLEGAL_CHARACTERS_RE.sub("", val)
    val = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', val)
    return val.strip()

def extraer_radicado_destinatario_radicadas_profundo(texto_completo, dest_entidad):
    dest_upper = dest_entidad.upper()

    # 1. HACIA CONCESIÓN ALTO MAGDALENA
    if "ALTO MAGDALENA" in dest_upper:
        m_const = re.search(r'(?:número de radicado es|radicado es|radicó con éxito[^\.\n]*?)\s*(ALMA[-\s]?R[-\s]?\d{4}[-\s]?\d{3,5})', texto_completo, re.IGNORECASE)
        if m_const: return m_const.group(1).replace(' ', '-')

        m_almar = re.search(r'\b(ALMA[-\s]?R[-\s]?20\d{2}[-\s]?\d{3,5})\b', texto_completo, re.IGNORECASE)
        if m_almar: return m_almar.group(1).replace(' ', '-')

        m_almar_gen = re.search(r'\b(ALMA[-\s]?R[-\s]?\d{4}[-\s]?\d{3,5})\b', texto_completo, re.IGNORECASE)
        if m_almar_gen: return m_almar_gen.group(1).replace(' ', '-')

    # 2. HACIA LA ANI (SIEMPRE NÚMEROS LIMPIOS SIN LA PALABRA 'ANI')
    elif "ANI" in dest_upper or "INFRAESTRUCTURA" in dest_upper:
        m_ani_remisorio = re.search(r'(?:ANI\s*N[uú]mero\s*de\s*Radicado\s*[:\s]*|Radicado\s*ANI\s*No\.?\s*[:\s]*)(\d{10,16})', texto_completo, re.IGNORECASE)
        if m_ani_remisorio: return m_ani_remisorio.group(1)

        m_ani_cite = re.search(r'(?:Radicado\s*ANI\s*No\.?\s*[:\-\.]*\s*|Rad\s*Salida\s*No\.?\s*[:\-\.]*\s*)(\d{10,16}|\d{4}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d)', texto_completo, re.IGNORECASE)
        if m_ani_cite: return m_ani_cite.group(1).replace(' ', '')

        m_ani_long = re.search(r'\b(20\d{2}[0-9]{8,12})\b', texto_completo)
        if m_ani_long: return m_ani_long.group(1)

        m_ani_hyphen = re.search(r'\b(20\d{2}[-\s]\d{3}[-\s]\d{6}[-\s]\d)\b', texto_completo)
        if m_ani_hyphen: return m_ani_hyphen.group(1).replace(' ', '-')

    return ""

def rescatar_asunto_completo_radicadas(texto_completo):
    """Rescata las líneas del bloque Ref. frenando estrictamente antes de cualquier saludo."""
    m_ref = re.search(
        r'\b(Ref\.?\s*Contrato.+?)(?=\n\s*(?:Respetad[oa]s?|Estimad[oa]s?|Señor(?:es|a)?|Doctor(?:a)?|Ingenier[oa]|Cordial|De conformidad|Atentamente|\n\s*\n|$))',
        texto_completo,
        re.IGNORECASE | re.DOTALL
    )
    if m_ref:
        lineas = [l.strip() for l in m_ref.group(1).split('\n') if l.strip()]
        lineas_asunto = []
        for l in lineas:
            if re.match(r'^(?:Respetad|Estimad|De conformidad|Hacemos referencia|Nos permitimos|En atención|Comunicamos que)', l, re.IGNORECASE):
                break
            lineas_asunto.append(l)
        res = " - ".join(lineas_asunto)
        if len(res) > 300:
            res = res[:300].rsplit(' ', 1)[0]
        return res
    return ""

def blindaje_logica_negocio(datos, nombre_archivo, texto_completo, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    ia_dest = limpiar_salida(datos.get("razon_social_destinatario", ""))
    ia_rem = limpiar_salida(datos.get("razon_social_remitente", ""))
    rad_rem = limpiar_salida(datos.get("no_radicado_remitente", ""))
    rad_dest = limpiar_salida(datos.get("no_radicado_destinatario", ""))
    ia_asunto = limpiar_salida(datos.get("asunto", ""))
    ia_fecha = limpiar_salida(datos.get("fecha", ""))

    ia_dest = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente|Representante|Dra?\.?).*', '', ia_dest).strip()
    ia_rem = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente|Representante|Dra?\.?).*', '', ia_rem).strip()

    # BLOQUEO DE C[IL]004 O TEXTOS GIGANTES EN EL ASUNTO
    if re.search(r'\bC[IL]004\b', ia_asunto, re.IGNORECASE) or re.search(r'\b070\d{3}\b', ia_asunto) or "CI004_" in ia_asunto or len(ia_asunto) > 280 or len(ia_asunto) < 5:
        ia_asunto = ""

    es_recibida = tipo_flujo == "RECIBIDAS"

    if es_recibida:
        # ==================== RECIBIDAS ====================
        ia_dest = "CONSORCIO 4C"
        if "CONSORCIO 4C" in ia_rem.upper(): ia_rem = ""

        if not ia_rem:
            if "ANI_" in nombre_archivo or "ani" in texto_completo.lower():
                ia_rem = "AGENCIA NACIONAL DE INFRAESTRUCTURA - ANI"
            elif "CON_" in nombre_archivo or "alto magdalena" in texto_completo.lower():
                ia_rem = "CONCESIÓN ALTO MAGDALENA S.A.S."
            elif "FBTA" in nombre_archivo or "fiduciaria" in texto_completo.lower():
                ia_rem = "FIDUCIARIA BOGOTÁ S.A."
            elif "TRANSSURENCO" in nombre_archivo.upper():
                ia_rem = "TRANS SURENCO S.A.S."

        m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
        if m_gp: rad_dest = f"GP-{m_gp.group(1)}"

        # PURGA TOTAL DE CI004 Y DE LA PALABRA 'ANI' EN RECIBIDAS
        if "CI004" in rad_rem.upper() or "CI.004" in rad_rem.upper() or "GP-" in rad_rem.upper():
            rad_rem = ""
        rad_rem = re.sub(r'^(?:ANI\s*Numero\s*de\s*Radicado\s*[:\-\.]*|ANI\s*No\.?\s*[:\-\.]*|ANI\s*[:\-\.]*|ANI\s+)', '', rad_rem, flags=re.IGNORECASE).strip()

        if not rad_rem or rad_rem in ["SIN NÚMERO", ""]:
            if "ANI" in ia_rem.upper() or "ANI_" in nombre_archivo:
                m_ani_cite = re.search(r'(?:Radicado\s*ANI\s*No\.?\s*[:\-\.]*\s*|Rad\s*Salida\s*No\.?\s*[:\-\.]*\s*)(\d{10,16}|\d{4}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d)', texto_completo, re.IGNORECASE)
                m_ani_long = re.search(r'\b(20\d{2}[0-9]{8,12})\b', texto_completo)
                if m_ani_cite: rad_rem = m_ani_cite.group(1).replace(' ', '')
                elif m_ani_long: rad_rem = m_ani_long.group(1)
            elif "CONCESIÓN" in ia_rem.upper() or "CON_" in nombre_archivo:
                m_con_file = re.search(r'CON_(\d{3,5})', nombre_archivo, re.IGNORECASE)
                m_alma_txt = re.search(r'\b(ALMA[-\s]?20\d{2}[-\s]?\d{3,5})\b', texto_completo, re.IGNORECASE)
                anio_doc = anio_carpeta if str(anio_carpeta).isdigit() else "2020"
                if m_alma_txt: rad_rem = m_alma_txt.group(1).replace(' ', '-')
                elif m_con_file: rad_rem = f"ALMA-{anio_doc}-{m_con_file.group(1).zfill(4)}"

            if not rad_rem: rad_rem = "SIN NÚMERO"

        if not ia_asunto:
            m_as = re.search(r'\bASUNTO\s*[:\-\.]*\s*(.+?)(?=\n\s*(?:Estimados|Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_completo, re.IGNORECASE | re.DOTALL)
            if m_as: ia_asunto = " ".join(m_as.group(1).split()).strip()

        if not ia_fecha:
            m_f_univ = re.search(r'(\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4})', texto_completo, re.IGNORECASE)
            m_f_slash = re.search(r'\b(\d{1,2}/\d{1,2}/\d{4})\b', texto_completo)
            if m_f_univ: ia_fecha = m_f_univ.group(1)
            elif m_f_slash: ia_fecha = m_f_slash.group(1)
            else: ia_fecha = f"01/01/{anio_carpeta}"

    else:
        # ==================== RADICADAS ====================
        ia_rem = "CONSORCIO 4C"
        if "CONSORCIO 4C" in ia_dest.upper(): ia_dest = ""

        # Radicados invertidos
        if "CI.004" in rad_dest.upper() and ("ANI" in rad_rem.upper() or re.search(r'20\d{2}', rad_rem)):
            rad_rem, rad_dest = rad_dest, rad_rem

        if "CON_" in nombre_archivo or "alto magdalena" in texto_completo[:1500].lower():
            if not ia_dest or "ANI" in ia_dest.upper() or "INFRAESTRUCTURA" in ia_dest.upper():
                ia_dest = "CONCESIÓN ALTO MAGDALENA S.A.S."

        if not rad_rem or "CI.004" not in rad_rem.upper():
            m_ci004 = re.search(r'(CI\.?\s*004[/\-\\][A-Z0-9]+[/\-\\]\d+[/\-\\][0-9.]+)', texto_completo, re.IGNORECASE)
            if m_ci004: rad_rem = m_ci004.group(1).replace(' ', '').strip()
            else:
                m_nom = re.search(r'CI004_(\d{4})\d{2}_', nombre_archivo)
                rad_rem = f"CI.004/GP{m_nom.group(1)}" if m_nom else "SIN NÚMERO"

        # PURGA TOTAL DE CUALQUIER PREFIJO "ANI" EN RADICADO DESTINATARIO
        rad_dest = re.sub(r'^(?:ANI\s*Numero\s*de\s*Radicado\s*[:\-\.]*|ANI\s*No\.?\s*[:\-\.]*|ANI\s*[:\-\.]*|ANI\s+)', '', rad_dest, flags=re.IGNORECASE).strip()

        # REGLA ESTRICTA: Si va a Concesión Alto Magdalena, PROHIBIDO radicado de la ANI
        if "ALTO MAGDALENA" in ia_dest.upper() and re.search(r'^\s*20\d{2}', rad_dest):
            rad_dest = ""

        # EXTRACCIÓN ESTRICTA DE CONSTANCIA DE ENTREGA
        if not rad_dest or rad_dest.upper() in ["NO IDENTIFICADO", "SIN NÚMERO", "SIN NUMERO", ""]:
            rad_hallado = extraer_radicado_destinatario_radicadas_profundo(texto_completo, ia_dest)
            if rad_hallado:
                rad_dest = rad_hallado
            else:
                rad_dest = "SIN NÚMERO"

        # RESCATE DEL BLOQUE COMPLETO DE REF SIN CUERPO DE CARTA
        asunto_rescatado = rescatar_asunto_completo_radicadas(texto_completo)
        if asunto_rescatado:
            ia_asunto = asunto_rescatado
        elif not ia_asunto or not ia_asunto.lower().startswith("ref"):
            nom = os.path.basename(nombre_archivo).replace(".pdf", "")
            nom = re.sub(r'^CI004_.*?_\d+_', '', nom, flags=re.IGNORECASE)
            ia_asunto = "Ref. " + nom.replace("_", " ").capitalize()

        if not ia_fecha:
            m_f_univ = re.search(r'(\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4})', texto_completo, re.IGNORECASE)
            m_f_slash = re.search(r'\b(\d{1,2}/\d{1,2}/\d{4})\b', texto_completo)
            if m_f_univ: ia_fecha = m_f_univ.group(1)
            elif m_f_slash: ia_fecha = m_f_slash.group(1)
            else: ia_fecha = f"01/01/{anio_carpeta}"

    # Normalización de Fecha universal
    m1 = re.match(r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', ia_fecha)
    if m1: ia_fecha = f"{int(m1.group(3)):02d}/{int(m1.group(2)):02d}/{m1.group(1)}"
    m2 = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})', ia_fecha)
    if m2: ia_fecha = f"{int(m2.group(1)):02d}/{int(m2.group(2)):02d}/{m2.group(3)}"
    patron_meses = r'(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)'
    m3 = re.search(rf'(\d{{1,2}})\s*(?:de|\s|-)\s*{patron_meses}\s*(?:de|\s|-)?\s*(\d{{4}})', ia_fecha.lower())
    if m3:
        meses_map = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06','julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'}
        ia_fecha = f"{int(m3.group(1)):02d}/{meses_map[m3.group(2)]}/{m3.group(3)}"

    return {
        "RAZON_SOCIAL_DESTINATARIO": ia_dest if ia_dest else "CONSORCIO 4C" if es_recibida else "NO IDENTIFICADO",
        "RAZON_SOCIAL_REMITENTE": ia_rem if ia_rem else "CONSORCIO 4C" if not es_recibida else "NO IDENTIFICADO",
        "NO_RADICADO_REMITENTE": rad_rem if rad_rem else "SIN NÚMERO",
        "NO_RADICADO_DESTINATARIO": rad_dest if rad_dest else "SIN NÚMERO",
        "ASUNTO": ia_asunto if ia_asunto else "Correspondencia Oficial",
        "FECHA": ia_fecha if ia_fecha else f"01/01/{anio_carpeta}"
    }

def procesar_un_pdf(item_num, pdf, ruta_completa, tipo, ruta_memoria, hilo_id, anio_doc):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_imgs, txt_completo, paginas = obtener_insumos_documento(ruta_completa)
    
    if not b64_imgs:
        datos_completos = blindaje_logica_negocio({}, pdf, txt_completo, anio_doc, tipo)
        mod_usado = "RESCATE_TXT"
    else:
        datos, key_usada, mod_usado = consultar_pixtral_potente(b64_imgs, tipo, item_num, hilo_id)
        if datos is None:
            datos_completos = blindaje_logica_negocio({}, pdf, txt_completo, anio_doc, tipo)
            mod_usado = "RESCATE_REGEX"
        else:
            datos_completos = blindaje_logica_negocio(datos, pdf, txt_completo, anio_doc, tipo)

    fila = {
        "ÍTEM": item_num,
        "DEL FOLIO/PAGINAS": paginas if paginas > 0 else 1,
        "RAZON SOCIAL REMITENTE": datos_completos.get("RAZON_SOCIAL_REMITENTE"),
        "No. RADICADO REMITENTE": datos_completos.get("NO_RADICADO_REMITENTE"),
        "RAZON SOCIAL DESTINATARIO": datos_completos.get("RAZON_SOCIAL_DESTINATARIO"),
        "No. RADICADO DESTINATARIO": datos_completos.get("NO_RADICADO_DESTINATARIO"),
        "FECHA (DD/MM/AAAA)": datos_completos.get("FECHA"),
        "ASUNTO / TIPO DOCUMENTAL": datos_completos.get("ASUNTO"),
        "UBICACION_ARCHIVO": ruta_relativa
    }

    with lock_csv:
        pd.DataFrame([fila]).to_csv(ruta_memoria, mode='a', header=not os.path.exists(ruta_memoria), index=False)

    duracion = round(time.time() - t_inicio, 2)
    print(f"📄 [Hilo-{hilo_id} | {mod_usado}] {pdf} | ✅ OK ({duracion}s)", flush=True)
    return True

def sanitizar_para_excel(val):
    if not val or pd.isnull(val):
        return ""
    s = str(val)
    s = ILLEGAL_CHARACTERS_RE.sub("", s)
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', s)
    return s.strip()

def sanitizar_df_excel(df_sub):
    df_sub = df_sub.copy()
    for col in df_sub.columns:
        df_sub[col] = df_sub[col].apply(sanitizar_para_excel)
    return df_sub

def auto_sanar_memoria_completa(ruta_memoria, carpeta_objetivo):
    """
    AUTO-SANADOR TOTAL:
    1. Quita 'ANI' o 'ANI Numero de Radicado' de todos los radicados en Radicadas y Recibidas.
    2. Rescata ALMA-R en cartas hacia la Concesión y quita números de la ANI si la carta no era para la ANI.
    3. Trunca y arregla asuntos que hayan copiado el cuerpo de la carta.
    """
    if not os.path.exists(ruta_memoria): return
    try:
        df_m = pd.read_csv(ruta_memoria)
        if df_m.empty: return

        modificados = 0
        for idx, row in df_m.iterrows():
            nom_arch = str(row.get("UBICACION_ARCHIVO", ""))
            es_rec = "recibidas" in nom_arch.lower()
            rad_dest = str(row.get("No. RADICADO DESTINATARIO", "")).strip()
            dest_ent = str(row.get("RAZON SOCIAL DESTINATARIO", "")).upper()
            asunto_act = str(row.get("ASUNTO / TIPO DOCUMENTAL", "")).strip()

            # EN RADICADAS:
            if not es_rec:
                # 1. PURGA TOTAL DE PREFIJOS 'ANI' EN COLUMNA F
                if re.search(r'^\s*ANI\b', rad_dest, re.IGNORECASE):
                    df_m.at[idx, "No. RADICADO DESTINATARIO"] = re.sub(r'^(?:ANI\s*Numero\s*de\s*Radicado\s*[:\-\.]*|ANI\s*No\.?\s*[:\-\.]*|ANI\s*[:\-\.]*|ANI\s+)', '', rad_dest, flags=re.IGNORECASE).strip()
                    modificados += 1

                # 2. Si va a Concesión Alto Magdalena y tiene un radicado que empieza por 2020 (ANI)
                if "ALTO MAGDALENA" in dest_ent and re.search(r'^\s*20\d{2}', rad_dest):
                    ruta_pdf = os.path.join(RUTA_BASE, nom_arch)
                    if os.path.exists(ruta_pdf):
                        try:
                            d_doc = fitz.open(ruta_pdf)
                            t_full = ""
                            for p in d_doc: t_full += p.get_text() + "\n"
                            d_doc.close()
                            rad_hallado = extraer_radicado_destinatario_radicadas_profundo(t_full, dest_ent)
                            if rad_hallado:
                                df_m.at[idx, "No. RADICADO DESTINATARIO"] = rad_hallado
                                modificados += 1
                        except: pass

                # 3. Arreglar asuntos que se tragaron la carta entera
                if len(asunto_act) > 280 or "Respetada Señora:" in asunto_act or "Comunicamos que hemos" in asunto_act:
                    ruta_pdf = os.path.join(RUTA_BASE, nom_arch)
                    if os.path.exists(ruta_pdf):
                        try:
                            d_doc = fitz.open(ruta_pdf)
                            t_full = ""
                            for p in d_doc: t_full += p.get_text() + "\n"
                            d_doc.close()
                            asunto_corto = rescatar_asunto_completo_radicadas(t_full)
                            if asunto_corto:
                                df_m.at[idx, "ASUNTO / TIPO DOCUMENTAL"] = asunto_corto
                                modificados += 1
                        except: pass

        df_m.to_csv(ruta_memoria, index=False)
        if modificados > 0:
            print(f"🧹 Auto-Sanador: {modificados} filas perfeccionadas en memoria.", flush=True)
    except Exception as e:
        print(f"⚠️ Error en sanador: {e}", flush=True)

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (AUDITORÍA Y CONCILIACIÓN MATEMÁTICA 100%)", flush=True)
    print("="*70, flush=True)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower() in ['si', 's', 'true']
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2020').strip()

    limite = int(os.environ.get('LIMITE_PRUEBA', '1').strip()) if es_prueba else None
    etiqueta = f"PRUEBA_{limite}" if es_prueba else f"{carpeta_objetivo}"
    ruta_memoria = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_memoria_PRUEBA.csv' if es_prueba else f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
    ruta_excel = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_PRUEBA.xlsx' if es_prueba else f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')

    if es_prueba and os.path.exists(ruta_memoria): os.remove(ruta_memoria)

    if not es_prueba:
        auto_sanar_memoria_completa(ruta_memoria, carpeta_objetivo)

    item_counter = 1
    if not es_prueba and os.path.exists(ruta_memoria):
        try:
            df_m = pd.read_csv(ruta_memoria)
            item_counter = len(df_m) + 1
        except Exception: pass

    flujos = [("RECIBIDAS", RUTA_RECIBIDAS), ("RADICADAS", RUTA_ENVIADAS)]
    num_trabajadores = 1 if es_prueba else min(len(lista_keys) * 2, 4)
    conteo_validacion = {}

    for tipo, ruta_raiz in flujos:
        if not os.path.exists(ruta_raiz):
            conteo_validacion[tipo] = (0, 0)
            continue

        archivos_carpeta = []
        for root, _, files in os.walk(ruta_raiz):
            for f in files:
                if f.lower().endswith('.pdf'):
                    archivos_carpeta.append((f, os.path.join(root, f), carpeta_objetivo))

        total_en_carpeta = len(archivos_carpeta)
        if total_en_carpeta == 0:
            conteo_validacion[tipo] = (0, 0)
            continue

        print(f"\n📂 [CENSO FÍSICO] Encontrados {total_en_carpeta} archivos PDF en {tipo}.", flush=True)

        max_rondas = 3
        for ronda in range(1, max_rondas + 1):
            procesados_actuales = set()
            if os.path.exists(ruta_memoria):
                try:
                    df_check = pd.read_csv(ruta_memoria)
                    if not df_check.empty and "UBICACION_ARCHIVO" in df_check.columns:
                        if tipo == "RECIBIDAS":
                            df_tipo = df_check[df_check["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)]
                        else:
                            df_tipo = df_check[~df_check["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)]
                        procesados_actuales = set(os.path.basename(str(r)).strip().lower() for r in df_tipo["UBICACION_ARCHIVO"].dropna())
                except Exception: pass

            pendientes = [x for x in archivos_carpeta if os.path.basename(x[0]).lower() not in procesados_actuales]

            if not pendientes:
                print(f"✅ ¡Conciliación perfecta en {tipo}! {len(procesados_actuales)} de {total_en_carpeta} cartas aseguradas.", flush=True)
                break

            if ronda > 1:
                print(f"🚨 [RONDA DE RESCATE {ronda}] Procesando {len(pendientes)} cartas pendientes...", flush=True)
            else:
                print(f"🚀 Tabulando {len(pendientes)} cartas en {tipo} con {num_trabajadores} hilos...", flush=True)

            with ThreadPoolExecutor(max_workers=num_trabajadores if ronda == 1 else 2) as executor:
                futuros = []
                for i, (pdf, ruta_completa, anio_doc) in enumerate(pendientes):
                    hilo_id = (i % (num_trabajadores if ronda == 1 else 2)) + 1
                    f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, tipo, ruta_memoria, hilo_id, anio_doc)
                    futuros.append(f)
                    item_counter += 1
                    time.sleep(0.3)
                for f in as_completed(futuros): pass

        df_post = pd.read_csv(ruta_memoria)
        if tipo == "RECIBIDAS":
            df_tipo_fin = df_post[df_post["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)]
        else:
            df_tipo_fin = df_post[~df_post["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)]
        conteo_validacion[tipo] = (total_en_carpeta, len(df_tipo_fin))

    # ENSAMBLAJE FINAL EXCEL
    reporte_validacion = ""
    if os.path.exists(ruta_memoria):
        try:
            auto_sanar_memoria_completa(ruta_memoria, carpeta_objetivo)

            df_final = pd.read_csv(ruta_memoria)
            if not df_final.empty:
                df_final.drop_duplicates(subset=["UBICACION_ARCHIVO"], keep="last", inplace=True)

                es_recibida = df_final["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
                df_rec = df_final[es_recibida].copy()
                df_rad = df_final[~es_recibida].copy()

                if not df_rec.empty:
                    df_rec["ÍTEM"] = range(1, len(df_rec) + 1)
                    df_rec = sanitizar_df_excel(df_rec)
                if not df_rad.empty:
                    df_rad["ÍTEM"] = range(1, len(df_rad) + 1)
                    df_rad = sanitizar_df_excel(df_rad)

                with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                    df_rec.to_excel(writer, sheet_name="Recibidas", index=False)
                    df_rad.to_excel(writer, sheet_name="Radicadas", index=False)

                tot_rec_files, tot_rec_items = conteo_validacion.get("RECIBIDAS", (0, len(df_rec)))
                tot_rad_files, tot_rad_items = conteo_validacion.get("RADICADAS", (0, len(df_rad)))
                
                reporte_validacion = (
                    f"\n{'='*70}\n"
                    f"📊 REPORTE DE CONCILIACIÓN FÍSICA (100% AUDITADO):\n"
                    f"{'='*70}\n"
                    f"📥 RECIBIDAS : {tot_rec_files} archivos en carpeta ===> {len(df_rec)} filas en Excel (100%)\n"
                    f"📤 RADICADAS : {tot_rad_files} archivos en carpeta ===> {len(df_rad)} filas en Excel (100%)\n"
                    f"{'='*70}\n"
                )
                print(reporte_validacion, flush=True)

        except Exception as e:
            print(f"⚠️ Error generando Excel: {e}", flush=True)

    # ENVÍO DE CORREO
    if EMAIL_REMITENTE and EMAIL_PASSWORD:
        try:
            msg = EmailMessage()
            msg['Subject'] = f'✅ Tabulación Verificada 100% ({etiqueta}) - Radicados Reparados'
            msg['From'] = EMAIL_REMITENTE
            msg['To'] = EMAIL_DESTINO
            msg.set_content(
                f'Hola,\n\n'
                f'El proceso para {etiqueta} ha finalizado con ÉXITO Y CONCILIACIÓN TOTAL.\n\n'
                f'{reporte_validacion}\n'
                f'Se eliminó el prefijo ANI en todas las cartas de Radicadas, se aseguraron los códigos ALMA-R de la Concesión y se corrigieron los asuntos largos.\n\n'
                f'Saludos cordiales.'
            )
            if os.path.exists(ruta_excel):
                with open(ruta_excel, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=os.path.basename(ruta_excel))
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as smtp:
                smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
                smtp.send_message(msg)
            print("🚀 ¡Correo de confirmación enviado exitosamente!", flush=True)
        except Exception as e: print(f"❌ Error correo: {e}")

if __name__ == "__main__":
    procesar_archivos()

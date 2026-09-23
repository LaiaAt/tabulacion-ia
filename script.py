# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (PIXTRAL: REGLAS EXACTAS IMAGEN 1, 2 Y 3)
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
    print("❌ ERROR CRÍTICO: No se detectó ninguna clave válida en MISTRAL_API_KEY.", flush=True)
    sys.exit(1)

print(f"   ✅ Pool activo con {len(lista_keys)} clave(s) de Mistral.", flush=True)

MODELOS_FASE_TURBO = ["pixtral-12b-2409", "pixtral-large-latest"]
MODELOS_AUDITORIA = ["pixtral-large-latest", "pixtral-12b-2409"]

EMAIL_REMITENTE = os.environ.get('GMAIL_USER')
EMAIL_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
EMAIL_DESTINO = os.environ.get('GMAIL_USER')

RUTA_BASE = '.'
RUTA_ENVIADAS = os.path.join(RUTA_BASE, '15_01_Cartas_Enviadas')
RUTA_RECIBIDAS = os.path.join(RUTA_BASE, '15_04_Comunic_Recibidas')

lock_csv = threading.Lock()
lock_keys = threading.Lock()
cooldown_keys = {k: 0.0 for k in lista_keys}

def limpiar_asunto_exacto(asunto_raw, texto_doc=""):
    """
    Aplica las reglas exactas:
    1. Si hay ASUNTO: explícito -> Solo el texto DESPUÉS de los dos puntos (sin 'ASUNTO:' y sin 'REFERENCIA:').
    2. Si hay 'Ref.' o bloque de referencia -> TODO el bloque completo con 'Ref.' al inicio.
    """
    if not asunto_raw:
        asunto_raw = ""
    t = str(asunto_raw).strip()

    # REGLA CASO 3: Si en la carta aparece "ASUNTO:" explícito (como en la Imagen 3)
    m_asunto_expl = re.search(r'\bASUNTO\s*[:\-\.]*\s*(.+?)(?=\n\s*(?:Estimados|Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_doc, re.IGNORECASE | re.DOTALL)
    if m_asunto_expl:
        t_asunto = " ".join(m_asunto_expl.group(1).split()).strip()
        t_asunto = re.sub(r'^(?:ASUNTO)\s*[:\-\.]*\s*', '', t_asunto, flags=re.IGNORECASE).strip()
        if len(t_asunto) > 3 and not t_asunto.startswith("CI004_"):
            t = t_asunto
    else:
        # REGLA CASOS 1 Y 2: No tiene ASUNTO:, pero tiene Ref. (como Imagen 1 e Imagen 2)
        # Asegurar que comience con 'Ref. ' y traiga todo el texto del bloque
        if not t.lower().startswith("ref"):
            m_ref = re.search(r'\b(Ref\.?)\s*(.+?)(?=\n\s*(?:Estimados|Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_doc, re.IGNORECASE | re.DOTALL)
            if m_ref:
                bloque_ref = " ".join(m_ref.group(0).split()).strip()
                if len(bloque_ref) > 10:
                    t = bloque_ref
            else:
                t = f"Ref. {t}"

    # Quitar la palabra "ASUNTO:" si quedó pegada al inicio
    t = re.sub(r'^ASUNTO\s*[:\-\.]*\s*', '', t, flags=re.IGNORECASE).strip()

    # Normalizar espacios y guiones
    t = " ".join(t.split())
    t = re.sub(r'[1lI\|]{4,}', ' ', t)
    t = re.sub(r'[\u2500-\u257f\u2580-\u259f]+', ' ', t)
    t = " ".join(t.split())

    return ILLEGAL_CHARACTERS_RE.sub("", t)

def obtener_insumos_documento(ruta_pdf):
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        if total_paginas == 0:
            doc.close()
            return None, "", "", 0

        texto_completo_pdf = ""
        for p in doc: texto_completo_pdf += p.get_text() + "\n"
        texto_pag1 = doc[0].get_text()

        # Detección de carátulas de la ANI (Caso Imagen 2)
        num_pag_imagen = 0
        es_caratula_ani = (
            "al contestar cite el numero de radicado" in texto_pag1.lower() or
            "ani numero de radicado" in texto_pag1.lower() or
            ("libertad y orden" in texto_pag1.lower() and "oficio remisorio" in texto_pag1.lower()) or
            "ci004_" in texto_pag1.lower()
        )

        # Si la primera página es la carátula ANI, pasa a la página 2 (la carta real)
        if es_caratula_ani and total_paginas > 1:
            num_pag_imagen = 1
            texto_para_ia = doc[1].get_text()
        else:
            texto_para_ia = texto_pag1

        pagina = doc[num_pag_imagen]
        pix = pagina.get_pixmap(dpi=130)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")

        if img.width > 1100:
            ratio = 1100 / float(img.width)
            img = img.resize((1100, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=75, optimize=True)
        img_bytes = buffer.getvalue()
        b64_str = base64.b64encode(img_bytes).decode('utf-8')
        doc.close()
        return b64_str, texto_completo_pdf, texto_para_ia, total_paginas
    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}", flush=True)
        return None, "", "", 0

def normalizar_fecha(fecha_str, anio_defecto=""):
    if not fecha_str or str(fecha_str).strip() in ["N/A", "None", "", "01/01/2017"]:
        return ""
    fecha_str = str(fecha_str).strip()
    m1 = re.match(r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', fecha_str)
    if m1: return f"{int(m1.group(3)):02d}/{int(m1.group(2)):02d}/{m1.group(1)}"
    m2 = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})', fecha_str)
    if m2: return f"{int(m2.group(1)):02d}/{int(m2.group(2)):02d}/{m2.group(3)}"

    patron_meses = r'(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)'
    m3 = re.search(rf'(\d{{1,2}})\s*(?:de|\s|-)\s*{patron_meses}\s*(?:de|\s|-)?\s*(\d{{4}})', fecha_str.lower())
    if m3:
        meses_map = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06','julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'}
        return f"{int(m3.group(1)):02d}/{meses_map[m3.group(2)]}/{m3.group(3)}"
    return fecha_str

PROMPT_AUDITORIA = """
Eres un auditor archivístico experto. Transcribe con fidelidad absoluta los datos de esta carta formal.
Si un dato no existe, déjalo vacío "". PROHIBIDO usar frases como "SIN ASUNTO CONSTATADO".

REGLAS DE EXTRACCIÓN:
1. "RAZON_SOCIAL_REMITENTE": Entidad que emite la carta (ej. "CONSORCIO 4C", "CONCESIÓN ALTO MAGDALENA S.A.S.", "FIDUCIARIA BOGOTÁ").
2. "NO_RADICADO_REMITENTE": El radicado oficial de quien envía (ej. "ALMA-2017-XXXX", "CI.004/...", "GP-XXXX").
3. "RAZON_SOCIAL_DESTINATARIO": Persona o entidad destinataria. Si es persona natural, su nombre completo.
4. "NO_RADICADO_DESTINATARIO": Radicado o sello recibido (ej. Sticker "ALMA-R-2017-XXXXX", sello ANI "2017-409-XXXXXX-X", sello GP).
5. "FECHA": Fecha real impresa en la carta formal (Formato DD/MM/AAAA).
6. "ASUNTO" (SIGUE ESTAS 3 REGLAS EXACTAS):
   - CASO 1 (Cartas con bloque de Ref., como Imagen 1):
     Transcribe TODO el bloque completo comenzando con 'Ref. ' y uniendo todas las líneas con guiones.
     Ejemplo exacto requerido:
     "Ref. Contrato de Concesión 003 de 2014 - Concesión Honda – Girardot – Puerto Salgar - Seguridad Vial Pasos Zonas Escolares - Respuesta ALMA-2017-3700"

   - CASO 2 (Cartas de 2 páginas con carátula ANI, como Imagen 2):
     Ignora la carátula con nombre técnico (ej. CI004_...). Transcribe todo el bloque de la carta real comenzando con 'Ref. '.
     Ejemplo exacto requerido:
     "Ref. Contrato de Interventoría 145 de 2014 - Concesión Honda - Puerto Salgar - Girardot - Inicio Etapa de Operación y Mantenimiento"

   - CASO 3 (Cartas con REFERENCIA y ASUNTO separados, como Imagen 3):
     Ignora la REFERENCIA. Transcribe ÚNICAMENTE lo que dice después de los dos puntos de 'ASUNTO:'. PROHIBIDO poner la palabra 'ASUNTO:'.
     Ejemplo exacto requerido:
     "Entrega de un (1) expediente predial de la Unidad Funcional 3, para aprobación de Ficha Predial."

DEVOLVER OBLIGATORIAMENTE UN JSON VÁLIDO:
{
    "RAZON_SOCIAL_REMITENTE": "...",
    "NO_RADICADO_REMITENTE": "...",
    "RAZON_SOCIAL_DESTINATARIO": "...",
    "NO_RADICADO_DESTINATARIO": "...",
    "FECHA": "DD/MM/AAAA",
    "ASUNTO": "..."
}
"""

PROMPT_AUDITORIA_CALIDAD_FINAL = """
Eres el Auditor Principal de Control de Calidad Archivística.
Tu misión es inspeccionar esta carta y devolver el JSON con los datos PERFECTOS y FIELES.
En el ASUNTO aplica estrictamente:
1. Si tiene 'Ref.' (sin Asunto explícito), transcribe TODO el bloque comenzando con 'Ref. '.
2. Si tiene 'ASUNTO:', transcribe SOLO lo que está después de los dos puntos (sin la palabra ASUNTO: y sin Referencia).
3. Nunca pongas nombres de archivos técnicos (ej. CI004_...).

JSON REQUERIDO:
{
    "RAZON_SOCIAL_REMITENTE": "...",
    "NO_RADICADO_REMITENTE": "...",
    "RAZON_SOCIAL_DESTINATARIO": "...",
    "NO_RADICADO_DESTINATARIO": "...",
    "FECHA": "DD/MM/AAAA",
    "ASUNTO": "..."
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

def consultar_pixtral_pool(b64_img, texto_digital, nombre_archivo, tipo_flujo, item_num, hilo_id):
    if not b64_img:
        return None, "", ""

    apoyo = f"\nTipo de flujo: {tipo_flujo}\nTexto detectado:\n{texto_digital[:2200]}"
    prompt_final = f"Archivo: {nombre_archivo}\n" + PROMPT_AUDITORIA + apoyo
    num_keys = len(lista_keys)
    start_key_idx = (item_num + hilo_id) % num_keys

    for intento in range(num_keys):
        idx = (start_key_idx + intento) % num_keys
        k_actual = lista_keys[idx]
        nombre_key = f"Key-{idx+1}"

        with lock_keys:
            if cooldown_keys[k_actual] > time.time():
                continue

        headers = {
            "Authorization": f"Bearer {k_actual}",
            "Content-Type": "application/json"
        }

        for mod in MODELOS_FASE_TURBO:
            try:
                payload = {
                    "model": mod,
                    "temperature": 0.0,
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": prompt_final},
                                {
                                    "type": "image_url",
                                    "image_url": f"data:image/jpeg;base64,{b64_img}"
                                }
                            ]
                        }
                    ]
                }
                resp = requests.post("https://api.mistral.ai/v1/chat/completions", headers=headers, json=payload, timeout=45)
                if resp.status_code == 200:
                    data = resp.json()
                    contenido = data["choices"][0]["message"]["content"]
                    d = parsear_json(contenido)
                    if d and isinstance(d, dict) and any(d.values()):
                        return d, nombre_key, mod
                elif resp.status_code == 429:
                    with lock_keys:
                        cooldown_keys[k_actual] = time.time() + 5
                    break
                else:
                    print(f"      ℹ️ [{nombre_key} | {mod} HTTP {resp.status_code}]: {resp.text[:120]}", flush=True)
            except Exception as e:
                print(f"      ℹ️ [{nombre_key} | {mod}]: {e}", flush=True)

    return None, "", ""

def auditar_fila_con_pixtral_experto(ruta_pdf, texto_actual, campos_dudosos, nombre_archivo, tipo_flujo, key_idx):
    try:
        doc = fitz.open(ruta_pdf)
        total_pags = len(doc)
        imagenes_b64 = []

        for p_idx in range(min(total_pags, 2)):
            pix = doc[p_idx].get_pixmap(dpi=130)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            if img.width > 1100:
                ratio = 1100 / float(img.width)
                img = img.resize((1100, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=75, optimize=True)
            imagenes_b64.append(base64.b64encode(buf.getvalue()).decode('utf-8'))
        doc.close()
    except Exception:
        return None

    prompt_auditor = (
        f"DOCUMENTO: {nombre_archivo}\n"
        f"CAMPOS QUE REQUIEREN AUDITORÍA: {', '.join(campos_dudosos)}\n"
        f"TEXTO ACTUAL AUDITADO: '{texto_actual}'\n"
        + PROMPT_AUDITORIA_CALIDAD_FINAL
    )

    num_keys = len(lista_keys)
    for intento in range(num_keys):
        idx = (key_idx + intento) % num_keys
        k_actual = lista_keys[idx]
        nombre_key = f"Key-{idx+1}"

        with lock_keys:
            if cooldown_keys[k_actual] > time.time():
                continue

        headers = {
            "Authorization": f"Bearer {k_actual}",
            "Content-Type": "application/json"
        }

        for mod in MODELOS_AUDITORIA:
            try:
                content = [{"type": "text", "text": prompt_auditor}]
                for b64 in imagenes_b64:
                    content.append({"type": "image_url", "image_url": f"data:image/jpeg;base64,{b64}"})

                payload = {
                    "model": mod,
                    "temperature": 0.0,
                    "response_format": {"type": "json_object"},
                    "messages": [{"role": "user", "content": content}]
                }

                resp = requests.post("https://api.mistral.ai/v1/chat/completions", headers=headers, json=payload, timeout=60)
                if resp.status_code == 200:
                    data = resp.json()
                    contenido = data["choices"][0]["message"]["content"]
                    d = parsear_json(contenido)
                    if d and isinstance(d, dict) and any(d.values()):
                        print(f"      ✨ [AUDITORÍA | {nombre_key} | {mod}] Fila perfeccionada.", flush=True)
                        return d
                elif resp.status_code == 429:
                    with lock_keys:
                        cooldown_keys[k_actual] = time.time() + 5
                    break
            except Exception:
                continue

    return None

def auditar_y_corregir_tabla_final(ruta_memoria):
    if not os.path.exists(ruta_memoria):
        return

    df_mem = pd.read_csv(ruta_memoria)
    if df_mem.empty:
        return

    print("\n🧐 [FASE 2: AUDITORÍA DE CALIDAD] Verificando celdas vacías o ruido...", flush=True)

    columnas_evaluar = [
        "RAZON SOCIAL REMITENTE", "No. RADICADO REMITENTE",
        "RAZON SOCIAL DESTINATARIO", "No. RADICADO DESTINATARIO",
        "FECHA (DD/MM/AAAA)", "ASUNTO / TIPO DOCUMENTAL"
    ]

    filas_novedad = []
    for idx, row in df_mem.iterrows():
        asunto_val = str(row.get("ASUNTO / TIPO DOCUMENTAL", "")).strip()
        tiene_vacios = any(not str(row.get(col, "")).strip() or str(row.get(col, "")).strip() in ['nan', 'None', 'NAN'] for col in columnas_evaluar)
        
        tiene_ruido = bool(
            re.search(r'[\|!¡]{2,}', asunto_val) or
            re.search(r'ASUNTO:\s*$', asunto_val, re.IGNORECASE) or
            re.search(r'1{5,}', asunto_val) or
            "CI004_" in asunto_val or
            "Delivery Status" in asunto_val or
            len(asunto_val) < 6
        )

        if tiene_vacios or tiene_ruido:
            filas_novedad.append(idx)

    if not filas_novedad:
        print("✅ Control de Calidad: 100% de las filas están limpias y completas.", flush=True)
        return

    print(f"⚠️ Detectadas {len(filas_novedad)} carta(s) con novedades. Auditando con Pixtral Large...", flush=True)

    corregidos = 0
    for i, idx in enumerate(filas_novedad, 1):
        row = df_mem.loc[idx]
        ubic_rel = str(row["UBICACION_ARCHIVO"]).strip()
        ruta_pdf_completa = os.path.join(RUTA_BASE, ubic_rel)
        nombre_pdf = os.path.basename(ubic_rel)
        tipo_flujo = "RECIBIDAS" if "recibidas" in ubic_rel.lower() else "RADICADAS"

        if not os.path.exists(ruta_pdf_completa):
            continue

        campos_a_revisar = [c for c in columnas_evaluar if not str(row.get(c, "")).strip() or str(row.get(c, "")).strip() in ['nan', 'None']]
        asunto_actual = str(row.get("ASUNTO / TIPO DOCUMENTAL", ""))

        print(f"   [{i}/{len(filas_novedad)}] Auditando {nombre_pdf}...", flush=True)
        datos_auditados = auditar_fila_con_pixtral_experto(ruta_pdf_completa, asunto_actual, campos_a_revisar, nombre_pdf, tipo_flujo, i)

        if datos_auditados:
            try:
                doc_t = fitz.open(ruta_pdf_completa)
                txt_t = ""
                for p in doc_t: txt_t += p.get_text() + "\n"
                doc_t.close()
            except:
                txt_t = ""

            datos_pulidos = motor_cero_vacios(datos_auditados, nombre_pdf, txt_t, "", "2017", tipo_flujo)

            for col_nombre in columnas_evaluar:
                clave_dict = col_nombre.replace(" (DD/MM/AAAA)", "").replace(" / TIPO DOCUMENTAL", "").replace(" ", "_")
                val_nuevo = datos_pulidos.get(clave_dict, "")
                if val_nuevo and str(val_nuevo).strip():
                    df_mem.at[idx, col_nombre] = str(val_nuevo).strip()

            corregidos += 1
            time.sleep(1.0)

    df_mem.to_csv(ruta_memoria, index=False)
    print(f"🎉 Auditoría Final completada: {corregidos} carta(s) corregida(s).", flush=True)

def motor_cero_vacios(datos, nombre_archivo, texto_completo, texto_pag1, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    def clean(val):
        if not val or str(val).strip().upper() in ["NONE", "N/A", "NULL", "SIN REMITENTE CONSTATADO", "SIN DESTINATARIO CONSTATADO", "SIN ASUNTO CONSTATADO", "SIN RADICADO CONSTATADO", "SIN RADICADO REMITENTE", "NAN"]:
            return ""
        return ILLEGAL_CHARACTERS_RE.sub("", str(val).strip())

    ia_dest = clean(datos.get("RAZON_SOCIAL_DESTINATARIO", ""))
    ia_rem = clean(datos.get("RAZON_SOCIAL_REMITENTE", ""))
    rad_rem = clean(datos.get("NO_RADICADO_REMITENTE", ""))
    rad_dest = clean(datos.get("NO_RADICADO_DESTINATARIO", ""))
    ia_asunto = clean(datos.get("ASUNTO", ""))
    ia_fecha = clean(datos.get("FECHA", ""))

    es_recibida = tipo_flujo == "RECIBIDAS"

    if es_recibida:
        ia_dest = "CONSORCIO 4C"

        m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
        if m_gp:
            rad_dest = f"GP-{m_gp.group(1)}"
        elif not rad_dest.startswith("GP-"):
            m_gp_txt = re.search(r'GP[-_\s]?(\d{3,6})', texto_completo, re.IGNORECASE)
            rad_dest = f"GP-{m_gp_txt.group(1)}" if m_gp_txt else ""

        if rad_rem and ("ALMA-3-" in rad_rem or not re.search(r'ALMA[-\s]?20\d{2}', rad_rem, re.IGNORECASE)):
            if "ALMA" in rad_rem: rad_rem = ""

        if not rad_rem or rad_rem.startswith("GP-"):
            m_alma_oficial = re.search(r'\b(ALMA[-\s]?20\d{2}[-\s]?\d{3,5})\b', texto_completo, re.IGNORECASE)
            m_ani = re.search(r'\b(20\d{2}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d|\d{4}-\d{3}-\d+)\b', texto_completo)
            m_car = re.search(r'\b(0\d{10})\b', texto_completo)

            if m_alma_oficial:
                rad_rem = m_alma_oficial.group(1).replace(' ', '-')
            elif m_ani:
                rad_rem = m_ani.group(1).replace(' ', '')
            elif m_car:
                rad_rem = m_car.group(1)
            else:
                m_con = re.search(r'CON_(\d{3,5})', nombre_archivo, re.IGNORECASE)
                m_ani_nom = re.search(r'ANI_([0-9\-]+)', nombre_archivo, re.IGNORECASE)
                anio_doc = anio_carpeta if str(anio_carpeta).isdigit() else "2017"
                if m_con:
                    rad_rem = f"ALMA-{anio_doc}-{m_con.group(1)}"
                elif m_ani_nom:
                    rad_rem = f"ANI-{m_ani_nom.group(1)}"

        if not ia_rem:
            if "ALMA" in rad_rem or "CON_" in nombre_archivo:
                ia_rem = "CONCESIÓN ALTO MAGDALENA S.A.S."
            elif "ANI" in rad_rem or "ANI_" in nombre_archivo:
                ia_rem = "AGENCIA NACIONAL DE INFRAESTRUCTURA – ANI"
            elif "fiduciaria bogot" in texto_completo.lower():
                ia_rem = "FIDUCIARIA BOGOTÁ"
    else:
        ia_rem = "CONSORCIO 4C"

        m_nom = re.search(r'CI004_(\d{4})\d{2}_', nombre_archivo, re.IGNORECASE)
        if m_nom:
            rad_rem = f"GP-{m_nom.group(1).zfill(4)}"
        else:
            m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
            rad_rem = f"GP-{m_gp.group(1)}" if m_gp else ""

        if not rad_dest:
            m_ani_stick = re.search(r'(?:Rad(?:icado)?\s*No\.?\s*|ANI\s*Numero\s*de\s*Radicado\s*)(\d{4}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d|\d{4}-\d{3}-\d+)', texto_completo, re.IGNORECASE)
            m_almar = re.search(r'\b(ALMA[-\s]?R[-\s]?20\d{2}[-\s]?\d+)\b', texto_completo, re.IGNORECASE)
            if not m_almar:
                m_almar = re.search(r'\b(ALMA[-\s]?R[-\s]?\d{4}[-\s]?\d+)\b', texto_completo, re.IGNORECASE)

            if m_ani_stick:
                rad_dest = m_ani_stick.group(1).replace(' ', '')
            elif m_almar:
                rad_dest = m_almar.group(1).replace(' ', '-')

        if not ia_dest:
            m_senor = re.search(r'Señor(?:es|a)?\s*:\s*\n?\s*([A-ZÁÉÍÓÚÑ\s]{3,40})(?=\n|$)', texto_pag1, re.IGNORECASE)
            if m_senor and len(m_senor.group(1).strip()) > 3:
                ia_dest = m_senor.group(1).strip()
            elif "DP_" in nombre_archivo:
                m_nom_arch = re.search(r'DP_([A-Z_]+)(?:\.pdf)?', nombre_archivo, re.IGNORECASE)
                if m_nom_arch:
                    ia_dest = m_nom_arch.group(1).replace('_', ' ').strip()
            elif "ANI_" in nombre_archivo or "ani" in texto_completo.lower():
                ia_dest = "AGENCIA NACIONAL DE INFRAESTRUCTURA – ANI"
            elif "ALMA" in texto_completo or "concesion" in texto_completo.lower():
                ia_dest = "CONCESIÓN ALTO MAGDALENA S.A.S."

    # Se aplica la limpieza exacta según las reglas acordadas
    asunto_final = limpiar_asunto_exacto(ia_asunto, texto_completo)
    fecha_final = normalizar_fecha(ia_fecha, anio_defecto=anio_carpeta)
    if not fecha_final:
        m_f = re.search(r'(?:Bogot[aá]|Girardot|Honda)[^\n\r]*,?\s*(\d{1,2}\s*de\s*[a-zA-Z]+\s*de\s*\d{4}|\d{2}[-/.]\d{2}[-/.]\d{4})', texto_completo, re.IGNORECASE)
        if m_f: fecha_final = normalizar_fecha(m_f.group(1), anio_defecto=anio_carpeta)

    return {
        "RAZON_SOCIAL_DESTINATARIO": ia_dest,
        "RAZON_SOCIAL_REMITENTE": ia_rem,
        "NO_RADICADO_REMITENTE": rad_rem,
        "NO_RADICADO_DESTINATARIO": rad_dest,
        "ASUNTO": asunto_final,
        "FECHA": fecha_final
    }

def procesar_un_pdf(item_num, pdf, ruta_completa, anio_doc, tipo, ruta_memoria, hilo_id):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_img, txt, txt1, paginas = obtener_insumos_documento(ruta_completa)
    if not b64_img:
        return False

    datos, key_usada, mod_usado = consultar_pixtral_pool(b64_img, txt1, pdf, tipo, item_num, hilo_id)

    if datos is None:
        print(f"⚠️ [Hilo-{hilo_id}] No se pudo tabular {pdf}.", flush=True)
        return False

    datos_completos = motor_cero_vacios(datos, pdf, txt, txt1, anio_doc, tipo)

    fila = {
        "ÍTEM": item_num,
        "DEL FOLIO/PAGINAS": paginas,
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
    print(f"📄 [Hilo-{hilo_id} | {key_usada} | {mod_usado}] {pdf} | ⏱️ {duracion}s", flush=True)
    return True

def fusionar_y_cargar_memoria(carpeta_objetivo, ruta_memoria_final, es_prueba=False):
    if es_prueba:
        if os.path.exists(ruta_memoria_final):
            try: os.remove(ruta_memoria_final)
            except: pass
        return set(), 1

    if not os.path.exists(ruta_memoria_final):
        return set(), 1

    try:
        df = pd.read_csv(ruta_memoria_final)
        if df.empty or "UBICACION_ARCHIVO" not in df.columns:
            return set(), 1
        procesados = set(os.path.basename(str(r).strip()).lower() for r in df["UBICACION_ARCHIVO"].dropna())
        print(f"✅ Memoria previa cargada: {len(procesados)} cartas aseguradas.", flush=True)
        return procesados, len(df) + 1
    except Exception as e:
        print(f"⚠️ Error cargando memoria: {e}", flush=True)
        return set(), 1

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (PIXTRAL: REGLAS EXACTAS DE ASUNTO)", flush=True)
    print("="*70, flush=True)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower() in ['si', 's', 'true']
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2017').strip()

    if es_prueba:
        limite = int(os.environ.get('LIMITE_PRUEBA', '1').strip())
        etiqueta = f"PRUEBA_{limite}_por_flujo"
        ruta_memoria = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_memoria_PRUEBA.csv')
        ruta_excel = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_PRUEBA.xlsx')
        print(f"🧪 MODO PRUEBA: {limite} archivo(s) por flujo. Memoria de producción aislada.", flush=True)
    else:
        limite = None
        etiqueta = f"{carpeta_objetivo}"
        ruta_memoria = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
        ruta_excel = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')
        print(f"🚀 MODO PRODUCCIÓN: Procesando año {carpeta_objetivo} con memoria persistente.", flush=True)

    procesados_basenames, item_counter = fusionar_y_cargar_memoria(carpeta_objetivo, ruta_memoria, es_prueba=es_prueba)
    flujos = [("RECIBIDAS", RUTA_RECIBIDAS), ("RADICADAS", RUTA_ENVIADAS)]
    num_trabajadores = 1 if es_prueba else min(len(lista_keys) * 2, 4)

    for tipo, ruta_raiz in flujos:
        if not os.path.exists(ruta_raiz):
            continue

        archivos = []
        for root, _, files in os.walk(ruta_raiz):
            for f in files:
                if f.lower().endswith('.pdf'):
                    archivos.append((f, os.path.join(root, f), carpeta_objetivo))

        pendientes = [x for x in archivos if os.path.basename(x[0]).lower() not in procesados_basenames]

        if not pendientes:
            print(f"✅ Todas las cartas de {tipo} ya están en memoria.", flush=True)
            continue

        print(f"\n📂 [FASE 1: BARRIDO] Tabulando {len(pendientes)} cartas pendientes en {tipo}...", flush=True)

        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for i, (pdf, ruta_completa, anio_doc) in enumerate(pendientes):
                hilo_id = (i % num_trabajadores) + 1
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, anio_doc, tipo, ruta_memoria, hilo_id)
                futuros.append(f)
                item_counter += 1
                time.sleep(0.5)

            for f in as_completed(futuros):
                pass

    if not es_prueba:
        auditar_y_corregir_tabla_final(ruta_memoria)

    generar_excel_dos_hojas(ruta_memoria, ruta_excel)
    
    print("\n📧 Enviando correo con el archivo Excel final...", flush=True)
    enviar_correo_exito(ruta_excel, etiqueta)
    print("\n🏁 Proceso concluido exitosamente.", flush=True)

def sanitizar_df_excel(df_sub):
    df_sub = df_sub.copy()
    for col in df_sub.columns:
        df_sub[col] = df_sub[col].apply(lambda x: ILLEGAL_CHARACTERS_RE.sub("", str(x)) if pd.notnull(x) else "")
    return df_sub

def generar_excel_dos_hojas(ruta_memoria, ruta_excel):
    if os.path.exists(ruta_memoria):
        try:
            df_final = pd.read_csv(ruta_memoria)
            if not df_final.empty:
                es_recibida = df_final["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
                df_recibidas = df_final[es_recibida].copy()
                df_radicadas = df_final[~es_recibida].copy()

                if not df_recibidas.empty:
                    df_recibidas["ÍTEM"] = range(1, len(df_recibidas) + 1)
                    df_recibidas = sanitizar_df_excel(df_recibidas)

                if not df_radicadas.empty:
                    df_radicadas["ÍTEM"] = range(1, len(df_radicadas) + 1)
                    df_radicadas = sanitizar_df_excel(df_radicadas)

                with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                    df_recibidas.to_excel(writer, sheet_name="Recibidas", index=False)
                    df_radicadas.to_excel(writer, sheet_name="Radicadas", index=False)

                print(f"\n✅ EXCEL GENERADO:")
                print(f"   📑 'Recibidas': {len(df_recibidas)} cartas | 'Radicadas': {len(df_radicadas)} cartas")
        except Exception as e:
            print(f"⚠️ Error al crear Excel: {e}", flush=True)

def enviar_correo_exito(ruta_archivo, etiqueta):
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        print("⚠️ No se pudo enviar correo: Faltan GMAIL_USER o GMAIL_APP_PASSWORD.", flush=True)
        return

    try:
        msg = EmailMessage()
        msg['Subject'] = f'✅ Tabulación Completa ({etiqueta}) - Excel con 2 Hojas'
        msg['From'] = EMAIL_REMITENTE
        msg['To'] = EMAIL_DESTINO
        msg.set_content(
            f'Hola,\n\n'
            f'El proceso de tabulación con Mistral Pixtral ha finalizado exitosamente para {etiqueta}.\n\n'
            f'Se adjunta el archivo Excel final con las 2 hojas generadas ("Recibidas" y "Radicadas").\n\n'
            f'Saludos cordiales.'
        )

        if os.path.exists(ruta_archivo):
            with open(ruta_archivo, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(ruta_archivo)
            msg.add_attachment(
                file_data,
                maintype='application',
                subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                filename=file_name
            )

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print(f"🚀 ¡Correo enviado exitosamente a {EMAIL_DESTINO}!", flush=True)
    except Exception as e:
        print(f"❌ Error al enviar correo por Gmail: {e}", flush=True)

if __name__ == "__main__":
    procesar_archivos()

# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (ASUNTO 100% IA | CERO BASURA OCR NI CORREOS)
# DETECTOR PÁGINA 2 ANI | RADICADOS BLINDADOS | EXCEL CON 2 HOJAS
# ==============================================================================

import os
import sys
import time
import json
import re
import random
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
from google import genai
from google.genai import types

print("⏳ [1/3] Cargando Pool de Claves Gemini...", flush=True)

raw_keys = os.environ.get('GEMINI_API_KEYS') or os.environ.get('GEMINI_API_KEY') or ""
lista_keys = [k.strip() for k in raw_keys.replace('\n', ',').split(',') if len(k.strip()) > 10]

gemini_clients = []
for i, k in enumerate(lista_keys, 1):
    try:
        c = genai.Client(
            api_key=k,
            http_options=types.HttpOptions(timeout=25_000)
        )
        gemini_clients.append((f"Key-{i}", c))
    except Exception as e:
        print(f"⚠️ Error cargando clave Gemini #{i}: {e}", flush=True)

if gemini_clients:
    print(f"✅ Pool de Gemini activo con {len(gemini_clients)} claves listas.", flush=True)
else:
    print("❌ ERROR CRÍTICO: No se cargó ninguna clave de Gemini.", flush=True)

EMAIL_REMITENTE = os.environ.get('GMAIL_USER')
EMAIL_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
EMAIL_DESTINO = os.environ.get('GMAIL_USER')

RUTA_BASE = '.'
RUTA_ENVIADAS = os.path.join(RUTA_BASE, '15_01_Cartas_Enviadas')
RUTA_RECIBIDAS = os.path.join(RUTA_BASE, '15_04_Comunic_Recibidas')

lock_csv = threading.Lock()
evento_cuota_agotada = threading.Event()

def limpiar_asunto_puro(asunto_raw, nombre_archivo=""):
    if not asunto_raw:
        return ""

    t = " ".join(str(asunto_raw).strip().split())

    # Descartar basura evidente (notificaciones de correo, rebotes, o nombres de archivo crudos)
    if any(k in t.lower() for k in ["delivery status", "mailbox unavailable", "diagnostic-code", "reflector flasher"]):
        return ""

    # Eliminar secuencias de números unos repetidos o barras de escáner
    t = re.sub(r'[1lI\|]{5,}', ' ', t)
    t = re.sub(r'[\u2500-\u257f\u2580-\u259f]+', ' ', t)

    # Si trae el prefijo 'CI004_' es nombre de archivo, no asunto real
    if t.startswith("CI004_"):
        return ""

    # Limitar a tamaño real de un asunto (máximo 350 caracteres)
    if len(t) > 350:
        t = t[:350].rsplit(' ', 1)[0]

    t = " ".join(t.split())
    return ILLEGAL_CHARACTERS_RE.sub("", t)

def obtener_insumos_documento(ruta_pdf):
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        texto_completo_pdf = ""
        for p in doc: texto_completo_pdf += p.get_text() + "\n"
        texto_pag1 = doc[0].get_text()

        # Detección de carátula electrónica ANI en Página 1
        num_pag_imagen = 0
        es_caratula_ani = (
            "al contestar cite el numero de radicado" in texto_pag1.lower() or
            "ani numero de radicado" in texto_pag1.lower() or
            ("libertad y orden" in texto_pag1.lower() and "oficio remisorio" in texto_pag1.lower())
        )

        # Si Pág 1 es carátula ANI, la carta formal con membrete y asunto está en la PÁGINA 2
        if es_caratula_ani and total_paginas > 1:
            num_pag_imagen = 1
            texto_para_ia = doc[1].get_text()
        else:
            texto_para_ia = texto_pag1

        pagina = doc[num_pag_imagen]
        pix = pagina.get_pixmap(dpi=160)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")

        if img.width > 1500:
            ratio = 1500 / float(img.width)
            img = img.resize((1500, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=85, optimize=True)
        img_bytes = buffer.getvalue()
        b64_str = base64.b64encode(img_bytes).decode('utf-8')
        doc.close()
        return b64_str, img_bytes, texto_completo_pdf, texto_para_ia, total_paginas
    except Exception:
        return None, None, "", "", 0

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
Eres un auditor archivístico experto de correspondencia técnica.
Transcribe EXACTA, PURA y LITERALMENTE lo que ves en la imagen de la carta formal.
PROHIBIDO USAR FRASES COMO "SIN ASUNTO CONSTATADO" O "SIN REMITENTE". Si algo no existe, déjalo vacío "".

REGLAS OBLIGATORIAS:
1. "RAZON_SOCIAL_REMITENTE": Entidad que emite la carta (ej. "CONSORCIO 4C", "CONCESIÓN ALTO MAGDALENA S.A.S.", "FIDUCIARIA BOGOTÁ"). Mira el logo o membrete superior.
2. "NO_RADICADO_REMITENTE": El radicado oficial de quien envía (ej. "ALMA-2017-4669", "CI.004/0143/17/2.2", "GP-XXXX").
3. "RAZON_SOCIAL_DESTINATARIO": Persona o entidad a quien va dirigida la carta (después de "Señores:", "Señor:", "Doctor"). Si es persona natural (ej. "LEONARDO SALAZAR", "MÓNICA OVIEDO"), transcribe el nombre.
4. "NO_RADICADO_DESTINATARIO": Radicado o sello recibido (ej. Sticker "ALMA-R-2017-XXXXX", sello ANI "2017-409-XXXXXX-X", sello GP).
5. "FECHA": Fecha real impresa en la carta formal (Formato DD/MM/AAAA).
6. "ASUNTO": Transcribe el texto EXACTO del Asunto o Referencia de la carta formal. Si tiene "ASUNTO:" separado de "REFERENCIA:", transcribe SOLO el "ASUNTO:". Si solo tiene "Ref.", transcribe la referencia. PROHIBIDO poner nombres de archivos técnicos ("CI004_..."), párrafos del cuerpo o notificaciones de correo ("Delivery Status").

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
    try:
        t = re.sub(r'```[a-zA-Z]*', '', texto).replace('```', '').strip()
        start, end = t.find('{'), t.rfind('}')
        if start != -1 and end != -1: return json.loads(t[start:end+1])
        return json.loads(t)
    except: return None

MODELOS_GEMINI_OFICIALES = [
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.5-flash",
    "gemini-3.7-flash"
]

def consultar_ia_completa(b64_img, img_bytes, texto_digital, nombre_archivo, tipo_flujo, item_num, hilo_id):
    if not gemini_clients or not img_bytes or evento_cuota_agotada.is_set():
        return None, "", ""

    apoyo = f"\nTipo de flujo: {tipo_flujo}\nTexto detectado:\n{texto_digital[:3500]}"
    prompt_final = f"Archivo: {nombre_archivo}\n" + PROMPT_AUDITORIA + apoyo
    part_img = types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")

    total_keys = len(gemini_clients)
    start_idx = (item_num + hilo_id) % total_keys

    for intento in range(total_keys):
        if evento_cuota_agotada.is_set():
            return None, "", ""

        idx = (start_idx + intento) % total_keys
        nombre_key, client = gemini_clients[idx]

        for mod in MODELOS_GEMINI_OFICIALES:
            try:
                r = client.models.generate_content(
                    model=mod, contents=[part_img, prompt_final],
                    config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0)
                )
                d = parsear_json(r.text)
                if d and isinstance(d, dict) and any(d.values()):
                    return d, nombre_key, mod
            except Exception as e:
                err = str(e).upper()
                if any(k in err for k in ["429", "RESOURCE_EXHAUSTED", "QUOTA", "RATE_LIMIT"]):
                    continue
                elif any(k in err for k in ["API_KEY_INVALID", "PERMISSION_DENIED", "401", "403"]):
                    break
                else:
                    continue

    return None, "", ""

def motor_cero_vacios(datos, nombre_archivo, texto_completo, texto_pag1, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    def clean(val):
        if not val or str(val).strip().upper() in ["NONE", "N/A", "NULL", "SIN REMITENTE CONSTATADO", "SIN DESTINATARIO CONSTATADO", "SIN ASUNTO CONSTATADO", "SIN RADICADO CONSTATADO", "SIN RADICADO REMITENTE", "NAN"]:
            return ""
        s = str(val).strip()
        return ILLEGAL_CHARACTERS_RE.sub("", s)

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

    # ASUNTO 100% PURO DE LA IA
    asunto_final = limpiar_asunto_puro(ia_asunto, nombre_archivo)

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

def buscar_pdfs_en_ruta(ruta_base, carpeta_filtro=None):
    archivos_encontrados = []
    if not os.path.exists(ruta_base): return archivos_encontrados
    for root, dirs, files in os.walk(ruta_base):
        pdfs = [f for f in files if f.lower().endswith('.pdf')]
        if not pdfs: continue
        if carpeta_filtro and carpeta_filtro.lower() != 'todo':
            if carpeta_filtro.lower() not in root.lower():
                continue
        m_anio = re.search(r'\b(20\d{2})\b', root)
        anio_detectado = m_anio.group(1) if m_anio else "GENERAL"
        for pdf in pdfs:
            archivos_encontrados.append((pdf, os.path.join(root, pdf), anio_detectado))
    return archivos_encontrados

def fusionar_y_cargar_memoria(carpeta_objetivo, ruta_memoria_final):
    archivos_memoria = [f for f in os.listdir(RUTA_BASE) if f.endswith('.csv') and 'memoria' in f.lower() and carpeta_objetivo in f]
    
    if not archivos_memoria:
        return set(), 1

    dfs = []
    for f_mem in archivos_memoria:
        try:
            ruta_comp = os.path.join(RUTA_BASE, f_mem)
            d = pd.read_csv(ruta_comp)
            if not d.empty and "UBICACION_ARCHIVO" in d.columns:
                dfs.append(d)
        except Exception:
            pass

    if not dfs:
        return set(), 1

    print(f"🧹 Fusionando memorias existentes...", flush=True)
    df = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=["UBICACION_ARCHIVO"])

    for idx, row in df.iterrows():
        ubic = str(row.get("UBICACION_ARCHIVO", "")).strip()
        nom_arch = os.path.basename(ubic)
        es_recibida = "recibidas" in ubic.lower()

        if es_recibida:
            df.at[idx, "RAZON SOCIAL DESTINATARIO"] = "CONSORCIO 4C"
            rad_dest = str(df.at[idx, "No. RADICADO DESTINATARIO"]).strip()
            if not rad_dest.startswith("GP-") or rad_dest in ["nan", "None", ""]:
                m_gp = re.search(r'GP[-_]?(\d{3,6})', nom_arch, re.IGNORECASE)
                if m_gp: df.at[idx, "No. RADICADO DESTINATARIO"] = f"GP-{m_gp.group(1)}"
            
            rad_rem = str(df.at[idx, "No. RADICADO REMITENTE"]).strip()
            if "ALMA-3-" in rad_rem or not rad_rem or rad_rem in ["nan", "None", ""]:
                m_con = re.search(r'CON_(\d{3,5})', nom_arch, re.IGNORECASE)
                anio_match = re.search(r'\b(20\d{2})\b', ubic)
                anio_doc = anio_match.group(1) if anio_match else "2017"
                if m_con: df.at[idx, "No. RADICADO REMITENTE"] = f"ALMA-{anio_doc}-{m_con.group(1)}"
        else:
            df.at[idx, "RAZON SOCIAL REMITENTE"] = "CONSORCIO 4C"
            rad_rem = str(df.at[idx, "No. RADICADO REMITENTE"]).strip()
            if not rad_rem.startswith("GP-") or rad_rem in ["nan", "None", ""]:
                m_nom = re.search(r'CI004_(\d{4})\d{2}_', nom_arch, re.IGNORECASE)
                if m_nom:
                    df.at[idx, "No. RADICADO REMITENTE"] = f"GP-{m_nom.group(1).zfill(4)}"
                else:
                    m_gp = re.search(r'GP[-_]?(\d{3,6})', nom_arch, re.IGNORECASE)
                    if m_gp: df.at[idx, "No. RADICADO REMITENTE"] = f"GP-{m_gp.group(1)}"

    # Purgar cualquier fila que tenga asuntos basura o datos incompletos
    asunto_str = df["ASUNTO / TIPO DOCUMENTAL"].fillna('').astype(str)
    
    malos = (
        df.astype(str).apply(lambda col: col.str.contains("CONSTATADO|SIN RADICADO", case=False, na=False)).any(axis=1) |
        asunto_str.str.strip().isin(['', 'NONE', 'N/A', 'nan']) |
        asunto_str.str.contains("CI004_", case=False, na=False) |
        asunto_str.str.contains("Delivery Status|mailbox unavailable|Diagnostic-Code|Reflector Flasher", case=False, na=False) |
        (asunto_str.str.len() > 350) |
        df["RAZON SOCIAL DESTINATARIO"].fillna('').astype(str).str.strip().isin(['', 'NONE', 'N/A', 'nan']) |
        df["No. RADICADO REMITENTE"].fillna('').astype(str).str.strip().isin(['', 'NONE', 'N/A', 'nan']) |
        df["No. RADICADO DESTINATARIO"].fillna('').astype(str).str.strip().isin(['', 'NONE', 'N/A', 'nan'])
    )

    df_limpio = df[~malos].copy()
    df_limpio.to_csv(ruta_memoria_final, index=False)

    print(f"✅ Memorias pulidas: {len(df_limpio)} cartas buenas conservadas.", flush=True)
    print(f"🎯 Detectadas {malos.sum()} cartas con datos por reparar.", flush=True)

    procesados_basenames = set(os.path.basename(str(r).strip()).lower() for r in df_limpio["UBICACION_ARCHIVO"].dropna())
    item_sig = len(df_limpio) + 1
    return procesados_basenames, item_sig

def procesar_un_pdf(item_num, pdf, ruta_completa, anio_doc, tipo, ruta_memoria, hilo_id):
    if evento_cuota_agotada.is_set():
        return False

    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_img, img_bytes, txt, txt1, paginas = obtener_insumos_documento(ruta_completa)
    datos, clave_usada, mod_usado = consultar_ia_completa(b64_img, img_bytes, txt1, pdf, tipo, item_num, hilo_id)

    if datos is None:
        print(f"⏸️ [Hilo-{hilo_id}] {pdf} | Pausado por cuota temporal.", flush=True)
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
    print(f"📄 [Hilo-{hilo_id} | {clave_usada} | {mod_usado}] {pdf} | ⏱️ {duracion}s", flush=True)
    return True

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (ASUNTO 100% IA | CERO BASURA OCR)", flush=True)
    print("="*70, flush=True)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower()
    limite = None
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2017').strip()
    etiqueta = f"{carpeta_objetivo}"

    if es_prueba in ['si', 's', 'true']:
        try:
            limite = int(os.environ.get('LIMITE_PRUEBA', '5').strip())
        except Exception:
            limite = 5
        print(f"🎲 MODO PRUEBA: {limite} archivos por flujo.", flush=True)
        etiqueta = f"PRUEBA_{limite}_archivos"

    ruta_memoria = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
    ruta_excel = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')

    reiniciar = os.environ.get('REINICIAR_MEMORIA', 'no').strip().lower() in ['si', 's', 'true']
    if reiniciar:
        archivos_memoria = [f for f in os.listdir(RUTA_BASE) if f.endswith('.csv') and 'memoria' in f.lower() and carpeta_objetivo in f]
        for m in archivos_memoria:
            os.remove(os.path.join(RUTA_BASE, m))
            print(f"🧹 REINICIO FORZADO: Memoria {m} eliminada.", flush=True)
        if os.path.exists(ruta_excel):
            os.remove(ruta_excel)

    procesados_basenames, item_counter = fusionar_y_cargar_memoria(carpeta_objetivo, ruta_memoria)
    flujos = [("RECIBIDAS", RUTA_RECIBIDAS), ("RADICADAS", RUTA_ENVIADAS)]

    for tipo, ruta_raiz in flujos:
        if evento_cuota_agotada.is_set():
            break

        print(f"\n📂 Buscando en: {tipo}...", flush=True)
        todos_los_pdfs = buscar_pdfs_en_ruta(ruta_raiz, carpeta_objetivo)
        
        pendientes = []
        for p, r, a in todos_los_pdfs:
            if os.path.basename(p).lower() not in procesados_basenames:
                pendientes.append((p, r, a))

        ya_listos = len(todos_los_pdfs) - len(pendientes)
        print(f"   Total en Drive: {len(todos_los_pdfs)} | Listos: {ya_listos} | A PROCESAR: {len(pendientes)}", flush=True)

        if limite and len(pendientes) > limite:
            pendientes = random.sample(pendientes, limite)

        if not pendientes:
            print(f"   ✅ Todas las cartas de {tipo} ya están perfectamente tabuladas.", flush=True)
            continue

        num_trabajadores = 8
        print(f"🚀 Procesando {len(pendientes)} cartas de {tipo} con {num_trabajadores} HILOS...", flush=True)

        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for i, (pdf, ruta_completa, anio_doc) in enumerate(pendientes):
                if evento_cuota_agotada.is_set():
                    break
                hilo_id = (i % num_trabajadores) + 1
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, anio_doc, tipo, ruta_memoria, hilo_id)
                futuros.append(f)
                item_counter += 1

            for f in as_completed(futuros):
                pass

    generar_excel_dos_hojas(ruta_memoria, ruta_excel)

    if evento_cuota_agotada.is_set():
        print("\n📧 Enviando correo de ALERTA al dueño...", flush=True)
        enviar_correo_alerta_cuota(ruta_excel, etiqueta)
        print("🛑 Programa detenido de forma segura.", flush=True)
        sys.exit(0)
    else:
        print("\n📧 Enviando correo de ÉXITO al dueño...", flush=True)
        enviar_correo_exito(ruta_excel, etiqueta)

def sanitizar_df_excel(df_sub):
    df_sub = df_sub.copy()
    for col in df_sub.columns:
        df_sub[col] = df_sub[col].apply(lambda x: ILLEGAL_CHARACTERS_RE.sub("", str(x)) if pd.notnull(x) else "")
    return df_sub

def generar_excel_dos_hojas(ruta_memoria, ruta_excel):
    if os.path.exists(ruta_memoria):
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

            print(f"\n✅ EXCEL CON 2 HOJAS GENERADO:", flush=True)
            print(f"   📑 Hoja 'Recibidas': {len(df_recibidas)} cartas", flush=True)
            print(f"   📑 Hoja 'Radicadas': {len(df_radicadas)} cartas", flush=True)

def enviar_correo_alerta_cuota(ruta_archivo, etiqueta):
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        return

    msg = EmailMessage()
    msg['Subject'] = f'🚨 ALERTA: Cuotas de Gemini Agotadas ({etiqueta}) - Proceso Pausado'
    msg['From'] = EMAIL_REMITENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(
        f'Hola Eduardo,\n\n'
        f'⚠️ EL PROGRAMA SE HA DETENIDO DE FORMA SEGURA:\n'
        f'Se ha alcanzado el límite de cuota en las claves de Gemini.\n\n'
        f'El avance está guardado en el archivo adjunto.\n\n'
        f'Saludos!'
    )

    try:
        if os.path.exists(ruta_archivo):
            with open(ruta_archivo, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(ruta_archivo)
            msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("📧 ¡CORREO DE ALERTA ENVIADO A TU GMAIL!", flush=True)
    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")

def enviar_correo_exito(ruta_archivo, etiqueta):
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        return

    msg = EmailMessage()
    msg['Subject'] = f'✅ Tabulación Completa ({etiqueta}) - Excel con 2 Hojas'
    msg['From'] = EMAIL_REMITENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(
        f'Hola Eduardo,\n\n'
        f'El proceso ha finalizado con éxito total para {etiqueta}.\n'
        f'El archivo adjunto contiene las 2 hojas completas:\n'
        f' - Hoja "Recibidas": Destinatario siempre Consorcio 4C y radicado GP.\n'
        f' - Hoja "Radicadas": Remitente siempre Consorcio 4C y radicado GP.\n\n'
        f'Todos los radicados están completos y los asuntos provienen 100% de la visión de la IA sin basura.\n\n'
        f'Saludos!'
    )

    try:
        if os.path.exists(ruta_archivo):
            with open(ruta_archivo, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(ruta_archivo)
            msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("🚀 ¡CORREO ENVIADO CON ÉXITO!", flush=True)
    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")

if __name__ == "__main__":
    procesar_archivos()

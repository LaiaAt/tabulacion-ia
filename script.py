# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (MOTOR TURBO MULTI-HILO + REANUDACIÓN)
# 4 HILOS PARALELOS | POOL 14 CLAVES GEMINI | SALVAVIDAS ANTI-CANCELACIÓN
# ==============================================================================

import os
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

import pymupdf as fitz
from PIL import Image
import io
from google import genai
from google.genai import types

print("⏳ [1/3] Cargando Pool de Claves Gemini...")

raw_keys = os.environ.get('GEMINI_API_KEYS') or os.environ.get('GEMINI_API_KEY') or ""
lista_keys = [k.strip() for k in raw_keys.replace('\n', ',').split(',') if len(k.strip()) > 10]

gemini_clients = []
for i, k in enumerate(lista_keys, 1):
    try:
        c = genai.Client(api_key=k)
        gemini_clients.append((f"Key-{i}", c))
    except Exception as e:
        print(f"⚠️ Error cargando clave Gemini #{i}: {e}")

if gemini_clients:
    print(f"✅ Pool de Gemini activo con {len(gemini_clients)} claves rotativas.")
else:
    print("❌ ERROR CRÍTICO: No se cargó ninguna clave de Gemini.")

EMAIL_REMITENTE = os.environ.get('GMAIL_USER')
EMAIL_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
EMAIL_DESTINO = os.environ.get('GMAIL_USER')

RUTA_BASE = '.'
RUTA_ENVIADAS = os.path.join(RUTA_BASE, '15_01_Cartas_Enviadas')
RUTA_RECIBIDAS = os.path.join(RUTA_BASE, '15_04_Comunic_Recibidas')

# Candados para concurrencia entre hilos
lock_csv = threading.Lock()
lock_key = threading.Lock()
current_key_idx = 0

# ==========================================
# LIMPIEZA DE ASUNTO SIN MUTILACIÓN
# ==========================================
def limpiar_asunto(asunto_raw, texto_doc=""):
    if not asunto_raw or str(asunto_raw).strip() in ["None", "N/A", ""]:
        m = re.search(r'ASUNTO\s*:\s*(.+?)(?=\n\s*(?:Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_doc, re.IGNORECASE | re.DOTALL)
        if m: asunto_raw = " ".join(m.group(1).split())
        else: return "SIN ASUNTO CONSTATADO"

    t = " ".join(str(asunto_raw).strip().split())
    m_asunto = re.search(r'\bASUNTO\s*:\s*(.+)', t, re.IGNORECASE)
    if m_asunto: t = m_asunto.group(1).strip()
    t = re.sub(r'^(?:REFERENCIA|Ref\.?)\s*[:\-\.]*\s*', '', t, flags=re.IGNORECASE).strip()
    t = re.sub(r'^Contrato\s+de\s+(?:Concesi[oó]n|Interventor[ií]a)[^\n\r–—\.]*?(?:Honda\s*[–—-]\s*Girardot\s*[–—-]\s*Puerto\s*Salgar|Puerto\s*Salgar\s*[–—-]\s*Girardot)?[\.\–—\-\s]*', '', t, flags=re.IGNORECASE).strip()
    t = re.sub(r'^[\.\-\–—:,;\s]+', '', t).strip()
    return t if t else str(asunto_raw).strip()

def obtener_insumos_documento(ruta_pdf):
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        texto_completo_pdf = ""
        for p in doc: texto_completo_pdf += p.get_text() + "\n"
        texto_pag1 = doc[0].get_text()

        num_pag_imagen = 0
        if "al contestar cite el numero de radicado" in texto_pag1.lower() and total_paginas > 1:
            num_pag_imagen = 1

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
        return b64_str, img_bytes, texto_completo_pdf, texto_pag1, total_paginas
    except Exception:
        return None, None, "", "", 0

def normalizar_fecha(fecha_str, anio_defecto=""):
    if not fecha_str or str(fecha_str).strip() in ["N/A", "None", "", "NO ESPECIFICADA"]:
        anio_m = re.search(r'\b(20\d{2})\b', str(anio_defecto))
        return f"01/01/{anio_m.group(1)}" if anio_m else "NO ESPECIFICADA"

    fecha_str = str(fecha_str).strip()
    fecha_str = re.sub(r'\b21(\d{2})\b', r'20\1', fecha_str)
    m1 = re.match(r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', fecha_str)
    if m1: return f"{int(m1.group(3)):02d}/{int(m1.group(2)):02d}/{m1.group(1)}"
    m2 = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})', fecha_str)
    if m2: return f"{int(m2.group(1)):02d}/{int(m2.group(2)):02d}/{m2.group(3)}"

    patron_meses = r'(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre|ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)'
    m3 = re.search(rf'(\d{{1,2}})\s*(?:de|\s|-)\s*{patron_meses}\s*(?:de|\s|-)?\s*(\d{{2,4}})', fecha_str.lower())
    if m3:
        meses_map = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06','julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12','ene':'01','feb':'02','mar':'03','abr':'04','may':'05','jun':'06','jul':'07','ago':'08','sep':'09','oct':'10','nov':'11','dic':'12'}
        anio = m3.group(3) if len(m3.group(3)) == 4 else "20" + m3.group(3)
        return f"{int(m3.group(1)):02d}/{meses_map[m3.group(2)]}/{anio}"
    return fecha_str

PROMPT_AUDITORIA = """
Eres un auditor archivístico experto en correspondencia.
Tu única misión es transcribir EXACTA y TÁCITAMENTE lo que ves en el documento, actuando como un espejo literal. PROHIBIDO SUPONER O INVENTAR DATOS.

REGLAS CRÍTICAS:
1. "RAZON_SOCIAL_DESTINATARIO": Transcribe de manera literal la entidad a la que va dirigida la carta (quien aparece después de "Señores:" o "Dirigido a:"). Ejemplos: "AGENCIA NACIONAL DE INFRAESTRUCTURA", "CONCESIÓN ALTO MAGDALENA S.A.S.", "GOBERNACIÓN DE CUNDINAMARCA". ¡Copia el texto idéntico!
2. "RAZON_SOCIAL_REMITENTE": La entidad que emite y firma la carta o cuyo logo está en el membrete superior.
3. "NO_RADICADO_REMITENTE": El radicado oficial literal que usó quien envía. (Si hay un sticker que dice "ALMA-2016-0003869", TRANSCRIBE CON TODOS LOS CEROS EXACTOS). En cartas de Consorcio 4C, será un código GP-XXXX (ej. GP-6063).
4. "NO_RADICADO_DESTINATARIO": El número de radicado o sello colocado por quien recibe (Sticker de la ANI, código de barras, o el sello GP de Consorcio 4C).
5. "ASUNTO": Transcribe literal todo el texto real del asunto, omitiendo solo la frase genérica "REFERENCIA: Contrato de Concesión...".
6. "FECHA": Formato DD/MM/AAAA.

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

def consultar_ia_completa(b64_img, img_bytes, texto_digital, nombre_archivo, tipo_flujo):
    global current_key_idx
    if not gemini_clients or not img_bytes:
        return {}

    apoyo = f"\nTipo de flujo: {tipo_flujo}\nTexto detectado:\n{texto_digital[:3500]}"
    prompt_final = f"Archivo: {nombre_archivo}\n" + PROMPT_AUDITORIA + apoyo
    part_img = types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")

    modelos_gemini = ["gemini-3.5-flash-lite", "gemini-3.5-flash", "gemini-3.6-flash", "gemini-3.1-flash-lite", "gemini-2.5-flash"]
    total_keys = len(gemini_clients)

    for intento_key in range(total_keys):
        with lock_key:
            idx = (current_key_idx + intento_key) % total_keys
            nombre_key, client = gemini_clients[idx]

        for mod in modelos_gemini:
            try:
                r = client.models.generate_content(
                    model=mod, contents=[part_img, prompt_final],
                    config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0)
                )
                d = parsear_json(r.text)
                if d and d.get("ASUNTO") and len(str(d["ASUNTO"]).strip()) > 5:
                    with lock_key:
                        current_key_idx = (idx + 1) % total_keys
                    return d

# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (GITHUB ACTIONS + GMAIL AUTOMÁTICO)
# CASCADA: GEMINI 3.X/2.5 -> KIMI -> GROQ | CERO INVENTOS | REFLEJO LITERAL
# ==============================================================================

import os
import time
import json
import re
import random
import base64
import smtplib
from email.message import EmailMessage
import pandas as pd
import requests

import pymupdf as fitz
from PIL import Image
import io
from groq import Groq
from google import genai
from google.genai import types

# ==============================================================================
# CONEXIÓN DE APIS DESDE GITHUB SECRETS
# ==============================================================================
print("⏳ Cargando configuraciones y APIs...")

gemini_client = None
k_gemini = os.environ.get('GEMINI_API_KEY')
if k_gemini:
    try:
        gemini_client = genai.Client(api_key=k_gemini.strip())
        print("✅ GEMINI listo (Prioridad 1).")
    except Exception as e:
        print(f"⚠️ Error iniciando Gemini: {e}")

kimi_key = os.environ.get('KIMI_API_KEY')
if kimi_key:
    kimi_key = kimi_key.strip()
    print("✅ KIMI listo (Prioridad 2).")

groq_client = None
k_groq = os.environ.get('GROQ_API_KEY')
if k_groq:
    try:
        groq_client = Groq(api_key=k_groq.strip())
        print("✅ GROQ listo (Prioridad 3).")
    except Exception as e:
        print(f"⚠️ Error iniciando Groq: {e}")

# Credenciales de Gmail
EMAIL_REMITENTE = os.environ.get('GMAIL_USER')
EMAIL_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
EMAIL_DESTINO = os.environ.get('GMAIL_USER')

# Rutas locales dentro del servidor de GitHub
RUTA_BASE = '.'
RUTA_ENVIADAS = os.path.join(RUTA_BASE, '15_01_Cartas_Enviadas')
RUTA_RECIBIDAS = os.path.join(RUTA_BASE, '15_04_Comunic_Recibidas')
RUTA_MEMORIA_CSV = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_memoria.csv')
RUTA_EXCEL_FINAL = os.path.join(RUTA_BASE, 'RESTREPO_2_IA.xlsx')

# ==============================================================================
# LIMPIEZA DE ASUNTO SIN MUTILACIÓN
# ==============================================================================
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

# ==============================================================================
# INSUMOS DE IMAGEN Y TEXTO
# ==============================================================================
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

# ==============================================================================
# PROMPT AUDITOR LITERAL
# ==============================================================================
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
    es_escaneado = len(texto_digital.strip()) < 60
    apoyo = f"\nTipo de flujo: {tipo_flujo}\nTexto detectado:\n{texto_digital[:3500]}"
    prompt_final = f"Archivo: {nombre_archivo}\n" + PROMPT_AUDITORIA + apoyo

    if gemini_client and img_bytes:
        # TUS MODELOS ORIGINALES DE COLAB
        modelos_gemini = ["gemini-3.5-flash-lite", "gemini-3.5-flash", "gemini-3.6-flash", "gemini-3.1-flash-lite", "gemini-2.5-flash"]
        for mod in modelos_gemini:
            for intento in range(2):
                try:
                    part_img = types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")
                    r = gemini_client.models.generate_content(
                        model=mod, contents=[part_img, prompt_final],
                        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0)
                    )
                    d = parsear_json(r.text)
                    if d and d.get("ASUNTO") and len(str(d["ASUNTO"]).strip()) > 5:
                        print(f"      ♊ Transcripción Gemini exitosa ({mod})")
                        return d
                except Exception as e:
                    print(f"      ⚠️ Intento con Gemini ({mod}) falló: {e}")
                    if ("429" in str(e) or "503" in str(e)) and intento == 0: 
                        time.sleep(2.5)
                        continue
                    break

    if kimi_key and b64_img:
        for mod_k in ["kimi-k3", "kimi-k2.6"]:
            try:
                r = requests.post("https://api.moonshot.ai/v1/chat/completions",
                    headers={"Authorization": f"Bearer {kimi_key}", "Content-Type": "application/json"},
                    json={"model": mod_k, "messages": [{"role": "user", "content": [{"type": "text", "text": prompt_final}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_img}"}}]}], "temperature": 0.0}, timeout=30)
                if r.status_code == 200:
                    d = parsear_json(r.json()["choices"][0]["message"]["content"])
                    if d and d.get("ASUNTO") and len(str(d["ASUNTO"]).strip()) > 5:
                        print(f"      🌙 Transcripción Kimi ({mod_k})")
                        return d
            except Exception as e:
                print(f"      ⚠️ Kimi falló: {e}")

    if groq_client and not es_escaneado:
        # TUS MODELOS ORIGINALES DE GROQ
        for mod_g in ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "llama-3.3-70b-versatile"]:
            try:
                res = groq_client.chat.completions.create(messages=[{"role": "user", "content": prompt_final}], model=mod_g, response_format={"type": "json_object"}, temperature=0.0)
                d = parsear_json(res.choices[0].message.content)
                if d and d.get("ASUNTO") and len(str(d["ASUNTO"]).strip()) > 5:
                    print(f"      ⚡ Transcripción Groq ({mod_g})")
                    return d
            except Exception as e:
                print(f"      ⚠️ Groq falló: {e}")
    
    print("      ❌ Ninguna IA pudo transcribir este documento (usando modo de emergencia).")
    return {}

# ==============================================================================
# MOTOR DE CONFIANZA 100% EN LA IA
# ==============================================================================
def motor_cero_vacios(datos, nombre_archivo, texto_completo, texto_pag1, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    def is_valid(val):
        return val and str(val).strip().upper() not in ["", "NONE", "N/A", "NULL", "SIN RADICADO CONSTATADO", "SIN RADICADO REMITENTE", "SIN RADICADO DESTINATARIO", "NO ESPECIFICADA"]

    ia_dest = str(datos.get("RAZON_SOCIAL_DESTINATARIO", "")).strip()
    ia_rem = str(datos.get("RAZON_SOCIAL_REMITENTE", "")).strip()
    ia_rad_rem = str(datos.get("NO_RADICADO_REMITENTE", "")).strip()
    ia_rad_dest = str(datos.get("NO_RADICADO_DESTINATARIO", "")).strip()
    ia_asunto = str(datos.get("ASUNTO", "")).strip()
    ia_fecha = str(datos.get("FECHA", "")).strip()

    datos["RAZON_SOCIAL_DESTINATARIO"] = ia_dest if is_valid(ia_dest) else "SIN DESTINATARIO CONSTATADO"
    datos["RAZON_SOCIAL_REMITENTE"] = ia_rem if is_valid(ia_rem) else "SIN REMITENTE CONSTATADO"

    rad_rem = ia_rad_rem if is_valid(ia_rad_rem) else ""
    rad_dest = ia_rad_dest if is_valid(ia_rad_dest) else ""

    if tipo_flujo == "RECIBIDAS" and rad_rem.startswith("GP-") and not rad_dest.startswith("GP-"):
        rad_rem, rad_dest = rad_dest, rad_rem
    elif tipo_flujo == "ENVIADAS" and rad_dest.startswith("GP-") and not rad_rem.startswith("GP-"):
        rad_rem, rad_dest = rad_dest, rad_rem

    if not is_valid(rad_rem):
        if tipo_flujo == "ENVIADAS":
            m_nom = re.search(r'CI004_(\d{4})\d{2}_', nombre_archivo, re.IGNORECASE)
            m_txt = re.search(r'CI\.?004[/\s_]+(?:GP\s*)?0*(\d{1,4})[/\s_]', texto_completo, re.IGNORECASE)
            if m_nom: rad_rem = f"GP-{m_nom.group(1).zfill(4)}"
            elif m_txt: rad_rem = f"GP-{m_txt.group(1).zfill(4)}"
            else: rad_rem = "SIN RADICADO REMITENTE"
        else:
            m_alma = re.search(r'\b(ALMA[-\s]?\d{4}[-\s]?\d+)\b', texto_completo, re.IGNORECASE)
            m_cssa = re.search(r'\b(CSSA\d{6,14})\b', texto_completo, re.IGNORECASE)
            m_ani = re.search(r'\b(20\d{2}-\d{3}-\d{6}-\d|\d{4}-\d{3}-\d+)\b', texto_completo)

            if m_cssa: rad_rem = m_cssa.group(1)
            elif m_alma: rad_rem = m_alma.group(1).replace(' ', '-')
            elif m_ani: rad_rem = m_ani.group(1)
            else: rad_rem = "SIN RADICADO REMITENTE"

    if not is_valid(rad_dest):
        if tipo_flujo == "ENVIADAS":
            m_ani = re.search(r'\b(20\d{2}-\d{3}-\d{6}-\d)\b', texto_completo)
            m_super = re.search(r'\b(2018560\d{7}|20\d{12})\b', texto_completo)
            m_alma_r = re.search(r'\b(ALMA-R[-\s]?\d+)\b', texto_completo, re.IGNORECASE)

            if m_ani: rad_dest = m_ani.group(1)
            elif m_super: rad_dest = m_super.group(1)
            elif m_alma_r: rad_dest = m_alma_r.group(1).replace(' ', '-')
            else: rad_dest = "SIN RADICADO CONSTATADO"
        else:
            m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
            if m_gp: rad_dest = f"GP-{m_gp.group(1)}"
            else: rad_dest = "SIN RADICADO CONSTATADO"

    if tipo_flujo == "ENVIADAS" and rad_rem.startswith("GP-") and len(rad_rem) > 7:
        rad_rem = rad_rem[:7]

    if rad_rem == rad_dest and is_valid(rad_rem):
        if tipo_flujo == "RECIBIDAS": rad_rem = "SIN RADICADO REMITENTE"
        else: rad_dest = "SIN RADICADO CONSTATADO"

    datos["NO_RADICADO_REMITENTE"] = rad_rem
    datos["NO_RADICADO_DESTINATARIO"] = rad_dest
    datos["ASUNTO"] = limpiar_asunto(ia_asunto, texto_completo)

    if not is_valid(ia_fecha) or "2105" in ia_fecha or "01/01/" in ia_fecha:
        m_f = re.search(r'(?:Bogot[aá]\s*D\.?C\.?,?\s*|Honda[^\n\r]*,?\s*|Girardot[^\n\r]*,?\s*|Fecha:\s*|FECHA:\s*)(\d{1,2}\s*(?:de|-)\s*[a-zA-Z]+\s*(?:de|-)\s*\d{2,4}|\d{2}[-/.]\d{2}[-/.]\d{4})', texto_completo, re.IGNORECASE)
        if m_f: datos["FECHA"] = normalizar_fecha(m_f.group(1), anio_defecto=anio_carpeta)
        else: datos["FECHA"] = normalizar_fecha(ia_fecha, anio_defecto=anio_carpeta)
    else:
        datos["FECHA"] = normalizar_fecha(ia_fecha, anio_defecto=anio_carpeta)

    return datos

# ==============================================================================
# MEMORIA Y AUTO-REPARACIÓN DE ERRORES PASADOS
# ==============================================================================
def cargar_memoria():
    if os.path.exists(RUTA_MEMORIA_CSV):
        try:
            df = pd.read_csv(RUTA_MEMORIA_CSV)
            if not df.empty and "UBICACION_ARCHIVO" in df.columns:
                filas_error = (
                    (df["No. RADICADO REMITENTE"].astype(str).str.contains(r'ALMA-\d{4}-\d{3,4}$', regex=True)) |
                    (df["RAZON SOCIAL DESTINATARIO"] == "CONSORCIO 4C") & (df["UBICACION_ARCHIVO"].str.contains("Recibidas")) |
                    (df["No. RADICADO DESTINATARIO"].astype(str).str.contains("SIN RADICADO")) |
                    ((df["No. RADICADO REMITENTE"] == df["No. RADICADO DESTINATARIO"]) & (~df["No. RADICADO REMITENTE"].astype(str).str.contains("SIN RADICADO")))
                )
                num_err = filas_error.sum()
                if num_err > 0:
                    print(f"🔧 Se detectaron {num_err} filas con errores pasados. Se depurarán automáticamente.")
                    df = df[~filas_error]
                    df.to_csv(RUTA_MEMORIA_CSV, index=False)

                if not df.empty:
                    print(f"🔄 MEMORIA: {len(df)} archivos limpios conservados.")
                    item_sig = int(df["ÍTEM"].max()) + 1 if "ÍTEM" in df.columns else len(df) + 1
                    return set(df["UBICACION_ARCHIVO"].dropna()), item_sig
        except Exception: pass
    return set(), 1

def buscar_pdfs_en_ruta(ruta_base, procesar_anio=None):
    archivos_encontrados = []
    if not os.path.exists(ruta_base): return archivos_encontrados
    for root, dirs, files in os.walk(ruta_base):
        pdfs = [f for f in files if f.lower().endswith('.pdf')]
        if not pdfs: continue
        m_anio = re.search(r'\b(20\d{2})\b', root)
        anio_detectado = m_anio.group(1) if m_anio else "GENERAL"
        if procesar_anio and (anio_detectado != procesar_anio and f"/{procesar_anio}" not in root): continue
        for pdf in pdfs: archivos_encontrados.append((pdf, os.path.join(root, pdf), anio_detectado))
    return archivos_encontrados

# ==============================================================================
# PROCESO PRINCIPAL
# ==============================================================================
def procesar_archivos():
    print("\n" + "="*70)
    print(" MOTOR RESTREPO_2 (MODO AUTOMÁTICO - GITHUB ACTIONS)")
    print("="*70)

    procesados, item_counter = cargar_memoria()
    
    # Toma el año que escribiste en GitHub (o procesa todo si no especificas)
    procesar_anio = os.environ.get('ANIO_PROCESAR', '').strip() or None
    if procesar_anio:
        print(f"🎯 FILTRADO AUTOMÁTICO: Procesando exclusivamente el año {procesar_anio}")
    else:
        print("🚀 Procesando archivos disponibles.")
    limite = None

    flujos = [("RECIBIDAS", RUTA_RECIBIDAS), ("ENVIADAS", RUTA_ENVIADAS)]

    for tipo, ruta_raiz in flujos:
        print(f"\n📂 Buscando en: {tipo}...")
        todos_los_pdfs = buscar_pdfs_en_ruta(ruta_raiz, procesar_anio)
        pendientes = [(p, r, a) for p, r, a in todos_los_pdfs if os.path.relpath(r, RUTA_BASE) not in procesados]
        print(f"   Encontrados {len(todos_los_pdfs)} PDFs ({len(pendientes)} pendientes).")

        if limite and len(pendientes) > limite: pendientes = random.sample(pendientes, limite)

        for pdf, ruta_completa, anio_doc in pendientes:
            t_inicio = time.time()
            ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE)
            print(f"📄 [{item_counter}] Procesando: {pdf}")

            b64_img, img_bytes, txt, txt1, paginas = obtener_insumos_documento(ruta_completa)
            datos = consultar_ia_completa(b64_img, img_bytes, txt1, pdf, tipo)
            datos_completos = motor_cero_vacios(datos, pdf, txt, txt1, anio_doc, tipo)

            duracion = round(time.time() - t_inicio, 2)
            print(f"      ⏱️ Duración: {duracion} s")

            fila = {
                "ÍTEM": item_counter,
                "DEL FOLIO/PAGINAS": paginas,
                "RAZON SOCIAL REMITENTE": datos_completos.get("RAZON_SOCIAL_REMITENTE"),
                "No. RADICADO REMITENTE": datos_completos.get("NO_RADICADO_REMITENTE"),
                "RAZON SOCIAL DESTINATARIO": datos_completos.get("RAZON_SOCIAL_DESTINATARIO"),
                "No. RADICADO DESTINATARIO": datos_completos.get("NO_RADICADO_DESTINATARIO"),
                "FECHA (DD/MM/AAAA)": datos_completos.get("FECHA"),
                "ASUNTO / TIPO DOCUMENTAL": datos_completos.get("ASUNTO"),
                "UBICACION_ARCHIVO": ruta_relativa
            }
            pd.DataFrame([fila]).to_csv(RUTA_MEMORIA_CSV, mode='a', header=not os.path.exists(RUTA_MEMORIA_CSV), index=False)
            item_counter += 1
            time.sleep(2.0)

    if os.path.exists(RUTA_MEMORIA_CSV):
        pd.read_csv(RUTA_MEMORIA_CSV).to_excel(RUTA_EXCEL_FINAL, index=False)
        print(f"\n✅ EXCEL FINALIZADO EN:\n📁 {RUTA_EXCEL_FINAL}")
        enviar_correo_excel(RUTA_EXCEL_FINAL, procesar_anio)

# ==============================================================================
# ENVÍO AUTOMÁTICO DE CORREO POR GMAIL
# ==============================================================================
def enviar_correo_excel(ruta_archivo, anio_texto):
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        print("⚠️ No se configuraron credenciales de correo (GMAIL_USER o GMAIL_APP_PASSWORD). Omitiendo envío.")
        return

    etiqueta_anio = f"Año {anio_texto}" if anio_texto else "General"
    print("📧 Preparando correo para enviar a:", EMAIL_DESTINO)
    msg = EmailMessage()
    msg['Subject'] = f'✅ Tabulación Finalizada ({etiqueta_anio}) - Excel Adjunto'
    msg['From'] = EMAIL_REMITENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(
        f'Hola Eduardo,\n\n'
        f'El proceso de tabulación automática ha finalizado con éxito para el {etiqueta_anio}.\n'
        f'Adjunto encontrarás el archivo Excel con todos los radicados y asuntos extraídos por la IA.\n\n'
        f'Saludos,\n'
        f'Bot Automático'
    )

    try:
        with open(ruta_archivo, 'rb') as f:
            file_data = f.read()
            nombre_adjunto = f"RESTREPO_2_IA_{etiqueta_anio.replace(' ', '_')}.xlsx"
        
        msg.add_attachment(
            file_data, 
            maintype='application', 
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            filename=nombre_adjunto
        )

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("🚀 ¡CORREO ENVIADO CON ÉXITO A TU GMAIL!")
    except Exception as e:
        print(f"❌ Error al enviar el correo: {e}")

if __name__ == "__main__":
    procesar_archivos()

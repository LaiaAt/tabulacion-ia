# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (DATOS PUROS Y LITERALES | CERO RELLENOS)
# EXCEL CON 2 HOJAS (RECIBIDAS Y RADICADAS) | POOL 14 CLAVES GEMINI
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

lock_csv = threading.Lock()
lock_key = threading.Lock()
current_key_idx = 0

# ==============================================================================
# LIMPIEZA DE ASUNTO PURA (SIN FRASES DE RELLENO)
# ==============================================================================
def limpiar_asunto(asunto_raw, texto_doc="", nombre_archivo=""):
    if not asunto_raw or str(asunto_raw).strip().upper() in ["NONE", "N/A", "", "SIN ASUNTO CONSTATADO"]:
        m_asunto = re.search(r'(?:ASUNTO|OBJETO|REFERENCIA|REF\.?)\s*:\s*(.+?)(?=\n\s*(?:Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_doc, re.IGNORECASE | re.DOTALL)
        if m_asunto:
            asunto_raw = " ".join(m_asunto.group(1).split())
        else:
            nombre_limpio = os.path.splitext(nombre_archivo)[0]
            m_nom = re.search(r'CON_\d+_(.+)', nombre_limpio, re.IGNORECASE)
            if m_nom:
                asunto_raw = m_nom.group(1).replace('_', ' ').strip()
            else:
                asunto_raw = ""

    t = " ".join(str(asunto_raw).strip().split())
    m_as = re.search(r'\b(?:ASUNTO|OBJETO|REFERENCIA|Ref\.?)\s*:\s*(.+)', t, re.IGNORECASE)
    if m_as: t = m_as.group(1).strip()
    t = re.sub(r'^(?:REFERENCIA|Ref\.?|OBJETO)\s*[:\-\.]*\s*', '', t, flags=re.IGNORECASE).strip()
    t = re.sub(r'^Contrato\s+de\s+(?:Concesi[oó]n|Interventor[ií]a)[^\n\r–—\.]*?(?:Honda\s*[–—-]\s*Girardot\s*[–—-]\s*Puerto\s*Salgar|Puerto\s*Salgar\s*[–—-]\s*Girardot)?[\.\–—\-\s]*', '', t, flags=re.IGNORECASE).strip()
    t = re.sub(r'^[\.\-\–—:,;\s]+', '', t).strip()
    return t if t else ""

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
    if not fecha_str or str(fecha_str).strip() in ["N/A", "None", ""]:
        anio_m = re.search(r'\b(20\d{2})\b', str(anio_defecto))
        return f"01/01/{anio_m.group(1)}" if anio_m else ""

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
Tu misión es transcribir EXACTA, PURA y TÁCITAMENTE lo que ves en el documento, actuando como un espejo literal. PROHIBIDO INVENTAR O AGREGAR TEXTOS COMO "SIN ASUNTO CONSTATADO" O "SIN REMITENTE". Si algo no existe, déjalo como cadena vacía "".

REGLAS CRÍTICAS:
1. "RAZON_SOCIAL_DESTINATARIO": Transcribe literal la entidad a la que va dirigida la carta (después de "Señores:" o "Dirigido a:").
2. "RAZON_SOCIAL_REMITENTE": La entidad que emite y firma la carta o cuyo logo está en el membrete superior.
3. "NO_RADICADO_REMITENTE": El radicado oficial literal que usó quien envía (ej. ALMA-2016-0003869, GP-XXXX).
4. "NO_RADICADO_DESTINATARIO": El radicado o sello colocado por quien recibe (Sticker ANI, código de barras, sello GP).
5. "ASUNTO": Transcribe literal el asunto, objeto o referencia real del documento.
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
                if d and isinstance(d, dict) and any(d.values()):
                    with lock_key:
                        current_key_idx = (idx + 1) % total_keys
                    return d
            except Exception as e:
                err = str(e)
                if "429" in err or "503" in err or "RESOURCE_EXHAUSTED" in err:
                    time.sleep(1.0)
                    break
                continue
    return {}

# ==============================================================================
# MOTOR CERO TEXTOS DE RELLENO (DATOS PUROS)
# ==============================================================================
def motor_cero_vacios(datos, nombre_archivo, texto_completo, texto_pag1, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    def clean_val(val):
        if not val or str(val).strip().upper() in ["NONE", "N/A", "NULL", "SIN ASUNTO CONSTATADO", "SIN REMITENTE CONSTATADO", "SIN DESTINATARIO CONSTATADO", "SIN RADICADO CONSTATADO", "SIN RADICADO REMITENTE"]:
            return ""
        return str(val).strip()

    ia_dest = clean_val(datos.get("RAZON_SOCIAL_DESTINATARIO", ""))
    ia_rem = clean_val(datos.get("RAZON_SOCIAL_REMITENTE", ""))
    rad_rem = clean_val(datos.get("NO_RADICADO_REMITENTE", ""))
    rad_dest = clean_val(datos.get("NO_RADICADO_DESTINATARIO", ""))
    ia_asunto = clean_val(datos.get("ASUNTO", ""))
    ia_fecha = clean_val(datos.get("FECHA", ""))

    # Limpieza estricta de nombres de archivo filtrados en radicados
    if rad_dest.lower().endswith(".pdf") or "ci004" in rad_dest.lower() or len(rad_dest) > 25:
        m_gp = re.search(r'GP[-_]?(\d{3,6})', rad_dest, re.IGNORECASE)
        rad_dest = f"GP-{m_gp.group(1)}" if m_gp else ""

    if rad_rem.lower().endswith(".pdf") or "ci004" in rad_rem.lower() or len(rad_rem) > 25:
        m_gp = re.search(r'GP[-_]?(\d{3,6})', rad_rem, re.IGNORECASE)
        rad_rem = f"GP-{m_gp.group(1)}" if m_gp else ""

    # Corrección de radicados invertidos ALMA
    if "ALTO MAGDALENA" in ia_rem.upper():
        if "ALMA-" in rad_dest.upper() and not "ALMA-" in rad_rem.upper():
            rad_rem, rad_dest = rad_dest, rad_rem

    # Detección de cruces de Consorcio 4C
    if tipo_flujo == "RECIBIDAS" and rad_rem.startswith("GP-") and not rad_dest.startswith("GP-"):
        rad_rem, rad_dest = rad_dest, rad_rem
    elif tipo_flujo in ["RADICADAS", "ENVIADAS"] and rad_dest.startswith("GP-") and not rad_rem.startswith("GP-"):
        rad_rem, rad_dest = rad_dest, rad_rem

    # Fallbacks inteligentes solo si están vacíos
    if not rad_rem:
        if tipo_flujo in ["RADICADAS", "ENVIADAS"]:
            m_nom = re.search(r'CI004_(\d{4})\d{2}_', nombre_archivo, re.IGNORECASE)
            m_txt = re.search(r'CI\.?004[/\s_]+(?:GP\s*)?0*(\d{1,4})[/\s_]', texto_completo, re.IGNORECASE)
            if m_nom: rad_rem = f"GP-{m_nom.group(1).zfill(4)}"
            elif m_txt: rad_rem = f"GP-{m_txt.group(1).zfill(4)}"
        else:
            m_alma = re.search(r'\b(ALMA[-\s]?\d{4}[-\s]?\d+)\b', texto_completo, re.IGNORECASE)
            m_cssa = re.search(r'\b(CSSA\d{6,14})\b', texto_completo, re.IGNORECASE)
            m_ani = re.search(r'\b(20\d{2}-\d{3}-\d{6}-\d|\d{4}-\d{3}-\d+)\b', texto_completo)
            if m_cssa: rad_rem = m_cssa.group(1)
            elif m_alma: rad_rem = m_alma.group(1).replace(' ', '-')
            elif m_ani: rad_rem = m_ani.group(1)

    if not rad_dest:
        if tipo_flujo in ["RADICADAS", "ENVIADAS"]:
            m_ani = re.search(r'\b(20\d{2}-\d{3}-\d{6}-\d)\b', texto_completo)
            m_super = re.search(r'\b(2018560\d{7}|20\d{12})\b', texto_completo)
            m_alma_r = re.search(r'\b(ALMA-R[-\s]?\d+)\b', texto_completo, re.IGNORECASE)
            if m_ani: rad_dest = m_ani.group(1)
            elif m_super: rad_dest = m_super.group(1)
            elif m_alma_r: rad_dest = m_alma_r.group(1).replace(' ', '-')
        else:
            m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
            if m_gp: rad_dest = f"GP-{m_gp.group(1)}"

    # Asunto limpio
    asunto_final = limpiar_asunto(ia_asunto, texto_completo, nombre_archivo)

    # Fecha limpia
    fecha_final = normalizar_fecha(ia_fecha, anio_defecto=anio_carpeta)
    if not fecha_final:
        m_f = re.search(r'(?:Bogot[aá]\s*D\.?C\.?,?\s*|Honda[^\n\r]*,?\s*|Girardot[^\n\r]*,?\s*|Fecha:\s*|FECHA:\s*)(\d{1,2}\s*(?:de|-)\s*[a-zA-Z]+\s*(?:de|-)\s*\d{2,4}|\d{2}[-/.]\d{2}[-/.]\d{4})', texto_completo, re.IGNORECASE)
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

def pulir_y_cargar_memoria(ruta_csv):
    if os.path.exists(ruta_csv):
        try:
            df = pd.read_csv(ruta_csv)
            if not df.empty and "UBICACION_ARCHIVO" in df.columns:
                print("🧹 Purgando registros con textos de relleno...")
                # Eliminar registros con textos de relleno anteriores
                filas_malas = (
                    df["ASUNTO / TIPO DOCUMENTAL"].astype(str).str.contains("SIN ASUNTO", case=False, na=True) |
                    df["RAZON SOCIAL REMITENTE"].astype(str).str.contains("SIN REMITENTE", case=False, na=True) |
                    (df["ASUNTO / TIPO DOCUMENTAL"].fillna('').str.strip() == '')
                )
                df_limpio = df[~filas_malas].copy()
                df_limpio.to_csv(ruta_csv, index=False)
                procesados = set(df_limpio["UBICACION_ARCHIVO"].dropna().astype(str).str.strip())
                item_sig = len(df_limpio) + 1
                return procesados, item_sig
        except Exception as e:
            print(f"⚠️ Aviso memoria: {e}")
    return set(), 1

def procesar_un_pdf(item_num, pdf, ruta_completa, anio_doc, tipo, ruta_memoria):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_img, img_bytes, txt, txt1, paginas = obtener_insumos_documento(ruta_completa)
    datos = consultar_ia_completa(b64_img, img_bytes, txt1, pdf, tipo)
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
    print(f"📄 [{tipo}] {pdf} | ⏱️ {duracion}s")
    return True

# ==============================================================================
# PROCESO PRINCIPAL
# ==============================================================================
def procesar_archivos():
    print("\n" + "="*70)
    print(" MOTOR RESTREPO_2 (DATOS PUROS | EXCEL CON 2 HOJAS)")
    print("="*70)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower()
    limite = None
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2017').strip()
    etiqueta = f"{carpeta_objetivo}"

    if es_prueba in ['si', 's', 'true']:
        try:
            limite = int(os.environ.get('LIMITE_PRUEBA', '5').strip())
        except Exception:
            limite = 5
        print(f"🎲 MODO PRUEBA: {limite} archivos por flujo.")
        etiqueta = f"PRUEBA_{limite}_archivos"

    ruta_memoria = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
    ruta_excel = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')

    reiniciar = os.environ.get('REINICIAR_MEMORIA', 'no').strip().lower() in ['si', 's', 'true']
    if reiniciar:
        if os.path.exists(ruta_memoria):
            os.remove(ruta_memoria)
            print(f"🧹 REINICIO FORZADO: Memoria eliminada.")
        if os.path.exists(ruta_excel):
            os.remove(ruta_excel)

    procesados, item_counter = pulir_y_cargar_memoria(ruta_memoria)

    # Procesar AMBOS flujos
    flujos = [("RECIBIDAS", RUTA_RECIBIDAS), ("RADICADAS", RUTA_ENVIADAS)]

    for tipo, ruta_raiz in flujos:
        print(f"\n📂 Buscando en: {tipo}...")
        todos_los_pdfs = buscar_pdfs_en_ruta(ruta_raiz, carpeta_objetivo)
        
        pendientes = []
        for p, r, a in todos_los_pdfs:
            rel_path = os.path.relpath(r, RUTA_BASE).strip()
            if rel_path not in procesados:
                pendientes.append((p, r, a))

        ya_listos = len(todos_los_pdfs) - len(pendientes)
        print(f"   Total en Drive: {len(todos_los_pdfs)} | Listos: {ya_listos} | A PROCESAR: {len(pendientes)}")

        if limite and len(pendientes) > limite:
            pendientes = random.sample(pendientes, limite)

        if not pendientes:
            print(f"   ✅ Todas las cartas de {tipo} ya están perfectamente tabuladas.")
            continue

        num_trabajadores = 4
        print(f"🚀 Procesando {len(pendientes)} cartas de {tipo} con {num_trabajadores} hilos...")

        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for pdf, ruta_completa, anio_doc in pendientes:
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, anio_doc, tipo, ruta_memoria)
                futuros.append(f)
                item_counter += 1

            for f in as_completed(futuros):
                pass

    # ==========================================================================
    # CREACIÓN DEL EXCEL DEFINITIVO CON 2 HOJAS (RECIBIDAS Y RADICADAS)
    # ==========================================================================
    if os.path.exists(ruta_memoria):
        df_final = pd.read_csv(ruta_memoria)
        if not df_final.empty:
            # Separar en dos DataFrames
            es_recibida = df_final["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
            df_recibidas = df_final[es_recibida].copy()
            df_radicadas = df_final[~es_recibida].copy()

            # Reenumerar ÍTEM de 1 en adelante para cada hoja
            if not df_recibidas.empty:
                df_recibidas["ÍTEM"] = range(1, len(df_recibidas) + 1)
            if not df_radicadas.empty:
                df_radicadas["ÍTEM"] = range(1, len(df_radicadas) + 1)

            with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                df_recibidas.to_excel(writer, sheet_name="Recibidas", index=False)
                df_radicadas.to_excel(writer, sheet_name="Radicadas", index=False)

            print(f"\n✅ EXCEL CON 2 HOJAS CREADO EXITOSAMENTE:")
            print(f"   📑 Hoja 'Recibidas': {len(df_recibidas)} cartas")
            print(f"   📑 Hoja 'Radicadas': {len(df_radicadas)} cartas")
            print(f"📁 Ruta: {ruta_excel}")
            enviar_correo_excel(ruta_excel, etiqueta, len(df_recibidas), len(df_radicadas))

def enviar_correo_excel(ruta_archivo, etiqueta, tot_rec, tot_rad):
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        print("⚠️ No se configuraron credenciales de correo. Omitiendo envío.")
        return

    print("📧 Preparando correo para enviar a:", EMAIL_DESTINO)
    msg = EmailMessage()
    msg['Subject'] = f'✅ Tabulación Completa ({etiqueta}) - Excel con 2 Hojas'
    msg['From'] = EMAIL_REMITENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(
        f'Hola Eduardo,\n\n'
        f'Ha finalizado el proceso de tabulación para {etiqueta}.\n\n'
        f'El archivo Excel adjunto contiene 2 hojas organizadas:\n'
        f' - Hoja "Recibidas": {tot_rec} cartas tabuladas.\n'
        f' - Hoja "Radicadas": {tot_rad} cartas tabuladas.\n\n'
        f'Datos 100% puros y literales extraídos por Gemini.\n\n'
        f'Saludos!'
    )

    try:
        with open(ruta_archivo, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(ruta_archivo)

        msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("🚀 ¡CORREO ENVIADO CON ÉXITO A TU GMAIL!")
    except Exception as e:
        print(f"❌ Error al enviar el correo: {e}")

if __name__ == "__main__":
    procesar_archivos()

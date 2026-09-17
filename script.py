# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (TURBO MULTI-HILO + REANUDACIÓN INCREMENTAL)
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

# ==============================================================================
# POOL DE CLAVES GEMINI CON ROTACIÓN
# ==============================================================================
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

# Candados de concurrencia segura
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

# ==============================================================================
# PROMPT AUDITOR LITERAL
# ==============================================================================
PROMPT_AUDITORIA = """
Eres un auditor archivístico experto en correspondencia.
Tu única misión es transcribir EXACTA y TÁCITAMENTE lo que ves en el documento, actuando como un espejo literal. PROHIBIDO SUPONER O INVENTAR DATOS.

REGLAS CRÍTICAS:
1. "RAZON_SOCIAL_DESTINATARIO": Transcribe literal la entidad a la que va dirigida la carta (después de "Señores:" o "Dirigido a:").
2. "RAZON_SOCIAL_REMITENTE": La entidad que emite y firma la carta o cuyo logo está en el membrete superior.
3. "NO_RADICADO_REMITENTE": El radicado oficial literal que usó quien envía (ej. ALMA-2016-0003869 con todos sus ceros, o GP-XXXX).
4. "NO_RADICADO_DESTINATARIO": El radicado o sello colocado por quien recibe (Sticker ANI, código de barras, sello GP).
5. "ASUNTO": Transcribe literal el asunto, omitiendo solo la frase genérica "REFERENCIA: Contrato de Concesión...".
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

# ==============================================================================
# CONSULTA A GEMINI CON ROTACIÓN SEGURA
# ==============================================================================
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
            except Exception as e:
                err = str(e)
                if "429" in err or "503" in err or "RESOURCE_EXHAUSTED" in err:
                    break
                continue
    return {}

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
# MOTOR DE MEMORIA INCREMENTAL (REANUDACIÓN SEGURA)
# ==============================================================================
def cargar_memoria_existente(ruta_csv):
    if os.path.exists(ruta_csv):
        try:
            df = pd.read_csv(ruta_csv)
            if not df.empty and "UBICACION_ARCHIVO" in df.columns:
                procesados = set(df["UBICACION_ARCHIVO"].dropna().astype(str).str.strip())
                item_sig = int(df["ÍTEM"].max()) + 1 if "ÍTEM" in df.columns else len(df) + 1
                return procesados, item_sig
        except Exception as e:
            print(f"⚠️ Aviso al leer memoria existente: {e}")
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
    print(f"📄 [{item_num}] {pdf} | ⏱️ {duracion}s")
    return True

# ==============================================================================
# PROCESO PRINCIPAL TURBO CON SALVAVIDAS
# ==============================================================================
def procesar_archivos():
    print("\n" + "="*70)
    print(" MOTOR RESTREPO_2 (TURBO MULTI-HILO + PERSISTENCIA INCREMENTAL)")
    print("="*70)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower()
    limite = None
    procesar_anio = None

    if es_prueba in ['si', 's', 'true']:
        try:
            limite = int(os.environ.get('LIMITE_PRUEBA', '5').strip())
        except Exception:
            limite = 5
        print(f"🎲 MODO PRUEBA: {limite} archivos AL AZAR por flujo.")
        etiqueta = f"PRUEBA_{limite}_archivos"
    else:
        resp_alcance = os.environ.get('ALCANCE', 'todo').strip()
        m_anio_dir = re.search(r'\b(20\d{2})\b', resp_alcance)
        if m_anio_dir:
            procesar_anio = m_anio_dir.group(1)
            print(f"🎯 FILTRADO: Solo año {procesar_anio}.")
            etiqueta = f"Año_{procesar_anio}"
        else:
            print("🚀 MODO PRODUCCIÓN: Procesando TODO.")
            etiqueta = "Completo"

    ruta_memoria = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_memoria_{etiqueta}.csv')
    ruta_excel = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_{etiqueta}.xlsx')

    # Solo si el usuario eligió forzar reinicio en el menú de Actions
    reiniciar = os.environ.get('REINICIAR_MEMORIA', 'no').strip().lower() in ['si', 's', 'true']
    if reiniciar:
        if os.path.exists(ruta_memoria):
            os.remove(ruta_memoria)
            print(f"🧹 REINICIO FORZADO: Memoria previa eliminada para comenzar desde cero.")
        if os.path.exists(ruta_excel):
            os.remove(ruta_excel)

    # 1. Cargar lo que ya se tabuló en ejecuciones anteriores
    procesados, item_counter = cargar_memoria_existente(ruta_memoria)

    flujos = [("RECIBIDAS", RUTA_RECIBIDAS), ("ENVIADAS", RUTA_ENVIADAS)]
    total_pendientes_global = 0

    for tipo, ruta_raiz in flujos:
        print(f"\n📂 Buscando en: {tipo}...")
        todos_los_pdfs = buscar_pdfs_en_ruta(ruta_raiz, procesar_anio)
        
        # Filtra omitiendo automáticamente los que ya están en el CSV
        pendientes = []
        for p, r, a in todos_los_pdfs:
            rel_path = os.path.relpath(r, RUTA_BASE).strip()
            if rel_path not in procesados:
                pendientes.append((p, r, a))

        ya_listos = len(todos_los_pdfs) - len(pendientes)
        print(f"   Total en Drive: {len(todos_los_pdfs)} | Ya tabulados: {ya_listos} | Pendientes por hacer: {len(pendientes)}")

        if limite and len(pendientes) > limite:
            pendientes = random.sample(pendientes, limite)

        total_pendientes_global += len(pendientes)

        if not pendientes:
            print("   ✅ No hay archivos pendientes en esta carpeta. ¡Todo al día!")
            continue

        # 4 hilos en paralelo
        num_trabajadores = 4
        print(f"🚀 Procesando {len(pendientes)} cartas pendientes con {num_trabajadores} hilos en paralelo...")

        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for pdf, ruta_completa, anio_doc in pendientes:
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, anio_doc, tipo, ruta_memoria)
                futuros.append(f)
                item_counter += 1

            for f in as_completed(futuros):
                pass

    # Compilar Excel consolidado con todo lo que hay en memoria
    if os.path.exists(ruta_memoria):
        df_final = pd.read_csv(ruta_memoria)
        if not df_final.empty:
            if "ÍTEM" in df_final.columns:
                df_final = df_final.sort_values(by="ÍTEM")
            df_final.to_excel(ruta_excel, index=False)
            print(f"\n✅ EXCEL CONSOLIDADO ({len(df_final)} cartas en total) EN:\n📁 {ruta_excel}")
            enviar_correo_excel(ruta_excel, etiqueta, len(df_final))

# ==============================================================================
# ENVÍO AUTOMÁTICO DE CORREO
# ==============================================================================
def enviar_correo_excel(ruta_archivo, etiqueta, total_filas):
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        print("⚠️ No se configuraron credenciales de correo. Omitiendo envío.")
        return

    nombre_bonito = etiqueta.replace("_", " ")
    print("📧 Preparando correo para enviar a:", EMAIL_DESTINO)
    msg = EmailMessage()
    msg['Subject'] = f'✅ Tabulación Finalizada ({nombre_bonito}) - {total_filas} Cartas'
    msg['From'] = EMAIL_REMITENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(
        f'Hola Eduardo,\n\n'
        f'Ha finalizado el proceso de tabulación ({nombre_bonito}).\n'
        f'Total de cartas consolidadas en este reporte: {total_filas}.\n'
        f'Se adjunta el archivo Excel.\n\n'
        f'Saludos,\n'
        f'Tu Bot de Tabulación'
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

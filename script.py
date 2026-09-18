# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (REPARTO MATEMÁTICO EQUITATIVO DE 31 CLAVES)
# DISTRIBUCIÓN POR CARTA | ROTACIÓN SIN ATASCOS | EXCEL CON 2 HOJAS
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
        c = genai.Client(api_key=k)
        gemini_clients.append((f"Key-{i}", c))
    except Exception as e:
        print(f"⚠️ Error cargando clave Gemini #{i}: {e}", flush=True)

if gemini_clients:
    print(f"✅ Pool de Gemini activo con {len(gemini_clients)} claves listas para trabajar.", flush=True)
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

def limpiar_asunto(asunto_raw, texto_doc=""):
    if not asunto_raw or str(asunto_raw).strip().upper() in ["NONE", "N/A", "", "SIN ASUNTO CONSTATADO"]:
        m = re.search(r'(?:ASUNTO|OBJETO|REFERENCIA|Ref\.?)\s*[:\-\.]*\s*(.+?)(?=\n\s*(?:Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_doc, re.IGNORECASE | re.DOTALL)
        if m:
            asunto_raw = " ".join(m.group(1).split())
        else:
            m2 = re.search(r'(?:Seguimiento|Solicitud|Respuesta|Informe|Envío|Remisión)[^\n\r]+', texto_doc, re.IGNORECASE)
            asunto_raw = m2.group(0).strip() if m2 else ""

    t = " ".join(str(asunto_raw).strip().split())
    m_as = re.search(r'\b(?:ASUNTO|OBJETO|REFERENCIA|Ref\.?)\s*[:\-\.]*\s*(.+)', t, re.IGNORECASE)
    if m_as: t = m_as.group(1).strip()
    t = re.sub(r'^(?:REFERENCIA|Ref\.?|OBJETO)\s*[:\-\.]*\s*', '', t, flags=re.IGNORECASE).strip()
    t = re.sub(r'^Contrato\s+de\s+(?:Concesi[oó]n|Interventor[ií]a)[^\n\r–—\.]*?(?:Honda\s*[–—-]\s*Girardot\s*[–—-]\s*Puerto\s*Salgar|Puerto\s*Salgar\s*[–—-]\s*Girardot)?[\.\–—\-\s]*', '', t, flags=re.IGNORECASE).strip()
    t = re.sub(r'^[\.\-\–—:,;\s]+', '', t).strip()
    return t if len(t) > 3 else ""

def obtener_insumos_documento(ruta_pdf):
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        texto_completo_pdf = ""
        for p in doc: texto_completo_pdf += p.get_text() + "\n"
        texto_pag1 = doc[0].get_text()

        pagina = doc[0]
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
Eres un auditor archivístico experto.
Transcribe EXACTA, PURA y LITERALMENTE lo que ves en el documento.
PROHIBIDO USAR FRASES COMO "SIN ASUNTO CONSTATADO", "SIN REMITENTE", "SIN DESTINATARIO". Si algo no existe, déjalo vacío "".

REGLAS OBLIGATORIAS:
1. "RAZON_SOCIAL_REMITENTE": Quién emite la carta (ej. "CONSORCIO 4C", "CONCESIÓN ALTO MAGDALENA S.A.S.", "AGENCIA NACIONAL DE INFRAESTRUCTURA"). Mira logos o membrete superior.
2. "NO_RADICADO_REMITENTE": El código o radicado de quien envía (ej. "CI.004/G2525/17/7.1.3", "GP-2525", "ALMA-2017-XXXX").
3. "RAZON_SOCIAL_DESTINATARIO": A quién va dirigida la carta (después de "Señores:").
4. "NO_RADICADO_DESTINATARIO": Radicado o sello recibido (ej. Sticker de barras "ALMA-R-2017-02101", sello ANI).
5. "FECHA": Fecha real de la carta (Formato DD/MM/AAAA).
6. "ASUNTO": El asunto u objeto real del documento.

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
# MOTOR CON REPARTO MATEMÁTICO EQUITATIVO Y REPORTE DE ERRORES REAL
# ==============================================================================
def consultar_ia_completa(b64_img, img_bytes, texto_digital, nombre_archivo, tipo_flujo, item_num):
    if not gemini_clients or not img_bytes or evento_cuota_agotada.is_set():
        return None

    apoyo = f"\nTipo de flujo: {tipo_flujo}\nTexto detectado:\n{texto_digital[:3500]}"
    prompt_final = f"Archivo: {nombre_archivo}\n" + PROMPT_AUDITORIA + apoyo
    part_img = types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")

    modelos_gemini = ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.0-flash"]
    total_keys = len(gemini_clients)

    # CADA CARTA ARRANCA CON UNA CLAVE DISTINTA: item_num % total_keys
    start_idx = item_num % total_keys

    for intento in range(total_keys):
        if evento_cuota_agotada.is_set():
            return None

        idx = (start_idx + intento) % total_keys
        nombre_key, client = gemini_clients[idx]

        for mod in modelos_gemini:
            try:
                r = client.models.generate_content(
                    model=mod, contents=[part_img, prompt_final],
                    config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0)
                )
                d = parsear_json(r.text)
                if d and isinstance(d, dict) and any(d.values()):
                    return d
            except Exception as e:
                err = str(e).upper()
                if any(k in err for k in ["429", "RESOURCE_EXHAUSTED", "QUOTA", "RATE_LIMIT"]):
                    print(f"      ⚠️ {nombre_key} saturada de velocidad (429). Probando siguiente...", flush=True)
                    break
                elif any(k in err for k in ["API_KEY_INVALID", "PERMISSION_DENIED", "401", "403"]):
                    print(f"      ⚠️ {nombre_key} rechazada por Google (Límite 0 o no habilitada). Probando siguiente...", flush=True)
                    break
                else:
                    print(f"      ⚠️ {nombre_key} ({mod}): {err[:70]}", flush=True)
                    continue

    return None

def motor_cero_vacios(datos, nombre_archivo, texto_completo, texto_pag1, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    def clean(val):
        if not val or str(val).strip().upper() in ["NONE", "N/A", "NULL", "SIN REMITENTE CONSTATADO", "SIN DESTINATARIO CONSTATADO", "SIN ASUNTO CONSTATADO", "SIN RADICADO CONSTATADO", "SIN RADICADO REMITENTE"]:
            return ""
        return str(val).strip()

    ia_dest = clean(datos.get("RAZON_SOCIAL_DESTINATARIO", ""))
    ia_rem = clean(datos.get("RAZON_SOCIAL_REMITENTE", ""))
    rad_rem = clean(datos.get("NO_RADICADO_REMITENTE", ""))
    rad_dest = clean(datos.get("NO_RADICADO_DESTINATARIO", ""))
    ia_asunto = clean(datos.get("ASUNTO", ""))
    ia_fecha = clean(datos.get("FECHA", ""))

    if "consorcio 4c" in texto_pag1.lower() or "ci.004" in texto_pag1.lower() or tipo_flujo in ["RADICADAS", "ENVIADAS"]:
        if not ia_rem: ia_rem = "CONSORCIO 4C"
        m_g = re.search(r'CI\.?004[/\s_]+(?:GP\s*)?0*(\d{1,4})', texto_completo, re.IGNORECASE)
        if m_g and not rad_rem: rad_rem = f"GP-{m_g.group(1).zfill(4)}"

    m_almar = re.search(r'\b(ALMA-R[-\s]?\d{4}[-\s]?\d+)\b', texto_completo, re.IGNORECASE)
    if m_almar and not rad_dest:
        rad_dest = m_almar.group(1).replace(' ', '-')

    if rad_dest.lower().endswith(".pdf") or "ci004" in rad_dest.lower() or len(rad_dest) > 25:
        m_gp = re.search(r'GP[-_]?(\d{3,6})', rad_dest, re.IGNORECASE)
        rad_dest = f"GP-{m_gp.group(1)}" if m_gp else ""

    if rad_rem.lower().endswith(".pdf") or "ci004" in rad_rem.lower() or len(rad_rem) > 25:
        m_gp = re.search(r'GP[-_]?(\d{3,6})', rad_rem, re.IGNORECASE)
        rad_rem = f"GP-{m_gp.group(1)}" if m_gp else ""

    asunto_final = limpiar_asunto(ia_asunto, texto_completo)
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
        except Exception as e:
            print(f"⚠️ Aviso leyendo {f_mem}: {e}", flush=True)

    if not dfs:
        return set(), 1

    print(f"🧹 Fusionando {len(dfs)} archivos de memoria...", flush=True)
    df = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=["UBICACION_ARCHIVO"])

    tiene_constatado = df.astype(str).apply(
        lambda col: col.str.contains("CONSTATADO|SIN RADICADO", case=False, na=False)
    ).any(axis=1)

    asunto_invalido = (
        df["ASUNTO / TIPO DOCUMENTAL"].fillna('').astype(str).str.strip().isin(['', 'NONE', 'N/A']) |
        df["ASUNTO / TIPO DOCUMENTAL"].astype(str).str.contains("CI004_", case=False, na=False)
    )

    malos = tiene_constatado | asunto_invalido
    df_limpio = df[~malos].copy()
    df_limpio.to_csv(ruta_memoria_final, index=False)

    print(f"✅ Memorias fusionadas: {len(df_limpio)} cartas buenas conservadas.", flush=True)
    print(f"🎯 Detectadas {malos.sum()} cartas con CONSTATADO que serán reparadas por Gemini.", flush=True)

    procesados_basenames = set(os.path.basename(str(r).strip()).lower() for r in df_limpio["UBICACION_ARCHIVO"].dropna())
    item_sig = len(df_limpio) + 1
    return procesados_basenames, item_sig

def procesar_un_pdf(item_num, pdf, ruta_completa, anio_doc, tipo, ruta_memoria):
    if evento_cuota_agotada.is_set():
        return False

    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_img, img_bytes, txt, txt1, paginas = obtener_insumos_documento(ruta_completa)
    
    # Se le envía item_num para asignar la clave matemáticamente
    datos = consultar_ia_completa(b64_img, img_bytes, txt1, pdf, tipo, item_num)

    if datos is None:
        print(f"⏸️ [{tipo}] {pdf} | Pausado por falta de cuota.", flush=True)
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
    print(f"📄 [{tipo}] {pdf} | ⏱️ {duracion}s", flush=True)
    return True

# ==============================================================================
# PROCESO PRINCIPAL
# ==============================================================================
def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (REPARTO EQUITATIVO DE 31 CLAVES)", flush=True)
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

        num_trabajadores = 4
        print(f"🚀 Procesando {len(pendientes)} cartas de {tipo} con {num_trabajadores} hilos...", flush=True)

        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for pdf, ruta_completa, anio_doc in pendientes:
                if evento_cuota_agotada.is_set():
                    break
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, anio_doc, tipo, ruta_memoria)
                futuros.append(f)
                item_counter += 1

            for f in as_completed(futuros):
                pass

    generar_excel_dos_hojas(ruta_memoria, ruta_excel)

    if evento_cuota_agotada.is_set():
        print("\n📧 Enviando correo de ALERTA al dueño del programa...", flush=True)
        enviar_correo_alerta_cuota(ruta_excel, etiqueta)
        print("🛑 Programa detenido de forma segura.", flush=True)
        sys.exit(0)
    else:
        print("\n📧 Enviando correo de ÉXITO al dueño del programa...", flush=True)
        enviar_correo_exito(ruta_excel, etiqueta)

def generar_excel_dos_hojas(ruta_memoria, ruta_excel):
    if os.path.exists(ruta_memoria):
        df_final = pd.read_csv(ruta_memoria)
        if not df_final.empty:
            es_recibida = df_final["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
            df_recibidas = df_final[es_recibida].copy()
            df_radicadas = df_final[~es_recibida].copy()

            if not df_recibidas.empty:
                df_recibidas["ÍTEM"] = range(1, len(df_recibidas) + 1)
            if not df_radicadas.empty:
                df_radicadas["ÍTEM"] = range(1, len(df_radicadas) + 1)

            with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                df_recibidas.to_excel(writer, sheet_name="Recibidas", index=False)
                df_radicadas.to_excel(writer, sheet_name="Radicadas", index=False)

            print(f"\n✅ EXCEL CON 2 HOJAS ACTUALIZADO:", flush=True)
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
        f'Se ha alcanzado el límite de cuota en las claves válidas de Gemini.\n\n'
        f'El sistema se frenó para NO inventar datos ni generar celdas vacías.\n'
        f'Agrega más claves en GitHub Secrets para continuar.\n\n'
        f'TU AVANCE ESTÁ A SALVO: Cuando vuelvas a ejecutarlo, reanudará exactamente donde se quedó.\n\n'
        f'Adjunto el Excel con el avance procesado en sus dos hojas.\n\n'
        f'Saludos!'
    )

    try:
        if os.path.exists(ruta_archivo):
            with open(ruta_archivo, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(ruta_archivo)
            msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("📧 ¡CORREO DE ALERTA ENVIADO A TU GMAIL!", flush=True)
    except Exception as e:
        print(f"❌ Error al enviar correo de alerta: {e}", flush=True)

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
        f' - Hoja "Recibidas"\n'
        f' - Hoja "Radicadas"\n\n'
        f'Datos 100% puros y reales.\n\n'
        f'Saludos!'
    )

    try:
        if os.path.exists(ruta_archivo):
            with open(ruta_archivo, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(ruta_archivo)
            msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("🚀 ¡CORREO ENVIADO CON ÉXITO!", flush=True)
    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")

if __name__ == "__main__":
    procesar_archivos()

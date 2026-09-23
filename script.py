# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (PIXTRAL: EXTRACTOR LIMPIO Y DESINFECTADO)
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

MODELOS_PIXTRAL = ["pixtral-12b-2409", "pixtral-large-latest"]

EMAIL_REMITENTE = os.environ.get('GMAIL_USER')
EMAIL_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')
EMAIL_DESTINO = os.environ.get('GMAIL_USER')

RUTA_BASE = '.'
RUTA_ENVIADAS = os.path.join(RUTA_BASE, '15_01_Cartas_Enviadas')
RUTA_RECIBIDAS = os.path.join(RUTA_BASE, '15_04_Comunic_Recibidas')

lock_csv = threading.Lock()
lock_keys = threading.Lock()
cooldown_keys = {k: 0.0 for k in lista_keys}

def depurar_ruido_asunto(asunto_raw, texto_doc=""):
    """
    Filtro quirúrgico contra ruido de stickers, códigos de barras y corchetes:
    - En Recibidas: elimina corchetes [ ], puntos y comillas iniciales.
    - En Radicadas: conserva 'Ref. ' y corta cualquier código de barras o leyenda de sticker.
    """
    if not asunto_raw or str(asunto_raw).strip() in ['nan', 'None', 'NAN']:
        asunto_raw = ""
    t = str(asunto_raw).strip()

    # 1. Eliminar corchetes, comillas y puntos iniciales (ej. [Respuesta a... o . Observaciones...)
    t = re.sub(r'^[\[\(\.\,\-\_\s\"\'\\]+', '', t)
    t = re.sub(r'[\]\)\s\"\'\\]+$', '', t)

    # 2. Si hay "ASUNTO:" explícito (como en Imagen 3), priorizar el texto tras los dos puntos
    m_asunto_expl = re.search(r'\bASUNTO\s*[:\-\.]*\s*(.+?)(?=\n\s*(?:Estimados|Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_doc, re.IGNORECASE | re.DOTALL)
    if m_asunto_expl:
        t_asunto = " ".join(m_asunto_expl.group(1).split()).strip()
        t_asunto = re.sub(r'^(?:ASUNTO)\s*[:\-\.]*\s*', '', t_asunto, flags=re.IGNORECASE).strip()
        if len(t_asunto) > 3 and not t_asunto.startswith("CI004_"):
            t = t_asunto

    # 3. Eliminar basura de código de barras, stickers y membretes
    t = re.sub(r'Este recibido no impl?i?ca aceptaci[oó]n[^\.\n]*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'Esle:? recibido[^\.\n]*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'Ele recibido[^\.\n]*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'ALMA-R-2017-\d{4,6}', '', t)
    t = re.sub(r'\b(?:RADICACION|Radicaci[oó]n)\s+(?:HONDA|Honda)\b', '', t)
    t = re.sub(r'[1lI\|i\'\:\!]{4,}', ' ', t)
    t = re.sub(r'[\u2500-\u257f\u2580-\u259f]+', ' ', t)

    # 4. Quitar la palabra "ASUNTO:" si quedó al inicio
    t = re.sub(r'^ASUNTO\s*[:\-\.]*\s*', '', t, flags=re.IGNORECASE).strip()

    # 5. Si la carta original empieza con Ref. (Radicadas), asegurar que empiece con Ref.
    if not m_asunto_expl and not t.lower().startswith("ref"):
        m_ref = re.search(r'\b(Ref\.?)\s*(.+?)(?=\n\s*(?:Estimados|Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_doc, re.IGNORECASE | re.DOTALL)
        if m_ref:
            bloque_ref = " ".join(m_ref.group(0).split()).strip()
            if len(bloque_ref) > 10:
                t = bloque_ref
        elif "contrato" in t.lower() or "concesi" in t.lower():
            t = f"Ref. {t}"

    # Limpieza final de espacios y caracteres ilegales
    t = " ".join(t.split())
    t = re.sub(r'^[\[\(\.\,\-\_\s\"\'\\]+', '', t).strip()
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

        num_pag_imagen = 0
        es_caratula_ani = (
            "al contestar cite el numero de radicado" in texto_pag1.lower() or
            "ani numero de radicado" in texto_pag1.lower() or
            ("libertad y orden" in texto_pag1.lower() and "oficio remisorio" in texto_pag1.lower()) or
            "ci004_" in texto_pag1.lower()
        )

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
PROHIBIDO USAR FRASES COMO "SIN ASUNTO CONSTATADO" O "SIN REMITENTE".

REGLAS DE EXTRACCIÓN:
1. "RAZON_SOCIAL_REMITENTE": Entidad que emite la carta (ej. "CONSORCIO 4C", "CONCESIÓN ALTO MAGDALENA S.A.S.", "FIDUCIARIA BOGOTÁ").
2. "NO_RADICADO_REMITENTE": El radicado oficial de quien envía (ej. "ALMA-2017-XXXX", "CI.004/...", "GP-XXXX").
3. "RAZON_SOCIAL_DESTINATARIO": Persona o entidad a quien va dirigida la carta. Si es persona natural, su nombre completo.
4. "NO_RADICADO_DESTINATARIO": Radicado o sello recibido (ej. Sticker "ALMA-R-2017-XXXXX", sello ANI "2017-409-XXXXXX-X", sello GP).
5. "FECHA": Fecha real impresa en la carta formal (Formato DD/MM/AAAA).
6. "ASUNTO" (SIGUE ESTAS 3 REGLAS ESTRICTAS):
   - CASO 1 (Cartas con bloque de Ref.): Transcribe TODO el bloque comenzando con 'Ref. ' y uniendo todas las líneas con guiones. NUNCA transcribas párrafos del cuerpo de la carta ni códigos de barras ni leyendas de stickers.
     Ejemplo: "Ref. Contrato de Concesión 003 de 2014 - Concesión Honda – Girardot – Puerto Salgar - Seguridad Vial Pasos Zonas Escolares - Respuesta ALMA-2017-3700"
   - CASO 2 (Cartas con carátula de la ANI): Ignora la carátula con nombre técnico (CI004_...). Transcribe todo el bloque de la carta real comenzando con 'Ref. '.
   - CASO 3 (Cartas con REFERENCIA y ASUNTO separados): Ignora la REFERENCIA. Transcribe ÚNICAMENTE lo que dice después de 'ASUNTO:'. NUNCA incluyas corchetes '[', ']' ni la palabra 'ASUNTO:'.
     Ejemplo: "Entrega de un (1) expediente predial de la Unidad Funcional 3, para aprobación de Ficha Predial."

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

        for mod in MODELOS_PIXTRAL:
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
            except Exception:
                continue

    return None, "", ""

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

    # Depuración de Asunto
    asunto_final = depurar_ruido_asunto(ia_asunto, texto_completo)
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

        # Limpieza masiva de corchetes en memoria previa
        for idx, row in df.iterrows():
            asunto_original = str(row.get("ASUNTO / TIPO DOCUMENTAL", ""))
            asunto_limpio = depurar_ruido_asunto(asunto_original)
            df.at[idx, "ASUNTO / TIPO DOCUMENTAL"] = asunto_limpio

        df.to_csv(ruta_memoria_final, index=False)
        procesados = set(os.path.basename(str(r).strip()).lower() for r in df["UBICACION_ARCHIVO"].dropna())
        print(f"✅ Memoria previa desinfectada: {len(procesados)} cartas aseguradas.", flush=True)
        return procesados, len(df) + 1
    except Exception as e:
        print(f"⚠️ Error cargando memoria: {e}", flush=True)
        return set(), 1

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (PIXTRAL: EXTRACTOR LIMPIO Y DESINFECTADO)", flush=True)
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

        print(f"\n📂 Tabulando {len(pendientes)} cartas pendientes/reparadas en {tipo}...", flush=True)

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
                # Limpieza final de corchetes y espacios antes de exportar
                for idx, row in df_final.iterrows():
                    asunto_orig = str(row.get("ASUNTO / TIPO DOCUMENTAL", ""))
                    df_final.at[idx, "ASUNTO / TIPO DOCUMENTAL"] = depurar_ruido_asunto(asunto_orig)

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

                print(f"\n✅ EXCEL DESINFECTADO Y GENERADO:")
                print(f"   📑 'Recibidas': {len(df_recibidas)} cartas | 'Radicadas': {len(df_radicadas)} cartas")
        except Exception as e:
            print(f"⚠️ Error al crear Excel: {e}", flush=True)

def enviar_correo_exito(ruta_archivo, etiqueta):
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        print("⚠️ No se pudo enviar correo: Faltan GMAIL_USER o GMAIL_APP_PASSWORD.", flush=True)
        return

    try:
        msg = EmailMessage()
        msg['Subject'] = f'✅ Tabulación Completa ({etiqueta}) - Excel Desinfectado con 2 Hojas'
        msg['From'] = EMAIL_REMITENTE
        msg['To'] = EMAIL_DESTINO
        msg.set_content(
            f'Hola,\n\n'
            f'El proceso de tabulación y desinfección con Mistral Pixtral ha finalizado exitosamente para {etiqueta}.\n\n'
            f'Se adjunta el archivo Excel final completamente limpio de ruido ("Recibidas" y "Radicadas").\n\n'
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

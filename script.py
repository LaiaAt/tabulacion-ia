# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (ALTA RESOLUCIÓN + REGLAS ESTRICTAS DE PÁGINA 1)
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

MODELOS_FASE_TURBO = ["pixtral-12b-2409"]

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
    """
    Convierte las páginas a alta resolución (140 DPI) para que los stickers
    y códigos pequeños debajo de códigos de barras sean totalmente legibles.
    """
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        if total_paginas == 0:
            doc.close()
            return [], "", 0

        texto_completo = ""
        for p in doc: texto_completo += p.get_text() + "\n"

        paginas_a_procesar = set()
        if total_paginas <= 4:
            paginas_a_procesar.update(range(total_paginas))
        else:
            # Página 1 (la más importante), Página 2 y última página
            paginas_a_procesar.update([0, 1, total_paginas - 1])

        imagenes_b64 = []
        for i in sorted(list(paginas_a_procesar)):
            pagina = doc[i]
            # 140 DPI para máxima nitidez de stickers
            pix = pagina.get_pixmap(dpi=140)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            if img.width > 1250:
                ratio = 1250 / float(img.width)
                img = img.resize((1250, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=80, optimize=True)
            imagenes_b64.append(base64.b64encode(buffer.getvalue()).decode('utf-8'))

        doc.close()
        return imagenes_b64, texto_completo, total_paginas
    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}", flush=True)
        return [], "", 0

PROMPT_MAESTRO = """
ACTÚA COMO UN AUDITOR Y EXTRACTOR DOCUMENTAL DE CORRESPONDENCIA CONTRACTUAL.
Analiza visualmente las imágenes del documento y extrae la información con FIDELIDAD ABSOLUTA.

REGLAS DE ORO OBLIGATORIAS:

1. "razon_social_remitente":
   - Es la entidad emisora según el LOGO O MEMBRETE DE LA PÁGINA 1.
   - En cartas con logo de Consorcio 4C: "CONSORCIO 4C".
   - En cartas con logo de Concesión: "CONCESIÓN ALTO MAGDALENA S.A.S.".
   - En cartas con logo de ANI: "AGENCIA NACIONAL DE INFRAESTRUCTURA - ANI".
   - NUNCA pongas nombres de personas ni "Atn.".

2. "no_radicado_remitente":
   - EN CARTAS DE CONSORCIO 4C (Radicadas): Está en la PÁGINA 1, arriba a la derecha (bajo el logo). Es un código como "CI.004/GPXXXX/XX/X.X" o "CI.004/XXXX/XX/X.X". Cópialo completo.
   - EN CARTAS DE CONCESIÓN ALTO MAGDALENA (Recibidas): Búscalo en el sticker arriba a la derecha. El número impreso DEBAJO DEL CÓDIGO DE BARRAS es el radicado (Ej: "ALMA-2017-5207", "ALMA-2017-0199"). Cópialo con su prefijo ALMA-.

3. "razon_social_destinatario":
   - Es la entidad a quien va dirigida la carta en la PÁGINA 1 (tras "Señores:").
   - REGLA CRÍTICA: Extrae SOLO la entidad de la PÁGINA 1. PROHIBIDO tomar entidades que aparezcan en constancias, firmas o anexos de páginas posteriores.
   - PROHIBIDO incluir "Atn.", "Ing.", "Doctor" o nombres de personas.

4. "no_radicado_destinatario":
   - Radicado o sticker que certifica la entrega:
     * Si la carta fue enviada a CONCESIÓN ALTO MAGDALENA: Es el sticker de barras "ALMA-R-AAAA-XXXXX" impreso en la PÁGINA 1. (NO tomes radicados de la ANI si la carta fue enviada a la Concesión).
     * Si fue enviada a la ANI: Es el número de radicado de entrada ANI (ej. 2017-409-XXXXXX o 2017409...).
     * Si fue recibida por CONSORCIO 4C: Es el radicado GP (ej. "GP-6735").

5. "fecha":
   - Fecha de emisión de la carta en la PÁGINA 1 (Formato DD/MM/AAAA).

6. "asunto":
   - Copia literal del bloque de Asunto o Referencia:
     * Si dice "Ref. Contrato...", copia todo el bloque completo comenzando con "Ref. ".
     * Si dice "ASUNTO: XYZ", copia "XYZ" (sin la palabra ASUNTO:).
     * PROHIBIDO copiar nombres de archivos técnicos (ej. CI004_...).

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

def consultar_pixtral_pool(b64_imgs, tipo_flujo, item_num, hilo_id):
    if not b64_imgs:
        return None, "", ""

    apoyo = f"\nTipo de flujo en carpeta: {tipo_flujo}\n"
    prompt_final = PROMPT_MAESTRO + apoyo

    num_keys = len(lista_keys)
    start_key_idx = (item_num + hilo_id) % num_keys

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

        for mod in MODELOS_FASE_TURBO:
            try:
                payload = {
                    "model": mod,
                    "temperature": 0.0,
                    "response_format": {"type": "json_object"},
                    "messages": [{"role": "user", "content": content_array}]
                }
                resp = requests.post("https://api.mistral.ai/v1/chat/completions", headers=headers, json=payload, timeout=50)
                if resp.status_code == 200:
                    d = parsear_json(resp.json()["choices"][0]["message"]["content"])
                    if d and isinstance(d, dict):
                        return d, nombre_key, mod
                elif resp.status_code == 429:
                    with lock_keys: cooldown_keys[k_actual] = time.time() + 5
                    break
            except Exception:
                continue

    return None, "", ""

def limpiar_salida(val):
    if not val: return ""
    val = str(val).strip()
    if val.upper() in ["NONE", "NULL", "NAN", "", "NO IDENTIFICADO"]: return ""
    return ILLEGAL_CHARACTERS_RE.sub("", val)

def blindaje_logica_negocio(datos, nombre_archivo, texto_completo, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    ia_dest = limpiar_salida(datos.get("razon_social_destinatario", ""))
    ia_rem = limpiar_salida(datos.get("razon_social_remitente", ""))
    rad_rem = limpiar_salida(datos.get("no_radicado_remitente", ""))
    rad_dest = limpiar_salida(datos.get("no_radicado_destinatario", ""))
    ia_asunto = limpiar_salida(datos.get("asunto", ""))
    ia_fecha = limpiar_salida(datos.get("fecha", ""))

    # Purga estricta de nombres y Atn. en entidades
    ia_dest = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente|Representante|Dra?\.?).*', '', ia_dest).strip()
    ia_rem = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente|Representante|Dra?\.?).*', '', ia_rem).strip()

    # Si la IA copió el nombre de archivo técnico como asunto, lo anulamos para rescate
    if "CI004_" in ia_asunto or len(ia_asunto) < 5:
        ia_asunto = ""

    es_recibida = tipo_flujo == "RECIBIDAS"

    if es_recibida:
        # ==================== RECIBIDAS ====================
        ia_dest = "CONSORCIO 4C"
        if "CONSORCIO 4C" in ia_rem.upper(): ia_rem = ""

        # Detección de Remitente
        if not ia_rem:
            if "CON_" in nombre_archivo or "alto magdalena" in texto_completo.lower():
                ia_rem = "CONCESIÓN ALTO MAGDALENA S.A.S."
            elif "ANI_" in nombre_archivo or "ani" in texto_completo.lower():
                ia_rem = "AGENCIA NACIONAL DE INFRAESTRUCTURA - ANI"

        # Radicado Destinatario: Consorcio 4C (GP-XXXX) extraído del nombre de archivo
        m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
        if m_gp:
            rad_dest = f"GP-{m_gp.group(1)}"

        # Radicado Remitente de la entidad emisora
        if "CI.004" in rad_rem.upper() or "GP-" in rad_rem.upper():
            rad_rem = ""

        if not rad_rem:
            # 1. Buscar en el texto digital
            m_alma_txt = re.search(r'\b(ALMA[-\s]?20\d{2}[-\s]?\d{3,5})\b', texto_completo, re.IGNORECASE)
            m_ani_rad = re.search(r'(?:Radicado\s*ANI\s*No\.?\s*[:\-\.]*\s*|Rad\s*No\.?\s*)(\d{4}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d|\d{4}-\d{3}-\d+)', texto_completo, re.IGNORECASE)
            
            if m_alma_txt:
                rad_rem = m_alma_txt.group(1).replace(' ', '-')
            elif m_ani_rad:
                rad_rem = m_ani_rad.group(1).replace(' ', '')
            else:
                # 2. Respaldo definitivo desde el nombre de archivo CON_XXXX
                m_con_file = re.search(r'CON_(\d{3,5})', nombre_archivo, re.IGNORECASE)
                m_ani_file = re.search(r'ANI_([0-9\-]+)', nombre_archivo, re.IGNORECASE)
                anio_doc = anio_carpeta if str(anio_carpeta).isdigit() else "2017"
                if m_con_file:
                    rad_rem = f"ALMA-{anio_doc}-{m_con_file.group(1).zfill(4)}"
                elif m_ani_file:
                    rad_rem = f"ANI-{m_ani_file.group(1)}"

        # Rescate de Asunto
        if not ia_asunto:
            m_as = re.search(r'\bASUNTO\s*[:\-\.]*\s*(.+?)(?=\n\s*(?:Estimados|Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_completo, re.IGNORECASE | re.DOTALL)
            if m_as:
                ia_asunto = " ".join(m_as.group(1).split()).strip()

    else:
        # ==================== RADICADAS ====================
        ia_rem = "CONSORCIO 4C"
        if "CONSORCIO 4C" in ia_dest.upper(): ia_dest = ""

        # Corrección de Destinatario si se confundió con anexos de la ANI
        if "CON_" in nombre_archivo or "alto magdalena" in texto_completo[:1500].lower():
            if not ia_dest or "ANI" in ia_dest.upper() or "INFRAESTRUCTURA" in ia_dest.upper():
                ia_dest = "CONCESIÓN ALTO MAGDALENA S.A.S."

        # Radicado Remitente (CI.004 de Consorcio 4C)
        if not rad_rem or "CI.004" not in rad_rem.upper():
            # Regex flexible para capturar el código completo aunque tenga variaciones de espaciado
            m_ci004 = re.search(r'(CI\.?\s*004[/\-\\][A-Z0-9]+[/\-\\]\d+[/\-\\][0-9.]+)', texto_completo, re.IGNORECASE)
            if m_ci004:
                rad_rem = m_ci004.group(1).replace(' ', '').strip()

        # Radicado Destinatario (Sticker de entrega)
        if "ALTO MAGDALENA" in ia_dest.upper():
            # Si el destinatario es la Concesión, el radicado DEBE ser ALMA-R-, NUNCA de la ANI
            if not rad_dest.startswith("ALMA-R-") or "2017409" in rad_dest or "202" in rad_dest:
                m_almar = re.search(r'(ALMA-R-\d{4}-\d{4,6})', texto_completo, re.IGNORECASE)
                if m_almar:
                    rad_dest = m_almar.group(1)
        elif not rad_dest:
            m_ani = re.search(r'\b(20\d{2}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d|\d{4}-\d{3}-\d+)\b', texto_completo)
            if m_ani:
                rad_dest = m_ani.group(1).replace(' ', '')

        # Rescate de Asunto en Radicadas
        if not ia_asunto or not ia_asunto.lower().startswith("ref"):
            m_ref = re.search(r'\b(Ref\.?|REFERENCIA)\s*[:\-]*\s*(.+?)(?=\n\s*(?:Respetados|Estimados|Señores|Cordial|De conformidad|Atentamente|$))', texto_completo, re.IGNORECASE | re.DOTALL)
            if m_ref:
                ia_asunto = "Ref. " + " ".join(m_ref.group(2).split()).strip()

    # Normalización de Fecha
    if not ia_fecha:
        m_f = re.search(r'(?:Bogot[aá]|Girardot|Honda)[^\n\r]*,?\s*(\d{1,2}\s*de\s*[a-zA-Z]+\s*de\s*\d{4}|\d{2}[-/.]\d{2}[-/.]\d{4})', texto_completo, re.IGNORECASE)
        ia_fecha = m_f.group(1) if m_f else ""

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
        "RAZON_SOCIAL_DESTINATARIO": ia_dest,
        "RAZON_SOCIAL_REMITENTE": ia_rem,
        "NO_RADICADO_REMITENTE": rad_rem,
        "NO_RADICADO_DESTINATARIO": rad_dest,
        "ASUNTO": ia_asunto,
        "FECHA": ia_fecha
    }

def procesar_un_pdf(item_num, pdf, ruta_completa, tipo, ruta_memoria, hilo_id, anio_doc):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_imgs, txt_completo, paginas = obtener_insumos_documento(ruta_completa)
    
    if not b64_imgs:
        datos_completos = blindaje_logica_negocio({}, pdf, txt_completo, anio_doc, tipo)
        key_usada, mod_usado = "FALLO_DOC", "RESCATE"
    else:
        datos, key_usada, mod_usado = consultar_pixtral_pool(b64_imgs, tipo, item_num, hilo_id)
        if datos is None:
            datos_completos = blindaje_logica_negocio({}, pdf, txt_completo, anio_doc, tipo)
            key_usada, mod_usado = "FALLO_IA", "RESCATE"
        else:
            datos_completos = blindaje_logica_negocio(datos, pdf, txt_completo, anio_doc, tipo)

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
    print(f"📄 [Hilo-{hilo_id} | {mod_usado}] {pdf} | ✅ OK ({duracion}s)", flush=True)
    return True

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (ALTA RESOLUCIÓN Y CORRECCIÓN TOTAL)", flush=True)
    print("="*70, flush=True)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower() in ['si', 's', 'true']
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2017').strip()

    limite = int(os.environ.get('LIMITE_PRUEBA', '1').strip()) if es_prueba else None
    etiqueta = f"PRUEBA_{limite}" if es_prueba else f"{carpeta_objetivo}"
    ruta_memoria = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_memoria_PRUEBA.csv' if es_prueba else f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
    ruta_excel = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_PRUEBA.xlsx' if es_prueba else f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')

    if es_prueba and os.path.exists(ruta_memoria): os.remove(ruta_memoria)

    # Cargar memoria previa
    procesados_basenames = set()
    item_counter = 1
    if not es_prueba and os.path.exists(ruta_memoria):
        try:
            df_m = pd.read_csv(ruta_memoria)
            procesados_basenames = set(os.path.basename(str(r).strip()).lower() for r in df_m["UBICACION_ARCHIVO"].dropna())
            item_counter = len(df_m) + 1
        except Exception: pass

    flujos = [("RECIBIDAS", RUTA_RECIBIDAS), ("RADICADAS", RUTA_ENVIADAS)]
    num_trabajadores = 1 if es_prueba else min(len(lista_keys) * 2, 4)

    for tipo, ruta_raiz in flujos:
        if not os.path.exists(ruta_raiz): continue
        archivos = []
        for root, _, files in os.walk(ruta_raiz):
            for f in files:
                if f.lower().endswith('.pdf'):
                    archivos.append((f, os.path.join(root, f), carpeta_objetivo))

        pendientes = [x for x in archivos if os.path.basename(x[0]).lower() not in procesados_basenames]

        if not pendientes:
            continue

        print(f"\n📂 Tabulando {len(pendientes)} cartas en {tipo}...", flush=True)
        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for i, (pdf, ruta_completa, anio_doc) in enumerate(pendientes):
                hilo_id = (i % num_trabajadores) + 1
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, tipo, ruta_memoria, hilo_id, anio_doc)
                futuros.append(f)
                item_counter += 1
                time.sleep(0.5)
            for f in as_completed(futuros): pass

    # ENSAMBLAJE FINAL EXCEL
    if os.path.exists(ruta_memoria):
        df_final = pd.read_csv(ruta_memoria)
        if not df_final.empty:
            es_recibida = df_final["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
            df_rec = df_final[es_recibida].copy()
            df_rad = df_final[~es_recibida].copy()

            if not df_rec.empty:
                df_rec["ÍTEM"] = range(1, len(df_rec) + 1)
            if not df_rad.empty:
                df_rad["ÍTEM"] = range(1, len(df_rad) + 1)

            with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                df_rec.to_excel(writer, sheet_name="Recibidas", index=False)
                df_rad.to_excel(writer, sheet_name="Radicadas", index=False)

    # ENVÍO DE CORREO
    if EMAIL_REMITENTE and EMAIL_PASSWORD:
        try:
            msg = EmailMessage()
            msg['Subject'] = f'✅ Tabulación Completa ({etiqueta})'
            msg['From'] = EMAIL_REMITENTE
            msg['To'] = EMAIL_DESTINO
            msg.set_content(f'Proceso concluido exitosamente con alta resolución y coherencia de radicación para {etiqueta}.')
            if os.path.exists(ruta_excel):
                with open(ruta_excel, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=os.path.basename(ruta_excel))
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as smtp:
                smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
                smtp.send_message(msg)
            print("🚀 ¡Correo enviado exitosamente!", flush=True)
        except Exception as e: print(f"❌ Error correo: {e}")

if __name__ == "__main__":
    procesar_archivos()

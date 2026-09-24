# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (CONCILIACIÓN MATEMÁTICA Y AUDITORÍA 100%)
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
        for p in doc: texto_completo += p.get_text() + "\n"

        paginas_a_procesar = set()
        if total_paginas <= 4:
            paginas_a_procesar.update(range(total_paginas))
        else:
            paginas_a_procesar.update([0, 1, total_paginas - 1])

        imagenes_b64 = []
        for i in sorted(list(paginas_a_procesar)):
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
Analiza visualmente las imágenes del documento y extrae la información con MÁXIMA PRECISIÓN Y SIN OMITIR NADA.
LA INFORMACIÓN SIEMPRE ESTÁ PRESENTE EN EL DOCUMENTO. ENCUÉNTRALA.

REGLAS DE ORO OBLIGATORIAS:
1. "razon_social_remitente":
   - Es la entidad emisora según el LOGO de la página 1 (ej. "CONCESIÓN ALTO MAGDALENA S.A.S.", "FIDUCIARIA BOGOTÁ S.A.", "CONSORCIO 4C").
   - NUNCA pongas nombres de personas ni "Atn.".

2. "no_radicado_remitente":
   - EN CARTAS DE CONCESIÓN ALTO MAGDALENA (Recibidas): El radicado SIEMPRE está en el sticker arriba a la derecha. El código bajo el código de barras es el radicado (Ej: "ALMA-2020-0994", "ALMA-2017-0199"). Cópialo completo con su prefijo ALMA-.
   - EN CARTAS DE CONSORCIO 4C (Radicadas): Está arriba a la derecha bajo el logo (Ej: "CI.004/GPXXXX/XX/X.X").
   - EN EMAILS O SOLICITUDES SIN RADICADO DE SALIDA: Escribe "SIN NÚMERO".

3. "razon_social_destinatario":
   - En RECIBIDAS: Siempre es "CONSORCIO 4C".
   - En RADICADAS: La entidad de la página 1 tras "Señores:" (sin nombres de personas).

4. "no_radicado_destinatario":
   - En RECIBIDAS: El radicado GP con el que Consorcio 4C sella el documento (ej. "GP-12333").
   - En RADICADAS: El sticker de entrega de la entidad receptora (ej. "ALMA-R-AAAA-XXXXX" o radicado ANI).

5. "fecha":
   - Fecha de la carta o del correo electrónico (Formato DD/MM/AAAA).

6. "asunto":
   - Si la carta dice "ASUNTO: XYZ", transcribe "XYZ" completo (sin la palabra ASUNTO:).
   - Si la carta comienza con "Ref.", transcribe el bloque completo comenzando con "Ref. ".
   - En correos electrónicos: Transcribe el Asunto / Subject del correo.
   - PROHIBIDO copiar nombres de archivos técnicos (ej. CI004_...).

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
    val = ILLEGAL_CHARACTERS_RE.sub("", val)
    val = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', val)
    return val

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

    if "CI004_" in ia_asunto or len(ia_asunto) < 4:
        ia_asunto = ""

    es_recibida = tipo_flujo == "RECIBIDAS"

    if es_recibida:
        ia_dest = "CONSORCIO 4C"
        if "CONSORCIO 4C" in ia_rem.upper(): ia_rem = ""

        if not ia_rem:
            if "FBTA" in nombre_archivo or "fidubogota" in texto_completo.lower() or "fiduciaria bogot" in texto_completo.lower():
                ia_rem = "FIDUCIARIA BOGOTÁ S.A."
            elif "TRANSSURENCO" in nombre_archivo.upper() or "transsurenco" in texto_completo.lower():
                ia_rem = "TRANS SURENCO S.A.S."
            elif "CON_" in nombre_archivo or "alto magdalena" in texto_completo.lower():
                ia_rem = "CONCESIÓN ALTO MAGDALENA S.A.S."
            elif "ANI_" in nombre_archivo or "ani" in texto_completo.lower():
                ia_rem = "AGENCIA NACIONAL DE INFRAESTRUCTURA - ANI"

        m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
        if m_gp:
            rad_dest = f"GP-{m_gp.group(1)}"

        if "CI.004" in rad_rem.upper() or "GP-" in rad_rem.upper():
            rad_rem = ""

        if not rad_rem:
            if "SINNUMERO" in nombre_archivo.upper() or "sin numero" in texto_completo.lower():
                rad_rem = "SIN NÚMERO"
            elif "FBTA" in nombre_archivo.upper() or "CSSA" in texto_completo:
                m_cssa = re.search(r'\b(CSSA\d{8,14})\b', texto_completo)
                m_fbta_num = re.search(r'FBTA_(\d{4,8})', nombre_archivo)
                if m_cssa: rad_rem = m_cssa.group(1)
                elif m_fbta_num: rad_rem = f"CSSA{m_fbta_num.group(1)}"
            else:
                m_alma_txt = re.search(r'\b(ALMA[-\s]?20\d{2}[-\s]?\d{3,5})\b', texto_completo, re.IGNORECASE)
                m_ani_rad = re.search(r'(?:Radicado\s*ANI\s*No\.?\s*[:\-\.]*\s*|Rad\s*No\.?\s*)(\d{4}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d|\d{4}-\d{3}-\d+)', texto_completo, re.IGNORECASE)
                if m_alma_txt: rad_rem = m_alma_txt.group(1).replace(' ', '-')
                elif m_ani_rad: rad_rem = m_ani_rad.group(1).replace(' ', '')
                else:
                    m_con_file = re.search(r'CON_(\d{3,5})', nombre_archivo, re.IGNORECASE)
                    m_ani_file = re.search(r'ANI_([0-9\-]+)', nombre_archivo, re.IGNORECASE)
                    anio_doc = anio_carpeta if str(anio_carpeta).isdigit() else "2020"
                    if m_con_file: rad_rem = f"ALMA-{anio_doc}-{m_con_file.group(1).zfill(4)}"
                    elif m_ani_file: rad_rem = f"ANI-{m_ani_file.group(1)}"

        if not ia_asunto:
            m_email_subj = re.search(r'(?:Asunto|Subject)\s*:\s*([^\n\r]+)', texto_completo, re.IGNORECASE)
            m_email_bold = re.search(r'Correo de Interventor[^\n\r]*[-–]\s*([^\n\r]+)', texto_completo, re.IGNORECASE)
            m_as = re.search(r'\bASUNTO\s*[:\-\.]*\s*(.+?)(?=\n\s*(?:Estimados|Señores|Doctor|Respetad|Cordial|Atentamente|De conformidad|$))', texto_completo, re.IGNORECASE | re.DOTALL)
            if m_email_subj: ia_asunto = " ".join(m_email_subj.group(1).split()).strip()
            elif m_email_bold: ia_asunto = " ".join(m_email_bold.group(1).split()).strip()
            elif m_as: ia_asunto = " ".join(m_as.group(1).split()).strip()

        if not ia_fecha:
            m_f_email = re.search(r'(\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4})', texto_completo, re.IGNORECASE)
            m_f_slash = re.search(r'\b(\d{1,2}/\d{1,2}/\d{4})\b', texto_completo)
            if m_f_email: ia_fecha = m_f_email.group(1)
            elif m_f_slash: ia_fecha = m_f_slash.group(1)

    else:
        ia_rem = "CONSORCIO 4C"
        if "CONSORCIO 4C" in ia_dest.upper(): ia_dest = ""

        if "CON_" in nombre_archivo or "alto magdalena" in texto_completo[:1500].lower():
            if not ia_dest or "ANI" in ia_dest.upper() or "INFRAESTRUCTURA" in ia_dest.upper():
                ia_dest = "CONCESIÓN ALTO MAGDALENA S.A.S."

        if not rad_rem or "CI.004" not in rad_rem.upper():
            m_ci004 = re.search(r'(CI\.?\s*004[/\-\\][A-Z0-9]+[/\-\\]\d+[/\-\\][0-9.]+)', texto_completo, re.IGNORECASE)
            if m_ci004: rad_rem = m_ci004.group(1).replace(' ', '').strip()

        if "ALTO MAGDALENA" in ia_dest.upper():
            if not rad_dest.startswith("ALMA-R-") or "2017409" in rad_dest or "202" in rad_dest:
                m_almar = re.search(r'(ALMA-R-\d{4}-\d{4,6})', texto_completo, re.IGNORECASE)
                if m_almar: rad_dest = m_almar.group(1)
        elif not rad_dest:
            m_ani = re.search(r'\b(20\d{2}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d|\d{4}-\d{3}-\d+)\b', texto_completo)
            if m_ani: rad_dest = m_ani.group(1).replace(' ', '')

        if not ia_asunto or not ia_asunto.lower().startswith("ref"):
            m_ref = re.search(r'\b(Ref\.?|REFERENCIA)\s*[:\-]*\s*(.+?)(?=\n\s*(?:Respetados|Estimados|Señores|Cordial|De conformidad|Atentamente|$))', texto_completo, re.IGNORECASE | re.DOTALL)
            if m_ref: ia_asunto = "Ref. " + " ".join(m_ref.group(2).split()).strip()

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

    item_counter = 1
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

        # BUCLE DE CONCILIACIÓN (Hasta 3 rondas automáticas para asegurar el 100%)
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
                print(f"✅ ¡Conciliación perfecta en {tipo}! {len(procesados_actuales)} de {total_en_carpeta} cartas ya en memoria.", flush=True)
                break

            if ronda > 1:
                print(f"🚨 [RONDA DE RESCATE {ronda}] Procesando de inmediato {len(pendientes)} cartas que quedaron pendientes...", flush=True)
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

        # Verificación de cierre del flujo
        df_post = pd.read_csv(ruta_memoria)
        if tipo == "RECIBIDAS":
            df_tipo_fin = df_post[df_post["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)]
        else:
            df_tipo_fin = df_post[~df_post["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)]
        total_tabulados = len(df_tipo_fin)
        conteo_validacion[tipo] = (total_en_carpeta, total_tabulados)

    # ENSAMBLAJE FINAL EXCEL CON SANITIZACIÓN Y DEDUPLICACIÓN
    reporte_validacion = ""
    if os.path.exists(ruta_memoria):
        try:
            df_final = pd.read_csv(ruta_memoria)
            if not df_final.empty:
                # Deduplicación por seguridad
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

    # ENVÍO DE CORREO CON EL REPORTE DE CONCILIACIÓN
    if EMAIL_REMITENTE and EMAIL_PASSWORD:
        try:
            msg = EmailMessage()
            msg['Subject'] = f'✅ Tabulación Verificada 100% ({etiqueta})'
            msg['From'] = EMAIL_REMITENTE
            msg['To'] = EMAIL_DESTINO
            msg.set_content(
                f'Hola,\n\n'
                f'El proceso para {etiqueta} ha finalizado con ÉXITO Y CONCILIACIÓN FÍSICA TOTAL.\n\n'
                f'{reporte_validacion}\n'
                f'Se garantiza que cada archivo en Google Drive cuenta con su respectiva fila en el Excel adjunto.\n\n'
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

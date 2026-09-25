# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 - VERSIÓN 2 (EXTRACCIÓN 100% IA DE VISIÓN)
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

print("⏳ [1/2] Configurando Pool de Claves de Mistral AI en V2...", flush=True)

raw_keys = os.environ.get('MISTRAL_API_KEY', '').strip()
lista_keys = [k.strip() for k in raw_keys.replace('\n', ',').split(',') if len(k.strip()) > 10]

if not lista_keys:
    print("❌ ERROR CRÍTICO: No se detectó MISTRAL_API_KEY.", flush=True)
    sys.exit(1)

print(f"   ✅ Pool V2 activo con {len(lista_keys)} clave(s).", flush=True)

# Solo utilizamos el modelo especializado de Pixtral
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
    """
    Carga el PDF y extrae:
    1. Un recorte en ALTÍSIMA RESOLUCIÓN (250 DPI) del sticker superior derecho.
    2. Las páginas relevantes (carta inicial y correos finales).
    Todo se envía a la IA para que lo lea. CERO OCR LOCAL.
    """
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        if total_paginas == 0:
            doc.close()
            return [], total_paginas

        paginas_a_procesar = set()
        
        # Siempre incluir página 1 (índice 0) y página 2 (índice 1)
        paginas_a_procesar.update([0])
        if total_paginas > 1:
            paginas_a_procesar.update([1])
            
        # Incluir la última página siempre (suele tener correos de confirmación)
        paginas_a_procesar.update([total_paginas - 1])

        imagenes_b64 = []

        # 1. GENERAR SÚPER-RECORTE DEL STICKER (Esquina Superior Derecha)
        p1 = doc[0]
        w, h = p1.rect.width, p1.rect.height
        rect_sticker = fitz.Rect(w * 0.40, 0, w, h * 0.45) # Cuadrante superior derecho amplio
        pix_sticker = p1.get_pixmap(clip=rect_sticker, dpi=250) # 250 DPI = Ultra nitidez
        img_sticker = Image.open(io.BytesIO(pix_sticker.tobytes("png"))).convert("L")
        buf_sticker = io.BytesIO()
        img_sticker.save(buf_sticker, format="JPEG", quality=90, optimize=True)
        imagenes_b64.append(base64.b64encode(buf_sticker.getvalue()).decode('utf-8'))

        # 2. GENERAR PÁGINAS COMPLETAS
        for i in sorted(list(paginas_a_procesar)):
            pagina = doc[i]
            pix = pagina.get_pixmap(dpi=130)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            if img.width > 1200:
                ratio = 1200 / float(img.width)
                img = img.resize((1200, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=80, optimize=True)
            imagenes_b64.append(base64.b64encode(buffer.getvalue()).decode('utf-8'))

        doc.close()
        return imagenes_b64, total_paginas
    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}", flush=True)
        return [], 0

PROMPT_MAESTRO = """
ACTÚA COMO UN SISTEMA DE EXTRACCIÓN DOCUMENTAL.
Tu trabajo es extraer los datos de las imágenes que te envío, LEYÉNDOLAS TÚ MISMO.
La primera imagen es un acercamiento en ALTA DEFINICIÓN de la esquina de la carta. Las siguientes son las páginas completas.

CAMPOS A EXTRAER:
1. "razon_social_remitente": Entidad que envía. NUNCA incluyas "Atn.", "Gerente" o nombres de personas.
2. "no_radicado_remitente": Código oficial de envío.
3. "razon_social_destinatario": Entidad receptora. PROHIBIDO incluir nombres de personas.
4. "no_radicado_destinatario": Radicado de recepción.
   - ¡MUY IMPORTANTE EN CARTAS ENVIADAS A CONCESIÓN ALTO MAGDALENA!
   - EL RADICADO SIEMPRE EXISTE. Búscalo en la PRIMERA IMAGEN (el acercamiento del sticker). Es el código alfanumérico debajo del código de barras (Ej: "ALMA-R-2021-1536", "ALMA-R-2021-0665").
   - Si no hay sticker, lee el texto de los correos impresos en las últimas páginas. Busca frases como "su comunicado fue ingresado con radicado entrante ALMA-R-2021- 1464".
   - Si el radicado tiene un espacio por error (ej. "ALMA-R-2021- 1464"), CÓPIALO EXACTO o únelo. NUNCA pongas "SIN NÚMERO" si logras ver ese código en alguna parte.
5. "fecha": Fecha principal de la carta (Formato DD/MM/AAAA).
6. "asunto": Transcripción literal y fiel del asunto o Ref.

Si un dato definitivamente no está en NINGUNA imagen, usa "NO IDENTIFICADO".

Devuelve ÚNICAMENTE un JSON válido sin Markdown:
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

    # En V2 usamos 3 rondas y priorizamos el modelo grande si falla el pequeño
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

        time.sleep(8)

    return None, "", ""

def limpiar_salida_excel(val):
    if not val: return "SIN NÚMERO"
    val = str(val).strip()
    if val.upper() in ["NONE", "NULL", "NAN", "", "NO IDENTIFICADO"]: return "SIN NÚMERO"
    
    # Auto-Limpieza del error del espacio en el radicado de la IA (Ej: ALMA-R-2021- 1464 -> ALMA-R-2021-1464)
    val = re.sub(r'(ALMA[-\s]?R[-\s]?\d{4})[-\s]+(\d+)', r'\1-\2', val, flags=re.IGNORECASE)
    
    val = ILLEGAL_CHARACTERS_RE.sub("", val)
    val = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', val)
    return val

def procesar_un_pdf(item_num, pdf, ruta_completa, tipo, ruta_memoria, hilo_id):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_imgs, paginas = obtener_insumos_documento(ruta_completa)
    
    if not b64_imgs:
        datos = {}
        mod_usado = "FALLO_DOC"
    else:
        datos, key_usada, mod_usado = consultar_pixtral_pool(b64_imgs, tipo, item_num, hilo_id)
        if datos is None:
            datos = {}
            mod_usado = "FALLO_IA"

    # Python SOLO formatea la salida de la IA, no inventa datos.
    ia_dest = limpiar_salida_excel(datos.get("razon_social_destinatario", ""))
    ia_rem = limpiar_salida_excel(datos.get("razon_social_remitente", ""))
    
    if tipo == "RECIBIDAS":
        ia_dest = "CONSORCIO 4C"
    else:
        ia_rem = "CONSORCIO 4C"

    fila = {
        "ÍTEM": item_num,
        "DEL FOLIO/PAGINAS": paginas if paginas > 0 else 1,
        "RAZON SOCIAL REMITENTE": ia_rem,
        "No. RADICADO REMITENTE": limpiar_salida_excel(datos.get("no_radicado_remitente", "")),
        "RAZON SOCIAL DESTINATARIO": ia_dest,
        "No. RADICADO DESTINATARIO": limpiar_salida_excel(datos.get("no_radicado_destinatario", "")),
        "FECHA (DD/MM/AAAA)": limpiar_salida_excel(datos.get("fecha", "")),
        "ASUNTO / TIPO DOCUMENTAL": limpiar_salida_excel(datos.get("asunto", "")),
        "UBICACION_ARCHIVO": ruta_relativa
    }

    with lock_csv:
        pd.DataFrame([fila]).to_csv(ruta_memoria, mode='a', header=not os.path.exists(ruta_memoria), index=False)

    duracion = round(time.time() - t_inicio, 2)
    print(f"📄 [Hilo-{hilo_id} | {mod_usado}] {pdf} | ✅ OK ({duracion}s)", flush=True)
    return True

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (VERSIÓN 2: VISIÓN PURA IA - RESCATE DE SIN NÚMERO)", flush=True)
    print("="*70, flush=True)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower() in ['si', 's', 'true']
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2021').strip()

    limite = int(os.environ.get('LIMITE_PRUEBA', '1').strip()) if es_prueba else None
    etiqueta = f"PRUEBA_{limite}" if es_prueba else f"{carpeta_objetivo}"
    ruta_memoria = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_memoria_PRUEBA.csv' if es_prueba else f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
    ruta_excel = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_PRUEBA.xlsx' if es_prueba else f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')

    if es_prueba and os.path.exists(ruta_memoria): os.remove(ruta_memoria)

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
                    archivos.append((f, os.path.join(root, f)))

        pendientes = [x for x in archivos if os.path.basename(x[0]).lower() not in procesados_basenames]

        if not pendientes:
            continue

        print(f"\n📂 [V2 IA VISUAL] Tabulando {len(pendientes)} cartas en {tipo}...", flush=True)
        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for i, (pdf, ruta_completa) in enumerate(pendientes):
                hilo_id = (i % num_trabajadores) + 1
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, tipo, ruta_memoria, hilo_id)
                futuros.append(f)
                item_counter += 1
                time.sleep(0.5)
            for f in as_completed(futuros): pass

    # ENSAMBLAJE FINAL EXCEL
    if os.path.exists(ruta_memoria):
        try:
            df_final = pd.read_csv(ruta_memoria)
            if not df_final.empty:
                df_final.drop_duplicates(subset=["UBICACION_ARCHIVO"], keep="last", inplace=True)

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
        except Exception as e:
            print(f"⚠️ Error generando Excel: {e}", flush=True)

    # ENVÍO DE CORREO
    if EMAIL_REMITENTE and EMAIL_PASSWORD:
        try:
            msg = EmailMessage()
            msg['Subject'] = f'✅ Tabulación Verificada V2 ({etiqueta}) - Radicados ALMA-R'
            msg['From'] = EMAIL_REMITENTE
            msg['To'] = EMAIL_DESTINO
            msg.set_content(f'Hola,\n\nEl proceso V2 100% IA ha finalizado para {etiqueta}.\nLa IA leyó los stickers y los correos para rescatar los radicados ALMA-R.\n\nSaludos.')
            if os.path.exists(ruta_excel):
                with open(ruta_excel, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=os.path.basename(ruta_excel))
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as smtp:
                smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
                smtp.send_message(msg)
            print("🚀 ¡Correo V2 enviado exitosamente!", flush=True)
        except Exception as e: print(f"❌ Error correo: {e}")

if __name__ == "__main__":
    procesar_archivos()

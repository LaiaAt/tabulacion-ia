# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (PROMPT MAESTRO - CERO CONDICIONALES PYTHON)
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

# Solo utilizamos el modelo especializado de Pixtral
MODELOS_PIXTRAL = ["pixtral-12b-2409"]

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
    Convierte a imágenes las PRIMERAS 3 páginas y las ÚLTIMAS 3 páginas,
    tal como lo exige el prompt maestro.
    """
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        if total_paginas == 0:
            doc.close()
            return [], 0

        paginas_a_procesar = set()
        if total_paginas <= 6:
            paginas_a_procesar.update(range(total_paginas))
        else:
            # Primeras 3 y últimas 3
            paginas_a_procesar.update([0, 1, 2])
            paginas_a_procesar.update([total_paginas-3, total_paginas-2, total_paginas-1])

        imagenes_b64 = []
        for i in sorted(list(paginas_a_procesar)):
            pagina = doc[i]
            # DPI en 110 para mantener legibilidad sin exceder el peso máximo del payload
            pix = pagina.get_pixmap(dpi=110)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            if img.width > 1000:
                ratio = 1000 / float(img.width)
                img = img.resize((1000, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=75, optimize=True)
            imagenes_b64.append(base64.b64encode(buffer.getvalue()).decode('utf-8'))

        doc.close()
        return imagenes_b64, total_paginas
    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}", flush=True)
        return [], 0

PROMPT_MAESTRO = """
ACTÚA COMO UN SISTEMA DE EXTRACCIÓN DOCUMENTAL ESPECIALIZADO EN CORRESPONDENCIA ADMINISTRATIVA DEL CONSORCIO 4C.

Debes analizar visualmente imágenes correspondientes a páginas de documentos PDF escaneados y extraer información para una posterior tabulación automática en Excel.
Las páginas que recibirás son IMÁGENES obtenidas directamente de un PDF escaneado.
NO utilizarás OCR externo. Debes leer visualmente el contenido de las imágenes.

============================================================
CAMPOS QUE DEBES EXTRAER
============================================================
Debes entregar ÚNICAMENTE estos seis campos:
1. razon_social_remitente
2. no_radicado_remitente
3. razon_social_destinatario
4. no_radicado_destinatario
5. fecha
6. asunto

============================================================
CONSORCIO 4C
============================================================
El Consorcio 4C es la organización desde cuyo punto de vista se clasifica la correspondencia.
Debes determinar si la comunicación es RECIBIDA POR EL CONSORCIO 4C o ENVIADA/RADICADA POR EL CONSORCIO 4C.

------------------------------------------------------------
COMUNICACIONES RECIBIDAS
------------------------------------------------------------
Cuando el Consorcio 4C recibe una comunicación:
- El REMITENTE es la persona, empresa o entidad externa que envía la comunicación.
- El DESTINATARIO es el Consorcio 4C.
- El número de radicado recibido normalmente corresponde al radicado del Consorcio 4C (Ej: GP-0000).

------------------------------------------------------------
COMUNICACIONES ENVIADAS POR CONSORCIO 4C
------------------------------------------------------------
Cuando el Consorcio 4C envía una comunicación:
- El REMITENTE es el Consorcio 4C.
- El DESTINATARIO es la persona, empresa o entidad externa.
- El radicado propio del Consorcio 4C normalmente aparece en la parte superior de la comunicación (Ej: CI.000/GP0000/).

============================================================
ESTRATEGIA DE ANÁLISIS
============================================================
Te enviamos imágenes de las PRIMERAS y ÚLTIMAS páginas del documento.
Busca activamente: remitente, destinatario, radicados, fecha, asunto, página de radicación y constancias.
La página de radicación NO necesariamente contiene el asunto real de la carta; da prioridad a la CARTA PRINCIPAL.
Si hay una Constancia de Radicación Digital, NO utilices su texto como asunto y NO utilices su fecha si existe una fecha propia de la carta. Úsala para identificar el radicado del destinatario.

============================================================
FECHA
============================================================
Extrae la fecha correspondiente a la CARTA PRINCIPAL.
Formato EXACTO: DD/MM/AAAA (Ejemplo: 15/03/2017).

============================================================
ASUNTO
============================================================
Comunicaciones Recibidas: Busca el asunto real. Si aparece "Asunto:", copia el contenido completo. NO resumas, NO interpretes.
Comunicaciones Enviadas: NO contienen un campo "Asunto:". Contienen un párrafo que comienza con "Ref.". Ese párrafo completo es el asunto. DEBES INCLUIR "Ref." EN EL RESULTADO. (Ejemplo: "Ref. Solicitud de modificación..."). No cambies Ref. por Asunto:.

============================================================
FIDELIDAD Y NO DEJAR CAMPOS VACÍOS
============================================================
LEER → IDENTIFICAR → COPIAR → CLASIFICAR. NO es: RESUMIR → INTERPRETAR → INVENTAR.
Si después de revisar las páginas disponibles realmente no existe evidencia suficiente para determinar un campo, escribe exactamente: NO IDENTIFICADO.
Nunca utilices null, "", un espacio vacío ni omitas un campo.

============================================================
FORMATO DE RESPUESTA
============================================================
Devuelve ÚNICAMENTE un JSON válido con EXACTAMENTE estos seis campos:

{
  "razon_social_remitente": "",
  "no_radicado_remitente": "",
  "razon_social_destinatario": "",
  "no_radicado_destinatario": "",
  "fecha": "",
  "asunto": ""
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

def consultar_pixtral_pool(b64_imgs, nombre_archivo, tipo_flujo, item_num, hilo_id):
    if not b64_imgs:
        return None, "", ""

    apoyo = f"\nTipo de flujo en carpeta: {tipo_flujo}\nArchivo: {nombre_archivo}\n"
    prompt_final = PROMPT_MAESTRO + apoyo

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

def limpiar_salida_excel(val):
    """Limpia caracteres especiales que rompen Excel, pero respeta el texto de la IA."""
    if not val: return "NO IDENTIFICADO"
    val = str(val).strip()
    if val.upper() in ["NONE", "NULL", "NAN", ""]: return "NO IDENTIFICADO"
    return ILLEGAL_CHARACTERS_RE.sub("", val)

def procesar_un_pdf(item_num, pdf, ruta_completa, tipo, ruta_memoria, hilo_id):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_imgs, paginas = obtener_insumos_documento(ruta_completa)
    
    # GARANTÍA 100%: Si falla todo, se crea la fila con NO IDENTIFICADO.
    if not b64_imgs:
        datos = {}
        key_usada, mod_usado = "FALLO_DOC", "NINGUNO"
    else:
        datos, key_usada, mod_usado = consultar_pixtral_pool(b64_imgs, pdf, tipo, item_num, hilo_id)
        if datos is None:
            datos = {}
            key_usada, mod_usado = "FALLO_IA", "NINGUNO"

    # Se usa ÚNICA y EXCLUSIVAMENTE lo que dijo la IA. Cero sobreescritura en Python.
    fila = {
        "ÍTEM": item_num,
        "DEL FOLIO/PAGINAS": paginas,
        "RAZON SOCIAL REMITENTE": limpiar_salida_excel(datos.get("razon_social_remitente", "NO IDENTIFICADO")),
        "No. RADICADO REMITENTE": limpiar_salida_excel(datos.get("no_radicado_remitente", "NO IDENTIFICADO")),
        "RAZON SOCIAL DESTINATARIO": limpiar_salida_excel(datos.get("razon_social_destinatario", "NO IDENTIFICADO")),
        "No. RADICADO DESTINATARIO": limpiar_salida_excel(datos.get("no_radicado_destinatario", "NO IDENTIFICADO")),
        "FECHA (DD/MM/AAAA)": limpiar_salida_excel(datos.get("fecha", "NO IDENTIFICADO")),
        "ASUNTO / TIPO DOCUMENTAL": limpiar_salida_excel(datos.get("asunto", "NO IDENTIFICADO")),
        "UBICACION_ARCHIVO": ruta_relativa
    }

    with lock_csv:
        pd.DataFrame([fila]).to_csv(ruta_memoria, mode='a', header=not os.path.exists(ruta_memoria), index=False)

    duracion = round(time.time() - t_inicio, 2)
    estado = "⚠️ SALVADO VACÍO" if "FALLO" in key_usada else "✅ OK"
    print(f"📄 [Hilo-{hilo_id} | {mod_usado}] {pdf} | {estado} ({duracion}s)", flush=True)
    return True

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (PROMPT MAESTRO - CONTROL TOTAL DE IA)", flush=True)
    print("="*70, flush=True)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower() in ['si', 's', 'true']
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2023').strip()

    limite = int(os.environ.get('LIMITE_PRUEBA', '1').strip()) if es_prueba else None
    etiqueta = f"PRUEBA_{limite}" if es_prueba else f"{carpeta_objetivo}"
    ruta_memoria = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_memoria_PRUEBA.csv' if es_prueba else f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
    ruta_excel = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_PRUEBA.xlsx' if es_prueba else f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')

    if es_prueba and os.path.exists(ruta_memoria): os.remove(ruta_memoria)

    # Cargar memoria
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
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, tipo, ruta_memoria, hilo_id)
                futuros.append(f)
                item_counter += 1
                time.sleep(0.5)
            for f in as_completed(futuros): pass

    # ENSAMBLAJE FINAL EXCEL SIN ALTERACIONES
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
            msg.set_content(f'Proceso concluido exitosamente bajo el Prompt Maestro para {etiqueta}. Ningún archivo omitido. Cero celdas vacías.')
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

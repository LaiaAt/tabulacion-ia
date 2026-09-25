# ==============================================================================
# SISTEMA DE RECTIFICACIÓN QUIRÚRGICA v2 - RADICADOS ALTO MAGDALENA
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
CARPETA_OBJETIVO = os.environ.get('CARPETA_OBJETIVO', '2022').strip()

lock_csv = threading.Lock()
lock_keys = threading.Lock()
cooldown_keys = {k: 0.0 for k in lista_keys}

def obtener_imagenes_inspeccion(ruta_pdf):
    """
    Extrae la Página 1 (donde va el sticker físico en la esquina superior derecha)
    y la(s) última(s) página(s) o páginas de constancia (donde va la radicación digital/correo).
    """
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        if total_paginas == 0:
            doc.close()
            return []

        paginas_a_procesar = [0]  # Página 1 (obligatoria para sticker físico)

        # Si el documento tiene más páginas, evaluar páginas intermedias con constancias y la última
        for idx in range(1, total_paginas):
            txt = doc[idx].get_text().lower()
            if any(k in txt for k in ["radicó con éxito", "número de radicado", "alma-r-", "ventanilla", "recibido", "correo"]):
                paginas_a_procesar.append(idx)

        # Incluir siempre la última página por si fue radicado digital/correo
        if total_paginas - 1 not in paginas_a_procesar:
            paginas_a_procesar.append(total_paginas - 1)

        paginas_a_procesar = sorted(list(set(paginas_a_procesar)))

        imagenes_b64 = []
        for i in paginas_a_procesar:
            pagina = doc[i]
            # Buena resolución (150 DPI) para leer códigos pequeños de stickers
            pix = pagina.get_pixmap(dpi=150)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            
            # Limitar tamaño manteniendo legibilidad
            if img.width > 1600:
                ratio = 1600 / float(img.width)
                img = img.resize((1600, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)
                
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=85, optimize=True)
            imagenes_b64.append(base64.b64encode(buffer.getvalue()).decode('utf-8'))

        doc.close()
        return imagenes_b64
    except Exception as e:
        print(f"⚠️ Error abriendo PDF {ruta_pdf}: {e}", flush=True)
        return []

PROMPT_ESPECIALIZADO_ALMA = """
ACTÚA COMO UN AUDITOR Y TRANSSCRIPTOR DOCUMENTAL FORENSE EXPERTO.
Tu ÚNICA Y EXCLUSIVA misión es localizar y transcribir el NÚMERO DE RADICADO asignado por CONCESIÓN ALTO MAGDALENA S.A.S.

UBICACIÓN VISUAL OBLIGATORIA:
1. RADICACIÓN FÍSICA (Página 1):
   - Examina con extremo detalle la ESQUINA SUPERIOR DERECHA de la carta.
   - Encontrarás un sticker adhesivo blanco pegado que dice: "CONCESION ALTO MAGDALENA S.A.S." con fecha, hora, un código de barras y un radicado impreso.
   - Ejemplo de código en el sticker: "ALMA-R-2022-1379", "ALMA-R-2017-00392", "ALMA-R-2022-1026", "ALMA-R-2017-01700".

2. RADICACIÓN DIGITAL (Últimas Páginas):
   - Si no hay sticker físico en la página 1, revisa las constancias de radicación digital, ventanilla única virtual o comprobante de correo en las últimas páginas.

REGLAS DE ORO OBLIGATORIAS:
- El radicado de Alto Magdalena SIEMPRE debe comenzar con el prefijo "ALMA-", comúnmente con el formato "ALMA-R-AAAA-XXXX".
- CUALQUIER COSA QUE NO COMIENCE CON "ALMA-" DEBE SER OMITIDA.
- NUNCA pongas radicados de Consorcio 4C (como CI.004...).
- Si el documento realmente no tiene ningún sticker ni constancia con prefijo "ALMA-", responde exactamente: "SIN NÚMERO".

Devuelve ÚNICAMENTE un JSON con esta estructura:
{
  "no_radicado_destinatario": "ALMA-R-..."
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

def consultar_pixtral_alma(b64_imgs, item_idx):
    if not b64_imgs:
        return "SIN NÚMERO"

    num_keys = len(lista_keys)
    start_key_idx = item_idx % num_keys

    for ronda in range(3):
        for intento in range(num_keys):
            idx = (start_key_idx + intento) % num_keys
            k_actual = lista_keys[idx]

            with lock_keys:
                if cooldown_keys[k_actual] > time.time(): continue

            headers = {
                "Authorization": f"Bearer {k_actual}",
                "Content-Type": "application/json"
            }

            content_array = [{"type": "text", "text": PROMPT_ESPECIALIZADO_ALMA}]
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
                            rad = str(d.get("no_radicado_destinatario", "")).strip().replace(" ", "-")
                            # VALIDACIÓN ESTRICTA: SI NO INICIA CON ALMA-, SE OMITE
                            if rad.upper().startswith("ALMA-"):
                                return rad.upper()
                            return "SIN NÚMERO"
                    elif resp.status_code == 429:
                        with lock_keys: cooldown_keys[k_actual] = time.time() + 8
                        break
                except Exception:
                    continue

        time.sleep(5)

    return "SIN NÚMERO"

def rectificar_registro(idx, row, ruta_memoria):
    nom_rel = str(row.get("UBICACION_ARCHIVO", "")).replace('\\', '/').strip()
    ruta_pdf = os.path.join(RUTA_BASE, nom_rel)

    if not os.path.exists(ruta_pdf):
        # Intentar ruta alternativa
        nombre_base = os.path.basename(nom_rel)
        ruta_pdf_alt = os.path.join(RUTA_BASE, "15_01_Cartas_Enviadas", nombre_base)
        if os.path.exists(ruta_pdf_alt):
            ruta_pdf = ruta_pdf_alt
        else:
            return idx, None, f"⚠️ No encontrado: {nombre_base}"

    b64_imgs = obtener_imagenes_inspeccion(ruta_pdf)
    if not b64_imgs:
        return idx, None, f"⚠️ Error renderizando: {os.path.basename(ruta_pdf)}"

    rad_obtenido = consultar_pixtral_alma(b64_imgs, idx)

    # DOBLE VALIDACIÓN: Debe iniciar con ALMA-
    if not rad_obtenido.upper().startswith("ALMA-"):
        rad_obtenido = "SIN NÚMERO"

    return idx, rad_obtenido, f"📄 {os.path.basename(ruta_pdf)} ➡️ {rad_obtenido}"

def sanitizar_para_excel(val):
    if not val or pd.isnull(val): return ""
    s = str(val)
    s = ILLEGAL_CHARACTERS_RE.sub("", s)
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', s)
    return s.strip()

def procesar_rectificacion():
    print("\n" + "="*70, flush=True)
    print(f" MOTOR RECTIFICADOR v2 (RADICADOS ALTO MAGDALENA - {CARPETA_OBJETIVO})", flush=True)
    print("="*70, flush=True)

    ruta_memoria = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_memoria_{CARPETA_OBJETIVO}.csv')
    ruta_excel = os.path.join(RUTA_BASE, f'RESTREPO_2_IA_{CARPETA_OBJETIVO}.xlsx')

    if not os.path.exists(ruta_memoria):
        print(f"❌ No se encontró el archivo de memoria {ruta_memoria}", flush=True)
        return

    df_m = pd.read_csv(ruta_memoria)
    if df_m.empty:
        print("❌ La memoria está vacía.", flush=True)
        return

    # Filtro: Radicadas + Concesión Alto Magdalena + Radicado 'SIN NÚMERO' o no empieza con ALMA-
    es_rad = ~df_m['UBICACION_ARCHIVO'].astype(str).str.contains('Recibidas', case=False, na=False)
    dest_alma = df_m['RAZON SOCIAL DESTINATARIO'].astype(str).str.contains('MAGDALENA|ALMA', case=False, na=False)
    rad_dest_actual = df_m['No. RADICADO DESTINATARIO'].fillna('').astype(str).str.strip()
    rad_invalido = (rad_dest_actual.str.upper() == 'SIN NÚMERO') | (~rad_dest_actual.str.upper().str.startswith('ALMA-'))

    indices_objetivo = df_m[es_rad & dest_alma & rad_invalido].index.tolist()
    total_objetivo = len(indices_objetivo)

    print(f"🔍 Registros identificados para rectificación: {total_objetivo}", flush=True)
    if total_objetivo == 0:
        print("✅ Todos los radicados de Alto Magdalena ya están correctamente tabulados.", flush=True)
        return

    num_trabajadores = min(len(lista_keys) * 2, 4)
    modificados = 0

    with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
        futuros = [
            executor.submit(rectificar_registro, idx, df_m.loc[idx], ruta_memoria)
            for idx in indices_objetivo
        ]

        for fut in as_completed(futuros):
            idx, nuevo_rad, log_msg = fut.result()
            print(log_msg, flush=True)

            if nuevo_rad and nuevo_rad != "SIN NÚMERO":
                with lock_csv:
                    # ACTUALIZACIÓN QUIRÚRGICA: SOLO ESTA COLUMNA
                    df_m.at[idx, "No. RADICADO DESTINATARIO"] = nuevo_rad
                    modificados += 1

    # Guardar CSV de memoria actualizado
    df_m.to_csv(ruta_memoria, index=False)
    print(f"\n💾 Memoria actualizada: {modificados} radicados ALMA- recuperados y asignados.", flush=True)

    # Actualizar Excel final
    try:
        es_recibida = df_m["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
        df_rec = df_m[es_recibida].copy()
        df_rad = df_m[~es_recibida].copy()

        for c in df_rec.columns: df_rec[c] = df_rec[c].apply(sanitizar_para_excel)
        for c in df_rad.columns: df_rad[c] = df_rad[c].apply(sanitizar_para_excel)

        with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
            df_rec.to_excel(writer, sheet_name="Recibidas", index=False)
            df_rad.to_excel(writer, sheet_name="Radicadas", index=False)

        print(f"📊 Excel regenerado con éxito: {ruta_excel}", flush=True)
    except Exception as e:
        print(f"⚠️ Error actualizando Excel: {e}", flush=True)

    # Envío de correo con confirmación
    if EMAIL_REMITENTE and EMAIL_PASSWORD:
        try:
            msg = EmailMessage()
            msg['Subject'] = f'✅ Rectificación de Radicados ALMA- ({CARPETA_OBJETIVO})'
            msg['From'] = EMAIL_REMITENTE
            msg['To'] = EMAIL_DESTINO
            msg.set_content(
                f'Hola,\n\n'
                f'Se ha completado la rectificación visual de radicados para {CARPETA_OBJETIVO}.\n'
                f'- Cartas inspeccionadas: {total_objetivo}\n'
                f'- Radicados ALMA- recuperados exitosamente: {modificados}\n\n'
                f'El resto de los datos en la memoria y en el Excel permanecieron intactos.\n\n'
                f'Saludos cordiales.'
            )
            if os.path.exists(ruta_excel):
                with open(ruta_excel, 'rb') as f:
                    msg.add_attachment(
                        f.read(),
                        maintype='application',
                        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        filename=os.path.basename(ruta_excel)
                    )
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30) as smtp:
                smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
                smtp.send_message(msg)
            print("🚀 Correo con reporte enviado exitosamente.", flush=True)
        except Exception as e:
            print(f"❌ Error al enviar correo: {e}", flush=True)

if __name__ == "__main__":
    procesar_rectificacion()

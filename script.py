# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 - VERSIÓN VISIÓN PURA (Pixtral Large + 12B)
# Sin OCR - Multi-página inteligente - 100% Vision
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

# Orden: primero el más potente, luego el más barato
MODELOS_FASE_TURBO = ["pixtral-large-latest", "pixtral-12b-2409"]
MODELOS_AUDITORIA = ["pixtral-large-latest", "pixtral-12b-2409"]

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
    Extrae hasta 3 páginas clave como imágenes JPEG (100% visión, sin OCR).
    - Página 0
    - Página 1 (si existe)
    - Última página (para capturar radicados que quedaron al final)
    """
    try:
        doc = fitz.open(ruta_pdf)
        total_paginas = len(doc)
        if total_paginas == 0:
            doc.close()
            return [], 0

        indices = set([0])
        if total_paginas > 1:
            indices.add(1)
        if total_paginas > 2:
            indices.add(total_paginas - 1)

        indices = sorted(list(indices))[:3]

        imagenes_b64 = []
        for idx in indices:
            pagina = doc[idx]
            pix = pagina.get_pixmap(dpi=150)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")

            if img.width > 1400:
                ratio = 1400 / float(img.width)
                img = img.resize((1400, int(img.height * ratio)), Image.Resampling.LANCZOS)

            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=82, optimize=True)
            b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
            imagenes_b64.append(b64)

        doc.close()
        return imagenes_b64, total_paginas

    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}", flush=True)
        return [], 0


PROMPT_AUDITORIA = """
Eres un experto en extracción de datos de cartas oficiales colombianas (ANI, Consorcio 4C, interventorías, CAR, etc.).

CONTEXTO CRÍTICO:
- Los PDFs suelen estar mal fusionados. La página de radicación (sello/sticker) puede estar en la PRIMERA o en la ÚLTIMA página.
- Siempre existe una carta real con membrete, ciudad, fecha, destinatario y asunto/referencia.
- Recibirás varias imágenes del mismo documento (páginas clave). Úsalas todas.

REGLAS DE ORO:
1. Copia ÚNICAMENTE lo que ves escrito. Si no está claro o no aparece → deja "".
2. NUNCA inventes números de radicado, fechas ni asuntos.
3. Distingue perfectamente:
   - REMITENTE = quien envía (tiene el membrete o firma).
   - DESTINATARIO = a quien va dirigida la carta.

CAMPOS EXACTOS A EXTRAER:

- RAZON_SOCIAL_REMITENTE: Solo el nombre de la empresa/entidad que envía. Sin nombres de personas ni cargos.
- NO_RADICADO_REMITENTE: Código de la carta del remitente (ej: CI.004/0427/17/2.1). Está casi siempre arriba a la derecha en el membrete.
- RAZON_SOCIAL_DESTINATARIO: Solo el nombre de la entidad receptora (ej: AGENCIA NACIONAL DE INFRAESTRUCTURA). Sin "Atn.", sin nombres de personas.
- NO_RADICADO_DESTINATARIO: Número de radicado de recepción (sello, sticker o código tipo GP-XXXX).
- FECHA: Fecha de la carta en formato DD/MM/AAAA.
- ASUNTO: Texto literal del asunto o referencia. 
  Ejemplo correcto: "Ref.: Contrato de Interventoría 145 de 2014 Interventoría Concesión Vial Girardot – Honda – Puerto Salgar"
  Si dice solo "ASUNTO: Reemplazo Auxiliar Administrativa" → pon "Reemplazo Auxiliar Administrativa".

Devuelve ÚNICAMENTE este JSON válido (sin markdown, sin texto extra):

{
  "RAZON_SOCIAL_REMITENTE": "",
  "NO_RADICADO_REMITENTE": "",
  "RAZON_SOCIAL_DESTINATARIO": "",
  "NO_RADICADO_DESTINATARIO": "",
  "FECHA": "",
  "ASUNTO": ""
}
"""


def parsear_json(texto):
    if not texto:
        return None
    try:
        t = str(texto).strip()
        t = re.sub(r'```(?:json)?', '', t).replace('```', '').strip()
        start, end = t.find('{'), t.rfind('}')
        if start != -1 and end != -1:
            return json.loads(t[start:end+1])
        return json.loads(t)
    except Exception:
        return None


def consultar_pixtral_pool(imagenes_b64, nombre_archivo, tipo_flujo, item_num, hilo_id, modelos=MODELOS_FASE_TURBO):
    """
    Recibe lista de imágenes base64.
    Prueba primero pixtral-large-latest, luego pixtral-12b-2409.
    """
    if not imagenes_b64:
        return None, "", ""

    apoyo = f"\nTipo de flujo: {tipo_flujo}\nNombre del archivo: {nombre_archivo}"
    prompt_final = PROMPT_AUDITORIA + apoyo

    content = [{"type": "text", "text": prompt_final}]
    for b64 in imagenes_b64:
        content.append({
            "type": "image_url",
            "image_url": f"data:image/jpeg;base64,{b64}"
        })

    num_keys = len(lista_keys)
    start_key_idx = (item_num + hilo_id) % num_keys

    for intento in range(num_keys):
        idx = (start_key_idx + intento) % num_keys
        k_actual = lista_keys[idx]

        with lock_keys:
            if cooldown_keys[k_actual] > time.time():
                continue

        headers = {
            "Authorization": f"Bearer {k_actual}",
            "Content-Type": "application/json"
        }

        for mod in modelos:
            try:
                payload = {
                    "model": mod,
                    "temperature": 0.0,
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {"role": "user", "content": content}
                    ]
                }
                resp = requests.post(
                    "https://api.mistral.ai/v1/chat/completions",
                    headers=headers,
                    json=payload,
                    timeout=70
                )
                if resp.status_code == 200:
                    d = parsear_json(resp.json()["choices"][0]["message"]["content"])
                    if d and isinstance(d, dict):
                        return d, f"Key-{idx+1}", mod
                elif resp.status_code == 429:
                    with lock_keys:
                        cooldown_keys[k_actual] = time.time() + 10
                    break
            except Exception:
                continue

    return None, "", ""


def limpieza_estricta_blindada(datos, nombre_archivo, tipo_flujo):
    """Limpieza de negocio. Ya no depende de texto OCR."""
    if not isinstance(datos, dict):
        datos = {}

    def clean(val):
        if not val or str(val).strip().upper() in ["NONE", "N/A", "NULL", "NAN", ""]:
            return ""
        return ILLEGAL_CHARACTERS_RE.sub("", str(val).strip())

    ia_dest = clean(datos.get("RAZON_SOCIAL_DESTINATARIO", ""))
    ia_rem = clean(datos.get("RAZON_SOCIAL_REMITENTE", ""))
    rad_rem = clean(datos.get("NO_RADICADO_REMITENTE", ""))
    rad_dest = clean(datos.get("NO_RADICADO_DESTINATARIO", ""))
    ia_asunto = clean(datos.get("ASUNTO", ""))
    ia_fecha = clean(datos.get("FECHA", ""))

    # Limpieza destinatario
    ia_dest = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente|Dr\.|Dra\.).*', '', ia_dest).strip()

    # Limpieza asunto
    ia_asunto = re.sub(r'^[\[\(\.\,\-\_\s\"\'\\]+', '', ia_asunto)
    ia_asunto = re.sub(r'^ASUNTO\s*[:\-\.]*\s*', '', ia_asunto, flags=re.IGNORECASE).strip()
    ia_asunto = " ".join(ia_asunto.split())

    es_recibida = tipo_flujo.upper() == "RECIBIDAS"

    if es_recibida:
        # REGLAS RECIBIDAS
        ia_dest = "CONSORCIO 4C"

        if "CONSORCIO 4C" in ia_rem.upper():
            ia_rem = ""

        if "CI.004" in rad_rem.upper() or rad_rem.upper().startswith("GP-"):
            rad_rem = ""

        # Intentar rescatar GP del nombre del archivo
        if not rad_dest.startswith("GP-"):
            m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
            if m_gp:
                rad_dest = f"GP-{m_gp.group(1)}"
    else:
        # REGLAS RADICADAS / ENVIADAS
        ia_rem = "CONSORCIO 4C"

        if "CONSORCIO 4C" in ia_dest.upper():
            ia_dest = ""

        # Intentar rescatar CI.004 del nombre del archivo si la IA falló
        if not rad_rem or "CI.004" not in rad_rem.upper():
            m_ci = re.search(r'CI\.?004[/_-]?(\d{4})', nombre_archivo, re.IGNORECASE)
            if m_ci:
                # No inventamos el resto, solo dejamos vacío si no vino de la IA
                pass

    return {
        "RAZON_SOCIAL_DESTINATARIO": ia_dest,
        "RAZON_SOCIAL_REMITENTE": ia_rem,
        "NO_RADICADO_REMITENTE": rad_rem,
        "NO_RADICADO_DESTINATARIO": rad_dest,
        "ASUNTO": ia_asunto,
        "FECHA": ia_fecha
    }


def procesar_un_pdf(item_num, pdf, ruta_completa, anio_doc, tipo, ruta_memoria, hilo_id):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    imagenes_b64, paginas = obtener_insumos_documento(ruta_completa)

    if not imagenes_b64:
        datos_completos = limpieza_estricta_blindada({}, pdf, tipo)
        key_usada, mod_usado = "FALLO_DOC", "NINGUNO"
    else:
        datos, key_usada, mod_usado = consultar_pixtral_pool(
            imagenes_b64, pdf, tipo, item_num, hilo_id
        )
        if datos is None:
            datos_completos = limpieza_estricta_blindada({}, pdf, tipo)
            key_usada, mod_usado = "FALLO_IA", "NINGUNO"
        else:
            datos_completos = limpieza_estricta_blindada(datos, pdf, tipo)

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
    estado = "⚠️ SALVADO VACÍO" if "FALLO" in key_usada else "✅ OK"
    print(f"📄 [Hilo-{hilo_id} | {mod_usado}] {pdf} | {estado} ({duracion}s)", flush=True)
    return True


def auditar_y_corregir_tabla_final(ruta_memoria):
    if not os.path.exists(ruta_memoria):
        return
    df_mem = pd.read_csv(ruta_memoria)
    if df_mem.empty:
        return

    print("\n🧐 [FASE 2: AUDITORÍA] Re-evaluando celdas vacías...", flush=True)

    filas_vacias = df_mem[df_mem['ASUNTO / TIPO DOCUMENTAL'].fillna('').str.strip() == ''].index.tolist()

    if not filas_vacias:
        print("✅ Control de Calidad: Todas las filas tienen Asunto.", flush=True)
        return

    print(f"⚠️ Detectadas {len(filas_vacias)} cartas vacías. Auditando con Pixtral Large...", flush=True)

    for i, idx in enumerate(filas_vacias, 1):
        row = df_mem.loc[idx]
        ubic_rel = str(row["UBICACION_ARCHIVO"]).strip()
        ruta_pdf_completa = os.path.join(RUTA_BASE, ubic_rel)
        nombre_pdf = os.path.basename(ubic_rel)
        tipo_flujo = "RECIBIDAS" if "recibidas" in ubic_rel.lower() else "RADICADAS"

        if not os.path.exists(ruta_pdf_completa):
            continue

        imagenes_b64, _ = obtener_insumos_documento(ruta_pdf_completa)
        datos, _, _ = consultar_pixtral_pool(
            imagenes_b64, nombre_pdf, tipo_flujo, i, 1, modelos=MODELOS_AUDITORIA
        )

        if datos:
            datos_pulidos = limpieza_estricta_blindada(datos, nombre_pdf, tipo_flujo)
            for col, key in zip(
                ["RAZON SOCIAL REMITENTE", "No. RADICADO REMITENTE", "RAZON SOCIAL DESTINATARIO",
                 "No. RADICADO DESTINATARIO", "FECHA (DD/MM/AAAA)", "ASUNTO / TIPO DOCUMENTAL"],
                ["RAZON_SOCIAL_REMITENTE", "NO_RADICADO_REMITENTE", "RAZON_SOCIAL_DESTINATARIO",
                 "NO_RADICADO_DESTINATARIO", "FECHA", "ASUNTO"]
            ):
                if datos_pulidos.get(key):
                    df_mem.at[idx, col] = datos_pulidos[key]
            time.sleep(1.2)

    df_mem.to_csv(ruta_memoria, index=False)
    print("🎉 Auditoría Final completada.", flush=True)


def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 - VISIÓN PURA (Pixtral Large + 12B)", flush=True)
    print("="*70, flush=True)

    es_prueba = os.environ.get('ES_PRUEBA', 'no').strip().lower() in ['si', 's', 'true']
    carpeta_objetivo = os.environ.get('CARPETA_OBJETIVO', '2023').strip()
    reiniciar = os.environ.get('REINICIAR_MEMORIA', 'no').strip().lower() in ['si', 's', 'true']

    limite = int(os.environ.get('LIMITE_PRUEBA', '1').strip()) if es_prueba else None
    etiqueta = f"PRUEBA_{limite}" if es_prueba else f"{carpeta_objetivo}"
    ruta_memoria = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_memoria_PRUEBA.csv' if es_prueba else f'RESTREPO_2_IA_memoria_{carpeta_objetivo}.csv')
    ruta_excel = os.path.join(RUTA_BASE, 'RESTREPO_2_IA_PRUEBA.xlsx' if es_prueba else f'RESTREPO_2_IA_{carpeta_objetivo}.xlsx')

    if es_prueba and os.path.exists(ruta_memoria):
        os.remove(ruta_memoria)

    if reiniciar and os.path.exists(ruta_memoria):
        os.remove(ruta_memoria)
        print("🗑️ Memoria borrada. Reiniciando desde cero.", flush=True)

    procesados_basenames = set()
    item_counter = 1
    if not es_prueba and os.path.exists(ruta_memoria):
        try:
            df_m = pd.read_csv(ruta_memoria)
            procesados_basenames = set(os.path.basename(str(r).strip()).lower() for r in df_m["UBICACION_ARCHIVO"].dropna())
            item_counter = len(df_m) + 1
        except Exception:
            pass

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
            continue

        if es_prueba:
            pendientes = pendientes[:limite]

        print(f"\n📂 Tabulando {len(pendientes)} cartas en {tipo}...", flush=True)
        with ThreadPoolExecutor(max_workers=num_trabajadores) as executor:
            futuros = []
            for i, (pdf, ruta_completa, anio_doc) in enumerate(pendientes):
                hilo_id = (i % num_trabajadores) + 1
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, anio_doc, tipo, ruta_memoria, hilo_id)
                futuros.append(f)
                item_counter += 1
                time.sleep(0.4)
            for f in as_completed(futuros):
                pass

    if not es_prueba:
        auditar_y_corregir_tabla_final(ruta_memoria)

    # Ensamblaje final Excel
    if os.path.exists(ruta_memoria):
        df_final = pd.read_csv(ruta_memoria)
        if not df_final.empty:
            es_recibida = df_final["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
            df_rec = df_final[es_recibida].copy()
            df_rad = df_final[\~es_recibida].copy()

            if not df_rec.empty:
                df_rec["ÍTEM"] = range(1, len(df_rec) + 1)
                for col in df_rec.columns:
                    df_rec[col] = df_rec[col].apply(lambda x: ILLEGAL_CHARACTERS_RE.sub("", str(x)) if pd.notnull(x) else "")
            if not df_rad.empty:
                df_rad["ÍTEM"] = range(1, len(df_rad) + 1)
                for col in df_rad.columns:
                    df_rad[col] = df_rad[col].apply(lambda x: ILLEGAL_CHARACTERS_RE.sub("", str(x)) if pd.notnull(x) else "")

            with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                df_rec.to_excel(writer, sheet_name="Recibidas", index=False)
                df_rad.to_excel(writer, sheet_name="Radicadas", index=False)

    # Envío de correo
    if EMAIL_REMITENTE and EMAIL_PASSWORD:
        try:
            msg = EmailMessage()
            msg['Subject'] = f'✅ Tabulación Completa 100% ({etiqueta})'
            msg['From'] = EMAIL_REMITENTE
            msg['To'] = EMAIL_DESTINO
            msg.set_content(f'Proceso concluido exitosamente para {etiqueta}.')
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
            print("🚀 ¡Correo enviado exitosamente!", flush=True)
        except Exception as e:
            print(f"❌ Error correo: {e}")


if __name__ == "__main__":
    procesar_archivos()

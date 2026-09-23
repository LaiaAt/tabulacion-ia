# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (INCLUSIÓN 100% + REGLAS ESTRICTAS)
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

MODELOS_FASE_TURBO = ["pixtral-12b-2409", "pixtral-large-latest"]
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
        es_caratula = ("al contestar cite el numero" in texto_pag1.lower() or "ci004_" in texto_pag1.lower())
        
        if es_caratula and total_paginas > 1:
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

PROMPT_AUDITORIA = """
Eres un transcriptor de datos EXTREMADAMENTE ESTRICTO.
Tu ÚNICA labor es copiar el texto visible. REGLA DE ORO: PROHIBIDO INVENTAR NÚMEROS O DATOS. Si no lo ves claramente, déjalo vacío "".

CAMPOS A EXTRAER:
1. "RAZON_SOCIAL_REMITENTE": Entidad que envía. NUNCA incluyas nombres de personas.
2. "NO_RADICADO_REMITENTE": Código oficial de envío. En cartas de Consorcio 4C suele estar en el encabezado (Ej: "CI.004/0371/25/1.4"). Cópialo idéntico.
3. "RAZON_SOCIAL_DESTINATARIO": SOLO el nombre de la empresa/entidad receptora. PROHIBIDO incluir "Atn.", "Ing.", "Gerente" o nombres de personas.
4. "NO_RADICADO_DESTINATARIO": Radicado de recepción (sticker o sello).
5. "FECHA": Fecha del documento (DD/MM/AAAA).
6. "ASUNTO": Transcripción literal y fiel. NO inventes "Ref.". Si la carta dice "Ref. XYZ", transcribe "Ref. XYZ". Si la carta dice "ASUNTO: XYZ", transcribe "XYZ" (sin la palabra asunto). NUNCA transcribas sellos aquí.

Devuelve SOLO JSON VÁLIDO:
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

def consultar_pixtral_pool(b64_img, texto_digital, nombre_archivo, tipo_flujo, item_num, hilo_id, modelos=MODELOS_FASE_TURBO):
    if not b64_img:
        return None, "", ""

    apoyo = f"\nTipo de flujo: {tipo_flujo}\nArchivo: {nombre_archivo}\nTexto base:\n{texto_digital[:2000]}"
    prompt_final = PROMPT_AUDITORIA + apoyo
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

        for mod in modelos:
            try:
                payload = {
                    "model": mod,
                    "temperature": 0.0,
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {"role": "user", "content": [{"type": "text", "text": prompt_final}, {"type": "image_url", "image_url": f"data:image/jpeg;base64,{b64_img}"}]}
                    ]
                }
                resp = requests.post("https://api.mistral.ai/v1/chat/completions", headers=headers, json=payload, timeout=45)
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

def limpieza_estricta_blindada(datos, nombre_archivo, texto_completo, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    def clean(val):
        if not val or str(val).strip().upper() in ["NONE", "N/A", "NULL", "NAN"]: return ""
        return ILLEGAL_CHARACTERS_RE.sub("", str(val).strip())

    ia_dest = clean(datos.get("RAZON_SOCIAL_DESTINATARIO", ""))
    ia_rem = clean(datos.get("RAZON_SOCIAL_REMITENTE", ""))
    rad_rem = clean(datos.get("NO_RADICADO_REMITENTE", ""))
    rad_dest = clean(datos.get("NO_RADICADO_DESTINATARIO", ""))
    ia_asunto = clean(datos.get("ASUNTO", ""))
    ia_fecha = clean(datos.get("FECHA", ""))

    # 1. PURGA ESTRICTA DE "ATN." EN EL DESTINATARIO
    ia_dest = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente).*', '', ia_dest).strip()

    # 2. LIMPIEZA DEL ASUNTO (Sin prefijos falsos, cero invenciones)
    ia_asunto = re.sub(r'^[\[\(\.\,\-\_\s\"\'\\]+', '', ia_asunto)
    ia_asunto = re.sub(r'^ASUNTO\s*[:\-\.]*\s*', '', ia_asunto, flags=re.IGNORECASE).strip()
    ia_asunto = " ".join(ia_asunto.split()) # Quitar saltos de línea innecesarios

    es_recibida = tipo_flujo == "RECIBIDAS"

    if es_recibida:
        # ======= REGLAS INQUEBRANTABLES PARA RECIBIDAS =======
        ia_dest = "CONSORCIO 4C"
        
        # El remitente NUNCA es Consorcio 4C
        if "CONSORCIO 4C" in ia_rem.upper():
            ia_rem = ""

        # Radicado Remitente de terceros NUNCA es CI.004 ni GP-
        if "CI.004" in rad_rem.upper() or "GP-" in rad_rem.upper():
            rad_rem = ""

        # Radicado Destinatario SIEMPRE es el GP del Consorcio
        if not rad_dest.startswith("GP-"):
            m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
            rad_dest = f"GP-{m_gp.group(1)}" if m_gp else rad_dest

    else:
        # ======= REGLAS INQUEBRANTABLES PARA RADICADAS =======
        ia_rem = "CONSORCIO 4C"

        # Destinatario NUNCA es Consorcio 4C
        if "CONSORCIO 4C" in ia_dest.upper():
            ia_dest = ""

        # Radicado Remitente SIEMPRE incluye CI.004... (Consorcio 4C)
        if not rad_rem or "CI.004" not in rad_rem:
            m_ci004 = re.search(r'(CI\.004/[^\s\n\r]+)', texto_completo, re.IGNORECASE)
            if m_ci004:
                rad_rem = m_ci004.group(1).strip()
            else:
                m_nom = re.search(r'CI004_(\d{4})\d{2}_', nombre_archivo, re.IGNORECASE)
                if m_nom:
                    rad_rem = f"GP-{m_nom.group(1).zfill(4)}"
                else:
                    m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
                    rad_rem = f"GP-{m_gp.group(1)}" if m_gp else ""

    # Rescate de fecha
    if not ia_fecha:
        m_f = re.search(r'(?:Bogot[aá]|Girardot|Honda)[^\n\r]*,?\s*(\d{1,2}\s*de\s*[a-zA-Z]+\s*de\s*\d{4}|\d{2}[-/.]\d{2}[-/.]\d{4})', texto_completo, re.IGNORECASE)
        if m_f:
            ia_fecha = m_f.group(1)

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

    b64_img, txt, txt1, paginas = obtener_insumos_documento(ruta_completa)
    
    # GARANTÍA 100%: Si falla todo, se crea la fila vacía con datos extraídos por Regex para auditarse después.
    if not b64_img:
        datos_completos = limpieza_estricta_blindada({}, pdf, txt, anio_doc, tipo)
        key_usada, mod_usado = "FALLO_DOC", "NINGUNO"
    else:
        datos, key_usada, mod_usado = consultar_pixtral_pool(b64_img, txt1, pdf, tipo, item_num, hilo_id)
        if datos is None:
            datos_completos = limpieza_estricta_blindada({}, pdf, txt, anio_doc, tipo)
            key_usada, mod_usado = "FALLO_IA", "NINGUNO"
        else:
            datos_completos = limpieza_estricta_blindada(datos, pdf, txt, anio_doc, tipo)

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
    if not os.path.exists(ruta_memoria): return
    df_mem = pd.read_csv(ruta_memoria)
    if df_mem.empty: return

    print("\n🧐 [FASE 2: AUDITORÍA DE CALIDAD] Re-evaluando celdas vacías por fallos de red...", flush=True)

    # Identificar filas que entraron en blanco por "FALLO_IA"
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

        if not os.path.exists(ruta_pdf_completa): continue

        b64_img, txt, txt1, _ = obtener_insumos_documento(ruta_pdf_completa)
        # Auditoría con el modelo pesado
        datos, _, _ = consultar_pixtral_pool(b64_img, txt1, nombre_pdf, tipo_flujo, i, 1, modelos=MODELOS_AUDITORIA)
        
        if datos:
            datos_pulidos = limpieza_estricta_blindada(datos, nombre_pdf, txt, "2023", tipo_flujo)
            for col, key in zip(
                ["RAZON SOCIAL REMITENTE", "No. RADICADO REMITENTE", "RAZON SOCIAL DESTINATARIO", "No. RADICADO DESTINATARIO", "FECHA (DD/MM/AAAA)", "ASUNTO / TIPO DOCUMENTAL"],
                ["RAZON_SOCIAL_REMITENTE", "NO_RADICADO_REMITENTE", "RAZON_SOCIAL_DESTINATARIO", "NO_RADICADO_DESTINATARIO", "FECHA", "ASUNTO"]
            ):
                if datos_pulidos.get(key):
                    df_mem.at[idx, col] = datos_pulidos[key]
            time.sleep(1.0)

    df_mem.to_csv(ruta_memoria, index=False)
    print("🎉 Auditoría Final completada.", flush=True)

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (INCLUSIÓN 100% Y REGLAS ESTRICTAS DE NEGOCIO)", flush=True)
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
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, anio_doc, tipo, ruta_memoria, hilo_id)
                futuros.append(f)
                item_counter += 1
                time.sleep(0.5)
            for f in as_completed(futuros): pass

    if not es_prueba:
        auditar_y_corregir_tabla_final(ruta_memoria)

    # ENSAMBLAJE FINAL EXCEL (Renumera los ítems de forma impecable sin huecos)
    if os.path.exists(ruta_memoria):
        df_final = pd.read_csv(ruta_memoria)
        if not df_final.empty:
            es_recibida = df_final["UBICACION_ARCHIVO"].str.contains("Recibidas", case=False, na=False)
            df_rec = df_final[es_recibida].copy()
            df_rad = df_final[~es_recibida].copy()

            if not df_rec.empty:
                df_rec["ÍTEM"] = range(1, len(df_rec) + 1)
                for col in df_rec.columns: df_rec[col] = df_rec[col].apply(lambda x: ILLEGAL_CHARACTERS_RE.sub("", str(x)) if pd.notnull(x) else "")
            if not df_rad.empty:
                df_rad["ÍTEM"] = range(1, len(df_rad) + 1)
                for col in df_rad.columns: df_rad[col] = df_rad[col].apply(lambda x: ILLEGAL_CHARACTERS_RE.sub("", str(x)) if pd.notnull(x) else "")

            with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
                df_rec.to_excel(writer, sheet_name="Recibidas", index=False)
                df_rad.to_excel(writer, sheet_name="Radicadas", index=False)

    # ENVÍO DE CORREO
    if EMAIL_REMITENTE and EMAIL_PASSWORD:
        try:
            msg = EmailMessage()
            msg['Subject'] = f'✅ Tabulación Completa 100% ({etiqueta})'
            msg['From'] = EMAIL_REMITENTE
            msg['To'] = EMAIL_DESTINO
            msg.set_content(f'Proceso concluido exitosamente y con reglas estrictas aplicadas para {etiqueta}.')
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

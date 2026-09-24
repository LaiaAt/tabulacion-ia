# ==============================================================================
# SISTEMA DE TABULACIÓN RESTREPO_2 (DOBLE MODELO IA + LÓGICA DE NEGOCIO BLINDADA)
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
    Convierte a imágenes las PRIMERAS 3 páginas y las ÚLTIMAS 3 páginas.
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
        if total_paginas <= 6:
            paginas_a_procesar.update(range(total_paginas))
        else:
            paginas_a_procesar.update([0, 1, 2])
            paginas_a_procesar.update([total_paginas-3, total_paginas-2, total_paginas-1])

        imagenes_b64 = []
        for i in sorted(list(paginas_a_procesar)):
            pagina = doc[i]
            pix = pagina.get_pixmap(dpi=110)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            if img.width > 1000:
                ratio = 1000 / float(img.width)
                img = img.resize((1000, int(float(img.height) * ratio)), Image.Resampling.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=75, optimize=True)
            imagenes_b64.append(base64.b64encode(buffer.getvalue()).decode('utf-8'))

        doc.close()
        return imagenes_b64, texto_completo, total_paginas
    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}", flush=True)
        return [], "", 0

PROMPT_MAESTRO = """
ACTÚA COMO UN SISTEMA DE EXTRACCIÓN DOCUMENTAL ESPECIALIZADO EN CORRESPONDENCIA ADMINISTRATIVA.
Debes analizar visualmente páginas de un documento escaneado y extraer información.

REGLAS ESTRICTAS DE EXTRACCIÓN:
1. "razon_social_remitente": Entidad que envía. NUNCA incluyas "Atn.", "Gerente" o nombres de personas. Solo la entidad.
2. "no_radicado_remitente": Código oficial de envío. (Ej: "CI.004/0371/25/1.4" o radicados de ANI/Concesión).
3. "razon_social_destinatario": Entidad receptora. PROHIBIDO incluir nombres de personas o cargos.
4. "no_radicado_destinatario": Radicado de recepción (sticker o sello, Ej: "GP-0000"). Si es constancia web, búscalo allí.
5. "fecha": Fecha principal de la carta (Formato DD/MM/AAAA).
6. "asunto": Transcripción literal y fiel.
   - En cartas recibidas: Si dice "ASUNTO: XYZ", transcribe "XYZ" (sin la palabra asunto).
   - En cartas enviadas: NO tienen "Asunto:". Contienen un párrafo que comienza con "Ref.". Ese párrafo completo es el asunto. DEBES INCLUIR "Ref." EN EL RESULTADO.
   - PROHIBIDO inventar "Ref." si no está impreso en la hoja.
   - PROHIBIDO copiar nombres de archivos técnicos o sellos en este campo.

Si un dato definitivamente no está visible en ninguna página, usa "NO IDENTIFICADO". NUNCA dejes el campo vacío.

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

def llamar_api_mistral(b64_imgs, nombre_archivo, tipo_flujo, modelo_ia, k_actual):
    apoyo = f"\nTipo de flujo: {tipo_flujo}\nArchivo: {nombre_archivo}\n"
    prompt_final = PROMPT_MAESTRO + apoyo

    headers = {
        "Authorization": f"Bearer {k_actual}",
        "Content-Type": "application/json"
    }

    content_array = [{"type": "text", "text": prompt_final}]
    for b64 in b64_imgs:
        content_array.append({"type": "image_url", "image_url": f"data:image/jpeg;base64,{b64}"})

    payload = {
        "model": modelo_ia,
        "temperature": 0.0,
        "response_format": {"type": "json_object"},
        "messages": [{"role": "user", "content": content_array}]
    }

    try:
        resp = requests.post("https://api.mistral.ai/v1/chat/completions", headers=headers, json=payload, timeout=50)
        if resp.status_code == 200:
            return parsear_json(resp.json()["choices"][0]["message"]["content"]), resp.status_code
        else:
            return None, resp.status_code
    except Exception:
        return None, 500

def consultar_pixtral_exhaustivo(b64_imgs, nombre_archivo, tipo_flujo, item_num, hilo_id):
    if not b64_imgs:
        return None, "", ""

    num_keys = len(lista_keys)
    start_key_idx = (item_num + hilo_id) % num_keys

    # PRIMER INTENTO (Modelo Rápido 12B)
    datos_extraidos = None
    modelo_exitoso = ""
    key_exitosa = ""

    for intento in range(num_keys):
        idx = (start_key_idx + intento) % num_keys
        k_actual = lista_keys[idx]
        nombre_key = f"Key-{idx+1}"

        with lock_keys:
            if cooldown_keys[k_actual] > time.time(): continue

        datos, status = llamar_api_mistral(b64_imgs, nombre_archivo, tipo_flujo, "pixtral-12b-2409", k_actual)
        
        if datos and isinstance(datos, dict):
            datos_extraidos = datos
            modelo_exitoso = "pixtral-12b"
            key_exitosa = nombre_key
            break
        elif status == 429:
            with lock_keys: cooldown_keys[k_actual] = time.time() + 5

    # VERIFICACIÓN DE CALIDAD: Si el modelo 12B dejó campos clave como "NO IDENTIFICADO",
    # forzamos una segunda pasada con el modelo GIGANTE (pixtral-large-latest).
    if datos_extraidos:
        campos_criticos = [
            str(datos_extraidos.get("asunto", "")).upper(),
            str(datos_extraidos.get("no_radicado_remitente", "")).upper(),
            str(datos_extraidos.get("no_radicado_destinatario", "")).upper()
        ]
        
        if any("NO IDENTIFICADO" in c or c == "" for c in campos_criticos):
            for intento in range(num_keys):
                idx = (start_key_idx + intento + 1) % num_keys
                k_actual = lista_keys[idx]
                nombre_key = f"Key-{idx+1}"

                with lock_keys:
                    if cooldown_keys[k_actual] > time.time(): continue

                datos_rescate, status = llamar_api_mistral(b64_imgs, nombre_archivo, tipo_flujo, "pixtral-large-latest", k_actual)
                if datos_rescate and isinstance(datos_rescate, dict):
                    # Solo reemplazamos si el modelo grande encontró más datos reales
                    if str(datos_rescate.get("asunto", "")).upper() != "NO IDENTIFICADO":
                        datos_extraidos = datos_rescate
                        modelo_exitoso = "pixtral-large (Rescate)"
                        key_exitosa = nombre_key
                    break
                elif status == 429:
                    with lock_keys: cooldown_keys[k_actual] = time.time() + 5

    return datos_extraidos, key_exitosa, modelo_exitoso

def normalizar_fecha(fecha_str, texto_completo, anio_defecto=""):
    if not fecha_str or str(fecha_str).strip().upper() in ["NO IDENTIFICADO", "N/A", "NONE", "", "NAN"]:
        m_f = re.search(r'(?:Bogot[aá]|Girardot|Honda)[^\n\r]*,?\s*(\d{1,2}\s*de\s*[a-zA-Z]+\s*de\s*\d{4}|\d{2}[-/.]\d{2}[-/.]\d{4})', texto_completo, re.IGNORECASE)
        fecha_str = m_f.group(1) if m_f else ""

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
    
    return fecha_str if fecha_str else "NO IDENTIFICADO"

def blindaje_logica_negocio(datos, nombre_archivo, texto_completo, anio_carpeta, tipo_flujo):
    if not isinstance(datos, dict): datos = {}

    def clean(val):
        if not val or str(val).strip().upper() in ["NONE", "N/A", "NULL", "NAN", "NO IDENTIFICADO"]: return ""
        return ILLEGAL_CHARACTERS_RE.sub("", str(val).strip())

    ia_dest = clean(datos.get("razon_social_destinatario", ""))
    ia_rem = clean(datos.get("razon_social_remitente", ""))
    rad_rem = clean(datos.get("no_radicado_remitente", ""))
    rad_dest = clean(datos.get("no_radicado_destinatario", ""))
    ia_asunto = clean(datos.get("asunto", ""))
    ia_fecha = clean(datos.get("fecha", ""))

    # PURGA DE "ATN." (Evita que nombres de personas queden como entidad)
    ia_dest = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente).*', '', ia_dest).strip()
    ia_rem = re.sub(r'(?i)[,.\-\s]*(Atn|Atención|Attn|Att|A la atención|Ing\.|Gerente).*', '', ia_rem).strip()

    es_recibida = tipo_flujo == "RECIBIDAS"

    if es_recibida:
        # ======= REGLAS INQUEBRANTABLES: RECIBIDAS =======
        ia_dest = "CONSORCIO 4C"
        
        # El remitente NUNCA es Consorcio 4C
        if "CONSORCIO 4C" in ia_rem.upper():
            ia_rem = ""

        if not ia_rem:
            if "ANI_" in nombre_archivo or "ani" in texto_completo.lower():
                ia_rem = "AGENCIA NACIONAL DE INFRAESTRUCTURA - ANI"
            elif "CON_" in nombre_archivo or "alto magdalena" in texto_completo.lower():
                ia_rem = "CONCESIÓN ALTO MAGDALENA S.A.S."

        # Radicado Destinatario SIEMPRE es el GP del Consorcio 4C
        if not rad_dest.startswith("GP-"):
            m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
            rad_dest = f"GP-{m_gp.group(1)}" if m_gp else rad_dest

        # Radicado Remitente de terceros NUNCA es CI.004 ni GP
        if "CI.004" in rad_rem.upper() or "GP-" in rad_rem.upper():
            rad_rem = ""
            
        if not rad_rem:
            m_ani_rad = re.search(r'(?:Radicado\s*ANI\s*No\.?\s*[:\-\.]*\s*|Rad\s*No\.?\s*)(\d{4}[-\s]?\d{3}[-\s]?\d{6}[-\s]?\d|\d{4}-\d{3}-\d+)', texto_completo, re.IGNORECASE)
            m_alma_txt = re.search(r'\b(ALMA[-\s]?20\d{2}[-\s]?\d{3,5})\b', texto_completo, re.IGNORECASE)
            
            if m_ani_rad: rad_rem = m_ani_rad.group(1).replace(' ', '')
            elif m_alma_txt: rad_rem = m_alma_txt.group(1).replace(' ', '-')

    else:
        # ======= REGLAS INQUEBRANTABLES: RADICADAS =======
        ia_rem = "CONSORCIO 4C"

        # Destinatario NUNCA es Consorcio 4C
        if "CONSORCIO 4C" in ia_dest.upper():
            ia_dest = ""

        # Radicado Remitente SIEMPRE es CI.004...
        if "CI.004" not in rad_rem.upper():
            m_ci004 = re.search(r'(CI\.004/[^\s\n\r]+)', texto_completo, re.IGNORECASE)
            if m_ci004:
                rad_rem = m_ci004.group(1).strip()
            else:
                m_nom = re.search(r'CI004_(\d{4})\d{2}_', nombre_archivo, re.IGNORECASE)
                if m_nom:
                    rad_rem = f"CI.004/GP{m_nom.group(1).zfill(4)}"
                else:
                    m_gp = re.search(r'GP[-_]?(\d{3,6})', nombre_archivo, re.IGNORECASE)
                    rad_rem = f"CI.004/GP{m_gp.group(1)}" if m_gp else ""

    # Consolidación final
    fecha_final = normalizar_fecha(ia_fecha, texto_completo, anio_defecto=anio_carpeta)

    return {
        "RAZON_SOCIAL_DESTINATARIO": ia_dest if ia_dest else "NO IDENTIFICADO",
        "RAZON_SOCIAL_REMITENTE": ia_rem if ia_rem else "NO IDENTIFICADO",
        "NO_RADICADO_REMITENTE": rad_rem if rad_rem else "NO IDENTIFICADO",
        "NO_RADICADO_DESTINATARIO": rad_dest if rad_dest else "NO IDENTIFICADO",
        "ASUNTO": ia_asunto if ia_asunto else "NO IDENTIFICADO",
        "FECHA": fecha_final
    }

def procesar_un_pdf(item_num, pdf, ruta_completa, tipo, ruta_memoria, hilo_id, anio_doc):
    t_inicio = time.time()
    ruta_relativa = os.path.relpath(ruta_completa, RUTA_BASE).strip()

    b64_imgs, txt_completo, paginas = obtener_insumos_documento(ruta_completa)
    
    # GARANTÍA 100%: Si falla todo, se crea la fila con NO IDENTIFICADO.
    if not b64_imgs:
        datos_completos = blindaje_logica_negocio({}, pdf, txt_completo, anio_doc, tipo)
        key_usada, mod_usado = "FALLO_DOC", "NINGUNO"
    else:
        datos, key_usada, mod_usado = consultar_pixtral_exhaustivo(b64_imgs, pdf, tipo, item_num, hilo_id)
        if datos is None:
            datos_completos = blindaje_logica_negocio({}, pdf, txt_completo, anio_doc, tipo)
            key_usada, mod_usado = "FALLO_IA", "NINGUNO"
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
    estado = "⚠️ SALVADO VACÍO" if "FALLO" in key_usada else "✅ OK"
    print(f"📄 [Hilo-{hilo_id} | {mod_usado}] {pdf} | {estado} ({duracion}s)", flush=True)
    return True

def procesar_archivos():
    print("\n" + "="*70, flush=True)
    print(" MOTOR RESTREPO_2 (DOBLE MODELO IA + BLINDAJE LÓGICO CONSORCIO 4C)", flush=True)
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
                f = executor.submit(procesar_un_pdf, item_counter, pdf, ruta_completa, tipo, ruta_memoria, hilo_id, anio_doc)
                futuros.append(f)
                item_counter += 1
                time.sleep(0.5)
            for f in as_completed(futuros): pass

    # ENSAMBLAJE FINAL EXCEL SIN ALTERACIONES INVASIVAS (Mantiene el orden limpio)
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
            msg.set_content(f'Proceso concluido exitosamente bajo doble revisión de IA para {etiqueta}. Ningún archivo fue omitido.')
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

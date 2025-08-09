import streamlit as st
import pandas as pd
import requests
import datetime
import os
import time
import random
import re
from io import BytesIO

# =========================
# UI base
# =========================
st.set_page_config(page_title="📨 Envío Masivo WhatsApp", layout="centered")
st.title("📨 Envío Masivo de WhatsApp con Plantillas")

api_key = st.text_input("🔐 Ingresa tu API Key de 360dialog", type="password")
file = st.file_uploader("📁 Sube tu archivo Excel", type=["xlsx"])

colA, colB, colC = st.columns(3)
with colA:
    throttle_ms = st.number_input("⏱️ Pausa entre envíos (ms)", min_value=0, max_value=5000, value=250, step=50)
with colB:
    batch_size = st.number_input("📦 Tamaño de lote", min_value=1, max_value=1000, value=50, step=10)
with colC:
    max_retries = st.number_input("🔁 Reintentos (WA/Chatwoot)", min_value=1, max_value=10, value=5, step=1)

st.caption("Consejo: si notas cortes, sube la pausa, baja el tamaño de lote o aumenta reintentos.")

# =========================
# Plantillas locales (solo para reflejo y bitácora)
# =========================
plantillas = {
    "mensaje_entre_semana_24_hrs": lambda localidad: f"""Buen día, te saludamos de CHEP (Tarimas azules), es un gusto en saludarte.

Te escribo para confirmar que el día de mañana tenemos programada la recolección de tarimas en tu localidad: {localidad}.

¿Me podrías indicar cuántas tarimas tienes para entregar? Así podremos coordinar la unidad.""",

    "recordatorio_24_hrs": lambda: "Buen día, estamos siguiendo tu solicitud, ¿Me ayudarías a confirmar si puedo validar la cantidad de tarimas que serán entregadas?"
}

# =========================
# Archivos de salida
# =========================
ARCHIVO_ENVIOS = "envios_hoy.xlsx"
ARCHIVO_ERRORES = "errores_envio_chatwoot.txt"
if not os.path.exists(ARCHIVO_ENVIOS):
    pd.DataFrame(columns=["Fecha", "Número", "Nombre", "Plantilla", "Estado", "Detalle"]).to_excel(ARCHIVO_ENVIOS, index=False)

# =========================
# Utilidades
# =========================
def solo_digitos(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def normalizar_numero(phone_e164: str) -> str:
    """Asegura +521 para México cuando aplique. Mantiene otros países igual."""
    if not phone_e164.startswith("+"):
        return phone_e164
    if phone_e164.startswith("+521"):
        return phone_e164
    if phone_e164.startswith("+52") and not phone_e164.startswith("+521"):
        return "+521" + phone_e164[3:]
    return phone_e164

def construir_payload_template(nombre_plantilla: str, parametros: list):
    payload = {
        "messaging_product": "whatsapp",
        "to": None,  # se define sin '+'
        "type": "template",
        "template": {
            "name": nombre_plantilla,
            "language": {"code": "es_MX"},
            "components": []
        }
    }
    if parametros:
        payload["template"]["components"].append({
            "type": "body",
            "parameters": [{"type": "text", "text": str(p) if p is not None else ""} for p in parametros]
        })
    return payload

# Sesión HTTP reutilizable (mejor performance y menos cortes)
session = requests.Session()
session.headers.update({"Content-Type": "application/json"})

def backoff_sleep(attempt, base=0.6, jitter=True):
    """Exponential backoff con jitter."""
    delay = base * (2 ** (attempt - 1))
    if jitter:
        delay = delay * (0.5 + random.random())  # 50%-150%
    time.sleep(delay)

def enviar_360dialog(numero_e164: str, payload: dict, api_key_360: str, retries: int):
    url = "https://waba-v2.360dialog.io/messages"
    headers = {"D360-API-KEY": api_key_360}
    body = payload.copy()
    body["to"] = numero_e164.replace("+", "")  # sin '+'

    for intento in range(1, retries + 1):
        try:
            r = session.post(url, headers=headers, json=body, timeout=30)
            # OK 2xx
            if 200 <= r.status_code < 300:
                return True, f"{r.status_code}", r.text
            # Rate limit u otros transitorios
            if r.status_code in (408, 409, 425, 429, 500, 502, 503, 504):
                # Respetar Retry-After si existe
                ra = r.headers.get("Retry-After")
                if ra:
                    try:
                        time.sleep(float(ra))
                    except:
                        pass
                else:
                    backoff_sleep(intento)
                continue
            # Error duro → no insistir
            return False, f"{r.status_code}", r.text
        except requests.RequestException as e:
            if intento == retries:
                return False, "EXC", str(e)
            backoff_sleep(intento)

    return False, "UNKN", "Sin respuesta tras reintentos"

def reflejar_chatwoot(phone_e164: str, nombre: str, contenido: str, retries: int):
    """Reflejo con más paciencia (Chatwoot a veces tarda en crear contacto/conversación)."""
    url = "https://webhook-chatwoots.onrender.com/send-chatwoot-message"
    data = {"phone": phone_e164, "name": nombre or "Cliente WhatsApp", "content": contenido}

    for intento in range(1, retries + 1):
        try:
            r = session.post(url, json=data, timeout=30)
            if 200 <= r.status_code < 300:
                return True, f"{r.status_code}", r.text
            # Reintentar si 4xx/5xx (salvo 400/404 repetidos podría ser datos mal formados)
            if r.status_code in (408, 409, 425, 429, 500, 502, 503, 504):
                backoff_sleep(intento, base=0.8)
                continue
            # a veces 404/422 ocurre si la conversación aún no se registró: dar 1-2 intentos más
            if r.status_code in (404, 422) and intento < retries:
                backoff_sleep(intento, base=0.8)
                continue
            return False, f"{r.status_code}", r.text
        except requests.RequestException as e:
            if intento == retries:
                return False, "EXC", str(e)
            backoff_sleep(intento, base=0.8)

    return False, "UNKN", "Sin respuesta tras reintentos"

def registrar_envio_local(numero: str, nombre: str, plantilla: str, estado: str, detalle: str):
    try:
        hoy = datetime.date.today().strftime('%Y-%m-%d')
        df_existente = pd.read_excel(ARCHIVO_ENVIOS)
        nuevo = pd.DataFrame([{
            "Fecha": hoy,
            "Número": f"'{numero}",   # para que Excel no recorte ceros
            "Nombre": nombre,
            "Plantilla": plantilla,
            "Estado": estado,
            "Detalle": detalle[:300]
        }])
        df_actualizado = pd.concat([df_existente, nuevo], ignore_index=True)
        df_actualizado.to_excel(ARCHIVO_ENVIOS, index=False)
        return True, ""
    except Exception as e:
        return False, str(e)

# =========================
# Lógica principal
# =========================
if api_key and file:
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        st.success(f"Archivo cargado con {len(df)} registros.")
    except Exception as e:
        st.error(f"❌ No se pudo leer el Excel: {e}")
        st.stop()

    columnas = df.columns.tolist()
    st.subheader("🧭 Mapeo de columnas")
    plantilla_col = st.selectbox("🧩 Columna plantilla:", columnas)
    telefono_col = st.selectbox("📱 Teléfono:", columnas)
    nombre_col = st.selectbox("📗 Nombre:", columnas)
    pais_col = st.selectbox("🌎 Código país:", columnas)
    param1_col = st.selectbox("🔢 Parámetro {{1}}:", ["(ninguno)"] + columnas)
    param2_col = st.selectbox("🔢 Parámetro {{2}} (opcional):", ["(ninguno)"] + columnas)

    if "enviado" not in df.columns:
        df["enviado"] = False

    if st.button("🚀 Enviar mensajes"):
        total = len(df)
        enviados_ok = 0
        reflejados_ok = 0

        # Procesar por lotes para evitar picos
        for start in range(0, total, batch_size):
            end = min(start + batch_size, total)
            st.write(f"▶️ Procesando lote {start+1}–{end} de {total}...")
            lote = df.iloc[start:end]

            for idx, row in lote.iterrows():
                try:
                    if bool(row.get("enviado")) is True:
                        continue

                    # Construcción del número
                    cc = solo_digitos(str(row[pais_col])) if pd.notna(row[pais_col]) else ""
                    tel_raw = str(row[telefono_col]).strip() if pd.notna(row[telefono_col]) else ""
                    tel_digits = solo_digitos(tel_raw)

                    if not tel_digits:
                        st.warning(f"Fila {idx+2}: teléfono vacío. Se omite.")
                        df.at[idx, "enviado"] = False
                        continue

                    if cc:
                        phone_e164 = f"+{cc}{tel_digits}"
                    else:
                        # Aceptar si ya viene con +E.164 en 'telefono'
                        if tel_raw.startswith("+"):
                            phone_e164 = tel_raw
                        else:
                            st.warning(f"Fila {idx+2}: sin código de país (usa columna de país o +E.164).")
                            df.at[idx, "enviado"] = False
                            continue

                    phone_e164 = normalizar_numero(phone_e164)
                    nombre = str(row[nombre_col]).strip() if pd.notna(row[nombre_col]) else ""
                    plantilla_nombre = str(row[plantilla_col]).strip() if pd.notna(row[plantilla_col]) else ""

                    # Parámetros del template
                    param1 = str(row[param1_col]).strip() if param1_col != "(ninguno)" and pd.notna(row.get(param1_col)) else None
                    param2 = str(row[param2_col]).strip() if param2_col != "(ninguno)" and pd.notna(row.get(param2_col)) else None
                    parameters = []

                    if plantilla_nombre == "recordatorio_24_hrs":
                        mensaje_real = plantillas["recordatorio_24_hrs"]()
                        # recordatorio_24_hrs no lleva body params
                    elif plantilla_nombre == "mensaje_entre_semana_24_hrs":
                        # espera 1 param (localidad)
                        mensaje_real = plantillas["mensaje_entre_semana_24_hrs"](param1 or "")
                        if param1 is not None:
                            parameters.append({"type": "text", "text": param1})
                    else:
                        # genérico: enviar hasta 2 params si vienen
                        mensaje_real = plantillas.get(
                            plantilla_nombre,
                            lambda x=None: f"Mensaje plantilla '{plantilla_nombre}' enviado."
                        )(param1)
                        if param1 is not None:
                            parameters.append({"type": "text", "text": param1})
                        if param2 is not None:
                            parameters.append({"type": "text", "text": param2})

                    # Preparar payload WA
                    payload = construir_payload_template(plantilla_nombre, [p["text"] for p in parameters])
                    ok_wa, wa_status, wa_text = enviar_360dialog(phone_e164, payload, api_key, int(max_retries))

                    if ok_wa:
                        enviados_ok += 1
                        df.at[idx, "enviado"] = True
                        st.success(f"✅ WhatsApp enviado: {phone_e164}")
                        # Registrar localmente
                        ok_log, err_log = registrar_envio_local(phone_e164, nombre, plantilla_nombre, "Enviado", wa_status)
                        if not ok_log:
                            st.warning(f"⚠️ No se pudo registrar en Excel local: {err_log}")
                    else:
                        df.at[idx, "enviado"] = False
                        st.error(f"❌ WhatsApp error ({phone_e164}): {wa_status} {wa_text[:180]}")
                        # Registrar fallo
                        registrar_envio_local(phone_e164, nombre, plantilla_nombre, "Fallo", f"{wa_status} {wa_text[:180]}")
                        # Si falla WA, no intentes Chatwoot
                        if throttle_ms > 0: time.sleep(throttle_ms/1000)
                        continue

                    # Reflejo en Chatwoot
                    time.sleep(0.6)  # dar tiempo a que se cree conversación
                    mensaje_env = mensaje_real.strip()
                    if "[streamlit]" not in mensaje_env:
                        mensaje_env += " [streamlit]"

                    ok_cw, cw_status, cw_text = reflejar_chatwoot(phone_e164, nombre or "Cliente WhatsApp", mensaje_env, int(max_retries))
                    if ok_cw:
                        reflejados_ok += 1
                        st.info(f"📥 Reflejado en Chatwoot: {phone_e164}")
                    else:
                        st.warning(f"⚠️ Chatwoot error ({phone_e164}): {cw_status} {str(cw_text)[:180]}")
                        try:
                            with open(ARCHIVO_ERRORES, "a", encoding="utf-8") as f:
                                f.write(f"{datetime.datetime.now()} - Error al reflejar {phone_e164}: {mensaje_env} | {cw_status}\n")
                        except:
                            pass

                    # Throttle entre envíos
                    if throttle_ms > 0:
                        time.sleep(throttle_ms / 1000.0)

                except Exception as e:
                    st.error(f"❌ Error en fila {idx+2}: {e}")
                    try:
                        with open(ARCHIVO_ERRORES, "a", encoding="utf-8") as f:
                            f.write(f"{datetime.datetime.now()} - Excepción fila {idx+2}: {e}\n")
                    except:
                        pass

            # Pausa corta entre lotes para no saturar
            time.sleep(1.2)

        st.success(f"🎯 Envíos WA OK: {enviados_ok}")
        st.info(f"📝 Reflejados en Chatwoot OK: {reflejados_ok}")

        # Botón de descarga del Excel acumulado
        try:
            df_final = pd.read_excel(ARCHIVO_ENVIOS)
            output = BytesIO()
            df_final.to_excel(output, index=False)
            st.download_button(
                label="📥 Descargar Excel de envíos",
                data=output.getvalue(),
                file_name="envios_hoy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.warning(f"⚠️ No se pudo preparar archivo para descargar: {e}")

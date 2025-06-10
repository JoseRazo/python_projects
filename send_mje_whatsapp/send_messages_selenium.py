import os
import pyperclip
import pyautogui
import pandas as pd
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess

def copy_file_to_clipboard_linux(filepath):
    """
    Copia el archivo al portapapeles en formato URI para que WhatsApp Web lo reconozca al pegar.
    Requiere tener instalado 'xclip'.
    """
    filepath_abs = os.path.abspath(filepath)
    uri = f"file://{filepath_abs}"
    subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'text/uri-list'], input=uri.encode())

# Elige el navegador: "firefox", "chrome", o "edge"
navegador = "firefox"  # Cambia a "chrome" o "edge" si lo deseas

if navegador == "firefox":
    options = webdriver.FirefoxOptions()
    options.set_preference("permissions.default.image", 2)
    options.set_preference("dom.webnotifications.enabled", False)
    options.set_preference("media.volume_scale", "0.0")
    options.set_preference("network.http.use-cache", False)
    driver = webdriver.Firefox(options=options)

elif navegador == "chrome":
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-images")
    driver = webdriver.Chrome(options=options)

elif navegador == "edge":
    options = webdriver.EdgeOptions()
    driver = webdriver.Edge(options=options)

else:
    raise Exception("Navegador no soportado. Usa 'firefox', 'chrome' o 'edge'.")

driver.get("https://web.whatsapp.com")
input("📲 Escanea el código QR y presiona ENTER cuando estés listo... (solo si es la primera vez)")

df = pd.read_excel("clientes.xlsx")

# Mensajes para pago pendiente
mensajes_base_pago_pendiente = [
    (
        "Aviso Netmiio Telecomunicaciones\n\n"
        "Estimado suscriptor, tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Te invitamos a realizar tu pago. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Netmiio Telecomunicaciones Informa\n\n"
        "Hola, tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Te invitamos a realizar tu pago. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Aviso Importante de Netmiio Telecomunicaciones\n\n"
        "Buen día, ya esta disponible tu recibo del servicio de internet tu fecha limite de pago es el {fecha_limite}. "
        "Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    )
]

# Mensajes para pago vencido
mensajes_base_pago_vencido = [
    (
        "Aviso Netmiio Telecomunicaciones\n\n"
        "Estimado suscriptor, le recordamos que la fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Por favor, realiza tu pago para evitar la suspensión del servicio. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Netmiio Telecomunicaciones Informa\n\n"
        "Hola, tu cuenta presenta un atraso. Tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Le solicitamos realizar el pago lo antes posible para mantener el servicio activo. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Aviso Importante de Netmiio Telecomunicaciones\n\n"
        "Buen día, según nuestros registros tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Para evitar interrupciones en el servicio, le pedimos ponerse al corriente. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    )
]

# Selección de tipo de mensaje
while True:
    print("Seleccione el tipo de mensaje que desea enviar:")
    print("1. Recordatorio de pago pendiente")
    print("2. Aviso de pago vencido")
    print("3. Mensaje personalizado")
    tipo_mensaje = input("Ingrese 1, 2 o 3: ")

    if tipo_mensaje == "1":
        mensajes_base = mensajes_base_pago_pendiente
        print("Seleccionaste: Recordatorio de pago pendiente")
        break
    elif tipo_mensaje == "2":
        mensajes_base = mensajes_base_pago_vencido
        print("Seleccionaste: Aviso de pago vencido")
        break
    elif tipo_mensaje == "3":
        mensaje_personalizado = input("Escribe el mensaje que deseas enviar: ")
        print("Seleccionaste: Mensaje personalizado")
        break
    else:
        print("Selección no válida. Por favor, ingrese 1, 2 o 3.")

for i, row in df.iterrows():
    if pd.isna(row["Telefono"]):
        print(f"[{i+1}] ❌ Sin número: {row['Nombre']}")
        continue

    raw_numero = str(int(row["Telefono"])).strip()
    if len(raw_numero) != 10:
        print(f"[{i+1}] ❌ Número inválido ({raw_numero})")
        continue

    numero = f"52{raw_numero}"
    nombre = str(row["Nombre"]).strip().title()
    suscriptor = str(row["Suscriptor"]).strip()

    if tipo_mensaje in ["1", "2"]:
        fecha_limite = str(row["Fecha Limite"]).strip()
        mensaje = random.choice(mensajes_base).format(
            nombre=nombre,
            suscriptor=suscriptor,
            fecha_limite=fecha_limite
        )
    else:
        mensaje = mensaje_personalizado

    archivo_pdf = None
    if tipo_mensaje == "1":
        for archivo in os.listdir("recibos"):
            if archivo.startswith(f"{suscriptor}_") and archivo.endswith(".pdf"):
                archivo_pdf = os.path.join("recibos", archivo)
                print(f"[Suscriptor {suscriptor}] Archivo PDF encontrado: {archivo_pdf}")
                ruta_pdf = os.path.abspath(archivo_pdf)
                break

    try:
        driver.get(f"https://web.whatsapp.com/send?phone={numero}")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='10']"))
        )
        time.sleep(2)

        # Pegar mensaje texto
        pyperclip.copy(mensaje)
        campo = driver.find_element(By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
        campo.click()
        campo.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)
        campo.send_keys(Keys.ENTER)
        print(f"[{i+1}] ✅ Mensaje enviado a {raw_numero} - {nombre}")
        time.sleep(3)

        # Pegar archivo PDF si existe
        if archivo_pdf:
            # Hacer clic en el botón de "Adjuntar"
            clip_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@title='Adjuntar']"))
            )
            clip_button.click()
            time.sleep(1)

            # Esperar a que aparezca el input tipo file
            attach_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
            )

            # Enviar la ruta del archivo
            attach_input.send_keys(os.path.abspath(ruta_pdf))

            # Esperar que se cargue la vista previa del archivo (puede ajustar el XPath si es necesario)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@aria-label, 'Documento') or contains(@aria-label, 'archivo')]"))
            )

            # Esperar y hacer clic en el botón de enviar (div con aria-label="Enviar")
            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Enviar' and @role='button']"))
            )
            send_button.click()

            time.sleep(3)  # Esperar a que se envíe el archivo

            print(f"[{i+1}] ✅ Archivo PDF enviado a {raw_numero} - {nombre}")

        time.sleep(3)

    except Exception as e:
        print(f"[{i+1}] ⚠ Error con {raw_numero} - {nombre}: {e}")
        continue

driver.quit()

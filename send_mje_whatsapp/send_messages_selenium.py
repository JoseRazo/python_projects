import pyperclip
import pandas as pd
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

# Abrir WhatsApp Web
driver.get("https://web.whatsapp.com")
input("📲 Escanea el código QR y presiona ENTER cuando estés listo... (solo si es la primera vez)")

# Leer clientes
df = pd.read_excel("clientes.xlsx")

# Mensajes para pago pendiente
mensajes_base_pago_pendiente = [
    (
        "Aviso Netmiio Telecomunicaciones\n\n"
        "Estimado suscriptor {nombre}, tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Por favor, realiza tu pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Netmiio Telecomunicaciones Informa\n\n"
        "Hola {nombre}, tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Le solicitamos realizar el pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Aviso Importante de Netmiio Telecomunicaciones\n\n"
        "Buen día {nombre}, según nuestros registros tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Realiza tu pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    )
]

# Mensajes para pago vencido
mensajes_base_pago_vencido = [
    (
        "Aviso Netmiio Telecomunicaciones\n\n"
        "Estimado suscriptor {nombre}, le recordamos que la fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Por favor, realiza tu pago para evitar la suspensión del servicio. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Netmiio Telecomunicaciones Informa\n\n"
        "Hola {nombre}, tu cuenta presenta un atraso. Tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Le solicitamos realizar el pago lo antes posible para mantener el servicio activo. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    ),
    (
        "Aviso Importante de Netmiio Telecomunicaciones\n\n"
        "Buen día {nombre}, según nuestros registros tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Para evitar interrupciones en el servicio, le pedimos ponerse al corriente. Contrato {suscriptor}, si ya pagaste haz caso omiso. "
        "Este mensaje se genero automáticamente, no es necesario responder."
    )
]

# Solicitar al usuario que seleccione el tipo de mensaje
while True:
    print("Seleccione el tipo de mensaje que desea enviar:")
    print("1. Recordatorio de pago pendiente")
    print("2. Aviso de pago vencido")
    tipo_mensaje = input("Ingrese 1 o 2: ")

    if tipo_mensaje == "1":
        mensajes_base = mensajes_base_pago_pendiente
        print("Seleccionaste: Recordatorio de pago pendiente")
        break  # Salir del bucle si la opción es válida
    elif tipo_mensaje == "2":
        mensajes_base = mensajes_base_pago_vencido
        print("Seleccionaste: Aviso de pago vencido")
        break  # Salir del bucle si la opción es válida
    else:
        print("Selección no válida. Por favor, ingrese 1 o 2.")

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
    fecha_limite = str(row["Fecha Limite"]).strip()
    mensaje = random.choice(mensajes_base).format(nombre=nombre, suscriptor=suscriptor, fecha_limite=fecha_limite)

    try:
        # Solo abrir el chat una vez
        driver.get(f"https://web.whatsapp.com/send?phone={numero}")
        
        # Esperar a que se cargue el campo de texto
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='10']"))
        )
        time.sleep(2)

        # Copiar el mensaje al portapapeles y pegarlo
        pyperclip.copy(mensaje)
        campo = driver.find_element(By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
        campo.click()
        campo.send_keys(Keys.CONTROL, 'v')  # Pega el mensaje completo
        time.sleep(1)
        campo.send_keys(Keys.ENTER)

        print(f"[{i+1}] ✅ Enviado a {raw_numero} - {nombre}")
        time.sleep(5)

    except Exception as e:
        print(f"[{i+1}] ⚠ Error con {raw_numero} - {nombre}: {e}")
        continue

driver.quit()

import pywhatkit as kit
import pandas as pd
import time
import random
import re

# Cargar números desde Excel
df = pd.read_excel("clientes.xlsx")  # Columnas: Nombre, Telefono, Suscriptor, Fecha Limite

# Mensajes para pago pendiente
mensajes_base_pago_pendiente = [
    (
        "Aviso Netmiio Telecomunicaciones\n\n"
        "Estimado suscriptor {nombre}, tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Por favor, realiza tu pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso."
    ),
    (
        "Netmiio Telecomunicaciones Informa\n\n"
        "Hola {nombre}, tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Le solicitamos realizar el pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso."
    ),
    (
        "Aviso Importante de Netmiio Telecomunicaciones\n\n"
        "Buen día {nombre}, según nuestros registros tu fecha límite de pago del servicio de internet es el {fecha_limite}. "
        "Realiza tu pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso."
    )
]

# Mensajes para pago vencido
mensajes_base_pago_vencido = [
    (
        "Aviso Netmiio Telecomunicaciones\n\n"
        "Estimado suscriptor {nombre}, le recordamos que la fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Por favor, realiza tu pago para evitar la suspensión del servicio. Contrato {suscriptor}, si ya pagaste haz caso omiso."
    ),
    (
        "Netmiio Telecomunicaciones Informa\n\n"
        "Hola {nombre}, tu cuenta presenta un atraso. Tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Le solicitamos realizar el pago lo antes posible para mantener el servicio activo. Contrato {suscriptor}, si ya pagaste haz caso omiso."
    ),
    (
        "Aviso Importante de Netmiio Telecomunicaciones\n\n"
        "Buen día {nombre}, según nuestros registros tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. "
        "Para evitar interrupciones en el servicio, le pedimos ponerse al corriente. Contrato {suscriptor}, si ya pagaste haz caso omiso."
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

# Función para validar que el número tenga exactamente 10 dígitos
def es_numero_valido(numero: str) -> bool:
    return bool(re.fullmatch(r"\d{10}", numero))

for i, row in df.iterrows():
    if pd.isna(row['Telefono']):
        print(f"[{i+1}] ❌ El suscriptor {row['Nombre']} no tiene número de teléfono.")
        continue
    
    raw_numero = str(int(row['Telefono'])).strip()
    
    if not es_numero_valido(raw_numero):
        print(f"[{i+1}] ❌ Número inválido ({raw_numero}) para el suscriptor {row['Nombre']} — debe tener exactamente 10 dígitos.")
        continue

    numero_formateado = f"+521{raw_numero}"
    nombre = str(row['Nombre']).strip().title()
    suscriptor = str(row['Suscriptor']).strip()
    fecha_limite = str(row['Fecha Limite']).strip()
    mensaje = random.choice(mensajes_base).format(nombre=nombre, suscriptor=suscriptor, fecha_limite=fecha_limite)

    print(f"[{i+1}] ✅ Enviando mensaje a {numero_formateado} - {nombre}")
    try:
        kit.sendwhatmsg_instantly(numero_formateado, mensaje, wait_time=10, tab_close=True)
        time.sleep(30)  # Espera entre mensajes
    except Exception as e:
        print(f"[{i+1}] ⚠️ Error al enviar a {numero_formateado}: {e}")
        time.sleep(10)

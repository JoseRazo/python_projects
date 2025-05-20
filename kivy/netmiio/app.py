# archivo: app.py

import kivy
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput

import pandas as pd
import pywhatkit as kit
import time
import random
import re

class WhatsAppSender(BoxLayout):

    def __init__(self, **kwargs):
        super().__init__(orientation='vertical', padding=10, spacing=10, **kwargs)

        self.filechooser = FileChooserIconView(filters=['*.xlsx'])
        self.add_widget(self.filechooser)

        self.spinner = Spinner(
            text='Selecciona tipo de mensaje',
            values=('Recordatorio de pago pendiente', 'Aviso de pago vencido'),
            size_hint=(1, None),
            height=44
        )
        self.add_widget(self.spinner)

        self.output = TextInput(readonly=True, size_hint=(1, 0.4))
        self.add_widget(self.output)

        self.send_button = Button(text='Enviar mensajes', size_hint=(1, None), height=50)
        self.send_button.bind(on_press=self.enviar_mensajes)
        self.add_widget(self.send_button)

    def log(self, mensaje):
        self.output.text += mensaje + '\n'

    def es_numero_valido(self, numero: str) -> bool:
        return bool(re.fullmatch(r"\d{10}", numero))

    def enviar_mensajes(self, instance):
        filepath = self.filechooser.selection[0] if self.filechooser.selection else None
        if not filepath:
            self.log("❌ Selecciona un archivo Excel primero.")
            return

        try:
            df = pd.read_excel(filepath)
        except Exception as e:
            self.log(f"❌ Error al leer el archivo: {e}")
            return

        tipo = self.spinner.text
        if tipo == 'Recordatorio de pago pendiente':
            mensajes_base = [
                "Aviso Netmiio Telecomunicaciones\n\nEstimado suscriptor {nombre}, tu fecha límite de pago del servicio de internet es el {fecha_limite}. Por favor, realiza tu pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso.",
                "Netmiio Telecomunicaciones Informa\n\nHola {nombre}, tu fecha límite de pago del servicio de internet es el {fecha_limite}. Le solicitamos realizar el pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso.",
                "Aviso Importante de Netmiio Telecomunicaciones\n\nBuen día {nombre}, según nuestros registros tu fecha límite de pago del servicio de internet es el {fecha_limite}. Realiza tu pago lo antes posible. Contrato {suscriptor}, si ya pagaste haz caso omiso."
            ]
        elif tipo == 'Aviso de pago vencido':
            mensajes_base = [
                "Aviso Netmiio Telecomunicaciones\n\nEstimado suscriptor {nombre}, le recordamos que la fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. Por favor, realiza tu pago para evitar la suspensión del servicio. Contrato {suscriptor}, si ya pagaste haz caso omiso.",
                "Netmiio Telecomunicaciones Informa\n\nHola {nombre}, tu cuenta presenta un atraso. Tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. Le solicitamos realizar el pago lo antes posible para mantener el servicio activo. Contrato {suscriptor}, si ya pagaste haz caso omiso.",
                "Aviso Importante de Netmiio Telecomunicaciones\n\nBuen día {nombre}, según nuestros registros tu fecha límite de pago del servicio de internet ya vencio el {fecha_limite}. Para evitar interrupciones en el servicio, le pedimos ponerse al corriente. Contrato {suscriptor}, si ya pagaste haz caso omiso."
            ]
        else:
            self.log("❌ Selecciona el tipo de mensaje.")
            return

        for i, row in df.iterrows():
            if pd.isna(row['Telefono']):
                self.log(f"[{i+1}] ❌ El suscriptor {row['Nombre']} no tiene número.")
                continue

            raw_numero = str(int(row['Telefono'])).strip()
            if not self.es_numero_valido(raw_numero):
                self.log(f"[{i+1}] ❌ Número inválido ({raw_numero}) para {row['Nombre']}")
                continue

            numero_formateado = f"+521{raw_numero}"
            nombre = str(row['Nombre']).strip().title()
            suscriptor = str(row['Suscriptor']).strip()
            fecha_limite = str(row['Fecha Limite']).strip()
            mensaje = random.choice(mensajes_base).format(
                nombre=nombre,
                suscriptor=suscriptor,
                fecha_limite=fecha_limite
            )

            self.log(f"[{i+1}] ✅ Enviando a {numero_formateado} - {nombre}")
            try:
                kit.sendwhatmsg_instantly(numero_formateado, mensaje, wait_time=10, tab_close=True)
                time.sleep(30)
            except Exception as e:
                self.log(f"[{i+1}] ⚠️ Error al enviar a {numero_formateado}: {e}")
                time.sleep(10)

class WhatsAppApp(App):
    def build(self):
        return WhatsAppSender()

if __name__ == '__main__':
    WhatsAppApp().run()

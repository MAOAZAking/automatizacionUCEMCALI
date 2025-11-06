# Ejecutar esto para instalar dependencias necesarias
# pip install twilio

from twilio.rest import Client
import pandas as pd
import os

# Cargar el archivo Excel
df = pd.read_excel('contactos.xlsx')  # El archivo Excel con las columnas: Nombre, Teléfono, Ruta del PDF

# Configuración de Twilio
account_sid = 'tu_account_sid'  # Tu SID de cuenta Twilio
auth_token = 'tu_auth_token'    # Tu token de autenticación de Twilio
client = Client(account_sid, auth_token)

# Iterar sobre cada fila del Excel
for index, row in df.iterrows():
    nombre = row['Nombre']
    telefonos = row['Teléfonos']  # Teléfonos (pueden estar separados por ;)
    pdf_path = row['Ruta del PDF']  # Ruta del archivo PDF

    # Verificar si el archivo PDF existe
    if os.path.exists(pdf_path):
        # Comprobar que el teléfono no es "N/A"
        if pd.notna(telefonos) and telefonos != "N/A":  # Solo proceder si hay teléfono y no es "N/A"
            # Separar los teléfonos por ';' y quitar espacios adicionales
            lista_telefonos = [t.strip() for t in telefonos.split(';')]

            # Generar el saludo (si el nombre no es "N/A", lo incluimos, si no, dejamos el mensaje sin nombre)
            if nombre != "N/A":
                saludo = f"Hola {nombre}, te envío el PDF solicitado."
            else:
                saludo = "Hola, te envío el PDF solicitado."

            # Enviar el mismo archivo a cada número
            for telefono in lista_telefonos:
                telefono = telefono.replace(" ", "")  # Eliminar espacios si los hay

                # Enviar el mensaje de WhatsApp con el archivo adjunto
                message = client.messages.create(
                    body=saludo,  # Mensaje con o sin el nombre
                    from_='whatsapp:+14155238886',  # Tu número Twilio de WhatsApp
                    to=f'whatsapp:{telefono}',  # Número al que se le envía el mensaje
                    media_url=[f'file:///{pdf_path}']  # Ruta del archivo PDF
                )

                print(f"Mensaje enviado a {telefono} con archivo {pdf_path}")
        else:
            print(f"Sin teléfono válido para {nombre}. No se enviará el mensaje.")
    else:
        print(f"El archivo no se encuentra: {pdf_path}")

print("Proceso completado.")
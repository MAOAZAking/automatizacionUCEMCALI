import pyautogui
import pyperclip
import time
import keyboard
import re

def asegurar_foco_ventana():
    try:
        # Asegurarnos de que la ventana de Python esté en foco
        print("Asegurando que la ventana de Python esté en foco...")
        pyautogui.getWindowsWithTitle("Python")[0].activate()  # Ajusta el nombre de la ventana si es necesario
        time.sleep(1)  # Pequeño retraso para asegurarse que la ventana tiene el foco
    except Exception as e:
        print(f"Error al asegurar el foco de la ventana: {e}")

def hacer_clic_en_ubicacion(x, y):
    try:
        # Hacer clic en la ubicación proporcionada
        print(f"Haciendo clic en la ubicación: X={x}, Y={y}")
        pyautogui.click(x, y)  # Hacer clic en las coordenadas dadas
        time.sleep(2)  # Esperar un poco a que la interfaz responda
    except Exception as e:
        print(f"Error al hacer clic en la ubicación: {e}")

def copiar_texto_del_correo():
    try:
        # Seleccionar todo el texto del correo usando Ctrl + A
        pyautogui.hotkey('ctrl', 'a')  # Seleccionar todo el texto
        time.sleep(1)  # Dar tiempo para la selección

        # Copiar el texto seleccionado al portapapeles con Ctrl + C
        pyautogui.hotkey('ctrl', 'c')  # Copiar
        time.sleep(1)  # Esperar que se copie al portapapeles

        # Obtener el contenido del portapapeles
        texto = pyperclip.paste()
        if texto:
            print("Texto copiado del correo:", texto)  # Para depuración
        return texto
    except Exception as e:
        print(f"Error al copiar el texto del correo: {e}")
        return None

def buscar_id_en_texto(texto):
    try:
        # Buscar la palabra clave "ID" seguida de un número con máximo un espacio después de la palabra "ID"
        # Expresión regular que acepta "ID" seguido de 1 o más dígitos, con un máximo de 1 espacio entre ellos.
        match = re.search(r'idnumero\s?(\d+)', texto)
        if match:
            idnumero = match.group(1)  # Obtener el ID
            print(f"ID encontrado: {idnumero}")
            return idnumero
        else:
            print("No se encontró un ID válido en el texto")
            return None
    except Exception as e:
        print(f"Error al buscar el ID en el texto: {e}")
        return None

def copiar_id_a_portapapeles(idnumero):
    try:
        pyperclip.copy(idnumero)  # Copiar al portapapeles
        print("ID copiado al portapapeles.")
    except Exception as e:
        print(f"Error al copiar al portapapeles: {e}")

def realizar_acciones_teclado(idnumero):
    try:
        # 1. Cambiar a la ventana de bloc de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')

        # 2. Inicia la informacion del correo
        pyautogui.write("***************************************")
        pyautogui.press('enter')

        # 3. Escribir "ID: " y pegar el ID desde el portapapeles
        pyautogui.write("ID: ")
        pyautogui.hotkey(idnumero)  # Pegar el ID copiado
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        
        # 4. Cambiar a la ventana del SIGT con Alt + Tab
        keyboard.press('alt')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        keyboard.release('alt')
        
        # 5. Hacer clic en el panel izquierdo del SIGT sobre *CONSULTAR CONTRATO* que esta en la mitad de la pantalla a lo alto y 200 pixeles desde el borde izquierdo hacia la derecha
        x_panel_izquierdo = pyautogui.size().width // 2 - 200
        y_panel_izquierdo = pyautogui.size().height // 2 - 350
        pyautogui.click(x_panel_izquierdo, y_panel_izquierdo)  # Hacer clic en el panel izquierdo
        
        # 6. Dar clic para desplegar con que va a buscar (para selecionar contrato)
        x_busqueda = pyautogui.size().width // 2
        y_busqueda = pyautogui.size().height // 2 + 300
        pyautogui.click(x_busqueda, y_busqueda)  # Hacer clic para mostrar menú

        # 7. Dar clic para selecionar la opcion deseada en el menu que se desplego
        x_busqueda = pyautogui.size().width // 2
        y_busqueda = pyautogui.size().height // 2 + 400
        pyautogui.click(x_busqueda, y_busqueda)  # Hacer clic en la opcion deseada

        # 8. Dar clic en la barra de busqueda y pegar el id
        x_busqueda = pyautogui.size().width // 2 + 400
        y_busqueda = pyautogui.size().height // 2 + 300
        pyautogui.click(x_busqueda, y_busqueda)  # Deja el cursor sobre la barra de búsqueda
        pyautogui.hotkey('ctrl', 'v')  # Pegar el ID copiado
        
        # 9. Tocar en el boton de buscar
        x_boton_buscar = x_busqueda + 600
        y_boton_buscar = y_busqueda
        pyautogui.click(x_boton_buscar, y_boton_buscar)  # Hacer clic en el botón de buscar
        
        # 10. Esperar un poco para que la búsqueda se complete y seleccionar el nombre
        time.sleep(3)
        x_medio_pantalla = pyautogui.size().width // 2 + 150
        y_medio_pantalla = pyautogui.size().height // 2 + 200

        pyautogui.click(x_medio_pantalla, y_medio_pantalla, button='left', duration=0.5)
        pyautogui.mouseDown()
        pyautogui.moveRel(250, 0, duration=0.5)
        pyautogui.mouseUp()

        time.sleep(0.6)

        pyautogui.hotkey('ctrl', 'c')

        nombre_contacto = pyperclip.paste()


        # 11. Cambiar a la ventana de block de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 12. Escribir el nombre del usuario
        pyautogui.write("Nombre: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        
         # 13. Cambiar a la ventana de SIGT con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 14. Seleccionar el número de contacto del usuario
        x_medio_pantalla = pyautogui.size().width // 2 + 400
        y_medio_pantalla = pyautogui.size().height // 2

        pyautogui.click(x_medio_pantalla, y_medio_pantalla, button='left', duration=0.5)
        pyautogui.mouseDown()
        pyautogui.moveRel(250, 0, duration=0.5)
        pyautogui.mouseUp()

        time.sleep(0.6)

        pyautogui.hotkey('ctrl', 'c')

        numero_contacto = pyperclip.paste()
        
        # Hacer la conversion de número telefonico a número celular
        numero_contacto_sin_espacios = numero_contacto.replace(" ", "")
        cantidad_caracteres = len(numero_contacto_sin_espacios)

        if cantidad_caracteres == 7:
            numero_contacto = "602" + numero_contacto_sin_espacios

        print("Número de contacto:", numero_contacto)
        
        
        pyperclip.copy(numero_contacto) #Copia el nuevo número de contacto al portapapeles


        # 15. Cambiar a la ventana de block de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 16. Escribir el número de contacto copiado
        pyautogui.write("Número de contacto: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        
        # 17. Cambiar a la ventana de SIGT con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 18. Seleccionar el email del usuario
        x_medio_pantalla = pyautogui.size().width // 2 + 150
        y_medio_pantalla = pyautogui.size().height // 2 + 200

        pyautogui.click(x_medio_pantalla, y_medio_pantalla, button='left', duration=0.5)
        pyautogui.mouseDown()
        pyautogui.moveRel(250, 0, duration=0.5)
        pyautogui.mouseUp()

        time.sleep(0.6)

        pyautogui.hotkey('ctrl', 'c')

        email_contacto = pyperclip.paste()

        # 19. Cambiar a la ventana de block de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 20. Escribir el correo electronico del cliente copiado
        pyautogui.write("Email de contacto: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        
        # 21. Cambiar a la ventana de SIGT con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 22. Seleccionar el tipo de usuario
        x_medio_pantalla = pyautogui.size().width // 2 + 150
        y_medio_pantalla = pyautogui.size().height // 2 - 200

        pyautogui.click(x_medio_pantalla, y_medio_pantalla, button='left', duration=0.5)
        pyautogui.mouseDown()
        pyautogui.moveRel(250, 0, duration=0.5)
        pyautogui.mouseUp()

        time.sleep(0.6)

        pyautogui.hotkey('ctrl', 'c')

        tipo_usuario = pyperclip.paste()


        # 23. Cambiar a la ventana de block de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 24. Escribir el tipo de cliente copiado
        pyautogui.write("Tipo de cliente: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        
        # 25. Cambiar a la ventana de SIGT con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 26. Seleccionar el plan del usuario
        x_medio_pantalla = pyautogui.size().width // 2 + 400
        y_medio_pantalla = pyautogui.size().height // 2 + 100

        pyautogui.click(x_medio_pantalla, y_medio_pantalla, button='left', duration=0.5)
        pyautogui.mouseDown()
        pyautogui.moveRel(250, 50, duration=0.5)
        pyautogui.mouseUp()

        time.sleep(0.6)

        pyautogui.hotkey('ctrl', 'c')

        plan_usuario = pyperclip.paste()


        # 27. Cambiar a la ventana de block de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 28. Escribir el plan copiado
        pyautogui.write("Plan del cliente: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 29. Cambiar a la ventana del Outlook con Alt + Tab
        keyboard.press('alt')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        keyboard.release('alt')

        # 30. Realizar la acción de Ctrl + L para enfocar la barra de direcciones
        pyautogui.hotkey('ctrl', 'l')  # Enfocar la barra de direcciones

        # 31. Copiar con Ctrl + C
        pyautogui.hotkey('ctrl', 'c')  # Copiar

        # 32. Cambiar a la ventada de bloc denotas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')

        # 33. Pegar la URL del correoabierto para que el usuario pueda acceder con facilidad
        pyautogui.press('enter')
        pyautogui.write("URL del correo:")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        print("Acciones de teclado realizadas.")
    except Exception as e:
        print(f"Error al realizar las acciones de teclado: {e}")

def main():
    try:
        # Asegurarse de que la ventana esté en foco antes de comenzar
        asegurar_foco_ventana()

        # Retrasar la ejecución del código por 7 segundos para dar tiempo a preparar el entorno
        print("Esperando 7 segundos antes de comenzar...")
        time.sleep(7)

        # 1. Hacer clic en la ubicación para seleccionar el correo más reciente
        x1 = 441
        y1 = 343
        hacer_clic_en_ubicacion(x1, y1)

        # 2. Hacer clic en la ubicación para seleccionar el correo específico
        x2 = 939
        y2 = 473
        hacer_clic_en_ubicacion(x2, y2)

        # 3. Copiar el texto del correo
        texto = copiar_texto_del_correo()

        if texto:
            # 4. Buscar el ID en el texto copiado
            idnumero = buscar_id_en_texto(texto)

            if idnumero:
                # 5. Copiar el ID al portapapeles
                copiar_id_a_portapapeles(idnumero)

                # 6. Realizar las acciones de teclado con el ID
                realizar_acciones_teclado(idnumero)
            else:
                print("No se encontró un ID válido en el correo.")
        else:
            print("No se pudo copiar el texto del correo.")
    except Exception as e:
        print(f"Error en el flujo principal: {e}")

if __name__ == "__main__":
    main()


######################### COSAS NECESARIAS PARA QUE FUNCIONE #########################
# Actualizar pip a la última versión
# python -m pip install --upgrade pip
# NOTA: Si aparece un arror acticvamos el entorno virtual aunque es solo ahora porque esta en github, pero igual este es el cmando: ".venv\Scripts\Activate" luego si vuelves a actualizar

# Instalar las librerías necesarias
# pip install pyautogui pyperclip keyboard

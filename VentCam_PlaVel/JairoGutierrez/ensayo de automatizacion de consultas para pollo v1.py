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
        
        # 5. Hacer clic en el panel izquierdo del SIGT que esta en la mitad de la pantalla a lo alto y 200 pixeles desde el borde izquierdo hacia la derecha
        x_panel_izquierdo = pyautogui.size().width // 2 - 200
        y_panel_izquierdo = pyautogui.size().height // 2
        pyautogui.click(x_panel_izquierdo, y_panel_izquierdo)  # Hacer clic en el panel izquierdo
        
        # 6. Dar clic en la barra de busqueda que esta en la mitad de la pantalla de ancho, pero unos 200 pixeles hacia abajo, desde la parte superior de la pantalla
        x_busqueda = pyautogui.size().width // 2
        y_busqueda = pyautogui.size().height // 2 + 200
        pyautogui.click(x_busqueda, y_busqueda)  # Hacer clic en la barra de búsqueda
        
        # 7. Tocar en el boton de buscar que esta en la misma ubicación que la barra de búsqueda solo que 300 pixeles hacia la derecha
        x_boton_buscar = x_busqueda + 300
        y_boton_buscar = y_busqueda
        pyautogui.click(x_boton_buscar, y_boton_buscar)  # Hacer clic en el botón de buscar
        
        # 8. Esperar un poco para que la búsqueda se complete y hacer un clic mantener el clic desde la mitad de toda la pantalla y sin solar el clic arrastrarlo 150 pixeles hacia la derecha, para que selecione el texto que esta desde la mitad de la pantalla hasta 150 pixeles mas hacia la derecha, una vez se recorran los 150 pixeles, manteniendo el clic para seleccionar, se suelta el clic y se copia el texto que se ha seleccionado
        time.sleep(3)
        x_medio_pantalla = pyautogui.size().width // 2
        y_medio_pantalla = pyautogui.size().height // 2

        pyautogui.click(x_medio_pantalla, y_medio_pantalla, button='left', duration=0.5)
        pyautogui.mouseDown()
        pyautogui.moveRel(150, 0, duration=0.5)
        pyautogui.mouseUp()

        time.sleep(1)

        pyautogui.hotkey('ctrl', 'c')

        nombre = pyperclip.paste()


        # 9. Cambiar a la ventana de block de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')
        
        # 10. Escribir el nombre del correo copiado
        pyautogui.write("Nombre: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 11. Cambiar a la ventana del Outlook con Alt + Tab
        keyboard.press('alt')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        keyboard.release('alt')

        # 12. Realizar la acción de Ctrl + L para enfocar la barra de direcciones
        pyautogui.hotkey('ctrl', 'l')  # Enfocar la barra de direcciones

        # 13. Copiar con Ctrl + C
        pyautogui.hotkey('ctrl', 'c')  # Copiar

        # 14. Cambiar a la ventada de bloc denotas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')

        # 8. Pegar la URL del correoabierto para que el usuario pueda acceder con facilidad
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

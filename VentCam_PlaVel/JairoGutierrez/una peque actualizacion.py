import pyautogui
import pyperclip
import time
import keyboard
import re
import ctypes  # Para la ventana emergente de advertencia

idnumero = 0  # Inicializar variable global

def asegurar_foco_ventana():
    try:
        print("Asegurando que la ventana de Python esté en foco...")
        pyautogui.getWindowsWithTitle("Python")[0].activate()
        time.sleep(1)
    except Exception as e:
        print(f"Error al asegurar el foco de la ventana: {e}")

def copiar_texto_del_correo():
    try:
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(1)
        texto = pyperclip.paste()
        if texto:
            print("Texto copiado del correo:", texto)
        return texto
    except Exception as e:
        print(f"Error al copiar el texto del correo: {e}")
        return None

def buscar_id_en_texto(texto):
    try:
        match = re.search(r'idnumero\s?(\d+)', texto)
        if match:
            id_encontrado = match.group(1)
            print(f"ID encontrado: {id_encontrado}")
            return id_encontrado
        else:
            print("No se encontró un ID válido en el texto")
            return None
    except Exception as e:
        print(f"Error al buscar el ID en el texto: {e}")
        return None

def copiar_id_a_portapapeles(idtexto):
    try:
        pyperclip.copy(idtexto)
        print("ID copiado al portapapeles.")
    except Exception as e:
        print(f"Error al copiar al portapapeles: {e}")

def mostrar_alerta_y_terminar():
    ctypes.windll.user32.MessageBoxW(0, "ATENCIÓN: El correo actual no tiene ID\nSe requiere atención inmediata de un usuario", "¡Advertencia!", 0)
    exit()

def copiar_y_pegar_dato(etiqueta, offset_x, offset_y, mover_y=0, convertir=False):
    centro_x = pyautogui.size().width // 2 + offset_x
    centro_y = pyautogui.size().height // 2 + offset_y

    pyautogui.click(centro_x, centro_y)
    pyautogui.mouseDown()
    pyautogui.moveRel(250, mover_y, duration=0.5)
    pyautogui.mouseUp()
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'c')
    valor = pyperclip.paste()

    if convertir:
        valor = valor.replace(" ", "")
        if len(valor) == 7:
            valor = "602" + valor
        pyperclip.copy(valor)

    pyautogui.hotkey('alt', 'tab')
    pyautogui.write(f"{etiqueta}: ")
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'tab')
    return valor

def realizar_acciones_teclado(idtexto):
    try:
        # 1. Cambiar a la ventana de bloc de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')

        # 2. Inicia la informacion del correo
        pyautogui.write("***************************************")
        pyautogui.press('enter')

        # 3. Escribir "ID: " y pegar el ID desde el portapapeles
        pyautogui.write("ID: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 4. Cambiar a la ventana del SIGT con Alt + Tab (2 veces)
        keyboard.press('alt'); time.sleep(0.1)
        pyautogui.press('tab'); time.sleep(0.1)
        pyautogui.press('tab'); time.sleep(0.1)
        keyboard.release('alt')

        # 5. Hacer clic en el panel izquierdo sobre CONSULTAR CONTRATO
        x, y = pyautogui.locateCenterOnScreen('consultar_contrato.png')  # consultar_contrato.png
        pyautogui.click(x, y)
        time.sleep(2)

        # 6. Dar clic para desplegar con qué va a buscar (para seleccionar contrato)
        x, y = pyautogui.locateCenterOnScreen('desplegar_busqueda.png')  # desplegar_busqueda.png
        pyautogui.click(x, y)
        time.sleep(2)

        # 7. Dar clic para seleccionar la opción deseada en el menú
        x, y = pyautogui.locateCenterOnScreen('opcion_contrato.png')  # opcion_contrato.png
        pyautogui.click(x, y)
        time.sleep(2)

        # 8. Dar clic en la barra de búsqueda y pegar el ID
        x, y = pyautogui.locateCenterOnScreen('barra_busqueda.png')  # barra_busqueda.png
        pyautogui.click(x, y)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)

        # 9. Tocar en el botón de buscar
        x, y = pyautogui.locateCenterOnScreen('boton_buscar.png')  # boton_buscar.png
        pyautogui.click(x, y)
        time.sleep(3)

        # 10. Seleccionar el nombre del contacto
        copiar_y_pegar_dato("Nombre", +150, +200)  # nombre_contacto.png

        # 11. Seleccionar el número de contacto del usuario
        copiar_y_pegar_dato("Número de contacto", +400, 0, convertir=True)  # numero_contacto.png

        # 12. Seleccionar el correo electrónico del cliente
        copiar_y_pegar_dato("Email de contacto", +150, +200)  # email_contacto.png

        # 13. Seleccionar el tipo de cliente
        copiar_y_pegar_dato("Tipo de cliente", +150, -200)  # tipo_cliente.png

        # 14. Seleccionar el plan del usuario
        plan = copiar_y_pegar_dato("Plan del cliente", +400, +100, mover_y=+50)  # plan_cliente.png

        # REFRESCAR LA PÁGINA DEL SIGT
        pyautogui.press('f5')  # refrescar SIGT
        time.sleep(3)

        # Cambiar al Bloc de Notas y pegar plan del cliente
        pyautogui.hotkey('alt', 'tab')
        pyautogui.write("Plan del cliente: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'tab')

        # 15. Cambiar a la ventana de Outlook con Alt + Tab (2 veces)
        keyboard.press('alt'); time.sleep(0.1)
        pyautogui.press('tab'); time.sleep(0.1)
        pyautogui.press('tab'); time.sleep(0.1)
        keyboard.release('alt')

        # 16. Realizar la acción de Ctrl + L para enfocar la barra de direcciones
        pyautogui.hotkey('ctrl', 'l')

        # 17. Copiar la URL del correo
        pyautogui.hotkey('ctrl', 'c')

        # 18. Cambiar a la ventana de bloc de notas con Alt + Tab
        pyautogui.hotkey('alt', 'tab')

        # 19. Escribir la URL del correo
        pyautogui.press('enter')
        pyautogui.write("URL del correo:")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 20. Dar clic en el botón de mover correo
        x, y = pyautogui.locateCenterOnScreen('boton_mover_correo.png')  # boton_mover_correo.png
        pyautogui.click(x, y)
        time.sleep(1)

        # 21. Dar clic en la carpeta a la que se moverá el correo
        x, y = pyautogui.locateCenterOnScreen('carpeta_destino.png')  # carpeta_destino.png
        pyautogui.click(x, y)

        print("Acciones completadas.")

    except Exception as e:
        print(f"Error al realizar acciones de teclado: {e}")

def main():
    asegurar_foco_ventana()
    print("Esperando 7 segundos para preparar el entorno...")
    time.sleep(7)

    while True:
        # 1. Hacer clic en la ubicación para seleccionar el correo más reciente
        x = pyautogui.size().width // 2 - 479
        y = pyautogui.size().height // 2 - 207
        pyautogui.click(x, y)
        time.sleep(2)

        # 2. Hacer clic en la ubicación para seleccionar el correo específico
        x = pyautogui.size().width // 2 + 139
        y = pyautogui.size().height // 2 - 27
        pyautogui.click(x, y)
        time.sleep(2)

        # 3. Copiar el texto del correo
        texto = copiar_texto_del_correo()

        if texto:
            # 4. Buscar el ID en el texto copiado
            id_encontrado = buscar_id_en_texto(texto)
            if id_encontrado:
                idnumero = id_encontrado
                # 5. Copiar el ID al portapapeles
                copiar_id_a_portapapeles(idnumero)
                # 6. Realizar las acciones de teclado con el ID
                realizar_acciones_teclado(idnumero)
                break
            else:
                idnumero = 0
        else:
            idnumero = 0

        # 7. Si no se encuentra un ID válido, mostrar alerta y terminar
        if idnumero == 0:
            mostrar_alerta_y_terminar()

if __name__ == "__main__":
    main()


######################### COSAS NECESARIAS PARA QUE FUNCIONE #########################
# Actualizar pip a la última versión
# python -m pip install --upgrade pip
# NOTA: Si aparece un arror acticvamos el entorno virtual aunque es solo ahora porque esta en github, pero igual este es el cmando: ".venv\Scripts\Activate" luego si vuelves a actualizar

# Instalar las librerías necesarias
# pip install pyautogui pyperclip keyboard

import pyautogui
import pyperclip
import time
import re
import ctypes  # Para la ventana emergente de advertencia
import cv2
import numpy as np
from PIL import ImageGrab
import sys

# Evito uso de keyboard para evitar dependencias problemáticas
# Usaré sólo pyautogui.hotkey para combinaciones de teclas

# Variables globales eliminadas para evitar confusión; manejo idnumero localmente

def asegurar_foco_ventana():
    try:
        print("Asegurando que la ventana de Python esté en foco...")
        ventana_python = pyautogui.getWindowsWithTitle("Python")
        if ventana_python:
            ventana_python[0].activate()
            time.sleep(1)
        else:
            print("Ventana Python no encontrada.")
    except Exception as e:
        print(f"Error al asegurar el foco de la ventana: {e}")

def copiar_texto_del_correo():
    try:
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)  # Reducido tiempo de espera (mejora 6)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)
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
    sys.exit()

# Mejora 2 y 3: Captura pantalla una vez y búsqueda con reintentos y escalado dinámico para robustez y velocidad
def hacer_clic_en_imagen(nombre_imagen, descripcion="", espera=2, pantalla=None, reintentos=3):
    """
    Busca la imagen en pantalla con escalado variable, con reintentos y espera.
    pantalla: imagen PIL para buscar (opcional para optimizar)
    """
    imagen_objetivo = cv2.imread(nombre_imagen, cv2.IMREAD_UNCHANGED)
    if imagen_objetivo is None:
        raise FileNotFoundError(f"No se pudo cargar la imagen {nombre_imagen}")

    for intento in range(reintentos):
        print(f"Buscando '{descripcion}', intento {intento+1}/{reintentos}...")

        # Captura pantalla sólo si no se pasa como argumento (mejora 2)
        if pantalla is None:
            pantalla = ImageGrab.grab()
        pantalla_cv = cv2.cvtColor(np.array(pantalla), cv2.COLOR_RGB2BGR)

        # Rango de escalas para robustez (mejora 3 y 6)
        escalas = np.linspace(0.9, 1.1, 11)  # 0.9,0.91,...,1.1
        mejor_valor = 0
        mejor_ubicacion = None
        mejor_escala = 1.0

        for escala in escalas:
            ancho = int(imagen_objetivo.shape[1] * escala)
            alto = int(imagen_objetivo.shape[0] * escala)
            if ancho < 10 or alto < 10:
                continue  # Imagen demasiado pequeña, saltar
            imagen_redimensionada = cv2.resize(imagen_objetivo, (ancho, alto), interpolation=cv2.INTER_AREA)

            res = cv2.matchTemplate(pantalla_cv, imagen_redimensionada, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

            if max_val > mejor_valor:
                mejor_valor = max_val
                mejor_ubicacion = max_loc
                mejor_escala = escala

            if mejor_valor >= 0.75:  # Umbral encontrado
                break

        if mejor_valor >= 0.75 and mejor_ubicacion:
            x, y = mejor_ubicacion
            ancho = int(imagen_objetivo.shape[1] * mejor_escala)
            alto = int(imagen_objetivo.shape[0] * mejor_escala)
            centro_x = x + ancho // 2
            centro_y = y + alto // 2
            print(f"{descripcion} encontrado en {centro_x, centro_y} con escala {mejor_escala:.2f} (confianza {mejor_valor:.2f})")
            pyautogui.click(centro_x, centro_y)
            time.sleep(espera)
            return True
        else:
            print(f"No se encontró {descripcion} en intento {intento+1}. Esperando y reintentando...")
            time.sleep(1)

    raise Exception(f"No se pudo encontrar la imagen {nombre_imagen} en pantalla después de {reintentos} intentos.")

# Mejora 7 y 8: Mejorar cálculo de posición evitando siempre relativo al centro absoluto, usa ventana activa si posible
def obtener_posicion_relativa(offset_x, offset_y):
    # Intentar obtener ventana activa (si no, fallback centro pantalla)
    try:
        ventana_activa = pyautogui.getActiveWindow()
        if ventana_activa:
            base_x = ventana_activa.left
            base_y = ventana_activa.top
            ancho_ventana = ventana_activa.width
            alto_ventana = ventana_activa.height
            pos_x = base_x + ancho_ventana // 2 + offset_x
            pos_y = base_y + alto_ventana // 2 + offset_y
            return pos_x, pos_y
    except Exception:
        pass
    # fallback absoluto a centro de pantalla
    pos_x = pyautogui.size().width // 2 + offset_x
    pos_y = pyautogui.size().height // 2 + offset_y
    return pos_x, pos_y

def copiar_y_pegar_dato(etiqueta, offset_x, offset_y, mover_y=0, convertir=False):
    centro_x, centro_y = obtener_posicion_relativa(offset_x, offset_y)

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

# Mejora 5 y 8: Unifico combinaciones de teclas con pyautogui, evito keyboard
def presionar_alt_tab_veces(veces=1):
    for _ in range(veces):
        pyautogui.hotkey('alt', 'tab')
        time.sleep(0.5)

def realizar_acciones_teclado(idtexto):
    try:
        # 1. Cambiar a la ventana de bloc de notas con Alt + Tab
        presionar_alt_tab_veces()

        # 2. Inicia la informacion del correo
        pyautogui.write("***************************************")
        pyautogui.press('enter')

        # 3. Escribir "ID: " y pegar el ID desde el portapapeles
        pyautogui.write("ID: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 4. Cambiar a la ventana del SIGT con Alt + Tab (2 veces)
        presionar_alt_tab_veces(2)

        # 5. Hacer clic en el panel izquierdo sobre CONSULTAR CONTRATO
        hacer_clic_en_imagen("imagenes/consultas.png", "Panel Consultar Contrato")

        # 6. Dar clic para desplegar con qué va a buscar
        hacer_clic_en_imagen("imagenes/formabuscar.png", "Desplegar forma de búsqueda")

        # 7. Seleccionar la opción deseada (contrato)
        hacer_clic_en_imagen("imagenes/seleccionarcontrato.png", "Seleccionar opción Contrato")

        # 8. Clic en la barra de búsqueda y pegar el ID
        hacer_clic_en_imagen("imagenes/valor.png", "Campo de Valor del contrato")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)

        # 9. Tocar el botón de buscar
        hacer_clic_en_imagen("imagenes/buscar.png", "Botón Buscar")
        time.sleep(3)

        # 10. Seleccionar el nombre del contacto
        copiar_y_pegar_dato("Nombre", +150, +200)

        # 11. Seleccionar el número de contacto del usuario
        copiar_y_pegar_dato("Número de contacto", +400, 0, convertir=True)

        # 12. Seleccionar el correo electrónico del cliente
        copiar_y_pegar_dato("Email de contacto", +150, +200)

        # 13. Seleccionar el tipo de cliente
        copiar_y_pegar_dato("Tipo de cliente", +150, -200)

        # 14. Seleccionar el plan del usuario
        copiar_y_pegar_dato("Plan del cliente", +400, +100, mover_y=+50)

        # 15. Cambiar a la ventana de Outlook con Alt + Tab (2 veces)
        presionar_alt_tab_veces(2)

        # 16. Enfocar la barra de direcciones y copiar URL
        pyautogui.hotkey('ctrl', 'l')
        pyautogui.hotkey('ctrl', 'c')

        # 17. Cambiar a bloc de notas y pegar la URL
        presionar_alt_tab_veces()
        pyautogui.press('enter')
        pyautogui.write("URL del correo:")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 18. Mover el correo
        hacer_clic_en_imagen("imagenes/movera.png", "Botón Mover Correo")
        hacer_clic_en_imagen("imagenes/seleccionarcarpeta.png", "Carpeta destino")

        print("Acciones completadas.")

    except Exception as e:
        print(f"Error al realizar acciones de teclado: {e}")

def main():
    asegurar_foco_ventana()
    print("Esperando 7 segundos para preparar el entorno...")
    time.sleep(7)

    intentos_max = 5  # Mejora 9: Limitar número de intentos para evitar bucle infinito
    intentos = 0
    idnumero = 0

    while intentos < intentos_max:
        intentos += 1
        print(f"Intento #{intentos} para procesar correo...")

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
            print("ID no encontrado, mostrando alerta.")
            mostrar_alerta_y_terminar()

    if intentos >= intentos_max:
        print(f"No se pudo procesar el correo después de {intentos_max} intentos.")
        mostrar_alerta_y_terminar()

if __name__ == "__main__":
    main()

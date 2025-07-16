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
    imagen_objetivo = cv2.imread(nombre_imagen, cv2.IMREAD_UNCHANGED)
    if imagen_objetivo is None:
        raise FileNotFoundError(f"No se pudo cargar la imagen {nombre_imagen}")

    for intento in range(reintentos):
        print(f"Buscando '{descripcion}', intento {intento+1}/{reintentos}...")

        if pantalla is None:
            pantalla = ImageGrab.grab()
        pantalla_cv = cv2.cvtColor(np.array(pantalla), cv2.COLOR_RGB2BGR)

        escalas = np.linspace(0.9, 1.1, 11)
        mejor_valor = 0
        mejor_ubicacion = None
        mejor_escala = 1.0

        for escala in escalas:
            ancho = int(imagen_objetivo.shape[1] * escala)
            alto = int(imagen_objetivo.shape[0] * escala)
            if ancho < 10 or alto < 10:
                continue
            imagen_redimensionada = cv2.resize(imagen_objetivo, (ancho, alto), interpolation=cv2.INTER_AREA)

            res = cv2.matchTemplate(pantalla_cv, imagen_redimensionada, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

            if max_val > mejor_valor:
                mejor_valor = max_val
                mejor_ubicacion = max_loc
                mejor_escala = escala

            if mejor_valor >= 0.75:
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
    pos_x = pyautogui.size().width // 2 + offset_x
    pos_y = pyautogui.size().height // 2 + offset_y
    return pos_x, pos_y

# Mejora 5 y 8: Unifico combinaciones de teclas con pyautogui, evito keyboard
def presionar_alt_tab_veces(veces=1):
    for _ in range(veces):
        pyautogui.hotkey('alt', 'tab')
        time.sleep(0.5)

def realizar_acciones_teclado(idtexto):
    try:
        presionar_alt_tab_veces()

        pyautogui.write("***************************************")
        pyautogui.press('enter')

        pyautogui.write("ID: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        presionar_alt_tab_veces(2)

        hacer_clic_en_imagen("imagenes/consultas.png", "Panel Consultar Contrato")
        hacer_clic_en_imagen("imagenes/formabuscar.png", "Desplegar forma de búsqueda")
        hacer_clic_en_imagen("imagenes/seleccionarcontrato.png", "Seleccionar opción Contrato")
        hacer_clic_en_imagen("imagenes/valor.png", "Campo de Valor del contrato")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)

        hacer_clic_en_imagen("imagenes/buscar.png", "Botón Buscar")
        time.sleep(3)

        # Datos desde etiquetas con imagen
        extraer_dato_desde_etiqueta("Nombre", "imagenes/nombre.png")
        extraer_dato_desde_etiqueta("Número de contacto", "imagenes/numero_contacto.png", convertir=True)
        extraer_dato_desde_etiqueta("Email de contacto", "imagenes/email.png")
        extraer_dato_desde_etiqueta("Tipo de cliente", "imagenes/tipo_cliente.png")

        # Plan desde tabla
        extraer_plan_desde_tabla()

        presionar_alt_tab_veces(2)
        pyautogui.hotkey('ctrl', 'l')
        pyautogui.hotkey('ctrl', 'c')

        presionar_alt_tab_veces()
        pyautogui.press('enter')
        pyautogui.write("URL del correo:")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        hacer_clic_en_imagen("imagenes/movera.png", "Botón Mover Correo")
        hacer_clic_en_imagen("imagenes/seleccionarcarpeta.png", "Carpeta destino")

        print("Acciones completadas.")

    except Exception as e:
        print(f"Error al realizar acciones de teclado: {e}")

# NUEVAS FUNCIONES PARA DETECCIÓN BASADA EN ETIQUETAS

def extraer_dato_desde_etiqueta(etiqueta, imagen_etiqueta, desplazamiento_x=250, desplazamiento_y=0, convertir=False):
    try:
        pantalla = ImageGrab.grab()
        encontrado = hacer_clic_en_imagen(imagen_etiqueta, f"Etiqueta {etiqueta}", pantalla=pantalla)

        if not encontrado:
            print(f"No se encontró la etiqueta {etiqueta}")
            return None

        pyautogui.moveRel(desplazamiento_x, desplazamiento_y, duration=0.3)
        pyautogui.click()
        pyautogui.mouseDown()
        pyautogui.moveRel(200, 0, duration=0.3)
        pyautogui.mouseUp()
        time.sleep(0.5)

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

    except Exception as e:
        print(f"Error al extraer {etiqueta}: {e}")
        return None

def extraer_plan_desde_tabla():
    try:
        pantalla = ImageGrab.grab()
        encontrado = hacer_clic_en_imagen("imagenes/informacion_contrato.png", "Tabla Información del contrato", pantalla=pantalla)

        if not encontrado:
            print("No se encontró la tabla Información del contrato")
            return None

        pyautogui.moveRel(200, 60, duration=0.3)
        pyautogui.click()
        pyautogui.mouseDown()
        pyautogui.moveRel(200, 0, duration=0.3)
        pyautogui.mouseUp()
        time.sleep(0.5)

        pyautogui.hotkey('ctrl', 'c')
        plan = pyperclip.paste()

        pyautogui.hotkey('alt', 'tab')
        pyautogui.write("Plan del cliente: ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'tab')

        return plan

    except Exception as e:
        print(f"Error al extraer el plan: {e}")
        return None

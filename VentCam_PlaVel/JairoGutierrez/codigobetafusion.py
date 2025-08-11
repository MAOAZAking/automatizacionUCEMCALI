import pyautogui
import pyperclip
import time
import re
import ctypes
import cv2
import numpy as np
from PIL import ImageGrab
import sys
import pygetwindow as gw  # Para manejar ventanas por título

# Configuraciones globales
pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = False

ESCALAS_POSIBLES = np.linspace(0.5, 1.5, num=11)
TOLERANCIA_DETECTADA = 0.75

def asegurar_foco_ventana(titulo_parcial="Python"):
    """
    Busca y activa ventana cuyo título contenga titulo_parcial.
    """
    try:
        log(f"Intentando enfocar ventana con título que contenga '{titulo_parcial}'...")
        ventanas = [v for v in gw.getAllWindows() if titulo_parcial.lower() in v.title.lower()]
        if ventanas:
            ventana = ventanas[0]
            ventana.activate()
            time.sleep(1)
            log(f"Ventana '{ventana.title}' activada.")
            return True
        else:
            log(f"No se encontró ventana con título que contenga '{titulo_parcial}'.")
            return False
    except Exception as e:
        log(f"Error al enfocar ventana: {e}")
        return False

    enfocado = asegurar_foco_ventana("Bloc de notas")
    if not enfocado:
        log("No se pudo enfocar la ventana Python, continuando de todas formas...")

    enfocado = asegurar_foco_ventana("SIGT")
    if not enfocado:
        log("No se pudo enfocar la ventana Python, continuando de todas formas...")

    enfocado = asegurar_foco_ventana("Correo: ")
    if not enfocado:
        log("No se pudo enfocar la ventana Python, continuando de todas formas...")

def log(msg):
    print(f"[LOG] {msg}")


def validar_imagenes(lista_imagenes):
    """
    Valida que todas las imágenes existan y puedan abrirse.
    """
    faltantes = []
    for img in lista_imagenes:
        try:
            if cv2.imread(img) is None:
                faltantes.append(img)
        except Exception:
            faltantes.append(img)
    if faltantes:
        log(f"ERROR: Las siguientes imágenes no se pudieron cargar: {faltantes}")
        return False
    return True

def copiar_texto_del_correo():
    try:
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        texto = pyperclip.paste()
        if texto.strip():
            log(f"Texto copiado del correo: {texto[:50]}...")  # Muestra solo primeros 50 chars
            return texto
        else:
            log("Texto copiado está vacío.")
            return None
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al copiar el contendio del correo para buscar su NÚMERO DE CONTRATO: {e}")

def buscar_id_en_texto(texto):
    try:
        # Convertir a minúsculas para evitar problemas con mayúsculas o tildes
        texto = texto.lower()

        # Lista de patrones flexibles para identificar el número de contrato
        patron = (
            r'(?:'                        # Grupo no capturante para las variantes
                r'n[úu]?m(?:ero)?\.?\s*'   # Variaciones de número: num, núm, número, núm. etc.
                r'(?:de\s*)?'              # Opcionalmente "de"
                r'contrato|'               # seguido de "contrato"
                r'contrato|'               # o solo la palabra "contrato"
                r'n°|'                     # símbolo N°
                r'#|'                      # símbolo #
                r'nro\.?|'                 # abreviatura nro, nro.
                r'nº'                      # símbolo Nº
                r'no'
                r'no.'
                r'no:'
                r'no...'
            r')\s*(\d{4,})'                # Captura un número de al menos 3 dígitos (ajustable)
        )

        match = re.search(patron, texto)
        if match:
            id_encontrado = match.group(1)
            log(f"ID encontrado: {id_encontrado}")
            return id_encontrado
        else:
            log("No se encontró un númeor de contrato válido en el texto")
            return None
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al buscar su NÚMERO DE CONTRATO, puede que este correo no diga exactamente NÚMERO DE CONTRATO: {e}")

def copiar_id_a_portapapeles(idtexto):
    try:
        pyperclip.copy(idtexto)
        log("Número de contrato copiado al portapapeles.")
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al copiar el número de contrato al portapapeles, para buscar la informacion en el SIGT: {e}")

def mostrar_alerta_y_terminar(mensaje="ATENCIÓN: El correo actual no tiene ID\nSe requiere atención inmediata de un usuario"):
    import ctypes, sys
    ctypes.windll.user32.MessageBoxW(0, mensaje, "¡Advertencia!", 0)
    sys.exit()

def hacer_clic_en_imagen(nombre_imagen, descripcion="", tiempo_espera=3, region=None, multiple=False):
    """
    Busca una imagen en pantalla con diferentes escalas y hace clic en ella si la encuentra.
    Soporta diferentes resoluciones, escalados y proporciones.
    """
    print(f"[🧠] Buscando: {descripcion} → Imagen: {nombre_imagen}")

    # Carga la imagen base en escala de grises
    template = cv2.imread(nombre_imagen, cv2.IMREAD_GRAYSCALE)
    if template is None:
        print(f"[❌] No se pudo cargar la imagen {nombre_imagen}")
        return None

    h_template, w_template = template.shape

    tiempo_inicio = time.time()
    escalas = np.linspace(0.6, 1.4, num=20)  # Escalas desde 60% a 140%

    while time.time() - tiempo_inicio < tiempo_espera:
        captura = pyautogui.screenshot(region=region)
        captura = cv2.cvtColor(np.array(captura), cv2.COLOR_RGB2GRAY)

        mejores_puntos = []
        for escala in escalas:
            resized_template = cv2.resize(template, (0, 0), fx=escala, fy=escala)
            h, w = resized_template.shape

            if captura.shape[0] < h or captura.shape[1] < w:
                continue

            resultado = cv2.matchTemplate(captura, resized_template, cv2.TM_CCOEFF_NORMED)
            umbral = 0.7
            ubicaciones = np.where(resultado >= umbral)

            for pt in zip(*ubicaciones[::-1]):
                mejores_puntos.append((pt[0] + w // 2, pt[1] + h // 2, resultado[pt[1], pt[0]]))

        if mejores_puntos:
            mejores_puntos.sort(key=lambda x: x[2], reverse=True)

            if multiple:
                return mejores_puntos  # Devolver todas las coincidencias

            # Solo una coincidencia: hacer clic
            mejor_pt = mejores_puntos[0]
            abs_x, abs_y = mejor_pt[0], mejor_pt[1]
            print(f"[✅] Coincidencia encontrada en ({abs_x}, {abs_y}) con confianza: {mejor_pt[2]:.2f}")
            pyautogui.moveTo(abs_x, abs_y, duration=0.2)
            pyautogui.click()
            return (abs_x, abs_y)


    print(f"[⚠️] No se encontró la imagen: {nombre_imagen}")
    return None


def hacer_clic_en_imagen_ignorando_primera(nombre_imagen, descripcion="", tiempo_espera=3):
    """
    Busca una imagen, ignora la primera coincidencia (más cercana arriba/izquierda),
    y hace clic en la segunda coincidencia si existe.
    """
    print(f"[🧠] Buscando (ignorando 1ra): {descripcion} → Imagen: {nombre_imagen}")
    puntos = hacer_clic_en_imagen(nombre_imagen, descripcion, tiempo_espera, multiple=True)

    if puntos and len(puntos) > 1:
        segundo = puntos[1]
        pyautogui.moveTo(segundo[0], segundo[1], duration=0.2)
        pyautogui.click()
        print(f"[✅] Se hizo clic en la segunda coincidencia.")
        return (segundo[0], segundo[1])
    elif puntos:
        print("[⚠️] Solo se encontró una coincidencia. Se hace clic en esa.")
        pyautogui.moveTo(puntos[0][0], puntos[0][1], duration=0.2)
        pyautogui.click()
        return (puntos[0][0], puntos[0][1])
    else:
        print("[❌] No se encontró ninguna coincidencia.")
        return None

def presionar_alt_tab_veces(veces=1):
    pyautogui.keyDown('alt')  # Mantiene presionado Alt
    for _ in range(veces):
        pyautogui.press('tab')  # Pulsa Tab sin soltar Alt
    pyautogui.keyUp('alt')     # Suelta Alt


def extraer_dato_desde_etiqueta(etiqueta, imagen_etiqueta, desplazamiento_x=75, desplazamiento_y=0, convertir=False):
    try:
        pantalla = ImageGrab.grab()
        encontrado = hacer_clic_en_imagen_ignorando_primera(imagen_etiqueta, f"Etiqueta {etiqueta}")

        if not encontrado:
            log(f"No se encontró la etiqueta {etiqueta}")
            return None

        pyautogui.moveRel(desplazamiento_x, desplazamiento_y, duration=0.3)
        pyautogui.click()
        pyautogui.mouseDown()
        pyautogui.moveRel(150, 0, duration=0.3)
        pyautogui.mouseUp()
        time.sleep(0.5)

        pyautogui.hotkey('ctrl', 'c')
        valor = pyperclip.paste()

        if convertir:
            valor = valor.replace(" ", "")
            if len(valor) == 7:
                valor = "602" + valor
            pyperclip.copy(valor)

        # Pegamos TODO el texto copiado (etiqueta + info) sin escribir la etiqueta manual
        presionar_alt_tab_veces(1)  # Cambiar al Bloc de notas
        pyautogui.write(f"{etiqueta} : ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        presionar_alt_tab_veces(1)

        log(f"{etiqueta} extraído y pegado: {valor.strip()[:40]}...")  # Log parcial

        return valor

    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al copiar el número de contrato al portapapeles, para buscar la informacion en el SIGT: {e}")
        log(f"Error al extraer {etiqueta}: {e}")
        return None

def extraer_plan_desde_tabla():
    try:
        pantalla = ImageGrab.grab()
        encontrado = hacer_clic_en_imagen("imagenes/informacion_contrato.png", "Plan")

        if not encontrado:
            log("No se encontró la tabla Información del contrato")
            return None

        pyautogui.moveRel(100, 0, duration=0.3)
        pyautogui.click()
        pyautogui.mouseDown()
        pyautogui.moveRel(100, 30, duration=0.3)
        pyautogui.mouseUp()
        time.sleep(0.5)

        pyautogui.hotkey('ctrl', 'c')
        plan = pyperclip.paste()

        presionar_alt_tab_veces(1)
        pyautogui.write("Plan : ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        log(f"Plan extraído y pegado: {plan.strip()[:40]}...")

        return plan

    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al extraer el plan del cliente: {e}")

def limpiar_estado_o_cerrar(mensaje_error="Se produjo un error crítico. El proceso ha sido detenido."):
    log(f"Error crítico: {mensaje_error}")
    mostrar_alerta_y_terminar(mensaje_error)


def realizar_acciones_teclado(idtexto):
    try:
        presionar_alt_tab_veces(2)
        pyautogui.press('enter')
        pyautogui.write("******************************************************************************")
        pyautogui.press('enter')
        pyautogui.write("Número de contrato: ")
        pyperclip.copy(idtexto)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        presionar_alt_tab_veces(2)

        hacer_clic_en_imagen("imagenes/consultas.png", "Panel Consultar Contrato")
        hacer_clic_en_imagen("imagenes/forma_buscar.png", "Desplegar forma de búsqueda")
        hacer_clic_en_imagen("imagenes/seleccionar_contrato.png", "Seleccionar opción Contrato")
        hacer_clic_en_imagen("imagenes/valor.png", "Campo de Valor del contrato")
        pyautogui.hotkey('ctrl', 'v')

        hacer_clic_en_imagen("imagenes/buscar.png", "Botón Buscar")
        time.sleep(2)

        #   Seleccionar contrato activo
        hacer_clic_en_imagen("imagenes/seleccionar_contrato_activo.png", "Seleccionar el Contrato que este activo")
        
        extraer_dato_desde_etiqueta("Nombre", "imagenes/nombre.png")
        extraer_dato_desde_etiqueta("Número de contacto", "imagenes/numero_contacto.png", convertir=True)
        extraer_dato_desde_etiqueta("Email de contacto", "imagenes/email.png")
        extraer_dato_desde_etiqueta("Tipo de cliente", "imagenes/tipo_cliente.png")
        extraer_plan_desde_tabla()

        presionar_alt_tab_veces(2)
        pyautogui.hotkey('ctrl', 'l')
        pyautogui.hotkey('ctrl', 'c')

        presionar_alt_tab_veces(1)
        pyautogui.press('enter')
        pyautogui.write("URL del correo:")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        pyautogui.press('enter')
        presionar_alt_tab_veces(1)

        hacer_clic_en_imagen("imagenes/mover_a.png", "Botón Mover Correo")
        hacer_clic_en_imagen("imagenes/mostrar_todas_las_carpetas.png", "Mostrar todas las carpetas")
        hacer_clic_en_imagen("imagenes/seleccionar_carpeta.png", "Carpeta destino")

        log("Acciones completadas.")
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error en realizar_acciones_teclado(): {e}")

def main():
    if not validar_imagenes([
        "imagenes/consultas.png",
        "imagenes/forma_buscar.png",
        "imagenes/seleccionar_contrato.png",
        "imagenes/seleccionar_contrato_activo.png",
        "imagenes/valor.png",
        "imagenes/buscar.png",
        "imagenes/nombre.png",
        "imagenes/email.png",
        "imagenes/numero_contacto.png",
        "imagenes/tipo_cliente.png",
        "imagenes/informacion_contrato.png",
        "imagenes/mover_a.png",
        "imagenes/mostrar_todas_las_carpetas.png",
        "imagenes/seleccionar_carpeta.png"
    ]):
        mostrar_alerta_y_terminar("Faltan imágenes necesarias. Revise la carpeta 'imagenes'.")

    #enfocado = asegurar_foco_ventana("Gestion ADSL")
    #if not enfocado:
        #log("No se pudo enfocar la ventana de Bloc de notas, continuando de todas formas...")

    log("Esperando 7 segundos para preparar el entorno...")
    time.sleep(7)

    intentos_max = 5
    intentos = 0
    numerocontrato = 0

    while intentos < intentos_max:
        intentos += 1
        log(f"Intento #{intentos} para procesar correo...")

        x = pyautogui.size().width // 2 - 419
        y = pyautogui.size().height // 2 - 107
        pyautogui.click(x, y)
        time.sleep(1)

        x = pyautogui.size().width // 2 + 139
        y = pyautogui.size().height // 2 - 27
        pyautogui.click(x, y)

        texto = copiar_texto_del_correo()

        if texto:
            id_encontrado = buscar_id_en_texto(texto)
            if id_encontrado:
                numerocontrato = id_encontrado
                copiar_id_a_portapapeles(numerocontrato)
                realizar_acciones_teclado(numerocontrato)
                break
            else:
                numerocontrato = 0
        else:
            numerocontrato = 0

        if numerocontrato == 0:
            log("NÚMERO DE CONTRATO no encontrado, mostrando alerta.")
            mostrar_alerta_y_terminar()

    if intentos >= intentos_max:
        log(f"No se pudo procesar el correo después de {intentos_max} intentos.")
        mostrar_alerta_y_terminar()

if __name__ == "__main__":
    main()

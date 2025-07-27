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

def log(msg):
    print(f"[LOG] {msg}")

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
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)
        texto = pyperclip.paste()
        if texto.strip():
            log(f"Texto copiado del correo: {texto[:50]}...")  # Muestra solo primeros 50 chars
            return texto
        else:
            log("Texto copiado está vacío.")
            return None
    except Exception as e:
        log(f"Error al copiar el texto del correo: {e}")
        return None

def buscar_id_en_texto(texto):
    try:
        match = re.search(r'idnumero\s?(\d+)', texto, re.IGNORECASE)
        if match:
            id_encontrado = match.group(1)
            log(f"ID encontrado: {id_encontrado}")
            return id_encontrado
        else:
            log("No se encontró un ID válido en el texto")
            return None
    except Exception as e:
        log(f"Error al buscar el ID en el texto: {e}")
        return None

def copiar_id_a_portapapeles(idtexto):
    try:
        pyperclip.copy(idtexto)
        log("ID copiado al portapapeles.")
    except Exception as e:
        log(f"Error al copiar al portapapeles: {e}")

def mostrar_alerta_y_terminar(mensaje="ATENCIÓN: El correo actual no tiene ID\nSe requiere atención inmediata de un usuario"):
    ctypes.windll.user32.MessageBoxW(0, mensaje, "¡Advertencia!", 0)
    sys.exit()

def hacer_clic_en_imagen(nombre_imagen, descripcion="", espera=2, pantalla=None, reintentos=3, porcentaje_inicio_x=0.5):
    imagen_objetivo = cv2.imread(nombre_imagen, cv2.IMREAD_UNCHANGED)
    if imagen_objetivo is None:
        raise FileNotFoundError(f"No se pudo cargar la imagen {nombre_imagen}")

    for intento in range(reintentos):
        log(f"Buscando '{descripcion}', intento {intento+1}/{reintentos}...")

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
            clic_x = int(x + ancho * porcentaje_inicio_x)
            clic_y = y + alto // 2
            log(f"{descripcion} encontrado en {clic_x, clic_y} con escala {mejor_escala:.2f} (confianza {mejor_valor:.2f})")
            pyautogui.click(clic_x, clic_y)
            time.sleep(espera)
            return True
        else:
            log(f"No se encontró {descripcion} en intento {intento+1}. Esperando y reintentando...")
            time.sleep(1)

    raise Exception(f"No se pudo encontrar la imagen {nombre_imagen} en pantalla después de {reintentos} intentos.")

def presionar_alt_tab_veces(veces=1):
    for _ in range(veces):
        pyautogui.hotkey('alt', 'tab')
        time.sleep(0.5)

def extraer_dato_desde_etiqueta(etiqueta, imagen_etiqueta, desplazamiento_x=250, desplazamiento_y=0, convertir=False):
    try:
        pantalla = ImageGrab.grab()
        encontrado = hacer_clic_en_imagen(imagen_etiqueta, f"Etiqueta {etiqueta}", pantalla=pantalla, porcentaje_inicio_x=0.05)

        if not encontrado:
            log(f"No se encontró la etiqueta {etiqueta}")
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

        # Pegamos TODO el texto copiado (etiqueta + info) sin escribir la etiqueta manual
        pyautogui.hotkey('alt', 'tab')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'tab')

        log(f"{etiqueta} extraído y pegado: {valor.strip()[:40]}...")  # Log parcial

        return valor

    except Exception as e:
        log(f"Error al extraer {etiqueta}: {e}")
        return None

def extraer_plan_desde_tabla():
    try:
        pantalla = ImageGrab.grab()
        encontrado = hacer_clic_en_imagen("imagenes/informacion_contrato.png", "Tabla Información del contrato", pantalla=pantalla, porcentaje_inicio_x=0.05)

        if not encontrado:
            log("No se encontró la tabla Información del contrato")
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
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'tab')

        log(f"Plan extraído y pegado: {plan.strip()[:40]}...")

        return plan

    except Exception as e:
        log(f"Error al extraer el plan: {e}")
        return None

def limpiar_estado_o_cerrar():
    """
    Función para cerrar ventana o limpiar estado si error crítico.
    (Implementar según contexto)
    """
    log("Limpiando estado o cerrando ventana por error crítico.")
    # Ejemplo simple: presionar ESC para cerrar popups
    pyautogui.press('esc')
    time.sleep(1)

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

        extraer_dato_desde_etiqueta("Nombre", "imagenes/nombre.png")
        extraer_dato_desde_etiqueta("Número de contacto", "imagenes/numero_contacto.png", convertir=True)
        extraer_dato_desde_etiqueta("Email de contacto", "imagenes/email.png")
        extraer_dato_desde_etiqueta("Tipo de cliente", "imagenes/tipo_cliente.png")
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
        hacer_clic_en_imagen("imagenes/mostrartodaslascarpetas.png", "Mostrar todas las carpetas")
        hacer_clic_en_imagen("imagenes/seleccionarcarpeta.png", "Carpeta destino")

        log("Acciones completadas.")
    except Exception as e:
        log(f"Error al realizar acciones de teclado: {e}")
        limpiar_estado_o_cerrar()

def main():
    if not validar_imagenes([
        "imagenes/consultas.png",
        "imagenes/formabuscar.png",
        "imagenes/seleccionarcontrato.png",
        "imagenes/valor.png",
        "imagenes/buscar.png",
        "imagenes/nombre.png",
        "imagenes/email.png",
        "imagenes/numero_contacto.png",
        "imagenes/tipo_cliente.png",
        "imagenes/informacion_contrato.png",
        "imagenes/movera.png",
        "imagenes/seleccionarcarpeta.png"
    ]):
        mostrar_alerta_y_terminar("Faltan imágenes necesarias. Revise la carpeta 'imagenes'.")

    enfocado = asegurar_foco_ventana("Python")
    if not enfocado:
        log("No se pudo enfocar la ventana Python, continuando de todas formas...")

    log("Esperando 7 segundos para preparar el entorno...")
    time.sleep(7)

    intentos_max = 5
    intentos = 0
    idnumero = 0

    while intentos < intentos_max:
        intentos += 1
        log(f"Intento #{intentos} para procesar correo...")

        x = pyautogui.size().width // 2 - 479
        y = pyautogui.size().height // 2 - 207
        pyautogui.click(x, y)
        time.sleep(2)

        x = pyautogui.size().width // 2 + 139
        y = pyautogui.size().height // 2 - 27
        pyautogui.click(x, y)
        time.sleep(2)

        texto = copiar_texto_del_correo()

        if texto:
            id_encontrado = buscar_id_en_texto(texto)
            if id_encontrado:
                idnumero = id_encontrado
                copiar_id_a_portapapeles(idnumero)
                realizar_acciones_teclado(idnumero)
                break
            else:
                idnumero = 0
        else:
            idnumero = 0

        if idnumero == 0:
            log("ID no encontrado, mostrando alerta.")
            mostrar_alerta_y_terminar()

    if intentos >= intentos_max:
        log(f"No se pudo procesar el correo después de {intentos_max} intentos.")
        mostrar_alerta_y_terminar()

if __name__ == "__main__":
    main()

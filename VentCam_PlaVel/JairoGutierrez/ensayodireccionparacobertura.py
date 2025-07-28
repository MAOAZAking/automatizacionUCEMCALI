import pyautogui
import time
import re
from PIL import ImageChops
import pyperclip
import subprocess

# Función para buscar la imagen en la pantalla en una región específica
def buscar_imagen_region(imagen_nombre, region=None):
    try:
        # Se busca la imagen en la pantalla, devuelve las coordenadas (x, y)
        ubicacion = pyautogui.locateOnScreen(f"imagenes/{imagen_nombre}", region=region)
        if ubicacion:
            centro = pyautogui.center(ubicacion)
            return centro
        else:
            print(f"Imagen {imagen_nombre} no encontrada en la región")
            return None
    except Exception as e:
        print(f"Error buscando imagen {imagen_nombre}: {e}")
        return None

# Función para hacer clic en la imagen encontrada
def hacer_click(imagen_nombre, region=None):
    ubicacion = buscar_imagen_region(imagen_nombre, region)
    if ubicacion:
        pyautogui.click(ubicacion)
        time.sleep(1)

# Función para escribir en un campo después de hacer clic
def escribir_texto(texto):
    pyautogui.write(texto, interval=0.1)
    pyautogui.press('enter')

# Función para comparar dos capturas de pantalla
def comparar_imagenes(imagen1, imagen2):
    try:
        im1 = pyautogui.screenshot(region=imagen1)
        im2 = pyautogui.screenshot(region=imagen2)
        diff = ImageChops.difference(im1, im2)
        return diff.getbbox() is not None  # Si hay diferencia, las imágenes no son iguales
    except Exception as e:
        print(f"Error al comparar imágenes: {e}")
        return False

# Función para procesar la dirección y validar la aceptación
def procesar_direccion(direccion):
    # Regex para capturar el tipo de dirección, número, y nomenclatura
    regex_calle = r"(cl|call|cll|calle|carrera|kr|cr)\.?(\s*\d+)\s?([a-zA-Z]?\s*\d*\w*)"
    regex_nomenclatura = r"(\d+[a-zA-Z]*)\s*(nº|num|numero|#)?\s*([a-zA-Z0-9\s\-]*)"
    
    match_calle = re.search(regex_calle, direccion.lower())
    if match_calle:
        tipo_direccion = match_calle.group(1).capitalize()
        numero_direccion = match_calle.group(2)
        nomenclatura_1 = match_calle.group(3).strip()
    else:
        return "Dirección no válida"
    
    match_nomenclatura = re.search(regex_nomenclatura, direccion)
    if match_nomenclatura:
        numero_con_letra = match_nomenclatura.group(1)
        nomenclatura_2 = match_nomenclatura.group(3)
    else:
        numero_con_letra = numero_direccion
        nomenclatura_2 = ""
    
    direccion_1 = f"{tipo_direccion} {numero_con_letra} {nomenclatura_1} {nomenclatura_2}"
    
    # Cambiar la dirección
    if 'calle' in tipo_direccion.lower():
        tipo_direccion = tipo_direccion.replace('calle', 'carrera')
    elif 'carrera' in tipo_direccion.lower():
        tipo_direccion = tipo_direccion.replace('carrera', 'calle')
    
    direccion_2 = f"{tipo_direccion} {nomenclatura_1} {numero_con_letra} {nomenclatura_2}"

    # Empezamos el proceso de automatización de clics e interacciones con el formulario
    if 'calle' in tipo_direccion.lower():
        # Omitir clic en tipo de dirección si ya es calle
        print("Tipo de dirección ya es Calle, omitiendo clic en selección de tipo de dirección.")
    else:
        hacer_click("selecciontipodireccion.png")  # Clic en el tipo de dirección
        hacer_click(f"seleccionar{tipo_direccion.lower()}.png")  # Clic en el tipo de dirección correcto
    
    # Ingresamos el número de la dirección
    hacer_click("seleccionarnumerodireccion.png")
    escribir_texto(numero_con_letra)

    # Ingresamos la nomenclatura 1
    hacer_click("seleccionarnomenclatura1.png")
    escribir_texto(nomenclatura_1)

    # Ingresamos la nomenclatura 2 (si existe)
    if nomenclatura_2:
        hacer_click("seleccionarnomenclatura2.png")
        escribir_texto(nomenclatura_2)

    # Capturamos la parte inferior de la pantalla (mitad inferior)
    print("Capturando la mitad inferior de la pantalla...")
    pantalla_inferior = pyautogui.screenshot(region=(0, pyautogui.size()[1] // 2, pyautogui.size()[0], pyautogui.size()[1] // 2))

    # 1. Buscamos si aparece el mensaje "No se encontró dirección" o similar
    if buscar_imagen_region("direccionnoencontrada.png", region=(0, pyautogui.size()[1] // 2, pyautogui.size()[0], pyautogui.size()[1] // 2)):
        print("Dirección no encontrada. Intentando con el tipo de dirección 2.")
        cambiar_tipo_direccion(direccion_2)
        return direccion_2

    # 2. Validamos si aparece alguna de las imágenes de aceptación
    aceptada = False
    for i in range(1, 4):
        if buscar_imagen_region(f"aceptardirecciontipo{i}.png", region=(0, pyautogui.size()[1] // 2, pyautogui.size()[0], pyautogui.size()[1] // 2)):
            print(f"Dirección aceptada con tipo {i}.")
            aceptada = True
            break
    
    # Si no se encontró ninguna aceptación, cambiamos a la segunda forma
    if not aceptada:
        print("Dirección no aceptada, cambiando al tipo de dirección 2.")
        cambiar_tipo_direccion(direccion_2)
        return direccion_2

    # Si la dirección es aceptada, devolvemos la dirección final
    return direccion_1

# Función para cambiar al tipo de dirección 2
def cambiar_tipo_direccion(direccion_2):
    tipo_direccion = 'calle' if 'carrera' in direccion_2.lower() else 'carrera'
    numero_direccion = direccion_2.split()[1]
    nomenclatura_1 = direccion_2.split()[2]

    # Copiar la dirección original al portapapeles
    pyperclip.copy(direccion_2)

    # Foco en Bloc de notas y pegar la dirección
    subprocess.Popen(['notepad.exe'])
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'v')  # Pega el contenido

    # Mostrar una alerta emergente si no se puede introducir la dirección
    pyautogui.alert(text='No se pudo ingresar una dirección válida, intenta manualmente.', title='Error', button='OK')

# Ejemplo de uso
direccion_entrada = "calle 123 29F num 50-60"
direccion_final = procesar_direccion(direccion_entrada)
print(f"Dirección final procesada: {direccion_final}")

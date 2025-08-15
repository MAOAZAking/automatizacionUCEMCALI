# -*- coding: utf-8 -*-
import pyautogui
import pyperclip
import time
import re
import ctypes
import cv2
import numpy as np
from PIL import ImageGrab
import sys
import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw

# Configuraciones globales
pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = False

# Esta lista de escalas permite que la búsqueda de imágenes sea flexible a diferentes resoluciones de pantalla.
ESCALAS_POSIBLES = np.linspace(0.5, 1.5, num=11)
TOLERANCIA_DETECTADA = 0.75

def mostrar_ventana_exito(cantidad):
    """
    Muestra una ventana de éxito con la cantidad de correos procesados.
    """
    ventana = tk.Tk()
    ventana.title("¡Proceso Completo!")
    
    ventana.geometry("400x200")
    ventana.config(bg="#A8E6CF")

    mensaje = f"¡Listo! {cantidad} correo(s) han sido procesados correctamente.\nSe requiere atención del operador."
    etiqueta = tk.Label(ventana, text=mensaje,
                        font=("Helvetica", 14), fg="white", bg="#A8E6CF", padx=20, pady=20)
    etiqueta.pack(pady=20)

    boton_cerrar = tk.Button(ventana, text="Cerrar", font=("Helvetica", 12), bg="#FF8C00", fg="white", command=ventana.quit)
    boton_cerrar.pack(pady=10)

    ventana.mainloop()


def asegurar_foco_ventana(titulo):
    """
    Asegura que se enfoca una ventana por nombre parcial.
    Si la ventana se encuentra, la activa y le da foco.
    """
    try:
        ventanas_encontradas = [v for v in gw.getAllWindows() if titulo.lower() in v.title.lower()]
        
        if ventanas_encontradas:
            v = ventanas_encontradas[0]  # Tomar la primera ventana que coincida (si hay varias)
            
            if "Bloc de notas" in titulo:
                # Minimizar y luego restaurar para mantenerlo al fondo
                if v.isMinimized:
                    v.restore()  # Restaurar si está minimizada
                v.activate()  # Activar la ventana
                time.sleep(1)
                v.minimize()  # Minimizar de nuevo para mantenerlo en el fondo
                print(f"Ventana activada y minimizada: {v.title}")
                return True
            
            elif "Gestion" in titulo:
                # Activar 'Gestion'
                if v.isMinimized:
                    v.restore()  # Restaurar si está minimizada
                v.activate()  # Activar la ventana
                time.sleep(1)
                print(f"Ventana activada: {v.title}")
                return True
            
            elif "correo" in titulo:
                # Finalmente, activar 'Correo' y ponerlo al frente
                if v.isMinimized:
                    v.restore()  # Restaurar si está minimizada
                v.activate()  # Activar la ventana
                time.sleep(1)
                print(f"Ventana activada al frente: {v.title}")
                return True
            
        else:
            print(f"No se encontró ninguna ventana que contenga: {titulo}")
            return False
            
    except Exception as e:
        print(f"Error al enfocar ventana '{titulo}': {e}")
        return False


def log(msg):
    """
    Imprime un mensaje con un prefijo de log.
    """
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
    """
    Selecciona todo el texto del correo activo y lo copia al portapapeles.
    """
    try:
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        texto = pyperclip.paste()
        if texto.strip():
            log(f"Texto copiado del correo: {texto[:50]}...")
            return texto
        else:
            log("Texto copiado está vacío.")
            return None
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al copiar el contenido del correo para buscar su NÚMERO DE CONTRATO: {e}")

def buscar_id_en_texto(texto):
    """
    Busca un número de contrato o un número de teléfono en el texto del correo.
    Devuelve una tupla (id_contrato, id_telefono).
    """
    try:
        # --- Lógica de búsqueda de Contrato y Teléfono ---
        texto_normalizado = texto.lower().replace('(', '').replace(')', '').replace('-', '')
        
        # Patrón para buscar el número de contrato
        patron_contrato = (
            r'(?:'
                r'n[úu]?m(?:ero)?\.?\s*'
                r'(?:de\s*)?'
                r'contrato|'
                r'contrato|'
                r'n°|'
                r'#|'
                r'nro\.?|'
                r'nº'
                r'no'
                r'no.'
                r'no:'
                r'no...'
            r')\s*(\d{4,})'
        )

        match_contrato = re.search(patron_contrato, texto.lower())
        if match_contrato:
            id_encontrado = match_contrato.group(1)
            log(f"ID de contrato encontrado: {id_encontrado}")
            return id_encontrado, None
        
        # Patrón para buscar el número de teléfono
        patron_telefono = (
            r'(?:'
                r'línea(?: base)?|'
                r'numero telef[oó]nico|'
                r'telefono|'
                r'celular|'
                r'm[oó]vil|'
                r'fono|'
                r'tel'
            r')\s*(?:[.:]?\s*)?'
            r'(\d{7,10})'
        )
        
        match_telefono = re.search(patron_telefono, texto_normalizado)
        if match_telefono:
            numero_telefonico = match_telefono.group(1)
            if len(numero_telefonico) == 7 and not numero_telefonico.startswith('602'):
                numero_telefonico = '602' + numero_telefonico
            
            log(f"Número telefónico encontrado: {numero_telefonico}")
            return None, numero_telefonico
        
        log("No se encontró ni un número de contrato válido ni un número telefónico en el texto.")
        return None, None
    
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al buscar NÚMERO DE CONTRATO o TELÉFONO: {e}")
        return None, None


def copiar_id_a_portapapeles(idtexto):
    """
    Copia un texto al portapapeles.
    """
    try:
        pyperclip.copy(idtexto)
        log("Número de contrato copiado al portapapeles.")
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al copiar el número de contrato al portapapeles, para buscar la informacion en el SIGT: {e}")

def mostrar_alerta_y_terminar(mensaje="ATENCIÓN: El correo actual no tiene ID\nSe requiere atención inmediata de un usuario"):
    """
    Muestra una alerta de Windows y termina el programa.
    """
    import ctypes, sys
    ctypes.windll.user32.MessageBoxW(0, mensaje, "¡Advertencia!", 0)
    sys.exit()

def hacer_clic_en_imagen(nombre_imagen, descripcion="", tiempo_espera=3, region=None, multiple=False):
    """
    Busca una imagen en pantalla con diferentes escalas y hace clic en ella si la encuentra.
    """
    print(f"[🧠] Buscando: {descripcion} → Imagen: {nombre_imagen}")

    template = cv2.imread(nombre_imagen, cv2.IMREAD_GRAYSCALE)
    if template is None:
        print(f"[❌] No se pudo cargar la imagen {nombre_imagen}")
        return None

    h_template, w_template = template.shape

    tiempo_inicio = time.time()
    escalas = np.linspace(0.6, 1.4, num=20)

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
                return mejores_puntos

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
    Busca la imagen, ignora la primera coincidencia y hace clic en la segunda.
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
    """
    Simula la combinación de teclas Alt+Tab para cambiar de ventana.
    """
    pyautogui.keyDown('alt')
    for _ in range(veces):
        pyautogui.press('tab')
    pyautogui.keyUp('alt')

def extraer_dato_desde_etiqueta(etiqueta, imagen_etiqueta, desplazamiento_x=75, desplazamiento_y=0, convertir=False):
    """
    Extrae un dato en pantalla a partir de la ubicación de una imagen (etiqueta).
    """
    try:
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

        presionar_alt_tab_veces(1)
        pyautogui.write(f"{etiqueta} : ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        presionar_alt_tab_veces(1)

        log(f"{etiqueta} extraído y pegado: {valor.strip()[:40]}...")

        return valor

    except Exception as e:
        limpiar_estado_o_cerrar(f"Error al copiar el número de contrato al portapapeles, para buscar la informacion en el SIGT: {e}")
        log(f"Error al extraer {etiqueta}: {e}")
        return None

def extraer_plan_desde_tabla():
    """
    Extrae el "Plan" de una tabla en pantalla.
    """
    try:
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
    """
    Muestra un mensaje de error y cierra el programa.
    """
    log(f"Error crítico: {mensaje_error}")
    mostrar_alerta_y_terminar(mensaje_error)


def realizar_acciones_teclado(idtexto):
    """
    Realiza todas las acciones de búsqueda y extracción de datos en la aplicación.
    """
    try:
        # Volver al bloc de notas y preparar la entrada
        presionar_alt_tab_veces(2)
        pyautogui.press('enter')
        pyautogui.write("******************************************************************************")
        pyautogui.press('enter')
        pyautogui.write("Número de contrato: ")
        pyperclip.copy(idtexto)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # Cambiar a la ventana de Gestión ADSL
        presionar_alt_tab_veces(2)

        # Lógica para adaptarse al estado de la ventana de Gestión ADSL
        log("Adaptando al estado de la ventana de Gestión ADSL...")
        if hacer_clic_en_imagen("imagenes/forma_buscar.png", "Formulario de búsqueda"):
            log("El formulario de búsqueda ya está abierto. Continuando.")
        elif hacer_clic_en_imagen("imagenes/consultas.png", "Botón 'Consultas'"):
            log("Se hizo clic en el botón 'Consultas'. Buscando el formulario de búsqueda ahora.")
            if not hacer_clic_en_imagen("imagenes/forma_buscar.png", "Formulario de búsqueda"):
                limpiar_estado_o_cerrar("No se encontró el formulario de búsqueda después de hacer clic en 'Consultas'.")
        elif hacer_clic_en_imagen("imagenes/mostrar_menu_para_Seleccionar_consultas.png", "Botón para mostrar menú"):
            log("Se hizo clic en el botón de menú. Buscando 'Consultas' ahora.")
            if not hacer_clic_en_imagen("imagenes/consultas.png", "Botón 'Consultas'"):
                limpiar_estado_o_cerrar("No se encontró el botón 'Consultas' después de abrir el menú.")
            log("Se hizo clic en el botón 'Consultas'. Buscando el formulario de búsqueda ahora.")
            if not hacer_clic_en_imagen("imagenes/forma_buscar.png", "Formulario de búsqueda"):
                limpiar_estado_o_cerrar("No se encontró el formulario de búsqueda después de hacer clic en 'Consultas'.")
        else:
            limpiar_estado_o_cerrar("No se encontró ningún punto de entrada conocido en la ventana de Gestión ADSL.")

        # Continuación del flujo normal una vez que el formulario de búsqueda está listo.
        hacer_clic_en_imagen("imagenes/seleccionar_contrato.png", "Seleccionar opción Contrato")
        hacer_clic_en_imagen("imagenes/valor.png", "Campo de Valor del contrato")
        pyautogui.hotkey('ctrl', 'v')

        hacer_clic_en_imagen("imagenes/buscar.png", "Botón Buscar")
        time.sleep(2)

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
        hacer_clic_en_imagen("imagenes/seleccionar_carpeta.png", "Seleccionar carpeta de destino")
        
        hacer_clic_en_imagen("imagenes/ver_correos_procesador.png", "Ver correos procesador")
        hacer_clic_en_imagen("imagenes/acomodar_orden_de_correos_segun_bloc_de_notas.png", "Poner orden descendente de correos procesados")
        hacer_clic_en_imagen("imagenes/elecciona_primer_correo_procesado.png", "Seleccionar primer correo procesado")
        
        # Enfocamos el bloc de notas al final para que el operador pueda ver el resultado.
        enfocado = asegurar_foco_ventana("Bloc de notas")
        if not enfocado:
            log("No se pudo enfocar la ventana 'BLOC DE NOTAS'")

        log("Acciones completadas.")
        return True
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error en realizar_acciones_teclado(): {e}")
        return False

def main():
    """
    Función principal que ejecuta el flujo de automatización.
    """
    if not validar_imagenes([
        "imagenes/consultas.png", "imagenes/forma_buscar.png", "imagenes/seleccionar_contrato.png",
        "imagenes/seleccionar_contrato_activo.png", "imagenes/valor.png", "imagenes/buscar.png",
        "imagenes/nombre.png", "imagenes/email.png", "imagenes/numero_contacto.png",
        "imagenes/tipo_cliente.png", "imagenes/informacion_contrato.png", "imagenes/mover_a.png",
        "imagenes/mostrar_todas_las_carpetas.png", "imagenes/seleccionar_carpeta.png",
        "imagenes/mostrar_menu_para_Seleccionar_consultas.png"
    ]):
        mostrar_alerta_y_terminar("Faltan imágenes necesarias. Revise la carpeta 'imagenes'.")

    log("Esperando 7 segundos para preparar el entorno...")


# Lista de ventanas a enfocar en el orden especificado
ventanas_a_enfocar = ["Bloc de notas", "Gestion ADSL", "Correo:"]
for titulo in ventanas_a_enfocar:
    enfocado = asegurar_foco_ventana(titulo)
    time.sleep(0.5)
    if not enfocado:
        print(f"No se pudo enfocar la ventana: '{titulo}'.")

    intentos_max = 5
    intentos = 0
    correos_procesados = 0

    for i in range(10):
        log(f"Repetición #{i + 1} para procesar el correo...")
        
        while intentos < intentos_max:
            intentos += 1
            log(f"Intento #{intentos} para procesar correo...")

            # Clics para seleccionar el correo.
            x = pyautogui.size().width // 2 - 419
            y = pyautogui.size().height // 2 - 107
            pyautogui.click(x, y)
            time.sleep(1)
            x = pyautogui.size().width // 2 + 139
            y = pyautogui.size().height // 2 - 27
            pyautogui.click(x, y)

            texto_del_correo = copiar_texto_del_correo()

            if texto_del_correo:
                contrato_id, telefono_encontrado = buscar_id_en_texto(texto_del_correo)
                
                if contrato_id:
                    log(f"Contrato encontrado: {contrato_id}")
                    copiar_id_a_portapapeles(contrato_id)
                    if realizar_acciones_teclado(contrato_id):
                        correos_procesados += 1
                        break
                elif telefono_encontrado:
                    log(f"Teléfono encontrado: {telefono_encontrado}")
                    pyperclip.copy(telefono_encontrado)
                    presionar_alt_tab_veces(1)
                    pyautogui.write("Número de Teléfono: ")
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press('enter')
                    presionar_alt_tab_veces(1)
                    
                    correos_procesados += 1
                    break
                else:
                    log("Ni NÚMERO DE CONTRATO ni TELÉFONO encontrados. Reintentando...")
            else:
                log("Texto copiado está vacío. Reintentando...")
        
        if not (contrato_id or telefono_encontrado) and intentos >= intentos_max:
            log("No se encontró NÚMERO DE CONTRATO ni TELÉFONO después de múltiples intentos. Finalizando el proceso.")
            mostrar_ventana_exito(correos_procesados)
            break
        elif not (contrato_id or telefono_encontrado):
            pass

        intentos = 0
        contrato_id = None
        telefono_encontrado = None

    log("Todos los correos han sido procesados. Mostrando ventana de éxito.")
    mostrar_ventana_exito(correos_procesados)



if __name__ == "__main__":
    main()

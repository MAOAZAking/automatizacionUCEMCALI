import pyautogui
import pyperclip
import time
import re
import ctypes
import cv2
import numpy as np
import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw

# Configuraciones globales
pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = False

# Esta lista de escalas permite que la búsqueda de imágenes sea flexible a diferentes resoluciones de pantalla.
ESCALAS_POSIBLES = np.linspace(0.7, 1.3, num=13)
TOLERANCIA_DETECTADA = 0.68

def mostrar_ventana_exito(cantidad):
    """
    Muestra una ventana de éxito con la cantidad de correos procesados.
    La ventana se fuerza a estar en primer plano (topmost) como una alerta modal.
    """
    ventana = tk.Tk()
    ventana.title("¡Proceso Completo!")

    # Tamaño y estilo
    ventana.geometry("600x200")
    ventana.config(bg="#A8E6CF")

    # Mantener la ventana al frente
    ventana.attributes('-topmost', True)
    ventana.update()
    ventana.focus_force()

    # Mensaje
    mensaje = f"¡Listo! {cantidad} correo(s) han sido procesados correctamente.\nSe requiere atención del operador."
    etiqueta = tk.Label(
        ventana,
        text=mensaje,
        font=("Helvetica", 14),
        fg="white",
        bg="#A8E6CF",
        padx=20,
        pady=20
    )
    etiqueta.pack(pady=20)

    # Botón de cierre
    boton_cerrar = tk.Button(
        ventana,
        text="Cerrar",
        font=("Helvetica", 12),
        bg="#FF8C00",
        fg="white",
        command=ventana.quit
    )
    boton_cerrar.pack(pady=10)

    # Forzar foco final antes de mostrar
    ventana.lift()
    ventana.attributes('-topmost', True)
    ventana.after(100, lambda: ventana.focus_force())

    ventana.mainloop()



def asegurar_foco_ventana(titulo_parcial):
    """
    Busca y activa la ventana cuyo título contenga titulo_parcial.
    Este método es el más confiable para traer una ventana al frente.
    """
    try:
        log(f"Intentando enfocar ventana con título que contenga '{titulo_parcial}'...")
        # Buscamos todas las ventanas que contengan el título parcial.
        ventanas_encontradas = [v for v in gw.getAllWindows() if titulo_parcial.lower() in v.title.lower()]
        
        if ventanas_encontradas:
            ventana = ventanas_encontradas[0] # Tomamos la primera coincidencia
            # Restaurar y activar si está minimizada.
            if ventana.isMinimized:
                ventana.restore()
            ventana.activate()
            time.sleep(0.5) # Damos un tiempo para que el sistema procese el cambio de foco.
            log(f"Ventana '{ventana.title}' activada.")
            return True
        else:
            log(f"No se encontró ventana con título que contenga '{titulo_parcial}'.")
            return False
    except Exception as e:
        log(f"Error al enfocar ventana: {e}")
        return False

def ventana_existe(titulo_parcial):
    """Verifica si una ventana con el título parcial existe sin activarla."""
    try:
        ventanas_encontradas = [v for v in gw.getAllWindows() if titulo_parcial.lower() in v.title.lower()]
        if len(ventanas_encontradas) > 0:
            return True
        return False
    except Exception as e:
        log(f"Error al verificar si la ventana '{titulo_parcial}' existe: {e}")
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
        texto_normalizado = texto.lower().replace('(', '').replace(')', '').replace('-', '').replace('—', '').replace('...', '')
        
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

def hacer_clic_en_imagen(nombre_imagen, descripcion="", tiempo_espera=3, region=None, multiple=False, click=True):
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

    while time.time() - tiempo_inicio < tiempo_espera:
        captura = pyautogui.screenshot(region=region)
        captura = cv2.cvtColor(np.array(captura), cv2.COLOR_RGB2GRAY)

        mejores_puntos = []
        for escala in ESCALAS_POSIBLES:
            resized_template = cv2.resize(template, (0, 0), fx=escala, fy=escala)
            h, w = resized_template.shape

            if captura.shape[0] < h or captura.shape[1] < w:
                continue

            resultado = cv2.matchTemplate(captura, resized_template, cv2.TM_CCOEFF_NORMED)
            ubicaciones = np.where(resultado >= TOLERANCIA_DETECTADA)

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
            if click:
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

        pyautogui.hotkey('ctrl', 'c')
        valor = pyperclip.paste()

        if convertir:
            valor = valor.replace(" ", "")
            if len(valor) == 7:
                valor = "602" + valor
            pyperclip.copy(valor)

        asegurar_foco_ventana("Bloc de notas")
        pyautogui.write(f"{etiqueta} : ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        asegurar_foco_ventana("Gestion ADSL")

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

        pyautogui.hotkey('ctrl', 'c')
        plan = pyperclip.paste()

        asegurar_foco_ventana("Bloc de notas")
        pyautogui.write("Plan : ")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        asegurar_foco_ventana("Bloc de notas")

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
        asegurar_foco_ventana("Bloc de notas")
        pyautogui.press('enter')
        pyautogui.write("******************************************************************************")
        pyautogui.press('enter')
        pyautogui.write("Número de contrato: ")
        pyperclip.copy(idtexto)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # Cambiar a la ventana de Gestión ADSL
        asegurar_foco_ventana("Gestion ADSL")

        # Lógica para adaptarse al estado de la ventana de Gestión ADSL
        log("Adaptando al estado de la ventana de Gestión ADSL...")
        if hacer_clic_en_imagen("imagenes/consultas.png", "Botón 'Consultas'"):
            log("Se hizo clic en el botón 'Consultas'. Buscando el formulario de búsqueda ahora.")
            if not hacer_clic_en_imagen("imagenes/forma_buscar.png", "Formulario de búsqueda"):
                limpiar_estado_o_cerrar("No se encontró el formulario de búsqueda después de hacer clic en 'Consultas'.")
        elif hacer_clic_en_imagen("imagenes/consultas_sin_seleccionar.png", "Botón 'Consultas sin seleccionar'"):
            log("Se hizo clic en el botón 'Consultas'. Buscando el formulario de búsqueda ahora.")
            if not hacer_clic_en_imagen("imagenes/forma_buscar.png", "Formulario de búsqueda"):
                limpiar_estado_o_cerrar("No se encontró el formulario de búsqueda después de hacer clic en 'Consultas'.")
        elif hacer_clic_en_imagen("imagenes/mostrar_menu_para_Seleccionar_consultas.png", "Botón para mostrar menú"):
            log("Se hizo clic en el botón de menú. Buscando 'Consultas' ahora.")
            if not hacer_clic_en_imagen("imagenes/consultas.png", "Botón 'Consultas'"):
                limpiar_estado_o_cerrar("No se encontró el botón 'Consultas' después de abrir el menú.")
            elif hacer_clic_en_imagen("imagenes/consultas_sin_seleccionar.png", "Botón 'Consultas sin seleccionar'"):
                if not hacer_clic_en_imagen("imagenes/forma_buscar.png", "Formulario de búsqueda"):
                    limpiar_estado_o_cerrar("No se encontró el formulario de búsqueda después de hacer clic en 'Consultas'.")
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

        asegurar_foco_ventana("Correo:")
        pyautogui.hotkey('ctrl', 'l')
        pyautogui.hotkey('ctrl', 'c')

        asegurar_foco_ventana("Bloc de notas")
        pyautogui.press('enter')
        pyautogui.write("URL del correo:")
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        pyautogui.press('enter')
        
        asegurar_foco_ventana("Correo:")
        hacer_clic_en_imagen("imagenes/mover_a.png", "Botón Mover Correo")
        hacer_clic_en_imagen("imagenes/mostrar_todas_las_carpetas.png", "Mostrar todas las carpetas")
        hacer_clic_en_imagen("imagenes/seleccionar_carpeta_correos_revisados.png", "Seleccionar carpeta de destino")
        
        hacer_clic_en_imagen("imagenes/ver_correos_procesador.png", "Ver correos procesador")
        hacer_clic_en_imagen("imagenes/acomodar_orden_de_correos_segun_bloc_de_notas.png", "Poner orden descendente de correos procesados")
        # Clics para seleccionar el correo.
        x = pyautogui.size().width // 2 - 419
        y = pyautogui.size().height // 2 - 107
        pyautogui.click(x, y)
                    
        # Enfocamos el bloc de notas al final para que el operador pueda ver el resultado.
        asegurar_foco_ventana("Bloc de notas")

        log("Acciones completadas.")
        return True
    except Exception as e:
        limpiar_estado_o_cerrar(f"Error en realizar_acciones_teclado(): {e}")
        return False

def validar_y_ordenar_correos_por_fecha():
    """
    Verifica si los correos están ordenados de más antiguo a más nuevo.
    Si están de reciente a antiguo, cambia el orden.
    Si no se encuentra el ícono de fecha, busca íconos alternativos de filtro y actúa en consecuencia.
    """
    log("Verificando y ajustando orden de correos por fecha...")

    # Primero, intenta encontrar el ícono que indica orden por fecha
    fecha_icono = hacer_clic_en_imagen("imagenes/cambiar_orden_de_correos_por_fecha.png", "Ícono de orden por fecha", tiempo_espera=2)

    if fecha_icono is None:
        log("No se encontró el ícono de fecha, buscando íconos de filtrado alternativos...")

        alternativas = [
            "imagenes/filtrado_de_reciente_a_antiguo.png",
            "imagenes/filtrado_de_antiguo_a_reciente.png"
        ]

        for img in alternativas:
            punto = hacer_clic_en_imagen(img, f"Ícono alternativo: {img}", tiempo_espera=2)
            if punto:
                # Clic 40px a la derecha del ícono encontrado
                x, y = punto[0] + 40, punto[1]
                pyautogui.click(x, y)
                log(f"Se hizo clic 40px a la derecha del ícono: {img}")
                break
        else:
            log("No se encontraron íconos alternativos para ordenar los correos.")
            return

        # Luego de usar íconos alternativos, intentar de nuevo encontrar el de fecha
        fecha_icono = hacer_clic_en_imagen("imagenes/cambiar_orden_de_correos_por_fecha.png", "Ícono de orden por fecha", tiempo_espera=2)

    if fecha_icono:
        # Ahora que se detectó el ícono de fecha, verificar el estado del ordenamiento
        log("Verificando tipo de orden (reciente a antiguo o viceversa)...")

        reciente_a_antiguo = hacer_clic_en_imagen("imagenes/filtrado_de_reciente_a_antiguo.png", "Orden de reciente a antiguo", tiempo_espera=1)

        if reciente_a_antiguo:
            # Si está en orden de reciente a antiguo, hacer clic para cambiarlo
            log("El orden es de reciente a antiguo. Cambiando a más antiguo a más nuevo...")
            pyautogui.click(fecha_icono[0], fecha_icono[1])
        else:
            log("El orden ya está en más antiguo a más nuevo. No se requiere acción.")
    else:
        log("No se pudo validar ni ajustar el orden por fecha.")


def abrir_bloc_de_notas():
    """Abre una nueva instancia del Bloc de notas."""
    log("Abriendo Bloc de notas...")
    pyautogui.hotkey('win'); pyautogui.write('Bloc de notas'); pyautogui.press('enter'); time.sleep(3)

def abrir_adsl():
    """Abre Chrome, navega a ADSL e intenta iniciar sesión."""
    log("Abriendo Google Chrome para ADSL...")
    pyautogui.hotkey('win'); pyautogui.write('Google Chrome'); pyautogui.press('enter'); time.sleep(3); pyautogui.hotkey('win', 'left'); pyautogui.hotkey('win', 'up'); pyautogui.hotkey('win', 'up'); pyautogui.hotkey('ctrl', 'l')
    pyautogui.write('http://adsl.emcali.net.co/'); pyautogui.press('enter'); time.sleep(5)

    # Reintentar el login de ADSL
    login_adsl_exitoso = False
    for intento_login in range(3):
        log(f"Intento de login en ADSL #{intento_login + 1}")
        if hacer_clic_en_imagen("imagenes/usuario_adsl.png", "Seleccionar usuario ADSL", tiempo_espera=7):
            pyautogui.press('down'); pyautogui.press('enter'); pyautogui.press('enter')
            login_adsl_exitoso = True
            log("Login en ADSL con autocompletado exitoso.")
            break
        else:
            log("No se encontró la imagen del usuario ADSL en este intento.")
            time.sleep(2)
    if not login_adsl_exitoso:
        log("No se pudo iniciar sesión en ADSL, se necesita que el operador inice sesión manualmente.")

def abrir_outlook():
    """Abre Chrome y navega a Outlook."""
    log("Abriendo Google Chrome para Outlook...")
    pyautogui.hotkey('win'); pyautogui.write('Google Chrome'); pyautogui.press('enter'); time.sleep(3); pyautogui.hotkey('win', 'left'); pyautogui.hotkey('win', 'up'); pyautogui.hotkey('win', 'up'); pyautogui.hotkey('ctrl', 'l')
    pyautogui.write('https://outlook.live.com/mail/0/'); pyautogui.press('enter'); time.sleep(5)

def preparar_entorno():
    """
    Función principal que ejecuta el flujo de automatización.
    """
    if not validar_imagenes([
        "imagenes/anuncio.png",
        "imagenes/eliminar.png",
        "imagenes/buscar.png",
        "imagenes/cambiar_orden_de_correos_por_fecha.png",
        "imagenes/consultas.png",
        "imagenes/consultas_sin_seleccionar.png",
        "imagenes/email.png",
        "imagenes/filtrado_de_antiguo_a_reciente.png",
        "imagenes/filtrado_de_reciente_a_antiguo.png",
        "imagenes/forma_buscar.png",
        "imagenes/informacion_contrato.png",
        "imagenes/mostrar_menu_para_seleccionar_consultas.png",
        "imagenes/mostrar_todas_las_carpetas.png",
        "imagenes/mover_a.png",
        "imagenes/nombre.png",
        "imagenes/numero_contacto.png",
        "imagenes/organizar_correos_del_mas_antiguo_al_mas_nuevo.png",
        "imagenes/seleccionar_carpeta_correos_revisados.png",
        "imagenes/seleccionar_carpeta_por_revisar.png",
        "imagenes/seleccionar_contrato.png",
        "imagenes/seleccionar_contrato_activo.png",
        "imagenes/tipo_cliente.png",
        "imagenes/valor.png",
        "imagenes/usuario_adsl.png"
    ]):
        mostrar_alerta_y_terminar("Faltan imágenes necesarias. Revise la carpeta 'imagenes'.")

    log("Paso 1: Verificando y abriendo las aplicaciones necesarias...")

    if not ventana_existe("Bloc de notas"):
        abrir_bloc_de_notas()
    else:
        log("Bloc de notas ya está abierto.")

    if not ventana_existe("Gestion ADSL"):
        abrir_adsl()
    else:
        log("Ventana de Gestion ADSL ya está abierta.")

    if not ventana_existe("Correo:"):
        abrir_outlook()
    else:
        log("Ventana de Correo (Outlook) ya está abierta.")
        
    log("Paso 2: Validar si la session de ADSL está iniciada...")
    asegurar_foco_ventana("Gestion ADSL")
    if not hacer_clic_en_imagen("imagenes/mostrar_menu_para_seleccionar_consultas.png", "Valida si aparece el menú de gestion que indica que la session ya etsa iniciada", tiempo_espera=5):
        # Reintentar el login de ADSL
        login_adsl_exitoso = False
        for intento_login in range(3):
            log(f"Intento de login en ADSL #{intento_login + 1}")
            if hacer_clic_en_imagen("imagenes/usuario_adsl.png", "Seleccionar usuario ADSL", tiempo_espera=7):
                pyautogui.press('down'); pyautogui.press('enter'); pyautogui.press('enter')
                login_adsl_exitoso = True
                log("Login en ADSL con autocompletado exitoso.")
                break
            else:
                log("No se encontró la imagen del usuario ADSL en este intento.")
                time.sleep(2)
        if not login_adsl_exitoso:
            log("No se pudo iniciar sesión en ADSL, se necesita que el operador inice sesión manualmente.")
            mostrar_alerta_y_terminar("No se pudo iniciar sesión en ADSL automáticamente. Por favor, inicie sesión manualmente.")

    log("Paso 3: Asegurando el foco de las ventanas en el orden correcto...")
    ventanas_a_enfocar = ["Bloc de notas", "Gestion ADSL", "Correo:"]
    for titulo in ventanas_a_enfocar:
        if not asegurar_foco_ventana(titulo):
            limpiar_estado_o_cerrar(f"Error CRÍTICO: No se pudo enfocar la ventana '{titulo}' después de la verificación.")
    log("Foco de ventanas asegurado. Iniciando el procesamiento de correos.")
    return True

def procesar_correos():
    """Bucle principal para procesar los correos uno por uno."""
    intentos_max = 5
    correos_procesados = 0

    # Bucle para procesar un número máximo de correos (ej. 20) para evitar ciclos infinitos.
    for _ in range(20):
        log(f"Iniciando ciclo de procesamiento de correo...")
        intentos = 0
        correo_procesado_exitosamente = False
        while intentos < intentos_max and not correo_procesado_exitosamente:
            intentos += 1
            log(f"Intento #{intentos} para procesar correo...")

            # Limpiar el portapapeles antes de realizar cualquier acción
            pyperclip.copy('0')  # Esto vacía el portapapeles


            # Aseguramos el foco en la ventana de Correo para cada intento
            asegurar_foco_ventana("Correo:")

            validar_y_ordenar_correos_por_fecha()
            time.sleep(2)

            # --- Bucle para eliminar todos los anuncios visibles ---
            log("Iniciando bucle de eliminación de anuncios...")
            intentos_fallidos_anuncio = 0
            while True:
                anuncio_encontrado = hacer_clic_en_imagen("imagenes/anuncio.png", "Anuncio en correo", tiempo_espera=2, click=False)
                
                if anuncio_encontrado:
                    log("Anuncio encontrado, intentando eliminarlo.")
                    intentos_fallidos_anuncio = 0  # Reiniciar contador de fallos porque encontramos uno
                    
                    if hacer_clic_en_imagen("imagenes/eliminar.png", "Botón para eliminar anuncio", tiempo_espera=2):
                        log("Anuncio eliminado. Buscando más anuncios...")
                        time.sleep(2)  # Esperar a que la UI se actualice
                    else:
                        log("No se encontró el botón para eliminar el anuncio. Se continuará buscando más anuncios.")
                else:
                    intentos_fallidos_anuncio += 1
                    log(f"No se encontraron anuncios en esta pasada. Intento de fallo #{intentos_fallidos_anuncio}.")
                    if intentos_fallidos_anuncio >= 2:
                        log("No se han encontrado anuncios en 2 intentos consecutivos. Saliendo del bucle de eliminación.")
                        break  # Salir del bucle while
            log("Fin del bucle de eliminación de anuncios. Procesando correo.")

            # --- Clics con coordenadas relativas para seleccionar el correo. Son frágiles a cambios de resolución o layout. ---
            # Clics para seleccionar el correo.
            x = pyautogui.size().width // 2 - 419
            y = pyautogui.size().height // 2 - 107
            pyautogui.click(x, y)

            # Clics sobre el contenido del correo.
            x = pyautogui.size().width // 2 + 139
            y = pyautogui.size().height // 2 - 27
            pyautogui.click(x, y)

            texto_del_correo = copiar_texto_del_correo()

            # Después de copiar el texto del correo, verificamos si el portapapeles está vacío
            if pyperclip.paste().strip() == "0":
                log("El portapapeles está vacío. No se pudo copiar el texto del correo.")

                # Si el portapapeles está vacío, significa que no hay más correos por revisar.
                log("No hay más correos por revisar, cambiando a la carpeta de correos revisados.")

                # Dar clic en la carpeta de "correos revisados"
                hacer_clic_en_imagen("imagenes/seleccionar_carpeta_correos_revisados.png", "Seleccionar carpeta 'Correos revisados'")

                # Esperar a que se carguen los correos revisados
                time.sleep(1)

                # Seleccionar el correo más arriba
                x = pyautogui.size().width // 2 - 419
                y = pyautogui.size().height // 2 - 107
                pyautogui.click(x, y)

                # Dar clic en la imagen para cambiar el orden de los correos por fecha
                hacer_clic_en_imagen("imagenes/cambiar_orden_de_correos_por_fecha.png", "Cambiar orden de correos por fecha")

                # Dar clic en la imagen para organizar los correos del más antiguo al más nuevo
                hacer_clic_en_imagen("imagenes/organizar_correos_del_mas_antiguo_al_mas_nuevo.png", "Organizar correos del más antiguo al más nuevo")

                # Esperar 2 segundos para que se organicen los correos
                time.sleep(2)

                # Mostrar el mensaje de que todos los correos han sido procesados
                mostrar_ventana_exito(correos_procesados)

                # Terminar el flujo si no hay más correos que revisar
                return # Finaliza la función de procesamiento

            elif texto_del_correo:
                contrato_id, telefono_encontrado = buscar_id_en_texto(texto_del_correo)

                if contrato_id:
                    log(f"Contrato encontrado: {contrato_id}")
                    copiar_id_a_portapapeles(contrato_id)
                    if realizar_acciones_teclado(contrato_id):
                        correos_procesados += 1
                        correo_procesado_exitosamente = True
                elif telefono_encontrado:
                    log(f"Teléfono encontrado: {telefono_encontrado}")
                    pyperclip.copy(telefono_encontrado)
                    asegurar_foco_ventana("Bloc de notas")
                    pyautogui.write("Número de Teléfono: ")
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press('enter')
                    asegurar_foco_ventana("Correo:")

                    correos_procesados += 1
                    correo_procesado_exitosamente = True
                else:
                    log("Ni NÚMERO DE CONTRATO ni TELÉFONO encontrados. Reintentando...")
            else:
                log("Texto copiado está vacío. Reintentando...")

        # Si se agotaron los intentos para un correo y no se procesó, finalizar.
        if not correo_procesado_exitosamente:
            log("No se encontró NÚMERO DE CONTRATO ni TELÉFONO después de múltiples intentos. Finalizando el proceso.")
            mostrar_ventana_exito(correos_procesados)
            return

    log("Se ha alcanzado el límite de correos a procesar en una ejecución. Finalizando.")
    mostrar_ventana_exito(correos_procesados)

def main():
    if preparar_entorno():
        procesar_correos()

if __name__ == "__main__":
    main()
    
############################  Se mejoro la funcion de remover los anuncios integrando la capacidad de eliminar todos los anuncios que puedan aparecer   #########################################################
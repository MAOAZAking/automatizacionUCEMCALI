import re
import time

# Expresiones regulares para detectar partes comunes de direcciones
regex_calle = r"(cl|call|cll|calle|carrera|kr|cr)\.?(\s*\d+)\s?([a-zA-Z]?\s*\d*\w*)"
regex_nomenclatura = r"(\d+[a-zA-Z]*)\s*(#|num|numero|nº)?\s*([a-zA-Z0-9\s\-]*)"

# Función para validar y separar la dirección
def procesar_direccion(direccion):
    # Comprobamos si se ajusta al patrón de calle o carrera
    match_calle = re.search(regex_calle, direccion.lower())
    if match_calle:
        tipo_direccion = match_calle.group(1).capitalize()  # Calle, carrera, etc.
        numero_direccion = match_calle.group(2)  # Número de dirección
        nomenclatura_1 = match_calle.group(3).strip()  # Parte de nomenclatura (si existe)
    else:
        # Si no encuentra coincidencias, devolvemos un error o mensaje
        return "Dirección no válida"
    
    # Buscamos la nomenclatura después del número de dirección
    match_nomenclatura = re.search(regex_nomenclatura, direccion)
    if match_nomenclatura:
        numero_con_letra = match_nomenclatura.group(1)  # Número con letra (ej: 29F)
        nomenclatura_2 = match_nomenclatura.group(3)  # Segunda parte de la nomenclatura
    else:
        # Si no se encuentra nomenclatura adicional
        numero_con_letra = numero_direccion
        nomenclatura_2 = ""
    
    # Vamos a organizar los resultados en el formato que se requiere
    direccion_1 = f"{tipo_direccion} {numero_con_letra} {nomenclatura_1} {nomenclatura_2}"
    
    # Cambiar la dirección: si se encuentra calle, cambiar por carrera y viceversa
    if 'calle' in tipo_direccion.lower():
        tipo_direccion = tipo_direccion.replace('calle', 'carrera')
    elif 'carrera' in tipo_direccion.lower():
        tipo_direccion = tipo_direccion.replace('carrera', 'calle')
    
    direccion_2 = f"{tipo_direccion} {nomenclatura_1} {numero_con_letra} {nomenclatura_2}"

    # Validamos con una espera de 7 segundos y comparación
    time.sleep(7)  # Simulamos el cambio
    # Verificamos si hubo un cambio (esto puede ser validado con una captura de pantalla en un entorno real)
    if direccion_1 != direccion_2:
        return direccion_1  # La primera forma fue aceptada
    else:
        return direccion_2  # Si no cambia, se acepta la segunda forma

# Ejemplo de uso
direccion_entrada = "calle 123 29F # 50-60"
direccion_final = procesar_direccion(direccion_entrada)

print(direccion_final)

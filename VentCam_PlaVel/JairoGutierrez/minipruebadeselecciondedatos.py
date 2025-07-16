from PIL import ImageGrab
import cv2
import numpy as np

pantalla = ImageGrab.grab()
pantalla_cv = cv2.cvtColor(np.array(pantalla), cv2.COLOR_RGB2BGR)
imagen = cv2.imread("imagenes/nombre.png")

res = cv2.matchTemplate(pantalla_cv, imagen, cv2.TM_CCOEFF_NORMED)
_, max_val, _, max_loc = cv2.minMaxLoc(res)

print("Confianza de detección:", max_val)

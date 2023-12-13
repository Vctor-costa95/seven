import pytesseract
import cv2

path = r'C:/Program Files/Tesseract-OCR'
imagem = cv2.imread(r"printseven.png")

pytesseract.pytesseract.tesseract_cmd = path + r'/tesseract.exe'

texto = pytesseract.image_to_string(imagem)

print(texto)


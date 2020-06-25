
from PIL import Image
import pytesseract

img = Image.open('photo_2020-06-19 00.36.07.jpeg')
text = pytesseract.image_to_string(img, lang='chi_tra')
print(text)

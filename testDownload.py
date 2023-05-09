import requests
from PIL import Image
from io import BytesIO

url = 'https://10.95.224.27:9083/bicp/verifycode?' # Esta es la URL que genera la imagen CAPTCHA
# Realiza la petición GET para obtener la imagen
response = requests.get(url, verify=False)
# Crea una instancia de la clase Image para procesar la imagen
img = Image.open(BytesIO(response.content))
# Guarda la imagen en un archivo en el disco
img.save('captcha.jpg', 'JPEG')

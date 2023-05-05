import time
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from twocaptcha import TwoCaptcha


# Especifica la ruta del controlador de Chrome
pathDriver = "112/chromedriver.exe"

# Especifica los detalles de la URL, nombre de usuario y contraseña
url = "https://2captcha.com/demo/normal"

# Configura el controlador de Chrome y abre la URL especificada
driver = webdriver.Chrome(executable_path=pathDriver)
driver.get(url)

# Espera a que se cargue completamente la página web
driver.implicitly_wait(10)


img = driver.find_element(By.XPATH,'//*[@id="root"]/div/main/div/section/form/div/img')
src = img.get_attribute('src')
img = requests.get(src)
with open('captcha.jpg', 'wb') as f:
    f.write(img.content)



solver = TwoCaptcha('8f63da7191fe11e63148c3d8b28c71f2')

id = solver.send(file='captcha.jpg')
time.sleep(20)

code = solver.get_result(id)

code_field = driver.find_element(By.XPATH, '//*[@id="simple-captcha-field"]')
code_field.send_keys(code)

#driver.close()

time.sleep(1000)

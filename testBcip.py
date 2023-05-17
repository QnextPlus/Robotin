import time
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from twocaptcha import TwoCaptcha
from PIL import Image
from io import BytesIO

# Especifica la ruta del controlador de Chrome
pathDriver = "113/chromedriver.exe"

#Especifica Directorio de descarga
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

#Para Ignorar los errores de certificado SSL (La conexion no es privada)
options.add_argument("--ignore-certificate-errors")

# Especifica los detalles de la URL, nombre de usuario y contraseña
url = "https://10.95.224.27:9083/bicp/login.action"
username = "E8234353"
password = "JJJmm88***1"


# Configura el controlador de Chrome y abre la URL especificada
driver = webdriver.Chrome(executable_path=pathDriver, options=options)
driver.get(url)
"""
# Espera a que se cargue completamente la página web
driver.implicitly_wait(10)

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                      'button#details-button')))\
    .click()
WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                      'a#proceed-link')))\
    .click()

"""    
time.sleep(2)


#cierra ventana emergente
elemento_cierre = driver.find_element(By.XPATH, '//*[@id="tipClose"]')
elemento_cierre.click()

# Busca los campos de usuario y contraseña y los llena con los valores especificados
campoInputUser = driver.find_element(By.ID, 'username2')
campoInputUser.clear()
actions = ActionChains(driver)
actions.move_to_element(campoInputUser)
actions.click()
actions.send_keys(username)
actions.perform()

campoInputPassword = driver.find_element(By.ID, 'password2')
campoInputPassword.clear()
actions = ActionChains(driver)
actions.move_to_element(campoInputPassword)
actions.click()
actions.send_keys(password)
actions.perform()

# Busca el elemento deseado utilizando un selector CSS
element = driver.find_element(By.ID,'validateimg')
# Captura una imagen del elemento
element.screenshot('capture.jpg')

solver = TwoCaptcha('8f63da7191fe11e63148c3d8b28c71f2')

id = solver.send(file='capture.jpg')
time.sleep(10)

code = solver.get_result(id)

code_field = driver.find_element(By.XPATH, '//*[@id="validate"]')
code_field.send_keys(code)

time.sleep(5)
#presiona boton submit 
submit_button = driver.find_element(By.CLASS_NAME, 'login_submit_btn')
submit_button.click()

time.sleep(5)
#Menu principal
systemMenu = driver.find_element(By.ID, 'systemMenu')
systemMenu.click()

time.sleep(5)
#selecciona Resource manager del menu principal
resourceManager = driver.find_element(By.ID, 's_800')
resourceManager.click()
time.sleep(5)

#Selecciona searchResource para busqueda de campaña
searchResource = driver.find_element(By.ID, 'searchResource')
searchResource.click()
time.sleep(5)

#Codigo de la campaña
codigoCampana = '1259'

#input donde se ingresa codigo de la campaña
inputCampana = driver.find_element(By.ID, 'searchCondtion_input_value')
inputCampana.send_keys(codigoCampana)
botonBusqueda = driver.find_element(By.ID, 'searchBtn')
botonBusqueda.click()
time.sleep(2)

#selecciona documento
#gridLink = driver.find_element(By.ID, 'AID_43205')
#gridLink.click()




time.sleep(10000)

print('termino')
exit()




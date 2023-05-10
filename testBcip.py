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

# Especifica los detalles de la URL, nombre de usuario y contraseña
url = "https://10.95.224.27:9083/bicp/login.action"
username = "E8234180X"
password = "RoboT$n2023X%"


# Configura el controlador de Chrome y abre la URL especificada
driver = webdriver.Chrome(executable_path=pathDriver, options=options)
driver.get(url)

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

#presiona boton submit
submit_button = driver.find_element(By.XPATH, '//*[@id="submitBtn"]/a')
submit_button.click()

time.sleep(10000)

print('termino')
exit()


# Envía el formulario y cierra el controlador de Chrome
password_field.send_keys(Keys.RETURN)

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                      'a#menu-button')))\
    .click()

menuReporte = driver.find_element(By.XPATH, '//*[@id="sidebar"]/div[1]/div[3]/div[2]/ul/li[3]/a')
menuReporte.click()

menuReporteNivel2 = driver.find_element(By.XPATH, '//*[@id="sidebar"]/div[1]/div[3]/div[2]/ul/li[3]/ul/li[2]')
menuReporteNivel2.click()

menuReporteNivel3 = driver.find_element(By.XPATH, '//*[@id="sidebar"]/div[1]/div[3]/div[2]/ul/li[3]/ul/li[2]/ul/li[2]/a')
menuReporteNivel3.click()

#menuPeriodo = driver.find_element(By.XPATH, '//*[@id="summary-section"]/div[1]/div[1]/div[1]/div[1]/button')
#menuPeriodo.click()

#menuPeriodoNivel2 = driver.find_element(By.XPATH, '//*[@id="summary-section"]/div[1]/div[1]/div[1]/div[2]/div/ul/li[3]')
#menuPeriodoNivel2.click()

menuSede = driver.find_element(By.XPATH, '//*[@id="summary-section"]/div[1]/div[2]/div/div[1]/button')
menuSede.click()

menuSedeNivel2 = driver.find_element(By.XPATH, '//*[@id="summary-section"]/div[1]/div[2]/div/div[2]/div[2]/ul/li[1]')
menuSedeNivel2.click()

btnDescargar = driver.find_element(By.XPATH, '//*[@id="summary-section"]/div[2]/div/div/button[3]')
btnDescargar.click()

archivo = driver.find_element(By.CLASS_NAME, 'file-name-span')
archivo.click()

#driver.close()

time.sleep(10000)

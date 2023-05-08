import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


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
url = "https://qnexweb.com/lapositiva/in-config"
username = "michael.luque"
password = "73031615"

# Configura el controlador de Chrome y abre la URL especificada
driver = webdriver.Chrome(executable_path=pathDriver, options=options)
driver.get(url)

# Espera a que se cargue completamente la página web
driver.implicitly_wait(10)

# Busca los campos de usuario y contraseña y los llena con los valores especificados
username_field = driver.find_element(By.NAME, 'username')
password_field = driver.find_element(By.NAME, 'password')
username_field.send_keys(username)
password_field.send_keys(password)

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

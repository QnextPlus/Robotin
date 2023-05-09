from selenium import webdriver
from selenium.webdriver.common.by import By

# Inicializa el driver de Selenium
driver = webdriver.Chrome()

# Navega a la página web deseada
driver.get('https://es.wikipedia.org/wiki/Wikipedia:Portada')

# Busca el elemento deseado utilizando un selector CSS
element = driver.find_element(By.XPATH,'/html/body/div[1]/header/div[1]/a/img')

# Captura una imagen del elemento
element.screenshot('capture.jpg')

# Cierra el driver de Selenium
driver.quit()

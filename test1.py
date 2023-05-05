from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from twocaptcha import TwoCaptcha

# Especifica los detalles de la URL, nombre de usuario, contraseña y clave API de 2captcha
url = "https://ejemplo.com"
username = "BE2456"
password = "3ER1ZA2019"
api_key = "8f63da7191fe11e63148c3d8b28c71f2"

# Configura el controlador de Chrome y abre la URL especificada
driver = webdriver.Chrome()
driver.get(url)

# Busca los campos de usuario y contraseña y los llena con los valores especificados
username_field = driver.find_element_by_name("username")
password_field = driver.find_element_by_name("password")
username_field.send_keys(username)
password_field.send_keys(password)

# Busca el campo de captcha y obtiene la URL de la imagen del captcha
captcha_element = driver.find_element_by_xpath("//img[@id='captcha']")
captcha_url = captcha_element.get_attribute("src")

# Usa la API de 2captcha para resolver el captcha y obtener el texto de la respuesta
solver = TwoCaptcha(api_key)
captcha_response = solver.normal(captcha_url)
captcha_text = captcha_response['code']

# Llena el campo de captcha con la respuesta obtenida
captcha_input = driver.find_element_by_name("captcha")
captcha_input.send_keys(captcha_text)

# Envía el formulario y cierra el controlador de Chrome
password_field.send_keys(Keys.RETURN)
driver.close()

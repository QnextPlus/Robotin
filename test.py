from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time

options = webdriver.ChromeOptions()
#options.add_argument('--start-maximized')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument('--disable-extensions')
options.add_experimental_option('detach', True)

driver_path = './chromedriver.exe'

driver = webdriver.Chrome(driver_path, chrome_options=options)

driver.get('https://10.95.224.27:9083/bicp/login.action')

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                      'button#details-button')))\
    .click()

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                      'a#proceed-link')))\
    .click()

WebDriverWait(driver, 5) \
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#username')))\
    .send_keys('E8221297')\
    .click()

WebDriverWait(driver, 5) \
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#password')))\
    .send_keys('3ERiza$#123')\
    .click()



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

"""


WebDriverWait(driver, 5) \
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#password')))\
    .send_keys('inspimichael@gmail.com')

WebDriverWait(driver, 5) \
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#pass')))\
    .send_keys('zaza1=')

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,
                                      '/html/body/div[1]/div[1]/div[1]/div/div/div/div[2]/div/div[1]/form/div[2]/button')))\
    .click()

#driver.quit()"""
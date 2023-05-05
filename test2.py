from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By


options = Options()
#options.add_argument("start-maximized")
options.add_experimental_option('detach', True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get("https://10.95.224.27:9083/bicp/login.action")


#driver.find_element(By.NAME, 'proceed-button')

driver.find_element(By.NAME, '#details-button').click()
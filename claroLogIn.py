#usuarios_password = {Jose, Lijhoan, Merlhy, renzo}
user_pass = {'E8234355': 'RRRjj77---', 'E8234354': 'LLLzz92---', 'E8235446': 'CLaro2023/*/', 'E8235447': 'Merlhy2023*//'}
elementos = user_pass.items()
for user,passw in elementos:
    import time
    import requests
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException
    from selenium.webdriver.common.alert import Alert
    from twocaptcha import TwoCaptcha, NetworkException
    from PIL import Image
    from io import BytesIO
    import pyautogui
    import os
    import glob
    import shutil
    import datetime
    import random
    import subprocess
    import keyboard

    # Especifica la ruta del controlador de Chrome
    pathDriver = "113/chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        #"download.default_directory": r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython\descargaRobotin",
        #"download.default_directory": defaultPathDownloads,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })


    # Para Ignorar los errores de certificado SSL (La conexion no es privada)
    options.add_argument("--ignore-certificate-errors")

    # Especifica los detalles de la URL, nombre de usuario y contraseña
    url = "http://intranetwebapp/aplicaciones/inicio.aspx"
    username = user
    password = passw

    # Configura el controlador de Chrome y abre la URL especificada
    driver = webdriver.Chrome(executable_path=pathDriver, options=options)
    driver.get(url)
    time.sleep(1)

    ###### user
    # Definir la coordenada deseada
    x = 550
    y = 200

    # Mover el mouse a la coordenada deseada
    actions = ActionChains(driver)
    # Mover el cursor a las coordenadas deseadas
    pyautogui.moveTo(x, y)

    # Escribir el nombre de usuario en la ubicación definida
    pyautogui.click()#.send_keys(username).perform()
    keyboard.write(username)
    time.sleep(1)

    ### pass
    # Definir la coordenada deseada
    x = 550
    y = 250

    # Mover el mouse a la coordenada deseada
    actions = ActionChains(driver)
    # Mover el cursor a las coordenadas deseadas
    pyautogui.moveTo(x, y)

    # Escribir el nombre de usuario en la ubicación definida
    pyautogui.click() #.send_keys(password).perform()
    keyboard.write(password)
    time.sleep(1)

    ### inicio sesion
    # Definir la coordenada deseada
    x = 500
    y = 300
    # Mover el cursor a las coordenadas deseadas
    pyautogui.moveTo(x, y)

    # Escribir el nombre de usuario en la ubicación definida
    pyautogui.click()
    time.sleep(60) # un minuto conectado
    
    # Cambiar al marco deseado
    frame = driver.find_element(By.NAME, "top")
    driver.switch_to.frame(frame)

    cerrarSesion = driver.find_element(By.CLASS_NAME,'LinkGrid')
    cerrarSesion.click()
    time.sleep(3)
    driver.quit()

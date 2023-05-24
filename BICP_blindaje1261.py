import time
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from twocaptcha import TwoCaptcha, NetworkException
from PIL import Image
from io import BytesIO
import pyautogui
from datetime import datetime, timedelta
import random

# Especifica la ruta del controlador de Chrome
pathDriver = "113/chromedriver.exe"

#Especifica Directorio de descarga
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython\blindaje\bicp\1261",
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
time.sleep(1)

#cierra ventana emergente
elemento_cierre = driver.find_element(By.XPATH, '//*[@id="tipClose"]')
elemento_cierre.click()
time.sleep(1)

#funcion para obtener el captcha
def obtenerCaptcha():
    # Busca el elemento deseado utilizando un selector CSS
    element = driver.find_element(By.ID,'validateimg')
    # Captura una imagen del elemento
    element.screenshot('capture.png')

    try:
        solver = TwoCaptcha('8f63da7191fe11e63148c3d8b28c71f2')
        id = solver.send(file='capture.png')
        time.sleep(10)
        code = solver.get_result(id)
    except NetworkException:
        code = 'error'

    return code

#funcion para inicio de sesion
def inicioSesion(usuario, contrasena):
    # Busca los campos de usuario y contraseña y los llena con los valores especificados
    campoInputUser = driver.find_element(By.ID, 'username2')
    campoInputUser.clear()
    actions = ActionChains(driver)
    actions.move_to_element(campoInputUser)
    actions.click()
    actions.send_keys(usuario)
    actions.perform()
    time.sleep(1)

    campoInputPassword = driver.find_element(By.ID, 'password2')
    campoInputPassword.clear()
    actions = ActionChains(driver)
    actions.move_to_element(campoInputPassword)
    actions.click()
    actions.send_keys(contrasena)
    actions.perform()

    validacion_logueo = True
    while validacion_logueo:
        code = obtenerCaptcha()
        while code == 'error' :
            refrescarCaptcha = driver.find_element(By.CLASS_NAME, 'validate_btn')
            refrescarCaptcha.click()
            code = obtenerCaptcha()
        else:
            pass
        code_field = driver.find_element(By.XPATH, '//*[@id="validate"]')
        code_field.send_keys(code)
        time.sleep(1)

        #presiona boton submit 
        submit_button = driver.find_element(By.CLASS_NAME, 'login_submit_btn')
        submit_button.click()
        time.sleep(5)

        nueva_url = driver.current_url
        if nueva_url == url:
            validacion_logueo = True
            code_field.clear()
        else:
            validacion_logueo = False

#funcion para validar Sesion activa
def validaSesionActiva():
#Si la sesion esta activa
    try:
        continuar = driver.find_element(By.ID,'usm_continue')
        sesionAbierta = True
    except NoSuchElementException:
        sesionAbierta = False

    if sesionAbierta:
        continuar.click()
    else:
        pass

#funcion para validar cookies
def validaSiExisteCookie():
    # para las cookies
    try:
        cookie = driver.find_element(By.CLASS_NAME, 'neterror')
        existeCookie = True
    except NoSuchElementException:
        existeCookie = False
    return existeCookie

inicioSesion(username, password)
validaSesionActiva()
hayCookie = validaSiExisteCookie()

while hayCookie:
    driver.quit()
    driver = webdriver.Chrome(executable_path=pathDriver, options=options)
    driver.get(url)
    elemento_cierre = driver.find_element(By.XPATH, '//*[@id="tipClose"]')
    elemento_cierre.click()
    time.sleep(1)

    inicioSesion(username, password)
    validaSesionActiva()
    time.sleep(1)
    hayCookie = validaSiExisteCookie()
else:
    pass

# +++++ Funcion Reporte 1261 +++++
def reporte1261(xpathBPO, xpathActivity, xpatkCampana):
    #Menu principal
    systemMenu = driver.find_element(By.ID, 'systemMenu')
    systemMenu.click()
    time.sleep(1)
    #selecciona Resource manager del menu principal
    resourceManager = driver.find_element(By.ID, 's_800')
    resourceManager.click()
    time.sleep(5)
    #Dentro del Resource Manager
    driver.switch_to.frame('tabPage_800_iframe')
    driver.switch_to.frame('container')
    lupa = driver.find_element(By.ID,'searchResource')
    lupa.click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()

    codigoCampana = 1261
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID,'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.CLASS_NAME, 'grid_link1')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    # *** MOV 1 ***
    # === Fecha D - 1 ===
    driver.switch_to.frame('view8adf609b6f1e1f27016f21e66fd10255_iframe')

    fechaActual = datetime.now().date()
    diaAnterior = fechaActual - timedelta(days=1)
    # Formato mes/dia/anio
    fechaD_1 = diaAnterior.strftime(r'%m/%d/%Y')

    textBoxFecha = driver.find_element(By.ID, 'rpt_param_Value0')
    textBoxFecha.clear()
    textBoxFecha.send_keys(fechaD_1)
    time.sleep(1)

    # === BPO ===
    btnBPO = driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
    btnBPO.click()
    time.sleep(2)

    #Selecciona BPO
    checkBoxBPO = driver.find_element(By.XPATH, xpathBPO)
    checkBoxBPO.click()
    time.sleep(2)

    BotonOk = driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(2)

    # === Activity ===
    btnActivity = driver.find_element(By.XPATH, '//*[@id="ActivityViewControl"]/table/tbody/tr/td[5]/input')
    btnActivity.click()
    time.sleep(2)

    texto = 'OUTBPOPERETEN'
    textBox = driver.find_element(By.NAME, 'startwith')
    textBox.send_keys(texto)
    time.sleep(2)

    botonSearch = driver.find_element(By.NAME, 'search')
    botonSearch.click()
    time.sleep(2)

    #Selecciona Activity
    checkBoxCampana = driver.find_element(By.XPATH, xpathActivity)
    checkBoxCampana.click()
    time.sleep(5)

    BotonOk = driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(2)

    # === Campaign ===
    btnCampaign = driver.find_element(By.XPATH, '//*[@id="CampaignViewControl"]/table/tbody/tr/td[5]/input')
    btnCampaign.click()
    time.sleep(2)

    texto = 'OUTBPOPERETEN'
    textBox = driver.find_element(By.NAME, 'startwith')
    textBox.send_keys(texto)
    time.sleep(2)

    botonSearch = driver.find_element(By.NAME, 'search')
    botonSearch.click()
    time.sleep(2)

    #Selecciona Campaign
    checkBoxCampana = driver.find_element(By.XPATH, xpatkCampana)
    checkBoxCampana.click()
    time.sleep(2)

    BotonOk = driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(2)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(10)

    #Descarga
    btnDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    btnDowloand.click()
    time.sleep(1)

    descargarExcel = driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
    descargarExcel.click()
    time.sleep(1)

    excel2013 = driver.find_element(By.ID, 'downExcel2007Id')
    excel2013.click()
    time.sleep(1)

    confirmacionFinal = driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
    confirmacionFinal.click()
    time.sleep(15)
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(By.ID, 'view8adf609b6f1e1f27016f21e66fd10255_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)
# +++++ Fin Reporte 1261 +++++

# === Reporte 1261 ===
#Blindaje
    #    ---- OUTBPOPERETENMOVSF01 ----
labelBPO = 'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="radio" and @label="{labelBPO}"]'

labelAct = 'OUTBPOPERETENMOVSF01'
xpathActivity = f'//input[@type="radio" and @label="{labelAct}"]'

labelC = labelAct
xpatkCampana = f'//input[@type="radio" and @label="{labelC}"]'
reporte1261(xpathBPO, xpathActivity, xpatkCampana)

    #    ---- OUTBPOPERETENMOVSF02 ----
labelBPO = 'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="radio" and @label="{labelBPO}"]'

labelAct = 'OUTBPOPERETENMOVSF02'
xpathActivity = f'//input[@type="radio" and @label="{labelAct}"]'

labelC = labelAct
xpatkCampana = f'//input[@type="radio" and @label="{labelC}"]'
reporte1261(xpathBPO, xpathActivity, xpatkCampana)

    #    ---- OUTBPOPERETENBITSF03 ----
labelBPO = 'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="radio" and @label="{labelBPO}"]'

labelAct = 'OUTBPOPERETENBITSF03'
xpathActivity = f'//input[@type="radio" and @label="{labelAct}"]'

labelC = labelAct
xpatkCampana = f'//input[@type="radio" and @label="{labelC}"]'
reporte1261(xpathBPO, xpathActivity, xpatkCampana)

# === Fin Reportes 1261 ===

#Cierra la aplicacion
fun_logout = driver.find_element(By.ID, 'fun_logout')
fun_logout.click()
time.sleep(1)

confirmaCierreSesion = driver.find_element(By.XPATH, '//*[@id="winmsg0"]/div/div[2]/div[3]/div[3]/span[1]/div/div')
confirmaCierreSesion.click()
time.sleep(3)

driver.quit()

print('Descarga Exitosa')

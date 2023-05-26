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
import os
import glob

directoryPath = os.getcwd()
defaultPathDownloads = directoryPath + r'\temp'
# Especifica la ruta del controlador de Chrome
pathDriver = "113/chromedriver.exe"

#Especifica Directorio de descarga
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
    "download.default_directory": defaultPathDownloads,
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

#Funcion cuenta cantidad de archivos xlsx en carpeta de descarga
def cantidadExcel():
    ruta_carpeta = defaultPathDownloads
    extension = '*.xlsx'
    patron_busqueda = os.path.join(ruta_carpeta, extension)
    archivos = glob.glob(patron_busqueda)
    cantidad_archivos = len(archivos)
    return cantidad_archivos

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

# +++++ Funcion Reporte 26 +++++
def reporte26(xpathCampana, xpathBPO):
    #Menu principal
    systemMenu = driver.find_element(By.ID, 'systemMenu')
    systemMenu.click()

    time.sleep(1)
    #selecciona Resource manager del menu principal
    resourceManager = driver.find_element(By.ID, 's_800')
    resourceManager.click()
    time.sleep(5)

    driver.switch_to.frame('tabPage_800_iframe')
    driver.switch_to.frame('container')
    lupa = driver.find_element(By.ID,'searchResource')
    lupa.click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()

    codigoCampana = 26
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

    #dentro del input resource
    driver.switch_to.frame('view8adf609c511ba97601511ba991d20056_iframe')

    # === Campaign Name ===
    comboBoxCN = driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId2"]/div/div')
    comboBoxCN.click()
    time.sleep(2)

    #desmarca selecion Default All
    checkBoxDefault = driver.find_element(By.XPATH, '/html/body/div[13]/ul/li[1]/span')
    checkBoxDefault.click()
    time.sleep(1)

    #Selecciona Campaign Name
    checkBoxCampana = driver.find_element(By.XPATH, xpathCampana)
    checkBoxCampana.click()
    time.sleep(1)

    comboBoxCN.click()
    time.sleep(2)

    # === BPO ====
    comboBoxBPO = driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId3"]/div/div')
    comboBoxBPO.click()
    time.sleep(2)

    #desmarca selecion Default All
    checkBoxDefault = driver.find_element(By.XPATH, '/html/body/div[14]/ul/li[1]/span')
    checkBoxDefault.click()
    time.sleep(1)

    #Selecciona BPO
    checkBoxCampana = driver.find_element(By.XPATH, xpathBPO)
    checkBoxCampana.click()
    time.sleep(1)

    comboBoxBPO.click()
    time.sleep(1)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    #time.sleep(10)
    # Esta funcion valida si existen datos para exportar
    def existeDatos():
        try:
            continua = driver.find_element(By.CLASS_NAME,'pageControl')
            tabla = True
        except NoSuchElementException:
            tabla = False
        return tabla
    
    #Espera a que se carguen los datos
    puedeDescargar = False
    while puedeDescargar:
        pass
    else:
        time.sleep(1)
        puedeDescargar = existeDatos()
    #cuenta la cantidad de xlsx que hay en la carpeta de descarga antes de descargar el reporte
    cantidadExcelinicial = cantidadExcel()

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
    #time.sleep(15)
    #Valida que la descarga concluya
    cantidadExcelFinal = cantidadExcelinicial
    while cantidadExcelFinal == cantidadExcelinicial:
        time.sleep(1)
        cantidadExcelFinal = cantidadExcel()
    else:
        pass

    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(By.ID, 'view8adf609c511ba97601511ba991d20056_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)
# +++++ Fin funcon Reporte 26 +++++

#descarga reporte 26
titleC = 'RETENCIONESMOVILES'
xpathCampana = f'//li[@title="{titleC}"]/span[@class="icon"]'

titleBPO = 'BPOPERURETENCION'
xpathBPO = f'//li[@title="{titleBPO}"]/span[@class="icon"]'
reporte26(xpathCampana ,xpathBPO)

#Cierra la aplicacion
fun_logout = driver.find_element(By.ID, 'fun_logout')
fun_logout.click()
time.sleep(1)

confirmaCierreSesion = driver.find_element(By.XPATH, '//*[@id="winmsg0"]/div/div[2]/div[3]/div[3]/span[1]/div/div')
confirmaCierreSesion.click()
time.sleep(3)

driver.quit()

print('Descarga Exitosa')

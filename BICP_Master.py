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
from datetime import datetime, timedelta
import random
import subprocess

directoryPath = os.getcwd()

# Funcion que reubicará las descargas en sus respectivas carpetas
def renombrarReubicar(nuevoNombre, carpetaDestino):
    # ruta_descargas = r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython\descargaRobotin"
    ruta_descargas = directoryPath + r'/temp'
    archivos_descargados = sorted(
        glob.glob(os.path.join(ruta_descargas, '*')), key=os.path.getmtime, reverse=True
    )
    # Comprobar si hay archivos descargados
    if len(archivos_descargados) > 0:
        ultimo_archivo = archivos_descargados[0]
        # Cambiar el nombre del archivo --1er argumento de la funcion
        nuevo_nombre = f'{nuevoNombre}.xlsx'
        carpeta_destino = carpetaDestino
        # Comprobar si la carpeta de destino existe, si no, crearla
        if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino)
        # Ruta completa del archivo de destino
        ruta_destino = os.path.join(carpeta_destino, nuevo_nombre)
        # Mover el archivo a la carpeta de destino con el nuevo nombre
        shutil.move(ultimo_archivo, ruta_destino)

# Funcion que crea el nombre del reporte
def nombreReporte(name):
    fechaHora = datetime.now()
    fecha = fechaHora.strftime("%Y%m%d_%H%M%S")
    aleatorio = str(random.randint(100, 999))
    nameFile = name + fecha + '_' + aleatorio
    print(name + fecha + '_' + aleatorio)
    return nameFile

# Especifica la ruta del controlador de Chrome
pathDriver = "113/chromedriver.exe"

#Directorio de descarga
defaultPathDownloads = directoryPath + r'\temp'

#Funcion cuenta cantidad de archivos xlsx en carpeta de descarga
def cantidadExcel():
    ruta_carpeta = defaultPathDownloads
    extension = '*.xlsx'
    patron_busqueda = os.path.join(ruta_carpeta, extension)
    archivos = glob.glob(patron_busqueda)
    cantidad_archivos = len(archivos)
    return cantidad_archivos

# Especifica Directorio de descarga
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
    # "download.default_directory": r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython\descargaRobotin",
    "download.default_directory": defaultPathDownloads,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})


# Para Ignorar los errores de certificado SSL (La conexion no es privada)
options.add_argument("--ignore-certificate-errors")

# Especifica los detalles de la URL, nombre de usuario y contraseña
url = "https://10.95.224.27:9083/bicp/login.action"
username = "E8234353"
password = "JJJmm88***1"
username2 = 'E8234180'

# Configura el controlador de Chrome y abre la URL especificada
driver = webdriver.Chrome(executable_path=pathDriver, options=options)
driver.get(url)
time.sleep(1)

# cierra ventana emergente
elemento_cierre = driver.find_element(By.XPATH, '//*[@id="tipClose"]')
elemento_cierre.click()
time.sleep(1)

# funcion para obtener el captcha
def obtenerCaptcha():
    # Busca el elemento deseado utilizando un selector CSS
    element = driver.find_element(By.ID, 'validateimg')
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

# funcion para inicio de sesion
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
        while code == 'error':
            refrescarCaptcha = driver.find_element(
                By.CLASS_NAME, 'validate_btn')
            refrescarCaptcha.click()
            code = obtenerCaptcha()
        else:
            pass
        code_field = driver.find_element(By.XPATH, '//*[@id="validate"]')
        code_field.send_keys(code)
        time.sleep(1)

        # presiona boton submit
        submit_button = driver.find_element(By.CLASS_NAME, 'login_submit_btn')
        submit_button.click()
        time.sleep(5)

        nueva_url = driver.current_url
        if nueva_url == url:
            validacion_logueo = True
            code_field.clear()
        else:
            validacion_logueo = False

# funcion para validar Sesion activa
def validaSesionActiva():
    # Si la sesion esta activa
    try:
        continuar = driver.find_element(By.ID, 'usm_continue')
        sesionAbierta = True
    except NoSuchElementException:
        sesionAbierta = False

    if sesionAbierta:
        continuar.click()
    else:
        pass

# funcion para validar cookies
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

# +++++ 1. Funcion para descargar Reporte 1259 +++++
def reporte1259(xpathCamapana):
    # Menu principal
    systemMenu = driver.find_element(By.ID, 'systemMenu')
    systemMenu.click()
    time.sleep(1)

    # selecciona Resource manager del menu principal
    resourceManager = driver.find_element(By.ID, 's_800')
    resourceManager.click()
    time.sleep(5)

    # Dentro de la seccion de Resource Manager
    driver.switch_to.frame('tabPage_800_iframe')
    driver.switch_to.frame('container')
    lupa = driver.find_element(By.ID, 'searchResource')
    lupa.click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()

    # Dentro de la seccion Search Resource
    # Completa el TextBox con el codigo de la campaña
    codigoCampana = 1259
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID, 'searchCondtion_input_value')
    textBox.clear()
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="1259"]')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    # Seccion de Input de parámetros
    driver.switch_to.frame('view8adf609b6f1e1f27016f21e66d51024c_iframe')
    # === BPO ===
    btnBPO = driver.find_element(
        By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
    btnBPO.click()
    time.sleep(2)

    # Desmarca la seleccion por default
    checkBoxDefault = driver.find_element(
        By.CSS_SELECTOR, 'input[type="checkbox"][value="9999"]')
    checkBoxDefault.click()
    time.sleep(1)
    # Marca el checbox de la campaña desseada -- 1er argumento funcion
    checkBoxCampana = driver.find_element(By.XPATH, xpathCamapana)
    checkBoxCampana.click()
    time.sleep(1)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(1)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
        
    #Descargar
    iconoDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    iconoDowloand.click()
    time.sleep(1)
    descargarExcel = driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
    descargarExcel.click()
    time.sleep(1)

    excel2013 = driver.find_element(By.ID, 'downExcel2007Id')
    excel2013.click()
    time.sleep(1)

    confirmacionFinal = driver.find_element(
        By.ID, 'btnOk_downExcel2003Or2007WinId')
    confirmacionFinal.click()
    #time.sleep(15)

    #Valida que la descarga concluya
    cantidadExcelFinal = cantidadExcelinicial
    while cantidadExcelFinal == cantidadExcelinicial:
        time.sleep(1)
        cantidadExcelFinal = cantidadExcel()
    else:
        pass

    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(
        By.ID, 'view8adf609b6f1e1f27016f21e66d51024c_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""
# +++++ Fin Funcion Reporte 1259 +++++

# +++++ 2. Funcion para descargar Reporte 401 +++++
def reporte401(xpathBPO, txtCampana, xpathCampana, xpathAgentWorkgroup='noDefinido'):
    # Menu Principal
    systemMenu = driver.find_element(By.ID, 'systemMenu')
    systemMenu.click()
    time.sleep(1)

    # selecciona Resource manager del menu principal
    resourceManager = driver.find_element(By.ID, 's_800')
    resourceManager.click()
    time.sleep(5)

    # Dentro del Resource Manager
    driver.switch_to.frame('tabPage_800_iframe')
    driver.switch_to.frame('container')
    lupa = driver.find_element(By.ID, 'searchResource')
    lupa.click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()

    # Dentro de la seccion Search Resource
    # Completa el TextBox con el codigo de la campaña
    codigoCampana = 401
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID, 'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="401"]')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    # Seccion de Input de parámetros
    driver.switch_to.frame('view8adf609c6387afda01638b7f47b90570_iframe')

    # === BPO ===
    btnBPO = driver.find_element(
        By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
    btnBPO.click()
    time.sleep(2)

    # desmarca selecion Default All
    checkBoxDefault = driver.find_element(
        By.CSS_SELECTOR, 'input[type="checkbox"][value="9999"]')
    checkBoxDefault.click()
    time.sleep(1)

    # Selecciona BPO
    checkBoxCampana = driver.find_element(By.XPATH, xpathBPO)
    checkBoxCampana.click()
    time.sleep(1)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(1)

    # === Campaign ===
    btnCampana = driver.find_element(
        By.XPATH, '//*[@id="CampaignViewControl"]/table/tbody/tr/td[5]/input')
    btnCampana.click()
    time.sleep(2)

    # desmarca selecion Default All
    checkBoxDefault = driver.find_element(
        By.CSS_SELECTOR, 'input[type="checkbox"][value="9999"]')
    checkBoxDefault.click()
    time.sleep(1)

    # ingresar texto busqueda en textbox
    texto = txtCampana
    textBox = driver.find_element(By.NAME, 'startwith')
    textBox.send_keys(texto)
    time.sleep(1)

    botonSearch = driver.find_element(By.NAME, 'search')
    botonSearch.click()
    time.sleep(2)

    # Selecciona Campaign
    checkBoxCampana = driver.find_element(By.XPATH, xpathCampana)
    checkBoxCampana.click()
    time.sleep(1)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(1)

    # === Agent Workgroup ===
    agentWkg = xpathAgentWorkgroup
    if agentWkg != 'noDefinido':
        btnAgentWorkgroup = driver.find_element(
            By.XPATH, '//*[@id="Agent WorkgroupViewControl"]/table/tbody/tr/td[5]/input')
        btnAgentWorkgroup.click()
        time.sleep(2)

        checkBox1034 = driver.find_element(By.XPATH, agentWkg)
        checkBox1034.click()
        time.sleep(1)

        BotonOk = driver.find_element(
            By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)
    else:
        pass

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

    # Descarga
    btnDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    btnDowloand.click()
    time.sleep(1)

    descargarExcel = driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
    descargarExcel.click()
    time.sleep(1)

    excel2013 = driver.find_element(By.ID, 'downExcel2007Id')
    excel2013.click()
    time.sleep(1)

    confirmacionFinal = driver.find_element(
        By.ID, 'btnOk_downExcel2003Or2007WinId')
    confirmacionFinal.click()
    #time.sleep(15)

    #Valida que la descarga concluya
    cantidadExcelFinal = cantidadExcelinicial
    while cantidadExcelFinal == cantidadExcelinicial:
        time.sleep(1)
        cantidadExcelFinal = cantidadExcel()
    else:
        pass

    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(
        By.ID, 'view8adf609c6387afda01638b7f47b90570_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""

# +++++ Fin Funcion Reporte 401 +++++

# +++++ 3. Funcion para descargar Reporte 112 +++++
def reporte112(xpathBPO):
    # Menu principal
    systemMenu = driver.find_element(By.ID, 'systemMenu')
    systemMenu.click()

    time.sleep(1)
    # selecciona Resource manager del menu principal
    resourceManager = driver.find_element(By.ID, 's_800')
    resourceManager.click()
    time.sleep(5)

    driver.switch_to.frame('tabPage_800_iframe')
    driver.switch_to.frame('container')
    lupa = driver.find_element(By.ID, 'searchResource')
    lupa.click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()

    codigoCampana = 112
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID, 'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="112"]')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    # === BPO ===
    # Seleccionar campaña
    driver.switch_to.frame('view8adf609c54f300b50154f46383e501a8_iframe')
    btnCampana = driver.find_element(
        By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
    btnCampana.click()
    time.sleep(2)

    # desmarca toda la selecion por Default SelectAll
    checkBoxselectAll = driver.find_element(
        By.CSS_SELECTOR, 'input[type="checkbox"][value="selectAll"]')
    checkBoxselectAll.click()
    checkBoxselectAll.click()
    time.sleep(1)

    # Selecciona BPO
    checkBoxBPO = driver.find_element(By.XPATH, xpathBPO)
    checkBoxBPO.click()
    time.sleep(1)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(1)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

    # Descarga
    btnDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    btnDowloand.click()
    time.sleep(1)

    descargarExcel = driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
    descargarExcel.click()
    time.sleep(1)

    excel2013 = driver.find_element(By.ID, 'downExcel2007Id')
    excel2013.click()
    time.sleep(1)

    confirmacionFinal = driver.find_element(
        By.ID, 'btnOk_downExcel2003Or2007WinId')
    confirmacionFinal.click()
    #time.sleep(15)

    #Valida que la descarga concluya
    cantidadExcelFinal = cantidadExcelinicial
    while cantidadExcelFinal == cantidadExcelinicial:
        time.sleep(1)
        cantidadExcelFinal = cantidadExcel()
    else:
        pass
    
    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(
        By.ID, 'view8adf609c54f300b50154f46383e501a8_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""
# +++++ Fin funcion reporte 112 +++++

# +++++ 4. Funcion Reporte 1261 +++++
def reporte1261(xpathBPO, xpathActivity, xpatkCampana):
    # Menu principal
    systemMenu = driver.find_element(By.ID, 'systemMenu')
    systemMenu.click()
    time.sleep(1)
    # selecciona Resource manager del menu principal
    resourceManager = driver.find_element(By.ID, 's_800')
    resourceManager.click()
    time.sleep(5)
    # Dentro del Resource Manager
    driver.switch_to.frame('tabPage_800_iframe')
    driver.switch_to.frame('container')
    lupa = driver.find_element(By.ID, 'searchResource')
    lupa.click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()

    codigoCampana = 1261
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID, 'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="1261"]')
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
    btnBPO = driver.find_element(
        By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
    btnBPO.click()
    time.sleep(2)

    # Selecciona BPO
    checkBoxBPO = driver.find_element(By.XPATH, xpathBPO)
    checkBoxBPO.click()
    time.sleep(2)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(2)

    # === Activity ===
    btnActivity = driver.find_element(
        By.XPATH, '//*[@id="ActivityViewControl"]/table/tbody/tr/td[5]/input')
    btnActivity.click()
    time.sleep(2)

    texto = 'OUTBPOPERETEN'
    textBox = driver.find_element(By.NAME, 'startwith')
    textBox.send_keys(texto)
    time.sleep(2)

    botonSearch = driver.find_element(By.NAME, 'search')
    botonSearch.click()
    time.sleep(2)

    # Selecciona Activity
    checkBoxCampana = driver.find_element(By.XPATH, xpathActivity)
    checkBoxCampana.click()
    time.sleep(5)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(2)

    # === Campaign ===
    btnCampaign = driver.find_element(
        By.XPATH, '//*[@id="CampaignViewControl"]/table/tbody/tr/td[5]/input')
    btnCampaign.click()
    time.sleep(2)

    texto = 'OUTBPOPERETEN'
    textBox = driver.find_element(By.NAME, 'startwith')
    textBox.send_keys(texto)
    time.sleep(2)

    botonSearch = driver.find_element(By.NAME, 'search')
    botonSearch.click()
    time.sleep(2)

    # Selecciona Campaign
    checkBoxCampana = driver.find_element(By.XPATH, xpatkCampana)
    checkBoxCampana.click()
    time.sleep(2)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(2)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

    # Descarga
    btnDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    btnDowloand.click()
    time.sleep(1)

    descargarExcel = driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
    descargarExcel.click()
    time.sleep(1)

    excel2013 = driver.find_element(By.ID, 'downExcel2007Id')
    excel2013.click()
    time.sleep(1)

    confirmacionFinal = driver.find_element(
        By.ID, 'btnOk_downExcel2003Or2007WinId')
    confirmacionFinal.click()
    #time.sleep(15)

    #Valida que la descarga concluya
    cantidadExcelFinal = cantidadExcelinicial
    while cantidadExcelFinal == cantidadExcelinicial:
        time.sleep(1)
        cantidadExcelFinal = cantidadExcel()
    else:
        pass 
    
    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(
        By.ID, 'view8adf609b6f1e1f27016f21e66fd10255_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""
# +++++ Fin Reporte 1261 +++++

# +++++ 5. Funcion Reporte 90 +++++
def reporte90(xpathCampana, xpathBPO, xpathAgentGroup):
    # Menu principal
    systemMenu = driver.find_element(By.ID, 'systemMenu')
    systemMenu.click()
    time.sleep(1)
    # selecciona Resource manager del menu principal
    resourceManager = driver.find_element(By.ID, 's_800')
    resourceManager.click()
    time.sleep(5)
    # Dentro del Resource Manager
    driver.switch_to.frame('tabPage_800_iframe')
    driver.switch_to.frame('container')
    lupa = driver.find_element(By.ID, 'searchResource')
    lupa.click()
    time.sleep(1)
    driver.switch_to.default_content()
    driver.switch_to.default_content()

    codigoCampana = 90
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID, 'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="90"]')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    # Dentro del Input source
    driver.switch_to.frame('view8adf649b5237e8b7015352f0feb700d9_iframe')

    # === Campaign ===
    btnCampaign = driver.find_element(
        By.XPATH, '//*[@id="rpt_param_ComboxId2"]/div/div')
    btnCampaign.click()
    time.sleep(2)

    # Selecciona Campaign
    seleccionDefault = driver.find_element(
        By.XPATH, '/html/body/div[13]/ul/li[1]/span')
    seleccionDefault.click()
    time.sleep(1)
    seleccionDefault.click()
    time.sleep(1)

    checkBoxCampana = driver.find_element(By.XPATH, xpathCampana)
    checkBoxCampana.click()
    time.sleep(2)

    btnCampaign.click()
    time.sleep(2)

    # === BPO ===
    btnBPO = driver.find_element(
        By.XPATH, '//*[@id="rpt_param_ComboxId3"]/div/div')
    btnBPO.click()
    time.sleep(2)

    # Selecciona BPO
    seleccionDefault = driver.find_element(
        By.XPATH, '/html/body/div[14]/ul/li[1]/span')
    seleccionDefault.click()
    time.sleep(1)

    checkBoxBPO = driver.find_element(By.XPATH, xpathBPO)
    checkBoxBPO.click()
    time.sleep(2)

    btnBPO.click()
    time.sleep(2)

    # === Agent Group ===
    btnActivity = driver.find_element(
        By.XPATH, '//*[@id="AgentWorkGroupViewControl"]/table/tbody/tr/td[5]/input')
    btnActivity.click()
    time.sleep(2)

    # Selecciona Agent Group
    checkBoxCampana = driver.find_element(By.XPATH, xpathAgentGroup)
    checkBoxCampana.click()
    time.sleep(5)

    BotonOk = driver.find_element(
        By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
    BotonOk.click()
    time.sleep(2)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

    # Descarga
    btnDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    btnDowloand.click()
    time.sleep(1)

    descargarExcel = driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
    descargarExcel.click()
    time.sleep(1)

    excel2013 = driver.find_element(By.ID, 'downExcel2007Id')
    excel2013.click()
    time.sleep(1)

    confirmacionFinal = driver.find_element(
        By.ID, 'btnOk_downExcel2003Or2007WinId')
    confirmacionFinal.click()
    #time.sleep(15)

    #Valida que la descarga concluya
    cantidadExcelFinal = cantidadExcelinicial
    while cantidadExcelFinal == cantidadExcelinicial:
        time.sleep(1)
        cantidadExcelFinal = cantidadExcel()
    else:
        pass
    
    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(
        By.ID, 'view8adf649b5237e8b7015352f0feb700d9_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""
# +++++ Fin Reporte 90 +++++

# +++++ 6. Funcion Reporte 43 +++++
def reporte43(xpathBPO):
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

    codigoCampana = 43
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID,'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="43"]')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    #dentro del input resource
    driver.switch_to.frame('view8adf609c511ba97601511ba994360060_iframe')
    
    # === BPO ====
    comboBox = driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId2"]/div/div')
    comboBox.click()
    time.sleep(2)

    #desmarca selecion Default All
    checkBoxDefault = driver.find_element(By.XPATH, '/html/body/div[13]/ul/li[1]/span')
    checkBoxDefault.click()
    time.sleep(1)

    #Selecciona BPO
    checkBoxCampana = driver.find_element(By.XPATH, xpathBPO)
    checkBoxCampana.click()
    time.sleep(1)

    comboBox.click()
    time.sleep(1)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

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

    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(By.ID, 'view8adf609c511ba97601511ba994360060_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""

# +++++ Fin funcion Reporte 43 +++++

# +++++ 7. Funcion Reporte 26 +++++
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

    reporte = driver.find_element(By.XPATH, '//a[@title="26"]')
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
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))    

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
    
    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(By.ID, 'view8adf609c511ba97601511ba991d20056_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""
# +++++ Fin funcon Reporte 26 +++++

# +++++ 8. Funcion Reporte 194 +++++
def reporte194(xpathBPO):
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

    codigoCampana = 194
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID,'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="194"]')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    #dentro del input resource
    driver.switch_to.frame('view8adf609c5becac34015bf11be608063a_iframe')

    # === BPO ====
    comboBoxBPO = driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[3]/div/span/button/span[1]')
    comboBoxBPO.click()
    time.sleep(2)

    #Selecciona BPO
    checkBoxCampana = driver.find_element(By.XPATH, xpathBPO)
    checkBoxCampana.click()
    time.sleep(1)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
    
    #Descarga
    iconoDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    iconoDowloand.click()
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
    
    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(By.ID, 'view8adf609c5becac34015bf11be608063a_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""
# +++++ Fin funcon Reporte 194 +++++

# +++++ 9. Funcion Reporte 192 +++++
def reporte192(xpathBPO, xpathAgentGroup):
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

    codigoCampana = 192
    driver.switch_to.frame('searchResource_iframe')
    textBox = driver.find_element(By.ID,'searchCondtion_input_value')
    textBox.send_keys(codigoCampana)
    time.sleep(1)

    searchBtn = driver.find_element(By.ID, 'searchBtn')
    searchBtn.click()
    time.sleep(3)

    reporte = driver.find_element(By.XPATH, '//a[@title="194"]')
    reporte.click()
    driver.switch_to.default_content()
    time.sleep(5)

    #dentro del input resource
    driver.switch_to.frame('view8adf609c5becac34015bf11be4bf0631_iframe')

    # === BPO ====
    comboBoxBPO = driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[3]/div/span/button/span[1]')
    comboBoxBPO.click()
    time.sleep(2)

    #Selecciona BPO
    checkBoxCampana = driver.find_element(By.XPATH, xpathBPO)
    checkBoxCampana.click()
    time.sleep(1)

    # === Agent Group ===
    comboAgentGroup = driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId4"]/div/div')
    comboAgentGroup.click()
    time.sleep(2)

    CheckboxDefault = driver.find_element(By.XPATH, '/html/body/div[13]/ul/li[1]/span')
    CheckboxDefault.click()
    time.sleep(1)

    checkboxAgentGroup = driver.find_element(By.XPATH, xpathAgentGroup)
    checkboxAgentGroup.click()
    time.sleep(1)

    comboAgentGroup.click()
    time.sleep(1)

    BotonOk2 = driver.find_element(By.ID, 'rpt_param_OkBtn')
    BotonOk2.click()
    time.sleep(1)

    cantidadExcelinicial = cantidadExcel()
    # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
    WebDriverWait(driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
    
    #Descarga
    iconoDowloand = driver.find_element(By.CLASS_NAME, 'ico_download')
    iconoDowloand.click()
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

    driver.refresh()
    time.sleep(1)
    """
    driver.switch_to.default_content()

    cerrarInputParametros = driver.find_element(By.ID, 'view8adf609c5becac34015bf11be4bf0631_close')
    cerrarInputParametros.click()
    time.sleep(2)

    cerrarSearchResourse = driver.find_element(By.ID, 'searchResource_close')
    cerrarSearchResourse.click()
    time.sleep(2)

    cerrarResourseManager = driver.find_element(By.ID, 'tabPage_800_close')
    cerrarResourseManager.click()
    time.sleep(2)"""
# +++++ Fin funcon Reporte 194 +++++


# Descarga de reportes
# === Reportes 1259 ===
# Blindaje
labelC = f'OUTBPOPERETENSF'
reporte1259(f'//input[@type="checkbox" and @label="{labelC}"]')
nombreAsignado = 'blindaje_bicp_1259_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\blindaje\bicp\1259'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/blindaje/bicp/1259/' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])

# Recupero Inbound
labelC = f'BPOPERETENCIONOUT'
reporte1259(f'//input[@type="checkbox" and @label="{labelC}"]')
nombreAsignado = 'RecuperoInbound_bicp_1259_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\recuperoInbound\bicp\1259'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/recuperoInbound/bicp/1259' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])

# Contactados
labelC = f'OUTBPOPEPREVENPOSTREDSF'
reporte1259(f'//input[@type="checkbox" and @label="{labelC}"]')
nombreAsignado = 'contactado_bicp_1259_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\contactado\bicp\1259'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/contactado/bicp/1259' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])

# Retenciones Inbound
# Pendiente

# === Fin Reporte 1259 ===

# === Reporte 401 ===
# Blindaje
labelBPO = f'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="checkbox" and @label="{labelBPO}"]'

txtCampana = 'OUTBPOPERETEN'
labelC = 'OUTBPOPERETENMOVSF01'
xpathCampana = f'//input[@type="checkbox" and @label="{labelC}"]'

labelAG = '1034^OUTBPOPERETENSF'
xpathAgentWorkgroup = f'//input[@type="checkbox" and @label="{labelAG}"]'

reporte401(xpathBPO, txtCampana, xpathCampana, xpathAgentWorkgroup)
nombreAsignado = 'blindaje_bicp_401_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\blindaje\bicp\401'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/blindaje/bicp/401' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])

# Recupero Inbound
labelBPO = f'BPOPERETENCIONOUT'
xpathBPO = f'//input[@type="checkbox" and @label="{labelBPO}"]'

txtCampana = 'BPOPERETENCIONOUT'
valueC = 'selectAll'
xpathCampana = f'//input[@type="checkbox" and @value="{valueC}"]'

reporte401(xpathBPO, txtCampana, xpathCampana)
nombreAsignado = 'RecuperoInbound_bicp_401_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\recuperoInbound\bicp\401'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/recuperoInbound/bicp/401' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])

# Contactado
labelBPO = f'OUTBPOPEPREVENPOSTREDSF'
xpathBPO = f'//input[@type="checkbox" and @label="{labelBPO}"]'

txtCampana = 'OUTBPOPEPREVENPOSTREDMOVSF'
valueC = 'selectAll'
xpathCampana = f'//input[@type="checkbox" and @value="{valueC}"]'

reporte401(xpathBPO, txtCampana, xpathCampana)
nombreAsignado = 'contactado_bicp_401_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\contactado\bicp\401'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/contactado/bicp/401' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])

# === Fin Reportes 401 ===

# === Reportes 112 ===
# Blindaje
labelC = f'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="checkbox" and @label="{labelC}"]'
reporte112(xpathBPO)
nombreAsignado = 'blindaje_bicp_112_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\blindaje\bicp\112'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/blindaje/bicp/112' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


# Contactado
labelC = f'OUTBPOPEPREVENPOSTREDSF'
xpathBPO = f'//input[@type="checkbox" and @label="{labelC}"]'
reporte112(xpathBPO)
nombreAsignado = 'contactado_bicp_112_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\contactado\bicp\112'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/contactado/bicp/112' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


# Retenciones Inbound
labelC = f'BPOPERURETENCION'
xpathBPO = f'//input[@type="checkbox" and @label="{labelC}"]'
reporte112(xpathBPO)
nombreAsignado = 'RetencionesInbound_bicp_112_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\retencionesInbound\bicp\112'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/retencionesInbound/bicp/112' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


# === Fin Reportes 112 ===

# === Descarga Reportes 1261 ===
# 1. Blindaje
#    ---- OUTBPOPERETENMOVSF01 ----
labelBPO = 'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="radio" and @label="{labelBPO}"]'

labelAct = 'OUTBPOPERETENMOVSF01'
xpathActivity = f'//input[@type="radio" and @label="{labelAct}"]'

labelC = labelAct
xpatkCampana = f'//input[@type="radio" and @label="{labelC}"]'
reporte1261(xpathBPO, xpathActivity, xpatkCampana)
nombreAsignado = 'blindajeMov01_bicp_1261_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\blindaje\BICP\1261'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/blindaje/BICP/1261' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


#    ---- OUTBPOPERETENMOVSF02 ----
labelBPO = 'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="radio" and @label="{labelBPO}"]'

labelAct = 'OUTBPOPERETENMOVSF02'
xpathActivity = f'//input[@type="radio" and @label="{labelAct}"]'

labelC = labelAct
xpatkCampana = f'//input[@type="radio" and @label="{labelC}"]'
reporte1261(xpathBPO, xpathActivity, xpatkCampana)
nombreAsignado = 'blindajeMov02_bicp_1261_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\blindaje\BICP\1261'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/blindaje/BICP/1261' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


#    ---- OUTBPOPERETENBITSF03 ----
labelBPO = 'OUTBPOPERETENSF'
xpathBPO = f'//input[@type="radio" and @label="{labelBPO}"]'

labelAct = 'OUTBPOPERETENBITSF03'
xpathActivity = f'//input[@type="radio" and @label="{labelAct}"]'

labelC = labelAct
xpatkCampana = f'//input[@type="radio" and @label="{labelC}"]'
reporte1261(xpathBPO, xpathActivity, xpatkCampana)
nombreAsignado = 'blindajeBit03_bicp_1261_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\blindaje\BICP\1261'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/blindaje/BICP/1261' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


# === Fin Reportes 1261 ===

# === Reporte 90 ===
# Retenciones Inbound
titleC = 'RETENCIONESMOVILES'
xpatkCampana = f'//li[@title="{titleC}"]'

titleBPO = 'BPOPERURETENCION'
xpathBPO = f'//li[@title="{titleBPO}"]'

labelAG = '1070^BPOPERURETENCION'
xpathAgentG = f'//input[@type="checkbox" and @label="{labelAG}"]'

reporte90(xpatkCampana, xpathBPO, xpathAgentG)
nombreAsignado = 'retencionesInbound_bicp_90_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\retencionesInbound\BICP\90'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/retencionesInbound/BICP/90' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


# === Fin Reportes 90 ===

# === Reportes 43 ===
#Retenciones Inbound
titleBPO = 'BPOPERURETENCION'
xpathBPO = f'//li[@title="{titleBPO}"]'
reporte43(xpathBPO)
nombreAsignado = 'retencionesInbound_bicp_43_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\retencionesInbound\BICP\43'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/retencionesInbound/BICP/43' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])


# === Fin Reportes 43 ===

# === Reporte 26 ===
#Retenciones Inbound
titleC = 'RETENCIONESMOVILES'
xpathCampana = f'//li[@title="{titleC}"]/span[@class="icon"]'

titleBPO = 'BPOPERURETENCION'
xpathBPO = f'//li[@title="{titleBPO}"]/span[@class="icon"]'
reporte26(xpathCampana ,xpathBPO)
nombreAsignado = 'retencionesInbound_bicp_26_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\retencionesInbound\BICP\26'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/retencionesInbound/BICP/26' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])

# === Fin Reportes 26 ===

# Cierra la aplicacion
def cerrarSesion():
    fun_logout = driver.find_element(By.ID, 'fun_logout')
    fun_logout.click()
    time.sleep(1)

    confirmaCierreSesion = driver.find_element(
        By.XPATH, '//*[@id="winmsg0"]/div/div[2]/div[3]/div[3]/span[1]/div/div')
    confirmaCierreSesion.click()
    time.sleep(3)

cerrarSesion()

# ********** Inicia sesion Username2 ***********
# cierra ventana emergente
elemento_cierre = driver.find_element(By.XPATH, '//*[@id="tipClose"]')
elemento_cierre.click()
time.sleep(1)

inicioSesion(username2, password)
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

#Descarga Reportes 

# === Reporte 194 ===
# RETENCIONES INBOUND
titleBPO = "BPOPERURETENCION"
xpathBPO = f"//li[@class='ui-menu-item']/a[@title='{titleBPO}']"
reporte194(xpathBPO)
nombreAsignado = 'retencionesInbound_bicp_194_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\retencionesInbound\BICP\194'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/retencionesInbound/BICP/194' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])
# === Fin Reportes 194 ===

# === Reporte 192 ===
# RETENCIONES INBOUND
titleBPO = "BPOPERURETENCION"
xpathBPO = f"//li[@class='ui-menu-item']/a[@title='{titleBPO}']"

titleAG = "1070^BPOPERURETENCION"
xpathAgentGroup = f"//li[@title='{titleAG}']/span"
reporte192(xpathBPO, xpathAgentGroup)

nombreAsignado = 'retencionesInbound_bicp_192_'
nombre = nombreReporte(nombreAsignado)
destino = directoryPath + r'/carga\retencionesInbound\BICP\192'
renombrarReubicar(nombre, destino)

# Enviamos los archivos descargados al Servidor
pathLocal = destino + '/' + f'{nombre}.xlsx'
pathServer = r'carga/retencionesInbound/BICP/192' + f'{nombre}.xlsx'
subprocess.call(['python', 'ftp.py', pathLocal, pathServer])
# === Fin Reportes 192 ===

cerrarSesion()
driver.quit()

print('Descarga Exitosa')

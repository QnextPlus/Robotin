#usuarios_password = {Isac, Michael}
user_pass = {'E8234353': 'JJJmm88***1', 'E8234180': 'JJJmm88***1'}
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

    directoryPath = os.getcwd()

    # Funcion que reubicará las descargas en sus respectivas carpetas
    def renombrarReubicar(nuevoNombre, carpetaDestino):
        # ruta_descargas = r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython\descargaRobotin"
        ruta_descargas = directoryPath + r'/pruebasDiario'
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
        fechaHora = datetime.datetime.now()
        fecha = fechaHora.strftime("%Y%m%d_%H%M%S")
        aleatorio = str(random.randint(100, 999))
        nameFile = name + fecha + '_' + aleatorio
        print(name + fecha + '_' + aleatorio)
        return nameFile

    # Funcion para obtener fechas
    def fecha():
        fechaActual = datetime.date.today()
        fechaSiguiente = fechaActual + datetime.timedelta(days=1)
        horaCero = datetime.time(0, 0, 0)
        fechaActualConHora = datetime.datetime.combine(fechaActual, horaCero)
        fechaSiguienteConHora = datetime.datetime.combine(fechaSiguiente, horaCero)
        formato = "%m/%d/%Y %H:%M:%S"
        formato2 = "%Y-%m-%d %H:%M:%S"
        faCH1 = fechaActualConHora.strftime(formato)
        fsCH1 = fechaSiguienteConHora.strftime(formato)
        faCH2 = fechaActualConHora.strftime(formato2)
        fsCH2 = fechaSiguienteConHora.strftime(formato2)
        faFormato1 = fechaActual.strftime("%m/%d/%Y")
        fsFormato1 = fechaSiguiente.strftime("%m/%d/%Y")
        fechaActual = str(fechaActual)
        fechaSiguiente = str(fechaSiguiente)
        fechas = {"hoyH":faCH1, 'mañanaH': fsCH1, 
                'hoyH2': faCH2, 'mañanaH2': fsCH2,
                'hoy': fechaActual, 'mañana': fechaSiguiente, 
                'hoyF1': faFormato1, 'mañanaF1': fsFormato1}
        return fechas

    # Especifica la ruta del controlador de Chrome
    pathDriver = "113/chromedriver.exe"

    #Directorio de descarga
    defaultPathDownloads = directoryPath + r'\pruebasDiario'

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


    username = user
    password = passw
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

        #Periodo
        fechaInicio = fecha()['hoyH']
        fechaFin = fecha()['mañanaH']

        startTime = driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

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
    # +++++ Fin Funcion Reporte 1259 +++++

    # Descarga de reportes
    # === Reportes 1259 ===
    # Blindaje
    labelC = f'OUTBPOPERETENSF'
    reporte1259(f'//input[@type="checkbox" and @label="{labelC}"]')
    """nombreAsignado = 'blindaje_bicp_1259_'
    nombre = nombreReporte(nombreAsignado)
    destino = directoryPath + r'/carga\blindaje\bicp\1259'
    renombrarReubicar(nombre, destino)

    # Enviamos los archivos descargados al Servidor
    pathLocal = destino + '/' + f'{nombre}.xlsx'
    pathServer = r'carga/blindaje/bicp/1259/' + f'{nombre}.xlsx'
    subprocess.call(['python', 'ftp.py', pathLocal, pathServer])"""

    # === Fin Reporte 1259 ===

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
    driver.quit()
    print(f'Se realizó la descarga con el usuario {user}')
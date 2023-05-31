import schedule
import time
import subprocess
import os

directoryPath = os.getcwd()
contador = 0
def ejecutaBICPMaster():
    global contador
    contador += 1
    script_path = 'BICP_Master.py'
    subprocess.call(['python', script_path])
    print(f'Numero de Ejecuciones: {contador}')

schedule.every().hour.at(":00").do(ejecutaBICPMaster)
# Ejecuta el bucle principal
while True:
    schedule.run_pending()
    time.sleep(1)

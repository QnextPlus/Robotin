from ftplib import FTP
import os
import sys

host = "66.206.10.130"
username = "bicp"
password = "fzMVE$0N3f26"

# Conectarse al servidor FTP
ftp = FTP(host)
login_status = ftp.login(user=username, passwd=password)

# Subir un archivo al servidor FTP
archivo_local = '/carga' + str(sys.args[1])
archivo_remoto = archivo_local
with open(archivo_local, 'rb') as archivo:
    ftp.storbinary('STOR ' + archivo_remoto, archivo)


# Cerrar la conexión FTP
ftp.quit()
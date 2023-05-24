from ftplib import FTP
import os

host = "66.206.10.130"
username = "bicp"
password = "fzMVE$0N3f26"

# Conectarse al servidor FTP
ftp = FTP(host)
login_status = ftp.login(user=username, passwd=password)

# Subir un archivo al servidor FTP
archivo_local = 'carga/ISAC_DE_LA_CRUZ_CDR_6_Level_Report_for_Outbound_Campaign_-_ALL_CALLS(2)_769813727.xlsx'
archivo_remoto = 'carga/archivo.xlsx'
with open(archivo_local, 'rb') as archivo:
    ftp.storbinary('STOR ' + archivo_remoto, archivo)


# Cerrar la conexión FTP
ftp.quit()
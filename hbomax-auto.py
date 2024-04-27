import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException  # Importa la excepción
import random


import gdown  #PARA DESCARGAR EL EXCEL DEL DRIVE
import gspread #CONECTAR A LA API DRIVE
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.http import MediaFileUpload

#CONEXIÓN ENVÍO CORREO:
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configurar el alcance y las credenciales
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('.json', scope)
client = gspread.authorize(creds)

# Acceder a la hoja de cálculo en la carpeta KgvStoreDemo
nombre_hoja = "Automatikgv"  # Nombre de tu hoja de cálculo
ruta_hoja = "KgvStoreDemo/Automatikgv"  # Ruta a la hoja de cálculo en Google Drive

# Abrir la hoja de cálculo por su nombre
sh = client.open(nombre_hoja)

# Acceder a una hoja específica por su título
worksheet = sh.worksheet("hbomax")

# Leer los datos de la hoja de cálculo y cargarlos en un DataFrame
data = worksheet.get_all_values()
headers = data.pop(0)
dfhbomax = pd.DataFrame(data, columns=headers)


###################################################################################


# Agregar nuevas columnas al DataFrame
dfhbomax["Error clave"] = ""
dfhbomax["Suscripcion"] = ""

# Ruta específica del ChromeDriver
driver = webdriver.Chrome()
time.sleep(random.uniform(3, 5))

# Abrir el sitio web de HBO Max
driver.get("https://www.hbomax.com/account")
time.sleep(random.uniform(3, 5))

try:
    accept_cookies_button = driver.find_element("xpath", '//*[@id="onetrust-accept-btn-handler"]')
    accept_cookies_button.click()
except NoSuchElementException:
    # El botón de aceptar cookies no aparece, continúa con el proceso
    pass


time.sleep(3)

# Iterar a través de las filas del DataFrame
for index, row in dfhbomax.iterrows():
    correo = row["correo"].strip()
    clave = str(row["clave"]).strip()

    for _ in range(2):
        try:
            email_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'email')))
            email_field.clear()

            for letra in correo:
                email_field.send_keys(letra)
                time.sleep(random.uniform(0.2, 0.9))

            password_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'password')))
            password_field.clear()

            for letra in clave:
                password_field.send_keys(letra)
                time.sleep(random.uniform(0.2, 0.9))
            password_field.send_keys(Keys.RETURN)
            time.sleep(6)

            # Pausa aleatoria
            time.sleep(random.uniform(1.5, 3.0))

            # Verificar el resultado
            current_url = driver.current_url
            if "plan-picker" in current_url:
                dfhbomax.at[index, "Suscripcion"] = "Caida cuenta"
                driver.find_element("xpath", '//*[@id="signOut"]').click()
                driver.get("https://www.hbomax.com/account")
            elif "manage" in current_url:
                dfhbomax.at[index, "Suscripcion"] = "Funciona Cuenta"
                driver.find_element("xpath", '//*[@id="signOut"]').click()
                driver.get("https://www.hbomax.com/account")
            else:
                try:
                    general_error = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'generalError')))
                    error_message = general_error.text
                    dfhbomax.at[index, "Error clave"] = error_message
                    driver.get("https://www.hbomax.com/account")
                except NoSuchElementException:
                    dfhbomax.at[index, "Error clave"] = "Clave incorrecta"

        except Exception as e:
            print(f"Error: {str(e)}")
            print("Esperando 60 segundos y volviendo a cargar la página...")
            time.sleep(60)

            # Verificar si la página se cargó después de esperar 60 segundos
            if driver.current_url == current_url:
                print("La página no ha cargado. detecto el ID DE FALLA, Volviendo a cargar...")
                driver.refresh()
            else:
                # Verificar si se ha producido un error específico
                if driver.find_elements(By.CLASS_NAME, 'ce__GeneralError__traceId'):
                    print("Error específico detectado. Volviendo a cargar el navegador y continuando...")
                    driver.quit()
                    driver = webdriver.Chrome()
                    time.sleep(random.uniform(3, 5))
                    driver.get("https://www.hbomax.com/account")
                    continue  # Reiniciar el bucle de intentos
                else:
                    print("La página ha cargado. Continuando desde donde quedó.")
                    break  # Salir del bucle de espera

    # Verificar si la página se cargó después de esperar 60 segundos
    if driver.current_url != current_url:
        print("La página ha cargado. Continuando desde donde quedó.")
    else:
        # Si después de esperar 60 segundos aún no hay respuesta, cerrar el navegador y guardar el DataFrame
        print("Cerramos el navegador por fallas en el Internet o servidor de la página.")
        dfhbomax.to_csv("resultados_hbomax.txt", sep='\t', index=False)
        driver.quit()
        break  # Salir del bucle principal


# Guardar el DataFrame en un archivo de texto
dfhbomax.to_csv("resultados_hbomax.txt", sep="\t", index=False)

# Dirección de correo electrónico y contraseña del remitente
email_address = "correo@gmail.com"
password = "clave"

# Crear el mensaje
msg = MIMEMultipart()
msg['From'] = email_address
msg['To'] = "correo destino"
msg['Subject'] = "Archivo de resultados Netflix"

# Cuerpo del mensaje
body = "Adjunto encontrarás los resultados del análisis de Netflix."
msg.attach(MIMEText(body, 'plain'))

# Adjuntar el archivo
filename = "resultados_hbomax.txt"
attachment = open(filename, "rb")

part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename= {filename}')

msg.attach(part)
text = msg.as_string()

# Conexión con el servidor SMTP de Gmail
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email_address, password)

# Enviar el correo electrónico
server.sendmail(email_address, "correodestino@gmail.com", text)
server.quit()

# Cerrar el navegador
driver.quit()

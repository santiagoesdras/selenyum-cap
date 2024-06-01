[Yesterday 3:42 PM] ALVAREZ VALIENTE, MIGUEL ALBERTO
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import os
import time
# Configurar ruta al archivo CSV
ruta_csv = r'C:\Users\MALVARE5\OneDrive - Capgemini\Desktop\Facturas\Invoices ELM.csv'
# Leer el archivo CSV
df = pd.read_csv(ruta_csv)
# Asegúrate de que las columnas del CSV se llamen 'NumeroFactura' y 'NombreCuenta'
facturas = df[['NumeroFactura', 'NombreCuenta']]
# Configurar el WebDriver (suponiendo que usas Chrome)
options = webdriver.ChromeOptions()
download_path = r'C:\Users\MALVARE5\OneDrive - Capgemini\Desktop\Facturas'
prefs = {
    "download.default_directory": download_path,
    "savefile.default_directory": download_path,
    "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local"}],"selectedDestinationId":"Save as PDF","version":2}',
    "printing.default_destination_selection_rules": '{"kind":"local","id":"Save as PDF"}'
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--kiosk-printing")
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
# URL de inicio de sesión de NetSuite
url_login = 'https://system.netsuite.com/pages/customerlogin.jsp'
# Credenciales de NetSuite
usuario = 'brandon.cortez@wolterskluwer.com'
contrasena = 'Cyndaquil236'
# Iniciar sesión en NetSuite
driver.get(url_login)
time.sleep(3)  # Esperar a que la página cargue
# Encontrar y llenar los campos de inicio de sesión
try:
    user_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, 'email'))
    )
    user_field.send_keys(usuario)
    password_field = driver.find_element(By.NAME, 'password')
    password_field.send_keys(contrasena)
    password_field.send_keys(Keys.RETURN)
except Exception as e:
    print(f"Error al iniciar sesión: {e}")
    driver.quit()
    exit()
time.sleep(5)  # Esperar a que se complete el inicio de sesión
# Crear una carpeta para guardar las facturas si no existe
os.makedirs(download_path, exist_ok=True)

# Función para esperar hasta que un elemento esté disponible
def esperar_elemento(by, value, tiempo=15):
    return WebDriverWait(driver, tiempo).until(EC.presence_of_element_located((by, value)))

# Función para limpiar nombres de archivos
def limpiar_nombre(nombre):
    return re.sub(r'[^a-zA-Z0-9_\-\.]', '_', nombre)

# Navegar y descargar cada factura
for index, row in facturas.iterrows():
    numero_factura = limpiar_nombre(str(row['NumeroFactura']))
    nombre_cuenta = limpiar_nombre(str(row['NombreCuenta']))
    # Usar la barra de búsqueda
    try:
        search_bar = esperar_elemento(By.ID, 'uif38 input')
    except:
        print("No se pudo encontrar la barra de búsqueda con ID 'uif38 input'. Intentando con CLASS.")
        try:
            search_bar = esperar_elemento(By.CLASS_NAME, 'uif750')
        except:
            print("No se pudo encontrar la barra de búsqueda con CLASS 'uif750'. Intentando con XPATH.")
            search_bar = esperar_elemento(By.XPATH, '//input[@placeholder="Search"]')
    search_bar.clear()
    search_bar.send_keys(numero_factura)
    search_bar.send_keys(Keys.RETURN)
    time.sleep(3)  # Esperar a que se carguen los resultados de la búsqueda
    # Hacer clic en el botón "Imprimir" para abrir la factura como PDF
    try:
        PDFboton_imprimir = esperar_elemento(By.XPATH,
                                             '//span[@id="spn_PRINT_d1"]//a[contains(@class, "pgm_menu_print")]')
    except:
        print("No se pudo encontrar el botón de impresión.")
        continue
    time.sleep(3)  # Esperar a que se abra el menú desplegable de impresión
    # Crear una instancia de ActionChains
    actions = ActionChains(driver)
    # Mover el cursor al botón de imprimir
    actions.move_to_element(PDFboton_imprimir).perform()
    # Seleccionar el primer resultado
    try:
        primer_resultado = esperar_elemento(By.XPATH, '//*[@id="nl1"]/a')
        primer_resultado.click()
    except:
        print("No se pudo encontrar el primer resultado de impresión.")
        continue
    time.sleep(3)  # Esperar a que se abra la página de la factura
    try:
        opcion_print = esperar_elemento(By.XPATH,
                                        '//span[@id="spn_PRINT_d1"]//a[contains(@class, "pgm_menu_print")]//div[contains(@class, "button-print")]')
        opcion_print.click()
    except:
        print("No se pudo encontrar la opción de impresión.")
        continue
    # Cambiar al nuevo tab donde se abre el PDF
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(3)  # Esperar a que se cargue la vista de impresión
    # Enviar la impresión a PDF
    driver.execute_script('window.print();')
    time.sleep(5)  # Esperar a que se complete la impresión a PDF
    # Verificar si el archivo se ha descargado correctamente
    files = [f for f in os.listdir(download_path) if f.endswith('.pdf')]
    if files:
        downloaded_file = max([os.path.join(download_path, f) for f in files], key=os.path.getctime)
        nuevo_nombre_archivo = f'{numero_factura}_{nombre_cuenta}.pdf'
        ruta_archivo = os.path.join(download_path, nuevo_nombre_archivo)
        os.rename(downloaded_file, ruta_archivo)
        print(f"Factura {numero_factura} de {nombre_cuenta} descargada correctamente como {nuevo_nombre_archivo}.")
    else:
        print(f"Error al descargar la factura {numero_factura} de {nombre_cuenta}.")
    # Cambiar de vuelta al tab principal
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
driver.quit()
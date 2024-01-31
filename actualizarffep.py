# Importo selenium para webscraping
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
#from webdriver_manager.chrome import ChromeDriverManager
# Importo autoinstalador del driver de Chrome
#import chromedriver_autoinstaller
import chromedriver_autoinstaller_fix
# Importo pandas para manejo de Excel
import pandas as pd
# Importo time para sleep
import time

# flags
pedirXls = True
maxClientes = 100

# pido el nombre del archivo Excel con las claves
if (pedirXls):
    archivoXls = input("Nombre del archivo Excel con las claves (ej: claves.xlsx): ")
else:
    archivoXls= "claves.xlsx"
print("")

print("Ingrese el nuevo valor del FFEP (separar parte decimal con . )")
ffep = float(input(""))

# seteo las opciones del browser
options = Options()
options.add_argument("--start-maximized")
options.headless = False  # True para modo silencioso (no muestra el browser)
options.add_experimental_option("prefs", {
  #"download.default_directory": r"C:\",
  "download.prompt_for_download": False,
  #"download.directory_upgrade": True,
  #"safebrowsing.enabled": True
})
options.add_experimental_option("excludeSwitches", ["enable-logging"])
#options.add_argument("--user-agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'")

# inicializo el driver del browser Chrome
chromedriver_autoinstaller_fix.install()
#chromedriver_autoinstaller.install(cwd=True)
driver = webdriver.Chrome(options=options)
#driver = webdriver.Chrome(options=options, executable_path='../116/chromedriver.exe')
#driver = webdriver.Chrome(options=options, service=ChromeService(ChromeDriverManager().install()))

# abro el archivo y extraigo los pares usuario, clave
errores = 0
clientes = 0
df = pd.read_excel(archivoXls, header=None)
for row in df.itertuples():
    clientes = clientes + 1
    print("Accediendo con clave fiscal de " + str(row[1]))
    # abro el navegador en la dirección que quiero
    driver.get('https://auth.afip.gob.ar/contribuyente_/login.xhtml')
    cuitElemento = driver.find_element(By.ID, 'F1:username')
    siguienteElemento = driver.find_element(By.ID, 'F1:btnSiguiente')
    cuitElemento.clear()
    cuitElemento.send_keys(str(row[1]))
    driver.execute_script("arguments[0].click();", siguienteElemento)
    try:
        # Ingreso con clave fiscal
        passwordElemento = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "F1:password"))
        )
        time.sleep(3)
        window_before = driver.window_handles[0]
        passwordElemento.send_keys(str(row[2]))
        ingresarElemento = driver.find_element(By.ID, 'F1:btnIngresar')
        driver.execute_script("arguments[0].click();", ingresarElemento)
        time.sleep(3)
        # Con la versión nueva de la página de AFIP, antes hay que ir a la pestaña "Ver todos"
        misserviciosElementos = driver.find_elements(By.CLASS_NAME, 'p-y-1')
        for msElemento in misserviciosElementos:
            if msElemento.get_attribute('innerHTML').find('Ver todos')>=0:
                driver.execute_script("arguments[0].click();", msElemento)
                break
        time.sleep(3)
        # Con la versión nueva de la página de AFIP, antes hay que ir a la pestaña "Ver todos"
        declaenlinea = driver.find_elements(By.CLASS_NAME, 'h5')
        for msElemento in declaenlinea:
            if msElemento.get_attribute('innerHTML').find('DECLARACIÓN EN LÍNEA')>=0:
                driver.execute_script("arguments[0].click();", msElemento)
                break
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])
        # Si hay presencia del form de seleccion de empresa, tengo que elegir una CUIT
        formElemento = driver.find_elements(By.XPATH, "//select[@id='ctl00_ContentPlaceHolder1_ddlCUIT']")
        if len(formElemento) > 0:
            #print("Hay CUITs para elegir")
            select = Select(driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ddlCUIT'))
            select.select_by_value(str(row[3]))
            # Boton ACEPTAR
            aceptarElemento = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_btnAceptar"))
            )
            driver.execute_script("arguments[0].click();", aceptarElemento)
            time.sleep(3)
        # Boton ACEPTAR
        aceptarElemento = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Aviso1_btnAceptar"))
        )
        driver.execute_script("arguments[0].click();", aceptarElemento)
        time.sleep(3)
        # Actualizar datos empleador
        links = driver.find_elements(By.TAG_NAME, 'a')
        for msElemento in links:
            #print (msElemento.get_attribute('href'))
            if msElemento.get_attribute('innerHTML').find('para actualizar sus datos de Empleador')>=0:
                driver.execute_script("arguments[0].click();", msElemento)
                break
        time.sleep(5)
        # Modifico el valor del FFEP
        ffepElemento = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_ucCuotaLRT_txtRem"))
        )
        time.sleep(3)
        driver.execute_script("arguments[0].setAttribute('value',arguments[1])", ffepElemento, str(ffep).replace(".", ","))
        # Boton GUARDAR
        guardarElemento = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_btnGrabar"))
        )
        driver.execute_script("arguments[0].click();", guardarElemento)
        time.sleep(3)
        # Cierro la sesión
        driver.close()
        driver.switch_to.window(window_before)
        usuarioElemento = driver.find_element(By.XPATH, "//div[@id='iconoChicoContribuyenteAFIP']")
        driver.execute_script("arguments[0].click();", usuarioElemento)
        time.sleep(1)
        salirElemento = driver.find_element(By.XPATH, "//button[@title='Salir']")
        driver.execute_script("arguments[0].click();", salirElemento)
        time.sleep(3)
    except Exception as e:
        print("Se ha producido un error con el contribuyente " + str(row[1]) + " -> " + str(e))
        errores = errores + 1

    # cierro sesión
    time.sleep(3)

    if (clientes >= maxClientes):
        print("Se ha llegado al límite máximo de contribuyentes")
        break

# termino
print("Tarea terminada con " + str(errores) + " errores")

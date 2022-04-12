#Recursos
import time
import warnings
import traceback
import pandas as pd
import math

from VAST import calc
from GOOGLE import gapi
from pytz import timezone
from datetime import datetime
from thefuzz import fuzz, process
from pickle import FALSE, TRUE
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementClickInterceptedException

#Ignorar DeprecationWarning
warnings.filterwarnings("ignore", category=DeprecationWarning)

#Definir Chromedriver
CHROMEDRIVER_PATH = '/usr/local/bin/chromedriver'
WINDOW_SIZE = "1920,1080"

chrome_options = Options()
chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument("--test-type")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-first-run")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--start-maximazed")

driver = webdriver.Chrome(executable_path=CHROMEDRIVER_PATH,chrome_options=chrome_options)
ignored_exceptions=(NoSuchElementException,StaleElementReferenceException)

#Definir ActionChains
action = ActionChains(driver)

#Definir TimeZone
tz = timezone('America/Caracas')

#Jalada de correos a sheets
def scrape():
    #Llamadas al API
    #trof = gapi(1).mail()
    gt = gapi(2).mail()
    #Vaciado en SHEETS
    vast = gapi(3).sheets(gt,'A','last')
scrape()

#Desk Login
def desk_login():
    #Bienvenida
    print("Hola operardor VAST, bienvenido a tu AUTOMANUAL para ZELLE, en breves momentos estaré a operativo.")           
    #Solicitar credenciales
    desk_link = gapi(3).credentials('link')
    user = gapi(3).credentials('user')
    password = gapi(3).credentials('password')

    #Inicio de sesion DESK
    try:
        driver.get(desk_link)
    except:
        print("Link para accesar el portal cambió")
    #Login
    try:
        #Usuario
        usertxt = WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div/div/div/div/div/form/div[1]/div/div/input'))
        )
        usertxt.send_keys(user)
        #Clave
        passtxt =  WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div/div/div/div/div/form/div[2]/div/div/input'))
        )
        passtxt.send_keys(password)
        #Continuar
        advance = driver.find_element_by_xpath('//*[@id="root"]/div/main/div/div/div/div/div/form/div[3]/button')
        advance.click()
    except:
        print("No se logró iniciar sesión en Desk, porfavor reiniciar bot")
        driver.quit
    #Ir a VOTC
    try:
        panel = WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH,'//a[@href="/votc_board"]'))
        )
        panel.click()
        print("Sesión iniciada en Desk")
    except:
        print("No se logró iniciar sesión en Desk, porfavor reiniciar bot")
        driver.quit
desk_login()

#Operar
def operate():
    #Confirmar transacción
    def confirmar():
        if monto != monto_sheet:
            #Corregir monto si hay discrepancia
            print("Discrepancia encontrada en el monto, procediendo a corregir")
            time.sleep(1)
            discrepancia = WebDriverWait(driver,5).until(
                EC.presence_of_element_located((By.XPATH,"(//button[@type='button'])[11]"))
            )
            discrepancia.click()
            time.sleep(1)
            correccion = WebDriverWait(driver,5).until(
                EC.presence_of_element_located((By.XPATH,"(//input[@type='text'])[2]"))
            )
            correccion.send_keys(monto_sheet)
            #Completar entrada
            completar_entrada = WebDriverWait(driver,5).until(
                EC.presence_of_element_located((By.XPATH,"//button[@class='MuiButtonBase-root MuiButton-root MuiButton-contained MuiButton-containedPrimary']"))
            )
            action.move_to_element(completar_entrada)
            completar_entrada.click
        else:
            completar_entrada = WebDriverWait(driver,5).until(
                EC.presence_of_element_located((By.XPATH,"//button[@class='MuiButtonBase-root MuiButton-root MuiButton-contained']"))
            )
            action.move_to_element(completar_entrada)
            completar_entrada.click()

    #Cancelar transacción
    def cancelar():
        accion_votc = WebDriverWait(driver,5).until(
            EC.presence_of_element_located((By.XPATH,"//div[@class='MuiSelect-root MuiSelect-select MuiSelect-selectMenu MuiInputBase-input MuiInput-input']"))
        )
        accion_votc.click()
        time.sleep(1)
        try:
            item_votc = WebDriverWait(driver,5).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="menu-"]/div[3]/ul/li[6]'))
            )
            action.move_to_element(item_votc).click().perform()
        except:
            item_votc = WebDriverWait(driver,5).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="menu-"]/div[3]/ul/li[6]'))
            )
            action.move_to_element(item_votc).click().perform()
        time.sleep(1)
        cancelar = WebDriverWait(driver,5).until(
            EC.presence_of_element_located((By.XPATH,"(//button[@type='button'])[13]"))
        )
        cancelar.click()
    
    #Buscar transacciones
    print("Buscando transacciones para hacer match")
    #Definir variable para cuenta de transacciones
    offset = 0
    #Loop de transacciones
    while TRUE:
        #Buscar nuevos transid en sheets
        id_list = (gapi(3).lookup()).to_dict('New Records')
        time.sleep(20)
        new_id_list = (gapi(3).lookup()).to_dict('New Records')
        for i in id_list:
            if i in new_id_list:
                new_id_list.remove(i)

        time.sleep(5)
        #Transacciones tomadas
        rows =  driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
        rows_count = len(rows)
        print("Transacciones en cola", round(rows_count/2))
        #Resetear offset
        if offset>(rows_count-2):
            offset = 0
            scrape()
        if rows_count > 0 and new_id_list != []:
            #Empezar a medir tiempo de operacion
            begin_time = datetime.now()
            #Busqueda de transacciones tomadas
            time.sleep(1)
            #Contar el numero de transacciones
            rows = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
            rows_count = len(rows)
            last = rows_count - offset - 1 
            try:
                drop_down = WebDriverWait(driver,3).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="root"]//tbody/tr['+str(last)+']/td[8]/button'))
                )
                drop_down.click()
            except:
                continue
            #Datos de primera fila
            try:
                tipo = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last)+']/td[2]/div').text
                banco = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last)+']/td[3]').text
                timer = datetime.strptime(str(driver.find_element_by_xpath("//*[@id='root']//tbody/tr["+str(last)+"]/td[6]//*[name()='svg']").get_attribute('title')).replace('+',''),'%M:%S').time()
            except:
                continue
            #Volver a contar transacciones
            time.sleep(4)
            rows = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
            try:
                trans_id = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[1]/div/p').text
                nombre_reserve = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[2]/div/p').text
                usuario_reserve = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[3]/div/p').text
                rsv = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[4]/div/p').text
                tasa = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[2]/div[4]/div/p').text
                telefono = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[3]/div/div/p').text
                remitente = str(driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[4]/div/div/div/p').text).upper()
                beneficiario = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[5]/div/div[1]/div/p').text
                correo = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[5]/div/div[2]/div/p').text
                monto = round(float(rsv)*float(tasa.replace(",",".")),2)
            except:
                print("Error recogiendo los datos, posible cambio en el portal, si persiste contactar a soporte")
                continue
            #Imprimir datos para feedback de user
            data_list = []
            data = {
                'Remitente': remitente,
                'Monto': monto,
                'TransId': trans_id,
                'Nombre': nombre_reserve,
                'Usuario': usuario_reserve,
                'RSV': monto,
                'Tasa': tasa,
                'Teléfono': telefono,
                'Beneficiario': beneficiario,
                'Correo': correo,
                'Fecha de cierre': str(datetime.today().date())
            }
            data_list.append(data)
            df = pd.DataFrame(data_list)

            #Comparar transids
            index=0
            for row in new_id_list:
                index+=1
                monto_sheet = float(row['Monto'])
                if trans_id==row['TransId']:
                    print('Transacción encontrada')
                    confirmar()
                    gapi(3).sheets(df,'F',str(row['Index']))
                    new_id_list.clear()
                    break
                elif index == len(new_id_list):
                    #Si no hay match, cuenta las filas, cierra el calim y pasa a la siguiente
                    print('Transacción no encontrada')
                    rows = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
                    drop_down.click()
                    offset +=2
        else:
            scrape()
operate()
             


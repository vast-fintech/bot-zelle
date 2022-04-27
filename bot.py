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
chrome_options.add_argument('--headless')
chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument("--test-type")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-first-run")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--start-maximazed")

#Definir TimeZone
tz = timezone('America/Caracas')

def main():
    #Llamar al webdriver
    driver = webdriver.Chrome(executable_path=CHROMEDRIVER_PATH,chrome_options=chrome_options)
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException)

    #Definir ActionChains
    action = ActionChains(driver)
    
    #Manejo de excepciones
    def handle_exception():
        print('Algo ocurrió pero no te preocupes, me estoy reiniciando')
        driver.quit()
        time.sleep(2)
        main()
        

    #Jalada de correos a sheets
    def scrape():
        #Llamadas al API
        gt = gapi(2).mail()
        #Vaciado en SHEETS
        vast = gapi(3).sheets(gt,'A','last')
    #scrape()

    #Desk Login
    def desk_login():
        #Bienvenida
        print("Hola operardor VAST, soy tu BOT ZELLE, en breves momentos empezaré a operar.")           
        #Solicitar credenciales
        try:
            desk_link = gapi(3).credentials('link')
            user = gapi(3).credentials('user')
            password = gapi(3).credentials('password')
        except:
            handle_exception()

        #Inicio de sesion DESK
        try:
            driver.get(desk_link)
        except:
            print("Link para accesar el portal cambió")
            handle_exception()
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
            print("No se logró iniciar sesión en Desk, reiniciando bot")
            handle_exception()
        #Ir a VOTC
        try:
            panel = WebDriverWait(driver,10).until(
                EC.presence_of_element_located((By.XPATH,'//a[@href="/votc_board"]'))
            )
            panel.click()
            print("Sesión iniciada en Desk")
        except:
            print("No se logró iniciar sesión en Desk, reiniciando bot")
            handle_exception()
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
                    EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[2]/div[1]/div[2]/div[1]/button'))
                )
                discrepancia.click()
                time.sleep(1)
                correccion = WebDriverWait(driver,5).until(
                    EC.presence_of_element_located((By.XPATH,"//input[@class='MuiInputBase-input MuiOutlinedInput-input MuiInputBase-inputMarginDense MuiOutlinedInput-inputMarginDense']"))
                )
                correccion.send_keys(str(monto_sheet).replace('.',','))
                #Completar entrada
                completar_entrada = WebDriverWait(driver,5).until(
                    EC.presence_of_element_located((By.XPATH,"//button[@class='MuiButtonBase-root MuiButton-root MuiButton-contained MuiButton-containedPrimary']"))
                )
                action.move_to_element(completar_entrada)
                completar_entrada.click()
                #Seguro?
                confirmacion=WebDriverWait(driver,5).until(
                    EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div[3]/div/div[3]/button[2]'))
                )
                confirmacion.click()
            else:
                try:
                    completar_entrada =driver.find_element_by_xpath("//button[@class='MuiButtonBase-root MuiButton-root MuiButton-contained']")
                    action.move_to_element(completar_entrada)
                    completar_entrada.click()
                except:
                    handle_exception()
            print('Operación CONFIRMADA con éxito')


        #Cancelar transacción
        def cancelar():
            a = 0
            while a == 0:
                try:
                    accion_votc = driver.find_element_by_xpath("//div[@class='MuiSelect-root MuiSelect-select MuiSelect-selectMenu MuiInputBase-input MuiInput-input']")
                    time.sleep(7)
                    action.move_to_element(accion_votc).click().perform()
                    time.sleep(1)
                except:
                    print('Buscando botón de cancelar')
                    continue
                print('Botón encontrado')
                a=1
            a = 0
            while a == 0:
                try:
                    item_votc = driver.find_element_by_xpath('//*[@id="menu-"]/div[3]/ul/li[1]')
                    item_votc.click()
                except:
                    print('Buscando item de cancelación')
                    continue
                print('Item encontrado')
                a=1

            time.sleep(2)
            try:
                cancelar = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[2]/div[2]/div/button')
                cancelar.click()
            except:
                handle_exception()
            print('Operación CANCELADA con éxito')

        #Cuenta de transacciones
        def count():
            rows =  driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
            rows_count = len(rows)
            return(rows_count)

        #Buscar transacciones
        print("Buscando transacciones")
        #Definir variable para cuenta de transacciones
        offset = 0
        #Loop de transacciones
        while TRUE:
            time.sleep(5)
            claims = count()
            print("Transacciones en cola", round(claims/2))
            #Resetear offset
            if offset>(claims-2):
                offset = 0
                #scrape()

            if claims > 0:
                #Empezar a medir tiempo de operacion
                begin_time = datetime.now()
                #Busqueda de transacciones tomadas
                time.sleep(1)
                #Contar el numero de transacciones
                last = claims - offset - 1 
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
                except:
                    continue
                try: 
                    timer = str(driver.find_element_by_xpath("//*[@id='root']//tbody/tr["+str(last)+"]/td[6]//*[name()='svg']").get_attribute('title')).replace('+','')
                except:
                    print('No se pudo recoger el timer')
                    pass
                
                #Volver a contar transacciones
                time.sleep(4)
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
                    confirmacion = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(last+1)+']/td/div/div/div/div/div[1]/div[1]/div[6]/div/div/div/p').text
                except:
                    print("Error recogiendo los datos, posible cambio en el portal")
                    continue

                #Cancelar si lleva mas de 20 minutos
                int_timer = int((timer.split(':'))[0])
                if int_timer>20:
                    print('Transacción lleva', timer, 'minutos, procediendo a cancelar')
                    cancelar()
                    continue

                def frame():
                #Imprimir datos para feedback de user
                    data_list = []
                    data = {
                        'Remitente': remitente,
                        'Monto': monto,
                        'TransId': trans_id,
                        'Nombre': nombre_reserve,
                        'Usuario': usuario_reserve,
                        'RSV': monto_sheet,
                        'Tasa': tasa,
                        'Fecha de cierre': str(datetime.today().date()),
                        'Beneficiario': beneficiario,
                        'Fuzzy Match': str(fuzz.ratio(str(row['Nombre']),str(remitente))),
                        'Fullfilment': 'BOT'
                    }
                    data_list.append(data)
                    df = pd.DataFrame(data_list)
                    return(df)

                #Comparar transacciones
                transactions = gapi(3).compare()
                transact_dict = transactions.to_dict('Records')
                index = 1

                for row in transact_dict:
                    #Preparar nombres
                    bank_name = str(row['Nombre'])
                    desk_name = remitente
                    #Separar nombres
                    bank_name_ls = (str(row['Nombre']).split(' '))
                    desk_name_ls = (remitente.split(' '))
                    #Reordenar si el nombre tiene más de 4 elementos
                    if len(bank_name_ls) >= 4:
                        bank_name = str(str(bank_name_ls[0])+' '+str(bank_name_ls[2]+' '+str(bank_name_ls[3])))
                        bank_name2 = str(str(bank_name_ls[0])+' '+str(bank_name_ls[1]+' '+str(bank_name_ls[3])))
                    else:
                        bank_name2 = bank_name
                    if len(desk_name_ls) >= 4:
                        desk_name = str(str(desk_name_ls[0])+' '+str(desk_name_ls[2]+' '+str(desk_name_ls[3])))
                        desk_name2 = str(str(desk_name_ls[0])+' '+str(desk_name_ls[1]+' '+str(desk_name_ls[3])))
                    else:
                        desk_name2 = desk_name
                    index+=1
                    monto_sheet = float(row['Monto'])
                    discrepancy = abs((monto-monto_sheet)/monto)
                    cascade = calc(monto).tolerance()
                    
                    if  fuzz.ratio(str(bank_name),str(desk_name))>=65 and discrepancy < cascade and str(row['Remitente Desk'])=='POR PAGAR':
                        print('Transacción encontrada')
                        print(frame())
                        try:
                            gapi(3).sheets(frame(),'G',index)
                        except:
                            handle_exception()
                        confirmar()
                        break
                    elif  fuzz.ratio(str(bank_name2),str(desk_name))>=65 and discrepancy < cascade and str(row['Remitente Desk'])=='POR PAGAR':
                        print('Transacción encontrada')
                        print(frame())
                        try:
                            gapi(3).sheets(frame(),'G',index)
                        except:
                            handle_exception()
                        confirmar()
                        break
                    elif  fuzz.ratio(str(bank_name),str(desk_name2))>=65 and discrepancy < cascade and str(row['Remitente Desk'])=='POR PAGAR':
                        print('Transacción encontrada')
                        print(frame())
                        try:
                            gapi(3).sheets(frame(),'G',index)
                        except:
                            handle_exception()
                        confirmar()
                        break
                    elif  fuzz.ratio(str(bank_name2),str(desk_name2))>=65 and discrepancy < cascade and str(row['Remitente Desk'])=='POR PAGAR':
                        print('Transacción encontrada')
                        print(frame())
                        try:
                            gapi(3).sheets(frame(),'G',index)
                        except:
                            handle_exception()
                        confirmar()
                        break
                    elif index == len(transact_dict)+1:
                        try:
                            #Si no hay match, cuenta las filas, cierra el calim y pasa a la siguiente
                            #scrape()
                            print('Transacción no encontrada')
                            print(frame())
                            drop_down.click()
                            offset +=2
                        except:
                            handle_exception()
            else:
                pass
    operate()
main()

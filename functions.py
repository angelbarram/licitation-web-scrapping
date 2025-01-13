import pandas as pd
import numpy as np
import re
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import ElementClickInterceptedException, WebDriverException, ElementNotInteractableException, TimeoutException
import requests
import os
import time
from datetime import datetime
import openai
import smtplib
from datetime import datetime, timedelta
import threading
import shutil
from email.mime.text import MIMEText
import sys

def descargar_licitaciones():
    try:
        options = Options()
        options.add_argument("--headless=new")
        try:
            driver = webdriver.Chrome(options=options)
        except Exception as e:
            print("Actualizar chromedriver.exe la información se enviará al correo.")
            advertencia = "Para actualizar chromedriver debe acceder a este sitio: \nhttps://googlechromelabs.github.io/chrome-for-testing/\ny descargar la version denominada como Stable y extraer del archivo .zip el ejecutable chromedriver.exe para reemplazarlo en la carpeta donde se ubique el código."
            send_email(advertencia, "Actualizar chromedriver")

            counter_value = 2
            with open("counter.txt","w") as file:
                file.write(str(counter_value))

            sys.exit()
        wait = WebDriverWait(driver, 10)
        wait_cola = WebDriverWait(driver, 1000)
        driver.get("https://www.mercadopublico.cl/Home")
        driver.maximize_window()

        time.sleep(30)

        try:
            inicio_sesion = driver.find_element("xpath",'//button[@value="Iniciar sesión"]')
            inicio_sesion.click()
        except ElementClickInterceptedException:
            print("Elemento no clickeable en este momento, manejando la excepción.")
            # time.sleep(10)
            # inicio_sesion.click()

        clave_unica = driver.find_element("xpath","//img[@class]")
        clave_unica.click()

        time.sleep(10)

        usuario_run = driver.find_element("xpath","//input[@id='uname']")
        usuario_claveunica = driver.find_element("xpath","//input[@id='pword']")

        ingresar_run = "rut"
        ingresar_clave_unica = "mercado-publico-clave"

        usuario_run.send_keys(ingresar_run)
        usuario_claveunica.send_keys(ingresar_clave_unica)

        ingresar = driver.find_element("xpath","//button")
        ingresar.click()

        time.sleep(30)

        entidad = wait.until(EC.visibility_of_element_located(('xpath','//span[@class="wrap-td"]')))
        entidad.click()

        ingresar2 = driver.find_element("xpath",'//a[@class="btn btn-pri"]')
        ingresar2.click()

        time.sleep(10)

        menu_licitaciones = driver.find_element("xpath",'//img[@alt="Expand Licitaciones "]')
        acciones = ActionChains(driver)
        acciones.move_to_element(menu_licitaciones).perform()

        oferta_licitaciones = driver.find_element("xpath",'//a[@href="/BID/Modules/RFB/NEwSearchProcurement.aspx"]')
        oferta_licitaciones.click()

        time.sleep(10)

        driver.switch_to.frame("fraDetalle")

        descargar_excel = wait.until(EC.visibility_of_element_located(("xpath",'//a[@id="lbExcel"]')))

        descargar_excel.click()

        time.sleep(30)

        driver.quit()

        counter_value = 1
        with open("counter.txt","w") as file:
            file.write(str(counter_value))

    except (ElementClickInterceptedException, ElementNotInteractableException, TimeoutException):
        print("Reiniciar ejecución, problemas con la carga de Mercadopublico.com")
        counter_value = 0
        with open("counter.txt","w") as file:
            file.write(str(counter_value))

def reordenamiento_de_datos(intereses):
    licitacion = "excel_licitaciones/licitacion_Publicada.xlsx"
    licitacion = pd.read_excel(licitacion)
    licitacion.columns = ['?',
                      'Número Adquisición',
                      'Nombre Adquisición',
                      '?',
                      'Nombre licitación',
                      'Descripción',
                      'Organismo',
                      'Región Compradora',
                      '?',
                      'Fecha Publicación',
                      'Fecha Cierre',
                      'Descripción del producto/servicio',
                      'Código ONU',
                      'Unidad de Medida',
                      'Cantidad',
                      'Genérico',
                      'Nivel 1',
                      'Nivel 2',
                      'Nivel 3']
    licitacion = licitacion.drop([0, 1, 2, 3, 4, 5, 6])
    licitacion = licitacion.reset_index(drop=True)
    licitacion = licitacion.dropna(axis=1,how='all')
    licitacion['Descripción del producto/servicio'] = licitacion['Descripción del producto/servicio'].fillna('')

    for i in range(len(licitacion['Descripción del producto/servicio'])):
        licitacion.loc[i, 'Descripción del producto/servicio'] = licitacion.loc[i, 'Descripción del producto/servicio'].lower()

    patron = r'\b(?:{})\b'.format('|'.join(intereses))
    filtro = licitacion['Descripción del producto/servicio'].str.contains(patron, case=False, regex=True)
    nuevo_licitacion = licitacion.loc[filtro, ['Número Adquisición', 'Descripción','Descripción del producto/servicio', 'Fecha Publicación', 'Fecha Cierre', 'Unidad de Medida', 'Cantidad']]
    nuevo_licitacion['Palabra clave'] = nuevo_licitacion['Descripción del producto/servicio'].str.extract('({})'.format('|'.join(intereses)), flags=re.IGNORECASE)
    nuevo_licitacion  = nuevo_licitacion.drop_duplicates(subset='Número Adquisición')
    nuevo_licitacion = nuevo_licitacion.sort_values('Palabra clave')
    nuevo_licitacion.reset_index(drop=True)
    nuevo_licitacion.to_excel("nuevo_licitacion.xlsx", sheet_name='Licitaciones filtradas')
    ids_de_licitacion = nuevo_licitacion['Número Adquisición'].tolist()

    return nuevo_licitacion, ids_de_licitacion

def ingresar_licitaciones(id_licitaciones):
    resultados=id_licitaciones
    try:
        options = Options()
        options.add_argument("--headless=new")
        options.add_argument('--ignore-certificate-errors')
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 30)
        wait_cola = WebDriverWait(driver, 1000)
        driver.get("https://www.mercadopublico.cl/Home")

        time.sleep(30)
        try:
            inicio_sesion = driver.find_element("xpath",'//button[@value="Iniciar sesión"]')
            inicio_sesion.click()
        except ElementClickInterceptedException:
            print("Elemento no clickeable en este momento, manejando la excepción.")
            # time.sleep(10)
            # inicio_sesion.click()

        clave_unica = driver.find_element("xpath","//img[@class]")
        clave_unica.click()

        usuario_run = driver.find_element("xpath","//input[@id='uname']")
        usuario_claveunica = driver.find_element("xpath","//input[@id='pword']")

        ingresar_run = "rut"
        ingresar_clave_unica = "mercado-publico-clave"

        usuario_run.send_keys(ingresar_run)
        usuario_claveunica.send_keys(ingresar_clave_unica)

        ingresar = driver.find_element("xpath","//button")
        ingresar.click()

        time.sleep(30)

        entidad = wait.until(EC.visibility_of_element_located(('xpath','//span[@class="wrap-td"]')))
        entidad.click()

        ingresar2 = driver.find_element("xpath",'//a[@class="btn btn-pri"]')
        ingresar2.click()

        time.sleep(30)

        menu_licitaciones = driver.find_element("xpath",'//img[@alt="Expand Licitaciones "]')
        acciones = ActionChains(driver)
        acciones.move_to_element(menu_licitaciones).perform()

        oferta_licitaciones = driver.find_element("xpath",'//a[@href="/BID/Modules/RFB/NEwSearchProcurement.aspx"]')
        oferta_licitaciones.click()

        # driver.switch_to.frame("fraDetalle")

        # with open(osit,'r',encoding='utf-8') as archivo:
        #     osforit = archivo.read()

        # with open(perfil, 'r',encoding='utf-8') as archivo:
        #     perfil_licitaciones = archivo.read()   

        montos = []
        descripciones = []
        divisas = []
        plazo_preguntas = []
        ids_licitacion = []
        tipo_licitacion = []
        tipo_convocatoria = []
        num_etapas = []
        razon_contraloria = []
        razon_social = []
        estado_licitacion = []

        directorio = "licitaciones_nuevas"

        if os.path.exists(directorio):
            shutil.rmtree(directorio)
        os.makedirs(directorio)

        for resultado in resultados:
            id_licitacion = resultado
            driver.switch_to.frame("fraDetalle")

            time.sleep(1)

            ingresar_id_licitacion = driver.find_element("xpath",'//input[@id="txtProcCode"]')
            ingresar_id_licitacion.send_keys(id_licitacion)

            boton_id_licitacion = driver.find_element("xpath",'//input[@id="btnSearchByCode"]')

            time.sleep(1)

            boton_id_licitacion.click()

            ingresar_id_licitacion.clear()

            time.sleep(1)

            descripcion_web = driver.find_element("xpath",'//span[@id="rptAcquisition_ctl01_lblDescription"]')
            descripcion = descripcion_web.text
            descripciones.append(descripcion)

            ver_ficha = driver.find_element("xpath",'//input[@id="rptAcquisition_ctl01_ibtnVerAdquisicion"]')
            ver_ficha.click()

            time.sleep(1)

            driver.switch_to.default_content()

            ventana_principal = driver.current_window_handle

            wait.until(EC.number_of_windows_to_be(2))

            ventanas = driver.window_handles

            driver.switch_to.window(ventanas[1])

            donde_estado = driver.find_element('xpath','//img[@id="imgEstado"]')

            estado = donde_estado.get_attribute('src')

            if 'cerrada' in estado:
                # print('La licitación está cerrada')
                estado_licitacion.append('Cerrada')
            elif 'adjudicada' in estado:
                #print('La licitación ya ha sido adjudicada')
                estado_licitacion.append('Abjudicada')
            elif 'publicadas' in estado:
                print('La licitación '+id_licitacion+' aún está disponible')
                estado_licitacion.append('Abierta')
            elif 'desierta' in estado:
                # print('La licitación está desierta')
                estado_licitacion.append('Desierta')
            elif 'revocada' in estado:
                # print('Licitación revocada')
                estado_licitacion.append('Revocada')

            try:
                # Encontrar todos los elementos <span> que contengan la palabra "Monto" (ignorando mayúsculas y minúsculas)
                spans_monto = driver.find_elements("xpath", "//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'monto')]")

                monto_encontrado = False

                tipo_moneda = driver.find_element("xpath",'//span[@id="lblFicha1Moneda"][@class="texto09"]')

                divisa = tipo_moneda.text

                fecha_fin_preguntas = driver.find_element("xpath",'//span[@id="lblFicha3Fin"][@class="texto09"]')
                fin_preguntas = fecha_fin_preguntas.text
                
                caja_ficha = driver.find_element('xpath', '//div[@class="caja_ficha00"]')

                # Encontrar los elementos span con la clase "texto09" dentro del div
                elementos_span = caja_ficha.find_elements('css selector', 'span.texto09')

                try:

                    licitacion_tipo = driver.find_element('xpath', '//span[@id="lblFicha1Tipo"]')
                    licitacion_tipo_texto = licitacion_tipo.text
                    tipo_licitacion.append(licitacion_tipo_texto)
                except NoSuchElementException:
                    tipo_licitacion.append('No se proporciona informacion')


                convocatoria_tipo = driver.find_element('xpath','//span[@id="lblFicha1Convocatoria"]')

                convocatoria_tipo_texto = convocatoria_tipo.text
                tipo_convocatoria.append(convocatoria_tipo_texto)

                etapas = driver.find_element('xpath','//span[@id="lblFicha1Etapas"]')

                etapas_texto = etapas.text
                num_etapas.append(etapas_texto)

                contraloria = driver.find_element('xpath','//span[@id="lblFicha1TR"]')

                contraloria_texto = contraloria.text
                razon_contraloria.append(contraloria_texto)

                razon = driver.find_element('xpath','//span[@id="lblFicha2Razon"]')

                razon_texto = razon.text
                razon_social.append(razon_texto)

                elementos_por_id = []
                # Filtrar y luego imprimir el contenido de los elementos span
                # try:
                for elemento in elementos_span:
                        # Obtener el id del elemento]
                    elemento_id = elemento.get_attribute('id')

                        # Comprobar si el id contiene las palabras no deseadas
                    if 'Categoria' not in elemento_id and 'TituloCategoria' not in elemento_id and 'Unidad' not in elemento_id and 'Cantidad' not in elemento_id:
                        elementos_por_id.append(elemento.text)
                        ids_licitacion.append(resultado)
                            # print(elemento.text)           
                # except Exception as e:
                #     elementos_por_id = "No se especificaron los elementos."

                # Recorrer todos los elementos <span> con la palabra "Monto" y obtener el siguiente <span> con class="texto09"
                for span_monto in spans_monto:
                    span_texto09_siguiente = span_monto.find_element("xpath", "following::span[@class='texto09'][1]")
                    texto_siguiente = span_texto09_siguiente.text
                    
                    # Verificar si el texto del monto es un número utilizando isdigit()
                    if texto_siguiente.isdigit():
                        # print("Monto encontrado:", texto_siguiente+' '+divisa)
                        montos.append(texto_siguiente)
                        divisas.append(divisa)
                        plazo_preguntas.append(fin_preguntas)
                        monto_encontrado = True
                        break  # Salir del bucle una vez que se encuentra el primer monto que es un número

                # Si no se encontró un monto válido, imprimir "Monto no especificado"
                if not monto_encontrado:
                    montos.append("Monto no especificado")
                    divisas.append('-')
                    plazo_preguntas.append(fin_preguntas)
                    # print("Monto no especificado")

            except Exception as e:
                montos.append("Monto no especificado")
                divisas.append('-')
                plazo_preguntas.append("No se especifica fin de preguntas")

            time.sleep(1)

            driver.switch_to.window(ventanas[0])

            try:
                if len(elementos_por_id) == 0:
                    elemento_final = "No se especifica"
                elif len(elementos_por_id) == 1:
                    elemento_final = str(elementos_por_id)
                elif len(elementos_por_id) == 2:
                    elemento_final = 'y '.join(elementos_por_id)
                elif len(elementos_por_id) > 2:
                    elemento_final = ', '.join(elementos_por_id[:-1]) + 'y ' + elementos_por_id[-1]


                pregunta = "La licitación de ID: \n"+resultado+"\n. \nCuya descripción es la siguiente: \n"+descripcion+".\n y consiste en los siguientes elementos:\n"+elemento_final+"\n El monto es el siguinte:\n"+texto_siguiente+" "+divisa+"\nEs un tipo de licitación "+licitacion_tipo_texto+"\nY de una convocatoria de tipo\n"+convocatoria_tipo_texto+"\nLas preguntas son hasta el:\n"+fin_preguntas+"\n y posee un total de \n"+etapas_texto
            except Exception as e:

                pregunta = "a licitación de ID: \n"+resultado+"\n. \nCuya descripción es la siguiente: \n"+descripcion
            # pregunta = "La licitación de número ID: "+resultado+". Cuya descripción es la siguiente: "+descripcion+" y consiste en los siguientes elementos: "+elemento_final+". El monto es el siguiente: "+texto_siguiente+" "+divisa+". Faltan 15 días para el fin de preguntas. El tipo de licitación es: "+licitacion_tipo_texto+". La convocatoria es: "+convocatoria_tipo_texto
            # pregunta_2 = ".  ¿Esta licitación podría considerarse para la empresa OS4IT que busca licitaciones con este perfil: "+perfil_licitaciones+"? Si la licitacion es considera buena para la empresa comienza con un 'Si' y realiza un pequeño reporte de la licitación mientras que si no se considera, empieza con un 'No'. También agrega el ID de la licitación. Si es 'No' ¿Por qué?"
            # pregunta = pregunta_1+pregunta_2

            # print(perfil_licitaciones)

            nombre_archivo = f"{directorio}/{resultado}.txt"
            with open(nombre_archivo, "w", encoding="utf-8") as archivo:
                archivo.write(pregunta)

            driver.set_window_position
        time.sleep(1)
        counter_value = 1
        with open("counter.txt","w") as file:
            file.write(str(counter_value))

    except (ElementClickInterceptedException, ElementNotInteractableException, TimeoutException):
        print("Reiniciar ejecución, problemas con la carga de Mercadopublico.com")
        counter_value = 0
        with open("counter.txt","w") as file:
            file.write(str(counter_value))

def get_completion(prompt, model="gpt-3.5-turbo"):

    openai.api_key = 'api-key'

    messages = [{"role": "user", "content": prompt}]

    response = openai.ChatCompletion.create(

    model=model,

    messages=messages,

    temperature=0,

    )

    return response.choices[0].message["content"]

def send_email(response, asunto):
    subject = "Nueva licitación: "+asunto

    # Crear un objeto MIMEText con el cuerpo del mensaje y la codificación UTF-8
    message = MIMEText(response, 'plain', 'utf-8')
    message['Subject'] = subject

    # Configurar el servidor SMTP y enviar el correo electrónico
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login('sendermail', 'api-key')
        server.sendmail('sendermail', 'receivermail', message.as_string())

    print('Mail sent')

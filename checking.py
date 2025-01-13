from functions import *
from zipfile import ZipFile
# from transformers import AutoTokenizer
import sys
import smtplib
import time
import threading

counter_value = 0
with open("counter.txt", "w") as file:
    file.write(str(counter_value))

intereses = ['Oracle',
             'VMWare',
             'Nutanix',
             'RedHat',
             'Dell',
             'Lenovo',
             'NetApp',
             'Veeam',
             'CommVault',
             'Cisco',
             'Nagios',
             'SolarWinds',
             'HPE',
             'HP',
             'Notebook',
             'Laptop',
             'Ordenador']

for i in range(len(intereses)):
    intereses[i] = intereses[i].lower()

path_descargas = "C:/Users/angel/Downloads/Licitacion_Publicada.zip"

if os.path.exists("archivo.txt"):
    os.remove("archivo.txt")

if os.path.exists(path_descargas):
    os.remove(path_descargas)    

# descargar_licitaciones()

def ejecutar_descarga():
    while True:
        start_time = datetime.now()  # Registra el tiempo de inicio de la ejecución
        
        t = threading.Thread(target=descargar_licitaciones)
        t.start()

        while t.is_alive():
            elapsed_time = datetime.now() - start_time  # Calcula el tiempo transcurrido

            # Si ha pasado más de 1.5 minutos, intenta detener la ejecución actual y reiniciarla
            if elapsed_time.total_seconds() > 420:
                print("La ejecución ha superado 8 minutos, reiniciando...")
                t.join(timeout=0)  # Intenta detener la ejecución actual
                break  # Sale del bucle interno

        if not t.is_alive():  # Si la ejecución ha finalizado, sale del bucle externo
            break


while counter_value == 0:
    ejecutar_descarga()
    with open("counter.txt","r") as file:
        counter_value = int(file.read().strip())
    if counter_value == 1:
        break
    elif counter_value == 2:
        sys.exit()


try:
    with ZipFile("C:/Users/angel/Downloads/Licitacion_Publicada.zip",'r') as zObject:
        zObject.extractall(path="C:/Users/angel/Desktop/Practica/Licitacion/excel_licitaciones")
except Exception as e:
    print("No se pudo descargar Licitacion_Publicada.zip")
    sys.exit()

print("Se descargaron la lista de licitaciones, buscando nuevas licitaciones...")

filtrado_por_intereses, licitaciones = reordenamiento_de_datos(intereses)

check_licitaciones = "licitaciones_vistas.txt"

lineas = []
with open(check_licitaciones, 'r') as archivo:
    for linea in archivo:
        lineas.append(linea.strip())

nuevas_licitaciones = [licitacion for licitacion in licitaciones if licitacion not in lineas]

counter_value=0
with open("counter.txt", "w") as file:
    file.write(str(counter_value))

if nuevas_licitaciones:
    print(nuevas_licitaciones)
    print('Se hallaron nuevas licitaciones, en breve se enviará un correo con la información de estas.')
    def ejecutar_descarga_2():
        while True:
            start_time = datetime.now()  # Registra el tiempo de inicio de la ejecución

            t = threading.Thread(target=ingresar_licitaciones, args=(nuevas_licitaciones,))
            t.start()

            while t.is_alive():
                elapsed_time = datetime.now() - start_time  # Calcula el tiempo transcurrido

                # Si ha pasado más de 1.5 minutos, intenta detener la ejecución actual y reiniciarla
                if elapsed_time.total_seconds() > 1200:
                    print("La ejecución ha superado 20 minutos, reiniciando...")
                    t.join(timeout=0)  # Intenta detener la ejecución actual
                    break  # Sale del bucle interno

            if not t.is_alive():  # Si la ejecución ha finalizado, sale del bucle externo
                break


    while counter_value == 0:
        ejecutar_descarga_2()
        with open("counter.txt","r") as file:
            counter_value = int(file.read().strip())
        if counter_value == 1:
            break


else:
    print("No se encontraron nuevas licitaciones")
    sys.exit()

directorio = "licitaciones_nuevas"
contenido_archivos = []

if os.path.exists(directorio) and os.path.isdir(directorio):
    archivos = os.listdir(directorio)

    for archivo in archivos:
        if archivo.endswith('.txt'):
            ruta_archivo = os.path.join(directorio, archivo)
            with open(ruta_archivo, "r", encoding="utf-8") as file:
                contenido_archivo = file.read()
                contenido_archivos.append(contenido_archivo)


lista_licitaciones = ", ".join(nuevas_licitaciones[:-1]) + " y " + nuevas_licitaciones[-1]

with open("archivo.txt", 'a') as archivo:
    archivo.write("Las licitaciones encontradas son:\n\n")

with open("archivo.txt", 'a') as archivo:
    archivo.write(lista_licitaciones+"\n\n")

with open("archivo.txt", 'a') as archivo:
    archivo.write("A continuación un breve resumen de cada una:\n\n")

for i in range(len(contenido_archivos)):
    resumen = get_completion("Haz un pequeño resumen de la siguiente licitación: "+contenido_archivos[i])
    with open("archivo.txt", 'a') as archivo:
        archivo.write(resumen+"\n\n")
    time.sleep(30)
#    send_email(resumen, nuevas_licitaciones[i])
    
with open("archivo.txt","a") as archivo:
    archivo.write('Para más información, ingresar a www.MercadoPublico.cl, "Licitaciones" y "Busqueda de Licitaciones para Ofertar"\n Ahí podrá ingresar el ID de la licitación de interés y leer más detalle sobre esta. ')

with open("archivo.txt", 'r') as archivo:
    content = archivo.read()

send_email(content, "Licitaciones encontradas")

with open('licitaciones_vistas.txt','a') as archivo:
    for elemento in nuevas_licitaciones:
        archivo.write('\n'+elemento) 

os.remove("archivo.txt")

exit()

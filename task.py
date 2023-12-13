import defRPAselenium
import moldesTerrenos
import models
import moldesTerrenos
from RPA.Browser.Selenium import Selenium;
import os
import shutil
from shutil import rmtree
from datetime import datetime
import time
import logging
import correo
import openpyxl
import pandas as pd

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s | %(name)s | %(levelname)s | %(message)s',
                    filename= 'log procesos' )

#Calculamos tiempo de ejecucion
tiempoInicio=time.time()
browser = Selenium()
Dt=models.master()
urlbase=str(defRPAselenium.Pyasset(asset="base"))
UrlMacro=defRPAselenium.Pyasset(asset="Ruta ")
libro=defRPAselenium.Pyasset(asset="LIBRO ")

def eliminarcarpetas():
    try:
        rmtree("PDF")
        rmtree("CSV")
        rmtree("Log Scraping")
        rmtree("Formato Solicitud")
        rmtree("Salida")  
        rmtree("Out Hojas Scraping")
        rmtree("Excel")
        os.remove("Correo\\Formato Reporte.xlsx")
        shutil.copy2("Correo\\Formato Reporte\\Formato Reporte.xlsx","Correo\\Formato Reporte.xlsx")
        logging.info("Eliminamos carpetas")


    except:
        pass

def Creacionescarpetas():
    logging.info("Creado las carpetas para PDF's")

    try:
        os.mkdir('PDF')
        os.mkdir('CSV')
        os.mkdir('Excel')
        os.mkdir("Formato Solicitud")  
        os.mkdir("Log Scraping")
        os.mkdir("Salida")
        os.mkdir("Out Hojas Scraping") 

    except:
        pass

def task():     
            for dtable in Dt:
                if dtable[5] == "SI":
                    strcomuna="{} [{}]"
                    strrolmatriz="{}- {}"
                    Rut=str(dtable[0])
                    Inmobiliaria=dtable[1]
                    Asset=dtable[2]
                    Carpeta=dtable[3]
                    Hoja=dtable[4]
                    Activo=dtable[5]
                    region=dtable[6]
                    comuna=strcomuna.format(dtable[7],dtable[10])
                    rolmatriz=strrolmatriz.format(dtable[8],dtable[9])
                    rol1=dtable[8]                               
                    rol2=dtable[9]
                    Codigo=dtable[10]
                    

                    
                    logging.info("creacion de hoja Resumen")
                    logging.info(defRPAselenium.LOGconsulta(region,comuna,rol1,rol2))
                    try:
                     os.remove("Log Scraping\total.txt")
                    except:
                        pass
                    nIntentos=0
                    cantidad=0
                    consulta=True
                    while consulta==True:                       
                        cantidad=1+cantidad
                        estadoConsulta="consulta de terrenos de la region {0} , comuna {1} y rolmatriz {2}-{3} --- consulta # {4} ".format(region,comuna,rol1,rol2,cantidad)
                        logging.info(estadoConsulta)
                        try:
                            nIntentos=+1
                            defRPAselenium.valoresReporte(Rut,Inmobiliaria)
                            defRPAselenium.cerrarinicio()
                            tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                            stado=False
                            consulta=False 
                            roles=str(rol1)+"-"+str(rol2)
                        except:                      
                            consulta=True
                            try:tabla.quit()
                            except:pass
                            roles=str(rol1)+"-"+str(rol2)
                            defRPAselenium.reportarError("Fallo al navegar por el menú de TGR, Contribuciones por ROL - Reintento N°"+str(nIntentos-1),True,roles)
                            nIntentos+=1
    
                    if consulta==stado :
                       consulta=False 
                       
                    logging.info("----------------Diligenciando Formato de solicitud--------------------------- ")
                    try:defRPAselenium.formatosolicitusd(Hoja,Carpeta)
                    except:logging.error("Fallo la funcio formatosolicitusd().")
                    """try:
                        try:
                            defRPAselenium.cerrarinicio()
                            logging.info("----------------Diligenciando resumen--------------------------------------- ")
                            defRPAselenium.diligenciarResumen(Hoja,Carpeta)
                            stado=False
                        except:
                            logging.error("Fallo la funcion diligenciarResumen()")

                        try:
                            logging.info("----------------Diligenciando hojas resumen por sheets de excel--------------- ")
                            defRPAselenium.diligenciarhojas(Hoja,Carpeta,region,comuna,str(rolmatriz),str(Rut),str(Inmobiliaria),str(rol1),str(rol2)) 
                        except:
                            logging.error("Fallo la funcion diligenciarhojas()")

                        logging.info("----------------Ejecutando Macros     -------------------------------------------- ")
                        try:
                            defRPAselenium.Macros(str(Hoja))
                          
                        except:
                            logging.error("Fallo la funcio Macros()")

                        logging.info("----------------Salidas de excel -------------------------------------------- ")
                       

                        
                        logging.info("----------------Diligenciando Formato de solicitud--------------------------- ")
                        try:
                            defRPAselenium.formatosolicitusd(Hoja,Carpeta)
                        except:
                            logging.error("Fallo la funcio formatosolicitusd().")

                        logging.info("----------------TEST TOTALES ------------------------------------------------ ")
                        
                        try:
                                moldesTerrenos.logscraping(Carpeta,str(rolmatriz))
                        except:
                                logging.error("Fallo la funcio logscraping()  .")
                        
                            
                        try: moldesTerrenos.lecturaALL(Carpeta,Rut,region,comuna,rolmatriz,Inmobiliaria,Hoja)
                        except:logging.error("Fallo la funcio lecturaALL().")
                       

                    except:
                            pass"""
                
def tgc():
   task()  

def enviandoCorreo(fecha_formateada_inicio, fecha_formateada_final):
    archivo_excel = "Correo/Formato Reporte.xlsx"
    nombre_hoja = 'Errores'
    df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
    if not df.empty:
        print("Enviando Correo con errores")
        correo.enviarCorreo(fecha_formateada_inicio, fecha_formateada_final,True)
    else:
        print("Enviando Correo sin errores")
        libro_excel=openpyxl.load_workbook(archivo_excel)
        hoja_a_borrar=libro_excel["Errores"]
        libro_excel.remove(hoja_a_borrar)
        libro_excel.save(archivo_excel)
        correo.enviarCorreo(fecha_formateada_inicio, fecha_formateada_final,False)

def procesoCompletado():
    import tkinter as tk
    from tkinter import messagebox

# Crear una ventana principal
    root = tk.Tk()
    root.title("Ejemplo Tkinter")

    # Establecer el tamaño de la ventana
    root.geometry("300x200")

# Establecer la ubicación de la ventana en la pantalla (posiciónX, posiciónY)
    root.geometry("+500+300")  # Ajusta estas coordenadas según tus necesidades
    # Hacer que la ventana esté siempre en la parte superior
    root.attributes("-topmost", True)

    # Función para mostrar el cuadro de mensaje
    def mostrar_mensaje():
        messagebox.showinfo(message="Ejecución Finalizada", title="Ejecución Bot Unidades")

    # Botón para mostrar el mensaje
    boton_mostrar = tk.Button(root, text="Mostrar Mensaje", command=mostrar_mensaje)
    boton_mostrar.pack(pady=20)

# Iniciar el bucle principal de Tkinter
    root.mainloop()

if __name__ == "__main__":
   fecha_inicio=datetime.now()
   fecha_formateada_inicio = fecha_inicio.strftime('%d/%m/%Y %H:%M:%S')
   eliminarcarpetas()
   Creacionescarpetas()
   moldesTerrenos.creacionExcelResumen()
   defRPAselenium.bakup()
   tgc()
   moldesTerrenos.sumatorias()
   moldesTerrenos.solicitud()
   logging.info('Ejecucion finalizada')
   tiempoFinal=time.time() 
   TiempoTotal=tiempoFinal-tiempoInicio
   fecha_final=datetime.now()
   fecha_formateada_final = fecha_final.strftime('%d/%m/%Y %H:%M:%S')
   enviandoCorreo(fecha_formateada_inicio, fecha_formateada_final)
   print("Tiempo total de ejecucion es " + str(TiempoTotal) + " seg")
   procesoCompletado()

from RPA.Excel.Files import Files;
import os
from shutil import rmtree
import csv
import openpyxl
from RPA.Tables import Tables
import shutil
from datetime import date
from datetime import datetime
import defRPAselenium



library = Tables()
lib = Files()

listMster= [{'RUT':'',
                   'Inmobiliaria':'',
                   'Asset':'',
                   'Carpeta':'',
                   'Hoja':'',
                   'Activo':'',
                   'Region':'',
                   'Comuna':'',
                   'RolMatriz':'',
                   'Rol':'',
                   'Codigo':'',
                   'status':''

             }]


listSCRAPIADO= ([{
               'CUOTA':"",
               'NRO FOLIO':"", 
               'VALOR':"",
               'VENCIMIENTO':"",
                'TOTA A PAGAR':"",
                 }])

def Dt_BaseTerreno():
    lib.open_workbook("")      #ubicacion del libro
    lib.read_worksheet("Base existencia")       #nombre de la hoja
    lista=lib.read_worksheet_as_table(name='Base existencia',header=True, start=1).data
    return lista
 
def LogSCraping():
    
    contenido = os.listdir('Log Scraping/')
    for dataTxt in contenido:
        mensaje = open('Log Scraping/'+dataTxt, "r")
        outmensaje=mensaje.read()
        outmensaje=outmensaje.replace("VALOR"," ")
        outmensaje=outmensaje.replace("CUOTA"," ")
        outmensaje=outmensaje.replace("VALOR CUOTA"," " )
        outmensaje=outmensaje.replace("NRO FOLIO"," " )
        outmensaje=outmensaje.replace("VENCIMIENTO"," " )
        outmensaje=outmensaje.replace("TOTAL A PAGAR"," " )
        outmensaje=outmensaje.replace("EMAIL"," " )
        outmensaje=outmensaje.replace("DESCARGAR"," " )
        outmensaje=outmensaje.replace("""CUOTA
                                            VALOR CUOTA
                                            NRO FOLIO
                                            VENCIMIENTO
                                            TOTAL A PAGAR
                                            EMAIL
                                            DESCARGAR"""," " )
        return(contenido)

def master():
   lib.open_workbook("Data\Master.xlsx")      #ubicacion del libro
   lib.read_worksheet("Listado")       #nombre de la hoja
   DtMaster=lib.read_worksheet_as_table(name='Listado',header=False, start=1).data

   return DtMaster

def txttocsv():

   contenido = os.listdir('Log Scraping/')
   for dataTxt in contenido:
        txt=dataTxt.replace(".txt",".csv")

        txt_file =r"Log Scraping/" + dataTxt
        csv_file =r"CSV/" + txt 

        in_txt = csv.reader(open(txt_file, "r"), delimiter = " ")
        out_csv = csv.writer(open(csv_file, 'w'))

        out_csv.writerows(in_txt)

        del out_csv

def ReadCSV():
    with open('CSV/94-76182178-4-Inversiones World Logistic.csv', newline='') as f:
     reader = csv.reader(f,delimiter=",")
     for row in reader:
         print(row[0])
       # print(f"CUOTA:{0},VALOR CUOTA:{1},NRO FOLIO:{2},VENCIMIENTO:{3},TOTAL A PAGAR:{4},EMAIL:{5}".format(row[0],row[1],row[2],row[3],row[4],row[5]))      

def txttest(f):

  scrp=[]
  for x in f:
       if x.find(" ") != 0:
           scrp.append(x)
  liscon=[]

  print(scrp.index)
  
  for u in scrp:
      final=u.find(" ")
      largo=len(u)

      Sumatoria=0

      
      CUOTA=str(u)[0:final]
      Sumatoria=Sumatoria+len(CUOTA)+1

      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ") 
      VALOR=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(VALOR)+1
      

      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ")
      NRO_FOLIO=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(NRO_FOLIO)+1
      


      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ")
      VENCIMIENTO=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(VENCIMIENTO)+1
    

      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ")
      TOTAPAGAR=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(TOTAPAGAR)+1

              
      listSCRAPIADO.append({
               'CUOTA':CUOTA,
               'NRO FOLIO':NRO_FOLIO, 
               'VALOR':VALOR,
               'VENCIMIENTO':VENCIMIENTO,
               'TOTA A PAGAR':TOTAPAGAR,
                 },
      )
     
               
  return listSCRAPIADO

def cambionombre(origen, destino):
    archivo = origen
    nombre_nuevo = destino
    os.rename(archivo, nombre_nuevo)

def FormatoSolicitud(h,CUOTA, valor):
    LibroMaster=master()

    for dtMaster in LibroMaster:
                idrut=str(dtMaster[0])
                Inmobiliaria=dtMaster[1]
                Asset=dtMaster[2]
                Carpeta=dtMaster[3]
                Hoja=dtMaster[4]
                Activo=dtMaster[5]
                region=dtMaster[6]
                rol1=dtMaster[8]                               
                rol2=dtMaster[9]
                Codigo=dtMaster[10]

                if Hoja==h:
                 Rut=idrut
                 origen='Data\\Formato Solicitud Pago.xlsx'
                 destino="Formato Solicitud\\"+Carpeta +" " +" Monto " + str(valor) + "Formato Solicitud Pago.xlsx"
                 #Copias el libro de formato de solicitud
                 shutil.copy(origen,destino )
                 lib.open_workbook(destino)    #ubicacion del libro
                 lib.read_worksheet("Hoja1")                                              #nombre de la hoja
                 formato=lib.read_worksheet_as_table(name='Hoja1',header=True, start=1).data

                 break
                else:
                    pass
            
 #formato de resumen 
    for rem in Resumen():
        N=rem[0]
        if N==h:
           NombreSolicitante=str(N[7])
           now = datetime.now()
           gerente=defRPAselenium.Pyasset(asset="Gerente")
           InmobiliariaGiradora=str(N[3])
           Monto=str(valor)
           Rut=str(N[2])
           RUTtesoria=defRPAselenium.Pyasset(asset="RutTesoreria")
           Dirección=defRPAselenium.Pyasset(asset="Dirección")
           Glosagasto="Pago de sobretasa cuota "+str(CUOTA)
           Detallegasto= Glosagasto
           CentroGestion=str(N[12])

           lib.set_cell_value(8,"D",str(NombreSolicitante))
           lib.set_cell_value(6,"H",str(now))
           lib.set_cell_value(10,"D",str(gerente))
           lib.set_cell_value(12,"D",str(InmobiliariaGiradora))
           lib.set_cell_value(12,"H",str(Rut))
           lib.set_cell_value(14,"C",str(Monto))
           lib.set_cell_value(14,"C",str(RUTtesoria))
           lib.set_cell_value(20,"C",str(Dirección))
           lib.set_cell_value(22,"D",str(Glosagasto))
           lib.set_cell_value(24,"D",str(Detallegasto))
           lib.set_cell_value(28,"D",str(CentroGestion))
           lib.set_cell_value(30,"D","Contribuciones")
    
    lib.save_workbook()

def Resumen():
    lib.open_workbook("Data/Resumen_Contribuciones_Terreno_2023.xlsx")      #ubicacion del libro
    lib.read_worksheet("Resumen")                                              #nombre de la hoja
    dtresumen=lib.read_worksheet_as_table(name='Resumen',header=True, start=1).data
    return dtresumen

def ejemplo(CUOTA,valor):
     for rem in Resumen():
         
        lib.open_workbook("Formato Solicitud\\6-2028  Monto 76345236 Formato Solicitud Pago.xlsx")        #ubicacion del libro
        lib.read_worksheet("Solicitud")                                                                   #nombre de la hoja                                          
        formato=lib.read_worksheet_as_table(name='Solicitud',header=True, start=1).data

        NombreSolicitante=str(rem[8])
        print(NombreSolicitante)
        now = datetime.now()
        print(now)
        gerente=defRPAselenium.Pyasset(asset="Gerente")
        print(gerente)
        InmobiliariaGiradora=str(rem[3])
        print(InmobiliariaGiradora)
        Monto=str(valor)
        print(Monto)
        Rut=str(rem[2])
        print(Rut)
        gerente=defRPAselenium.Pyasset(asset="Gerente")
        RUTtesoria=defRPAselenium.Pyasset(asset="RutTesoreria")
        print(RUTtesoria)
        Dirección=defRPAselenium.Pyasset(asset="Dirección")
        print(Dirección)
        Glosagasto="Pago de sobretasa cuota "+str(CUOTA)
        print(Glosagasto)
        Detallegasto= Glosagasto
        print(Detallegasto)
        CentroGestion=str(rem[12])
        print(CentroGestion)




        lib.set_cell_value(8,"D",str(NombreSolicitante))
        lib.set_cell_value(6,"H",str(now))
        lib.set_cell_value(10,"D",str(gerente))
        lib.set_cell_value(12,"D",str(InmobiliariaGiradora))
        lib.set_cell_value(12,"H",str(Rut))
        lib.set_cell_value(14,"C",str(Monto))
        lib.set_cell_value(14,"C",str(RUTtesoria))
        lib.set_cell_value(20,"C",str(Dirección))
        lib.set_cell_value(22,"D",str(Glosagasto))
        lib.set_cell_value(24,"D",str(Detallegasto))
        lib.set_cell_value(28,"D",str(CentroGestion))
        lib.set_cell_value(30,"D","Contribuciones")
        lib.save_workbook()





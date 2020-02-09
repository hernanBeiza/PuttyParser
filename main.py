import xlrd
import datetime
from sys import argv
from datetime import time

def main(argv):
  print("Main")
  print("Para parsear Excel del Putty")
  #print(argv)
  script,archivoXLXS,salidaCSV = argv
  print(archivoXLXS)
  print(salidaCSV)
  limpiarSalida(salidaCSV)
  leerExcel(archivoXLXS, salidaCSV)

def leerExcel(archivoXLXS, salidaCSV):
  print("Leyendo exxcel" + archivoXLXS)
  workbook = xlrd.open_workbook(archivoXLXS)
  for hoja in workbook.sheets():
    print ('Hoja:',hoja.name)
    print ('Rows:',hoja.nrows)
    print ('Cols:',hoja.ncols)

    #TODO ver como solucionar para más de un mes por hoja
    mes = hoja.cell(0,2).value
    dia = ''
    barra = ''
    hora = ''
    valor = ''

    print('Mes:', mes)

    dias = []
    for col in range(hoja.ncols):
      if(col>1 and col<7):
        dias.append(hoja.cell(1,col).value)

    print ('Dias:', dias)

    barras = []
    for row in range(hoja.nrows):
      if(row>1):
        if(hoja.cell(row,0).value!=""):
          barras.append(hoja.cell(row,0).value)

    print ('Barras:', barras)

    horas = []
    for row in range(hoja.nrows):
      if(row>1):
        if(hoja.cell(row,1).value!=""):
          horaFloat = hoja.cell(row,1).value
          fechaCompleta = xlrd.xldate_as_datetime(horaFloat, workbook.datemode)
          #print(fechaCompleta.time())
          #print(fechaCompleta.time().strftime("%H:%M:%S"))
          horas.append(fechaCompleta.time().strftime("%H:%M:%S"))

          #print (horaTuple[3],horaTuple[4],horaTuple[5])
          """
          horaTuple = xlrd.xldate_as_tuple(horaFloat, workbook.datemode)
          print(horaTuple)
          hora = time(*horaTuple[3:])
          horaString = str(horaTuple[3])+":"+str(horaTuple[4])+":"+str(horaTuple[5])
          horas.append(horaString)
          """

    print ('Horas:', horas)

    datos = []
    datosMulti = []
    for row in range(hoja.nrows):
      datosPorHora = []
      for col in range(hoja.ncols):
        if(row>1 and col > 1):
          dato = int(hoja.cell(row,col).value)
          datosPorHora.append(dato)
          datos.append(dato);
      if(row>1 and col > 1):
        datosMulti.append(datosPorHora)

    print('Datos:', datos)


    print ('Salida')
    for barra in barras:
      for i in range(len(horas)):
        for j in range(len(dias)):
          print(barra,dias[j],horas[i],datosMulti[i][j])
          linea = str(mes)+","+str(barra)+","+str(dias[j])+","+str(horas[i])+","+str(datosMulti[i][j])
          guardarLineaEnCSV(salidaCSV,linea)

def guardarLineaEnCSV(archivoCSV,lineaString):
  print("Guardando línea en csv"+archivoCSV)
  with open(archivoCSV, 'a') as archivoCSV:
    archivoCSV.write(lineaString+"\n")

def limpiarSalida(archivoCSV):
  print("Borrando archivo "+archivoCSV)
  try:
    raw = open(archivoCSV, "r+")
    contents = raw.read().split("\n")
    raw.seek(0)                        # <- This is the missing piece
    raw.truncate()
  except IOError:
    print ("Error: Archivo "+archivoCSV+ " no existe.")

if __name__ == '__main__':
  if len(argv) < 2:
    print(argv);
    #argv por defecto
    #argv = ['script', '1', '2', '3']
    print("Falta especificar ruta de archivo.xlsx");
  else:
    main(argv)
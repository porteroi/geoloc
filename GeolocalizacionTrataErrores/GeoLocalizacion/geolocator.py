import requests
import xlwt
import xlrd
import copy
import csv
import sys
import time
####Funcion de geolocalizacion de direcciones fisicas. Se le pasa una direccion como parametro y 
## se devuelve las coordenadas en el mapa.
def geoloc(direccion):
	params={'address':direccion}
	cont=0
	location=[]
	coordenadas=[]
	while True:
		r=requests.get(url,params=params)
		results=r.json()['results']
		if(r.json()['status']=='OVER_QUERY_LIMIT'):
			print("LIMITE DE DIRECCIONES ALCANZADO")
			exit()
			
		
		print(results)
		cont=cont+1
		if(len(results)!=0 or cont==10):
			break
	
	try:
		con=0
		while True:
			#Este if filtra para las ubicaciones que hay en ESPAÑA
			if("Spain" in results[con]['formatted_address']):
				
				location=results[con]['geometry']['location']
				#print("Cogiendo la location") 
				if(len(location)!=0):
					break
			con=con+1
	except IndexError:
		pass
	
	#print(location['lat'],location['lng'])
	
	try:
		while True:
			coordenadasX = (location['lat'])
			coordenadasY =(location['lng'])
			#print("Pillando coordenadas")
			coordenadas.append(coordenadasX)
			coordenadas.append(coordenadasY)
			print(coordenadas)
			if(len(coordenadas)!=0):
				#print("Pasando por el break")
				break
	except TypeError:
		coordX.append(0)
		coordY.append(0)
		pass
		
	
	cont=0
	#print(coordenadas)
	
	return coordenadas
	

##Comienzo del programa principal donde se cogen los datos de las empresas del csv, se utilizan en la 
##funcion y se vuelve a escribir en otro csv.
url='https://maps.googleapis.com/maps/api/geocode/json'
wb1= xlrd.open_workbook('Libro1.xlsx')
copia= copy.copy(wb1)
coordX=[]
coordY=[]
hoja=copia.sheet_by_name('Hoja1')
direcciones=[]
contador=0
cuentaCoord=0
with open ("empresasCoordenadas.csv","w+",newline='') as empresasf:
	for i in range(1,hoja.nrows):
		coord=[]
		var=hoja.cell(i,5).value + "," + hoja.cell(i,6).value + " " + (hoja.cell(i,7).value)
		direcciones.append(hoja.cell(i,5).value + "," + hoja.cell(i,6).value + " " + (hoja.cell(i,7).value))
		
		try:
			while(coord==[] and contador<=7):
				contador=contador+1
				coord=geoloc(var)
				coordX.append(coord[0])
				coordY.append(coord[1])
			contador=0
		except IndexError:
			pass
			
		
		
		
		
		#direcciones.append(coord) 
		
		
		
		
		#print(direcciones)
	writer= csv.writer(empresasf,lineterminator='\n',delimiter=' ' ,quotechar=' ', quoting=csv.QUOTE_MINIMAL)
	writer.writerow(['Direccion,','Latitud,','Longitud'])
	for val in direcciones:
		#print(coordX)
		#print(coordY)
		val=str(val)
		val=val.replace(",","").replace("\t","").replace(";","").replace('\n','').replace('\v','').replace('\r','').replace(',','')
		val= val+ " "
		## ESTO PUEDE SER PARA HACER LAS COLUMNAS writer.writerow['Direccion', 'Coordenadas']
		writer.writerow([val,',',coordX[cuentaCoord],',',coordY[cuentaCoord]]) #AÑADIMOS ,coord DENTRO DE LOS CORCHETES
		cuentaCoord=cuentaCoord+1
	del writer
	empresasf.close()

	
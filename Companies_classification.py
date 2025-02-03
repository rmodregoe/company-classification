'''El objetivo de este programa es crear dos archivos excel:
-El primero se llama Listado_DDMMAAAA_HHMMSS y contiene un listado descargado de la página de El Economista que contiene las empresas españolas con más de 50M 
de facturación con datos de cada empresa referentes a Posición Nacional, Evolución Posiciones, Nombre de la empresa, Provincia,	Sector y Facturación (€).

-El segundo se llama Análisis_empresas_DDMMAAAA_HHMMSS y, a partir del excel anterior, clasifica estas empresas por sectores menos específicos y más accesibles para analizar.
El archivo Excel ontiene tres hojas:
    -Sheet1: contiene las empresas españolas con más de 50M de facturación con datos de cada empresa referentes a Posición Nacional, Evolución Posiciones, Nombre de la empresa, Provincia,	Sector , Facturación (€)
    y Grupo, que es la clasificación por sectores más útil y accesible.
    -Suma Facturación: contiene una tabla con distintos los sectores y la facturación de cada uno, es decir, la suma de la facturación de todas las empresas de cada sector.
    -Empresas Sector: contiene una tabla con distintos los sectores y el número de empresas de cada uno, es decir, la suma de todas las empresas de cada sector.

Tarda alrededor de 10 minutos en ejecutarse
'''


import datetime
import os
import re
import string
import time
from tkinter import StringVar

import numpy as np
import openpyxl
import pandas as pd
import xlsxwriter
from playwright.sync_api import Error, sync_playwright

#Obtener fecha de hoy para poner nombre al archivo
fechaHoy = datetime.datetime.now()
fechaHoyStr = fechaHoy.strftime("%d%m%Y_%H%M%S")
nombreExcel = f"Listado_{fechaHoyStr}.xlsx"

workbook = xlsxwriter.Workbook(f"./{nombreExcel}")
workbook.close()

i=0
l=0
s=0
v=[1,2,3,4,5,6]
h=[0]
headerPuesto = False

with sync_playwright() as p:
    with pd.ExcelWriter(nombreExcel,engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        navegador = p.chromium.launch(timeout= 20000, headless=False)
        context = navegador.new_context()
        browser = context.new_page()
        browser.goto("https://ranking-empresas.eleconomista.es/ranking_empresas_nacional.html?qVentasNorm=corporativas")
        browser.set_default_timeout(60000)
        browser.wait_for_timeout(2000)
        browser.click('#didomi-notice-agree-button')       
        browser.wait_for_timeout(3000)
        Resultados = browser.inner_text("#tabla-ranking > h2")
        result= re.search('\((.* )',Resultados)
        result = str(result.group(1)) #número de empresas bajo el filtro del sector utilizado
        result = result.strip() #quitar el espacio final

        if len(result)-1 == 4: ##Como hay 7.634 Resultados, siempre va a entrar a este bucle, pero por si cambia el número, está preparado para cualquier caso
            d= slice(30,31) #Primera cifra
            e= slice(32,35) #De la tercera a la sexta, se comería el punto, debería ser (33,36)

            Result = Resultados[d] + Resultados[e] #concatena las cadenas con la posición del número de resultados
            NumResultados=int(Result)
            UltimoDig = NumResultados % 100 #El resto
            NumPaginas = int(NumResultados/100) #Si tengo 1543 pues el número de centenas-->15+1
            NumPaginas = NumPaginas + 1 #Por eso le suma 1 luego 
            if UltimoDig == 0:  #A no ser que ult digito=0 (decenas y unidades)
                NumPaginas = int(NumResultados/100)

        elif len(result)-1 > 6: #filtra según la longitud de la cadena (que es el nº de cifras de resultados + el punto,
                #si tiene más de 6 cifras que repita porque lo máximo son 510.000)
                browser.goto("https://ranking-empresas.eleconomista.es/ranking_empresas_nacional.html?qVentasNorm=corporativas")
                browser.wait_for_timeout(3000)
                Resultados = browser.inner_text("#tabla-ranking > h2")
                result= re.search('\((.* )',Resultados)
                result = str(result.group(1))
                result = result.strip()


        elif len(result)-1 == 6: #Si tiene 6 cifras: 100.000-->Tiene punto. 
            a= slice(30,33) #3 primeras cifras
            b= slice(34,37) #De la cuarta a la sexta cifras

            Result = Resultados[a] + Resultados[b] #concatena las cadenas con la posición del número de resultados. 100000
            NumResultados=int(Result)
            NumPaginas = int(NumResultados/100) #Si tengo 143250, pues el número de centenas-->1432+1
            UltimoDig = NumResultados % 100 #El resto 
            NumPaginas = NumPaginas + 1
            if UltimoDig == 0:
                NumPaginas = int(NumResultados/100)

        elif len(result)-1 == 5: #Si tiene 5 cifras: 10.000-->Tiene punto. No hay más en ningún sector
            a= slice(30,32) #Dos primeras cifras
            b= slice(33,36) #De la tercera a la sexta, se comería el punto, debería ser (33,36)

            Result = Resultados[a] + Resultados[b] #concatena las cadenas con la posición del número de resultados. 10000
            NumResultados=int(Result)
            UltimoDig = NumResultados % 100 #El resto 
            NumPaginas = int(NumResultados/100) #Si tengo 14325, pues el número de centenas-->143+1
            NumPaginas = NumPaginas + 1
            if UltimoDig == 0:
                NumPaginas = int(NumResultados/100)



        elif len(result)==3 : #no tiene punto 845
            rangoResultados = slice(30,33) #coge los índices 30,31 y 32: las 3 cifras
            NumResultados = int(Resultados[rangoResultados])
            UltimoDig = NumResultados % 100 #El resto
            NumPaginas = int(NumResultados/100) #redondea a la baja la división entre 100 para quedarte con las centenas
            NumPaginas = NumPaginas + 1 #Por eso le suma 1 luego 
            if UltimoDig == 0:  #A no ser que ult digito=0 (decenas y unidades)
                NumPaginas = int(NumResultados/100)


        else: #No tiene punto, 93
            NumPaginas = 1
            NumResultados = int(result)

        


        m=0
        A=np.ones((NumResultados+3*NumPaginas,6),dtype=StringVar)
        while l < NumPaginas:
            numeroFilas = browser.locator("#tabla-ranking > table > tbody > tr").count()

            numeroFilas = numeroFilas-1
            
            l=l+1
            i=0
            while i <= numeroFilas:
                i=i+1
                m=m+1
                for j in v:
                    valorCelda = browser.inner_text(f"#tabla-ranking > table > tbody > tr:nth-child({i}) > td:nth-child({j})")
                    if valorCelda=="":
                        for k in range(6):
                            A[m-1, k] = ""
                        break        
                    A[m-1,j-1]=valorCelda                   

                    if j == 4 and (int(valorCelda.replace('.', '')) < 50000000):
                            break  
                    
                if j == 4 and (int(valorCelda.replace('.', '')) < 50000000):
                            break  
                
            if j == 4 and (int(valorCelda.replace('.', '')) < 50000000):
                            break  
            
        
            if l< NumPaginas:
                browser.click("#tabla-ranking >> a:has-text('»')")
                browser.wait_for_timeout(3000)
        
        
        df = pd.DataFrame(A, columns = ['Posición Nacional', 'Evolución de posiciones', 'Nombre de la empresa', 'Facturación', 'Sector Actividad', 'Provincia' ])
        df.replace("",np.nan,inplace=True)
        df.dropna(inplace=True)
        df.replace(1,np.nan,inplace=True)
        df.dropna(inplace=True)
        df['Facturación'] = df['Facturación'].replace(['corporativa', 'grande', 'mediana', 'pequeña'], np.nan)  # Reemplaza los valores por NaN
        df.dropna(inplace=True)
        df['Facturación'] = df['Facturación'].str.replace('.', '')  # Elimina los puntos de los valores
        df['Facturación'] = df['Facturación'].astype(float)  # Convierte la columna "Facturación" a tipo float
        df['Sector Actividad'] = df['Sector Actividad'].astype(str)
        df['Sector Actividad'] = df['Sector Actividad'].str.zfill(4) #introduce un cero al principio de los sectores con 3 códigos

        df.reset_index(drop=True)
        if headerPuesto == False:
            df.to_excel(writer,header=True,index=False,startrow=sum(h))
            headerPuesto = True
        else:
            df.to_excel(writer,header=False,index=False,startrow=sum(h)+1)

        h.append(len(df.index))        
        browser.wait_for_timeout(1500)

''' Generación del análisis a partir del listado de empresas'''

df = pd.read_excel(f'{nombreExcel}')

sectores = {
    "01": "Agricultura, ganadería, caza y servicios relacionados con las mismas",
    "02": "Silvicultura y explotación forestal",
    "03": "Pesca y acuicultura",
    "05": "Extracción de antracita, hulla y lignito",
    "06": "Extracción de crudo de petróleo y gas natural",
    "07": "Extracción de minerales metálicos",
    "08": "Otras industrias extractivas",
    "09": "Actividades de apoyo a las industrias extractivas",
    "10": "Industria de la alimentación",
    "11": "Fabricación de bebidas",
    "12": "Industria del tabaco",
    "13": "Industria textil",
    "14": "Confección de prendas de vestir",
    "15": "Industria del cuero y del calzado",
    "16": "Industria de la madera y del corcho",
    "17": "Industria del papel",
    "18": "Artes gráficas y reproducción de soportes grabados",
    "19": "Coquerías y refino de petróleo",
    "20": "Industria química",
    "21": "Fabricación de productos farmacéuticos",
    "22": "Fabricación de productos de caucho y plásticos",
    "23": "Fabricación de otros productos minerales no metálicos",
    "24": "Metalurgia; fabricación de productos de hierro, acero y ferroaleaciones",
    "25": "Fabricación de productos metálicos, excepto maquinaria y equipo",
    "26": "Fabricación de productos informáticos, electrónicos y ópticos",
    "27": "Fabricación de material y equipo eléctrico",
    "28": "Fabricación de maquinaria y equipo n.c.o.p.",
    "29": "Fabricación de vehículos de motor, remolques y semirremolques",
    "30": "Fabricación de otro material de transporte",
    "31": "Fabricación de muebles",
    "32": "Otras industrias manufactureras",
    "33": "Reparación e instalación de maquinaria y equipo",
    "35": "Suministro de energía eléctrica, gas, vapor y aire acondicionado",
    "36": "Captación, depuración y distribución de agua",
    "37": "Recogida y tratamiento de aguas residuales",
    "38": "Recogida, tratamiento y eliminación de residuos; valorización",
    "39": "Actividades de descontaminación y otros servicios de gestión de residuos",
    "41": "Construcción de edificios",
    "42": "Ingeniería civil",
    "43": "Actividades de construcción especializada",
    "45": "Venta y reparación de vehículos de motor y motocicletas",
    "46": "Comercio al por mayor e intermediarios del comercio",
    "47": "Comercio al por menor",
    "49": "Transporte terrestre y por tubería",
    "50": "Transporte marítimo y por vías navegables interiores",
    "51": "Transporte aéreo",
    "52": "Almacenamiento y actividades anexas al transporte",
    "53": "Actividades postales y de correos",
    "55": "Servicios de alojamiento",
    "56": "Servicios de comidas y bebidas",
    "58": "Edición",
    "59": "Actividades cinematográficas, de vídeo y de programas de televisión, grabación de sonido y edición musical",
    "60": "Actividades de programación y emisión de radio y televisión",
    "61": "Telecomunicaciones",
    "62": "Programación, consultoría y otras actividades relacionadas con la informática",
    "63": "Servicios de información",
    "64": "Servicios financieros, excepto seguros y fondos de pensiones",
    "65": "Seguros, reaseguros y fondos de pensiones, excepto Seguridad Social obligatoria",
    "66": "Actividades auxiliares a los servicios financieros y a los seguros",
    "68": "Actividades inmobiliarias",
    "69": "Actividades jurídicas y de contabilidad",
    "70": "Actividades de las sedes centrales; actividades de consultoría de gestión empresarial",
    "71": "Servicios técnicos de arquitectura e ingeniería; ensayos y análisis técnicos",
    "72": "Investigación y desarrollo",
    "73": "Publicidad y estudios de mercado",
    "74": "Otras actividades profesionales, científicas y técnicas",
    "75": "Actividades veterinarias",
    "77": "Actividades de alquiler",
    "78": "Actividades relacionadas con el empleo",
    "79": "Actividades de agencias de viajes, operadores turísticos, servicios de reservas y actividades relacionadas con los mismos",
    "80": "Actividades de seguridad e investigación",
    "81": "Servicios a edificios y actividades de jardinería",
    "82": "Actividades administrativas de oficina y otras actividades auxiliares a las empresas",
    "84": "Administración Pública y defensa; Seguridad Social obligatoria",
    "85": "Educación",
    "86": "Actividades sanitarias",
    "87": "Asistencia en establecimientos residenciales",
    "88": "Actividades de servicios sociales sin alojamiento",
    "90": "Actividades de creación, artísticas y espectáculos",
    "91": "Actividades de bibliotecas, archivos, museos y otras actividades culturales",
    "92": "Actividades de juegos de azar y apuestas",
    "93": "Actividades deportivas, recreativas y de entretenimiento",
    "94": "Actividades asociativas",
    "95": "Reparación de ordenadores, efectos personales y artículos de uso doméstico",
    "96": "Otros servicios personales",
    "97": "Actividades de los hogares como empleadores de personal doméstico",
    "99": "Actividades de organizaciones y organismos extraterritoriales"
}

# Crear una nueva columna 'Grupo' con los dos primeros dígitos del código de sector
df['Sector Actividad'] = df['Sector Actividad'].astype(str)
df['Sector Actividad'] = df['Sector Actividad'].str.zfill(4)
df['Grupo'] = df['Sector Actividad'].str[:2]

# Mapear los códigos de sector con los nombres correspondientes
df['Grupo'] = df['Grupo'].map(sectores)

# Agrupar por sector y contar el número de empresas en cada sector
empresas_por_sector = df.groupby('Grupo')['Nombre de la empresa'].count()
suma_facturacion_por_sector = df.groupby('Grupo')['Facturación'].sum()


# Crear un diccionario con los grupos y las actividades correspondientes
grupos_actividades = {
    "Energy and utilities": ["Suministro de energía eléctrica, gas, vapor y aire acondicionado",
                             "Captación, depuración y distribución de agua",
                             "Recogida y tratamiento de aguas residuales",
                             "Recogida, tratamiento y eliminación de residuos; valorización",
                             "Actividades de descontaminación y otros servicios de gestión de residuos"],
    "Renewables": ["Energías renovables y servicios relacionados"],
    "Telco & Media": ["Telecomunicaciones",
                      "Actividades cinematográficas, de vídeo y de programas de televisión, grabación de sonido y edición musical",
                      "Edición",
                      "Actividades de programación y emisión de radio y televisión"],
    "Public Sector": ["Administración Pública y defensa; Seguridad Social obligatoria",
                      "Actividades de bibliotecas, archivos, museos y otras actividades culturales"],
    "Real State": ["Actividades inmobiliarias"],
    "Infra & Construction": ["Construcción de edificios",
                             "Ingeniería civil",
                             "Actividades de construcción especializada"],
    "Industry & Logistics": ["Agricultura, ganadería, caza y servicios relacionados con las mismas",
                             "Silvicultura y explotación forestal",
                             "Pesca y acuicultura",
                             "Extracción de antracita, hulla y lignito",
                             "Extracción de crudo de petróleo y gas natural",
                             "Extracción de minerales metálicos",
                             "Otras industrias extractivas",
                             "Actividades de apoyo a las industrias extractivas",
                             "Industria de la alimentación",
                             "Fabricación de bebidas",
                             "Industria del tabaco",
                             "Industria textil",
                             "Confección de prendas de vestir",
                             "Industria del cuero y del calzado",
                             "Industria de la madera y del corcho",
                             "Industria del papel",
                             "Artes gráficas y reproducción de soportes grabados",
                             "Fabricación de productos de caucho y plásticos",
                             "Fabricación de otros productos minerales no metálicos",
                             "Metalurgia; fabricación de productos de hierro, acero y ferroaleaciones",
                             "Fabricación de productos metálicos, excepto maquinaria y equipo",
                             "Fabricación de productos informáticos, electrónicos y ópticos",
                             "Fabricación de material y equipo eléctrico",
                             "Fabricación de maquinaria y equipo n.c.o.p.",
                             "Fabricación de vehículos de motor, remolques y semirremolques",
                             "Fabricación de otro material de transporte",
                             "Fabricación de muebles",
                             "Otras industrias manufactureras",
                             "Reparación e instalación de maquinaria y equipo",
                             "Actividades postales y de correos"],
    "Health Services": ["Actividades sanitarias",
                        "Actividades veterinarias"],
    "Life Sciences & Chemical": ["Investigación y desarrollo",
                                 "Otras actividades profesionales, científicas y técnicas",
                                 "Coquerías y refino de petróleo",
                                 "Industria química",
                                 "Fabricación de productos farmacéuticos"],
    "CPG & Retail": ["Comercio al por mayor e intermediarios del comercio",
                     "Comercio al por menor",
                     "Reparación de ordenadores, efectos personales y artículos de uso doméstico"],
    "Banking & Insurance": ["Servicios financieros, excepto seguros y fondos de pensiones",
                            "Seguros, reaseguros y fondos de pensiones, excepto Seguridad Social obligatoria",
                            "Actividades auxiliares a los servicios financieros y a los seguros"],
    "Services": ["Actividades de alquiler",
                 "Actividades relacionadas con el empleo",
                 "Actividades de agencias de viajes, operadores turísticos, servicios de reservas y actividades relacionadas con los mismos",
                 "Actividades de seguridad e investigación",
                 "Servicios a edificios y actividades de jardinería",
                 "Actividades administrativas de oficina y otras actividades auxiliares a las empresas",
                 "Actividades de los hogares como empleadores de personal doméstico",
                 "Actividades de organizaciones y organismos extraterritoriales",
                 "Servicios de comidas y bebidas",
                 "Actividades de servicios sociales sin alojamiento",
                 "Actividades de creación, artísticas y espectáculos",
                 "Actividades asociativas",
                 "Actividades de juegos de azar y apuestas",
                 "Actividades de las sedes centrales; actividades de consultoría de gestión empresarial",
                 "Otros servicios personales",
                 "Programación, consultoría y otras actividades relacionadas con la informática",
                 "Publicidad y estudios de mercado",
                 "Servicios de información",
                 "Servicios técnicos de arquitectura e ingeniería; ensayos y análisis técnicos",
                 "Venta y reparación de vehículos de motor y motocicletas",
                 "Asistencia en establecimientos residenciales",
                 "Actividades jurídicas y de contabilidad"],
    "Travel, Hotels & Mobility": ["Servicios de alojamiento",
                                  "Actividades deportivas, recreativas y de entretenimiento",
                                  "Transporte aéreo",
                                  "Transporte marítimo y por vías navegables interiores",
                                  "Transporte terrestre y por tubería",
                                  "Almacenamiento y actividades anexas al transporte"]
}

# Crear el DataFrame de ejemplo


# Definir la función que asigna el valor del grupo correspondiente
def asignar_industria(row):
    for grupo, actividades in grupos_actividades.items():
        if row["Grupo"] in actividades:
            return grupo
    return "No clasificado"  # Valor por defecto si no se encuentra en ningún grupo

# Aplicar la función a la columna "Grupo" para crear la nueva columna "Grupo clasificado"
df["Industria"] = df.apply(asignar_industria, axis=1)

empresas_por_industria = df.groupby('Industria')['Nombre de la empresa'].count()
suma_facturacion_por_industria = df.groupby('Industria')['Facturación'].sum()

fechaHoy = datetime.datetime.now()
fechaHoyStr = fechaHoy.strftime("%d%m%Y_%H%M%S")
nombreExcel = f"Análisis_empresas_{fechaHoyStr}.xlsx"

workbook = xlsxwriter.Workbook(f"./{nombreExcel}")
workbook.close()
h=[0]
with pd.ExcelWriter(nombreExcel,engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
    df.to_excel(writer,header=True,index=False)
    headerPuesto = True
    suma_facturacion_por_sector.to_excel(writer, sheet_name='Suma Facturación', index=True)
    empresas_por_sector.to_excel(writer, sheet_name='Empresas Sector', index=True)
    suma_facturacion_por_industria.to_excel(writer, sheet_name='Industrias Facturación', index=True)
    empresas_por_industria.to_excel(writer, sheet_name='Empresas Industria', index=True)
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  1 15:38:22 2021

@author: hlopez


Informacion y ayuda
Descarga de archivos zip
https://www.simplifiedpython.net/python-download-file/

PDF con la descripcion de archivos
https://www.camaraderiesgo.com/wp-content/uploads/2021/03/MEFFStation-FD-Liquidacion.pdf

"""

import pandas as pd
# import numpy as np
import requests
import zipfile
import os
from datetime import date, timedelta, datetime

# Carpetas
carpeta_descarga = 'C:\\Users\\hlopez\\Documents\\Garantias\\MEFF\\'
carpeta_precio = 'C:\\Users\\hlopez\\Documents\\Precios\\MEFF\\C7\\'
carpeta_exportacion = 'C:\\Users\\hlopez\\Documents\\Precios\\MEFF\\'

# Fechas
fecha = date.today() - timedelta(days=1)
fecha_url = fecha.strftime('%y%m%d')
fecha_exportacion = fecha.strftime('%Y%m%d')

# Archivos
nombre_zip = 'ME{}.zip'.format(fecha_url)
archivo_zip = carpeta_descarga + nombre_zip
archivo_precio = 'CCONTRSTAT.C7' # Datos diarios de los contratos del grupo de contratos
archivo_multiplicador = 'CCONTRTYP.C7' # tipos de contrato del subgrupo contratos
archivo_contratos = 'CCONTRACTS.C7' # información general de los contratos del grupo de contratos disponibles en la sesion
archivo_exportacion = 'Precios_MEFF_{}.xlsx'.format(fecha_exportacion)
# Nombres de las columnas segun el pdf
columnas_multiplicador = ['SessionDate', 'ContractGroup', 'ContractSubgroupCode', 'ContractTypeCode', 'ContractTypeDescription',
                          'PriceMultiplier', 'Nominal', 'Currency', 'CalcMethod', 'Filler', 'ContractFamily', 'All', 'PriceType',
                          'SecurityType', 'FlexibleIndicator', 'ExerciseStyle', 'SettMethod', 'PutorCall', 'Periodicity',
                          'AdjustmentsRule', 'CFICode', 'UnitOfMeasure', 'BaseCurrency', 'SettlCurrency']
columnas_contratos = ['SessionDate', 'ContractGroup', 'ContractCode', 'ContractSubgroupCode',
                      'ContractTypeCode', 'StrikePrice', 'MaturityDate', 'TradingEndDate',
                      'ExerciseUnderlyingContractCode', 'MarginUnderlyingContractCode',
                      'ArrayCode', 'Filler1', 'Filler2', 'ExpirySpan', 'MaturityMonthYear',
                      'ISINCode', 'StartMaturityMonthYear', 'EndMaturityMonthYear', 'VersionNumber',
                      'ForwardMaturityDate', 'SpotMaturityDate']
columnas_precio = ['SessionDate', 'ContractGroup', 'ContractCode', 'HighPrice', 'LowPrice', 
                   'FirstPrice', 'LastPrice', 'SettlPrice', 'SettlVolatility', 'SettlDelta', 
                   'PreviousDaySettlPrice', 'PreviousDaySettlVolatility', 'PreviousDaySettlDelta', 
                   'TotalRegVolume', 'NumberOfTrades', 'OpenInterest', 'AccruedInterest',
                   'Yield', 'ForwardPrice', 'PreviousDayForwardPrice', 'NextDaySwapPoints']
# Seleccion de las columnas que se necesitan
col_multiplicador_seleccion = ['SessionDate', 'ContractSubgroupCode', 'ContractTypeCode', 'ContractTypeDescription', 'PriceMultiplier', 'Periodicity']
col_contratos_seleccion = ['SessionDate', 'ContractCode', 'ContractSubgroupCode', 'ContractTypeCode', 'MaturityDate', 'TradingEndDate', 'MarginUnderlyingContractCode']
col_precio_seleccion = ['SessionDate', 'ContractCode', 'SettlPrice']
columnas_exportacion = ['ContractCode', 'Tipo', 'Periodo', 'TradingEndDate', 'MaturityDate', 'PriceMultiplier', 'SettlPrice']
# Nombre de las columnas del archivo de exportacion
nombre_columnas = ['Valor', 'Tipo', 'Periodo', 'Fin Registro', 'Fin Entrega', 'Multiplicador', 'Precio']
# Elementos necesarios para generar nuevas columnas
lista_tipo = ['Base', 'Punta', 'Gas Natural']
lista_periodo = ['Diario', 'Resto Mes', 'Fin de Semana', 'Semanal', 'Anual', 'Mensual', 'Trimestral', 'Temporada', 'Semanal']
lista_periodicity = ['D', 'm', 'E', 'K', 'Y', 'M', 'Q', 'S', 'B']
tabla_periodo = pd.DataFrame({'CodPeriodo':lista_periodicity, 
                              'Periodo':lista_periodo})

# Genera la url
url = 'https://www.bmeclearing.es/docs/Ficheros/Descarga/dME/{}?-mlOKQ!!'.format(nombre_zip)

# Funciones
def extraccion_zip(archivo_zip, carpeta_descarga):
    try:
        with zipfile.ZipFile(archivo_zip) as z:
            # Lista el contenido
            lista_contenido = zipfile.ZipFile.namelist(z)
            z.extractall(carpeta_descarga)
            print('Extraidos todos los archivos')
    except:
        print('Archivo invalido')
    return lista_contenido


#%% 1. Descarga de datos

# Descarga el contenido del archivo en formato binario
r = requests.get(url)
with open(archivo_zip, 'wb') as zip:
    zip.write(r.content)

# Abre los archivos zip
lista_contenido = extraccion_zip(archivo_zip, carpeta_descarga)
# Si el contenido es otro zip
if (len(lista_contenido)==1) & (lista_contenido[0][-4:]=='.zip'):
    archivo_zip = carpeta_descarga + lista_contenido[0]
    lista_contenido = extraccion_zip(archivo_zip, carpeta_precio)

# Lista los productos
lista_archivos = os.listdir(carpeta_precio)
precios = pd.read_csv(carpeta_precio+archivo_precio, sep=';', decimal=',', header=None)
precios.columns = columnas_precio
precios = precios[col_precio_seleccion]
multiplicador = pd.read_csv(carpeta_precio+archivo_multiplicador, sep=';', decimal=',', header=None)
multiplicador.columns = columnas_multiplicador
multiplicador = multiplicador[col_multiplicador_seleccion]
contratos = pd.read_csv(carpeta_precio+archivo_contratos, sep=';', header=None)
contratos.columns = columnas_contratos
contratos = contratos[col_contratos_seleccion]
# Mete todas las tablas en un diccionario para luego hacer un bucle sobre el
dict_tablas = {'precios': precios, 'multiplicador': multiplicador, 'contratos':contratos}

# Comprueba la fechas y elimina las columnas
for k in dict_tablas:
    fecha_tabla = datetime.strptime(str(dict_tablas[k]['SessionDate'][0]), '%Y%m%d').date()
    if fecha_tabla == fecha: print('La fecha de hoy y de {} coinciden'.format(k))
    else: print('ERROR: La fecha de hoy y de {} NO coinciden'.format(k), '¡Actualiza los archivos!')
# Eliminacion de las columnas sobrantes
precios.drop('SessionDate', axis=1, inplace=True)
multiplicador.drop('SessionDate', axis=1, inplace=True)
contratos.drop('SessionDate', axis=1, inplace=True)


#%% 2. Generacion de las tablas

# Union de contrato cons multiplicador en base a 'Contracttypecode'
tabla_precios = contratos.merge(multiplicador, on=['ContractSubgroupCode', 'ContractTypeCode'], how='left')
tabla_precios = tabla_precios.merge(precios, on='ContractCode', how='left')
tabla_precios[['MaturityDate', 'TradingEndDate']] = tabla_precios[['MaturityDate', 'TradingEndDate']].apply(lambda x: pd.to_datetime(x, format='%Y%m%d')) # .dt.date
# Mete el tipo
tabla_precios['Tipo'] = 'Gas Natural'
tabla_precios.loc[tabla_precios['MarginUnderlyingContractCode']=='ELECB', 'Tipo'] = 'Base'
tabla_precios.loc[tabla_precios['MarginUnderlyingContractCode']=='ELECP', 'Tipo'] = 'Punta'
# Mete el Periodo
tabla_precios = tabla_precios.merge(tabla_periodo, left_on='Periodicity', right_on='CodPeriodo', how='left')#.dropna()
# Filtra las columnas y cambia el nombre de las columnas
tabla_precios = tabla_precios[columnas_exportacion]
tabla_precios.columns = nombre_columnas
# Elimina los nan
tabla_precios = tabla_precios.dropna()

# Exporta a excel el archivo de precios
tabla_precios.to_excel(carpeta_exportacion+archivo_exportacion, index=False)

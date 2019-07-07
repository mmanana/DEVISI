# -*- coding: utf-8 -*-
"""
Created on Mon Jun 22 18:23:53 2015

@author: Mario

@prerrequisitos:

If pyodbc is not installed then: pip install pyodbc
If openpyxl is not installed then: pip install openpyxl

"""

import pyodbc
from openpyxl import load_workbook
import re
import os

print "Conectando a la base de datos..."
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost\PQS;DATABASE=UC_CIR;UID=uccir;PWD=1234')
#cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost\PQS;DATABASE=REE;UID=sa;PWD=PQSpqs12345')
cursor = cnxn.cursor()
print "Conectado a la base de datos..."



#cursor.execute("SELECT TOP 10 * FROM [UC_CIR].[dbo].[alumnos]")
#row = cursor.fetchone()
#print 'name:', row[1]          # access by column index
#print 'name:', row.user_name   # or access by name
# rows = cursor.fetchall()
# for row in rows:
#    print(row.AlumnoId, row.Nombre)
    
os.chdir( "c:\\Mario\\trabajos2\\innovacion_docente_2014\\bin")
    
# Fichero de alumnos
# Nombre del fichero
NombreFichero = '..\\base_datos\\Listado_de_Alumnos_de_la_Asignatura.xlsx'
Inicio=10
Final=112
#Final=112 # NÚMERO DE FILAS + 1

wb = load_workbook(filename = NombreFichero)
sheet_ranges = wb['Relacion de Alumnos']

indice = 0
for filas in range(Inicio,Final,1):
    indice = indice + 1
    DNI = sheet_ranges['B'+str(filas)].value
    newDNI = DNI.replace(" ", "")
    NOMBRE_COMPLETO = re.split(',' , str(sheet_ranges['C'+str(filas)].value))
    APELLIDOS = NOMBRE_COMPLETO[0]
    NOMBRE =  NOMBRE_COMPLETO[1]
    EMAIL = sheet_ranges['D'+str(filas)].value
    ASIGNATURA = sheet_ranges['B5'].value
    ANIO = '2015'
    valores = '( \' ' + newDNI + '\' , \'' + APELLIDOS + '\' , \'' + NOMBRE + '\' , \'' + EMAIL + '\' , \'' + str(ASIGNATURA) + '\' , \'' + ANIO + '\' )'
    print indice, valores
    sql_statement = "insert into alumnos( AlumnoId, Apellidos, Nombre, Email, AsignaturaId, Anio) values " + valores
    cursor.execute( str( sql_statement))
    cnxn.commit()    
    

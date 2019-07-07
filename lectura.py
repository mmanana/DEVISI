# -*- coding: utf-8 -*-
"""
Created on Mon Jun 22 18:23:53 2015

@author: Mario
"""

import pyodbc
from openpyxl import load_workbook

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost\PQS;DATABASE=UC_CIR;UID=sa;PWD=PQSpqs12345')
cursor = cnxn.cursor()

cursor.execute("SELECT TOP 10 * FROM [UC_CIR].[dbo].[alumnos]")
#row = cursor.fetchone()
#print 'name:', row[1]          # access by column index
#print 'name:', row.user_name   # or access by name

rows = cursor.fetchall()
for row in rows:
    print(row.AlumnoId, row.Nombre)
    
    

# Fichero de alumnos


    
    
cursor.execute("insert into alumnos( AlumnoId, Nombre) values ('1', 'Pedro')")
cnxn.commit()    




wb = load_workbook(filename = 'Listado_de_Alumnos_de_la_Asignatura.xlsx')
sheet_ranges = wb['Relacion de Alumnos']
print(sheet_ranges['A10'].value)

print wb.get_sheet_names()

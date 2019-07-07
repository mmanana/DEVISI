# -*- coding: utf-8 -*-
"""
Created on Mon Jun 22 18:23:53 2015

@author: Mario
"""

import pyodbc
from openpyxl import load_workbook
import re
import random
import os
import subprocess
import shlex
import time 
import sys
import math

# ****************************************************************************
# ****************************************************************************
# Codigo de la asignatura o asignaturas (docencia compartida)
ASIGNATURAID = 'G412G280'
# Año
ANIO = '2015'
# Configuración acceso SQL server
# SERVER .- Debe coincidir con el nombre de la instancia de la base de datos
# Si la instancia es un nombre simple utilizar "localhost"
# Si la instancia es nombremaquina\instancia utilizar "localhost\instancia"
sql_configuracion = 'DRIVER={SQL Server};SERVER=localhost\PQS;DATABASE=UC_CIR;UID=uccir;PWD=1234'

# Fichero de configuración de casos
NombreFichero = 'C:\\Mario\\trabajos2\\innovacion_docente_2014\\casos\\casos.xlsx'
# Fila del fichero excel que define el caso
caso = 18
# Ruta a la carpeta de casos
rutacasos = 'C:\\Mario\\trabajos2\\innovacion_docente_2014\\casos\\'
# ****************************************************************************
# ****************************************************************************


def tail(f, lines=1, _buffer=4098):
    """Tail a file and get X lines from the end"""
    # place holder for the lines found
    lines_found = []

    # block counter will be multiplied by buffer
    # to get the block size from the end
    block_counter = -1

    # loop until we find X lines
    while len(lines_found) < lines:
        try:
            f.seek(block_counter * _buffer, os.SEEK_END)
        except IOError:  # either file is too small, or too many lines requested
            f.seek(0)
            lines_found = f.readlines()
            break

        lines_found = f.readlines()

        # we found enough lines, get out
        if len(lines_found) > lines:
            break

        # decrement the block counter to get the
        # next X bytes
        block_counter -= 1

    return lines_found[-lines:]


def find(word, letter):
    """Find the position of a character in a string"""
    index = 0
    while index < len(word):
        if word[index] == letter:
            return index
        index = index + 1
    return -1


print 'abriendo conexion con la bbdd'
cnxn = pyodbc.connect( sql_configuracion)
cursor = cnxn.cursor()
print 'conexion abierta'



# Recupera lista de alumnos
# Configuración script escritura
cursor_sql = 'SELECT ALL * FROM [UC_CIR].[dbo].[alumnos] WHERE AsignaturaId = ' 
cursor_sql = cursor_sql + "\'" + ASIGNATURAID + "\'"  + ' AND Anio = ' + "\'" + ANIO + "\'" 
cursor.execute( cursor_sql)
#row = cursor.fetchone()
#print 'name:', row[1]          # access by column index
#print 'name:', row.user_name   # or access by name
listado_alumnos = cursor.fetchall()
VariablesEn = [[0 for x in range(3)] for x in range(len(listado_alumnos))] 
VariablesSa = [[0 for x in range(1)] for x in range(len(listado_alumnos))] 

    

wb = load_workbook(filename = NombreFichero)
sheet_ranges = wb['Hoja1']

NumVarEn = sheet_ranges['E'+str(caso)].value
if NumVarEn >= 1:
    var1name = sheet_ranges['F'+str(caso)].value
    var1min = sheet_ranges['G'+str(caso)].value
    var1max = sheet_ranges['H'+str(caso)].value            
if NumVarEn >= 2:
    var2name = sheet_ranges['I'+str(caso)].value
    var2min = sheet_ranges['J'+str(caso)].value
    var2max = sheet_ranges['K'+str(caso)].value            
if NumVarEn >= 3:
    var3name = sheet_ranges['L'+str(caso)].value
    var3min = sheet_ranges['M'+str(caso)].value
    var3max = sheet_ranges['N'+str(caso)].value            

NumVarSa = sheet_ranges['O'+str(caso)].value

# Cambia al directorio del caso analizado
directorio = rutacasos  +  sheet_ranges['A'+str(caso)].value
os.chdir( directorio)
retval = os.getcwd()
print "Directory changed successfully %s" % retval

# Genera ficheros .cir para cada alumno
print "Ficheros .cir para cada alumno..."
indice=0
# listado_alumnos=listado_alumnos[1:4] # solo para pruebas
for alumno in listado_alumnos:
    print(alumno.AlumnoId, alumno.Nombre)
    fichero = '.\\' + sheet_ranges['C'+str(caso)].value               
    fichero_sal = '.\\' + sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.cir'
    fichero_sal = fichero_sal.replace(" ", "")
    f_sal = open( fichero_sal, 'w')
    f = open(fichero)
    var1 = str(round(random.uniform( var1min, var1max),0))
    VariablesEn[indice][0] = var1
    var2 = str(round(random.uniform( var2min, var2max),0))
    VariablesEn[indice][1] = var2    
    var3 = str(round(random.uniform( var3min, var3max),0))
    VariablesEn[indice][2] = var3    
    

    VarSa=round(random.uniform( 0.501, NumVarSa+0.499),0)
    if VarSa == 1:
        VarSaName = sheet_ranges['P'+str(caso)].value
        VarSaEtiqueta = sheet_ranges['S'+str(caso)].value
    if VarSa == 2:
        VarSaName = sheet_ranges['Q'+str(caso)].value
        VarSaEtiqueta = sheet_ranges['T'+str(caso)].value
    if VarSa == 3:
        VarSaName = sheet_ranges['R'+str(caso)].value
        VarSaEtiqueta = sheet_ranges['U'+str(caso)].value

    VarSaName = str( VarSaName)
    print VarSaName
    VariablesSa[indice][0] = VarSaEtiqueta
    indice = indice+1

    fsalida = sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.txt'    
    fsalida = fsalida.replace( " ", "")
    for line in f:
        # print line
        new_line = line.replace( var1name, var1)
        new_line = new_line.replace( var2name, var2)
        new_line = new_line.replace( var3name, var3)
        new_line = new_line.replace( "$fsalida$", fsalida)   
        new_line = new_line.replace( "$sal1$", VarSaName)           
        f_sal.write( new_line)
    f.close()
    f_sal.close()
    print indice
    
    f = open( fichero_sal)
    f.close()
 #   f = open( fsalida)
 #   f.close()
    
    
# raw_input("Press ENTER to continue.")    
indice = 0    
print "Genera los ficheros de salida .tex,.pdf,.png,.cir..."
for alumno in listado_alumnos:
    #print(alumno.AlumnoId, alumno.Nombre)
    fichero = '.\\' + sheet_ranges['A'+str(caso)].value + '.tex'               
    fichero_sal = '.\\' + sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.tex'
    fichero_sal = fichero_sal.replace(" ", "")
    f_sal = open( fichero_sal, 'w')
    f = open(fichero)
    var1 = VariablesEn[indice][0]
    var2 = VariablesEn[indice][1]
    var3 = VariablesEn[indice][2]

    for line in f:
        # print line
        new_line = line.replace( var1name, var1)
        new_line = new_line.replace( var2name, var2)
        new_line = new_line.replace( var3name, var3)
        new_line = new_line.replace( "$Nombre$", alumno.Nombre)
        new_line = new_line.replace( "$Apellidos$", alumno.Apellidos)        
        new_line = new_line.replace( "$Asignatura$", alumno.AsignaturaId)                
        new_line = new_line.replace( "$Fecha$", str(alumno.Anio))      
        new_line = new_line.replace( "$resultado1$", VariablesSa[indice][0])      
        f_sal.write( new_line)

    indice = indice+1                  
    f.close()
    f_sal.close()
    ftex =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.tex'
    ftex = str(ftex)
    ftex = ftex.replace(" ", "")
    comando = 'pdflatex ' + ftex
    comando = str( comando)
    print comando
    os.system( comando)
    os.system( comando)
    # proc=subprocess.call('pdflatex C:\\Mario\\trabajos2\\innovacion_docente_2014\\casos\\caso001\\aa.tex')
    fpng =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.png'
    fpng = str(fpng)
    fpng = fpng.replace(" ", "")

    fjpg =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.jpg'
    fjpg = str(fjpg)
    fjpg = fjpg.replace(" ", "")

    fpdf =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.pdf'
    fpdf = str(fpdf)
    fpdf = fpdf.replace(" ", "")
    comando = 'C:\\ImageMagick-7.0.1-0-portable-Q16-x64\\convert -density 300 -depth 8 -quality 85 ' + fpdf + ' ' + fjpg
    comando = str(comando)
    print comando
    os.system( comando)
    
    comando = 'C:\\ImageMagick-7.0.1-0-portable-Q16-x64\\convert ' + fjpg + ' ' + fpng
    comando = str(comando)
    print comando
    os.system( comando)
       
    os.system("del *.log")
    os.system("del *.aux")
    fraw =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.raw'
    fraw = str(fraw)
    fraw = fraw.replace(" ", "")
    ftxt =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.txt'
    ftxt = str(ftxt)
    ftxt = ftxt.replace(" ", "")
    fcir =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.cir'
    fcir = str(fcir)
    fcir = fcir.replace(" ", "")
    spice = "c:\\spice\\bin\\ngspice -b -r " + fraw + " " +  fcir
    # spice = "c:\\spice\\bin\\ngspice " + directorio + "\\" + fcir + " ; taskkill /F /IM ngspice.exe"
    print spice    
    os.system( spice)
    time.sleep(2)
#    os.system( "taskkill /F /IM ngspice.exe")
    print indice
 #   f = open( fichero_sal)
 #   f.close()
 #   f = open( fpng)
 #   f.close()
 #   f = open( fjpg)
 #   f.close()
 #   f = open( fpdf)
 #   f.close()
 #   f = open( fraw)
 #   f.close()
 #   f = open( ftxt)
 #   f.close()
 #   f = open( fcir)
 #   f.close()
    

# raw_input("Press ENTER to continue.")
# Busca fichero de resultados
print "Busca soluciones spice de cada alumno..."
Resultados = [[0 for x in range(1)] for x in range(len(listado_alumnos))]
indice = 0
for alumno in listado_alumnos:
    #print(alumno.AlumnoId, alumno.Nombre)
    print indice
    ftxt =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.txt'
    ftxt = str(ftxt)
    ftxt = ftxt.replace(" ", "")
    f_txt = open( ftxt, 'r')
    Res_alumno = tail(f_txt, lines=3)
    if caso==17:
        Resultados[indice] = round( float( Res_alumno[1]),2)
    elif caso==18:
        Res_c = Res_alumno[1]
        N = len(Res_c)
        ih = find( Res_c, ",")
        real_c = Res_c[1:ih]
        imag_c = Res_c[ih+1:N-1]
        real = float( real_c)
        imag = float( imag_c)
        Resultados[indice] = math.sqrt( real*real + imag*imag) 
    indice = indice + 1
    f_txt.close()

print Resultados

# Actualiza bbdd casos-resultados

cnxn2 = pyodbc.connect( sql_configuracion)
cursor2 = cnxn2.cursor()

# raw_input("Press ENTER to continue.")
print "Almacena resultados en la tabla [ejercicios] de la bbdd"
error_por = 5
indice = -1
campos = "insert into ejercicios( EjercicioId, CasoId, AlumnoId, AsignaturaId,"
campos = campos + "Anio, FicheroEnu, FicheroCir, NumVarEn, VarEn01, VarEn02,"
campos = campos + "VarEn03,VarEn04, VarEn05, NumVarSa, VarSa01, VarSa02,"
campos = campos + "VarSa03, VarSa04, VarSa05,Sol01, Sol02, Sol03, Sol04, Sol05,"
campos = campos + "Err01, Err02, Err03, Err04, Err05, IntentosMax, Intentos) values"
for alumno in listado_alumnos:
    indice = indice + 1
    fcir =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.cir'
    fcir = str(fcir)
    fcir = fcir.replace(" ", "")
    fpdf =  sheet_ranges['A'+str(caso)].value + '_'  + str(alumno.AlumnoId) + '.pdf'
    fpdf = str(fpdf)
    fpdf = fpdf.replace(" ", "")
    EJERCICIOID = str( indice)
    CASOID = sheet_ranges['A'+str(caso)].value
    ALUMNOID = alumno.AlumnoId
    FICHEROENU = fpdf
    FICHEROCIR = fcir
    NUMVAREN = str(NumVarEn)
    VAREN01 = str( VariablesEn[indice][0])
    VAREN02 = str( VariablesEn[indice][1])
    VAREN03 = str( VariablesEn[indice][2])
    VAREN04 = ' '
    VAREN05 = ' '
    NUMVARSA = str( 1)
    VARSA01 = str( Resultados[indice])
    VARSA02 = ' '
    VARSA03 = ' '
    VARSA04 = ' '
    VARSA05 = ' '
    SOL01 = ' '
    SOL02 = ' '
    SOL03 = ' '
    SOL04 = ' '
    SOL05 = ' '
    ERR01 = str( error_por)
    ERR02 = str( error_por)
    ERR03 = str( error_por)
    ERR04 = str( error_por)
    ERR05 = str( error_por)
    INTENTOSMAX = str(5)
    INTENTOS = str(0)
    
    valores = '( \' ' + EJERCICIOID + '\' , \''
    valores = valores + CASOID + '\' , \''
    valores = valores + ALUMNOID + '\' , \''
    valores = valores + ASIGNATURAID + '\' , \''    
    valores = valores + ANIO + '\' , \''    
    valores = valores + FICHEROENU + '\' , \''    
    valores = valores + FICHEROCIR + '\' , \''    
    valores = valores + NUMVAREN + '\' , \''    
    valores = valores + VAREN01 + '\' , \''
    valores = valores + VAREN02 + '\' , \''
    valores = valores + VAREN03 + '\' , \''
    valores = valores + VAREN04 + '\' , \''
    valores = valores + VAREN05 + '\' , \''    
    valores = valores + NUMVARSA + '\' , \''
    valores = valores + VARSA01 + '\' , \''
    valores = valores + VARSA02 + '\' , \''
    valores = valores + VARSA03 + '\' , \''
    valores = valores + VARSA04 + '\' , \''
    valores = valores + VARSA05 + '\' , \''        
    valores = valores + SOL01 + '\' , \''    
    valores = valores + SOL02 + '\' , \''    
    valores = valores + SOL03 + '\' , \''    
    valores = valores + SOL04 + '\' , \''    
    valores = valores + SOL05 + '\' , \''            
    valores = valores + ERR01 + '\' , \''        
    valores = valores + ERR02 + '\' , \''        
    valores = valores + ERR03 + '\' , \''        
    valores = valores + ERR04 + '\' , \''        
    valores = valores + ERR05 + '\' , \''
    valores = valores + INTENTOSMAX +  '\' , \''
    valores = valores + INTENTOS + '\' )'   
          
    print valores
    print indice
    sql_statement = campos + valores
    cursor2.execute( str( sql_statement))
    cnxn2.commit()    




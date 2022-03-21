# -*- coding: utf-8 -*-
"""
Created on Fri Mar 18 15:10:43 2022

@author: Andres
"""

import os #Permite realizar operaciones dependientes del Sistema Operativo como crear una carpeta, listar contenidos de una carpeta, entre otros
import pandas as pd #Permite obtener columnas o filas de nuestros datos.
import glob #devuelve una lista con las entradas que coincidan con el patrón especificado en pathname.
import pdfplumber #extraer información desde archivos PDF

def pag_num():  
    for files in glob.glob("C:\Cursos\pdfs\*.pdf"):
        with pdfplumber.open(files) as pdf:
            var = 0
            cont = 0
            
            while var < 1:
                try:
                    page = pdf.pages[cont]
                    print (page)
                    cont = cont + 1
                except:
                    var = var + 1
                    print("Núm Pág. "+str(cont))
                
if __name__ == '__main__':
    pag_num()
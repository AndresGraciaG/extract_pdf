# -*- coding: utf-8 -*-
"""
Created on Wed Mar 16 12:47:10 2022

@author: Andres
"""

import os #Permite realizar operaciones dependientes del Sistema Operativo como crear una carpeta, listar contenidos de una carpeta, entre otros
import pandas as pd #Permite obtener columnas o filas de nuestros datos.
import glob #devuelve una lista con las entradas que coincidan con el patrón especificado en pathname.
import pdfplumber #extraer información desde archivos PDF

"""
start = palabra anterior a la requerida.

end = palabra posterior a la requerida.

text = texto extraído de la pág. seleccionada.
"""

def get_keyword(start, end, text):
    
    for i in range(len(start)):
        try:
            field = ((text.split(start[i]))[1].split(end[i])[0])
            return field
        except:
            continue



def especific_pag (pag_num):
    my_dataframe = pd.DataFrame()
    
    for files in glob.glob("C:\Cursos\pdfs\*.pdf"):
        with pdfplumber.open(files) as pdf:
            page = pdf.pages[pag_num]
            print (page)
            text = page.extract_text()
            text = " ".join(text.split())
            
            #Obteniendo la palabra clave #1: requiere
            start = ['Para los casos en donde Colmena no']
            end = ['el']
            keyword1 = get_keyword(start, end, text)
            
            #Obteniendo la palabra clave #2: quien asume el riesgo y paga la indemnización a los beneficiarios
            start = ['Es']
            end = ['en caso de presentarse un siniestro']
            keyword2 = get_keyword(start, end, text)
            
            #Obteniendo la palabra clave #3: la empresa o entidad que le presta dinero a los asegurados
            start = ['Normalmente,']
            end = ['.']
            keyword3 = get_keyword(start, end, text)
            
            
            my_list = [keyword1, keyword2, keyword3]
            
            my_list = pd.Series(my_list)
            
            my_dataframe = my_dataframe.append(my_list, ignore_index=True)
            
            print("Keywords extracted successfully")
            
    
    my_dataframe = my_dataframe.rename(columns={0:'Keyword 1',
                                                1:'Keyword 2',
                                                2:'Tomador'})
    
    save_path = ('C:\Cursos\docExcel')
    os.chdir(save_path)
    
    my_dataframe.to_excel('documento.xlsx', sheet_name= 'my dataframe')


def main():
    especific_pag(0)

if __name__ == '__main__':
    main()
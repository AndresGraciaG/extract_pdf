# -*- coding: utf-8 -*-
"""
Created on Thu Mar 10 17:51:56 2022

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
        
def main():
    my_dataframe = pd.DataFrame()
    
    for files in glob.glob("C:\Cursos\pdfs\*.pdf"):
        with pdfplumber.open(files) as pdf:
            
            var = 0
            cont = 0
            var1 = 0
            cont1 = 0
            
            while var < 1:
                try:
                    page = pdf.pages[cont]
                    #print (page)
                    cont = cont + 1
                except:
                    var = var + 1
                    print("Núm Pág. "+str(cont))
            
            while var1 < cont:
                page = pdf.pages[cont1]
                #print (page)
                text = page.extract_text()
                text = " ".join(text.split())
                        
                #Obteniendo la palabra clave #1: Diplomados
                start = ['Por la cual se ordena ofrecer']
                end = ['como opción de grado']
                keyword1 = get_keyword(start, end, text)
                        
                #Obteniendo la palabra clave #2:  General
                start = ['Que la Rectora']
                end = ['de la Universidad del Sinú']
                keyword2 = get_keyword(start, end, text)
                        
                #Obteniendo la palabra clave #3: para los Diplomados
                start = ['Fijar las fechas de inscripción y matriculas']
                end = ['como opción de grado que se ofrecen']
                keyword3 = get_keyword(start, end, text)
                
        
                my_list = [keyword1, keyword2, keyword3]
                
                my_list = pd.Series(my_list)
                        
                my_dataframe = my_dataframe.append(my_list, ignore_index=True)
                        
                print("Keywords extracted successfully")
                    
                cont1 = cont1 + 1
                    
                    
                var1 = var1 + 1
            
            my_dataframe = my_dataframe.rename(columns={0:'Keyword 1',
                                                        1:'Keyword 2',
                                                        2:'Tomador'})
                            
            save_path = ('C:\Cursos\docExcel')
            os.chdir(save_path)
            my_dataframe.to_excel('documento.xlsx', sheet_name= 'Hoja 1')

if __name__ == '__main__':
    main()
            
            
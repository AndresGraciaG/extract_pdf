# -*- coding: utf-8 -*-
"""
Created on Mon Mar 21 22:10:55 2022

@author: A2G7G7R5
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
    print(start)
    for i in range(len(start)):
        try:
            field = ((text.split(start[i]))[1].split(end[i])[0])
            return field
        except:
            continue

def main():
    my_dataframe = pd.DataFrame()
    find = False
    for files in glob.glob("C:\Cursos\pdfs\*.pdf"):
        with pdfplumber.open(files) as pdf:

            var = 0
            cont = 0
            var1 = 0
            #cont1 = 0
            
            while var < 1:
                try:
                    page = pdf.pages[cont]
                    #print (page)
                    cont = cont + 1
                except:
                    var = var + 1
                    print("Núm Pág. "+str(cont))
            
            
            while var1 < cont:
                page = pdf.pages[var1]
                #print (page)
                text = page.extract_text()
                text = " ".join(text.split())
                if  find == False:
                    #Obteniendo la palabra clave #1: Vida Grupo Deudor
                    start = ['PRODUCTO']
                    end = ['Producto 3406']
                    keyword1 = get_keyword(start, end, text)
                    find = True
                    print('salir '+str(find))
                    
                #Obteniendo la palabra clave #2: 3406
                start = ['PRODUCTO VIDA GRUPO DEUDOR Producto']
                end = ['']
                keyword2 = get_keyword(start, end, text)

                #Obteniendo la palabra clave #3: 2158-8-2019
                start = ['PROPUESTA – (']
                end = [') TOMADOR BANCO MUNDO MUJER S.A.']
                keyword3 = get_keyword(start, end, text)
                
                #Obteniendo la palabra clave #4: BANCO MUNDO MUJER S.A.
                start = [') TOMADOR']
                end = ['. NIT 900.768.933 - 8']
                keyword4 = get_keyword(start, end, text)
                
                #Obteniendo la palabra clave #5: 900.768.933 - 8
                start = ['NIT']
                end = ['ACTIVIDAD']
                keyword5 = get_keyword(start, end, text)
                
                #Obteniendo la palabra clave #6: Menores 74 años
                start = ['Máximo valor por deudor']
                end = ['El máximo valor asegurado por persona en uno o varios créditos']
                keyword6 = get_keyword(start, end, text)
                
                #Obteniendo la palabra clave #7: Mayores o iguales 74 años
                start = ['será de 450 SMMLV.']
                end = ['El máximo valor asegurado por persona en uno o varios créditos']
                keyword7 = get_keyword(start, end, text)
                
                my_list = [keyword1, keyword2, keyword3, keyword4, keyword5, keyword6, keyword7]
               
                my_list = pd.Series(my_list)

                my_dataframe = my_dataframe.append(my_list, ignore_index=True)

                    #cont1 = cont1 + 1
                var1 = var1 + 1
                keyword1 = None
                
                
            print("Palabras extraídas con éxito")
            
            
            my_dataframe = my_dataframe.rename(columns={0:'Producto',
                                                        1:'# Producto',
                                                        2:'#Cotización',
                                                        3:'Tomador',
                                                        4:'NIT',
                                                        5:'Cobertura',
                                                        6:'Cobertura'})

            save_path = ('C:\Cursos\docExcel')
            os.chdir(save_path)
            my_dataframe.to_excel('documento.xlsx', sheet_name= 'Hoja 1')

if __name__ == '__main__':
    main()
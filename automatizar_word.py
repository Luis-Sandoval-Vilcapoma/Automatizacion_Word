# Importamos las librerias que vamos a usar

import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

doc = DocxTemplate("plantilla.docx")

#Creo las variables que van a ser constante en los diversos archivos
nombre = "Luis Sandoval"
correo = "luis@gmail.com"
telefono = "(123)456-7890"
fecha = datetime.today().strftime("%d/%m/%y")

#DocxTemplate solo acepta diccionarios, las constantes las ponemos en un diccionario
#Las keys deben ser iguales al que se coloco en el word y los values son las constantes creadas

constantes = {
                "nombre":nombre,
                "telefono":telefono,
                "correo":correo,
                "fecha":fecha
            }

# creamos la variable df con el archivo de excel con las notas
df =pd.read_excel('Notas.xlsx')

#Se crea un loop para que se itere cada fila del dataframe y devuelve una tupla con el indice de la fila y los valores de la fila en formato de serie de pandas
for indice,fila in df.iterrows():
    contenido = {
                "nombre_alumno":fila["Alumno"],
                "nota_mat":fila["Mat"],
                "nota_fis":fila["Fis"],
                "nota_qui":fila["Qui"]
                }
    #unimos las constanntes a cada contenido de loop
    contenido.update(constantes)

    #Importamos los valores del dict al archivo de word que se va a generar en cada loop
    doc.render(contenido)

    # Guardamos el word con los valores en un nuevo archivo, se genera un archivo por alumno
    doc.save(f"notas_de_{fila["Alumno"]}.docx")

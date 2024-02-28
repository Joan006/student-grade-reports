from dataclasses import dataclass
import pandas as pd 
from datetime import datetime 
from docxtpl import DocxTemplate
import os



# guaradamos el archivo en una variable - llamando a la funcion DocxTemplate
nombre = "Joan Mtz"
telefono = "5611922947"
correo = "joan@hotmail.com"
fecha = datetime.today().strftime("%d-%m/%Y")
# guardamos el valor del dia actual 

# constantes
DOC = DocxTemplate("plantilla.docx")
CARPETA = fecha

constantes = {
    "nombre" : nombre,
    "telefono" : telefono,
    "correo" : correo,
    "fecha" : fecha
}

df = pd.read_excel("Alumnos.xlsx")
# con pandas y la funcion read_excel - guardamos el contenido en un DataFrame

for indice, fila in df.iterrows():
    # iteramos en la lista del data frame crado df
    contenido = {
        'nombre_alumno' : fila["Nombre del Alumno"],
        'calificacion_mat' : fila["Mat"],
        'calificacion_fis' : fila ["Fis"],
        'calificacion_qui' : fila ["Qui"]
    }
    contenido.update(constantes)

    DOC.render(contenido)

    if not os.path.exists(CARPETA):
        os.makedirs(CARPETA)
    
    DOC.save(os.path.join(CARPETA, f"Notas_de_{fila['Nombre del Alumno']}.docx"))




import pandas as pa
from unidecode  import unidecode
import os

#cargar los datos
ruta = os.path.join(os.path.dirname(__file__), "notas_estudiantes.csv")
dataset = pa.read_csv(ruta)
print(dataset)

#limpiar columna carrera espacio en blanco,primera en mayuscula y quitar las tildes
dataset["Carrera"] = dataset["Carrera"].apply(lambda x: unidecode(x.strip().title()) if isinstance(x, str)else x)
print(dataset)
dataset["Nombre"] = dataset["Nombre"].apply(lambda x: unidecode(x.strip().title())if isinstance(x, str)else x)
print(dataset)
#eliminar las filas con nombre,edad,carrera y notas vacios
dataset= dataset.dropna(subset=["Nombre","Edad","Carrera","Nota1","Nota2","Nota3"])
print(dataset)
#cambiar tipo de dato float a int
dataset["Edad"]=dataset["Edad"].astype(int)
print(dataset)
#eliminar duplicados
dataset= dataset.drop_duplicates()
print(dataset)

#crear funcion para calcular el promedio de la materia
dataset["Promedio"] =  dataset[["Nota1","Nota2","Nota3"]].mean(axis=1)
dataset["Promedio"] = dataset["Promedio"].round(1)
print(dataset)
#funcion para clasificar el promedio excelente,bueno,regular
def clasificarpro(prome):
    if prome >= 4.5:
        return "Excelente"
    elif prome >=3.5:
        return "Bueno"
    elif prome >= 2.5:
        return "Regular"
    else:
        return "Deficiente"
    
#llamar la funcion y crear una columna desempeño
dataset["Desempeño"] = dataset["Promedio"].apply(clasificarpro)   
print(dataset) 

Nuevo = dataset.to_excel('notas_limpio.xlsx',index=False)
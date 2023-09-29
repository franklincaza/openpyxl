# openpyxl
manual de como usar la libreria openpyxl

import openpyxl

# Nombre del archivo de Excel en el que deseas escribir
nombre_archivo = "mi_archivo.xlsx"

# Abre el archivo de Excel
workbook = openpyxl.load_workbook(nombre_archivo)

# Selecciona la hoja con la que deseas trabajar (por nombre o índice)
hoja = workbook["NombreDeLaHoja"]  # Reemplaza "NombreDeLaHoja" con el nombre de tu hoja

# selecionar celdas recomendable para iterar.
hoja.cell(row=1, column=1).value

# Escribe un valor entero en una celda específica (por ejemplo, en la celda A1)
valor_entero = 42
hoja["A1"].value = valor_entero


# Guarda los cambios en el archivo de Excel
workbook.save(nombre_archivo)

# Cierra el archivo de Excel cuando hayas terminado
workbook.close()

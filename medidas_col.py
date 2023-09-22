import openpyxl

def obtener_anchos_columnas(nombre_archivo):
    workbook = openpyxl.load_workbook(nombre_archivo)
    sheet = workbook.active

    anchos_columnas = []

    for col_idx, column_dimension in enumerate(sheet.column_dimensions.values(), 1):
        anchos_columnas.append(column_dimension.width)

    return anchos_columnas


# Ruta de un archivo Excel
nombre_archivo = 'ARANDELA CHAPISTA y comunes.xlsx'

# Obtener las anchuras de las columnas
anchos_columnas = obtener_anchos_columnas(nombre_archivo)

# Imprimir las anchuras de las columnas
print(anchos_columnas)

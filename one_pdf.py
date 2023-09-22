import PyPDF2
import os

def eliminar_paginas_en_blanco(archivo_entrada):
    pdf_reader = PyPDF2.PdfReader(archivo_entrada)
    pdf_writer = PyPDF2.PdfWriter()

    for num_pagina in range(len(pdf_reader.pages)):
        pagina = pdf_reader.pages[num_pagina]
        contenido = pagina.extract_text()

        # Verificar si el contenido de la página está vacío o solo contiene espacios en blanco
        if not contenido.strip():
            continue

        pdf_writer.add_page(pagina)


    pdf_writer.write(archivo_entrada)


def combinar_archivos_pdf(lista_archivos, archivo_salida):
    pdf_mezclado = PyPDF2.PdfWriter()

    for archivo in lista_archivos:
        eliminar_paginas_en_blanco(archivo)
        pdf_reader = PyPDF2.PdfReader(archivo)

        for num_pagina in range(len(pdf_reader.pages)):
            pagina = pdf_reader.pages[num_pagina]
            pdf_mezclado.add_page(pagina)

    with open(archivo_salida, 'wb') as archivo_salida:
        pdf_mezclado.write(archivo_salida)


def obtener_archivos_pdf(directorio):
    lista_archivos_pdf = []

    for nombre_archivo in os.listdir(directorio):
        ruta_archivo = os.path.join(directorio, nombre_archivo)

        if os.path.isfile(ruta_archivo) and nombre_archivo.lower().endswith(".pdf"):
            lista_archivos_pdf.append(ruta_archivo)

    return lista_archivos_pdf


if __name__ == "__main__":

    list_pdfs = obtener_archivos_pdf("C:\\Users\\notebook\\Desktop\\catalogo\\pdfs")
    for pdf in list_pdfs:
        eliminar_paginas_en_blanco(pdf)

    combinar_archivos_pdf(list_pdfs, "Catalogo.pdf")
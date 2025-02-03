# # import sys
# # print(sys.path)
# import os
# from docx2pdf import convert
# from PyPDF2 import PdfReader, PdfWriter
# from PIL import Image
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import letter
#
# # Configuración de rutas
# CARPETA_ENTRADA = "GUIAS CONFIRMADAS SIN CUÑAR"
# CARPETA_SALIDA = "GUIAS CONFIRMADAS"
# IMAGEN_CUÑO = "cuño.png"
#
# def crear_carpetas():
#     """Crea las carpetas de entrada y salida si no existen."""
#     os.makedirs(CARPETA_ENTRADA, exist_ok=True)
#     os.makedirs(CARPETA_SALIDA, exist_ok=True)
#
# def agregar_cuño_a_pdf(pdf_entrada, pdf_salida, imagen_cuño):
#     """Agrega una imagen en la parte inferior derecha de la primera página de un PDF."""
#     reader = PdfReader(pdf_entrada)
#     writer = PdfWriter()
#
#     # Crear un archivo temporal con ReportLab
#     temp_pdf = "temp.pdf"
#     c = canvas.Canvas(temp_pdf, pagesize=letter)
#     width, height = letter
#
#     # Verificar que la imagen existe y cargarla
#     if not os.path.exists(imagen_cuño):
#         print(f"Error: No se encontró la imagen del cuño '{imagen_cuño}'.")
#         return
#
#     cuño = Image.open(imagen_cuño)
#     cuño_width, cuño_height = cuño.size
#
#     # Redimensionar la imagen si es demasiado grande
#     max_width, max_height = 139, 50
#     if cuño_width > max_width or cuño_height > max_height:
#         cuño = cuño.resize((max_width, max_height))
#
#     x_cuño = width - cuño.width - 75
#     y_cuño = 230
#
#     c.drawImage(imagen_cuño, x_cuño, y_cuño, width=cuño.width, height=cuño.height)
#     c.save()
#
#     # Fusionar el PDF original con el cuño
#     cuño_reader = PdfReader(temp_pdf)
#     for i, page in enumerate(reader.pages):
#         if i == 0:
#             page.merge_page(cuño_reader.pages[0])
#         writer.add_page(page)
#
#     # Guardar el PDF final
#     with open(pdf_salida, "wb") as f:
#         writer.write(f)
#
#     os.remove(temp_pdf)
#
# def convertir_word_a_pdf():
#     """Convierte archivos Word a PDF, agrega el cuño y los guarda."""
#     for archivo in os.listdir(CARPETA_ENTRADA):
#         if archivo.endswith(".docx"):
#             ruta_word = os.path.join(CARPETA_ENTRADA, archivo)
#             nombre_base = os.path.splitext(archivo)[0]
#             ruta_pdf_temporal = os.path.join(CARPETA_ENTRADA, f"{nombre_base}.pdf")
#             ruta_pdf_final = os.path.join(CARPETA_SALIDA, f"{nombre_base}.pdf")
#
#             # Convertir Word a PDF (docx2pdf genera el PDF en la misma carpeta del .docx)
#             convert(CARPETA_ENTRADA)
#
#             # Mover el archivo a la carpeta de salida
#             if os.path.exists(ruta_pdf_temporal):
#                 os.rename(ruta_pdf_temporal, ruta_pdf_final)
#
#                 # Agregar el cuño
#                 agregar_cuño_a_pdf(ruta_pdf_final, ruta_pdf_final, IMAGEN_CUÑO)
#                 print(f"Procesado: {archivo}")
#             else:
#                 print(f"Error: No se generó el PDF de {archivo}")
#
# if __name__ == "__main__":
#     crear_carpetas()
#     convertir_word_a_pdf()

import os
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Configuración de rutas
CARPETA_ENTRADA = "GUIAS CONFIRMADAS SIN CUÑAR"
CARPETA_SALIDA = "GUIAS CONFIRMADAS"
IMAGEN_CUÑO = "cuño.png"

def crear_carpetas():
    """Crea las carpetas de entrada y salida si no existen."""
    os.makedirs(CARPETA_ENTRADA, exist_ok=True)
    os.makedirs(CARPETA_SALIDA, exist_ok=True)

def agregar_cuño_a_pdf(pdf_entrada, pdf_salida, imagen_cuño):
    """Agrega una imagen en la parte inferior derecha de la primera página de un PDF."""
    reader = PdfReader(pdf_entrada)
    writer = PdfWriter()

    # Crear un archivo temporal con ReportLab
    temp_pdf = "temp.pdf"
    c = canvas.Canvas(temp_pdf, pagesize=letter)
    width, height = letter

    # Verificar que la imagen existe y cargarla
    if not os.path.exists(imagen_cuño):
        print(f"Error: No se encontró la imagen del cuño '{imagen_cuño}'.")
        return

    cuño = Image.open(imagen_cuño)
    cuño_width, cuño_height = cuño.size

    # Redimensionar la imagen si es demasiado grande
    max_width, max_height = 139, 50
    if cuño_width > max_width or cuño_height > max_height:
        cuño = cuño.resize((max_width, max_height))

    x_cuño = width - cuño.width - 75
    y_cuño = 230
    c.drawImage(imagen_cuño, x_cuño, y_cuño, width=cuño.width, height=cuño.height)
    c.save()

    # Fusionar el PDF original con el cuño
    cuño_reader = PdfReader(temp_pdf)
    for i, page in enumerate(reader.pages):
        if i == 0:
            page.merge_page(cuño_reader.pages[0])
        writer.add_page(page)

    # Guardar el PDF final
    with open(pdf_salida, "wb") as f:
        writer.write(f)

    # Eliminar el archivo temporal
    os.remove(temp_pdf)

def convertir_word_a_pdf():
    """Convierte archivos Word a PDF, agrega el cuño y los guarda."""
    for archivo in os.listdir(CARPETA_ENTRADA):
        if archivo.endswith(".docx"):
            ruta_word = os.path.join(CARPETA_ENTRADA, archivo)
            nombre_base = os.path.splitext(archivo)[0]
            ruta_pdf_final = os.path.join(CARPETA_SALIDA, f"{nombre_base}.pdf")

            # Convertir Word a PDF directamente en la carpeta de salida
            convert(ruta_word, ruta_pdf_final)

            # Agregar el cuño al PDF
            if os.path.exists(ruta_pdf_final):
                agregar_cuño_a_pdf(ruta_pdf_final, ruta_pdf_final, IMAGEN_CUÑO)
                print(f"Procesado: {archivo}")
            else:
                print(f"Error: No se generó el PDF de {archivo}")

if __name__ == "__main__":
    crear_carpetas()
    convertir_word_a_pdf()
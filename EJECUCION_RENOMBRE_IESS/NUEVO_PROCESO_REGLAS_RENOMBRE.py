import json
import fitz  # PyMuPDF
import pandas as pd
import os
import shutil
import io
from PIL import Image
import PyPDF2

def limpiar_contenido_carpeta(ruta_carpeta):
    for carpeta_raiz, carpetas, archivos in os.walk(ruta_carpeta, topdown=False):
        for archivo in archivos:
            os.remove(os.path.join(carpeta_raiz, archivo))
        for carpeta in carpetas:
            shutil.rmtree(os.path.join(carpeta_raiz, carpeta), ignore_errors=True)

def pdf_to_images(pdf_path, dpi=70):
    doc = fitz.open(pdf_path)
    images = []
    
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    
    return images

def convert_to_grayscale(image):
    return image.convert('L')

def compress_image(image, quality=60):
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='JPEG', quality=quality)
    
    while len(img_byte_arr.getvalue()) > 100 * 827 and quality > 10:
        img_byte_arr = io.BytesIO()
        quality -= 5
        image.save(img_byte_arr, format='JPEG', quality=quality)
    
    return Image.open(img_byte_arr)

def save_images_to_pdf(images, output_pdf_path):
    images[0].save(output_pdf_path, save_all=True, append_images=images[1:], format='PDF')

def resize_pdf(input_pdf_path, output_pdf_path, dpi=400, quality=90):
    if not os.path.exists(input_pdf_path):
        print(f"El archivo no existe: {input_pdf_path}")
        return
    
    if os.path.getsize(input_pdf_path) == 0:
        print(f"El archivo está vacío: {input_pdf_path}")
        return
    
    try:
        images = pdf_to_images(input_pdf_path)

        if not images:
            print(f"Error: No se generaron imágenes del archivo: {input_pdf_path}")
            return

        grouped_images = []
        for page_num, img in enumerate(images):
            grayscale_img = convert_to_grayscale(img)
            compressed_img = compress_image(grayscale_img, quality)
            grouped_images.append(compressed_img)

            if (page_num + 1) % 20 == 0 or page_num == len(images) - 1:
                part_number = (page_num // 20) + 1
                output_pdf_path_with_part = os.path.splitext(output_pdf_path)[0] + f"_{part_number}.pdf"
                save_images_to_pdf(grouped_images, output_pdf_path_with_part)
                grouped_images = []

        save_images_to_pdf(grouped_images, output_pdf_path)
    except Exception as e:
        print(f"Guardando archivo:")

def renombrar_archivos_con_ori(ruta):
    for carpeta_raiz, carpetas, archivos in os.walk(ruta):
        for archivo in archivos:
            if "_ori" in archivo:
                ruta_actual = os.path.join(carpeta_raiz, archivo)
                nuevo_nombre = archivo.replace("_ori", "")
                ruta_nueva = os.path.join(carpeta_raiz, nuevo_nombre)
                os.rename(ruta_actual, ruta_nueva)

def dividir_pdf_en_partes(ruta_destino, max_size_kb=1750):
    try:
        tamaño_archivo = os.path.getsize(ruta_destino)

        if tamaño_archivo <= max_size_kb * 1024:
            resize_pdf(ruta_destino, ruta_destino)
            print("El archivo es menor o igual al tamaño máximo permitido, no se dividirá.")
            return
        
        with open(ruta_destino, 'rb') as archivo:
            lector = PyPDF2.PdfReader(archivo)
            num_paginas = len(lector.pages)

            tamaño_promedio_pagina = tamaño_archivo / num_paginas
            max_size_bytes = max_size_kb * 1024
            paginas_por_particion = int(max_size_bytes // tamaño_promedio_pagina)

            base_name = os.path.splitext(os.path.basename(ruta_destino))[0]
            output_dir = os.path.dirname(ruta_destino)

            for i in range(0, num_paginas, paginas_por_particion):
                pdf_writer = PyPDF2.PdfWriter()
                start_page = i
                end_page = min(i + paginas_por_particion, num_paginas)

                for page in range(start_page, end_page):
                    pdf_writer.add_page(lector.pages[page])

                output_ruta = os.path.join(output_dir, f"{base_name}_{start_page + 1}_{end_page}.pdf")
                with open(output_ruta, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)

                tamaño_particion = os.path.getsize(output_ruta)

                if tamaño_particion > max_size_bytes:
                    print(f"Advertencia: La partición supera el tamaño máximo permitido de {max_size_kb} KB.")
                else:
                    print(f"La partición cumple con el tamaño máximo permitido.")

        if os.path.exists(ruta_destino):
            try:
                os.remove(ruta_destino)
                print(f"Archivo eliminado: {ruta_destino}")
            except Exception as e:
                print(f"Error al eliminar el archivo: {e}")
        else:
            print(f"El archivo no existe: {ruta_destino}")
    except FileNotFoundError:
        print(f"El archivo {ruta_destino} no se encontró.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Leer rutas desde el archivo JSON
with open('config.json', 'r') as f:
    rutas_hospital = json.load(f)

# Procesar cada área del hospital
for area, rutas in rutas_hospital['areas'].items():
    ruta_base = rutas['ruta_base']
    ruta_renombre = rutas['ruta_renombre2']
    ruta_revision = rutas['ruta_revision2']
    excel_path = rutas['excel_path']

    # Crear carpetas si no existen
    os.makedirs(ruta_renombre, exist_ok=True)
    os.makedirs(ruta_revision, exist_ok=True)

    # Limpiar el contenido de las carpetas
    limpiar_contenido_carpeta(ruta_renombre)
    limpiar_contenido_carpeta(ruta_revision)

    # Cargar Excel con nombres para renombrar archivos
    df_1 = pd.read_excel(excel_path)

    # Procesar archivos en la carpeta base
    for carpeta_raiz, carpetas, archivos in os.walk(ruta_base):
        secuencial = 4
        for carpeta in carpetas:
            os.makedirs(os.path.join(ruta_renombre, carpeta), exist_ok=True)
        for archivo in archivos:
            ruta_actual = os.path.join(carpeta_raiz, archivo)

            nombre_actual = os.path.splitext(archivo)[0].strip().upper()
            fila = df_1[df_1['ANTES'].str.strip().str.upper() == nombre_actual]
            if not fila.empty:
                if fila['ACTUAL'].values[0].strip().upper() == "PLANILLA_INDIVIDUAL":
                    nuevo_nombre = "1_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                elif fila['ACTUAL'].values[0].strip().upper() in {"CODIGO_DE_VALIDACION", "CODIGO DE VALIDACION_2", "CODIGO DE VALIDACION_3"}:
                    nuevo_nombre = "2_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                elif fila['ACTUAL'].values[0].strip().upper() in {"COBERTURAS_1", "COBERTURAS_2", "COBERTURAS", "COBERTURAS_CONYUGE", "COBERTURA_PADRE", "COBERTURA_MADRE"}:
                    nuevo_nombre = "3_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                else:
                    if fila['ACTUAL'].values[0].strip().upper() in {"ACTA_ENTREGA", "ACTA ENTREGA_1", "ACTA ENTREGA_2"}:
                        nuevo_nombre = "4_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                        secuencial = 5
                    else:
                        nuevo_nombre = str(secuencial) + "_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                        secuencial += 1

                ruta_destino = os.path.join(ruta_renombre, os.path.basename(carpeta_raiz), nuevo_nombre)

                try:
                    # Verificar si el archivo ya existe antes de copiar
                    if not os.path.exists(ruta_destino):
                        shutil.copy(ruta_actual, ruta_destino)
                    else:
                        print(f"El archivo {nuevo_nombre} ya existe en {ruta_renombre}, no se copia.")
                except Exception as e:
                    print(f"Error al copiar {archivo}: {e}")

                # Redimensionar el PDF
                try:
                    resize_pdf(ruta_destino, ruta_destino)
                except Exception as e:
                    print(f"Error al redimensionar el PDF: {e}")

                # Dividir el PDF en partes
                try:
                    dividir_pdf_en_partes(ruta_destino)
                except Exception as e:
                    print(f"Error al dividir el PDF: {e}")

                # Eliminar el archivo original en la carpeta RENOMBRE
                try:
                    if os.path.exists(ruta_destino):
                        os.remove(ruta_destino)
                except Exception as e:
                    print(f"Error al eliminar archivo: {e}")

    # Renombrar archivos que tienen "_ori"
    renombrar_archivos_con_ori(ruta_renombre)

print("Proceso completado para todas las áreas del hospital.")
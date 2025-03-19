import fitz
import pandas as pd
import os
import shutil
import io
from PIL import Image
import PyPDF2
import json

# Función para limpiar el contenido de una carpeta (sin eliminar la carpeta)
def limpiar_contenido_carpeta(ruta_carpeta):
    for carpeta_raiz, carpetas, archivos in os.walk(ruta_carpeta):
        for archivo in archivos:
            os.remove(os.path.join(carpeta_raiz, archivo))
        for carpeta in carpetas:
            shutil.rmtree(os.path.join(carpeta_raiz, carpeta), ignore_errors=True)

# Función para eliminar carpetas de pacientes
def eliminar_carpeta_pacientes(ruta_base):
    for carpeta in os.listdir(ruta_base):
        ruta_area = os.path.join(ruta_base, carpeta)
        if os.path.isdir(ruta_area):
            shutil.rmtree(ruta_area, ignore_errors=True)
            print(f"Carpeta eliminada: {ruta_area}")

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
        for img in images:
            grayscale_img = convert_to_grayscale(img)
            compressed_img = compress_image(grayscale_img, quality)
            grouped_images.append(compressed_img)

        save_images_to_pdf(grouped_images, output_pdf_path)
    except Exception as e:
        print(f"Error al procesar el archivo PDF: {input_pdf_path}. Error: {e}")

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
        tamaño_archivo = os.path.getsize(ruta_destino)
        print(f"Tamaño del archivo original: {tamaño_archivo} bytes")

        if tamaño_archivo <= max_size_kb * 1024:
            print("El archivo es menor o igual al tamaño máximo permitido, no se dividirá.")
            return  # Retornar si no se divide
        
        with open(ruta_destino, 'rb') as archivo:
            lector = PyPDF2.PdfReader(archivo)
            num_paginas = len(lector.pages)

            tamaño_promedio_pagina = tamaño_archivo / num_paginas
            print(f"Tamaño promedio por página: {tamaño_promedio_pagina} bytes")

            max_size_bytes = max_size_kb * 1024
            paginas_por_particion = int(max_size_bytes // tamaño_promedio_pagina)
            print(f"Páginas por partición (aprox.): {paginas_por_particion}")

            base_name = os.path.splitext(os.path.basename(ruta_destino))[0]
            output_dir = os.path.dirname(ruta_destino)

            parte = 1  # Contador para numerar las partes
            for i in range(0, num_paginas, paginas_por_particion):
                pdf_writer = PyPDF2.PdfWriter()
                start_page = i
                end_page = min(i + paginas_por_particion, num_paginas)

                for page in range(start_page, end_page):
                    pdf_writer.add_page(lector.pages[page])

                output_ruta = os.path.join(output_dir, f"{base_name}_parte{parte}.pdf")
                with open(output_ruta, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)

                tamaño_particion = os.path.getsize(output_ruta)
                print(f"Tamaño de la partición: {tamaño_particion} bytes")
                print(f"nombre archivo particionado: {output_ruta}")

                parte += 1  # Incrementar el contador de partes

        # Eliminar el archivo original después de dividirlo
        os.remove(ruta_destino)
        print(f"Archivo original eliminado: {ruta_destino}")

    except FileNotFoundError:
        print(f"El archivo {ruta_destino} no se encontró.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Cargar configuraciones desde un archivo JSON
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# Obtener la ruta de procesados desde el archivo JSON
ruta_procesados = config.get("ruta_procesados", "")
print(f"Ruta de procesados: '{ruta_procesados}'")

# Abrir archivo de texto para registrar el histórico
with open('historico_archivos_reglas.txt', 'a', encoding='utf-8') as historico_file:
    # Procesar cada área definida en el archivo JSON
    for area, paths in config.get("areas", {}).items():
        print(f"\nProcesando área: {area}")
        
        ruta_base = paths.get("ruta_base", "")
        ruta_renombre = paths.get("ruta_renombre_reglas", "")
        ruta_revision = paths.get("ruta_revision_reglas", "")
        excel_path = paths.get("excel_path", "")
        
        os.makedirs(ruta_renombre, exist_ok=True)
        os.makedirs(ruta_revision, exist_ok=True)

        # Limpiar el contenido de las carpetas de renombre y revisión
        limpiar_contenido_carpeta(ruta_renombre)
        limpiar_contenido_carpeta(ruta_revision)

        # Cargar Excel con nombres para renombrar archivos
        df_1 = pd.read_excel(excel_path)

        # Procesar archivos en la carpeta base
        for carpeta_raiz, carpetas, archivos in os.walk(ruta_base):
            secuencial = 5
            for carpeta in carpetas:
                os.makedirs(os.path.join(ruta_renombre, carpeta), exist_ok=True)
            for archivo in archivos:
                ruta_actual = os.path.join(carpeta_raiz, archivo)

                nombre_actual = os.path.splitext(archivo)[0].strip().upper()
                fila = df_1[df_1['ANTES'].str.strip().str.upper() == nombre_actual]
                if not fila.empty:
                    if fila['ACTUAL'].values[0].strip().upper() == "PLANILLA_INDIVIDUAL":
                        nuevo_nombre = "1_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                    elif fila['ACTUAL'].values[0].strip().upper() in {"CODIGO_DE_VALIDACION","CODIGO DE VALIDACION_2","CODIGO DE VALIDACION_3"}:
                        nuevo_nombre = "2_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                    elif fila['ACTUAL'].values[0].strip().upper() in {"COBERTURAS_1","COBERTURAS_2","COBERTURAS","COBERTURAS_CONYUGE","COBERTURA_PADRE","COBERTURA_MADRE"}:
                        nuevo_nombre = "3_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"                
                    else:
                        if fila['ACTUAL'].values[0].strip().upper() in {"ACTA_ENTREGA","ACTA ENTREGA_1","ACTA ENTREGA_2"}:
                            nuevo_nombre = "4_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                        else:
                            nuevo_nombre = str(secuencial) + "_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                            secuencial += 1

                    ruta_destino = os.path.join(ruta_renombre, os.path.basename(carpeta_raiz), nuevo_nombre)
                    
                    try:
                        shutil.copy(ruta_actual, ruta_destino)
                        historico_file.write(f"{archivo} ; {ruta_destino}\n")  # Registrar en el histórico
                    except Exception as e:
                        print(f"Error al copiar {archivo}: {e}")

                    # Convertir a escala de grises y luego a PDF
                    output_pdf_path = os.path.splitext(ruta_destino)[0] + ".pdf"  # Guardar como PDF sin sufijo
                    try:
                        resize_pdf(ruta_destino, output_pdf_path)
                        print(f"Archivo convertido a escala de grises y guardado como: {output_pdf_path}")

                        # Evaluar si se necesita dividir el PDF
                        dividir_pdf_en_partes(output_pdf_path)

                    except Exception as e:
                        print(f"Error al procesar el archivo: {e}")

                else:
                    # Si no se encuentra el archivo en el Excel, mover a la carpeta de revisión
                    ruta_revision_destino = os.path.join(ruta_revision, archivo)
                    try:
                        shutil.copy(ruta_actual, ruta_revision_destino)
                        historico_file.write(f"{archivo} (no mapeado) ; {ruta_revision_destino}\n")  # Registrar en el histórico
                        print(f"Archivo no mapeado movido a revisión: {ruta_revision_destino}")
                    except Exception as e:
                        print(f"Error al mover {archivo} a revisión: {e}")

        # Renombrar archivos que tienen "_ori"
        renombrar_archivos_con_ori(ruta_renombre)

        # Copiar las carpetas de la ruta base a la nueva dirección
        for carpeta_raiz, carpetas, archivos in os.walk(ruta_base):
            for carpeta in carpetas:
                ruta_carpeta_origen = os.path.join(carpeta_raiz, carpeta)
                
                # Crear una subcarpeta para el área hospitalaria en PROCESADOS
                area_hospitalaria = os.path.basename(carpeta_raiz)  # Suponiendo que el nombre del área está en la ruta
                ruta_carpeta_destino = os.path.join(ruta_procesados, area_hospitalaria, carpeta)

                # Verificar si la carpeta de destino ya existe
                if os.path.exists(ruta_carpeta_destino):
                    # Si existe, agregar un número al final del nombre de la carpeta
                    contador = 1
                    nuevo_nombre = f"{carpeta}_{contador}"
                    while os.path.exists(os.path.join(ruta_procesados, area_hospitalaria, nuevo_nombre)):
                        contador += 1
                        nuevo_nombre = f"{carpeta}_{contador}"
                    ruta_carpeta_destino = os.path.join(ruta_procesados, area_hospitalaria, nuevo_nombre)

                os.makedirs(os.path.dirname(ruta_carpeta_destino), exist_ok=True)  # Crear la carpeta del área si no existe
                
                try:
                    shutil.copytree(ruta_carpeta_origen, ruta_carpeta_destino)  # Cambiar move por copytree
                    historico_file.write(f"Carpeta {carpeta} ; {ruta_carpeta_destino}\n")  # Registrar en el histórico
                    print(f"Carpeta copiada: {ruta_carpeta_origen} a {ruta_carpeta_destino}")
                except Exception as e:
                    print(f"Error al copiar la carpeta {ruta_carpeta_origen}: {e}")

        # Eliminar las carpetas de los pacientes en la ruta base
        eliminar_carpeta_pacientes(ruta_base)

print("Proceso completado para todas las áreas del hospital.")
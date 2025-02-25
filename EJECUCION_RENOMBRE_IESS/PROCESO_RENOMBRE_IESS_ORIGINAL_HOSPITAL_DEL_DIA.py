import fitz  # PyMuPDF
import pandas as pd
import os
import shutil
import io
from PIL import Image
import PyPDF2

# Función para limpiar el contenido de una carpeta
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
    
    # Verificar si el archivo PDF está vacío
    if os.path.getsize(input_pdf_path) == 0:
        print(f"El archivo está vacío: {input_pdf_path}")
        return
    
    try:
        # Convertir las páginas del PDF en imágenes
        images = pdf_to_images(input_pdf_path)

        # Si no se generaron imágenes, mostrar un mensaje de error
        if not images:
            print(f"Error: No se generaron imágenes del archivo: {input_pdf_path}")
            return

        # Procesar cada imagen y agrupar en lotes de 20
        grouped_images = []
        for page_num, img in enumerate(images):
            grayscale_img = convert_to_grayscale(img)
            compressed_img = compress_image(grayscale_img, quality)
            grouped_images.append(compressed_img)

            # Cada vez que se juntan 20 páginas, guardarlas en un PDF
            if (page_num + 1) % 600 == 0 or page_num == len(images) - 1:
                part_number = (page_num // 20) + 1
                output_pdf_path_with_part = os.path.splitext(output_pdf_path)[0] + f"_ori.pdf"
                save_images_to_pdf(grouped_images, output_pdf_path_with_part)
                grouped_images = []  # Resetear la lista para el siguiente grupo

        # Procesar cada imagen y convertirla a escala de grises
        grayscale_images = [convert_to_grayscale(img) for img in images]

        # Guardar las imágenes convertidas a PDF
        save_images_to_pdf(grayscale_images, output_pdf_path)
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
        tamaño_archivo = os.path.getsize(ruta_destino)
        print(f"Tamaño del archivo original: {tamaño_archivo} bytes")

        # Si el archivo es menor que el límite, no lo divide
        if tamaño_archivo <= max_size_kb * 1024:
            try:
                resize_pdf(ruta_destino,ruta_destino)
                print("El archivo es menor o igual al tamaño máximo permitido, no se dividirá.")
                return
            except Exception as e:
                print(f"error resize:  {e}")
                return
        
        with open(ruta_destino, 'rb') as archivo:
            lector = PyPDF2.PdfReader(archivo)
            num_paginas = len(lector.pages)

            # Calcular el tamaño promedio de una página
            tamaño_promedio_pagina = tamaño_archivo / num_paginas
            print(f"Tamaño promedio por página: {tamaño_promedio_pagina} bytes")

            # Calcular cuántas páginas caben en una partición de 2MB
            max_size_bytes = max_size_kb * 1024
            paginas_por_particion = int(max_size_bytes // tamaño_promedio_pagina)
            print(f"Páginas por partición (aprox.): {paginas_por_particion}")

            # Crear las particiones del PDF
            base_name = os.path.splitext(os.path.basename(ruta_destino))[0]
            output_dir = os.path.dirname(ruta_destino)

            for i in range(0, num_paginas, paginas_por_particion):
                pdf_writer = PyPDF2.PdfWriter()
                start_page = i
                end_page = min(i + paginas_por_particion, num_paginas)  # Última partición puede ser menor

                # Agregar páginas a la partición
                for page in range(start_page, end_page):
                    pdf_writer.add_page(lector.pages[page])

                # Guardar cada partición
                output_ruta = os.path.join(output_dir, f"{base_name}_{start_page + 1}_{end_page}.pdf")
                with open(output_ruta, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)

                # Verificar el tamaño de la partición creada
                tamaño_particion = os.path.getsize(output_ruta)
                print(f"Tamaño de la partición: {tamaño_particion} bytes")
                print(f"nombre archivo particionado: "+output_ruta)

                # Validar que la partición sea menor a 2MB
                if tamaño_particion > max_size_bytes:
                    print(f"Advertencia: La partición supera el tamaño máximo permitido de {max_size_kb} KB.")
                else:
                    print(f"La partición cumple con el tamaño máximo permitido.")

        # Eliminar el archivo copiado en la carpeta RENOMBRE
        if os.path.exists(ruta_destino):
            try:
                os.remove(ruta_destino)  # Eliminar solo desde la carpeta RENOMBRE
                print(f"Archivo eliminado: {ruta_destino}")
            except Exception:
                print(f"Error al eliminar el archivo: {e}")
        else:
            print(f"El archivo no existe: {ruta_destino}")
    except FileNotFoundError:
        print(f"El archivo {ruta_destino} no se encontró.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Rutas principales
ruta_base_hospital_del_dia = '/run/user/1000/gvfs/smb-share:server=192.168.10.249,share=carterauio/EJECUCION_RENOMBRE_IESS/HOSPITAL_DEL_DIA'
ruta_renombre_hospital_del_dia = '/run/user/1000/gvfs/smb-share:server=192.168.10.249,share=carterauio/EJECUCION_RENOMBRE_IESS/RENOMBRE_ORIGINAL_HOSPITAL_DEL_DIA'
ruta_revision_hospital_del_dia = '/run/user/1000/gvfs/smb-share:server=192.168.10.249,share=carterauio/EJECUCION_RENOMBRE_IESS/REVISION_ORIGINAL_HOSPITAL_DEL_DIA'
excel_path_hospital_del_dia = '/run/user/1000/gvfs/smb-share:server=192.168.10.249,share=carterauio/EJECUCION_RENOMBRE_IESS/RenombreIESS_HOSPITAL_DEL_DIA.xlsx'

# Crear carpetas si no existen
os.makedirs(ruta_renombre_hospital_del_dia, exist_ok=True)
os.makedirs(ruta_revision_hospital_del_dia, exist_ok=True)

# Limpiar el contenido de las carpetas
limpiar_contenido_carpeta(ruta_renombre_hospital_del_dia)
limpiar_contenido_carpeta(ruta_revision_hospital_del_dia)

# Cargar Excel con nombres para renombrar archivos
df_1 = pd.read_excel(excel_path_hospital_del_dia)

# Procesar archivos en la carpeta base emergencia
for carpeta_raiz, carpetas, archivos in os.walk(ruta_base_hospital_del_dia):
    secuencial = 4
    for carpeta in carpetas:
        os.makedirs(os.path.join(ruta_renombre_hospital_del_dia, carpeta), exist_ok=True)
    for archivo in archivos:
        ruta_actual = os.path.join(carpeta_raiz, archivo)
        #resize_pdf(ruta_actual,ruta_actual)

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
                    secuencial = 5
                else:
                    nuevo_nombre = str(secuencial) + "_" + fila['ACTUAL'].values[0].strip().upper() + ".pdf"
                    secuencial += 1

            ruta_actual = os.path.join(carpeta_raiz, archivo)
            ruta_destino = os.path.join(ruta_renombre_hospital_del_dia, os.path.basename(carpeta_raiz), nuevo_nombre)
            
            try:
                shutil.copy(ruta_actual, ruta_destino)
            except Exception as e:
                print(f"Error al copiar {archivo}: {e}")

            try:
                dividir_pdf_en_partes(ruta_destino)
            except Exception as e:
                print(f"error divide partes: {e}")

            try:
                os.remove(ruta_destino)
            except Exception as e:
                print(f"error al eliminar archivo: {e}")

# Renombrar archivos que tienen "_ori"
renombrar_archivos_con_ori(ruta_renombre_hospital_del_dia)

print("Proceso completado hospital del dia original.")
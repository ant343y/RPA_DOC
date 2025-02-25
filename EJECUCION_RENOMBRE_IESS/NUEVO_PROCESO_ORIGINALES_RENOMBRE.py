import os
import shutil
import json
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

# Cargar configuración desde config.json
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

# Función para limpiar contenido de una carpeta
def limpiar_contenido_carpeta(ruta):
    if os.path.exists(ruta):
        for archivo in os.listdir(ruta):
            ruta_archivo = os.path.join(ruta, archivo)
            if os.path.isfile(ruta_archivo):
                os.remove(ruta_archivo)
            elif os.path.isdir(ruta_archivo):
                shutil.rmtree(ruta_archivo)
    else:
        os.makedirs(ruta)

# Función para dividir PDF en partes basadas en tamaño máximo
def dividir_pdf_en_partes(ruta_pdf, max_size_kb=1750):
    try:
        base, ext = os.path.splitext(ruta_pdf)
        ruta_ori = f"{base}_ori{ext}"
        os.rename(ruta_pdf, ruta_ori)

        reader = PdfReader(ruta_ori)
        num_paginas = len(reader.pages)

        tamaño_archivo = os.path.getsize(ruta_ori)
        print(f"Tamaño del archivo original: {tamaño_archivo} bytes")

        if tamaño_archivo <= max_size_kb * 1024:
            print(f"El archivo {ruta_ori} no necesita dividirse.")
            return

        tamaño_promedio_pagina = tamaño_archivo / num_paginas
        max_size_bytes = max_size_kb * 1024
        paginas_por_particion = int(max_size_bytes // tamaño_promedio_pagina)
        print(f"Páginas por partición (aprox.): {paginas_por_particion}")

        for i in range(0, num_paginas, paginas_por_particion):
            writer = PdfWriter()
            for j in range(i, min(i + paginas_por_particion, num_paginas)):
                writer.add_page(reader.pages[j])

            nombre_parte = f"{base}_parte{i//paginas_por_particion + 1}.pdf"
            with open(nombre_parte, "wb") as output_pdf:
                writer.write(output_pdf)
            print(f"Parte {i//paginas_por_particion + 1} guardada como {nombre_parte}")

    except Exception as e:
        print(f"Error al dividir PDF {ruta_pdf}: {e}")

# Función para renombrar archivos con "_ORI"
def renombrar_archivos_con_ori(ruta_pdf):
    try:
        if os.path.exists(ruta_pdf):
            print(f"Renombrando archivo: {ruta_pdf}")
            base, ext = os.path.splitext(ruta_pdf)
            nuevo_nombre = f"{base}_ORI{ext}"
            os.rename(ruta_pdf, nuevo_nombre)
            print(f"Archivo renombrado a: {nuevo_nombre}")
        else:
            print(f"Error: El archivo no existe para renombrar: {ruta_pdf}")
    except Exception as e:
        print(f"Error al renombrar {ruta_pdf}: {e}")

# Función para generar nuevo nombre basado en el Excel
def generar_nuevo_nombre(fila, secuencial):
    nombre_actualizado = fila['ACTUAL'].values[0].strip().upper()
    
    # Asignar números específicos a ciertos nombres
    if nombre_actualizado == "PLANILLA_INDIVIDUAL":
        nuevo_nombre = f"1_{nombre_actualizado}.pdf"
    elif nombre_actualizado in {"CODIGO_DE_VALIDACION", "CODIGO DE VALIDACION_2", "CODIGO DE VALIDACION_3"}:
        nuevo_nombre = f"2_{nombre_actualizado}.pdf"
    elif nombre_actualizado in {"COBERTURAS_1", "COBERTURAS_2", "COBERTURAS", "COBERTURAS_CONYUGE", "COBERTURA_PADRE", "COBERTURA_MADRE"}:
        nuevo_nombre = f"3_{nombre_actualizado}.pdf"
    elif nombre_actualizado in {"ACTA_ENTREGA", "ACTA ENTREGA_1", "ACTA ENTREGA_2"}:
        nuevo_nombre = f"4_{nombre_actualizado}.pdf"
        secuencial = 5  # Reiniciar el secuencial si se encuentra un acta de entrega
    else:
        nuevo_nombre = f"{secuencial}_{nombre_actualizado}.pdf"
        secuencial += 1

    return nuevo_nombre, secuencial

# Procesar cada área definida en config.json
for area, paths in config.get("areas", {}).items():
    print(f"\nProcesando área: {area}")

    ruta_base = paths.get("ruta_base", "")
    ruta_renombre = paths.get("ruta_renombre", "")
    ruta_revision = paths.get("ruta_revision", "")
    excel_path = paths.get("excel_path", "")

    os.makedirs(ruta_renombre, exist_ok=True)
    os.makedirs(ruta_revision, exist_ok=True)

    limpiar_contenido_carpeta(ruta_renombre)
    limpiar_contenido_carpeta(ruta_revision)

    if not os.path.exists(excel_path):
        print(f"Error: No se encontró el archivo {excel_path}")
        continue

    df = pd.read_excel(excel_path)

    for carpeta_raiz, carpetas, archivos in os.walk(ruta_base):
        secuencial = 5  # Comenzar desde 5, ya que 1-4 están ocupados
        for carpeta in carpetas:
            os.makedirs(os.path.join(ruta_renombre, carpeta), exist_ok=True)

        for archivo in archivos:
            ruta_actual = os.path.join(carpeta_raiz, archivo)
            nombre_actual = os.path.splitext(archivo)[0].strip().upper()
            fila = df[df['ANTES'].str.strip().str.upper() == nombre_actual]

            if not fila.empty:
                nuevo_nombre, secuencial = generar_nuevo_nombre(fila, secuencial)
                ruta_destino = os.path.join(ruta_renombre, os.path.basename(carpeta_raiz), nuevo_nombre)

                try:
                    carpeta_paciente = os.path.dirname(ruta_destino)
                    os.makedirs(carpeta_paciente, exist_ok=True)

                    if not os.path.exists(ruta_destino):
                        shutil.copy(ruta_actual, ruta_destino)
                        print(f"Archivo copiado a: {ruta_destino}")
                        dividir_pdf_en_partes(ruta_destino)
                        renombrar_archivos_con_ori(ruta_destino)
                    else:
                        print(f"El archivo {nuevo_nombre} ya existe en {carpeta_paciente}, no se copia.")
                except Exception as e:
                    print(f"Error al procesar {archivo}: {e}")
            else:
                shutil.copy(ruta_actual, os.path.join(ruta_revision, archivo))
                print(f"Archivo no encontrado en Excel, copiado a revisión: {archivo}")

print("Proceso completado para todas las áreas del hospital.")
import pandas as pd
import os
import shutil
import json

# Función para limpiar el contenido de una carpeta
def limpiar_contenido_carpeta(ruta_carpeta):
    for carpeta_raiz, carpetas, archivos in os.walk(ruta_carpeta, topdown=False):
        for archivo in archivos:
            os.remove(os.path.join(carpeta_raiz, archivo))
        for carpeta in carpetas:
            shutil.rmtree(os.path.join(ruta_carpeta, carpeta), ignore_errors=True)

# Cargar configuraciones desde un archivo JSON
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# Abrir archivo de texto para registrar el histórico
with open('historico_archivos_original.txt', 'a', encoding='utf-8') as historico_file:
    # Procesar cada área definida en el archivo JSON
    for area, paths in config.get("areas", {}).items():
        print(f"\nProcesando área: {area}")
        
        ruta_base = paths.get("ruta_base", "")
        ruta_renombre = paths.get("ruta_renombre", "")
        ruta_revision = paths.get("ruta_revision", "")
        excel_path = paths.get("excel_path", "")
        
        os.makedirs(ruta_renombre, exist_ok=True)
        os.makedirs(ruta_revision, exist_ok=True)

        # Limpiar el contenido de las carpetas
        limpiar_contenido_carpeta(ruta_renombre)
        limpiar_contenido_carpeta(ruta_revision)

        # Cargar Excel con nombres para renombrar archivos
        df_1 = pd.read_excel(excel_path)

        # Procesar archivos en la carpeta base
        for carpeta_raiz, carpetas, archivos in os.walk(ruta_base):
            secuencial = 5  # Comenzar desde 5 porque 1-4 son números reservados
            for carpeta in carpetas:
                os.makedirs(os.path.join(ruta_renombre, carpeta), exist_ok=True)
            for archivo in archivos:
                ruta_actual = os.path.join(carpeta_raiz, archivo)

                nombre_actual = os.path.splitext(archivo)[0].strip().upper()
                fila = df_1[df_1['ANTES'].str.strip().str.upper() == nombre_actual]
                if not fila.empty:
                    actual_value = fila['ACTUAL'].values[0].strip().upper()
                    if actual_value == "PLANILLA_INDIVIDUAL":
                        nuevo_nombre = "1_" + actual_value + ".pdf"
                    elif actual_value in {"CODIGO_DE_VALIDACION", "CODIGO DE VALIDACION_2", "CODIGO DE VALIDACION_3"}:
                        nuevo_nombre = "2_" + actual_value + ".pdf"
                    elif actual_value in {"COBERTURAS_1", "COBERTURAS_2", "COBERTURAS", "COBERTURAS_CONYUGE", "COBERTURA_PADRE", "COBERTURA_MADRE"}:
                        nuevo_nombre = "3_" + actual_value + ".pdf"
                    elif actual_value in {"ACTA_ENTREGA", "ACTA ENTREGA_1", "ACTA ENTREGA_2"}:
                        nuevo_nombre = "4_" + actual_value + ".pdf"
                    else:
                        nuevo_nombre = str(secuencial) + "_" + actual_value + ".pdf"
                        secuencial += 1  # Incrementar solo si no es un número reservado

                    ruta_destino = os.path.join(ruta_renombre, os.path.basename(carpeta_raiz), nuevo_nombre)
                    
                    try:
                        shutil.copy(ruta_actual, ruta_destino)
                        historico_file.write(f"{archivo} ; {ruta_destino}\n")  # Registrar en el histórico
                        print(f"Archivo renombrado y copiado a: {ruta_destino}")
                    except Exception as e:
                        print(f"Error al copiar {archivo}: {e}")

                else:
                    # Si no se encuentra el archivo en el Excel, mover a la carpeta de revisión
                    ruta_revision_destino = os.path.join(ruta_revision, archivo)
                    try:
                        shutil.copy(ruta_actual, ruta_revision_destino)
                        historico_file.write(f"{archivo} (no mapeado) ; {ruta_revision_destino}\n")  # Registrar en el histórico
                        print(f"Archivo no mapeado movido a revisión: {ruta_revision_destino}")
                    except Exception as e:
                        print(f"Error al mover {archivo} a revisión: {e}")

print("Proceso completado para todas las áreas del hospital.")
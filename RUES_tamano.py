import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

UVT = 47065  # UVT para 2024

# -------------------------------------------
# Clasificador de tamaño empresarial
# -------------------------------------------
def clasificar_por_tamano(rama, ingresos):
    if pd.isna(ingresos):
        return "Sin información"
    ingresos = float(ingresos)
    if rama == "Servicio":
        if ingresos <= UVT * 32988:
            return "Microempresa"
        elif ingresos <= UVT * 131951:
            return "Pequeña empresa"
        elif ingresos <= UVT * 483034:
            return "Mediana empresa"
        else:
            return "Gran empresa"
    elif rama == "Comercio":
        if ingresos <= UVT * 44769:
            return "Microempresa"
        elif ingresos <= UVT * 431196:
            return "Pequeña empresa"
        elif ingresos <= UVT * 2160692:
            return "Mediana empresa"
        else:
            return "Gran empresa"
    else:  # Industria
        if ingresos <= UVT * 23563:
            return "Microempresa"
        elif ingresos <= UVT * 204995:
            return "Pequeña empresa"
        elif ingresos <= UVT * 1736565:
            return "Mediana empresa"
        else:
            return "Gran empresa"

# -------------------------------------------
# Interfaz selección de archivos
# -------------------------------------------
def cargar_archivo(titulo, tipo_archivo):
    root = tk.Tk()
    root.withdraw()
    archivo = filedialog.askopenfilename(title=titulo, filetypes=[(tipo_archivo[0], tipo_archivo[1])])
    if not archivo:
        messagebox.showerror("Error", f"No se seleccionó archivo para {titulo}")
        exit()
    return archivo

def seleccionar_directorio_salida():
    root = tk.Tk()
    root.withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona carpeta para guardar archivos de salida")
    if not carpeta:
        messagebox.showerror("Error", "No se seleccionó carpeta de salida.")
        exit()
    return carpeta

# -------------------------------------------
# Script principal
# -------------------------------------------
def main():
    archivo_rues = cargar_archivo("Selecciona archivo RUES (.txt)", ("Archivos TXT", "*.txt"))
    archivo_ciiu = cargar_archivo("Selecciona archivo CIIU_PIB (.csv)", ("Archivos CSV", "*.csv"))
    archivo_provincias = cargar_archivo("Selecciona Municipios_Provincias (.csv)", ("Archivos CSV", "*.csv"))
    carpeta_salida = seleccionar_directorio_salida()

    # Cargar base RUES
    df = pd.read_csv(archivo_rues, sep='|', encoding='latin1', dtype=str)
    df.columns = df.columns.str.strip().str.upper()
    df['INGRESOS_ACTIVIDAD_ORDINARIA'] = pd.to_numeric(df['INGRESOS_ACTIVIDAD_ORDINARIA'], errors='coerce')
    df['COD_CIIU_ACT_ECON_PRI'] = pd.to_numeric(df['COD_CIIU_ACT_ECON_PRI'], errors='coerce')

    # Cargar y preparar CIIU_PIB
    ciiu_pib = pd.read_csv(archivo_ciiu, sep=';', encoding='utf-8', dtype=str)
    ciiu_pib.columns = ciiu_pib.columns.str.upper().str.strip()
    ciiu_pib['CIIU_INT'] = pd.to_numeric(ciiu_pib['CIIU_INT'], errors='coerce')

    # Merge con CIIU_PIB
    df = df.merge(ciiu_pib, how='left', left_on='COD_CIIU_ACT_ECON_PRI', right_on='CIIU_INT')
    df['TAMANO_EMPRESA'] = df.apply(lambda x: clasificar_por_tamano(x['RAMA'], x['INGRESOS_ACTIVIDAD_ORDINARIA']), axis=1)

    # Cargar Municipios_Provincias
    provincias = pd.read_csv(archivo_provincias, sep=';', encoding='utf-8', dtype=str)
    provincias.columns = provincias.columns.str.upper().str.strip()
    df = df.merge(provincias, how='left', left_on='MUNICIPIO_COMERCIAL', right_on='MPIO_CDPMP')

    # Filtro por Bogotá y Cundinamarca
    df_filtrado = df[df['DPTO_COMERCIAL'].isin(['BOGOTA', 'CUNDINAMARCA'])]

    # Columnas solicitadas
    columnas_principales = [
        'MATRICULA', 'RAZON_SOCIAL','CODIGO_CLASE_IDENTIFICACION', 'NUMERO_IDENTIFICACION', 'FECHA_RENOVACION',
        'ULTIMO_ANO_RENOVADO', 'FECHA_MATRICULA', 'FECHA_CANCELACION', 'CODIGO_ORGANIZACION_JURIDICA',
        'DESC_ORGANIZACION_JURIDICA', 'CODIGO_TIPO_SOCIEDAD', 'DESC_TIPO_SOCIEDAD',
        'CODIGO_CATEGORIA_MATRICULA', 'DESC_CATEGORIA_MATRICULA', 'CODIGO_ESTADO_MATRICULA',
        'MUNICIPIO_COMERCIAL', 'DPTO_COMERCIAL', 'COD_CIIU_ACT_ECON_PRI', 'ACTIVOS_TOTAL',
        'COD_CIIU_ACT_ECON_SEC', 'CIIU3', 'CIIU4', 'FECHA_INICIO_ACT_ECON_PRI',
        'CIIU_MAYORES_INGRESOS', 'INGRESOS_ACTIVIDAD_ORDINARIA', 'UTILIDAD_PERDIDA_OPERACIONAL',
        'RESULTADO_DEL_PERIODO', 'EMPLEADOS', 'TAMANO_EMPRESA'
    ]

    columnas_adicionales = [col for col in ciiu_pib.columns if col not in columnas_principales] + \
                           [col for col in provincias.columns if col not in columnas_principales]

    columnas_finales = columnas_principales + columnas_adicionales
    columnas_finales = [col for col in columnas_finales if col in df_filtrado.columns]  # asegurar existencia

    df_final = df_filtrado[columnas_finales]

    # Exportar
    df_final.to_csv(os.path.join(carpeta_salida, "02_empresas_filtradas_BOG_CUND.csv"),
                    index=False, sep=';', encoding='utf-8-sig')

    df_final.to_excel(os.path.join(carpeta_salida, "03_empresas_filtradas_BOG_CUND.xlsx"),
                      index=False, engine='openpyxl')

    messagebox.showinfo("Proceso finalizado", f"Exportación completada.\nArchivos guardados en:\n{carpeta_salida}")

# -------------------------------------------
# Ejecutar
# -------------------------------------------
if __name__ == "__main__":
    main()

# ============================================================
# PROCESADOR DE CURSADAS
# Versión 4.0
#
# Autor: Francisco Lombroni



# ============================================================
# MÓDULO 0
# VENTANA DE PROGRESO
# ============================================================

from tkinter import ttk

ventana_progreso = None
barra_progreso = None
etiqueta_estado = None


def crear_ventana_progreso():
    """
    Crea la ventana de progreso del programa.
    """

    global ventana_progreso
    global barra_progreso
    global etiqueta_estado

    ventana_progreso = tk.Tk()

    ventana_progreso.title(TITULO)

    ventana_progreso.geometry("500x180")

    ventana_progreso.resizable(False, False)

    ventana_progreso.attributes("-topmost", True)

    ttk.Label(
        ventana_progreso,
        text="Procesador de Cursadas",
        font=("Segoe UI", 12, "bold")
    ).pack(pady=(15, 10))

    etiqueta_estado = ttk.Label(
        ventana_progreso,
        text="Inicializando...",
        font=("Segoe UI", 10)
    )

    etiqueta_estado.pack()

    barra_progreso = ttk.Progressbar(
        ventana_progreso,
        orient="horizontal",
        mode="determinate",
        length=420
    )

    barra_progreso.pack(pady=20)

    barra_progreso["maximum"] = 100

    barra_progreso["value"] = 0

    ttk.Label(
        ventana_progreso,
        text="Espere por favor...",
        font=("Segoe UI", 9)
    ).pack()

    ventana_progreso.update()

def mostrar_error(mensaje):

    cerrar_ventana_progreso()

    messagebox.showerror(
        TITULO,
        mensaje
    )


def mostrar_advertencia(mensaje):

    messagebox.showwarning(
        TITULO,
        mensaje
    )


def mostrar_info(mensaje):

    messagebox.showinfo(
        TITULO,
        mensaje
    )
  
def actualizar_progreso(mensaje, porcentaje):
    """
    Actualiza el mensaje y la barra de progreso.
    """

    etiqueta_estado.config(text=mensaje)

    barra_progreso["value"] = porcentaje

    ventana_progreso.update_idletasks()


def cerrar_ventana_progreso():
    """
    Cierra la ventana de progreso si todavía existe.
    """

    global ventana_progreso

    if ventana_progreso is None:
        return

    try:
        if ventana_progreso.winfo_exists():
            ventana_progreso.destroy()
    except tk.TclError:
        pass

    ventana_progreso = None


# ============================================================
# MÓDULO 0.5
# FUNCIONES AUXILIARES
# ============================================================

def obtener_lista_alumnos(df, columna_alumno="ALUMNO"):
    """
    Devuelve un diccionario con los apellidos normalizados
    y la fila correspondiente dentro del DataFrame.
    """

    alumnos = {}

    for indice, valor in df[columna_alumno].items():

        if valor is None:
            continue

        nombre = str(valor).strip()

        if nombre == "":
            continue

        apellido = normalizar_texto(
            obtener_apellido(nombre)
        )

        alumnos[apellido] = indice

    return alumnos
    
def quitar_tildes(texto):
    """
    Elimina las tildes de un texto.
    """

    if texto is None:
        return ""

    texto = str(texto)

    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


# ------------------------------------------------------------

def normalizar_texto(texto):
    """
    Convierte un texto a un formato estándar:
    - Sin tildes
    - Sin espacios al principio y al final
    - En mayúsculas
    """

    texto = quitar_tildes(texto)

    return texto.strip().upper()


# ------------------------------------------------------------

def buscar_fila_encabezado(df):
    """
    Busca automáticamente la fila donde comienza la tabla.
    """

    for fila in range(min(15, len(df))):

        valores = [
            normalizar_texto(x)
            for x in df.iloc[fila].fillna("")
        ]

        if normalizar_texto(COLUMNA_ALUMNO) in valores:
            return fila

    return None


# ------------------------------------------------------------

def convertir_nota(valor):
    """
    Convierte una nota al formato numérico.

    A -> 0

    Devuelve None si está vacía.
    """

    if pd.isna(valor):
        return None

    texto = str(valor).strip().upper()

    if texto == "":
        return None

    if texto == "A":
        return 0

    try:
        return float(texto)
    except:
        return None


# ------------------------------------------------------------

def es_nota_valida(valor):
    """
    Verifica si una nota es válida.
    """

    nota = convertir_nota(valor)

    if nota is None:
        return False

    return NOTA_MINIMA <= nota <= NOTA_MAXIMA

def obtener_apellido(nombre_completo):
    """
    Devuelve el primer apellido de un alumno.
    """

    if nombre_completo is None:
        return ""

    partes = normalizar_texto(nombre_completo).split()

    if len(partes) == 0:
        return ""

    return partes[0]

# ------------------------------------------------------------

def obtener_nota(fila, columna):
    """
    Obtiene una nota de una fila del DataFrame.

    Devuelve:
        - Número (0 a 10)
        - None si la columna no existe o la celda está vacía.

    Convierte automáticamente:
        A -> 0
    """

    if columna not in fila:
        return None

    return convertir_nota(fila[columna])


# ------------------------------------------------------------

def calcular_promedio(nota1, nota2):
    """
    Calcula el promedio de dos notas.

    Devuelve el resultado con un decimal.
    """

    if nota1 is None or nota2 is None:
        return None

    return round((nota1 + nota2) / 2, 1)


# ============================================================
# MÓDULO 1 - IMPORTACIONES
# ============================================================

import os
import sys
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# Compatibilidad con archivos ODS
try:
    import odf
except ImportError:
    odf = None


# ============================================================
# MÓDULO 1 - CONFIGURACIÓN GENERAL
# ============================================================

VERSION = "4.0"

TITULO = f"Procesador de Cursadas V{VERSION}"


# ============================================================
# HOJAS DEL ARCHIVO
# ============================================================

HOJA_REPORTE = "REPORTE1"
HOJA_LIB = "LIB"
HOJA_REG2ET = "REG 2ET"
HOJA_INAS = "INAS"
HOJA_INAS2 = "INAS2"
HOJA_RCP = "RCP"
HOJA_PRN = "PRN"


HOJAS_OBLIGATORIAS = [
    HOJA_REPORTE,
    HOJA_LIB,
    HOJA_REG2ET,
    HOJA_INAS,
    HOJA_INAS2
]


# ============================================================
# COLUMNAS
# ============================================================

COLUMNA_ALUMNO = "ALUMNO"

COLUMNA_PROPUESTA = "PROPUESTA"

COLUMNAS_TP = [
    "TP1",
    "TP2",
    "TP3",
    "TP4"
]

COLUMNAS_PARCIALES = [
    "1P",
    "2P"
]

COLUMNAS_PROMEDIOS = [
    "PRM1",
    "PRM2"
]

COLUMNA_INAS = "INAS"


COLUMNAS_REPORTE = [
    COLUMNA_ALUMNO,
    "TP1",
    "TP2",
    "PRM1",
    "1P",
    "TP3",
    "TP4",
    "PRM2",
    "2P",
    COLUMNA_INAS
]


# ============================================================
# COLORES
# ============================================================

COLOR_VERDE = "92D050"
COLOR_AMARILLO = "FFD966"
COLOR_ROJO = "FF6666"

COLOR_ENCABEZADO = "D9EAD3"


# ============================================================
# VALIDACIONES
# ============================================================

NOTA_MINIMA = 0
NOTA_MAXIMA = 10

PROMEDIO_APROBACION = 4
PROMEDIO_PROMOCION = 7

NOTA_MINIMA_PROMOCION = 6

VALORES_ESPECIALES = [
    "A"
]


# ============================================================
# VARIABLES GLOBALES
# ============================================================

archivo_notas = None
archivo_inasistencias = None

workbook = None
worksheet = None

df_notas = None
df_reporte = None
df_inas1 = None
df_inas2 = None

alumnos_rcp = []

apellidos_repetidos = []

errores = []


# ============================================================
# FIN DEL MÓDULO 1
# ============================================================

print("=" * 55)
print(TITULO)
print("Módulo 1 cargado correctamente.")
print("=" * 55)

# ============================================================
# MÓDULO 2
# SELECCIÓN DE ARCHIVOS
# ============================================================

def seleccionar_archivo(titulo):
    """
    Abre una ventana para seleccionar un archivo Excel u ODS.
    Devuelve la ruta completa o None si el usuario cancela.
    """

    ruta = filedialog.askopenfilename(
        title=titulo,
        filetypes=[
            ("Archivos Excel y ODS", "*.xlsx *.ods"),
            ("Excel", "*.xlsx"),
            ("OpenDocument", "*.ods")
        ]
    )

    if not ruta:
        return None

    return ruta


# ------------------------------------------------------------

def seleccionar_archivos():

    global archivo_notas
    global archivo_inasistencias

    actualizar_progreso(
        "Seleccionando archivo de notas...",
        5
    )

    root = tk.Tk()
    root.withdraw()

    archivo_notas = seleccionar_archivo(
        "Seleccione el archivo de NOTAS"
    )

    if archivo_notas is None:
        mostrar_advertencia("Operación cancelada.")
        return False

    actualizar_progreso(
        "Seleccionando archivo de inasistencias...",
        10
    )

    archivo_inasistencias = seleccionar_archivo(
        "Seleccione el archivo de INASISTENCIAS"
    )

    if archivo_inasistencias is None:
        mostrar_advertencia("Operación cancelada.")
        return False

    mostrar_info("Archivos seleccionados correctamente.")

    actualizar_progreso(
        "Archivos seleccionados.",
        15
    )

    return True

# ============================================================
# MÓDULO 3
# VALIDACIÓN DE ARCHIVOS
# ============================================================

def validar_archivo_notas():

    global archivo_notas

    try:

        if archivo_notas.lower().endswith(".xlsx"):

            df = pd.read_excel(
                archivo_notas,
                header=None
            )

        else:

            df = pd.read_excel(
                archivo_notas,
                engine="odf",
                header=None
            )

    except Exception as e:

        mostrar_error(f"No fue posible abrir el archivo de notas.\n\n{e}")

        return False

    encontrado = False

    for fila in range(min(15, len(df))):

        for valor in df.iloc[fila]:

            if pd.isna(valor):
                continue

            if str(valor).strip().upper() == COLUMNA_PROPUESTA:

                encontrado = True
                break

        if encontrado:
            break

    if not encontrado:

        mostrar_error(
           "El archivo seleccionado no corresponde al archivo de NOTAS.\n\n"
            "No se encontró la columna PROPUESTA."
        )

        return False

    return True


# ------------------------------------------------------------

def validar_archivo_inasistencias():

    global archivo_inasistencias

    try:

        libro = load_workbook(
            archivo_inasistencias,
            read_only=True
        )

    except Exception as e:

        mostrar_error(
            f"No fue posible abrir el archivo de inasistencias.\n\n{e}"
        )

        return False

    hojas = libro.sheetnames

    faltantes = []

    for hoja in HOJAS_OBLIGATORIAS:

        if hoja not in hojas:

            faltantes.append(hoja)

    if faltantes:

        texto = "\n".join(faltantes)

        mostrar_error(
           "El archivo de inasistencias no es válido.\n\n"
            "Faltan las siguientes hojas:\n\n"
            f"{texto}"
        )

        return False

    return True


# ------------------------------------------------------------

def validar_cantidad_alumnos():

    """
    Se implementará en el Módulo 4,
    cuando ya tengamos ambos DataFrames cargados.
    """

    return True


# ------------------------------------------------------------

def validar_archivos():

    if not validar_archivo_notas():
        return False

    if not validar_archivo_inasistencias():
        return False

    actualizar_progreso(
        "Archivos validados correctamente.",
        45
    )

    return True

# ============================================================
# MÓDULO 4
# LECTURA DE DATOS
# ============================================================

def cargar_archivo_notas():

    global df_notas

    actualizar_progreso(
        "Leyendo archivo de notas...",
        50
    )

    try:

        if archivo_notas.lower().endswith(".xlsx"):

            bruto = pd.read_excel(
                archivo_notas,
                header=None
            )

        else:

            bruto = pd.read_excel(
                archivo_notas,
                engine="odf",
                header=None
            )

    except Exception as e:

        mostrar_error(
            f"No fue posible abrir el archivo de notas.\n\n{e}"
        )

        return False

    fila = buscar_fila_encabezado(bruto)

    if fila is None:

        mostrar_error(
            "No fue posible localizar el encabezado del archivo."
        )

        return False

    if archivo_notas.lower().endswith(".xlsx"):

        df_notas = pd.read_excel(
            archivo_notas,
            header=fila
        )

    else:

        df_notas = pd.read_excel(
            archivo_notas,
            engine="odf",
            header=fila
        )

    return True


def cargar_archivo_inasistencias():

    global df_reporte
    global df_inas1
    global df_inas2

    actualizar_progreso(
        "Leyendo archivo de inasistencias...",
        60
    )

    df_reporte = pd.read_excel(
        archivo_inasistencias,
        sheet_name=HOJA_REPORTE
    )

    df_inas1 = pd.read_excel(
        archivo_inasistencias,
        sheet_name=HOJA_INAS
    )

    df_inas2 = pd.read_excel(
        archivo_inasistencias,
        sheet_name=HOJA_INAS2
    )

    return True

def validar_cantidad_alumnos():

    cantidad_notas = len(df_notas)

    cantidad_inas = len(df_reporte)

    if cantidad_notas != cantidad_inas:

        mostrar_error(
            "La cantidad de alumnos no coincide.\n\n"
            f"Notas: {cantidad_notas}\n"
            f"Inasistencias: {cantidad_inas}"
        )

        return False

    return True

# ------------------------------------------------------------

def leer_datos():
    """
    Lee ambos archivos y verifica que puedan procesarse.
    """

    if not cargar_archivo_notas():
        return False

    if not cargar_archivo_inasistencias():
        return False

    if not validar_cantidad_alumnos():
        return False

    actualizar_progreso(
        "Datos cargados correctamente.",
        65
    )

    return True





if __name__ == "__main__":

    crear_ventana_progreso()

    if not seleccionar_archivos():
        cerrar_ventana_progreso()
        sys.exit()

    if not validar_archivos():
        cerrar_ventana_progreso()
        sys.exit()

    if not leer_datos():
        cerrar_ventana_progreso()
        sys.exit()

    actualizar_progreso(
        "Preparando validación de notas...",
        70
    )

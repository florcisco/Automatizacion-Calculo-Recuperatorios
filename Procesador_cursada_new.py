# ============================================================
# PROCESADOR DE CURSADAS
# Versión 4.0
#
# Autor: Francisco Lombroni



# ============================================================
# MÓDULO 0
# VENTANA DE PROGRESO
# ============================================================

from tkinter import tk, simpledialog

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
    Obtiene los alumnos de un DataFrame.

    Devuelve un diccionario cuya clave es el apellido
    normalizado y cuyo valor contiene la fila y el
    nombre completo del alumno.
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

        alumnos[apellido] = {
            "fila": indice,
            "nombre": nombre
        }

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

limite_faltas = None
archivo_salida = None

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

lista_lib = []
inas_segunda_etapa = {}

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

def validar_coincidencia_alumnos():

    actualizar_progreso(
        "Verificando coincidencia de alumnos...",
        35
    )

    alumnos_notas = obtener_lista_alumnos(df_notas)
    alumnos_reporte = obtener_lista_alumnos(df_reporte)

    faltan_en_reporte = sorted(
        set(alumnos_notas.keys()) -
        set(alumnos_reporte.keys())
    )

    faltan_en_notas = sorted(
        set(alumnos_reporte.keys()) -
        set(alumnos_notas.keys())
    )

    if not faltan_en_reporte and not faltan_en_notas:
        return True

    mensaje = ""

    if faltan_en_reporte:

        mensaje += (
            "Alumnos presentes en NOTAS y ausentes "
            "en REPORTE:\n\n"
        )

        for alumno in faltan_en_reporte:
            mensaje += f"• {alumno}\n"

    if faltan_en_notas:

        if mensaje != "":
            mensaje += "\n-----------------------------\n\n"

        mensaje += (
            "Alumnos presentes en REPORTE y ausentes "
            "en NOTAS:\n\n"
        )

        for alumno in faltan_en_notas:
            mensaje += f"• {alumno}\n"

    mostrar_error(mensaje)

    return False
    
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

def leer_datos():
    """
    Lee ambos archivos y verifica que puedan procesarse.
    """

    if not cargar_archivo_notas():
        return False

    if not cargar_archivo_inasistencias():
        return False
        
    if not validar_coincidencia_alumnos():
        return False

    actualizar_progreso(
        "Datos cargados correctamente.",
        65
    )

    return True

# ============================================================
# MÓDULO 5
# PROCESAMIENTO DE SEGUNDA ETAPA
# ============================================================

def calcular_prm2():
    """
    Calcula PRM2 como el promedio de TP3 y TP4.
    El resultado se redondea a un decimal.
    """

    global df_notas

    actualizar_progreso(
        "Calculando PRM2...",
        45
    )

    for indice in df_notas.index:

        tp3 = obtener_nota(
            df_notas.at[indice, "TP3"]
        )

        tp4 = obtener_nota(
            df_notas.at[indice, "TP4"]
        )

        prm2 = round((tp3 + tp4) / 2, 1)

        df_notas.at[indice, "PRM2"] = prm2

    return True


def determinar_habilitados_segundo_parcial():
    """
    Determina qué alumnos pueden rendir el segundo parcial.

    PRM2 >= 4:
        Puede rendir.

    PRM2 < 4:
        No puede rendir y pasa a LIB.
    """

    global df_notas
    global lista_lib

    actualizar_progreso(
        "Verificando alumnos habilitados para el segundo parcial...",
        50
    )

    lista_lib = []

    for indice in df_notas.index:

        alumno = df_notas.at[indice, "ALUMNO"]

        prm2 = obtener_nota(
            df_notas.at[indice, "PRM2"]
        )

        if prm2 < 4:

            lista_lib.append({
                "alumno": alumno,
                "fila": indice,
                "prm2": prm2
            })

    return True


def calcular_inasistencias_segunda_etapa():
    """
    Calcula las inasistencias correspondientes
    únicamente a la segunda etapa.

    INAS segunda etapa = INAS2 - INAS primera etapa.
    """

    global df_reporte
    global df_inas
    global df_inas2
    global inas_segunda_etapa

    actualizar_progreso(
        "Calculando inasistencias de la segunda etapa...",
        55
    )

    inas_segunda_etapa = {}

    alumnos_reporte = obtener_lista_alumnos(
        df_reporte
    )

    alumnos_inas = obtener_lista_alumnos(
        df_inas
    )

    alumnos_inas2 = obtener_lista_alumnos(
        df_inas2
    )

    for apellido in alumnos_reporte.keys():

        valor_inas = 0
        valor_inas2 = 0

        # ----------------------------------------
        # INAS PRIMERA ETAPA
        # ----------------------------------------

        if apellido in alumnos_inas:

            datos = alumnos_inas[apellido]

            if isinstance(datos, dict):
                fila = datos.get("fila")
            else:
                fila = datos

            try:
                valor_inas = obtener_nota(
                    df_inas.loc[fila, "INAS"]
                )
            except Exception:
                valor_inas = 0

        # ----------------------------------------
        # INAS ACUMULADAS / SEGUNDA ETAPA
        # ----------------------------------------

        if apellido in alumnos_inas2:

            datos = alumnos_inas2[apellido]

            if isinstance(datos, dict):
                fila = datos.get("fila")
            else:
                fila = datos

            try:
                valor_inas2 = obtener_nota(
                    df_inas2.loc[fila, "INAS"]
                )
            except Exception:
                valor_inas2 = 0

        # ----------------------------------------
        # DIFERENCIA
        # ----------------------------------------

        resultado = valor_inas2 - valor_inas

        if resultado < 0:
            resultado = 0

        inas_segunda_etapa[apellido] = resultado

    return True

def solicitar_limite_faltas():
    """
    Solicita al usuario el límite de faltas permitido
    para la materia.
    """

    global limite_faltas

    while True:

        limite = simpledialog.askinteger(
            TITULO,
            "¿Cuál es el límite de faltas para esta materia?",
            minvalue=0
        )

        if limite is None:

            mostrar_error(
                "No se indicó el límite de faltas.\n\n"
                "La operación será cancelada."
            )

            return False

        limite_faltas = limite

        return True

def preparar_datos_segunda_etapa():
    """
    Ejecuta todo el procesamiento correspondiente
    a la segunda etapa.
    """

    actualizar_progreso(
        "Procesando segunda etapa...",
        40
    )

    if not solicitar_limite_faltas():
        return False

    if not calcular_prm2():
        return False

    if not determinar_habilitados_segundo_parcial():
        return False

    if not calcular_inasistencias_segunda_etapa():
        return False

    actualizar_progreso(
        "Procesamiento de segunda etapa completado.",
        65
    )

    return True


# ============================================================
# MÓDULO 6
# GENERACIÓN DEL ARCHIVO FINAL
# ============================================================

import os
import shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font


def detectar_columnas_notas():
    """
    Detecta las columnas del archivo de notas a partir
    del texto de sus encabezados.

    No depende de que las columnas se llamen exactamente
    TP1, TP2, TP3, TP4 o 1P.
    """

    columnas = {}

    for columna in df_notas.columns:

        texto = str(columna).upper().strip()

        if texto == "ALUMNO":
            columnas["ALUMNO"] = columna

        elif texto.startswith("TP1"):
            columnas["TP1"] = columna

        elif texto.startswith("TP2"):
            columnas["TP2"] = columna

        elif texto.startswith("TP3"):
            columnas["TP3"] = columna

        elif texto.startswith("TP4"):
            columnas["TP4"] = columna

        elif "PRIMER PARCIAL" in texto:
            columnas["1P"] = columna

        elif "PROPUESTA" in texto:
            columnas["PROPUESTA"] = columna

    columnas_faltantes = []

    for nombre in ["ALUMNO", "TP1", "TP2", "1P", "TP3", "TP4"]:
        if nombre not in columnas:
            columnas_faltantes.append(nombre)

    if columnas_faltantes:

        mostrar_error(
            "No se pudieron identificar todas las columnas "
            "necesarias del archivo de notas.\n\n"
            "Faltan:\n"
            + "\n".join(columnas_faltantes)
        )

        return None

    return columnas


def obtener_datos_notas():
    """
    Convierte el DataFrame de notas en un diccionario
    indexado por apellido normalizado.

    Esto permite relacionar las notas con REPORTE,
    INAS e INAS2.
    """

    columnas = detectar_columnas_notas()

    if columnas is None:
        return None

    datos = {}

    for indice in df_notas.index:

        alumno = df_notas.at[
            indice,
            columnas["ALUMNO"]
        ]

        if alumno is None:
            continue

        alumno = str(alumno).strip()

        if alumno == "":
            continue

        apellido = normalizar_texto(
            obtener_apellido(alumno)
        )

        datos[apellido] = {

            "alumno": alumno,

            "TP1": obtener_nota(
                df_notas.at[
                    indice,
                    columnas["TP1"]
                ]
            ),

            "TP2": obtener_nota(
                df_notas.at[
                    indice,
                    columnas["TP2"]
                ]
            ),

            "1P": obtener_nota(
                df_notas.at[
                    indice,
                    columnas["1P"]
                ]
            ),

            "TP3": obtener_nota(
                df_notas.at[
                    indice,
                    columnas["TP3"]
                ]
            ),

            "TP4": obtener_nota(
                df_notas.at[
                    indice,
                    columnas["TP4"]
                ]
            )
        }

    return datos


def crear_archivo_salida():
    """
    Crea una copia del archivo de inasistencias
    que será utilizada como archivo final.

    De esta manera se conserva el formato original.
    """

    actualizar_progreso(
        "Creando archivo de salida...",
        70
    )

    carpeta = os.path.dirname(
        archivo_inasistencias
    )

    nombre = os.path.splitext(
        os.path.basename(archivo_inasistencias)
    )[0]

    archivo_salida = os.path.join(
        carpeta,
        nombre + "_procesado.xlsx"
    )

    shutil.copy2(
        archivo_inasistencias,
        archivo_salida
    )

    return archivo_salida


def eliminar_hojas_no_necesarias(wb):
    """
    Deja únicamente las cuatro hojas necesarias:

        REPORTE
        LIB
        INAS
        INAS2
    """

    hojas_necesarias = {
        "REPORTE",
        "LIB",
        "INAS",
        "INAS2"
    }

    for nombre in list(wb.sheetnames):

        if nombre not in hojas_necesarias:

            del wb[nombre]


def obtener_fila_encabezado(ws):
    """
    Busca la fila que contiene ALUMNO.

    Se utiliza para no depender rígidamente
    de una posición determinada.
    """

    for fila in range(1, ws.max_row + 1):

        for columna in range(1, ws.max_column + 1):

            valor = ws.cell(
                fila,
                columna
            ).value

            if valor is None:
                continue

            texto = str(valor).strip().upper()

            if texto == "ALUMNO":
                return fila

    return None


def obtener_columnas_reporte(ws):
    """
    Obtiene las columnas de la hoja REPORTE
    según sus encabezados.
    """

    fila_encabezado = obtener_fila_encabezado(ws)

    if fila_encabezado is None:

        mostrar_error(
            "No se encontró la fila de encabezados "
            "en la hoja REPORTE."
        )

        return None

    columnas = {}

    for columna in range(
        1,
        ws.max_column + 1
    ):

        valor = ws.cell(
            fila_encabezado,
            columna
        ).value

        if valor is None:
            continue

        texto = str(valor).strip().upper()

        if texto == "ALUMNO":
            columnas["ALUMNO"] = columna

        elif texto == "TP1":
            columnas["TP1"] = columna

        elif texto == "TP2":
            columnas["TP2"] = columna

        elif texto == "PRM1":
            columnas["PRM1"] = columna

        elif texto == "1P":
            columnas["1P"] = columna

        elif texto == "TP3":
            columnas["TP3"] = columna

        elif texto == "TP4":
            columnas["TP4"] = columna

        elif texto == "PRM2":
            columnas["PRM2"] = columna

        elif texto == "INAS":
            columnas["INAS"] = columna

    return fila_encabezado, columnas


def preparar_reporte(ws):
    """
    Prepara la hoja REPORTE para recibir
    la información de la segunda etapa.
    """

    resultado = obtener_columnas_reporte(ws)

    if resultado is None:
        return False

    fila_encabezado, columnas = resultado

    actualizar_progreso(
        "Preparando hoja REPORTE...",
        75
    )

    columnas_necesarias = [
        "ALUMNO",
        "TP1",
        "TP2",
        "PRM1",
        "1P",
        "TP3",
        "TP4",
        "PRM2",
        "INAS",
        "OBSERVACION"
    ]

    ultima_columna = ws.max_column

    for nombre in columnas_necesarias:

        if nombre not in columnas:

            ultima_columna += 1

            ws.cell(
                fila_encabezado,
                ultima_columna
            ).value = nombre

            columnas[nombre] = ultima_columna

    return fila_encabezado, columnas


def escribir_notas_reporte(ws):
    """
    Escribe las notas del archivo de notas
    en la hoja REPORTE.
    """

    resultado = preparar_reporte(ws)

    if resultado is False:
        return False

    fila_encabezado, columnas = resultado

    datos_notas = obtener_datos_notas()

    if datos_notas is None:
        return False

    actualizar_progreso(
        "Copiando notas al REPORTE...",
        80
    )

    for fila in range(
        fila_encabezado + 1,
        ws.max_row + 1
    ):

        valor_alumno = ws.cell(
            fila,
            columnas["ALUMNO"]
        ).value

        if valor_alumno is None:
            continue

        alumno = str(
            valor_alumno
        ).strip()

        if alumno == "":
            continue

        apellido = normalizar_texto(
            obtener_apellido(alumno)
        )

        if apellido not in datos_notas:
            continue

        datos = datos_notas[apellido]

        ws.cell(
            fila,
            columnas["TP1"]
        ).value = datos["TP1"]

        ws.cell(
            fila,
            columnas["TP2"]
        ).value = datos["TP2"]

        ws.cell(
            fila,
            columnas["1P"]
        ).value = datos["1P"]

        ws.cell(
            fila,
            columnas["TP3"]
        ).value = datos["TP3"]

        ws.cell(
            fila,
            columnas["TP4"]
        ).value = datos["TP4"]

        prm1 = round(
            (
                datos["TP1"]
                +
                datos["TP2"]
            ) / 2,
            1
        )

        prm2 = round(
            (
                datos["TP3"]
                +
                datos["TP4"]
            ) / 2,
            1
        )

        ws.cell(
            fila,
            columnas["PRM1"]
        ).value = prm1

        ws.cell(
            fila,
            columnas["PRM2"]
        ).value = prm2

    return True


def escribir_inasistencias_reporte(ws):
    """
    Escribe en REPORTE las inasistencias
    correspondientes únicamente a la segunda etapa.

    INAS = INAS2 - INAS
    """

    resultado = preparar_reporte(ws)

    if resultado is False:
        return False

    fila_encabezado, columnas = resultado

    actualizar_progreso(
        "Copiando inasistencias...",
        85
    )

    alumnos_reporte = obtener_lista_alumnos(
        df_reporte
    )

    for fila in range(
        fila_encabezado + 1,
        ws.max_row + 1
    ):

        valor_alumno = ws.cell(
            fila,
            columnas["ALUMNO"]
        ).value

        if valor_alumno is None:
            continue

        alumno = str(
            valor_alumno
        ).strip()

        if alumno == "":
            continue

        apellido = normalizar_texto(
            obtener_apellido(alumno)
        )

        if apellido in inas_segunda_etapa:

            ws.cell(
                fila,
                columnas["INAS"]
            ).value = inas_segunda_etapa[
                apellido
            ]

        else:

            ws.cell(
                fila,
                columnas["INAS"]
            ).value = 0

    return True

def marcar_alumnos_excedidos_en_faltas(ws):
    """
    Marca en REPORTE a los alumnos que:

        PRM2 >= 4
        INAS > limite_faltas

    Escribe la observación:
        ALUMNO REGULAR EXCEDIDO EN FALTAS

    Toda la fuente de la fila queda en rojo.
    """

    resultado = preparar_reporte(ws)

    if resultado is False:
        return False

    fila_encabezado, columnas = resultado

    actualizar_progreso(
        "Verificando límite de faltas...",
        88
    )

    fuente_roja = Font(
        color="FF0000"
    )

    for fila in range(
        fila_encabezado + 1,
        ws.max_row + 1
    ):

        valor_alumno = ws.cell(
            fila,
            columnas["ALUMNO"]
        ).value

        if valor_alumno is None:
            continue

        alumno = str(
            valor_alumno
        ).strip()

        if alumno == "":
            continue

        prm2 = obtener_nota(
            ws.cell(
                fila,
                columnas["PRM2"]
            ).value
        )

        inas = obtener_nota(
            ws.cell(
                fila,
                columnas["INAS"]
            ).value
        )

        # ----------------------------------------------------
        # Alumno regular excedido en faltas
        # ----------------------------------------------------

        if (
            prm2 >= 4
            and inas > limite_faltas
        ):

            ws.cell(
                fila,
                columnas["OBSERVACION"]
            ).value = (
                "ALUMNO REGULAR EXCEDIDO EN FALTAS"
            )

            # Toda la fila en fuente roja
            for columna in range(
                1,
                ws.max_column + 1
            ):

                celda = ws.cell(
                    fila,
                    columna
                )

                celda.font = fuente_roja

        else:

            ws.cell(
                fila,
                columnas["OBSERVACION"]
            ).value = None

    return True

def preparar_lib(ws):
    """
    Completa la hoja LIB con los alumnos
    que tienen PRM2 menor que 4.

    Si no hay alumnos, escribe SIN ALUMNOS.

    Los alumnos libres se muestran con fondo rojo.
    """

    actualizar_progreso(
        "Preparando hoja LIB...",
        90
    )

    # --------------------------------------------------------
    # Limpiar solamente el contenido de la hoja.
    # No se elimina la hoja para conservar su formato.
    # --------------------------------------------------------

    for fila in range(
        1,
        ws.max_row + 1
    ):

        for columna in range(
            1,
            ws.max_column + 1
        ):

            ws.cell(
                fila,
                columna
            ).value = None

    # --------------------------------------------------------
    # Si no hay alumnos libres
    # --------------------------------------------------------

    if not lista_lib:

        ws.cell(
            1,
            1
        ).value = "SIN ALUMNOS"

        return True

    # --------------------------------------------------------
    # Color rojo para alumnos libres
    # --------------------------------------------------------

    relleno_rojo = PatternFill(
        fill_type="solid",
        fgColor="FF0000"
    )

    # --------------------------------------------------------
    # Encabezados
    # --------------------------------------------------------

    ws.cell(
        1,
        1
    ).value = "ALUMNO"

    ws.cell(
        1,
        2
    ).value = "PRM2"

    # --------------------------------------------------------
    # Alumnos libres
    # --------------------------------------------------------

    fila = 2

    for alumno in lista_lib:

        ws.cell(
            fila,
            1
        ).value = alumno["alumno"]

        ws.cell(
            fila,
            2
        ).value = alumno["prm2"]

        # Pintar toda la fila de rojo
        for columna in range(
            1,
            ws.max_column + 1
        ):

            ws.cell(
                fila,
                columna
            ).fill = relleno_rojo

        fila += 1

    return True

def generar_archivo_final():
    """
    Genera el archivo final completo.
    """

    global archivo_salida

    actualizar_progreso(
        "Generando archivo final...",
        70
    )

    try:

        archivo_salida = crear_archivo_salida()

        wb = load_workbook(
            archivo_salida
        )

        # ----------------------------------------------------
        # Eliminar hojas que no necesitamos
        # ----------------------------------------------------

        eliminar_hojas_no_necesarias(wb)

        # ----------------------------------------------------
        # Verificar hojas necesarias
        # ----------------------------------------------------

        hojas_necesarias = [
            "REPORTE",
            "LIB",
            "INAS",
            "INAS2"
        ]

        for nombre in hojas_necesarias:

            if nombre not in wb.sheetnames:

                mostrar_error(
                    "No se encontró la hoja "
                    f"'{nombre}' en el archivo "
                    "de inasistencias."
                )

                wb.close()

                return False

        # ----------------------------------------------------
        # REPORTE
        # ----------------------------------------------------

        ws_reporte = wb["REPORTE"]

        if not escribir_notas_reporte(
            ws_reporte
        ):
            wb.close()
            return False

        if not escribir_inasistencias_reporte(
            ws_reporte
        ):
            wb.close()
            return False

        if not marcar_alumnos_excedidos_en_faltas(
            ws_reporte
        ):
            wb.close()
            return False

        # ----------------------------------------------------
        # LIB
        # ----------------------------------------------------

        ws_lib = wb["LIB"]

        if not preparar_lib(
            ws_lib
        ):
            wb.close()
            return False

        # ----------------------------------------------------
        # Guardar
        # ----------------------------------------------------

        actualizar_progreso(
            "Guardando archivo final...",
            95
        )

        wb.save(
            archivo_salida
        )

        wb.close()

        actualizar_progreso(
            "Proceso terminado correctamente.",
            100
        )

        return True

    except Exception as error:

        mostrar_error(
            "Se produjo un error al generar "
            "el archivo final:\n\n"
            f"{error}"
        )

        return False





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

import os
import re
import copy
import tkinter as tk

from tkinter import filedialog, messagebox, simpledialog

import openpyxl

from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment
)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def seleccionar_archivo(titulo):
    """
    Abre una ventana para seleccionar un archivo Excel.
    """

    ruta = filedialog.askopenfilename(
        title=titulo,
        filetypes=[
            ("Archivos Excel", "*.xlsx *.xlsm"),
            ("Todos los archivos", "*.*")
        ]
    )

    return ruta


def extraer_dni(valor):
    """
    Extrae el DNI de un texto.

    Puede encontrar:

        DNI 12345678

    o directamente:

        12345678
    """

    if valor is None:
        return None

    texto = str(valor)

    # DNI escrito explícitamente
    match = re.search(
        r"\bDNI\s*(\d{7,9})\b",
        texto,
        re.IGNORECASE
    )

    if match:
        return match.group(1)

    # Si directamente es un número
    match = re.fullmatch(
        r"\s*(\d{7,9})\s*",
        texto
    )

    if match:
        return match.group(1)

    return None


def convertir_nota(valor):
    """
    Convierte una nota a número.

    Si aparece 'A', se considera 0.
    """

    if valor is None or valor == "":
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip().upper()

    if texto == "A":
        return 0.0

    try:
        return float(
            texto.replace(",", ".")
        )

    except ValueError:
        return 0.0


def promedio(nota1, nota2):
    """
    Calcula el promedio de dos notas.
    """

    n1 = convertir_nota(nota1)
    n2 = convertir_nota(nota2)

    return (n1 + n2) / 2


def convertir_inasistencia(valor):
    """
    Convierte una cantidad de inasistencias a número.
    """

    if valor is None or valor == "":
        return 0.0

    try:
        return float(
            str(valor).replace(",", ".")
        )

    except ValueError:
        return 0.0


def numero_limpio(valor):
    """
    Si el número es entero, devuelve un int.

    Si tiene decimal, devuelve float.

    Ejemplos:

        10.0 -> 10
        8.0  -> 8
        8.5  -> 8.5
    """

    valor = float(valor)

    if valor.is_integer():
        return int(valor)

    return valor


# ============================================================
# PROCESAMIENTO PRINCIPAL
# ============================================================

def procesar(archivo_notas, archivo_inas, limite_inasistencias):

    print()
    print("==========================================")
    print("INICIANDO PROCESAMIENTO")
    print("==========================================")
    print()

    # --------------------------------------------------------
    # ABRIR ARCHIVOS
    # --------------------------------------------------------

    wb_inas = openpyxl.load_workbook(
        archivo_inas
    )

    wb_notas = openpyxl.load_workbook(
        archivo_notas,
        data_only=True
    )


    # --------------------------------------------------------
    # OBTENER HOJAS
    # --------------------------------------------------------

    ws_reporte = wb_inas["REPORTE1"]
    ws_inas = wb_inas["INAS"]
    ws_inas2 = wb_inas["INAS2"]
    ws_notas = wb_notas["Reporte"]


    # ========================================================
    # CARGAR NOTAS
    # ========================================================

    print("Cargando notas...")

    notas = {}

    for fila in range(
        11,
        ws_notas.max_row + 1
    ):

        dni = extraer_dni(
            ws_notas.cell(
                fila,
                1
            ).value
        )

        if dni is None:
            continue

        tp3 = ws_notas.cell(
            fila,
            5
        ).value

        tp4 = ws_notas.cell(
            fila,
            6
        ).value

        notas[dni] = {
            "tp3": tp3,
            "tp4": tp4,
            "prm2": promedio(
                tp3,
                tp4
            )
        }


    # ========================================================
    # CARGAR INASISTENCIAS - PRIMERA ETAPA
    # ========================================================

    print(
        "Cargando inasistencias "
        "de la primera etapa..."
    )

    inas1 = {}

    for fila in range(
        2,
        ws_inas.max_row + 1
    ):

        legajo = ws_inas.cell(
            fila,
            2
        ).value

        faltas = ws_inas.cell(
            fila,
            4
        ).value

        if legajo is None:
            continue

        dni = str(legajo).strip()

        inas1[dni] = convertir_inasistencia(
            faltas
        )


    # ========================================================
    # CARGAR INASISTENCIAS - SEGUNDA ETAPA
    # ========================================================

    print(
        "Cargando inasistencias "
        "de la segunda etapa..."
    )

    inas2 = {}

    for fila in range(
        2,
        ws_inas2.max_row + 1
    ):

        legajo = ws_inas2.cell(
            fila,
            2
        ).value

        faltas = ws_inas2.cell(
            fila,
            4
        ).value

        if legajo is None:
            continue

        dni = str(legajo).strip()

        inas2[dni] = convertir_inasistencia(
            faltas
        )


    # ========================================================
    # ELIMINAR HOJAS ANTERIORES
    # ========================================================

    print(
        "Eliminando hojas LIB y REG 2ET..."
    )

    for nombre in [
        "LIB",
        "REG 2ET"
    ]:

        if nombre in wb_inas.sheetnames:
            del wb_inas[nombre]


    # ========================================================
    # LIMPIAR DESDE COLUMNA E
    # ========================================================

    print(
        "Limpiando columnas desde E..."
    )

    for fila in range(
        1,
        ws_reporte.max_row + 1
    ):

        for columna in range(
            5,
            ws_reporte.max_column + 1
        ):

            ws_reporte.cell(
                fila,
                columna
            ).value = None


    # ========================================================
    # ENCABEZADOS
    # ========================================================

    ws_reporte["E10"] = "TP3"
    ws_reporte["F10"] = "TP4"
    ws_reporte["G10"] = "PRM2"
    ws_reporte["H10"] = "INAS"
    ws_reporte["I10"] = "OBSERVACIONES"


    # ========================================================
    # COLORES
    # ========================================================

    fondo_blanco = PatternFill(
        fill_type="solid",
        fgColor="FFFFFF"
    )

    fondo_rojo = PatternFill(
        fill_type="solid",
        fgColor="FF0000"
    )

    fuente_negra = Font(
        name="Arial",
        size=10,
        color="000000"
    )

    fuente_roja = Font(
        name="Arial",
        size=10,
        color="FF0000"
    )


    # ========================================================
    # COPIAR ESTILO DE COLUMNA D A E-I
    # ========================================================

    for columna in range(
        5,
        10
    ):

        for fila in range(
            1,
            ws_reporte.max_row + 1
        ):

            origen = ws_reporte.cell(
                fila,
                4
            )

            destino = ws_reporte.cell(
                fila,
                columna
            )

            if origen.has_style:

                destino._style = copy.copy(
                    origen._style
                )

            if origen.alignment:

                destino.alignment = copy.copy(
                    origen.alignment
                )


    # ========================================================
    # PROCESAR ALUMNOS
    # ========================================================

    print(
        "Procesando alumnos..."
    )

    alumnos_procesados = 0
    alumnos_sin_notas = 0
    alumnos_libres = 0
    alumnos_excedidos = 0


    for fila in range(
        11,
        ws_reporte.max_row + 1
    ):

        # ----------------------------------------------------
        # OBTENER DNI
        # ----------------------------------------------------

        dni = extraer_dni(
            ws_reporte.cell(
                fila,
                1
            ).value
        )

        if dni is None:
            continue


        # ----------------------------------------------------
        # PINTAR TODA LA FILA DEL ALUMNO
        # DE BLANCO Y FUENTE NEGRA
        # ----------------------------------------------------

        for columna in range(
            1,
            10
        ):

            celda = ws_reporte.cell(
                fila,
                columna
            )

            celda.fill = copy.copy(
                fondo_blanco
            )

            celda.font = copy.copy(
                fuente_negra
            )


        # ====================================================
        # TP3 / TP4 / PRM2
        # ====================================================

        if dni in notas:

            tp3 = convertir_nota(
                notas[dni]["tp3"]
            )

            tp4 = convertir_nota(
                notas[dni]["tp4"]
            )

            prm2 = promedio(
                tp3,
                tp4
            )


            # -----------------------------------------------
            # TP3
            # -----------------------------------------------

            ws_reporte.cell(
                fila,
                5
            ).value = numero_limpio(
                tp3
            )


            # -----------------------------------------------
            # TP4
            # -----------------------------------------------

            ws_reporte.cell(
                fila,
                6
            ).value = numero_limpio(
                tp4
            )


            # -----------------------------------------------
            # PRM2
            # -----------------------------------------------

            prm2 = numero_limpio(
                prm2
            )

            ws_reporte.cell(
                fila,
                7
            ).value = prm2


        else:

            tp3 = None
            tp4 = None
            prm2 = None

            ws_reporte.cell(
                fila,
                5
            ).value = None

            ws_reporte.cell(
                fila,
                6
            ).value = None

            ws_reporte.cell(
                fila,
                7
            ).value = None

            alumnos_sin_notas += 1


        # ====================================================
        # INASISTENCIAS
        #
        # INAS = INAS2 - INAS
        # ====================================================

        faltas_1 = inas1.get(
            dni,
            0
        )

        faltas_2 = inas2.get(
            dni,
            0
        )

        diferencia = faltas_2 - faltas_1

        diferencia = numero_limpio(
            diferencia
        )

        ws_reporte.cell(
            fila,
            8
        ).value = diferencia


        # ====================================================
        # OBSERVACIONES
        # ====================================================

        ws_reporte.cell(
            fila,
            9
        ).value = None


        # ====================================================
        # CASO 1:
        #
        # PRM2 MENOR A 4
        #
        # -> FONDO ROJO
        # -> SE AGREGA A LIB
        # ====================================================

        if prm2 is not None and float(prm2) < 4:

            # Pintar toda la fila de rojo
            for columna in range(
                1,
                10
            ):

                ws_reporte.cell(
                    fila,
                    columna
                ).fill = copy.copy(
                    fondo_rojo
                )

                # Fuente negra
                ws_reporte.cell(
                    fila,
                    columna
                ).font = copy.copy(
                    fuente_negra
                )

            alumnos_libres += 1


        # ====================================================
        # CASO 2:
        #
        # PRM2 MAYOR A 4
        # Y SUPERA EL LIMITE DE INASISTENCIAS
        #
        # -> OBSERVACIÓN
        # -> FUENTE ROJA
        # ====================================================

        elif (
            prm2 is not None
            and float(prm2) >= 4
            and float(diferencia) > limite_inasistencias
        ):

            ws_reporte.cell(
                fila,
                9
            ).value = (
                "ALUMNO REGULAR EXCEDIDO EN FALTAS"
            )


            # Poner toda la fila con fuente roja
            for columna in range(
                1,
                10
            ):

                ws_reporte.cell(
                    fila,
                    columna
                ).font = copy.copy(
                    fuente_roja
                )

            alumnos_excedidos += 1


        # ====================================================
        # FORMATO GENERAL
        # ====================================================

        for columna in range(
            5,
            9
        ):

            ws_reporte.cell(
                fila,
                columna
            ).number_format = "General"


        alumnos_procesados += 1


    # ========================================================
    # FORMATO GENERAL DE E-I
    # ========================================================

    for fila in range(
        1,
        ws_reporte.max_row + 1
    ):

        for columna in range(
            5,
            9
        ):

            ws_reporte.cell(
                fila,
                columna
            ).number_format = "General"


    # ========================================================
    # ANCHO DE COLUMNAS
    # ========================================================

    ws_reporte.column_dimensions[
        "E"
    ].width = 12

    ws_reporte.column_dimensions[
        "F"
    ].width = 12

    ws_reporte.column_dimensions[
        "G"
    ].width = 12

    ws_reporte.column_dimensions[
        "H"
    ].width = 12

    ws_reporte.column_dimensions[
        "I"
    ].width = 42


    # ========================================================
    # CREAR HOJA LIB
    # ========================================================

    print(
        "Creando hoja LIB..."
    )

    ws_lib = wb_inas.create_sheet(
        "LIB"
    )


    # --------------------------------------------------------
    # ENCABEZADO
    # --------------------------------------------------------

    ws_lib["A1"] = "ALUMNOS LIBRES"

    ws_lib["A1"].font = Font(
        name="Arial",
        size=12,
        bold=True,
        color="000000"
    )

    ws_lib["A1"].alignment = Alignment(
        horizontal="center"
    )


    # --------------------------------------------------------
    # ENCABEZADO DE COLUMNA
    # --------------------------------------------------------

    ws_lib["A3"] = "ALUMNO"

    ws_lib["A3"].font = Font(
        name="Arial",
        size=10,
        bold=True,
        color="000000"
    )


    # ========================================================
    # COPIAR NOMBRES DE LOS ALUMNOS LIBRES
    # ========================================================

    fila_lib = 4

    for fila in range(
        11,
        ws_reporte.max_row + 1
    ):

        dni = extraer_dni(
            ws_reporte.cell(
                fila,
                1
            ).value
        )

        if dni is None:
            continue

        if dni not in notas:
            continue

        prm2 = promedio(
            notas[dni]["tp3"],
            notas[dni]["tp4"]
        )

        if prm2 < 4:

            nombre = ws_reporte.cell(
                fila,
                1
            ).value

            ws_lib.cell(
                fila_lib,
                1
            ).value = nombre

            ws_lib.cell(
                fila_lib,
                1
            ).font = copy.copy(
                fuente_negra
            )

            fila_lib += 1


    # ========================================================
    # FORMATO DE HOJA LIB
    # ========================================================

    ws_lib.column_dimensions[
        "A"
    ].width = 45


    # ========================================================
    # NOMBRE DEL ARCHIVO DE SALIDA
    # ========================================================

    carpeta = os.path.dirname(
        archivo_inas
    )

    nombre = os.path.basename(
        archivo_inas
    )

    nombre_base, extension = os.path.splitext(
        nombre
    )

    archivo_salida = os.path.join(
        carpeta,
        f"{nombre_base} procesado{extension}"
    )


    # ========================================================
    # GUARDAR
    # ========================================================

    print(
        "Guardando archivo..."
    )

    wb_inas.save(
        archivo_salida
    )


    # ========================================================
    # RESULTADO
    # ========================================================

    print()
    print(
        "=========================================="
    )
    print(
        "PROCESO TERMINADO"
    )
    print(
        "=========================================="
    )

    print(
        f"Alumnos procesados: "
        f"{alumnos_procesados}"
    )

    print(
        f"Alumnos libres: "
        f"{alumnos_libres}"
    )

    print(
        f"Alumnos regulares excedidos "
        f"en faltas: {alumnos_excedidos}"
    )

    print(
        f"Alumnos sin notas encontradas: "
        f"{alumnos_sin_notas}"
    )

    print()
    print(
        f"Límite de inasistencias: "
        f"{limite_inasistencias}"
    )

    print()
    print(
        "Archivo generado:"
    )

    print(
        archivo_salida
    )


    messagebox.showinfo(
        "Proceso terminado",
        "El archivo fue procesado correctamente.\n\n"

        f"Alumnos procesados: "
        f"{alumnos_procesados}\n"

        f"Alumnos libres: "
        f"{alumnos_libres}\n"

        f"Regulares excedidos en faltas: "
        f"{alumnos_excedidos}\n"

        f"Alumnos sin notas: "
        f"{alumnos_sin_notas}\n\n"

        f"Límite de inasistencias: "
        f"{limite_inasistencias}\n\n"

        "Archivo generado:\n"
        f"{archivo_salida}"
    )


# ============================================================
# INTERFAZ
# ============================================================

def main():

    # Crear ventana principal oculta
    root = tk.Tk()
    root.withdraw()


    # ========================================================
    # PRIMERO: ARCHIVO DE NOTAS
    # ========================================================

    messagebox.showinfo(
        "Paso 1 de 3",
        "Primero seleccioná el archivo de NOTAS."
    )

    archivo_notas = seleccionar_archivo(
        "Seleccionar archivo de NOTAS"
    )


    if not archivo_notas:

        messagebox.showwarning(
            "Cancelado",
            "No se seleccionó el archivo de notas."
        )

        root.destroy()

        return


    # ========================================================
    # SEGUNDO: ARCHIVO DE INASISTENCIAS
    # ========================================================

    messagebox.showinfo(
        "Paso 2 de 3",
        "Ahora seleccioná el archivo de INASISTENCIAS."
    )

    archivo_inas = seleccionar_archivo(
        "Seleccionar archivo de INASISTENCIAS"
    )


    if not archivo_inas:

        messagebox.showwarning(
            "Cancelado",
            "No se seleccionó el archivo de inasistencias."
        )

        root.destroy()

        return


    # ========================================================
    # TERCERO: LÍMITE DE INASISTENCIAS
    # ========================================================

    limite = simpledialog.askfloat(
        "Paso 3 de 3",
        "¿Cuál es el límite de inasistencias?",
        parent=root,
        minvalue=0
    )


    if limite is None:

        messagebox.showwarning(
            "Cancelado",
            "No se indicó un límite de inasistencias."
        )

        root.destroy()

        return


    # ========================================================
    # PROCESAR
    # ========================================================

    try:

        procesar(
            archivo_notas,
            archivo_inas,
            limite
        )

    except Exception as e:

        messagebox.showerror(
            "Error",
            "Ocurrió un error durante "
            "el procesamiento:\n\n"

            f"{type(e).__name__}: {e}"
        )

        print()
        print(
            "ERROR:"
        )

        print(
            type(e).__name__,
            e
        )


    root.destroy()


# ============================================================
# EJECUTAR
# ============================================================

if __name__ == "__main__":
    main()
import os
import re
import copy
import queue
import threading
import tkinter as tk

from tkinter import (
    filedialog,
    messagebox,
    simpledialog
)

import openpyxl

from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment
)


# ============================================================
# CONFIGURACIÓN
# ============================================================

ETAPAS = [
    "Elegir archivos",
    "Validar archivos",
    "Validar contenido",
    "Calcular",
    "Comparar",
    "Creando reporte"
]


# ============================================================
# VENTANA DE PROGRESO
# ============================================================

class VentanaProgreso:

    def __init__(self, root):

        self.root = root

        self.ventana = tk.Toplevel(
            root
        )

        self.ventana.title(
            "Procesando archivos"
        )

        self.ventana.geometry(
            "560x430"
        )

        self.ventana.resizable(
            False,
            False
        )

        self.ventana.protocol(
            "WM_DELETE_WINDOW",
            self.intentar_cerrar
        )

        # ----------------------------------------------------
        # TÍTULO
        # ----------------------------------------------------

        titulo = tk.Label(
            self.ventana,
            text="PROCESANDO ARCHIVOS",
            font=("Arial", 16, "bold")
        )

        titulo.pack(
            pady=(25, 20)
        )


        # ----------------------------------------------------
        # ETAPAS
        # ----------------------------------------------------

        frame_etapas = tk.Frame(
            self.ventana
        )

        frame_etapas.pack(
            fill="x",
            padx=45
        )


        self.indicadores = []
        self.textos = []


        for etapa in ETAPAS:

            frame_etapa = tk.Frame(
                frame_etapas
            )

            frame_etapa.pack(
                fill="x",
                pady=3
            )


            indicador = tk.Label(
                frame_etapa,
                text="○",
                font=("Arial", 12),
                width=3,
                anchor="w"
            )

            indicador.pack(
                side="left"
            )


            texto = tk.Label(
                frame_etapa,
                text=etapa,
                font=("Arial", 11),
                anchor="w"
            )

            texto.pack(
                side="left"
            )


            self.indicadores.append(
                indicador
            )

            self.textos.append(
                texto
            )


        # ----------------------------------------------------
        # BARRA DE PROGRESO
        # ----------------------------------------------------

        self.barra_fondo = tk.Frame(
            self.ventana,
            bg="#E6E6E6",
            height=22
        )

        self.barra_fondo.pack(
            fill="x",
            padx=45,
            pady=(30, 8)
        )

        self.barra_fondo.pack_propagate(
            False
        )


        self.barra = tk.Frame(
            self.barra_fondo,
            bg="#4CAF50"
        )

        self.barra.place(
            x=0,
            y=0,
            width=0,
            relheight=1
        )


        # ----------------------------------------------------
        # PORCENTAJE
        # ----------------------------------------------------

        self.porcentaje = tk.Label(
            self.ventana,
            text="0%",
            font=("Arial", 11, "bold")
        )

        self.porcentaje.pack(
            pady=(0, 10)
        )


        # ----------------------------------------------------
        # ESTADO
        # ----------------------------------------------------

        self.estado = tk.Label(
            self.ventana,
            text="Esperando...",
            font=("Arial", 10),
            fg="#555555"
        )

        self.estado.pack(
            pady=5
        )


        # ----------------------------------------------------
        # CENTRAR
        # ----------------------------------------------------

        self.ventana.update_idletasks()

        ancho = self.ventana.winfo_width()
        alto = self.ventana.winfo_height()

        pantalla_ancho = (
            self.ventana.winfo_screenwidth()
        )

        pantalla_alto = (
            self.ventana.winfo_screenheight()
        )

        x = (
            pantalla_ancho - ancho
        ) // 2

        y = (
            pantalla_alto - alto
        ) // 2

        self.ventana.geometry(
            f"{ancho}x{alto}+{x}+{y}"
        )


    # ========================================================
    # ACTUALIZAR PROGRESO
    # ========================================================

    def actualizar(
        self,
        etapa,
        porcentaje,
        mensaje
    ):

        # ----------------------------------------------------
        # ETAPAS TERMINADAS
        # ----------------------------------------------------

        for i in range(
            etapa
        ):

            self.indicadores[i].config(
                text="✓",
                fg="#228B22"
            )

            self.textos[i].config(
                fg="#228B22",
                font=("Arial", 11)
            )


        # ----------------------------------------------------
        # ETAPA ACTUAL
        # ----------------------------------------------------

        if etapa < len(ETAPAS):

            self.indicadores[
                etapa
            ].config(
                text="●",
                fg="#0066CC"
            )

            self.textos[
                etapa
            ].config(
                fg="#0066CC",
                font=("Arial", 11, "bold")
            )


        # ----------------------------------------------------
        # ETAPAS PENDIENTES
        # ----------------------------------------------------

        for i in range(
            etapa + 1,
            len(ETAPAS)
        ):

            self.indicadores[i].config(
                text="○",
                fg="#777777"
            )

            self.textos[i].config(
                fg="#777777",
                font=("Arial", 11)
            )


        # ----------------------------------------------------
        # BARRA
        # ----------------------------------------------------

        porcentaje = max(
            0,
            min(
                100,
                porcentaje
            )
        )


        self.barra.place(
            relx=0,
            rely=0,
            relwidth=porcentaje / 100,
            relheight=1
        )


        self.porcentaje.config(
            text=f"{int(porcentaje)}%"
        )


        self.estado.config(
            text=mensaje,
            fg="#555555"
        )


    # ========================================================
    # FINALIZAR
    # ========================================================

    def finalizar(
        self,
        mensaje
    ):

        for i in range(
            len(ETAPAS)
        ):

            self.indicadores[i].config(
                text="✓",
                fg="#228B22"
            )

            self.textos[i].config(
                fg="#228B22",
                font=("Arial", 11)
            )


        self.barra.place(
            relwidth=1
        )


        self.porcentaje.config(
            text="100%"
        )


        self.estado.config(
            text=mensaje,
            fg="#228B22",
            font=("Arial", 10, "bold")
        )


    # ========================================================
    # INTENTAR CERRAR
    # ========================================================

    def intentar_cerrar(self):

        messagebox.showwarning(
            "Proceso en ejecución",
            "El proceso todavía está en ejecución.\n\n"
            "Esperá a que termine antes de cerrar.",
            parent=self.ventana
        )


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def seleccionar_archivo(titulo):

    ruta = filedialog.askopenfilename(
        title=titulo,
        filetypes=[
            ("Archivos Excel", "*.xlsx *.xlsm"),
            ("Todos los archivos", "*.*")
        ]
    )

    return ruta


def extraer_dni(valor):

    if valor is None:
        return None

    texto = str(valor)


    match = re.search(
        r"\bDNI\s*(\d{7,9})\b",
        texto,
        re.IGNORECASE
    )

    if match:
        return match.group(1)


    match = re.fullmatch(
        r"\s*(\d{7,9})\s*",
        texto
    )

    if match:
        return match.group(1)


    return None


def convertir_nota(valor):

    if valor is None or valor == "":
        return 0.0


    if isinstance(
        valor,
        (int, float)
    ):

        return float(valor)


    texto = str(
        valor
    ).strip().upper()


    if texto == "A":
        return 0.0


    try:

        return float(
            texto.replace(
                ",",
                "."
            )
        )

    except ValueError:

        return 0.0


def promedio(
    nota1,
    nota2
):

    n1 = convertir_nota(
        nota1
    )

    n2 = convertir_nota(
        nota2
    )

    return (
        n1 + n2
    ) / 2


def convertir_inasistencia(valor):

    if valor is None or valor == "":
        return 0.0


    try:

        return float(
            str(valor).replace(
                ",",
                "."
            )
        )

    except ValueError:

        return 0.0


def numero_limpio(valor):

    valor = float(
        valor
    )

    if valor.is_integer():
        return int(valor)

    return valor


def detectar_fila_inicio_inasistencias(ws):
    """
    Detecta automáticamente si la hoja de inasistencias
    tiene encabezado.

    Si la primera fila contiene un encabezado,
    comienza a leer desde la fila 2.

    Si la primera fila ya contiene datos de un alumno,
    comienza desde la fila 1.
    """

    valores = []

    for columna in range(
        1,
        ws.max_column + 1
    ):

        valor = ws.cell(
            1,
            columna
        ).value

        if valor is not None:

            valores.append(
                str(valor).strip().upper()
            )


    # --------------------------------------------------------
    # PALABRAS QUE INDICAN QUE LA FILA ES UN ENCABEZADO
    # --------------------------------------------------------

    encabezados = [
        "LEGAJO",
        "ALUMNO",
        "INASISTENCIAS",
        "INASISTENCIA",
        "FALTAS",
        "NOMBRE",
        "APELLIDO"
    ]


    for valor in valores:

        for encabezado in encabezados:

            if encabezado in valor:

                return 2


    # --------------------------------------------------------
    # SI NO HAY ENCABEZADO
    # --------------------------------------------------------

    return 1


# ============================================================
# PROCESAMIENTO PRINCIPAL
# ============================================================

def procesar(
    archivo_notas,
    archivo_inas,
    limite_inasistencias,
    cola
):

    # ========================================================
    # ETAPA 2 - VALIDAR ARCHIVOS
    # ========================================================

    cola.put(
        (
            "progreso",
            1,
            15,
            "Validando archivos..."
        )
    )


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

    ws_reporte = wb_inas[
        "REPORTE"
    ]

    ws_inas = wb_inas[
        "INAS"
    ]

    ws_inas2 = wb_inas[
        "INAS2"
    ]

    ws_notas = wb_notas[
        "Reporte"
    ]


    # ========================================================
    # ETAPA 3 - VALIDAR CONTENIDO
    # ========================================================

    cola.put(
        (
            "progreso",
            2,
            30,
            "Validando contenido..."
        )
    )


    # Por ahora no agregamos nuevas
    # validaciones.


    # ========================================================
    # ETAPA 4 - CALCULAR
    # ========================================================

    cola.put(
        (
            "progreso",
            3,
            35,
            "Cargando notas..."
        )
    )


    # ========================================================
    # CARGAR NOTAS
    # ========================================================

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
    # INASISTENCIAS - PRIMERA ETAPA
    # ========================================================

    inas1 = {}


    fila_inicio_inas1 = detectar_fila_inicio_inasistencias(
        ws_inas
    )


    for fila in range(
        fila_inicio_inas1,
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


        dni = str(
            legajo
        ).strip()


        inas1[dni] = convertir_inasistencia(
            faltas
        )


    # ========================================================
    # INASISTENCIAS - SEGUNDA ETAPA
    # ========================================================

    inas2 = {}


    fila_inicio_inas2 = detectar_fila_inicio_inasistencias(
        ws_inas2
    )


    for fila in range(
        fila_inicio_inas2,
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


        dni = str(
            legajo
        ).strip()


        inas2[dni] = convertir_inasistencia(
            faltas
        )


    # ========================================================
    # ELIMINAR HOJAS ANTERIORES
    # ========================================================

    for nombre in [
        "LIB",
        "REG 2ET"
    ]:

        if nombre in wb_inas.sheetnames:

            del wb_inas[
                nombre
            ]


    # ========================================================
    # LIMPIAR DESDE COLUMNA E
    # ========================================================

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
    # COPIAR ESTILO D -> E-I
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

    alumnos_procesados = 0
    alumnos_sin_notas = 0
    alumnos_libres = 0
    alumnos_excedidos = 0


    total_filas = max(
        1,
        ws_reporte.max_row - 10
    )


    for fila in range(
        11,
        ws_reporte.max_row + 1
    ):

        porcentaje = (
            35
            + (
                (fila - 10)
                / total_filas
            ) * 25
        )


        cola.put(
            (
                "progreso",
                3,
                porcentaje,
                "Calculando resultados..."
            )
        )


        # ----------------------------------------------------
        # DNI
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
        # FILA BLANCA
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


            ws_reporte.cell(
                fila,
                5
            ).value = numero_limpio(
                tp3
            )


            ws_reporte.cell(
                fila,
                6
            ).value = numero_limpio(
                tp4
            )


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
        # ====================================================

        faltas_1 = inas1.get(
            dni,
            0
        )


        faltas_2 = inas2.get(
            dni,
            0
        )


        diferencia = (
            faltas_2
            - faltas_1
        )


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
        # ALUMNO LIBRE
        # ====================================================

        if (
            prm2 is not None
            and float(prm2) < 4
        ):

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


                ws_reporte.cell(
                    fila,
                    columna
                ).font = copy.copy(
                    fuente_negra
                )


            alumnos_libres += 1


        # ====================================================
        # REGULAR EXCEDIDO
        # ====================================================

        elif (
            prm2 is not None
            and float(prm2) >= 4
            and float(diferencia)
            > limite_inasistencias
        ):

            ws_reporte.cell(
                fila,
                9
            ).value = (
                "ALUMNO REGULAR EXCEDIDO EN FALTAS"
            )


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
        # FORMATO
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
    # ETAPA 5 - COMPARAR
    # ========================================================

    cola.put(
        (
            "progreso",
            4,
            65,
            "Comparando resultados..."
        )
    )


    # ========================================================
    # FORMATO GENERAL
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
    # ANCHOS
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
    # ETAPA 6 - CREANDO REPORTE
    # ========================================================

    cola.put(
        (
            "progreso",
            5,
            78,
            "Creando reporte..."
        )
    )


    # ========================================================
    # CREAR LIB
    # ========================================================

    ws_lib = wb_inas.create_sheet(
        "LIB"
    )


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


    ws_lib["A3"] = "ALUMNO"


    ws_lib["A3"].font = Font(
        name="Arial",
        size=10,
        bold=True,
        color="000000"
    )


    # ========================================================
    # NOMBRES LIBRES
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
    # FORMATO LIB
    # ========================================================

    ws_lib.column_dimensions[
        "A"
    ].width = 45


    for fila in range(
        4,
        fila_lib
    ):

        ws_lib.row_dimensions[
            fila
        ].height = 36


        ws_lib.cell(
            fila,
            1
        ).alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )


    # ========================================================
    # ARCHIVO DE SALIDA
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

    cola.put(
        (
            "progreso",
            5,
            92,
            "Guardando archivo..."
        )
    )


    wb_inas.save(
        archivo_salida
    )


    # ========================================================
    # TERMINADO
    # ========================================================

    cola.put(
        (
            "terminado",
            {
                "alumnos_procesados":
                    alumnos_procesados,

                "alumnos_libres":
                    alumnos_libres,

                "alumnos_excedidos":
                    alumnos_excedidos,

                "alumnos_sin_notas":
                    alumnos_sin_notas,

                "limite":
                    limite_inasistencias,

                "archivo_salida":
                    archivo_salida
            }
        )
    )


# ============================================================
# INTERFAZ PRINCIPAL
# ============================================================

def main():

    root = tk.Tk()

    root.withdraw()


    # ========================================================
    # COLA DE COMUNICACIÓN
    # ========================================================

    cola = queue.Queue()


    # ========================================================
    # VENTANA DE PROGRESO
    # ========================================================

    progreso = VentanaProgreso(
        root
    )


    # ========================================================
    # SELECCIÓN DE ARCHIVOS
    # ========================================================

    progreso.actualizar(
        0,
        0,
        "Seleccioná el archivo de NOTAS..."
    )


    messagebox.showinfo(
        "Paso 1 de 3",
        "Primero seleccioná el archivo de NOTAS.",
        parent=progreso.ventana
    )


    archivo_notas = seleccionar_archivo(
        "Seleccionar archivo de NOTAS"
    )


    if not archivo_notas:

        progreso.ventana.destroy()
        root.destroy()

        return


    progreso.actualizar(
        0,
        5,
        "Seleccioná el archivo de INASISTENCIAS..."
    )


    messagebox.showinfo(
        "Paso 2 de 3",
        "Ahora seleccioná el archivo de INASISTENCIAS.",
        parent=progreso.ventana
    )


    archivo_inas = seleccionar_archivo(
        "Seleccionar archivo de INASISTENCIAS"
    )


    if not archivo_inas:

        progreso.ventana.destroy()
        root.destroy()

        return


    progreso.actualizar(
        0,
        10,
        "Indicá el límite de inasistencias..."
    )


    limite = simpledialog.askfloat(
        "Paso 3 de 3",
        "¿Cuál es el límite de inasistencias?",
        parent=progreso.ventana,
        minvalue=0
    )


    if limite is None:

        progreso.ventana.destroy()
        root.destroy()

        return


    # ========================================================
    # ELEGIR ARCHIVOS TERMINADO
    # ========================================================

    progreso.actualizar(
        0,
        15,
        "Archivos seleccionados. Iniciando procesamiento..."
    )


    # ========================================================
    # PROCESAMIENTO EN SEGUNDO PLANO
    # ========================================================

    def ejecutar():

        try:

            procesar(
                archivo_notas,
                archivo_inas,
                limite,
                cola
            )

        except Exception as e:

            cola.put(
                (
                    "error",
                    type(e).__name__,
                    str(e)
                )
            )


    hilo = threading.Thread(
        target=ejecutar,
        daemon=True
    )


    hilo.start()


    # ========================================================
    # REVISAR COLA
    # ========================================================

    def revisar_cola():

        try:

            while True:

                mensaje = cola.get_nowait()


                # --------------------------------------------
                # ACTUALIZAR PROGRESO
                # --------------------------------------------

                if mensaje[0] == "progreso":

                    _, etapa, porcentaje, texto = mensaje


                    progreso.actualizar(
                        etapa,
                        porcentaje,
                        texto
                    )


                # --------------------------------------------
                # TERMINADO
                # --------------------------------------------

                elif mensaje[0] == "terminado":

                    resultado = mensaje[1]


                    progreso.finalizar(
                        "Proceso terminado correctamente."
                    )


                    messagebox.showinfo(
                        "Proceso terminado",

                        "El archivo fue procesado correctamente.\n\n"

                        f"Alumnos procesados: "
                        f"{resultado['alumnos_procesados']}\n"

                        f"Alumnos libres: "
                        f"{resultado['alumnos_libres']}\n"

                        f"Regulares excedidos en faltas: "
                        f"{resultado['alumnos_excedidos']}\n"

                        f"Alumnos sin notas: "
                        f"{resultado['alumnos_sin_notas']}\n\n"

                        f"Límite de inasistencias: "
                        f"{resultado['limite']}\n\n"

                        "Archivo generado:\n"
                        f"{resultado['archivo_salida']}",

                        parent=progreso.ventana
                    )


                    progreso.ventana.destroy()
                    root.destroy()

                    return


                # --------------------------------------------
                # ERROR
                # --------------------------------------------

                elif mensaje[0] == "error":

                    _, tipo, detalle = mensaje


                    messagebox.showerror(
                        "Error",

                        "Ocurrió un error durante "
                        "el procesamiento:\n\n"

                        f"{tipo}: {detalle}",

                        parent=progreso.ventana
                    )


                    progreso.ventana.destroy()
                    root.destroy()

                    return


        except queue.Empty:

            pass


        root.after(
            100,
            revisar_cola
        )


    # ========================================================
    # INICIAR CONTROL DE COLA
    # ========================================================

    root.after(
        100,
        revisar_cola
    )


    # ========================================================
    # LOOP PRINCIPAL
    # ========================================================

    root.mainloop()


# ============================================================
# EJECUTAR
# ============================================================

if __name__ == "__main__":
    main()

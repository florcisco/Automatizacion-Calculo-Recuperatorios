# Automatización del Procesamiento de Cursadas con Python

## Descripción

Este proyecto consiste en una aplicación desarrollada en **Python**, para automatizar el procesamiento de planillas académicas exportadas desde un sistema de gestión.

El objetivo fue reemplazar un proceso manual repetitivo por una herramienta capaz de generar automáticamente los reportes necesarios, reduciendo tiempos de trabajo y minimizando errores humanos.

El programa fue diseñado a partir de necesidades reales detectadas durante su uso, y mejorado iterativamente incorporando nuevas funcionalidades.

---

## Características

- Lectura de archivos **Excel (.xlsx)** y **OpenDocument (.ods)**.
- Conversión automática de ausencias (**A**) a **0** para el cálculo de promedios.
- Cálculo automático del promedio (**PRM1**).
- Generación automática de las hojas:
  - REPORTE
  - LIB
  - REG 2ET
  - INAS
- Conservación de la estructura principal del archivo original.
- Resaltado automático de alumnos libres.
- Detección de alumnos con el mismo primer apellido.
- Detección de apellidos repetidos ignorando diferencias de tildes (por ejemplo, **García** y **Garcia**).
- Distribución como aplicación ejecutable para Windows mediante **PyInstaller**.

---

## Tecnologías utilizadas

- Python
- Pandas
- OpenPyXL
- ODFPy
- Tkinter

---

## Capturas

### Selección del archivo

![Programa](capturas/elegir_archivo.png)

---

### Archivo original

![Original](capturas/archivo_original.png)

---

### Ventana de Finalización

![Proceso Finalizado](capturas/ventana_finalizacion.png)

---

### Archivo procesado

![Procesado](capturas/archivo_procesado.png)

---

### Detección de apellidos repetidos

![Alerta](capturas/alerta_apellidos.png)

---

## Instalación

Clonar el repositorio:

```bash
git clone https://github.com/TU_USUARIO/procesador-cursadas-python.git
```

Instalar las dependencias:

```bash
pip install -r requirements.txt
```

Ejecutar:

```bash
python procesar_cursada.py
```

---

## Mejoras implementadas durante el desarrollo

Durante las pruebas se incorporaron distintas mejoras a partir del uso real de la aplicación:

- Conservación del encabezado y formato del archivo original.
- Compatibilidad con archivos **.xlsx** y **.ods**.
- Ajuste automático del ancho de columnas.
- Separación del DNI para mejorar la legibilidad.
- Creación automática de hojas auxiliares.
- Detección de alumnos con el mismo primer apellido.
- Normalización de tildes para evitar omitir casos como **García** / **Garcia**.

---

## Posibles mejoras futuras

- Procesamiento de múltiples archivos simultáneamente.
- Vista previa del reporte antes de guardarlo.
- Registro automático de actividades (logs).
- Validación adicional de datos.
- Compatibilidad completa con archivos Excel 97-2003 (.xls).

---

## Licencia

Este proyecto se distribuye bajo la licencia **MIT**.

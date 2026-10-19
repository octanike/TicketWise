# GINEFOBIA — Registro de entradas (versión de escritorio)

App de escritorio hecha con Python (Tkinter). Los datos se guardan en una
base de datos SQLite local llamada `ginefobia.db`, que se crea sola la
primera vez que abrís la app, en la misma carpeta donde está `ginefobia.py`.

## 1. Requisitos

- Python 3.8 o más nuevo.
- **Tkinter**: en Windows y Mac ya viene incluido con Python. En Linux a
  veces hay que instalarlo aparte:
  ```
  sudo apt install python3-tk
  ```

## 2. Instalar dependencias

Desde una terminal, parado en la carpeta donde están estos archivos:

```
pip install -r requirements.txt
```

(instala `python-docx` y `openpyxl`, que son las que se usan para generar
los archivos de Word y Excel).

## 3. Ejecutar

```
python ginefobia.py
```

Se abre la ventana de la app. La primera vez va a crear el archivo
`ginefobia.db` al lado del script — ahí es donde queda guardado todo, y
va a seguir estando la próxima vez que abras la app.

## Cómo se usa

- **Lobby**: título GINEFOBIA arriba, y dos opciones — *Nueva Entrada* y
  *Entradas*.
- **Nueva Entrada**: formulario con Nombre, Apellido, Cantidad, Precio,
  Método de pago y Fecha (día/mes/año). Al tocar **GUARDAR** queda
  registrada al toque en la base de datos.
- **Entradas**: lista todo lo cargado, con el total recaudado arriba.
  Abajo hay tres botones:
  - **Imprimir en Word (.docx)** — te deja elegir dónde guardar un Word
    con la tabla completa.
  - **Imprimir en Excel (.xlsx)** — igual, pero como planilla de Excel.
  - **Borrar todo** — pide confirmación antes de borrar (no se puede
    deshacer).

## Notas

- Si querés llevarte la base de datos a otra computadora, alcanza con
  copiar el archivo `ginefobia.db` junto con `ginefobia.py`.
- El "Método de pago" es una lista desplegable (Efectivo, Transferencia,
  Débito, Crédito, Mercado Pago). Se puede editar la lista `METODOS` al
  principio de `ginefobia.py` si querés otras opciones.

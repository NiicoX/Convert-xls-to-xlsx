# Convert-xls-to-xlsx

Aplicación de consola (estilo CMD) para convertir archivos Excel del formato antiguo **.xls** (Excel 97-2003) al formato moderno **.xlsx**.

## Requisitos

- Windows
- Python 3.8 o superior
- Microsoft Excel (opcional, pero da la conversión más fiel)

## Cómo ejecutar

Doble clic en `run.bat`, o desde CMD:

```bat
run.bat
```

La primera ejecución crea un entorno virtual e instala dependencias.

### Menú interactivo

1. Convertir un archivo `.xls`
2. Convertir todos los `.xls` de una carpeta
3. Convertir una carpeta incluyendo subcarpetas
4. Ayuda
0. Salir

Puedes pegar la ruta o arrastrar el archivo/carpeta a la ventana de consola.

### Uso directo (sin menú)

```bat
run.bat "C:\ruta\archivo.xls"
run.bat "C:\ruta\carpeta" -r
run.bat "C:\ruta\carpeta" -r -y -o "C:\salida"
```

| Opción | Descripción |
| --- | --- |
| `-r` / `--recursive` | Incluye subcarpetas |
| `-y` / `--overwrite` | Sobrescribe `.xlsx` existentes |
| `-o` / `--output` | Carpeta de destino |

## Cómo convierte

1. **Microsoft Excel (COM)** si está instalado: conserva fórmulas, gráficos y formato con mayor fidelidad.
2. **xlrd + openpyxl** si Excel no está disponible: copia hojas, valores, celdas combinadas y estilos básicos.

Los archivos `.xls` originales no se eliminan.

## Desarrollo

```bat
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m unittest discover -s tests
```

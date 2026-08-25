# Sales Operations Tool

> Personal portfolio project created by Da Ye.

Aplicación de escritorio en Python para automatizar tareas repetitivas basadas en Excel dentro de operaciones de ventas.

## Sobre el proyecto

Desarrollé esta herramienta para convertir procesos manuales y repetitivos de mi día a día en flujos más consistentes, rápidos y fáciles de revisar.

El proyecto refleja mi forma de trabajar: identificar tareas que consumen tiempo, entender las reglas necesarias, construir una solución práctica y mejorarla de forma iterativa.

Además de la interfaz de escritorio, la aplicación valida ficheros de entrada, aplica reglas de asignación y genera resultados en Excel para distintos procesos operativos.

## Tecnologías y habilidades aplicadas

- Python y desarrollo de aplicaciones de escritorio con Tkinter.
- Procesamiento y generación de archivos Excel con `openpyxl`.
- Consultas SQL ejecutables con DuckDB para documentar parte de la lógica de datos.
- Validación de datos y gestión de errores.
- Automatización de procesos operativos.
- Organización de código en módulos reutilizables.
- Control de versiones con Git y GitHub.

## Origen de la lógica

Parte de la lógica de validación y asignación se diseñó inicialmente mediante consultas SQL. Posteriormente la adapté a Python para crear una herramienta de escritorio autónoma, con validación de archivos Excel y generación de resultados.

Este proceso me permitió trasladar reglas orientadas a datos a una aplicación reutilizable y más accesible para usuarios no técnicos.

## Ejemplos SQL

El trabajo original incluyó varios scripts SQL orientados a distintos procesos operativos. Este repositorio contiene una selección reducida de ejemplos reconstruidos para fines de portfolio.

Los scripts publicados muestran técnicas de normalización, validación, CTEs, agregaciones y funciones de ventana, sin incluir datos operativos, rutas locales, nombres internos de tablas ni configuraciones confidenciales.

Consulta [sql/README.md](sql/README.md) para ejecutar los ejemplos con datos completamente ficticios.

## Uso de IA durante el desarrollo

Se utilizaron herramientas de IA como apoyo para explorar alternativas, revisar código, refactorizar y resolver problemas puntuales.

La definición del problema, las decisiones funcionales, la integración de los cambios y la validación final del comportamiento han sido realizadas por el autor. La IA se ha empleado como herramienta de apoyo dentro del proceso de desarrollo.

## Funcionalidades

La aplicación incluye varios procesos:

- Preallocation
- Vehicle Allocation
- Dealer Stock
- Vehicle Preallocation
- Check Free Cars
- Leads Analysis

Cada proceso solicita los ficheros necesarios, valida su disponibilidad y genera un archivo de resultado en la carpeta de trabajo seleccionada.

El proceso de Leads admite fuentes CSV o Excel de España y Portugal, consolida sus métricas por modelo y periodo, y permite configurar de forma validada la fórmula de una métrica derivada.

La lectura de libros de entrada también comprueba los encabezados esperados para localizar la hoja de datos cuando la pestaña activa no es la correcta, lo que hace la importación más resistente a plantillas con distinta organización.

## Requisitos

- Python 3 con Tkinter disponible.
- Las dependencias indicadas en `requirements.txt`.

Tkinter no se instala mediante `pip` porque forma parte de la instalación de Python. En Windows, el instalador oficial de Python suele incluir Tcl/Tk; puedes comprobarlo con:

```powershell
python -m tkinter
```

Si no se abre una ventana de prueba, modifica o reinstala Python incluyendo la opción **Tcl/Tk and IDLE**. En distribuciones Linux puede ser necesario instalar el paquete del sistema `python3-tk` antes de crear el entorno virtual.

## Instalación

Desde PowerShell, dentro de la carpeta del proyecto:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Ejecución

```powershell
python sales_operations_app.py
```

## Tests

Las pruebas automatizadas usan datos sintéticos y no requieren archivos operativos:

```powershell
python -m unittest discover -s tests
```

GitHub Actions ejecuta esta suite y comprueba que el código Python compile en cada cambio publicado.

Al abrir la aplicación:

1. Selecciona una carpeta de trabajo.
2. Elige el proceso que quieres ejecutar.
3. Revisa o selecciona los ficheros de entrada necesarios.
4. Ejecuta el proceso.
5. El archivo de resultado se guardará en la carpeta de trabajo.

## Datos y configuración local

Este repositorio contiene únicamente código preparado para fines de demostración. No incluye archivos operativos, datos de clientes, resultados, configuraciones locales ni información confidencial.

Consulta `.gitignore` para ver los archivos excluidos del repositorio.

## Estructura principal

- `sales_operations_app.py`: interfaz gráfica y orquestación de procesos.
- `asignaciones_excel.py`: motor de preasignación.
- `allocation_excel.py`: motor de asignación de vehículos.
- `dealer_stock_excel.py`: generación de dealer stock.
- `vehicle_preallocation_excel.py`: preasignación de vehículos.
- `check_free_cars_excel.py`: comprobación de vehículos libres.
- `leads_analysis_excel.py`: consolidación de leads y generación de informes por periodo.

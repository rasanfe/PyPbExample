# PyPbExample — PowerBuilder + PyPb → Python → Excel

Ejemplo autocontenido de cómo **llamar a Python desde PowerBuilder 2025 R2**
usando **[PyPb](https://github.com/Appeon/PyPb)** (el wrapper *beta* de Appeon
sobre Python.NET), con un **runtime de Python embebido y portable** para que el
usuario final **no tenga que instalar nada**.

El ejemplo carga un JSON (300 facturas) en un **DataWindow** y lo exporta a Excel
de dos formas, para responder a una pregunta: **¿generar el Excel con Python es
una mejora frente al export nativo de PowerBuilder?**

## Qué demuestra

| | **Nativo (PB)** `SaveDisplayedDataAs(…, XLSX!)` | **Python** (openpyxl) |
|---|---|---|
| `.xlsx` real | ✅ | ✅ |
| Cabeceras mostradas, formato moneda, totales | ✅ | ✅ |
| Respeta columnas ocultas / reordenadas | ✅ | ✅ |
| Colores / negrita / fuentes en celda | ❌ | ✅ |
| Autofiltro / cabecera inmovilizada | ❌ | ✅ |
| Líneas de código | 1 | un módulo `.py` + una fachada PB |

**Conclusión:** el nativo `SaveDisplayedDataAs(XLSX!)` ya te da una hoja Excel
**correcta** en una sola línea. Python compensa cuando necesitas **estilo visual**
(colores, totales destacados, filas alternas, filtros) — es decir, un informe
**presentable**.

## Cómo funciona

- **`n_cst_pyton`** — fachada fina en PowerScript sobre PyPb (init / import / run /
  invoke), con manejo de errores al estilo ERP (`0/-1` + `of_lasterror()`) y sin
  excepciones hacia fuera.
- **`n_cst_pyton_excel`** — objeto de negocio de ejemplo que usa openpyxl a través
  de la fachada.
- **`python.runtime/`** — un **Python 3.13 *embeddable*** con **openpyxl incluido**
  (`pip install … -t python.runtime`). Viaja junto a la aplicación; nada que instalar.
- **`bin.pypb.appeon/`** — los *assemblies* .NET de PyPb (se despliegan junto al EXE).
- **`export_facturas.py`** — exportador estilizado y *data-driven*: PB le pasa las
  **columnas visibles** del DataWindow para que el Excel refleje lo que se ve en
  pantalla.

## Funciones del ejemplo (ventana `w_main`)

| Botón | Qué hace |
|---|---|
| **Probar Python** | Arranca el runtime e imprime la versión de Python (`platform.python_version()`). |
| **Crear Excel (openpyxl)** | Crea un `.xlsx` básico con openpyxl vía `n_cst_pyton_excel`. |
| **Recargar facturas (JSON)** | Carga `data2026.json` en el DataWindow (`JSONParser`). |
| **Excel nativo (PB XLSX)** | `dw_1.SaveDisplayedDataAs(…, XLSX!)`. |
| **Excel Python (openpyxl)** | Exporta el grid a un `.xlsx` con estilos (cabecera color, filas zebra, moneda €, totales, autofiltro, cabecera fija). |

## Requisitos

- PowerBuilder **2022 R3 / 2025 / 2025 R2**
- Python **3.11–3.13**, misma arquitectura que la app (aquí 64-bit)
- .NET **8.0**

## Cómo ejecutarlo

1. Abre el *workspace* en el IDE de PowerBuilder.
2. Asegúrate de que **`pypblib.pbl`** está en la *library list* del *target*.
3. Compila y ejecuta. El grid se llena al abrir; usa los dos botones de export y
   compara `facturas_nativo.xlsx` con `facturas_python.xlsx`.

## Despliegue (sin instalar Python en el cliente)

Se entrega un **Python *embeddable* portable** dentro del proyecto:

1. Descarga el *Python Embeddable Package* (misma versión que en desarrollo).
2. Extráelo en `python.runtime/`.
3. Instala ahí las dependencias: `python.exe -m pip install openpyxl -t python.runtime`.
4. Apunta el runtime con una ruta relativa al EXE.
5. Distribuye `python.runtime/` **y** `bin.pypb.appeon/` junto al EXE.

> *Gotcha:* en el *embeddable*, descomenta `import site` en `pythonXXX._pth`.

## Notas y *gotchas* (PyPb es beta)

- **`of_executestatement` usa `eval()` de Python → solo EXPRESIONES.** Una asignación
  (`ws.title = x`) lanza `SyntaxError`; usa `setattr(...)` o métodos.
- **`of_set(prop, string)` pasa un `System.String` de .NET**, no un `str` de Python →
  rompe el código que valida con `re`/`isinstance(str)` (p. ej. openpyxl). Pasa el
  valor como argumento con nombre de un *invocation request*.
- **PyPb mantiene un único contexto de Python por proceso** y lo reutiliza: no mezcles
  runtimes distintos.
- **No existe `dw.ExportString(JSON!)`** en esta versión; el ejemplo pasa rutas y Python
  lee el fichero.

> PyPb es una característica **beta** de Appeon. Este repositorio es un ejemplo de
> aprendizaje, no código de producción.

---

Para estar al tanto de lo que publico puedes seguir mi blog:
<https://rsrsystem.blogspot.com/>

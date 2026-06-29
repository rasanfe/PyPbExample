# PyPbExample — PowerBuilder llamando a Python con PyPb 🐍

![PowerBuilder](https://img.shields.io/badge/PowerBuilder-2025%20R2-2D6FB3?style=flat-square&logo=appian&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.13%20embebido-3776AB?style=flat-square&logo=python&logoColor=white)
![PyPb](https://img.shields.io/badge/PyPb-beta%20(Appeon)-FFD43B?style=flat-square)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20con%20estilo-1D6F42?style=flat-square&logo=microsoftexcel&logoColor=white)
![Blog](https://img.shields.io/badge/blog-rsrsystem-FF5722?style=flat-square&logo=blogger&logoColor=white)

Esta vez me he puesto a trastear con **PyPb**, el *bridge* (todavía en **beta**) de
Appeon que deja **llamar a Python directamente desde PowerBuilder 2025 R2**, sin
montar APIs web, sin escribir una línea de C# y, lo mejor de todo, **sin que el
cliente tenga que instalar Python**. Os cuento cómo ha ido.

## 📋 ¿Qué es esto?

Un ejemplo **didáctico, en modo solución**, de cómo integrar Python en PowerBuilder
usando [PyPb](https://github.com/Appeon/PyPb) (el *bridge* de Appeon sobre
Python.NET). La idea central es una clase reutilizable, **`n_cst_pyton`**, que
convierte "hablar con Python" en algo tan simple como `of_import` + `of_invoke`:
importar módulos, llamar funciones, instanciar clases y recoger resultados.

Y para que no se quede en teoría, lo validamos con un caso ejecutable: **exportar un
DataWindow a Excel**, comparando dos caminos:

- El **export nativo** que estrena PB 2025: `dw_1.SaveDisplayedDataAs(…, XLSX!)`.
- El camino **Python + openpyxl**, que va un paso más allá y genera un `.xlsx`
  estilizado (cabecera con color, filas zebra, formato moneda €, totales, autofiltro
  y cabecera fija).

> Pero ojo: **el Excel es solo la excusa**. Lo importante es lo cómodo que resulta
> llamar a Python una vez tienes `n_cst_pyton` montada. A partir de ahí, integrar
> **pandas, numpy, openpyxl, OpenCV…** es cuestión de dos líneas.

## ✨ La pieza central: `n_cst_pyton`

Es una **fachada fina** en PowerScript sobre PyPb, con manejo de errores al estilo
ERP (`0/-1` + `of_lasterror()`) y **sin lanzar excepciones hacia fuera**. Fijaos en
lo poco que hace falta para arrancar Python y pedirle la versión:

```powerscript
n_cst_pyton lnv_py
string ls_ver

lnv_py = CREATE n_cst_pyton
lnv_py.of_init("…\python.runtime\python313.dll")        // arranca el runtime
lnv_py.of_invoke("platform", "python_version", ls_ver)  // import + llamada
// ls_ver -> "3.13.1"
DESTROY lnv_py
```

| Método | Para qué |
|---|---|
| `of_init(dll)` | Arranca (o reutiliza) el runtime de Python |
| `of_import(modulo, ref mod)` | Importa un módulo (con caché) |
| `of_invoke(modulo, func, ref res)` | Atajo: importa + llama función → string |
| `of_run(sentencia, ref res)` | Ejecuta una expresión Python suelta |
| `of_exec_req(sentencia, req, ref res)` | Ejecuta Python con variables locales |
| `of_lasterror()` | Último error capturado |

## 🎯 Qué incluye el ejemplo (ventana `w_main`)

| Botón | Qué hace |
|---|---|
| **Probar Python** | Arranca el runtime e imprime la versión (`platform.python_version()`). |
| **Crear Excel (openpyxl)** | Crea un `.xlsx` básico con openpyxl vía `n_cst_pyton_excel`. |
| **Recargar facturas (JSON)** | Carga `data2026.json` (300 facturas) en un DataWindow. |
| **Excel nativo (PB XLSX)** | La nueva función de PB 2025: `dw_1.SaveDisplayedDataAs(…, XLSX!)`. |
| **Excel Python (openpyxl)** | Exporta el grid a un `.xlsx` con estilo: cabecera color, filas zebra, moneda €, totales, autofiltro y cabecera fija. |

`n_cst_pyton_excel` es un objeto de negocio de ejemplo construido **encima** de
`n_cst_pyton`: enseña el patrón de **envolver una librería Python en una clase PB**
de uso cómodo.

## 🧠 Cómo funciona por dentro

- **`n_cst_pyton`** — la fachada reutilizable (lo importante del repo).
- **`n_cst_pyton_excel`** — ejemplo de objeto de negocio sobre openpyxl.
- **`python.runtime/`** — Python 3.13 *embeddable* con **openpyxl incluido**
  (`pip install … -t python.runtime`). Viaja junto a la app; nada que instalar.
- **`bin.pypb.appeon/`** — los *assemblies* .NET de PyPb (se despliegan con el EXE).
- **`export_facturas.py`** — exportador estilizado *data-driven*: PB le pasa las
  columnas **visibles** del DataWindow para que el Excel refleje lo que se ve.

## 🛠️ Requisitos

- PowerBuilder **2022 R3 / 2025 / 2025 R2**
- Python **3.11–3.13**, misma arquitectura que la app (aquí 64-bit)
- .NET **8.0**

## ▶️ Cómo probarlo

1. Abre el *workspace* (`PyPbExample.pbsln`) en el IDE de PowerBuilder.
2. Asegúrate de que **`pypblib.pbl`** está en la *library list* del *target*.
3. Compila y ejecuta.

### Despliegue sin instalar Python en el cliente

Se entrega un **Python *embeddable* portable** dentro del proyecto:

1. Descarga el *Python Embeddable Package* (misma versión que en desarrollo).
2. Extráelo en `python.runtime/`.
3. Instala ahí las dependencias: `python.exe -m pip install openpyxl -t python.runtime`.
4. Apunta el runtime con una ruta relativa al EXE.
5. Distribuye `python.runtime/` **y** `bin.pypb.appeon/` junto al EXE.

> *Gotcha:* en el *embeddable*, descomenta `import site` en `pythonXXX._pth`.

## ⚠️ Lo que he aprendido por el camino (PyPb es beta)

### PyPb y el export XLSX nativo se pelean por .NET

Al probar los dos exports me topé con algo llamativo:

- **Python → nativo:** funciona. Y a partir de ahí, en cualquier orden.
- **Nativo → Python:** el botón de Python falla con
  `Could not create instance of .NET Assembly: Load bin.pypb.appeon\…dll failed`.

La causa: **PowerBuilder carga el runtime .NET una sola vez por proceso.** El export
XLSX nativo (`SaveDisplayedDataAs`) y PyPb usan .NET, y **manda quien lo inicialice
primero**. Si el nativo va primero, PyPb ya no puede cargar su *assembly*.

Mi **primer intento** fue restaurar el *directorio de trabajo* tras el export (PyPb
carga su DLL por ruta relativa) — pero **no era eso**: no era el directorio, era el
**orden de inicialización**.

La **solución**: **pre-cargar PyPb al arrancar la app** (en el `open` del objeto
aplicación, antes de abrir la ventana), para que su .NET se inicialice el primero:

```powerscript
n_cst_pyton lnv_pyinit
lnv_pyinit = CREATE n_cst_pyton
lnv_pyinit.of_init(gs_appdir + "\python.runtime\python313.dll")
DESTROY lnv_pyinit
```

El contexto de Python es global y persiste tras el `DESTROY`, así que después los
botones funcionan en cualquier orden.

### Otros *gotchas* anotados

- **`of_executestatement` usa `eval()` de Python → solo EXPRESIONES.** Una asignación
  (`ws.title = x`) lanza `SyntaxError`; usa `setattr(...)` o métodos (`of_exec_req`).
- **`of_set(prop, string)` pasa un `System.String` de .NET**, no un `str` de Python →
  rompe el código que valida tipos con `re`/`isinstance(str)` (p. ej. openpyxl).
- **PyPb mantiene un único contexto de Python por proceso** y lo reutiliza.
- Para pasar los datos del grid a Python se usa **`dw_1.ExportJson()`** (en esta
  versión `ExportString(JSON!)` no existe), así el Excel sale **fiel al DataWindow**.

> PyPb es una característica **beta** de Appeon. Este repositorio es un ejemplo de
> aprendizaje, no código de producción.

## 🔗 Repo PowerBuilder

Tenéis el ejemplo publicado **en modo solución** (clonáis y compila):
<https://github.com/rasanfe/PyPbExample>

## 🙌 Créditos

- **PyPb** es un proyecto de **Appeon**: <https://github.com/Appeon/PyPb>
- Excel estilizado con **openpyxl**.

---

> ¡Nos vemos en el próximo artículo! Y recuerda: en PowerBuilder, los límites solo están en nuestra imaginación. 🚀

📨 **Blog:** <https://rsrsystem.blogspot.com/>

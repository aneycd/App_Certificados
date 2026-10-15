# AGENTS.md

## Proyecto
- Academia Horizonte: automatizar emision de certificados por cohorte (hoy manual en Excel/Word) como aplicacion web.

## Arquitectura (3 piezas)
- Bodega: `INSUMOS/*.xlsx`. Solo lectura, nunca modificar.
- Cocina: backend Python + Flask en `app.py` (aun no existe). Toda la logica va aqui, NUNCA en la interfaz.
- Salon: HTML servido por Flask. Solo muestra; no calcula ni decide.

## Insumos
- `INSUMOS/Maestro_Estudiantes.xlsx` → hoja `Estudiantes` (`A1:E25`): `Identificacion | Nombre_Completo | Correo | Programa | Cohorte`.
- `INSUMOS/Registro_Evaluaciones.xlsx` → hoja `Evaluaciones`: `Identificacion | Programa | Modulo | Nota | Asistencia_Pct | Fecha_Cierre`.
- `Programa`: `Tecnico en IA Aplicada`, `Excel Avanzado para Negocios`. `Cohorte` observada: `2026-A`. `Modulo 1..4`, cierres mar–jun 2026.
- Clave de agrupacion: `(Identificacion, Programa)`. Un certificado por estudiante y programa.
- Dato conocido: `Registro_Evaluaciones` trae `Identificacion=999880777` sin match en maestro — caso huerfano, no borrar en silencio ni romper.
- Correos `@ejemplo.cr` son fixtures; tildes salen mojibake en XML crudo — leer con libreria Excel, no con parseo XML manual.

## Reglas de negocio
- `Promedio = suma(Nota) / cantidad de modulos cursados`. `Asistencia = promedio(Asistencia_Pct)`. `Nota` 0–100.
- `Aprobacion`: `Promedio >= 70` y `Asistencia >= 80` (limites inclusivos).
- `Participacion`: `Promedio < 70` y `Asistencia >= 80`.
- Sin certificado si `Asistencia < 80`.

## Decisiones técnicas (ver PLAN.md)
- Excel: `openpyxl` (lee `.xlsx`, liviana). Descartadas: `pandas` (pesada, esconde lógica), `xlrd` (ya no lee `.xlsx`).
- Web: `Flask` + Jinja (viene con Flask). Descartados: `FastAPI` (sobra), `Django` (gigante).
- Estructura: `app.py` + `templates/base.html,index.html,estudiante.html,certificado.html` + `static/style.css`. Rutas: `/`, `/estudiante/<id>/<programa>`, `/certificado/<id>/<programa>`.
- IA promedia 4 módulos, Excel 3 (divisor = módulos cursados, no fijo).

## Estilo de trabajo
- Codigo simple y comentado en espanol, marcando ladrillos: variables, tipos, condicionales, bucles y funciones.
- Diseno web azul marino y dorado.
- Explicar en espanol sencillo lo que se haga.

## Entorno
- Shell Windows PowerShell; la ruta del repo tiene espacios — entrecomillar paths; `&&`/`||` no valen, usar `;`.
- Toolchain verificado: Python 3.12 en `%LOCALAPPDATA%\Programs\Python\Python312\python.exe` + `pip install -r requirements.txt` (`Flask`, `openpyxl`). Arranca con `python app.py` o doble clic en `iniciar_app.bat` → http://localhost:5000 (o http://127.0.0.1:5000).
- `.gitignore` cubre `.env`, `.venv`, `__pycache__`, `*.pyc`, `.vercel`. No commitear certificados/salidas generadas salvo pedido.

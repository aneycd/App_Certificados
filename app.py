# app.py — LA COCINA (backend con Flask).
# Toda la lógica va aquí. La página solo muestra lo que aquí se calcula.
# Para arrancar con doble clic usa iniciar_app.bat. Para detener: Ctrl+C.

from collections import defaultdict  # LADRILLO: caja para agrupar por clave
from pathlib import Path  # LADRILLO: rutas que funcionan desde cualquier carpeta
from flask import Flask, render_template, request  # LADRILLO: Flask = servidor web
from openpyxl import load_workbook  # LADRILLO: librería para leer Excel .xlsx

# LADRILLO variable: creamos la aplicación web.
app = Flask(__name__)

# LADRILLO variable: carpeta donde vive app.py (no depende de dónde se ejecuta).
BASE_DIR = Path(__file__).resolve().parent
RUTA_MAESTRO = BASE_DIR / "INSUMOS" / "Maestro_Estudiantes.xlsx"
RUTA_EVAL = BASE_DIR / "INSUMOS" / "Registro_Evaluaciones.xlsx"


# LADRILLO función: calcula un promedio (suma / cantidad).
def calcular_promedio(numeros):
    # LADRILLO variables y tipos: numeros es lista de float, devuelve float.
    # LADRILLO condicional: si no hay datos, evita dividir entre cero.
    if not numeros:
        return 0.0
    return sum(numeros) / len(numeros)


# LADRILLO función: decide el certificado según AGENTS.md (límites inclusivos).
def decidir_certificado(promedio, asistencia):
    # LADRILLO variables y tipos: promedio (float), asistencia (float) -> texto.
    # LADRILLO condicionales: el orden importa (primero la asistencia).
    if asistencia < 80:
        return "Sin certificado"
    if promedio >= 70:
        return "Aprobación"
    return "Participación"


def cargar_maestro(ruta):
    """Lee el Excel oficial. Devuelve dict con clave (identificacion, programa)."""
    wb = load_workbook(ruta, data_only=True, read_only=True)
    ws = wb["Estudiantes"]
    maestro = {}
    # LADRILLO bucle: recorre cada fila del Excel (salta el encabezado).
    for fila in list(ws.iter_rows(values_only=True))[1:]:
        ident, nombre, correo, programa, cohorte = fila
        maestro[(str(ident), programa)] = {
            "nombre": nombre, "correo": correo, "cohorte": cohorte,
        }
    return maestro


def cargar_evaluaciones(ruta):
    """Lee notas. Devuelve notas y asistencias agrupadas por (id, programa)."""
    wb = load_workbook(ruta, data_only=True, read_only=True)
    ws = wb["Evaluaciones"]
    notas = defaultdict(list)
    asist = defaultdict(list)
    # LADRILLO bucle: junta las notas de cada estudiante y programa.
    for fila in list(ws.iter_rows(values_only=True))[1:]:
        ident, programa, _modulo, nota, asistencia, _cierre = fila
        clave = (str(ident), programa)
        notas[clave].append(float(nota))
        asist[clave].append(float(asistencia))
    return notas, asist


def calcular_resultados():
    """Cruza maestro + evaluaciones y prepara todo para la página."""
    maestro = cargar_maestro(RUTA_MAESTRO)
    notas, asist = cargar_evaluaciones(RUTA_EVAL)

    resultados = []
    # LADRILLO bucle: solo los del maestro generan certificado (el huérfano no).
    for clave, datos in maestro.items():
        if clave not in notas:
            continue  # LADRILLO condicional: sin notas -> va a inconsistencias
        promedio = calcular_promedio(notas[clave])  # LADRILLO función
        asistencia = calcular_promedio(asist[clave])
        tipo = decidir_certificado(promedio, asistencia)  # LADRILLO función
        resultados.append({
            "id": clave[0], "programa": clave[1], "nombre": datos["nombre"],
            "modulos": len(notas[clave]),
            "promedio": round(promedio, 2), "asistencia": round(asistencia, 2),
            "tipo": tipo,
        })
    # Orden: programa y luego nombre.
    resultados.sort(key=lambda r: (r["programa"], r["nombre"]))

    # LADRILLO condicionales para los contadores de arriba.
    n_aprob = sum(1 for r in resultados if r["tipo"] == "Aprobación")
    n_part = sum(1 for r in resultados if r["tipo"] == "Participación")
    n_sin = sum(1 for r in resultados if r["tipo"] == "Sin certificado")

    # Las 3 inconsistencias detectadas.
    sin_eval = [  # maestro SIN ninguna evaluación
        {"id": c[0], "programa": c[1], "nombre": d["nombre"]}
        for c, d in maestro.items() if c not in notas
    ]
    huerfanos = [  # evaluación SIN maestro (no genera certificado)
        {"id": c[0], "programa": c[1]}
        for c in notas if c not in maestro
    ]
    # LADRILLO bucle + condicional: detecta ceros (posible abandono).
    con_ceros = [
        {"id": c[0], "programa": c[1], "nombre": maestro[c]["nombre"]}
        for c in notas if c in maestro
        and (0 in notas[c] or 0 in asist[c])
    ]

    programas = sorted({r["programa"] for r in resultados})
    return resultados, n_aprob, n_part, n_sin, sin_eval, huerfanos, con_ceros, programas


# LADRILLO función con decorador: atiende la página principal.
@app.route("/")
def index():
    resultados, n_aprob, n_part, n_sin, sin_eval, huerfanos, con_ceros, programas = \
        calcular_resultados()
    # Filtro por programa (?programa=...). La página no decide, solo filtra vista.
    filtro = request.args.get("programa", "Todos")
    if filtro != "Todos":  # LADRILLO condicional
        resultados = [r for r in resultados if r["programa"] == filtro]
    return render_template(
        "index.html", resultados=resultados,
        n_aprob=n_aprob, n_part=n_part, n_sin=n_sin,
        sin_eval=sin_eval, huerfanos=huerfanos, con_ceros=con_ceros,
        programas=programas, filtro=filtro,
    )


if __name__ == "__main__":
    # LADRILLO: arranca el servidor en http://localhost:5000
    app.run(debug=True)

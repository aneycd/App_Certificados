# PLAN.md — App Certificados (Academia Horizonte)

> Modo Plan: este archivo manda. No escribir código hasta validar cada etapa.

## Orden de construcción

1. Etapa 0 — Preparar la mesa (entorno).
2. Etapa 1 — Validar el cruce de datos (sin web todavía).
3. Etapa 2 — Cocina: backend Flask en `app.py`.
4. Etapa 3 — Salón: interfaz HTML azul marino y dorado.
5. Etapa 4 — Repaso final y cierre de cohorte.

## Etapa 0 — Preparar la mesa

**Qué haré:**
- Verificar `python --version` y `pip --version` en esta máquina.
- Crear `.venv`, `requirements.txt` (`Flask`, `openpyxl`) y `app.py` vacío.
- Probar que Flask arranca con una página "hola" temporal (se quita después).

**Cómo la validamos:**
- Me muestras el resultado de `pip install -r requirements.txt` sin errores.
- `python app.py` abre la página de prueba en el navegador.
- Solo sigues si ves la página de prueba con tus ojos.

## Etapa 1 — Validar el cruce de datos (lo más importante)

**Qué haré:**
- Función para leer los dos Excel con `openpyxl` (sin modificar `INSUMOS/`).
- Agrupar por `(Identificacion, Programa)`, calcular `Promedio` y `Asistencia` con las reglas de `AGENTS.md`.
- Informe en pantalla con 4 listas: Aprobación, Participación, Sin certificado, Avisos (huérfano + sin notas).

**Números esperados (ya calculados del análisis):**
- 22 certificados: 19 Aprobación + 3 Participación.
- 1 Sin certificado: `111230984` (44.5 / 36.25).
- 1 pendiente sin notas: `304560321`.
- 1 huérfano a reportar: `999880777` (no genera certificado).

**Cómo la validamos:**
- Comparamos 3 casos a mano: `101230456` → 91.25/96.25 Aprobación; `307650129` → 59.67/85.0 Participación; `111230984` → Sin certificado.
- Revisas que el huérfano salga como aviso y no como certificado.
- No pasamos a la etapa 2 hasta que los 4 números cuadren.

## Etapa 2 — Cocina: backend Flask

**Qué haré:**
- `app.py` con funciones: `cargar_maestro()`, `cargar_evaluaciones()`, `calcular_resultados()`, `clasificar(promedio, asistencia)`.
- Rutas: `/` (resumen), `/estudiante/<id>/<programa>` (detalle), `/certificado/<id>/<programa>` (solo si aplica).
- Toda decisión en Python; las plantillas solo muestran datos.

**Cómo la validamos:**
- Abres `/` y cuentas 22 filas con sus promedios.
- Pruebas un caso de cada tipo + un huérfano (debe dar error amable, no caerse).
- Confirmas que IA divide entre 4 y Excel entre 3.

## Etapa 3 — Salón: interfaz

**Qué haré:**
- `templates/base.html`, `index.html`, `estudiante.html`, `certificado.html` + `static/style.css`.
- Diseño azul marino y dorado, botones grandes, tabla clara, certificado imprimible.

**Cómo la validamos:**
- Navegas las 3 páginas sin errores y el certificado se ve bien al imprimir en PDF.
- Verificas que la página no calcula nada (ver código fuente: solo texto, sin fórmulas).

## Etapa 4 — Repaso final

**Qué haré:**
- Repaso de comentarios en español (ladrillos: variables, tipos, condicionales, bucles, funciones).
- Prueba final con los Excel originales intactos (`git status` limpio en `INSUMOS/`).

**Cómo la validamos:**
- Lista de chequeo: 22 certificados, colores correctos, explicación en español sencillo.
- Tú das el visto bueno para emitir.

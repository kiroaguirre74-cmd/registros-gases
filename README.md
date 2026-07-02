# Registros de Gases — Generador semanal

Repositorio con registros de un detector de gases multigás (H2S, CO, O2, LEL)
en formato **semanal** (un libro de Excel por semana), más las herramientas
para volver a generarlos para cualquier año y dispositivo.

## Contenido

```
Dispositivo_MA215-012525/   Registros del equipo MA215-012525 (hoja FMexport54977)
  2026/ 2027/ 2028/ 2029/   Un .xlsx por semana + un .zip por año
Dispositivo_BH-4S-181334/   Registros del equipo BH-4S-181334 (hoja CMexport54977)
  2026/ 2027/ 2028/ 2029/   Un .xlsx por semana + un .zip por año

plantilla/                  Archivo .xlsx base (una semana real de ejemplo)
generar_registros.py        Script generador (solo biblioteca estándar de Python)
PROMPT_GENERADOR.txt        Prompt exacto y reutilizable para usar en cualquier IA
```

## Cómo generar más años / dispositivos

Requiere Python 3 (sin dependencias externas). Desde la raíz del repo:

```bash
# Dispositivo BH-4S-181334, hoja CM, año 2030, filas variables 15-60
SERIAL="BH-4S-181334" SHEET_NAME="CMexport54977" \
  OUTDIR="Dispositivo_BH-4S-181334/2030" \
  VARY_ROWS=1 ROW_MIN=15 ROW_MAX=60 YEAR=2030 \
  python3 generar_registros.py

# Dispositivo MA215-012525, hoja FM, año 2030
SERIAL="MA215-012525" SHEET_NAME="FMexport54977" \
  OUTDIR="Dispositivo_MA215-012525/2030" \
  VARY_ROWS=1 ROW_MIN=15 ROW_MAX=60 YEAR=2030 \
  python3 generar_registros.py
```

### Parámetros (variables de entorno)

| Variable | Descripción | Ejemplo |
|----------|-------------|---------|
| `SERIAL` | Número de serie (columna A) | `BH-4S-181334` |
| `SHEET_NAME` | Nombre de la hoja de Excel | `CMexport54977` |
| `YEAR` | Año a generar | `2030` |
| `MES_INI` / `MES_FIN` | Rango de meses (1–12) | `1` / `12` |
| `VARY_ROWS` | `1` filas variables, `0` fijas | `1` |
| `ROW_MIN` / `ROW_MAX` | Rango de filas si `VARY_ROWS=1` | `15` / `60` |
| `OUTDIR` | Carpeta de salida | `Dispositivo_.../2030` |
| `TEMPLATE` | Ruta de la plantilla | `plantilla/ENE_...xlsx` |

## Reglas aplicadas

- **Semanas** de lunes a domingo, recortadas al mes; las semanas parciales
  de 1–2 días se fusionan con la semana vecina del mismo mes.
- **Fechas** en orden cronológico y siempre dentro de la semana del archivo.
- **Lecturas de gas** aleatorias y realistas (H2S normalmente 0.0 y a veces
  en alarma ≥10.0, O2 ~20.9, CO/LEL 0.0).
- Se conservan estructura (55 columnas), formato y configuración del equipo.

## Usar en otra IA

Abre **`PROMPT_GENERADOR.txt`**, copia el bloque de prompt y pégalo en
cualquier IA (ChatGPT, Claude, Gemini, etc.), cambiando los valores entre
`<...>`. También puedes entregarle a esa IA el `generar_registros.py` y la
carpeta `plantilla/` para que lo ejecute directamente.

# Reporte Gerencial Unifrutti

Aplicación Flask para el panel gerencial del Fundo Unifrutti. La portada enlaza los reportes complementarios y las cinco áreas de gestión.

## Estructura

```text
data/
  abastecimiento/
    BD_ EMBALSE PALO VERDE.xlsx
    BD_ OFERTAS.xlsx
  monitoreo/
  electromecanica/
  proyectos/
  infraestructura/
templates/
  index.html
  reports/              Un template por reporte
static/                 Logo e imágenes del fundo
app.py                  Rutas Flask, lectura Excel y API
```

## Abastecimiento

El reporte combina dos matrices y sincroniza ambas con los filtros de año y mes:

1. **WEEKLY OPERATION REPORT H2OLMOS** — lee `H2OMOS` y presenta las lecturas diarias directas de `Flow transferred`, `Inflow` y `Reservoir capacity (Hm3)`. Al ingresar, selecciona el mes y la semana actuales y conserva las siete fechas, incluso si alguna lectura aún no está disponible. `Capacity %` se calcula como la capacidad del embalse dividida entre la capacidad de reserva definida para el `YEAR`.
2. **WATER LICENSE USAGE** — lee `EXE lice` de `BD_ OFERTAS.xlsx`, sin responder a los filtros del reporte principal. Agrupa solo por Mes y Licencia (m³/día), usa cada Fundo como columna y presenta el promedio de `Ejecucion de licencia (%)` sobre todos los años disponibles.

La lectura de los libros Excel queda en caché por versión del archivo y se actualiza automáticamente si cambia un libro. `Flow transferred` muestra banderas según los rangos de consumo configurados en la vista.

- Página: <http://127.0.0.1:5000/abastecimiento>
- API JSON: <http://127.0.0.1:5000/api/abastecimiento>
- Filtro anual: `/abastecimiento?anio=2026`

Las futuras fuentes se guardan en `data/<area>/` y sus reportes en `templates/reports/`.

## Ejecutar en Windows

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

Abre <http://127.0.0.1:5000>. El panel usa el azul `#1A428A` y el rojo `#c1121f`.




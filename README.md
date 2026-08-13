# Efectividad Mecanicos

Dashboard web que replica el tablero de Power BI "Eficiencia Operativa".
Los Exceles fuente se commitean al repo. Una GitHub Action los procesa y
genera JSONs. Vercel sirve la UI + los JSONs como sitio estatico.

## Estructura

```
Efectividad Mecanicos/
+- .github/workflows/build.yml   <- Regenera app/data/ cuando se suben Excels
+- vercel.json                   <- Config de deploy (output: app/)
+- app/
   +- excel/
   |  +- efectividad.xlsx        <- Fuente: pegar aca la version actualizada
   |  +- ordenes.xlsx            <- Fuente: pegar aca la version actualizada
   +- data/
   |  +- tareas.json             <- Generado, NO editar a mano
   |  +- ordenes.json            <- Generado, NO editar a mano
   |  +- metricas.json           <- Generado, NO editar a mano
   +- build_data.py              <- Script que lee excels y genera JSONs
   +- index.html                 <- UI (filtros, tabla, export CSV)
   +- README.md
```

## Flujo de actualizacion de datos

1. Reemplazar `app/excel/efectividad.xlsx` y/o `app/excel/ordenes.xlsx` con la
   version nueva (mismos nombres).
2. Hacer commit + push a `main`.
3. La GitHub Action `.github/workflows/build.yml` corre automaticamente:
   - `python app/build_data.py` (lee los Exceles y genera los 3 JSONs).
   - Commit + push de `app/data/` de vuelta al repo.
4. Vercel detecta el push y redespliega (< 30 s).
5. La web ya muestra los datos nuevos.

Si queres, podes disparar la Action manualmente desde la pestana *Actions* del
repo de GitHub con `workflow_dispatch`. No requiere secrets.

## Como correrlo en local (sin Vercel)

Para desarrollo local rapido:

```
pip install openpyxl
python app/build_data.py            # genera app/data/*.json
python -m http.server -d app 8000   # sirve app/ en localhost:8000
```

Abrir `http://localhost:8000`.

## Como deployar en Vercel

1. Importar el repo en https://vercel.com/new. Sin settings extra:
   - Framework preset: **Other**
   - Output Directory: ya viene de `vercel.json` (`app`).
2. Cada push a `main` redespliega.
3. El `vercel.json` ya configura `cleanUrls` y cache de 5 min para `/data/*`.

## Endpoint del sitio

| URL              | Descripcion                                          |
|------------------|------------------------------------------------------|
| `GET /`          | UI principal (`index.html`)                          |
| `GET /data/tareas.json`   | Filas de tareas (regenerado por CI)           |
| `GET /data/ordenes.json`  | Mapa DocNum -> Tipo de orden                  |
| `GET /data/metricas.json` | Metricas DAX replicadas, con key `__total__`  |

## Calculo (replica DAX)

| Metrica                  | Formula                                                                                              |
|--------------------------|------------------------------------------------------------------------------------------------------|
| Hs Totales               | `SUMMARIZE` unico por (DocNum, Fecha, Codigo Tarea, Codigo Empleado) MAX(Std) / DISTINCTCOUNT(Emp en mismo DocNum+Fecha+Tarea) |
| Hs Mecanico              | `SUM(Horas Trabajadas) / 60`                                                                        |
| % Horas Mes              | `Hs Mecanico / 156`                                                                                  |
| Margen Productividad     | `(SUM(Std) WHERE dif>=0 / SUM(Trab) WHERE dif>=0 - 1) * 100`                                        |
| % Procutivo              | `COUNT(fila dif>=0) / COUNT(total filas) * 100`                                                     |
| Supera las 80Hs          | `IF(HsTot >= 80, "SI", "NO")`                                                                        |
| Supera %50 Optim.        | `IF(%Procutivo >= 50, "SI", "NO")`                                                                   |

Las formulas replican las DAX del modelo PBI y fueron validadas contra el
tablero original.

## Filtros disponibles (multi-select)

- Empleado (con busqueda)
- Mes y Anio
- Tipo de orden (cruzado via DocNum con `ordenes.xlsx`)
- Tareas con tiempo estandar menor a 1 minuto

## Exportar

El boton **Exportar CSV** descarga las filas filtradas, separadas por `;`.

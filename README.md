# CENTRO DE COSTOS - MR. PEANUTT

Centro de costos de Mr. Peanutt (mantequillas de maní y almendra): productos, canal directo,
tiendas (B2B), restaurantes (baldes), punto de equilibrio, promociones "Arma tu pack" y hallazgos.

El dashboard es un **solo archivo** (`dashboard/index.html`) con toda la data embebida, listo para
desplegar como sitio estático en Vercel.

## Deploy en Vercel

1. En Vercel: **Add New → Project → Import** este repositorio.
2. Framework Preset: **Other**. No hace falta Build Command ni Install Command.
3. `vercel.json` ya define `outputDirectory: dashboard`, así que solo se publica el dashboard
   (la BD, el Excel y los scripts no quedan expuestos).
4. Deploy. Cada `git push` a `main` vuelve a publicar.

## Estructura

```
DATOS INICIALES/            Excel original de Piero (se toma el primer .xlsx)
scripts/
  etl_build_db.py           1) Excel -> BD/mrpeanutt_costos.db (todas las constantes de negocio al inicio)
  export_dashboard_json.py  2) BD -> dashboard/data.json
  build_dashboard.py        3) data.json + dashboard_template.html -> dashboard/index.html
  export_excel.py           4) BD -> CENTRO_DE_COSTOS_MRPEANUTT.xlsx (una hoja por tabla)
  dashboard_template.html   Plantilla del dashboard (marcador /*__DATA__*/)
BD/mrpeanutt_costos.db      Base de datos SQLite generada
dashboard/                  Sitio estático que publica Vercel
CENTRO_DE_COSTOS_MRPEANUTT.xlsx  Exportación Excel generada
```

## Regenerar el dashboard

Cualquier cambio de precio, margen, tramo o costo fijo se hace en las constantes al inicio de
`scripts/etl_build_db.py` y se corre la cadena completa (no editar `index.html` ni el Excel a mano):

```bash
pip install -r requirements.txt
set PYTHONIOENCODING=utf-8
python scripts/etl_build_db.py
python scripts/export_dashboard_json.py
python scripts/build_dashboard.py
python scripts/export_excel.py
```

Luego `git commit` + `git push` y Vercel publica la nueva versión.

## Vista local

```bash
python -m http.server 8766 --directory dashboard
```

y abrir <http://localhost:8766>.

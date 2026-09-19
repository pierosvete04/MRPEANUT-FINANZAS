# -*- coding: utf-8 -*-
"""
Exporta la BD SQLite (BD/mrpeanutt_costos.db) al JSON que consume el dashboard
(dashboard/data.json). Correr después de etl_build_db.py:

    python etl_build_db.py
    python export_dashboard_json.py
    python build_dashboard.py
    python export_excel.py
"""
import json
import sqlite3
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DB_PATH = BASE / "BD" / "mrpeanutt_costos.db"
OUT_PATH = BASE / "dashboard" / "data.json"

TABLAS = [
    ("parametros", "id"),
    ("productos_costeo", "id"),
    ("costeo_balde_detalle", "id"),
    ("canal_directo", "id"),
    ("promociones", "id"),
    ("promociones_propuesta", "id"),
    ("promociones_ejemplos", "id"),
    ("promociones_terminos", "orden"),
    ("packs_escenarios", "id"),
    ("escala_b2b_excel", "id"),
    ("escala_b2b_propuesta", "id"),
    ("cadena_valor", "id"),
    ("restaurantes_porcion", "id"),
    ("resumen_margenes", "id"),
    ("costos_fijos", "id"),
    ("punto_equilibrio", "id"),
    ("issues", "CASE severidad WHEN 'alta' THEN 0 WHEN 'media' THEN 1 ELSE 2 END, id"),
    ("recomendaciones", "prioridad"),
]


def rows(cx, sql):
    cx.row_factory = sqlite3.Row
    return [dict(r) for r in cx.execute(sql).fetchall()]


def main():
    cx = sqlite3.connect(DB_PATH)
    data = {t: rows(cx, f"SELECT * FROM {t} ORDER BY {orden}") for t, orden in TABLAS}
    data["meta"] = dict(cx.execute("SELECT clave, valor FROM meta").fetchall())

    directo = data["canal_directo"]
    prop = data["escala_b2b_propuesta"]
    excel = data["escala_b2b_excel"]
    issues = data["issues"]

    tiendas_prop = [p for p in prop if p["segmento"] == "TIENDAS"]
    rest_prop = [p for p in prop if p["segmento"] == "RESTAURANTES" and p["clave"] == "balde4"]
    tiendas_excel = [e for e in excel if e["segmento"] == "TIENDAS"]

    def mn(lst, k):
        v = [x[k] for x in lst if x[k] is not None]
        return min(v) if v else None

    def mx(lst, k):
        v = [x[k] for x in lst if x[k] is not None]
        return max(v) if v else None

    # Margen más bajo de toda la escala propuesta (criterio Suplevet: el punto de equilibrio se
    # calcula siempre con el margen más conservador).
    peor = min(prop, key=lambda p: p["margen_mrpeanutt_pct"])

    pe = data["punto_equilibrio"]
    pe_cons = next(p for p in pe if p["es_conservador"])
    pe_mani = next(p for p in pe if p["canal"] == "Directo" and p["clave"] == "mani")
    cf_total = sum(c["monto_mensual"] for c in data["costos_fijos"])

    data["kpis"] = {
        "n_productos": len(data["productos_costeo"]),
        "costos_fijos_total": round(cf_total, 2),
        "pe_conservador_sin_igv": pe_cons["venta_equilibrio_sin_igv"],
        "pe_conservador_con_igv": pe_cons["venta_equilibrio_con_igv"],
        "pe_conservador_margen": pe_cons["margen_pct"],
        "pe_conservador_unidades": pe_cons["unidades_equilibrio"],
        "pe_directo_mani_con_igv": pe_mani["venta_equilibrio_con_igv"],
        "pe_directo_mani_unidades": pe_mani["unidades_equilibrio"],
        "margen_directo_min": mn(directo, "margen_pct"),
        "margen_directo_max": mx(directo, "margen_pct"),
        "promos_margen_min": mn(data["promociones_propuesta"], "margen_min_pct"),
        "promos_margen_max": mx(data["promociones_propuesta"], "margen_max_pct"),
        "margen_tiendas_excel_min": mn(tiendas_excel, "margen_mrpeanutt_pct"),
        "margen_tiendas_excel_max": mx(tiendas_excel, "margen_mrpeanutt_pct"),
        "margen_tienda_excel_min": mn(tiendas_excel, "margen_tienda_pct"),
        "margen_tienda_excel_max": mx(tiendas_excel, "margen_tienda_pct"),
        "margen_tiendas_prop_min": mn(tiendas_prop, "margen_mrpeanutt_pct"),
        "margen_tiendas_prop_max": mx(tiendas_prop, "margen_mrpeanutt_pct"),
        "margen_tienda_prop_min": mn(tiendas_prop, "margen_tienda_pct"),
        "margen_tienda_prop_max": mx(tiendas_prop, "margen_tienda_pct"),
        "margen_restaurantes_min": mn(rest_prop, "margen_mrpeanutt_pct"),
        "margen_restaurantes_max": mx(rest_prop, "margen_mrpeanutt_pct"),
        "balde4_precio_kg_min": mn(rest_prop, "precio_por_kg_con_igv"),
        "balde4_precio_kg_max": mx(rest_prop, "precio_por_kg_con_igv"),
        "margen_conservador": peor["margen_mrpeanutt_pct"],
        "margen_conservador_origen": f"{peor['producto']} · {peor['nombre_tramo']}",
        "issues_alta": sum(1 for i in issues if i["severidad"] == "alta"),
        "issues_media": sum(1 for i in issues if i["severidad"] == "media"),
        "issues_baja": sum(1 for i in issues if i["severidad"] == "baja"),
        "issues_total": len(issues),
        "issues_pendientes": sum(1 for i in issues if i["estado"] == "pendiente"),
    }

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print("Exportado:", OUT_PATH, f"({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()

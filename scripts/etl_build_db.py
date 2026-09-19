# -*- coding: utf-8 -*-
"""
ETL: lee el Excel de márgenes de MR. PEANUTT (DATOS INICIALES/) y construye la base de datos
local SQLite (BD/mrpeanutt_costos.db) con todas las tablas del centro de costos.

Cómo usarlo cuando actualices el Excel:
    1. Reemplaza el archivo dentro de "DATOS INICIALES" (se toma el primer .xlsx de la carpeta).
    2. Corre la cadena completa (con PYTHONIOENCODING=utf-8):
           python etl_build_db.py
           python export_dashboard_json.py
           python build_dashboard.py
           python export_excel.py

Qué hace este script:
  - Lee SOLO los insumos (costos unitarios, márgenes, umbrales de compra) del Excel y recalcula
    todas las fórmulas aquí, para no depender de los valores cacheados del archivo.
  - Reconstruye tal cual la escala B2B del Excel ("as-is") y, aparte, arma la PROPUESTA de escala
    por tramo para los dos segmentos B2B (tiendas = frascos 150 g, restaurantes = baldes) con el
    margen que se lleva Mr. Peanutt y el margen que le queda a la tienda al revender al PVP.
  - Registra en la tabla `issues` cada inconsistencia encontrada en el Excel y en
    `recomendaciones` las decisiones sugeridas.

Supuestos de negocio (todos marcados como tales y revisables al inicio del archivo):
  A1. El canal directo NO paga IGV (régimen NRUS, decisión del usuario 2026-09-18): el PVP es
      precio neto y el margen del Excel (p. ej. 74.8%) es el margen real. El canal directo fija el
      precio de mercado: las tiendas revenden al mismo PVP (almendra S/25, no S/29.90).
  A2. Los umbrales de compra mínima de la escala B2B se comparan contra el total del pedido CON
      IGV (lo que realmente paga el cliente). El Excel dividía el umbral entre el precio sin IGV.
  A3. Por B2B (tiendas) SOLO se venden maní y almendra (decisión del usuario, 2026-09-18);
      chocomaní y crunchy van únicamente por el canal directo. El tramo de descuento se define por
      el TOTAL del pedido sumando ambos productos, y el descuento del tramo se aplica a los dos.
  A4. El balde de 1 kg solo trae el costo del envase (S/3.00): se le propone el costeo completo
      envase + 1 kg de mantequilla + etiqueta (misma etiqueta de S/1.70 que el de 4 kg).
"""
import datetime
import sqlite3
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
DATOS = BASE / "DATOS INICIALES"
DB_PATH = BASE / "BD" / "mrpeanutt_costos.db"

IGV = 0.18
_IGV = 1 + IGV

# ------------------------------------------------------------------------------------------------
# PARÁMETROS DE LA PROPUESTA (decisiones sugeridas; se pueden ajustar y volver a correr el ETL)
# ------------------------------------------------------------------------------------------------

# PVP del canal directo (sin IGV, NRUS) y precio de reventa sugerido para las tiendas. Sale del
# precio de lista del Excel redondeado a múltiplos de S/0.50 (maní 23, chocomaní/crunchy 23.50,
# almendra 25). El canal directo es el que fija el precio de mercado.
def redondear_pvp(x, paso=0.50):
    return round(round(x / paso) * paso, 2)

# Promociones del canal directo (sin IGV): mecánica "ARMA TU PACK" (decisión de diseño, 2026-09-18).
# En vez de combos con nombre o % de descuento, se vende un PRECIO FIJO POR CANTIDAD: el cliente
# elige los sabores del catálogo y el precio del pack no cambia. Máximo 1 almendra por pack mixto
# (cuesta casi el doble); los packs de SOLO almendra tienen su propio precio. El precio por frasco
# baja con la cantidad (21.00 -> 20.67 -> 20.50) para empujar al pack siguiente. Regla dura: ningún
# pack, con ninguna mezcla posible, baja de MARGEN_PROMO_MINIMO.
# PRECIOS (decisión del usuario, 2026-09-18): "el mercado peruano no va a pagar esos precios". El
# Pack 2 se queda en S/42. Pack 3 y Pack 4 se fijan con MARGEN_OBJETIVO_PACK sobre la PEOR mezcla
# posible (2 chocomaní + 1 almendra / 3 chocomaní + 1 almendra): 55% -> S/51 y S/64. Con 50% serían
# S/46 y S/57 (ver tabla packs_escenarios). El precio se redondea a soles enteros.
MARGEN_OBJETIVO_PACK = {3: 0.55, 4: 0.55}
MARGENES_ESCENARIO_PACK = [0.65, 0.60, 0.55, 0.50, 0.45]
#   (pack, n_frascos, precio_pack (None = se calcula con MARGEN_OBJETIVO_PACK), solo_almendra, max_almendra, ejemplo)
PACKS_PROPUESTA = [
    ("Pack 2", 2, 42.0, False, 1, {"mani": 1, "chocomani": 1}),
    ("Pack 3", 3, None, False, 1, {"mani": 1, "chocomani": 1, "almendra": 1}),
    ("Pack 4", 4, None, False, 1, {"mani": 1, "chocomani": 1, "crunchy": 1, "almendra": 1}),
    # Packs de SOLO almendra: combos oficiales con precio fijo (decisión del usuario, 2026-09-18, según el
    # término n.° 3). No siguen el 55%: con ese margen el x2 saldría a S/49 (2% de descuento, no motiva).
    # Se aceptan con ~50%: x2 S/46 (52.2%, ahorra S/4), x3 S/66 (50%, ahorra S/9).
    ("Pack almendra x2", 2, 46.0, True, 2, {"almendra": 2}),
    ("Pack almendra x3", 3, 66.0, True, 3, {"almendra": 3}),
]
MARGEN_PROMO_MINIMO = 0.50

# Términos y condiciones sugeridos para la mecánica de packs (se muestran tal cual en el dashboard).
#   (orden, termino, por_que)
PROMOS_TERMINOS = [
    (1, "Precio fijo por cantidad: 2 frascos S/42, 3 frascos S/51, 4 frascos S/64, elige los sabores que quieras del catálogo.",
     "Un solo mensaje, sin porcentajes ni decimales. El cliente arma su pack y siente que gana; Mr. Peanutt controla el margen."),
    (2, "Máximo 1 mantequilla de almendra por pack mixto. Puedes repetir sabor (p. ej. 2 de maní).",
     "La almendra cuesta S/11 (casi el doble). Con 1 por pack el margen nunca baja del objetivo (55%); con 2 sí bajaría."),
    (3, "Packs de solo almendra como combos oficiales con precio propio: Pack almendra x2 por S/46, Pack almendra x3 por S/66.",
     "Al que solo quiere almendra se le da su promo sin romper la regla anterior. No siguen el 55% de los mixtos (saldrían a S/49 y S/73, "
     "casi sin descuento): se aceptan con 50-52% porque la almendra cuesta S/11."),
    (4, "La almendra es el 'upgrade' del pack: elegirla no cuesta más.",
     "Quien pone la almendra siente más descuento y Mr. Peanutt sigue por encima del margen objetivo: es el gancho para hacerla probar."),
    (5, "El precio unitario (S/23 / S/23.50 / S/25) no se toca ni se descuenta nunca: el pack es el único descuento.",
     "Si el suelto también tiene descuento, el pack pierde sentido. El pack existe para subir el ticket, no para bajar el precio."),
    (6, "Comunicar siempre el ahorro en soles con el ejemplo más común ('3 frascos por S/51, ahorras S/20'), nunca el porcentaje.",
     "S/8 se entiende al instante; 11.4% obliga a calcular. Se cuenta el ahorro sobre el pack más barato posible para no exagerar."),
    (7, "Vigencia permanente en el catálogo, no 'solo por hoy'.",
     "Es la escala de precios del consumidor final (igual que la escala de tiendas): lo que se busca es que el ticket promedio suba siempre."),
    (8, "No acumulable con otros descuentos, cupones ni con la escala B2B.",
     "Evita que un descuento se monte sobre otro y el margen caiga por debajo del piso sin darse cuenta."),
    (9, "Si hay costo de delivery, incluirlo gratis desde el Pack 3 (a validar con el costo real del envío).",
     "El delivery gratis es el motivo n.° 1 para subir de 2 a 3 frascos; hoy no hay dato del costo de envío, por eso queda condicionado."),
    (10, "Cualquier pack nuevo se valida en el armador de promociones: margen mínimo 50% con la peor mezcla posible.",
     "Regla de control: la promo se diseña desde el margen, no desde el descuento (así se evita lo que pasó con la promo de 3 almendras)."),
    (11, "El precio por frasco dentro del pack no debe bajar de lo que paga la tienda en Comercial (S/15.20 con IGV).",
     "Si el público consigue en el canal directo un precio igual o menor al que paga la tienda, la tienda deja de comprar. "
     "Pack 3 a S/51 = S/17.00 por frasco y Pack 4 a S/64 = S/16.00: todavía por encima, pero ya cerca del límite."),
]

# Tramos de la escala B2B PROPUESTA. La compra mínima es POR PEDIDO, con IGV (supuesto A2).
#   Tiendas (frascos): se cierra el hueco S/500-700 del Excel dejando Comercial en "< S/700".
#   Restaurantes (baldes): los tramos se definen por CANTIDAD DE BALDES, que es más fácil de
#   comunicar y evita la circularidad "umbral / precio del tramo" del Excel (S/400 / 97.69 = 4.09).
TRAMOS_TIENDAS = [
    # (tramo, nombre, compra_minima_con_igv, etiqueta condición)
    (1, "Comercial", 0.0, "Pedidos menores a S/700"),
    (2, "Mayorista", 700.0, "Pedidos desde S/700"),
    (3, "Distribuidor", 1000.0, "Pedidos desde S/1,000"),
    (4, "Exclusivo", 1600.0, "Pedidos desde S/1,600"),
]
TRAMOS_RESTAURANTES = [
    # (tramo, nombre, baldes_minimos, etiqueta condición)
    (1, "Comercial", 1, "1 a 3 baldes"),
    (2, "Mayorista", 4, "4 a 6 baldes"),
    (3, "Distribuidor", 7, "7 a 9 baldes"),
    (4, "Exclusivo", 10, "10 baldes o más"),
]

# Margen objetivo de Mr. Peanutt (sobre precio sin IGV) por tramo, en la PROPUESTA.
#   Frascos de maní/chocomaní/crunchy: se baja el Comercial de 58.6% a 55% para que la tienda que
#   compra al precio comercial gane al menos ~32% al PVP (hoy gana 28%). El resto queda como el
#   Excel (50 / 45 / 40).
#   Almendra: PVP fijo en S/25 (lo fija el canal directo). Se mantienen los márgenes del Excel
#   (30 / 22 / 19 / 15): la tienda gana 26-39%. Subir el margen dejaría a la tienda por debajo de 25%.
#   Baldes: se mantienen los márgenes del Excel (59 / 52 / 45.8 / 41.2), solo se redondea el precio.
MARGEN_PROPUESTA = {
    "mani": {1: 0.55, 2: 0.50, 3: 0.45, 4: 0.40},
    "almendra": {1: 0.30, 2: 0.22, 3: 0.19, 4: 0.15},
    "balde": {1: 0.59, 2: 0.52, 3: 0.458, 4: 0.412},
}
# Redondeo del precio con IGV en la propuesta (lo que se comunica al cliente).
PASO_REDONDEO_FRASCO = 0.10
PASO_REDONDEO_BALDE = 1.00

# Productos que se venden por B2B a tiendas (decisión del usuario, 2026-09-18): solo estos dos.
# Chocomaní y crunchy quedan solo en el canal directo.
PRODUCTOS_B2B_TIENDAS = ["mani", "almendra"]

# Costos fijos mensuales (decisión del usuario, 2026-09-18). La empresa no tiene más costos fijos:
# un sueldo de S/1,020, la cuota del NRUS (régimen tributario) que se toma en su tope, S/50, y el
# servicio Cloud de USD 20/mes (se mensualiza en soles con TIPO_CAMBIO_USD_PEN, el mismo que usa el
# centro de costos de Suplevet: BCRP 15-sep-2026).
TIPO_CAMBIO_USD_PEN = 3.38
CLOUD_USD_MENSUAL = 20.0
#   (concepto, categoria, monto_mensual, nota)
COSTOS_FIJOS = [
    ("Sueldo", "Personal", 1020.0, "Decisión del usuario 2026-09-18"),
    ("NRUS (cuota mensual del régimen tributario)", "Tributario", 50.0,
     "Se toma el tope de S/50 (la cuota real hoy es S/20)"),
    ("Cloud (USD 20/mes)", "Tecnología", round(CLOUD_USD_MENSUAL * TIPO_CAMBIO_USD_PEN, 2),
     f"USD {CLOUD_USD_MENSUAL:.0f} x S/{TIPO_CAMBIO_USD_PEN} (tipo de cambio BCRP 15-sep-2026)"),
]

# Porción de referencia para restaurantes (untado de un sándwich / crepe / bowl).
PORCION_GR = 30

# Margen mínimo "sano" que se le quiere dejar a la tienda al PVP (referencia retail de alimentos
# especializados: 30-40%). Se usa para marcar en rojo/ámbar las celdas de la escala.
MARGEN_TIENDA_MINIMO = 0.30


# ------------------------------------------------------------------------------------------------
# Utilidades
# ------------------------------------------------------------------------------------------------
def r2(x):
    return None if x is None else round(float(x), 2)


def r4(x):
    return None if x is None else round(float(x), 4)


def precio_desde_margen(costo, margen):
    """Precio sin IGV que deja `margen` sobre el precio: costo / (1 - margen)."""
    return costo / (1 - margen)


def margen_sobre_precio(costo, precio):
    return (precio - costo) / precio if precio else None


def excel_path():
    xs = sorted(DATOS.glob("*.xlsx"))
    if not xs:
        raise SystemExit(f"No hay ningún .xlsx en {DATOS}")
    return xs[0]


# ------------------------------------------------------------------------------------------------
# Lectura del Excel: solo insumos
# ------------------------------------------------------------------------------------------------
def leer_insumos(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    v = lambda c: ws[c].value

    insumos = {
        "archivo": path.name,
        "hoja": ws.title,
        # Canal directo (costo, margen de lista)
        "frascos": [
            # clave, nombre Excel, celda costo, costo, margen lista, familia de escala
            ("mani", v("B3"), "C3", float(v("C3")), float(v("D3")), "mani"),
            ("almendra", v("B4"), "C4", float(v("C4")), float(v("D4")), "almendra"),
            ("chocomani", v("B5"), "C5", float(v("C5")), float(v("D5")), "mani"),
            ("crunchy", v("B6"), "C6", float(v("C6")), float(v("D6")), "mani"),
        ],
        "promo2_almendra_margen": float(v("D9")),
        "promo3_almendra_margen": float(v("D10")),
        "costo_almendra_kg": float(v("H7")),
        "ref_I7": v("I7"),
        "ref_J7": v("J7"),
        "costo_mantequilla_kg": float(v("H8")),
        # Baldes
        "balde4_envase": float(v("C16")),
        "balde4_kg": 4.0,
        "balde_etiqueta": float(v("C18")),
        "balde1_envase": float(v("F16")),
        # Combo 3 mantequillas
        "combo3_margen": float(v("N19")),
        # Escala B2B as-is: (segmento, producto clave, fila, condición, margen)
        "escala_balde": [(1, v("A24"), v("B24"), float(v("D24")), v("H24")),
                         (2, v("A25"), v("B25"), float(v("D25")), None),
                         (3, v("A26"), v("B26"), float(v("D26")), None),
                         (4, v("A27"), v("B27"), float(v("D27")), None)],
        "escala_mani": [(1, v("A31"), v("B31"), float(v("D31"))),
                        (2, v("A32"), v("B32"), float(v("D32"))),
                        (3, v("A33"), v("B33"), float(v("D33"))),
                        (4, v("A34"), v("B34"), float(v("D34")))],
        "escala_almendra": [(1, v("A38"), v("B38"), float(v("D38"))),
                            (2, v("A39"), v("B39"), float(v("D39"))),
                            (3, v("A40"), v("B40"), float(v("D40"))),
                            (4, v("A41"), v("B41"), float(v("D41")))],
    }
    # Umbrales de compra que trae el Excel (texto "Compras mayores a 400 soles" -> 400)
    insumos["umbral_balde"] = {1: 0.0, 2: 400.0, 3: 650.0, 4: 900.0}
    insumos["umbral_frasco"] = {1: 0.0, 2: 700.0, 3: 1000.0, 4: 1600.0}
    insumos["umbral_frasco_comercial_txt"] = v("A31")   # "Pedidos menores a 500" (hueco 500-700)
    return insumos


# ------------------------------------------------------------------------------------------------
# Esquema
# ------------------------------------------------------------------------------------------------
SCHEMA = """
DROP TABLE IF EXISTS parametros;
DROP TABLE IF EXISTS productos_costeo;
DROP TABLE IF EXISTS costeo_balde_detalle;
DROP TABLE IF EXISTS canal_directo;
DROP TABLE IF EXISTS promociones;
DROP TABLE IF EXISTS promociones_propuesta;
DROP TABLE IF EXISTS promociones_ejemplos;
DROP TABLE IF EXISTS promociones_terminos;
DROP TABLE IF EXISTS packs_escenarios;
DROP TABLE IF EXISTS escala_b2b_excel;
DROP TABLE IF EXISTS escala_b2b_propuesta;
DROP TABLE IF EXISTS cadena_valor;
DROP TABLE IF EXISTS restaurantes_porcion;
DROP TABLE IF EXISTS resumen_margenes;
DROP TABLE IF EXISTS costos_fijos;
DROP TABLE IF EXISTS punto_equilibrio;
DROP TABLE IF EXISTS issues;
DROP TABLE IF EXISTS recomendaciones;
DROP TABLE IF EXISTS meta;

CREATE TABLE meta (clave TEXT PRIMARY KEY, valor TEXT);

CREATE TABLE parametros (
    id INTEGER PRIMARY KEY,
    clave TEXT, descripcion TEXT, valor REAL, unidad TEXT, fuente TEXT, nota TEXT
);

-- Costeo base de cada presentación.
CREATE TABLE productos_costeo (
    id INTEGER PRIMARY KEY,
    clave TEXT, producto TEXT, familia TEXT, presentacion TEXT, peso_kg REAL,
    costo_unitario REAL, costo_por_kg REAL,
    costo_insumo_principal_est REAL, costo_envase_y_otros_est REAL,
    estado TEXT, fuente TEXT
);

CREATE TABLE costeo_balde_detalle (
    id INTEGER PRIMARY KEY,
    producto TEXT, componente TEXT, cantidad REAL, unidad TEXT, costo_unitario REAL,
    costo_total REAL, estado TEXT, fuente TEXT
);

-- Canal directo (venta al consumidor final).
CREATE TABLE canal_directo (
    id INTEGER PRIMARY KEY,
    clave TEXT, producto TEXT, costo REAL,
    margen_excel_pct REAL, precio_lista REAL, pvp_redondeado REAL,
    margen_pct REAL, ganancia_unit REAL, precio_por_kg REAL, fuente TEXT
);

CREATE TABLE promociones (
    id INTEGER PRIMARY KEY,
    promocion TEXT, canal TEXT, composicion TEXT, costo REAL,
    margen_excel_pct REAL, precio_excel REAL,
    suma_precios_individuales REAL, descuento_vs_individual_pct REAL,
    ganancia_unit REAL, observacion TEXT, fuente TEXT
);

-- Mecánica PROPUESTA del canal directo: precio fijo por cantidad ("arma tu pack"). Una fila por
-- pack con el rango de descuento y margen sobre TODAS las mezclas posibles.
CREATE TABLE promociones_propuesta (
    id INTEGER PRIMARY KEY,
    pack TEXT, n_frascos INTEGER, precio_pack REAL, precio_por_frasco REAL, regla TEXT, combos_posibles INTEGER,
    costo_min REAL, costo_max REAL, sueltos_min REAL, sueltos_max REAL,
    descuento_min_pct REAL, descuento_max_pct REAL, ahorro_min REAL, ahorro_max REAL,
    margen_min_pct REAL, margen_max_pct REAL, mezcla_peor_margen TEXT, mezcla_mayor_descuento TEXT,
    ganancia_min REAL, ganancia_max REAL, cumple_minimo INTEGER,
    ejemplo TEXT, ejemplo_sueltos REAL, ejemplo_ahorro REAL, ejemplo_margen_pct REAL
);

-- Todas las mezclas posibles de cada pack, con su descuento y margen.
CREATE TABLE promociones_ejemplos (
    id INTEGER PRIMARY KEY,
    pack TEXT, n_frascos INTEGER, precio_pack REAL, mezcla TEXT, n_almendra INTEGER,
    costo REAL, sueltos REAL, descuento_pct REAL, ahorro REAL, margen_pct REAL, ganancia REAL, ganancia_cedida REAL
);

-- Escenarios de precio de cada pack según el margen objetivo sobre la peor mezcla.
CREATE TABLE packs_escenarios (
    id INTEGER PRIMARY KEY,
    pack TEXT, n_frascos INTEGER, margen_objetivo_pct REAL, precio_pack REAL, precio_por_frasco REAL,
    costo_peor_mezcla REAL, margen_min_pct REAL, margen_max_pct REAL,
    descuento_min_pct REAL, descuento_max_pct REAL, ahorro_min REAL, ahorro_max REAL,
    precio_tienda_comercial REAL, sobre_precio_tienda_pct REAL, es_propuesto INTEGER
);

-- Términos y condiciones sugeridos de la mecánica de packs.
CREATE TABLE promociones_terminos (
    id INTEGER PRIMARY KEY,
    orden INTEGER, termino TEXT, por_que TEXT
);

-- Escala B2B tal como está en el Excel (as-is), con el margen de la tienda al PVP actual.
CREATE TABLE escala_b2b_excel (
    id INTEGER PRIMARY KEY,
    segmento TEXT, producto TEXT, tramo INTEGER, nombre_tramo TEXT, condicion_excel TEXT,
    compra_minima_soles REAL, costo REAL, margen_mrpeanutt_pct REAL,
    precio_sin_igv REAL, precio_con_igv REAL, ganancia_unit_sin_igv REAL,
    descuento_vs_comercial_pct REAL, unidades_minimas REAL,
    precio_por_kg_con_igv REAL,
    pvp_referencia REAL, margen_tienda_pct REAL, ganancia_tienda_unit REAL, markup_tienda_pct REAL,
    fuente TEXT
);

-- Escala B2B PROPUESTA (precios redondeados, tramos sin huecos, márgenes objetivo).
CREATE TABLE escala_b2b_propuesta (
    id INTEGER PRIMARY KEY,
    segmento TEXT, producto TEXT, clave TEXT, tramo INTEGER, nombre_tramo TEXT, condicion TEXT,
    compra_minima_con_igv REAL, unidades_minimas INTEGER,
    costo REAL, margen_objetivo_pct REAL,
    precio_con_igv REAL, precio_sin_igv REAL, margen_mrpeanutt_pct REAL, ganancia_unit_sin_igv REAL,
    descuento_vs_comercial_pct REAL, precio_por_kg_con_igv REAL,
    pvp_sugerido REAL, margen_tienda_pct REAL, ganancia_tienda_unit REAL, markup_tienda_pct REAL,
    semaforo_tienda TEXT, cambio_vs_excel TEXT
);

-- Quién gana qué en cada eslabón cuando el producto pasa por la escala propuesta.
CREATE TABLE cadena_valor (
    id INTEGER PRIMARY KEY,
    escenario TEXT, producto TEXT, clave TEXT,
    pvp_con_igv REAL, pvp_sin_igv REAL, costo REAL, contribucion_total_sin_igv REAL,
    eslabon_1 TEXT, gana_1 REAL, margen_1_pct REAL,
    eslabon_2 TEXT, gana_2 REAL, margen_2_pct REAL,
    eslabon_3 TEXT, gana_3 REAL, margen_3_pct REAL,
    comentario TEXT
);

-- Qué le cuesta al restaurante cada porción según formato y tramo.
CREATE TABLE restaurantes_porcion (
    id INTEGER PRIMARY KEY,
    formato TEXT, tramo INTEGER, nombre_tramo TEXT, precio_con_igv REAL, peso_kg REAL,
    precio_por_kg_con_igv REAL, porcion_gr REAL, costo_porcion REAL,
    ahorro_vs_frasco_pct REAL, margen_mrpeanutt_pct REAL, ganancia_mrpeanutt_por_kg REAL
);

-- Resumen de márgenes de Mr. Peanutt por canal (para KPIs y gráfico).
CREATE TABLE resumen_margenes (
    id INTEGER PRIMARY KEY,
    canal TEXT, segmento TEXT, producto TEXT, clave TEXT, version TEXT,
    margen_min_pct REAL, margen_max_pct REAL, margen_tienda_min_pct REAL, margen_tienda_max_pct REAL
);

CREATE TABLE costos_fijos (
    id INTEGER PRIMARY KEY,
    concepto TEXT, categoria TEXT, monto_mensual REAL, nota TEXT
);

-- Punto de equilibrio mensual por escenario de canal (costos fijos / margen bruto del escenario).
CREATE TABLE punto_equilibrio (
    id INTEGER PRIMARY KEY,
    escenario TEXT, canal TEXT, producto TEXT, clave TEXT, tramo TEXT,
    costos_fijos REAL, margen_pct REAL,
    venta_equilibrio_sin_igv REAL, venta_equilibrio_con_igv REAL,
    precio_unit_sin_igv REAL, precio_unit_con_igv REAL, unidades_equilibrio REAL, unidades_por_dia REAL,
    es_conservador INTEGER, comentario TEXT
);

CREATE TABLE issues (
    id INTEGER PRIMARY KEY,
    severidad TEXT, area TEXT, archivo TEXT, hoja TEXT, celda TEXT,
    descripcion TEXT, estado TEXT DEFAULT 'pendiente', resolucion TEXT
);

CREATE TABLE recomendaciones (
    id INTEGER PRIMARY KEY,
    prioridad INTEGER, area TEXT, recomendacion TEXT, impacto TEXT
);
"""


# ------------------------------------------------------------------------------------------------
# Carga
# ------------------------------------------------------------------------------------------------
def load_parametros(cx, I):
    rows = [
        ("igv", "IGV", IGV, "%", "Ley", "Todos los precios B2B se cotizan sin IGV y se muestran con IGV"),
        ("costo_mantequilla_kg", "Costo de la mantequilla de maní a granel", I["costo_mantequilla_kg"], "S/ por kg", "H8", ""),
        ("costo_almendra_kg", "Costo de almendras", I["costo_almendra_kg"], "S/ por kg", "H7", ""),
        ("ref_I7", "Celda I7 (=H7/5) sin uso en las fórmulas", float(I["ref_I7"]) if I["ref_I7"] is not None else None, "S/", "I7", "Referencia suelta; ver issues"),
        ("ref_J7", "Celda J7 sin uso en las fórmulas", float(I["ref_J7"]) if I["ref_J7"] is not None else None, "S/", "J7", "Referencia suelta; ver issues"),
        ("balde4_envase", "Envase balde 4 kg", I["balde4_envase"], "S/ por unidad", "C16", ""),
        ("balde1_envase", "Envase balde 1 kg", I["balde1_envase"], "S/ por unidad", "F16", ""),
        ("balde_etiqueta", "Etiqueta del balde", I["balde_etiqueta"], "S/ por unidad", "C18", "Asumida igual para el balde de 1 kg (A4)"),
        ("canal_directo_igv", "IGV en el canal directo (NRUS: no se cobra)", 0.0, "%", "Decisión 2026-09-18", "El PVP es precio neto"),
        ("margen_promo_minimo", "Margen mínimo de cualquier promoción del canal directo", MARGEN_PROMO_MINIMO, "%", "Propuesta", ""),
        ("porcion_gr", "Porción de referencia para restaurantes", PORCION_GR, "gramos", "Propuesta", "Untado de un sándwich / crepe"),
        ("tipo_cambio_usd_pen", "Tipo de cambio USD -> PEN (para el Cloud)", TIPO_CAMBIO_USD_PEN, "S/ por USD", "BCRP 15-sep-2026", "Mismo que Suplevet"),
        ("margen_tienda_minimo", "Margen mínimo sano para la tienda al PVP", MARGEN_TIENDA_MINIMO, "%", "Propuesta", "Referencia retail 30-40%"),
    ]
    cx.executemany(
        "INSERT INTO parametros (clave, descripcion, valor, unidad, fuente, nota) VALUES (?,?,?,?,?,?)", rows)


def load_productos(cx, I):
    """Devuelve dict clave -> costo unitario para las tablas siguientes."""
    costos = {}
    rows = []
    kg_mant = I["costo_mantequilla_kg"]
    kg_alm = I["costo_almendra_kg"]
    for clave, nombre, celda, costo, _m, fam in I["frascos"]:
        insumo = 0.15 * (kg_alm if clave == "almendra" else kg_mant)
        rows.append((clave, nombre.strip(), "FRASCO", "Frasco 150 g", 0.15, costo, costo / 0.15,
                     r2(insumo), r2(costo - insumo), "Excel", f"Hoja 1!{celda}"))
        costos[clave] = costo
    balde4 = I["balde4_envase"] + I["balde4_kg"] * kg_mant + I["balde_etiqueta"]
    balde1 = I["balde1_envase"] + 1.0 * kg_mant + I["balde_etiqueta"]
    rows.append(("balde4", "Mr. Peanutt Balde 4 kg", "BALDE", "Balde 4 kg", 4.0, r2(balde4), r2(balde4 / 4),
                 r2(4 * kg_mant), r2(I["balde4_envase"] + I["balde_etiqueta"]), "Excel", "Hoja 1!C16:C19"))
    rows.append(("balde1", "Mr. Peanutt Balde 1 kg", "BALDE", "Balde 1 kg", 1.0, r2(balde1), r2(balde1),
                 r2(kg_mant), r2(I["balde1_envase"] + I["balde_etiqueta"]), "Propuesto (A4)", "Hoja 1!F16 + supuesto"))
    costos["balde4"], costos["balde1"] = balde4, balde1
    cx.executemany("""INSERT INTO productos_costeo
        (clave, producto, familia, presentacion, peso_kg, costo_unitario, costo_por_kg,
         costo_insumo_principal_est, costo_envase_y_otros_est, estado, fuente)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)""", rows)

    det = [
        ("Balde 4 kg", "Envase (balde)", 1, "unidad", I["balde4_envase"], I["balde4_envase"], "Excel", "C16"),
        ("Balde 4 kg", "Mantequilla de maní", 4, "kg", kg_mant, 4 * kg_mant, "Excel", "C17 (=4*10)"),
        ("Balde 4 kg", "Etiqueta", 1, "unidad", I["balde_etiqueta"], I["balde_etiqueta"], "Excel", "C18"),
        ("Balde 4 kg", "COSTO TOTAL", None, None, None, r2(balde4), "Excel", "C19"),
        ("Balde 1 kg", "Envase (balde)", 1, "unidad", I["balde1_envase"], I["balde1_envase"], "Excel", "F16"),
        ("Balde 1 kg", "Mantequilla de maní", 1, "kg", kg_mant, kg_mant, "Propuesto", "H8 x 1 kg"),
        ("Balde 1 kg", "Etiqueta", 1, "unidad", I["balde_etiqueta"], I["balde_etiqueta"], "Propuesto (A4)", "misma etiqueta C18"),
        ("Balde 1 kg", "COSTO TOTAL", None, None, None, r2(balde1), "Propuesto", "suma"),
    ]
    cx.executemany("""INSERT INTO costeo_balde_detalle
        (producto, componente, cantidad, unidad, costo_unitario, costo_total, estado, fuente)
        VALUES (?,?,?,?,?,?,?,?)""", det)
    return costos


def load_canal_directo(cx, I):
    """Devuelve dict clave -> PVP redondeado (sin IGV: el canal directo está en NRUS)."""
    pvp = {}
    rows = []
    for clave, nombre, celda, costo, margen, _fam in I["frascos"]:
        lista = precio_desde_margen(costo, margen)
        red = redondear_pvp(lista)
        rows.append((clave, nombre.strip(), costo, margen, r2(lista), red,
                     r4(margen_sobre_precio(costo, red)), r2(red - costo), r2(red / 0.15),
                     f"Hoja 1!{celda}:E{celda[1:]}"))
        pvp[clave] = red
    cx.executemany("""INSERT INTO canal_directo
        (clave, producto, costo, margen_excel_pct, precio_lista, pvp_redondeado, margen_pct, ganancia_unit,
         precio_por_kg, fuente) VALUES (?,?,?,?,?,?,?,?,?,?)""", rows)
    return pvp


def load_promociones(cx, I, costos, pvp):
    rows = []
    c_alm = costos["almendra"]
    # Promo 2 almendras
    costo = 2 * c_alm
    m = I["promo2_almendra_margen"]
    precio = precio_desde_margen(costo, m)
    suma = 2 * pvp["almendra"]
    rows.append(("Promoción 2 mantequillas de almendra", "Directo", "2 x Almendra 150 g", r2(costo), m, r2(precio),
                 r2(suma), r4(1 - precio / suma), r2(precio - costo),
                 "Sale MÁS CARA que comprar 2 frascos sueltos: el margen 61.2% está mal puesto (no es promoción). "
                 "Reemplazo propuesto: Pack almendra x2 a S/46 (mecánica arma tu pack).",
                 "Hoja 1!B9:E9"))
    # Promo 3 almendras
    costo = 3 * c_alm
    m = I["promo3_almendra_margen"]
    precio = precio_desde_margen(costo, m)
    suma = 3 * pvp["almendra"]
    rows.append(("Promoción 3 mantequillas de almendra", "Directo", "3 x Almendra 150 g", r2(costo), m, r2(precio),
                 r2(suma), r4(1 - precio / suma), r2(precio - costo),
                 "Margen 23%: regala casi toda la ganancia (43% de descuento). Reemplazo propuesto: Pack almendra x3 a S/66.",
                 "Hoja 1!B10:E10"))
    # Combo 3 mantequillas (maní + chocomaní + almendra), ubicado bajo CANAL B2B en el Excel
    costo = costos["mani"] + costos["chocomani"] + costos["almendra"]
    m = I["combo3_margen"]
    precio = precio_desde_margen(costo, m)
    suma = pvp["mani"] + pvp["chocomani"] + pvp["almendra"]
    rows.append(("Combo 3 mantequillas (máx. 1 almendra)", "Excel: bajo CANAL B2B",
                 "Maní 150 g + Chocomaní 150 g + Almendra 150 g", r2(costo), m, r2(precio),
                 r2(suma), r4(1 - precio / suma), r2(precio - costo),
                 "Para el canal directo es un 44% de descuento sobre los sueltos (S/71.50): demasiado. "
                 "Reemplazo propuesto: Pack 3 a S/51 con máx. 1 almendra (margen 55-66%).",
                 "Hoja 1!L14:O19"))
    cx.executemany("""INSERT INTO promociones
        (promocion, canal, composicion, costo, margen_excel_pct, precio_excel, suma_precios_individuales,
         descuento_vs_individual_pct, ganancia_unit, observacion, fuente)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)""", rows)


NOMBRE_CORTO = {"mani": "Maní", "almendra": "Almendra", "chocomani": "Chocomaní", "crunchy": "Crunchy"}


def _mezclas(n, solo_almendra, max_almendra):
    """Todas las mezclas (multiconjuntos) de n frascos que respetan la regla de almendra."""
    from itertools import combinations_with_replacement
    claves = ["almendra"] if solo_almendra else ["mani", "chocomani", "crunchy", "almendra"]
    out = []
    for combo in combinations_with_replacement(claves, n):
        comp = {k: combo.count(k) for k in claves if combo.count(k)}
        if comp.get("almendra", 0) <= max_almendra:
            out.append(comp)
    return out


def _texto_mezcla(comp):
    orden = ["mani", "chocomani", "crunchy", "almendra"]
    return " + ".join(f"{comp[k]} {NOMBRE_CORTO[k]}" if comp[k] > 1 else NOMBRE_CORTO[k] for k in orden if comp.get(k))


def load_promociones_propuesta(cx, costos, pvp):
    packs, ejemplos, escenarios = [], [], []
    # Precio al que la tienda compra en Comercial (con IGV): piso de referencia para el precio por frasco del pack.
    p_tienda = cx.execute("SELECT precio_con_igv FROM escala_b2b_propuesta WHERE clave='mani' AND tramo=1").fetchone()[0]
    for pack, n, precio, solo_alm, max_alm, ejemplo in PACKS_PROPUESTA:
        mezclas = _mezclas(n, solo_alm, max_alm)
        costo_peor = max(sum(costos[k] * v for k, v in c.items()) for c in mezclas)
        costo_mejor = min(sum(costos[k] * v for k, v in c.items()) for c in mezclas)
        s_min = min(sum(pvp[k] * v for k, v in c.items()) for c in mezclas)
        s_max = max(sum(pvp[k] * v for k, v in c.items()) for c in mezclas)
        if precio is None:
            precio = float(round(costo_peor / (1 - MARGEN_OBJETIVO_PACK[n])))
        if True:   # escenarios para todos los packs, incluidos los de solo almendra
            for m_obj in MARGENES_ESCENARIO_PACK:
                p_esc = float(round(costo_peor / (1 - m_obj)))
                escenarios.append((pack, n, m_obj, p_esc, r2(p_esc / n), r2(costo_peor),
                                   r4((p_esc - costo_peor) / p_esc), r4((p_esc - costo_mejor) / p_esc),
                                   r4(1 - p_esc / s_max), r4(1 - p_esc / s_min), r2(s_min - p_esc), r2(s_max - p_esc),
                                   p_tienda, r4(p_esc / n / p_tienda - 1), 1 if p_esc == precio else 0))
        filas = []
        for comp in mezclas:
            costo = sum(costos[k] * v for k, v in comp.items())
            suma = sum(pvp[k] * v for k, v in comp.items())
            filas.append((comp, costo, suma, 1 - precio / suma, suma - precio, margen_sobre_precio(costo, precio),
                          precio - costo))
            ejemplos.append((pack, n, precio, _texto_mezcla(comp), comp.get("almendra", 0), r2(costo), r2(suma),
                             r4(1 - precio / suma), r2(suma - precio), r4(margen_sobre_precio(costo, precio)),
                             r2(precio - costo), r2(suma - precio)))
        peor = min(filas, key=lambda f: f[5])
        mas_dto = max(filas, key=lambda f: f[3])
        ej_costo = sum(costos[k] * v for k, v in ejemplo.items())
        ej_suma = sum(pvp[k] * v for k, v in ejemplo.items())
        regla = ("Solo almendra" if solo_alm else f"Elige del catálogo; máx. {max_alm} almendra; se puede repetir sabor")
        packs.append((pack, n, precio, r2(precio / n), regla, len(filas),
                      r2(min(f[1] for f in filas)), r2(max(f[1] for f in filas)),
                      r2(min(f[2] for f in filas)), r2(max(f[2] for f in filas)),
                      r4(min(f[3] for f in filas)), r4(max(f[3] for f in filas)),
                      r2(min(f[4] for f in filas)), r2(max(f[4] for f in filas)),
                      r4(peor[5]), r4(max(f[5] for f in filas)), _texto_mezcla(peor[0]), _texto_mezcla(mas_dto[0]),
                      r2(min(f[6] for f in filas)), r2(max(f[6] for f in filas)),
                      1 if peor[5] >= MARGEN_PROMO_MINIMO else 0,
                      _texto_mezcla(ejemplo), r2(ej_suma), r2(ej_suma - precio), r4(margen_sobre_precio(ej_costo, precio))))
    cx.executemany("""INSERT INTO promociones_propuesta
        (pack, n_frascos, precio_pack, precio_por_frasco, regla, combos_posibles, costo_min, costo_max, sueltos_min, sueltos_max,
         descuento_min_pct, descuento_max_pct, ahorro_min, ahorro_max, margen_min_pct, margen_max_pct, mezcla_peor_margen,
         mezcla_mayor_descuento, ganancia_min, ganancia_max, cumple_minimo, ejemplo, ejemplo_sueltos, ejemplo_ahorro,
         ejemplo_margen_pct) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", packs)
    cx.executemany("""INSERT INTO promociones_ejemplos
        (pack, n_frascos, precio_pack, mezcla, n_almendra, costo, sueltos, descuento_pct, ahorro, margen_pct, ganancia, ganancia_cedida)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""", ejemplos)
    cx.executemany("INSERT INTO promociones_terminos (orden, termino, por_que) VALUES (?,?,?)", PROMOS_TERMINOS)
    cx.executemany("""INSERT INTO packs_escenarios
        (pack, n_frascos, margen_objetivo_pct, precio_pack, precio_por_frasco, costo_peor_mezcla, margen_min_pct, margen_max_pct,
         descuento_min_pct, descuento_max_pct, ahorro_min, ahorro_max, precio_tienda_comercial, sobre_precio_tienda_pct, es_propuesto)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", escenarios)


def _fila_escala(segmento, producto, tramo, nombre, condicion, umbral, costo, margen, precio_comercial_sin,
                 peso_kg, pvp_ref, fuente):
    sin = precio_desde_margen(costo, margen)
    con = sin * _IGV
    desc = 1 - sin / precio_comercial_sin if precio_comercial_sin else 0.0
    unidades = (umbral / con) if umbral else None
    if pvp_ref:
        m_tienda = (pvp_ref - con) / pvp_ref
        g_tienda = pvp_ref - con
        markup = (pvp_ref - con) / con
    else:
        m_tienda = g_tienda = markup = None
    return (segmento, producto, tramo, nombre, condicion, umbral, r2(costo), margen,
            r2(sin), r2(con), r2(sin - costo), r4(desc), r2(unidades) if unidades is not None else None,
            r2(con / peso_kg), pvp_ref, r4(m_tienda), r2(g_tienda), r4(markup), fuente)


def load_escala_excel(cx, I, costos, pvp):
    rows = []
    # Restaurantes: balde 4 kg
    c = costos["balde4"]
    base_sin = precio_desde_margen(c, I["escala_balde"][0][3])
    for tramo, cond, nombre, margen, _h in I["escala_balde"]:
        cond = cond or "Compra base (1-2 baldes)"
        rows.append(_fila_escala("RESTAURANTES", "Balde 4 kg", tramo, nombre, cond, I["umbral_balde"][tramo],
                                 c, margen, base_sin, 4.0, None, f"Hoja 1!A{23 + tramo}:H{23 + tramo}"))
    # Tiendas: maní
    c = costos["mani"]
    base_sin = precio_desde_margen(c, I["escala_mani"][0][3])
    for tramo, cond, nombre, margen in I["escala_mani"]:
        rows.append(_fila_escala("TIENDAS", "Mantequilla de maní 150 g", tramo, nombre, cond, I["umbral_frasco"][tramo],
                                 c, margen, base_sin, 0.15, pvp["mani"], f"Hoja 1!A{30 + tramo}:H{30 + tramo}"))
    # Tiendas: almendra
    c = costos["almendra"]
    base_sin = precio_desde_margen(c, I["escala_almendra"][0][3])
    for tramo, cond, nombre, margen in I["escala_almendra"]:
        rows.append(_fila_escala("TIENDAS", "Mantequilla de almendra 150 g", tramo, nombre, cond, I["umbral_frasco"][tramo],
                                 c, margen, base_sin, 0.15, pvp["almendra"], f"Hoja 1!A{37 + tramo}:H{37 + tramo}"))
    cx.executemany("""INSERT INTO escala_b2b_excel
        (segmento, producto, tramo, nombre_tramo, condicion_excel, compra_minima_soles, costo, margen_mrpeanutt_pct,
         precio_sin_igv, precio_con_igv, ganancia_unit_sin_igv, descuento_vs_comercial_pct, unidades_minimas,
         precio_por_kg_con_igv, pvp_referencia, margen_tienda_pct, ganancia_tienda_unit, markup_tienda_pct, fuente)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", rows)


def _semaforo_tienda(m):
    if m is None:
        return None
    if m >= MARGEN_TIENDA_MINIMO:
        return "ok"
    if m >= MARGEN_TIENDA_MINIMO - 0.05:
        return "ajustado"
    return "bajo"


def load_escala_propuesta(cx, I, costos, pvp):
    """Escala propuesta. Devuelve dict (clave, tramo) -> precio con IGV para las tablas de abajo."""
    precios = {}
    rows = []
    # Márgenes del Excel para comparar (por familia y tramo)
    m_excel = {
        "mani": {t: m for t, _c, _n, m in I["escala_mani"]},
        "almendra": {t: m for t, _c, _n, m in I["escala_almendra"]},
        "balde": {t: m for t, _c, _n, m, _h in I["escala_balde"]},
    }

    def fila(segmento, producto, clave, familia, costo, peso_kg, tramo, nombre, cond, minimo, unidades, pvp_sug, paso,
             cambio_extra=""):
        m_obj = MARGEN_PROPUESTA[familia][tramo]
        sin_obj = precio_desde_margen(costo, m_obj)
        con = round(round(sin_obj * _IGV / paso) * paso, 2)
        sin = con / _IGV
        m_real = margen_sobre_precio(costo, sin)
        if pvp_sug:
            m_tienda = (pvp_sug - con) / pvp_sug
            g_tienda = pvp_sug - con
            markup = (pvp_sug - con) / con
        else:
            m_tienda = g_tienda = markup = None
        m_ex = m_excel[familia].get(tramo)
        cambio = f"Margen Excel {m_ex*100:.1f}% -> propuesta {m_real*100:.1f}%" if m_ex is not None else "Nuevo"
        if abs((m_ex or 0) - m_real) < 0.005 and m_ex is not None:
            cambio = "Igual al Excel (solo redondeo de precio)"
        if cambio_extra:
            cambio += "; " + cambio_extra
        precios[(clave, tramo)] = con
        return [segmento, producto, clave, tramo, nombre, cond, minimo, unidades, r2(costo), m_obj,
                con, r2(sin), r4(m_real), r2(sin - costo), None, r2(con / peso_kg),
                pvp_sug, r4(m_tienda), r2(g_tienda), r4(markup), _semaforo_tienda(m_tienda), cambio]

    # --- Tiendas: frascos ---
    nombres = {"mani": "Mantequilla de maní 150 g", "almendra": "Mantequilla de almendra 150 g"}
    frascos = []
    for clave in PRODUCTOS_B2B_TIENDAS:
        frascos.append((clave, nombres[clave], "almendra" if clave == "almendra" else "mani", pvp[clave], ""))
    for clave, nombre, fam, pvp_sug, extra in frascos:
        costo = costos[clave]
        grupo = []
        for tramo, nombre_t, minimo, cond in TRAMOS_TIENDAS:
            grupo.append(fila("TIENDAS", nombre, clave, fam, costo, 0.15, tramo, nombre_t, cond, minimo, None,
                              pvp_sug, PASO_REDONDEO_FRASCO, extra))
        base_sin = grupo[0][11]
        for g in grupo:
            g[14] = r4(1 - g[11] / base_sin)                      # descuento vs comercial
            if g[6]:
                g[7] = int(-(-g[6] // g[10]))                     # unidades mínimas = ceil(mínimo / precio con IGV)
            else:
                g[7] = 1
        rows.extend(grupo)

    # --- Restaurantes: baldes ---
    for clave, nombre, peso in (("balde4", "Balde 4 kg", 4.0), ("balde1", "Balde 1 kg", 1.0)):
        costo = costos[clave]
        extra = "costeo propuesto (A4)" if clave == "balde1" else ""
        grupo = []
        for tramo, nombre_t, baldes_min, cond in TRAMOS_RESTAURANTES:
            g = fila("RESTAURANTES", nombre, clave, "balde", costo, peso, tramo, nombre_t, cond, None, baldes_min,
                     None, PASO_REDONDEO_BALDE, extra)
            grupo.append(g)
        base_sin = grupo[0][11]
        for g in grupo:
            g[14] = r4(1 - g[11] / base_sin)
            g[6] = r2(g[7] * g[10])                               # compra mínima = baldes mínimos x precio con IGV
        rows.extend(grupo)

    cx.executemany("""INSERT INTO escala_b2b_propuesta
        (segmento, producto, clave, tramo, nombre_tramo, condicion, compra_minima_con_igv, unidades_minimas,
         costo, margen_objetivo_pct, precio_con_igv, precio_sin_igv, margen_mrpeanutt_pct, ganancia_unit_sin_igv,
         descuento_vs_comercial_pct, precio_por_kg_con_igv, pvp_sugerido, margen_tienda_pct, ganancia_tienda_unit,
         markup_tienda_pct, semaforo_tienda, cambio_vs_excel) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        rows)
    return precios


def load_cadena_valor(cx, costos, pvp, precios):
    """Tres escenarios por frasco: venta directa, tienda que compra Comercial, y distribuidor
    (compra Exclusivo, revende a la tienda al precio Comercial, la tienda vende al PVP)."""
    rows = []
    productos = [
        ("mani", "Mantequilla de maní 150 g", pvp["mani"]),
        ("almendra", "Mantequilla de almendra 150 g", pvp["almendra"]),
    ]
    for clave, nombre, p in productos:
        costo = costos[clave]
        # La tienda sí está afecta a IGV: su ingreso neto es PVP / 1.18. El canal directo (NRUS) no.
        p_sin = p / _IGV
        contrib = p_sin - costo
        com_sin = precios[(clave, 1)] / _IGV
        exc_sin = precios[(clave, 4)] / _IGV
        # 1) Directo (sin IGV: todo el PVP es ingreso)
        rows.append(("1. Venta directa al público (sin IGV, NRUS)", nombre, clave, p, p, costo, r2(p - costo),
                     "Mr. Peanutt", r2(p - costo), r4((p - costo) / p),
                     None, None, None, None, None, None,
                     "Mr. Peanutt se queda con toda la contribución (pero asume venta, delivery y marketing al consumidor)."))
        # 2) Tienda compra al precio Comercial
        g_mp = com_sin - costo
        g_t = p_sin - com_sin
        rows.append(("2. Tienda compra al precio Comercial", nombre, clave, p, r2(p_sin), costo, r2(contrib),
                     "Mr. Peanutt", r2(g_mp), r4(g_mp / com_sin),
                     "Tienda", r2(g_t), r4(g_t / p_sin),
                     None, None, None,
                     "La tienda pone la góndola y el cliente; Mr. Peanutt vende por volumen."))
        # 3) Distribuidor compra Exclusivo y revende a la tienda al precio Comercial
        g_mp = exc_sin - costo
        g_d = com_sin - exc_sin
        g_t = p_sin - com_sin
        m_d = g_d / com_sin
        if m_d >= 0.22:
            com3 = (f"Los tramos Distribuidor/Exclusivo existen para quien revende a tiendas: aquí el distribuidor "
                    f"se queda con {m_d*100:.0f}% y la tienda sigue ganando {g_t / p_sin*100:.0f}%.")
        else:
            com3 = (f"Con costo S/{costo:.2f} y PVP S/{p:.0f} el producto NO aguanta tres eslabones: al distribuidor "
                    f"solo le quedan {m_d*100:.0f}%. Sugerido: vender la almendra por B2B solo en Comercial/Mayorista.")
        rows.append(("3. Distribuidor (Exclusivo) -> tienda (Comercial) -> público", nombre, clave, p, r2(p_sin), costo,
                     r2(contrib),
                     "Mr. Peanutt", r2(g_mp), r4(g_mp / exc_sin),
                     "Distribuidor", r2(g_d), r4(m_d),
                     "Tienda", r2(g_t), r4(g_t / p_sin), com3))
    cx.executemany("""INSERT INTO cadena_valor
        (escenario, producto, clave, pvp_con_igv, pvp_sin_igv, costo, contribucion_total_sin_igv,
         eslabon_1, gana_1, margen_1_pct, eslabon_2, gana_2, margen_2_pct, eslabon_3, gana_3, margen_3_pct, comentario)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", rows)


def load_restaurantes_porcion(cx, costos, precios):
    rows = []
    # Referencia: frasco de maní 150 g al precio Comercial B2B propuesto
    frasco_con = precios[("mani", 1)]
    frasco_kg = frasco_con / 0.15
    rows.append(("Frasco 150 g (precio Comercial tiendas)", 1, "Comercial", frasco_con, 0.15, r2(frasco_kg), PORCION_GR,
                 r2(frasco_kg * PORCION_GR / 1000), 0.0,
                 r4(margen_sobre_precio(costos["mani"], frasco_con / _IGV)), r2((frasco_con / _IGV - costos["mani"]) / 0.15)))
    for clave, nombre, peso in (("balde4", "Balde 4 kg", 4.0), ("balde1", "Balde 1 kg (propuesto)", 1.0)):
        for tramo, nombre_t, _b, _c in TRAMOS_RESTAURANTES:
            con = precios[(clave, tramo)]
            por_kg = con / peso
            rows.append((nombre, tramo, nombre_t, con, peso, r2(por_kg), PORCION_GR, r2(por_kg * PORCION_GR / 1000),
                         r4(1 - por_kg / frasco_kg), r4(margen_sobre_precio(costos[clave], con / _IGV)),
                         r2((con / _IGV - costos[clave]) / peso)))
    cx.executemany("""INSERT INTO restaurantes_porcion
        (formato, tramo, nombre_tramo, precio_con_igv, peso_kg, precio_por_kg_con_igv, porcion_gr, costo_porcion,
         ahorro_vs_frasco_pct, margen_mrpeanutt_pct, ganancia_mrpeanutt_por_kg) VALUES (?,?,?,?,?,?,?,?,?,?,?)""", rows)


def load_resumen_margenes(cx):
    rows = []
    for (clave, producto, m) in cx.execute(
            "SELECT clave, producto, margen_pct FROM canal_directo ORDER BY id").fetchall():
        rows.append(("DIRECTO", "Público", producto, clave, "actual", m, m, None, None))
    for (seg, producto, mn, mx, tmn, tmx) in cx.execute("""
            SELECT segmento, producto, MIN(margen_mrpeanutt_pct), MAX(margen_mrpeanutt_pct),
                   MIN(margen_tienda_pct), MAX(margen_tienda_pct)
            FROM escala_b2b_excel GROUP BY segmento, producto ORDER BY MIN(id)""").fetchall():
        rows.append(("B2B", seg, producto, None, "excel", mn, mx, tmn, tmx))
    for (seg, producto, clave, mn, mx, tmn, tmx) in cx.execute("""
            SELECT segmento, producto, clave, MIN(margen_mrpeanutt_pct), MAX(margen_mrpeanutt_pct),
                   MIN(margen_tienda_pct), MAX(margen_tienda_pct)
            FROM escala_b2b_propuesta GROUP BY segmento, producto, clave ORDER BY MIN(id)""").fetchall():
        rows.append(("B2B", seg, producto, clave, "propuesta", mn, mx, tmn, tmx))
    cx.executemany("""INSERT INTO resumen_margenes
        (canal, segmento, producto, clave, version, margen_min_pct, margen_max_pct, margen_tienda_min_pct,
         margen_tienda_max_pct) VALUES (?,?,?,?,?,?,?,?,?)""", rows)


def load_costos_fijos(cx):
    cx.executemany("INSERT INTO costos_fijos (concepto, categoria, monto_mensual, nota) VALUES (?,?,?,?)", COSTOS_FIJOS)
    return sum(m for _c, _k, m, _n in COSTOS_FIJOS)


def load_punto_equilibrio(cx, cf):
    """PE mensual por escenario: costos fijos / margen del escenario. El escenario 'conservador' usa el
    margen más bajo de toda la escala B2B propuesta (criterio heredado de Suplevet: siempre el margen
    más bajo). Unidades por día sobre 26 días de venta."""
    DIAS = 26
    rows = []

    def fila(escenario, canal, producto, clave, tramo, margen, p_sin, p_con, comentario, cons=0, igv=True):
        venta = cf / margen
        u = venta / p_sin
        rows.append((escenario, canal, producto, clave, tramo, r2(cf), r4(margen), r2(venta),
                     r2(venta * _IGV) if igv else r2(venta),
                     r2(p_sin), r2(p_con), r2(u), r2(u / DIAS), cons, comentario))

    for (clave, producto, p, m) in cx.execute(
            "SELECT clave, producto, pvp_redondeado, margen_pct FROM canal_directo ORDER BY id"):
        fila(f"Directo - {producto}", "Directo", producto, clave, "PVP", m, p, p,
             "Si todo se vendiera por canal directo a este producto (sin IGV, NRUS).", igv=False)
    for (seg, producto, clave, tramo, nombre, m, p_sin, p_con) in cx.execute(
            "SELECT segmento, producto, clave, tramo, nombre_tramo, margen_mrpeanutt_pct, precio_sin_igv, precio_con_igv "
            "FROM escala_b2b_propuesta WHERE tramo IN (1, 4) AND clave <> 'balde1' ORDER BY id"):
        canal = "Tiendas" if seg == "TIENDAS" else "Restaurantes"
        fila(f"{canal} - {producto} - {nombre}", canal, producto, clave, nombre, m, p_sin, p_con,
             f"Si todo se vendiera a {canal.lower()} en el tramo {nombre}.")
    peor = cx.execute("SELECT producto, nombre_tramo, margen_mrpeanutt_pct, precio_sin_igv, precio_con_igv, clave "
                      "FROM escala_b2b_propuesta ORDER BY margen_mrpeanutt_pct LIMIT 1").fetchone()
    fila("CONSERVADOR (margen más bajo de toda la escala)", "Conservador", peor[0], peor[5], peor[1], peor[2],
         peor[3], peor[4],
         f"Piso de venta mensual que cubre los costos fijos aun si TODO se vendiera al margen más bajo "
         f"({peor[0]}, tramo {peor[1]}). Cualquier mezcla real necesita vender MENOS que esto.", cons=1)
    cx.executemany("""INSERT INTO punto_equilibrio
        (escenario, canal, producto, clave, tramo, costos_fijos, margen_pct, venta_equilibrio_sin_igv,
         venta_equilibrio_con_igv, precio_unit_sin_igv, precio_unit_con_igv, unidades_equilibrio, unidades_por_dia,
         es_conservador, comentario) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", rows)


def load_issues(cx, I, pvp):
    arch, hoja = I["archivo"], I["hoja"]
    p2 = cx.execute("SELECT precio_excel, suma_precios_individuales FROM promociones WHERE id=1").fetchone()
    p3 = cx.execute("SELECT margen_excel_pct FROM promociones WHERE id=2").fetchone()[0]
    m_dir = dict(cx.execute("SELECT clave, margen_pct FROM canal_directo").fetchall())
    alm = {t: (p, mt) for t, p, mt in cx.execute(
        "SELECT tramo, precio_con_igv, margen_tienda_pct FROM escala_b2b_propuesta WHERE clave='almendra'").fetchall()}
    t_com = cx.execute("SELECT margen_tienda_pct FROM escala_b2b_excel WHERE segmento='TIENDAS' AND tramo=1 ORDER BY id").fetchall()
    prop_mani = cx.execute("SELECT precio_con_igv, margen_tienda_pct FROM escala_b2b_propuesta "
                           "WHERE clave='mani' AND tramo=1").fetchone()
    rows = [
        ("alta", "Promociones", arch, hoja, "D9:E9",
         f"La 'Promoción 2 mantequillas de almendra' sale a S/{p2[0]:.2f}, MÁS CARA que 2 frascos sueltos "
         f"(S/{p2[1]:.2f}). El margen 61.2% no corresponde a una promoción.",
         "mitigado", "Propuesta: Pack almendra x2 a S/46 (8% de descuento, margen 52%). Ver promociones propuestas."),
        ("alta", "Promociones", arch, hoja, "D10:E10",
         f"La 'Promoción 3 almendras' tiene margen {p3*100:.2f}% (43% de descuento sobre 3 sueltas): casi no deja ganancia.",
         "mitigado", "Propuesta: Pack almendra x3 a S/66 (12% de descuento, margen 50%). Regla: ninguna promo baja de 50%."),
        ("alta", "Escala B2B", arch, hoja, "D38:D41",
         "Almendra: el margen de Mr. Peanutt en la escala B2B es 30 / 22 / 19 / 15% y la tienda gana solo "
         f"{t_com[1][0]*100:.0f}% al PVP de S/{pvp['almendra']:.0f}. Con costo S/11 y PVP S/25 no hay espacio para "
         "dos intermediarios.",
         "mitigado", f"Decisión del usuario: el PVP lo fija el canal directo (S/25) y la tienda revende a ese precio. Se mantienen "
                     f"los márgenes del Excel: la tienda gana {alm[1][1]*100:.0f}% (Comercial) a {alm[4][1]*100:.0f}% (Exclusivo). "
                     "La almendra no aguanta un distribuidor intermedio: venderla por B2B solo en Comercial/Mayorista."),
        ("media", "Escala B2B", arch, hoja, "A31:A32",
         "Tramos de tiendas con hueco: 'Pedidos menores a 500' y luego 'mayores a 700'. Un pedido de S/600 no cae en ningún tramo.",
         "mitigado", "Propuesta: Comercial = menos de S/700; Mayorista desde S/700."),
        ("media", "Escala B2B", arch, hoja, "A24:B24 / H24",
         "El tramo Comercial del balde no tiene condición, y la celda de 'Unidades mínimas' quedó como fecha (01/02/2026): "
         "probablemente se escribió '1-2'.",
         "mitigado", "Propuesta: tramos del balde por cantidad (1-3 / 4-6 / 7-9 / 10+)."),
        ("media", "Escala B2B", arch, hoja, "H25:H27",
         "Las unidades mínimas se calculan dividiendo el umbral entre el precio SIN IGV del mismo tramo (circular y sin IGV): "
         "4.09 baldes para S/400, cuando el cliente paga con IGV.",
         "mitigado", "Propuesta: umbrales sobre el total del pedido con IGV y redondeados a unidades enteras (A2)."),
        ("media", "Canal directo", arch, hoja, "E3:E6",
         "El precio de lista del canal directo no separa el IGV.",
         "resuelto", "Decisión del usuario (2026-09-18): el canal directo está en NRUS y no cobra IGV, así que el PVP es precio neto "
                     f"y el margen del Excel es el real (maní {m_dir['mani']*100:.1f}%, almendra {m_dir['almendra']*100:.1f}% "
                     "con el PVP redondeado). Supuesto A1."),
        ("baja", "Tributario", arch, hoja, "-",
         "El canal directo está en NRUS (sin IGV), pero la escala B2B del Excel cotiza precios con IGV: en NRUS no se emiten "
         "facturas, y las tiendas y restaurantes suelen pedir factura para usar el crédito fiscal.",
         "pendiente", "Confirmar con el contador si las ventas B2B se facturan desde otro RUC/régimen o si se venden con boleta."),
        ("media", "Escala B2B", arch, hoja, "D31",
         f"Tramo Comercial del maní: la tienda gana {t_com[0][0]*100:.1f}% al PVP, por debajo del 30% que suele pedir el retail.",
         "mitigado", f"Propuesta: margen Comercial 55% -> precio S/{prop_mani[0]:.2f} con IGV -> la tienda gana {prop_mani[1]*100:.1f}%."),
        ("media", "Combos", arch, hoja, "L14:O19",
         "El 'Combo 3 mantequillas' (S/39.82) no dice si es sin IGV ni a qué canal pertenece (está bajo CANAL B2B).",
         "pendiente", "Confirmar canal y si el precio lleva IGV. Con IGV sería S/46.99."),
        ("baja", "Escala B2B", arch, hoja, "-",
         "Chocomaní y Crunchy no tienen escala B2B en el Excel.",
         "resuelto", "Decisión del usuario (2026-09-18): por B2B solo van maní y almendra; chocomaní y crunchy "
                     "se venden únicamente por el canal directo. Supuesto A3."),
        ("baja", "Costeo", arch, hoja, "E15:F16",
         "El balde de 1 kg solo tiene el costo del envase (S/3.00); falta mantequilla y etiqueta.",
         "mitigado", "Propuesta: S/3.00 + 1 kg x S/10 + etiqueta S/1.70 = S/14.70. Supuesto A4."),
        ("baja", "Costeo", arch, hoja, "I7:J7",
         "Las celdas I7 (=H7/5 = 8) y J7 (10.2) no se usan en ninguna fórmula; parecen un tanteo del costo de la almendra.",
         "pendiente", "Confirmar si el costo S/11 del frasco de almendra ya incluye envase y mano de obra."),
        ("baja", "Punto de equilibrio", arch, hoja, "-",
         "No hay costos fijos mensuales en el Excel, así que no se podía calcular el punto de equilibrio de Mr. Peanutt.",
         "resuelto", f"Decisión del usuario (2026-09-18): sueldo S/1,020 + NRUS S/50 (tope) + Cloud USD 20 = "
                     f"S/{sum(m for _c, _k, m, _n in COSTOS_FIJOS):,.2f}/mes. Se calcula el PE por escenario y el "
                     "conservador con el margen más bajo (tabla punto_equilibrio)."),
    ]
    cx.executemany("""INSERT INTO issues (severidad, area, archivo, hoja, celda, descripcion, estado, resolucion)
        VALUES (?,?,?,?,?,?,?,?)""", rows)


def load_recomendaciones(cx):
    P = {(c, t): (p, mt) for c, t, p, mt in cx.execute(
        "SELECT clave, tramo, precio_con_igv, margen_tienda_pct FROM escala_b2b_propuesta").fetchall()}
    pvp = dict(cx.execute("SELECT clave, pvp_redondeado FROM canal_directo").fetchall())
    porc = {t: c for t, c in cx.execute(
        "SELECT tramo, costo_porcion FROM restaurantes_porcion WHERE formato='Balde 4 kg'").fetchall()}
    b4 = " / ".join(f"S/{P[('balde4', t)][0]:.0f}" for t in (1, 2, 3, 4))
    b1 = " / ".join(f"S/{P[('balde1', t)][0]:.0f}" for t in (1, 2, 3, 4))
    kg4 = cx.execute("SELECT precio_por_kg_con_igv FROM escala_b2b_propuesta WHERE clave='balde4' AND tramo=1").fetchone()[0]
    kg1 = cx.execute("SELECT precio_por_kg_con_igv FROM escala_b2b_propuesta WHERE clave='balde1' AND tramo=1").fetchone()[0]
    cf_total = cx.execute("SELECT SUM(monto_mensual) FROM costos_fijos").fetchone()[0]
    cx.row_factory = sqlite3.Row
    pe_cons = cx.execute("SELECT * FROM punto_equilibrio WHERE es_conservador=1").fetchone()
    pe_dir = cx.execute("SELECT * FROM punto_equilibrio WHERE canal='Directo' AND clave='mani'").fetchone()
    cx.row_factory = None
    rows = [
        (1, "Tiendas", f"El canal directo fija el precio de mercado: PVP maní S/{pvp['mani']:.2f}, chocomaní/crunchy "
            f"S/{pvp['chocomani']:.2f}, almendra S/{pvp['almendra']:.2f}. Exigirlo a las tiendas como precio mínimo de reventa "
            "para que el canal directo y las tiendas no se pisen.",
         f"Protege el margen del canal directo y da a la tienda un margen conocido ({P[('mani',1)][1]*100:.0f}-{P[('mani',4)][1]*100:.0f}%)."),
        (2, "Tiendas", "Escala de 4 tramos por pedido con IGV, sin huecos: Comercial < S/700, Mayorista desde S/700, "
            "Distribuidor desde S/1,000, Exclusivo desde S/1,600. Precios con IGV redondeados a S/0.10.",
         f"Mr. Peanutt gana 55 / 50 / 45 / 40% en maní; la tienda gana {P[('mani',1)][1]*100:.1f}% comprando Comercial "
         f"(S/{P[('mani',1)][0]:.2f})."),
        (3, "Almendra", f"Con PVP fijo en S/{pvp['almendra']:.0f} y costo S/11, ofrecer la almendra a tiendas solo en Comercial "
            f"(S/{P[('almendra',1)][0]:.2f}) y Mayorista (S/{P[('almendra',2)][0]:.2f}); no darle tramo Distribuidor/Exclusivo "
            "(dejan 19% y 15% a Mr. Peanutt).",
         f"La tienda gana {P[('almendra',1)][1]*100:.0f}-{P[('almendra',2)][1]*100:.0f}% y Mr. Peanutt no baja de 22%. "
         "Si baja el costo de la almendra, se revisa."),
        (4, "Restaurantes", f"Vender el balde de 4 kg por cantidad (1-3 / 4-6 / 7-9 / 10+ baldes): {b4} con IGV. "
            f"Comunicar el costo por porción de {PORCION_GR:.0f} g (S/{porc[1]:.2f} -> S/{porc[4]:.2f}), no el descuento.",
         "Margen Mr. Peanutt 59 -> 41%; el restaurante paga ~70-77% menos por kg que en frasco."),
        (5, "Restaurantes", f"Lanzar el balde de 1 kg (costo propuesto S/14.70) con la misma escala: {b1} con IGV.",
         f"Formato de entrada para cafeterías chicas; S/{kg1:.0f}/kg vs S/{kg4:.2f}/kg del balde de 4 kg."),
        (6, "Promociones", "Mecánica 'arma tu pack' con precio fijo por cantidad: 2 frascos S/42, 3 frascos S/51, 4 frascos S/64 "
            "(elige sabores, máx. 1 almendra, se puede repetir); packs de solo almendra 2 por S/46 y 3 por S/66. Pack 3 y 4 "
            "fijados al 55% de margen sobre la peor mezcla (con 50% serían S/46 y S/57). El suelto no se descuenta nunca; "
            "se comunica el ahorro en soles, no el porcentaje.",
         "Pack 3: cliente ahorra S/18-21 (26-29%), margen 55-66%. Pack 4: ahorra S/28-32 (30-33%), margen 55-64%. "
         "Ojo: S/17 y S/16 por frasco ya se acercan a lo que paga la tienda (S/15.20)."),
        (7, "Control", f"Con costos fijos de S/{cf_total:,.0f}/mes, el punto de equilibrio conservador (margen más bajo de "
            f"la escala, {pe_cons['margen_pct']*100:.1f}%) es S/{pe_cons['venta_equilibrio_con_igv']:,.0f}/mes con IGV. "
            "Revisar cada mes si la venta B2B supera ese piso; si se vende más por canal directo, el piso baja.",
         f"Vendiendo solo por canal directo el equilibrio baja a S/{pe_dir['venta_equilibrio_con_igv']:,.0f}/mes con IGV "
         f"({pe_dir['unidades_equilibrio']:.0f} frascos)."),
        (8, "Control", "Regla para distribuidores: quien compra en Distribuidor/Exclusivo debe revender a tiendas al precio "
            "Comercial (no por debajo) y al público al PVP sugerido.",
         "Mantiene los ~25% del distribuidor y los 30%+ de la tienda sin canibalizar el canal directo."),
    ]
    cx.executemany("INSERT INTO recomendaciones (prioridad, area, recomendacion, impacto) VALUES (?,?,?,?)", rows)


def main():
    path = excel_path()
    I = leer_insumos(path)
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    if DB_PATH.exists():
        DB_PATH.unlink()
    cx = sqlite3.connect(DB_PATH)
    cx.executescript(SCHEMA)

    load_parametros(cx, I)
    costos = load_productos(cx, I)
    pvp = load_canal_directo(cx, I)
    load_promociones(cx, I, costos, pvp)
    load_escala_excel(cx, I, costos, pvp)
    precios = load_escala_propuesta(cx, I, costos, pvp)
    load_promociones_propuesta(cx, costos, pvp)   # después de la escala: usa el precio Comercial de tiendas
    load_cadena_valor(cx, costos, pvp, precios)
    load_restaurantes_porcion(cx, costos, precios)
    load_resumen_margenes(cx)
    cf = load_costos_fijos(cx)
    load_punto_equilibrio(cx, cf)
    load_issues(cx, I, pvp)
    load_recomendaciones(cx)

    cx.executemany("INSERT INTO meta (clave, valor) VALUES (?,?)", [
        ("generado", datetime.datetime.now().isoformat(timespec="seconds")),
        ("excel_origen", str(path)),
        ("igv", str(IGV)),
        ("canal_directo_igv", "0"),
        ("porcion_gr", str(PORCION_GR)),
        ("margen_tienda_minimo", str(MARGEN_TIENDA_MINIMO)),
    ])
    cx.commit()

    for t in ("productos_costeo", "canal_directo", "promociones", "promociones_propuesta", "escala_b2b_excel", "escala_b2b_propuesta",
              "cadena_valor", "restaurantes_porcion", "costos_fijos", "punto_equilibrio", "issues", "recomendaciones"):
        n = cx.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        print(f"  {t:24s} {n:3d} filas")
    cx.close()
    print("BD generada:", DB_PATH)


if __name__ == "__main__":
    main()

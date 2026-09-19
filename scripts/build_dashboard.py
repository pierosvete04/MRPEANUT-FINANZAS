# -*- coding: utf-8 -*-
"""
Inserta dashboard/data.json dentro de la plantilla y genera dashboard/index.html
(un solo archivo, se puede abrir desde disco o enviar por WhatsApp/correo).

    python etl_build_db.py
    python export_dashboard_json.py
    python build_dashboard.py
"""
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
TEMPLATE = Path(__file__).resolve().parent / "dashboard_template.html"
DATA_JSON = BASE / "dashboard" / "data.json"
OUT = BASE / "dashboard" / "index.html"


def main():
    template = TEMPLATE.read_text(encoding="utf-8")
    data = json.loads(DATA_JSON.read_text(encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False)
    out = template.replace("/*__DATA__*/", payload)
    OUT.write_text(out, encoding="utf-8")
    print("Generado:", OUT, f"({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()

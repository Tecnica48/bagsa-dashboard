"""
sync_extensiones.py — Seguimiento de Extensiones BAGSA (versión web)
=====================================================================
Reemplaza a "actualizar.py" para la versión del dashboard que vive en la
web (GitHub Pages): en vez de pegar los datos dentro del HTML, genera
data/extensiones.json y lo sube al repo. extensiones.html ya está
preparado para leer ese JSON con fetch().

Requisitos
----------
- Este archivo, actualizar.py y github_sync.py en la MISMA carpeta.
- Completar REPO_OWNER / REPO_NAME en github_sync.py.
- Tener la variable de entorno GITHUB_TOKEN configurada (ver github_sync.py).
- pip install openpyxl   (ya lo tenés si usás actualizar.py)

Uso
---
    python sync_extensiones.py
    python sync_extensiones.py --excel "C:\\ruta\\Seguimiento Extensiones.xlsm"
    python sync_extensiones.py --sin-pausa      (para Task Scheduler)
"""

import argparse
import importlib.util
import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))


def _importar(nombre_modulo, nombre_archivo):
    ruta = os.path.join(BASE, nombre_archivo)
    spec = importlib.util.spec_from_file_location(nombre_modulo, ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


actualizar = _importar("actualizar_extensiones_core", "actualizar.py")
import github_sync  # noqa: E402  (después de fijar sys.path arriba)

EXCEL_NAME = actualizar.EXCEL_NAME
REPO_PATH_JSON = "data/extensiones.json"


def main(argv=None):
    ap = argparse.ArgumentParser(description="Publica extensiones.json en GitHub.")
    ap.add_argument("--excel", default=os.path.join(BASE, EXCEL_NAME), help="ruta del Excel")
    ap.add_argument("--sin-pausa", action="store_true", help="no esperar un Enter al terminar")
    args = ap.parse_args(argv)
    pausar = sys.stdin is not None and sys.stdin.isatty() and not args.sin_pausa

    codigo = 0
    try:
        if not os.path.exists(args.excel):
            raise actualizar.ErrorActualizar(f"No encontré el Excel en:\n  {args.excel}")

        data = actualizar.calcular(args.excel)

        json_local = os.path.join(BASE, "extensiones.json")
        with open(json_local, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

        kb = os.path.getsize(json_local) // 1024
        print(f"✓ {data['total']} extensiones · {kb} KB de datos")

        resultado = github_sync.subir_archivo(
            json_local, REPO_PATH_JSON,
            f"extensiones: {data['total']} registros al {data['generatedAt']} {data.get('generatedTime', '')}"
        )
        print(f"✓ {REPO_PATH_JSON} {resultado} en GitHub")
        print("  El dashboard mostrará estos datos la próxima vez que alguien lo abra (o recargue la página).")
    except (actualizar.ErrorActualizar, github_sync.ErrorGithubSync) as e:
        print(f"\nERROR: {e}")
        codigo = 1
    except Exception as e:
        print(f"\nERROR inesperado ({type(e).__name__}): {e}")
        codigo = 1
    if pausar:
        input("\nPresioná Enter para cerrar...")
    return codigo


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    sys.exit(main())

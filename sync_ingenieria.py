"""
sync_ingenieria.py — Visor de Seguimiento de Ingeniería (versión web)
=======================================================================
Reemplaza a "actualizar_ingenieria.py" para la versión del visor que vive
en la web (GitHub Pages): en vez de pegar los datos dentro del HTML,
genera data/ingenieria_<obra>.json y lo sube al repo. ingenieria.html ya
está preparado para leer ese JSON según el parámetro ?obra= de la URL.

Se corre UNA VEZ POR OBRA (una carpeta por obra, con su Excel
*_Seg_doc_Ingenieria.xlsm, este script y actualizar_ingenieria.py adentro).

Requisitos
----------
- Este archivo, actualizar_ingenieria.py y github_sync.py en la MISMA
  carpeta que el Excel de la obra.
- Completar REPO_OWNER / REPO_NAME en github_sync.py (una sola vez, en
  la copia de github_sync.py de cualquiera de las carpetas).
- Tener la variable de entorno GITHUB_TOKEN configurada.

Uso
---
    python sync_ingenieria.py --obra anchorena
    python sync_ingenieria.py --obra mechita --excel "C:\\ruta\\MECHITA_Seg_doc_Ingenieria.xlsm"
    python sync_ingenieria.py --obra chascomus --sin-pausa
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


actualizar_ing = _importar("actualizar_ingenieria_core", "actualizar_ingenieria.py")
import github_sync  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(description="Publica ingenieria_<obra>.json en GitHub.")
    ap.add_argument("--obra", required=True,
                     help="slug de la obra (minúsculas, sin espacios/tildes), ej: anchorena, mechita, chascomus")
    ap.add_argument("--excel", default=None, help="ruta del Excel de la obra")
    ap.add_argument("--sin-pausa", action="store_true", help="no esperar un Enter al terminar")
    args = ap.parse_args(argv)
    pausar = sys.stdin is not None and sys.stdin.isatty() and not args.sin_pausa

    obra = args.obra.strip().lower()
    repo_path_json = f"data/ingenieria_{obra}.json"

    codigo = 0
    try:
        excel_path = args.excel or actualizar_ing.encontrar_excel()
        if not excel_path or not os.path.exists(excel_path):
            raise actualizar_ing.ErrorActualizar(
                f"No encontré el Excel de la obra '{obra}'. Pasalo con --excel \"ruta\\archivo.xlsm\"."
            )

        data = actualizar_ing.calcular(excel_path)

        json_local = os.path.join(BASE, f"ingenieria_{obra}.json")
        with open(json_local, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

        kb = os.path.getsize(json_local) // 1024
        print(f"✓ obra '{obra}': {kb} KB de datos")

        resultado = github_sync.subir_archivo(
            json_local, repo_path_json, f"ingenieria/{obra}: actualización automática"
        )
        print(f"✓ {repo_path_json} {resultado} en GitHub")
        print(f"  Verlo en: ingenieria.html?obra={obra}")
    except (actualizar_ing.ErrorActualizar, github_sync.ErrorGithubSync) as e:
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

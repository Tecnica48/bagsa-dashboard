"""
github_sync.py — sube archivos a un repo de GitHub (Contents API)
==================================================================
Lo usan sync_extensiones.py y sync_ingenieria.py para publicar los JSON
(y, la primera vez, los HTML) en el repo que después sirve GitHub Pages.

CONFIGURACIÓN (una sola vez)
----------------------------
1. Generar un token en GitHub: Settings -> Developer settings ->
   Personal access tokens -> Fine-grained tokens. Permisos necesarios:
   "Contents: Read and write" sobre el repo del dashboard.

2. Guardar el token como variable de entorno, NO en este archivo:
     Windows (PowerShell), una sola vez:
         setx GITHUB_TOKEN "ghp_xxxxxxxxxxxx"
     (después de setx hay que abrir una terminal nueva para que tome el valor)

3. Completar REPO_OWNER y REPO_NAME más abajo.
"""

import base64
import json
import os
import urllib.error
import urllib.request

REPO_OWNER = "TU_USUARIO_DE_GITHUB"     # <-- completar
REPO_NAME = "bagsa-dashboard"            # <-- completar si le pusiste otro nombre
REPO_BRANCH = "main"

API_BASE = "https://api.github.com"


class ErrorGithubSync(Exception):
    pass


def _token():
    tok = os.environ.get("GITHUB_TOKEN")
    if not tok:
        raise ErrorGithubSync(
            "No encontré la variable de entorno GITHUB_TOKEN.\n"
            "Creá un token en GitHub (Settings > Developer settings > "
            "Personal access tokens) con permiso 'Contents: Read and write' "
            "sobre el repo, y guardalo como variable de entorno GITHUB_TOKEN."
        )
    return tok


def _api_request(method, path, body=None):
    url = f"{API_BASE}{path}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {_token()}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        cuerpo = e.read().decode("utf-8", errors="replace")
        if e.code == 404:
            return 404, None
        raise ErrorGithubSync(f"GitHub API {method} {path} -> {e.code}: {cuerpo}") from e
    except urllib.error.URLError as e:
        raise ErrorGithubSync(f"No se pudo conectar a GitHub ({e.reason}). Revisá la conexión a internet.") from e


def _sha_actual(repo_path):
    """SHA del archivo en el repo si ya existe, o None si no existe todavía."""
    status, data = _api_request(
        "GET", f"/repos/{REPO_OWNER}/{REPO_NAME}/contents/{repo_path}?ref={REPO_BRANCH}"
    )
    if status == 404:
        return None
    return data["sha"]


def subir_archivo(path_local, repo_path, mensaje_commit):
    """Crea o actualiza `repo_path` en el repo con el contenido de `path_local`."""
    with open(path_local, "rb") as f:
        contenido = f.read()
    sha = _sha_actual(repo_path)
    body = {
        "message": mensaje_commit,
        "content": base64.b64encode(contenido).decode("ascii"),
        "branch": REPO_BRANCH,
    }
    if sha:
        body["sha"] = sha
    _api_request("PUT", f"/repos/{REPO_OWNER}/{REPO_NAME}/contents/{repo_path}", body)
    return "actualizado" if sha else "creado"

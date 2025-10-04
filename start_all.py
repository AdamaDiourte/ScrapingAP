import os
import sys
import time
import site
import logging
import subprocess
import threading
import webbrowser
from urllib.request import urlopen
from urllib.error import URLError


PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
VENDOR_DIR = os.path.join(PROJECT_ROOT, "vendor")
FRONTEND_INDEX = os.path.join(PROJECT_ROOT, "frontend", "index.html")


def add_paths() -> None:
    if os.path.isdir(VENDOR_DIR) and VENDOR_DIR not in sys.path:
        sys.path.insert(0, VENDOR_DIR)
    try:
        user_site = site.getusersitepackages()
        if user_site and user_site not in sys.path:
            sys.path.append(user_site)
    except Exception:
        pass


def ensure_installed() -> None:
    logging.info("[deps] Vérification/installation dépendances")
    # uvicorn + fastapi en vendor pour garantir l'import
    if not os.path.isdir(VENDOR_DIR):
        os.makedirs(VENDOR_DIR, exist_ok=True)
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--target", VENDOR_DIR, "uvicorn[standard]", "fastapi"],
        cwd=PROJECT_ROOT,
    )
    # dépendances métiers via requirements.txt (si dispo)
    req = os.path.join(PROJECT_ROOT, "requirements.txt")
    if os.path.isfile(req):
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], cwd=PROJECT_ROOT)


def start_backend(host: str = "127.0.0.1", port: int = 8000) -> subprocess.Popen:
    logging.info("[backend] Démarrage FastAPI %s:%s", host, port)
    env = os.environ.copy()
    env.setdefault("HOST", host)
    env.setdefault("PORT", str(port))
    # utilise serve_backend.py qui ajoute vendor au sys.path
    return subprocess.Popen([sys.executable, os.path.join(PROJECT_ROOT, "serve_backend.py")], cwd=PROJECT_ROOT, env=env)


def wait_backend_ready(url: str, timeout_sec: int = 30) -> bool:
    logging.info("[backend] Attente disponibilité: %s", url)
    t0 = time.time()
    while time.time() - t0 < timeout_sec:
        try:
            with urlopen(url, timeout=2) as resp:
                if resp.status == 200:
                    logging.info("[backend] OK")
                    return True
        except URLError:
            pass
        except Exception:
            pass
        time.sleep(0.5)
    logging.error("[backend] Timeout d'attente")
    return False


def open_frontend(index_path: str) -> None:
    url = index_path
    if os.name == "nt" and not url.lower().startswith("file:"):
        # build file:// URL
        url = "file:///" + index_path.replace("\\", "/")

    def try_open(u: str) -> None:
        try:
            webbrowser.open_new(u)
        except Exception:
            pass
        if os.name == "nt":
            try:
                os.startfile(index_path)  # type: ignore[attr-defined]
            except Exception:
                pass
            try:
                subprocess.Popen(["cmd", "/c", "start", "", index_path], shell=True)
            except Exception:
                pass

    # plusieurs tentatives espacées
    for delay in (0.2, 1.0, 2.0):
        threading.Timer(delay, lambda u=url: try_open(u)).start()
    logging.info("[frontend] Ouverture: %s", url)


def main() -> None:
    logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"), format="%(asctime)s %(levelname)s %(message)s")
    add_paths()
    ensure_installed()
    proc = start_backend()
    ready = wait_backend_ready("http://127.0.0.1:8000/health", timeout_sec=40)
    if not ready:
        logging.warning("Le backend ne répond pas encore, tentative d'ouverture du frontend quand même")
    if os.path.isfile(FRONTEND_INDEX):
        open_frontend(FRONTEND_INDEX)
        logging.info("Interface ouverte. Si rien ne s'ouvre, ouvrez manuellement: %s", FRONTEND_INDEX)
    else:
        logging.error("Fichier frontend introuvable: %s", FRONTEND_INDEX)

    # garde le processus backend en vie
    try:
        proc.wait()
    except KeyboardInterrupt:
        try:
            proc.terminate()
        except Exception:
            pass


if __name__ == "__main__":
    main()







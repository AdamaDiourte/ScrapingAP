import os
import sys
import site
import logging


def ensure_paths() -> None:
    project_root = os.path.dirname(os.path.abspath(__file__))
    vendor = os.path.join(project_root, "vendor")
    if os.path.isdir(vendor) and vendor not in sys.path:
        sys.path.insert(0, vendor)
    try:
        user_site = site.getusersitepackages()
        if user_site and user_site not in sys.path:
            sys.path.append(user_site)
    except Exception:
        pass


def main() -> None:
    logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"))
    ensure_paths()
    # Charge les variables depuis un fichier .env si présent
    try:
        from dotenv import load_dotenv  # type: ignore
        load_dotenv()
    except Exception:
        pass
    import uvicorn  # type: ignore

    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    # active le rechargement automatique pour prendre en compte les corrections
    uvicorn.run("backend.main:app", host=host, port=port, reload=True)


if __name__ == "__main__":
    main()



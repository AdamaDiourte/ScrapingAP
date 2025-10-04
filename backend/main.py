import logging
import asyncio
import os
import tempfile
import shutil
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from .finder import AppelsProjetFinder


LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO")
logging.basicConfig(
    level=LOG_LEVEL, format="%(asctime)s %(levelname)s %(name)s: %(message)s"
)
logger = logging.getLogger("backend")

app = FastAPI(title="Appels à Projet API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # conservé, mais regex ci-dessous couvre file:// (Origin: null)
    allow_origin_regex=".*",
    allow_credentials=False,  # pas de cookies; permet Access-Control-Allow-Origin: *
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=[
        "Content-Disposition",
        "Content-Type",
        "Content-Length",
        "X-AI-Provider",
        "X-AI-Calls",
        "X-AI-Success",
        "X-AI-Heuristic-Used",
        "X-AI-Errors",
    ],
    max_age=600,
)


_heartbeat_task: asyncio.Task | None = None


@app.on_event("startup")
async def _on_startup() -> None:
    global _heartbeat_task
    async def _heartbeat() -> None:
        while True:
            logger.info("Heartbeat: backend alive")
            await asyncio.sleep(30)
    try:
        _heartbeat_task = asyncio.create_task(_heartbeat())
    except Exception:
        logger.exception("Impossible de démarrer le heartbeat")


@app.on_event("shutdown")
async def _on_shutdown() -> None:
    global _heartbeat_task
    if _heartbeat_task is not None:
        try:
            _heartbeat_task.cancel()
        except Exception:
            pass

@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/process")
async def process(
    api_key: str = Form("") ,
    api_provider: str = Form("openai"),
    excel_file: UploadFile = File(...),
):
    logger.info("/process reçu provider=%s file=%s content_type=%s", api_provider, excel_file.filename, getattr(excel_file, "content_type", "?"))

    try:
        # Crée un répertoire temporaire persistant (sera supprimé après l'envoi)
        tmpdir = tempfile.mkdtemp(prefix="ap_ui_")
        in_path = Path(tmpdir) / Path(excel_file.filename).name
        out_path = Path(tmpdir) / "resultats_appels_projet.docx"

        content = await excel_file.read()
        logger.info("Lecture upload: %d octets", len(content))
        in_path.write_bytes(content)
        logger.info("Fichier Excel sauvegardé: %s (%d octets)", in_path, len(content))

        # Si une clé n'est pas fournie par l'UI, tente depuis l'environnement (.env)
        effective_key = (api_key or os.environ.get("AP_FINDER_API_KEY") or os.environ.get("OPENAI_API_KEY") or os.environ.get("ANTHROPIC_API_KEY") or "").strip()
        effective_provider = api_provider

        # Si toujours vide, essaye de lire le fichier .env directement (parfois non chargé selon l'environnement)
        if not effective_key:
            try:
                from dotenv import find_dotenv, dotenv_values  # type: ignore
                env_path = find_dotenv(usecwd=True) or ".env"
                values = dotenv_values(env_path)
                effective_key = (values.get("AP_FINDER_API_KEY") or values.get("OPENAI_API_KEY") or values.get("ANTHROPIC_API_KEY") or "").strip()  # type: ignore[attr-defined]
            except Exception:
                pass

        # Détecte OpenRouter si clé sk-or- fournie via .env
        if not effective_provider and effective_key.lower().startswith("sk-or-"):
            effective_provider = "openrouter"

        logger.info("Initialisation finder(api_provider=%s, key=%s)", effective_provider or "openai", "****" if effective_key else "(vide)")
        finder = AppelsProjetFinder(api_key=effective_key, api_provider=effective_provider or "openai")
        logger.info("Traitement du fichier Excel...")
        finder.traiter_fichier(str(in_path))
        logger.info("Génération du document Word...")
        finder.generer_document_word(str(out_path))

        # Log de synthèse sur l'utilisation de l'IA
        try:
            ai_info = {
                "provider": finder.api_provider,
                "ai_calls": getattr(finder, "ai_calls", None),
                "ai_success": getattr(finder, "ai_success", None),
                "heuristic_used": getattr(finder, "heuristic_used", None),
                "ai_errors": getattr(finder, "ai_errors", None),
            }
            logger.info("Synthèse IA: %s", ai_info)
        except Exception:
            pass

        if not out_path.exists():
            logger.error("Fichier Word non généré")
            # Nettoie si génération échoue
            shutil.rmtree(tmpdir, ignore_errors=True)
            return JSONResponse(
                status_code=500, content={"error": "generation_failed"}
            )

        # Lit le document en mémoire pour éviter tout problème de fichier éphémère
        try:
            data = out_path.read_bytes()
            logger.info("DOCX prêt: %s octets", len(data))
        except Exception as read_exc:
            logger.exception("Lecture DOCX échouée: %s", read_exc)
            shutil.rmtree(tmpdir, ignore_errors=True)
            return JSONResponse(status_code=500, content={"error": "read_failed"})

        # Supprime le dossier temporaire après que la réponse ait été envoyée
        background = BackgroundTask(lambda: shutil.rmtree(tmpdir, ignore_errors=True))
        headers = {
            "Content-Disposition": "attachment; filename=\"resultats_appels_projet.docx\"",
            "Cache-Control": "no-store",
        }
        # Ajoute des en-têtes de diagnostic IA
        try:
            headers["X-AI-Provider"] = str(getattr(finder, "api_provider", ""))
            headers["X-AI-Calls"] = str(getattr(finder, "ai_calls", ""))
            headers["X-AI-Success"] = str(getattr(finder, "ai_success", ""))
            headers["X-AI-Heuristic-Used"] = str(getattr(finder, "heuristic_used", ""))
            errors_val = getattr(finder, "ai_errors", None)
            if errors_val:
                headers["X-AI-Errors"] = "; ".join(errors_val)[:900]
        except Exception:
            pass
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers,
            background=background,
        )
    except Exception as exc:
        logger.exception("Erreur /process: %s", exc)
        return JSONResponse(status_code=500, content={"error": str(exc)})



# Sert l'interface frontend en HTTP local (évite file://)
app.mount(
    "/",
    StaticFiles(directory="frontend", html=True),
    name="frontend",
)


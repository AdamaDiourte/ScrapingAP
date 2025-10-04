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
    expose_headers=["Content-Disposition", "Content-Type", "Content-Length"],
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
    api_key: str = Form(...),
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

        logger.info("Initialisation finder(api_provider=%s)", api_provider)
        finder = AppelsProjetFinder(api_key=api_key, api_provider=api_provider)
        logger.info("Traitement du fichier Excel...")
        finder.traiter_fichier(str(in_path))
        logger.info("Génération du document Word...")
        finder.generer_document_word(str(out_path))

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


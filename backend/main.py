import logging
import asyncio
import os
import tempfile
import shutil
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask
from starlette.responses import StreamingResponse
try:
    from watchfiles import awatch  # type: ignore
except Exception:  # pragma: no cover
    awatch = None  # type: ignore

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
_reload_watch_task: asyncio.Task | None = None
_reload_version: int = 0
_last_ai_info: dict[str, str] | None = None


@app.on_event("startup")
async def _on_startup() -> None:
    global _heartbeat_task, _reload_watch_task, _reload_version
    async def _heartbeat() -> None:
        while True:
            logger.info("Heartbeat: backend alive")
            await asyncio.sleep(30)
    try:
        _heartbeat_task = asyncio.create_task(_heartbeat())
    except Exception:
        logger.exception("Impossible de démarrer le heartbeat")

    # Démarre la surveillance des fichiers frontend pour Live Reload (dev)
    if awatch is not None:
        async def _watch_frontend() -> None:
            nonlocal_logger = logging.getLogger("frontend-watch")
            try:
                global _reload_version
                async for _ in awatch("frontend"):
                    _reload_version += 1
                    nonlocal_logger.info("Frontend changé → version=%s", _reload_version)
            except asyncio.CancelledError:
                pass
            except Exception:
                nonlocal_logger.exception("Watch frontend échoué")

        try:
            _reload_watch_task = asyncio.create_task(_watch_frontend())
        except Exception:
            logger.exception("Impossible de démarrer le watcher frontend")


@app.on_event("shutdown")
async def _on_shutdown() -> None:
    global _heartbeat_task, _reload_watch_task
    if _heartbeat_task is not None:
        try:
            _heartbeat_task.cancel()
        except Exception:
            pass
    if _reload_watch_task is not None:
        try:
            _reload_watch_task.cancel()
        except Exception:
            pass

@app.get("/health")
def health() -> JSONResponse:
    headers: dict[str, str] = {}
    content: dict[str, object] = {"status": "ok"}
    try:
        if _last_ai_info:
            content["ai"] = _last_ai_info
            if _last_ai_info.get("provider"):
                headers["X-AI-Provider"] = _last_ai_info.get("provider", "")
            if _last_ai_info.get("ai_calls"):
                headers["X-AI-Calls"] = _last_ai_info.get("ai_calls", "")
            if _last_ai_info.get("ai_success"):
                headers["X-AI-Success"] = _last_ai_info.get("ai_success", "")
            if _last_ai_info.get("heuristic_used"):
                headers["X-AI-Heuristic-Used"] = _last_ai_info.get("heuristic_used", "")
            if _last_ai_info.get("ai_errors"):
                headers["X-AI-Errors"] = _last_ai_info.get("ai_errors", "")
    except Exception:
        pass
    return JSONResponse(content=content, headers=headers)


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
        effective_provider = (api_provider or "").strip()
        env = os.environ
        effective_key = api_key.strip() if api_key else ""
        key_source = "ui" if effective_key else ""

        # Priorité: clé spécifique au fournisseur sélectionné (env)
        if not effective_key:
            prov = effective_provider.lower()
            if prov == "openrouter":
                effective_key = (env.get("OPENROUTER_API_KEY") or "").strip()
                key_source = "env:OPENROUTER_API_KEY" if effective_key else key_source
            elif prov == "openai":
                effective_key = (env.get("OPENAI_API_KEY") or "").strip()
                key_source = "env:OPENAI_API_KEY" if effective_key else key_source
            elif prov == "anthropic":
                effective_key = (env.get("ANTHROPIC_API_KEY") or "").strip()
                key_source = "env:ANTHROPIC_API_KEY" if effective_key else key_source

        # Lecture .env directe si toujours vide (provider-spécifique)
        if not effective_key:
            try:
                from dotenv import find_dotenv, dotenv_values  # type: ignore
                env_path = find_dotenv(usecwd=True) or ".env"
                values = dotenv_values(env_path)
                prov = effective_provider.lower()
                if prov == "openrouter":
                    effective_key = (values.get("OPENROUTER_API_KEY") or "").strip()  # type: ignore[attr-defined]
                    key_source = "dotenv:OPENROUTER_API_KEY" if effective_key else key_source
                elif prov == "openai":
                    effective_key = (values.get("OPENAI_API_KEY") or "").strip()  # type: ignore[attr-defined]
                    key_source = "dotenv:OPENAI_API_KEY" if effective_key else key_source
                elif prov == "anthropic":
                    effective_key = (values.get("ANTHROPIC_API_KEY") or "").strip()  # type: ignore[attr-defined]
                    key_source = "dotenv:ANTHROPIC_API_KEY" if effective_key else key_source
                # Fallback générique AP_FINDER_API_KEY si rien (utilisé aussi pour auto-détection)
                if not effective_key:
                    generic = (values.get("AP_FINDER_API_KEY") or "").strip()  # type: ignore[attr-defined]
                    if generic:
                        effective_key = generic
                        key_source = "dotenv:AP_FINDER_API_KEY"
            except Exception:
                pass

        # Auto-détection du provider si la clé vient de l'environnement/.env (UI vide)
        if key_source and key_source != "ui":
            k = effective_key.lower()
            if k.startswith("sk-or-"):
                effective_provider = "openrouter"
            elif k.startswith("sk-ant-"):
                effective_provider = "anthropic"
            elif not effective_provider:
                effective_provider = "openai"
        
        logger.info("Clé IA chargée via=%s provider=%s", key_source or "(none)", (effective_provider or ""))

        # Si toujours vide, essaye de lire le fichier .env directement (parfois non chargé selon l'environnement)
        if not effective_key:
            try:
                from dotenv import find_dotenv, dotenv_values  # type: ignore
                env_path = find_dotenv(usecwd=True) or ".env"
                values = dotenv_values(env_path)
                # Priorité au provider courant
                provider_lower = (effective_provider or "").lower()
                if provider_lower == "openrouter":
                    effective_key = (values.get("OPENROUTER_API_KEY") or "").strip()  # type: ignore[attr-defined]
                elif provider_lower == "openai":
                    effective_key = (values.get("OPENAI_API_KEY") or "").strip()  # type: ignore[attr-defined]
                elif provider_lower == "anthropic":
                    effective_key = (values.get("ANTHROPIC_API_KEY") or "").strip()  # type: ignore[attr-defined]
                # Fallbacks
                if not effective_key:
                    effective_key = (
                        values.get("AP_FINDER_API_KEY")
                        or values.get("OPENROUTER_API_KEY")
                        or values.get("OPENAI_API_KEY")
                        or values.get("ANTHROPIC_API_KEY")
                        or ""
                    ).strip()  # type: ignore[attr-defined]
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
            ai_info_raw = {
                "provider": getattr(finder, "api_provider", None),
                "ai_calls": getattr(finder, "ai_calls", None),
                "ai_success": getattr(finder, "ai_success", None),
                "heuristic_used": getattr(finder, "heuristic_used", None),
                "ai_errors": getattr(finder, "ai_errors", None),
            }
            logger.info("Synthèse IA: %s", ai_info_raw)
            # Normalise et mémorise pour /health et diagnostics
            errors_val = ai_info_raw.get("ai_errors")
            if isinstance(errors_val, (list, tuple)):
                errors_str = "; ".join(str(x) for x in errors_val)
            else:
                errors_str = str(errors_val or "")
            global _last_ai_info
            _last_ai_info = {
                "provider": str(ai_info_raw.get("provider") or ""),
                "ai_calls": str(ai_info_raw.get("ai_calls") or ""),
                "ai_success": str(ai_info_raw.get("ai_success") or ""),
                "heuristic_used": str(ai_info_raw.get("heuristic_used") or ""),
                "ai_errors": errors_str,
                "key_source": key_source or "",
            }
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
            # Source de la clé si connue
            try:
                # réutilise la dernière info si dispo
                if _last_ai_info:
                    src = _last_ai_info.get("key_source") or ""
                    if src:
                        headers["X-AI-Key-Source"] = src
            except Exception:
                pass
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


# En-têtes no-cache pour HTML/CSS/JS afin d'éviter le cache navigateur en dev
@app.middleware("http")
async def _no_cache_static(request: Request, call_next):  # type: ignore[override]
    response = await call_next(request)
    try:
        ctype = response.headers.get("content-type", "")
        if "text/html" in ctype or "javascript" in ctype or "text/css" in ctype:
            response.headers["Cache-Control"] = "no-store"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
    except Exception:
        pass
    return response


# Endpoint SSE pour Live Reload côté frontend
@app.get("/dev/reload")
async def dev_reload(request: Request) -> StreamingResponse:
    async def event_stream():
        last_sent = -1
        # ping initial pour établir la connexion
        yield ": connected\n\n"
        while True:
            if await request.is_disconnected():
                break
            try:
                if _reload_version != last_sent:
                    last_sent = _reload_version
                    yield f"data: {last_sent}\n\n"
                await asyncio.sleep(0.5)
            except asyncio.CancelledError:
                break
            except Exception:
                await asyncio.sleep(1.0)

    return StreamingResponse(event_stream(), media_type="text/event-stream", headers={"Cache-Control": "no-store"})


from pathlib import Path
import shutil
import uuid
import traceback
import logging

from fastapi import (
    FastAPI,
    UploadFile,
    File,
    HTTPException,
    BackgroundTasks,
)

from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware

from payment_processor import PaymentProcessor

# ==========================================================
# Logging
# ==========================================================

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

logger.info("APP.PY VERSION 2026-07-22")

# ==========================================================
# Paths
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "output"
STATIC_FOLDER = BASE_DIR / "static"
TEMPLATE_FOLDER = BASE_DIR / "templates"

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

# ==========================================================
# Upload limits
# ==========================================================

MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB — adjust if your reports run larger

# ==========================================================
# FastAPI
# ==========================================================

app = FastAPI(
    title="Myntra Payment Register",
    version="1.0.0"
)


class LogMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        logger.info(f"REQUEST START : {request.method} {request.url.path}")

        response = await call_next(request)

        logger.info(
            f"REQUEST END   : {request.method} {request.url.path} -> {response.status_code}"
        )

        return response


app.add_middleware(LogMiddleware)

# ==========================================================
# Static
# ==========================================================

app.mount(
    "/static",
    StaticFiles(directory=STATIC_FOLDER),
    name="static",
)

# ==========================================================
# CORS
# ==========================================================
# NOTE: wildcard origins ("*") combined with allow_credentials=True
# is an invalid combination — browsers reject it, and some frameworks
# silently echo back the request Origin instead, defeating the point
# of the wildcard. Since this endpoint doesn't rely on cookies/auth
# headers, credentials are disabled here to keep the wildcard honest.
# If you later need cookie-based auth from a browser frontend, replace
# allow_origins=["*"] with your specific frontend origin(s) instead.

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================================================
# Helpers
# ==========================================================


def delete_file(path: Path):
    try:
        if path.exists():
            path.unlink()
            logger.info(f"Deleted : {path}")
    except Exception as e:
        logger.error(f"Delete failed : {e}")


# ==========================================================
# Home
# ==========================================================


@app.get("/", include_in_schema=False)
async def home():
    return FileResponse(str(TEMPLATE_FOLDER / "index.html"))


# ==========================================================
# Health
# ==========================================================


@app.get("/health")
async def health():
    return {"status": "Healthy"}


# ==========================================================
# Upload
# ==========================================================


@app.post("/upload")
async def upload_excel(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
):

    logger.info("Upload route reached")

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file selected.",
        )

    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(
            status_code=400,
            detail="Only .xlsx files are supported.",
        )

    unique_id = str(uuid.uuid4())

    uploaded_file = UPLOAD_FOLDER / f"{unique_id}.xlsx"
    output_file = OUTPUT_FOLDER / f"Payment_Register_{unique_id}.xlsx"

    try:

        logger.info("STEP 1 - Saving uploaded file")

        # Stream to disk with a hard size cap instead of trusting
        # shutil.copyfileobj to stop on its own — protects against
        # a very large upload filling disk before we notice.
        bytes_written = 0
        with uploaded_file.open("wb") as buffer:
            while chunk := await file.read(1024 * 1024):
                bytes_written += len(chunk)
                if bytes_written > MAX_UPLOAD_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail=f"File exceeds the {MAX_UPLOAD_BYTES // (1024*1024)} MB upload limit.",
                    )
                buffer.write(chunk)

        file.file.close()

        logger.info(f"STEP 2 - File saved ({bytes_written} bytes)")

        processor = PaymentProcessor(uploaded_file)

        logger.info("STEP 3 - Processor initialized")

        processor.save_excel(output_file)

        logger.info("STEP 4 - Excel generated")

        if not output_file.exists():
            raise Exception("Output file was not created.")

        if output_file.stat().st_size == 0:
            raise Exception("Generated file is empty.")

        logger.info(
            f"Output File : {output_file} ({round(output_file.stat().st_size/1024,2)} KB)"
        )

        background_tasks.add_task(delete_file, uploaded_file)
        background_tasks.add_task(delete_file, output_file)

        logger.info("STEP 5 - Returning FileResponse")

        return FileResponse(
            path=str(output_file),
            filename="Payment_Register.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=background_tasks,
        )

    except HTTPException:
        delete_file(uploaded_file)
        delete_file(output_file)
        raise

    except Exception as e:

        logger.error(f"UPLOAD FAILED — {type(e).__name__}: {e}")

        logger.exception("UPLOAD FAILED")

        delete_file(uploaded_file)
        delete_file(output_file)

        raise HTTPException(
            status_code=500,
            detail=str(e),
        )
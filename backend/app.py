from pathlib import Path
import shutil
import uuid
import traceback

from fastapi import (
    FastAPI,
    UploadFile,
    File,
    HTTPException,
    BackgroundTasks
)

from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from payment_processor import PaymentProcessor

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
# FastAPI
# ==========================================================

app = FastAPI(
    title="Myntra Payment Register",
    version="1.0.0"
)

# ==========================================================
# Static Files
# ==========================================================

app.mount(
    "/static",
    StaticFiles(directory=STATIC_FOLDER),
    name="static"
)

# ==========================================================
# CORS
# ==========================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================================================
# Helper
# ==========================================================

def delete_file(path: Path):
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass

# ==========================================================
# Home
# ==========================================================

@app.get("/", include_in_schema=False)
async def home():
    return FileResponse(TEMPLATE_FOLDER / "index.html")

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
#async def upload_excel(
    #background_tasks: BackgroundTasks,
    #file: UploadFile = File(...)
#):
async def upload_excel(file: UploadFile = File(...)):

    print("Upload route reached")

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file selected."
        )

    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(
            status_code=400,
            detail="Only .xlsx files are supported."
        )

    unique_id = str(uuid.uuid4())

    uploaded_file = UPLOAD_FOLDER / f"{unique_id}.xlsx"

    output_file = OUTPUT_FOLDER / f"Payment_Register_{unique_id}.xlsx"

    try:

        print("Step 1: File received")
        with uploaded_file.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        file.file.close()

        print("Step 2: File saved")

        processor = PaymentProcessor(uploaded_file)

        print("Step 3: Processor created")

        processor.save_excel(output_file)

        print("Step 4: Excel generated")

        if not output_file.exists():
            raise Exception("Output file not found")
        
        if output_file.stat().st_size == 0:
            raise Exception("Generated Excel file is empty.")

        print("Step 5: Returning file")

        background_tasks.add_task(delete_file, uploaded_file)
        background_tasks.add_task(delete_file, output_file)

        print("Output File :", output_file)
        print("Exists      :", output_file.exists())
        print("Size (KB)   :", round(output_file.stat().st_size / 1024, 2))

        return FileResponse(
            path=output_file,
            filename="Payment_Register.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=background_tasks
        )

    except Exception as e:

        print("\n" + "=" * 80)
        print("UPLOAD ERROR")
        print("Exception Type :", type(e).__name__)
        print("Exception      :", str(e))
        traceback.print_exc()
        print("=" * 80 + "\n")
        delete_file(uploaded_file)
        delete_file(output_file)

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )
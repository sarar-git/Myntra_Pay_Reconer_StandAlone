from pathlib import Path
import shutil
import uuid

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

from payment_processor import PaymentProcessor

# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent

UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "output"

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

# --------------------------------------------------

app = FastAPI(
    title="Myntra Payment Register API",
    version="1.0.0"
)

# --------------------------------------------------

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --------------------------------------------------

@app.get("/")
def home():

    return {
        "application": "Myntra Payment Register",
        "version": "1.0.0",
        "status": "Running"
    }

# --------------------------------------------------

@app.get("/health")
def health():

    return {
        "status": "Healthy"
    }

# --------------------------------------------------

@app.post("/upload")
async def upload_excel(
    file: UploadFile = File(...)
):

    if not file.filename.lower().endswith(".xlsx"):

        raise HTTPException(
            status_code=400,
            detail="Only Excel (.xlsx) files are allowed."
        )

    unique_id = str(uuid.uuid4())

    uploaded_file = (
        UPLOAD_FOLDER /
        f"{unique_id}.xlsx"
    )

    output_file = (
        OUTPUT_FOLDER /
        f"Payment_Register_{unique_id}.xlsx"
    )

    with uploaded_file.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    processor = PaymentProcessor(uploaded_file)

    processor.save_excel(output_file)

    return FileResponse(
        output_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Payment_Register.xlsx"
    )
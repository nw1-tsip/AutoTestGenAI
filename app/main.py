from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import uuid
import os
import shutil

from core.generator import generate_testcases

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JOB_DIR = os.path.join(BASE_DIR, "jobs")
TEMPLATE_PATH = "templates/TestCases_Template.xlsx"

os.makedirs(JOB_DIR, exist_ok=True)


@app.post("/generate-testcases")
async def generate(srs: UploadFile = File(...)):

    job_id = str(uuid.uuid4())
    job_path = os.path.join(JOB_DIR, job_id)
    os.makedirs(job_path, exist_ok=True)

    srs_path = os.path.join(job_path, srs.filename)
    output_path = os.path.join(job_path, "Generated_TestCases.xlsx")

    # Save uploaded SRS
    with open(srs_path, "wb") as f:
        shutil.copyfileobj(srs.file, f)

    # Run your GitHub code
    generate_testcases(
        srs_path=srs_path,
        template_path=TEMPLATE_PATH,
        output_path=output_path
    )

    return FileResponse(
        output_path,
        filename="Generated_TestCases.xlsx"
    )


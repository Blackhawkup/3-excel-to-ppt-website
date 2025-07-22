from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
import os
import uuid
import shutil
import tempfile

# Import your model logic
from model1_monthly.excel_to_ppt import generate_ppt as generate_monthly_ppt
from model2_ria.ria_automation import generate_ppt as generate_ria_ppt
from ready_reckoner import run_pms, run_hybrid   # <-- ensures this file is in your project

app = FastAPI()

# Serve your frontend (HTML) files
app.mount("/frontend", StaticFiles(directory="frontend", html=True), name="frontend")


@app.get("/")
async def root():
    return RedirectResponse(url="/frontend/index.html")


########################
# Endpoint 1 (/process-ppt/)
# RIA Automation endpoint
@app.post("/process-ppt/")
async def process_files(
    background_tasks: BackgroundTasks,
    ppt_file: UploadFile = File(...),
    excel_file1: UploadFile = File(...),
    excel_file2: UploadFile = File(...),
    excel_file3: UploadFile = File(...),
    owner_no: str = Form(...)
):
    temp_dir = "temp_ppt"
    os.makedirs(temp_dir, exist_ok=True)

    def save_uploaded(file: UploadFile):
        fname = f"{uuid.uuid4().hex}_{file.filename}"
        fpath = os.path.join(temp_dir, fname)
        with open(fpath, "wb") as f:
            shutil.copyfileobj(file.file, f)
        return fpath

    ppt_path = save_uploaded(ppt_file)
    excel1_path = save_uploaded(excel_file1)
    excel2_path = save_uploaded(excel_file2)
    excel3_path = save_uploaded(excel_file3)
    output_path = os.path.join(temp_dir, f"{uuid.uuid4().hex}_output.pptx")

    generate_ria_ppt(
        ppt_template_path=ppt_path,
        excel1_path=excel1_path,
        excel2_path=excel2_path,
        excel3_path=excel3_path,
        owner_no=owner_no,
        output_path=output_path
    )

    def cleanup():
        for p in [ppt_path, excel1_path, excel2_path, excel3_path, output_path]:
            try:
                os.remove(p)
            except Exception:
                pass

    background_tasks.add_task(cleanup)

    return FileResponse(
        path=output_path,
        filename="modified.pptx",
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        background=background_tasks
    )


########################
# Endpoint 2 (/monthly/)
# Model1 Monthly endpoint
@app.post("/monthly/")
async def generate_monthly_ppt_endpoint(
    file: UploadFile = File(...),
    background_tasks: BackgroundTasks = BackgroundTasks()
):
    unique_id = uuid.uuid4().hex
    input_filename = f"{unique_id}_data.xlsx"
    input_path = os.path.join("model1_monthly", input_filename)

    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    output_path = generate_monthly_ppt(input_filename)

    if isinstance(output_path, dict) and "error" in output_path:
        return output_path

    background_tasks.add_task(os.remove, input_path)
    background_tasks.add_task(os.remove, output_path)

    return FileResponse(
        path=output_path,
        filename="generated.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        background=background_tasks
    )


########################
# Endpoint 3 (/generate_pptx/)
# ReadyReckoner (PMS/Hybrid) endpoint
@app.post("/generate_pptx/")
async def generate_pptx(
    background_tasks: BackgroundTasks,
    process_type: str = Form(...),
    n_pms: int = Form(1),
    n_hybrid: int = Form(1),
    excel_file: UploadFile = File(...),
    pms_template: UploadFile = File(...),
    hybrid_template: UploadFile = File(...)
):
    tmp_dir = tempfile.mkdtemp()
    try:
        excel_path = os.path.join(tmp_dir, excel_file.filename)
        pms_template_path = os.path.join(tmp_dir, pms_template.filename)
        hybrid_template_path = os.path.join(tmp_dir, hybrid_template.filename)

        with open(excel_path, "wb") as f:
            f.write(await excel_file.read())
        with open(pms_template_path, "wb") as f:
            f.write(await pms_template.read())
        with open(hybrid_template_path, "wb") as f:
            f.write(await hybrid_template.read())

        if process_type == "PMS":
            output_file = os.path.join(tmp_dir, "Client_Associates_All_PMS_Funds.pptx")
            run_pms(excel_path, pms_template_path, n_pms, output_file)
            out_name = "Client_Associates_All_PMS_Funds.pptx"
        else:
            output_file = os.path.join(tmp_dir, "Client_Associates_All_Hybrid_Funds.pptx")
            run_hybrid(excel_path, hybrid_template_path, n_hybrid, output_file)
            out_name = "Client_Associates_All_Hybrid_Funds.pptx"

        background_tasks.add_task(shutil.rmtree, tmp_dir, ignore_errors=True)

        return FileResponse(
            output_file,
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise e


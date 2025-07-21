from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks
from fastapi.responses import FileResponse
import os, uuid, shutil
from model2_ria.ria_automation import generate_ppt

app = FastAPI()

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

    generate_ppt(
        ppt_template_path=ppt_path,
        excel1_path=excel1_path,
        excel2_path=excel2_path,
        excel3_path=excel3_path,
        owner_no=owner_no,
        output_path=output_path
    )

    def cleanup():
        os.remove(ppt_path)
        os.remove(excel1_path)
        os.remove(excel2_path)
        os.remove(excel3_path)
        os.remove(output_path)

    background_tasks.add_task(cleanup)

    return FileResponse(
        path=output_path,
        filename="modified.pptx",
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

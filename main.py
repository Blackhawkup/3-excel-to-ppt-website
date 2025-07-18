from fastapi import FastAPI, UploadFile, File, BackgroundTasks
from fastapi.responses import FileResponse
import shutil
import os
import uuid
from model1_monthly.excel_to_ppt import generate_ppt

app = FastAPI()

@app.post("/monthly/")
async def generate_monthly_ppt(file: UploadFile = File(...), background_tasks: BackgroundTasks = BackgroundTasks()):
    # Create a unique filename
    unique_id = uuid.uuid4().hex
    input_filename = f"{unique_id}_data.xlsx"
    input_path = os.path.join("model1_monthly", input_filename)

    # Save uploaded file
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # Generate PPT
    output_path = generate_ppt(input_filename)

    if isinstance(output_path, dict) and "error" in output_path:
        return output_path

    # Schedule deletion of both files
    background_tasks.add_task(os.remove, input_path)
    background_tasks.add_task(os.remove, output_path)

    # Return the file as response
    return FileResponse(
        path=output_path,
        filename="generated.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        background=background_tasks
    )

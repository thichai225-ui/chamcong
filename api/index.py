from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from utils.excel_handler import process_file
import os

app = FastAPI()

@app.get("/", response_class=HTMLResponse)
async def home():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/api/process")
async def process_excel(file: UploadFile = File(...)):
    input_path = f"/tmp/{file.filename}"
    output_path = f"/tmp/processed_{file.filename}"

    # Lưu file upload vào /tmp
    with open(input_path, "wb") as f:
        f.write(await file.read())

    # Gọi hàm xử lý Excel
    process_file(input_path, output_path)

    # Trả về file kết quả
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"processed_{file.filename}",
    )

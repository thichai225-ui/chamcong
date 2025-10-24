from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from utils.excel_handler import process_file
import os, tempfile

app = FastAPI()

@app.get("/", response_class=HTMLResponse)
async def home():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/api/process")
async def process_excel(file: UploadFile = File(...)):
    # Thư mục tạm hợp lệ cho cả Windows & Vercel
    tmp_dir = tempfile.gettempdir()
    input_path = os.path.join(tmp_dir, file.filename)
    output_path = os.path.join(tmp_dir, f"processed_{file.filename}")

    # Ghi file upload vào ổ tạm
    with open(input_path, "wb") as f:
        f.write(await file.read())

    # Xử lý Excel (hàm bạn đã có sẵn)
    process_file(input_path, output_path)

    # Trả về file đã xử lý cho client
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"processed_{file.filename}",
    )

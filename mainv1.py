from pydantic import BaseModel
import requests
import time
import os
from fastapi import FastAPI, UploadFile, File, Query, BackgroundTasks
from pathlib import Path
from typing import List
import shutil
import uuid
from convert import convert_office_folder_to_pdf, merge_pdfs, extract_text_from_folder
import zipfile
import tempfile
from fastapi.responses import FileResponse

from convert import (
    convert_office_folder_to_pdf,
    merge_pdfs,
    extract_text_from_folder
)

# Bạn set API key ở biến môi trường hoặc gán trực tiếp ở đây
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "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")

BASE_URL = "https://api.openai.com/v1"

HEADERS = {
    "Authorization": f"Bearer {OPENAI_API_KEY}",
    "Content-Type": "application/json",
    "OpenAI-Beta": "assistants=v2"
}


app = FastAPI(title="PPTX to PDF API")

OUTPUT_DIR = Path(__file__).parent / "pptx-to-pdf"
TEMP_UPLOAD_DIR = Path(__file__).parent / "temp_uploads"
TEMP_UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


class Message(BaseModel):
    role: str
    content: str

class ChatRequest(BaseModel):
    assistant_id: str
    messages: List[Message]

@app.post("/chat/completions")
def chat_with_assistant(req: ChatRequest):
    # 1. Tạo thread
    res = requests.post(f"{BASE_URL}/threads", headers=HEADERS, json={})
    res.raise_for_status()
    thread_id = res.json()["id"]

    # 2. Thêm tất cả messages vào thread
    for msg in req.messages:
        payload_msg = {
            "role": msg.role,
            "content": msg.content
        }
        r = requests.post(f"{BASE_URL}/threads/{thread_id}/messages", headers=HEADERS, json=payload_msg)
        r.raise_for_status()

    # 3. Tạo run
    payload_run = {
        "assistant_id": req.assistant_id
    }
    res = requests.post(f"{BASE_URL}/threads/{thread_id}/runs", headers=HEADERS, json=payload_run)
    res.raise_for_status()
    run_id = res.json()["id"]

    # 4. Poll cho đến khi hoàn tất
    while True:
        res = requests.get(f"{BASE_URL}/threads/{thread_id}/runs/{run_id}", headers=HEADERS)
        res.raise_for_status()
        status = res.json()["status"]
        if status == "completed":
            break
        elif status in ["failed", "cancelled", "expired"]:
            return {"error": f"Run ended with status: {status}"}
        time.sleep(1)

    # 5. Lấy câu trả lời từ assistant
    res = requests.get(f"{BASE_URL}/threads/{thread_id}/messages", headers=HEADERS)
    res.raise_for_status()
    data = res.json()["data"]

    for msg in data:
        if msg["role"] == "assistant":
            return {
                "assistant_reply": msg["content"][0]["text"]["value"]
            }

    return {"error": "No assistant reply found"}



@app.post("/convert-folder")
def convert_folder_api(
    folder_path: str = Query(..., description="Đường dẫn thư mục chứa file Office")
):
    output_dir = Path(__file__).parent / "office-to-pdf"
    try:
        pdf_files = convert_office_folder_to_pdf(folder_path, output_dir)
        return {
            "message": f"Đã convert {len(pdf_files)} file.",
            "output_folder": str(output_dir),
            "pdf_files": [str(p) for p in pdf_files]
        }
    except Exception as e:
        return {"error": str(e)}

@app.post("/merge-pdf")
def merge_pdf_api(
    folder_path: str = Query(..., description="Thư mục chứa file PDF"),
    merged_name: str = Query("merged.pdf", description="Tên file PDF gộp")
):
    try:
        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            return {"error": f"'{folder_path}' không tồn tại hoặc không phải thư mục."}

        pdf_list = sorted(folder.glob("*.pdf"))
        if not pdf_list:
            return {"message": "Không tìm thấy file PDF nào để ghép."}

        # Ghi vào thư mục 'output' cùng cấp với main.py
        output_dir = Path(__file__).parent / "output"
        output_path = output_dir / merged_name

        merged_pdf = merge_pdfs(pdf_list, output_path)

        return {
            "message": f"Gộp {len(pdf_list)} file PDF thành công.",
            "merged_pdf": str(merged_pdf)
        }
    except Exception as e:
        return {"error": str(e)}


@app.post("/convert-and-merge")
def convert_and_merge_api(
    folder_path: str = Query(..., description="Đường dẫn thư mục chứa file Office"),
    merged_name: str = Query("merged.pdf", description="Tên file PDF gộp")
):
    output_dir = Path(__file__).parent / "office-to-pdf"
    try:
        pdf_files = convert_office_folder_to_pdf(folder_path, output_dir)
        if not pdf_files:
            return {"message": "Không tìm thấy file Office nào để convert."}

        merged_pdf = merge_pdfs(pdf_files, output_dir / merged_name)
        return {
            "message": f"Đã convert {len(pdf_files)} file và ghép thành công.",
            "output_folder": str(output_dir),
            "pdf_files": [str(p) for p in pdf_files],
            "merged_pdf": str(merged_pdf)
        }
    except Exception as e:
        return {"error": str(e)}


@app.post("/extract-pdf")
def extract_pdf_api(
    folder_path: str = Query(..., description="Thư mục chứa file PDF"),
    output_dir: str = Query("output_pdfs", description="Thư mục lưu PDF xuất ra")
):
    try:
        new_files = extract_text_from_folder(folder_path, output_dir)
        return {
            "message": f"Đã trích xuất {len(new_files)} file PDF text-only.",
            "files": [str(f) for f in new_files]
        }
    except Exception as e:
        return {"error": str(e)}

@app.post("/convert-extract-download")
def convert_extract_download_api(
    folder_path: str = Query(..., description="Đường dẫn chứa file txt, doc, pptx, pdf...")
):
    # Tạo thư mục tạm
    temp_dir = Path(tempfile.mkdtemp())

    try:
        pdf_dir = temp_dir / "pdfs"
        text_pdf_dir = temp_dir / "text_pdfs"
        merged_pdf_path = temp_dir / "merged.pdf"

        # Convert Office -> PDF
        converted_pdfs = convert_office_folder_to_pdf(folder_path, pdf_dir)

        # PDF có sẵn
        existing_pdfs = list(Path(folder_path).glob("*.pdf"))
        all_pdfs = converted_pdfs + existing_pdfs

        if not all_pdfs:
            return {"error": "Không tìm thấy file PDF hoặc Office hợp lệ."}

        # Merge PDF
        merge_pdfs(all_pdfs, merged_pdf_path)

        # Extract text
        extract_text_from_folder(pdf_dir, text_pdf_dir)
        extract_text_from_folder(folder_path, text_pdf_dir)

        # Tạo file zip trong thư mục CỐ ĐỊNH
        zip_path = Path("result.zip")  # Lưu ngay tại thư mục chạy server
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(merged_pdf_path, arcname="merged.pdf")
            for pdf in all_pdfs:
                zipf.write(pdf, arcname=f"pdfs/{pdf.name}")
            for text_pdf in text_pdf_dir.glob("*.pdf"):
                zipf.write(text_pdf, arcname=f"text_pdfs/{text_pdf.name}")

        # Trả file zip
        return FileResponse(
            zip_path,
            media_type='application/zip',
            filename="result.zip"
        )

    except Exception as e:
        return {"error": str(e)}

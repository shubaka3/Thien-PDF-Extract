# main.py
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
    extract_text_from_folder,
    # ADD THIS LINE:
    extract_text_from_pdf, # Import the function
    save_texts_to_pdf # Import save_texts_to_pdf as well, as it's used directly
)
# Bạn set API key ở biến môi trường hoặc gán trực tiếp ở đây
app = FastAPI(title="PPTX to PDF API")

OUTPUT_DIR = Path(__file__).parent / "pptx-to-pdf"
TEMP_UPLOAD_DIR = Path(__file__).parent / "temp_uploads"
TEMP_UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


class Message(BaseModel):
    role: str
    content: str


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

        # Ensure only text-only PDFs are merged
        pdf_list = sorted(folder.glob("*_text_only.pdf")) # Modified to glob for text-only PDFs
        if not pdf_list:
            return {"message": "Không tìm thấy file PDF text-only nào để ghép."}

        # Ghi vào thư mục 'output' cùng cấp với main.py
        output_dir = Path(__file__).parent / "output"
        output_dir.mkdir(exist_ok=True) # Ensure output directory exists
        output_path = output_dir / merged_name

        merged_pdf = merge_pdfs(pdf_list, output_path)

        return {
            "message": f"Gộp {len(pdf_list)} file PDF text-only thành công.",
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

        # This endpoint still merges the converted PDFs, not necessarily text-only ones.
        # If the intention is to convert, extract, then merge extracted, this endpoint logic needs reevaluation.
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
        # Adjusted to create output_pdfs in the same parent as main.py
        actual_output_dir = Path(__file__).parent / output_dir
        new_files = extract_text_from_folder(folder_path, actual_output_dir)
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
    # Create temporary directory
    temp_dir = Path(tempfile.mkdtemp())

    try:
        pdf_dir = temp_dir / "converted_pdfs" # Store converted PDFs
        text_pdf_dir = temp_dir / "text_only_pdfs" # Store text-only PDFs
        merged_text_pdf_path = temp_dir / "merged_text_only.pdf" # Path for the merged text-only PDF

        # Convert Office -> PDF
        converted_pdfs = convert_office_folder_to_pdf(folder_path, pdf_dir)

        # Existing PDFs in the source folder
        existing_pdfs = list(Path(folder_path).glob("*.pdf"))

        # Process all PDFs (converted and existing) to extract text
        all_pdfs_for_extraction = converted_pdfs + existing_pdfs
        
        # Extract text from all PDFs and create text-only PDFs in text_pdf_dir
        text_only_pdfs = []
        for pdf_file in all_pdfs_for_extraction:
            texts = extract_text_from_pdf(pdf_file)
            # Ensure the output directory for text-only PDFs exists
            text_pdf_dir.mkdir(parents=True, exist_ok=True)
            output_text_pdf = save_texts_to_pdf(texts, text_pdf_dir, pdf_file.stem)
            text_only_pdfs.append(output_text_pdf)

        if not text_only_pdfs:
            return {"error": "Không tìm thấy file PDF hoặc Office hợp lệ để trích xuất văn bản."}

        # Merge only the text-only PDFs
        merge_pdfs(text_only_pdfs, merged_text_pdf_path)

        # Create zip file in a fixed directory
        zip_path = Path("result.zip")  # Save in the server's running directory
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            if merged_text_pdf_path.exists():
                zipf.write(merged_text_pdf_path, arcname="merged_text_only.pdf")
            
            # Add all individual text-only PDFs to the zip
            for text_pdf in text_pdf_dir.glob("*.pdf"):
                zipf.write(text_pdf, arcname=f"text_only_pdfs/{text_pdf.name}")

        # Return the zip file
        return FileResponse(
            zip_path,
            media_type='application/zip',
            filename="result.zip"
        )

    except Exception as e:
        return {"error": str(e)}
    finally:
        # Clean up temporary directory
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
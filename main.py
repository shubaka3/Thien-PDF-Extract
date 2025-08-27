# main.py
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Query, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask
from fastapi.middleware.cors import CORSMiddleware


# Các hàm từ convert.py vẫn được import và sử dụng như cũ
from convert import (
    convert_office_folder_to_pdf,
    extract_text_from_pdf,
    merge_pdfs,
    save_texts_to_pdf,
)


app = FastAPI(title="File Conversion and Extraction API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Cho phép tất cả domain, hoặc list domain cụ thể
    allow_credentials=True,
    allow_methods=["*"],  # Cho phép GET, POST, PUT, DELETE, ...
    allow_headers=["*"],  # Cho phép tất cả headers
)

# --- HELPER FUNCTIONS ---
# Các hàm này đã tốt, giữ nguyên để sử dụng cho các endpoint mới
def remove_temp_dir(path: Path):
    """Xóa toàn bộ thư mục tạm một cách an toàn."""
    try:
        shutil.rmtree(path)
    except Exception as e:
        print(f"Error removing temp dir {path}: {e}")

def save_and_extract_files(files: List[UploadFile]) -> Path:
    """
    Lưu tất cả file upload vào 1 thư mục tạm và giải nén nếu là file zip.
    Trả về đường dẫn đến thư mục tạm.
    """
    temp_dir = Path(tempfile.mkdtemp(prefix="api_upload_"))
    for file in files:
        # Đảm bảo tên file an toàn
        filename = Path(file.filename).name
        file_path = temp_dir / filename
        with open(file_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        
        # Nếu là file zip thì giải nén và xóa file zip gốc
        if filename.lower().endswith(".zip"):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            file_path.unlink()
    return temp_dir

# --- REFACTORED API ENDPOINTS ---

@app.post("/convert-files", summary="Convert Office files to PDF")
async def convert_files_api(files: List[UploadFile] = File(..., description="Upload Office files (doc, docx, pptx) or a single ZIP")):
    """
    Chuyển đổi các file Office được upload thành PDF.
    - Nếu kết quả là 1 file PDF, trả về file đó.
    - Nếu kết quả là nhiều file PDF, trả về một file ZIP chứa tất cả chúng.
    """
    temp_dir = save_and_extract_files(files)
    
    try:
        output_dir = temp_dir / "converted_pdfs"
        pdf_files = convert_office_folder_to_pdf(str(temp_dir), str(output_dir))

        if not pdf_files:
            remove_temp_dir(temp_dir)
            return JSONResponse(status_code=400, content={"error": "No valid Office files found to convert."})

        # Logic trả về: 1 file hoặc ZIP
        if len(pdf_files) == 1:
            file_path = pdf_files[0]
            return FileResponse(
                file_path,
                filename=file_path.name,
                media_type='application/pdf',
                background=BackgroundTask(remove_temp_dir, temp_dir)
            )
        else:
            zip_path = temp_dir / "converted_files.zip"
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for pdf_file in pdf_files:
                    zipf.write(pdf_file, arcname=pdf_file.name)
            
            return FileResponse(
                zip_path,
                filename="converted_files.zip",
                media_type='application/zip',
                background=BackgroundTask(remove_temp_dir, temp_dir)
            )

    except Exception as e:
        remove_temp_dir(temp_dir)
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.post("/merge-files", summary="Merge multiple PDF files")
async def merge_files_api(
    files: List[UploadFile] = File(..., description="Upload multiple PDF files or a single ZIP"),
    merged_name: str = Query("merged.pdf", description="Output file name for the merged PDF")
):
    """
    Gộp nhiều file PDF được upload thành một file PDF duy nhất.
    """
    temp_dir = save_and_extract_files(files)

    try:
        pdf_list = sorted(temp_dir.glob("*.pdf"))
        if len(pdf_list) < 2:
            remove_temp_dir(temp_dir)
            return JSONResponse(status_code=400, content={"error": "At least two PDF files are required to merge."})
        
        output_path = temp_dir / merged_name
        merge_pdfs(pdf_list, output_path)

        return FileResponse(
            output_path,
            filename=merged_name,
            media_type='application/pdf',
            background=BackgroundTask(remove_temp_dir, temp_dir)
        )

    except Exception as e:
        remove_temp_dir(temp_dir)
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.post("/convert-and-merge", summary="Convert Office files and merge them into a single PDF")
async def convert_and_merge_api(
    files: List[UploadFile] = File(..., description="Upload Office files or a single ZIP"),
    merged_name: str = Query("merged.pdf", description="Output file name for the merged PDF")
):
    """
    Chuyển đổi tất cả file Office được upload thành PDF, sau đó gộp chúng lại thành 1 file PDF duy nhất.
    """
    temp_dir = save_and_extract_files(files)
    
    try:
        output_dir = temp_dir / "converted_pdfs"
        pdf_files = convert_office_folder_to_pdf(str(temp_dir), str(output_dir))

        if not pdf_files:
            remove_temp_dir(temp_dir)
            return JSONResponse(status_code=400, content={"error": "No valid Office files found to convert."})
        
        if len(pdf_files) == 1:
            # Nếu chỉ có 1 file, không cần merge, trả về file đó luôn
            file_path = pdf_files[0]
            return FileResponse(
                file_path,
                filename=file_path.name,
                media_type='application/pdf',
                background=BackgroundTask(remove_temp_dir, temp_dir)
            )

        output_path = temp_dir / merged_name
        merge_pdfs(pdf_files, output_path)

        return FileResponse(
            output_path,
            filename=merged_name,
            media_type='application/pdf',
            background=BackgroundTask(remove_temp_dir, temp_dir)
        )

    except Exception as e:
        remove_temp_dir(temp_dir)
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.post("/extract-text", summary="Extract text from PDF files")
async def extract_text_api(
    files: List[UploadFile] = File(..., description="Upload PDF files or a single ZIP"),
    return_format: str = Query("file", enum=["file", "text"], description="Return format: 'file' (text-only PDF) or 'text' (JSON)")
):
    """
    Trích xuất văn bản từ các file PDF.
    - `return_format='text'`: Trả về JSON chứa nội dung text.
    - `return_format='file'`: Trả về file PDF chỉ chứa text (hoặc ZIP nếu nhiều file).
    """
    temp_dir = save_and_extract_files(files)
    
    try:
        all_pdfs = sorted(temp_dir.glob("*.pdf"))
        if not all_pdfs:
            remove_temp_dir(temp_dir)
            return JSONResponse(status_code=400, content={"error": "No PDF files found to extract text from."})

        # --- Lựa chọn 1: Trả về JSON chứa text ---
        if return_format == "text":
            extracted_data = {}
            for pdf_file in all_pdfs:
                texts = extract_text_from_pdf(pdf_file)
                # Ghép text từ các trang lại thành một chuỗi duy nhất
                extracted_data[pdf_file.name] = "\n".join(texts)
            
            remove_temp_dir(temp_dir) # Dọn dẹp ngay lập tức
            return JSONResponse(content=extracted_data)

        # --- Lựa chọn 2: Trả về file PDF text-only ---
        elif return_format == "file":
            text_pdf_dir = temp_dir / "text_only_pdfs"
            text_pdf_dir.mkdir()
            
            output_files = []
            for pdf_file in all_pdfs:
                texts = extract_text_from_pdf(pdf_file)
                output_path = save_texts_to_pdf(texts, text_pdf_dir, pdf_file.stem)
                output_files.append(output_path)
            
            if len(output_files) == 1:
                file_path = output_files[0]
                return FileResponse(
                    file_path,
                    filename=file_path.name,
                    media_type='application/pdf',
                    background=BackgroundTask(remove_temp_dir, temp_dir)
                )
            else:
                zip_path = temp_dir / "extracted_text_files.zip"
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for text_pdf in output_files:
                        zipf.write(text_pdf, arcname=text_pdf.name)
                
                return FileResponse(
                    zip_path,
                    filename="extracted_text_files.zip",
                    media_type='application/zip',
                    background=BackgroundTask(remove_temp_dir, temp_dir)
                )

    except Exception as e:
        remove_temp_dir(temp_dir)
        return JSONResponse(status_code=500, content={"error": str(e)})

# Endpoint gốc của bạn vẫn hoạt động tốt, giữ lại làm ví dụ cho một quy trình phức tạp
@app.post("/convert-extract-download", summary="All-in-one: Convert, Extract, Merge, and Download")
async def convert_extract_download_api(
    files: List[UploadFile] = File(..., description="Upload Office/PDF files or a single ZIP")
):
    """
    Quy trình đầy đủ:
    1. Upload file Office và/hoặc PDF.
    2. Convert các file Office thành PDF.
    3. Trích xuất text từ tất cả các file PDF (cả mới convert và có sẵn).
    4. Tạo các file PDF chỉ chứa text.
    5. Gộp tất cả các file text-only PDF lại.
    6. Trả về một file ZIP chứa file đã gộp và tất cả các file text-only PDF riêng lẻ.
    """
    temp_dir = save_and_extract_files(files)

    try:
        # 1. Convert Office -> PDF
        pdf_dir = temp_dir / "converted_pdfs"
        converted_pdfs = convert_office_folder_to_pdf(str(temp_dir), str(pdf_dir))

        # 2. Lấy tất cả PDF để trích xuất (cả có sẵn và mới convert)
        existing_pdfs = list(temp_dir.glob("*.pdf"))
        all_pdfs_for_extraction = converted_pdfs + existing_pdfs

        if not all_pdfs_for_extraction:
            remove_temp_dir(temp_dir)
            return JSONResponse(status_code=400, content={"error": "No valid Office or PDF files found."})

        # 3. Trích xuất và tạo text-only PDF
        text_pdf_dir = temp_dir / "text_only_pdfs"
        text_pdf_dir.mkdir(parents=True, exist_ok=True)
        text_only_pdfs = []
        for pdf_file in all_pdfs_for_extraction:
            texts = extract_text_from_pdf(pdf_file)
            output_text_pdf = save_texts_to_pdf(texts, text_pdf_dir, pdf_file.stem)
            text_only_pdfs.append(output_text_pdf)
        
        if not text_only_pdfs:
            remove_temp_dir(temp_dir)
            return JSONResponse(status_code=400, content={"error": "Could not extract any text from the provided files."})
        
        # 4. Merge các text-only PDF
        merged_text_pdf_path = temp_dir / "merged_text_only.pdf"
        merge_pdfs(text_only_pdfs, merged_text_pdf_path)

        # 5. Tạo file ZIP kết quả
        zip_path = temp_dir / "result.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            if merged_text_pdf_path.exists():
                zipf.write(merged_text_pdf_path, arcname="merged_text_only.pdf")
            for text_pdf in text_only_pdfs:
                zipf.write(text_pdf, arcname=f"text_only_pdfs/{text_pdf.name}")
        
        # 6. Trả về ZIP và lên lịch xóa thư mục tạm
        return FileResponse(
            zip_path,
            media_type='application/zip',
            filename="result.zip",
            background=BackgroundTask(remove_temp_dir, temp_dir)
        )

    except Exception as e:
        remove_temp_dir(temp_dir)
        return JSONResponse(status_code=500, content={"error": str(e)})

# --- ENDPOINT MỚI (extractor_service) ---

from extractor_service import (
    extract_text_from_pdf,
    extract_text_from_word,
    extract_text_from_pptx,
    extract_data_from_excel_as_markdown,
    chunk_text,
)

@app.post("/super-extract", summary="Extract and chunk data from various file types for RAG")
async def super_extract_api(
    files: List[UploadFile] = File(..., description="Upload files (.pdf, .docx, .pptx, .xlsx) or a ZIP"),
    custom_prefix: str = Query("", description="Văn bản tùy biến để thêm vào đầu mỗi chunk dữ liệu"),
    chunk_size: int = Query(0, description="Số ký tự tối đa cho mỗi chunk text. Bỏ qua nếu bằng 0."),
    max_tokens: int = Query(256, description="Số từ (token) tối đa cho mỗi chunk text. Ưu tiên hơn chunk_size."),
    xlsx_row_limit: int = Query(50, description="Số dòng tối đa cho mỗi bảng Markdown từ file Excel.")
) -> JSONResponse:
    """
    API đa năng để trích xuất và chuẩn bị dữ liệu cho RAG:
    - **Hỗ trợ**: PDF, Word, PowerPoint, Excel.
    - **Chunking**: Chia nhỏ văn bản theo số token (từ) hoặc ký tự.
    - **Excel to Markdown**: Chuyển đổi bảng Excel thành Markdown, tự động lặp lại header khi chia nhỏ.
    - **Custom Prefix**: Cho phép thêm metadata/context tùy chỉnh vào đầu mỗi chunk.
    """
    temp_dir = save_and_extract_files(files)
    results: Dict[str, List[str]] = {}

    try:
        # Lấy danh sách tất cả file sau khi đã giải nén (nếu có)
        all_files = [p for p in temp_dir.iterdir() if p.is_file()]

        for file_path in all_files:
            file_name = file_path.name
            file_ext = file_path.suffix.lower()
            
            extracted_chunks = []

            if file_ext == ".pdf":
                text = extract_text_from_pdf(file_path)
                extracted_chunks = chunk_text(text, chunk_size, max_tokens)
            
            elif file_ext == ".docx":
                text = extract_text_from_word(file_path)
                extracted_chunks = chunk_text(text, chunk_size, max_tokens)

            elif file_ext == ".pptx":
                text = extract_text_from_pptx(file_path)
                extracted_chunks = chunk_text(text, chunk_size, max_tokens)

            elif file_ext == ".xlsx":
                # Hàm excel đã tự xử lý chunking, không cần gọi chunk_text
                extracted_chunks = extract_data_from_excel_as_markdown(file_path, xlsx_row_limit)
            
            else:
                results[file_name] = [f"File type '{file_ext}' is not supported."]
                continue

            # Thêm prefix vào đầu mỗi chunk nếu có
            if custom_prefix:
                # Thêm dấu cách sau prefix nếu nó chưa có để phân tách với nội dung
                prefix = custom_prefix if custom_prefix.endswith(" ") else f"{custom_prefix} "
                results[file_name] = [prefix + chunk for chunk in extracted_chunks]
            else:
                results[file_name] = extracted_chunks

    finally:
        # Luôn đảm bảo thư mục tạm được xóa
        remove_temp_dir(temp_dir)

    return JSONResponse(content=results)


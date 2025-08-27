# extractor.py
from pathlib import Path
from typing import List
import docx
import openpyxl
import zipfile
import xml.etree.ElementTree as ET
from PyPDF2 import PdfReader

# PDF ------------------------------------------------------------------
def extract_text_from_pdf(path: Path) -> str:
    """Trích xuất toàn bộ text từ một file PDF, bỏ qua lỗi trang."""
    texts = []
    try:
        reader = PdfReader(path)
        for i, page in enumerate(reader.pages, start=1):
            try:
                page_text = page.extract_text()
                if page_text:
                    texts.append(page_text.strip())
            except Exception as e:
                texts.append(f"[⚠️ Lỗi đọc trang {i}: {e}]")
    except Exception as e:
        return f"Error reading PDF {path.name}: {e}"
    return "\n\n".join(texts)


# DOCX -----------------------------------------------------------------
def extract_text_from_word(path: Path) -> str:
    """Trích xuất toàn bộ text từ một file .docx."""
    texts = []
    try:
        doc = docx.Document(path)
        for para in doc.paragraphs:
            if para.text.strip():
                texts.append(para.text.strip())
    except Exception as e:
        return f"Error reading DOCX {path.name}: {e}"
    return "\n\n".join(texts)


# PPTX -----------------------------------------------------------------
def extract_text_from_pptx(path: Path) -> str:
    """Trích xuất text từ file .pptx, bỏ qua media/audio."""
    text_chunks = []
    try:
        with zipfile.ZipFile(path, 'r') as z:
            # lọc tất cả slide XML
            slide_files = [f for f in z.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")]
            for slide_file in sorted(slide_files):  # sort để giữ thứ tự slide
                try:
                    xml_content = z.read(slide_file)
                    tree = ET.fromstring(xml_content)
                    # namespace pptx
                    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                    # lấy text trong <a:t>
                    texts = [node.text for node in tree.findall(".//a:t", ns) if node.text]
                    if texts:
                        text_chunks.append("\n".join(texts))
                except Exception as inner_e:
                    text_chunks.append(f"[⚠️ Lỗi đọc {slide_file}: {inner_e}]")
    except Exception as e:
        return f"Error reading PPTX {path.name}: {e}"

    return "\n\n".join(text_chunks)


# XLSX -----------------------------------------------------------------
def extract_data_from_excel_as_markdown(path: Path, row_limit: int = 50) -> List[str]:
    """
    Trích xuất dữ liệu từ các sheet trong Excel và chuyển thành Markdown.
    Mỗi chunk <= row_limit, header được lặp lại để giữ ngữ cảnh.
    """
    if row_limit <= 0:
        row_limit = 50

    all_markdown_chunks = []
    try:
        workbook = openpyxl.load_workbook(path, data_only=True)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            data = list(sheet.values)

            if not data:
                continue

            header = [str(cell) if cell is not None else "" for cell in data[0]]
            rows = [[str(cell) if cell is not None else "" for cell in row] for row in data[1:]]

            if not rows:
                continue

            # Header Markdown
            md_header = f"| {' | '.join(header)} |"
            md_separator = f"| {' | '.join(['---'] * len(header))} |"

            # Chunk rows
            for i in range(0, len(rows), row_limit):
                chunk_rows = rows[i:i + row_limit]

                chunk_md_lines = [f"## Sheet: {sheet_name}\n"]
                chunk_md_lines.append(md_header)
                chunk_md_lines.append(md_separator)

                for row in chunk_rows:
                    chunk_md_lines.append(f"| {' | '.join(row)} |")

                all_markdown_chunks.append("\n".join(chunk_md_lines))

    except Exception as e:
        all_markdown_chunks.append(f"Error reading XLSX {path.name}: {e}")

    return all_markdown_chunks


# Chunk helper ---------------------------------------------------------
def chunk_text(text: str, chunk_size: int = 0, max_tokens: int = 0) -> List[str]:
    """
    Chia nhỏ văn bản dựa trên số ký tự (chunk_size) hoặc số từ (max_tokens).
    max_tokens được ưu tiên nếu cả hai được cung cấp.
    """
    if not text:
        return []

    # Ưu tiên chunk theo token (giả lập = đếm từ)
    if max_tokens > 0:
        words = text.split()
        return [" ".join(words[i:i + max_tokens]) for i in range(0, len(words), max_tokens)]

    # Chunk theo ký tự
    if chunk_size > 0:
        return [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]

    return [text]

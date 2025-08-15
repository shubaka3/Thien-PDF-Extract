import subprocess
import textwrap
from pathlib import Path
from typing import List
from PyPDF2 import PdfMerger, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4


def convert_office_folder_to_pdf(folder_path: str, output_dir: str) -> List[Path]:
    """
    Convert tất cả file Office (pptx, doc, docx) trong thư mục sang PDF.
    Trả về danh sách đường dẫn PDF đã tạo.
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Folder '{folder_path}' không tồn tại hoặc không phải thư mục.")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    office_files = list(folder.glob("*.pptx")) + list(folder.glob("*.doc")) + list(folder.glob("*.docx"))
    if not office_files:
        return []

    LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
    pdf_files = []

    for office_file in office_files:
        subprocess.run([
            LIBREOFFICE_PATH,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(office_file)
        ], check=True)
        pdf_path = output_dir / (office_file.stem + ".pdf")
        pdf_files.append(pdf_path)

    return pdf_files


def merge_pdfs(pdf_list: List[Path], output_path: Path):
    """Ghép nhiều file PDF thành 1."""
    if not pdf_list:
        raise ValueError("Danh sách PDF rỗng, không thể ghép.")
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(str(pdf))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    merger.write(str(output_path))
    merger.close()
    return output_path


def extract_text_from_pdf(pdf_path: Path) -> list[str]:
    """Trích xuất text từ PDF, trả về list từng trang"""
    reader = PdfReader(str(pdf_path))
    texts = []
    for page in reader.pages:
        text = page.extract_text()
        texts.append(text.strip() if text else "")
    return texts


def save_texts_to_pdf(pages_text: list[str], output_path: Path):
    """Lưu danh sách trang text thành PDF text-only"""
    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4
    c.setFont("Helvetica", 11)
    y = height - 50

    for page_text in pages_text:
        lines = textwrap.wrap(page_text, width=90)
        for line in lines:
            c.drawString(50, y, line)
            y -= 15
            if y < 50:
                c.showPage()
                y = height - 50
                c.setFont("Helvetica", 11)
        c.showPage()

    c.save()


def extract_text_from_folder(folder_path: str, output_dir: str) -> List[Path]:
    """
    Trích xuất text từ tất cả PDF trong thư mục, lưu ra file PDF text-only.
    Trả về danh sách file text-only PDF.
    """
    folder = Path(folder_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = list(folder.glob("*.pdf"))
    output_files = []

    for pdf_file in pdf_files:
        texts = extract_text_from_pdf(pdf_file)
        out_path = output_dir / f"{pdf_file.stem}_text.pdf"
        save_texts_to_pdf(texts, out_path)
        output_files.append(out_path)

    return output_files

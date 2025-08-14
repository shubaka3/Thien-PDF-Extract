# convert.py
import subprocess
import textwrap
from pathlib import Path
from typing import List
from PyPDF2 import PdfMerger, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

def convert_office_folder_to_pdf(folder_path: str, output_dir: str) -> List[Path]:
    """
    Convert all Office files (pptx, doc, docx) in a folder to PDF.
    Returns a list of paths to the created PDFs.
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Folder '{folder_path}' does not exist or is not a directory.")

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
    """Merges multiple PDF files into one. This function is now specifically for text-only PDFs."""
    if not pdf_list:
        raise ValueError("Empty PDF list, cannot merge.")
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(str(pdf))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    merger.write(str(output_path))
    merger.close()
    return output_path

def extract_text_from_folder(folder_path: str, output_dir: str) -> List[Path]:
    """
    Extracts text from all PDFs in a folder, saves them as text-only PDFs.
    Returns a list of text-only PDF files.
    """
    folder = Path(folder_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = list(folder.glob("*.pdf"))
    output_files = []

    for pdf_file in pdf_files:
        texts = extract_text_from_pdf(pdf_file)
        # Pass the original file stem to save_texts_to_pdf for naming
        out_path = save_texts_to_pdf(texts, output_dir, pdf_file.stem)
        output_files.append(out_path)

    return output_files


def extract_text_from_pdf(pdf_path: Path) -> list[str]:
    """Extracts text from a PDF, returning a list of strings per page."""
    reader = PdfReader(str(pdf_path))
    texts = []
    for page in reader.pages:
        text = page.extract_text()
        texts.append(text.strip() if text else "")
    return texts

def save_texts_to_pdf(pages_text: list[str], output_dir: Path, original_file_stem: str, lines_per_chunk: int = 10) -> Path:
    """
    Saves a list of text pages into a text-only PDF.
    Combines multiple short lines into a single page and prefixes each line with the original file name.
    """
    output_path = output_dir / f"{original_file_stem}_text_only.pdf"
    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4
    font_size = 11 # Default font size
    line_height = font_size + 4 # Space between lines for readability
    margin_left = 50
    margin_right = width - 50
    text_width_limit = margin_right - margin_left

    c.setFont("Helvetica", font_size)
    y_position = height - 50

    current_chunk_lines = []

    for page_num, page_text in enumerate(pages_text):
        # Prefix each line with the original file stem ONLY
        # textwrap will handle breaking long lines within the page_text
        # We also add a small indent for the actual text after the filename
        prefixed_text_block = f"{original_file_stem}: {page_text}"

        # Use textwrap to wrap the prefixed text block
        # The width parameter needs to be carefully adjusted based on font size and margins
        # A rough character count that fits within the text_width_limit
        # For Helvetica 11pt, roughly 10 characters per inch. A4 width is ~8.27 inches.
        # (8.27 - 2*0.5 inch margin) * 10 char/inch = ~72 characters.
        # Let's be generous with width for wrapping.
        wrapped_lines = textwrap.wrap(prefixed_text_block, width=int(text_width_limit / (font_size * 0.6))) # Adjusted width calculation
        
        current_chunk_lines.extend(wrapped_lines)

        # Check if we have enough lines for a chunk or it's the last original page
        if len(current_chunk_lines) >= lines_per_chunk or page_num == len(pages_text) - 1:
            for line in current_chunk_lines:
                # If current line causes overflow on the PDF page, start a new PDF page
                if y_position < 50:
                    c.showPage()
                    y_position = height - 50
                    c.setFont("Helvetica", font_size) # Reset font after new page
                
                # Draw the line
                c.drawString(margin_left, y_position, line)
                y_position -= line_height # Move to the next line

            current_chunk_lines = [] # Reset chunk after writing
            
            # After writing a chunk, ensure there's a new page if more content follows
            # This prevents merging unrelated content on the same page unnecessarily
            if page_num < len(pages_text) - 1:
                c.showPage()
                y_position = height - 50 # Reset Y for the new page

    # Ensure to save even if no content was added to a new page at the end of the loop
    if current_chunk_lines: # If there's any remaining chunk that didn't trigger a page break
        for line in current_chunk_lines:
            if y_position < 50:
                c.showPage()
                y_position = height - 50
                c.setFont("Helvetica", font_size)
            c.drawString(margin_left, y_position, line)
            y_position -= line_height
        c.showPage() # Add a final showPage to flush content if needed
    
    # Only save if there's at least one page with content, otherwise PyPDF2 might complain about empty PDF
    # The canvas.Canvas constructor already implicitly starts the first page.
    # We should ensure showPage is called at least once if there's content.
    # A simple check for at least one page being drawn should be sufficient.
    
    c.save()
    return output_path
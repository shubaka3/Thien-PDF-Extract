from fastapi import UploadFile, HTTPException
from fastapi.responses import JSONResponse
from paddleocr import PaddleOCR
from PIL import Image
import pytesseract
import numpy as np
import io
import time

# Khởi tạo sẵn model PaddleOCR cho tiếng Anh và tiếng Việt
ocr_models = {
    "eng": PaddleOCR(use_angle_cls=True, lang="en"),
    "vie": PaddleOCR(use_angle_cls=True, lang="vi")
}

def read_image(contents: bytes):
    """Đọc bytes thành ảnh numpy array RGB"""
    return np.array(Image.open(io.BytesIO(contents)).convert("RGB"))


async def ocr_full(file: UploadFile, model: str, lang: str):
    """OCR mode: trả về box + text + confidence"""
    if model not in ("paddle", "tesseract"):
        raise HTTPException(status_code=400, detail="Model must be 'paddle' or 'tesseract'")
    if lang not in ("eng", "vie"):
        raise HTTPException(status_code=400, detail="Lang must be 'eng' or 'vie'")

    try:
        start_time = time.time()
        contents = await file.read()
        image_np = read_image(contents)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid image file")

    result = []

    if model == "paddle":
        paddle = ocr_models[lang]
        raw = paddle.ocr(image_np, cls=True)
        for block in raw:
            for line in block:
                result.append({
                    "box": line[0],
                    "text": line[1][0],
                    "confidence": line[1][1]
                })

    elif model == "tesseract":
        paddle = ocr_models[lang]
        raw = paddle.ocr(image_np, cls=True)
        for block in raw:
            for line in block:
                box = line[0]
                try:
                    x_coords = [int(pt[0]) for pt in box]
                    y_coords = [int(pt[1]) for pt in box]
                    x_min, x_max = max(0, min(x_coords)), max(x_coords)
                    y_min, y_max = max(0, min(y_coords)), max(y_coords)
                    cropped = image_np[y_min:y_max, x_min:x_max]
                except:
                    continue

                try:
                    text = pytesseract.image_to_string(cropped, lang=lang)
                except Exception as e:
                    text = f"[ERROR: {str(e)}]"

                result.append({
                    "box": box,
                    "text": text.strip()
                })

    return JSONResponse(content={
        "result": result,
        "time_ms": round((time.time() - start_time) * 1000, 2)
    })


async def ocr_fullV2(file: UploadFile, model: str, lang: str):
    """OCR mode: trả về format gốc (raw PaddleOCR style)"""
    if model not in ("paddle", "tesseract"):
        raise HTTPException(status_code=400, detail="Model must be 'paddle' or 'tesseract'")
    if lang not in ("eng", "vie"):
        raise HTTPException(status_code=400, detail="Lang must be 'eng' or 'vie'")

    try:
        start_time = time.time()
        contents = await file.read()
        image_np = read_image(contents)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid image file")

    result = []

    if model == "paddle":
        paddle = ocr_models[lang]
        result = paddle.ocr(image_np, cls=True)

    elif model == "tesseract":
        paddle = ocr_models[lang]
        raw = paddle.ocr(image_np, cls=True)
        for block in raw:
            for line in block:
                box = line[0]
                try:
                    x_coords = [int(pt[0]) for pt in box]
                    y_coords = [int(pt[1]) for pt in box]
                    x_min, x_max = max(0, min(x_coords)), max(x_coords)
                    y_min, y_max = max(0, min(y_coords)), max(y_coords)
                    cropped = image_np[y_min:y_max, x_min:x_max]
                except:
                    continue

                try:
                    text = pytesseract.image_to_string(cropped, lang=lang)
                except Exception as e:
                    text = f"[ERROR: {str(e)}]"

                result.append([
                    box,
                    [text.strip(), 1.0]
                ])

    return JSONResponse(content={
        "result": result,
        "time_ms": round((time.time() - start_time) * 1000, 2)
    })


def split_text_by_token(text: str, token: int):
    """Chia text thành token-length segments"""
    if token <= 0:
        return [text]
    words = text.split()
    return [" ".join(words[i:i+token]) for i in range(0, len(words), token)]

def split_text_by_chunk(text: str, chunk: int):
    """Chia text thành chunk-length segments theo ký tự"""
    if chunk <= 0:
        return [text]
    return [text[i:i+chunk] for i in range(0, len(text), chunk)]

async def ocr_fulltext(file: UploadFile, model: str, lang: str, token: int = 0, chunk: int = 0):
    if model not in ("paddle", "tesseract"):
        raise HTTPException(status_code=400, detail="Model must be 'paddle' or 'tesseract'")
    if lang not in ("eng", "vie"):
        raise HTTPException(status_code=400, detail="Lang must be 'eng' or 'vie'")

    try:
        start_time = time.time()
        contents = await file.read()
        image_np = read_image(contents)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid image file")

    full_text = ""

    if model == "paddle":
        paddle = ocr_models[lang]
        raw = paddle.ocr(image_np, cls=True)
        lines = [line[1][0] for block in raw for line in block]
        full_text = "\n".join(lines)

    elif model == "tesseract":
        try:
            full_text = pytesseract.image_to_string(Image.fromarray(image_np), lang=lang)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Tesseract error: {str(e)}")

    # Split text nếu token hoặc chunk > 0
    if token > 0:
        segments = split_text_by_token(full_text, token)
    elif chunk > 0:
        segments = split_text_by_chunk(full_text, chunk)
    else:
        segments = [full_text]

    return JSONResponse({
        "result": segments,
        "time_ms": round((time.time() - start_time) * 1000, 2)
    })
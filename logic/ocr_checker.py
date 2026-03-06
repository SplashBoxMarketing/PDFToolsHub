"""
logic/ocr_checker.py
Analyzes PDF pages and flags those that likely need OCR remediation.

Strategy: render each page as a raster image using pdf2image (poppler), then
run Tesseract on the rendered image. This sidesteps all compressed-stream
decoding issues and works regardless of how images are encoded in the PDF.

To isolate scanned content from readable text we render the page TWICE:
  - Full render:        what the page actually looks like
  - Text-hidden render: same page but with the PDF text layer removed

If Tesseract finds text in the text-hidden render, that text must be baked
into the image (i.e. a scan) rather than coming from the PDF text layer.

Requires:
  pip install pytesseract pillow pdfplumber pdf2image
  + Tesseract binary (Windows: https://github.com/UB-Mannheim/tesseract/wiki)
  + Poppler binary  (Windows: https://github.com/oschwartz10612/poppler-windows/releases)
"""

import io
import sys
import pdfplumber
from PIL import Image, ImageDraw

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

try:
    import pytesseract
    if sys.platform == "win32":
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False

# Minimum characters Tesseract must find to consider image as "contains text"
TESSERACT_MIN_CHARS = 15

# Minimum extractable text chars to consider the page itself "has readable text"
PAGE_TEXT_MIN_CHARS = 20

# DPI for page rendering — 150 is fast and sufficient for Tesseract
RENDER_DPI = 150


def _render_page(file_bytes: bytes, page_index: int) -> Image.Image | None:
    """Render a single PDF page (0-based index) to a PIL Image."""
    try:
        images = convert_from_bytes(
            file_bytes,
            dpi=RENDER_DPI,
            first_page=page_index + 1,
            last_page=page_index + 1,
        )
        return images[0].convert("RGB") if images else None
    except Exception as e:
        print(f"  [Page {page_index+1}] render failed: {e}")
        return None


def _tesseract_text(img: Image.Image) -> str:
    """Run Tesseract on a PIL image and return the extracted string."""
    try:
        return pytesseract.image_to_string(img, timeout=15)
    except Exception:
        return ""


def _classify(page_char_count: int, image_has_text: bool) -> str:
    if not image_has_text:
        return "OK"
    if page_char_count < PAGE_TEXT_MIN_CHARS:
        return "NEEDS OCR"
    return "REVIEW"


def analyze_pdf_bytes(file_bytes: bytes, filename: str) -> dict:
    """
    Analyze a PDF given its raw bytes.
    Returns a dict with summary counts and per-page details.
    """
    missing = []
    if not TESSERACT_AVAILABLE:
        missing.append("pytesseract  →  pip install pytesseract")
    if not PDF2IMAGE_AVAILABLE:
        missing.append("pdf2image    →  pip install pdf2image  (also needs Poppler)")
    if missing:
        return {
            "filename":  filename,
            "error":     "Missing dependencies: " + "; ".join(missing),
            "pages":     [], "total": 0,
            "needs_ocr": 0, "review": 0, "ok": 0,
        }

    pages = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            total_pages = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                page_num   = i + 1
                text       = page.extract_text() or ""
                char_count = len(text.strip())
                img_count  = len(page.images)

                # Only bother rendering pages that actually have embedded images
                if img_count == 0:
                    status = "OK"
                    image_has_text = False
                    print(f"  [Page {page_num}/{total_pages}] no images → OK")
                else:
                    # Render the full page
                    rendered = _render_page(file_bytes, i)
                    if rendered is None:
                        status = "OK"
                        image_has_text = False
                    else:
                        # Blank out the PDF text layer by drawing white rectangles
                        # over every word bounding box, leaving only image content
                        masked = rendered.copy()
                        draw  = ImageDraw.Draw(masked)
                        scale_x = rendered.width  / float(page.width)
                        scale_y = rendered.height / float(page.height)
                        for word in (page.extract_words() or []):
                            x0 = int(float(word["x0"]) * scale_x)
                            y0 = int(float(word["top"]) * scale_y)
                            x1 = int(float(word["x1"]) * scale_x)
                            y1 = int(float(word["bottom"]) * scale_y)
                            draw.rectangle([x0, y0, x1, y1], fill="white")

                        tess_text      = _tesseract_text(masked)
                        image_has_text = len(tess_text.strip()) >= TESSERACT_MIN_CHARS
                        status         = _classify(char_count, image_has_text)
                        print(f"  [Page {page_num}/{total_pages}] char_count={char_count}, "
                              f"img_count={img_count}, tess_chars={len(tess_text.strip())}, "
                              f"image_has_text={image_has_text}, status={status}")

                pages.append({
                    "page":           page_num,
                    "status":         status,
                    "char_count":     char_count,
                    "image_count":    img_count,
                    "image_has_text": image_has_text,
                })

    except Exception as e:
        return {
            "filename":  filename,
            "error":     str(e),
            "pages":     [], "total": 0,
            "needs_ocr": 0, "review": 0, "ok": 0,
        }

    needs_ocr = sum(1 for p in pages if p["status"] == "NEEDS OCR")
    review    = sum(1 for p in pages if p["status"] == "REVIEW")
    ok        = sum(1 for p in pages if p["status"] == "OK")

    return {
        "filename":  filename,
        "error":     None,
        "pages":     pages,
        "total":     len(pages),
        "needs_ocr": needs_ocr,
        "review":    review,
        "ok":        ok,
    }




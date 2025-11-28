#!/usr/bin/env python3
"""
auditram_highlight.py

Usage:
    python auditram_highlight.py --input /path/to/file --text "search phrase" [--case-sensitive]

Produces a new file (next to the input) with bounding boxes (red, unfilled) around all matches.
Supported input types: .pdf, .png, .jpg, .jpeg, .tiff, .docx, .xlsx

Notes:
 - For DOCX/XLSX the script will attempt to convert to PDF using LibreOffice (soffice).
 - For image OCR it uses pytesseract.
 - For PDF text-position search it uses PyMuPDF (fitz).
"""

import argparse
import os
import sys
import tempfile
import subprocess
import shutil
from pathlib import Path
from typing import List, Tuple

# External libraries:
# pip install PyMuPDF pillow pytesseract pdf2image python-docx openpyxl
# System dependencies (may be required):
# - LibreOffice (soffice) for docx/xlsx -> pdf conversion
# - Tesseract OCR (tesseract) for image text detection
# - poppler-utils (pdftoppm) for pdf2image on some platforms

try:
    import fitz  # PyMuPDF
except Exception as e:
    print("Missing dependency: PyMuPDF (fitz). Install with: pip install PyMuPDF")
    raise

from PIL import Image, ImageDraw
import pytesseract
from pdf2image import convert_from_path

# -------------------------
# Helpers for conversion
# -------------------------
def convert_to_pdf_with_libreoffice(input_path: str, out_dir: str) -> str:
    """
    Convert DOCX/XLSX to PDF using LibreOffice (soffice).
    Returns path to converted PDF or raises if conversion fails.
    """
    input_path = str(input_path)
    if shutil.which("soffice") is None:
        raise EnvironmentError("LibreOffice 'soffice' not found in PATH. Install LibreOffice or provide PDF input.")
    cmd = [
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        input_path
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    base = Path(input_path).stem
    converted = Path(out_dir) / f"{base}.pdf"
    if not converted.exists():
        raise FileNotFoundError(f"Conversion failed; expected {converted}")
    return str(converted)

# -------------------------
# PDF processing
# -------------------------
def search_pdf_and_draw(input_pdf: str, search_text: str, output_pdf: str, case_sensitive: bool=False):
    """
    Search for text in PDF pages and draw red, unfilled rectangles around occurrences.
    Saves a new PDF to output_pdf.
    """
    doc = fitz.open(input_pdf)
    # Work on a copy of the doc in memory, then save to output file
    for page_num in range(len(doc)):
        page = doc[page_num]
        # Try page.search_for first (quicker)
        search_target = search_text if case_sensitive else search_text.lower()

        rects = page.search_for(search_text, hit_max=4096)  # exact search
        if not rects and not case_sensitive:
            # fallback: search by words with lowercase matching
            words = page.get_text("words")  # list of (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            if words:
                # Build sequence matching for multi-word phrase
                word_texts = [w[4] for w in words]
                lowered = [w.lower() for w in word_texts]
                tokens = search_text.split()
                ltokens = [t.lower() for t in tokens]
                # slide over words to find sequences
                i = 0
                while i <= len(lowered) - len(ltokens):
                    if lowered[i:i+len(ltokens)] == ltokens:
                        # compute bounding box over those words
                        x0 = min(w[0] for w in words[i:i+len(ltokens)])
                        y0 = min(w[1] for w in words[i:i+len(ltokens)])
                        x1 = max(w[2] for w in words[i:i+len(ltokens)])
                        y1 = max(w[3] for w in words[i:i+len(ltokens)])
                        rects.append(fitz.Rect(x0, y0, x1, y1))
                        i += len(ltokens)
                    else:
                        i += 1
        # If still empty and case_sensitive True, try words-based exact match
        if not rects and case_sensitive:
            words = page.get_text("words")
            if words:
                word_texts = [w[4] for w in words]
                tokens = search_text.split()
                i = 0
                while i <= len(word_texts) - len(tokens):
                    if word_texts[i:i+len(tokens)] == tokens:
                        x0 = min(w[0] for w in words[i:i+len(tokens)])
                        y0 = min(w[1] for w in words[i:i+len(tokens)])
                        x1 = max(w[2] for w in words[i:i+len(tokens)])
                        y1 = max(w[3] for w in words[i:i+len(tokens)])
                        rects.append(fitz.Rect(x0, y0, x1, y1))
                        i += len(tokens)
                    else:
                        i += 1

        # Draw rectangles on the page for each rect found
        for r in rects:
            # Draw a red, unfilled rectangle (stroke only)
            page.draw_rect(r, color=(1, 0, 0), width=1.5)  # RGB with values in 0..1
    # Save to output_pdf
    doc.save(output_pdf)
    doc.close()

# -------------------------
# Image processing
# -------------------------
def search_image_and_draw(input_image_path: str, search_text: str, output_image_path: str, case_sensitive: bool=False):
    """
    Use pytesseract to get boxes, then draw rectangles around matches; saves output image.
    """
    img = Image.open(input_image_path).convert("RGB")
    ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    n_boxes = len(ocr_data['text'])
    draw = ImageDraw.Draw(img)
    target = search_text if case_sensitive else search_text.lower()
    # We will match tokens; also allow multi-word phrase matching across tokens
    words = [ocr_data['text'][i] for i in range(n_boxes)]
    # get cleaned list (indices)
    indices = list(range(n_boxes))
    # Compose normalized words for matching
    norm_words = [w for w in words]  # keep original for case-sensitive
    if not case_sensitive:
        norm_words = [w.lower() for w in words]
    tokens = target.split()
    i = 0
    while i <= len(norm_words) - len(tokens):
        if norm_words[i:i+len(tokens)] == tokens:
            # bounding box across these tokens
            xs = []
            ys = []
            xe = []
            ye = []
            for j in range(i, i+len(tokens)):
                if int(ocr_data['conf'][j]) > 0 and ocr_data['text'][j].strip() != "":
                    x = ocr_data['left'][j]
                    y = ocr_data['top'][j]
                    w = ocr_data['width'][j]
                    h = ocr_data['height'][j]
                    xs.append(x)
                    ys.append(y)
                    xe.append(x+w)
                    ye.append(y+h)
            if xs:
                bbox = (min(xs), min(ys), max(xe), max(ye))
                # draw unfilled rectangle (stroke only)
                draw.rectangle(bbox, outline=(255,0,0), width=3)
            i += len(tokens)
        else:
            i += 1
    img.save(output_image_path)

# -------------------------
# Orchestrator
# -------------------------
def process_file(input_path: str, search_text: str, case_sensitive: bool=False) -> str:
    """
    Determine file type, process accordingly, and return path to generated output file.
    """
    p = Path(input_path)
    if not p.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    ext = p.suffix.lower()
    out_dir = p.parent
    base = p.stem
    # Prepare output naming
    if ext in [".pdf"]:
        output_pdf = out_dir / f"{base}_boxed.pdf"
        search_pdf_and_draw(str(p), search_text, str(output_pdf), case_sensitive=case_sensitive)
        return str(output_pdf)
    elif ext in [".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif"]:
        output_img = out_dir / f"{base}_boxed{ext}"
        search_image_and_draw(str(p), search_text, str(output_img), case_sensitive=case_sensitive)
        return str(output_img)
    elif ext in [".docx", ".xlsx"]:
        # convert to pdf using LibreOffice, then process pdf
        with tempfile.TemporaryDirectory() as td:
            converted_pdf = convert_to_pdf_with_libreoffice(str(p), td)
            output_pdf = out_dir / f"{base}_boxed.pdf"
            search_pdf_and_draw(converted_pdf, search_text, str(output_pdf), case_sensitive=case_sensitive)
            return str(output_pdf)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

# -------------------------
# CLI
# -------------------------
def main():
    parser = argparse.ArgumentParser(description="AuditRAM: highlight search text across files with red bounding boxes.")
    parser.add_argument("--input", "-i", required=True, help="Path to input file (pdf, docx, xlsx, png, jpg, jpeg, tiff)")
    parser.add_argument("--text", "-t", required=True, help="Search text (string)")
    parser.add_argument("--case-sensitive", action="store_true", help="Make search case-sensitive (default: case-insensitive)")
    args = parser.parse_args()

    try:
        out = process_file(args.input, args.text, case_sensitive=args.case_sensitive)
        print(f"Output created: {out}")
    except Exception as e:
        print("ERROR:", str(e))
        sys.exit(2)

if __name__ == "__main__":
    main()

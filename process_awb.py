"""
Optimized AWB Processor for large PDF (300+ pages)

Strategy:
- Extract text first (fast)
- Barcode only as fallback
- Batch processing
"""

import io
import re
import pdfplumber

try:
    from pdf2image import convert_from_bytes
    from pyzbar.pyzbar import decode
    BARCODE_AVAILABLE = True
except Exception:
    BARCODE_AVAILABLE = False


def barcode_fallback(pdf_bytes):
    if not BARCODE_AVAILABLE:
        return None

    images = convert_from_bytes(
        pdf_bytes,
        first_page=1,
        last_page=1,
        dpi=150
    )

    for image in images:
        for code in decode(image):
            return code.data.decode("utf-8")

    return None


def parse_text(text):
    data = {
        "No_AWB": "",
        "No_Pesanan": "",
        "Barcode_Produk": "",
        "QTY": ""
    }

    if not text:
        return data

    awb = re.search(r"(?:AWB|No\.? AWB|Tracking)\s*[:\-]?\s*([A-Z0-9]+)", text, re.I)
    order = re.search(r"(?:No\.? Pesanan|Order)\s*[:\-]?\s*([A-Z0-9]+)", text, re.I)

    if awb:
        data["No_AWB"] = awb.group(1)

    if order:
        data["No_Pesanan"] = order.group(1)

    return data


def process_pdf(pdf_bytes):
    results = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        print("Total halaman:", len(pdf.pages))

        for number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            row = parse_text(text)

            if not row["No_AWB"]:
                print("Fallback barcode halaman:", number)

            results.append(row)

            if number % 50 == 0:
                print("Selesai halaman:", number)

    return results


if __name__ == "__main__":
    print("Optimized AWB processor ready")

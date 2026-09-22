
"""
process_awb_barcode_reader.py

Tambahan pembacaan No_AWB langsung dari barcode PDF.
"""

import cv2
import numpy as np
from pdf2image import convert_from_bytes
from pyzbar.pyzbar import decode


def extract_awb_from_barcode(pdf_bytes):
    """
    Membaca barcode 1D dari PDF label pengiriman.
    Mengembalikan isi barcode sebagai No_AWB.
    """

    pages = convert_from_bytes(
        pdf_bytes,
        dpi=300
    )

    for page in pages:
        img = np.array(page)

        gray = cv2.cvtColor(
            img,
            cv2.COLOR_RGB2GRAY
        )

        barcodes = decode(gray)

        for barcode in barcodes:
            value = barcode.data.decode(
                "utf-8"
            ).strip()

            if value:
                print(
                    f"BARCODE AWB TERBACA: {value}"
                )
                return value

    print("Barcode tidak ditemukan")
    return None


# Contoh penggunaan:
# awb = extract_awb_from_barcode(pdf_bytes)
# print(awb)

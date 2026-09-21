"""
process_dn.py

Khusus memproses PDF Delivery Note (DN) Lemonilo.
Tidak mengganggu process_awb.py.

Output:
Pesanan:
A No_AWB
B No_Pesanan

Detail_Pesanan:
A ID_Detail
B No_AWB
C Barcode_Produk
D Kode_Rak
E QTY_Diminta
F Qty_Discan
G Scan_Barcode
H Status
I UOM
"""

import io
import json
import os
import re

import pdfplumber
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

SHEET_PESANAN = "Pesanan"
SHEET_DETAIL = "Detail_Pesanan"


def get_services():
    key_json = os.environ["GCP_SA_KEY_JSON"]
    info = json.loads(key_json)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=SCOPES
    )
    return (
        build("drive", "v3", credentials=creds),
        build("sheets", "v4", credentials=creds),
    )


def list_new_pdfs(drive_service, folder_id):
    query = (
        f"'{folder_id}' in parents "
        "and mimeType='application/pdf' "
        "and trashed=false"
    )
    result = drive_service.files().list(
        q=query,
        fields="files(id,name)"
    ).execute()
    return result.get("files", [])


def download_pdf(drive_service, file_id):
    request = drive_service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    buf.seek(0)
    return buf


def extract_dn(pdf_bytes):
    text = ""

    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            text += "\n" + (page.extract_text() or "")

    # DN NO
    dn_match = re.search(
        r"DN\s*NO\s*[:\s]*([A-Z0-9\-]+)",
        text,
        re.IGNORECASE
    )

    if not dn_match:
        return None

    dn_no = dn_match.group(1).strip()

    items = []

    # Format umum Lemonilo:
    # ITEM ID ... QTY ... UOM
    rows = re.findall(
        r"(IFD[A-Z0-9\-\/]+).*?(\d+)\s+(PCS|CTN|BOX)",
        text,
        re.IGNORECASE | re.DOTALL
    )

    for item_id, qty, uom in rows:
        items.append({
            "barcode": item_id.strip(),
            "qty": int(qty),
            "uom": uom.upper()
        })

    return {
        "no_awb": dn_no,
        "no_pesanan": dn_no,
        "items": items
    }


def next_row(sheets_service, spreadsheet_id, sheet):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet}!A:A"
    ).execute()

    return len(result.get("values", [])) + 1


def write_sheet(sheets_service, spreadsheet_id, data):
    pesanan = []
    detail = []

    counter = 0

    pesanan.append([
        data["no_awb"],
        data["no_pesanan"]
    ])

    for item in data["items"]:
        counter += 1

        detail.append([
            f"DT-{data['no_awb']}-{counter}",
            data["no_awb"],
            item["barcode"],
            "",
            item["qty"],
            0,
            "",
            "Belum",
            item["uom"]
        ])

    if pesanan:
        row = next_row(
            sheets_service,
            spreadsheet_id,
            SHEET_PESANAN
        )

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_PESANAN}!A{row}:B{row}",
            valueInputOption="RAW",
            body={"values": pesanan},
        ).execute()

    if detail:
        row = next_row(
            sheets_service,
            spreadsheet_id,
            SHEET_DETAIL
        )

        last = row + len(detail) - 1

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_DETAIL}!A{row}:I{last}",
            valueInputOption="RAW",
            body={"values": detail},
        ).execute()


def main():
    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    folder_id = os.environ["FOLDER_ID_AWB_MASUK"]

    drive, sheets = get_services()

    files = list_new_pdfs(drive, folder_id)

    for file in files:
        pdf = download_pdf(drive, file["id"])

        data = extract_dn(pdf)

        if data:
            write_sheet(
                sheets,
                spreadsheet_id,
                data
            )

            print(
                f"Berhasil proses DN {data['no_awb']}"
            )
        else:
            print(
                f"Bukan format DN Lemonilo: {file['name']}"
            )


if __name__ == "__main__":
    main()

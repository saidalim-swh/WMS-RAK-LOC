"""
process_awb.py
================
Dijalankan otomatis oleh GitHub Actions tiap beberapa menit.
Tugas: cek folder Google Drive -> proses PDF AWB baru -> tulis ke Google Sheets
       (tab Pesanan & Detail_Pesanan) -> pindahkan PDF ke folder "Processed".

Tidak butuh Cloud Functions / billing account - script ini jalan di komputer
virtual gratis milik GitHub (GitHub Actions runner).

ENV VARS yang dibutuhkan (diisi lewat GitHub Secrets, lihat README):
- GCP_SA_KEY_JSON     : isi lengkap file JSON service account (sebagai teks)
- SPREADSHEET_ID      : ID Google Sheets WMS
- FOLDER_ID_AWB_MASUK : ID folder Drive tempat admin upload PDF AWB
"""
import io
import json
import os
import re
from collections import Counter

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
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return drive_service, sheets_service


def list_new_pdfs(drive_service, folder_id):
    query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    return results.get("files", [])


def get_or_create_processed_folder(drive_service, parent_folder_id):
    query = (
        f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' "
        f"and name='Processed' and trashed=false"
    )
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    folder_metadata = {
        "name": "Processed",
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }
    folder = drive_service.files().create(body=folder_metadata, fields="id").execute()
    return folder["id"]


def download_pdf(drive_service, file_id):
    request = drive_service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


def move_file(drive_service, file_id, new_parent_id, old_parent_id):
    drive_service.files().update(
        fileId=file_id,
        addParents=new_parent_id,
        removeParents=old_parent_id,
        fields="id, parents",
    ).execute()


def extract_page(text):
    candidates = re.findall(r"\b[A-Z]{2,5}\d{6,12}\b", text)
    no_awb = Counter(candidates).most_common(1)[0][0] if candidates else None

    m = re.search(r"No\.?\s*Pesanan\s*:\s*(\d+)", text)
    no_pesanan = m.group(1) if m else None
    if not no_pesanan:
        m2 = re.search(r"Order\s*Id\s*:\s*(\d+)", text, re.IGNORECASE)
        no_pesanan = m2.group(1) if m2 else None

    raw_items = re.findall(r"\b([A-Z]{2,}\d{4,})\b[^\n]*?\1\D+(\d+)\b", text)
    items, seen = [], set()
    for sku, qty in raw_items:
        key = (sku, qty)
        if key in seen:
            continue
        seen.add(key)
        items.append({"sku": sku, "qty": int(qty)})

    return {"no_awb": no_awb, "no_pesanan": no_pesanan, "items": items}


def extract_all_pages(pdf_bytes):
    results = []
    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            data = extract_page(text)
            if data["no_awb"]:
                results.append(data)
    return results


def write_to_sheets(sheets_service, spreadsheet_id, results):
    pesanan_rows, detail_rows = [], []
    counter = 0
    for data in results:
        pesanan_rows.append([data["no_awb"], "", "", "Pending"])
        for item in data["items"]:
            counter += 1
            id_detail = f"DT-{data['no_awb']}-{counter}"
            detail_rows.append(
                [id_detail, data["no_awb"], item["sku"], "", item["qty"], 0, "", "Belum"]
            )

    if pesanan_rows:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_PESANAN}!A:D",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": pesanan_rows},
        ).execute()

    if detail_rows:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_DETAIL}!A:H",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": detail_rows},
        ).execute()

    return len(pesanan_rows), len(detail_rows)


def main():
    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    folder_id = os.environ["FOLDER_ID_AWB_MASUK"]

    drive_service, sheets_service = get_services()
    processed_folder_id = get_or_create_processed_folder(drive_service, folder_id)

    files = list_new_pdfs(drive_service, folder_id)
    print(f"Ditemukan {len(files)} PDF baru di folder.")

    for f in files:
        print(f"Memproses: {f['name']} ({f['id']})")
        try:
            pdf_bytes = download_pdf(drive_service, f["id"])
            results = extract_all_pages(pdf_bytes)
            n_pesanan, n_detail = write_to_sheets(sheets_service, spreadsheet_id, results)
            print(f"  -> {len(results)} AWB ditemukan, {n_pesanan} baris Pesanan, {n_detail} baris Detail")
            move_file(drive_service, f["id"], processed_folder_id, folder_id)
            print(f"  -> dipindah ke folder Processed")
        except Exception as e:
            print(f"  GAGAL memproses {f['name']}: {e}")


if __name__ == "__main__":
    main()

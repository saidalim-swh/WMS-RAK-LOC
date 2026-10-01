"""
process_dn.py

Khusus memproses PDF Delivery Note (DN) Lemonilo.
Tidak mengganggu process_awb.py.

Perubahan dari versi sebelumnya:
- extract_dn() sekarang mendukung BEBERAPA DN dalam 1 file PDF
  (sebelumnya re.search() cuma menangkap DN pertama, dan item barang
  dari seluruh PDF tercampur jadi satu entry).
- Ditambahkan pengecekan duplikat (mirip process_awb.py), supaya kalau
  workflow kebetulan jalan dua kali untuk file yang sama sebelum file
  dipindah ke folder Processed, DN yang sama tidak tertulis dobel.

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


def get_or_create_processed_folder(drive_service, parent_folder_id):
    query = (
        f"'{parent_folder_id}' in parents "
        "and name='Processed' "
        "and mimeType='application/vnd.google-apps.folder' "
        "and trashed=false"
    )

    result = drive_service.files().list(
        q=query,
        fields="files(id,name)"
    ).execute()

    folders = result.get("files", [])

    if folders:
        return folders[0]["id"]

    folder = drive_service.files().create(
        body={
            "name": "Processed",
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_folder_id]
        },
        fields="id"
    ).execute()

    return folder["id"]


def move_to_processed(drive_service, file_id):
    masuk_folder = os.environ["FOLDER_ID_DN_MASUK"]

    processed_folder = get_or_create_processed_folder(
        drive_service,
        masuk_folder
    )

    drive_service.files().update(
        fileId=file_id,
        addParents=processed_folder,
        removeParents=masuk_folder,
        fields="id, parents"
    ).execute()


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
    """
    Ekstrak SEMUA DN dalam satu PDF (bisa lebih dari satu).
    Mengembalikan list of dict, bisa kosong kalau tidak ada DN yang
    terdeteksi sama sekali.
    """
    text = ""

    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            text += "\n" + (page.extract_text() or "")

    # Cari SEMUA kemunculan "DN NO", bukan cuma yang pertama
    matches = list(re.finditer(
        r"DN\s*NO\s*[:\s]*([A-Z0-9\-]+)",
        text,
        re.IGNORECASE
    ))

    if not matches:
        return []

    results = []
    for i, m in enumerate(matches):
        dn_no = m.group(1).strip()

        # Potong teks: dari akhir match DN ini sampai sebelum DN
        # berikutnya (atau sampai akhir teks kalau ini DN terakhir).
        # Supaya item barang dicari PER-DN, bukan dari teks gabungan
        # seluruh PDF.
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        segment = text[start:end]

        items = []
        rows = re.findall(
            r"(IFD[A-Z0-9\-\/]+).*?(\d+)\s+(PCS|CTN|BOX)",
            segment,
            re.IGNORECASE | re.DOTALL
        )
        for item_id, qty, uom in rows:
            clean_barcode = re.sub(r"[-/]+$", "", item_id.strip())
            items.append({
                "barcode": clean_barcode,
                "qty": int(qty),
                "uom": uom.upper()
            })

        results.append({
            "no_awb": dn_no,
            "no_pesanan": dn_no,
            "items": items
        })

    return results


def next_row(sheets_service, spreadsheet_id, sheet):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet}!A:A"
    ).execute()

    return len(result.get("values", [])) + 1


def get_existing_dn(sheets_service, spreadsheet_id):
    """Baca kolom A & B sheet Pesanan, kembalikan set No_AWB & No_Pesanan
    yang sudah pernah tercatat - dipakai untuk cek duplikat."""
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{SHEET_PESANAN}!A:B"
    ).execute()

    rows = result.get("values", [])
    existing = set()

    for row in rows[1:]:
        if len(row) >= 1 and row[0]:
            existing.add(("AWB", str(row[0]).strip()))
        if len(row) >= 2 and row[1]:
            existing.add(("PESANAN", str(row[1]).strip()))

    return existing


def filter_duplicate_dn(sheets_service, spreadsheet_id, data_list):
    """Buang DN yang No_AWB atau No_Pesanan-nya sudah ada di sheet Pesanan."""
    existing = get_existing_dn(sheets_service, spreadsheet_id)
    filtered = []

    for data in data_list:
        awb = str(data.get("no_awb") or "").strip()
        pesanan = str(data.get("no_pesanan") or "").strip()

        if ("AWB", awb) in existing or ("PESANAN", pesanan) in existing:
            print(f"SKIP DUPLICATE DN: {awb}")
            continue

        filtered.append(data)
        existing.add(("AWB", awb))
        if pesanan:
            existing.add(("PESANAN", pesanan))

    return filtered


def write_sheet(sheets_service, spreadsheet_id, data_list):
    pesanan = []
    detail = []

    for data in data_list:
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
        row = next_row(sheets_service, spreadsheet_id, SHEET_PESANAN)
        last = row + len(pesanan) - 1

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_PESANAN}!A{row}:B{last}",
            valueInputOption="RAW",
            body={"values": pesanan},
        ).execute()

    if detail:
        row = next_row(sheets_service, spreadsheet_id, SHEET_DETAIL)
        last = row + len(detail) - 1

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_DETAIL}!A{row}:I{last}",
            valueInputOption="RAW",
            body={"values": detail},
        ).execute()


def main():
    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    folder_id = os.environ["FOLDER_ID_DN_MASUK"]

    drive, sheets = get_services()

    files = list_new_pdfs(drive, folder_id)
    print(f"Ditemukan {len(files)} PDF baru di folder DN masuk.")

    for file in files:
        print(f"Memproses: {file['name']} ({file['id']})")

        try:
            pdf = download_pdf(drive, file["id"])
            data_list = extract_dn(pdf)

            if not data_list:
                print(f"  Bukan format DN Lemonilo: {file['name']}")
                continue

            print(f"  -> {len(data_list)} DN terdeteksi di file ini.")

            data_list = filter_duplicate_dn(sheets, spreadsheet_id, data_list)

            if not data_list:
                print("  Semua DN duplicate, tidak ditulis.")
                move_to_processed(drive, file["id"])
                continue

            write_sheet(sheets, spreadsheet_id, data_list)

            for data in data_list:
                print(f"  Berhasil proses DN {data['no_awb']} "
                      f"({len(data['items'])} item)")

            move_to_processed(drive, file["id"])
            print(f"  -> {file['name']} dipindahkan ke Processed")

        except Exception as e:
            print(f"  GAGAL memproses {file['name']}: {e}")


if __name__ == "__main__":
    main()

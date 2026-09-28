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
SHEET_PRODUK = "Produk"
SHEET_STOK = "Stok"


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

    if not no_awb:
        # Fallback untuk format AWB murni angka (Shopee Economy/Standard/Regular,
        # Cargo, FastTrack) yang tidak punya awalan huruf sama sekali.
        fallback_patterns = [
            r"\|\s*[A-Z]{2,5}\s+(\d{10,15})\b",              # "MKSX9 | MKS 11004386937208"
            r"[A-Z]{2,4}-[A-Z0-9]+-[A-Z0-9]+\s+(\d{10,15})\b",  # "PDG-PMM001A-KU 201773297878"
            r"FastTrack\s*\n\s*(\d{10,15})\b",                # "FastTrack\n570606842138"
            r"[A-Z0-9\-]+\s+(\d{10,15})\b",                  # "SU2-TBG-A 004663229510"
        ]
        for pat in fallback_patterns:
            fm = re.search(pat, text)
            if fm:
                no_awb = fm.group(1)
                break

    m = re.search(r"No\.?\s*Pesanan\s*:\s*([A-Za-z0-9]+)", text)
    no_pesanan = m.group(1) if m else None
    if not no_pesanan:
        m2 = re.search(r"Order\s*Id\s*:\s*([A-Za-z0-9]+)", text, re.IGNORECASE)
        no_pesanan = m2.group(1) if m2 else None

    # Item fisik: SEMUA kode SKU barang fisik Lemonilo selalu diawali "IFD"
    # (beda dari kode SKU paket/bundle yang diawali "OS..."). Cukup cari pola
    # "IFD... <spasi> <qty>" di mana saja di teks - TIDAK perlu 2 kemunculan
    # token yang sama (beda dari pendekatan lama yang gagal kalau bagian
    # "pembuka" rusak/tumpang tindih akibat PDF sumbernya sendiri).
    # Contoh yang harus tertangkap:
    #   "IFDBC00016 30"        -> normal
    #   "IFDBC00016- 10"       -> varian TTS/renceng, strip tanda '-' di akhir
    #   "...garbled...IFDBC00016 30" -> tetap ketemu walau teks sebelumnya rusak
    raw_items = re.findall(
    r"\b(IFD[A-Z0-9]+)\b(?:\s+\S+){0,15}?\s+(\d+)\b",
    text
    )
    items, seen = [], set()
    for sku, qty in raw_items:
        if sku == no_awb:
            continue  # jaga-jaga, walau harusnya tidak akan pernah collide
        key = (sku, qty)
        if key in seen:
            continue
        seen.add(key)
        items.append({"sku": sku, "qty": int(qty)})

    # Pengaman: kalau produk disebut "campur beberapa rasa" (mis. "9 Renceng
    # (3 Chocochips + 3 keju + 3 Strawberry)") tapi jumlah item yang berhasil
    # dibaca tidak sesuai jumlah rasa yang disebut - ini indikasi ada item
    # yang gagal terbaca akibat teks PDF yang tumpang tindih (bukan bug regex,
    # tapi masalah di PDF sumbernya). Kasih peringatan supaya dicek manual.
    warning = None
    # (?!\+) mencegah nomor telepon format "(+62)81..." ikut kedeteksi -
    # bundle asli selalu diawali angka (mis. "(3 Chocochips + 3 keju...)"),
    # bukan simbol '+' di awal seperti kode telepon.
    bundle_match = re.search(r"\((?!\+)([^()]*\+[^()]*)\)", text)
    if bundle_match:
        expected_flavors = bundle_match.group(1).count("+") + 1
        if len(items) < expected_flavors:
            warning = (
                f"Terdeteksi varian campur ({expected_flavors} rasa) tapi cuma "
                f"{len(items)} item yang terbaca - kemungkinan ada item hilang, cek manual."
            )

    return {"no_awb": no_awb, "no_pesanan": no_pesanan, "items": items, "warning": warning}


def extract_all_pages(pdf_bytes):
    """
    Ekstrak semua halaman. Kalau 1 AWB punya banyak barang, daftarnya bisa
    'meluber' ke halaman berikutnya - halaman lanjutan itu TIDAK punya
    barcode/No_AWB lagi. Halaman seperti itu dianggap lanjutan dari AWB
    terakhir yang terdeteksi, bukan dibuang.
    """
    results = []
    current = None
    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            data = extract_page(text)

            # Halaman "asli" (bukan lanjutan) selalu punya info pengirim/penerima.
            # Halaman lanjutan cuma berisi sambungan tabel barang saja.
            has_header = bool(re.search(r"Penerima|Pengirim|Receiver|Sender", text, re.IGNORECASE))

            if not data["no_awb"] and has_header and data["no_pesanan"]:
                # Halaman baru asli tapi tidak ada AWB terpisah (mis. tipe "INSTANT")
                # - No_Pesanan dipakai sebagai identitas fisik pengganti.
                data["no_awb"] = data["no_pesanan"]

            if data["no_awb"]:
                # halaman baru dengan No_AWB sendiri -> mulai entry baru
                current = data
                results.append(current)
            elif current is not None and data["items"]:
                # halaman tanpa No_AWB tapi ada data item -> lanjutan AWB sebelumnya
                existing_keys = {(it["sku"], it["qty"]) for it in current["items"]}
                for item in data["items"]:
                    key = (item["sku"], item["qty"])
                    if key not in existing_keys:
                        current["items"].append(item)
                        existing_keys.add(key)
                if data.get("warning") and not current.get("warning"):
                    current["warning"] = data["warning"]
    return results


def _get_next_empty_row(sheets_service, spreadsheet_id, sheet_name):
    """Hitung baris kosong berikutnya berdasarkan kolom A saja - lebih bisa
    diandalkan daripada mengandalkan auto-detect dari values.append()."""
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A:A"
    ).execute()
    values = result.get("values", [])
    return len(values) + 1  # +1 karena baris berikutnya setelah data terakhir


def build_sku_to_barcode_lookup(sheets_service, spreadsheet_id):
    """Baca tabel Produk, bikin mapping SKU -> Barcode_Produk (nomor barcode asli)."""
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"{SHEET_PRODUK}!A:C"
    ).execute()
    rows = result.get("values", [])
    if not rows:
        return {}
    header = rows[0]
    try:
        idx_barcode = header.index("Barcode_Produk")
        idx_sku = header.index("SKU")
    except ValueError:
        print("PERINGATAN: kolom Barcode_Produk/SKU tidak ditemukan di tabel Produk, lookup dilewati.")
        return {}

    lookup = {}
    for row in rows[1:]:
        if len(row) > max(idx_barcode, idx_sku):
            sku = str(row[idx_sku]).strip()
            barcode = str(row[idx_barcode]).strip()
            if sku:
                lookup[sku] = barcode
    return lookup


def build_sku_to_rak_lookup(sheets_service, spreadsheet_id):
    """Baca tabel Stok, bikin mapping SKU -> Kode_Rak (lokasi barang di gudang).
    Kalau 1 SKU ada di beberapa rak, dipakai kemunculan pertama saja."""
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"{SHEET_STOK}!A:E"
    ).execute()
    rows = result.get("values", [])
    if not rows:
        return {}
    header = rows[0]
    try:
        idx_sku = header.index("SKU")
        idx_rak = header.index("Kode_Rak")
    except ValueError:
        print("PERINGATAN: kolom SKU/Kode_Rak tidak ditemukan di tabel Stok, lookup dilewati.")
        return {}

    lookup = {}
    for row in rows[1:]:
        if len(row) > max(idx_sku, idx_rak):
            sku = str(row[idx_sku]).strip()
            rak = str(row[idx_rak]).strip()
            if sku and sku not in lookup:  # ambil kemunculan pertama saja
                lookup[sku] = rak
    return lookup



def get_existing_awb_pesanan(sheets_service, spreadsheet_id):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{SHEET_PESANAN}!A:B"
    ).execute()

    rows = result.get("values", [])
    existing = set()

    for row in rows[1:]:
        if len(row) >= 2:
            if row[0]:
                existing.add(("AWB", str(row[0]).strip()))
            if row[1]:
                existing.add(("PESANAN", str(row[1]).strip()))

    return existing


def filter_duplicate_results(sheets_service, spreadsheet_id, results):
    existing = get_existing_awb_pesanan(
        sheets_service,
        spreadsheet_id
    )

    filtered = []

    for data in results:
        awb = str(data.get("no_awb") or "").strip()
        pesanan = str(data.get("no_pesanan") or "").strip()

        if ("AWB", awb) in existing or ("PESANAN", pesanan) in existing:
            print(f"SKIP DUPLICATE: {awb} / {pesanan}")
            continue

        filtered.append(data)
        existing.add(("AWB", awb))

        if pesanan:
            existing.add(("PESANAN", pesanan))

    return filtered


def write_to_sheets(sheets_service, spreadsheet_id, results):
    sku_to_barcode = build_sku_to_barcode_lookup(sheets_service, spreadsheet_id)
    sku_to_rak = build_sku_to_rak_lookup(sheets_service, spreadsheet_id)

    pesanan_rows, detail_rows = [], []
    counter = 0
    for data in results:
        # Urutan kolom Pesanan: No_AWB | No_Pesanan | ID_Shopify | Tanggal | Picker | Status
        # ID_Shopify dikosongkan (bukan hasil ekstraksi PDF, diisi manual/integrasi lain kalau ada)
        pesanan_rows.append([data["no_awb"], data["no_pesanan"] or "", "", "", "", "Pending"])
        for item in data["items"]:
            counter += 1
            sku = item["sku"]
            barcode = sku_to_barcode.get(sku)
            if not barcode:
                print(f"PERINGATAN: SKU {sku} tidak ditemukan di tabel Produk, dipakai apa adanya.")
                barcode = sku
            kode_rak = sku_to_rak.get(sku, "")
            if not kode_rak:
                print(f"PERINGATAN: SKU {sku} tidak ditemukan di tabel Stok, Kode_Rak dikosongkan.")

            id_detail = f"DT-{data['no_awb']}-{counter}"
            detail_rows.append(
                [id_detail, data["no_awb"], barcode, kode_rak, item["qty"], 0, "", "Belum"]
            )

    if pesanan_rows:
        next_row = _get_next_empty_row(sheets_service, spreadsheet_id, SHEET_PESANAN)
        last_row = next_row + len(pesanan_rows) - 1
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_PESANAN}!A{next_row}:F{last_row}",
            valueInputOption="RAW",
            body={"values": pesanan_rows},
        ).execute()

    if detail_rows:
        next_row = _get_next_empty_row(sheets_service, spreadsheet_id, SHEET_DETAIL)
        last_row = next_row + len(detail_rows) - 1
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_DETAIL}!A{next_row}:H{last_row}",
            valueInputOption="RAW",
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
            results = filter_duplicate_results(
                sheets_service,
                spreadsheet_id,
                results
            )

            if not results:
                print("Semua data duplicate, tidak ditulis.")
                move_file(drive_service, f["id"], processed_folder_id, folder_id)
                continue

            n_pesanan, n_detail = write_to_sheets(sheets_service, spreadsheet_id, results)
            print(f"  -> {len(results)} AWB ditemukan, {n_pesanan} baris Pesanan, {n_detail} baris Detail")
            for r in results:
                if r.get("warning"):
                    print(f"  ⚠ PERINGATAN untuk No_AWB {r['no_awb']}: {r['warning']}")
            move_file(drive_service, f["id"], processed_folder_id, folder_id)
            print(f"  -> dipindah ke folder Processed")
        except Exception as e:
            print(f"  GAGAL memproses {f['name']}: {e}")


if __name__ == "__main__":
    main()

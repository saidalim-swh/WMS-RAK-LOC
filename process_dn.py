"""
process_dn.py

Khusus memproses PDF Delivery Note (DN) Lemonilo.
Tidak mengganggu process_awb.py.

Perubahan terbaru:
- Item DN sekarang dibaca berdasarkan POSISI KOLOM & BARIS tabel, bukan regex
  teks. Versi lama menangkap kode "IFD..." dari kolom LOT (mis.
  "IFDBC00013-SPR/HRC (5)") sebagai item, lalu mengambil qty dari baris
  berikutnya, sehingga muncul baris "IFDBC00013-SPR/HRC" dan item lain hilang.
- Kode SKU dipotong di tanda "-" atau "/" pertama, jadi suffix gudang seperti
  "-SPR/HRC" otomatis terbuang: "IFDBC00013-SPR/HRC" -> "IFDBC00013".
- Mendukung banyak DN dalam satu PDF (satu halaman satu DN, beberapa DN
  bertumpuk dalam satu halaman, maupun tabel yang bersambung ke halaman lain).
- Cek duplikat sebelum menulis ke sheet.

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

# Kode SKU = "IFD" + huruf/angka, BERHENTI di "-" atau "/" pertama.
# Jadi "IFDBC00013-SPR/HRC" -> "IFDBC00013" (suffix gudang terbuang).
SKU_RE = re.compile(r"^(IFD[A-Z0-9]+)")


# --------------------------------------------------------------------------
# Google Drive / Sheets
# --------------------------------------------------------------------------

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


# --------------------------------------------------------------------------
# Ekstraksi PDF
# --------------------------------------------------------------------------

def _cx(w):
    return (w["x0"] + w["x1"]) / 2


def _cy(w):
    return (w["top"] + w["bottom"]) / 2


def clean_sku(raw):
    """'IFDBC00013-SPR/HRC' -> 'IFDBC00013'. None kalau bukan kode IFD."""
    m = SKU_RE.match(str(raw).strip().upper())
    return m.group(1) if m else None


def find_dn_markers(words):
    """
    Cari semua penanda "DN NO <nomor>" di halaman berdasarkan posisi kata.
    Return list of (top, dn_no), urut dari atas ke bawah.
    """
    markers = []
    for w in words:
        if w["text"].upper() != "DN":
            continue

        # kata "NO" tepat di kanan "DN", di baris yang sama
        no_candidates = [
            v for v in words
            if v["text"].upper().rstrip(":") == "NO"
            and abs(v["top"] - w["top"]) <= 3
            and 0 <= v["x0"] - w["x1"] <= 15
        ]
        if not no_candidates:
            continue
        no_w = min(no_candidates, key=lambda v: v["x0"])

        # nilai nomor DN: kata terdekat di kanan "NO", di baris yang sama
        value_candidates = [
            v for v in words
            if abs(v["top"] - w["top"]) <= 3
            and v["x0"] >= no_w["x1"] - 1
            and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-]*", v["text"])
        ]
        if not value_candidates:
            continue
        value = min(value_candidates, key=lambda v: v["x0"])["text"].strip()
        markers.append((w["top"], value))

    markers.sort(key=lambda m: m[0])
    return markers


def extract_items_from_words(words):
    """
    Baca item tabel DN berdasarkan POSISI kolom/baris.

    Cara kerja:
    - Header tabel dikenali dari kata "QTY" dan "UOM" yang sebaris.
    - Kolom ITEM ID = kolom PALING KIRI yang berisi kode "IFD...". (Kode IFD
      juga muncul di kolom ITEM NAME dan LOT, tapi keduanya ada di sebelah
      kanan, jadi otomatis diabaikan.)
    - Tiap kode di kolom ITEM ID jadi "jangkar" satu baris. Qty & UOM diambil
      dari kolom QTY/UOM yang posisi vertikalnya ada di antara jangkar ini dan
      jangkar baris berikutnya. Ini aman walau ITEM ID wrap dua baris dan
      qty/UOM berada di tengah baris.
    Return list of {"barcode", "qty", "uom"}.
    """
    # --- header QTY + UOM yang sebaris ---
    qty_hdrs = [w for w in words if w["text"].upper() == "QTY"]
    uom_hdrs = [w for w in words if w["text"].upper() == "UOM"]
    header = None
    for q in sorted(qty_hdrs, key=lambda w: w["top"]):
        for u in uom_hdrs:
            if abs(q["top"] - u["top"]) <= 4 and u["x0"] > q["x0"]:
                header = (q, u)
                break
        if header:
            break
    if not header:
        return []

    q_hdr, u_hdr = header
    header_bottom = max(q_hdr["bottom"], u_hdr["bottom"])
    qty_cx, uom_cx = _cx(q_hdr), _cx(u_hdr)
    col_tol = max(15.0, abs(uom_cx - qty_cx) * 0.45)

    region = [w for w in words if w["top"] >= header_bottom - 1]

    # batas bawah tabel: kata "MEMO" (kalau ada)
    memo_tops = [w["top"] for w in region if w["text"].upper().startswith("MEMO")]
    table_end = min(memo_tops) if memo_tops else float("inf")
    region = [w for w in region if w["top"] < table_end]

    # --- jangkar baris: kode IFD di kolom paling kiri ---
    ifd_words = [w for w in region if SKU_RE.match(w["text"].upper())]
    if not ifd_words:
        return []
    col_x = min(w["x0"] for w in ifd_words)
    anchors = sorted(
        [w for w in ifd_words if abs(w["x0"] - col_x) <= 8],
        key=lambda w: w["top"],
    )

    items = []
    for i, a in enumerate(anchors):
        y0 = a["top"] - 3
        y1 = anchors[i + 1]["top"] - 3 if i + 1 < len(anchors) else float("inf")
        sku = clean_sku(a["text"])

        qty_cands = [
            w for w in region
            if re.fullmatch(r"\d+(?:[.,]\d+)?", w["text"])
            and abs(_cx(w) - qty_cx) <= col_tol
            and y0 <= _cy(w) < y1
        ]
        if not qty_cands:
            print(f"PERINGATAN: Qty untuk {sku} tidak ketemu di kolom QTY, baris dilewati.")
            continue
        qty_w = min(qty_cands, key=_cy)
        qty = int(float(qty_w["text"].replace(",", ".")))

        uom_cands = [
            w for w in region
            if re.fullmatch(r"[A-Za-z]{2,6}", w["text"])
            and abs(_cx(w) - uom_cx) <= col_tol
            and y0 <= _cy(w) < y1
        ]
        if uom_cands:
            uom = min(uom_cands, key=_cy)["text"].upper()
        else:
            uom = ""
            print(f"PERINGATAN: UOM untuk {sku} tidak ketemu, dikosongkan.")

        items.append({"barcode": sku, "qty": qty, "uom": uom})

    return items


def extract_items_from_text(text):
    """
    CADANGAN (hanya dipakai kalau header tabel tidak bisa dikenali lewat posisi).
    Kurang akurat dibanding pembacaan kolom.
    """
    print("PERINGATAN: header tabel tidak terbaca, memakai cara teks (kurang akurat).")
    items = []
    for item_id, qty, uom in re.findall(
        r"(IFD[A-Z0-9]+)[^\n]*?\b(\d+)\s+(PCS|CTN|BOX)\b",
        text,
        re.IGNORECASE,
    ):
        items.append({"barcode": clean_sku(item_id), "qty": int(qty), "uom": uom.upper()})
    return items


def merge_items(existing, new):
    """Gabungkan list item; qty dijumlahkan kalau (barcode, uom) sama."""
    index = {(it["barcode"], it["uom"]): it for it in existing}
    merged = list(existing)
    for it in new:
        key = (it["barcode"], it["uom"])
        if key in index:
            index[key]["qty"] += it["qty"]
        else:
            copy = dict(it)
            index[key] = copy
            merged.append(copy)
    return merged


def extract_dn(pdf_bytes):
    """
    Ekstrak SEMUA DN dalam satu PDF. Return list of dict:
    {"no_awb", "no_pesanan", "items"}. Bisa kosong kalau bukan PDF DN.
    """
    results = []
    current = None

    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            markers = find_dn_markers(words)

            if markers:
                # Satu halaman bisa memuat >1 DN (bertumpuk). Tiap DN
                # memegang kata-kata dari posisinya sampai DN berikutnya.
                for k, (top, dn_no) in enumerate(markers):
                    bottom = markers[k + 1][0] if k + 1 < len(markers) else float("inf")
                    band = [w for w in words if top - 1 <= w["top"] < bottom - 1]
                    items = extract_items_from_words(band)

                    if not items and len(markers) == 1:
                        items = extract_items_from_text(page.extract_text() or "")

                    # DN yang sama muncul lagi (mis. halaman lanjutan yang
                    # mengulang "DN NO") -> gabungkan, jangan buat entry baru.
                    if current is not None and current["no_awb"] == dn_no and k == 0:
                        current["items"] = merge_items(current["items"], items)
                    else:
                        current = {
                            "no_awb": dn_no,
                            "no_pesanan": dn_no,
                            "items": merge_items([], items),
                        }
                        results.append(current)
            elif current is not None:
                # Halaman tanpa "DN NO": sambungan tabel DN sebelumnya.
                items = extract_items_from_words(words)
                if items:
                    current["items"] = merge_items(current["items"], items)

    return results


# --------------------------------------------------------------------------
# Tulis ke Google Sheets
# --------------------------------------------------------------------------

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

            for data in data_list:
                if not data["items"]:
                    print(f"  PERINGATAN: DN {data['no_awb']} TIDAK ADA item yang terbaca, cek manual.")

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

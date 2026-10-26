# ==============================================================================
# APLIKASI SMART GENERATOR LPJ DWP KPPBC TMC MALANG (V3.0 - PRODUCTION)
# Arsitektur: Template-Cloning Engine + Smart Financial 3-Mode + Google Sheets
# Aturan Koding Mutlak: 0 Backslash Character di Seluruh Baris Kode Python
# ==============================================================================

import os
import io
import json
import uuid
import base64
import shutil
import tempfile
import subprocess
from datetime import date, datetime

import pandas as pd
import streamlit as st
from PIL import Image

import docx
from docx.shared import Cm, Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except Exception:
    HAS_GSPREAD = False

# ==============================================================================
# BAGIAN 1: KONSTANTA, ASET BAWAAN & KONFIGURASI HALAMAN
# ==============================================================================

st.set_page_config(
    page_title="Smart LPJ DWP KPPBC TMC Malang",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

NL = chr(10)
TAB = chr(9)
BS = chr(92)

DEFAULT_SPREADSHEET_ID = "1_I2f5Gb_F2hMCIIpwwjLF_LpWeCSlWc_E60rYh9hTxc"
SHEET_DB_LPJ = "DATABASE_LPJ"
SHEET_REKAP_DANA = "REKAP_ITEM_DANA"

COLS_DB_LPJ = [
    "ID_LPJ",
    "WAKTU_SIMPAN",
    "TGL_LAPORAN",
    "NAMA_KEGIATAN",
    "TEMA_KEGIATAN",
    "KONTEKS_HUT",
    "TGL_ACARA",
    "WAKTU_ACARA",
    "TEMPAT_ACARA",
    "SUMBER_DANA",
    "KET_TABEL_DANA",
    "TOTAL_DANA",
    "PENANDATANGAN",
    "PAYLOAD_JSON",
]

COLS_REKAP_DANA = [
    "ID_LPJ",
    "TGL_ACARA",
    "NAMA_KEGIATAN",
    "NO_ITEM",
    "KEBUTUHAN_KEGIATAN",
    "JUMLAH",
    "HARGA_SATUAN",
    "HARGA_TOTAL",
    "KETERANGAN",
]

TEMPLATE_FILENAME = "LPJ_DWP GABUNGAN DI BLITAR_2026-1.docx"

EMBEDDED_LOGO1_B64 = "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"
EMBEDDED_LOGO2_B64 = "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"
EMBEDDED_SIG_B64 = "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"

DEFAULT_UCAPAN_DONATUR = (
    "Kami mengucapkan terima kasih yang sebesar-besarnya kepada para donatur "
    "yang telah menyisihkan sedikit rezekinya untuk kegiatan ini sehingga dapat "
    "berjalan dengan sangat baik. Semoga bantuan dari Bapak/Ibu dan Keluarga "
    "dihitung Allah SWT sebagai amal kebaikan yang bernilai pahala, bermanfaat "
    "untuk sesama, dan menjadi berkah untuk kita semua. Selain itu, semoga "
    "dengan kegiatan ini semakin memperkuat tali silaturahmi dan kekeluargaan kita."
)

PRESET_BLITAR_2026 = {
    "id_lpj": "LPJ-20260813-BLITAR",
    "tgl_laporan": date(2026, 8, 13),
    "nama_kegiatan": "Pertemuan Rutin Paguyuban Dharma Wanita Persatuan DJBC di lingkungan Kanwil DJBC Jawa Timur II",
    "tema_kegiatan": "Kebersamaan dalam Keceriaan, Bersinergi dalam Kebahagiaan",
    "konteks_hut": "menyambut HUT ke-81 Kemerdekaan Republik Indonesia",
    "tgl_acara": date(2026, 8, 10),
    "tgl_acara_manual": "",
    "waktu_acara": "08.30 - selesai",
    "tempat_acara": "KPPBC Tipe Madya Pabean C Blitar",
    "rangkaian_kegiatan": NL.join([
        "Pembukaan acara dengan sambutan dari Pembina DWP BC Blitar Bapak Nurtjahjo, Ketua DWP BC Blitar Ibu Hasrifah Nurtjahjo, dan Ibu Ketua DWP Kanwil DJBC Jatim II Ibu Ana Lukman.",
        "Kegiatan lomba menghias tumpeng dan lomba sambung lirik lagu Nasional untuk menyambut HUT ke-81 Kemerdekaan Republik Indonesia.",
        "Kunjungan ke Museum Bung Karno.",
    ]),
    "sumber_dana": "Kas Dharma Wanita KPPBC TMC Malang dan donatur",
    "ket_tabel_dana": "Kas DWP BC Malang & Donatur Anggota",
    "merge_ket_vertikal": True,
    "tampilkan_ucapan_donatur": True,
    "teks_ucapan_donatur": DEFAULT_UCAPAN_DONATUR,
    "catatan_ig": "*Note : dokumentasi kegiatan dapat dilihat di Instagram @dwpbcmalang",
    "jabatan_baris_1": "Ketua Dharma Wanita Persatuan",
    "jabatan_baris_2": "KPPBC Tipe Madya Cukai Malang",
    "nama_penandatangan": "Ny. Rini Johan Pandores",
    "tampilkan_ttd": True,
    "items_dana": [
        {"NO": 1, "KEBUTUHAN KEGIATAN": "Iuran DWP Gabungan di Blitar", "JUMLAH": 10, "HARGA SATUAN": 0, "HARGA TOTAL": 675000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 2, "KEBUTUHAN KEGIATAN": "Doorprize", "JUMLAH": 10, "HARGA SATUAN": 0, "HARGA TOTAL": 0, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 3, "KEBUTUHAN KEGIATAN": "Mobil", "JUMLAH": 2, "HARGA SATUAN": 0, "HARGA TOTAL": 0, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 4, "KEBUTUHAN KEGIATAN": "Hadiah Peserta Lomba Tumpeng", "JUMLAH": 3, "HARGA SATUAN": 250000, "HARGA TOTAL": 750000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 5, "KEBUTUHAN KEGIATAN": "Kerudung Merah", "JUMLAH": 10, "HARGA SATUAN": 50000, "HARGA TOTAL": 500000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 6, "KEBUTUHAN KEGIATAN": "Baju Atasan Merah (Blouse)", "JUMLAH": 9, "HARGA SATUAN": 90000, "HARGA TOTAL": 810000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 7, "KEBUTUHAN KEGIATAN": "Sarapan Pagi (Roti)", "JUMLAH": 20, "HARGA SATUAN": 13000, "HARGA TOTAL": 260000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 8, "KEBUTUHAN KEGIATAN": "Snack (Keripik Usus)", "JUMLAH": 12, "HARGA SATUAN": 15000, "HARGA TOTAL": 180000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 9, "KEBUTUHAN KEGIATAN": "Makan Malam (Ayam Pemuda)", "JUMLAH": 13, "HARGA SATUAN": 100000, "HARGA TOTAL": 1300000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 10, "KEBUTUHAN KEGIATAN": "Bahan Latihan Tumpeng, Support Makan Siang & Kopi Untuk Latihan", "JUMLAH": 3, "HARGA SATUAN": 200000, "HARGA TOTAL": 600000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
    ],
}


# ==============================================================================
# BAGIAN 2: HELPER FUNGSI BAHASA & KEUANGAN INDONESIA
# ==============================================================================

def parse_num(val) -> int:
    """Mengubah nilai string format Indonesia (Rp 250.000, 13rb, -) menjadi integer murni."""
    if val is None:
        return 0
    if isinstance(val, bool):
        return 0
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return 0
        return int(round(float(val)))

    s = str(val).strip().lower()
    if not s or s in ("-", "rp -", "rp-", "nihil", "gratis", "free", "nan", "none"):
        return 0

    s = s.replace("rp", "").replace("idr", "").replace(TAB, "").strip()

    multiplier = 1
    if "juta" in s or "jt" in s:
        multiplier = 1000000
        s = s.replace("juta", "").replace("jt", "").strip()
    elif "ribu" in s or "rb" in s:
        multiplier = 1000
        s = s.replace("ribu", "").replace("rb", "").strip()

    if multiplier > 1:
        s_clean = s.replace(".", "").replace(",", ".")
        digits_dot = "".join(ch for ch in s_clean if ch.isdigit() or ch == ".")
        if not digits_dot:
            return 0
        try:
            return int(round(float(digits_dot) * multiplier))
        except Exception:
            return 0

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.split(",")[0]
    elif "," in s and len(s.split(",")[-1]) <= 2 and len(s.split(",")[0]) > 3:
        s = s.split(",")[0]

    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return 0
    return int(digits)


def fmt_ribuan(val) -> str:
    """Mengubah integer menjadi string titik ribuan Indonesia (5075000 -> 5.075.000)."""
    n = parse_num(val)
    if n < 0:
        return "-" + f"{abs(n):,}".replace(",", ".")
    return f"{n:,}".replace(",", ".")


def fmt_rp_cell(val, is_dash_if_zero: bool = True) -> str:
    """Format sel Rupiah tabel LPJ tanpa karakter Tab liar agar anti-potong di PDF."""
    n = parse_num(val)
    if n <= 0:
        return "Rp -" if is_dash_if_zero else "Rp 0"
    return "Rp " + fmt_ribuan(n)


def _terbilang_dasar(n: int) -> str:
    satuan = [
        "", "satu", "dua", "tiga", "empat", "lima",
        "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"
    ]
    if n < 12:
        return satuan[n]
    if n < 20:
        return _terbilang_dasar(n - 10) + " belas"
    if n < 100:
        hasil = _terbilang_dasar(n // 10) + " puluh"
        if n % 10 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 10)
        return hasil
    if n < 200:
        sisa = _terbilang_dasar(n - 100)
        return ("seratus " + sisa).strip()
    if n < 1000:
        hasil = _terbilang_dasar(n // 100) + " ratus"
        if n % 100 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 100)
        return hasil
    if n < 2000:
        sisa = _terbilang_dasar(n - 1000)
        return ("seribu " + sisa).strip()
    if n < 1000000:
        hasil = _terbilang_dasar(n // 1000) + " ribu"
        if n % 1000 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 1000)
        return hasil
    if n < 1000000000:
        hasil = _terbilang_dasar(n // 1000000) + " juta"
        if n % 1000000 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 1000000)
        return hasil
    if n < 1000000000000:
        hasil = _terbilang_dasar(n // 1000000000) + " miliar"
        if n % 1000000000 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 1000000000)
        return hasil
    return str(n)


def terbilang(n) -> str:
    """Mengubah nominal angka menjadi teks Terbilang Rupiah Title Case."""
    angka = parse_num(n)
    if angka <= 0:
        return "Nol Rupiah"
    teks = " ".join(_terbilang_dasar(angka).split())
    return teks.title() + " Rupiah"


def format_hari_tanggal_indo(dt) -> str:
    """Contoh: date(2026, 8, 13) -> Kamis, 13 Agustus 2026."""
    if isinstance(dt, str):
        s = dt.strip()
        if "," in s and any(b in s for b in ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]):
            return s
        try:
            dt = datetime.strptime(s[:10], "%Y-%m-%d").date()
        except Exception:
            return s
    if isinstance(dt, datetime):
        dt = dt.date()
    if not isinstance(dt, date):
        dt = date.today()

    nama_hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    nama_bulan = [
        "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    hari = nama_hari[dt.weekday()]
    bulan = nama_bulan[dt.month]
    return f"{hari}, {dt.day} {bulan} {dt.year}"


def normalize_items_df(records, default_ket: str = "Kas DWP BC Malang & Donatur Anggota") -> pd.DataFrame:
    """Menerapkan Logika 3-Mode secara konsisten pada seluruh baris rincian dana."""
    rows = []
    if isinstance(records, pd.DataFrame):
        raw_list = records.to_dict(orient="records")
    elif isinstance(records, list):
        raw_list = records
    else:
        raw_list = []

    idx_counter = 1
    for item in raw_list:
        kebutuhan = str(item.get("KEBUTUHAN KEGIATAN", item.get("kebutuhan", "")) or "").strip()
        if not kebutuhan and parse_num(item.get("HARGA TOTAL", 0)) == 0 and parse_num(item.get("HARGA SATUAN", 0)) == 0:
            continue

        jml = max(0, parse_num(item.get("JUMLAH", item.get("jumlah", 1))))
        h_sat = max(0, parse_num(item.get("HARGA SATUAN", item.get("harga_satuan", 0))))
        h_tot_in = max(0, parse_num(item.get("HARGA TOTAL", item.get("harga_total", 0))))
        ket = str(item.get("KETERANGAN", item.get("keterangan", default_ket)) or default_ket).strip()

        # Logika 3-Mode Smart Financial Engine:
        # Mode 1 (Normal): Harga Satuan > 0 -> Harga Total = Jumlah * Harga Satuan
        # Mode 2 (Lump Sum / Borongan): Harga Satuan == 0 dan Harga Total > 0 -> Harga Total tetap
        # Mode 3 (In-Kind / Fasilitas): Harga Satuan == 0 dan Harga Total == 0 -> Rp -
        if h_sat > 0:
            qty_calc = jml if jml > 0 else 1
            h_tot = qty_calc * h_sat
        elif h_tot_in > 0:
            h_tot = h_tot_in
        else:
            h_tot = 0

        rows.append({
            "NO": idx_counter,
            "KEBUTUHAN KEGIATAN": kebutuhan,
            "JUMLAH": jml,
            "HARGA SATUAN": h_sat,
            "HARGA TOTAL": h_tot,
            "KETERANGAN": ket,
        })
        idx_counter += 1

    if not rows:
        rows.append({
            "NO": 1,
            "KEBUTUHAN KEGIATAN": "Kebutuhan Kegiatan",
            "JUMLAH": 1,
            "HARGA SATUAN": 0,
            "HARGA TOTAL": 0,
            "KETERANGAN": default_ket,
        })

    return pd.DataFrame(rows)


def smart_paste_parser(raw_text: str, default_ket: str) -> list:
    """Mengekstrak daftar belanja dari teks WhatsApp/Catatan ke format 3-Mode secara cerdas."""
    parsed = []
    if not raw_text or not raw_text.strip():
        return parsed

    lines = [ln.strip() for ln in raw_text.split(NL) if ln.strip()]
    for ln in lines:
        # Buang nomor urut di depan jika ada (misal: "1. Roti | 20 | 13000")
        clean_ln = ln
        if "." in clean_ln[:4]:
            left_num, rest = clean_ln.split(".", 1)
            if left_num.strip().isdigit():
                clean_ln = rest.strip()

        # Coba format pemisah pipa (|) atau titik koma (;)
        if "|" in clean_ln or ";" in clean_ln:
            sep = "|" if "|" in clean_ln else ";"
            parts = [p.strip() for p in clean_ln.split(sep)]
            nama = parts[0] if len(parts) > 0 else "Item Kegiatan"
            jml = parse_num(parts[1]) if len(parts) > 1 else 1
            val_3 = parse_num(parts[2]) if len(parts) > 2 else 0
            val_4 = parse_num(parts[3]) if len(parts) > 3 else 0
            ket = parts[4] if len(parts) > 4 and parts[4] else default_ket

            if len(parts) >= 4 and val_3 == 0 and val_4 > 0:
                h_sat, h_tot = 0, val_4
            elif "total" in (parts[2].lower() if len(parts) > 2 else "") or "borongan" in clean_ln.lower():
                h_sat, h_tot = 0, val_3
            else:
                h_sat = val_3
                h_tot = (jml if jml > 0 else 1) * h_sat if h_sat > 0 else 0

            parsed.append({
                "KEBUTUHAN KEGIATAN": nama,
                "JUMLAH": jml if jml > 0 else 1,
                "HARGA SATUAN": h_sat,
                "HARGA TOTAL": h_tot,
                "KETERANGAN": ket,
            })
            continue

        # Coba deteksi pola natural WhatsApp:
        # 1) "Roti 20 x 13.000" atau "Roti 20 @ 13000"
        lower_ln = clean_ln.lower()
        found_mult = False
        for sym in [" x ", " @ ", " * "]:
            if sym in lower_ln:
                idx_sym = lower_ln.rfind(sym)
                left_part = clean_ln[:idx_sym].strip()
                right_part = clean_ln[idx_sym + len(sym):].strip()
                h_sat = parse_num(right_part)
                tokens = left_part.split()
                if tokens and tokens[-1].isdigit():
                    jml = int(tokens[-1])
                    nama = " ".join(tokens[:-1]).strip()
                else:
                    jml = 1
                    nama = left_part
                parsed.append({
                    "KEBUTUHAN KEGIATAN": nama or "Item Kegiatan",
                    "JUMLAH": jml,
                    "HARGA SATUAN": h_sat,
                    "HARGA TOTAL": jml * h_sat,
                    "KETERANGAN": default_ket,
                })
                found_mult = True
                break
        if found_mult:
            continue

        # 2) "Iuran Blitar 10 total 675.000"
        if " total " in lower_ln:
            idx_tot = lower_ln.rfind(" total ")
            left_part = clean_ln[:idx_tot].strip()
            right_part = clean_ln[idx_tot + 7:].strip()
            h_tot = parse_num(right_part)
            tokens = left_part.split()
            if tokens and tokens[-1].isdigit():
                jml = int(tokens[-1])
                nama = " ".join(tokens[:-1]).strip()
            else:
                jml = 1
                nama = left_part
            parsed.append({
                "KEBUTUHAN KEGIATAN": nama or "Item Borongan",
                "JUMLAH": jml,
                "HARGA SATUAN": 0,
                "HARGA TOTAL": h_tot,
                "KETERANGAN": default_ket,
            })
            continue

        # 3) "Mobil 2 -" atau "Doorprize 10"
        tokens = clean_ln.split()
        if len(tokens) >= 2:
            last_tok = tokens[-1]
            second_last = tokens[-2] if len(tokens) >= 2 else ""
            if last_tok in ("-", "0", "nihil") and second_last.isdigit():
                jml = int(second_last)
                nama = " ".join(tokens[:-2]).strip()
                parsed.append({
                    "KEBUTUHAN KEGIATAN": nama or "Fasilitas",
                    "JUMLAH": jml,
                    "HARGA SATUAN": 0,
                    "HARGA TOTAL": 0,
                    "KETERANGAN": default_ket,
                })
                continue
            elif last_tok.isdigit() and int(last_tok) < 500:
                jml = int(last_tok)
                nama = " ".join(tokens[:-1]).strip()
                parsed.append({
                    "KEBUTUHAN KEGIATAN": nama or "Fasilitas",
                    "JUMLAH": jml,
                    "HARGA SATUAN": 0,
                    "HARGA TOTAL": 0,
                    "KETERANGAN": default_ket,
                })
                continue
            else:
                nominal = parse_num(last_tok)
                if second_last.isdigit() and int(second_last) < 1000 and nominal > 0:
                    jml = int(second_last)
                    nama = " ".join(tokens[:-2]).strip()
                    parsed.append({
                        "KEBUTUHAN KEGIATAN": nama or "Item Kegiatan",
                        "JUMLAH": jml,
                        "HARGA SATUAN": nominal,
                        "HARGA TOTAL": jml * nominal,
                        "KETERANGAN": default_ket,
                    })
                    continue

        parsed.append({
            "KEBUTUHAN KEGIATAN": clean_ln,
            "JUMLAH": 1,
            "HARGA SATUAN": 0,
            "HARGA TOTAL": 0,
            "KETERANGAN": default_ket,
        })

    return parsed


def build_paragraf_pembuka(nama_kegiatan: str, tema_kegiatan: str, konteks_hut: str) -> str:
    """Menyusun paragraf pembuka adaptif sesuai ketersediaan Tema dan Konteks HUT."""
    nama = (nama_kegiatan or "").strip() or "Kegiatan Dharma Wanita Persatuan KPPBC TMC Malang"
    tema = (tema_kegiatan or "").strip().strip('"')
    hut = (konteks_hut or "").strip()

    kalimat_1 = (
        "Puji syukur kehadirat Allah SWT yang telah melimpahkan rahmat dan rizki-Nya "
        f"kepada kita semua, yang mana telah terlaksana {nama}."
    )
    if tema and hut:
        kalimat_2 = f' Kegiatan ini mengusung tema "{tema}" serta {hut}, yang dilaksanakan pada :'
    elif tema:
        kalimat_2 = f' Kegiatan ini mengusung tema "{tema}", yang dilaksanakan pada :'
    elif hut:
        kalimat_2 = f" Kegiatan ini dilaksanakan dalam rangka {hut}, pada :"
    else:
        kalimat_2 = " Kegiatan tersebut dilaksanakan pada :"
    return kalimat_1 + kalimat_2


# ==============================================================================
# BAGIAN 3: KONEKSI GOOGLE SHEETS API & OFFLINE RESILIENCE
# ==============================================================================

@st.cache_resource(show_spinner=False)
def get_gspread_client(creds_json_str: str = ""):
    """Membuka koneksi gspread dengan dukungan st.secrets atau JSON Service Account."""
    if not HAS_GSPREAD:
        return None, "Library gspread belum terpasang di environment ini."

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = None

    if creds_json_str and creds_json_str.strip():
        try:
            creds_dict = json.loads(creds_json_str)
        except Exception as e:
            return None, f"Format JSON Service Account tidak valid: {e}"
    else:
        try:
            if "gcp_service_account" in st.secrets:
                creds_dict = dict(st.secrets["gcp_service_account"])
        except Exception:
            creds_dict = None

    if not creds_dict:
        return None, "Mode Lokal Aktif (Kredensial Google Sheets belum dihubungkan)."

    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client, "Terhubung ke Google Sheets API"
    except Exception as e:
        return None, f"Gagal otentikasi Google Sheets: {e}"


def ensure_worksheets_ready(sh):
    """Memastikan worksheet DATABASE_LPJ dan REKAP_ITEM_DANA tersedia beserta header-nya."""
    existing_titles = [ws.title for ws in sh.worksheets()]

    if SHEET_DB_LPJ not in existing_titles:
        ws_db = sh.add_worksheet(title=SHEET_DB_LPJ, rows=500, cols=len(COLS_DB_LPJ))
        ws_db.batch_update([{"range": "A1:N1", "values": [COLS_DB_LPJ]}], value_input_option="USER_ENTERED")
    else:
        ws_db = sh.worksheet(SHEET_DB_LPJ)

    if SHEET_REKAP_DANA not in existing_titles:
        ws_rk = sh.add_worksheet(title=SHEET_REKAP_DANA, rows=2000, cols=len(COLS_REKAP_DANA))
        ws_rk.batch_update([{"range": "A1:I1", "values": [COLS_REKAP_DANA]}], value_input_option="USER_ENTERED")
    else:
        ws_rk = sh.worksheet(SHEET_REKAP_DANA)

    return ws_db, ws_rk


@st.cache_data(ttl=180, show_spinner=False)
def fetch_sheet_bundle_cached(spreadsheet_id: str, creds_json_str: str = "") -> dict:
    """Membaca DATABASE_LPJ dan REKAP_ITEM_DANA sekaligus dengan values_batch_get hemat kuota."""
    client, status_msg = get_gspread_client(creds_json_str)
    if client is None:
        return {"ok": False, "msg": status_msg, "db_rows": [], "rekap_rows": []}

    try:
        sh = client.open_by_key(spreadsheet_id)
        ensure_worksheets_ready(sh)
        batch_res = sh.values_batch_get([f"{SHEET_DB_LPJ}!A1:N500", f"{SHEET_REKAP_DANA}!A1:I2000"])
        value_ranges = batch_res.get("valueRanges", [])

        db_vals = value_ranges[0].get("values", []) if len(value_ranges) > 0 else []
        rk_vals = value_ranges[1].get("values", []) if len(value_ranges) > 1 else []

        def _to_records(vals, expected_cols):
            if not vals or len(vals) < 2:
                return []
            headers = vals[0]
            out = []
            for r in vals[1:]:
                if not any(str(c).strip() for c in r):
                    continue
                row_dict = {}
                for i, col_name in enumerate(expected_cols):
                    col_idx = headers.index(col_name) if col_name in headers else i
                    row_dict[col_name] = r[col_idx] if col_idx < len(r) else ""
                out.append(row_dict)
            return out

        return {
            "ok": True,
            "msg": "Data berhasil disinkronkan dari Google Sheets.",
            "db_rows": _to_records(db_vals, COLS_DB_LPJ),
            "rekap_rows": _to_records(rk_vals, COLS_REKAP_DANA),
        }
    except Exception as e:
        return {"ok": False, "msg": f"Gagal membaca Google Sheets: {e}", "db_rows": [], "rekap_rows": []}


def get_sheet_bundle(spreadsheet_id: str, creds_json_str: str = "") -> dict:
    """Mengambil data Google Sheets dengan fallback ke _backup_sheet_bundle & session_state lokal."""
    res = fetch_sheet_bundle_cached(spreadsheet_id, creds_json_str)
    if res.get("ok"):
        st.session_state["_backup_sheet_bundle"] = res
        return res

    if "_backup_sheet_bundle" in st.session_state:
        backup = dict(st.session_state["_backup_sheet_bundle"])
        backup["msg"] = backup.get("msg", "") + " (Menggunakan cadangan memori lokal)"
        return backup

    return {
        "ok": False,
        "msg": res.get("msg", "Mode Offline Lokal Aktif"),
        "db_rows": st.session_state.get("local_db_lpj", []),
        "rekap_rows": st.session_state.get("local_rekap_dana", []),
    }


def save_lpj_to_storage(payload: dict, spreadsheet_id: str, creds_json_str: str = "") -> tuple:
    """Menyimpan LPJ ke Session State Lokal dan Google Sheets (via batch_update di luar loop)."""
    id_lpj = payload.get("id_lpj") or ("LPJ-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
    payload["id_lpj"] = id_lpj
    waktu_simpan = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    tgl_lap_str = str(payload.get("tgl_laporan", date.today()))
    tgl_acara_str = payload.get("tgl_acara_manual", "").strip() or format_hari_tanggal_indo(payload.get("tgl_acara", date.today()))

    df_items = normalize_items_df(payload.get("items_dana", []), payload.get("ket_tabel_dana", ""))
    total_dana = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

    serializable_payload = dict(payload)
    if isinstance(serializable_payload.get("tgl_laporan"), (date, datetime)):
        serializable_payload["tgl_laporan"] = serializable_payload["tgl_laporan"].strftime("%Y-%m-%d")
    if isinstance(serializable_payload.get("tgl_acara"), (date, datetime)):
        serializable_payload["tgl_acara"] = serializable_payload["tgl_acara"].strftime("%Y-%m-%d")
    serializable_payload["items_dana"] = df_items.to_dict(orient="records")

    db_row_dict = {
        "ID_LPJ": id_lpj,
        "WAKTU_SIMPAN": waktu_simpan,
        "TGL_LAPORAN": tgl_lap_str,
        "NAMA_KEGIATAN": payload.get("nama_kegiatan", ""),
        "TEMA_KEGIATAN": payload.get("tema_kegiatan", ""),
        "KONTEKS_HUT": payload.get("konteks_hut", ""),
        "TGL_ACARA": tgl_acara_str,
        "WAKTU_ACARA": payload.get("waktu_acara", ""),
        "TEMPAT_ACARA": payload.get("tempat_acara", ""),
        "SUMBER_DANA": payload.get("sumber_dana", ""),
        "KET_TABEL_DANA": payload.get("ket_tabel_dana", ""),
        "TOTAL_DANA": total_dana,
        "PENANDATANGAN": payload.get("nama_penandatangan", ""),
        "PAYLOAD_JSON": json.dumps(serializable_payload, ensure_ascii=False),
    }

    rekap_dicts = []
    for _, r in df_items.iterrows():
        rekap_dicts.append({
            "ID_LPJ": id_lpj,
            "TGL_ACARA": tgl_acara_str,
            "NAMA_KEGIATAN": payload.get("nama_kegiatan", ""),
            "NO_ITEM": int(r["NO"]),
            "KEBUTUHAN_KEGIATAN": str(r["KEBUTUHAN KEGIATAN"]),
            "JUMLAH": int(r["JUMLAH"]),
            "HARGA_SATUAN": int(r["HARGA SATUAN"]),
            "HARGA_TOTAL": int(r["HARGA TOTAL"]),
            "KETERANGAN": str(r["KETERANGAN"]),
        })

    # 1. Simpan ke Session State Lokal (Offline Resilience 100%)
    local_db = [row for row in st.session_state.get("local_db_lpj", []) if row.get("ID_LPJ") != id_lpj]
    local_db.insert(0, db_row_dict)
    st.session_state["local_db_lpj"] = local_db

    local_rk = [row for row in st.session_state.get("local_rekap_dana", []) if row.get("ID_LPJ") != id_lpj]
    local_rk.extend(rekap_dicts)
    st.session_state["local_rekap_dana"] = local_rk

    # 2. Jika Google Sheets terhubung, tulis sekaligus menggunakan batch_update di luar loop
    client, status_msg = get_gspread_client(creds_json_str)
    if client is None:
        return True, f"Tersimpan di Memori Lokal Aplikasi ({id_lpj}). {status_msg}"

    try:
        sh = client.open_by_key(spreadsheet_id)
        ws_db, ws_rk = ensure_worksheets_ready(sh)

        existing_db_vals = ws_db.get_all_values()
        new_db_matrix = [COLS_DB_LPJ]
        updated_existing = False

        for r in existing_db_vals[1:]:
            if not any(str(c).strip() for c in r):
                continue
            if len(r) > 0 and r[0] == id_lpj:
                new_db_matrix.append([db_row_dict[c] for c in COLS_DB_LPJ])
                updated_existing = True
            else:
                padded = list(r) + [""] * max(0, len(COLS_DB_LPJ) - len(r))
                new_db_matrix.append(padded[:len(COLS_DB_LPJ)])

        if not updated_existing:
            new_db_matrix.append([db_row_dict[c] for c in COLS_DB_LPJ])

        ws_db.clear()
        ws_db.batch_update(
            [{"range": f"A1:N{len(new_db_matrix)}", "values": new_db_matrix}],
            value_input_option="USER_ENTERED",
        )

        existing_rk_vals = ws_rk.get_all_values()
        new_rk_matrix = [COLS_REKAP_DANA]
        for r in existing_rk_vals[1:]:
            if not any(str(c).strip() for c in r):
                continue
            if len(r) > 0 and r[0] != id_lpj:
                padded = list(r) + [""] * max(0, len(COLS_REKAP_DANA) - len(r))
                new_rk_matrix.append(padded[:len(COLS_REKAP_DANA)])

        for item_d in rekap_dicts:
            new_rk_matrix.append([item_d[c] for c in COLS_REKAP_DANA])

        ws_rk.clear()
        ws_rk.batch_update(
            [{"range": f"A1:I{len(new_rk_matrix)}", "values": new_rk_matrix}],
            value_input_option="USER_ENTERED",
        )

        fetch_sheet_bundle_cached.clear()
        return True, f"Berhasil disimpan ke Google Sheets ({SHEET_DB_LPJ} & {SHEET_REKAP_DANA}) dengan ID: {id_lpj}"
    except Exception as e:
        return True, f"Tersimpan di Lokal ({id_lpj}), namun gagal sinkron ke GSheet: {e}"


# ==============================================================================
# BAGIAN 4: ENGINE DOKUMEN WORD ("ZERO-LOSS TEMPLATE CLONING" & FALLBACK KOP)
# ==============================================================================

def find_template_docx_path() -> str:
    """Mencari file LPJ_DWP GABUNGAN DI BLITAR_2026-1.docx di direktori kerja atau subdirektori."""
    candidates = [
        TEMPLATE_FILENAME,
        os.path.join(os.getcwd(), TEMPLATE_FILENAME),
    ]
    for root_dir, _, files in os.walk(os.getcwd()):
        if TEMPLATE_FILENAME in files:
            candidates.append(os.path.join(root_dir, TEMPLATE_FILENAME))
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return ""


def get_default_assets_bytes() -> dict:
    """Mengembalikan byte gambar Logo Kiri, Logo Kanan, dan Tanda Tangan Asli Ny. Rini Johan Pandores."""
    logo1_bytes = base64.b64decode(EMBEDDED_LOGO1_B64)
    logo2_bytes = base64.b64decode(EMBEDDED_LOGO2_B64)
    sig_bytes = base64.b64decode(EMBEDDED_SIG_B64)

    tpl_bytes = st.session_state.get("uploaded_template_bytes")
    tpl_path = find_template_docx_path()

    try:
        if tpl_bytes:
            doc_tmp = docx.Document(io.BytesIO(tpl_bytes))
        elif tpl_path:
            doc_tmp = docx.Document(tpl_path)
        else:
            doc_tmp = None

        if doc_tmp is not None:
            for _, part in doc_tmp.part.related_parts.items():
                if "image" in part.content_type:
                    sig_bytes = part.blob
    except Exception:
        pass

    return {
        "logo1": logo1_bytes,
        "logo2": logo2_bytes,
        "sig": sig_bytes,
    }


def _set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m_name, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m_name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_cell_shading(cell, hex_color: str):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def _set_table_borders(table, color="505050", sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def _build_fallback_header(section, logo1_bytes: bytes, logo2_bytes: bytes):
    """Membangun Kop Surat Resmi DWP KPPBC TMC Malang secara terprogram jika file template tidak ditemukan."""
    header = section.header
    header.is_linked_to_previous = False
    for p in list(header.paragraphs):
        p._element.getparent().remove(p._element)

    htbl = header.add_table(rows=1, cols=3, width=Cm(19.5))
    htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2.5), Cm(14.2), Cm(2.8)]
    for i, w in enumerate(widths):
        htbl.columns[i].width = w
        htbl.rows[0].cells[i].width = w
        htbl.rows[0].cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Logo Kiri
    p_left = htbl.rows[0].cells[0].paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.add_run().add_picture(io.BytesIO(logo1_bytes), width=Cm(1.95))

    # Teks Kop Tengah
    c_mid = htbl.rows[0].cells[1]
    p1 = c_mid.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_after = Pt(1)
    r1 = p1.add_run("Dharma Wanita Persatuan")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(16)
    r1.font.bold = True
    r1.font.italic = True

    p2 = c_mid.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(1)
    r2 = p2.add_run(
        "DIREKTORAT JENDERAL BEA DAN CUKAI KANTOR WILAYAH DJBC JAWA TIMUR II" + NL +
        "KANTOR PENGAWASAN DAN PELAYANAN BEA DAN CUKAI TIPE MADYA CUKAI MALANG"
    )
    r2.font.name = "Arial"
    r2.font.size = Pt(9)
    r2.font.bold = True

    p3 = c_mid.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(0)
    r3 = p3.add_run(
        "JALAN SURABAYA NOMOR 2, MALANG 65145" + NL +
        "TELEPON (0341) 551628; FAKSIMILE (0341)554193; LAMAN WWW.BEACUKAIMALANG.COM" + NL +
        "PUSAT KONTAK LAYANAN 1500225; SUREL INFO@CUSTOMS.GO.ID"
    )
    r3.font.name = "Calibri"
    r3.font.size = Pt(6.5)

    # Logo Kanan
    p_right = htbl.rows[0].cells[2].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right.add_run().add_picture(io.BytesIO(logo2_bytes), width=Cm(2.25))

    # Garis bawah kop surat
    p_line = header.add_paragraph()
    p_line.paragraph_format.space_before = Pt(2)
    p_line.paragraph_format.space_after = Pt(4)
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        '<w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/>'
        "</w:pBdr>"
    )
    p_line._p.get_or_add_pPr().append(pBdr)


def generate_lpj_docx_bytes(payload: dict, custom_sig_bytes: bytes = None, attachment_images: list = None) -> bytes:
    """
    Mesin utama pembuat dokumen Word (.docx) dengan Zero-Loss Template Cloning.
    Menjaga 100% keaslian Kop Surat & Tanda Tangan dari LPJ_DWP GABUNGAN DI BLITAR_2026-1.docx.
    """
    assets = get_default_assets_bytes()
    sig_to_use = custom_sig_bytes if custom_sig_bytes else assets["sig"]

    tpl_bytes = st.session_state.get("uploaded_template_bytes")
    tpl_path = find_template_docx_path()

    used_cloned_template = False
    if tpl_bytes:
        doc = docx.Document(io.BytesIO(tpl_bytes))
        used_cloned_template = True
    elif tpl_path:
        doc = docx.Document(tpl_path)
        used_cloned_template = True
    else:
        doc = docx.Document()

    if used_cloned_template:
        # Di template asli, referensi header berada pada sectPr di dalam paragraf ke-17.
        # Kita pindahkan headerReference tersebut ke sectPr utama dokumen sebelum menghapus isi body lama!
        hdr_ref = None
        for p in doc.paragraphs:
            pPr = p._p.find(qn("w:pPr"))
            if pPr is not None:
                sPr = pPr.find(qn("w:sectPr"))
                if sPr is not None:
                    href = sPr.find(qn("w:headerReference"))
                    if href is not None:
                        hdr_ref = href

        body_el = doc._body._body
        final_sectPr = body_el.find(qn("w:sectPr"))
        if hdr_ref is not None and final_sectPr is not None:
            if final_sectPr.find(qn("w:headerReference")) is None:
                final_sectPr.insert(0, hdr_ref)

        for child in list(body_el):
            if child.tag != qn("w:sectPr"):
                body_el.remove(child)
    else:
        sec = doc.sections[0]
        sec.page_width = Cm(21.59)
        sec.page_height = Cm(27.94)
        sec.top_margin = Cm(3.85)
        sec.bottom_margin = Cm(1.20)
        sec.left_margin = Cm(0.88)
        sec.right_margin = Cm(0.50)
        sec.header_distance = Cm(1.28)
        _build_fallback_header(sec, assets["logo1"], assets["logo2"])

    # Pastikan margin bawah cukup aman untuk konversi PDF
    for sec in doc.sections:
        if sec.bottom_margin < Cm(1.0):
            sec.bottom_margin = Cm(1.0)

    def _add_styled_p(text="", align=WD_ALIGN_PARAGRAPH.LEFT, left_cm=0.73, right_cm=0.80, first_line_cm=0.0, space_before_pt=0, space_after_pt=3, italic=False, bold=False, size_pt=11.0):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.left_indent = Cm(left_cm)
        pf.right_indent = Cm(right_cm)
        pf.first_line_indent = Cm(first_line_cm)
        pf.space_before = Pt(space_before_pt)
        pf.space_after = Pt(space_after_pt)
        pf.line_spacing = 1.15
        pf.keep_together = True
        if text:
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(size_pt)
            r.font.italic = italic
            r.font.bold = bold
            r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    # 1. Tanggal Laporan (Rata Kanan)
    tgl_laporan_indo = format_hari_tanggal_indo(payload.get("tgl_laporan", date.today()))
    _add_styled_p(
        tgl_laporan_indo,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
        left_cm=0.73,
        right_cm=1.32,
        space_before_pt=6,
        space_after_pt=10,
        size_pt=11.0,
    )

    # 2. Salam Pembuka
    _add_styled_p(
        "Assalamu’alaikum Wr. Wb.",
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=0.73,
        right_cm=0.80,
        space_before_pt=2,
        space_after_pt=8,
        italic=True,
        size_pt=11.0,
    )

    # 3. Paragraf Pembuka Adaptif
    teks_pembuka = build_paragraf_pembuka(
        payload.get("nama_kegiatan", ""),
        payload.get("tema_kegiatan", ""),
        payload.get("konteks_hut", ""),
    )
    _add_styled_p(
        teks_pembuka,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_cm=0.73,
        right_cm=0.80,
        first_line_cm=1.27,
        space_before_pt=0,
        space_after_pt=6,
        size_pt=11.0,
    )

    # 4. Blok Waktu & Tempat (Menggunakan Tab-Stop Presisi)
    tgl_acara_indo = (payload.get("tgl_acara_manual") or "").strip()
    if not tgl_acara_indo:
        tgl_acara_indo = format_hari_tanggal_indo(payload.get("tgl_acara", date.today()))

    waktu_acara = (payload.get("waktu_acara") or "08.30 - selesai").strip()
    tempat_acara = (payload.get("tempat_acara") or "-").strip()

    for label_kiri, nilai_kanan in [
        ("Hari, Tanggal", tgl_acara_indo),
        ("Waktu", waktu_acara),
        ("Tempat", tempat_acara),
    ]:
        p_info = _add_styled_p(
            "",
            align=WD_ALIGN_PARAGRAPH.LEFT,
            left_cm=1.98,
            right_cm=0.80,
            space_before_pt=1,
            space_after_pt=2,
            size_pt=11.0,
        )
        p_info.paragraph_format.tab_stops.add_tab_stop(Cm(4.85), WD_TAB_ALIGNMENT.LEFT)
        r_lbl = p_info.add_run(label_kiri + TAB + ": " + nilai_kanan)
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(11.0)

    # 5. Rangkaian Kegiatan
    _add_styled_p(
        "Dengan rangkaian kegiatan sebagai berikut :",
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=1.98,
        right_cm=0.80,
        space_before_pt=6,
        space_after_pt=3,
        size_pt=11.0,
    )

    rangkaian_raw = (payload.get("rangkaian_kegiatan") or "").strip()
    rangkaian_list = [ln.strip() for ln in rangkaian_raw.split(NL) if ln.strip()]
    for idx_r, item_r in enumerate(rangkaian_list, start=1):
        clean_r = item_r
        if "." in clean_r[:4]:
            l_num, r_rest = clean_r.split(".", 1)
            if l_num.strip().isdigit():
                clean_r = r_rest.strip()

        p_r = doc.add_paragraph()
        p_r.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf_r = p_r.paragraph_format
        pf_r.left_indent = Cm(2.61)
        pf_r.right_indent = Cm(0.80)
        pf_r.first_line_indent = Cm(-0.63)
        pf_r.space_before = Pt(1)
        pf_r.space_after = Pt(2)
        pf_r.line_spacing = 1.12
        pf_r.tab_stops.add_tab_stop(Cm(2.61), WD_TAB_ALIGNMENT.LEFT)
        run_r = p_r.add_run(f"{idx_r}." + TAB + clean_r)
        run_r.font.name = "Arial"
        run_r.font.size = Pt(11.0)

    # 6. Pengantar Rincian Dana
    sumber_dana = (payload.get("sumber_dana") or "Kas Dharma Wanita KPPBC TMC Malang dan donatur").strip()
    _add_styled_p(
        f"Dana untuk kegiatan ini berasal dari {sumber_dana}. Dengan rincian dana sebagai berikut :",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_cm=0.73,
        right_cm=0.80,
        first_line_cm=1.27,
        space_before_pt=6,
        space_after_pt=6,
        size_pt=11.0,
    )

    # 7. Konstruksi Tabel Rincian Dana 6 Kolom (Anti-Potong & Anti-Lompat Baris)
    df_items = normalize_items_df(payload.get("items_dana", []), payload.get("ket_tabel_dana", ""))
    n_items = len(df_items)
    total_rows = n_items + 3  # 2 baris header + N baris item + 1 baris TOTAL

    col_widths_cm = [1.05, 5.80, 1.85, 2.95, 3.10, 4.45]  # Total 19.20 cm, NO & JUMLAH anti-wrap
    table = doc.add_table(rows=total_rows, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_borders(table, color="505050", sz="4")

    # Kunci tblLayout fixed & indentasi tabel sesuai template
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="400" w:type="dxa"/>'))

    for col_idx, w_cm in enumerate(col_widths_cm):
        table.columns[col_idx].width = Cm(w_cm)

    def _format_cell_text(cell, text: str, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5, no_wrap=False):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_margins(cell, top=42, bottom=42, left=65, right=65)
        if no_wrap:
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(OxmlElement("w:noWrap"))
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        for r_old in list(p.runs):
            p._p.remove(r_old._r)
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(size_pt)
        r.font.bold = bold
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # Nested Header Baris 0 & Baris 1
    r0 = table.rows[0]
    r1 = table.rows[1]
    for h_row in (r0, r1):
        trPr = h_row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:tblHeader"))
        trPr.append(OxmlElement("w:cantSplit"))

    # Merge horizontal DANA (Kolom 3 & 4 pada Baris 0)
    r0.cells[3].merge(r0.cells[4])
    # Merge vertikal Kolom 0, 1, 2, 5 antara Baris 0 dan Baris 1
    for c_idx in (0, 1, 2, 5):
        r0.cells[c_idx].merge(r1.cells[c_idx])

    _format_cell_text(r0.cells[0], "NO", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5, no_wrap=True)
    _format_cell_text(r0.cells[1], "KEBUTUHAN KEGIATAN", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)
    _format_cell_text(r0.cells[2], "JUMLAH", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5, no_wrap=True)
    _format_cell_text(r0.cells[3], "DANA", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)
    _format_cell_text(r0.cells[5], "KETERANGAN", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)

    _format_cell_text(r1.cells[3], "HARGA SATUAN", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.0, no_wrap=True)
    _format_cell_text(r1.cells[4], "HARGA TOTAL", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.0, no_wrap=True)

    for h_r in (r0, r1):
        for c_idx, w_cm in enumerate(col_widths_cm):
            h_r.cells[c_idx].width = Cm(w_cm)
            _set_cell_shading(h_r.cells[c_idx], "F8FAFC")

    # Isi Baris Data (Item 1 s.d. N)
    merge_ket = bool(payload.get("merge_ket_vertikal", True))
    ket_global = (payload.get("ket_tabel_dana") or "Kas DWP BC Malang & Donatur Anggota").strip()

    for idx_df, row_data in df_items.iterrows():
        row_pos = 2 + idx_df
        tr_row = table.rows[row_pos]
        tr_row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))

        for c_idx, w_cm in enumerate(col_widths_cm):
            tr_row.cells[c_idx].width = Cm(w_cm)

        no_str = str(int(row_data["NO"]))
        keb_str = str(row_data["KEBUTUHAN KEGIATAN"])
        jml_str = str(int(row_data["JUMLAH"])) if int(row_data["JUMLAH"]) > 0 else "-"
        hsat_str = fmt_rp_cell(row_data["HARGA SATUAN"], is_dash_if_zero=True)
        htot_str = fmt_rp_cell(row_data["HARGA TOTAL"], is_dash_if_zero=True)
        ket_item = str(row_data.get("KETERANGAN", ket_global) or ket_global)

        _format_cell_text(tr_row.cells[0], no_str, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)
        _format_cell_text(tr_row.cells[1], keb_str, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size_pt=10.5)
        _format_cell_text(tr_row.cells[2], jml_str, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)
        _format_cell_text(tr_row.cells[3], hsat_str, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False, size_pt=10.5, no_wrap=True)
        _format_cell_text(tr_row.cells[4], htot_str, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False, size_pt=10.5, no_wrap=True)

        if not merge_ket:
            _format_cell_text(tr_row.cells[5], ket_item, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.0)

    if merge_ket and n_items > 0:
        first_ket_cell = table.rows[2].cells[5]
        if n_items > 1:
            last_ket_cell = table.rows[2 + n_items - 1].cells[5]
            first_ket_cell.merge(last_ket_cell)
        _format_cell_text(first_ket_cell, ket_global, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)

    # Baris Footer (TOTAL)
    r_tot = table.rows[total_rows - 1]
    r_tot._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    for c_idx, w_cm in enumerate(col_widths_cm):
        r_tot.cells[c_idx].width = Cm(w_cm)

    r_tot.cells[0].merge(r_tot.cells[3])
    total_keseluruhan = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

    _format_cell_text(r_tot.cells[0], "TOTAL", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)
    _format_cell_text(r_tot.cells[4], fmt_rp_cell(total_keseluruhan, is_dash_if_zero=False), align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size_pt=10.5, no_wrap=True)
    _format_cell_text(r_tot.cells[5], "", align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)

    # 8. Paragraf Ucapan Terima Kasih Donatur (Opsional/Default Aktif)
    if payload.get("tampilkan_ucapan_donatur", True):
        teks_donatur = (payload.get("teks_ucapan_donatur") or DEFAULT_UCAPAN_DONATUR).strip()
        if teks_donatur:
            p_don = _add_styled_p(
                teks_donatur,
                align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                left_cm=0.73,
                right_cm=0.80,
                first_line_cm=1.25,
                space_before_pt=10,
                space_after_pt=6,
                size_pt=11.0,
            )
            p_don.paragraph_format.keep_with_next = True

    # 9. Paragraf Penutup (Otomatis menyisipkan nama_kegiatan)
    nama_keg = (payload.get("nama_kegiatan") or "Kegiatan Dharma Wanita Persatuan").strip()
    teks_penutup = (
        f"Demikian laporan pertanggung jawaban {nama_keg} ini kami sampaikan, "
        "atas perhatian dan bantuan yang telah diberikan kami ucapkan terima kasih."
    )
    _add_styled_p(
        teks_penutup,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_cm=0.73,
        right_cm=0.80,
        first_line_cm=1.25,
        space_before_pt=4,
        space_after_pt=6,
        size_pt=11.0,
    )

    # 10. Salam Penutup
    _add_styled_p(
        "Wassalamu'alaikum Wr. Wb.",
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=0.73,
        right_cm=0.80,
        space_before_pt=2,
        space_after_pt=6,
        italic=True,
        size_pt=11.0,
    )

    # 11. Catatan Instagram
    catatan_ig = (payload.get("catatan_ig") or "").strip()
    if catatan_ig:
        p_ig = _add_styled_p(
            "",
            align=WD_ALIGN_PARAGRAPH.LEFT,
            left_cm=0.73,
            right_cm=0.80,
            space_before_pt=4,
            space_after_pt=8,
            size_pt=10.0,
        )
        if "@dwpbcmalang" in catatan_ig:
            parts_ig = catatan_ig.split("@dwpbcmalang", 1)
            r_ig1 = p_ig.add_run(parts_ig[0])
            r_ig1.font.name = "Arial"
            r_ig1.font.size = Pt(10.0)
            r_ig1.font.italic = True
            r_ig2 = p_ig.add_run("@dwpbcmalang")
            r_ig2.font.name = "Arial"
            r_ig2.font.size = Pt(10.0)
            r_ig2.font.bold = True
            if parts_ig[1]:
                r_ig3 = p_ig.add_run(parts_ig[1])
                r_ig3.font.name = "Arial"
                r_ig3.font.size = Pt(10.0)
        else:
            r_ig = p_ig.add_run(catatan_ig)
            r_ig.font.name = "Arial"
            r_ig.font.size = Pt(10.0)
            r_ig.font.italic = True

    # 12. Blok Tanda Tangan (Inden Kanan Presisi & Keep With Next agar tidak terpotong halaman)
    jab_1 = (payload.get("jabatan_baris_1") or "Ketua Dharma Wanita Persatuan").strip()
    jab_2 = (payload.get("jabatan_baris_2") or "KPPBC Tipe Madya Cukai Malang").strip()
    nama_ttd = (payload.get("nama_penandatangan") or "Ny. Rini Johan Pandores").strip()

    p_jab = _add_styled_p(
        jab_1 + NL + jab_2,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=12.14,
        right_cm=0.80,
        space_before_pt=6,
        space_after_pt=2,
        size_pt=11.0,
    )
    p_jab.paragraph_format.keep_with_next = True

    p_sig_img = doc.add_paragraph()
    p_sig_img.paragraph_format.left_indent = Cm(12.14)
    p_sig_img.paragraph_format.space_before = Pt(2)
    p_sig_img.paragraph_format.space_after = Pt(2)
    p_sig_img.paragraph_format.keep_with_next = True

    if payload.get("tampilkan_ttd", True) and sig_to_use:
        try:
            r_sig = p_sig_img.add_run()
            r_sig.add_picture(io.BytesIO(sig_to_use), width=Cm(3.35))
        except Exception:
            p_sig_img.paragraph_format.space_before = Pt(18)
            p_sig_img.paragraph_format.space_after = Pt(18)
    else:
        p_sig_img.paragraph_format.space_before = Pt(20)
        p_sig_img.paragraph_format.space_after = Pt(20)

    _add_styled_p(
        nama_ttd,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=12.14,
        right_cm=0.80,
        space_before_pt=2,
        space_after_pt=4,
        size_pt=11.0,
    )

    # 13. Opsional Lampiran Foto Dokumentasi / Nota Kwitansi (Auto-Grid 2x2)
    if attachment_images:
        doc.add_page_break()
        p_lamp = _add_styled_p(
            "LAMPIRAN DOKUMENTASI & BUKTI KWITANSI KEGIATAN",
            align=WD_ALIGN_PARAGRAPH.CENTER,
            left_cm=0.73,
            right_cm=0.80,
            space_before_pt=6,
            space_after_pt=10,
            bold=True,
            size_pt=12.0,
        )
        p_lamp.paragraph_format.keep_with_next = True

        n_imgs = len(attachment_images)
        grid_rows = (n_imgs + 1) // 2
        gtbl = doc.add_table(rows=grid_rows, cols=2)
        gtbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(gtbl, color="CBD5E1", sz="4")

        for idx_img, img_item in enumerate(attachment_images):
            r_i = idx_img // 2
            c_i = idx_img % 2
            cell = gtbl.rows[r_i].cells[c_i]
            cell.width = Cm(9.2)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
            p_c = cell.paragraphs[0]
            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                p_c.add_run().add_picture(io.BytesIO(img_item["bytes"]), width=Cm(8.2))
                p_cap = cell.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_cap = p_cap.add_run(img_item.get("caption", f"Dokumentasi {idx_img + 1}"))
                r_cap.font.name = "Arial"
                r_cap.font.size = Pt(9.5)
                r_cap.font.italic = True
            except Exception:
                p_c.add_run("[Gambar tidak dapat dimuat]")

    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue()


def convert_docx_bytes_to_pdf_bytes(docx_bytes: bytes) -> tuple:
    """Mengonversi byte .docx menjadi byte .pdf secara presisi menggunakan LibreOffice Headless."""
    lo_bin = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo_bin:
        return None, "LibreOffice Headless tidak ditemukan di server. Gunakan tombol Cetak PDF Browser di bawah."

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            in_docx = os.path.join(tmpdir, "LPJ_DWP.docx")
            out_pdf = os.path.join(tmpdir, "LPJ_DWP.pdf")
            with open(in_docx, "wb") as f:
                f.write(docx_bytes)

            cmd = [
                lo_bin,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmpdir,
                in_docx,
            ]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=45)
            if os.path.isfile(out_pdf) and os.path.getsize(out_pdf) > 0:
                with open(out_pdf, "rb") as f_pdf:
                    return f_pdf.read(), "PDF berhasil dibuat dengan LibreOffice Headless."
            err_txt = proc.stderr.decode("utf-8", errors="ignore")
            return None, f"Konversi PDF gagal: {err_txt}"
    except Exception as e:
        return None, f"Kendala konversi PDF: {e}"


# ==============================================================================
# BAGIAN 5: DIRECT DOWNLOAD BASE64 & LIVE PREVIEW KERTAS A4
# ==============================================================================

def make_direct_download_html(file_bytes: bytes, filename: str, mime_type: str, label: str, bg_gradient: str) -> str:
    """Membuat tombol unduh langsung berbasis Base64 Data URI (Anti-Error 404 Streamlit)."""
    b64_data = base64.b64encode(file_bytes).decode("ascii")
    return (
        f'<a href="data:{mime_type};base64,{b64_data}" download="{filename}" '
        f'style="display:flex;align-items:center;justify-content:center;gap:8px;'
        f'width:100%;min-height:52px;padding:12px 18px;border-radius:12px;'
        f'background:{bg_gradient};color:#FFFFFF !important;'
        f'-webkit-text-fill-color:#FFFFFF !important;font-weight:700;font-size:15px;'
        f'text-decoration:none;box-shadow:0 4px 12px rgba(124,45,18,0.22);'
        f'border:1px solid rgba(245,158,11,0.45);text-align:center;">'
        f"{label}</a>"
    )


def build_a4_live_preview_html(payload: dict, custom_sig_bytes: bytes = None) -> str:
    """Membangun pratinjau visual HTML Kertas A4 yang persis dengan LPJ DWP KPPBC TMC Malang."""
    assets = get_default_assets_bytes()
    logo1_b64 = base64.b64encode(assets["logo1"]).decode("ascii")
    logo2_b64 = base64.b64encode(assets["logo2"]).decode("ascii")
    sig_bytes = custom_sig_bytes if custom_sig_bytes else assets["sig"]
    sig_b64 = base64.b64encode(sig_bytes).decode("ascii") if sig_bytes else ""

    tgl_laporan_indo = format_hari_tanggal_indo(payload.get("tgl_laporan", date.today()))
    tgl_acara_indo = (payload.get("tgl_acara_manual") or "").strip() or format_hari_tanggal_indo(payload.get("tgl_acara", date.today()))
    waktu_acara = (payload.get("waktu_acara") or "08.30 - selesai").strip()
    tempat_acara = (payload.get("tempat_acara") or "-").strip()
    teks_pembuka = build_paragraf_pembuka(
        payload.get("nama_kegiatan", ""),
        payload.get("tema_kegiatan", ""),
        payload.get("konteks_hut", ""),
    )

    rangkaian_raw = (payload.get("rangkaian_kegiatan") or "").strip()
    rangkaian_items = [ln.strip() for ln in rangkaian_raw.split(NL) if ln.strip()]
    ol_html = ""
    for item_r in rangkaian_items:
        clean_r = item_r
        if "." in clean_r[:4]:
            l_num, r_rest = clean_r.split(".", 1)
            if l_num.strip().isdigit():
                clean_r = r_rest.strip()
        ol_html += f"<li style='margin-bottom:3px;'>{clean_r}</li>"

    df_items = normalize_items_df(payload.get("items_dana", []), payload.get("ket_tabel_dana", ""))
    n_items = len(df_items)
    merge_ket = bool(payload.get("merge_ket_vertikal", True))
    ket_global = (payload.get("ket_tabel_dana") or "Kas DWP BC Malang & Donatur Anggota").strip()
    total_dana = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

    rows_html = ""
    for idx_df, r in df_items.iterrows():
        no_str = str(int(r["NO"]))
        keb_str = str(r["KEBUTUHAN KEGIATAN"])
        jml_str = str(int(r["JUMLAH"])) if int(r["JUMLAH"]) > 0 else "-"
        hsat_str = fmt_rp_cell(r["HARGA SATUAN"], is_dash_if_zero=True)
        htot_str = fmt_rp_cell(r["HARGA TOTAL"], is_dash_if_zero=True)
        ket_item = str(r.get("KETERANGAN", ket_global) or ket_global)

        if merge_ket:
            if idx_df == 0:
                ket_td = f"<td rowspan='{n_items}' style='text-align:center;vertical-align:middle;padding:6px;'>{ket_global}</td>"
            else:
                ket_td = ""
        else:
            ket_td = f"<td style='text-align:center;vertical-align:middle;padding:4px 6px;'>{ket_item}</td>"

        rows_html += (
            "<tr>"
            f"<td style='text-align:center;'>{no_str}</td>"
            f"<td style='text-align:left;'>{keb_str}</td>"
            f"<td style='text-align:center;'>{jml_str}</td>"
            f"<td style='text-align:right;white-space:nowrap;'>{hsat_str}</td>"
            f"<td style='text-align:right;white-space:nowrap;'>{htot_str}</td>"
            f"{ket_td}"
            "</tr>"
        )

    ucapan_html = ""
    if payload.get("tampilkan_ucapan_donatur", True):
        teks_don = (payload.get("teks_ucapan_donatur") or DEFAULT_UCAPAN_DONATUR).strip()
        if teks_don:
            ucapan_html = f"<p class='lpj-p-indent'>{teks_don}</p>"

    nama_keg = (payload.get("nama_kegiatan") or "Kegiatan Dharma Wanita Persatuan").strip()
    penutup_html = (
        f"<p class='lpj-p-indent'>Demikian laporan pertanggung jawaban {nama_keg} ini kami sampaikan, "
        "atas perhatian dan bantuan yang telah diberikan kami ucapkan terima kasih.</p>"
    )

    catatan_ig = (payload.get("catatan_ig") or "").strip()
    ig_html = ""
    if catatan_ig:
        ig_formatted = catatan_ig.replace("@dwpbcmalang", "<b>@dwpbcmalang</b>")
        ig_html = f"<p style='margin:10px 0 12px 24px;font-size:12.5px;font-style:italic;'>{ig_formatted}</p>"

    jab_1 = (payload.get("jabatan_baris_1") or "Ketua Dharma Wanita Persatuan").strip()
    jab_2 = (payload.get("jabatan_baris_2") or "KPPBC Tipe Madya Cukai Malang").strip()
    nama_ttd = (payload.get("nama_penandatangan") or "Ny. Rini Johan Pandores").strip()

    if payload.get("tampilkan_ttd", True) and sig_b64:
        sig_img_html = f"<img src='data:image/jpeg;base64,{sig_b64}' style='height:58px;margin:4px 0;display:block;'/>"
    else:
        sig_img_html = "<div style='height:54px;'></div>"

    html_doc = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8"/>
    <style>
        :root {{ color-scheme: light !important; }}
        body {{
            background: #E2E8F0;
            margin: 0;
            padding: 16px 8px;
            font-family: Arial, Helvetica, sans-serif;
            color: #000000 !important;
        }}
        .toolbar-print {{
            max-width: 215mm;
            margin: 0 auto 12px auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #FFFFFF;
            padding: 10px 16px;
            border-radius: 10px;
            border: 1px solid #CBD5E1;
            box-shadow: 0 2px 6px rgba(15,23,42,0.06);
        }}
        .btn-print {{
            background: linear-gradient(135deg, #7C2D12, #9A3412, #B45309);
            color: #FFFFFF !important;
            border: none;
            padding: 9px 16px;
            border-radius: 8px;
            font-weight: 700;
            font-size: 13.5px;
            cursor: pointer;
        }}
        .a4-sheet {{
            max-width: 215mm;
            min-height: 279mm;
            margin: 0 auto;
            background: #FFFFFF !important;
            color: #000000 !important;
            padding: 12mm 14mm 15mm 16mm;
            box-shadow: 0 8px 28px rgba(15,23,42,0.15);
            box-sizing: border-box;
            font-size: 13.5px;
            line-height: 1.42;
        }}
        .kop-table {{
            width: 100%;
            border-collapse: collapse;
            border-bottom: 2.2px solid #000000;
            margin-bottom: 12px;
        }}
        .kop-table td {{
            border: none !important;
            vertical-align: middle;
            padding-bottom: 6px;
        }}
        .lpj-p-indent {{
            text-align: justify;
            text-indent: 36px;
            margin: 6px 24px;
        }}
        .info-table {{
            margin: 6px 0 6px 64px;
            border-collapse: collapse;
        }}
        .info-table td {{
            padding: 1.5px 6px 1.5px 0;
            vertical-align: top;
            font-size: 13.5px;
        }}
        .dana-table {{
            width: calc(100% - 24px);
            margin: 8px 12px 12px 12px;
            border-collapse: collapse;
            font-size: 12.8px;
        }}
        .dana-table th, .dana-table td {{
            border: 1px solid #475569;
            padding: 4px 6px;
            color: #000000 !important;
        }}
        .dana-table th {{
            background: #F8FAFC;
            font-weight: 700;
            text-align: center;
        }}
        @media print {{
            body {{ background: #FFFFFF; padding: 0; }}
            .toolbar-print {{ display: none !important; }}
            .a4-sheet {{ box-shadow: none; margin: 0; max-width: 100%; padding: 8mm 10mm; }}
        }}
    </style>
    </head>
    <body>
        <div class="toolbar-print">
            <div style="font-size:13px;font-weight:600;color:#334155;">
                📄 Pratinjau Kertas Surat Resmi LPJ DWP KPPBC TMC Malang (100% Sesuai Template)
            </div>
            <button class="btn-print" onclick="window.print()">🖨️ Cetak / Simpan PDF via Browser</button>
        </div>
        <div class="a4-sheet">
            <table class="kop-table">
                <tr>
                    <td style="width:78px;text-align:center;">
                        <img src="data:image/jpeg;base64,{logo1_b64}" style="width:68px;height:auto;"/>
                    </td>
                    <td style="text-align:center;">
                        <div style="font-family:'Times New Roman',serif;font-size:20px;font-weight:700;font-style:italic;letter-spacing:0.2px;">
                            Dharma Wanita Persatuan
                        </div>
                        <div style="font-size:11px;font-weight:700;line-height:1.25;margin-top:2px;">
                            DIREKTORAT JENDERAL BEA DAN CUKAI KANTOR WILAYAH DJBC JAWA TIMUR II<br/>
                            KANTOR PENGAWASAN DAN PELAYANAN BEA DAN CUKAI TIPE MADYA CUKAI MALANG
                        </div>
                        <div style="font-size:8.5px;margin-top:3px;line-height:1.25;">
                            JALAN SURABAYA NOMOR 2, MALANG 65145<br/>
                            TELEPON (0341) 551628; FAKSIMILE (0341)554193; LAMAN <span style="color:#1D4ED8;text-decoration:underline;">WWW.BEACUKAIMALANG.COM</span><br/>
                            PUSAT KONTAK LAYANAN 1500225; SUREL <span style="color:#1D4ED8;text-decoration:underline;">INFO@CUSTOMS.GO.ID</span>
                        </div>
                    </td>
                    <td style="width:88px;text-align:center;">
                        <img src="data:image/png;base64,{logo2_b64}" style="width:78px;height:auto;"/>
                    </td>
                </tr>
            </table>

            <div style="text-align:right;margin:8px 28px 14px 24px;">{tgl_laporan_indo}</div>
            <div style="margin:6px 24px 10px 24px;font-style:italic;">Assalamu’alaikum Wr. Wb.</div>

            <p class="lpj-p-indent">{teks_pembuka}</p>

            <table class="info-table">
                <tr><td style="width:110px;">Hari, Tanggal</td><td>: {tgl_acara_indo}</td></tr>
                <tr><td>Waktu</td><td>: {waktu_acara}</td></tr>
                <tr><td>Tempat</td><td>: {tempat_acara}</td></tr>
            </table>

            <div style="margin:8px 24px 4px 64px;">Dengan rangkaian kegiatan sebagai berikut :</div>
            <ol style="margin:2px 28px 10px 82px;padding-left:14px;">
                {ol_html}
            </ol>

            <p class="lpj-p-indent">
                Dana untuk kegiatan ini berasal dari {payload.get("sumber_dana", "")}. Dengan rincian dana sebagai berikut :
            </p>

            <table class="dana-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="width:6%;">NO</th>
                        <th rowspan="2" style="width:31%;">KEBUTUHAN KEGIATAN</th>
                        <th rowspan="2" style="width:9%;">JUMLAH</th>
                        <th colspan="2" style="width:32%;">DANA</th>
                        <th rowspan="2" style="width:22%;">KETERANGAN</th>
                    </tr>
                    <tr>
                        <th style="width:16%;">HARGA SATUAN</th>
                        <th style="width:16%;">HARGA TOTAL</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                    <tr>
                        <td colspan="4" style="text-align:center;font-weight:700;background:#F8FAFC;">TOTAL</td>
                        <td style="text-align:right;font-weight:700;white-space:nowrap;background:#F8FAFC;">{fmt_rp_cell(total_dana, is_dash_if_zero=False)}</td>
                        <td style="background:#F8FAFC;"></td>
                    </tr>
                </tbody>
            </table>

            {ucapan_html}
            {penutup_html}

            <div style="margin:8px 24px 6px 24px;font-style:italic;">Wassalamu'alaikum Wr. Wb.</div>
            {ig_html}

            <div style="margin-left:58%;margin-top:14px;line-height:1.35;">
                <div>{jab_1}</div>
                <div>{jab_2}</div>
                {sig_img_html}
                <div style="font-weight:600;">{nama_ttd}</div>
            </div>
        </div>
    </body>
    </html>
    """
    return html_doc


# ==============================================================================
# BAGIAN 6: ARSITEKTUR UI/UX (DWP ELEGANCE + MOBILE-FIRST + ANTI DARK-MODE)
# ==============================================================================

def inject_dwp_elegance_css():
    st.markdown(
        """
        <style>
        :root {
            color-scheme: light !important;
        }
        html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
            background-color: #FDFBF7 !important;
            color: #0F172A !important;
            font-family: 'Inter', 'Segoe UI', Roboto, sans-serif;
        }
        [data-testid="stHeader"] {
            background-color: rgba(253, 251, 247, 0.92) !important;
        }
        [data-testid="stSidebar"] {
            background-color: #FFF8F1 !important;
            border-right: 1px solid #FED7AA !important;
        }
        /* Paksa seluruh teks, label, markdown, metrik berwarna gelap pekat #0F172A */
        p, span, label, h1, h2, h3, h4, h5, h6, li, td, th,
        [data-testid="stWidgetLabel"] p,
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMetricValue"],
        [data-testid="stMetricLabel"] {
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
        }
        /* Ergonomi Mobile-First (Anti Auto-Zoom di HP: font-size 16px & min-height 48px) */
        input, textarea, select,
        div[data-baseweb="input"] input,
        div[data-baseweb="textarea"] textarea,
        div[data-baseweb="select"] > div {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
            font-size: 16px !important;
            min-height: 48px !important;
            border-radius: 10px !important;
            border: 1.5px solid #CBD5E1 !important;
        }
        div[data-baseweb="popover"], ul[data-baseweb="menu"], li[role="option"] {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
            font-size: 16px !important;
        }
        /* Tombol aksi utama minimal tinggi 52px dengan gradasi DWP Terracotta ke Deep Burgundy */
        .stButton > button {
            min-height: 52px !important;
            border-radius: 12px !important;
            font-size: 15.5px !important;
            font-weight: 700 !important;
            background: linear-gradient(135deg, #7C2D12 0%, #9A3412 55%, #B45309 100%) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border: 1px solid #F59E0B !important;
            box-shadow: 0 4px 12px rgba(124, 45, 18, 0.18) !important;
            transition: all 0.2s ease-in-out !important;
        }
        .stButton > button p, .stButton > button span {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        .stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 16px rgba(124, 45, 18, 0.28) !important;
        }
        /* Styling Tabs */
        button[data-baseweb="tab"] {
            min-height: 50px !important;
            font-size: 15.5px !important;
            font-weight: 700 !important;
            padding: 10px 20px !important;
            border-radius: 10px 10px 0 0 !important;
            background-color: #FFFFFF !important;
            margin-right: 6px !important;
            border: 1px solid #E2E8F0 !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            background: linear-gradient(135deg, #7C2D12 0%, #9A3412 100%) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border-bottom: 3px solid #F59E0B !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] p {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        /* Kartu putih bersih */
        .dwp-card {
            background: #FFFFFF !important;
            border: 1px solid #E2E8F0;
            border-left: 5px solid #9A3412;
            border-radius: 14px;
            padding: 18px 20px;
            margin-bottom: 16px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
        }
        .dwp-hero {
            background: linear-gradient(135deg, #7C2D12 0%, #9A3412 52%, #B45309 100%);
            border-radius: 16px;
            padding: 22px 26px;
            margin-bottom: 18px;
            border: 1.5px solid #F59E0B;
            box-shadow: 0 8px 24px rgba(124, 45, 18, 0.20);
        }
        .dwp-hero h1, .dwp-hero p, .dwp-hero span {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        .metric-box {
            background: #FFFFFF;
            border: 1.5px solid #FED7AA;
            border-radius: 12px;
            padding: 14px 16px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(124, 45, 18, 0.06);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# BAGIAN 7: INISIALISASI STATE & PRESET LOADER
# ==============================================================================

def load_payload_into_state(data: dict):
    """Memuat seluruh field LPJ ke st.session_state agar langsung tampil di semua Tab."""
    st.session_state["id_lpj"] = data.get("id_lpj", "LPJ-" + datetime.now().strftime("%Y%m%d-%H%M%S"))

    tgl_lap = data.get("tgl_laporan", date(2026, 8, 13))
    if isinstance(tgl_lap, str):
        try:
            tgl_lap = datetime.strptime(tgl_lap[:10], "%Y-%m-%d").date()
        except Exception:
            tgl_lap = date(2026, 8, 13)
    st.session_state["tgl_laporan"] = tgl_lap

    st.session_state["nama_kegiatan"] = data.get("nama_kegiatan", "")
    st.session_state["tema_kegiatan"] = data.get("tema_kegiatan", "")
    st.session_state["konteks_hut"] = data.get("konteks_hut", "")

    tgl_acr = data.get("tgl_acara", date(2026, 8, 10))
    if isinstance(tgl_acr, str):
        try:
            tgl_acr = datetime.strptime(tgl_acr[:10], "%Y-%m-%d").date()
        except Exception:
            tgl_acr = date(2026, 8, 10)
    st.session_state["tgl_acara"] = tgl_acr

    st.session_state["tgl_acara_manual"] = data.get("tgl_acara_manual", "")
    st.session_state["waktu_acara"] = data.get("waktu_acara", "08.30 - selesai")
    st.session_state["tempat_acara"] = data.get("tempat_acara", "")
    st.session_state["rangkaian_kegiatan"] = data.get("rangkaian_kegiatan", "")
    st.session_state["sumber_dana"] = data.get("sumber_dana", "Kas Dharma Wanita KPPBC TMC Malang dan donatur")
    st.session_state["ket_tabel_dana"] = data.get("ket_tabel_dana", "Kas DWP BC Malang & Donatur Anggota")
    st.session_state["merge_ket_vertikal"] = bool(data.get("merge_ket_vertikal", True))
    st.session_state["tampilkan_ucapan_donatur"] = bool(data.get("tampilkan_ucapan_donatur", True))
    st.session_state["teks_ucapan_donatur"] = data.get("teks_ucapan_donatur", DEFAULT_UCAPAN_DONATUR)
    st.session_state["catatan_ig"] = data.get("catatan_ig", "*Note : dokumentasi kegiatan dapat dilihat di Instagram @dwpbcmalang")
    st.session_state["jabatan_baris_1"] = data.get("jabatan_baris_1", "Ketua Dharma Wanita Persatuan")
    st.session_state["jabatan_baris_2"] = data.get("jabatan_baris_2", "KPPBC Tipe Madya Cukai Malang")
    st.session_state["nama_penandatangan"] = data.get("nama_penandatangan", "Ny. Rini Johan Pandores")
    st.session_state["tampilkan_ttd"] = bool(data.get("tampilkan_ttd", True))

    df_norm = normalize_items_df(data.get("items_dana", []), st.session_state["ket_tabel_dana"])
    st.session_state["df_items_dana"] = df_norm


def init_app_state():
    if "app_initialized" not in st.session_state:
        load_payload_into_state(PRESET_BLITAR_2026)
        st.session_state["local_db_lpj"] = []
        st.session_state["local_rekap_dana"] = []
        st.session_state["app_initialized"] = True


def collect_current_payload() -> dict:
    return {
        "id_lpj": st.session_state.get("id_lpj", "LPJ-20260813-BLITAR"),
        "tgl_laporan": st.session_state.get("tgl_laporan", date(2026, 8, 13)),
        "nama_kegiatan": st.session_state.get("nama_kegiatan", ""),
        "tema_kegiatan": st.session_state.get("tema_kegiatan", ""),
        "konteks_hut": st.session_state.get("konteks_hut", ""),
        "tgl_acara": st.session_state.get("tgl_acara", date(2026, 8, 10)),
        "tgl_acara_manual": st.session_state.get("tgl_acara_manual", ""),
        "waktu_acara": st.session_state.get("waktu_acara", "08.30 - selesai"),
        "tempat_acara": st.session_state.get("tempat_acara", ""),
        "rangkaian_kegiatan": st.session_state.get("rangkaian_kegiatan", ""),
        "sumber_dana": st.session_state.get("sumber_dana", ""),
        "ket_tabel_dana": st.session_state.get("ket_tabel_dana", ""),
        "merge_ket_vertikal": st.session_state.get("merge_ket_vertikal", True),
        "tampilkan_ucapan_donatur": st.session_state.get("tampilkan_ucapan_donatur", True),
        "teks_ucapan_donatur": st.session_state.get("teks_ucapan_donatur", DEFAULT_UCAPAN_DONATUR),
        "catatan_ig": st.session_state.get("catatan_ig", ""),
        "jabatan_baris_1": st.session_state.get("jabatan_baris_1", "Ketua Dharma Wanita Persatuan"),
        "jabatan_baris_2": st.session_state.get("jabatan_baris_2", "KPPBC Tipe Madya Cukai Malang"),
        "nama_penandatangan": st.session_state.get("nama_penandatangan", "Ny. Rini Johan Pandores"),
        "tampilkan_ttd": st.session_state.get("tampilkan_ttd", True),
        "items_dana": st.session_state.get("df_items_dana", pd.DataFrame()).to_dict(orient="records"),
    }


# ==============================================================================
# BAGIAN 8: MAIN APPLICATION (SIDEBAR + 3 TAB WORKFLOW 3 MENIT SELESAI)
# ==============================================================================

def main():
    inject_dwp_elegance_css()
    init_app_state()

    # --- SIDEBAR: KONFIGURASI TEMPLATE & DATABASE GOOGLE SHEETS ---
    with st.sidebar:
        st.markdown("### 🏛️ Panel Kontrol DWP BC Malang")
        st.caption("Template-Cloning & Smart Financial Engine V3.0")

        tpl_found = find_template_docx_path()
        if tpl_found:
            st.success(f"✅ Template Asli Terdeteksi: `{os.path.basename(tpl_found)}`")
        elif st.session_state.get("uploaded_template_bytes"):
            st.success("✅ Template Asli Terhubung dari Upload Sidebar")
        else:
            st.info("ℹ️ Menggunakan Engine Kop Surat & Tanda Tangan Asli Bawaan (Zero-Loss Embedded)")

        with st.expander("📂 Upload File Template .docx (Opsional)", expanded=False):
            up_tpl = st.file_uploader("Pilih file `LPJ_DWP GABUNGAN DI BLITAR_2026-1.docx`", type=["docx"], key="up_tpl_sidebar")
            if up_tpl is not None:
                st.session_state["uploaded_template_bytes"] = up_tpl.read()
                st.success("Template berhasil dimuat ke memori!")

        st.markdown("### ☁️ Koneksi Google Sheets")
        sheet_id_in = st.text_input("ID Spreadsheet Google Sheets", value=DEFAULT_SPREADSHEET_ID)

        up_sa = st.file_uploader("Upload JSON Service Account (Opsional)", type=["json"], key="up_sa_json")
        creds_str = ""
        if up_sa is not None:
            creds_str = up_sa.getvalue().decode("utf-8", errors="ignore")
            st.session_state["creds_json_str"] = creds_str
        else:
            creds_str = st.session_state.get("creds_json_str", "")

        bundle = get_sheet_bundle(sheet_id_in, creds_str)
        if bundle.get("ok"):
            st.success("🟢 " + bundle.get("msg", "Terhubung"))
        else:
            st.warning("🟡 " + bundle.get("msg", "Mode Offline Lokal Aktif"))

        if st.button("🔄 Segarkan Data Google Sheets", use_container_width=True):
            fetch_sheet_bundle_cached.clear()
            st.rerun()

    # --- HERO HEADER ---
    st.markdown(
        """
        <div class="dwp-hero">
            <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;">
                <div>
                    <span style="background:rgba(245,158,11,0.28);border:1px solid #FCD34D;padding:4px 12px;border-radius:999px;font-size:12.5px;font-weight:700;">
                        ✨ DWP KPPBC TIPE MADYA CUKAI MALANG — V3.0 PRODUCTION
                    </span>
                    <h1 style="margin:8px 0 4px 0;font-size:26px;font-weight:800;">
                        Smart Generator LPJ Kegiatan Dharma Wanita Persatuan
                    </h1>
                    <p style="margin:0;font-size:14.5px;opacity:0.95;">
                        Zero-Loss Template Cloning • Smart Financial 3-Mode Engine • Anti-Potong PDF • Integrasi Google Sheets
                    </p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --- 3 TAB UTAMA ("WORKFLOW 3 MENIT SELESAI") ---
    tab1, tab2, tab3 = st.tabs([
        "📝 TAB 1: Identitas & Rangkaian Kegiatan",
        "💰 TAB 2: Tabel Rincian Dana Pintar",
        "📥 TAB 3: Preview Persis Asli, Simpan & Download",
    ])

    # ==========================================================================
    # TAB 1: IDENTITAS & RANGKAIAN KEGIATAN
    # ==========================================================================
    with tab1:
        st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
        col_preset, col_hist = st.columns([1.15, 1.35])

        with col_preset:
            st.markdown("#### ⚡ Preset Dokumen Resmi 1-Klik")
            if st.button("📋 Muat Contoh Persis LPJ DWP Gabungan di Blitar (2026)", use_container_width=True):
                load_payload_into_state(PRESET_BLITAR_2026)
                st.toast("Seluruh data LPJ DWP Gabungan di Blitar (2026) berhasil dimuat!", icon="✅")
                st.rerun()

        with col_hist:
            st.markdown("#### 🗂️ Muat dari Riwayat Database (GSheet / Lokal)")
            db_rows = bundle.get("db_rows", [])
            if db_rows:
                options_map = {}
                for r in db_rows:
                    lbl = f"{r.get('ID_LPJ', '')} | {r.get('TGL_ACARA', '')} — {str(r.get('NAMA_KEGIATAN', ''))[:45]}"
                    options_map[lbl] = r
                selected_lbl = st.selectbox("Pilih Riwayat LPJ untuk Diedit / Diduplikasi", options=["-- Pilih Riwayat LPJ --"] + list(options_map.keys()))
                if selected_lbl != "-- Pilih Riwayat LPJ --":
                    if st.button("📂 Buka & Muat Riwayat Terpilih", use_container_width=True):
                        row_sel = options_map[selected_lbl]
                        raw_json = row_sel.get("PAYLOAD_JSON", "")
                        if raw_json:
                            try:
                                parsed_payload = json.loads(raw_json)
                                load_payload_into_state(parsed_payload)
                                st.toast("Riwayat LPJ berhasil dimuat ke dalam form!", icon="📂")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Gagal membaca payload riwayat: {e}")
            else:
                st.info("Belum ada riwayat LPJ tersimpan di sesi ini. Klik Simpan di Tab 3 untuk mulai merekam riwayat.")
        st.markdown("</div>", unsafe_allow_html=True)

        # Form Identitas & Rangkaian Kegiatan
        st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
        st.markdown("#### 📌 Identitas Surat & Pelaksanaan Kegiatan")

        c1, c2 = st.columns([1, 2])
        with c1:
            st.date_input("Tanggal Surat LPJ (Kanan Atas)", key="tgl_laporan")
            st.caption(f"Format Cetak: **{format_hari_tanggal_indo(st.session_state.get('tgl_laporan', date.today()))}**")
        with c2:
            st.text_input(
                "Nama Kegiatan Utama",
                key="nama_kegiatan",
                placeholder="Contoh: Pertemuan Rutin Paguyuban Dharma Wanita Persatuan DJBC...",
            )

        c3, c4 = st.columns(2)
        with c3:
            st.text_input(
                "Tema Kegiatan (Opsional — kosongkan jika tanpa tema)",
                key="tema_kegiatan",
                placeholder="Contoh: Kebersamaan dalam Keceriaan, Bersinergi dalam Kebahagiaan",
            )
        with c4:
            st.text_input(
                "Konteks Tambahan / HUT (Opsional)",
                key="konteks_hut",
                placeholder="Contoh: menyambut HUT ke-81 Kemerdekaan Republik Indonesia",
            )

        c5, c6, c7 = st.columns([1, 1, 1.2])
        with c5:
            st.date_input("Pilih Tanggal Kegiatan", key="tgl_acara")
            st.text_input(
                "Atau Ketik Hari/Tanggal Manual (Jika Multi-Hari)",
                key="tgl_acara_manual",
                placeholder="Kosongkan jika memakai pemilih tanggal di atas",
            )
        with c6:
            st.text_input("Waktu Kegiatan", key="waktu_acara", placeholder="08.30 - selesai")
        with c7:
            st.text_input("Tempat Kegiatan", key="tempat_acara", placeholder="KPPBC Tipe Madya Pabean C Blitar")

        st.markdown("#### 📋 Rangkaian Kegiatan (1 Baris = 1 Nomor Urut Otomatis)")
        st.text_area(
            "Ketik atau tempel rangkaian acara (setiap baris baru otomatis menjadi nomor 1, 2, 3 dst. di dokumen Word & PDF)",
            key="rangkaian_kegiatan",
            height=135,
        )

        # Pratinjau Kalimat Pembuka Adaptif
        preview_pembuka = build_paragraf_pembuka(
            st.session_state.get("nama_kegiatan", ""),
            st.session_state.get("tema_kegiatan", ""),
            st.session_state.get("konteks_hut", ""),
        )
        st.info(f"**Pratinjau Paragraf Pembuka Otomatis:**{NL}{preview_pembuka}")
        st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================================================
    # TAB 2: TABEL RINCIAN DANA PINTAR (SMART FINANCIAL ENGINE)
    # ==========================================================================
    with tab2:
        st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
        st.markdown("#### 🏦 Sumber Dana & Pengaturan Kolom Keterangan")
        cd1, cd2, cd3 = st.columns([1.4, 1.4, 0.9])
        with cd1:
            st.text_input("Sumber Dana (Kalimat Pengantar)", key="sumber_dana")
        with cd2:
            st.text_input("Keterangan Kolom Tabel (Merged)", key="ket_tabel_dana")
        with cd3:
            st.write("")
            st.checkbox("Merge Vertikal Kolom Keterangan", key="merge_ket_vertikal")
        st.markdown("</div>", unsafe_allow_html=True)

        col_paste, col_quick = st.columns([1.15, 1.15])

        # Fitur 1: Smart Paste dari WhatsApp / Catatan
        with col_paste:
            st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
            st.markdown("#### ⚡ Fitur 1: Smart Paste dari WhatsApp / Catatan")
            st.caption("Contoh format bebas: `Sarapan Pagi (Roti) | 20 | 13000` atau `Roti 20 x 13.000`, `Iuran Blitar 10 total 675.000`, `Mobil 2 -`")
            wa_text = st.text_area(
                "Tempel daftar rincian belanja dari WA di sini:",
                height=120,
                placeholder="Iuran DWP Gabungan di Blitar | 10 | 0 | 675000" + NL + "Doorprize | 10 | 0 | 0" + NL + "Sarapan Pagi (Roti) 20 x 13.000",
                key="wa_paste_box",
            )
            c_p1, c_p2 = st.columns(2)
            with c_p1:
                if st.button("🔮 Ekstrak & Tambah ke Tabel", use_container_width=True):
                    new_items = smart_paste_parser(wa_text, st.session_state.get("ket_tabel_dana", ""))
                    if new_items:
                        cur_list = st.session_state["df_items_dana"].to_dict(orient="records")
                        cur_list.extend(new_items)
                        st.session_state["df_items_dana"] = normalize_items_df(cur_list, st.session_state.get("ket_tabel_dana", ""))
                        st.toast(f"{len(new_items)} item berhasil diekstrak ke tabel!", icon="🔮")
                        st.rerun()
            with c_p2:
                if st.button("🔄 Ganti Seluruh Isi Tabel", use_container_width=True):
                    new_items = smart_paste_parser(wa_text, st.session_state.get("ket_tabel_dana", ""))
                    if new_items:
                        st.session_state["df_items_dana"] = normalize_items_df(new_items, st.session_state.get("ket_tabel_dana", ""))
                        st.toast("Tabel rincian dana berhasil diperbarui!", icon="✅")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        # Fitur 2: Form Tambah Item Cepat (1 Baris)
        with col_quick:
            st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
            st.markdown("#### ➕ Fitur 2: Tambah Item Cepat (Logika 3-Mode)")
            q_nama = st.text_input("Kebutuhan Kegiatan", key="q_nama_in", placeholder="Contoh: Hadiah Peserta Lomba Tumpeng")
            cq1, cq2 = st.columns([0.8, 1.5])
            with cq1:
                q_jml = st.number_input("Jumlah", min_value=1, value=1, step=1, key="q_jml_in")
            with cq2:
                q_mode = st.selectbox(
                    "Tipe Perhitungan Item",
                    options=[
                        "Normal (Jumlah x Harga Satuan)",
                        "Borongan / Lump Sum (Langsung Harga Total, Satuan Rp -)",
                        "Fasilitas / In-Kind (Harga Satuan Rp - & Total Rp -)",
                    ],
                    key="q_mode_in",
                )
            q_nom_str = st.text_input("Nominal (Rp) — boleh ketik 250000 atau 250.000", value="0", key="q_nom_in")

            if st.button("➕ Tambahkan Baris ke Tabel Rincian Dana", use_container_width=True):
                if q_nama.strip():
                    nom_val = parse_num(q_nom_str)
                    if q_mode.startswith("Normal"):
                        hsat, htot = nom_val, int(q_jml) * nom_val
                    elif q_mode.startswith("Borongan"):
                        hsat, htot = 0, nom_val
                    else:
                        hsat, htot = 0, 0

                    cur_list = st.session_state["df_items_dana"].to_dict(orient="records")
                    cur_list.append({
                        "KEBUTUHAN KEGIATAN": q_nama.strip(),
                        "JUMLAH": int(q_jml),
                        "HARGA SATUAN": hsat,
                        "HARGA TOTAL": htot,
                        "KETERANGAN": st.session_state.get("ket_tabel_dana", ""),
                    })
                    st.session_state["df_items_dana"] = normalize_items_df(cur_list, st.session_state.get("ket_tabel_dana", ""))
                    st.toast(f"Item '{q_nama.strip()}' berhasil ditambahkan!", icon="➕")
                    st.rerun()
                else:
                    st.warning("Mohon isi nama Kebutuhan Kegiatan terlebih dahulu.")
            st.markdown("</div>", unsafe_allow_html=True)

        # Fitur 3: Tabel Interaktif (st.data_editor)
        st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
        st.markdown("#### 📊 Fitur 3: Tabel Rincian Dana Interaktif (Klik Sel untuk Edit Langsung)")
        st.caption(
            "**Tips 3-Mode:** "
            "(1) Isi **HARGA SATUAN > 0** agar **HARGA TOTAL** dihitung otomatis (`JUMLAH x HARGA SATUAN`). "
            "(2) Biarkan **HARGA SATUAN = 0** dan isi **HARGA TOTAL > 0** untuk item Borongan/Iuran (`Rp -` pada Harga Satuan). "
            "(3) Isi **0** pada keduanya untuk item Fasilitas/Doorprize/Mobil (`Rp -`)."
        )

        edited_df = st.data_editor(
            st.session_state["df_items_dana"],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "NO": st.column_config.NumberColumn("NO", width="small", disabled=True),
                "KEBUTUHAN KEGIATAN": st.column_config.TextColumn("KEBUTUHAN KEGIATAN", width="large", required=True),
                "JUMLAH": st.column_config.NumberColumn("JUMLAH", min_value=0, step=1, width="small"),
                "HARGA SATUAN": st.column_config.NumberColumn("HARGA SATUAN (Rp)", min_value=0, step=1000, format="%d"),
                "HARGA TOTAL": st.column_config.NumberColumn("HARGA TOTAL (Rp)", min_value=0, step=1000, format="%d"),
                "KETERANGAN": st.column_config.TextColumn("KETERANGAN", width="medium"),
            },
            key="data_editor_lpj",
        )

        df_recalc = normalize_items_df(edited_df, st.session_state.get("ket_tabel_dana", ""))
        if not df_recalc.equals(st.session_state["df_items_dana"]):
            st.session_state["df_items_dana"] = df_recalc
            st.rerun()

        total_items = len(df_recalc)
        total_pengeluaran = int(df_recalc["HARGA TOTAL"].sum()) if not df_recalc.empty else 0
        teks_terbilang = terbilang(total_pengeluaran)

        m1, m2, m3 = st.columns([1, 1.3, 1.8])
        with m1:
            st.markdown(
                f'<div class="metric-box"><div style="font-size:13px;font-weight:600;color:#64748B;">Total Item Kebutuhan</div>'
                f'<div style="font-size:24px;font-weight:800;color:#7C2D12;">{total_items} Item</div></div>',
                unsafe_allow_html=True,
            )
        with m2:
            st.markdown(
                f'<div class="metric-box"><div style="font-size:13px;font-weight:600;color:#64748B;">Total Pengeluaran Dana</div>'
                f'<div style="font-size:24px;font-weight:800;color:#7C2D12;">{fmt_rp_cell(total_pengeluaran, is_dash_if_zero=False)}</div></div>',
                unsafe_allow_html=True,
            )
        with m3:
            st.markdown(
                f'<div class="metric-box"><div style="font-size:13px;font-weight:600;color:#64748B;">Terbilang Rupiah</div>'
                f'<div style="font-size:16px;font-weight:700;color:#0F172A;margin-top:4px;">{teks_terbilang}</div></div>',
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================================================
    # TAB 3: PREVIEW PERSIS ASLI, SIMPAN GSHEET & DOWNLOAD CENTER (WORD & PDF)
    # ==========================================================================
    with tab3:
        st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
        st.markdown("#### 🖋️ Pengaturan Penutup, Tanda Tangan & Lampiran Dokumentasi")

        cp1, cp2 = st.columns([1.3, 1])
        with cp1:
            st.checkbox("Tampilkan Paragraf Ucapan Terima Kasih Donatur", key="tampilkan_ucapan_donatur")
            if st.session_state.get("tampilkan_ucapan_donatur", True):
                st.text_area("Redaksi Ucapan Terima Kasih Donatur", key="teks_ucapan_donatur", height=95)
            st.text_input("Catatan Dokumentasi Instagram", key="catatan_ig")

        with cp2:
            st.text_input("Jabatan Baris 1", key="jabatan_baris_1")
            st.text_input("Jabatan Baris 2", key="jabatan_baris_2")
            st.text_input("Nama Penandatangan", key="nama_penandatangan")
            st.checkbox("Sertakan Gambar Tanda Tangan Asli (Ny. Rini Johan Pandores)", key="tampilkan_ttd")
            up_sig = st.file_uploader("Ganti Gambar Tanda Tangan (PNG/JPG Opsional)", type=["png", "jpg", "jpeg"], key="up_custom_sig")

        custom_sig_bytes = None
        if up_sig is not None:
            custom_sig_bytes = up_sig.getvalue()

        with st.expander("📸 Lampiran Foto Dokumentasi / Nota Kwitansi Tambahan (Auto-Grid 2x2 di Halaman Baru — Opsional)", expanded=False):
            up_photos = st.file_uploader(
                "Upload Foto Kegiatan atau Kwitansi (Bisa pilih beberapa foto sekaligus)",
                type=["png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key="up_lampiran_fotos",
            )
            attachment_images = []
            if up_photos:
                for idx_p, f_img in enumerate(up_photos, start=1):
                    attachment_images.append({
                        "bytes": f_img.getvalue(),
                        "caption": f"Dokumentasi / Kwitansi {idx_p}: {f_img.name}",
                    })
                st.success(f"{len(attachment_images)} foto siap dilampirkan dalam format Grid 2x2 rapi.")
        st.markdown("</div>", unsafe_allow_html=True)

        # Generate Dokumen Word & PDF secara Real-Time
        current_payload = collect_current_payload()
        docx_bytes = generate_lpj_docx_bytes(
            current_payload,
            custom_sig_bytes=custom_sig_bytes,
            attachment_images=attachment_images,
        )
        pdf_bytes, pdf_status_msg = convert_docx_bytes_to_pdf_bytes(docx_bytes)

        clean_keg = "".join(ch if ch.isalnum() else "_" for ch in current_payload.get("nama_kegiatan", "LPJ_DWP")[:35]).strip("_")
        file_base = f"LPJ_DWP_{clean_keg or 'MALANG'}_{date.today().strftime('%Y%m%d')}"

        # Tombol Aksi Utama (Simpan GSheet + Direct Download Word & PDF)
        st.markdown('<div class="dwp-card">', unsafe_allow_html=True)
        st.markdown("#### 🚀 Pusat Simpan Database & Unduh Dokumen Resmi (Anti-Error 404)")

        ca1, ca2, ca3 = st.columns([1.1, 1.15, 1.15])
        with ca1:
            if st.button("💾 Simpan / Update ke Google Sheets", use_container_width=True):
                ok_save, msg_save = save_lpj_to_storage(current_payload, sheet_id_in, creds_str)
                if ok_save:
                    st.success(msg_save)
                else:
                    st.error(msg_save)

        with ca2:
            html_dl_docx = make_direct_download_html(
                file_bytes=docx_bytes,
                filename=f"{file_base}.docx",
                mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                label="📄 Unduh File Word Resmi (.docx)",
                bg_gradient="linear-gradient(135deg, #1E3A8A 0%, #1D4ED8 100%)",
            )
            st.markdown(html_dl_docx, unsafe_allow_html=True)

        with ca3:
            if pdf_bytes:
                html_dl_pdf = make_direct_download_html(
                    file_bytes=pdf_bytes,
                    filename=f"{file_base}.pdf",
                    mime_type="application/pdf",
                    label="📕 Unduh File PDF Siap Sebar (.pdf)",
                    bg_gradient="linear-gradient(135deg, #7C2D12 0%, #9A3412 55%, #B45309 100%)",
                )
                st.markdown(html_dl_pdf, unsafe_allow_html=True)
            else:
                st.info(pdf_status_msg)
        st.markdown("</div>", unsafe_allow_html=True)

        # Live Preview Visual HTML Kertas A4
        st.markdown("#### 🖥️ Live Preview Visual Kertas A4 (Dilengkapi Tombol Cetak PDF Browser)")
        a4_html = build_a4_live_preview_html(current_payload, custom_sig_bytes=custom_sig_bytes)
        st.components.v1.html(a4_html, height=1080, scrolling=True)


if __name__ == "__main__":
    main()
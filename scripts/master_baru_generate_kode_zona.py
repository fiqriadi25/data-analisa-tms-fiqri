"""Master Baru & Generate Kode Zona Baru.

Script ini ditujukan untuk dijalankan di Google Colab. Data pada sheet
``DATA MASTER`` akan diberi kode zona pada kolom ``P2`` berdasarkan kombinasi
OP, kecamatan, dan seluruh nilai UJP yang tersedia.
"""

from __future__ import annotations

import io
import re
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


SHEET_NAME = "DATA MASTER"
OUTPUT_FILE = "data_master_generate_zona.xlsx"
MAX_ZONE_LENGTH = 15

REQUIRED_COLUMNS = ("OP", "KECAMATAN")
DISPLAY_COLUMNS = ("KODE TOKO", "NAMA PENERIMA", "OP", "KECAMATAN", "P2")
UJP_COLUMNS = (
    "UJP CDD",
    "UJP CDD-L",
    "UJP CDD-C",
    "UJP CDE",
    "UJP CDE-L",
    "UJP CDE-C",
    "UJP L300",
    "UJP FUSO",
    "UJP CDD-LC",
    "UJP CDE-LC",
    "UJP Tronton",
    "UJP Wingbox",
    "UJP BlindVan",
)


def normalize_text(value: Any) -> str:
    """Ubah nilai menjadi teks kapital alfanumerik tanpa spasi."""
    if pd.isna(value):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(value).strip().upper())


def normalize_ujp(value: Any) -> Any:
    """Samakan representasi nilai UJP agar NaN, 1, dan 1.0 konsisten."""
    if pd.isna(value) or value == "":
        return 0
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def generate_kecamatan_code(kecamatan_name: Any) -> str:
    """Buat kode kecamatan: maksimal 4 karakter awal + 3 karakter akhir."""
    kecamatan = normalize_text(kecamatan_name)
    if not kecamatan:
        raise ValueError("Nama kecamatan tidak boleh kosong.")
    return kecamatan if len(kecamatan) <= 7 else kecamatan[:4] + kecamatan[-3:]


def validate_columns(df: pd.DataFrame) -> list[str]:
    """Validasi kolom wajib dan kembalikan daftar kolom UJP yang tersedia."""
    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"Kolom wajib tidak ditemukan: {', '.join(missing)}")

    available_ujp = [column for column in UJP_COLUMNS if column in df.columns]
    if not available_ujp:
        raise ValueError("Tidak ada kolom UJP yang dikenali pada file input.")
    return available_ujp


def build_ujp_pattern(row: pd.Series, ujp_columns: Iterable[str]) -> tuple[Any, ...]:
    """Buat pola UJP yang stabil dan hashable untuk satu baris."""
    return tuple(normalize_ujp(row[column]) for column in ujp_columns)


def fit_zone_code(base_code: str, suffix: int, max_length: int) -> str:
    """Gabungkan base dan suffix dengan panjang maksimum yang ditentukan."""
    suffix_text = str(suffix)
    available_base_length = max_length - len(suffix_text)
    if available_base_length < 1:
        raise ValueError("Suffix terlalu panjang untuk batas kode zona.")
    return f"{base_code[:available_base_length]}{suffix_text}"


def generate_zone_codes(
    df: pd.DataFrame,
    ujp_columns: Iterable[str],
    max_length: int = MAX_ZONE_LENGTH,
) -> pd.Series:
    """Generate kode zona konsisten berdasarkan OP, kecamatan, dan pola UJP.

    Urutan suffix mengikuti urutan pertama pola muncul di data. Counter dipakai
    per base code agar singkatan kecamatan yang kebetulan sama tidak menghasilkan
    kode zona ganda.
    """
    ujp_columns = list(ujp_columns)
    pattern_to_code: dict[tuple[Any, ...], str] = {}
    base_counters: defaultdict[str, int] = defaultdict(int)
    used_codes: dict[str, tuple[Any, ...]] = {}
    result: list[str] = []

    for row_number, (_, row) in enumerate(df.iterrows(), start=2):
        op_code = normalize_text(row["OP"])
        if not op_code:
            raise ValueError(f"Nilai OP kosong pada baris Excel {row_number}.")

        kecamatan_raw = str(row["KECAMATAN"]).strip()
        kecamatan_code = generate_kecamatan_code(kecamatan_raw)
        ujp_pattern = build_ujp_pattern(row, ujp_columns)
        pattern_key = (op_code, kecamatan_raw.upper(), *ujp_pattern)

        if pattern_key in pattern_to_code:
            result.append(pattern_to_code[pattern_key])
            continue

        base_code = f"{op_code}{kecamatan_code}"
        while True:
            base_counters[base_code] += 1
            zone_code = fit_zone_code(base_code, base_counters[base_code], max_length)
            owner = used_codes.get(zone_code)
            if owner is None or owner == pattern_key:
                break

        pattern_to_code[pattern_key] = zone_code
        used_codes[zone_code] = pattern_key
        result.append(zone_code)

    return pd.Series(result, index=df.index, dtype="string")


def process_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Validasi data dan tambahkan hasil kode zona pada kolom P2."""
    available_ujp = validate_columns(df)
    processed = df.copy()
    processed["P2"] = generate_zone_codes(processed, available_ujp)
    return processed, available_ujp


def validate_result(df: pd.DataFrame, ujp_columns: Iterable[str]) -> None:
    """Pastikan satu P2 tidak dimiliki oleh lebih dari satu pola data."""
    key_columns = ["P2", "OP", "KECAMATAN", *ujp_columns]
    unique_patterns = df[key_columns].copy()
    for column in ujp_columns:
        unique_patterns[column] = unique_patterns[column].map(normalize_ujp)
    unique_patterns = unique_patterns.drop_duplicates()

    duplicated = unique_patterns[unique_patterns.duplicated("P2", keep=False)]
    if not duplicated.empty:
        raise ValueError(
            "Ditemukan kode P2 yang digunakan oleh lebih dari satu pola data:\n"
            f"{duplicated.to_string(index=False)}"
        )

    too_long = df[df["P2"].str.len() > MAX_ZONE_LENGTH]
    if not too_long.empty:
        raise ValueError("Ditemukan kode zona dengan panjang lebih dari 15 karakter.")


def print_summary(df: pd.DataFrame, ujp_columns: list[str]) -> None:
    """Tampilkan ringkasan hasil proses."""
    print("\n=== RINGKASAN HASIL ===")
    print(f"Total baris data       : {len(df):,}")
    print(f"Total kode zona unik   : {df['P2'].nunique():,}")
    print(f"Total kecamatan unik   : {df['KECAMATAN'].nunique():,}")
    print(f"Kolom UJP digunakan    : {len(ujp_columns)}")
    print(f"Panjang kode maksimum  : {df['P2'].str.len().max()}")

    preview_columns = [column for column in DISPLAY_COLUMNS if column in df.columns]
    preview_columns.extend(ujp_columns)
    print("\nContoh hasil (20 baris pertama):")
    print(df[preview_columns].head(20).to_string(index=False))


def process_excel_bytes(
    file_content: bytes,
    sheet_name: str = SHEET_NAME,
) -> tuple[pd.DataFrame, list[str]]:
    """Baca byte Excel, proses, validasi, lalu kembalikan hasil."""
    source_df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet_name)
    processed_df, available_ujp = process_dataframe(source_df)
    validate_result(processed_df, available_ujp)
    return processed_df, available_ujp


def run_colab() -> None:
    """Jalankan alur upload dan download khusus Google Colab."""
    try:
        from google.colab import files
    except ImportError as exc:
        raise RuntimeError("Fungsi ini hanya dapat dijalankan di Google Colab.") from exc

    print(f"Silakan upload file Excel yang memiliki sheet '{SHEET_NAME}'.")
    uploaded = files.upload()
    if not uploaded:
        print("Upload dibatalkan. Tidak ada file yang diproses.")
        return

    file_name, file_content = next(iter(uploaded.items()))
    print(f'File "{file_name}" berhasil diunggah ({len(file_content):,} bytes).')

    try:
        result_df, available_ujp = process_excel_bytes(file_content)
        result_df.to_excel(OUTPUT_FILE, index=False)
        print_summary(result_df, available_ujp)
        print(f"\nFile hasil berhasil dibuat: {OUTPUT_FILE}")
        files.download(OUTPUT_FILE)
    except Exception as exc:
        print(f"\nProses gagal: {exc}")
        raise


if __name__ == "__main__":
    run_colab()


"""Analisis update UJP master dan penyesuaian kode zona toko.

Script hanya mengubah ``Kode Zona`` bila diperlukan dan menambahkan kolom
``Keterangan``. Nilai pada kolom UJP asli tidak diubah.

Lokal:
    python scripts/update_ujp_master.py data_master.xlsx -o hasil_update_ujp.xlsx

Google Colab:
    %run scripts/update_ujp_master.py
"""

from __future__ import annotations

import argparse
import io
import re
import sys
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import pandas as pd


SHEET_NAME = "DATA MASTER"
MAX_ZONE_LENGTH = 15

UJP_ALL_COLS = (
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
)

UJP_UPDATE_MAP = {
    "UJP CDE-LC": "UPDATE UJP CDE-LC",
    "UJP CDD-C": "UPDATE UJP CDD-C",
}

REQUIRED_ID_COLUMNS = ("Kode Toko", "Kode Zona", "Operating Point", "Kecamatan")


def clean(value: Any) -> str:
    """Ubah nilai menjadi teks bersih; NaN menjadi string kosong."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def numeric_ujp(value: Any) -> Decimal:
    """Normalisasi nilai UJP kosong, numerik, dan format angka Indonesia."""
    if pd.isna(value) or clean(value) == "":
        return Decimal(0)

    if isinstance(value, (int, float, Decimal)):
        raw = str(value)
    else:
        raw = clean(value).replace(" ", "")
        if "," in raw:
            raw = raw.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+", raw):
            raw = raw.replace(".", "")

    try:
        result = Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Nilai UJP tidak valid: {value!r}") from exc

    if not result.is_finite() or result < 0:
        raise ValueError(f"Nilai UJP negatif/tidak terhingga: {value!r}")
    return result


def resolve_new_value(row: pd.Series, original_column: str, old_value: Decimal) -> Decimal:
    """Ambil nilai update; kosong atau ``Tidak update`` mempertahankan nilai lama."""
    update_column = UJP_UPDATE_MAP.get(original_column)
    if update_column is None:
        return old_value

    raw = row[update_column]
    if pd.isna(raw) or clean(raw).lower() == "tidak update":
        return old_value
    return numeric_ujp(raw)


def build_old_signature(row: pd.Series) -> tuple[Decimal, ...]:
    """Buat signature UJP sebelum update."""
    return tuple(numeric_ujp(row[column]) for column in UJP_ALL_COLS)


def build_new_signature(row: pd.Series) -> tuple[Decimal, ...]:
    """Buat signature UJP setelah mempertimbangkan kolom UPDATE UJP."""
    old_signature = build_old_signature(row)
    return tuple(
        resolve_new_value(row, column, old_signature[index])
        for index, column in enumerate(UJP_ALL_COLS)
    )


def parse_zone_suffix(zone_code: Any) -> tuple[str, int | None]:
    """Pisahkan prefix kode zona dan suffix angka."""
    code = clean(zone_code)
    match = re.fullmatch(r"(.*?)(\d+)", code)
    return (match.group(1), int(match.group(2))) if match else (code, None)


def next_available_suffix(existing_suffixes: set[int], current_suffix: int | None) -> int:
    """Cari suffix berikutnya yang belum digunakan."""
    candidate = (current_suffix or 0) + 1
    while candidate in existing_suffixes:
        candidate += 1
    return candidate


def validate_dataframe(dataframe: pd.DataFrame) -> None:
    """Validasi struktur dan nilai identitas wajib."""
    required = [*REQUIRED_ID_COLUMNS, *UJP_ALL_COLS, *UJP_UPDATE_MAP.values()]
    missing = [column for column in required if column not in dataframe.columns]
    if missing:
        raise ValueError(f"Kolom wajib tidak ditemukan: {missing}")

    empty_rows: list[str] = []
    for column in REQUIRED_ID_COLUMNS:
        mask = dataframe[column].map(clean).eq("")
        if mask.any():
            excel_rows = [str(index + 2) for index in dataframe.index[mask][:10]]
            empty_rows.append(f"{column} kosong pada baris {', '.join(excel_rows)}")
    if empty_rows:
        raise ValueError("; ".join(empty_rows))

    invalid_zones = [
        clean(code)
        for code in dataframe["Kode Zona"]
        if parse_zone_suffix(code)[1] is None
        or len(clean(code)) > MAX_ZONE_LENGTH
    ]
    if invalid_zones:
        raise ValueError(
            "Kode Zona harus memiliki suffix angka dan maksimal 15 karakter: "
            f"{sorted(set(invalid_zones))[:10]}"
        )


def process_dataframe(source: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Analisis perubahan UJP dan kembalikan dataframe hasil serta ringkasan."""
    dataframe = source.copy()
    dataframe.columns = [str(column).strip() for column in dataframe.columns]
    validate_dataframe(dataframe)

    try:
        dataframe["_old_sig"] = dataframe.apply(build_old_signature, axis=1)
        dataframe["_new_sig"] = dataframe.apply(build_new_signature, axis=1)
    except ValueError as exc:
        raise ValueError(f"Gagal membaca nilai UJP: {exc}") from exc

    dataframe["_row_changed"] = dataframe["_old_sig"] != dataframe["_new_sig"]

    zone_type: dict[str, str] = {}
    zone_majority_signature: dict[str, tuple[Decimal, ...]] = {}

    for zone_code, group in dataframe.groupby("Kode Zona", sort=False):
        changed_count = int(group["_row_changed"].sum())
        total_count = len(group)
        signature_counts = group["_new_sig"].value_counts(sort=True)
        unique_new_count = len(signature_counts)

        if changed_count == 0:
            zone_type[zone_code] = "A"
        elif unique_new_count == 1:
            zone_type[zone_code] = "B1" if changed_count == total_count else "B2"
        else:
            zone_type[zone_code] = "C"

        zone_majority_signature[zone_code] = signature_counts.index[0]

    dataframe["_zona_type"] = dataframe["Kode Zona"].map(zone_type)
    dataframe["_zona_maj_sig"] = dataframe["Kode Zona"].map(
        zone_majority_signature
    )

    parsed = dataframe["Kode Zona"].map(parse_zone_suffix)
    dataframe["_zona_prefix"] = parsed.map(lambda item: item[0])
    dataframe["_zona_suffix"] = parsed.map(lambda item: item[1])
    dataframe["_op_key"] = dataframe["Operating Point"].map(
        lambda value: clean(value).upper()
    )
    dataframe["_kec_key"] = dataframe["Kecamatan"].map(
        lambda value: clean(value).upper()
    )

    suffix_lookup: dict[tuple[str, str, str], set[int]] = {}
    for _, row in dataframe.iterrows():
        key = (row["_op_key"], row["_kec_key"], row["_zona_prefix"])
        suffix_lookup.setdefault(key, set()).add(int(row["_zona_suffix"]))

    zone_signature_map: dict[tuple[str, tuple[Decimal, ...]], str] = {}
    new_zones_log: list[dict[str, str]] = []

    for zone_code, group in dataframe.groupby("Kode Zona", sort=False):
        if zone_type[zone_code] != "C":
            continue

        op = group["_op_key"].iloc[0]
        kecamatan = group["_kec_key"].iloc[0]
        prefix = group["_zona_prefix"].iloc[0]
        current_suffix = int(group["_zona_suffix"].iloc[0])
        lookup_key = (op, kecamatan, prefix)
        majority_signature = zone_majority_signature[zone_code]

        for signature in group["_new_sig"].drop_duplicates():
            if signature == majority_signature:
                continue

            cache_key = (zone_code, signature)
            if cache_key in zone_signature_map:
                continue

            new_suffix = next_available_suffix(
                suffix_lookup.setdefault(lookup_key, set()), current_suffix
            )
            new_zone_code = f"{prefix}{new_suffix}"
            if len(new_zone_code) > MAX_ZONE_LENGTH:
                raise ValueError(
                    f"Kode zona baru {new_zone_code} melebihi {MAX_ZONE_LENGTH} karakter"
                )

            zone_signature_map[cache_key] = new_zone_code
            suffix_lookup[lookup_key].add(new_suffix)
            new_zones_log.append(
                {
                    "zona_lama": clean(zone_code),
                    "zona_baru": new_zone_code,
                    "operating_point": clean(group["Operating Point"].iloc[0]),
                    "kecamatan": clean(group["Kecamatan"].iloc[0]),
                }
            )

    def apply_result(row: pd.Series) -> tuple[str, str]:
        zone_code = clean(row["Kode Zona"])
        classification = row["_zona_type"]
        new_signature = row["_new_sig"]
        majority_signature = row["_zona_maj_sig"]
        changed = bool(row["_row_changed"])

        if classification == "A":
            return zone_code, "TIDAK UPDATE"
        if classification == "B1":
            return zone_code, "UPDATE UJP"
        if classification == "B2":
            return zone_code, "NO UJP UPDATE"

        if new_signature == majority_signature:
            return zone_code, "UPDATE UJP" if changed else "TIDAK UPDATE"

        new_zone = zone_signature_map[(zone_code, new_signature)]
        status = (
            "NEW ZONE - UPDATE UJP - MOVE ZONE"
            if changed
            else "NEW ZONE - MOVE ZONE"
        )
        return new_zone, status

    result = dataframe.apply(apply_result, axis=1, result_type="expand")
    dataframe["_zona_baru"] = result[0]
    dataframe["Keterangan"] = result[1]

    # Gabungkan zona dengan OP, kecamatan, prefix, dan signature akhir yang sama.
    dataframe["_temp_zone"] = dataframe["_zona_baru"]
    merge_columns = ["_op_key", "_kec_key", "_zona_prefix", "_new_sig"]
    for (op, kecamatan, prefix, new_signature), group in dataframe.groupby(
        merge_columns, sort=False
    ):
        zones = list(dict.fromkeys(group["_temp_zone"]))
        if len(zones) <= 1:
            continue

        main_zone = min(
            zones,
            key=lambda zone: (
                parse_zone_suffix(zone)[1] or 0,
                clean(zone),
            ),
        )
        mask = (
            dataframe["_op_key"].eq(op)
            & dataframe["_kec_key"].eq(kecamatan)
            & dataframe["_zona_prefix"].eq(prefix)
            & dataframe["_new_sig"].eq(new_signature)
            & dataframe["_temp_zone"].ne(main_zone)
        )
        dataframe.loc[mask, "_temp_zone"] = main_zone
        dataframe.loc[mask, "Keterangan"] = "MOVE ZONE"

    dataframe["Kode Zona"] = dataframe["_temp_zone"]

    temporary_columns = [
        column for column in dataframe.columns if column.startswith("_")
    ]
    dataframe.drop(columns=temporary_columns, inplace=True, errors="ignore")

    status_counts = Counter(dataframe["Keterangan"])
    type_counts = Counter(zone_type.values())
    summary = {
        "rows": len(dataframe),
        "status": dict(status_counts),
        "zone_types": {
            "A": type_counts.get("A", 0),
            "B1": type_counts.get("B1", 0),
            "B2": type_counts.get("B2", 0),
            "C": type_counts.get("C", 0),
        },
        "new_zones": new_zones_log,
    }
    return dataframe, summary


def process_excel_bytes(file_content: bytes) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Baca sheet DATA MASTER dari byte Excel dan proses datanya."""
    source = pd.read_excel(io.BytesIO(file_content), sheet_name=SHEET_NAME)
    return process_dataframe(source)


def update_workbook(input_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    """Proses workbook lokal dan simpan hasil ke file baru."""
    input_path = Path(input_path)
    output_path = Path(output_path)
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Output harus berbeda dari file input")
    if not input_path.exists():
        raise FileNotFoundError(f"File input tidak ditemukan: {input_path}")

    result, summary = process_excel_bytes(input_path.read_bytes())
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(output_path, sheet_name=SHEET_NAME, index=False)
    summary["output"] = str(output_path)
    return summary


def print_summary(summary: dict[str, Any]) -> None:
    """Tampilkan statistik hasil proses."""
    print("\n" + "=" * 64)
    print("STATISTIK PERUBAHAN (NILAI UJP ASLI TIDAK DIUBAH)")
    print("=" * 64)
    for status, count in summary["status"].items():
        print(f"  {status:<45}: {count:>5} toko")
    print(f"  {'-' * 45}   {'-' * 5}")
    print(f"  {'TOTAL':<45}: {summary['rows']:>5} toko")

    print("\nRingkasan klasifikasi zona:")
    labels = {
        "A": "TIDAK UPDATE",
        "B1": "UPDATE UJP (semua berubah)",
        "B2": "NO UJP UPDATE (campuran)",
        "C": "SPLIT -> ZONA BARU",
    }
    for zone_type in ("A", "B1", "B2", "C"):
        count = summary["zone_types"].get(zone_type, 0)
        print(f"  Tipe {zone_type} ({labels[zone_type]}): {count} zona")

    if summary["new_zones"]:
        print(f"\nZona baru dibuat ({len(summary['new_zones'])}):")
        for item in summary["new_zones"]:
            print(
                f"  {item['zona_lama']} -> {item['zona_baru']} "
                f"(OP={item['operating_point']}, Kec={item['kecamatan']})"
            )


def run_colab() -> None:
    """Jalankan upload, proses, dan download di Google Colab."""
    from google.colab import files

    print(f"Silakan upload satu file Excel dengan sheet '{SHEET_NAME}'.")
    uploaded = files.upload()
    excel_files = [name for name in uploaded if name.lower().endswith(".xlsx")]
    if len(excel_files) != 1:
        raise ValueError("Upload tepat satu file .xlsx")

    file_name = excel_files[0]
    result, summary = process_excel_bytes(uploaded[file_name])
    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Updated_UJP_Master_{timestamp}.xlsx"
    result.to_excel(output_file, sheet_name=SHEET_NAME, index=False)
    print_summary(summary)
    print(f"\nSelesai. File dibuat: {output_file}")
    print(
        "Nilai UJP asli tidak diubah; hasil hanya menyesuaikan Kode Zona "
        "dan menambahkan Keterangan."
    )
    files.download(output_file)


def main() -> None:
    """Entry point untuk Colab dan command line lokal."""
    try:
        from google.colab import files as colab_files
    except ImportError:
        colab_files = None

    if colab_files is not None:
        run_colab()
        return

    if "-f" in sys.argv[1:] and "ipykernel" in sys.modules:
        raise RuntimeError(
            "Di Jupyter, panggil update_workbook('input.xlsx', 'hasil.xlsx') "
            "secara langsung."
        )

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help=f"Excel dengan sheet {SHEET_NAME}")
    parser.add_argument("-o", "--output", default="hasil_update_ujp.xlsx")
    args = parser.parse_args()
    summary = update_workbook(args.input, args.output)
    print_summary(summary)
    print(f"\nOutput: {summary['output']}")


if __name__ == "__main__":
    main()


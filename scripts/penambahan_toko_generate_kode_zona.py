"""Penambahan toko dan generate kode zona untuk OP yang sudah extend.

Lokal:
    python penambahan_toko_generate_kode_zona.py data_master_olah.xlsx \
        -o hasil_zona.xlsx

Google Colab:
    Impor modul ini lalu panggil ``run_colab()``. Jika kode ditempel sebagai
    satu sel, dialog upload akan dibuka otomatis.

Input tidak ditimpa. Hasil selalu disimpan sebagai file Excel baru.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterator

import pandas as pd


DATA_MASTER_SHEET = "DATA MASTER"
MASTER_TOKO_SHEET = "MASTER TOKO"
AUDIT_SHEET = "AUDIT REFERENSI"
MAX_CODE_LENGTH = 15

UJP_TYPES = (
    "CDD",
    "CDD-L",
    "CDD-C",
    "CDE",
    "CDE-L",
    "CDE-C",
    "L300",
    "FUSO",
    "CDD-LC",
    "CDE-LC",
    "Tronton",
    "Wingbox",
    "Traga",
)
UJP_COLS = tuple(f"UJP {kind}" for kind in UJP_TYPES)


def clean(value: Any) -> str:
    """Ubah nilai menjadi teks bersih; NaN menjadi string kosong."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def canonical(value: Any) -> str:
    """Buat key stabil tanpa perbedaan kapitalisasi, spasi, tanda baca, dan aksen."""
    normalized = unicodedata.normalize("NFKD", clean(value))
    return re.sub(r"[^A-Z0-9]", "", normalized.upper())


def tariff(value: Any) -> Decimal | None:
    """Normalisasi tarif; ``None`` berarti kosong, sedangkan nol tetap eksplisit."""
    if pd.isna(value) or str(value).strip() == "":
        return None

    if isinstance(value, (int, float, Decimal)):
        raw = str(value)
    else:
        raw = str(value).strip().replace(" ", "")
        # Format umum Indonesia: 68.000,00 / 68.000 / 68000,00.
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


def signature(row: pd.Series, columns: tuple[str, ...]) -> tuple[Decimal, ...]:
    """Buat profil UJP; sel kosong dan nol sama-sama berarti tipe tidak dipakai."""
    return tuple(tariff(row[column]) or Decimal(0) for column in columns)


def suffix_of(code: str) -> int | None:
    """Ambil suffix angka di akhir kode zona."""
    match = re.search(r"(\d+)$", code)
    return int(match.group(1)) if match else None


def prefix_of(code: str) -> str:
    """Ambil prefix kode zona dengan menghapus suffix angka."""
    return re.sub(r"\d+$", "", code)


def abbreviation_candidates(kecamatan: str, width: int) -> Iterator[str]:
    """Hasilkan kandidat singkatan deterministik dari yang paling mudah dikenali."""
    name = canonical(kecamatan)
    words = re.findall(
        r"[A-Z0-9]+", unicodedata.normalize("NFKD", kecamatan.upper())
    )
    seen: set[str] = set()
    candidates = [name[:width]]

    if len(words) > 1:
        for first_size in range(1, width):
            candidates.append((words[0][:first_size] + "".join(words[1:]))[:width])

    # Jendela geser membedakan kecamatan dengan karakter awal yang sama.
    candidates.extend(name[index : index + width] for index in range(1, len(name)))

    # Fallback hash deterministik bila singkatan yang mudah dibaca telah dipakai.
    digest = hashlib.sha256(name.encode()).hexdigest().upper()
    candidates.extend(
        (name[: max(0, width - 3)] + digest[index : index + 3])[:width]
        for index in range(len(digest) - 2)
    )

    for candidate in candidates:
        if candidate and candidate not in seen:
            seen.add(candidate)
            yield candidate


def choose_prefix(
    op: str,
    kecamatan: str,
    reserved: dict[str, tuple[str, str]],
    suffix_digits: int = 1,
) -> str:
    """Pilih prefix unik untuk OP + kecamatan dengan panjang total maksimal 15."""
    available = MAX_CODE_LENGTH - len(op) - suffix_digits
    if available < 2:
        raise ValueError(
            f"OP {op!r} terlalu panjang untuk membuat kode zona <=15 karakter"
        )

    for candidate in abbreviation_candidates(kecamatan, available):
        prefix = op + candidate
        if prefix not in reserved:
            reserved[prefix] = (op, kecamatan)
            return prefix

    raise ValueError(f"Tidak ada singkatan zona unik untuk {op}/{kecamatan}")


def validate_required_columns(
    sheet_name: str,
    dataframe: pd.DataFrame,
    required: tuple[str, ...],
) -> None:
    """Pastikan seluruh kolom wajib tersedia."""
    absent = sorted(set(required) - set(dataframe.columns))
    if absent:
        raise ValueError(f"{sheet_name}: kolom wajib tidak ditemukan: {absent}")


def generate(input_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    """Proses workbook master toko dan simpan hasil sebagai workbook baru."""
    input_path = Path(input_path)
    output_path = Path(output_path)

    if input_path.resolve() == output_path.resolve():
        raise ValueError("Output harus file lain; input tidak boleh ditimpa")
    if not input_path.exists():
        raise FileNotFoundError(f"File input tidak ditemukan: {input_path}")

    data_master = pd.read_excel(
        input_path, sheet_name=DATA_MASTER_SHEET, dtype=object
    )
    master_toko = pd.read_excel(
        input_path, sheet_name=MASTER_TOKO_SHEET, dtype=object
    )
    data_master.columns = [str(column).strip() for column in data_master.columns]
    master_toko.columns = [str(column).strip() for column in master_toko.columns]

    validate_required_columns(
        DATA_MASTER_SHEET,
        data_master,
        ("OP", "KECAMATAN", "KODE ZONA", "NAMA ZONA"),
    )
    validate_required_columns(
        MASTER_TOKO_SHEET,
        master_toko,
        ("Operating Point", "Kecamatan", "Zona"),
    )

    common_ujp = tuple(
        column
        for column in UJP_COLS
        if column in data_master.columns and column in master_toko.columns
    )
    if not common_ujp:
        raise ValueError("Tidak ada kolom UJP yang sama pada kedua sheet")

    # Tipe yang nol/kosong pada seluruh master suatu OP belum memiliki referensi.
    op_columns: dict[str, tuple[str, ...]] = {}
    grouped_master = master_toko.groupby(master_toko["Operating Point"].map(canonical))
    for op, part in grouped_master:
        if op:
            op_columns[op] = tuple(
                column
                for column in common_ujp
                if any(
                    tariff(value) not in (None, Decimal(0))
                    for value in part[column]
                )
            )

    zones: defaultdict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    used_codes: dict[str, tuple[str, str]] = {}
    reserved: dict[str, tuple[str, str]] = {}
    suffix_max: defaultdict[str, int] = defaultdict(int)
    warnings: list[str] = []
    master_codes: set[str] = set()
    master_districts: set[tuple[str, str]] = set()

    def register(
        code: Any,
        op: str,
        kecamatan: str,
        sig: tuple[Decimal, ...] | None = None,
        source: str = MASTER_TOKO_SHEET,
    ) -> bool:
        normalized_code = canonical(code)
        if not normalized_code:
            return False

        suffix = suffix_of(normalized_code)
        if (
            len(normalized_code) > MAX_CODE_LENGTH
            or suffix is None
            or not normalized_code.startswith(op)
        ):
            warnings.append(
                f"{source}: kode {normalized_code} tidak sesuai OP/format/panjang; "
                "tidak dipakai sebagai referensi"
            )
            return False

        owner = used_codes.get(normalized_code)
        if owner is not None and owner != (op, kecamatan):
            warnings.append(
                f"{source}: kode {normalized_code} digunakan dua kecamatan; "
                "referensi kedua diabaikan"
            )
            return False

        prefix = prefix_of(normalized_code)
        prefix_owner = reserved.get(prefix)
        if prefix_owner is not None and prefix_owner != (op, kecamatan):
            warnings.append(
                f"{source}: prefix {prefix} digunakan dua kecamatan; "
                "prefix tetap dicadangkan"
            )

        reserved.setdefault(prefix, (op, kecamatan))
        used_codes[normalized_code] = (op, kecamatan)
        suffix_max[prefix] = max(suffix_max[prefix], suffix)

        existing = zones[(op, kecamatan)]
        if sig is not None and not any(
            zone["code"] == normalized_code and zone["sig"] == sig
            for zone in existing
        ):
            existing.append({"code": normalized_code, "prefix": prefix, "sig": sig})
        return True

    # Daftarkan zona referensi dari MASTER TOKO.
    for index, row in master_toko.iterrows():
        op = canonical(row["Operating Point"])
        kecamatan = canonical(row["Kecamatan"])
        code = clean(row["Zona"])
        if not op or not kecamatan or not code:
            warnings.append(
                f"{MASTER_TOKO_SHEET} baris {index + 2}: "
                "OP/Kecamatan/Zona kosong; diabaikan"
            )
            continue

        columns = op_columns.get(op, ())
        if not columns:
            warnings.append(
                f"{MASTER_TOKO_SHEET} baris {index + 2}: "
                f"OP {op} tanpa tarif UJP aktif"
            )
            continue

        try:
            sig = signature(row, columns)
        except ValueError as exc:
            warnings.append(f"{MASTER_TOKO_SHEET} baris {index + 2}: {exc}; diabaikan")
            continue

        if register(code, op, kecamatan, sig):
            normalized_code = canonical(code)
            master_codes.add(normalized_code)
            master_districts.add((op, kecamatan))

    # Reservasi semua kode yang telah terisi pada DATA MASTER sebelum generate.
    for index, row in data_master.iterrows():
        op = canonical(row["OP"])
        kecamatan = canonical(row["KECAMATAN"])
        code = clean(row["KODE ZONA"])
        if op and kecamatan and code:
            register(
                code,
                op,
                kecamatan,
                source=f"{DATA_MASTER_SHEET} baris {index + 2}",
            )

    output_codes: list[str] = []
    output_names: list[str] = []
    statuses: list[str] = []
    notes: list[str] = []
    counts: Counter[str] = Counter()
    generated: dict[tuple[Any, ...], str] = {}
    generated_status: dict[str, str] = {}

    for index, row in data_master.iterrows():
        line = index + 2
        op = canonical(row["OP"])
        kecamatan = canonical(row["KECAMATAN"])
        old_code = clean(row["KODE ZONA"])
        code = ""
        name = ""
        status = ""
        note = ""

        try:
            if not op or not kecamatan:
                raise ValueError("OP/Kecamatan kosong")

            columns = op_columns.get(op, ())
            if not columns:
                raise ValueError(
                    f"OP {op} tidak punya referensi UJP aktif di {MASTER_TOKO_SHEET}"
                )

            sig = signature(row, columns)
            if not any(value != 0 for value in sig):
                raise ValueError(
                    "Semua UJP yang didukung OP kosong/nol; perlu penentuan tarif"
                )

            unavailable = [
                column
                for column in UJP_COLS
                if column in data_master.columns
                and column not in columns
                and tariff(row[column]) not in (None, Decimal(0))
            ]
            if unavailable:
                note = (
                    "Tipe tanpa tarif referensi OP (tidak dibandingkan): "
                    + ", ".join(unavailable)
                )

            if old_code:
                code = canonical(old_code)
                status = "SUDAH ADA"
                if (
                    len(code) > MAX_CODE_LENGTH
                    or suffix_of(code) is None
                    or not code.startswith(op)
                ):
                    raise ValueError("Kode zona lama tidak sesuai OP/format/panjang")
                if used_codes.get(code) != (op, kecamatan):
                    raise ValueError("Kode zona lama bertabrakan dengan kecamatan lain")
            else:
                key = (op, kecamatan)
                profile = (op, kecamatan, sig)
                options = zones[key]
                exact = [zone for zone in options if zone["sig"] == sig]

                if exact:
                    code = min(
                        (zone["code"] for zone in exact),
                        key=lambda candidate: (suffix_of(candidate), candidate),
                    )
                    status = (
                        "PAKAI ZONA MASTER (UJP SAMA)"
                        if code in master_codes
                        else generated_status[code]
                    )
                    if len({zone["code"] for zone in exact}) > 1:
                        note += "; " if note else ""
                        note += (
                            "Profil UJP ada pada beberapa zona master; "
                            "dipilih suffix terkecil"
                        )
                elif profile in generated:
                    code = generated[profile]
                    status = generated_status[code]
                else:
                    if options:
                        prefixes = {
                            zone["prefix"]
                            for zone in options
                            if reserved.get(zone["prefix"]) == key
                        }
                        next_suffix = (
                            max((suffix_max[prefix] for prefix in prefixes), default=0)
                            + 1
                        )
                        eligible = [
                            prefix
                            for prefix in prefixes
                            if len(prefix) + len(str(next_suffix)) <= MAX_CODE_LENGTH
                        ]

                        if eligible:
                            prefix = min(
                                eligible,
                                key=lambda candidate: (
                                    min(
                                        suffix_of(zone["code"])
                                        for zone in options
                                        if zone["prefix"] == candidate
                                    ),
                                    candidate,
                                ),
                            )
                        elif prefixes:
                            prefix = min(prefixes, key=lambda candidate: (len(candidate), candidate))
                        else:
                            prefix = choose_prefix(
                                op, kecamatan, reserved, len(str(next_suffix))
                            )
                            note += "; " if note else ""
                            note += "Prefix master bertabrakan dengan kecamatan lain"

                        status = (
                            "SUFFIX BARU (UJP BERBEDA)"
                            if key in master_districts
                            else "KODE ZONA BARU"
                        )
                    else:
                        next_suffix = 1
                        prefix = choose_prefix(
                            op, kecamatan, reserved, len(str(next_suffix))
                        )
                        status = "KODE ZONA BARU"

                    code = prefix + str(next_suffix)
                    if len(code) > MAX_CODE_LENGTH:
                        old_prefix = prefix
                        prefix = choose_prefix(
                            op, kecamatan, reserved, len(str(next_suffix))
                        )
                        code = prefix + str(next_suffix)
                        note += "; " if note else ""
                        note += (
                            f"Prefix {old_prefix} penuh; digunakan prefix singkat {prefix}"
                        )
                        status = "KODE BARU (BATAS 15 KARAKTER)"

                    if code in used_codes:
                        raise ValueError(f"Kode {code} sudah digunakan")

                    register(
                        code,
                        op,
                        kecamatan,
                        sig,
                        source=f"hasil baris {line}",
                    )
                    generated[profile] = code
                    generated_status[code] = status

            suffix = suffix_of(code)
            if suffix is None:
                raise ValueError("Kode zona tidak memiliki suffix angka")
            name = kecamatan + str(suffix)

        except ValueError as exc:
            code = old_code
            name = clean(row["NAMA ZONA"])
            status = "PERLU REVIEW"
            note = (note + "; " if note else "") + str(exc)

        output_codes.append(code)
        output_names.append(name)
        statuses.append(status)
        notes.append(note)
        counts[status] += 1

    data_master["KODE ZONA"] = output_codes
    data_master["NAMA ZONA"] = output_names
    data_master["STATUS KODE ZONA"] = statuses
    data_master["CATATAN VALIDASI"] = notes

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        data_master.to_excel(writer, sheet_name=DATA_MASTER_SHEET, index=False)
        master_toko.to_excel(writer, sheet_name=MASTER_TOKO_SHEET, index=False)
        audit_notes = sorted(set(warnings)) or ["Tidak ada catatan"]
        pd.DataFrame({"CATATAN REFERENSI": audit_notes}).to_excel(
            writer, sheet_name=AUDIT_SHEET, index=False
        )

    return {
        "rows": len(data_master),
        "status": dict(counts),
        "warnings": len(set(warnings)),
        "ujp_compared_by_op": op_columns,
        "output": str(output_path),
    }


def run_colab() -> None:
    """Jalankan alur upload, proses, dan download di Google Colab."""
    from google.colab import files

    uploaded = files.upload()
    if len(uploaded) != 1:
        raise ValueError("Upload tepat satu file .xlsx")

    name = next(iter(uploaded))
    if not name.lower().endswith(".xlsx"):
        raise ValueError("File harus .xlsx")

    output_name = f"hasil_zona_{Path(name).stem}.xlsx"
    print(generate(name, output_name))
    files.download(output_name)


def main() -> None:
    """Entry point untuk Google Colab atau command line lokal."""
    try:
        from google.colab import files as colab_files
    except ImportError:
        colab_files = None

    if colab_files is not None:
        run_colab()
        return

    if "-f" in sys.argv[1:] and "ipykernel" in sys.modules:
        raise RuntimeError(
            "Di Jupyter, jangan jalankan argparse dari sel. "
            "Panggil generate('input.xlsx', 'hasil.xlsx') secara langsung."
        )

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "input", help=f"Excel dengan sheet {DATA_MASTER_SHEET} dan {MASTER_TOKO_SHEET}"
    )
    parser.add_argument("-o", "--output", default="hasil_zona.xlsx")
    args = parser.parse_args()
    print(generate(args.input, args.output))


if __name__ == "__main__":
    main()


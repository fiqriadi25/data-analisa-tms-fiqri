from pathlib import Path

import pandas as pd
import pytest

from scripts.penambahan_toko_generate_kode_zona import generate, tariff


def write_workbook(path: Path, data_master: pd.DataFrame, master_toko: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        data_master.to_excel(writer, sheet_name="DATA MASTER", index=False)
        master_toko.to_excel(writer, sheet_name="MASTER TOKO", index=False)


def test_tariff_accepts_indonesian_number_format():
    assert tariff("68.000,00") == 68000
    assert tariff("68.000") == 68000
    assert tariff("68000,50") == 68000.5


def test_generate_reuses_master_and_creates_new_suffix(tmp_path: Path):
    input_file = tmp_path / "input.xlsx"
    output_file = tmp_path / "output.xlsx"

    master_toko = pd.DataFrame(
        {
            "Operating Point": ["TGR"],
            "Kecamatan": ["Serpong"],
            "Zona": ["TGRSERPONG1"],
            "UJP CDD": [100_000],
            "UJP CDE": [75_000],
        }
    )
    data_master = pd.DataFrame(
        {
            "OP": ["TGR", "TGR", "TGR"],
            "KECAMATAN": ["Serpong", "Serpong", "Pamulang"],
            "KODE ZONA": [None, None, None],
            "NAMA ZONA": [None, None, None],
            "UJP CDD": [100_000, 125_000, 90_000],
            "UJP CDE": [75_000, 75_000, 70_000],
        }
    )
    write_workbook(input_file, data_master, master_toko)

    summary = generate(input_file, output_file)
    result = pd.read_excel(output_file, sheet_name="DATA MASTER")

    assert summary["rows"] == 3
    assert list(result["STATUS KODE ZONA"]) == [
        "PAKAI ZONA MASTER (UJP SAMA)",
        "SUFFIX BARU (UJP BERBEDA)",
        "KODE ZONA BARU",
    ]
    assert result.loc[0, "KODE ZONA"] == "TGRSERPONG1"
    assert result.loc[1, "KODE ZONA"] == "TGRSERPONG2"
    assert result.loc[2, "KODE ZONA"].startswith("TGRPAMULANG")
    assert result["KODE ZONA"].str.len().max() <= 15
    assert "AUDIT REFERENSI" in pd.ExcelFile(output_file).sheet_names


def test_input_file_cannot_be_overwritten(tmp_path: Path):
    input_file = tmp_path / "input.xlsx"
    input_file.touch()
    with pytest.raises(ValueError, match="tidak boleh ditimpa"):
        generate(input_file, input_file)


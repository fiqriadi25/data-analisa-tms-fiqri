from pathlib import Path

import pandas as pd
import pytest

from scripts.update_ujp_master import UJP_ALL_COLS, process_dataframe, update_workbook


def make_row(
    toko: str,
    zona: str,
    kecamatan: str,
    old_cdd_c: int,
    update_cdd_c,
) -> dict:
    row = {
        "Kode Toko": toko,
        "Kode Zona": zona,
        "Operating Point": "TGR",
        "Kecamatan": kecamatan,
        "UPDATE UJP CDD-C": update_cdd_c,
        "UPDATE UJP CDE-LC": "Tidak update",
    }
    row.update({column: 0 for column in UJP_ALL_COLS})
    row["UJP CDD-C"] = old_cdd_c
    return row


def test_process_all_zone_classifications_without_changing_original_ujp():
    rows = [
        # Tipe A: tidak ada perubahan.
        make_row("A01", "TGRAAA1", "AAA", 100, "Tidak update"),
        # Tipe B1: semua baris berubah menjadi signature yang sama.
        make_row("B01", "TGRBBB1", "BBB", 100, 200),
        make_row("B02", "TGRBBB1", "BBB", 100, 200),
        # Tipe B2: hanya sebagian berubah, tetapi signature akhir sama.
        make_row("C01", "TGRCCC1", "CCC", 100, 200),
        make_row("C02", "TGRCCC1", "CCC", 200, "Tidak update"),
        # Tipe C: signature minoritas dipindahkan ke suffix baru.
        make_row("D01", "TGRDDD1", "DDD", 100, 200),
        make_row("D02", "TGRDDD1", "DDD", 100, "Tidak update"),
        make_row("D03", "TGRDDD1", "DDD", 100, "Tidak update"),
    ]
    source = pd.DataFrame(rows)
    original_ujp = source["UJP CDD-C"].copy()

    result, summary = process_dataframe(source)

    assert summary["zone_types"] == {"A": 1, "B1": 1, "B2": 1, "C": 1}
    assert list(result["UJP CDD-C"]) == list(original_ujp)
    assert result.loc[result["Kode Toko"] == "A01", "Keterangan"].iloc[0] == "TIDAK UPDATE"
    assert set(result.loc[result["Kode Zona"] == "TGRBBB1", "Keterangan"]) == {"UPDATE UJP"}
    assert set(result.loc[result["Kode Toko"].isin(["C01", "C02"]), "Keterangan"]) == {"NO UJP UPDATE"}

    split_row = result.loc[result["Kode Toko"] == "D01"].iloc[0]
    assert split_row["Kode Zona"] == "TGRDDD2"
    assert split_row["Keterangan"] == "NEW ZONE - UPDATE UJP - MOVE ZONE"
    assert len(split_row["Kode Zona"]) <= 15


def test_update_workbook_creates_new_file(tmp_path: Path):
    source = pd.DataFrame(
        [make_row("A01", "TGRAAA1", "AAA", 100, "Tidak update")]
    )
    input_file = tmp_path / "input.xlsx"
    output_file = tmp_path / "output.xlsx"
    source.to_excel(input_file, sheet_name="DATA MASTER", index=False)

    summary = update_workbook(input_file, output_file)

    assert summary["rows"] == 1
    assert output_file.exists()
    assert "Keterangan" in pd.read_excel(output_file).columns


def test_output_cannot_overwrite_input(tmp_path: Path):
    input_file = tmp_path / "input.xlsx"
    input_file.touch()
    with pytest.raises(ValueError, match="berbeda"):
        update_workbook(input_file, input_file)


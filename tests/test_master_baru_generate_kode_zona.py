import pandas as pd
import pytest

from scripts.master_baru_generate_kode_zona import (
    generate_kecamatan_code,
    process_dataframe,
    validate_result,
)


def test_generate_kecamatan_code():
    assert generate_kecamatan_code("Sungai Raya") == "SUNGAYA"
    assert generate_kecamatan_code("Serpong") == "SERPONG"


def test_same_pattern_gets_same_zone_and_different_ujp_gets_suffix():
    source = pd.DataFrame(
        {
            "OP": ["TGR", "TGR", "TGR"],
            "KECAMATAN": ["Sungai Raya", "Sungai Raya", "Sungai Raya"],
            "UJP CDD": [100_000, 100_000.0, 125_000],
            "UJP CDE": [0, None, 0],
        }
    )

    result, ujp_columns = process_dataframe(source)

    assert result.loc[0, "P2"] == result.loc[1, "P2"]
    assert result.loc[0, "P2"] != result.loc[2, "P2"]
    assert result["P2"].str.len().max() <= 15
    validate_result(result, ujp_columns)


def test_different_op_gets_different_zone():
    source = pd.DataFrame(
        {
            "OP": ["TGR", "BKS"],
            "KECAMATAN": ["Serpong", "Serpong"],
            "UJP CDD": [100_000, 100_000],
        }
    )

    result, _ = process_dataframe(source)
    assert result.loc[0, "P2"] != result.loc[1, "P2"]


def test_missing_required_column_raises_error():
    source = pd.DataFrame({"OP": ["TGR"], "UJP CDD": [100_000]})
    with pytest.raises(ValueError, match="KECAMATAN"):
        process_dataframe(source)

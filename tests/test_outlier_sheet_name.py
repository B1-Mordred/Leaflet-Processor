from __future__ import annotations

from pathlib import Path

from src.parser import parse_workbook


def test_outlier_sheet_name_is_handled():
    result = parse_workbook(
        Path("92028-XT Lot1125 3PLUS1 Neuroleptics 1-XT Calibrator Set (Excel).xlsx")
    )
    assert result.normalized_values

    sheet_names = {cell.sheet_name.lower() for cell in result.raw_cells}
    assert any("sorted by column" in name for name in sheet_names)
    assert any("sorted by rows" in name for name in sheet_names)


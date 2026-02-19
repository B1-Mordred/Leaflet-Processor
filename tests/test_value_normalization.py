from __future__ import annotations

from pathlib import Path

from src.parser import parse_workbook


def test_nd_and_separator_mapping():
    nd_result = parse_workbook(
        Path("0037 Lot3124 Serum Control LV2 - Vitamins A and E (Excel).xlsx")
    )
    numeric_result = parse_workbook(
        Path("0036 Lot3124 Serum Control LV1 - Vitamins A and E (Excel).xlsx")
    )

    nd_records = [r for r in nd_result.normalized_values if (r.raw_value or "").lower() == "n.d."]
    sep_records = [r for r in nd_result.normalized_values if r.raw_value == "-"]
    numeric_records = [r for r in numeric_result.normalized_values if r.value_status == "ok"]

    assert nd_records
    assert all(r.value_status == "nd" and r.numeric_value is None for r in nd_records)

    assert sep_records
    assert all(r.value_status == "separator" and r.numeric_value is None for r in sep_records)

    assert numeric_records
    assert any(r.numeric_value is not None for r in numeric_records)


def test_metric_role_pattern_for_control_file():
    result = parse_workbook(
        Path("0036 Lot3124 Serum Control LV1 - Vitamins A and E (Excel).xlsx")
    )
    roles = {r.metric_role for r in result.normalized_values}
    assert "target" in roles
    assert "range_low" in roles
    assert "range_sep" in roles
    assert "range_high" in roles

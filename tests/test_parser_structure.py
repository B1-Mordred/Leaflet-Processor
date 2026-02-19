from __future__ import annotations


def test_parse_all_workbooks(parsed_results, workbook_paths):
    assert len(parsed_results) == len(workbook_paths) == 15

    for result in parsed_results:
        assert result.workbook_meta.sheet_name_rows
        assert result.analytes
        assert result.normalized_values
        assert result.raw_cells


def test_raw_cells_include_version_sheet_when_present(parsed_results):
    with_version = [r for r in parsed_results if any(c.sheet_name == "Version" for c in r.raw_cells)]
    assert with_version, "Expected at least one workbook with a Version sheet in raw capture."


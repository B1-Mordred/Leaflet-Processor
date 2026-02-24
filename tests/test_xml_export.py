from __future__ import annotations

from pathlib import Path

from xml.etree import ElementTree as ET

from src.config import XmlConfig
from src.models import AnalyteDef, MeasurementRecord, WorkbookMeta, WorkbookParseResult
from src.xml_exporter import (
    _EMBEDDED_ADDON_XSD,
    build_addon_xml,
    build_consolidated_addon_xml,
    validate_addon_xml,
)


def _build_result() -> WorkbookParseResult:
    return WorkbookParseResult(
        source_file="example.xlsx",
        workbook_meta=WorkbookMeta(),
        analytes=[
            AnalyteDef(name="Retinol", group_name=None, column_index=1, units_seen=["mg/L", "μmol/L"]),
            AnalyteDef(name="Another", group_name=None, column_index=2, units_seen=[]),
        ],
    )


def test_xml_has_expected_root_and_ids():
    xml_text = build_addon_xml(_build_result(), XmlConfig())

    root = ET.fromstring(xml_text)
    assert root.tag == "AddOn"
    assert root.findtext("Id") == "1"
    assert root.findtext("MethodId") == "Beta Therapeutic Drug Monitoring"
    assert root.findtext("MethodVersion") == "1.2"

    assays = root.findall("./Assays/Assay")
    assert len(assays) == 2
    assert assays[0].findtext("Id") == "1"
    assert assays[0].findtext("AddOnRef") == "1"

    first_analyte = assays[0].find("./Analytes/Analyte")
    assert first_analyte is not None
    assert first_analyte.findtext("AssayRef") == "1"


def test_xml_contains_multi_unit_analyte():
    xml_text = build_addon_xml(_build_result(), XmlConfig())
    root = ET.fromstring(xml_text)

    target = None
    for assay_el in root.findall("./Assays/Assay"):
        analyte_el = assay_el.find("./Analytes/Analyte")
        if analyte_el is not None and analyte_el.findtext("Name") == "Retinol":
            target = analyte_el
            break

    assert target is not None
    unit_names = [u.findtext("Name") for u in target.findall("./AnalyteUnits/AnalyteUnit")]
    assert "mg/L" in unit_names
    assert "μmol/L" in unit_names


def test_generated_xml_validates_against_addon_xsd():
    xml_text = build_addon_xml(_build_result(), XmlConfig())
    validate_addon_xml(xml_text)


def test_generated_xml_validation_does_not_require_external_xsd(tmp_path, monkeypatch):
    xml_text = build_addon_xml(_build_result(), XmlConfig())
    monkeypatch.chdir(tmp_path)

    # Validation should still work even if no template/AddOn.xsd exists near cwd.
    validate_addon_xml(xml_text)


def test_embedded_schema_matches_template_xsd():
    template_xsd = Path("template/AddOn.xsd").read_text(encoding="utf-8-sig")
    assert _EMBEDDED_ADDON_XSD == template_xsd


def test_consolidated_xml_uses_group_name_sample_code_and_unit_mappings():
    rec1 = MeasurementRecord(
        source_file="one.xlsx",
        sample_label="S1",
        sample_code="11",
        unit="mg/L",
        analyte_name="Retinol",
        group_name="Vitamin Assay",
        metric_role="target",
        raw_value="1.2",
        numeric_value=1.2,
        value_status="ok",
        sheet_row=10,
        sheet_col=5,
    )
    rec2 = MeasurementRecord(
        source_file="two.xlsx",
        sample_label="S2",
        sample_code="ABC",
        unit="μmol/L",
        analyte_name="Retinol",
        group_name="Vitamin Assay",
        metric_role="target",
        raw_value="3.4",
        numeric_value=3.4,
        value_status="ok",
        sheet_row=11,
        sheet_col=5,
    )
    rec3 = MeasurementRecord(
        source_file="two.xlsx",
        sample_label="S2",
        sample_code=None,
        unit=None,
        analyte_name="B12",
        group_name=None,
        metric_role="target",
        raw_value="0.9",
        numeric_value=0.9,
        value_status="ok",
        sheet_row=12,
        sheet_col=6,
    )
    results = [
        WorkbookParseResult(source_file="one.xlsx", workbook_meta=WorkbookMeta(), normalized_values=[rec1]),
        WorkbookParseResult(
            source_file="two.xlsx", workbook_meta=WorkbookMeta(), normalized_values=[rec2, rec3]
        ),
    ]

    xml_text = build_consolidated_addon_xml(results, XmlConfig())
    root = ET.fromstring(xml_text)

    assays = root.findall("./Assays/Assay")
    assert len(assays) == 2
    assert [assay.findtext("Name") for assay in assays] == ["", "Vitamin Assay"]
    assert all(assay.findtext("Id") == "0" for assay in assays)

    vitamin_assay = assays[1]
    retinol = vitamin_assay.find("./Analytes/Analyte")
    assert retinol is not None
    assert retinol.findtext("AssayRef") == "11"
    retinol_units = [u.findtext("Name") for u in retinol.findall("./AnalyteUnits/AnalyteUnit")]
    assert retinol_units == ["mg/L", "μmol/L"]
    assert all(u.findtext("Id") == "0" for u in retinol.findall("./AnalyteUnits/AnalyteUnit"))

    blank_assay_analyte = assays[0].find("./Analytes/Analyte")
    assert blank_assay_analyte is not None
    assert blank_assay_analyte.findtext("AssayRef") == "0"

    validate_addon_xml(xml_text)


def test_consolidated_xml_deduplicates_analytes_within_assay():
    rec1 = MeasurementRecord(
        source_file="one.xlsx",
        sample_label="S1",
        sample_code="",
        unit="mg/L",
        analyte_name=" Retinol ",
        group_name="Vitamin Assay",
        metric_role="target",
        raw_value="1.2",
        numeric_value=1.2,
        value_status="ok",
        sheet_row=10,
        sheet_col=5,
    )
    rec2 = MeasurementRecord(
        source_file="two.xlsx",
        sample_label="S2",
        sample_code="27",
        unit="μmol/L",
        analyte_name="retinol",
        group_name="Vitamin Assay",
        metric_role="target",
        raw_value="3.4",
        numeric_value=3.4,
        value_status="ok",
        sheet_row=11,
        sheet_col=5,
    )

    result = WorkbookParseResult(
        source_file="mix.xlsx",
        workbook_meta=WorkbookMeta(),
        normalized_values=[rec1, rec2],
    )

    xml_text = build_consolidated_addon_xml([result], XmlConfig())
    root = ET.fromstring(xml_text)

    analytes = root.findall("./Assays/Assay/Analytes/Analyte")
    assert len(analytes) == 1
    assert analytes[0].findtext("Name") == "Retinol"
    assert analytes[0].findtext("AssayRef") == "27"

    unit_names = [u.findtext("Name") for u in analytes[0].findall("./AnalyteUnits/AnalyteUnit")]
    assert unit_names == ["mg/L", "μmol/L"]


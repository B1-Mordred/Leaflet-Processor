from __future__ import annotations

from xml.etree import ElementTree as ET

from src.config import XmlConfig
from src.models import AnalyteDef, WorkbookMeta, WorkbookParseResult
from src.xml_exporter import build_addon_xml, validate_addon_xml


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

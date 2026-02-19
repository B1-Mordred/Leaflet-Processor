from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET

from src.config import XmlConfig
from src.parser import parse_workbook
from src.xml_exporter import build_addon_xml


def test_xml_has_expected_root_and_ids():
    result = parse_workbook(
        Path("0211-XT Lot1125 MassCheck Neuroleptics 1-XT LV1 (Excel).xlsx")
    )
    xml_text = build_addon_xml(result, XmlConfig())

    root = ET.fromstring(xml_text)
    assert root.tag == "AddOn"
    assert root.findtext("Id") == "1"
    assert root.findtext("MethodId") == "Beta Therapeutic Drug Monitoring"
    assert root.findtext("MethodVersion") == "1.2"

    analytes = root.findall("./Analytes/Analyte")
    assert len(analytes) == len(result.analytes)
    assert analytes[0].findtext("Id") == "1"
    assert analytes[0].findtext("AddOnRef") == "1"


def test_xml_contains_multi_unit_analyte():
    result = parse_workbook(
        Path("0036 Lot3124 Serum Control LV1 - Vitamins A and E (Excel).xlsx")
    )
    xml_text = build_addon_xml(result, XmlConfig())
    root = ET.fromstring(xml_text)

    target = None
    for analyte_el in root.findall("./Analytes/Analyte"):
        if analyte_el.findtext("Name") == "Retinol":
            target = analyte_el
            break

    assert target is not None
    unit_names = [u.findtext("Name") for u in target.findall("./AnalyteUnits/AnalyteUnit")]
    assert "mg/L" in unit_names
    assert "μmol/L" in unit_names


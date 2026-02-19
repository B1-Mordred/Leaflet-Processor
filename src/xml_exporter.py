from __future__ import annotations

import io
from pathlib import Path
from xml.etree import ElementTree as ET

from .config import XmlConfig
from .models import WorkbookParseResult


def build_addon_xml(result: WorkbookParseResult, cfg: XmlConfig) -> str:
    root = ET.Element(
        "AddOn",
        attrib={
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
        },
    )

    ET.SubElement(root, "Id").text = "1"
    ET.SubElement(root, "MethodId").text = cfg.method_id
    ET.SubElement(root, "MethodVersion").text = cfg.method_version

    sample_tube_types = ET.SubElement(root, "SampleTubeTypes")
    _append_string_items(sample_tube_types, cfg.sample_tube_types)

    measurement_sample_lists = ET.SubElement(root, "MeasurementSampleLists")
    _append_string_items(measurement_sample_lists, cfg.measurement_sample_lists)

    ET.SubElement(root, "RunResultsExportPath").text = cfg.run_results_export_path

    analytes_el = ET.SubElement(root, "Analytes")
    for analyte_id, analyte in enumerate(result.analytes, start=1):
        analyte_el = ET.SubElement(analytes_el, "Analyte")
        ET.SubElement(analyte_el, "Id").text = str(analyte_id)
        ET.SubElement(analyte_el, "Name").text = analyte.name
        ET.SubElement(analyte_el, "AddOnRef").text = "1"

        analyte_units_el = ET.SubElement(analyte_el, "AnalyteUnits")
        units = analyte.units_seen if analyte.units_seen else []
        for unit_id, unit_name in enumerate(units, start=1):
            unit_el = ET.SubElement(analyte_units_el, "AnalyteUnit")
            ET.SubElement(unit_el, "Id").text = str(unit_id)
            ET.SubElement(unit_el, "Name").text = unit_name
            ET.SubElement(unit_el, "AnalyteRef").text = str(analyte_id)

    ET.indent(root, space="  ")

    buffer = io.BytesIO()
    tree = ET.ElementTree(root)
    tree.write(buffer, encoding="utf-8", xml_declaration=True)
    return buffer.getvalue().decode("utf-8")


def write_addon_xml(result: WorkbookParseResult, cfg: XmlConfig, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    out_name = Path(result.source_file).stem + ".xml"
    out_path = out_dir / out_name
    xml_content = build_addon_xml(result=result, cfg=cfg)
    out_path.write_text(xml_content, encoding="utf-8")
    return out_path


def _append_string_items(parent: ET.Element, values: list[str]) -> None:
    for value in values:
        item = ET.SubElement(parent, "string")
        item.text = value


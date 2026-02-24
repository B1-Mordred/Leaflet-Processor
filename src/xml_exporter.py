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

    if cfg.run_results_export_path:
        ET.SubElement(root, "RunResultsExportPath").text = cfg.run_results_export_path

    assays_el = ET.SubElement(root, "Assays")
    for assay_id, analyte in enumerate(result.analytes, start=1):
        assay_el = ET.SubElement(assays_el, "Assay")
        ET.SubElement(assay_el, "Id").text = str(assay_id)
        ET.SubElement(assay_el, "Name").text = analyte.name
        ET.SubElement(assay_el, "AddOnRef").text = "1"

        analytes_el = ET.SubElement(assay_el, "Analytes")
        analyte_el = ET.SubElement(analytes_el, "Analyte")
        ET.SubElement(analyte_el, "Id").text = str(assay_id)
        ET.SubElement(analyte_el, "Name").text = analyte.name
        ET.SubElement(analyte_el, "AssayRef").text = str(assay_id)

        if analyte.units_seen:
            analyte_units_el = ET.SubElement(analyte_el, "AnalyteUnits")
            for unit_id, unit_name in enumerate(analyte.units_seen, start=1):
                unit_el = ET.SubElement(analyte_units_el, "AnalyteUnit")
                ET.SubElement(unit_el, "Id").text = str(unit_id)
                ET.SubElement(unit_el, "Name").text = unit_name
                ET.SubElement(unit_el, "AnalyteRef").text = str(assay_id)

    ET.indent(root, space="  ")

    buffer = io.BytesIO()
    ET.ElementTree(root).write(buffer, encoding="utf-8", xml_declaration=True)
    return buffer.getvalue().decode("utf-8")


def write_addon_xml(result: WorkbookParseResult, cfg: XmlConfig, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    out_name = Path(result.source_file).stem + ".xml"
    out_path = out_dir / out_name
    xml_content = build_addon_xml(result=result, cfg=cfg)
    validate_addon_xml(xml_content)
    out_path.write_text(xml_content, encoding="utf-8")
    return out_path


def validate_addon_xml(xml_text: str, xsd_path: Path | None = None) -> None:
    _ = ET.parse(xsd_path or _default_xsd_path())
    root = ET.fromstring(xml_text)

    if root.tag != "AddOn":
        raise ValueError("Generated AddOn XML failed XSD validation: root element must be AddOn")

    _require_int(root, "Id")

    assays_el = root.find("Assays")
    if assays_el is None:
        raise ValueError("Generated AddOn XML failed XSD validation: missing Assays element")

    for assay_el in assays_el.findall("Assay"):
        _require_int(assay_el, "Id")
        _require_int(assay_el, "AddOnRef")
        analytes_el = assay_el.find("Analytes")
        if analytes_el is None:
            continue

        for analyte_el in analytes_el.findall("Analyte"):
            _require_int(analyte_el, "Id")
            _require_int(analyte_el, "AssayRef")
            analyte_units_el = analyte_el.find("AnalyteUnits")
            if analyte_units_el is None:
                continue

            for analyte_unit_el in analyte_units_el.findall("AnalyteUnit"):
                _require_int(analyte_unit_el, "Id")
                _require_int(analyte_unit_el, "AnalyteRef")


def _require_int(parent: ET.Element, child_name: str) -> None:
    text = parent.findtext(child_name)
    if text is None:
        raise ValueError(
            f"Generated AddOn XML failed XSD validation: missing {child_name} under {parent.tag}"
        )
    try:
        int(text)
    except ValueError as exc:
        raise ValueError(
            f"Generated AddOn XML failed XSD validation: {child_name} must be an integer under {parent.tag}"
        ) from exc


def _default_xsd_path() -> Path:
    return Path(__file__).resolve().parents[1] / "template" / "AddOn.xsd"

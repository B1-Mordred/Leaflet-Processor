from __future__ import annotations

import io
from pathlib import Path
from xml.etree import ElementTree as ET

from .config import XmlConfig
from .models import WorkbookParseResult


_EMBEDDED_ADDON_XSD = '<?xml version="1.0" encoding="utf-8"?>\n<xs:schema elementFormDefault="qualified" xmlns:xs="http://www.w3.org/2001/XMLSchema">\n\t<xs:element name="AddOn" nillable="true" type="AddOn" />\n\t<xs:complexType name="AddOn">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Id" type="xs:int" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="MethodId" type="xs:string" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="MethodVersion" type="xs:string" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="SampleTubeTypes" type="ArrayOfSampleTubeType" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="MeasurementSampleLists" type="ArrayOfMeasurementSampleList" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="RunResultsExportPath" type="xs:string" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Assays" type="ArrayOfAssay" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="ArrayOfSampleTubeType">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="SampleTubeType" nillable="true" type="SampleTubeType" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="SampleTubeType">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Id" type="xs:int" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="DisplayName" type="xs:string" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="BarcodeMask" type="xs:string" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="FullFilename" type="xs:string" />\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="SampleCarrierType" type="SampleCarrierType" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="BarcodeRegex" type="xs:string" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="GeneralConfigs" type="ArrayOfGeneralConfiguration" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="AddOns" type="ArrayOfAddOn" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:simpleType name="SampleCarrierType">\n\t\t<xs:restriction base="xs:string">\n\t\t\t<xs:enumeration value="Undefined" />\n\t\t\t<xs:enumeration value="Positions24" />\n\t\t\t<xs:enumeration value="Positions32" />\n\t\t\t<xs:enumeration value="ErrorCarrierPositions24" />\n\t\t\t<xs:enumeration value="ErrorCarrierPositions32" />\n\t\t</xs:restriction>\n\t</xs:simpleType>\n\t<xs:complexType name="ArrayOfGeneralConfiguration">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="GeneralConfiguration" nillable="true" type="GeneralConfiguration" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="GeneralConfiguration">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Id" type="xs:int" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Version" type="xs:string" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="SampleTubeTypes" type="ArrayOfSampleTubeType" />\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="IsRequestListUsed" type="xs:boolean" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="RequestListFilePath" type="xs:string" />\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="IsLimsFileUsed" type="xs:boolean" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="LimsFilePath" type="xs:string" />\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="ValidateRequestList" type="xs:boolean" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="ArrayOfAddOn">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="AddOn" nillable="true" type="AddOn" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\n\t<!-- New: ArrayOfAssay / Assay -->\n\t<xs:complexType name="ArrayOfAssay">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="Assay" nillable="true" type="Assay" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\n\t<xs:complexType name="Assay">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="AddOnRef" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Analytes" type="ArrayOfAnalyte" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\n\t<xs:complexType name="ArrayOfMeasurementSampleList">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="MeasurementSampleList" nillable="true" type="MeasurementSampleList" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="MeasurementSampleList">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="AddOnRef" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="ExportPath" type="xs:string" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Header" type="xs:string" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Footer" type="xs:string" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="AdditionalInjections" type="ArrayOfAdditionalInjection" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="WarmUps" type="ArrayOfWarmUp" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="ParameterMappings" type="ArrayOfMeasurementSampleListItem" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="DelimiterType" type="DelimiterType" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="FileType" type="FileType" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="IsSelected" type="xs:boolean" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Assay" type="xs:string" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\t<xs:complexType name="NamedItemOfInt32">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Id" type="xs:int" />\n\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Name" type="xs:string" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="AnalyteUnit">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="AnalyteRef" type="xs:int" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\n\t<!-- Modified Analyte: now references AssayRef and may include AssayInformationType -->\n\t<xs:complexType name="Analyte">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="AssayRef" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="AnalyteUnits" type="ArrayOfAnalyteUnit" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="AssayInformationType" type="xs:string" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\n\t<xs:complexType name="ArrayOfAnalyteUnit">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="AnalyteUnit" nillable="true" type="AnalyteUnit" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="ArrayOfAnalyte">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="Analyte" nillable="true" type="Analyte" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\n\t<xs:complexType name="MeasurementSampleListItemComponent">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="MeasurementSampleListItemRef" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="MeasurementSampleListComponentType" type="MeasurementSampleListComponentType" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Value" type="xs:string" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Parameter" type="xs:string" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\n\t<xs:simpleType name="MeasurementSampleListComponentType">\n\t\t<xs:restriction base="xs:string">\n\t\t\t<xs:enumeration value="Undefined" />\n\t\t\t<xs:enumeration value="StaticText" />\n\t\t\t<xs:enumeration value="SampleId" />\n\t\t\t<xs:enumeration value="SampleType" />\n\t\t\t<xs:enumeration value="FinalPlateBarcode" />\n\t\t\t<xs:enumeration value="RunTimeStamp" />\n\t\t\t<xs:enumeration value="UserName" />\n\t\t\t<xs:enumeration value="AnalyteConcentration" />\n\t\t\t<xs:enumeration value="SamplePosition" />\n\t\t\t<xs:enumeration value="Level" />\n\t\t\t<xs:enumeration value="State" />\n\t\t\t<!-- Added entries to match C# enum -->\n\t\t\t<xs:enumeration value="SourcePosition" />\n\t\t\t<xs:enumeration value="DilutionFactor" />\n\t\t</xs:restriction>\n\t</xs:simpleType>\n\n\t<xs:complexType name="MeasurementSampleListItem">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="Components" type="ArrayOfMeasurementSampleListItemComponent" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="MeasurementSampleListRef" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Position" type="xs:int" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\t<xs:complexType name="ArrayOfMeasurementSampleListItemComponent">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="MeasurementSampleListItemComponent" nillable="true" type="MeasurementSampleListItemComponent" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="WarmUp">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="RequiredItemType" type="REQUIRED_ITEM_TYPE" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Level" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="MeasurementSampleListRef" type="xs:int" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\t<xs:simpleType name="REQUIRED_ITEM_TYPE">\n\t\t<xs:restriction base="xs:string">\n\t\t\t<xs:enumeration value="Undefined" />\n\t\t\t<xs:enumeration value="Calibrator" />\n\t\t\t<xs:enumeration value="Control" />\n\t\t\t<xs:enumeration value="Reagent" />\n\t\t\t<xs:enumeration value="Plate" />\n\t\t\t<xs:enumeration value="TipRack" />\n\t\t</xs:restriction>\n\t</xs:simpleType>\n\t<xs:complexType name="AdditionalInjection">\n\t\t<xs:complexContent mixed="false">\n\t\t\t<xs:extension base="NamedItemOfInt32">\n\t\t\t\t<xs:sequence>\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Frequency" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Offset" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Prepend" type="xs:boolean" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Append" type="xs:boolean" />\n\t\t\t\t\t<xs:element minOccurs="0" maxOccurs="1" name="DisplayName" type="xs:string" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="RequiredItemType" type="REQUIRED_ITEM_TYPE" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="Level" type="xs:int" />\n\t\t\t\t\t<xs:element minOccurs="1" maxOccurs="1" name="MeasurementSampleListRef" type="xs:int" />\n\t\t\t\t</xs:sequence>\n\t\t\t</xs:extension>\n\t\t</xs:complexContent>\n\t</xs:complexType>\n\t<xs:complexType name="ArrayOfAdditionalInjection">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="AdditionalInjection" nillable="true" type="AdditionalInjection" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="ArrayOfWarmUp">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="WarmUp" nillable="true" type="WarmUp" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:complexType name="ArrayOfMeasurementSampleListItem">\n\t\t<xs:sequence>\n\t\t\t<xs:element minOccurs="0" maxOccurs="unbounded" name="MeasurementSampleListItem" nillable="true" type="MeasurementSampleListItem" />\n\t\t</xs:sequence>\n\t</xs:complexType>\n\t<xs:simpleType name="DelimiterType">\n\t\t<xs:restriction base="xs:string">\n\t\t\t<xs:enumeration value="Undefined" />\n\t\t\t<xs:enumeration value="Semicolon" />\n\t\t\t<xs:enumeration value="Comma" />\n\t\t\t<xs:enumeration value="Blank" />\n\t\t\t<xs:enumeration value="Tab" />\n\t\t</xs:restriction>\n\t</xs:simpleType>\n\t<xs:simpleType name="FileType">\n\t\t<xs:restriction base="xs:string">\n\t\t\t<xs:enumeration value="Undefined" />\n\t\t\t<xs:enumeration value="Csv" />\n\t\t\t<xs:enumeration value="Txt" />\n\t\t\t<xs:enumeration value="Xlsx" />\n\t\t\t<xs:enumeration value="Json" />\n\t\t\t<xs:enumeration value="Xml" />\n\t\t</xs:restriction>\n\t</xs:simpleType>\n\t\n</xs:schema>'



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
    if xsd_path is not None:
        _ = ET.parse(xsd_path)
    else:
        _ = ET.fromstring(_EMBEDDED_ADDON_XSD)

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


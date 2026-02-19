from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Literal

MetricRole = Literal["target", "range_low", "range_sep", "range_high", "other"]
ValueStatus = Literal["ok", "nd", "separator", "blank", "text"]


@dataclass(slots=True)
class WorkbookMeta:
    title_lines: list[str] = field(default_factory=list)
    order_no: str | None = None
    lot_no: str | None = None
    exp_date: date | None = None
    consisting_of: str | None = None
    date_of_creation: date | None = None
    sheet_name_rows: str = ""


@dataclass(slots=True)
class AnalyteDef:
    name: str
    group_name: str | None
    column_index: int
    units_seen: list[str] = field(default_factory=list)


@dataclass(slots=True)
class MeasurementRecord:
    source_file: str
    sample_label: str | None
    sample_code: str | None
    unit: str | None
    analyte_name: str
    group_name: str | None
    metric_role: MetricRole
    raw_value: str | None
    numeric_value: float | None
    value_status: ValueStatus
    sheet_row: int
    sheet_col: int


@dataclass(slots=True)
class RawCellRecord:
    sheet_name: str
    row_idx: int
    col_idx: int
    raw_value: str | None
    python_type: str | None


@dataclass(slots=True)
class WorkbookParseResult:
    source_file: str
    workbook_meta: WorkbookMeta
    analytes: list[AnalyteDef] = field(default_factory=list)
    normalized_values: list[MeasurementRecord] = field(default_factory=list)
    raw_cells: list[RawCellRecord] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


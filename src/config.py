from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path


@dataclass(slots=True)
class XmlConfig:
    method_id: str = "Beta Therapeutic Drug Monitoring"
    method_version: str = "1.2"
    sample_tube_types: list[str] = field(default_factory=list)
    measurement_sample_lists: list[str] = field(default_factory=list)
    run_results_export_path: str = ""


_DEFAULT_PATH = Path("config/gui_defaults.json")


def load_gui_defaults() -> XmlConfig:
    if not _DEFAULT_PATH.exists():
        return XmlConfig()
    try:
        data = json.loads(_DEFAULT_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return XmlConfig()

    return XmlConfig(
        method_id=str(data.get("method_id", XmlConfig.method_id)),
        method_version=str(data.get("method_version", XmlConfig.method_version)),
        sample_tube_types=_normalize_list(data.get("sample_tube_types", [])),
        measurement_sample_lists=_normalize_list(data.get("measurement_sample_lists", [])),
        run_results_export_path=str(data.get("run_results_export_path", "")),
    )


def save_gui_defaults(cfg: XmlConfig) -> None:
    _DEFAULT_PATH.parent.mkdir(parents=True, exist_ok=True)
    _DEFAULT_PATH.write_text(
        json.dumps(asdict(cfg), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _normalize_list(value: object) -> list[str]:
    if isinstance(value, list):
        out: list[str] = []
        for item in value:
            text = str(item).strip()
            if text:
                out.append(text)
        return out
    return []


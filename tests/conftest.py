from __future__ import annotations

from pathlib import Path

import pytest

from src.parser import parse_folder


@pytest.fixture(scope="session")
def workbook_paths() -> list[Path]:
    return sorted(
        p for p in Path(".").glob("*.xlsx") if p.is_file() and not p.name.startswith("~$")
    )


@pytest.fixture(scope="session")
def parsed_results():
    return parse_folder(Path("."))


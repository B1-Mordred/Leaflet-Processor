# Changelog

## Unreleased

### Changed
- XML export now produces a single consolidated XML (`consolidated.xml`) from all imported Excel workbooks instead of writing one XML per workbook.
- Consolidated XML mappings now use: Excel `group name` -> `<Assay><Name>`, Excel `sample code` -> `<Analyte><AssayRef>` (non-integer or empty values default to `0`), and Excel unit -> `<AnalyteUnit><Name>`.
- Unassigned integer identifiers in consolidated export are explicitly written as `0`.

### Fixed
- Consolidated XML export now deduplicates analytes within the same assay (case/whitespace-insensitive analyte names), merging units into a single analyte entry.
- When multiple rows for the same analyte exist, consolidated export now prefers the first non-zero parsed sample code for `<AssayRef>`.
- XML export no longer depends on an external `template/AddOn.xsd` file at runtime; validation now uses an embedded schema fallback so compiled executables can export XML even when the template folder is unavailable.
- Windows build configuration now bundles `template/AddOn.xsd` into the PyInstaller output for compatibility with explicit XSD-path validation.
- Embedded XML fallback schema now exactly mirrors `template/AddOn.xsd` to keep fallback behavior aligned with the canonical template definition.

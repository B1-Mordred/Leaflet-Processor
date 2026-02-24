# Changelog

## Unreleased

### Fixed
- XML export no longer depends on an external `template/AddOn.xsd` file at runtime; validation now uses an embedded schema fallback so compiled executables can export XML even when the template folder is unavailable.
- Windows build configuration now bundles `template/AddOn.xsd` into the PyInstaller output for compatibility with explicit XSD-path validation.
- Embedded XML fallback schema now exactly mirrors `template/AddOn.xsd` to keep fallback behavior aligned with the canonical template definition.

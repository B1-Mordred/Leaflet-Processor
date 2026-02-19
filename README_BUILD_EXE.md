# Build Windows EXE

## 1) Install runtime dependencies

```powershell
python -m pip install -r requirements.txt
```

## 2) Install build dependency

```powershell
python -m pip install -r requirements-build.txt
```

## 3) Build app as a folder distribution (recommended)

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Output executable:

`dist\Chromsystems Universal Leaflet Parser 0.1.0\Chromsystems Universal Leaflet Parser 0.1.0.exe`

The build script uses `assets\favicon.ico` as the `.exe` icon and includes it in the bundle so the Tk title bar icon is also set at runtime.
It also excludes optional heavy scientific/test packages (`numpy`, `pandas`, `scipy`, `matplotlib`, `PIL`, `pytest`, `setuptools`, `yaml`, `reportlab`) to reduce PyInstaller warning noise and bundle size.

## 4) Build single-file EXE (optional)

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1 -OneFile
```

Output executable:

`dist\Chromsystems Universal Leaflet Parser 0.1.0.exe`

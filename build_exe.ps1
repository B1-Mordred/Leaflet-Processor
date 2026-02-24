param(
    [switch]$OneFile
)

$ErrorActionPreference = "Stop"

$AppName = "Chromsystems Universal Leaflet Parser 0.1.0"
$EntryScript = "run_app.py"
$IconPath = "assets\favicon.ico"

$ExcludedModules = @(
    "numpy",
    "pandas",
    "scipy",
    "matplotlib",
    "PIL",
    "pytest",
    "setuptools",
    "yaml",
    "reportlab"
)

if (-not (Test-Path $EntryScript)) {
    throw "Entry script not found: $EntryScript"
}

$baseArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--name", $AppName,
    $EntryScript
)

if (Test-Path $IconPath) {
    $baseArgs += @("--icon", $IconPath)
    $baseArgs += @("--add-data", "assets\favicon.ico;assets")
    $baseArgs += @("--add-data", "template\AddOn.xsd;template")
} else {
    Write-Warning "Icon not found at '$IconPath'. Building without custom icon."
}

if ($OneFile) {
    $baseArgs += "--onefile"
}

foreach ($moduleName in $ExcludedModules) {
    $baseArgs += @("--exclude-module", $moduleName)
}

Write-Host "Building executable for '$AppName'..."
python @baseArgs

Write-Host "Build complete."
if ($OneFile) {
    Write-Host "Output: dist\$AppName.exe"
} else {
    Write-Host "Output: dist\$AppName\$AppName.exe"
}

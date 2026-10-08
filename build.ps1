$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$venv = Join-Path $root ".venv-build"
$python = Join-Path $venv "Scripts\python.exe"
$dist = Join-Path $root "dist"
$work = Join-Path $root "build\pyinstaller"

if (-not (Test-Path $python)) {
    python -m venv $venv
}

& $python -m pip install --disable-pip-version-check -r (Join-Path $root "requirements-build.txt")
if ($LASTEXITCODE -ne 0) {
    throw "Could not install the pinned build dependencies."
}

& $python -m PyInstaller --noconfirm --clean --distpath $dist --workpath $work (Join-Path $root "md-xlsx-convertor.spec")
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed."
}

& $python (Join-Path $root "scripts\collect_licenses.py") $dist
if ($LASTEXITCODE -ne 0) {
    throw "Could not collect third-party license notices."
}

Write-Host "Created $(Join-Path $dist 'md-xlsx-convertor.exe')"
Write-Host "Distribute it with THIRD_PARTY_NOTICES.txt and ThirdPartyLicenses."

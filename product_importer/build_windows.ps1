$ErrorActionPreference = 'Stop'
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $appDir
$rootPython = Join-Path (Split-Path -Parent $appDir) '.venv\Scripts\python.exe'
if (Test-Path $rootPython) {
  $python = $rootPython
} else {
  if (-not (Test-Path '.venv\Scripts\python.exe')) { py -3.11 -m venv .venv }
  $python = Join-Path $appDir '.venv\Scripts\python.exe'
}
& $python -m pip install -r requirements.txt
& $python -m PyInstaller --noconfirm --clean --onefile --windowed --paths $appDir --name 'Xudaai产品资料整理工具-AI对话版' --distpath (Join-Path $appDir 'dist') --workpath (Join-Path $appDir 'build') --specpath $appDir --collect-all docx --collect-all openpyxl --collect-all pptx --collect-all reportlab --collect-all qrcode desktop.py
Write-Host "Build complete: $appDir\dist\Xudaai产品资料整理工具-AI对话版.exe"

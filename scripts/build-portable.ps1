param(
    [string]$OutputRoot = "C:\Projects",
    [string]$NamePrefix = "watermark-portable-deploy",
    [switch]$IncludeArchivedWatermark = $true
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "Creating packaging venv..."
python -m venv .packenv

Write-Host "Installing packaging dependencies..."
.\.packenv\Scripts\python.exe -m pip install --upgrade pip
.\.packenv\Scripts\python.exe -m pip install . pyinstaller

Write-Host "Writing launcher..."
@'
from watermark_app.main import main

if __name__ == "__main__":
    main()
'@ | Set-Content .\run_watermark.py -Encoding UTF8

Write-Host "Building standalone executable..."
.\.packenv\Scripts\pyinstaller.exe `
  --noconfirm --clean --onedir `
  --name watermark-app `
  --collect-all msal `
  --collect-all cryptography `
  --collect-all openpyxl `
  --collect-all docx `
  --collect-all PIL `
  .\run_watermark.py

$ts = Get-Date -Format "yyyyMMdd-HHmmss"
$deploy = Join-Path $OutputRoot "$NamePrefix-$ts"
$zip = "$deploy.zip"

Write-Host "Staging deploy folder: $deploy"
New-Item -ItemType Directory -Path $deploy | Out-Null

Copy-Item .\dist\watermark-app\* $deploy -Recurse -Force
Copy-Item .\.env.example "$deploy\.env.example" -Force
Copy-Item .\classified_watermark.png "$deploy\classified_watermark.png" -Force
if ($IncludeArchivedWatermark -and (Test-Path .\archived_watermark.png)) {
    Copy-Item .\archived_watermark.png "$deploy\archived_watermark.png" -Force
}

@'
Portable Deploy Instructions
============================

1) Copy this folder to VM, e.g. C:\Apps\watermark-app
2) Copy .env.example to .env and fill values.
3) Ensure SP_LIBRARY_WATERMARKS points to PNG paths on VM.
4) Dry run:
   .\watermark-app.exe --dry-run --log-level INFO
5) Real run:
   .\watermark-app.exe --log-level INFO

Task Scheduler:
Program/script: C:\Apps\watermark-app\watermark-app.exe
Arguments: --log-level INFO
Start in: C:\Apps\watermark-app
'@ | Set-Content "$deploy\RUN_ME_FIRST.txt" -Encoding UTF8

Write-Host "Creating zip: $zip"
Compress-Archive -Path "$deploy\*" -DestinationPath $zip -CompressionLevel Optimal

Write-Host "Done."
Write-Host "Deploy folder: $deploy"
Write-Host "Deploy zip: $zip"

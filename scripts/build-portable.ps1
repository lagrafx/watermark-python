param(
    [string]$OutputRoot = "C:\Projects",
    [string]$NamePrefix = "watermark-portable-deploy",
    [switch]$IncludeArchivedWatermark = $true,
    [string]$PortableVersion = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$sourceCommit = (git rev-parse --short HEAD).Trim()
$sourceBranch = (git rev-parse --abbrev-ref HEAD).Trim()
$sourceTag = ""
try {
    $sourceTag = (git describe --tags --exact-match 2>$null).Trim()
} catch {
    $sourceTag = ""
}

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
$portableTag = if ($PortableVersion) { $PortableVersion } else { "portable-v$ts+$sourceCommit" }

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

$manifest = @{
    portable_tag = $portableTag
    build_timestamp_utc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    source_commit = $sourceCommit
    source_branch = $sourceBranch
    source_tag = $sourceTag
    source_repo = "https://github.com/lagrafx/watermark-python"
    portable_repo = "https://github.com/lagrafx/watermark-python-portable"
    zip_name = [System.IO.Path]::GetFileName($zip)
}
$manifest | ConvertTo-Json -Depth 5 | Set-Content "$deploy\PORTABLE_RELEASE_MANIFEST.json" -Encoding UTF8

@"
# Portable Release Notes

- Portable tag: $portableTag
- Source commit: $sourceCommit
- Source branch: $sourceBranch
- Source tag: $(if ($sourceTag) { $sourceTag } else { "(none)" })
- Build timestamp (UTC): $((Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"))

## Publish Checklist
1. Create/update release in `watermark-python-portable` with tag: `$portableTag`.
2. Upload artifact: `$(Split-Path -Leaf $zip)`.
3. Paste the metadata above into the release notes.
4. Attach `PORTABLE_RELEASE_MANIFEST.json` as a release asset (optional but recommended).
"@ | Set-Content "$deploy\PORTABLE_RELEASE_NOTES.md" -Encoding UTF8

Write-Host "Creating zip: $zip"
Compress-Archive -Path "$deploy\*" -DestinationPath $zip -CompressionLevel Optimal

Write-Host "Done."
Write-Host "Deploy folder: $deploy"
Write-Host "Deploy zip: $zip"
Write-Host "Suggested portable tag: $portableTag"

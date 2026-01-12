$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseDir = Split-Path -Parent (Split-Path -Parent $scriptDir)

$inputPath = Join-Path $baseDir "out\output.ods"
$outDir = Join-Path $baseDir "out"
$profileDir = Join-Path $baseDir "profile"

$profileUri = "file:///$($profileDir -replace '\\', '/')"

& soffice --headless --convert-to pdf --outdir "$outDir" "--env:UserInstallation=$profileUri" "$inputPath"

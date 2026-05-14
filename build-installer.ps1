$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$appDir = Join-Path $root "Installer\App"
$outputDir = Join-Path $root "Installer\Output"
$wixExe = Join-Path $root ".tools-wix4\wix.exe"

if (!(Test-Path $wixExe)) {
    dotnet tool install wix --tool-path (Join-Path $root ".tools-wix4") --version 4.0.6
}

foreach ($path in @($appDir, $outputDir)) {
    if (Test-Path $path) {
        Remove-Item -LiteralPath $path -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $path | Out-Null
}

dotnet publish (Join-Path $root "GovDocFormatter\GovDocFormatter.csproj") `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:EnableCompressionInSingleFile=true `
    -o $appDir

$pdb = Join-Path $appDir "GovDocFormatter.pdb"
if (Test-Path $pdb) {
    Remove-Item -LiteralPath $pdb -Force
}

& $wixExe build (Join-Path $root "Installer\Product.wxs") `
    -arch x64 `
    -d "AppSourceDir=$appDir" `
    -o (Join-Path $outputDir "GovDocFormatterSetup.msi") `
    -pdbtype none

Write-Host "Installer created:" (Join-Path $outputDir "GovDocFormatterSetup.msi")

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$version = "1.0.12"
$distDir = Join-Path $root "Dist"
$publishDir = Join-Path $distDir "GovDocFormatter-linux-x64-v$version"
$archivePath = Join-Path $distDir "GovDocFormatter-linux-x64-v$version.tar.gz"

if (Test-Path $publishDir) {
    Remove-Item -LiteralPath $publishDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $publishDir | Out-Null

dotnet publish (Join-Path $root "GovDocFormatter.Cli\GovDocFormatter.Cli.csproj") `
    -c Release `
    -r linux-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=None `
    -p:DebugSymbols=false `
    -o $publishDir

$pdb = Join-Path $publishDir "GovDocFormatter.pdb"
if (Test-Path $pdb) {
    Remove-Item -LiteralPath $pdb -Force
}

Copy-Item -LiteralPath (Join-Path $root "README-Linux.md") -Destination (Join-Path $publishDir "README-Linux.md") -Force

if (Test-Path $archivePath) {
    Remove-Item -LiteralPath $archivePath -Force
}

Push-Location $distDir
try {
    tar -czf (Split-Path -Leaf $archivePath) (Split-Path -Leaf $publishDir)
}
finally {
    Pop-Location
}

Write-Host "Linux package created:" $archivePath

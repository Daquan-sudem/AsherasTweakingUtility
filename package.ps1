param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectFile = Join-Path $projectDir "WinOptApp.csproj"
$publishRoot = Join-Path $projectDir "dist"
$publishDir = Join-Path $publishRoot "$Runtime-$Configuration"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$zipPath = Join-Path $publishRoot "AshereTweakingUtility-$Runtime-$stamp.zip"
$portableExePath = Join-Path $publishRoot "AshereTweakingUtility-$Runtime-$stamp.exe"

Write-Host "Publishing $projectFile ..."
dotnet publish $projectFile `
  -c $Configuration `
  -r $Runtime `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true `
  -o $publishDir

$builtExe = Join-Path $publishDir "AshereTweakingUtility.exe"
if (!(Test-Path $builtExe)) {
    throw "Expected published exe not found: $builtExe"
}

Write-Host "Creating portable exe: $portableExePath"
Copy-Item -Path $builtExe -Destination $portableExePath -Force

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Write-Host "Creating zip: $zipPath"
Compress-Archive -Path (Join-Path $publishDir "*") -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Done."
Write-Host "Publish folder: $publishDir"
Write-Host "Portable exe: $portableExePath"
Write-Host "Zip file: $zipPath"

Param(
    [string]$ProjectFolder = $PSScriptRoot,
    [string]$Configuration = "Release",
    [string]$SolutionPath = "",
    [string]$MsbuildPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\MSBuild.exe",
    [string]$OutputDir = ""
)

$ErrorActionPreference = "Stop"

$ProjectFolder = (Resolve-Path $ProjectFolder).Path
if ([string]::IsNullOrWhiteSpace($SolutionPath)) {
    $SolutionPath = Join-Path $ProjectFolder "NcTalkOutlookAddIn.sln"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $ProjectFolder "dist"
}

if (-not (Test-Path $MsbuildPath)) {
    throw "MSBuild.exe not found at $MsbuildPath"
}
if (-not (Test-Path $SolutionPath)) {
    throw "Solution not found at $SolutionPath"
}

Write-Host "Building solution using $MsbuildPath ($Configuration)..."
& $MsbuildPath $SolutionPath "/m" "/t:Rebuild" "/p:Configuration=$Configuration" | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild exited with code $LASTEXITCODE."
}

$buildOutputDir = Join-Path $ProjectFolder "src\\NcTalkOutlookAddIn\\bin\\$Configuration"
$dllPath = Join-Path $buildOutputDir "NcTalkOutlookAddIn.dll"
if (-not (Test-Path $dllPath)) {
    throw "Build succeeded but assembly not found at $dllPath."
}

$assemblyInfo = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath).Version
$assemblyVersionFull = $assemblyInfo.ToString()
$assemblyVersionShort = "{0}.{1}.{2}" -f $assemblyInfo.Major, $assemblyInfo.Minor, $assemblyInfo.Build
Write-Host "Assembly version detected: $assemblyVersionFull"

$wixProject = Join-Path $ProjectFolder "installer\\NcConnectorOutlookInstaller.wixproj"
if (-not (Test-Path $wixProject)) {
    throw "WiX project not found at $wixProject."
}

Write-Host "Building MSI via WiX v4 SDK (dotnet build)..."
& dotnet build $wixProject -c $Configuration "/p:BuildOutputDir=$buildOutputDir\\" "/p:ProductVersion=$assemblyVersionShort" "/p:AssemblyVersion=$assemblyVersionFull" | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "dotnet build (WiX) exited with code $LASTEXITCODE."
}

$builtMsiPath = Join-Path $ProjectFolder "installer\\bin\\$Configuration\\NCConnectorForOutlook.msi"
if (-not (Test-Path $builtMsiPath)) {
    throw "MSI build succeeded but output not found at $builtMsiPath."
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$finalName = "NCConnectorForOutlook-$assemblyVersionShort.msi"
$finalPath = Join-Path $OutputDir $finalName
Copy-Item -Force $builtMsiPath $finalPath

Write-Host "MSI erstellt: $finalPath"

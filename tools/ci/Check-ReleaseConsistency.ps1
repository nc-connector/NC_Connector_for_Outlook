Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path

function Assert-Check {
    Param(
        [bool]$Condition,
        [string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Read-Text {
    Param([string]$RelativePath)
    return Get-Content -Raw -Path (Join-Path $ProjectRoot $RelativePath)
}

$assemblyInfo = Read-Text "src\NcTalkOutlookAddIn\Properties\AssemblyInfo.cs"
$changelog = Read-Text "CHANGELOG.md"
$wixProject = [xml](Read-Text "installer\NcConnectorOutlookInstaller.wixproj")

$assemblyVersionMatch = [regex]::Match($assemblyInfo, 'AssemblyVersion\("(?<version>\d+\.\d+\.\d+\.\d+)"\)')
$fileVersionMatch = [regex]::Match($assemblyInfo, 'AssemblyFileVersion\("(?<version>\d+\.\d+\.\d+\.\d+)"\)')
Assert-Check $assemblyVersionMatch.Success "AssemblyVersion not found in AssemblyInfo.cs."
Assert-Check $fileVersionMatch.Success "AssemblyFileVersion not found in AssemblyInfo.cs."

$assemblyVersion = [version]$assemblyVersionMatch.Groups["version"].Value
$fileVersion = [version]$fileVersionMatch.Groups["version"].Value
Assert-Check ($assemblyVersion -eq $fileVersion) "AssemblyVersion ($assemblyVersion) and AssemblyFileVersion ($fileVersion) differ."
Assert-Check ($assemblyVersion.Revision -eq 0) "AssemblyVersion revision must stay 0 because MSI ProductVersion supports only three fields."

$shortVersion = "{0}.{1}.{2}" -f $assemblyVersion.Major, $assemblyVersion.Minor, $assemblyVersion.Build
$changelogMatch = [regex]::Match($changelog, '(?m)^## \[(?<version>\d+\.\d+\.\d+)\] - (?<date>\d{4}-\d{2}-\d{2})')
Assert-Check $changelogMatch.Success "No release heading found in CHANGELOG.md."
Assert-Check ($changelogMatch.Groups["version"].Value -eq $shortVersion) "Latest CHANGELOG version ($($changelogMatch.Groups["version"].Value)) does not match AssemblyVersion ($shortVersion)."

$propertyGroup = $wixProject.Project.PropertyGroup | Select-Object -First 1
$wixProductVersion = $propertyGroup.ProductVersion.InnerText
$wixAssemblyVersion = $propertyGroup.AssemblyVersion.InnerText
Assert-Check ($wixProductVersion -eq $shortVersion) "WiX default ProductVersion ($wixProductVersion) does not match AssemblyVersion ($shortVersion)."
Assert-Check ($wixAssemblyVersion -eq $assemblyVersion.ToString()) "WiX default AssemblyVersion ($wixAssemblyVersion) does not match AssemblyVersion ($assemblyVersion)."

$buildScript = Read-Text "build.ps1"
Assert-Check ($buildScript -match 'assemblyVersionShort\s*=\s*"\{0\}\.\{1\}\.\{2\}"') "build.ps1 must derive MSI ProductVersion from Major.Minor.Build."
Assert-Check ($buildScript -match 'AssemblyVersion=\$assemblyVersionFull') "build.ps1 must pass the full assembly version into WiX."

Write-Host "Release consistency OK: $shortVersion"

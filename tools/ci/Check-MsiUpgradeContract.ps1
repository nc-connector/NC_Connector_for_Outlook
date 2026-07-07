Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$ProductWxs = Join-Path $ProjectRoot "installer\Product.wxs"
$text = Get-Content -Raw -Path $ProductWxs
$failures = New-Object System.Collections.Generic.List[string]

function Assert-Contains {
    Param([string]$Pattern, [string]$Message)
    if ($text -notmatch $Pattern) {
        $failures.Add($Message)
    }
}

Assert-Contains 'UpgradeCode="\{8C9D7AA6-EBB2-4F7D-9B5F-A34C59F314B3\}"' "MSI UpgradeCode changed or is missing."
Assert-Contains '<MajorUpgrade\s+AllowDowngrades="yes"\s*/>' "MajorUpgrade rule is missing."
Assert-Contains '<Component\s+Id="cmpOutlookAddinReg"\s+Guid="\{8ED9BE3B-8138-4F3D-92FD-A5DA0CECA9F6\}"' "64-bit Outlook add-in registry component GUID changed or is missing."
Assert-Contains '<Component\s+Id="cmpOutlookAddinReg32"\s+Guid="\{BC89C35C-3ACF-4F5B-B851-AFE8576EE014\}"\s+Bitness="always32"' "32-bit Outlook add-in registry component is missing or not always32."
Assert-Contains 'Key="Software\\Microsoft\\Office\\Outlook\\Addins\\NcTalkOutlook\.AddIn"\s+Name="LoadBehavior"\s+Type="integer"\s+Value="3"' "Outlook LoadBehavior=3 registration is missing."
Assert-Contains 'Key="Software\\Classes\\CLSID\\\{A8CC9257-A153-4A01-AB35-D66CB3D44AAA\}\\InprocServer32"\s+Name="CodeBase"\s+Type="string"\s+Value="file:///\[INSTALLFOLDER\]NcTalkOutlookAddIn\.dll"' "COM CodeBase registration is missing."
Assert-Contains '<http:UrlReservation\s+Url="http://127\.0\.0\.1:7777/nc-ifb/"' "IFB URL reservation is missing."
Assert-Contains '<ComponentRef\s+Id="cmpNcTalkCore"\s*/>' "Main feature does not include core files."
Assert-Contains '<ComponentRef\s+Id="cmpOutlookAddinReg"\s*/>' "Main feature does not include 64-bit registry component."
Assert-Contains '<ComponentRef\s+Id="cmpOutlookAddinReg32"\s*/>' "Main feature does not include 32-bit registry component."
Assert-Contains '<ComponentRef\s+Id="cmpHttpUrlAcl"\s*/>' "Main feature does not include IFB URLACL component."

$requiredFiles = @(
    "NcTalkOutlookAddIn.dll",
    "NcTalkOutlookAddIn.dll.config",
    "AngleSharp.dll",
    "AngleSharp.Css.dll",
    "HtmlSanitizer.dll",
    "System.Buffers.dll",
    "System.Collections.Immutable.dll",
    "System.Memory.dll",
    "System.Runtime.CompilerServices.Unsafe.dll",
    "System.Text.Encoding.CodePages.dll",
    "LICENSE.txt",
    "VENDOR.md"
)

foreach ($fileName in $requiredFiles) {
    if ($text -notmatch [regex]::Escape('Source="$(var.BuildOutputDir)' + $fileName + '"')) {
        $failures.Add("MSI does not package $fileName.")
    }
}

if ($text -match '\.pdb"') {
    $failures.Add("MSI packages PDB/debug files.")
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "MSI upgrade contract check failed with $($failures.Count) issue(s)."
}

Write-Host "MSI upgrade contract OK: stable upgrade code, COM registry, URLACL and package inputs are present."

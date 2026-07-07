Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path

function Assert-Check {
    Param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Read-Text {
    Param([string]$RelativePath)
    return Get-Content -Raw -Path (Join-Path $ProjectRoot $RelativePath)
}

$csproj = Read-Text "src\NcTalkOutlookAddIn\NcTalkOutlookAddIn.csproj"
$wxs = Read-Text "installer\Product.wxs"
$vendor = Read-Text "VENDOR.md"

$expectedDlls = @(
    "AngleSharp.dll",
    "AngleSharp.Css.dll",
    "HtmlSanitizer.dll",
    "System.Buffers.dll",
    "System.Collections.Immutable.dll",
    "System.Memory.dll",
    "System.Runtime.CompilerServices.Unsafe.dll",
    "System.Text.Encoding.CodePages.dll"
)

foreach ($dll in $expectedDlls) {
    $relativeVendorPath = "src\NcTalkOutlookAddIn\vendor\htmlsanitizer\$dll"
    Assert-Check (Test-Path (Join-Path $ProjectRoot $relativeVendorPath)) "Vendored dependency missing: $relativeVendorPath"
    Assert-Check ($csproj -match [regex]::Escape("vendor\htmlsanitizer\$dll")) "Project file does not reference vendor DLL: $dll"
    Assert-Check ($wxs -match [regex]::Escape('$' + "(var.BuildOutputDir)" + $dll)) "Installer Product.wxs does not package DLL: $dll"
    Assert-Check ($vendor -match [regex]::Escape($relativeVendorPath.Replace("\", "/")) -or $vendor -match [regex]::Escape($relativeVendorPath)) "VENDOR.md does not document DLL: $dll"
}

Assert-Check ($csproj -match '<Content Include="LICENSE.txt"') "Project file must copy LICENSE.txt to output."
Assert-Check ($csproj -match '<Link>VENDOR\.md</Link>') "Project file must copy VENDOR.md to output."
Assert-Check ($wxs -match 'filLicenseTxt') "Installer must include LICENSE.txt."
Assert-Check ($wxs -match 'filVendorTxt') "Installer must include VENDOR.md."
Assert-Check ($wxs -match 'Bitness="always32"') "Installer must contain explicit 32-bit Outlook registry component."
Assert-Check ($wxs -match 'cmpOutlookAddinReg32') "Installer must contain 32-bit Outlook add-in registration."
Assert-Check ($wxs -match 'cmpOutlookAddinReg"') "Installer must contain 64-bit/default Outlook add-in registration."
Assert-Check ($wxs -match 'http:UrlReservation') "Installer must contain the IFB URL reservation."
Assert-Check (Test-Path (Join-Path $ProjectRoot "installer\assets\nc-connector.ico")) "Installer icon is missing."

Write-Host "Vendor/package check OK: $($expectedDlls.Count) DLLs documented, referenced and packaged."

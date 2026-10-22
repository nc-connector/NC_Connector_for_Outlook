Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$Failures = New-Object System.Collections.Generic.List[string]

function Test-Contains {
    Param(
        [string]$Name,
        [string]$Content,
        [string]$Expected
    )

    if ($Content.IndexOf($Expected, [System.StringComparison]::Ordinal) -lt 0) {
        $Failures.Add("$Name is missing: $Expected")
    }
}

$stringsPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\Strings.cs"
$readmePath = Join-Path $ProjectRoot "README.md"
$readmeDePath = Join-Path $ProjectRoot "README.de.md"
$adminPath = Join-Path $ProjectRoot "docs\ADMIN.md"
$adminDePath = Join-Path $ProjectRoot "docs\ADMIN.de.md"
$developmentPath = Join-Path $ProjectRoot "docs\DEVELOPMENT.md"
$developmentDePath = Join-Path $ProjectRoot "docs\DEVELOPMENT.de.md"

$strings = Get-Content -LiteralPath $stringsPath -Raw
$readme = Get-Content -LiteralPath $readmePath -Raw
$readmeDe = Get-Content -LiteralPath $readmeDePath -Raw
$admin = Get-Content -LiteralPath $adminPath -Raw
$adminDe = Get-Content -LiteralPath $adminDePath -Raw
$development = Get-Content -LiteralPath $developmentPath -Raw
$developmentDe = Get-Content -LiteralPath $developmentDePath -Raw

Test-Contains "Strings.cs" $strings "private static bool UsesGermanAdminGuide()"
Test-Contains "Strings.cs" $strings "docs/ADMIN.md#system-address-book"
Test-Contains "Strings.cs" $strings "docs/ADMIN.de.md#systemadressbuch"
Test-Contains "Strings.cs" $strings "docs/ADMIN.md"
Test-Contains "Strings.cs" $strings "docs/ADMIN.de.md"
Test-Contains "ADMIN.md" $admin "### System address book"
Test-Contains "ADMIN.de.md" $adminDe "### Systemadressbuch"

Test-Contains "README.md" $readme "64-bit Windows 10 or 11"
Test-Contains "README.de.md" $readmeDe "Windows 10 oder Windows 11 (64-Bit)"
Test-Contains "README.md" $readme "docs/ADMIN.md"
Test-Contains "README.de.md" $readmeDe "docs/ADMIN.de.md"
Test-Contains "README.md" $readme "docs/DEVELOPMENT.md"
Test-Contains "README.de.md" $readmeDe "docs/DEVELOPMENT.de.md"

$developmentSections = @(
    @("DEVELOPMENT.md", $development, "## Project purpose"),
    @("DEVELOPMENT.md", $development, "## Quick start"),
    @("DEVELOPMENT.md", $development, "## Repository structure"),
    @("DEVELOPMENT.md", $development, "## Architecture"),
    @("DEVELOPMENT.md", $development, "## Network endpoints"),
    @("DEVELOPMENT.md", $development, "## Localization (i18n)"),
    @("DEVELOPMENT.md", $development, "## Logging"),
    @("DEVELOPMENT.md", $development, "## Compatibility & version checks"),
    @("DEVELOPMENT.md", $development, "## Build & release"),
    @("DEVELOPMENT.md", $development, "## Local testing"),
    @("DEVELOPMENT.md", $development, "## X-NCTALK-* property reference"),
    @("DEVELOPMENT.md", $development, "## Extension points"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Projektzweck"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Schnellstart"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Repository-Struktur"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Architektur"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Netzwerk-Endpunkte"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Lokalisierung (i18n)"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Logging"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Kompatibilität und Versionsprüfungen"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Build und Release"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Lokales Testen"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Referenz der X-NCTALK-*-Eigenschaften"),
    @("DEVELOPMENT.de.md", $developmentDe, "## Erweiterungspunkte")
)

foreach ($section in $developmentSections) {
    Test-Contains $section[0] $section[1] $section[2]
}

if ($development.Contains("Release 3.1.0 delta summary") -or
    $developmentDe.Contains("Delta-Zusammenfassung für Release 3.1.0")) {
    $Failures.Add("DEVELOPMENT guides still contain a stale release-delta section.")
}

if ($Failures.Count -gt 0) {
    $Failures | ForEach-Object { Write-Host ("[FAIL] " + $_) -ForegroundColor Red }
    exit 1
}

Write-Host "Documentation links, requirements, and guide sections are consistent."

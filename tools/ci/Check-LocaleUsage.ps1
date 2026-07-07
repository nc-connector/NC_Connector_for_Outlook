Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"
$LocalesRoot = Join-Path $SourceRoot "Resources\_locales"
$EnglishFile = Join-Path $LocalesRoot "en\messages.json"
$english = Get-Content -Raw -Path $EnglishFile | ConvertFrom-Json
$englishKeys = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::Ordinal)
foreach ($property in $english.PSObject.Properties) {
    [void]$englishKeys.Add($property.Name)
}

$usedKeys = New-Object System.Collections.Generic.SortedSet[string]([StringComparer]::Ordinal)
$sourceFiles = Get-ChildItem -Path $SourceRoot -Recurse -Filter "*.cs" |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }

foreach ($file in $sourceFiles) {
    $text = Get-Content -Raw -Path $file.FullName
    foreach ($match in [regex]::Matches($text, '\bGet\s*\(\s*"(?<key>[a-zA-Z0-9_]+)"\s*,')) {
        [void]$usedKeys.Add($match.Groups["key"].Value)
    }
    foreach ($match in [regex]::Matches($text, '\bGetInLanguage\s*\([^,]+,\s*"(?<key>[a-zA-Z0-9_]+)"\s*,')) {
        [void]$usedKeys.Add($match.Groups["key"].Value)
    }
}

$failures = New-Object System.Collections.Generic.List[string]
foreach ($key in $usedKeys) {
    if (-not $englishKeys.Contains($key)) {
        $failures.Add("Code uses locale key '$key', but en/messages.json does not define it.")
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Locale usage check failed with $($failures.Count) issue(s)."
}

Write-Host "Locale usage OK: $($usedKeys.Count) code-used key(s) exist in en/messages.json."

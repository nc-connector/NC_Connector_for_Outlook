Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$LocalesRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Resources\_locales"
$EnglishFile = Join-Path $LocalesRoot "en\messages.json"

function Assert-Check {
    Param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Read-Locale {
    Param([string]$Path)
    try {
        return Get-Content -Raw -Path $Path | ConvertFrom-Json
    } catch {
        throw "Locale JSON is invalid: $Path. $($_.Exception.Message)"
    }
}

function Get-MessageKeys {
    Param($Json)
    return @($Json.PSObject.Properties.Name | Sort-Object)
}

function Get-Placeholders {
    Param([string]$Message)
    if ($null -eq $Message) {
        return @()
    }
    return @([regex]::Matches($Message, '(\$\d+|\{\d+\})') | ForEach-Object { $_.Value } | Sort-Object -Unique)
}

$english = Read-Locale $EnglishFile
$englishKeys = Get-MessageKeys $english
Assert-Check ($englishKeys.Count -gt 0) "English locale has no messages."

$failures = New-Object System.Collections.Generic.List[string]
foreach ($file in Get-ChildItem -Path $LocalesRoot -Recurse -Filter "messages.json") {
    $locale = Split-Path (Split-Path $file.FullName -Parent) -Leaf
    $json = Read-Locale $file.FullName
    $keys = Get-MessageKeys $json

    $missing = @($englishKeys | Where-Object { $_ -notin $keys })
    $extra = @($keys | Where-Object { $_ -notin $englishKeys })
    foreach ($key in $missing) {
        $failures.Add("$locale is missing key '$key'.")
    }
    foreach ($key in $extra) {
        $failures.Add("$locale has extra key '$key'.")
    }

    foreach ($key in $englishKeys) {
        if ($key -notin $keys) {
            continue
        }

        $entry = $json.$key
        if ($null -eq $entry -or $null -eq $entry.message -or -not ($entry.message -is [string])) {
            $failures.Add("$locale key '$key' has no string 'message' value.")
            continue
        }

        $expectedPlaceholders = Get-Placeholders $english.$key.message
        $actualPlaceholders = Get-Placeholders $entry.message
        if (($expectedPlaceholders -join "|") -ne ($actualPlaceholders -join "|")) {
            $failures.Add("$locale key '$key' placeholder mismatch. Expected [$($expectedPlaceholders -join ', ')], got [$($actualPlaceholders -join ', ')].")
        }
    }

    Write-Host "Locale checked: $locale ($($keys.Count) keys)"
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "Locale issue: $_" }
    throw "Locale consistency check failed with $($failures.Count) issue(s)."
}

Write-Host "Locale consistency OK: $($englishKeys.Count) keys across $((Get-ChildItem -Path $LocalesRoot -Directory).Count) locales."

Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SettingsStoragePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Settings\SettingsStorage.cs"
$AddinSettingsPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Settings\AddinSettings.cs"

$storage = Get-Content -Raw -Path $SettingsStoragePath
$settings = Get-Content -Raw -Path $AddinSettingsPath
$failures = New-Object System.Collections.Generic.List[string]

$savedKeys = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach ($match in [regex]::Matches($storage, 'Append(?:OptionalBool)?Element\s*\([^;]*?"(?<key>[^"]+)"', 'Singleline')) {
    [void]$savedKeys.Add($match.Groups["key"].Value)
}

$loadedKeys = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach ($match in [regex]::Matches($storage, 'case\s+"(?<key>[^"]+)"\s*:')) {
    [void]$loadedKeys.Add($match.Groups["key"].Value)
}
foreach ($match in [regex]::Matches($storage, 'string\.Equals\(key,\s*"(?<key>[^"]+)"')) {
    [void]$loadedKeys.Add($match.Groups["key"].Value)
}

$properties = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach ($match in [regex]::Matches($settings, 'public\s+(?:[\w\?<>]+)\s+(?<name>\w+)\s*\{\s*get;\s*set;\s*\}')) {
    [void]$properties.Add($match.Groups["name"].Value)
}
foreach ($match in [regex]::Matches($settings, 'internal\s+(?:[\w\?<>]+)\s+(?<name>Managed\w+)\s*\{\s*get;\s*private\s+set;\s*\}')) {
    [void]$properties.Add($match.Groups["name"].Value)
}

$specialSavedKeyToProperty = @{
    AppPasswordProtected = "AppPassword"
}
$nonPersistedProperties = @(
    "ManagedNextcloudUrl",
    "ManagedNextcloudUrlSource",
    "ManagedNextcloudUrlLocked"
)

foreach ($key in $savedKeys) {
    $propertyName = if ($specialSavedKeyToProperty.ContainsKey($key)) { $specialSavedKeyToProperty[$key] } else { $key }
    if (-not $properties.Contains($propertyName)) {
        $failures.Add("SettingsStorage persists '$key', but AddinSettings has no '$propertyName' property.")
    }
    if (-not $loadedKeys.Contains($key)) {
        $failures.Add("SettingsStorage saves '$key', but LoadFromXmlFile/ApplySettingValue does not load it.")
    }
}

foreach ($property in $properties) {
    if ($property -in $nonPersistedProperties) {
        continue
    }
    $savedName = $property
    if ($property -eq "AppPassword") {
        $savedName = "AppPasswordProtected"
    }
    if (-not $savedKeys.Contains($savedName)) {
        $failures.Add("AddinSettings.$property is not persisted by SettingsStorage.")
    }
}

if (-not ($storage -match 'ProtectedData\.Protect') -or -not ($storage -match 'ProtectedData\.Unprotect')) {
    $failures.Add("SettingsStorage must protect and unprotect AppPassword via DPAPI.")
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Settings persistence check failed with $($failures.Count) issue(s)."
}

Write-Host "Settings persistence OK: $($savedKeys.Count) persisted key(s) match AddinSettings and load paths."

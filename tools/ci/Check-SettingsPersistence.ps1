Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SettingsStoragePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Settings\SettingsStorage.cs"
$SettingsTransactionPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Settings\SettingsFileTransaction.cs"
$SettingsWorkflowPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Controllers\SettingsWorkflowController.cs"
$SettingsFormPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\SettingsForm.cs"
$AddinSettingsPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Settings\AddinSettings.cs"
$FileLinkWizardPolicyPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.Policy.cs"
$AddinLifecyclePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.Lifecycle.cs"

$storage = Get-Content -Raw -Path $SettingsStoragePath
$transaction = Get-Content -Raw -Path $SettingsTransactionPath
$workflow = Get-Content -Raw -Path $SettingsWorkflowPath
$settingsForm = Get-Content -Raw -Path $SettingsFormPath
$settings = Get-Content -Raw -Path $AddinSettingsPath
$fileLinkWizardPolicy = Get-Content -Raw -Path $FileLinkWizardPolicyPath
$addinLifecycle = Get-Content -Raw -Path $AddinLifecyclePath
$failures = New-Object System.Collections.Generic.List[string]

$savedKeys = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach ($match in [regex]::Matches($storage, 'Append(?:OptionalBool|LocalChoice)?Element\s*\([^;]*?"(?<key>[^"]+)"', 'Singleline')) {
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
foreach ($match in [regex]::Matches($settings, 'public\s+(?:[\w\?<>]+)\s+(?<name>\w+)\s*\{\s*get(?:;|\s*\{)')) {
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

if (-not ($storage -match 'catch\s*\(FormatException') -or -not ($storage -match 'catch\s*\(CryptographicException')) {
    $failures.Add("A malformed or undecryptable DPAPI app password must not abort loading all other settings.")
}

if (-not ($storage -match 'SaveUserInitiated') -or -not ($storage -match '_automaticSaveBlocked')) {
    $failures.Add("SettingsStorage must distinguish explicit saves from blocked automatic saves after recovery failures.")
}

if (-not ($storage -match 'TryRestorePrimaryFromBackup') -or -not ($storage -match '\.Commit\(')) {
    $failures.Add("SettingsStorage must use backup recovery and atomic commits.")
}

if (-not ($transaction -match 'new Mutex') -or -not ($transaction -match 'FileOptions\.WriteThrough') -or -not ($transaction -match 'Flush\(true\)')) {
    $failures.Add("Settings file commits must be profile-locked and flushed durably.")
}

if (-not ($transaction -match 'File\.Replace') -or -not ($transaction -match '\.bak')) {
    $failures.Add("Settings file commits must atomically replace the primary and retain a backup.")
}

$persistNextIndex = $workflow.IndexOf("_persistSettings(nextSettings);", [StringComparison]::Ordinal)
$applyRuntimeNextIndex = $workflow.IndexOf("ApplyRuntimeSettings(nextSettings);", [StringComparison]::Ordinal)
if ($persistNextIndex -lt 0 -or $applyRuntimeNextIndex -le $persistNextIndex) {
    $failures.Add("Settings must be persisted before durable runtime settings are applied.")
}

if (-not ($workflow -match 'catch\s*\(Exception ex\)[\s\S]{0,600}?Strings\.SettingsSaveFailed[\s\S]{0,300}?MessageBoxIcon\.Error')) {
    $failures.Add("Settings persistence failures must be reported visibly inside the settings workflow.")
}

$validationStart = $workflow.IndexOf("private bool ValidateTransportSecurityBeforeSave", [StringComparison]::Ordinal)
$runtimeApplyStart = $workflow.IndexOf("private void ApplyRuntimeSettings", [StringComparison]::Ordinal)
if ($validationStart -lt 0 -or $runtimeApplyStart -le $validationStart) {
    $failures.Add("Transport-security validation method could not be inspected.")
}
else {
    $validationMethod = $workflow.Substring($validationStart, $runtimeApplyStart - $validationStart)
    if ($validationMethod -match '_setCurrentSettings') {
        $failures.Add("Pre-save transport validation must not temporarily replace the shared runtime settings.")
    }
    if (-not ($validationMethod -match '_applyTransportSecurityFromSettings\s*\(\s*nextSettings\s*,')) {
        $failures.Add("Pre-save transport validation must evaluate the candidate settings directly.")
    }
}

if (-not ($addinLifecycle -match 'TryApplyTransportSecurityFromSettings\(\s*"startup"\s*,\s*false\s*\)')) {
    $failures.Add("Add-in startup must apply the persisted transport-security settings.")
}

if (-not ($workflow -match '_applyTransportSecurityFromSettings\(\s*nextSettings\s*,\s*"settings_save_commit"')) {
    $failures.Add("A successful settings save must apply the persisted transport-security settings.")
}

if (-not ($storage -match 'case\s+"SharingAttachmentLinkTarget"[\s\S]{0,500}?AttachmentLinkTargetPolicy\.TryParse[\s\S]{0,500}?\(AttachmentLinkTarget\?\)null')) {
    $failures.Add("An invalid SharingAttachmentLinkTarget value must load as an unset nullable value.")
}

$refreshStart = $settingsForm.IndexOf("private async Task<bool> RefreshSettingsServerStateAsync", [StringComparison]::Ordinal)
$applyPolicyStart = $settingsForm.IndexOf("private void ApplyBackendPolicyStatus", [StringComparison]::Ordinal)
if ($refreshStart -lt 0 -or $applyPolicyStart -le $refreshStart) {
    $failures.Add("Settings server-state refresh method could not be inspected.")
}
else {
    $refreshMethod = $settingsForm.Substring($refreshStart, $applyPolicyStart - $refreshStart)
    if ($refreshMethod -match '!policyStatus\.FetchSucceeded\)\s*\{[\s\S]{0,500}?return false;') {
        $failures.Add("An unavailable optional backend policy endpoint must not block local settings saves.")
    }
    if (-not ($refreshMethod -match 'Settings server state refresh failed\.[\s\S]{0,500}?return true;')) {
        $failures.Add("A settings server-state refresh exception must leave local settings save available.")
    }
}

$wizardDefaultsStart = $fileLinkWizardPolicy.IndexOf("private void ApplyPolicyDefaultsToSettings", [StringComparison]::Ordinal)
$wizardWarningStart = $fileLinkWizardPolicy.IndexOf("private void ApplyPolicyWarningUi", [StringComparison]::Ordinal)
if ($wizardDefaultsStart -lt 0 -or $wizardWarningStart -le $wizardDefaultsStart) {
    $failures.Add("FileLink wizard policy-default method could not be inspected.")
}
else {
    $wizardDefaultsMethod = $fileLinkWizardPolicy.Substring(
        $wizardDefaultsStart,
        $wizardWarningStart - $wizardDefaultsStart)
    $wizardSource = Get-Content (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/UI/FileLinkWizardForm.cs') -Raw
    if ($wizardSource -notmatch '_defaults\s*=\s*\(defaults\s*\?\?\s*new AddinSettings\(\)\)\.ResolvePolicyDefaults\(policyStatus\)') {
        $failures.Add('FileLink must resolve defaults through the shared local-choice-aware policy path.')
    }
    if ($wizardDefaultsMethod -match 'TryGetPolicy|GetPolicyString') {
        $failures.Add('FileLink must not apply a second independent backend-default overlay.')
    }
}

if ($settings -notmatch 'status\.IsLocked\(domain, key\)\s*\|\|\s*ResolveDefaultsSource\(status\)\s*==\s*"backend"\s*\|\|\s*!HasLocalValue\(propertyName\)' -or
    $storage -notmatch 'if\s*\(settings\.HasLocalValue\(name\)\)') {
    $failures.Add('Policy resolution and XML persistence must retain the distinction between missing and explicit local choices.')
}
if ($storage -notmatch 'if\s*\(settings\.DefaultsSource\s*!=\s*null\)\s*\{\s*AppendElement\(document, root, "DefaultsSource", settings\.DefaultsSource\);' -or
    $storage -match 'AppendElement\([^\r\n]*(ResolveDefaultsSource|HasManagedDefaultsSource|IsManagedDefaultsSourceValid)') {
    $failures.Add('Defaults source persistence must write only an explicit raw user choice, never the effective managed source.')
}
if ($settingsForm -notmatch 'TrackLocalPolicyChoices\(\)' -or $settingsForm -notmatch '!_applyingPolicyDefaults') {
    $failures.Add('Settings must track user choices without persisting programmatic backend overlays.')
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Settings persistence check failed with $($failures.Count) issue(s)."
}

Write-Host "Settings persistence OK: $($savedKeys.Count) persisted key(s) match AddinSettings and load paths."

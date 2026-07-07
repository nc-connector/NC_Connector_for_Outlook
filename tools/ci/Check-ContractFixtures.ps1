Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$ContractsRoot = Join-Path $ProjectRoot "tests\contracts"

function Assert-Check {
    Param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Read-Json {
    Param([string]$RelativePath)
    $path = Join-Path $ContractsRoot $RelativePath
    Assert-Check (Test-Path $path) "Contract fixture missing: $RelativePath"
    try {
        return Get-Content -Raw -Path $path | ConvertFrom-Json
    } catch {
        throw "Contract fixture is invalid JSON: $RelativePath. $($_.Exception.Message)"
    }
}

function Assert-Url {
    Param([string]$Value, [string]$Name)
    $uri = $null
    Assert-Check ([Uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$uri) -and $uri.Scheme -eq "https") "$Name must be an absolute HTTPS URL."
}

$update = Read-Json "update-check.outlook.json"
Assert-Check ($update.latest_version -match '^\d+\.\d+\.\d+$') "Update check latest_version must be semantic Major.Minor.Patch."
Assert-Check ($update.update_available -is [bool]) "Update check update_available must be boolean."
Assert-Check ($update.counted -is [bool]) "Update check counted must be boolean."
Assert-Url $update.release_url "Update check release_url"
Assert-Url $update.download_url "Update check download_url"
Assert-Check ($null -ne $update.changelog -and $null -ne $update.changelog.sections) "Update check changelog.sections is required."
foreach ($section in @("added", "changed", "fixed")) {
    Assert-Check ($null -ne $update.changelog.sections.$section) "Update check changelog.sections.$section is required."
}

$policy = Read-Json "backend-policy-status.outlook.json"
Assert-Check ($policy.status.seat_assigned -is [bool]) "Backend status seat_assigned must be boolean."
Assert-Check ($policy.status.is_valid -is [bool]) "Backend status is_valid must be boolean."
Assert-Check ($policy.status.seat_state -is [string]) "Backend status seat_state must be string."
foreach ($domain in @("share", "talk", "email_signature")) {
    Assert-Check ($null -ne $policy.policy.$domain) "Backend policy domain missing: $domain"
    Assert-Check ($null -ne $policy.policy_editable.$domain) "Backend policy_editable domain missing: $domain"
}
Assert-Check ($policy.policy.share.share_send_password_mode -in @("plain", "secrets", $null)) "Share password mode contract must be plain, secrets or null."

$secrets = Read-Json "secrets-create-response.json"
Assert-Check ($secrets.ocs.data.uuid -match '^[A-Za-z0-9_-]+$') "Secrets response uuid must be URL-safe."
if ($secrets.ocs.data.expires -is [DateTime]) {
    Assert-Check ($secrets.ocs.data.expires.Kind -ne [DateTimeKind]::Unspecified) "Secrets response expires must include timezone information."
} else {
    Assert-Check ([string]$secrets.ocs.data.expires -match '^\d{4}-\d{2}-\d{2}T') "Secrets response expires must be ISO-like."
}

Write-Host "Contract fixtures OK: update-check, backend-policy-status, secrets-create-response."

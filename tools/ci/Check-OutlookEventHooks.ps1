Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"

$runtimeFiles = Get-ChildItem -Path $SourceRoot -Filter "NextcloudTalkAddIn*.cs"
if ($runtimeFiles.Count -eq 0) {
    throw "No NextcloudTalkAddIn runtime files found."
}

$additions = New-Object System.Collections.Generic.List[object]
$removals = New-Object System.Collections.Generic.List[object]
$eventPattern = '(?<target>[A-Za-z_][\w\.]*)\.(?<event>\w+)\s*(?<op>\+=|-=)\s*(?<handler>[A-Za-z_]\w+)\s*;'

foreach ($file in $runtimeFiles) {
    $relative = $file.FullName.Substring($ProjectRoot.Length + 1)
    $lines = Get-Content -Path $file.FullName
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $match = [regex]::Match($line, $eventPattern)
        if (-not $match.Success) {
            continue
        }

        $entry = [PSCustomObject]@{
            File = $relative
            Line = $i + 1
            Target = $match.Groups["target"].Value
            Event = $match.Groups["event"].Value
            Handler = $match.Groups["handler"].Value
            Key = $match.Groups["event"].Value + "::" + $match.Groups["handler"].Value
        }

        if ($match.Groups["op"].Value -eq "+=") {
            $additions.Add($entry)
        } else {
            $removals.Add($entry)
        }
    }
}

$removalKeys = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::Ordinal)
foreach ($removal in $removals) {
    [void]$removalKeys.Add($removal.Key)
}

$failures = New-Object System.Collections.Generic.List[string]
foreach ($addition in $additions) {
    if (-not $removalKeys.Contains($addition.Key)) {
        $failures.Add("$($addition.File):$($addition.Line) subscribes $($addition.Event) += $($addition.Handler), but no matching unsubscribe was found.")
    }
}

foreach ($requiredEvent in @("InlineResponse", "InlineResponseClose")) {
    $matchingAddition = $additions | Where-Object { $_.Event -eq $requiredEvent } | Select-Object -First 1
    if ($null -eq $matchingAddition) {
        $failures.Add("Compose lifecycle does not subscribe Explorer.$requiredEvent.")
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Outlook event hook symmetry check failed with $($failures.Count) issue(s)."
}

Write-Host "Outlook event hook symmetry OK: $($additions.Count) subscription(s), $($removals.Count) unsubscribe(s)."

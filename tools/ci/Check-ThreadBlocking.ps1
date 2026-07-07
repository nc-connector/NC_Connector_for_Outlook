Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"
$failures = New-Object System.Collections.Generic.List[string]

$sourceFiles = Get-ChildItem -Path $SourceRoot -Recurse -Filter "*.cs" |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }

function Is-TaskResultAfterWhenAll {
    Param(
        [string[]]$Lines,
        [int]$LineIndex,
        [string]$TaskName
    )

    $start = [Math]::Max(0, $LineIndex - 12)
    $context = ($Lines[$start..$LineIndex] -join "`n")
    return $context -match ('await\s+Task\.WhenAll\s*\([^;]*\b' + [regex]::Escape($TaskName) + '\b')
}

$allowedGetResult = @(
    "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs"
)
$allowedWait = @(
    "src\NcTalkOutlookAddIn\Services\FreeBusyServer.cs"
)

foreach ($file in $sourceFiles) {
    $relative = $file.FullName.Substring($ProjectRoot.Length + 1)
    $lines = Get-Content -Path $file.FullName

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]

        if ($line -match '\.GetAwaiter\(\)\.GetResult\(\)' -and $relative -notin $allowedGetResult) {
            $failures.Add("${relative}:$($i + 1) uses GetAwaiter().GetResult() outside the documented BeforeAttachmentAdd sync boundary.")
        }

        if ($line -match '\.Wait\s*\(' -and $relative -notin $allowedWait) {
            $failures.Add("${relative}:$($i + 1) uses Task.Wait() outside the listener shutdown path.")
        }

        foreach ($match in [regex]::Matches($line, '\b(?<name>\w+Task)\.Result\b')) {
            $taskName = $match.Groups["name"].Value
            if (-not (Is-TaskResultAfterWhenAll -Lines $lines -LineIndex $i -TaskName $taskName)) {
                $failures.Add("${relative}:$($i + 1) reads $taskName.Result without a nearby await Task.WhenAll(...).")
            }
        }
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Thread-blocking check failed with $($failures.Count) issue(s)."
}

Write-Host "Thread-blocking check OK: checked $($sourceFiles.Count) C# files."

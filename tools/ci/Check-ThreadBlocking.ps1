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

function Add-UiDialogBoundaryFailures {
    Param(
        [string]$ControllerRelativePath,
        [string]$AsyncMethodPattern,
        [string]$UiMethodPattern,
        [string]$DispatcherPattern,
        [string]$DialogType,
        [string]$FlowName
    )

    $controllerPath = Join-Path $SourceRoot $ControllerRelativePath
    $controller = Get-Content -LiteralPath $controllerPath -Raw
    $asyncFlowMatch = [regex]::Match(
        $controller,
        ('(?s)' + $AsyncMethodPattern + '.*?(?=\s+' + $UiMethodPattern + ')'))
    if (-not $asyncFlowMatch.Success) {
        $failures.Add("$ControllerRelativePath does not expose the expected async/UI $FlowName flow boundary.")
        return
    }

    if ($asyncFlowMatch.Value -notmatch $DispatcherPattern) {
        $failures.Add("$ControllerRelativePath does not marshal the $FlowName dialog back to Outlook's UI thread.")
    }
    if ($asyncFlowMatch.Value -match ('new\s+' + [regex]::Escape($DialogType) + '\s*\(')) {
        $failures.Add("$ControllerRelativePath constructs $DialogType before entering the Outlook UI-thread boundary.")
    }
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

Add-UiDialogBoundaryFailures `
    -ControllerRelativePath "Controllers\FileLinkLaunchController.cs" `
    -AsyncMethodPattern 'internal\s+async\s+Task<bool>\s+RunFileLinkWizardForMailAsync' `
    -UiMethodPattern 'private\s+bool\s+RunFileLinkWizardOnUiThread' `
    -DispatcherPattern '\.RunOnOutlookUiThreadAsync\s*\(' `
    -DialogType "FileLinkWizardForm" `
    -FlowName "FileLink"

Add-UiDialogBoundaryFailures `
    -ControllerRelativePath "Controllers\TalkRibbonController.cs" `
    -AsyncMethodPattern 'internal\s+async\s+Task\s+OnTalkButtonPressedAsync' `
    -UiMethodPattern 'private\s+bool\s+RunTalkDialogOnUiThread' `
    -DispatcherPattern '\.RunOnOutlookUiThreadAsync\s*\(' `
    -DialogType "TalkLinkForm" `
    -FlowName "Talk"

Add-UiDialogBoundaryFailures `
    -ControllerRelativePath "Controllers\SettingsWorkflowController.cs" `
    -AsyncMethodPattern 'internal\s+async\s+Task\s+RunAsync' `
    -UiMethodPattern 'private\s+void\s+RunSettingsDialogOnUiThread' `
    -DispatcherPattern '_runOnOutlookUiThreadAsync\s*\(' `
    -DialogType "SettingsForm" `
    -FlowName "settings"

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Thread-blocking check failed with $($failures.Count) issue(s)."
}

Write-Host "Thread-blocking check OK: checked $($sourceFiles.Count) C# files."

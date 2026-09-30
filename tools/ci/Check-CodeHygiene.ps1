Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"
$ProjectFile = Join-Path $SourceRoot "NcTalkOutlookAddIn.csproj"

$failures = New-Object System.Collections.Generic.List[string]
$sourceFiles = Get-ChildItem -Path $SourceRoot -Recurse -Filter "*.cs" |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }

[xml]$projectXml = Get-Content -LiteralPath $ProjectFile -Raw
$compileItems = New-Object System.Collections.Generic.HashSet[string](
    [System.StringComparer]::OrdinalIgnoreCase)
foreach ($node in $projectXml.SelectNodes("//*[local-name()='Compile']")) {
    $include = [string]$node.Include
    if (-not [string]::IsNullOrWhiteSpace($include)) {
        $compileItems.Add($include.Replace("/", "\")) | Out-Null
    }
}

foreach ($file in $sourceFiles) {
    $projectRelative = $file.FullName.Substring($SourceRoot.Length + 1)
    if (-not $compileItems.Contains($projectRelative)) {
        $failures.Add("$projectRelative is not compiled by NcTalkOutlookAddIn.csproj.")
    }
}

foreach ($include in $compileItems) {
    if (-not (Test-Path -LiteralPath (Join-Path $SourceRoot $include))) {
        $failures.Add("NcTalkOutlookAddIn.csproj references missing source file '$include'.")
    }
}

foreach ($file in $sourceFiles) {
    $relative = $file.FullName.Substring($ProjectRoot.Length + 1)
    $text = Get-Content -Raw -Path $file.FullName

    foreach ($match in [regex]::Matches($text, 'catch\s*\(\s*(?<type>[\w\.]+(?:Exception)?)\s+(?<name>\w+)\s*\)')) {
        $name = $match.Groups["name"].Value
        if ($name -in @("error", "err", "exception", "caught")) {
            $failures.Add("$relative uses catch variable '$name'. Outlook code should use 'ex' for caught exceptions.")
        }
    }

    if ($text -match 'Console\.Write(Line)?\s*\(') {
        $failures.Add("$relative writes to Console. Use DiagnosticsLogger instead.")
    }

}

$asyncVoidAllowList = @(
    'OnTalkButtonPressed',
    'OnSettingsButtonPressed',
    'OnFileLinkButtonPressed',
    'OnAttachmentEvalTimerTick',
    'OnBeforeAddShareTimerTick',
    'OnEmailSignatureTimerTick',
    'OnSaveButtonClick',
    'OnSelectedTabChanged',
    'OnUpdateCheckButtonClick',
    'OnLoginFlowButtonClick',
    'OnShown',
    'OnTestButtonClick',
    'HandleFileListViewDragDrop'
)

foreach ($file in $sourceFiles) {
    $relative = $file.FullName.Substring($ProjectRoot.Length + 1)
    $text = Get-Content -Raw -Path $file.FullName
    foreach ($match in [regex]::Matches($text, 'async\s+void\s+(?<name>\w+)\s*\(')) {
        $name = $match.Groups["name"].Value
        if ($name -notin $asyncVoidAllowList) {
            $failures.Add("$relative declares new async void method '$name'. Add explicit review or refactor to Task.")
        }
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Code hygiene check failed with $($failures.Count) issue(s)."
}

Write-Host "Code hygiene OK: checked $($sourceFiles.Count) C# files."

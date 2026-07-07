Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"
$EntryPoint = Join-Path $SourceRoot "NextcloudTalkAddIn.cs"

if (-not (Test-Path $EntryPoint)) {
    throw "Ribbon entry point not found at $EntryPoint"
}

$entryText = Get-Content -Raw -Path $EntryPoint
$sourceText = Get-ChildItem -Path $SourceRoot -Recurse -Filter "*.cs" |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' } |
    ForEach-Object { Get-Content -Raw -Path $_.FullName } |
    Out-String

$callbackAttributes = @(
    "onLoad",
    "onAction",
    "getImage",
    "getLabel",
    "getVisible",
    "getEnabled",
    "getScreentip",
    "getSupertip"
)

$callbacks = New-Object System.Collections.Generic.List[object]
foreach ($attribute in $callbackAttributes) {
    $pattern = "(?i)\b$attribute\s*=\s*'(?<name>[^']+)'"
    foreach ($match in [regex]::Matches($entryText, $pattern)) {
        $callbacks.Add([PSCustomObject]@{
            Attribute = $attribute
            Name = $match.Groups["name"].Value
        })
    }
}

if ($callbacks.Count -eq 0) {
    throw "No ribbon callbacks were found in NextcloudTalkAddIn.cs."
}

$failures = New-Object System.Collections.Generic.List[string]
$uniqueCallbacks = $callbacks | Sort-Object Attribute, Name -Unique

foreach ($callback in $uniqueCallbacks) {
    $name = [regex]::Escape($callback.Name)
    $methodPattern = "(?ms)(?<signature>\b(public|internal|private|protected)\s+(async\s+)?(?<return>[\w\.]+)\s+$name\s*\((?<params>[^)]*)\))"
    $methodMatch = [regex]::Match($sourceText, $methodPattern)
    if (-not $methodMatch.Success) {
        $failures.Add("Ribbon callback '$($callback.Name)' referenced by $($callback.Attribute) has no matching method.")
        continue
    }

    $returnType = $methodMatch.Groups["return"].Value
    $parameters = $methodMatch.Groups["params"].Value

    switch ($callback.Attribute) {
        "onLoad" {
            if ($parameters -notmatch "IRibbonUI") {
                $failures.Add("Ribbon onLoad callback '$($callback.Name)' should receive IRibbonUI.")
            }
            if ($returnType -ne "void") {
                $failures.Add("Ribbon onLoad callback '$($callback.Name)' should return void.")
            }
        }
        "onAction" {
            if ($parameters -notmatch "IRibbonControl") {
                $failures.Add("Ribbon onAction callback '$($callback.Name)' should receive IRibbonControl.")
            }
            if ($returnType -ne "void") {
                $failures.Add("Ribbon onAction callback '$($callback.Name)' should return void.")
            }
        }
        "getImage" {
            if ($parameters -notmatch "IRibbonControl") {
                $failures.Add("Ribbon getImage callback '$($callback.Name)' should receive IRibbonControl.")
            }
            if ($returnType -notmatch "(IPictureDisp|Object|object)$") {
                $failures.Add("Ribbon getImage callback '$($callback.Name)' should return IPictureDisp/object, got '$returnType'.")
            }
        }
        "getLabel" {
            if ($returnType -ne "string") {
                $failures.Add("Ribbon getLabel callback '$($callback.Name)' should return string.")
            }
        }
        "getScreentip" {
            if ($returnType -ne "string") {
                $failures.Add("Ribbon getScreentip callback '$($callback.Name)' should return string.")
            }
        }
        "getSupertip" {
            if ($returnType -ne "string") {
                $failures.Add("Ribbon getSupertip callback '$($callback.Name)' should return string.")
            }
        }
        "getVisible" {
            if ($returnType -ne "bool") {
                $failures.Add("Ribbon getVisible callback '$($callback.Name)' should return bool.")
            }
        }
        "getEnabled" {
            if ($returnType -ne "bool") {
                $failures.Add("Ribbon getEnabled callback '$($callback.Name)' should return bool.")
            }
        }
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Ribbon callback contract check failed with $($failures.Count) issue(s)."
}

Write-Host "Ribbon callback contract OK: checked $($uniqueCallbacks.Count) callback reference(s)."

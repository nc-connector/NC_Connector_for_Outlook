Param(
    [string]$OutputDirectory = ".ci\refs"
)

$ErrorActionPreference = "Stop"

$searchRoots = @(
    (Join-Path $env:ProgramFiles "Microsoft Visual Studio"),
    ${env:ProgramFiles(x86)}
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path $_) }

$extensibilityDll = $null
foreach ($root in $searchRoots) {
    $extensibilityDll = Get-ChildItem $root -Recurse -Filter "Extensibility.dll" -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($extensibilityDll) {
        break
    }
}

if (-not $extensibilityDll) {
    throw "Extensibility.dll not found. The Outlook add-in build needs the Office extensibility reference on the CI runner."
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
Copy-Item -Force -Path $extensibilityDll.FullName -Destination $OutputDirectory

Write-Host "Using Extensibility.dll from $($extensibilityDll.FullName)"
Write-Output (Resolve-Path $OutputDirectory).Path

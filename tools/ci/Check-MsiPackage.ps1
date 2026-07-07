Param(
    [string]$ProjectRoot = ".",
    [string]$MsiPath = "installer\bin\Release\NCConnectorForOutlook.msi"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$MsiPath = Join-Path $ProjectRoot $MsiPath

function Assert-Check {
    Param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

Assert-Check (Test-Path $MsiPath) "MSI file not found: $MsiPath"
Assert-Check ((Get-Item $MsiPath).Length -gt 500KB) "MSI file is unexpectedly small: $MsiPath"

$installer = New-Object -ComObject WindowsInstaller.Installer
$database = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))

function Invoke-MsiQuery {
    Param([string]$Sql)
    $view = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $database, @($Sql))
    $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
    $rows = @()
    while ($true) {
        $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
        if ($null -eq $record) {
            break
        }
        $fieldCount = $record.GetType().InvokeMember("FieldCount", "GetProperty", $null, $record, $null)
        $values = @()
        for ($i = 1; $i -le $fieldCount; $i++) {
            $values += $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, @($i))
        }
        $rows += ,$values
    }
    $view.GetType().InvokeMember("Close", "InvokeMethod", $null, $view, $null) | Out-Null
    return $rows
}

$files = Invoke-MsiQuery "SELECT ``FileName`` FROM ``File``" | ForEach-Object { $_[0] }
foreach ($expected in @(
    "NcTalkOutlookAddIn.dll",
    "NcTalkOutlookAddIn.dll.config",
    "HtmlSanitizer.dll",
    "AngleSharp.dll",
    "AngleSharp.Css.dll",
    "System.Memory.dll",
    "System.Text.Encoding.CodePages.dll",
    "LICENSE.txt",
    "VENDOR.md"
)) {
    Assert-Check (@($files | Where-Object { $_ -like "*$expected*" }).Count -gt 0) "MSI File table does not contain $expected."
}

$registryRows = Invoke-MsiQuery "SELECT ``Key``, ``Name``, ``Value`` FROM ``Registry``"
$loadBehaviorRows = @($registryRows | Where-Object { $_[0] -eq "Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn" -and $_[1] -eq "LoadBehavior" })
Assert-Check ($loadBehaviorRows.Count -ge 2) "MSI must register Outlook add-in LoadBehavior in both registry views."
Assert-Check (@($registryRows | Where-Object { $_[0] -like "Software\Classes\CLSID\{A8CC9257-A153-4A01-AB35-D66CB3D44AAA}*" -and $_[1] -eq "Assembly" }).Count -ge 2) "MSI must register COM assembly entries."

Write-Host "MSI package check OK: $MsiPath"

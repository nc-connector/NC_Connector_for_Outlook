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

$binaries = Invoke-MsiQuery "SELECT ``Name`` FROM ``Binary``" | ForEach-Object { $_[0] }
Assert-Check ($binaries -contains "IfbCleanupExe") "MSI must embed its IFB cleanup executable."
$actions = Invoke-MsiQuery "SELECT ``Action``, ``Type``, ``Source``, ``Target`` FROM ``CustomAction``"
$sequences = Invoke-MsiQuery "SELECT ``Action``, ``Condition``, ``Sequence`` FROM ``InstallExecuteSequence``"
$expectedTypes = @{ CheckOutlookClosed = 2; RollbackIfbRegistry = 3330; CleanupIfbRegistry = 3074; CommitIfbRegistry = 3586 }
$condition = 'NOT UPGRADINGPRODUCTCODE AND (NOT Installed OR REINSTALL OR REMOVE~="ALL")'
foreach ($name in $expectedTypes.Keys) {
    $action = @($actions | Where-Object { $_[0] -eq $name })
    Assert-Check ($action.Count -eq 1 -and [int]$action[0][1] -eq $expectedTypes[$name] -and $action[0][2] -eq "IfbCleanupExe") "Incorrect IFB action type or binary: $name"
    $sequence = @($sequences | Where-Object { $_[0] -eq $name })
    Assert-Check ($sequence.Count -eq 1 -and $sequence[0][1] -eq $condition) "IFB action must cover install, upgrade, repair and uninstall: $name"
}
$previous = -1
foreach ($name in @("CheckOutlookClosed", "InstallValidate", "InstallInitialize", "RollbackIfbRegistry", "CleanupIfbRegistry", "CommitIfbRegistry", "InstallFinalize")) {
    $sequence = @($sequences | Where-Object { $_[0] -eq $name })
    Assert-Check ($sequence.Count -eq 1 -and [int]$sequence[0][2] -gt $previous) "Incorrect IFB scheduling at $name"
    $previous = [int]$sequence[0][2]
}

# Evaluate the compiled condition without executing any installation action.
$session = $installer.GetType().InvokeMember("OpenPackage", "InvokeMethod", $null, $installer, @($MsiPath, 1))
foreach ($scenario in @(
    @{ Name = "fresh install"; Installed = ""; REINSTALL = ""; REMOVE = ""; UPGRADINGPRODUCTCODE = ""; Expected = 1 },
    @{ Name = "new side of upgrade"; Installed = ""; REINSTALL = ""; REMOVE = ""; UPGRADINGPRODUCTCODE = ""; Expected = 1 },
    @{ Name = "repair"; Installed = "1"; REINSTALL = "ALL"; REMOVE = ""; UPGRADINGPRODUCTCODE = ""; Expected = 1 },
    @{ Name = "full uninstall"; Installed = "1"; REINSTALL = ""; REMOVE = "ALL"; UPGRADINGPRODUCTCODE = ""; Expected = 1 },
    @{ Name = "retiring upgrade package"; Installed = "1"; REINSTALL = ""; REMOVE = "ALL"; UPGRADINGPRODUCTCODE = "1"; Expected = 0 }
)) {
    foreach ($property in @("Installed", "REINSTALL", "REMOVE", "UPGRADINGPRODUCTCODE")) {
        $session.GetType().InvokeMember("Property", "SetProperty", $null, $session, @($property, $scenario[$property])) | Out-Null
    }
    $result = $session.GetType().InvokeMember("EvaluateCondition", "InvokeMethod", $null, $session, @($condition))
    Assert-Check ($result -eq $scenario.Expected) "Unexpected cleanup condition for $($scenario.Name)."
}
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($session)

Write-Host "MSI package check OK: files, registration, IFB actions and maintenance conditions: $MsiPath"

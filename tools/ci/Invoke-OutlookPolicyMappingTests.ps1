Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-policy-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "OutlookPolicyMappingTests.cs"
    @'
using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;

internal static class OutlookPolicyMappingTests
{
    private static int failures;

    private static void Check(string name, bool condition, string detail = "")
    {
        if (condition)
        {
            Console.WriteLine("[OK] " + name);
            return;
        }
        failures++;
        Console.Error.WriteLine("[FAIL] " + name + (string.IsNullOrEmpty(detail) ? "" : ": " + detail));
    }

    public static int Main()
    {
        TestPasswordDeliveryMode();
        TestSecretsExpireDays();
        TestLockedBackendValueWins();
        TestEditableBackendValueKeepsLocalChoice();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " policy mapping test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook policy mapping tests passed.");
        return 0;
    }

    private static void TestPasswordDeliveryMode()
    {
        Check("Password delivery parses secrets", SharePasswordDeliveryPolicy.ParseMode("secrets") == SharePasswordDeliveryMode.Secrets);
        Check("Password delivery defaults unknown values to plain", SharePasswordDeliveryPolicy.ParseMode("bad") == SharePasswordDeliveryMode.Plain);
        Check("Password delivery storage value plain", SharePasswordDeliveryPolicy.ToStorageValue(SharePasswordDeliveryMode.Plain) == "plain");
        Check("Password delivery storage value secrets", SharePasswordDeliveryPolicy.ToStorageValue(SharePasswordDeliveryMode.Secrets) == "secrets");
    }

    private static void TestSecretsExpireDays()
    {
        Check("Secrets expire clamps low values", SharePasswordDeliveryPolicy.ClampSecretsExpireDays(-5) == 1);
        Check("Secrets expire clamps high values", SharePasswordDeliveryPolicy.ClampSecretsExpireDays(999) == 365);
        Check("Secrets expire keeps valid values", SharePasswordDeliveryPolicy.ClampSecretsExpireDays(14) == 14);
    }

    private static void TestLockedBackendValueWins()
    {
        var status = BuildStatus("secrets", false, "14");
        SharePasswordDeliveryPolicy policy = SharePasswordDeliveryPolicy.Resolve(status, SharePasswordDeliveryMode.Plain);
        Check("Locked backend delivery mode wins", policy.Mode == SharePasswordDeliveryMode.Secrets, policy.Mode.ToString());
        Check("Locked backend expire days are used", policy.SecretsExpireDays == 14, policy.SecretsExpireDays.ToString());
        Check("Locked backend policy uses secrets", policy.UseSecrets);
    }

    private static void TestEditableBackendValueKeepsLocalChoice()
    {
        var status = BuildStatus("secrets", true, "30");
        SharePasswordDeliveryPolicy policy = SharePasswordDeliveryPolicy.Resolve(status, SharePasswordDeliveryMode.Plain);
        Check("Editable backend delivery mode keeps local choice", policy.Mode == SharePasswordDeliveryMode.Plain, policy.Mode.ToString());
        Check("Editable backend still supplies shared expire days", policy.SecretsExpireDays == 30, policy.SecretsExpireDays.ToString());
    }

    private static BackendPolicyStatus BuildStatus(string mode, bool editable, string expireDays)
    {
        return new BackendPolicyStatus(
            true,
            true,
            true,
            false,
            string.Empty,
            "policy",
            "policy_active",
            true,
            true,
            "active",
            new Dictionary<string, object>
            {
                { "share_send_password_mode", mode },
                { "share_secrets_expire_days", expireDays }
            },
            new Dictionary<string, object>(),
            new Dictionary<string, object>(),
            new Dictionary<string, object>
            {
                { "share_send_password_mode", editable },
                { "share_secrets_expire_days", false }
            },
            new Dictionary<string, object>(),
            new Dictionary<string, object>());
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\BackendPolicyStatus.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryMode.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryPolicy.cs")
    )

    $exe = Join-Path $TempRoot "OutlookPolicyMappingTests.exe"
    & $csc /nologo /target:exe "/out:$exe" /reference:System.dll /reference:System.Core.dll @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    & $exe
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    if (Test-Path $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}

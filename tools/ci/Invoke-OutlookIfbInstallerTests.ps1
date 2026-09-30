Param([string]$ProjectRoot = ".")

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("NC4OL-IfbInstaller-" + [Guid]::NewGuid().ToString("N"))
[void](New-Item -ItemType Directory -Path $tempRoot)
try {
    $source = @'
using System;
using System.Globalization;
using Microsoft.Win32;
using NcConnectorOutlookInstaller;
using NcTalkOutlookAddIn.Services;

internal static class IfbInstallerTests
{
    private const string Sid = "S-1-5-21-100-200-300-1001";
    private const string Calendar = @"Software\Microsoft\Office\16.0\Outlook\Options\Calendar";
    private const string Internet = Calendar + @"\Internet Free/Busy";
    private static readonly string Own = "http://127.0.0.1:7777/nc-ifb/" + new string('a', 64) + "/freebusy/%NAME%@%SERVER%.vfb";

    private static void Check(bool condition, string description)
    {
        if (!condition) { throw new Exception(description); }
        Console.WriteLine("[OK] " + description);
    }

    private static void Put(RegistryKey hive, string path, string name, object value, RegistryValueKind kind)
    {
        using (RegistryKey key = hive.CreateSubKey(path)) { key.SetValue(name, value, kind); }
    }

    private static object Get(RegistryKey hive, string path, string name)
    {
        using (RegistryKey key = hive.OpenSubKey(path))
        {
            return key == null ? null : key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
        }
    }

    public static void Main()
    {
        foreach (string value in new[] { Own,
            "http://localhost:8888/nc-ifb/freebusy/%NAME%@example.org.vfb",
            "http://[::1]:7777/nc-ifb/freebusy/%NAME%@%SERVER%.vfb" })
        {
            Check(IfbRegistryEndpoint.IsSearchPath(value), "Recognize supported own IFB path");
        }
        foreach (string value in new[] { "", null, Own.Replace("127.0.0.1", "example.org"),
            Own.Replace("http:", "https:"), Own + "?query=1", Own + "#fragment",
            Own.Replace(":7777", ":0"), Own.Replace(":7777", ":65536"),
            Own.Replace(new string('a', 64), "unrelated"),
            Own.Replace("127.0.0.1", "user@127.0.0.1"), Own.Replace("/nc-ifb/", "/other/../nc-ifb/") })
        {
            Check(!IfbRegistryEndpoint.IsSearchPath(value), "Reject foreign or malformed search path");
        }
        Check(IfbCleanupProgram.IsUserSid(Sid), "Domain/local user SID supported");
        Check(IfbCleanupProgram.IsUserSid("S-1-12-1-1-2-3-4"), "Entra user SID supported");
        Check(!IfbCleanupProgram.IsUserSid(Sid + "_Classes") && !IfbCleanupProgram.IsUserSid("S-1-5-18"),
            "Exclude classes and system hives");
        foreach (string culture in new[] { "de", "en", "fr", "cs", "es", "hu", "it", "ja", "nl", "pl", "pt-BR", "pt-PT", "ru", "zh-CN", "zh-TW" })
        {
            foreach (string key in new[] { "outlook", "rollback" })
            {
                string text = IfbCleanupProgram.SetupMessage(CultureInfo.GetCultureInfo(culture), key);
                Check(!string.IsNullOrWhiteSpace(text) && (culture == "en" || text != IfbCleanupProgram.SetupMessage(CultureInfo.GetCultureInfo("en"), key)),
                    "Installer message translated: " + culture + " / " + key);
            }
        }

        // Production cleanup receives an isolated test hive, never HKU/profile enumeration.
        string testPath = @"Software\NC4OL_IfbInstallerTests\" + Guid.NewGuid().ToString("N");
        try
        {
            using (RegistryKey root = Registry.CurrentUser.CreateSubKey(testPath))
            using (RegistryKey hive = root.CreateSubKey("User"))
            using (RegistryKey journal = root.CreateSubKey("Journal"))
            {
                Put(hive, Calendar, "FreeBusySearchPath", Own, RegistryValueKind.String);
                Put(hive, Internet, "Read URL", Own, RegistryValueKind.ExpandString);
                Put(hive, Calendar, "OtherValue", Own, RegistryValueKind.String);
                string policy = Calendar.Replace(@"Software\Microsoft", @"Software\Policies\Microsoft");
                Put(hive, policy, "FreeBusySearchPath", Own, RegistryValueKind.String);
                string foreign = Calendar.Replace("16.0", "15.0");
                Put(hive, foreign, "FreeBusySearchPath", "https://example.org/freebusy.vfb", RegistryValueKind.String);
                string invalidVersion = Calendar.Replace("16.0", "OtherProduct");
                Put(hive, invalidVersion, "FreeBusySearchPath", Own, RegistryValueKind.String);
                string nonString = Calendar.Replace("16.0", "14.0");
                Put(hive, nonString, "FreeBusySearchPath", new[] { Own }, RegistryValueKind.MultiString);

                IfbCleanupProgram.CleanUserHive(hive, Sid, journal);
                Check(Get(hive, Calendar, "FreeBusySearchPath") == null && Get(hive, Internet, "Read URL") == null,
                    "Fresh install / upgrade / uninstall removes both own values without data directory");
                Check(journal.GetValueNames().Length == 2, "Journal contains only removed entries");
                Check((string)Get(hive, policy, "FreeBusySearchPath") == Own && (string)Get(hive, Calendar, "OtherValue") == Own,
                    "Policies and unrelated values remain unchanged");
                Check((string)Get(hive, foreign, "FreeBusySearchPath") == "https://example.org/freebusy.vfb"
                    && (string)Get(hive, invalidVersion, "FreeBusySearchPath") == Own
                    && Get(hive, nonString, "FreeBusySearchPath") is string[], "Foreign providers, product keys and types remain unchanged");
                IfbCleanupProgram.CleanUserHive(hive, Sid, journal);
                Check(journal.GetValueNames().Length == 2, "Repeated cleanup adds no duplicate snapshots");
                IfbCleanupProgram.RestoreUserHive(hive, journal, "S-1-5-21-100-200-300-1002");
                Check(journal.GetValueNames().Length == 2, "Rollback is scoped to the original user");
                IfbCleanupProgram.RestoreUserHive(hive, journal, Sid);
                Check((string)Get(hive, Calendar, "FreeBusySearchPath") == Own && (string)Get(hive, Internet, "Read URL") == Own,
                    "Failed MSI rollback restores exact prior own URLs");
                using (RegistryKey key = hive.OpenSubKey(Internet))
                {
                    Check(key.GetValueKind("Read URL") == RegistryValueKind.ExpandString, "Rollback preserves original registry type");
                }
                Check(journal.GetValueNames().Length == 0, "Successful rollback clears entries");
                IfbCleanupProgram.CleanUserHive(hive, Sid, journal);
                Put(hive, Calendar, "FreeBusySearchPath", "https://new.example.org/freebusy.vfb", RegistryValueKind.String);
                IfbCleanupProgram.RestoreUserHive(hive, journal, Sid);
                Check((string)Get(hive, Calendar, "FreeBusySearchPath") == "https://new.example.org/freebusy.vfb",
                    "Rollback does not overwrite a later external change");
                IfbCleanupProgram.CleanUserHive(hive, Sid, journal);
                hive.DeleteSubKeyTree(Internet);
                IfbCleanupProgram.RestoreUserHive(hive, journal, Sid);
                Check(hive.OpenSubKey(Internet) == null, "Rollback respects a later removed registry key");
            }
        }
        finally { Registry.CurrentUser.DeleteSubKeyTree(testPath, false); }
        Console.WriteLine("All IFB installer tests passed.");
    }
}
'@
    $sourcePath = Join-Path $tempRoot "IfbInstallerTests.cs"
    [IO.File]::WriteAllText($sourcePath, $source, (New-Object Text.UTF8Encoding($false)))
    $exePath = Join-Path $tempRoot "IfbInstallerTests.exe"
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    & $csc /nologo /target:exe /main:IfbInstallerTests "/out:$exePath" /reference:System.dll /reference:System.Core.dll /reference:System.Security.dll /reference:System.Xml.dll `
        "/resource:$ProjectRoot\installer\IfbCleanup\SetupMessages.xml,NcConnectorIfbCleanup.SetupMessages.xml" `
        $sourcePath "$ProjectRoot\installer\IfbCleanup\IfbCleanupProgram.cs" "$ProjectRoot\src\NcTalkOutlookAddIn\Services\IfbRegistryEndpoint.cs"
    if ($LASTEXITCODE -ne 0) { throw "IFB installer test compilation failed." }
    & $exePath
    if ($LASTEXITCODE -ne 0) { throw "IFB installer tests failed." }
}
finally {
    if ((Split-Path $tempRoot -Leaf) -like "NC4OL-IfbInstaller-*" -and [IO.Path]::GetFullPath($tempRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()))) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}

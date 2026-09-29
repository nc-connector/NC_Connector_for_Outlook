Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-security-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "OutlookSecurityRegressionTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class AppDataPaths
    {
        internal static string GetLocalRootDirectory()
        {
            return Path.GetTempPath();
        }
    }

    internal static class AddinVersionInfo
    {
        internal static string GetVersion()
        {
            return "3.3.0";
        }
    }

    internal static class Strings
    {
        internal static string TalkVersionUnknown { get { return "unknown"; } }
        internal static string UpdateChangelogEmpty { get { return "empty"; } }
        internal static string UpdateAvailableMessageFormat { get { return "{0}: {1}"; } }
        internal static string UpdateChangelogAdded { get { return "Added"; } }
        internal static string UpdateChangelogChanged { get { return "Changed"; } }
        internal static string UpdateChangelogFixed { get { return "Fixed"; } }
        internal static string ConnectionFailureCertificateSummary { get { return "certificate"; } }
        internal static string ConnectionFailureCertificateGuidance { get { return "certificate guidance"; } }
        internal static string ConnectionFailureDnsSummary { get { return "dns"; } }
        internal static string ConnectionFailureDnsGuidance { get { return "dns guidance"; } }
        internal static string ConnectionFailureProxySummary { get { return "proxy"; } }
        internal static string ConnectionFailureProxyGuidance { get { return "proxy guidance"; } }
        internal static string ConnectionFailureTimeoutSummary { get { return "timeout"; } }
        internal static string ConnectionFailureTimeoutGuidance { get { return "timeout guidance"; } }
        internal static string ConnectionFailureTlsSummary { get { return "tls"; } }
        internal static string ConnectionFailureTlsGuidance { get { return "tls guidance"; } }
        internal static string ConnectionFailureGenericSummary { get { return "generic"; } }
        internal static string ConnectionFailureGenericGuidance { get { return "generic guidance"; } }
        internal static string ManagedTlsPolicyInvalid { get { return "Managed TLS policy is invalid."; } }
    }
}

namespace NcTalkOutlookAddIn.Settings
{
    internal sealed class AddinSettings
    {
        internal AddinSettings() { IsManagedTransportTlsValid = true; }
        internal bool UpdateNotifyEnabled { get; set; }
        internal string UpdateInstallId { get; set; }
        internal string UpdateLastCheckedAtUtc { get; set; }
        internal string UpdateLatestVersion { get; set; }
        internal string UpdateReleaseUrl { get; set; }
        internal string UpdateDownloadUrl { get; set; }
        internal string UpdatePublishedAt { get; set; }
        internal string UpdateChangelogTitle { get; set; }
        internal string UpdateChangelogText { get; set; }
        internal string UpdateLastNotifiedVersion { get; set; }
        internal string UpdateLastNotifiedDateUtc { get; set; }
        internal bool TransportTlsUseSystemDefault { get; set; }
        internal bool TransportTlsEnable12 { get; set; }
        internal bool TransportTlsEnable13 { get; set; }
        internal bool HasManagedTransportTls { get; set; }
        internal bool IsManagedTransportTlsValid { get; set; }
    }
}

internal static class OutlookSecurityRegressionTests
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
        TestNextcloudUriBoundary();
        TestStructuredSecretRedaction();
        TestUpdateTargetPolicy();
        TestAtomicSettingsTransaction();
        TestTransportSecurityConfigurator();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " security regression test(s) failed.");
            return 1;
        }

        Console.WriteLine("All Outlook security regression tests passed.");
        return 0;
    }

    private static void TestNextcloudUriBoundary()
    {
        string normalized;
        Check(
            "Nextcloud base URL defaults to HTTPS",
            NextcloudUriValidator.TryNormalizeBaseUrl("cloud.example.test/nextcloud/", out normalized)
                && normalized == "https://cloud.example.test/nextcloud",
            normalized);
        Check(
            "Explicit HTTP Nextcloud base URL is rejected",
            !NextcloudUriValidator.TryNormalizeBaseUrl("http://cloud.example.test", out normalized));
        Check(
            "Nextcloud URL credentials are rejected",
            !NextcloudUriValidator.TryNormalizeBaseUrl("https://user:secret@cloud.example.test", out normalized));

        string resolved;
        Check(
            "Relative endpoint remains on configured Nextcloud path",
            NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "index.php/login/v2",
                "https://cloud.example.test/nextcloud",
                out resolved)
                && resolved == "https://cloud.example.test/nextcloud/index.php/login/v2",
            resolved);
        Check(
            "Cross-origin endpoint is rejected",
            !NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "https://evil.example.test/token",
                "https://cloud.example.test",
                out resolved));
        Check(
            "Scheme-relative cross-origin endpoint is rejected",
            !NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "//evil.example.test/token",
                "https://cloud.example.test",
                out resolved));
        Check(
            "Different Nextcloud port is a different origin",
            !NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "https://cloud.example.test:8443/token",
                "https://cloud.example.test",
                out resolved));
    }

    private static void TestStructuredSecretRedaction()
    {
        DiagnosticsLogger.SetAnonymization(false, string.Empty);
        MethodInfo sanitizer = typeof(DiagnosticsLogger).GetMethod(
            "SanitizeMessage",
            BindingFlags.NonPublic | BindingFlags.Static);
        string input =
            "pollToken=alpha token: 'bravo' appPassword=\"charlie\" "
            + "{\"roomToken\":\"delta\"} https://cloud.example.test/path?shareToken=echo";
        string output = Convert.ToString(sanitizer.Invoke(null, new object[] { input }));

        Check(
            "Structured secrets are redacted with PII anonymization disabled",
            !output.Contains("alpha")
                && !output.Contains("bravo")
                && !output.Contains("charlie")
                && !output.Contains("delta")
                && !output.Contains("echo"),
            output);
        Check("Redaction marker is present", output.Contains("<REDACTED>"), output);
    }

    private static void TestUpdateTargetPolicy()
    {
        var result = new UpdateCheckResult
        {
            DownloadUrl = "https://evil.example.test/update.msi",
            ReleaseUrl = "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/tag/v3.4.0"
        };
        Check(
            "Untrusted download falls back to trusted release page",
            UpdateCheckService.GetPreferredOpenUrl(result)
                == "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/tag/v3.4.0");

        result.DownloadUrl =
            "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/download/v3.4.0/NC-Connector.msi";
        Check(
            "Trusted GitHub release download is accepted",
            UpdateCheckService.GetPreferredOpenUrl(result) == result.DownloadUrl);

        result.DownloadUrl = "http://github.com/nc-connector/NC_Connector_for_Outlook/releases/download/v3.4.0/file.msi";
        result.ReleaseUrl = "https://github.com/another/repository/releases/tag/v3.4.0";
        Check(
            "HTTP and wrong-repository update targets are rejected",
            UpdateCheckService.GetPreferredOpenUrl(result) == string.Empty);

        var settings = new AddinSettings
        {
            UpdateLatestVersion = "3.2.9",
            UpdateDownloadUrl =
                "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/download/v3.2.9/file.msi"
        };
        UpdateCheckResult cached = UpdateCheckService.BuildCachedResult(settings);
        Check("Cached server state cannot downgrade local version comparison", !cached.UpdateAvailable);

        settings.UpdateLatestVersion = "3.4.0";
        cached = UpdateCheckService.BuildCachedResult(settings);
        Check("Newer cached version is detected locally", cached.UpdateAvailable);
    }

    private static void TestAtomicSettingsTransaction()
    {
        string directory = Path.Combine(Path.GetTempPath(), "nc4ol-settings-transaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string primary = Path.Combine(directory, "settings_profile.xml");
            var transaction = new SettingsFileTransaction(primary);
            Func<string, bool> healthy = path =>
                File.Exists(path) && File.ReadAllText(path, Encoding.UTF8).StartsWith("valid:", StringComparison.Ordinal);

            using (transaction.AcquireLock())
            {
                transaction.Commit(stream => Write(stream, "valid:first"), healthy);
                transaction.Commit(stream => Write(stream, "valid:second"), healthy);
            }

            Check("Atomic settings commit writes the new primary", File.ReadAllText(primary) == "valid:second");
            Check("Atomic settings commit keeps the previous valid backup", File.ReadAllText(transaction.BackupPath) == "valid:first");

            File.WriteAllText(primary, "corrupt", Encoding.UTF8);
            using (transaction.AcquireLock())
            {
                transaction.Commit(stream => Write(stream, "valid:third"), healthy);
            }
            Check("Corrupt primary does not replace the last valid backup", File.ReadAllText(transaction.BackupPath) == "valid:first");

            File.WriteAllText(primary, "corrupt-again", Encoding.UTF8);
            using (transaction.AcquireLock())
            {
                Check("Valid backup restores the primary", transaction.TryRestorePrimaryFromBackup(healthy));
            }
            Check("Restored primary equals the valid backup", File.ReadAllText(primary) == "valid:first");
            Check("Atomic settings transaction leaves no temp files", Directory.GetFiles(directory, "*.tmp").Length == 0);
        }
        finally
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
    }

    private static void TestTransportSecurityConfigurator()
    {
        SecurityProtocolType previous = ServicePointManager.SecurityProtocol;
        var previousSnapshot = new Dictionary<FieldInfo, object>();
        foreach (FieldInfo field in typeof(TransportSecurityConfigurator).GetFields(BindingFlags.Static | BindingFlags.NonPublic))
        {
            if (!field.IsLiteral && !field.IsInitOnly) previousSnapshot[field] = field.GetValue(null);
        }
        try
        {
            var settings = new AddinSettings
            {
                TransportTlsUseSystemDefault = false,
                TransportTlsEnable12 = true,
                TransportTlsEnable13 = false
            };
            SecurityProtocolType applied =
                TransportSecurityConfigurator.ApplyFromSettings(settings, "security_regression_test");

            Check(
                "TLS 1.2 setting changes the active runtime protocol",
                applied == SecurityProtocolType.Tls12
                    && ServicePointManager.SecurityProtocol == SecurityProtocolType.Tls12,
                ServicePointManager.SecurityProtocol.ToString());

            bool systemDefaultTlsDisabled;
            bool strongCryptoDisabled;
            Check(
                "Runtime enables system-default TLS support",
                AppContext.TryGetSwitch(
                    "Switch.System.Net.DontEnableSystemDefaultTlsVersions",
                    out systemDefaultTlsDisabled)
                    && !systemDefaultTlsDisabled);
            Check(
                "Runtime enables strong cryptography support",
                AppContext.TryGetSwitch(
                    "Switch.System.Net.DontEnableSchUseStrongCrypto",
                    out strongCryptoDisabled)
                    && !strongCryptoDisabled);
            Check(
                "System-default TLS maps to the runtime default",
                TransportSecurityConfigurator.BuildProtocol(true, false, false)
                    == SecurityProtocolType.SystemDefault);
            Check(
                "TLS 1.3 selection keeps its runtime protocol flag",
                (int)TransportSecurityConfigurator.BuildProtocol(false, false, true) == 12288);

            ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault;
            var unmanagedRequest = TransportSecurityConfigurator.CreateRequest("https://example.invalid/");
            Check("Unmanaged request creation keeps the existing runtime protocol", ServicePointManager.SecurityProtocol == SecurityProtocolType.SystemDefault);
            unmanagedRequest.Abort();
            TransportSecurityConfigurator.Restore(SecurityProtocolType.Tls12);
            Check("Unmanaged cancellation restores the previous runtime protocol", ServicePointManager.SecurityProtocol == SecurityProtocolType.Tls12);

            settings.HasManagedTransportTls = true;
            TransportSecurityConfigurator.ApplyFromSettings(settings, "managed_tls12_test");
            settings.TransportTlsUseSystemDefault = true;
            settings.TransportTlsEnable12 = false;
            settings.TransportTlsEnable13 = true;
            settings.HasManagedTransportTls = false;
            settings.IsManagedTransportTlsValid = false;
            applied = TransportSecurityConfigurator.Apply(true, false, false, "managed_preview_test");
            Check("Preview cannot replace the immutable managed TLS snapshot",
                applied == SecurityProtocolType.Tls12 && ServicePointManager.SecurityProtocol == SecurityProtocolType.Tls12);
            TransportSecurityConfigurator.Restore(SecurityProtocolType.SystemDefault);
            Check("Cancellation cannot undo a managed TLS 1.2 snapshot", ServicePointManager.SecurityProtocol == SecurityProtocolType.Tls12);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault;
            var managedRequest = TransportSecurityConfigurator.CreateRequest("https://example.invalid/");
            Check("Managed TLS 1.2 is reasserted before request construction", ServicePointManager.SecurityProtocol == SecurityProtocolType.Tls12);
            managedRequest.Abort();

            settings.HasManagedTransportTls = true;
            settings.IsManagedTransportTlsValid = true;
            settings.TransportTlsEnable13 = false;
            applied = TransportSecurityConfigurator.ApplyFromSettings(settings, "managed_system_default_test");
            Check("Managed system default overrides explicit preview values",
                applied == SecurityProtocolType.SystemDefault
                    && TransportSecurityConfigurator.Apply(false, true, false, "managed_system_preview_test") == SecurityProtocolType.SystemDefault);
            TransportSecurityConfigurator.Restore(SecurityProtocolType.Tls12);
            Check("Cancellation cannot undo managed OS-default TLS", ServicePointManager.SecurityProtocol == SecurityProtocolType.SystemDefault);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            managedRequest = TransportSecurityConfigurator.CreateRequest("https://example.invalid/");
            Check("Managed system default is reasserted before request construction", ServicePointManager.SecurityProtocol == SecurityProtocolType.SystemDefault);
            managedRequest.Abort();

            settings.IsManagedTransportTlsValid = false;
            settings.TransportTlsUseSystemDefault = false;
            settings.TransportTlsEnable12 = false;
            SecurityProtocolType beforeInvalid = ServicePointManager.SecurityProtocol;
            ExpectInvalidManagedTls("Invalid managed policy rejects startup application", delegate {
                TransportSecurityConfigurator.ApplyFromSettings(settings, "invalid_managed_test");
            });
            Check("Invalid managed policy does not silently apply a fallback", ServicePointManager.SecurityProtocol == beforeInvalid);
            ExpectInvalidManagedTls("Invalid managed policy cannot be bypassed by preview", delegate {
                TransportSecurityConfigurator.Apply(false, true, false, "invalid_managed_preview_test");
            });
            TransportSecurityConfigurator.Restore(SecurityProtocolType.Tls12);
            Check("Cancellation cannot substitute a protocol for invalid managed policy", ServicePointManager.SecurityProtocol == beforeInvalid);
            ExpectInvalidManagedTls("Invalid policy rejects before URL parsing or request construction", delegate {
                TransportSecurityConfigurator.CreateRequest("not a valid absolute request URL");
            });

            // A live loopback listener proves the request guard itself never opens a connection.
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            try
            {
                int port = ((IPEndPoint)listener.LocalEndpoint).Port;
                managedRequest = null;
                ExpectInvalidManagedTls("Invalid managed policy rejects requests before network I/O", delegate {
                    managedRequest = TransportSecurityConfigurator.CreateRequest("https://127.0.0.1:" + port + "/");
                });
                Check("Invalid managed policy does not construct a request", managedRequest == null);
                Check("Invalid managed request does not contact the network", !listener.Pending());
                Check("Invalid managed request does not get a fallback protocol", ServicePointManager.SecurityProtocol == beforeInvalid);
                if (managedRequest != null) managedRequest.Abort();
            }
            finally { listener.Stop(); }

            settings.HasManagedTransportTls = false;
            settings.IsManagedTransportTlsValid = true;
            settings.TransportTlsEnable12 = true;
            TransportSecurityConfigurator.ApplyFromSettings(settings, "managed_policy_removed_test");
            TransportSecurityConfigurator.Restore(SecurityProtocolType.SystemDefault);
            Check("Removing policy allows ordinary cancellation restoration again", ServicePointManager.SecurityProtocol == SecurityProtocolType.SystemDefault);
            unmanagedRequest = TransportSecurityConfigurator.CreateRequest("https://example.invalid/");
            Check("Removing policy clears the request override and invalid snapshot", ServicePointManager.SecurityProtocol == SecurityProtocolType.SystemDefault);
            unmanagedRequest.Abort();
        }
        finally
        {
            foreach (KeyValuePair<FieldInfo, object> entry in previousSnapshot) entry.Key.SetValue(null, entry.Value);
            ServicePointManager.SecurityProtocol = previous;
        }
    }

    private static void ExpectInvalidManagedTls(string name, Action action)
    {
        try { action(); Check(name, false, "No exception was thrown."); }
        catch (InvalidOperationException ex) { Check(name, ex.Message == Strings.ManagedTlsPolicyInvalid, ex.Message); }
    }

    private static void Write(Stream stream, string content)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\UpdateCheckResult.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\UpdateCheckService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Settings\SettingsFileTransaction.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\DiagnosticsLogger.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HttpFailureDiagnostics.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\LogCategories.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudUriValidator.cs")
    )
    $references = @(
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Security.dll",
        "/reference:System.Web.Extensions.dll"
    )

    $exe = Join-Path $TempRoot "OutlookSecurityRegressionTests.exe"
    & $csc /nologo /target:exe "/out:$exe" @references @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    & $exe
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    foreach ($clientFile in @('NcHttpClient.cs', 'UpdateCheckService.cs')) {
        $clientSource = Get-Content -LiteralPath (Join-Path $ProjectRoot ("src\NcTalkOutlookAddIn\Services\" + $clientFile)) -Raw
        $guard = [regex]::Match($clientSource, '\bTransportSecurityConfigurator\s*\.\s*CreateRequest\s*\(')
        if (-not $guard.Success -or $clientSource -match '\b(?:Http)?WebRequest\s*\.\s*Create(?:Http)?\s*\(') {
            throw "$clientFile must create every request through the central managed TLS guard."
        }
        $networkCalls = [regex]::Matches($clientSource, '\brequest\s*\.\s*(?:GetRequestStream(?:Async)?|GetResponse(?:Async)?)\s*\(')
        if ($networkCalls.Count -eq 0) {
            throw "$clientFile network-call checks no longer match the production request path."
        }
        foreach ($networkCall in $networkCalls) {
            if ($networkCall.Index -lt $guard.Index) {
                throw "$clientFile can contact the network before enforcing managed TLS policy."
            }
        }
    }
    Write-Host "[OK] Nextcloud and update HTTP paths enforce the central managed TLS guard before network I/O"

    $settingsFormPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\SettingsForm.cs"
    $settingsFormSource = Get-Content -LiteralPath $settingsFormPath -Raw
    $settingsGeneralSource = Get-Content -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\SettingsForm.General.cs") -Raw
    if ($settingsFormSource -notmatch [regex]::Escape("_saveButton.DialogResult = DialogResult.None;")) {
        throw "Settings save button can close the form before URL validation completes."
    }
    if ($settingsFormSource -notmatch "TryNormalizeBaseUrl\(requestedServerUrl,\s*out normalizedServerUrl\)") {
        throw "Settings save path does not validate and normalize the configured Nextcloud URL."
    }
    if ($settingsGeneralSource -notmatch "TryNormalizeBaseUrl\(baseUrl,\s*out normalizedUrl\)") {
        throw "Settings connection test does not validate and normalize the configured Nextcloud URL."
    }
    Write-Host "[OK] Settings UI rejects invalid Nextcloud URLs before save or connection test"
}
finally {
    if (Test-Path $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}

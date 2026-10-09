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

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class NcHttpRequestOptions
    {
        internal string Method, Url;
        internal int TimeoutMs;
        internal bool IncludeAuthHeader, IncludeOcsApiHeader, ParseJson;
    }

    internal sealed class NcHttpResponse
    {
        internal bool HasHttpResponse;
        internal HttpStatusCode StatusCode;
        internal Exception TransportException;
        internal IDictionary<string, object> ParsedJson;
        internal IDictionary<string, string> Headers;
    }

    internal sealed class NcHttpClient
    {
        internal static NcHttpResponse NextResponse;
        internal static NcHttpRequestOptions LastRequest;
        internal NcHttpClient(TalkServiceConfiguration configuration) { }
        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            LastRequest = options;
            return NextResponse;
        }
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
        TestBackendStatusResponseValidation();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " security regression test(s) failed.");
            return 1;
        }

        Console.WriteLine("All Outlook security regression tests passed.");
        return 0;
    }

    private static void TestBackendStatusResponseValidation()
    {
        foreach (int code in new[] { 200, 404 })
        foreach (bool wrapped in new[] { false, true })
        foreach (string mode in new[] { "community", "pro" })
        {
            string context = "HTTP " + code + (wrapped ? " OCS " : " plain ") + mode;
            string payload = StatusPayload(true, true, "active", mode);
            if (wrapped) payload = "{\"ocs\":{\"data\":" + payload + "}}";
            BackendPolicyStatus active = FetchBackendStatus(code, payload);
            Check(context + " accepts valid personal status", active.FetchSucceeded && active.EndpointAvailable
                && active.SeatAssigned && active.IsValid && active.SeatState == "active" && active.PolicyActive);
            Check(context + " retains independent policy domains", active.IsDomainActive("share")
                && active.IsDomainActive("talk") && active.IsDomainActive("email_signature"));
            Check(context + " does not require optional license metadata", active.LicenseStatus == string.Empty
                && active.AccessStatus == string.Empty && !active.CanManageLicense && active.DefaultsSource == null);

            BackendPolicyStatus unassigned = FetchBackendStatus(code, StatusPayload(false, true, "none", mode));
            Check(context + " explicit missing seat remains a successful refusal", unassigned.FetchSucceeded
                && unassigned.EndpointAvailable && !unassigned.SeatAssigned && unassigned.IsValid && !unassigned.PolicyActive);
            BackendPolicyStatus paused = FetchBackendStatus(code, StatusPayload(true, true, "suspended_overlimit", mode));
            Check(context + " personal suspension remains a successful refusal", paused.FetchSucceeded
                && paused.SeatAssigned && paused.IsValid && !paused.PolicyActive);
            BackendPolicyStatus invalidLicense = FetchBackendStatus(code, StatusPayload(true, false, "active", mode));
            Check(context + " invalid access remains a successful refusal", invalidLicense.FetchSucceeded
                && invalidLicense.SeatAssigned && !invalidLicense.IsValid && !invalidLicense.PolicyActive);
        }

        foreach (int code in new[] { 200, 404 })
        {
            BackendPolicyStatus oldBackend = FetchBackendStatus(code,
                "{\"status\":{\"seat_assigned\":true,\"is_valid\":true,\"seat_state\":\"active\"}}");
            Check("HTTP " + code + " accepts older backend without policy domains", oldBackend.FetchSucceeded
                && oldBackend.SeatAssigned && oldBackend.IsValid && !oldBackend.PolicyActive
                && oldBackend.Reason == "policy_domains_unavailable");
            BackendPolicyStatus unknownSeat = FetchBackendStatus(code, StatusPayload(true, true, "future_state", "community"));
            Check("HTTP " + code + " unknown seat state stays confirmed but cannot grant access", unknownSeat.FetchSucceeded
                && unknownSeat.SeatState == "future_state" && !unknownSeat.PolicyActive);

            foreach (string field in new[] { "seat_assigned", "is_valid", "seat_state" })
            {
                string[] invalidValues = field == "seat_state"
                    ? new[] { "null", "true", "1", "{}", "[]", "\"\"", "\"   \"" }
                    : new[] { "null", "\"true\"", "\"false\"", "1", "0", "{}", "[]" };
                foreach (string value in invalidValues)
                {
                    IDictionary<string, object> payload = NcJson.DeserializeObject(StatusPayload(true, true, "active", "community"));
                    NcJson.GetDictionary(payload, "status")[field] = NcJson.DeserializeObject("{\"value\":" + value + "}")["value"];
                    CheckInvalidBackendStatus("HTTP " + code + " invalid " + field + "=" + value,
                        FetchBackendStatus(code, NcJson.Serialize(payload)));
                }
                IDictionary<string, object> missing = NcJson.DeserializeObject(StatusPayload(true, true, "active", "pro"));
                NcJson.GetDictionary(missing, "status").Remove(field);
                CheckInvalidBackendStatus("HTTP " + code + " missing " + field,
                    FetchBackendStatus(code, NcJson.Serialize(missing)));
            }

            foreach (string statusValue in new[] { "{}", "null", "[]", "\"broken\"" })
            foreach (bool wrapped in new[] { false, true })
            {
                string invalid = "{\"status\":" + statusValue + "}";
                if (wrapped) invalid = "{\"ocs\":{\"data\":" + invalid + "}}";
                CheckInvalidBackendStatus("HTTP " + code + " invalid status " + statusValue + " wrapped=" + wrapped,
                    FetchBackendStatus(code, invalid));
            }
        }

        foreach (string body in new[] { "", "<html>Not found</html>", "{\"status\":", "{}", "{\"ocs\":{\"data\":{}}}" })
        {
            CheckInvalidBackendStatus("Successful HTTP requires a complete status payload", FetchBackendStatus(200, body));
            BackendPolicyStatus missing = FetchBackendStatus(404, body);
            Check("Ordinary HTTP 404 identifies the missing endpoint, not a confirmed policy or Seat refusal", !missing.FetchSucceeded
                && !missing.EndpointAvailable && !missing.PolicyActive && !missing.SeatAssigned
                && missing.Reason == "backend_missing" && missing.IsEndpointMissing && missing.IsServiceUnavailable);
        }

        foreach (int code in new[] { 301, 400, 401, 403, 409, 429, 500, 502, 503, 504, 507 })
        {
            BackendPolicyStatus rejected = FetchBackendStatus(code, StatusPayload(true, true, "active", "pro"));
            string expectedReason = code == 401 ? "authentication_rejected"
                : code == 429 ? "rate_limited"
                : code == 500 || code == 502 || code == 503 || code == 504 ? "backend_unavailable" : "check_failed";
            bool outage = code == 429 || code == 500 || code == 502 || code == 503 || code == 504;
            Check("HTTP " + code + " cannot be rescued by a valid-looking body", !rejected.FetchSucceeded
                && rejected.EndpointAvailable && !rejected.PolicyActive && !rejected.SeatAssigned
                && rejected.Reason == expectedReason && rejected.IsServiceUnavailable == outage);
            if (code == 429) Check("Rate limit without a usable server delay uses a bounded fallback", rejected.RetryAfterUtc > DateTime.UtcNow
                && rejected.RetryAfterUtc <= DateTime.UtcNow.AddMinutes(1));
        }
        NcHttpClient.NextResponse = new NcHttpResponse { TransportException = new IOException("Simulated offline transport") };
        BackendPolicyStatus offline = new BackendPolicyService(new TalkServiceConfiguration("https://cloud.example.test", "test-user", "test-only")).FetchStatus();
        Check("Transport failure is not a confirmed backend absence or refusal", !offline.FetchSucceeded
            && offline.EndpointAvailable && !offline.PolicyActive && !offline.SeatAssigned
            && offline.Reason == "nextcloud_unavailable" && offline.IsServiceUnavailable);

        DateTime before = DateTime.UtcNow;
        BackendPolicyStatus delayed = FetchBackendStatus(429, "{}", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "Retry-After", "120" } });
        Check("HTTP 429 honors server Retry-After seconds", delayed.Reason == "rate_limited" && delayed.IsServiceUnavailable
            && delayed.RetryAfterUtc >= before.AddSeconds(120) && delayed.RetryAfterUtc <= DateTime.UtcNow.AddSeconds(120));
        DateTime retryAt = DateTime.UtcNow.AddMinutes(5);
        retryAt = new DateTime(retryAt.Ticks - retryAt.Ticks % TimeSpan.TicksPerSecond, DateTimeKind.Utc);
        delayed = FetchBackendStatus(429, "{}", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "retry-after", retryAt.ToString("R", System.Globalization.CultureInfo.InvariantCulture) } });
        Check("HTTP 429 honors Retry-After HTTP dates", delayed.RetryAfterUtc == retryAt);
        foreach (string retryValue in new[] { "not-a-delay", "-1", "999999999999999999999999" })
        {
            before = DateTime.UtcNow;
            delayed = FetchBackendStatus(429, "{}", new Dictionary<string, string> { { "Retry-After", retryValue } });
            Check("Malformed Retry-After uses the fallback without accepting a negative delay", delayed.RetryAfterUtc >= before.AddMinutes(1)
                && delayed.RetryAfterUtc <= DateTime.UtcNow.AddMinutes(1));
        }
    }

    private static string StatusPayload(bool assigned, bool valid, string seatState, string mode)
    {
        return "{\"status\":{\"seat_assigned\":" + (assigned ? "true" : "false")
            + ",\"is_valid\":" + (valid ? "true" : "false") + ",\"seat_state\":\"" + seatState
            + "\",\"mode\":\"" + mode + "\",\"overlicensed\":true},"
            + "\"policy\":{\"share\":{},\"talk\":{},\"email_signature\":{}},"
            + "\"policy_editable\":{\"share\":{},\"talk\":{},\"email_signature\":{}}}";
    }

    private static BackendPolicyStatus FetchBackendStatus(int code, string body, IDictionary<string, string> headers = null)
    {
        IDictionary<string, object> parsed = null;
        try { parsed = NcJson.DeserializeObject(body); }
        catch (ArgumentException) { }
        NcHttpClient.NextResponse = new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = (HttpStatusCode)code,
            ParsedJson = parsed,
            Headers = headers
        };
        BackendPolicyStatus status = new BackendPolicyService(new TalkServiceConfiguration(
            "https://cloud.example.test/nextcloud", "test-user", "test-only")).FetchStatus();
        Check("Backend request retains authenticated status endpoint", NcHttpClient.LastRequest.Method == "GET"
            && NcHttpClient.LastRequest.Url == "https://cloud.example.test/nextcloud/apps/ncc_backend_4mc/api/v1/status"
            && NcHttpClient.LastRequest.IncludeAuthHeader && NcHttpClient.LastRequest.IncludeOcsApiHeader
            && NcHttpClient.LastRequest.ParseJson && NcHttpClient.LastRequest.TimeoutMs == 45000);
        return status;
    }

    private static void CheckInvalidBackendStatus(string name, BackendPolicyStatus status)
    {
        Check(name, !status.FetchSucceeded && status.EndpointAvailable && !status.PolicyActive
            && !status.SeatAssigned && status.Reason == "invalid_payload" && !status.IsServiceUnavailable);
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
            using (transaction.AcquireLock())
            {
                transaction.Commit(stream => Write(stream, "valid:preferences-without-credentials"), healthy, true);
            }
            Check("Credential removal writes cleared primary", File.ReadAllText(primary) == "valid:preferences-without-credentials");
            Check("Credential removal replaces the recoverable backup", File.ReadAllText(transaction.BackupPath) == "valid:preferences-without-credentials");
            File.WriteAllText(primary, "corrupt-after-removal", Encoding.UTF8);
            using (transaction.AcquireLock())
            {
                Check("Cleared backup still supports recovery", transaction.TryRestorePrimaryFromBackup(healthy));
            }
            Check("Backup recovery cannot restore removed credentials", File.ReadAllText(primary) == "valid:preferences-without-credentials");
            bool invalidRemovalRejected = false;
            try
            {
                using (transaction.AcquireLock())
                {
                    transaction.Commit(stream => Write(stream, "invalid-removal"), healthy, true);
                }
            }
            catch (InvalidDataException) { invalidRemovalRejected = true; }
            Check("Invalid removal leaves valid settings and backup intact", invalidRemovalRejected
                && File.ReadAllText(primary) == "valid:preferences-without-credentials"
                && File.ReadAllText(transaction.BackupPath) == "valid:preferences-without-credentials");
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
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\BackendPolicyStatus.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\UpdateCheckResult.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\BackendPolicyService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceConfiguration.cs"),
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

    $httpPauseSource = Join-Path $TempRoot "HttpRequestPauseTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
namespace NcTalkOutlookAddIn.Settings {
    internal sealed class AddinSettings {
        internal bool TransportTlsUseSystemDefault { get; set; }
        internal bool TransportTlsEnable12 { get; set; }
        internal bool TransportTlsEnable13 { get; set; }
        internal bool HasManagedTransportTls { get; set; }
        internal bool IsManagedTransportTlsValid { get { return true; } }
    }
}
namespace NcTalkOutlookAddIn.Utilities {
    internal static class AppDataPaths { internal static string EnsureLocalRootDirectory() { throw new InvalidOperationException("Use the isolated fixture path."); } }
    internal static class LogCategories { internal const string Api = "api", Core = "core"; }
    internal static class DiagnosticsLogger {
        internal static void Log(string category, string message) { }
        internal static void LogException(string category, string message, Exception ex) { }
    }
    internal static class Strings {
        internal const string ConnectionFailureCertificateSummary = "certificate", ConnectionFailureCertificateGuidance = "certificate";
        internal const string ConnectionFailureDnsSummary = "dns", ConnectionFailureDnsGuidance = "dns";
        internal const string ConnectionFailureProxySummary = "proxy", ConnectionFailureProxyGuidance = "proxy";
        internal const string ConnectionFailureTimeoutSummary = "timeout", ConnectionFailureTimeoutGuidance = "timeout";
        internal const string ConnectionFailureTlsSummary = "tls", ConnectionFailureTlsGuidance = "tls";
        internal const string ConnectionFailureGenericSummary = "generic", ConnectionFailureGenericGuidance = "generic";
        internal const string ManagedTlsPolicyInvalid = "invalid tls";
        internal const string ErrorCredentialsNotVerified = "credentials not verified", ConnectionAuthRequired = "sign in again", ConnectionRateLimited = "wait and retry";
    }
}
namespace NcTalkOutlookAddIn.Services {
    internal static class NextcloudCapabilitiesService { internal static void ClearCache() { } }
    internal static class NextcloudUserIdentityService {
        internal static void ClearCache() { }
        internal static string ResolveCurrentUserId(TalkServiceConfiguration configuration, bool forceRefresh) {
            NextcloudConnectionState.AssertRequestAllowed(configuration);
            NextcloudConnectionState.RecordVerifiedIdentity(configuration, "canonical-user", 0, false);
            return "canonical-user";
        }
    }
}
internal static class HttpRequestPauseTests {
    private static int failures, requests;
    private static volatile int statusCode = 200;
    private static volatile string retryAfter = "";
    private static void Check(string name, bool result) {
        Console.WriteLine((result ? "[OK] " : "[FAIL] ") + name);
        if (!result) failures++;
    }
    private static NcHttpRequestOptions Options(string url, bool verify = false) {
        return new NcHttpRequestOptions { Url = url, TimeoutMs = 2000, ParseJson = false,
            VerifyRejectedCredentials = verify, ForceFreshConnection = true };
    }
    public static int Main() {
        string stateDirectory = Path.Combine(Path.GetTempPath(), "nc4ol-pause-state-" + Guid.NewGuid().ToString("N"));
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        string url = "http://127.0.0.1:" + ((IPEndPoint)listener.LocalEndpoint).Port + "/pause-fixture";
        var server = new Thread(() => {
            while (true) {
                TcpClient client;
                try { client = listener.AcceptTcpClient(); }
                catch (SocketException) { break; }
                catch (ObjectDisposedException) { break; }
                catch (InvalidOperationException) { break; }
                using (client) using (NetworkStream stream = client.GetStream()) {
                    var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, true);
                    while (!string.IsNullOrEmpty(reader.ReadLine())) { }
                    Interlocked.Increment(ref requests);
                    string headers = string.IsNullOrEmpty(retryAfter) ? "" : "Retry-After: " + retryAfter + "\r\n";
                    byte[] bytes = Encoding.ASCII.GetBytes("HTTP/1.1 " + statusCode + " Fixture\r\nContent-Length: 0\r\nConnection: close\r\n" + headers + "\r\n");
                    stream.Write(bytes, 0, bytes.Length);
                }
            }
        });
        server.IsBackground = true;
        server.Start();
        try {
            var configuration = new TalkServiceConfiguration("https://pause.example.test/nextcloud", "fixture-user", "fixture-password");
            NextcloudConnectionState.Initialize(stateDirectory, "fixture-profile", configuration);
            int restored = 0;
            NextcloudConnectionState.ConnectionRestored += identity => { restored++; };
            var client = new NcHttpClient(configuration);
            statusCode = 401;
            NcHttpResponse rejected = client.Send(Options(url));
            Check("An actual HTTP 401 records the rejected authentication", rejected.StatusCode == HttpStatusCode.Unauthorized && requests == 1);
            statusCode = 200;
            int before = requests;
            for (int i = 0; i < 4; i++) {
                NcHttpResponse paused = new NcHttpClient(configuration).Send(Options(url));
                Check("New HTTP client instances cannot retry rejected credentials", paused.StatusCode == HttpStatusCode.Unauthorized && paused.RequestSequence == 0);
            }
            Check("Fresh connections do not bypass the authentication pause", requests == before);
            NcHttpResponse verified = client.Send(Options(url, true));
            Check("Explicit verification is permitted for rejected credentials", verified.StatusCode == HttpStatusCode.OK && requests == before + 1);
            Check("An HTTP success alone cannot resume unverified authentication", client.Send(Options(url)).StatusCode == HttpStatusCode.Unauthorized && requests == before + 1);
            NextcloudConnectionState.RecordVerifiedIdentity(configuration, "canonical-user", verified.RequestSequence, true);
            Check("A successful draft test records proof but does not resume requests", NextcloudConnectionState.HasFreshVerification(configuration)
                && client.Send(Options(url)).StatusCode == HttpStatusCode.Unauthorized && restored == 0);
            var proofsField = typeof(NextcloudConnectionState).GetField("Proofs", BindingFlags.Static | BindingFlags.NonPublic);
            var proofs = (System.Collections.IDictionary)proofsField.GetValue(null);
            foreach (System.Collections.DictionaryEntry entry in proofs) {
                entry.Value.GetType().GetField("VerifiedAtUtc", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(entry.Value, DateTime.UtcNow.AddMinutes(-6));
            }
            Check("Expired verification proof cannot authorize a save", !NextcloudConnectionState.HasFreshVerification(configuration));
            NextcloudConnectionState.RecordVerifiedIdentity(configuration, "canonical-user", verified.RequestSequence, true);
            NextcloudConnectionState.CommitVerifiedSavedConnection(configuration);
            Check("Verified saved authentication resumes matching background requests", client.Send(Options(url)).StatusCode == HttpStatusCode.OK && requests == before + 2 && restored == 1);
            Check("Committed identity retains the canonical account", NextcloudConnectionState.GetVerifiedIdentity().UserId == "canonical-user");

            NextcloudConnectionState.RecordResponse(configuration, url, true, rejected);
            Check("A late old rejection cannot undo a newer verified sign-in", client.Send(Options(url)).StatusCode == HttpStatusCode.OK);
            var badDraft = new TalkServiceConfiguration(configuration.BaseUrl, configuration.Username, "bad-draft-password");
            statusCode = 401;
            new NcHttpClient(badDraft).Send(Options(url, true));
            Check("A rejected draft cannot pause the working saved connection", !NextcloudConnectionState.GetStatus().IsPaused);
            rejected = client.Send(Options(url));
            Check("Rejection invalidates an older verification proof", !NextcloudConnectionState.HasFreshVerification(configuration));
            before = requests;
            NextcloudConnectionState.Shutdown();
            NextcloudConnectionState.Initialize(stateDirectory, "fixture-profile", configuration);
            Check("Rejected authentication stays paused after restart", client.Send(Options(url)).StatusCode == HttpStatusCode.Unauthorized && requests == before);
            Check("Known cleanup ownership survives a paused restart", NextcloudConnectionState.GetKnownIdentity(configuration).UserId == "canonical-user"
                && NextcloudConnectionState.GetVerifiedIdentity() == null);
            statusCode = 200;
            var updatedConfiguration = new TalkServiceConfiguration(configuration.BaseUrl, configuration.Username, "new-fixture-password");
            var updatedClient = new NcHttpClient(updatedConfiguration);
            verified = updatedClient.Send(Options(url, true));
            Check("Changed credentials can be checked without clearing old rejected credentials", verified.StatusCode == HttpStatusCode.OK
                && client.Send(Options(url)).StatusCode == HttpStatusCode.Unauthorized);
            NextcloudConnectionState.RecordVerifiedIdentity(updatedConfiguration, "canonical-user", verified.RequestSequence, true);
            NextcloudConnectionState.BeginVerification(updatedConfiguration);
            Check("Starting a new test invalidates any previous successful proof", !NextcloudConnectionState.HasFreshVerification(updatedConfiguration));
            NextcloudConnectionState.RecordVerifiedIdentity(updatedConfiguration, "canonical-user", verified.RequestSequence, true);
            NextcloudConnectionState.CommitVerifiedSavedConnection(updatedConfiguration);
            before = requests;
            Check("A delayed worker cannot replay a replaced password", client.Send(Options(url)).StatusCode == HttpStatusCode.Unauthorized && requests == before);
            Check("Saved replacement credentials resume ordinary requests", updatedClient.Send(Options(url)).StatusCode == HttpStatusCode.OK);
            NextcloudConnectionState.Shutdown();
            NextcloudConnectionState.Initialize(stateDirectory, "fixture-profile", updatedConfiguration);
            Check("Committed recovery remains usable after restart", updatedClient.Send(Options(url)).StatusCode == HttpStatusCode.OK);
            var storeField = typeof(NextcloudConnectionState).GetField("_store", BindingFlags.Static | BindingFlags.NonPublic);
            object store = storeField.GetValue(null);
            var writable = store.GetType().GetField("_writeAllowed", BindingFlags.Instance | BindingFlags.NonPublic);
            writable.SetValue(store, false);
            statusCode = 401;
            NcHttpResponse writeFailure = updatedClient.Send(Options(url));
            Check("State write failure retains the readable HTTP 401 and runtime pause", writeFailure.StatusCode == HttpStatusCode.Unauthorized
                && NextcloudConnectionState.GetStatus().Reason == "auth_required");
            statusCode = 200;
            verified = updatedClient.Send(Options(url, true));
            NextcloudConnectionState.RecordVerifiedIdentity(updatedConfiguration, "canonical-user", verified.RequestSequence, true);
            bool failedCommit = false;
            int restoredBeforeFailedCommit = restored;
            try { NextcloudConnectionState.CommitVerifiedSavedConnection(updatedConfiguration); }
            catch (InvalidOperationException) { failedCommit = true; }
            Check("Failed persistent commit cannot clear a pause or emit recovery", failedCommit && restored == restoredBeforeFailedCommit
                && NextcloudConnectionState.GetStatus().Reason == "auth_required");
            writable.SetValue(store, true);
            NextcloudConnectionState.CommitVerifiedSavedConnection(updatedConfiguration);

            statusCode = 429;
            retryAfter = "120";
            NcHttpResponse limited = updatedClient.Send(Options(url));
            DateTime firstDeadline = HttpFailureDiagnostics.ReadRetryAfterUtc(limited.Headers);
            Check("An actual HTTP 429 retains its server delay", (int)limited.StatusCode == 429 && firstDeadline > DateTime.UtcNow.AddSeconds(115));
            before = requests;
            statusCode = 200;
            retryAfter = "";
            var anotherConfiguration = new TalkServiceConfiguration(configuration.BaseUrl, "another-user", "another-password");
            NcHttpResponse delayed = new NcHttpClient(anotherConfiguration).Send(Options(url, true));
            Check("Changed user, password and explicit verification cannot bypass server backoff", (int)delayed.StatusCode == 429 && requests == before);
            DateTime retainedDeadline = HttpFailureDiagnostics.ReadRetryAfterUtc(delayed.Headers);
            Check("Locally paused requests retain rather than extend the deadline", Math.Abs((retainedDeadline - firstDeadline).TotalSeconds) < 1.1);
            var anonymous = Options(url);
            anonymous.IncludeAuthHeader = false;
            Check("Login flow also honors the server retry deadline", (int)updatedClient.Send(anonymous).StatusCode == 429 && requests == before);
            NextcloudConnectionState.Shutdown();
            NextcloudConnectionState.Initialize(stateDirectory, "fixture-profile", updatedConfiguration);
            Check("Server retry deadlines survive restart", (int)updatedClient.Send(anonymous).StatusCode == 429 && requests == before);
            Check("Another origin is not rate-limited", !NextcloudConnectionState.GetRequestStatus(updatedConfiguration,
                "https://other.example.test/nextcloud", true, false).IsPaused);
            var emptyConfiguration = new TalkServiceConfiguration(updatedConfiguration.BaseUrl, "", "");
            NextcloudConnectionState.RemoveSavedCredentials(emptyConfiguration);
            Check("Removing credentials does not reset server backoff", (int)updatedClient.Send(anonymous).StatusCode == 429 && requests == before);
            NextcloudConnectionState.Shutdown();
            NextcloudConnectionState.Initialize(stateDirectory, "fixture-profile", emptyConfiguration);
            Check("Removed credentials and retained backoff stay effective after restart", (int)updatedClient.Send(anonymous).StatusCode == 429 && requests == before);

            var stateField = typeof(NextcloudConnectionState).GetField("_state", BindingFlags.Static | BindingFlags.NonPublic);
            var state = (PersistedNextcloudConnectionState)stateField.GetValue(null);
            state.RetryAfterUtc = DateTime.UtcNow.AddSeconds(-1);
            Check("Expired server backoff permits explicit sign-in again", updatedClient.Send(Options(url, true)).StatusCode == HttpStatusCode.OK);
            Check("Expiry of server backoff does not authorize stale credentials", client.Send(Options(url)).StatusCode == HttpStatusCode.Unauthorized);
            Check("Unauthenticated login flow requests do not reuse rejected authentication", client.Send(anonymous).StatusCode == HttpStatusCode.OK);
            before = requests;
            Check("Removing credentials blocks delayed workers without network requests", updatedClient.Send(Options(url)).StatusCode == HttpStatusCode.Unauthorized && requests == before);
            Check("Removing credentials retains known cleanup ownership", NextcloudConnectionState.GetKnownIdentity(updatedConfiguration).UserId == "canonical-user");
            Check("Connection-state persistence contains no clear-text credentials", !File.ReadAllText(Directory.GetFiles(stateDirectory, "connection-state-*.dat")[0]).Contains("fixture-password"));
        }
        finally { NextcloudConnectionState.Shutdown(); listener.Stop(); server.Join(2000); Directory.Delete(stateDirectory, true); }
        return failures == 0 ? 0 : 1;
    }
}
'@ | Set-Content -LiteralPath $httpPauseSource -Encoding UTF8
    $httpPauseSources = @(
        $httpPauseSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\NcHttpClient.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\NextcloudConnectionState.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\ProtectedJsonStateStore.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\DurableFileReplace.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceException.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceConfiguration.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HttpFailureDiagnostics.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HttpAuthUtilities.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudUriValidator.cs")
    )
    $httpPauseExe = Join-Path $TempRoot "HttpRequestPauseTests.exe"
    & $csc /nologo /target:exe "/out:$httpPauseExe" @references @httpPauseSources
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & $httpPauseExe
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    $policyRefreshSource = Get-Content -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.PolicyTemplates.cs") -Raw
    $policyRefreshEnd = $policyRefreshSource.IndexOf('        internal PasswordPolicyInfo FetchPasswordPolicyForTalkWizard', [StringComparison]::Ordinal)
    if ($policyRefreshEnd -lt 0) { throw 'Production policy refresh section was not found.' }
    $policyRefreshSource = $policyRefreshSource.Substring(0, $policyRefreshEnd) + @'
        private static void LogCore(string message) { }
    }
}
namespace NcTalkOutlookAddIn.Controllers { }
namespace NcTalkOutlookAddIn.Utilities {
    internal static class LogCategories { internal const string Core = "core"; }
    internal static class DiagnosticsLogger { internal static void LogException(string category, string message, Exception ex) { } }
}
namespace NcTalkOutlookAddIn.Services {
    internal sealed class TalkServiceConfiguration {
        internal string Username = "fixture-user", AppPassword = "fixture-password";
        internal string GetNormalizedBaseUrl() { return "https://refresh.example.test/nextcloud"; }
    }
    internal sealed class BackendPolicyService {
        internal static BackendPolicyStatus Next;
        internal static int Calls;
        internal BackendPolicyService(TalkServiceConfiguration configuration) { }
        internal BackendPolicyStatus FetchStatus() { Calls++; return Next; }
    }
}
internal static class PolicyRefreshPauseTests {
    private static BackendPolicyStatus Status(bool success, string reason) {
        return new BackendPolicyStatus(true, success, success, "community", reason, success, success,
            success ? "active" : "", null, null, null, null, null, null);
    }
    public static int Main() {
        foreach (bool warm in new[] { false, true }) {
            foreach (string reason in new[] { "authentication_rejected", "rate_limited", "nextcloud_unavailable", "check_failed" }) {
                var owner = new NcTalkOutlookAddIn.NextcloudTalkAddIn();
                var configuration = new TalkServiceConfiguration();
                BackendPolicyStatus confirmed = Status(true, "active");
                if (warm) {
                    BackendPolicyService.Next = confirmed;
                    owner.FetchBackendPolicyStatus(configuration, "fixture_warm");
                }
                BackendPolicyService.Next = Status(false, reason);
                if (reason == "rate_limited") BackendPolicyService.Next.RetryAfterUtc = DateTime.UtcNow.AddMinutes(10);
                owner.FetchBackendPolicyStatus(configuration, "fixture_failure");
                int before = BackendPolicyService.Calls;
                for (int i = 0; i < 4; i++) {
                    BackendPolicyStatus resolved = owner.FetchEnterpriseRolloutPolicyStatus(configuration, "fixture_background");
                    if (warm ? !object.ReferenceEquals(resolved, confirmed) : resolved.FetchSucceeded) {
                        throw new Exception("Failed refresh lost known policy or invented unknown policy.");
                    }
                }
                if (BackendPolicyService.Calls != before) throw new Exception("Enterprise refresh bypassed failed-check cooldown: " + reason);
                Console.WriteLine("[OK] Enterprise background refresh honors " + reason + " cooldown; known policy=" + warm);
                var checkedAt = typeof(NcTalkOutlookAddIn.NextcloudTalkAddIn).GetField("_backendPolicyLastCheckedAtUtc",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                checkedAt.SetValue(owner, DateTime.UtcNow.AddMinutes(-1));
                owner.FetchEnterpriseRolloutPolicyStatus(configuration, "fixture_after_cooldown");
                int expectedCalls = before + (reason == "rate_limited" ? 0 : 1);
                if (BackendPolicyService.Calls != expectedCalls) throw new Exception("Retry-After and failed-check expiry differ from their deadlines: " + reason);
                BackendPolicyService.Next = Status(true, "verified");
                owner.FetchBackendPolicyStatus(configuration, "fixture_explicit_verified_refresh");
                before = BackendPolicyService.Calls;
                owner.FetchEnterpriseRolloutPolicyStatus(configuration, "fixture_background_resumed");
                if (BackendPolicyService.Calls != before + 1) throw new Exception("Successful refresh unexpectedly cached the enterprise operational check.");
            }
        }
        Console.WriteLine("[OK] Verified refresh resumes the existing fresh enterprise access checks");
        return 0;
    }
}
'@
    $policyRefreshPath = Join-Path $TempRoot "PolicyRefreshPauseTests.cs"
    Set-Content -LiteralPath $policyRefreshPath -Value $policyRefreshSource -Encoding UTF8
    $policyRefreshExe = Join-Path $TempRoot "PolicyRefreshPauseTests.exe"
    & $csc /nologo /target:exe "/out:$policyRefreshExe" @references $policyRefreshPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\BackendPolicyStatus.cs")
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & $policyRefreshExe
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

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

    $lifecycleSource = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\NcTalkOutlookAddIn\NextcloudTalkAddIn.Lifecycle.cs') -Raw
    $startupTls = $lifecycleSource.IndexOf('TryApplyTransportSecurityFromSettings("startup", false)', [StringComparison]::Ordinal)
    $connectionStateInit = $lifecycleSource.IndexOf('NextcloudConnectionState.Initialize(', [StringComparison]::Ordinal)
    $shareCleanupInit = $lifecycleSource.IndexOf('InitializeComposeShareCleanup(', [StringComparison]::Ordinal)
    if ($startupTls -lt 0 -or $connectionStateInit -le $startupTls -or $shareCleanupInit -le $connectionStateInit) {
        throw 'Managed TLS must be applied before connection-recovery timers and cleanup workers start.'
    }
    Write-Host '[OK] Startup applies managed TLS before connection and cleanup recovery can start network requests'

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

Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-policy-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $talkFormSource = Get-Content -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\TalkLinkForm.cs") -Raw
    if ($talkFormSource -notmatch '(?s)protected override void OnShown\(EventArgs e\)\s*\{[^}]*ApplyDialogLayout\(true\);') {
        throw "Talk must lay out visible warning panels when the dialog is first shown."
    }
    $fileLinkLayoutSource = Get-Content -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.Layout.cs") -Raw
    if ($fileLinkLayoutSource -notmatch '(?s)private void ReflowWizardLayout\(\)\s*\{[^}]*LayoutPolicyWarningPanel\(\);[^}]*UpdateStepHostBounds\(\);') {
        throw "FileLink must lay out its visible warning before positioning the step host."
    }
    Write-Host "[OK] Talk and FileLink reflow their warning panels on first display."

    $testSource = Join-Path $TempRoot "OutlookPolicyMappingTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Threading;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    internal sealed class AddinSettings
    {
        internal bool? EmailSignatureOnCompose { get; set; }
        internal bool? EmailSignatureOnReply { get; set; }
        internal bool? EmailSignatureOnForward { get; set; }
    }
}

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class TalkServiceConfiguration
    {
        internal bool IsComplete() { return true; }
        internal string GetNormalizedBaseUrl() { return "https://cloud.example.test/nextcloud"; }
    }

    internal sealed class NcHttpRequestOptions
    {
        internal string Method { get; set; }
        internal string Url { get; set; }
        internal int TimeoutMs { get; set; }
        internal bool IncludeAuthHeader { get; set; }
        internal bool IncludeOcsApiHeader { get; set; }
        internal bool ParseJson { get; set; }
    }

    internal sealed class NcHttpResponse
    {
        internal bool HasHttpResponse { get; set; }
        internal HttpStatusCode StatusCode { get; set; }
        internal Exception TransportException { get; set; }
        internal IDictionary<string, object> ParsedJson { get; set; }
    }

    internal sealed class NcHttpClient
    {
        internal static NcHttpResponse NextResponse;
        internal NcHttpClient(TalkServiceConfiguration configuration) { }
        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            if (NextResponse == null)
            {
                throw new InvalidOperationException("No HTTP response configured for policy test.");
            }
            return NextResponse;
        }
    }
}

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static bool IsEnabled { get { return false; } }
        internal static void Log(string category, string message) { }
        internal static void LogException(string category, string message, Exception ex) { }
    }

    internal static class LogCategories
    {
        internal const string Core = "CORE";
    }

    internal static class BrowserLauncher
    {
        internal static string LastUrl;
        internal static void OpenUrl(string url, string category, string message) { LastUrl = url; }
    }
}

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

    [STAThread]
    public static int Main()
    {
        Strings.SetPreferredUiLanguage("en");
        TestPasswordDeliveryMode();
        TestSecretsExpireDays();
        TestLockedBackendValueWins();
        TestEditableBackendValueKeepsLocalChoice();
        TestAttachmentLinkTargetMapping();
        TestAttachmentLinkTargetPrecedence();
        TestEditableSignatureDefaultCanBeEnabledLocally();
        TestLockedSignatureDefaultCannotBeEnabledLocally();
        TestSignaturePolicyAvailability();
        TestLicenseStatusParsing();
        TestLicenseNotices();
        TestLicenseMetadataDoesNotGrantAccess();
        TestLicenseSeatAndConnectionPrecedence();
        TestSuspendedSeatNoticePrecedence();
        TestLicenseAdminLinks();
        TestLicenseWarningRendering();
        TestWarningPanelLayout();

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
        Check("Locked backend policy uses secrets", policy.Mode == SharePasswordDeliveryMode.Secrets);
    }

    private static void TestEditableBackendValueKeepsLocalChoice()
    {
        var status = BuildStatus("secrets", true, "30");
        SharePasswordDeliveryPolicy policy = SharePasswordDeliveryPolicy.Resolve(status, SharePasswordDeliveryMode.Plain);
        Check("Editable backend delivery mode keeps local choice", policy.Mode == SharePasswordDeliveryMode.Plain, policy.Mode.ToString());
        Check("Editable backend still supplies shared expire days", policy.SecretsExpireDays == 30, policy.SecretsExpireDays.ToString());
    }

    private static void TestAttachmentLinkTargetMapping()
    {
        AttachmentLinkTarget parsedTarget;
        Check(
            "Attachment target parses ZIP",
            AttachmentLinkTargetPolicy.TryParse("zip_download", out parsedTarget)
            && parsedTarget == AttachmentLinkTarget.ZipDownload);
        Check(
            "Attachment target parses share page",
            AttachmentLinkTargetPolicy.TryParse("share_page", out parsedTarget)
            && parsedTarget == AttachmentLinkTarget.SharePage);
        Check(
            "Attachment target rejects unknown values and initializes ZIP",
            !AttachmentLinkTargetPolicy.TryParse("bad", out parsedTarget)
            && parsedTarget == AttachmentLinkTarget.ZipDownload);
        Check("Attachment target serializes ZIP", AttachmentLinkTargetPolicy.ToStorageValue(AttachmentLinkTarget.ZipDownload) == "zip_download");
        Check("Attachment target serializes share page", AttachmentLinkTargetPolicy.ToStorageValue(AttachmentLinkTarget.SharePage) == "share_page");
        Check("Missing attachment target defaults to ZIP", AttachmentLinkTargetPolicy.Resolve(null, null) == AttachmentLinkTarget.ZipDownload);

        AttachmentLinkTarget parsedInvalid;
        AttachmentLinkTarget? invalidLocalValue = AttachmentLinkTargetPolicy.TryParse("bad", out parsedInvalid)
            ? parsedInvalid
            : (AttachmentLinkTarget?)null;
        Check("Invalid persisted attachment target stays unset", !invalidLocalValue.HasValue);
        Check(
            "Editable backend target seeds an invalid persisted value",
            AttachmentLinkTargetPolicy.Resolve(invalidLocalValue, BuildAttachmentTargetStatus("share_page", true)) == AttachmentLinkTarget.SharePage);
    }

    private static void TestAttachmentLinkTargetPrecedence()
    {
        BackendPolicyStatus locked = BuildAttachmentTargetStatus("share_page", false);
        BackendPolicyStatus editable = BuildAttachmentTargetStatus("share_page", true);
        BackendPolicyStatus invalidLocked = BuildAttachmentTargetStatus("bad", false);

        Check(
            "Locked backend attachment target wins",
            AttachmentLinkTargetPolicy.Resolve(AttachmentLinkTarget.ZipDownload, locked) == AttachmentLinkTarget.SharePage);
        Check(
            "Editable backend attachment target keeps explicit local choice",
            AttachmentLinkTargetPolicy.Resolve(AttachmentLinkTarget.ZipDownload, editable) == AttachmentLinkTarget.ZipDownload);
        Check(
            "Editable backend attachment target seeds absent local choice",
            AttachmentLinkTargetPolicy.Resolve(null, editable) == AttachmentLinkTarget.SharePage);
        Check(
            "Invalid locked backend attachment target fails safe to ZIP",
            AttachmentLinkTargetPolicy.Resolve(AttachmentLinkTarget.SharePage, invalidLocked) == AttachmentLinkTarget.ZipDownload);
    }

    private static void TestEditableSignatureDefaultCanBeEnabledLocally()
    {
        BackendPolicyStatus status = BuildSignatureStatus(false, true, true, true);
        var settings = new AddinSettings
        {
            EmailSignatureOnCompose = true,
            EmailSignatureOnReply = true,
            EmailSignatureOnForward = true
        };

        EmailSignaturePolicy policy = new EmailSignaturePolicyService(status, settings).Resolve();

        Check("Editable disabled signature default can be enabled locally", policy.Active && policy.OnCompose, policy.Reason);
        Check("Editable reply and forward flags keep local choices", policy.OnReply && policy.OnForward);
    }

    private static void TestLockedSignatureDefaultCannotBeEnabledLocally()
    {
        BackendPolicyStatus status = BuildSignatureStatus(false, false, true, true);
        var settings = new AddinSettings { EmailSignatureOnCompose = true };

        EmailSignaturePolicy policy = new EmailSignaturePolicyService(status, settings).Resolve();

        Check("Locked disabled signature default remains disabled", !policy.Active && !policy.OnCompose, policy.Reason);
        Check("Locked disabled signature reports backend reason", policy.Reason == "signature_disabled_by_backend", policy.Reason);
    }

    private static void TestSignaturePolicyAvailability()
    {
        BackendPolicyStatus editableDisabled = BuildSignatureStatus(false, true, true, true);
        BackendPolicyStatus lockedDisabled = BuildSignatureStatus(false, false, true, true);
        BackendPolicyStatus missingTemplate = BuildSignatureStatus(false, true, false, true);
        BackendPolicyStatus missingUserEmail = BuildSignatureStatus(false, true, true, false);

        Check(
            "Editable disabled signature policy remains configurable",
            EmailSignaturePolicyService.IsAvailableForConfiguration(editableDisabled));
        Check(
            "Locked disabled signature policy remains available for lock-state rendering",
            EmailSignaturePolicyService.IsAvailableForConfiguration(lockedDisabled));
        Check(
            "Signature policy without template is unavailable",
            !EmailSignaturePolicyService.IsAvailableForConfiguration(missingTemplate));
        Check(
            "Signature policy without user email is unavailable",
            !EmailSignaturePolicyService.IsAvailableForConfiguration(missingUserEmail));
    }

    private static void TestLicenseStatusParsing()
    {
        Dictionary<string, object> payload = LicensePayload(" GRACE ", " ACTIVE ", true, true, "active", true);
        IDictionary<string, object> fields = NcJson.GetDictionary(payload, "status");
        fields["grace_until_iso"] = "2099-04-10T12:30:00Z";
        fields["license_activation"] = new Dictionary<string, object> { { "state", " conflict " } };
        fields["license_connection_error"] = true;
        fields["license_last_sync_at_iso"] = "2026-09-01T09:00:00Z";
        fields["license_offline_until_iso"] = "2099-04-12T12:30:00Z";
        var ocsPayload = new Dictionary<string, object>
        {
            { "ocs", new Dictionary<string, object> { { "data", payload } } }
        };
        foreach (IDictionary<string, object> shape in new[] { payload, ocsPayload })
        {
            BackendPolicyStatus parsed = BackendPolicyService.ParseStatus(NcJson.DeserializeObject(NcJson.Serialize(shape)));
            Check("Root and OCS status preserve normalized license fields",
                parsed.LicenseStatus == "GRACE" && parsed.AccessStatus == "ACTIVE"
                && parsed.CanManageLicense && parsed.LicenseActivationState == "conflict");
            Check("Root and OCS status preserve informational dates and connection state",
                parsed.GraceUntilIso == "2099-04-10T12:30:00Z" && parsed.LicenseConnectionError
                && parsed.LicenseLastSyncAtIso == "2026-09-01T09:00:00Z"
                && parsed.LicenseOfflineUntilIso == "2099-04-12T12:30:00Z");
            Check("Root and OCS status retain all policy domains", parsed.IsDomainActive("share")
                && parsed.IsDomainActive("talk") && parsed.IsDomainActive("email_signature"));
        }

        foreach (object value in new object[] { true, false, "true", "TRUE", "1", 1, 1L, 1.0, null })
        {
            fields["can_manage_license"] = value;
            BackendPolicyStatus parsed = BackendPolicyService.ParseStatus(payload);
            Check("License management requires a JSON boolean: " + (value == null ? "null" : value.GetType().Name + " " + value),
                parsed.CanManageLicense == (value is bool && (bool)value));
            Check("Non-boolean permission never exposes license management", string.IsNullOrEmpty(PolicyUiHelper.GetLicenseAdminUrl(parsed, "https://cloud.example.test"))
                == !(value is bool && (bool)value));
        }
        fields.Remove("can_manage_license");
        Check("Absent license management permission stays false", !BackendPolicyService.ParseStatus(payload).CanManageLicense);

        var oldPayload = LicensePayload(null, null, true, true, "active", false);
        NcJson.GetDictionary(oldPayload, "status").Remove("can_manage_license");
        BackendPolicyStatus oldBackend = BackendPolicyService.ParseStatus(oldPayload);
        Check("Old backend without additive fields retains access", oldBackend.PolicyActive && PolicyUiHelper.HasBackendSeatEntitlement(oldBackend));
        Check("Old backend without additive fields has no license notice", PolicyUiHelper.GetPolicyWarningMessage(oldBackend) == string.Empty);
        Check("Old backend additive metadata remains empty", oldBackend.LicenseStatus == string.Empty
            && oldBackend.AccessStatus == string.Empty && oldBackend.GraceUntilIso == string.Empty
            && oldBackend.LicenseActivationState == string.Empty && !oldBackend.LicenseConnectionError
            && oldBackend.LicenseLastSyncAtIso == string.Empty && oldBackend.LicenseOfflineUntilIso == string.Empty);
        Check("Missing status fails closed", !BackendPolicyService.ParseStatus(null).PolicyActive);
    }

    private static void TestLicenseNotices()
    {
        Check("Active valid license is silent", PolicyUiHelper.GetPolicyWarningMessage(ParseLicense("ACTIVE", "ACTIVE", true)) == string.Empty);
        CheckLicenseNotice("Expired", ParseLicense("ACTIVE", "EXPIRED", false), Strings.PolicyLicenseExpired);
        CheckLicenseNotice("Inactive", ParseLicense("INACTIVE", "INACTIVE", false), Strings.PolicyLicenseInactive);
        CheckLicenseNotice("Invalid", ParseLicense("INVALID", "INVALID", false), Strings.PolicyLicenseInvalid);
        CheckLicenseNotice("Offline expired", ParseLicense("ACTIVE", "OFFLINE_EXPIRED", false), Strings.PolicyLicenseOfflineExpired);
        CheckLicenseNotice("Activation required", ParseLicense("ACTIVE", "ACTIVATION_REQUIRED", false), Strings.PolicyLicenseActivationRequired);

        var conflictPayload = LicensePayload("ACTIVE", "ACTIVATION_REQUIRED", false, true, "active", false);
        NcJson.GetDictionary(conflictPayload, "status")["license_activation"] = new Dictionary<string, object> { { "state", "conflict" } };
        CheckLicenseNotice("Activation conflict", BackendPolicyService.ParseStatus(conflictPayload), Strings.PolicyLicenseActivationConflict);
        CheckLicenseNotice("Missing access status uses license status", ParseLicense("EXPIRED", null, false), Strings.PolicyLicenseExpired);
        CheckLicenseNotice("Unknown access status does not reuse expired license status", ParseLicense("EXPIRED", "future_state", false), Strings.PolicyWarningLicenseInvalid);
        CheckLicenseNotice("Old backend invalid license stays generic", ParseLicense(null, null, false), Strings.PolicyWarningLicenseInvalid);

        var gracePayload = LicensePayload("GRACE", "GRACE", true, true, "active", false);
        BackendPolicyStatus grace = BackendPolicyService.ParseStatus(gracePayload);
        CheckLicenseNotice("GRACE without a date", grace, Strings.PolicyLicenseGrace);
        Check("GRACE keeps separate password delivery enabled", PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(grace) == string.Empty
            && PolicyUiHelper.GetPasswordDeliveryModeUnavailableTooltip(grace) == string.Empty);
        IDictionary<string, object> graceFields = NcJson.GetDictionary(gracePayload, "status");
        graceFields["grace_until_iso"] = "not-a-date";
        CheckLicenseNotice("GRACE invalid date uses localized fallback", BackendPolicyService.ParseStatus(gracePayload), Strings.PolicyLicenseGrace);

        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
        try
        {
            const string graceDate = "2099-04-10T12:30:00Z";
            graceFields["grace_until_iso"] = graceDate;
            foreach (string cultureName in new[] { "en-US", "de-DE" })
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(cultureName);
                Strings.SetPreferredUiLanguage(cultureName == "de-DE" ? "de" : "en");
                string localizedDate = DateTimeOffset.Parse(graceDate, CultureInfo.InvariantCulture).LocalDateTime.ToString("g", CultureInfo.CurrentCulture);
                string expected = string.Format(CultureInfo.CurrentCulture, Strings.PolicyLicenseGraceFormat, localizedDate);
                CheckLicenseNotice("GRACE localized date " + cultureName, BackendPolicyService.ParseStatus(gracePayload), expected);
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = originalCulture;
            Strings.SetPreferredUiLanguage("en");
        }

        graceFields["license_connection_error"] = true;
        graceFields["license_last_sync_at_iso"] = "2026-09-01T09:00:00Z";
        graceFields["license_offline_until_iso"] = "2099-04-12T12:30:00Z";
        string connectionNotice = PolicyUiHelper.GetPolicyWarningMessage(BackendPolicyService.ParseStatus(gracePayload));
        Check("GRACE keeps its notice when license synchronization fails", connectionNotice.Contains(Strings.PolicyLicenseConnectionError)
            && connectionNotice.Contains(string.Format(CultureInfo.CurrentCulture, Strings.PolicyLicenseLastSyncFormat,
                DateTimeOffset.Parse("2026-09-01T09:00:00Z", CultureInfo.InvariantCulture).LocalDateTime.ToString("g", CultureInfo.CurrentCulture)))
            && connectionNotice.Contains(string.Format(CultureInfo.CurrentCulture, Strings.PolicyLicenseOfflineUntilFormat,
                DateTimeOffset.Parse("2099-04-12T12:30:00Z", CultureInfo.InvariantCulture).LocalDateTime.ToString("g", CultureInfo.CurrentCulture))));
        var connectionPayload = LicensePayload("ACTIVE", "ACTIVE", true, true, "active", false);
        NcJson.GetDictionary(connectionPayload, "status")["license_connection_error"] = true;
        BackendPolicyStatus connectionStatus = BackendPolicyService.ParseStatus(connectionPayload);
        CheckLicenseNotice("Connection error remains informational for valid access", connectionStatus, Strings.PolicyLicenseConnectionError);
        Check("Connection error does not disable valid separate password delivery", PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(connectionStatus) == string.Empty);
    }

    private static void TestLicenseMetadataDoesNotGrantAccess()
    {
        foreach (bool assigned in new[] { false, true })
        foreach (bool valid in new[] { false, true })
        foreach (string seatState in new[] { "active", "suspended_overlimit", "pending", "" })
        foreach (string metadata in new[] { "ACTIVE", "GRACE", "EXPIRED" })
        {
            var payload = LicensePayload(metadata, metadata, valid, assigned, seatState, true);
            IDictionary<string, object> fields = NcJson.GetDictionary(payload, "status");
            fields["grace_until_iso"] = "2099-12-31T23:59:59Z";
            fields["license_last_sync_at_iso"] = "2099-12-01T00:00:00Z";
            fields["license_offline_until_iso"] = "2099-12-31T23:59:59Z";
            fields["license_connection_error"] = true;
            BackendPolicyStatus status = BackendPolicyService.ParseStatus(payload);
            bool expected = assigned && valid && seatState == "active";
            string name = "Metadata does not change gates: assigned=" + assigned + ", valid=" + valid + ", seat=" + seatState + ", license=" + metadata;
            Check(name, status.PolicyActive == expected
                && status.IsDomainActive("share") == expected
                && status.IsDomainActive("talk") == expected
                && status.IsDomainActive("email_signature") == expected
                && PolicyUiHelper.HasBackendSeatEntitlement(status) == expected
                && PolicyUiHelper.HasPasswordDeliveryMode(status) == expected
                && new EmailSignaturePolicyService(status, new AddinSettings()).Resolve().Active == expected);
        }
    }

    private static void TestLicenseSeatAndConnectionPrecedence()
    {
        BackendPolicyStatus noSeat = ParseLicense("EXPIRED", "EXPIRED", false, false);
        Check("Normal user without seat sees only existing no-seat notice", PolicyUiHelper.GetPolicyWarningMessage(noSeat) == Strings.PolicyWarningNoSeat);
        Check("Normal user without seat retains no-seat tooltip", PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(noSeat) == Strings.SharingPasswordSeparateNoSeatTooltip);
        BackendPolicyStatus adminNoSeat = ParseLicense("EXPIRED", "EXPIRED", false, false, "active", true);
        CheckLicenseNotice("Administrator without seat sees license cause", adminNoSeat, Strings.PolicyLicenseExpired);
        BackendPolicyStatus suspended = ParseLicense("ACTIVE", "ACTIVE", true, true, "suspended_overlimit");
        Check("Actual suspended seat has explicit suspended notice", PolicyUiHelper.GetPolicyWarningMessage(suspended).StartsWith(Strings.PolicyWarningSeatSuspended, StringComparison.Ordinal));
        Check("Actual suspended seat tooltip agrees with notice", PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(suspended) == PolicyUiHelper.GetPolicyWarningMessage(suspended));
        CheckLicenseNotice("Expired license on suspended seat reports license first", ParseLicense("EXPIRED", "EXPIRED", false, true, "suspended_overlimit"), Strings.PolicyLicenseExpired);
        foreach (string state in new[] { "pending", "revoked", "suspended", "", "unknown" })
        {
            BackendPolicyStatus unavailable = ParseLicense("ACTIVE", "ACTIVE", true, true, state);
            string message = PolicyUiHelper.GetPolicyWarningMessage(unavailable);
            Check("Other seat states never claim suspension: " + state, message.StartsWith(Strings.PolicyWarningSeatUnavailable, StringComparison.Ordinal)
                && !message.Contains(Strings.PolicyWarningSeatSuspended)
                && PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(unavailable) == message);
        }

        NcHttpClient.NextResponse = new NcHttpResponse { HasHttpResponse = true, StatusCode = HttpStatusCode.NotFound };
        BackendPolicyStatus missingBackend = new BackendPolicyService(new TalkServiceConfiguration()).FetchStatus();
        Check("Missing backend keeps local behavior without a license notice", !missingBackend.EndpointAvailable && missingBackend.FetchSucceeded
            && !missingBackend.PolicyActive && PolicyUiHelper.GetPolicyWarningMessage(missingBackend) == string.Empty
            && PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(missingBackend) == Strings.SharingPasswordSeparateBackendRequiredTooltip);
        Check("Null backend keeps backend-required tooltip", PolicyUiHelper.GetPolicyWarningMessage(null) == string.Empty
            && PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(null) == Strings.SharingPasswordSeparateBackendRequiredTooltip);
        NcHttpClient.NextResponse = new NcHttpResponse { TransportException = new InvalidOperationException("Offline test") };
        BackendPolicyStatus unavailableBackend = new BackendPolicyService(new TalkServiceConfiguration()).FetchStatus();
        Check("Fetch failure is not misreported as no seat or expired license", unavailableBackend.EndpointAvailable && !unavailableBackend.FetchSucceeded
            && PolicyUiHelper.GetPolicyWarningMessage(unavailableBackend) == Strings.PolicyWarningBackendUnavailable
            && PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(unavailableBackend) == Strings.PolicyWarningBackendUnavailable
            && !PolicyUiHelper.HasBackendSeatEntitlement(unavailableBackend));
        NcHttpClient.NextResponse = null;
    }

    private static void TestSuspendedSeatNoticePrecedence()
    {
        foreach (bool admin in new[] { false, true })
        foreach (string access in new[] { "GRACE", "ACTIVE" })
        foreach (bool connectionError in new[] { false, true })
        {
            var payload = LicensePayload(access, access, true, true, "suspended_overlimit", admin);
            IDictionary<string, object> fields = NcJson.GetDictionary(payload, "status");
            fields["grace_until_iso"] = "2099-12-31T23:59:59Z";
            fields["license_connection_error"] = connectionError;
            BackendPolicyStatus status = BackendPolicyService.ParseStatus(payload);
            string expected = Strings.PolicyWarningSeatSuspended + Environment.NewLine
                + (admin ? Strings.PolicyLicenseAdminHint : Strings.PolicyLicenseUserHint);
            string name = "Suspended seat takes priority: admin=" + admin + ", access=" + access + ", syncError=" + connectionError;
            Check(name, PolicyUiHelper.GetPolicyWarningMessage(status) == expected
                && PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(status) == expected
                && !PolicyUiHelper.HasBackendSeatEntitlement(status)
                && PolicyUiHelper.GetLicenseAdminUrl(status, "https://cloud.example.test") == string.Empty);
            fields["is_valid"] = false;
            fields["access_status"] = "EXPIRED";
            CheckLicenseNotice("Invalid license still precedes suspension: " + name, BackendPolicyService.ParseStatus(payload), Strings.PolicyLicenseExpired);
        }
    }

    private static void TestLicenseAdminLinks()
    {
        BackendPolicyStatus admin = ParseLicense("EXPIRED", "EXPIRED", false, false, "active", true);
        const string suffix = "/index.php/settings/admin/ncc_backend_4mc";
        Check("License admin link preserves Nextcloud subpath and port",
            PolicyUiHelper.GetLicenseAdminUrl(admin, "https://cloud.example.test:8443/nextcloud/") == "https://cloud.example.test:8443/nextcloud" + suffix);
        Check("License admin link supports root installations", PolicyUiHelper.GetLicenseAdminUrl(admin, "https://cloud.example.test/") == "https://cloud.example.test" + suffix);
        foreach (string invalidUrl in new[] { null, "", "http://cloud.example.test", "ftp://cloud.example.test", "file:///C:/test",
            "javascript:alert(1)", "https://user:password@cloud.example.test", "https://cloud.example.test/?redirect=outside",
            "https://cloud.example.test/#fragment", "https://cloud.example.test/\r\nheader" })
        {
            Check("License admin link rejects untrusted base: " + invalidUrl, PolicyUiHelper.GetLicenseAdminUrl(admin, invalidUrl) == string.Empty);
        }
        Check("Normal users never receive license management link", PolicyUiHelper.GetLicenseAdminUrl(ParseLicense("EXPIRED", "EXPIRED", false), "https://cloud.example.test") == string.Empty);
        Check("Administrator with active license has no management notice link", PolicyUiHelper.GetLicenseAdminUrl(ParseLicense("ACTIVE", "ACTIVE", true, true, "active", true), "https://cloud.example.test") == string.Empty);
        Check("Seat-only warning does not offer license management", PolicyUiHelper.GetLicenseAdminUrl(ParseLicense("ACTIVE", "ACTIVE", true, true, "suspended_overlimit", true), "https://cloud.example.test") == string.Empty);
        Check("No-seat-only warning does not offer license management", PolicyUiHelper.GetLicenseAdminUrl(ParseLicense("ACTIVE", "ACTIVE", true, false, "active", true), "https://cloud.example.test") == string.Empty);
        Check("GRACE administrator can open license management", PolicyUiHelper.GetLicenseAdminUrl(ParseLicense("GRACE", "GRACE", true, true, "active", true), "https://cloud.example.test") == "https://cloud.example.test" + suffix);
        Check("Missing backend has no license management link", PolicyUiHelper.GetLicenseAdminUrl(null, "https://cloud.example.test") == string.Empty);
    }

    private static void TestLicenseWarningRendering()
    {
        using (var panel = new Panel())
        using (var text = new Label())
        using (var title = new Label())
        using (var link = new LinkLabel())
        {
            panel.Controls.AddRange(new Control[] { text, title, link });
            foreach (string locale in Strings.SupportedLanguageCodes)
            {
                Strings.SetPreferredUiLanguage(locale);
                foreach (bool canManageLicense in new[] { false, true })
                {
                    BackendPolicyStatus noSeat = ParseLicense("ACTIVE", "ACTIVE", true, false, "none", canManageLicense);
                    PolicyUiHelper.ApplyPolicyWarningState(noSeat, panel, text, title, link, "https://cloud.example.test");
                    Check(locale + " no-seat banner explains local use separately from feature hints", panel.Visible && !link.Visible
                        && text.Text == Strings.PolicyWarningNoSeat && text.Text != Strings.SharingPasswordSeparateNoSeatTooltip);
                    Check(locale + " no-seat feature tooltip remains unchanged", PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(noSeat) == Strings.SharingPasswordSeparateNoSeatTooltip);
                    Check(locale + " no-seat notice leaves policies local and Pro disabled", !noSeat.PolicyActive && !PolicyUiHelper.HasBackendSeatEntitlement(noSeat));
                }
            }
            Strings.SetPreferredUiLanguage("en");
            BackendPolicyStatus admin = ParseLicense("EXPIRED", "EXPIRED", false, true, "active", true);
            bool visible = PolicyUiHelper.ApplyPolicyWarningState(admin, panel, text, title, link, "https://cloud.example.test/nextcloud");
            Check("License warning renders cause and admin-only action", visible && panel.Visible && link.Visible
                && text.Text == PolicyUiHelper.GetPolicyWarningMessage(admin)
                && (string)link.Tag == "https://cloud.example.test/nextcloud/index.php/settings/admin/ncc_backend_4mc");
            BrowserLauncher.LastUrl = null;
            PolicyUiHelper.OpenLicenseAdministration(link, LogCategories.Core);
            Check("License action opens only rendered admin target", BrowserLauncher.LastUrl == (string)link.Tag);
            System.Drawing.Color expiredColor = title.ForeColor;
            BackendPolicyStatus grace = ParseLicense("GRACE", "GRACE", true);
            PolicyUiHelper.ApplyPolicyWarningState(grace, panel, text, title, link, "https://cloud.example.test");
            Check("GRACE uses informational color without admin action", panel.Visible && title.ForeColor != expiredColor && !link.Visible && (string)link.Tag == string.Empty);
            string englishText = text.Text;
            Strings.SetPreferredUiLanguage("de");
            PolicyUiHelper.ApplyPolicyWarningState(grace, panel, text, title, link, "https://cloud.example.test");
            Check("Warning text follows language changes without refetch", text.Text == PolicyUiHelper.GetPolicyWarningMessage(grace) && text.Text != englishText);
            Strings.SetPreferredUiLanguage("en");
            PolicyUiHelper.ApplyPolicyWarningState(ParseLicense("ACTIVE", "ACTIVE", true), panel, text, title, link, "https://cloud.example.test");
            Check("Active state clears prior warning and admin target", !panel.Visible && text.Text == string.Empty && (string)link.Tag == string.Empty);
            BrowserLauncher.LastUrl = null;
            PolicyUiHelper.OpenLicenseAdministration(link, LogCategories.Core);
            Check("Cleared action cannot reopen stale admin target", BrowserLauncher.LastUrl == null);
        }
    }

    private static void TestWarningPanelLayout()
    {
        using (var hiddenForm = new Form())
        using (var panel = new Panel())
        using (var text = new Label())
        using (var title = new Label())
        using (var link = new LinkLabel())
        {
            WarningPanelUiHelper.Initialize(panel, title, text, link, Strings.PolicyWarningTitle, string.Empty,
                Strings.PolicyWarningAdminLinkLabel + " " + Strings.PolicyWarningAdminLinkLabel + " " + Strings.PolicyWarningAdminLinkLabel);
            hiddenForm.Controls.Add(panel);
            BackendPolicyStatus admin = ParseLicense("EXPIRED", "EXPIRED", false, true, "active", true);
            PolicyUiHelper.ApplyPolicyWarningState(admin, panel, text, title, link, "https://cloud.example.test");
            text.Text += Environment.NewLine + Strings.PolicyLicenseConnectionError + Environment.NewLine + Strings.PolicyLicenseUserHint;
            int hiddenHeight = WarningPanelUiHelper.Layout(panel, title, text, link, 12, 20, 240, 8, 160, 4, 6);
            Check("Hidden parent defers warning layout", !hiddenForm.Visible && hiddenHeight == 0 && panel.Height == 0);

            // Detaching exposes the control's visibility without opening a desktop window.
            hiddenForm.Controls.Remove(panel);
            int wideHeight = WarningPanelUiHelper.Layout(panel, title, text, link, 12, 20, 600, 8, 160, 4, 6);
            Check("Visible warning reflow restores previously deferred height", panel.Visible && link.Visible && wideHeight > 0 && panel.Height == wideHeight);
            int narrowHeight = WarningPanelUiHelper.Layout(panel, title, text, link, 12, 20, 240, 8, 160, 4, 6);
            Check("Narrow warning wraps long text and admin action", narrowHeight > wideHeight
                && text.Height > text.Font.Height && link.Height > link.Font.Height
                && text.Right <= panel.Width - 8 && link.Right <= panel.Width - 8);
            Check("Visible link contributes its complete height", narrowHeight == link.Bottom + 8 && link.Top == text.Bottom + 6);
            link.Visible = false;
            int withoutLink = WarningPanelUiHelper.Layout(panel, title, text, link, 12, 20, 240, 8, 160, 4, 6);
            Check("Hidden link leaves no reserved gap or action height", withoutLink == text.Bottom + 8 && withoutLink < narrowHeight);
            link.Visible = true;
            int restoredHeight = WarningPanelUiHelper.Layout(panel, title, text, link, 12, 20, 240, 8, 160, 4, 6);
            Check("Repeated layout restores identical warning height", restoredHeight == narrowHeight);
            panel.Visible = false;
            Check("Hidden warning collapses completely", WarningPanelUiHelper.Layout(panel, title, text, link, 12, 20, 240, 8, 160, 4, 6) == 0 && panel.Height == 0);
        }
    }

    private static void CheckLicenseNotice(string name, BackendPolicyStatus status, string expectedCause)
    {
        string message = PolicyUiHelper.GetPolicyWarningMessage(status);
        Check(name + " displays its license cause", message.StartsWith(expectedCause, StringComparison.Ordinal), message);
        Check(name + " displays the matching role hint", message.Contains(status.CanManageLicense ? Strings.PolicyLicenseAdminHint : Strings.PolicyLicenseUserHint), message);
        Check(name + " is not described as a suspended seat", !message.Contains(Strings.PolicyWarningSeatSuspended), message);
        if (!status.IsValid)
        {
            Check(name + " disabled tooltip uses the same cause", PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(status) == message);
        }
    }

    private static BackendPolicyStatus ParseLicense(string license, string access, bool valid, bool assigned = true, string seat = "active", bool admin = false)
    {
        return BackendPolicyService.ParseStatus(LicensePayload(license, access, valid, assigned, seat, admin));
    }

    private static Dictionary<string, object> LicensePayload(string license, string access, bool valid, bool assigned, string seat, bool admin)
    {
        var fields = new Dictionary<string, object>
        {
            { "seat_assigned", assigned }, { "is_valid", valid }, { "seat_state", seat }, { "can_manage_license", admin }
        };
        if (license != null) fields["license_status"] = license;
        if (access != null) fields["access_status"] = access;
        var policy = new Dictionary<string, object>
        {
            { "share", new Dictionary<string, object> { { "share_send_password_mode", "secrets" } } },
            { "talk", new Dictionary<string, object>() },
            { "email_signature", new Dictionary<string, object>
                {
                    { "email_signature_on_compose", true }, { "email_signature_template", "<p>Signature</p>" }, { "user_email", "sender@example.test" }
                }
            }
        };
        var editable = new Dictionary<string, object>
        {
            { "share", new Dictionary<string, object>() }, { "talk", new Dictionary<string, object>() }, { "email_signature", new Dictionary<string, object>() }
        };
        return new Dictionary<string, object> { { "status", fields }, { "policy", policy }, { "policy_editable", editable } };
    }

    private static BackendPolicyStatus BuildSignatureStatus(
        bool onCompose,
        bool onComposeEditable,
        bool includeTemplate,
        bool includeUserEmail)
    {
        var policy = new Dictionary<string, object>
        {
            { "email_signature_on_compose", onCompose },
            { "email_signature_on_reply", false },
            { "email_signature_on_forward", false }
        };
        if (includeTemplate)
        {
            policy["email_signature_template"] = "<p>Backend signature</p>";
        }
        if (includeUserEmail)
        {
            policy["user_email"] = "sender@example.test";
        }

        return new BackendPolicyStatus(
            true,
            true,
            true,
            "policy",
            "policy_active",
            true,
            true,
            "active",
            new Dictionary<string, object>(),
            new Dictionary<string, object>(),
            policy,
            new Dictionary<string, object>(),
            new Dictionary<string, object>(),
            new Dictionary<string, object>
            {
                { "email_signature_on_compose", onComposeEditable },
                { "email_signature_on_reply", true },
                { "email_signature_on_forward", true }
            });
    }

    private static BackendPolicyStatus BuildAttachmentTargetStatus(string target, bool editable)
    {
        return new BackendPolicyStatus(
            true,
            true,
            true,
            "policy",
            "policy_active",
            true,
            true,
            "active",
            new Dictionary<string, object> { { "attachment_link_target", target } },
            new Dictionary<string, object>(),
            new Dictionary<string, object>(),
            new Dictionary<string, object> { { "attachment_link_target", editable } },
            new Dictionary<string, object>(),
            new Dictionary<string, object>());
    }

    private static BackendPolicyStatus BuildStatus(string mode, bool editable, string expireDays)
    {
        return new BackendPolicyStatus(
            true,
            true,
            true,
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
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\AttachmentLinkTargetPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\EmailSignaturePolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryMode.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\EmailSignaturePolicyService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\BackendPolicyService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudUriValidator.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\PolicyUiHelper.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\WarningPanelUiHelper.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\Strings.cs")
    )

    $resources = @(Get-ChildItem -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Resources\_locales") -Directory | ForEach-Object {
        $localeFile = Join-Path $_.FullName "messages.json"
        "/resource:$localeFile,OutlookPolicyMappingTests.Resources._locales.$($_.Name).messages.json"
    })
    $exe = Join-Path $TempRoot "OutlookPolicyMappingTests.exe"
    & $csc /nologo /target:exe "/out:$exe" /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Web.Extensions.dll @resources @sources
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

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
        internal bool IsEnterpriseRollout { get; set; }
        internal bool IsManagedTransportTlsValid { get { return true; } }
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
        TestSeatAccessMatrix();
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

    private static void TestSeatAccessMatrix()
    {
        int cases = 0;
        foreach (string access in new[] { "ACTIVE", "GRACE", "EXPIRED", "INACTIVE", "INVALID", "ACTIVATION_REQUIRED", "OFFLINE_EXPIRED", "UNKNOWN" })
        foreach (string seat in new[] { "active", "suspended_overlimit", "none" })
        foreach (bool overlicensed in new[] { false, true })
        foreach (bool admin in new[] { false, true })
        foreach (bool syncError in new[] { false, true })
        {
            string pairedBanner = null;
            string pairedTooltip = null;
            foreach (string mode in new[] { "community", "pro" })
            {
                bool valid = access == "ACTIVE" || access == "GRACE";
                bool assigned = seat != "none";
                bool usable = valid && seat == "active";
                var payload = LicensePayload(access, access, valid, assigned, seat, admin);
                IDictionary<string, object> fields = NcJson.GetDictionary(payload, "status");
                fields["mode"] = mode;
                fields["overlicensed"] = overlicensed;
                fields["license_connection_error"] = syncError;
                BackendPolicyStatus status = BackendPolicyService.ParseStatus(payload);
                string banner = PolicyUiHelper.GetPolicyWarningMessage(status);
                string tooltip = PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(status);
                string name = mode + "/" + access + "/" + seat + "/over=" + overlicensed + "/admin=" + admin + "/sync=" + syncError;
                Check(name + " personal access and all policy domains", PolicyUiHelper.HasBackendSeatEntitlement(status) == usable
                    && status.PolicyActive == usable && status.IsDomainActive("share") == usable
                    && status.IsDomainActive("talk") == usable && status.IsDomainActive("email_signature") == usable
                    && PolicyUiHelper.HasPasswordDeliveryMode(status) == usable
                    && new EmailSignaturePolicyService(status, new AddinSettings()).Resolve().Active == usable);
                Check(name + " feature tooltip follows personal access", usable ? tooltip == string.Empty
                    : !assigned ? tooltip == Strings.SharingPasswordSeparateNoSeatTooltip : tooltip == banner);
                if (!assigned)
                {
                    Check(name + " banner explains missing seat and local settings", banner.Contains(Strings.PolicyWarningNoSeat));
                    if (admin && (access == "GRACE" || syncError || !valid))
                    {
                        Check(name + " admin also receives license diagnostics", banner.Contains(Strings.PolicyLicenseAdminHint));
                    }
                }
                if (usable && access == "ACTIVE" && !syncError)
                {
                    Check(name + " global capacity alone produces no personal warning", banner == string.Empty);
                }
                if (pairedBanner != null)
                {
                    Check(name + " Community and Pro messages agree", banner == pairedBanner && tooltip == pairedTooltip);
                }
                pairedBanner = banner;
                pairedTooltip = tooltip;
                cases++;
            }
        }
        Check("Seat access matrix covers 384 paired status combinations", cases == 384);
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
        foreach (var malformed in new Dictionary<string, object>[] {
            null,
            new Dictionary<string, object>(),
            new Dictionary<string, object> { { "status", "not-an-object" } },
            new Dictionary<string, object> { { "ocs", new Dictionary<string, object> {
                { "data", new Dictionary<string, object>() }
            } } }
        })
        {
            BackendPolicyStatus rejected = BackendPolicyService.ParseStatus(malformed);
            Check("Missing status object is a fetch failure, never a confirmed missing seat", !rejected.FetchSucceeded
                && rejected.EndpointAvailable && !rejected.PolicyActive && !PolicyUiHelper.HasBackendSeatEntitlement(rejected)
                && PolicyUiHelper.GetPolicyWarningMessage(rejected) == Strings.PolicyWarningBackendUnavailable
                && PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(rejected) == Strings.PolicyWarningBackendUnavailable);
        }
        NcHttpClient.NextResponse = new NcHttpResponse {
            HasHttpResponse = true, StatusCode = HttpStatusCode.OK, ParsedJson = new Dictionary<string, object>()
        };
        Check("HTTP 200 without status cannot replace a last-success cache", !new BackendPolicyService(new TalkServiceConfiguration()).FetchStatus().FetchSucceeded);
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
            Check(name + " disabled tooltip uses the personal cause", PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(status)
                == (status.SeatAssigned ? message : Strings.SharingPasswordSeparateNoSeatTooltip));
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
    # Build the production assembly in an isolated directory for real settings/wizard tests.
    # No Outlook instance, real account, or persisted user profile is used.
    $referencePath = & (Join-Path $ProjectRoot 'tools/ci/Resolve-OfficeExtensibilityReference.ps1') -OutputDirectory (Join-Path $TempRoot 'refs')
    $uiOutput = Join-Path $TempRoot 'bin'
    $msbuild = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/MSBuild.exe'
    & $msbuild (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/NcTalkOutlookAddIn.csproj') /t:Rebuild /v:minimal /p:Configuration=Release "/p:ReferencePath=$referencePath" "/p:OutputPath=$uiOutput\" "/p:IntermediateOutputPath=$TempRoot\obj\"
    if ($LASTEXITCODE -ne 0) { throw 'Production assembly build for policy tests failed.' }
    $uiSource = Join-Path $TempRoot 'OutlookPolicyUiTests.cs'
    @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

internal static class OutlookPolicyUiTests
{
    private const BindingFlags Flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
    private static Assembly Product;
    private static int Checks;
    private static Type T(string name) { return Product.GetType("NcTalkOutlookAddIn." + name, true); }
    private static object New(string name, params object[] args)
    {
        return T(name).GetConstructors(Flags).Single(c => !c.IsStatic && c.GetParameters().Length == args.Length).Invoke(args);
    }
    private static MethodInfo Method(Type type, string name, int count)
    {
        return type.GetMethods(Flags).Single(m => m.Name == name && m.GetParameters().Length == count);
    }
    private static object Call(object target, string name, params object[] args)
    {
        Type type = target as Type ?? target.GetType();
        return Method(type, name, args.Length).Invoke(target is Type ? null : target, args);
    }
    private static object Get(object target, string name) { return target.GetType().GetProperty(name, Flags).GetValue(target, null); }
    private static void Set(object target, string name, object value) { target.GetType().GetProperty(name, Flags).SetValue(target, value, null); }
    private static object Field(object target, string name) { return target.GetType().GetField(name, Flags).GetValue(target); }
    private static void Check(bool condition, string message)
    {
        Checks++;
        if (!condition) throw new InvalidOperationException(message);
    }
    private static object Resolve(object settings, object status) { return Call(settings, "ResolvePolicyDefaults", status); }
    private static Dictionary<string, object> D(params object[] pairs)
    {
        var result = new Dictionary<string, object>();
        for (int i = 0; i < pairs.Length; i += 2) result[(string)pairs[i]] = pairs[i + 1];
        return result;
    }
    private static object Status(Dictionary<string, object> share, Dictionary<string, object> talk, bool editable, string mode, string seat)
    {
        var shareEdit = share.ToDictionary(p => p.Key, p => (object)editable);
        var talkEdit = talk.ToDictionary(p => p.Key, p => (object)editable);
        return Call(T("Services.BackendPolicyService"), "ParseStatus", D(
            "status", D("is_valid", seat != "invalid", "seat_assigned", seat != "none", "seat_state", seat == "paused" ? "suspended_overlimit" : "active", "mode", mode, "overlicensed", true),
            "policy", D("share", share, "talk", talk),
            "policy_editable", D("share", shareEdit, "talk", talkEdit)));
    }
    private static string Serialize(object settings)
    {
        using (var stream = new MemoryStream())
        {
            Call(T("Settings.SettingsStorage"), "SaveToXmlStream", stream, settings, "policy-test");
            return Encoding.UTF8.GetString(stream.ToArray());
        }
    }
    private static object RoundTrip(object settings, string root)
    {
        string path = Path.Combine(root, "policy.xml");
        File.WriteAllText(path, Serialize(settings));
        return Call(T("Settings.SettingsStorage"), "LoadFromXmlFile", path);
    }
    private static object Addressbook() { return New("Services.IfbAddressBookCache+SystemAddressbookStatus", true, 1, ""); }
    private static object Configuration() { return New("Services.TalkServiceConfiguration", "", "", ""); }
    private static Form Settings(object local, object status) { return (Form)New("UI.SettingsForm", local, null, status, null, Addressbook()); }
    private static Form Share(object local, object status, bool attachment)
    {
        object launch = New("Models.FileLinkWizardLaunchOptions");
        Set(launch, "AttachmentMode", attachment);
        return (Form)New("UI.FileLinkWizardForm", local, Configuration(), null, null, status, launch);
    }
    private static Form Talk(object local, object status)
    {
        return (Form)New("UI.TalkLinkForm", local, Configuration(), null, status, null, Addressbook(), "Meeting", DateTime.Today.AddDays(1), DateTime.Today.AddDays(1).AddHours(1));
    }
    private static void TestLocalChoices(string root)
    {
        object roomEvent = Enum.Parse(T("Models.TalkRoomType"), "EventConversation");
        object roomGroup = Enum.Parse(T("Models.TalkRoomType"), "StandardRoom");
        object plain = Enum.Parse(T("Models.SharePasswordDeliveryMode"), "Plain");
        object secrets = Enum.Parse(T("Models.SharePasswordDeliveryMode"), "Secrets");
        object[][] bindings = {
            new object[] { "share", "share_base_directory", "FileLinkBasePath", "Managed", "Local", "Managed" },
            new object[] { "share", "share_name_template", "SharingDefaultShareName", "Managed", "Local", "Managed" },
            new object[] { "share", "share_permission_upload", "SharingDefaultPermCreate", true, false, true },
            new object[] { "share", "share_permission_edit", "SharingDefaultPermWrite", true, false, true },
            new object[] { "share", "share_permission_delete", "SharingDefaultPermDelete", true, false, true },
            new object[] { "share", "share_set_password", "SharingDefaultPasswordEnabled", true, false, true },
            new object[] { "share", "share_send_password_separately", "SharingDefaultPasswordSeparateEnabled", true, false, true },
            new object[] { "share", "share_send_password_mode", "SharingDefaultPasswordDeliveryMode", "secrets", plain, secrets },
            new object[] { "share", "share_expire_days", "SharingDefaultExpireDays", 19, 7, 19 },
            new object[] { "share", "attachments_always_via_ncconnector", "SharingAttachmentsAlwaysConnector", true, false, true },
            new object[] { "share", "language_share_html_block", "ShareBlockLang", "fr", "default", "fr" },
            new object[] { "talk", "talk_lobby_active", "TalkDefaultLobbyEnabled", true, false, true },
            new object[] { "talk", "talk_show_in_search", "TalkDefaultSearchVisible", true, false, true },
            new object[] { "talk", "talk_set_password", "TalkDefaultPasswordEnabled", true, false, true },
            new object[] { "talk", "talk_add_users", "TalkDefaultAddUsers", true, false, true },
            new object[] { "talk", "talk_add_guests", "TalkDefaultAddGuests", true, false, true },
            new object[] { "talk", "talk_delete_room_on_event_delete", "TalkDeleteRoomOnEventDelete", true, false, true },
            new object[] { "talk", "language_talk_description", "EventDescriptionLang", "fr", "default", "fr" },
            new object[] { "talk", "talk_room_type", "TalkDefaultRoomType", "group", roomEvent, roomGroup }
        };
        foreach (object[] binding in bindings)
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (bool editable in new[] { false, true })
        foreach (bool explicitChoice in new[] { false, true })
        {
            string property = (string)binding[2];
            object local = New("Settings.AddinSettings");
            object productDefault = Get(local, property);
            if (explicitChoice) Set(local, property, binding[4]);
            string saved = Serialize(local);
            var policy = D(binding[1], binding[3]);
            object status = Status((string)binding[0] == "share" ? policy : D(), (string)binding[0] == "talk" ? policy : D(), editable, mode, seat);
            object effective = Resolve(local, status);
            object expected = seat == "active" && (!editable || !explicitChoice) ? binding[5] : explicitChoice ? binding[4] : productDefault;
            Check(object.Equals(Get(effective, property), expected), "Resolver precedence: " + property + "/" + mode + "/" + seat + "/" + editable + "/" + explicitChoice);
            Check(Serialize(local) == saved, "Resolution must not modify persisted choices: " + property);
            object reloaded = RoundTrip(local, root);
            Check((bool)Call(reloaded, "HasLocalValue", property) == explicitChoice, "Round-trip retains presence: " + property);
            Check(object.Equals(Get(Resolve(reloaded, status), property), expected), "Round-trip keeps effective choice: " + property);
        }
        object initial = New("Settings.AddinSettings");
        object clone = Call(initial, "Clone");
        Set(clone, "TalkDefaultLobbyEnabled", false);
        Check(!(bool)Call(initial, "HasLocalValue", "TalkDefaultLobbyEnabled"), "Clone has independent presence state");
        object missing = Resolve(initial, Status(D(), D(), true, "community", "active"));
        Check((int)Get(missing, "SharingDefaultExpireDays") == 7, "Missing backend expiry keeps product default");
        foreach (object value in new object[] { 0, "0", null, 1, 19, 3650 })
        {
            var share = D("share_expire_days", value);
            object status = Status(share, D(), false, "pro", "active");
            object actual = Call(status, "GetPolicyValue", "share", "share_expire_days");
            Check(object.Equals(actual, value != null && value.ToString() == "0" ? (object)1 : value), "Legacy expiry normalization");
            Check(object.Equals(share["share_expire_days"], value), "Parser must not mutate incoming policy dictionary");
        }
        Console.WriteLine("[OK] Local-choice precedence and XML round trips across both modes and all personal seat states");
    }
    private static void TestWizards()
    {
        string[][] bools = {
            new[] { "TalkDefaultPasswordEnabled", "_talkDefaultPasswordCheckBox", "_passwordToggleCheckBox" },
            new[] { "TalkDefaultAddUsers", "_talkDefaultAddUsersCheckBox", "_addUsersCheckBox" },
            new[] { "TalkDefaultAddGuests", "_talkDefaultAddGuestsCheckBox", "_addGuestsCheckBox" },
            new[] { "TalkDefaultLobbyEnabled", "_talkDefaultLobbyCheckBox", "_lobbyCheckBox" },
            new[] { "TalkDefaultSearchVisible", "_talkDefaultSearchCheckBox", "_searchCheckBox" }
        };
        foreach (string mode in new[] { "community", "pro" })
        foreach (bool editable in new[] { false, true })
        foreach (bool explicitChoice in new[] { false, true })
        {
            object local = New("Settings.AddinSettings");
            if (explicitChoice) {
                foreach (string[] pair in bools) Set(local, pair[0], false);
                Set(local, "TalkDefaultRoomType", Enum.Parse(T("Models.TalkRoomType"), "EventConversation"));
                Set(local, "SharingDefaultPermCreate", false);
                Set(local, "SharingDefaultExpireDays", 7);
            }
            object status = Status(D("share_permission_upload", true, "share_expire_days", 19), D("talk_set_password", true, "talk_add_users", true, "talk_add_guests", true, "talk_lobby_active", true, "talk_show_in_search", true, "talk_room_type", "group"), editable, mode, "active");
            string before = Serialize(local);
            using (Form options = Settings(local, status))
            using (Form talk = Talk(local, status))
            using (Form share = Share(local, status, false))
            {
                bool expected = !editable || !explicitChoice;
                foreach (string[] pair in bools) {
                    Check(((CheckBox)Field(options, pair[1])).Checked == expected, "Settings choice: " + pair[0]);
                    var control = (CheckBox)Field(talk, pair[2]);
                    Check(control.Checked == expected && control.Enabled == editable, "Talk control choice and lock: " + pair[0]);
                }
                Check(((ComboBox)Field(options, "_talkDefaultRoomTypeCombo")).SelectedIndex == ((ComboBox)Field(talk, "_roomTypeComboBox")).SelectedIndex, "Room-type UI agreement");
                Call(talk, "OnOkButtonClick", null, EventArgs.Empty);
                Check((bool)Get(talk, "LobbyUntilStart") == expected && (bool)Get(talk, "AddUsers") == expected && (bool)Get(talk, "AddGuests") == expected && (bool)Get(talk, "SearchVisible") == expected, "Talk confirmation retains UI selections");
                Check(Get(talk, "SelectedRoomType").ToString() == (expected ? "StandardRoom" : "EventConversation"), "Confirmed room type");
                Call(share, "ApplyFormData");
                object request = Field(share, "_request");
                Check(((Convert.ToInt32(Get(request, "Permissions")) & 4) != 0) == expected, "Actual share request upload permission");
                Check(((DateTime)Get(request, "ExpireDate") - DateTime.Today).Days == (expected ? 19 : 7), "Actual share request expiration");
                Check(Serialize(Get(options, "Result")) == before && Serialize(local) == before, "Opening settings and wizards preserves raw local state");
            }
        }
        foreach (int days in new[] { 0, 1, 19, 3650 })
        foreach (bool attachment in new[] { false, true })
        {
            object status = Status(D("share_expire_days", days), D(), false, "community", "active");
            using (Form share = Share(New("Settings.AddinSettings"), status, attachment)) {
                Call(share, "ApplyFormData");
                object request = Field(share, "_request");
                Check((bool)Get(request, "ExpireEnabled") && ((DateTime)Get(request, "ExpireDate") - DateTime.Today).Days == Math.Max(1, days), "Legacy expiry agrees in manual and automated share request");
            }
        }
        object noExpiry = New("Settings.AddinSettings");
        Set(noExpiry, "SharingDefaultExpireDays", 0);
        using (Form share = Share(noExpiry, Status(D("share_expire_days", 19), D(), true, "pro", "active"), false)) {
            Call(share, "ApplyFormData");
            Check(!(bool)Get(Field(share, "_request"), "ExpireEnabled"), "Explicit local zero still disables expiration");
        }
        Console.WriteLine("[OK] Real Settings, Talk and Sharing controls, lock state, confirmation and request payloads");
    }
    private static void TestSettingsEdits(string root)
    {
        object local = New("Settings.AddinSettings");
        object editable = Status(D("share_permission_upload", true, "share_expire_days", 19, "language_share_html_block", "fr"), D("talk_lobby_active", true), true, "pro", "active");
        using (Form options = Settings(local, editable))
        {
            object result = Get(options, "Result");
            Check(!(bool)Call(result, "HasLocalValue", "SharingDefaultPermCreate"), "Opening options does not select a backend default");
            ((CheckBox)Field(options, "_sharingDefaultPermCreateCheckBox")).Checked = false;
            ((CheckBox)Field(options, "_talkDefaultLobbyCheckBox")).Checked = false;
            ((NumericUpDown)Field(options, "_sharingDefaultExpireDaysUpDown")).Value = 7;
            Check((bool)Call(result, "HasLocalValue", "SharingDefaultPermCreate") && !(bool)Get(result, "SharingDefaultPermCreate"), "User checkbox edit is explicit even when false");
            Check((bool)Call(result, "HasLocalValue", "SharingDefaultExpireDays") && (int)Get(result, "SharingDefaultExpireDays") == 7, "Product-default numeric choice is explicit");
            object locked = Status(D("share_permission_upload", true, "share_expire_days", 19), D("talk_lobby_active", true), false, "pro", "active");
            options.GetType().GetField("_backendPolicyStatus", Flags).SetValue(options, locked);
            Call(options, "ApplyBackendPolicyStatus", "test_lock");
            Check(((CheckBox)Field(options, "_sharingDefaultPermCreateCheckBox")).Checked && !((CheckBox)Field(options, "_sharingDefaultPermCreateCheckBox")).Enabled, "Locked overlay is visible and disabled");
            Check(!(bool)Get(result, "SharingDefaultPermCreate") && !(bool)Get(result, "TalkDefaultLobbyEnabled"), "Lock never replaces saved choices");
            options.GetType().GetField("_backendPolicyStatus", Flags).SetValue(options, editable);
            Call(options, "ApplyBackendPolicyStatus", "test_unlock");
            Check(!((CheckBox)Field(options, "_sharingDefaultPermCreateCheckBox")).Checked && !((CheckBox)Field(options, "_talkDefaultLobbyCheckBox")).Checked, "Unlock restores local choices in both domains");
            object restored = RoundTrip(result, root);
            Check(!(bool)Get(Resolve(restored, editable), "SharingDefaultPermCreate"), "User choice survives persistence and reopening");
        }
        using (Form options = Settings(New("Settings.AddinSettings"), editable))
        {
            ((TextBox)Field(options, "_usernameTextBox")).Text = "test-only-login";
            Task save = (Task)Call(options, "SaveSettingsAsync");
            DateTime deadline = DateTime.UtcNow.AddSeconds(15);
            while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
            Check(save.IsCompleted, "Local settings save completes without a configured server");
            save.GetAwaiter().GetResult();
            string xml = Serialize(Get(options, "Result"));
            Check(xml.Contains("test-only-login") && !xml.Contains("<SharingDefault") && !xml.Contains("<TalkDefault") && !xml.Contains("<ShareBlockLang>"), "Actual credentials-only save leaves untouched policy choices absent");
        }
        Console.WriteLine("[OK] User edits, lock/unlock, XML persistence and actual credentials-only save");
    }
    private static void TestAttachmentAutomation()
    {
        Type subscription = T("NextcloudTalkAddIn+MailComposeSubscription");
        int cases = 0;
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (bool editable in new[] { false, true })
        foreach (bool explicitChoice in new[] { false, true })
        foreach (bool always in new[] { false, true })
        foreach (object threshold in new object[] { null, 0, 1, 19, 10240 })
        {
            object local = New("Settings.AddinSettings");
            if (explicitChoice) {
                Set(local, "SharingAttachmentsAlwaysConnector", false);
                Set(local, "SharingAttachmentsOfferAboveEnabled", false);
                Set(local, "SharingAttachmentsOfferAboveMb", 20);
            }
            object status = Status(D("attachments_always_via_ncconnector", always, "attachments_min_size_mb", threshold), D(), editable, mode, seat);
            object snapshot = Call(subscription, "BuildAttachmentAutomationSettings", local, local);
            object actual = Call(subscription, "ApplyAttachmentAutomationPolicy", snapshot, status);
            bool useBackend = seat == "active" && (!editable || !explicitChoice);
            bool expectedAlways = useBackend && always;
            bool expectedEnabled = !expectedAlways && (useBackend ? threshold != null : !explicitChoice);
            int expectedThreshold = useBackend && threshold != null ? ((int)threshold == 0 ? 5 : (int)threshold) : 20;
            Check((bool)Get(actual, "AlwaysConnector") == expectedAlways && (bool)Get(actual, "OfferAboveEnabled") == expectedEnabled && (int)Get(actual, "ThresholdMb") == expectedThreshold && (long)Get(actual, "ThresholdBytes") == expectedThreshold * 1024L * 1024L, "Operative attachment policy case " + cases);
            cases++;
        }
        foreach (object threshold in new object[] { null, 0, 1, 19, 10240 }) {
            object status = Status(D("attachments_min_size_mb", threshold), D(), false, "pro", "active");
            using (Form options = Settings(New("Settings.AddinSettings"), status)) {
                Check(((CheckBox)Field(options, "_sharingAttachmentsOfferAboveCheckBox")).Checked == (threshold != null), "Threshold UI null semantics");
                Check(((NumericUpDown)Field(options, "_sharingAttachmentsOfferAboveMbUpDown")).Value == (threshold == null ? 20 : (int)threshold == 0 ? 5 : (int)threshold), "Threshold UI matches operative value");
            }
        }
        Console.WriteLine("[OK] " + cases + " operative attachment combinations plus real threshold controls");
    }
    private static object Managed(object url, object locked, object ribbon, string source)
    {
        return New("Settings.ManagedSetupPolicy", url, locked, ribbon, null, null, null, source);
    }
    private static object ManagedTls(object system, object tls12, object tls13, string source = "TLS test")
    {
        return New("Settings.ManagedSetupPolicy", null, null, null, system, tls12, tls13, source);
    }
    private static readonly string[] TlsProperties = {
        "TransportTlsUseSystemDefault", "TransportTlsEnable12", "TransportTlsEnable13"
    };
    private static void CheckTls(object target, bool system, bool tls12, bool tls13, string message, bool raw = false)
    {
        bool[] expected = { system, tls12, tls13 };
        for (int i = 0; i < TlsProperties.Length; i++)
            Check((bool)Get(target, (raw ? "Local" : "") + TlsProperties[i]) == expected[i], message + ": " + TlsProperties[i]);
    }
    private static void TestManagedTls(string root)
    {
        CheckTls(New("Settings.AddinSettings"), false, true, false, "Unmanaged TLS product defaults are unchanged");
        foreach (object system in new object[] { null, 0, 1 })
        foreach (object tls12 in new object[] { null, 0, 1 })
        foreach (object tls13 in new object[] { null, 0, 1 })
        {
            bool present = system != null || tls12 != null || tls13 != null;
            bool expectedSystem = object.Equals(system, 1);
            bool expected12 = tls12 == null || object.Equals(tls12, 1);
            bool expected13 = object.Equals(tls13, 1);
            bool valid = expectedSystem || expected12 || expected13;
            object policy = ManagedTls(system, tls12, tls13);
            Check((bool)Get(policy, "HasTransportTlsPolicy") == present && (bool)Get(policy, "IsEnterpriseRollout") == present,
                "TLS policy presence, including zero, controls enterprise rollout");
            Check((bool)Get(policy, "IsTransportTlsPolicyValid") == valid, "All-false managed TLS is invalid; missing siblings use product defaults");
            CheckTls(policy, expectedSystem, expected12, expected13, "Managed TLS maps missing/zero/one independently");
            object local = New("Settings.AddinSettings");
            Set(local, TlsProperties[0], !expectedSystem);
            Set(local, TlsProperties[1], !expected12);
            Set(local, TlsProperties[2], !expected13);
            Call(local, "ApplyManagedSetupPolicy", policy);
            Check((bool)Get(local, "HasManagedTransportTls") == present && (bool)Get(local, "IsManagedTransportTlsValid") == valid,
                "Settings preserve managed TLS presence and validity");
            CheckTls(local, present ? expectedSystem : !expectedSystem, present ? expected12 : !expected12,
                present ? expected13 : !expected13, "Policy overlay uses product defaults, not local siblings");
            CheckTls(local, !expectedSystem, !expected12, !expected13, "Applying TLS policy retains local choices", true);
            object clone = Call(local, "Clone");
            Check((bool)Get(clone, "HasManagedTransportTls") == present && (bool)Get(clone, "IsManagedTransportTlsValid") == valid,
                "Clone retains managed TLS state including invalid policy");
            CheckTls(clone, !expectedSystem, !expected12, !expected13, "Clone retains raw TLS choices", true);
            Call(clone, "ApplyManagedSetupPolicy", (object)null);
            Check(!(bool)Get(clone, "HasManagedTransportTls") && (bool)Get(clone, "IsManagedTransportTlsValid"), "Removing TLS policy clears validity failure");
            CheckTls(clone, !expectedSystem, !expected12, !expected13, "Removing TLS policy restores all local values");
            Check((bool)Get(local, "HasManagedTransportTls") == present, "Removing policy from clone does not mutate original");
        }

        string[] sources = { "HKLM64", "HKLM32", "HKCU64", "HKCU32" };
        for (int field = 0; field < 3; field++)
        for (int location = 0; location < sources.Length; location++)
        foreach (object selected in new object[] { 0, 1, "malformed" })
        {
            Array policies = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 4);
            object[] values = { null, null, null };
            values[field] = selected;
            policies.SetValue(ManagedTls(values[0], values[1], values[2], sources[location]), location);
            for (int lower = location + 1; lower < sources.Length; lower++)
                policies.SetValue(ManagedTls(1, 1, 1, sources[lower]), lower);
            object resolved = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
            Check((bool)Get(resolved, "HasTransportTlsPolicy") && (bool)Get(resolved, "IsEnterpriseRollout"), "Any TLS value in each hive/view activates enterprise rollout");
            if (selected is string)
                Check(!(bool)Get(resolved, "IsTransportTlsPolicyValid"), "Malformed first-present TLS cannot fall through to a lower-priority valid value");
            else
                Check((bool)Get(resolved, TlsProperties[field]) == object.Equals(selected, 1), "First-present TLS wins even when zero: " + sources[location] + "/" + TlsProperties[field]);
        }
        Array mixed = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 4);
        mixed.SetValue(ManagedTls(0, null, null, "HKLM64"), 0);
        mixed.SetValue(ManagedTls(null, 0, null, "HKLM32"), 1);
        mixed.SetValue(ManagedTls(null, null, 1, "HKCU64"), 2);
        mixed.SetValue(ManagedTls(1, 1, 0, "HKCU32"), 3);
        object mixedPolicy = Call(T("Settings.ManagedSetupPolicy"), "Resolve", mixed);
        CheckTls(mixedPolicy, false, false, true, "TLS fields resolve independently across all four registry locations");
        Check((bool)Get(mixedPolicy, "IsTransportTlsPolicyValid"), "Resolved TLS validity uses the merged group");
        mixed.SetValue(ManagedTls(0, 1, 0, "HKLM64"), 0);
        mixed.SetValue(ManagedTls("bad", "bad", "bad", "HKLM32"), 1);
        mixedPolicy = Call(T("Settings.ManagedSetupPolicy"), "Resolve", mixed);
        CheckTls(mixedPolicy, false, true, false, "Lower-priority malformed TLS cannot replace selected valid fields");
        Check((bool)Get(mixedPolicy, "IsTransportTlsPolicyValid"), "Only selected malformed values invalidate TLS policy");
        for (int field = 0; field < 3; field++)
        foreach (object malformed in new object[] { "", "garbage", new byte[] { 1, 2 } })
        {
            object[] values = { 1, 1, 1 };
            values[field] = malformed;
            object policy = ManagedTls(values[0], values[1], values[2]);
            Check((bool)Get(policy, "HasTransportTlsPolicy") && (bool)Get(policy, "IsEnterpriseRollout")
                && !(bool)Get(policy, "IsTransportTlsPolicyValid"), "Malformed TLS is managed and invalid even if another protocol or system default is enabled");
        }

        for (int rawBits = 0; rawBits < 8; rawBits++)
        {
            object local = New("Settings.AddinSettings");
            bool[] raw = { (rawBits & 1) != 0, (rawBits & 2) != 0, (rawBits & 4) != 0 };
            for (int field = 0; field < 3; field++) Set(local, TlsProperties[field], raw[field]);
            Call(local, "ApplyManagedSetupPolicy", ManagedTls(1, 0, 0));
            string xml = Serialize(local);
            var doc = new XmlDocument(); doc.LoadXml(xml);
            for (int field = 0; field < 3; field++)
                Check(bool.Parse(doc.DocumentElement[TlsProperties[field]].InnerText) == raw[field], "XML writes raw TLS under the existing element name");
            Check(!xml.Contains("HasManagedTransportTls") && !xml.Contains("IsManagedTransportTlsValid")
                && !xml.Contains("LocalTransportTls") && !xml.Contains("IsEnterpriseRollout"), "Registry TLS state is never serialized into user XML");
            object loaded = RoundTrip(local, root);
            Check(!(bool)Get(loaded, "HasManagedTransportTls"), "Loading XML alone cannot create a registry TLS policy");
            CheckTls(loaded, raw[0], raw[1], raw[2], "XML round trip retains the user's TLS values");
            object clone = Call(local, "Clone");
            for (int field = 0; field < 3; field++) Set(clone, TlsProperties[field], !raw[field]);
            CheckTls(clone, true, false, false, "Setters cannot bypass managed TLS getters");
            CheckTls(clone, !raw[0], !raw[1], !raw[2], "Setters retain local choices while managed", true);
            CheckTls(local, raw[0], raw[1], raw[2], "Clone edits do not alter original local TLS values", true);
            Call(clone, "ApplyManagedSetupPolicy", Managed(null, null, null, "empty"));
            CheckTls(clone, !raw[0], !raw[1], !raw[2], "An empty policy restores updated local TLS choices");
        }

        TestManagedTlsUi();
        Console.WriteLine("[OK] Registry TLS presence, validity, per-field priority, overlay, clone, XML and locked UI regressions");
    }
    private static void TestManagedTlsUi()
    {
        string hint = (string)T("Utilities.Strings").GetProperty("AdvancedTlsManagedHint", Flags).GetValue(null, null);
        string invalid = (string)T("Utilities.Strings").GetProperty("ManagedTlsPolicyInvalid", Flags).GetValue(null, null);
        string[] controls = { "_tlsUseSystemDefaultCheckBox", "_tlsEnable12CheckBox", "_tlsEnable13CheckBox" };
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (object[] values in new[] {
            new object[] { null, null, null }, new object[] { 0, null, null },
            new object[] { 1, 0, 0 }, new object[] { 0, 0, 0 }, new object[] { "bad", 1, 0 }
        })
        {
            object local = New("Settings.AddinSettings");
            object policy = ManagedTls(values[0], values[1], values[2]);
            Call(local, "ApplyManagedSetupPolicy", policy);
            // Complete credentials avoid onboarding taking precedence over policy warning tests.
            Set(local, "ServerUrl", "https://cloud.example.test");
            Set(local, "Username", "tls-policy-test");
            Set(local, "AppPassword", "test-only");
            object status = Status(D(), D(), true, mode, seat);
            bool managed = (bool)Get(local, "HasManagedTransportTls");
            bool valid = (bool)Get(local, "IsManagedTransportTlsValid");
            using (Form form = Settings(local, status))
            {
                for (int field = 0; field < 3; field++)
                    Check(((CheckBox)Field(form, controls[field])).Checked == (bool)Get(local, TlsProperties[field]), "UI displays effective TLS including invalid all-false policy");
                foreach (bool busy in new[] { false, true, false })
                {
                    Call(form, "SetBusy", busy);
                    Call(form, "UpdateControlState");
                    Call(form, "UpdateTlsOptionsState");
                    for (int field = 0; field < 3; field++)
                    {
                        bool enabled = !managed && !busy && (field == 0 || !(bool)Get(local, TlsProperties[0]));
                        Check(((CheckBox)Field(form, controls[field])).Enabled == enabled, "Every TLS control remains locked across state/busy updates: " + mode + "/" + seat);
                    }
                }
                Check(((Label)Field(form, "_tlsHintLabel")).Text.Contains(hint) == managed, "TLS hint identifies organization management only when policy exists");
                if (managed && !valid)
                {
                    Check(((Label)Field(form, "_policyWarningTextLabel")).Text.Contains(invalid), "Invalid managed TLS has an actionable banner independent of mode or seat");
                    Check(RolloutNotice(local, status) == invalid && RolloutNotice(local, null) == invalid,
                        "Invalid TLS takes priority over cached backend/seat state and missing backend status");
                }
                CheckTls(Get(form, "Result"), false, true, false, "Opening settings and updating controls retain raw local TLS choices", true);
                CheckTls(local, false, true, false, "Opening settings does not mutate original raw TLS choices", true);
                var persisted = new XmlDocument(); persisted.LoadXml(Serialize(Get(form, "Result")));
                Check(!bool.Parse(persisted.DocumentElement[TlsProperties[0]].InnerText)
                    && bool.Parse(persisted.DocumentElement[TlsProperties[1]].InnerText)
                    && !bool.Parse(persisted.DocumentElement[TlsProperties[2]].InnerText), "Settings result persists raw TLS rather than its managed display");
                if (managed && valid)
                    Check(RolloutNotice(local, status) == RolloutNotice(ManagedSettings(null, true), status), "TLS-only enterprise rollout retains Community/Pro personal-seat rules");
            }
        }
    }
    private static object ManagedSettings(object url, object ribbon, object locked = null)
    {
        object settings = New("Settings.AddinSettings");
        Call(settings, "ApplyManagedSetupPolicy", Managed(url, locked, ribbon, "test"));
        return settings;
    }
    private static string RolloutNotice(object settings, object status)
    {
        return (string)Call(T("Utilities.PolicyUiHelper"), "GetEnterpriseRolloutNotice", settings, status);
    }
    private static void TestEnterpriseRollout(string root)
    {
        foreach (object url in new object[] { null, "", "invalid url", "https://cloud.example.test/nextcloud" })
        foreach (object locked in new object[] { null, true, false, 1, 0, "true", "false", "", "bad value" })
        foreach (object ribbon in new object[] { null, true, false, 1, 0, "true", "false", "bad value" })
        {
            object local = ManagedSettings(url, ribbon, locked);
            bool managed = url != null || locked != null || ribbon != null;
            bool visible = !(object.Equals(ribbon, false) || object.Equals(ribbon, 0) || object.Equals(ribbon, "false"));
            Check((bool)Get(local, "IsEnterpriseRollout") == managed, "Managed detection depends on presence");
            Check((bool)Get(local, "ShowMainRibbonTab") == visible, "Ribbon visibility default and boolean value");
            object cloned = Call(local, "Clone");
            Check((bool)Get(cloned, "IsEnterpriseRollout") == managed && (bool)Get(cloned, "ShowMainRibbonTab") == visible, "Clone preserves registry state");
            string xml = Serialize(local);
            Check(!xml.Contains("IsEnterpriseRollout") && !xml.Contains("ShowMainRibbonTab"), "Managed flags cannot be persisted in profile XML");
            object addin = New("NextcloudTalkAddIn");
            addin.GetType().GetField("_currentSettings", Flags).SetValue(addin, local);
            Check((bool)Call(addin, "OnGetMainRibbonTabVisible", (object)null) == visible, "Main tab callback follows policy");
            string explorer = (string)Call(addin, "GetCustomUI", "Microsoft.Outlook.Explorer");
            var doc = new XmlDocument(); doc.LoadXml(explorer);
            var tab = (XmlElement)doc.SelectSingleNode("//*[local-name()='tab' and @id='NcTalkExplorerTab']");
            var group = (XmlElement)tab.SelectSingleNode("*[local-name()='group' and @id='NcTalkExplorerGroup']");
            var button = (XmlElement)group.SelectSingleNode("*[local-name()='button' and @id='NcTalkSettingsExplorerButton']");
            Check(tab.GetAttribute("getVisible") == "OnGetMainRibbonTabVisible", "Main tab owns Settings visibility");
            Check(!group.HasAttribute("getVisible") && !group.HasAttribute("visible") && !button.HasAttribute("getVisible") && !button.HasAttribute("visible"), "Visible main tab retains the Settings button without a separate visibility rule");
            Check(button.GetAttribute("onAction") == "OnSettingsButtonPressed" && button.GetAttribute("getImage") == "OnGetButtonImage", "Settings action and icon remain wired");
            Check(explorer.Contains("OnFileLinkButtonPressed") && !explorer.Contains("OnEnterpriseStatus"), "Inline action retained without replacement button");
        }
        for (int location = 0; location < 4; location++)
        foreach (object locked in new object[] { true, false, 1, 0, "true", "false", "", "bad value" })
        {
            Array lockPolicies = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 4);
            lockPolicies.SetValue(Managed(null, locked, null, "lock-only"), location);
            object lockPolicy = Call(T("Settings.ManagedSetupPolicy"), "Resolve", lockPolicies);
            object justLock = New("Settings.AddinSettings");
            Set(justLock, "ServerUrl", "https://saved.example.test");
            Call(justLock, "ApplyManagedSetupPolicy", lockPolicy);
            Check((bool)Get(justLock, "IsEnterpriseRollout") && (bool)Get(justLock, "ShowMainRibbonTab"), "Lock-only policy activates rollout in every hive/view with visible tab default");
            Check((string)Get(justLock, "ServerUrl") == "https://saved.example.test" && !(bool)Get(justLock, "ManagedNextcloudUrlLocked"), "Lock-only policy preserves the editable profile URL");
        }
        Array policies = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 3);
        policies.SetValue(Managed(null, null, false, "HKLM64"), 0);
        policies.SetValue(Managed("https://machine.example.test", 1, null, "HKLM32"), 1);
        policies.SetValue(Managed("https://user.example.test", 0, true, "HKCU64"), 2);
        object resolved = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
        Check((string)Get(resolved, "NextcloudUrl") == "https://machine.example.test" && (bool)Get(resolved, "NextcloudUrlLocked") && !(bool)Get(resolved, "ShowMainRibbonTab"), "Machine precedence resolved independently per setting");
        policies.SetValue(Managed(null, 1, null, "HKLM64"), 0);
        policies.SetValue(null, 1);
        policies.SetValue(Managed("https://user.example.test", 0, false, "HKCU64"), 2);
        resolved = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
        Check((bool)Get(resolved, "IsEnterpriseRollout") && (string)Get(resolved, "NextcloudUrl") == "https://user.example.test" && !(bool)Get(resolved, "NextcloudUrlLocked") && !(bool)Get(resolved, "ShowMainRibbonTab"), "Machine lock-only policy does not lock a separate user URL or discard its ribbon value");
        policies.SetValue(Managed("https://machine.example.test", 0, null, "HKLM64"), 0);
        policies.SetValue(Managed(null, 1, null, "HKCU64"), 2);
        resolved = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
        Check((string)Get(resolved, "NextcloudUrl") == "https://machine.example.test" && !(bool)Get(resolved, "NextcloudUrlLocked"), "User lock-only policy cannot change the machine URL lock");
        object clear = ManagedSettings(null, false);
        Call(clear, "ApplyManagedSetupPolicy", (object)null);
        Check(!(bool)Get(clear, "IsEnterpriseRollout") && (bool)Get(clear, "ShowMainRibbonTab"), "Removing registry policy restores defaults");

        int states = 0;
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (bool managed in new[] { false, true })
        foreach (bool lockOnly in new[] { false, true })
        {
            object local = ManagedSettings(null, managed && !lockOnly ? (object)true : null, managed && lockOnly ? (object)false : null);
            object status = Status(D("attachments_always_via_ncconnector", true), D(), false, mode, seat);
            bool blocked = managed && seat != "active";
            Check(string.IsNullOrEmpty(RolloutNotice(local, status)) == !blocked, "Community/Pro personal rollout access " + states);
            Type subscription = T("NextcloudTalkAddIn+MailComposeSubscription");
            object snapshot = Call(subscription, "BuildAttachmentAutomationSettings", local, local);
            object effective = Call(subscription, "ApplyAttachmentAutomationPolicy", snapshot, status);
            Check((bool)Get(effective, "EnterpriseRolloutBlocked") == blocked, "Automation uses common rollout gate " + states);
            if (blocked) Check(!(bool)Get(effective, "AlwaysConnector") && !(bool)Get(effective, "OfferAboveEnabled"), "Blocked automation leaves Outlook attachments alone");
            states++;
        }
        object rollout = ManagedSettings(null, null, false);
        object good = Status(D(), D(), true, "community", "active");
        object missing = Status(D(), D(), true, "pro", "none");
        Set(missing, "EndpointAvailable", false);
        object unavailable = Status(D(), D(), true, "pro", "active");
        Set(unavailable, "FetchSucceeded", false);
        string backendMessage = RolloutNotice(rollout, missing);
        string seatMessage = RolloutNotice(rollout, Status(D(), D(), true, "pro", "none"));
        string connectionMessage = RolloutNotice(rollout, unavailable);
        Check(backendMessage.Length > 0 && seatMessage.Length > 0 && connectionMessage.Length > 0 && backendMessage != seatMessage && backendMessage != connectionMessage && seatMessage != connectionMessage, "Missing backend, absent seat and connection error have distinct messages");
        Check(RolloutNotice(rollout, null) == connectionMessage, "Unconfirmed state is not a missing backend or seat");
        foreach (object status in new[] { good, missing, unavailable, null })
            Check(RolloutNotice(New("Settings.AddinSettings"), status) == "", "Unmanaged installations unchanged");
        foreach (object url in new object[] { null, "https://cloud.example.test" })
        foreach (object locked in new object[] { null, true, false })
        foreach (object ribbon in new object[] { null, true, false })
        {
            using (Form form = Settings(ManagedSettings(url, ribbon, locked), good))
            {
                var tabs = (TabControl)Field(form, "_tabControl");
                bool hidden = object.Equals(ribbon, false);
                Check(tabs.TabPages.Count == (hidden ? 1 : 8) && tabs.TabPages[0] == Field(form, "_generalTab"), "Only a hidden main tab restricts settings to authentication");
                Check(!((Button)Field(form, "_loginFlowButton")).IsDisposed && !((Button)Field(form, "_saveButton")).IsDisposed, "Existing login and save controls retained");
                var serverUrl = (TextBox)Field(form, "_serverUrlTextBox");
                Check(serverUrl.Enabled == !(url != null && object.Equals(locked, true)), "URL editing follows its own lock, not ribbon visibility");
            }
        }
        foreach (object ribbon in new object[] { null, true, false })
        using (Form form = Settings(ManagedSettings(null, ribbon, false), good))
        {
            object result = Get(form, "Result");
            bool originalDebug = (bool)Get(result, "DebugLoggingEnabled");
            bool originalUpdate = (bool)Get(result, "UpdateNotifyEnabled");
            ((TextBox)Field(form, "_usernameTextBox")).Text = "rollout-login";
            ((CheckBox)Field(form, "_debugLogCheckBox")).Checked = !originalDebug;
            ((CheckBox)Field(form, "_updateNotifyCheckBox")).Checked = !originalUpdate;
            Task save = (Task)Call(form, "SaveSettingsAsync");
            DateTime deadline = DateTime.UtcNow.AddSeconds(15);
            while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
            Check(save.IsCompleted, "Managed settings save completes without network credentials");
            save.GetAwaiter().GetResult();
            bool hidden = object.Equals(ribbon, false);
            Check((form.DialogResult == DialogResult.OK) == !hidden, "Only authentication-only setup requires complete credentials before saving");
            Check((bool)Get(result, "DebugLoggingEnabled") == (hidden ? originalDebug : !originalDebug) && (bool)Get(result, "UpdateNotifyEnabled") == (hidden ? originalUpdate : !originalUpdate), "Visible managed Settings saves non-authentication preferences");
            Check((string)Get(result, "Username") == (hidden ? "" : "rollout-login"), "Incomplete hidden setup cannot replace credentials");
        }
        object owner = New("NextcloudTalkAddIn");
        object config = New("Services.TalkServiceConfiguration", "https://cloud.example.test", "alice", "test-only");
        Call(owner, "StoreBackendPolicySnapshot", config, good, "test");
        DateTime previousSuccess = DateTime.UtcNow.AddMinutes(-30);
        owner.GetType().GetField("_emailSignaturePolicyCacheFetchedAtUtc", Flags).SetValue(owner, previousSuccess);
        Check(object.ReferenceEquals(Call(owner, "StoreBackendPolicySnapshot", config, unavailable, "test"), good), "Failed refresh retains confirmed snapshot");
        Check((DateTime)Field(owner, "_emailSignaturePolicyCacheFetchedAtUtc") == previousSuccess, "Failed refresh does not mark the retained snapshot fresh");
        Call(owner, "StoreBackendPolicySnapshot", config, missing, "test");
        Check(object.ReferenceEquals(Call(owner, "StoreBackendPolicySnapshot", config, unavailable, "test"), missing), "Fresh refusal replaces previous success");
        object other = New("Services.TalkServiceConfiguration", "https://cloud.example.test", "bob", "test-only");
        Check(object.ReferenceEquals(Call(owner, "StoreBackendPolicySnapshot", other, unavailable, "test"), unavailable), "Rollout cache cannot cross account identity");
        Set(rollout, "ServerUrl", "https://cloud.example.test");
        Set(rollout, "Username", "alice");
        Set(rollout, "AppPassword", "test-only");
        owner.GetType().GetField("_currentSettings", Flags).SetValue(owner, rollout);
        Type subscriptionType = T("NextcloudTalkAddIn+MailComposeSubscription");
        object openCompose = System.Runtime.Serialization.FormatterServices.GetUninitializedObject(subscriptionType);
        subscriptionType.GetField("_owner", Flags).SetValue(openCompose, owner);
        object oldRules = Call(subscriptionType, "BuildAttachmentAutomationSettings", rollout, rollout);
        Set(oldRules, "AlwaysConnector", true);
        subscriptionType.GetField("_attachmentAutomationSettingsSnapshot", Flags).SetValue(openCompose, oldRules);
        subscriptionType.GetField("_attachmentAutomationSettingsSnapshotUtc", Flags).SetValue(openCompose, DateTime.UtcNow);
        object newRules = Call(openCompose, "ReadAttachmentAutomationSettings");
        Check((bool)Get(newRules, "EnterpriseRolloutBlocked") && !(bool)Get(newRules, "AlwaysConnector"), "Confirmed refusal immediately overrides an open compose's older routing rules");
        Console.WriteLine("[OK] Enterprise presence, registry precedence, ribbon, onboarding, seat parity, automation and cache checks");
    }
    private static void TestConnectionOnboarding()
    {
        string title = (string)T("Utilities.Strings").GetProperty("ConnectionSetupTitle", Flags).GetValue(null, null);
        string message = (string)T("Utilities.Strings").GetProperty("ConnectionSetupMessage", Flags).GetValue(null, null);
        string rejected = (string)T("Utilities.Strings").GetProperty("ConnectionSignInRequired", Flags).GetValue(null, null);
        foreach (object url in new object[] { null, "https://cloud.example.test" })
        foreach (object ribbon in new object[] { null, true, false })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        {
            object local = ManagedSettings(url, ribbon);
            object status = Status(D(), D(), true, "community", seat);
            using (Form form = Settings(local, status))
            {
                Check(((Label)Field(form, "_policyWarningTitleLabel")).Text == title, "Incomplete setup starts with a connection invitation, not a seat failure");
                Check(((Label)Field(form, "_policyWarningTextLabel")).Text == message, "Managed and local setup share the friendly message");
                Check(((LinkLabel)Field(form, "_policyWarningLinkLabel")).Tag == null, "Setup cannot expose a stale admin action");
                Call(form, "BeginAuthentication", true);
                Check(((Label)Field(form, "_policyWarningTextLabel")).Text == rejected, "Rejected authentication has specific guidance");
                ((TextBox)Field(form, "_usernameTextBox")).Text = "onboarding-test";
                Check(!(bool)Field(form, "_authenticationRejected") && (bool)Field(form, "_connectionSetupPending"), "Editing credentials clears rejection but still requires verification");
                Check(((Label)Field(form, "_policyWarningTextLabel")).Text == message, "Editing credentials restores connection invitation");
                Task save = (Task)Call(form, "SaveSettingsAsync");
                Check(save.IsCompleted, "Incomplete onboarding save does not start network calls");
                save.GetAwaiter().GetResult();
                Check(form.DialogResult != DialogResult.OK && (string)Get(Get(form, "Result"), "Username") == "", "Incomplete onboarding cannot report or persist success");
                Check(((TabControl)Field(form, "_tabControl")).TabPages.Count == (object.Equals(ribbon, false) ? 1 : 8), "Onboarding does not change ribbon-based settings visibility");
                Check((bool)Get(Get(form, "Result"), "TransportTlsEnable12") == (bool)Get(local, "TransportTlsEnable12")
                    && (bool)Get(Get(form, "Result"), "TransportTlsEnable13") == (bool)Get(local, "TransportTlsEnable13")
                    && (bool)Get(Get(form, "Result"), "TransportTlsUseSystemDefault") == (bool)Get(local, "TransportTlsUseSystemDefault"), "Onboarding preserves configured TLS options");
            }
        }
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        {
            object local = ManagedSettings("https://cloud.example.test", true);
            Set(local, "Username", "onboarding-test");
            Set(local, "AppPassword", "test-only");
            object status = Status(D(), D(), true, mode, seat);
            using (Form form = Settings(local, status))
            {
                Check(((Label)Field(form, "_policyWarningTextLabel")).Text == RolloutNotice(local, status), "Complete setup retains the actual backend/seat notice");
                Call(form, "BeginAuthentication", false);
                Check((bool)Field(form, "_connectionSetupPending"), "Explicit reauthentication requires a new verification");
                object failure = New("Services.TalkServiceException", "test rejection", true, System.Net.HttpStatusCode.Unauthorized, "", false);
                Call(form, "HandleServiceFailure", "{0}", failure);
                Check((bool)Field(form, "_authenticationRejected") && Field(form, "_backendPolicyStatus") == null, "Authentication failure clears old account policy");
                Check(((Label)Field(form, "_policyWarningTextLabel")).Text == rejected, "Authentication failure is not presented as a seat failure");
                Call(form, "BeginAuthentication", false);
                failure = New("Services.TalkServiceException", "server unavailable", false, System.Net.HttpStatusCode.ServiceUnavailable, "", false);
                Call(form, "HandleServiceFailure", "{0}", failure);
                Check(!(bool)Field(form, "_authenticationRejected"), "Server failure is not classified as rejected credentials");
                form.Dispose();
                Call(form, "SetBusy", false);
                Call(form, "SetStatus", "late result", false);
                Call(form, "HandleServiceFailure", "{0}", failure);
                Check(form.IsDisposed, "Late network completion cannot reopen a cancelled settings form");
            }
        }
        Console.WriteLine("[OK] Friendly managed/local onboarding, incomplete saves, credential changes, seat parity and closed-form callbacks");
    }
    [STAThread]
    public static int Main(string[] args)
    {
        Product = Assembly.LoadFrom(args[0]);
        string root = args[1];
        try {
            TestLocalChoices(root);
            TestWizards();
            TestSettingsEdits(root);
            TestAttachmentAutomation();
            TestEnterpriseRollout(root);
            TestConnectionOnboarding();
            TestManagedTls(root);
            Console.WriteLine("[OK] " + Checks + " production policy/persistence/UI assertions passed");
            return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex.ToString()); return 1; }
    }
}
'@ | Set-Content -LiteralPath $uiSource -Encoding UTF8
    $uiExe = Join-Path $TempRoot 'OutlookPolicyUiTests.exe'
    & $csc /nologo /target:exe "/out:$uiExe" /reference:System.dll /reference:System.Core.dll /reference:System.Xml.dll /reference:System.Windows.Forms.dll $uiSource
    if ($LASTEXITCODE -ne 0) { throw 'Policy UI test harness compilation failed.' }
    & $uiExe (Join-Path $uiOutput 'NcTalkOutlookAddIn.dll') $TempRoot
    if ($LASTEXITCODE -ne 0) { throw 'Production policy/persistence/UI tests failed.' }

    # Exercise the production workflow with dialog/network doubles; never touch a user's profile.
    $workflowSource = Join-Path $TempRoot 'SettingsWorkflowTests.cs'
    @'
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using System.Windows.Forms;
namespace Microsoft.Office.Interop.Outlook { public class Application {} }
namespace NcTalkOutlookAddIn.Models { public class BackendPolicyStatus {} }
namespace System.Windows.Forms {
    public enum DialogResult { Cancel, OK }
    public enum MessageBoxButtons { OK }
    public enum MessageBoxIcon { Error }
    public static class MessageBox {
        public static int Calls;
        public static void Show(object owner, string text, string title, MessageBoxButtons buttons, MessageBoxIcon icon) { Calls++; }
    }
}
namespace NcTalkOutlookAddIn.Settings {
    public class AddinSettings {
        public string ServerUrl = "https://example.test", Username = "old", AppPassword = "test-only";
        public bool HasManagedNextcloudUrl, ManagedNextcloudUrlLocked, IfbEnabled, DebugLoggingEnabled, LogAnonymizationEnabled;
        public string ManagedNextcloudUrlSource = "test", AuthMode = "manual";
        public int IfbCacheHours = 24, IfbPort = 5000;
        public AddinSettings Clone() { return (AddinSettings)MemberwiseClone(); }
        public void ApplyManagedSetupPolicy(object policy) {}
    }
    public static class ManagedSetupPolicy { public static object Load() { return null; } }
}
namespace NcTalkOutlookAddIn.Services {
    public class TalkServiceConfiguration {
        private readonly bool complete;
        public TalkServiceConfiguration(string url, string user, string password) { complete = url.Length > 0 && user.Length > 0 && password.Length > 0; }
        public bool IsComplete() { return complete; }
    }
    public class IfbAddressBookCache {
        public class SystemAddressbookStatus {}
        public IfbAddressBookCache(string directory, string profile) {}
        public SystemAddressbookStatus GetSystemAddressbookStatus(TalkServiceConfiguration config, int hours, bool force) { throw new Exception("Unexpected pre-authentication address-book request"); }
    }
}
namespace NcTalkOutlookAddIn.Utilities {
    public static class LogCategories { public const string Core = "core"; }
    public static class DiagnosticsLogger { public static void LogException(string category, string message, Exception ex) {} }
    public static class Strings { public const string SettingsSaveFailed = "save failed", SettingsFormTitle = "settings"; }
}
namespace NcTalkOutlookAddIn.UI {
    public sealed class SettingsForm : IDisposable {
        public static DialogResult NextResult;
        public static bool AuthenticationStarted, Rejected, Disposed;
        public AddinSettings Result;
        public SettingsForm(AddinSettings current, object app, object policy, object cache, object book) {
            Result = current.Clone(); Result.Username = "new";
        }
        public void BeginAuthentication(bool rejected) { AuthenticationStarted = true; Rejected = rejected; }
        public DialogResult ShowDialog() { return NextResult; }
        public void Dispose() { Disposed = true; }
    }
}
internal static class SettingsWorkflowTests {
    private static int checks;
    private static void Check(bool value, string name) { checks++; if (!value) throw new Exception(name); }
    public static int Main() {
        try {
            foreach (string scenario in new[] { "cancel", "save", "persist-failure", "validate-failure", "commit-failure", "incomplete-cancel" })
            foreach (bool rejected in new[] { false, true }) {
                var current = new AddinSettings();
                bool requireAuth = scenario != "incomplete-cancel";
                if (!requireAuth) current.AppPassword = "";
                var events = new List<string>();
                SettingsForm.NextResult = scenario.EndsWith("cancel") ? DialogResult.Cancel : DialogResult.OK;
                SettingsForm.AuthenticationStarted = SettingsForm.Rejected = SettingsForm.Disposed = false;
                MessageBox.Calls = 0;
                var workflow = new SettingsWorkflowController(null,
                    () => current,
                    next => { events.Add("runtime:" + next.Username); current = next; },
                    (config, source) => { throw new Exception("Unexpected pre-authentication backend request"); },
                    next => events.Add("diagnostics"),
                    (next, source, interactive) => {
                        events.Add(source);
                        return !(scenario == "validate-failure" && source == "settings_save_validate")
                            && !(scenario == "commit-failure" && source == "settings_save_commit");
                    },
                    () => events.Add("ifb"),
                    next => { events.Add("persist:" + next.Username); if (scenario == "persist-failure") throw new Exception("test write failure"); },
                    action => { events.Add("dispatch"); action(); return Task.FromResult(0); },
                    message => {}, "test-directory", "test-profile");
                bool saved = workflow.RunAsync(requireAuth, rejected).GetAwaiter().GetResult();
                Check(saved == (scenario == "save"), scenario + " success result");
                Check(SettingsForm.AuthenticationStarted == requireAuth && SettingsForm.Rejected == (requireAuth && rejected), "Authentication context reaches dialog");
                Check(SettingsForm.Disposed, "Dialog disposed on every outcome");
                Check(current.Username == (saved ? "new" : "old"), "Runtime state follows successful persistence only");
                Check(events.Contains("ifb") == saved, "IFB applies only after successful commit");
                Check(MessageBox.Calls == (scenario == "persist-failure" ? 1 : 0), "Only a write error displays the persistence error");
                if (scenario.EndsWith("cancel")) Check(events.Count == 1, "Cancellation has no persistence or runtime side effects");
                if (scenario == "save") Check(events.IndexOf("persist:new") < events.IndexOf("runtime:new") && events.IndexOf("settings_save_commit") < events.IndexOf("ifb"), "Persistence precedes runtime and IFB");
                if (scenario == "validate-failure") Check(!events.Contains("persist:new"), "Validation failure does not save");
                if (scenario == "commit-failure") Check(events.Contains("persist:old") && events.Contains("settings_save_commit_revert"), "Failed commit restores previous settings");
            }
            Console.WriteLine("[OK] " + checks + " production settings-workflow save/cancel/failure assertions passed");
            return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
'@ | Set-Content -LiteralPath $workflowSource -Encoding UTF8
    $workflowExe = Join-Path $TempRoot 'SettingsWorkflowTests.exe'
    & $csc /noconfig /nologo /target:exe "/out:$workflowExe" /reference:System.dll /reference:System.Core.dll $workflowSource (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Controllers/SettingsWorkflowController.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Settings workflow test harness compilation failed.' }
    & $workflowExe
    if ($LASTEXITCODE -ne 0) { throw 'Production settings workflow tests failed.' }
}
finally {
    if (Test-Path $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}

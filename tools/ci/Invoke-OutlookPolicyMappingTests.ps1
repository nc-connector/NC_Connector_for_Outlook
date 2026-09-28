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
        return T(name).GetConstructors(Flags).Single(c => c.GetParameters().Length == args.Length).Invoke(args);
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
}
finally {
    if (Test-Path $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}

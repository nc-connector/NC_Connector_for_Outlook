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
        internal string ResolveDefaultsSource(BackendPolicyStatus status) { return "local"; }
        internal AddinSettings ResolvePolicyDefaults(BackendPolicyStatus status)
        {
            return new AddinSettings {
                EmailSignatureOnCompose = EmailSignaturePolicyService.ResolveFlag(status, "email_signature_on_compose", EmailSignatureOnCompose),
                EmailSignatureOnReply = EmailSignaturePolicyService.ResolveFlag(status, "email_signature_on_reply", EmailSignatureOnReply),
                EmailSignatureOnForward = EmailSignaturePolicyService.ResolveFlag(status, "email_signature_on_forward", EmailSignatureOnForward)
            };
        }
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
        internal IDictionary<string, string> Headers { get; set; }
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
        foreach (string state in new[] { "pending", "revoked", "suspended", "unknown" })
        {
            BackendPolicyStatus unavailable = ParseLicense("ACTIVE", "ACTIVE", true, true, state);
            string message = PolicyUiHelper.GetPolicyWarningMessage(unavailable);
            Check("Other seat states never claim suspension: " + state, message.StartsWith(Strings.PolicyWarningSeatUnavailable, StringComparison.Ordinal)
                && !message.Contains(Strings.PolicyWarningSeatSuspended)
                && PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(unavailable) == message);
        }

        BackendPolicyStatus emptySeatState = ParseLicense("ACTIVE", "ACTIVE", true, true, "");
        Check("Empty seat state is an invalid response, not a confirmed seat refusal", !emptySeatState.FetchSucceeded
            && emptySeatState.Reason == "invalid_payload" && !PolicyUiHelper.HasBackendSeatEntitlement(emptySeatState));

        NcHttpClient.NextResponse = new NcHttpResponse { HasHttpResponse = true, StatusCode = HttpStatusCode.NotFound };
        BackendPolicyStatus missingBackend = new BackendPolicyService(new TalkServiceConfiguration()).FetchStatus();
        Check("Missing backend remains a failed availability check without a license or seat refusal", !missingBackend.EndpointAvailable && !missingBackend.FetchSucceeded
            && missingBackend.Reason == "backend_unavailable" && missingBackend.IsServiceUnavailable
            && !missingBackend.PolicyActive && !PolicyUiHelper.HasBackendSeatEntitlement(missingBackend)
            && missingBackend.LicenseStatus == string.Empty && missingBackend.AccessStatus == string.Empty && missingBackend.SeatState == string.Empty
            && PolicyUiHelper.GetPolicyWarningMessage(missingBackend) == string.Empty
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
    private static object Status(Dictionary<string, object> share, Dictionary<string, object> talk, bool editable, string mode, string seat,
        object defaultsSource = null, object defaultsSourceEditable = null, Dictionary<string, object> signature = null)
    {
        var shareEdit = share.ToDictionary(p => p.Key, p => (object)editable);
        var talkEdit = talk.ToDictionary(p => p.Key, p => (object)editable);
        var payload = D(
            "status", D("is_valid", seat != "invalid", "seat_assigned", seat != "none", "seat_state", seat == "paused" ? "suspended_overlimit" : "active", "mode", mode, "overlicensed", true),
            "policy", D("share", share, "talk", talk, "email_signature", signature),
            "policy_editable", D("share", shareEdit, "talk", talkEdit, "email_signature", signature == null ? null : signature.ToDictionary(p => p.Key, p => (object)editable)));
        if (defaultsSource != null) payload["defaults_source"] = defaultsSource;
        if (defaultsSourceEditable != null) payload["defaults_source_editable"] = defaultsSourceEditable;
        return Call(T("Services.BackendPolicyService"), "ParseStatus", payload);
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
    private static object PolicyFetcher(object status, object owner = null)
    {
        return typeof(OutlookPolicyUiTests).GetMethod("TypedPolicyFetcher", BindingFlags.Static | BindingFlags.NonPublic)
            .MakeGenericMethod(T("Services.TalkServiceConfiguration"), T("Models.BackendPolicyStatus"))
            .Invoke(null, new[] { status, owner });
    }
    private static object TypedPolicyFetcher<TConfig, TStatus>(object status, object owner)
    {
        return new Func<TConfig, string, TStatus>((configuration, trigger) => {
            if (owner != null && status != null && (bool)Get(status, "FetchSucceeded"))
                Call(owner, "StoreBackendPolicySnapshotIfCurrent", configuration, status, trigger, 2L);
            return (TStatus)status;
        });
    }
    private static Form Settings(object local, object status) { return (Form)New("UI.SettingsForm", local, null, status, null, Addressbook(), PolicyFetcher(status)); }
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
    private static void TestDefaultsSourceMetadata()
    {
        foreach (bool wrapped in new[] { false, true })
        foreach (object source in new object[] { null, "inherit", "unknown", "", "local", "backend", true, 1 })
        foreach (object editable in new object[] { null, false, true, "true", 1 })
        {
            var payload = D("status", D("is_valid", true, "seat_assigned", true, "seat_state", "active", "mode", "community"),
                "policy", D("share", D()), "policy_editable", D("share", D()), "defaults_source", source, "defaults_source_editable", editable);
            object status = Call(T("Services.BackendPolicyService"), "ParseStatus", wrapped ? D("ocs", D("data", payload)) : payload);
            string expected = object.Equals(source, "local") || object.Equals(source, "backend") ? (string)source : null;
            Check(object.Equals(Get(status, "DefaultsSource"), expected), "Only explicit valid top-level defaults source is retained");
            Check((bool)Get(status, "DefaultsSourceEditable") == (expected != null && object.Equals(editable, true)), "Only a boolean source permission accompanies a valid source");
            Check((bool)Get(status, "PolicyActive"), "Optional source metadata does not change seat access");
        }
        object nested = Call(T("Services.BackendPolicyService"), "ParseStatus", D("status", D("is_valid", true, "seat_assigned", true,
            "seat_state", "active", "mode", "pro", "defaults_source", "backend", "defaults_source_editable", false), "policy", D("share", D())));
        Check(Get(nested, "DefaultsSource") == null && !(bool)Get(nested, "DefaultsSourceEditable"), "Nested source metadata does not replace the top-level API fields");
        object legacy = Status(D(), D(), true, "pro", "active");
        Check(Get(legacy, "DefaultsSource") == null && !(bool)Get(legacy, "DefaultsSourceEditable"), "Old backend responses retain the local source default");
        Console.WriteLine("[OK] Optional top-level source metadata, legacy payloads and plain/OCS parsing");
    }
    private static void TestDefaultsSourcePrecedence(string root)
    {
        int cases = 0;
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (string registry in new string[] { null, "local", "backend", "invalid", "" })
        foreach (string backend in new string[] { null, "inherit", "unknown", "local", "backend" })
        foreach (bool editable in new[] { false, true })
        foreach (string user in new string[] { null, "local", "backend" })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "DefaultsSource", user);
            Call(local, "ApplyManagedSetupPolicy", ManagedDefaultsSource(registry));
            object status = Status(D(), D(), true, mode, seat, backend, editable);
            bool hasBackend = backend == "local" || backend == "backend";
            string expected = seat != "active" ? "local" : hasBackend ? (editable && user != null ? user : backend)
                : registry != null ? (registry == "backend" ? "backend" : "local") : user ?? "local";
            bool canEdit = seat == "active" && (hasBackend ? editable : registry == null);
            Check((string)Call(local, "ResolveDefaultsSource", status) == expected, "Source precedence case " + cases);
            Check((bool)Call(local, "CanEditDefaultsSource", status) == canEdit, "Source permission case " + cases);
            Check(object.Equals(Get(local, "DefaultsSource"), user), "Source resolution preserves raw user preference");
            Check((bool)Get(local, "HasManagedDefaultsSource") == (registry != null), "Registry presence is independent of source validity");
            Check((bool)Get(local, "IsEnterpriseRollout") == (registry != null), "Every present source policy activates rollout");
            cases++;
        }
        object active = Status(D(), D(), true, "community", "active");
        foreach (string user in new string[] { null, "local", "backend" })
        foreach (string registry in new string[] { null, "local", "backend", "bad" })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "DefaultsSource", user);
            Call(local, "ApplyManagedSetupPolicy", ManagedDefaultsSource(registry));
            string xml = Serialize(local);
            Check(user == null ? !xml.Contains("<DefaultsSource") : xml.Contains("<DefaultsSource>" + user + "</DefaultsSource>"), "XML writes only the explicit user source");
            Check(!xml.Contains("ManagedDefaultsSource") && !xml.Contains("EnterpriseRollout"), "XML omits managed source state");
            object clone = Call(local, "Clone");
            Check(Serialize(clone) == xml && object.Equals(Get(clone, "DefaultsSource"), user)
                && (bool)Get(clone, "HasManagedDefaultsSource") == (registry != null), "Clone retains source overlay and raw preference separately");
            Check((bool)Get(clone, "IsManagedDefaultsSourceValid") == (registry != "bad"), "Clone retains managed source validity");
            object restored = RoundTrip(local, root);
            Check(object.Equals(Get(restored, "DefaultsSource"), user) && !(bool)Get(restored, "HasManagedDefaultsSource"), "XML reload never turns managed source into user choice");
            Call(clone, "ApplyManagedSetupPolicy", (object)null);
            Check((string)Call(clone, "ResolveDefaultsSource", active) == (user ?? "local") && !(bool)Get(clone, "HasManagedDefaultsSource"), "Policy removal restores the raw user source");
            Check((bool)Get(local, "HasManagedDefaultsSource") == (registry != null), "Removing a cloned policy does not alter its original");
            Check((string)Call(local, "ResolveDefaultsSource", (object)null) == "local" && !(bool)Call(local, "CanEditDefaultsSource", (object)null), "Missing backend falls back locally without granting source access");
        }
        Console.WriteLine("[OK] " + cases + " source precedence/seat combinations plus raw XML, clone and removal checks");
    }
    private static void TestDefaultsSourceValues()
    {
        object zip = Enum.Parse(T("Models.AttachmentLinkTarget"), "ZipDownload");
        object sharePage = Enum.Parse(T("Models.AttachmentLinkTarget"), "SharePage");
        foreach (string source in new[] { "local", "backend" })
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (bool editable in new[] { false, true })
        foreach (bool backend in new[] { false, true })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "DefaultsSource", source);
            foreach (string name in new[] { "SharingDefaultPermCreate", "TalkDefaultLobbyEnabled", "EmailSignatureOnCompose", "EmailSignatureOnReply", "EmailSignatureOnForward" }) Set(local, name, !backend);
            Set(local, "SharingDefaultExpireDays", 0);
            Set(local, "SharingAttachmentLinkTarget", zip);
            Set(local, "ShareBlockLang", "de");
            Set(local, "EventDescriptionLang", "en");
            object status = Status(D("share_permission_upload", backend, "share_expire_days", 19, "attachment_link_target", "share_page", "language_share_html_block", "fr"),
                D("talk_lobby_active", backend, "language_talk_description", "it"), editable, mode, seat, null, null,
                D("email_signature_on_compose", backend, "email_signature_on_reply", backend, "email_signature_on_forward", backend,
                    "email_signature_template", "<p>Test signature</p>", "user_email", "user@example.test"));
            string before = Serialize(local);
            bool useBackend = seat == "active" && (!editable || source == "backend");
            object effective = Resolve(local, status);
            Check((bool)Get(effective, "SharingDefaultPermCreate") == (useBackend ? backend : !backend)
                && (bool)Get(effective, "TalkDefaultLobbyEnabled") == (useBackend ? backend : !backend), "Backend preference retains explicit false and locked-value precedence");
            Check((int)Get(effective, "SharingDefaultExpireDays") == (useBackend ? 19 : 0), "Local source preserves explicit zero expiration");
            Check(object.Equals(Get(effective, "SharingAttachmentLinkTarget"), useBackend ? sharePage : zip), "Attachment target uses the selected source and existing field lock");
            Check((string)Get(effective, "ShareBlockLang") == (useBackend ? "fr" : "de")
                && (string)Get(effective, "EventDescriptionLang") == (useBackend ? "it" : "en"), "Generated block languages share source resolution");
            object signature = Call(New("Services.EmailSignaturePolicyService", status, local), "Resolve");
            foreach (string name in new[] { "OnCompose", "OnReply", "OnForward" })
                Check((bool)Get(signature, name) == (seat == "active" && (useBackend ? backend : !backend)), "Signature flag source/seat parity: " + name);
            Check((bool)Get(signature, "Active") == (seat == "active" && (useBackend ? backend : !backend)), "Source metadata never grants inactive-seat signatures");
            Check(Serialize(local) == before, "Effective source values never replace raw false/zero/language settings");
            object missing = Resolve(local, Status(D(), D(), true, mode, "active"));
            Check((int)Get(missing, "SharingDefaultExpireDays") == 0 && (bool)Get(missing, "SharingDefaultPermCreate") == !backend,
                "Backend preference falls back to local when a value is absent");
        }
        Console.WriteLine("[OK] Source-aware false/zero values, field locks, signatures, attachment targets and languages");
    }
    private static void TestDefaultsSourceUi(string root, string previewPath)
    {
        string[] blockedTabs = { "_fileLinkTab", "_talkTab", "_signatureTab" };
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (string source in new[] { "local", "backend" })
        foreach (bool editable in new[] { false, true })
        {
            object local = New("Settings.AddinSettings");
            object status = Status(D("share_permission_upload", true), D("talk_lobby_active", true), true, mode, seat, source, editable);
            string before = Serialize(local);
            using (Form options = Settings(local, status))
            {
                var combo = (ComboBox)Field(options, "_defaultsSourceCombo");
                var tabs = (TabControl)Field(options, "_tabControl");
                var advanced = (TabPage)Field(options, "_advancedTab");
                bool backend = seat == "active" && source == "backend";
                Check(advanced.Contains(combo) && advanced.Contains((Control)Field(options, "_defaultsSourceLabel"))
                    && advanced.Contains((Control)Field(options, "_defaultsSourceHintLabel")), "Advanced owns the source selector and explanation");
                Check(combo.Items.Count == 2 && combo.SelectedIndex == (backend ? 1 : 0)
                    && combo.Enabled == (seat == "active" && editable), "Source selector shows effective source and personal edit permission");
                Check(!string.IsNullOrWhiteSpace(((Label)Field(options, "_defaultsSourceHintLabel")).Text), "Source selection has a visible reason/help hint");
                foreach (string name in blockedTabs)
                {
                    var page = (TabPage)Field(options, name);
                    Check(page.Enabled == !backend, "Only backend defaults disable the defaults tab: " + name);
                    var selecting = new TabControlCancelEventArgs(page, tabs.TabPages.IndexOf(page), false, TabControlAction.Selecting);
                    Call(tabs, "OnSelecting", selecting);
                    Check(selecting.Cancel == backend, "Mouse/keyboard/programmatic tab selection shares the disabled-tab guard");
                    if (backend) Check(!string.IsNullOrWhiteSpace(page.ToolTipText) && page.AccessibleDescription == page.ToolTipText, "Disabled tabs expose tooltip and accessible reason");
                }
                foreach (TabPage page in tabs.TabPages)
                    if (!blockedTabs.Any(name => ReferenceEquals(page, Field(options, name)))) Check(page.Enabled, "Source policy leaves other settings tabs accessible");
                Check(tabs.DrawMode == (backend ? TabDrawMode.OwnerDrawFixed : TabDrawMode.Normal), "Only disabled defaults tabs need gray header drawing");
                Call(options, "SetBusy", true);
                Check(!combo.Enabled, "Busy source control is disabled");
                Call(options, "SetBusy", false);
                Check(combo.Enabled == (seat == "active" && editable), "Busy changes preserve the source lock");
                Check(Serialize(Get(options, "Result")) == before && Serialize(local) == before, "Opening source controls never stores effective defaults");
            }
        }
        foreach (string registry in new[] { "backend", "bad" })
        {
            object local = New("Settings.AddinSettings");
            Call(local, "ApplyManagedSetupPolicy", ManagedDefaultsSource(registry));
            using (Form options = Settings(local, Status(D(), D(), true, "pro", "active", "inherit", true)))
                Check(!((ComboBox)Field(options, "_defaultsSourceCombo")).Enabled, "Inherit metadata cannot unlock valid or invalid registry policy");
        }
        foreach (string mode in new[] { "community", "pro" })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "SharingDefaultPermCreate", false);
            Set(local, "SharingDefaultExpireDays", 0);
            Set(local, "TalkDefaultLobbyEnabled", false);
            Set(local, "EmailSignatureOnCompose", false);
            Set(local, "ShareBlockLang", "de");
            object status = Status(D("share_permission_upload", true, "share_expire_days", 19, "language_share_html_block", "fr"),
                D("talk_lobby_active", true), true, mode, "active", "backend", true);
            using (Form options = Settings(local, status))
            {
                object result = Get(options, "Result");
                var combo = (ComboBox)Field(options, "_defaultsSourceCombo");
                var tabs = (TabControl)Field(options, "_tabControl");
                tabs.SelectedTab = (TabPage)Field(options, "_advancedTab");
                foreach (int index in new[] { 0, 1, 0, 1 })
                {
                    combo.SelectedIndex = index;
                    Call(combo, "OnSelectionChangeCommitted", EventArgs.Empty);
                    Check((string)Get(result, "DefaultsSource") == (index == 1 ? "backend" : "local"), "Committed source selection records only the raw source choice");
                    Check(((CheckBox)Field(options, "_sharingDefaultPermCreateCheckBox")).Checked == (index == 1)
                        && ((CheckBox)Field(options, "_talkDefaultLobbyCheckBox")).Checked == (index == 1), "Source selection refreshes visible defaults immediately");
                    Check((string)Call(T("UI.SettingsForm"), "GetSelectedLanguageChoice", Field(options, "_shareBlockLangCombo")) == (index == 1 ? "fr" : "de"), "Source selection refreshes language defaults immediately");
                    Check(!(bool)Get(result, "SharingDefaultPermCreate") && (int)Get(result, "SharingDefaultExpireDays") == 0
                        && !(bool)Get(result, "TalkDefaultLobbyEnabled") && !(bool)Get(result, "EmailSignatureOnCompose"), "Toggling source preserves raw false and zero choices");
                    foreach (string name in blockedTabs) Check(((TabPage)Field(options, name)).Enabled == (index == 0), "Source selection immediately changes tab availability");
                }
                Task save = (Task)Call(options, "SaveSettingsAsync");
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);
                while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Check(save.IsCompleted, "Source-only save completes without a configured server");
                save.GetAwaiter().GetResult();
                Check(options.DialogResult == DialogResult.OK, "Source-only save succeeds");
                object restored = RoundTrip(result, root);
                Check((string)Get(restored, "DefaultsSource") == "backend" && !(bool)Get(restored, "SharingDefaultPermCreate")
                    && (int)Get(restored, "SharingDefaultExpireDays") == 0 && !(bool)Get(restored, "EmailSignatureOnCompose"), "Saving source preserves dormant raw local defaults");
                Set(restored, "DefaultsSource", "local");
                Check(!(bool)Get(Resolve(restored, status), "SharingDefaultPermCreate") && (int)Get(Resolve(restored, status), "SharingDefaultExpireDays") == 0,
                    "Returning to local after save restores false and zero values");
            }
        }
        using (Form preview = Settings(New("Settings.AddinSettings"), Status(D(), D(), true, "community", "active", "backend", true)))
        {
            ((TabControl)Field(preview, "_tabControl")).SelectedTab = (TabPage)Field(preview, "_advancedTab");
            Directory.CreateDirectory(Path.GetDirectoryName(previewPath));
            preview.Show(); Application.DoEvents();
            using (var bitmap = new System.Drawing.Bitmap(preview.Width, preview.Height))
            {
                preview.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height));
                bitmap.Save(previewPath, System.Drawing.Imaging.ImageFormat.Png);
            }
        }
        Console.WriteLine("[OK] Real source controls, tab guards/tooltips, live switching and raw-value save; preview: " + previewPath);
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
        foreach (string source in new[] { "local", "backend" })
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (bool editable in new[] { false, true })
        foreach (bool explicitChoice in new[] { false, true })
        {
            string property = (string)binding[2];
            object local = New("Settings.AddinSettings");
            Set(local, "DefaultsSource", source);
            object productDefault = Get(local, property);
            if (explicitChoice) Set(local, property, binding[4]);
            string saved = Serialize(local);
            var policy = D(binding[1], binding[3]);
            object status = Status((string)binding[0] == "share" ? policy : D(), (string)binding[0] == "talk" ? policy : D(), editable, mode, seat);
            object effective = Resolve(local, status);
            object expected = seat == "active" && (!editable || !explicitChoice || source == "backend") ? binding[5] : explicitChoice ? binding[4] : productDefault;
            Check(object.Equals(Get(effective, property), expected), "Resolver precedence: " + property + "/" + source + "/" + mode + "/" + seat + "/" + editable + "/" + explicitChoice);
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
    private static void TestSettingsLanguageControls(string root)
    {
        string[][] bindings = {
            new[] { "_shareBlockLangCombo", "_shareBlockLangLabel", "_fileLinkTab", "ShareBlockLang", "en", "fr", "nl" },
            new[] { "_eventDescriptionLangCombo", "_eventDescriptionLangLabel", "_talkTab", "EventDescriptionLang", "de", "it", "es" }
        };
        foreach (string mode in new[] { "community", "pro" })
        {
            object local = New("Settings.AddinSettings");
            foreach (string[] binding in bindings) Set(local, binding[3], binding[4]);
            object editable = Status(D("language_share_html_block", "fr"), D("language_talk_description", "it"), true, mode, "active");
            object locked = Status(D("language_share_html_block", "fr"), D("language_talk_description", "it"), false, mode, "active");
            using (Form options = Settings(local, editable))
            {
                object result = Get(options, "Result");
                var advanced = (Control)Field(options, "_advancedTab");
                foreach (string[] binding in bindings)
                {
                    var combo = (ComboBox)Field(options, binding[0]);
                    var label = (Label)Field(options, binding[1]);
                    var tab = (Control)Field(options, binding[2]);
                    Check(tab.Contains(combo) && tab.Contains(label), "Language controls belong to their feature tab: " + binding[3]);
                    Check(!advanced.Contains(combo) && !advanced.Contains(label), "Advanced no longer owns language controls: " + binding[3]);
                    Check(combo.Enabled && (string)Call(T("UI.SettingsForm"), "GetSelectedLanguageChoice", combo) == binding[4], "Moved language control loads the local choice: " + binding[3]);
                }
                foreach (int width in new[] { 800, 1100 })
                {
                    options.ClientSize = new System.Drawing.Size(width, options.ClientSize.Height);
                    Call(options, "ApplyResponsiveLayout", false);
                    foreach (string[] binding in bindings)
                    {
                        var combo = (ComboBox)Field(options, binding[0]);
                        var label = (Label)Field(options, binding[1]);
                        Check(label.Bottom < combo.Top && combo.Right <= combo.Parent.ClientSize.Width,
                            "Language label and selection fit without overlap: " + binding[3]);
                    }
                    Check(((Control)Field(options, "_shareBlockLangCombo")).Bottom < ((Control)Field(options, "_sharingAttachmentAutomationGroup")).Top,
                        "Sharing language stays above attachment automation");
                    Check(((Control)Field(options, "_eventDescriptionLangCombo")).Bottom < ((Control)Field(options, "_talkDefaultsGroup")).ClientSize.Height,
                        "Talk defaults include the complete language selection");
                }
                options.GetType().GetField("_backendPolicyStatus", Flags).SetValue(options, locked);
                Call(options, "ApplyBackendPolicyStatus", "language_test_lock");
                foreach (string[] binding in bindings)
                {
                    var combo = (ComboBox)Field(options, binding[0]);
                    Check(!combo.Enabled && (string)Call(T("UI.SettingsForm"), "GetSelectedLanguageChoice", combo) == binding[5], "Moved language control retains its policy lock: " + binding[3]);
                    Check((string)Get(result, binding[3]) == binding[4], "Language policy overlay preserves the saved local choice: " + binding[3]);
                }
                options.GetType().GetField("_backendPolicyStatus", Flags).SetValue(options, editable);
                Call(options, "ApplyBackendPolicyStatus", "language_test_unlock");
                foreach (string[] binding in bindings)
                {
                    var combo = (ComboBox)Field(options, binding[0]);
                    Check(combo.Enabled && (string)Call(T("UI.SettingsForm"), "GetSelectedLanguageChoice", combo) == binding[4], "Language unlock restores the local choice: " + binding[3]);
                    Call(T("UI.SettingsForm"), "SelectLanguageChoice", combo, binding[6]);
                    Call(combo, "OnSelectionChangeCommitted", EventArgs.Empty);
                    Check((string)Get(result, binding[3]) == binding[6], "Moved language control records committed selection: " + binding[3]);
                }
                Task save = (Task)Call(options, "SaveSettingsAsync");
                DateTime deadline = DateTime.UtcNow.AddSeconds(5);
                while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Check(save.IsCompleted, "Language-only settings save completes without a configured server");
                save.GetAwaiter().GetResult();
                Check(options.DialogResult == DialogResult.OK, "Language-only settings save succeeds");
                object restored = RoundTrip(result, root);
                foreach (string[] binding in bindings)
                    Check((string)Get(restored, binding[3]) == binding[6], "Moved language choice survives settings save and XML round-trip: " + binding[3]);
            }
        }
        Console.WriteLine("[OK] Talk and Sharing language placement, local selections, policy locks and settings save");
    }
    private static void TestAttachmentAutomation()
    {
        Type subscription = T("NextcloudTalkAddIn+MailComposeSubscription");
        int cases = 0;
        foreach (string source in new[] { "local", "backend" })
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (bool editable in new[] { false, true })
        foreach (bool explicitChoice in new[] { false, true })
        foreach (bool always in new[] { false, true })
        foreach (object threshold in new object[] { null, 0, 1, 19, 10240 })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "DefaultsSource", source);
            if (explicitChoice) {
                Set(local, "SharingAttachmentsAlwaysConnector", false);
                Set(local, "SharingAttachmentsOfferAboveEnabled", false);
                Set(local, "SharingAttachmentsOfferAboveMb", 20);
            }
            object status = Status(D("attachments_always_via_ncconnector", always, "attachments_min_size_mb", threshold), D(), editable, mode, seat);
            object snapshot = Call(subscription, "BuildAttachmentAutomationSettings", local, local);
            object actual = Call(subscription, "ApplyAttachmentAutomationPolicy", snapshot, status);
            bool useBackend = seat == "active" && (!editable || !explicitChoice || source == "backend");
            bool expectedAlways = useBackend && always;
            bool expectedEnabled = !expectedAlways && (useBackend ? threshold != null : !explicitChoice);
            int expectedThreshold = useBackend && threshold != null ? ((int)threshold == 0 ? 5 : (int)threshold) : 20;
            Check((bool)Get(actual, "AlwaysConnector") == expectedAlways && (bool)Get(actual, "OfferAboveEnabled") == expectedEnabled && (int)Get(actual, "ThresholdMb") == expectedThreshold && (long)Get(actual, "ThresholdBytes") == expectedThreshold * 1024L * 1024L, "Operative attachment policy case " + cases);
            bool expectedMandatoryThreshold = seat == "active" && !editable && threshold != null && !expectedAlways;
            Check((bool)Get(actual, "ThresholdMandatory") == expectedMandatoryThreshold, "Only an applicable locked backend threshold creates mandatory routing: " + cases);
            cases++;
        }
        foreach (object threshold in new object[] { null, 0, 1, 19, 10240 }) {
            object status = Status(D("attachments_min_size_mb", threshold), D(), false, "pro", "active");
            using (Form options = Settings(New("Settings.AddinSettings"), status)) {
                Check(((CheckBox)Field(options, "_sharingAttachmentsOfferAboveCheckBox")).Checked == (threshold != null), "Threshold UI null semantics");
                Check(((NumericUpDown)Field(options, "_sharingAttachmentsOfferAboveMbUpDown")).Value == (threshold == null ? 20 : (int)threshold == 0 ? 5 : (int)threshold), "Threshold UI matches operative value");
            }
        }
        foreach (string mode in new[] { "community", "pro" })
        foreach (bool editable in new[] { false, true })
        foreach (object threshold in new object[] { null, 0, 1, 20 })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "SharingAttachmentsOfferAboveEnabled", true);
            Set(local, "SharingAttachmentsOfferAboveMb", 7);
            object rules = Call(subscription, "BuildAttachmentAutomationSettings", local, local);
            Check(!(bool)Get(rules, "ThresholdMandatory"), "A local optional size offer cannot become a central obligation");
            object status = Status(D("attachments_min_size_mb", threshold), D(), editable, mode, "active");
            Set(status, "FetchSucceeded", false);
            object unconfirmed = Call(subscription, "ApplyAttachmentAutomationPolicy", rules, status);
            Check(!(bool)Get(unconfirmed, "ThresholdMandatory"), "Failed initial backend check cannot invent a mandatory threshold");
        }
        Console.WriteLine("[OK] " + cases + " operative attachment combinations plus real threshold controls");
    }
    private static object Managed(object url, object locked, object ribbon, string source)
    {
        return New("Settings.ManagedSetupPolicy", url, locked, ribbon, null, null, null, null, null, null, null, null, null, null, null, null, null, source);
    }
    private static object ManagedTls(object system, object tls12, object tls13, string source = "TLS test")
    {
        return New("Settings.ManagedSetupPolicy", null, null, null, system, tls12, tls13, null, null, null, null, null, null, null, null, null, null, source);
    }
    private static object ManagedLogging(object debug, object anonymize, string source = "Logging test")
    {
        return New("Settings.ManagedSetupPolicy", null, null, null, null, null, null, debug, anonymize, null, null, null, null, null, null, null, null, source);
    }
    private static object ManagedUpdateNotify(object value)
    {
        return New("Settings.ManagedSetupPolicy", null, null, null, null, null, null, null, null, value, null, null, null, null, null, null, null, "Update notification test");
    }
    private static object ManagedIfb(object enabled, object days, object cacheHours, object port, string source = "IFB test")
    {
        return New("Settings.ManagedSetupPolicy", null, null, null, null, null, null, null, null, null, enabled, days, cacheHours, port, null, null, null, source);
    }
    private static object ManagedDefaultsSource(object value)
    {
        return New("Settings.ManagedSetupPolicy", null, null, null, null, null, null, null, null, null, null, null, null, null, value, null, null, "Defaults source test");
    }
    private static object ManagedAuthMode(object value, object url = null, object locked = null, object ribbon = null)
    {
        return New("Settings.ManagedSetupPolicy", url, locked, ribbon, null, null, null, null, null, null, null, null, null, null, null, value, null, "Authentication mode test");
    }
    private static object ManagedSendPolicyFailureMode(object value, string source = "Send policy test")
    {
        return New("Settings.ManagedSetupPolicy", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, value, source);
    }
    private static void TestManagedSendPolicyFailureMode(string root)
    {
        foreach (object value in new object[] { null, "failopen", "failclosed", " FAILOPEN ", " FAILCLOSED ", "", "invalid", "0", 0, false, new byte[] { 1 } })
        {
            string text = value as string;
            bool valid = value == null || string.Equals(text == null ? null : text.Trim(), "failopen", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text == null ? null : text.Trim(), "failclosed", StringComparison.OrdinalIgnoreCase);
            bool present = value != null;
            bool failClosed = valid && string.Equals(text == null ? null : text.Trim(), "failclosed", StringComparison.OrdinalIgnoreCase);
            string expected = failClosed ? "failclosed" : "failopen";
            object policy = ManagedSendPolicyFailureMode(value);
            Check((bool)Get(policy, "HasSendPolicyFailureModePolicy") == present
                && (bool)Get(policy, "IsSendPolicyFailureModePolicyValid") == valid
                && (string)Get(policy, "SendPolicyFailureMode") == expected, "Send failure mode distinguishes presence, validity and effective mode");
            object local = New("Settings.AddinSettings");
            Call(local, "ApplyManagedSetupPolicy", policy);
            Check((bool)Get(local, "SendPolicyFailClosed") == failClosed && (string)Get(local, "SendPolicyFailureMode") == expected
                && (bool)Get(local, "HasManagedSendPolicyFailureMode") == present
                && (bool)Get(local, "IsManagedSendPolicyFailureModeValid") == valid
                && (bool)Get(local, "IsEnterpriseRollout") == present, "Send policy presence activates rollout without changing default failopen");
            Check((string)Get(local, "ManagedSendPolicyFailureModeSource") == (present ? "Send policy test" : ""), "Send mode keeps its own registry source");
            string xml = Serialize(local);
            Check(!xml.Contains("SendPolicy"), "XML does not persist the administrative send failure mode");
            object restored = RoundTrip(local, root);
            Check(!(bool)Get(restored, "SendPolicyFailClosed") && !(bool)Get(restored, "HasManagedSendPolicyFailureMode"), "XML reload cannot make a managed send mode permanent");
            object clone = Call(local, "Clone");
            Check((bool)Get(clone, "SendPolicyFailClosed") == failClosed
                && (bool)Get(clone, "HasManagedSendPolicyFailureMode") == present
                && (bool)Get(clone, "IsManagedSendPolicyFailureModeValid") == valid, "Clone retains effective send mode and management metadata");
            Call(clone, "ApplyManagedSetupPolicy", (object)null);
            Check(!(bool)Get(clone, "SendPolicyFailClosed") && (string)Get(clone, "SendPolicyFailureMode") == "failopen"
                && !(bool)Get(clone, "HasManagedSendPolicyFailureMode"), "Removing send policy restores failopen without a user override");
            Check((bool)Get(local, "SendPolicyFailClosed") == failClosed, "Removing a cloned send policy does not alter the original");
        }
        Array policies = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 3);
        policies.SetValue(ManagedAuthMode("Manual", "https://cloud.example.test"), 0);
        policies.SetValue(ManagedSendPolicyFailureMode("invalid", "first send mode"), 1);
        policies.SetValue(ManagedSendPolicyFailureMode("failclosed", "lower send mode"), 2);
        object merged = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
        Check((string)Get(merged, "SendPolicyFailureMode") == "failopen" && !(bool)Get(merged, "IsSendPolicyFailureModePolicyValid")
            && (string)Get(merged, "SendPolicyFailureModeSource") == "first send mode" && (bool)Get(merged, "HasNextcloudUrl"), "Invalid selected send mode masks lower mode independently of URL/authentication");
        policies.SetValue(ManagedSendPolicyFailureMode("failclosed", "first send mode"), 1);
        policies.SetValue(ManagedSendPolicyFailureMode("invalid", "lower send mode"), 2);
        merged = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
        Check((string)Get(merged, "SendPolicyFailureMode") == "failclosed" && (bool)Get(merged, "IsSendPolicyFailureModePolicyValid"), "Unused invalid send mode cannot override a valid higher-priority mode");
    }
    private static void TestManagedAuthMode(string root)
    {
        object loginFlow = Enum.Parse(T("Settings.AuthenticationMode"), "LoginFlow");
        object manual = Enum.Parse(T("Settings.AuthenticationMode"), "Manual");
        foreach (object value in new object[] { null, "LoginFlow", "Manual", " loginflow ", " MANUAL ", "", "invalid", "0", 0, 1, false, new byte[] { 1 } })
        foreach (object raw in new[] { loginFlow, manual })
        {
            string text = value as string;
            bool valid = value == null || string.Equals(text == null ? null : text.Trim(), "LoginFlow", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text == null ? null : text.Trim(), "Manual", StringComparison.OrdinalIgnoreCase);
            bool present = value != null;
            object expected = present ? (valid && string.Equals(text.Trim(), "Manual", StringComparison.OrdinalIgnoreCase) ? manual : loginFlow) : raw;
            object policy = ManagedAuthMode(value);
            Check((bool)Get(policy, "HasAuthModePolicy") == present && (bool)Get(policy, "IsAuthModePolicyValid") == valid,
                "Auth mode preserves value presence separately from validity");
            object local = New("Settings.AddinSettings");
            Set(local, "AuthMode", raw);
            Call(local, "ApplyManagedSetupPolicy", policy);
            Check(object.Equals(Get(local, "AuthMode"), expected) && object.Equals(Get(local, "LocalAuthMode"), raw), "Auth mode overlay does not replace the local choice");
            Check((bool)Get(local, "HasManagedAuthMode") == present && (bool)Get(local, "IsManagedAuthModeValid") == valid
                && (bool)Get(local, "IsEnterpriseRollout") == present, "Auth mode presence alone activates managed rollout");
            Check((bool)Get(local, "ShowMainRibbonTab") && !(bool)Get(local, "HasManagedTransportTls") && (bool)Get(local, "IsManagedTransportTlsValid"),
                "Invalid authentication mode does not hide the ribbon or invalidate TLS");
            var xml = new XmlDocument(); xml.LoadXml(Serialize(local));
            Check(xml.DocumentElement["AuthMode"].InnerText == raw.ToString(), "XML persists raw authentication mode");
            object restored = RoundTrip(local, root);
            Check(object.Equals(Get(restored, "AuthMode"), raw) && !(bool)Get(restored, "HasManagedAuthMode"), "XML reload does not preserve a policy as local choice");
            object clone = Call(local, "Clone");
            Check(object.Equals(Get(clone, "AuthMode"), expected) && object.Equals(Get(clone, "LocalAuthMode"), raw)
                && (bool)Get(clone, "HasManagedAuthMode") == present && (bool)Get(clone, "IsManagedAuthModeValid") == valid, "Clone preserves auth mode overlay, raw choice and validity");
            Call(clone, "ApplyManagedSetupPolicy", (object)null);
            Check(object.Equals(Get(clone, "AuthMode"), raw) && !(bool)Get(clone, "HasManagedAuthMode"), "Policy removal restores the local authentication mode");
            Check((bool)Get(local, "HasManagedAuthMode") == present, "Removing cloned auth policy does not alter the original");
        }
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (object value in new object[] { null, "LoginFlow", "Manual", "invalid" })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "AuthMode", manual);
            Call(local, "ApplyManagedSetupPolicy", ManagedAuthMode(value));
            object status = Status(D(), D(), true, mode, seat);
            using (Form form = Settings(local, status))
            {
                var manualRadio = (RadioButton)Field(form, "_manualRadio");
                var flowRadio = (RadioButton)Field(form, "_loginFlowRadio");
                bool effectiveManual = value == null || object.Equals(value, "Manual");
                Check(manualRadio.Checked == effectiveManual && flowRadio.Checked == !effectiveManual, "Auth controls display effective managed mode");
                foreach (bool busy in new[] { false, true, false })
                {
                    Call(form, "SetBusy", busy); Call(form, "UpdateControlState");
                    Check(manualRadio.Enabled == (value == null && !busy) && flowRadio.Enabled == (value == null && !busy), "Both auth mode radios retain policy locks across busy changes");
                    Check(((Button)Field(form, "_loginFlowButton")).Enabled == (!effectiveManual && !busy), "Managed LoginFlow retains the explicit login button");
                    Check(((Button)Field(form, "_testButton")).Enabled == !busy, "Authentication mode alone does not block connection tests");
                }
                string hint = ((ToolTip)Field(form, "_toolTip")).GetToolTip(manualRadio);
                Check(string.IsNullOrEmpty(hint) == (value == null), "Managed authentication radios have a reachable hint");
                if (object.Equals(value, "invalid")) Check(hint == (string)T("Utilities.Strings").GetProperty("ManagedAuthModeInvalid", Flags).GetValue(null, null), "Invalid auth mode has a configuration hint");
                object previousTls = Call(form, "ApplySelectedTransportSecurity", "auth_mode_test");
                Call(T("UI.SettingsForm"), "RestoreTemporaryTls", previousTls, "auth_mode_test");
                if (value == null) flowRadio.Checked = true;
                Task save = (Task)Call(form, "SaveSettingsAsync");
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);
                while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Check(save.IsCompleted, "Auth mode settings save finishes without configured credentials"); save.GetAwaiter().GetResult();
                Check(object.Equals(Get(Get(form, "Result"), "LocalAuthMode"), value == null ? loginFlow : manual), "Actual save preserves managed raw mode and accepts unmanaged edits");
            }
            if (value != null) Check(RolloutNotice(local, status) == RolloutNotice(ManagedSettings(null, true), status), "Auth-mode-only rollout retains existing Community/Pro and personal seat rules");
        }
        Console.WriteLine("[OK] Managed authentication mode presence, validity, XML, clone/removal, live radios, save and seat parity");
    }
    private static void TestManagedLoginEligibility()
    {
        const string cloud = "https://cloud.example.test/nextcloud";
        object[][] cases = {
            new object[] { "LoginFlow", cloud, null, "", "", "", true },
            new object[] { " loginflow ", cloud, true, "", "", "", true },
            new object[] { "LoginFlow", cloud, false, cloud + "/", "", "", true },
            new object[] { "LoginFlow", cloud, false, "https://other.example.test", "", "", false },
            new object[] { "LoginFlow", cloud, true, "https://other.example.test", "", "", true },
            new object[] { "LoginFlow", cloud, null, "", "existing-user", "", true },
            new object[] { "LoginFlow", cloud, null, "", "", "existing-test-password", true },
            new object[] { "LoginFlow", cloud, null, "", "existing-user", "existing-test-password", false },
            new object[] { "Manual", cloud, true, "", "", "", false },
            new object[] { "invalid", cloud, true, "", "", "", false },
            new object[] { "", cloud, true, "", "", "", false },
            new object[] { 1, cloud, true, "", "", "", false },
            new object[] { null, cloud, true, "", "", "", false },
            new object[] { "LoginFlow", null, true, "", "", "", false },
            new object[] { "LoginFlow", null, true, cloud, "", "", false },
            new object[] { "LoginFlow", "", true, "", "", "", false },
            new object[] { "LoginFlow", "http://cloud.example.test", true, cloud, "", "", false },
            new object[] { "LoginFlow", "https://user:password@cloud.example.test", true, "", "", "", false },
            new object[] { null, null, true, cloud, "", "", false }
        };
        foreach (object[] values in cases)
        foreach (bool ribbon in new[] { false, true })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "ServerUrl", values[3]); Set(local, "Username", values[4]); Set(local, "AppPassword", values[5]);
            Call(local, "ApplyManagedSetupPolicy", ManagedAuthMode(values[0], values[1], values[2], ribbon));
            using (Form form = Settings(local, null))
            {
                Check(!(bool)Field(form, "_automaticLoginFlowPending"), "Opening normal settings never requests automatic login");
                Check((bool)Call(form, "ShouldStartManagedLoginFlow") == (bool)values[6], "Managed login eligibility matches mode, actual URL and credential presence");
                foreach (bool rejected in new[] { false, true })
                {
                    Call(form, "BeginAuthentication", rejected);
                    Check((bool)Field(form, "_automaticLoginFlowPending") == (bool)values[6], "Authentication-only entry schedules only an eligible managed LoginFlow");
                    Check(((TabControl)Field(form, "_tabControl")).TabPages.Count == (ribbon ? 8 : 1), "Auth policy retains ribbon-based dialog scope");
                }
                object result = Get(form, "Result");
                string[] names = { "LocalAuthMode", "ServerUrl", "Username", "AppPassword" };
                object[] before = names.Select(name => Get(result, name)).ToArray();
                form.Close(); form.Dispose();
                Check(names.Select(name => Get(result, name)).SequenceEqual(before), "Cancelling authentication does not change raw mode, URL or credentials");
                var xml = new XmlDocument(); xml.LoadXml(Serialize(result));
                Check(xml.DocumentElement["AuthMode"].InnerText == before[0].ToString(), "Cancelled authentication still serializes only the raw mode");
            }
        }
        Console.WriteLine("[OK] Automatic login eligibility: normal settings, URL-only/lock-only, invalid modes, mismatched URLs and existing credentials");
    }
    private static void CaptureManagedAuthPreview(string path)
    {
        object local = New("Settings.AddinSettings");
        Call(local, "ApplyManagedSetupPolicy", ManagedAuthMode("Manual", "https://cloud.example.test/nextcloud", true));
        using (Form form = Settings(local, null))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            form.Show(); Application.DoEvents();
            using (var bitmap = new System.Drawing.Bitmap(form.Width, form.Height))
            {
                form.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height));
                bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }
        }
    }
    private static readonly string[] IfbProperties = { "IfbEnabled", "IfbDays", "IfbCacheHours", "IfbPort" };
    private static void CheckIfb(object target, object[] values, string message, bool raw = false)
    {
        for (int i = 0; i < IfbProperties.Length; i++)
            Check(object.Equals(Get(target, (raw ? "Local" : "") + IfbProperties[i]), values[i]), message + ": " + IfbProperties[i]);
    }
    private static void TestManagedIfb(string root)
    {
        object[] defaults = { false, 30, 24, 7777 };
        object[][] validInputs = {
            new object[] { null, 0, 1, false, true, "false", "true" },
            new object[] { null, 10, 30, 60, 90 },
            new object[] { null, 1, 12, 24 },
            new object[] { null, 1024, 7777, 49151 }
        };
        object[][] invalidInputs = {
            new object[] { "", "bad", new byte[] { 1 } },
            new object[] { 0, 9, 11, 29, 31, 59, 61, 89, 91, -1, int.MaxValue, "30", 30L, 30U, 30.0, true, "", new byte[] { 30 } },
            new object[] { 0, 25, -1, int.MaxValue, "24", 24L, 24U, 24.0, true, "", new byte[] { 24 } },
            new object[] { 0, 1023, 49152, -1, int.MaxValue, "7777", 7777L, 7777U, 7777.0, true, "", new byte[] { 1 } }
        };
        for (int field = 0; field < IfbProperties.Length; field++)
        foreach (bool valid in new[] { true, false })
        foreach (object selected in valid ? validInputs[field] : invalidInputs[field])
        foreach (bool rawEnabled in new[] { false, true })
        foreach (bool decision in new[] { false, true })
        {
            object[] values = { null, null, null, null };
            values[field] = selected;
            object policy = ManagedIfb(values[0], values[1], values[2], values[3]);
            bool present = selected != null;
            Check((bool)Get(policy, "HasIfbPolicy") == present && (bool)Get(policy, "IsEnterpriseRollout") == present,
                "Any present IFB field activates rollout, including false, empty and invalid values");
            Check((bool)Get(policy, "IsIfbPolicyValid") == valid, "IFB validity rejects unsupported values without losing presence");
            object[] raw = { rawEnabled, 90, 3, 8888 };
            object[] effective = (object[])(present ? defaults : raw).Clone();
            if (present && valid) effective[field] = field == 0
                ? (object)(object.Equals(selected, 1) || object.Equals(selected, true) || object.Equals(selected, "true"))
                : selected;
            object local = New("Settings.AddinSettings");
            for (int i = 0; i < IfbProperties.Length; i++) Set(local, IfbProperties[i], raw[i]);
            Set(local, "IfbUserDecisionRecorded", decision);
            Call(local, "ApplyManagedSetupPolicy", policy);
            Check((bool)Get(local, "HasManagedIfb") == present && (bool)Get(local, "IsManagedIfbValid") == valid,
                "Settings retains IFB presence and validation separately");
            CheckIfb(local, effective, "Effective IFB uses selected fields and missing-sibling defaults");
            CheckIfb(local, raw, "Managed IFB preserves raw local values", true);
            Check((bool)Get(local, "IfbUserDecisionRecorded") == decision, "Applying IFB policy is not a user decision");
            Check((bool)Get(local, "ShowMainRibbonTab") && !(bool)Get(local, "HasManagedLogging")
                && !(bool)Get(local, "HasManagedTransportTls") && !(bool)Get(local, "HasManagedUpdateNotify"),
                "IFB-only policy neither hides Settings nor manages other groups");
            object clone = Call(local, "Clone");
            Check((bool)Get(clone, "HasManagedIfb") == present && (bool)Get(clone, "IsManagedIfbValid") == valid,
                "Clone retains IFB presence and validity");
            CheckIfb(clone, effective, "Clone retains effective IFB");
            CheckIfb(clone, raw, "Clone retains local IFB", true);
            string serialized = Serialize(local);
            var xml = new XmlDocument(); xml.LoadXml(serialized);
            for (int i = 0; i < IfbProperties.Length; i++)
                Check(xml.DocumentElement[IfbProperties[i]].InnerText == Convert.ToString(raw[i], System.Globalization.CultureInfo.InvariantCulture),
                    "Existing IFB XML elements write local choices only: " + IfbProperties[i]);
            Check(!serialized.Contains("HasManagedIfb") && !serialized.Contains("IsManagedIfbValid") && !serialized.Contains("LocalIfb"),
                "IFB policy metadata never enters profile XML");
            object restored = RoundTrip(local, root);
            CheckIfb(restored, raw, "XML alone restores saved IFB choices");
            Check(!(bool)Get(restored, "HasManagedIfb") && (bool)Get(restored, "IfbUserDecisionRecorded") == decision,
                "XML preserves the user decision without manufacturing a policy");
            Call(clone, "ApplyManagedSetupPolicy", (object)null);
            CheckIfb(clone, raw, "Removing IFB policy restores raw choices");
            Check(!(bool)Get(clone, "HasManagedIfb") && (bool)Get(clone, "IsManagedIfbValid")
                && (bool)Get(clone, "IfbUserDecisionRecorded") == decision, "Removal clears invalid IFB state without recording a decision");
            Check((bool)Get(local, "HasManagedIfb") == present, "Removing a clone overlay does not mutate its source");
        }
        for (int field = 0; field < IfbProperties.Length; field++)
        for (int location = 0; location < 4; location++)
        foreach (bool valid in new[] { true, false })
        {
            object[] values = { null, null, null, null };
            values[field] = valid ? new object[] { 0, 10, 1, 1024 }[field] : invalidInputs[field][0];
            Array policies = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 4);
            policies.SetValue(ManagedIfb(values[0], values[1], values[2], values[3]), location);
            for (int lower = location + 1; lower < 4; lower++) policies.SetValue(ManagedIfb(1, 90, 24, 49151), lower);
            object resolved = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
            object expected = valid ? (field == 0 ? (object)false : values[field]) : defaults[field];
            Check(object.Equals(Get(resolved, IfbProperties[field]), expected) && (bool)Get(resolved, "IsIfbPolicyValid") == valid,
                "First present IFB field wins in every hive/view, even when false or invalid");
        }
        Array mixed = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 4);
        mixed.SetValue(ManagedIfb(1, null, null, null, "HKLM64"), 0);
        mixed.SetValue(ManagedIfb("bad", 60, null, null, "HKLM32"), 1);
        mixed.SetValue(ManagedIfb(null, "bad", 6, null, "HKCU64"), 2);
        mixed.SetValue(ManagedIfb(null, null, "bad", 49151, "HKCU32"), 3);
        object merged = Call(T("Settings.ManagedSetupPolicy"), "Resolve", mixed);
        CheckIfb(merged, new object[] { true, 60, 6, 49151 }, "IFB fields independently merge across policy locations");
        Check((bool)Get(merged, "IsIfbPolicyValid"), "Ignored malformed lower-priority IFB values cannot invalidate the selected group");
        mixed.SetValue(ManagedTls(1, null, null), 0);
        mixed.SetValue(ManagedLogging(1, null), 1);
        mixed.SetValue(ManagedUpdateNotify(1), 2);
        mixed.SetValue(ManagedIfb(0, null, null, null), 3);
        merged = Call(T("Settings.ManagedSetupPolicy"), "Resolve", mixed);
        Check((bool)Get(merged, "HasIfbPolicy") && (bool)Get(merged, "HasTransportTlsPolicy")
            && (bool)Get(merged, "HasLoggingPolicy") && (bool)Get(merged, "HasUpdateNotifyPolicy"),
            "IFB, TLS, logging and update policy groups resolve independently");
        object invalidCache = New("Settings.AddinSettings");
        Call(invalidCache, "ApplyManagedSetupPolicy", ManagedIfb(1, 60, 0, 8888));
        Check(!(bool)Get(invalidCache, "IsManagedIfbValid") && (int)Get(invalidCache, "IfbCacheHours") == 24
            && (bool)Get(invalidCache, "IsManagedTransportTlsValid"), "Malformed IFB cache uses 24 hours without blocking unrelated transport");
        TestManagedIfbUi(root);
        Console.WriteLine("[OK] Managed IFB presence, DWORD validation, precedence, defaults, raw XML, clone/removal, UI and seat parity");
    }
    private static void TestManagedIfbUi(string root)
    {
        string[] fields = { "_ifbEnabledCheckBox", "_ifbDaysCombo", "_ifbCacheHoursCombo", "_ifbPortUpDown" };
        object[][] policies = {
            new object[] { 0, null, null, null }, new object[] { 1, 60, 6, 8888 },
            new object[] { null, 90, null, null }, new object[] { null, null, 1, null },
            new object[] { null, null, null, 1024 }, new object[] { 1, "bad", 25, 49152 }
        };
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (object[] values in policies)
        foreach (bool decision in new[] { false, true })
        {
            object local = New("Settings.AddinSettings");
            object[] raw = { false, 10, 2, 9999 };
            for (int i = 0; i < IfbProperties.Length; i++) Set(local, IfbProperties[i], raw[i]);
            Set(local, "IfbUserDecisionRecorded", decision);
            Call(local, "ApplyManagedSetupPolicy", ManagedIfb(values[0], values[1], values[2], values[3]));
            object status = Status(D(), D(), true, mode, seat);
            using (Form form = Settings(local, status))
            {
                var enabled = (CheckBox)Field(form, fields[0]);
                Check(((TabControl)Field(form, "_tabControl")).TabPages.Count == 8, "IFB-only rollout retains every settings tab");
                foreach (bool busy in new[] { false, true, false })
                {
                    Call(form, "SetBusy", busy); Call(form, "UpdateControlState");
                    foreach (string field in fields) Check(!((Control)Field(form, field)).Enabled,
                        "IFB group remains locked through busy and state refresh: " + field);
                    Check(enabled.Checked == (bool)Get(local, "IfbEnabled"), "Managed enabled value remains visible without credentials");
                    Check((int)Call(T("UI.SettingsForm"), "ParseComboValue", Field(form, fields[1]), 30) == (int)Get(local, "IfbDays")
                        && (int)Call(T("UI.SettingsForm"), "ParseComboValue", Field(form, fields[2]), 24) == (int)Get(local, "IfbCacheHours")
                        && ((NumericUpDown)Field(form, fields[3])).Value == (int)Get(local, "IfbPort"), "IFB controls display effective values");
                    Check(((CheckBox)Field(form, "_debugLogCheckBox")).Enabled == !busy
                        && ((Button)Field(form, "_updateCheckButton")).Enabled == !busy, "IFB policy leaves unrelated controls available");
                }
                string expectedHint = (string)T("Utilities.Strings").GetProperty(
                    (bool)Get(local, "IsManagedIfbValid") ? "PolicyAdminControlledTooltip" : "ManagedIfbPolicyInvalid", Flags).GetValue(null, null);
                foreach (string field in fields)
                    Check(((ToolTip)Field(form, "_toolTip")).GetToolTip((Control)Field(form, field)) == expectedHint,
                        "Every managed IFB control explains its lock or invalid policy: " + field);
                Call(form, "ApplyBackendPolicyStatus", "managed_ifb_test");
                foreach (string field in fields) Check(!((Control)Field(form, field)).Enabled, "Backend refresh cannot unlock IFB policy");
                Task save = (Task)Call(form, "SaveSettingsAsync");
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);
                while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Check(save.IsCompleted, "Managed IFB settings save completes without network credentials"); save.GetAwaiter().GetResult();
                Check(form.DialogResult == DialogResult.OK, "IFB policy errors do not block unrelated settings saves");
                object result = Get(form, "Result");
                CheckIfb(result, raw, "Saving managed controls never bakes IFB into local preferences", true);
                Check((bool)Get(result, "IfbUserDecisionRecorded") == decision, "Managed settings save does not record an IFB user decision");
                CheckIfb(RoundTrip(result, root), raw, "Actual managed settings save retains raw IFB XML");
            }
            Check(RolloutNotice(local, status) == RolloutNotice(ManagedSettings(null, true), status),
                "IFB-only rollout retains shared Community/Pro access and personal-seat rules");
            using (Form form = Settings(local, status))
            {
                ((TextBox)Field(form, "_serverUrlTextBox")).Text = "https://cloud.example.test";
                ((TextBox)Field(form, "_usernameTextBox")).Text = "ifb-test";
                ((TextBox)Field(form, "_appPasswordTextBox")).Text = "test-only";
                Call(form, "UpdateControlState");
                Check(((CheckBox)Field(form, fields[0])).Checked == (bool)Get(local, "IfbEnabled"),
                    "Completing credentials cannot auto-enable a managed false IFB value");
                foreach (string field in fields) Check(!((Control)Field(form, field)).Enabled, "Completing credentials cannot unlock managed IFB");
            }
        }
    }
    private static void TestManagedUpdateNotify(string root)
    {
        object[] inputs = { null, 0, 1, false, true, "false", "true", "", "bad", new byte[] { 1 } };
        foreach (bool raw in new[] { false, true })
        for (int index = 0; index < inputs.Length; index++)
        {
            object input = inputs[index];
            bool present = input != null;
            bool valid = index < 7;
            bool effective = present ? (index == 2 || index == 4 || index == 6) : raw;
            object policy = ManagedUpdateNotify(input);
            Check((bool)Get(policy, "HasUpdateNotifyPolicy") == present && (bool)Get(policy, "IsEnterpriseRollout") == present, "Update notification presence including false or malformed activates rollout");
            Check((bool)Get(policy, "IsUpdateNotifyPolicyValid") == valid, "Malformed update notification values retain a validity flag");
            object local = New("Settings.AddinSettings");
            Set(local, "UpdateNotifyEnabled", raw);
            Call(local, "ApplyManagedSetupPolicy", policy);
            Check((bool)Get(local, "UpdateNotifyEnabled") == effective && (bool)Get(local, "LocalUpdateNotifyEnabled") == raw, "Effective update notification overlay preserves local choice");
            Check((bool)Get(local, "ShowMainRibbonTab") && !(bool)Get(local, "HasManagedLogging") && !(bool)Get(local, "HasManagedTransportTls"), "Update policy alone does not hide the ribbon or lock TLS/logging");
            object clone = Call(local, "Clone");
            Check((bool)Get(clone, "HasManagedUpdateNotify") == present && (bool)Get(clone, "IsManagedUpdateNotifyValid") == valid, "Clone retains update policy presence and validity");
            Check((bool)Get(clone, "UpdateNotifyEnabled") == effective, "Clone uses effective update policy");
            object restored = RoundTrip(local, root);
            Check((bool)Get(restored, "UpdateNotifyEnabled") == raw && !(bool)Get(restored, "HasManagedUpdateNotify"), "XML roundtrip restores only the saved local preference");
            Check(!Serialize(local).Contains("HasManagedUpdateNotify") && !Serialize(local).Contains("LocalUpdateNotify"), "Registry update metadata is not persisted");
            Call(clone, "ApplyManagedSetupPolicy", (object)null);
            Check((bool)Get(clone, "UpdateNotifyEnabled") == raw && (bool)Get(local, "HasManagedUpdateNotify") == present, "Policy removal restores local update choice without mutating the original");
            object result = New("Models.UpdateCheckResult");
            Set(result, "LatestVersion", "9.9.9"); Set(result, "UpdateAvailable", true);
            Check((bool)Call(T("Services.UpdateCheckService"), "ShouldNotify", local, result) == effective, "Operative update notifications follow the effective policy");
            Call(T("Services.UpdateCheckService"), "MarkNotified", local, result);
            Check(!(bool)Call(T("Services.UpdateCheckService"), "ShouldNotify", local, result), "Policy does not bypass daily notification deduplication");
            Check((bool)Get(local, "LocalUpdateNotifyEnabled") == raw, "Marking notifications does not persist the policy as a local choice");
        }
        for (int location = 0; location < 4; location++)
        foreach (object value in new object[] { 0, 1, "bad" })
        {
            Array policies = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 4);
            policies.SetValue(ManagedUpdateNotify(value), location);
            for (int lower = location + 1; lower < 4; lower++) policies.SetValue(ManagedUpdateNotify(1), lower);
            object selected = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
            Check((bool)Get(selected, "UpdateNotifyEnabled") == object.Equals(value, 1) && (bool)Get(selected, "IsUpdateNotifyPolicyValid") == (value is int), "First-present update policy wins even when false or invalid");
        }
        Array mixed = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 3);
        mixed.SetValue(ManagedTls(1, null, null), 0);
        mixed.SetValue(ManagedLogging(1, null), 1);
        mixed.SetValue(ManagedUpdateNotify(1), 2);
        object combined = Call(T("Settings.ManagedSetupPolicy"), "Resolve", mixed);
        Check((bool)Get(combined, "HasTransportTlsPolicy") && (bool)Get(combined, "HasLoggingPolicy") && (bool)Get(combined, "UpdateNotifyEnabled"), "TLS, logging and update policies resolve independently across hives");
        mixed.SetValue(ManagedUpdateNotify(1), 0); mixed.SetValue(ManagedUpdateNotify("bad"), 1);
        combined = Call(T("Settings.ManagedSetupPolicy"), "Resolve", mixed);
        Check((bool)Get(combined, "UpdateNotifyEnabled") && (bool)Get(combined, "IsUpdateNotifyPolicyValid"), "Ignored malformed lower update value cannot invalidate selected policy");
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (object value in new object[] { null, 0, 1, "bad" })
        {
            object local = New("Settings.AddinSettings");
            Set(local, "UpdateNotifyEnabled", !object.Equals(value, 1));
            Call(local, "ApplyManagedSetupPolicy", ManagedUpdateNotify(value));
            object status = Status(D(), D(), true, mode, seat);
            using (Form form = Settings(local, status))
            {
                var box = (CheckBox)Field(form, "_updateNotifyCheckBox");
                foreach (bool busy in new[] { false, true, false })
                {
                    Call(form, "SetBusy", busy); Call(form, "UpdateControlState");
                    Check(box.Enabled == (value == null && !busy) && box.Checked == (bool)Get(local, "UpdateNotifyEnabled"), "Update notification checkbox displays and locks effective policy across state refreshes");
                    Check(((Button)Field(form, "_updateCheckButton")).Enabled == !busy, "Notification policy leaves manual update check available");
                    Check(((CheckBox)Field(form, "_debugLogCheckBox")).Enabled == !busy, "Update-only policy does not lock unrelated logging controls");
                }
                string hint = ((ToolTip)Field(form, "_toolTip")).GetToolTip(box);
                Check(string.IsNullOrEmpty(hint) == (value == null), "Update checkbox has a reachable managed tooltip");
                if (value is string) Check(hint == (string)T("Utilities.Strings").GetProperty("ManagedUpdateNotifyPolicyInvalid", Flags).GetValue(null, null), "Malformed update policy has configuration guidance");
                bool original = (bool)Get(local, "LocalUpdateNotifyEnabled");
                if (value == null) box.Checked = !box.Checked;
                Task save = (Task)Call(form, "SaveSettingsAsync");
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);
                while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Check(save.IsCompleted, "Update policy settings save completes without network"); save.GetAwaiter().GetResult();
                Check((bool)Get(Get(form, "Result"), "LocalUpdateNotifyEnabled") == (value == null ? !original : original), "Actual save preserves managed raw choice and accepts an unmanaged user edit");
            }
            if (value != null) Check(RolloutNotice(local, status) == RolloutNotice(ManagedSettings(null, true), status), "Update-only rollout retains Community/Pro parity and personal access checks");
        }
        Console.WriteLine("[OK] Managed update notification presence, precedence, XML, UI, actual notification decisions and seat parity");
    }
    private static void TestManagedLogging(string root)
    {
        string[] properties = { "DebugLoggingEnabled", "LogAnonymizationEnabled" };
        string[] controls = { "_debugLogCheckBox", "_debugAnonymizeCheckBox" };
        foreach (object debug in new object[] { null, 0, 1, "invalid", new byte[] { 1 } })
        foreach (object anonymize in new object[] { null, 0, 1, "invalid", new byte[] { 1 } })
        foreach (bool rawDebug in new[] { false, true })
        foreach (bool rawAnonymize in new[] { false, true })
        {
            bool present = debug != null || anonymize != null;
            bool valid = (debug == null || debug is int) && (anonymize == null || anonymize is int);
            bool[] raw = { rawDebug, rawAnonymize };
            bool[] effective = { present ? object.Equals(debug, 1) : rawDebug, present ? !object.Equals(anonymize, 0) : rawAnonymize };
            object policy = ManagedLogging(debug, anonymize);
            Check((bool)Get(policy, "HasLoggingPolicy") == present && (bool)Get(policy, "IsEnterpriseRollout") == present, "Logging presence controls managed mode, including false and malformed values");
            Check((bool)Get(policy, "IsLoggingPolicyValid") == valid, "Only malformed values invalidate logging; both false is valid");
            object local = New("Settings.AddinSettings");
            Set(local, properties[0], rawDebug); Set(local, properties[1], rawAnonymize);
            Call(local, "ApplyManagedSetupPolicy", policy);
            object clone = Call(local, "Clone");
            object restored = RoundTrip(local, root);
            var xml = new XmlDocument(); xml.LoadXml(Serialize(local));
            Check((bool)Get(clone, "HasManagedLogging") == present && (bool)Get(clone, "IsManagedLoggingValid") == valid, "Clone preserves logging management and validity");
            Check((bool)Get(local, "ShowMainRibbonTab") && !(bool)Get(local, "HasManagedTransportTls"), "Logging-only policy does not hide the ribbon or manage TLS");
            for (int i = 0; i < 2; i++)
            {
                Check((bool)Get(local, properties[i]) == effective[i] && (bool)Get(clone, properties[i]) == effective[i], "Logging getters expose managed values and missing sibling defaults");
                Check((bool)Get(local, "Local" + properties[i]) == raw[i] && (bool)Get(restored, properties[i]) == raw[i], "Logging raw values survive application and XML roundtrip");
                Check(bool.Parse(xml.DocumentElement[properties[i]].InnerText) == raw[i], "Existing XML logging elements persist local values, not registry overlay");
            }
            Check(!Serialize(local).Contains("HasManagedLogging") && !Serialize(local).Contains("LocalDebugLogging"), "Runtime logging policy is not serialized");
            Call(clone, "ApplyManagedSetupPolicy", (object)null);
            for (int i = 0; i < 2; i++) Check((bool)Get(clone, properties[i]) == raw[i], "Policy removal restores local logging choices");
            Check((bool)Get(local, "HasManagedLogging") == present, "Removing a clone policy does not alter the original");
        }
        for (int field = 0; field < 2; field++)
        for (int location = 0; location < 4; location++)
        foreach (object selected in new object[] { 0, 1, "bad", "", new byte[] { 1 } })
        {
            Array policies = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 4);
            policies.SetValue(ManagedLogging(field == 0 ? selected : null, field == 1 ? selected : null), location);
            for (int lower = location + 1; lower < 4; lower++) policies.SetValue(ManagedLogging(1, 0), lower);
            object resolved = Call(T("Settings.ManagedSetupPolicy"), "Resolve", policies);
            bool expected = selected is int ? object.Equals(selected, 1) : field == 1;
            Check((bool)Get(resolved, properties[field]) == expected, "First present logging field wins across all registry views, even false or malformed");
            Check((bool)Get(resolved, "IsLoggingPolicyValid") == (selected is int), "Lower registry value cannot hide a selected malformed logging policy");
        }
        Array mixed = Array.CreateInstance(T("Settings.ManagedSetupPolicy"), 3);
        mixed.SetValue(ManagedLogging(0, null), 0);
        mixed.SetValue(ManagedLogging(null, 0), 1);
        mixed.SetValue(ManagedLogging("bad", "bad"), 2);
        object merged = Call(T("Settings.ManagedSetupPolicy"), "Resolve", mixed);
        Check(!(bool)Get(merged, properties[0]) && !(bool)Get(merged, properties[1]) && (bool)Get(merged, "IsLoggingPolicyValid"), "Fields merge independently; ignored malformed lower values do not invalidate logging");
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "active", "none", "paused", "invalid" })
        foreach (object[] values in new[] { new object[] { null, null }, new object[] { 0, null }, new object[] { 1, 0 }, new object[] { "bad", "bad" } })
        {
            object local = New("Settings.AddinSettings");
            Set(local, properties[0], true); Set(local, properties[1], false);
            Call(local, "ApplyManagedSetupPolicy", ManagedLogging(values[0], values[1]));
            bool managed = (bool)Get(local, "HasManagedLogging");
            object status = Status(D(), D(), true, mode, seat);
            using (Form form = Settings(local, status))
            {
                foreach (bool busy in new[] { false, true, false })
                {
                    Call(form, "SetBusy", busy);
                    Call(form, "UpdateControlState");
                    for (int i = 0; i < 2; i++)
                    {
                        var box = (CheckBox)Field(form, controls[i]);
                        Check(box.Enabled == (!managed && !busy) && box.Checked == (bool)Get(local, properties[i]), "Logging UI shows effective values and stays locked across state refreshes for both seat types");
                    }
                }
                string hint = ((Label)Field(form, "_debugPolicyHintLabel")).Text;
                Check(string.IsNullOrEmpty(hint) == !managed, "Logging management is explained in the Debug tab");
                if (!(bool)Get(local, "IsManagedLoggingValid"))
                    Check(hint == (string)T("Utilities.Strings").GetProperty("ManagedLoggingPolicyInvalid", Flags).GetValue(null, null), "Malformed logging uses the configuration hint");
                Task save = (Task)Call(form, "SaveSettingsAsync");
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);
                while (!save.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Check(save.IsCompleted, "Managed logging settings save completes without network access");
                save.GetAwaiter().GetResult();
                object result = Get(form, "Result");
                Check((bool)Get(result, "LocalDebugLoggingEnabled") && !(bool)Get(result, "LocalLogAnonymizationEnabled"), "Actual settings save does not bake managed logging into local preferences");
            }
            if (managed) Check(RolloutNotice(local, status) == RolloutNotice(ManagedSettings(null, true), status), "Logging-only rollout retains existing Community/Pro access rules without a new logging network gate");
        }
        object managedSettings = New("Settings.AddinSettings");
        foreach (object[] values in new[] { new object[] { 0, 1 }, new object[] { 1, 0 } })
        {
            Call(managedSettings, "ApplyManagedSetupPolicy", ManagedLogging(values[0], values[1]));
            Call(T("NextcloudTalkAddIn"), "ConfigureDiagnosticsLogger", managedSettings);
            Check((bool)T("Utilities.DiagnosticsLogger").GetProperty("IsEnabled", Flags).GetValue(null, null) == (bool)Get(managedSettings, properties[0]), "Runtime logger uses effective debug policy");
            Check((bool)T("Utilities.DiagnosticsLogger").GetField("_anonymizationEnabled", Flags).GetValue(null) == (bool)Get(managedSettings, properties[1]), "Runtime logger uses effective anonymization policy");
        }
        Call(T("NextcloudTalkAddIn"), "ConfigureDiagnosticsLogger", New("Settings.AddinSettings"));
        Console.WriteLine("[OK] Managed logging priority, defaults, privacy, persistence, runtime logger, UI and seat parity");
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
        Call(owner, "StoreBackendPolicySnapshotIfCurrent", config, good, "test", 1L);
        DateTime previousSuccess = DateTime.UtcNow.AddMinutes(-30);
        owner.GetType().GetField("_emailSignaturePolicyCacheFetchedAtUtc", Flags).SetValue(owner, previousSuccess);
        Check(object.ReferenceEquals(Call(owner, "StoreBackendPolicySnapshotIfCurrent", config, unavailable, "test", 2L), good), "Failed refresh retains confirmed snapshot");
        Check((DateTime)Field(owner, "_emailSignaturePolicyCacheFetchedAtUtc") == previousSuccess, "Failed refresh does not mark the retained snapshot fresh");
        Call(owner, "StoreBackendPolicySnapshotIfCurrent", config, missing, "test", 3L);
        Check(object.ReferenceEquals(Call(owner, "StoreBackendPolicySnapshotIfCurrent", config, unavailable, "test", 4L), missing), "Fresh refusal replaces previous success");
        object other = New("Services.TalkServiceConfiguration", "https://cloud.example.test", "bob", "test-only");
        Check(object.ReferenceEquals(Call(owner, "StoreBackendPolicySnapshotIfCurrent", other, unavailable, "test", 5L), unavailable), "Rollout cache cannot cross account identity");
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
    private static void TestBackendPolicyAvailabilitySnapshot()
    {
        foreach (string mode in new[] { "community", "pro" })
        foreach (string reason in new[] { "nextcloud_unavailable", "backend_unavailable", "rate_limited", "authentication_rejected", "invalid_payload", "check_failed" })
        {
            object owner = New("NextcloudTalkAddIn");
            object configuration = New("Services.TalkServiceConfiguration", "https://cloud.example.test/nextcloud", "alice", "test-only");
            object confirmed = Status(D(), D(), true, mode, "active");
            Call(owner, "StoreBackendPolicySnapshotIfCurrent", configuration, confirmed, "test_confirmed", 1L);
            DateTime successAt = DateTime.UtcNow.AddMinutes(-30);
            owner.GetType().GetField("_emailSignaturePolicyCacheFetchedAtUtc", Flags).SetValue(owner, successAt);
            object failure = Status(D(), D(), true, mode, "active");
            Set(failure, "FetchSucceeded", false);
            Set(failure, "Reason", reason);
            if (reason == "rate_limited") Set(failure, "RetryAfterUtc", DateTime.UtcNow.AddMinutes(2));
            Check(object.ReferenceEquals(Call(owner, "StoreBackendPolicySnapshotIfCurrent", configuration, failure, "test_failed_refresh", 2L), confirmed), "Failed current check retains the account's last confirmed policy");
            Check((DateTime)Field(owner, "_emailSignaturePolicyCacheFetchedAtUtc") == successAt, "Failed current check does not renew the confirmed policy timestamp");
            object[] currentArgs = { configuration, null };
            Check((bool)Method(owner.GetType(), "TryGetCurrentBackendPolicyCheck", 2).Invoke(owner, currentArgs)
                && object.ReferenceEquals(currentArgs[1], failure), "Current failed check is observable separately from the last confirmed policy");
            bool outage = reason == "nextcloud_unavailable" || reason == "backend_unavailable" || reason == "rate_limited";
            Check((bool)Get(currentArgs[1], "IsServiceUnavailable") == outage, "Authentication, invalid payload and technical failure cannot become an offline exception");
            var retained = (Task)Call(owner, "GetEmailSignaturePolicyStatusAsync", configuration, "test_recent_failure");
            Check(retained.IsCompleted && object.ReferenceEquals(Get(retained, "Result"), confirmed), "Recent failure does not refetch or hide confirmed policy knowledge");

            owner.GetType().GetField("_backendPolicyLastCheckedAtUtc", Flags).SetValue(owner, DateTime.UtcNow.AddSeconds(-30));
            currentArgs = new object[] { configuration, null };
            Check((bool)Method(owner.GetType(), "TryGetCurrentBackendPolicyCheck", 2).Invoke(owner, currentArgs) == (reason == "rate_limited"), "Ordinary failed checks expire while server retry delays remain current");
            foreach (object changed in new[] {
                New("Services.TalkServiceConfiguration", "https://other.example.test/nextcloud", "alice", "test-only"),
                New("Services.TalkServiceConfiguration", "https://cloud.example.test/other", "alice", "test-only"),
                New("Services.TalkServiceConfiguration", "https://cloud.example.test/nextcloud", "bob", "test-only"),
                New("Services.TalkServiceConfiguration", "https://cloud.example.test/nextcloud", "alice", "different-test-only") })
            {
                currentArgs = new object[] { changed, null };
                Check(!(bool)Method(owner.GetType(), "TryGetCurrentBackendPolicyCheck", 2).Invoke(owner, currentArgs)
                    && currentArgs[1] == null, "URL, subpath, account and credential changes cannot reuse another current check");
                object[] cachedArgs = { changed, null };
                Check(!(bool)Method(owner.GetType(), "TryGetCachedEmailSignaturePolicyStatus", 2).Invoke(owner, cachedArgs)
                    && cachedArgs[1] == null, "URL, subpath, account and credential changes cannot reuse another policy snapshot");
            }
            object denied = Status(D(), D(), true, mode, "none");
            Call(owner, "StoreBackendPolicySnapshotIfCurrent", configuration, denied, "test_confirmed_refusal", 3L);
            currentArgs = new object[] { configuration, null };
            Check((bool)Method(owner.GetType(), "TryGetCurrentBackendPolicyCheck", 2).Invoke(owner, currentArgs)
                && object.ReferenceEquals(currentArgs[1], denied), "Fresh successful refusal replaces failed availability metadata");
            Check(object.ReferenceEquals(Field(owner, "_emailSignaturePolicyCache"), denied), "Fresh personal Seat refusal replaces an older permission");
        }
        object initialOwner = New("NextcloudTalkAddIn");
        object initialConfiguration = New("Services.TalkServiceConfiguration", "https://cloud.example.test", "alice", "test-only");
        object initialFailure = Status(D(), D(), true, "community", "active");
        Set(initialFailure, "FetchSucceeded", false);
        Set(initialFailure, "Reason", "backend_unavailable");
        Call(initialOwner, "StoreBackendPolicySnapshotIfCurrent", initialConfiguration, initialFailure, "test_first_start_backend_disabled", 1L);
        Check(Field(initialOwner, "_emailSignaturePolicyCache") == null, "First-start backend absence cannot invent a confirmed no-policy or no-Seat state");
        object[] firstCheckArgs = { initialConfiguration, null };
        Check((bool)Method(initialOwner.GetType(), "TryGetCurrentBackendPolicyCheck", 2).Invoke(initialOwner, firstCheckArgs)
            && object.ReferenceEquals(firstCheckArgs[1], initialFailure), "Unknown policy still retains truthful current backend failure metadata");
        object orderedOwner = New("NextcloudTalkAddIn");
        object olderAllowed = Status(D(), D(), true, "community", "active");
        object newerRefused = Status(D(), D(), true, "community", "none");
        Call(orderedOwner, "StoreBackendPolicySnapshotIfCurrent", initialConfiguration, newerRefused, "test_newer_refusal", 2L);
        DateTime refusalAt = (DateTime)Field(orderedOwner, "_emailSignaturePolicyCacheFetchedAtUtc");
        DateTime checkedAt = (DateTime)Field(orderedOwner, "_backendPolicyLastCheckedAtUtc");
        foreach (object olderResult in new[] { olderAllowed, initialFailure, null })
        {
            Check(object.ReferenceEquals(Call(orderedOwner, "StoreBackendPolicySnapshotIfCurrent", initialConfiguration, olderResult, "test_older_completion", 1L), newerRefused), "An older in-flight completion cannot undo a newly confirmed personal refusal");
            Check(object.ReferenceEquals(Field(orderedOwner, "_emailSignaturePolicyCache"), newerRefused)
                && object.ReferenceEquals(Field(orderedOwner, "_backendPolicyLastCheck"), newerRefused)
                && (DateTime)Field(orderedOwner, "_emailSignaturePolicyCacheFetchedAtUtc") == refusalAt
                && (DateTime)Field(orderedOwner, "_backendPolicyLastCheckedAtUtc") == checkedAt, "Ignored older completion cannot renew or replace current policy and availability timestamps");
        }
        Call(orderedOwner, "StoreBackendPolicySnapshotIfCurrent", initialConfiguration, initialFailure, "test_newer_outage", 3L);
        Check(object.ReferenceEquals(Field(orderedOwner, "_emailSignaturePolicyCache"), newerRefused)
            && object.ReferenceEquals(Field(orderedOwner, "_backendPolicyLastCheck"), initialFailure), "A newer failed check updates availability without erasing the confirmed refusal");
        Console.WriteLine("[OK] Account-scoped last-success policy and current availability/backoff transitions");
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
    private static void TestSettingsRefreshSnapshot()
    {
        foreach (string mode in new[] { "community", "pro" })
        foreach (string seat in new[] { "none", "paused", "invalid" })
        {
            object owner = New("NextcloudTalkAddIn");
            object local = New("Settings.AddinSettings");
            Set(local, "ServerUrl", "https://cloud.example.test/subpath");
            Set(local, "Username", "snapshot-test"); Set(local, "AppPassword", "test-only");
            object config = New("Services.TalkServiceConfiguration", Get(local, "ServerUrl"), Get(local, "Username"), Get(local, "AppPassword"));
            object good = Status(D(), D(), true, mode, "active");
            object denied = Status(D(), D(), true, mode, seat);
            Call(owner, "StoreBackendPolicySnapshotIfCurrent", config, good, "test_seed", 1L);
            using (Form form = (Form)New("UI.SettingsForm", local, null, good, null, Addressbook(), PolicyFetcher(denied, owner)))
            {
                var refresh = (System.Threading.Tasks.Task<bool>)Call(form, "RefreshSettingsServerStateAsync", config, false, "test_refresh");
                DateTime deadline = DateTime.UtcNow.AddSeconds(5);
                while (!refresh.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Check(refresh.IsCompleted && refresh.GetAwaiter().GetResult(), "Settings refresh completes using the supplied shared fetch path");
                Check(object.ReferenceEquals(Field(form, "_backendPolicyStatus"), denied), "Settings display the newly confirmed refusal");
                form.Close();
            }
            Check(object.ReferenceEquals(Field(owner, "_emailSignaturePolicyCache"), denied), "Cancelled settings still update the account snapshot after confirmed refusal");
            var next = (System.Threading.Tasks.Task)Call(owner, "GetEmailSignaturePolicyStatusAsync", config, "new_compose");
            Check(next.IsCompleted && object.ReferenceEquals(next.GetType().GetProperty("Result").GetValue(next, null), denied),
                "New compose uses refusal already received in Settings without another fetch");
        }
        Console.WriteLine("[OK] Settings refresh shares the account snapshot for subsequent compose actions after cancellation");
    }
    [STAThread]
    public static int Main(string[] args)
    {
        Product = Assembly.LoadFrom(args[0]);
        string root = args[1];
        try {
            TestDefaultsSourceMetadata();
            TestDefaultsSourcePrecedence(root);
            TestDefaultsSourceValues();
            TestManagedAuthMode(root);
            TestManagedSendPolicyFailureMode(root);
            TestManagedLoginEligibility();
            TestLocalChoices(root);
            TestWizards();
            TestSettingsEdits(root);
            TestSettingsLanguageControls(root);
            TestAttachmentAutomation();
            TestEnterpriseRollout(root);
            TestBackendPolicyAvailabilitySnapshot();
            TestConnectionOnboarding();
            TestSettingsRefreshSnapshot();
            TestManagedTls(root);
            TestManagedLogging(root);
            TestManagedUpdateNotify(root);
            TestManagedIfb(root);
            TestDefaultsSourceUi(root, args[2]);
            CaptureManagedAuthPreview(Path.Combine(Path.GetDirectoryName(args[2]), "auth-mode-settings.png"));
            Console.WriteLine("[OK] " + Checks + " production policy/persistence/UI assertions passed");
            return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex.ToString()); return 1; }
    }
}
'@ | Set-Content -LiteralPath $uiSource -Encoding UTF8
    $uiExe = Join-Path $TempRoot 'OutlookPolicyUiTests.exe'
    & $csc /nologo /target:exe "/out:$uiExe" /reference:System.dll /reference:System.Core.dll /reference:System.Xml.dll /reference:System.Windows.Forms.dll $uiSource
    if ($LASTEXITCODE -ne 0) { throw 'Policy UI test harness compilation failed.' }
    & $uiExe (Join-Path $uiOutput 'NcTalkOutlookAddIn.dll') $TempRoot (Join-Path $ProjectRoot '.tmp/defaults-source-settings.png')
    if ($LASTEXITCODE -ne 0) { throw 'Production policy/persistence/UI tests failed.' }

    # Compile the production registry reader and IFB manager against in-memory platform doubles.
    $ifbSource = Join-Path $TempRoot 'ManagedIfbRuntimeTests.cs'
    @'
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Services;
namespace Microsoft.Win32 {
    public sealed class RegistryKey : IDisposable {
        public static readonly Dictionary<string, Dictionary<string, object[]>> Fixtures = new Dictionary<string, Dictionary<string, object[]>>();
        private readonly string location;
        private RegistryKey(string value) { location = value; }
        public static RegistryKey OpenBaseKey(RegistryHive hive, RegistryView view) { return new RegistryKey(hive + "/" + view); }
        public RegistryKey OpenSubKey(string path, bool writable) {
            if (path != @"Software\Policies\NC Connector" || writable) throw new Exception("Unexpected registry access");
            return Fixtures.ContainsKey(location) ? this : null;
        }
        public string[] GetValueNames() { return Fixtures[location].Keys.ToArray(); }
        public RegistryValueKind GetValueKind(string name) { return (RegistryValueKind)Fixtures[location][name][1]; }
        public object GetValue(string name, object missing, RegistryValueOptions options) {
            if (options != RegistryValueOptions.DoNotExpandEnvironmentNames) throw new Exception("Unexpected registry expansion");
            object[] value; return Fixtures[location].TryGetValue(name, out value) ? value[0] : missing;
        }
        public void Dispose() {}
    }
}
namespace Microsoft.Office.Interop.Outlook { public sealed class Application { public string Version { get { return "16.0"; } } } }
namespace NcTalkOutlookAddIn.Settings {
    internal sealed class AddinSettings {
        internal const int DefaultIfbDays = 30, DefaultIfbCacheHours = 24, DefaultIfbPort = 7777, MinIfbPort = 1024, MaxIfbPort = 49151;
        internal string ServerUrl = "https://cloud.example.test", Username = "test", AppPassword = "test-only";
        internal bool IfbEnabled, IsEnterpriseRollout, IsManagedIfbValid = true;
        internal int IfbDays = DefaultIfbDays, IfbCacheHours = DefaultIfbCacheHours, IfbPort = DefaultIfbPort;
        internal static int NormalizeIfbPort(int port) { return port >= MinIfbPort && port <= MaxIfbPort ? port : DefaultIfbPort; }
    }
}
namespace NcTalkOutlookAddIn.Utilities {
    internal static class LogCategories { internal const string Core = "core", Ifb = "ifb"; }
    internal static class DiagnosticsLogger {
        internal static void Log(string category, string message) {}
        internal static void LogException(string category, string message, Exception ex) {}
    }
    internal static class Strings { internal const string ManagedIfbPolicyInvalid = "Invalid managed IFB policy"; }
}
namespace NcTalkOutlookAddIn.Services {
    internal sealed class TalkServiceConfiguration {
        private readonly bool complete;
        internal TalkServiceConfiguration(string url, string username, string password) { complete = url.StartsWith("https://") && username.Length > 0 && password.Length > 0; }
        internal bool IsComplete() { return complete; }
    }
    internal sealed class IfbAddressBookCache { internal IfbAddressBookCache(string root, string profile) {} }
    internal sealed class FreeBusyServer {
        internal static int Starts, Stops, Updates, Port, Days, CacheHours;
        internal static bool Rollout;
        internal static string Secret;
        internal FreeBusyServer(IfbAddressBookCache cache, Func<TalkServiceConfiguration, NcTalkOutlookAddIn.Models.BackendPolicyStatus> policy) {}
        internal void UpdateSettings(TalkServiceConfiguration config, int days, int cacheHours, string secret, bool rollout) {
            Updates++; Days = days; CacheHours = cacheHours; Secret = secret; Rollout = rollout;
        }
        internal void Start(int port) { Starts++; Port = port; }
        internal void Stop() { Stops++; }
    }
    internal sealed class IfbRegistryOwnershipManager {
        internal static int Applies, Restores;
        internal static string Url;
        internal IfbRegistryOwnershipManager(string root, string profile) {}
        internal void Apply(string version, string url, AddinSettings settings) { Applies++; Url = url; }
        internal void Restore() { Restores++; }
    }
}
internal static class ManagedIfbRuntimeTests {
    private static int checks;
    private static void Check(bool value, string message) { checks++; if (!value) throw new Exception(message); }
    private static string[] Locations() {
        return Environment.Is64BitOperatingSystem
            ? new[] { "LocalMachine/Registry64", "LocalMachine/Registry32", "CurrentUser/Registry64", "CurrentUser/Registry32" }
            : new[] { "LocalMachine/Registry32", "CurrentUser/Registry32" };
    }
    private static void Put(int location, string name, object value, RegistryValueKind kind) {
        string key = Locations()[location];
        Dictionary<string, object[]> values;
        if (!RegistryKey.Fixtures.TryGetValue(key, out values)) {
            values = new Dictionary<string, object[]>(StringComparer.OrdinalIgnoreCase); RegistryKey.Fixtures[key] = values;
        }
        values[name] = new object[] { value, kind };
    }
    private static AddinSettings Settings(ManagedSetupPolicy policy) {
        return new AddinSettings { IfbEnabled = policy.IfbEnabled, IfbDays = policy.IfbDays, IfbCacheHours = policy.IfbCacheHours,
            IfbPort = policy.IfbPort, IsManagedIfbValid = policy.IsIfbPolicyValid, IsEnterpriseRollout = policy.IsEnterpriseRollout };
    }
    private static void TestRegistryKindsAndPrecedence() {
        RegistryKey.Fixtures.Clear();
        Check(!ManagedSetupPolicy.Load().HasIfbPolicy, "Absent IFB registry group remains unmanaged");
        string[] fields = { "IfbEnabled", "IfbDays", "IfbCacheHours", "IfbPort" };
        object[] validValues = { 1, 60, 6, 8888 };
        foreach (string field in fields)
        foreach (RegistryValueKind kind in new[] { RegistryValueKind.DWord, RegistryValueKind.QWord, RegistryValueKind.String,
            RegistryValueKind.ExpandString, RegistryValueKind.MultiString, RegistryValueKind.Binary, RegistryValueKind.None, RegistryValueKind.Unknown })
        for (int location = 0; location < Locations().Length; location++) {
            RegistryKey.Fixtures.Clear();
            int index = Array.IndexOf(fields, field);
            object value = kind == RegistryValueKind.QWord ? (object)Convert.ToInt64(validValues[index])
                : kind == RegistryValueKind.String || kind == RegistryValueKind.ExpandString ? (object)validValues[index].ToString()
                : kind == RegistryValueKind.MultiString ? (object)new[] { validValues[index].ToString() }
                : kind == RegistryValueKind.Binary ? (object)new byte[] { 1 } : validValues[index];
            Put(location, field, value, kind);
            for (int lower = location + 1; lower < Locations().Length; lower++) Put(lower, field, validValues[index], RegistryValueKind.DWord);
            ManagedSetupPolicy policy = ManagedSetupPolicy.Load();
            bool valid = kind == RegistryValueKind.DWord || index == 0 && (kind == RegistryValueKind.QWord
                || kind == RegistryValueKind.String || kind == RegistryValueKind.ExpandString);
            Check(policy.HasIfbPolicy && policy.IsEnterpriseRollout && policy.IsIfbPolicyValid == valid,
                "Production reader honors IFB kind and first-present priority: " + field + "/" + kind + "/" + location);
            if (valid) Check(index == 0 ? policy.IfbEnabled : index == 1 ? policy.IfbDays == 60 : index == 2 ? policy.IfbCacheHours == 6 : policy.IfbPort == 8888,
                "Production registry reader retains selected valid value");
        }
        foreach (string field in fields) {
            RegistryKey.Fixtures.Clear(); Put(0, field.ToLowerInvariant(), null, RegistryValueKind.DWord);
            ManagedSetupPolicy policy = ManagedSetupPolicy.Load();
            Check(policy.HasIfbPolicy && !policy.IsIfbPolicyValid, "Present null IFB value cannot disappear or fall back");
        }
        RegistryKey.Fixtures.Clear();
        Put(0, "IfbEnabled", 0, RegistryValueKind.DWord);
        Put(Locations().Length - 1, "IfbEnabled", 1, RegistryValueKind.DWord);
        Put(Locations().Length - 1, "IfbCacheHours", 3, RegistryValueKind.DWord);
        ManagedSetupPolicy mixed = ManagedSetupPolicy.Load();
        Check(!mixed.IfbEnabled && mixed.IfbCacheHours == 3 && mixed.IfbDays == 30 && mixed.IfbPort == 7777 && mixed.IsIfbPolicyValid,
            "Managed false masks lower enabled while siblings resolve independently");
    }
    private static void TestDefaultsSourceRegistry() {
        RegistryKey.Fixtures.Clear();
        ManagedSetupPolicy absent = ManagedSetupPolicy.Load();
        Check(!absent.HasDefaultsSourcePolicy && !absent.IsEnterpriseRollout, "Absent source registry value leaves rollout inactive");
        foreach (RegistryValueKind kind in new[] { RegistryValueKind.String, RegistryValueKind.ExpandString, RegistryValueKind.MultiString,
            RegistryValueKind.DWord, RegistryValueKind.QWord, RegistryValueKind.Binary, RegistryValueKind.None, RegistryValueKind.Unknown })
        foreach (object source in new object[] { "local", "backend", "invalid", "inherit", "", null, 0, false })
        for (int location = 0; location < Locations().Length; location++) {
            RegistryKey.Fixtures.Clear();
            Put(location, "defaultssource", source, kind);
            for (int lower = location + 1; lower < Locations().Length; lower++) Put(lower, "DefaultsSource", "backend", RegistryValueKind.String);
            ManagedSetupPolicy policy = ManagedSetupPolicy.Load();
            bool valid = kind == RegistryValueKind.String && (object.Equals(source, "local") || object.Equals(source, "backend"));
            Check(policy.HasDefaultsSourcePolicy && policy.IsEnterpriseRollout && policy.IsDefaultsSourcePolicyValid == valid,
                "Only REG_SZ local/backend is valid; all present values activate rollout: " + kind + "/" + source + "/" + location);
            Check(policy.DefaultsSource == (valid ? (string)source : "local"), "Invalid first-present source masks lower-priority source and falls back locally");
        }
        RegistryKey.Fixtures.Clear();
        Put(0, "DefaultsSource", "local", RegistryValueKind.String);
        Put(Locations().Length - 1, "DefaultsSource", "backend", RegistryValueKind.String);
        Put(Locations().Length - 1, "IfbCacheHours", 3, RegistryValueKind.DWord);
        Put(Locations().Length - 1, "ShowMainRibbonTab", 0, RegistryValueKind.DWord);
        ManagedSetupPolicy mixed = ManagedSetupPolicy.Load();
        Check(mixed.DefaultsSource == "local" && mixed.IfbCacheHours == 3 && !mixed.ShowMainRibbonTab,
            "Defaults source, IFB and ribbon values retain independent per-value precedence");
        RegistryKey.Fixtures.Clear();
        Check(!ManagedSetupPolicy.Load().HasDefaultsSourcePolicy && !ManagedSetupPolicy.Load().IsEnterpriseRollout,
            "Removing the last source trigger removes its managed state and rollout");
    }
    private static void TestAuthModeRegistry() {
        RegistryKey.Fixtures.Clear();
        Check(!ManagedSetupPolicy.Load().HasAuthModePolicy && !ManagedSetupPolicy.Load().IsEnterpriseRollout,
            "Absent authentication mode leaves managed rollout inactive");
        foreach (RegistryValueKind kind in new[] { RegistryValueKind.String, RegistryValueKind.ExpandString, RegistryValueKind.MultiString,
            RegistryValueKind.DWord, RegistryValueKind.QWord, RegistryValueKind.Binary, RegistryValueKind.None, RegistryValueKind.Unknown })
        foreach (object value in new object[] { "LoginFlow", "Manual", " loginflow ", " MANUAL ", "invalid", "", "0", null, 0, false, new[] { "LoginFlow" } })
        for (int location = 0; location < Locations().Length; location++) {
            RegistryKey.Fixtures.Clear();
            Put(location, "authmode", value, kind);
            for (int lower = location + 1; lower < Locations().Length; lower++) Put(lower, "AuthMode", "Manual", RegistryValueKind.String);
            ManagedSetupPolicy policy = ManagedSetupPolicy.Load();
            string mode = value as string;
            bool valid = kind == RegistryValueKind.String && (string.Equals(mode == null ? null : mode.Trim(), "LoginFlow", StringComparison.OrdinalIgnoreCase)
                || string.Equals(mode == null ? null : mode.Trim(), "Manual", StringComparison.OrdinalIgnoreCase));
            Check(policy.HasAuthModePolicy && policy.IsEnterpriseRollout && policy.IsAuthModePolicyValid == valid,
                "AuthMode accepts only REG_SZ names and retains first-present priority: " + kind + "/" + location);
            Check(policy.AuthMode == (valid && string.Equals(mode.Trim(), "Manual", StringComparison.OrdinalIgnoreCase) ? AuthenticationMode.Manual : AuthenticationMode.LoginFlow),
                "Invalid selected auth mode masks lower entries and uses LoginFlow without claiming validity");
        }
        RegistryKey.Fixtures.Clear();
        Put(0, "AuthMode", "Manual", RegistryValueKind.String);
        Put(Locations().Length - 1, "AuthMode", "invalid", RegistryValueKind.String);
        Put(Locations().Length - 1, "NextcloudUrl", "https://cloud.example.test", RegistryValueKind.String);
        Put(Locations().Length - 1, "DefaultsSource", "backend", RegistryValueKind.String);
        ManagedSetupPolicy mixed = ManagedSetupPolicy.Load();
        Check(mixed.AuthMode == AuthenticationMode.Manual && mixed.IsAuthModePolicyValid && mixed.HasNextcloudUrl && mixed.DefaultsSource == "backend",
            "Auth mode priority is independent of URL and defaults source; unused invalid auth entries do not invalidate it");
        RegistryKey.Fixtures.Clear();
        Check(!ManagedSetupPolicy.Load().HasAuthModePolicy && !ManagedSetupPolicy.Load().IsEnterpriseRollout,
            "Removing the final auth policy trigger removes managed rollout");
    }
    private static void TestSendPolicyFailureModeRegistry() {
        RegistryKey.Fixtures.Clear();
        ManagedSetupPolicy absent = ManagedSetupPolicy.Load();
        Check(!absent.HasSendPolicyFailureModePolicy && absent.IsSendPolicyFailureModePolicyValid
            && absent.SendPolicyFailureMode == "failopen" && !absent.IsEnterpriseRollout, "Absent send policy defaults to failopen without managed rollout");
        foreach (RegistryValueKind kind in new[] { RegistryValueKind.String, RegistryValueKind.ExpandString, RegistryValueKind.MultiString,
            RegistryValueKind.DWord, RegistryValueKind.QWord, RegistryValueKind.Binary, RegistryValueKind.None, RegistryValueKind.Unknown })
        foreach (object value in new object[] { "failopen", "failclosed", " FAILOPEN ", " FAILCLOSED ", "invalid", "", "0", null, 0, false, new[] { "failclosed" } })
        for (int location = 0; location < Locations().Length; location++) {
            RegistryKey.Fixtures.Clear();
            Put(location, "sendpolicyfailuremode", value, kind);
            for (int lower = location + 1; lower < Locations().Length; lower++) Put(lower, "SendPolicyFailureMode", "failclosed", RegistryValueKind.String);
            ManagedSetupPolicy policy = ManagedSetupPolicy.Load();
            string mode = value as string;
            bool valid = kind == RegistryValueKind.String && (string.Equals(mode == null ? null : mode.Trim(), "failopen", StringComparison.OrdinalIgnoreCase)
                || string.Equals(mode == null ? null : mode.Trim(), "failclosed", StringComparison.OrdinalIgnoreCase));
            bool failClosed = valid && string.Equals(mode.Trim(), "failclosed", StringComparison.OrdinalIgnoreCase);
            Check(policy.HasSendPolicyFailureModePolicy && policy.IsEnterpriseRollout && policy.IsSendPolicyFailureModePolicyValid == valid,
                "Send mode accepts only REG_SZ values and retains first-present priority: " + kind + "/" + location);
            Check(policy.SendPolicyFailureMode == (failClosed ? "failclosed" : "failopen"), "Malformed selected send mode masks lower entries and falls back to failopen");
            Check(policy.SendPolicyFailureModeSource.StartsWith(location < (Environment.Is64BitOperatingSystem ? 2 : 1) ? "HKLM\\" : "HKCU\\"), "Send policy source belongs to its selected registry hive");
        }
        RegistryKey.Fixtures.Clear();
        Put(0, "SendPolicyFailureMode", "failopen", RegistryValueKind.String);
        Put(Locations().Length - 1, "SendPolicyFailureMode", "invalid", RegistryValueKind.String);
        Put(Locations().Length - 1, "AuthMode", "Manual", RegistryValueKind.String);
        Put(Locations().Length - 1, "NextcloudUrl", "https://cloud.example.test", RegistryValueKind.String);
        ManagedSetupPolicy mixed = ManagedSetupPolicy.Load();
        Check(mixed.SendPolicyFailureMode == "failopen" && mixed.IsSendPolicyFailureModePolicyValid && mixed.AuthMode == AuthenticationMode.Manual && mixed.HasNextcloudUrl,
            "Valid send mode ignores lower invalid entries while other fields retain independent precedence");
        RegistryKey.Fixtures.Clear();
        Check(!ManagedSetupPolicy.Load().HasSendPolicyFailureModePolicy && !ManagedSetupPolicy.Load().IsEnterpriseRollout,
            "Removing the last send-mode policy removes its managed state and rollout");
    }
    private static void TestManager() {
        RegistryKey.Fixtures.Clear(); Put(0, "IfbEnabled", 1, RegistryValueKind.DWord);
        Put(0, "IfbDays", 60, RegistryValueKind.DWord); Put(0, "IfbCacheHours", 6, RegistryValueKind.DWord); Put(0, "IfbPort", 8888, RegistryValueKind.DWord);
        AddinSettings settings = Settings(ManagedSetupPolicy.Load());
        using (var manager = new FreeBusyManager("in-memory", "managed-ifb-test")) {
            manager.Initialize(new Microsoft.Office.Interop.Outlook.Application()); manager.ApplySettings(settings);
            Check(FreeBusyServer.Starts == 1 && FreeBusyServer.Updates == 1 && IfbRegistryOwnershipManager.Applies == 1,
                "Valid managed IFB starts and owns the endpoint through the production manager");
            Check(FreeBusyServer.Port == 8888 && FreeBusyServer.Days == 60 && FreeBusyServer.CacheHours == 6 && FreeBusyServer.Rollout,
                "Production manager forwards managed port, days, shared cache hours and rollout gate");
            Check(FreeBusyServer.Secret.Length == 64 && IfbRegistryOwnershipManager.Url == "http://127.0.0.1:8888/nc-ifb/" + FreeBusyServer.Secret + "/freebusy/%NAME%@%SERVER%.vfb",
                "Managed IFB retains the process secret and Outlook attendee placeholders");
            foreach (bool enabled in new[] { false, true })
            foreach (bool credentials in new[] { false, true }) {
                settings.IfbEnabled = enabled; settings.Username = credentials ? "test" : ""; settings.IsManagedIfbValid = false;
                int restores = IfbRegistryOwnershipManager.Restores, stops = FreeBusyServer.Stops;
                bool rejected = false;
                try { manager.ApplySettings(settings); }
                catch (InvalidOperationException ex) { rejected = ex.Message == NcTalkOutlookAddIn.Utilities.Strings.ManagedIfbPolicyInvalid; }
                Check(rejected && FreeBusyServer.Starts == 1 && FreeBusyServer.Updates == 1 && IfbRegistryOwnershipManager.Applies == 1,
                    "Invalid managed IFB rejects before credentials, listener updates or registry writes");
                Check(IfbRegistryOwnershipManager.Restores == restores + 1 && FreeBusyServer.Stops == stops + 1,
                    "Invalid policy stops the running listener and restores owned state");
            }
            settings.IsManagedIfbValid = true; settings.IfbEnabled = true; settings.Username = ""; manager.ApplySettings(settings);
            Check(FreeBusyServer.Starts == 1, "Managed enabled without credentials cannot start a listener");
            settings.Username = "test"; settings.IfbEnabled = false; manager.ApplySettings(settings);
            Check(FreeBusyServer.Starts == 1, "Managed disabled with complete credentials cannot start a listener");
            settings.IfbEnabled = true; manager.ApplySettings(settings);
            Check(FreeBusyServer.Starts == 2, "Corrected managed settings can start again");
        }
        Check(FreeBusyServer.Stops > 0 && IfbRegistryOwnershipManager.Restores > 0, "Manager disposal stops and restores without rewriting managed enabled");
    }
    public static int Main() {
        try { TestRegistryKindsAndPrecedence(); TestDefaultsSourceRegistry(); TestAuthModeRegistry(); TestSendPolicyFailureModeRegistry(); TestManager(); Console.WriteLine("[OK] " + checks + " in-memory production IFB/defaults-source/auth-mode/send-mode registry/manager assertions passed"); return 0; }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
'@ | Set-Content -LiteralPath $ifbSource -Encoding UTF8
    $ifbExe = Join-Path $TempRoot 'ManagedIfbRuntimeTests.exe'
    & $csc /noconfig /nologo /nowarn:0436 /target:exe "/out:$ifbExe" /reference:System.dll /reference:System.Core.dll $ifbSource (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Settings/ManagedSetupPolicy.cs') (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Settings/AuthenticationMode.cs') (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Models/BackendPolicyStatus.cs') (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Services/FreeBusyManager.cs') (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Utilities/NextcloudUriValidator.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Managed IFB runtime test harness compilation failed.' }
    & $ifbExe
    if ($LASTEXITCODE -ne 0) { throw 'Production managed IFB registry/runtime tests failed.' }

    # Exercise the production workflow with dialog/network doubles; never touch a user's profile.
    $workflowSource = Join-Path $TempRoot 'SettingsWorkflowTests.cs'
    @'
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
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
        public SettingsForm(AddinSettings current, object app, object policy, object cache, object book,
            Func<TalkServiceConfiguration, string, BackendPolicyStatus> fetch) {
            if (fetch == null) throw new Exception("Settings refresh callback missing");
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

    # Run the production authentication entry and login flow against controlled services.
    $authFormSource = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/UI/SettingsForm.cs') -Raw
    $authGeneralSource = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/UI/SettingsForm.General.cs') -Raw
    $authMethods = foreach ($entry in @(
        @($authFormSource, 'internal void BeginAuthentication('),
        @($authFormSource, 'protected override async void OnShown('),
        @($authFormSource, 'private async Task SaveSettingsWithErrorHandlingAsync('),
        @($authFormSource, 'private async Task SaveSettingsAsync('),
        @($authGeneralSource, 'private bool ShouldStartManagedLoginFlow('),
        @($authGeneralSource, 'private async void OnLoginFlowButtonClick('),
        @($authGeneralSource, 'private async Task StartLoginFlowAsync(')
    )) {
        $methodText = $entry[0]
        $methodStart = $methodText.IndexOf($entry[1], [StringComparison]::Ordinal)
        if ($methodStart -lt 0) { throw "Authentication method not found: $($entry[1])" }
        $methodLine = $methodText.LastIndexOf([char]10, $methodStart) + 1
        $methodIndent = $methodText.Substring($methodLine, $methodStart - $methodLine)
        $methodClosing = [string][char]10 + $methodIndent + '}'
        $methodEnd = $methodText.IndexOf($methodClosing, $methodStart, [StringComparison]::Ordinal)
        if ($methodEnd -lt 0) { throw "Authentication method end not found: $($entry[1])" }
        $methodText.Substring($methodStart, $methodEnd + $methodClosing.Length - $methodStart)
    }
    $authSource = Join-Path $TempRoot 'ManagedAuthenticationTests.cs'
    $authHarness = @'
using System;
using System.Globalization;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

internal sealed class AuthSettings {
    internal bool HasManagedAuthMode = true, IsManagedAuthModeValid = true, HasManagedNextcloudUrl = true;
    internal AuthenticationMode AuthMode = AuthenticationMode.LoginFlow;
    internal string ManagedNextcloudUrl = "https://cloud.example.test/nextcloud";
    internal bool IsManagedTransportTlsValid = true, ShowMainRibbonTab = true;
    internal bool HasManagedIfb = true, HasManagedLogging = true, HasManagedTransportTls = true, HasManagedUpdateNotify = true;
    internal bool ManagedNextcloudUrlLocked, IfbEnabled, IfbUserDecisionRecorded, DebugLoggingEnabled, LogAnonymizationEnabled;
    internal bool TransportTlsUseSystemDefault, TransportTlsEnable12, TransportTlsEnable13, UpdateNotifyEnabled;
    internal int IfbDays, IfbPort, IfbCacheHours;
    internal string ServerUrl, Username, AppPassword;
}
internal static class AddinSettings { internal static int NormalizeIfbPort(int port) { return port; } }
internal sealed class TalkServiceConfiguration {
    private readonly string url, user, password;
    internal TalkServiceConfiguration(string url, string user, string password) { this.url = url; this.user = user; this.password = password; }
    internal bool IsComplete() { return GetNormalizedBaseUrl().Length > 0 && !string.IsNullOrWhiteSpace(user) && !string.IsNullOrEmpty(password); }
    internal string GetNormalizedBaseUrl() { string normalized; return NextcloudUriValidator.TryNormalizeBaseUrl(url, out normalized) ? normalized : ""; }
}
internal sealed class TalkServiceException : Exception { internal TalkServiceException(string message) : base(message) {} }
internal sealed class LoginFlowStart { internal string LoginUrl = "https://cloud.example.test/nextcloud/login/isolated-test"; }
internal sealed class LoginFlowCredentials { internal string LoginName = "test-login", AppPassword = "test-only-password"; }
internal sealed class TalkLoginFlowService {
    internal static readonly ManualResetEventSlim StartRelease = new ManualResetEventSlim(true);
    internal static readonly ManualResetEventSlim PollRelease = new ManualResetEventSlim(true);
    internal static int Starts, Polls, StartThread, PollThread;
    internal static bool FailStart, FailPoll;
    internal TalkLoginFlowService(string url) {}
    internal LoginFlowStart StartLoginFlow() {
        StartThread = Thread.CurrentThread.ManagedThreadId; Interlocked.Increment(ref Starts);
        if (!StartRelease.Wait(TimeSpan.FromSeconds(10))) throw new Exception("Test start wait expired");
        if (FailStart) throw new TalkServiceException("Test start failure");
        return new LoginFlowStart();
    }
    internal LoginFlowCredentials CompleteLoginFlow(LoginFlowStart start, TimeSpan timeout, TimeSpan interval) {
        PollThread = Thread.CurrentThread.ManagedThreadId; Interlocked.Increment(ref Polls);
        if (!PollRelease.Wait(TimeSpan.FromSeconds(10))) throw new Exception("Test poll wait expired");
        if (FailPoll) throw new TalkServiceException("Test poll failure");
        return new LoginFlowCredentials();
    }
}
internal sealed class TalkService {
    internal static int Verifications;
    internal static bool VerifyResult = true;
    internal TalkService(TalkServiceConfiguration configuration) {}
    internal bool VerifyConnection(out string response) { Interlocked.Increment(ref Verifications); response = ""; return VerifyResult; }
}
internal static class BrowserLauncher {
    internal static int Calls, ThreadId;
    internal static void OpenUrl(string url, string category, string error) { Calls++; ThreadId = Thread.CurrentThread.ManagedThreadId; }
}
internal static class LogCategories { internal const string Core = "core"; }
internal static class DiagnosticsLogger {
    internal static void Log(string category, string message) {}
    internal static void LogException(string category, string message, Exception ex) {}
}
internal static class Strings {
    internal const string StatusServerUrlRequired = "url required", StatusInvalidServerUrl = "invalid url", StatusLoginFlowStarting = "starting";
    internal const string StatusLoginFlowBrowser = "browser", ErrorCredentialsNotVerified = "not verified", StatusLoginFlowFailure = "failure: {0}", StatusLoginFlowSuccess = "success";
    internal const string SettingsSaveFailed = "save failed", ManagedTlsPolicyInvalid = "invalid TLS", TransportTlsSelectionRequired = "TLS required", DialogTitle = "Test";
}
internal sealed class AuthForm : Form {
    internal AuthSettings Result = new AuthSettings();
    internal readonly TextBox _serverUrlTextBox = new TextBox(), _usernameTextBox = new TextBox(), _appPasswordTextBox = new TextBox();
    internal readonly RadioButton _manualRadio = new RadioButton(), _loginFlowRadio = new RadioButton();
    internal readonly TabControl _tabControl = new TabControl();
    internal readonly TabPage _generalTab = new TabPage();
    internal readonly CheckBox _tlsUseSystemDefaultCheckBox = new CheckBox(), _tlsEnable12CheckBox = new CheckBox(), _tlsEnable13CheckBox = new CheckBox();
    internal readonly CheckBox _ifbEnabledCheckBox = new CheckBox(), _debugLogCheckBox = new CheckBox(), _debugAnonymizeCheckBox = new CheckBox(), _updateNotifyCheckBox = new CheckBox();
    internal readonly ComboBox _ifbDaysCombo = new ComboBox(), _ifbCacheHoursCombo = new ComboBox();
    internal readonly NumericUpDown _ifbPortUpDown = new NumericUpDown();
    internal bool _isBusy, _authenticationRequired, _authenticationRejected, _connectionSetupPending, _automaticLoginFlowPending;
    internal bool TlsValid = true, SaveRefreshResult = true, SaveRefreshThrows, _ifbDefaultApplied;
    internal int TlsApplies, TlsRestores, VerifiedTransitions, SaveRefreshes, CloseCalls, RepeatedVerification;
    internal string StatusText;
    internal AuthForm() {
        _serverUrlTextBox.Text = Result.ManagedNextcloudUrl;
        _loginFlowRadio.Checked = true;
        _tlsEnable12CheckBox.Checked = true;
        _tabControl.TabPages.Add(_generalTab);
    }
    internal void ShowEvent() { OnShown(EventArgs.Empty); }
    internal void LoginButton() { OnLoginFlowButtonClick(null, EventArgs.Empty); }
    internal Task LoginTask() { return StartLoginFlowAsync(); }
    internal new void Close() { CloseCalls++; base.Close(); }
    private Task<bool> TestConnectionAsync() { RepeatedVerification++; return Task.FromResult(true); }
    private Task<bool> RefreshSettingsServerStateAsync(TalkServiceConfiguration configuration, bool save, string source) {
        SaveRefreshes++;
        if (_isBusy || TlsApplies != TlsRestores) throw new Exception("Save started before login cleanup");
        if (SaveRefreshThrows) throw new InvalidOperationException("Controlled save failure");
        return Task.FromResult(SaveRefreshResult);
    }
    private int ParseComboValue(ComboBox combo, int fallback) { return fallback; }
    private void ApplyResponsiveLayout(bool width) {}
    private void ApplyBackendPolicyStatus(string source) { if (source == "login_verified") VerifiedTransitions++; }
    private void SetBusy(bool value) { _isBusy = value; }
    private void SetStatus(string message, bool error) { StatusText = message; }
    private SecurityProtocolType ApplySelectedTransportSecurity(string source) {
        TlsApplies++; if (!TlsValid) throw new InvalidOperationException("Invalid managed TLS"); return ServicePointManager.SecurityProtocol;
    }
    private void RestoreTemporaryTls(SecurityProtocolType previous, string source) { TlsRestores++; }
    private void HandleServiceFailure(string format, TalkServiceException ex) { SetStatus(string.Format(format, ex.Message), true); }
    __AUTH_METHODS__
}
internal static class ManagedAuthenticationTests {
    private static int checks;
    private static void Check(bool value, string message) { checks++; if (!value) throw new Exception(message); }
    private static void PumpUntil(Func<bool> condition, string message) {
        DateTime deadline = DateTime.UtcNow.AddSeconds(5);
        while (!condition() && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(1); }
        Check(condition(), message);
    }
    private static void Reset() {
        TalkLoginFlowService.StartRelease.Set(); TalkLoginFlowService.PollRelease.Set();
        TalkLoginFlowService.Starts = TalkLoginFlowService.Polls = BrowserLauncher.Calls = TalkService.Verifications = 0;
        TalkLoginFlowService.FailStart = TalkLoginFlowService.FailPoll = false; TalkService.VerifyResult = true;
    }
    private static void Complete(AuthForm form) { PumpUntil(() => !form._isBusy, "Login operation completes"); }
    [STAThread]
    public static int Main() {
        try {
            Reset();
            using (var form = new AuthForm()) {
                form.ShowEvent(); form.ShowEvent();
                Check(TalkLoginFlowService.Starts == 0, "Normal settings never start browser authentication automatically");
                form.LoginButton(); Complete(form);
                Check(TalkLoginFlowService.Starts == 1 && BrowserLauncher.Calls == 1 && form.VerifiedTransitions == 1, "Explicit login button retains the original verified flow");
                Check(form.CloseCalls == 0 && form.SaveRefreshes == 0, "Normal Settings never save or close automatically");
            }
            Reset();
            using (var form = new AuthForm()) {
                int uiThread = Thread.CurrentThread.ManagedThreadId;
                TalkLoginFlowService.StartRelease.Reset(); TalkLoginFlowService.PollRelease.Reset();
                form.BeginAuthentication(false); Check(form._automaticLoginFlowPending, "Eligible authentication entry schedules login");
                form.ShowEvent();
                PumpUntil(() => TalkLoginFlowService.Starts == 1, "Automatic entry reaches the controlled service");
                Check(!form._automaticLoginFlowPending && form._isBusy, "Pending automatic start is consumed before awaiting service work");
                form.ShowEvent(); form.LoginButton();
                Check(TalkLoginFlowService.Starts == 1, "Repeated show and login clicks do not duplicate an active operation");
                TalkLoginFlowService.StartRelease.Set();
                PumpUntil(() => TalkLoginFlowService.Polls == 1, "Login start hands off to browser and poll");
                Check(BrowserLauncher.Calls == 1 && BrowserLauncher.ThreadId == uiThread && TalkLoginFlowService.StartThread != uiThread
                    && TalkLoginFlowService.PollThread != uiThread, "Network work leaves the UI thread while browser launch returns to it");
                TalkLoginFlowService.PollRelease.Set(); Complete(form);
                Check(!form._connectionSetupPending && !form._authenticationRejected && form.VerifiedTransitions == 1
                    && form.TlsApplies == 1 && form.TlsRestores == 1, "Verified automatic login finishes setup and restores temporary TLS");
                Check(form.CloseCalls == 1 && form.DialogResult == DialogResult.OK && form.SaveRefreshes == 1
                    && form.Result.Username == "test-login" && form.Result.AppPassword == "test-only-password"
                    && form.RepeatedVerification == 0, "Managed login uses the real save path and closes once without duplicate verification");
                form.ShowEvent(); form.BeginAuthentication(true); form.ShowEvent();
                Check(TalkLoginFlowService.Starts == 1 && !form._automaticLoginFlowPending, "Complete existing credentials suppress automatic login even on rejected-credential entry");
            }
            foreach (string failure in new[] { "start", "poll", "verification" }) {
                Reset();
                using (var form = new AuthForm()) {
                    TalkLoginFlowService.FailStart = failure == "start"; TalkLoginFlowService.FailPoll = failure == "poll";
                    TalkService.VerifyResult = failure != "verification";
                    form.BeginAuthentication(false); form.ShowEvent(); Complete(form);
                    Check(TalkLoginFlowService.Starts == 1 && form._connectionSetupPending && form.VerifiedTransitions == 0, "Failed automatic login cannot complete setup: " + failure);
                    Check(form.CloseCalls == 0 && form.SaveRefreshes == 0, "Failed login never saves or closes: " + failure);
                    form.ShowEvent(); Check(TalkLoginFlowService.Starts == 1, "A failed attempt never starts an automatic retry: " + failure);
                    TalkLoginFlowService.FailStart = TalkLoginFlowService.FailPoll = false; TalkService.VerifyResult = true;
                    form.LoginButton(); Complete(form);
                    Check(TalkLoginFlowService.Starts == 2 && form.VerifiedTransitions == 1, "Explicit button retries failed automatic authentication: " + failure);
                    Check(form.CloseCalls == 1 && form.DialogResult == DialogResult.OK, "Successful managed retry saves and closes: " + failure);
                }
            }
            Reset();
            using (var form = new AuthForm()) {
                TalkLoginFlowService.FailStart = true;
                form.BeginAuthentication(false); form.ShowEvent(); Complete(form);
                form.BeginAuthentication(false); form.ShowEvent(); Complete(form);
                Check(TalkLoginFlowService.Starts == 2, "A new authentication invocation may make one new automatic attempt");
            }
            foreach (string cancellation in new[] { "before-show", "during-start", "during-poll" }) {
                Reset();
                using (var form = new AuthForm()) {
                    form.BeginAuthentication(false);
                    if (cancellation == "before-show") { form.Dispose(); form.ShowEvent(); }
                    else {
                        if (cancellation == "during-start") TalkLoginFlowService.StartRelease.Reset();
                        else TalkLoginFlowService.PollRelease.Reset();
                        form.ShowEvent();
                        PumpUntil(() => cancellation == "during-start" ? TalkLoginFlowService.Starts == 1 : TalkLoginFlowService.Polls == 1, "Cancellation fixture reaches pending stage");
                        form.Dispose(); TalkLoginFlowService.StartRelease.Set(); TalkLoginFlowService.PollRelease.Set(); Complete(form);
                    }
                    Check(!form._automaticLoginFlowPending && TalkService.Verifications == 0 && form.VerifiedTransitions == 0, "Cancellation cannot verify or finish setup: " + cancellation);
                    Check(BrowserLauncher.Calls == (cancellation == "during-poll" ? 1 : 0), "Closed forms cannot open a later login browser: " + cancellation);
                    Check(form._usernameTextBox.Text == "" && form._appPasswordTextBox.Text == "", "Cancelled login cannot overwrite credentials: " + cancellation);
                    Check(form.SaveRefreshes == 0 && form.CloseCalls == 0, "Cancelled login cannot save or accept the form: " + cancellation);
                }
            }
            foreach (string reason in new[] { "manual", "invalid-policy", "url-missing", "url-changed", "busy", "tls-invalid" }) {
                Reset();
                using (var form = new AuthForm()) {
                    if (reason == "manual") { form.Result.AuthMode = AuthenticationMode.Manual; form._manualRadio.Checked = true; }
                    if (reason == "invalid-policy") form.Result.IsManagedAuthModeValid = false;
                    if (reason == "url-missing") { form.Result.HasManagedNextcloudUrl = false; form._serverUrlTextBox.Text = ""; }
                    if (reason == "tls-invalid") form.TlsValid = false;
                    form.BeginAuthentication(false);
                    if (reason == "url-changed") form._serverUrlTextBox.Text = "https://other.example.test";
                    if (reason == "busy") form._isBusy = true;
                    form.ShowEvent();
                    if (reason != "busy") Complete(form);
                    Check(!form._automaticLoginFlowPending && TalkLoginFlowService.Starts == 0 && BrowserLauncher.Calls == 0, "Ineligible automatic entry cannot contact the service: " + reason);
                    form._isBusy = false; form.ShowEvent();
                    Check(TalkLoginFlowService.Starts == 0, "State changes after first show do not replay automatic entry: " + reason);
                    if (reason == "invalid-policy" || reason == "busy" || reason == "url-changed") {
                        form.LoginButton(); Complete(form);
                        Check(TalkLoginFlowService.Starts == 1, "Explicit login remains available without an automatic attempt: " + reason);
                        Check(form.CloseCalls == (reason == "invalid-policy" ? 0 : 1), "Only valid managed onboarding closes after explicit login: " + reason);
                    } else {
                        form.LoginButton(); Complete(form);
                        Check(TalkLoginFlowService.Starts == 0, "Manual/URL/TLS guard remains in the shared explicit login flow: " + reason);
                    }
                }
            }
            foreach (bool ribbon in new[] { false, true }) {
                Reset();
                using (var form = new AuthForm()) {
                    form.Result.ShowMainRibbonTab = ribbon;
                    form.Result.HasManagedNextcloudUrl = false;
                    form.BeginAuthentication(false); form.ShowEvent(); form.LoginButton(); Complete(form);
                    Check(form.CloseCalls == 1 && form.SaveRefreshes == 1 && form.DialogResult == DialogResult.OK,
                        "Managed action login saves in full and authentication-only dialogs even without URL autostart");
                }
            }
            foreach (string outcome in new[] { "refused", "exception", "unmanaged" }) {
                Reset();
                using (var form = new AuthForm()) {
                    form.SaveRefreshResult = outcome != "refused";
                    form.SaveRefreshThrows = outcome == "exception";
                    if (outcome == "unmanaged") form.Result.HasManagedAuthMode = false;
                    form.BeginAuthentication(false); form.LoginButton(); Complete(form);
                    Check(form.VerifiedTransitions == 1 && form.CloseCalls == 0 && form.DialogResult != DialogResult.OK,
                        "Unsuccessful save and unmanaged login never close automatically: " + outcome);
                    Check(form.SaveRefreshes == (outcome == "unmanaged" ? 0 : 1), "Save uses existing refresh only for managed action login: " + outcome);
                    if (outcome == "exception") Check(form.StatusText == Strings.SettingsSaveFailed, "Automatic save shares the save-button error handler");
                }
            }
            Console.WriteLine("[OK] " + checks + " production authentication entry/login-flow assertions passed with isolated services");
            return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { TalkLoginFlowService.StartRelease.Set(); TalkLoginFlowService.PollRelease.Set(); }
    }
}
'@
    $authHarness.Replace('__AUTH_METHODS__', ($authMethods -join "`r`n")) | Set-Content -LiteralPath $authSource -Encoding UTF8
    $authExe = Join-Path $TempRoot 'ManagedAuthenticationTests.exe'
    & $csc /noconfig /nologo /target:exe "/out:$authExe" /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll $authSource (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Settings/AuthenticationMode.cs') (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Utilities/NextcloudUriValidator.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Managed authentication test harness compilation failed.' }
    & $authExe
    if ($LASTEXITCODE -ne 0) { throw 'Production managed authentication tests failed.' }
}
finally {
    if (Test-Path $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}

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
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;

namespace NcTalkOutlookAddIn.Settings
{
    internal sealed class AddinSettings
    {
        internal bool? EmailSignatureOnCompose { get; set; }
        internal bool? EmailSignatureOnReply { get; set; }
        internal bool? EmailSignatureOnForward { get; set; }
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

    public static int Main()
    {
        TestPasswordDeliveryMode();
        TestSecretsExpireDays();
        TestLockedBackendValueWins();
        TestEditableBackendValueKeepsLocalChoice();
        TestAttachmentLinkTargetMapping();
        TestAttachmentLinkTargetPrecedence();
        TestEditableSignatureDefaultCanBeEnabledLocally();
        TestLockedSignatureDefaultCannotBeEnabledLocally();
        TestSignaturePolicyAvailability();

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
            false,
            string.Empty,
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
            false,
            string.Empty,
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
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\AttachmentLinkTargetPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\EmailSignaturePolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryMode.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\EmailSignaturePolicyService.cs")
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

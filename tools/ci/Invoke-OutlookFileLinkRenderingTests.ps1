Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-filelink-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "OutlookFileLinkRenderingTests.cs"
    @'
using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

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
        internal const string Core = "core";
        internal const string FileLink = "filelink";
    }
}

internal static class OutlookFileLinkRenderingTests
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
        Strings.SetPreferredUiLanguage("en");
        TestNormalModeUsesNextcloudLinkWording();
        TestAttachmentModeKeepsNextcloudSubpath();
        TestPlainTextKeepsNextcloudSubpath();
        TestCustomTemplateResolvesModeAwareLinkVariables();
        TestBackendEffectiveLanguageLocalizesCustomTemplateCopy();
        TestOlderBackendModeAwareTemplateStillRenders();
        TestLegacyCustomTemplateStillRenders();
        TestSecretLinkLabelHidesLongUrlInHtml();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " FileLink rendering test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook FileLink rendering tests passed.");
        return 0;
    }

    private static void TestNormalModeUsesNextcloudLinkWording()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = new FileLinkRequest
        {
            ShareName = "Folder",
            AttachmentMode = false,
            PasswordSeparateEnabled = false,
            NoteEnabled = false,
            Permissions = FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create
        };
        string html = FileLinkHtmlBuilder.Build(result, request, "en");
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "en");

        Check("Normal HTML labels the share page as a Nextcloud link", html.Contains(">Nextcloud link<"), html);
        Check("Normal plain text labels the share page as a Nextcloud link", plainText.Contains("Nextcloud link: https://cloud.example.test/nc/s/AbCd1234"), plainText);
        Check("Normal share URL does not gain a ZIP suffix", !html.Contains("/AbCd1234/download") && !plainText.Contains("/AbCd1234/download"));
    }

    private static void TestAttachmentModeKeepsNextcloudSubpath()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = BuildAttachmentRequest();
        string html = FileLinkHtmlBuilder.Build(result, request, "en");

        Check("Attachment ZIP URL keeps /nc subpath in HTML", html.Contains("https://cloud.example.test/nc/s/AbCd1234/download"), html);
        Check("Attachment ZIP URL does not drop /nc subpath in HTML", !html.Contains("https://cloud.example.test/s/AbCd1234/download"), html);
        Check("Attachment HTML labels the link as ZIP download", html.Contains(">ZIP download<"), html);
        Check("Attachment HTML explains ZIP download behavior", html.Contains("Download the shared files as a ZIP archive"), html);
    }

    private static void TestPlainTextKeepsNextcloudSubpath()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = BuildAttachmentRequest();
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "en");

        Check("Attachment ZIP URL keeps /nc subpath in plain text", plainText.Contains("https://cloud.example.test/nc/s/AbCd1234/download"), plainText);
        Check("Attachment plain text labels the link as ZIP download", plainText.Contains("ZIP download: https://cloud.example.test/nc/s/AbCd1234/download"), plainText);
        Check("Plain text stays plain", !plainText.Contains("<a "), plainText);
    }

    private static void TestCustomTemplateResolvesModeAwareLinkVariables()
    {
        const string template = "<p>{LINK_INTRO}</p><p>{LINK_LABEL}: <a href=\"{URL}\">{URL}</a></p>";
        BackendPolicyStatus policy = BuildCustomTemplatePolicy("<p>Legacy template: {URL}</p>", template);
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);

        string normalHtml = FileLinkHtmlBuilder.Build(result, new FileLinkRequest(), "custom", policy);
        string zipHtml = FileLinkHtmlBuilder.Build(result, BuildAttachmentRequest(), "custom", policy);
        string normal = FileLinkHtmlBuilder.BuildPlainText(result, new FileLinkRequest(), "custom", policy);
        string zip = FileLinkHtmlBuilder.BuildPlainText(result, BuildAttachmentRequest(), "custom", policy);

        Check("Custom normal template resolves LINK_INTRO", normal.Contains("Open the Nextcloud link below to view the share."), normal);
        Check("Custom normal template resolves LINK_LABEL", normal.Contains("Nextcloud link: https://cloud.example.test/nc/s/AbCd1234"), normal);
        Check("Custom attachment template resolves ZIP LINK_INTRO", zip.Contains("Download the shared files as a ZIP archive"), zip);
        Check("Custom attachment template resolves ZIP LINK_LABEL", zip.Contains("ZIP download: https://cloud.example.test/nc/s/AbCd1234/download"), zip);
        Check("Custom normal HTML uses the versioned template", normalHtml.Contains("Open the Nextcloud link below to view the share."), normalHtml);
        Check("Custom attachment HTML resolves the versioned template in ZIP mode", zipHtml.Contains("ZIP download"), zipHtml);
        Check("Versioned template takes precedence over compatibility template", !normal.Contains("Legacy template") && !normalHtml.Contains("Legacy template"), normal + normalHtml);
    }

    private static void TestLegacyCustomTemplateStillRenders()
    {
        BackendPolicyStatus policy = BuildCustomTemplatePolicy("<p>Legacy link: {URL}</p>");
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, new FileLinkRequest(), "custom", policy);

        Check("Legacy custom template still resolves its existing URL variable", plainText.Contains("Legacy link: https://cloud.example.test/nc/s/AbCd1234"), plainText);
        Check("Legacy custom template is not forced to contain new variables", !plainText.Contains("LINK_INTRO") && !plainText.Contains("LINK_LABEL"), plainText);
    }

    private static void TestBackendEffectiveLanguageLocalizesCustomTemplateCopy()
    {
        const string template = "<p>{LINK_INTRO}</p><p>{LINK_LABEL}: {URL}</p><p>{PASSWORD}</p><p>{RIGHTS}</p>";
        BackendPolicyStatus policy = BuildCustomTemplatePolicy("<p>Legacy: {URL}</p>", template, "de");
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        var request = new FileLinkRequest
        {
            PasswordSeparateEnabled = true,
            Permissions = FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create
        };

        string html = FileLinkHtmlBuilder.Build(result, request, "custom", policy);
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "custom", policy);

        foreach (string output in new[] { html, plainText })
        {
            Check("Backend template language localizes LINK_INTRO", output.Contains("Öffnen Sie den untenstehenden Nextcloud-Link"), output);
            Check("Backend template language localizes LINK_LABEL", output.Contains("Nextcloud-Link"), output);
            Check("Backend template language localizes separate-password hint", output.Contains("Das Passwort wird in einer separaten E-Mail gesendet."), output);
            Check("Backend template language localizes permission names", output.Contains("Lesen") && output.Contains("Hochladen") && output.Contains("Bearbeiten") && output.Contains("Löschen"), output);
        }
    }

    private static void TestOlderBackendModeAwareTemplateStillRenders()
    {
        const string template = "<p>{LINK_INTRO}</p><p>{LINK_LABEL}: <a href=\"{URL}\">{URL}</a></p>";
        BackendPolicyStatus policy = BuildCustomTemplatePolicy(template);
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", string.Empty);
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, new FileLinkRequest(), "custom", policy);

        Check("Older backend template field still resolves LINK_INTRO", plainText.Contains("Open the Nextcloud link below to view the share."), plainText);
        Check("Older backend template field still resolves LINK_LABEL", plainText.Contains("Nextcloud link: https://cloud.example.test/nc/s/AbCd1234"), plainText);
    }

    private static void TestSecretLinkLabelHidesLongUrlInHtml()
    {
        const string secretUrl = "https://cloud.example.test/index.php/apps/secrets/share/1234567890#VeryLongLocalKey";
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", secretUrl);
        string html = FileLinkHtmlBuilder.BuildPasswordOnly(result, "en", null, true);

        Check("Secret password mail uses compact link label", html.Contains(">Secret link<"), html);
        Check("Secret password mail keeps URL in href", html.Contains("href=\"" + secretUrl + "\""), html);
        Check("Secret password mail does not render the long URL as visible text", !html.Contains(">" + secretUrl + "<"), html);
    }

    private static FileLinkResult BuildResult(string shareUrl, string token, string password)
    {
        return new FileLinkResult(
            shareUrl,
            "42",
            token,
            password,
            new DateTime(2026, 7, 7),
            FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create,
            "Folder",
            "NC Connector/Folder");
    }

    private static FileLinkRequest BuildAttachmentRequest()
    {
        return new FileLinkRequest
        {
            ShareName = "Folder",
            AttachmentMode = true,
            PasswordSeparateEnabled = false,
            NoteEnabled = false,
            Permissions = FileLinkPermissionFlags.Read | FileLinkPermissionFlags.Create
        };
    }

    private static BackendPolicyStatus BuildCustomTemplatePolicy(string template, string versionedTemplate = null, string effectiveLanguage = null)
    {
        var sharePolicy = new Dictionary<string, object>
        {
            { "share_html_block_template", template }
        };
        if (!string.IsNullOrWhiteSpace(versionedTemplate))
        {
            sharePolicy.Add("share_html_block_template_v2", versionedTemplate);
        }
        if (!string.IsNullOrWhiteSpace(effectiveLanguage))
        {
            sharePolicy.Add("share_html_block_effective_language", effectiveLanguage);
        }
        var empty = new Dictionary<string, object>();
        return new BackendPolicyStatus(
            true,
            true,
            true,
            false,
            string.Empty,
            "policy",
            string.Empty,
            true,
            true,
            "active",
            sharePolicy,
            empty,
            empty,
            empty,
            empty,
            empty);
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $vendorDir = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\vendor\htmlsanitizer"
    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\BackendPolicyStatus.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkPermissions.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkSelection.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkRequest.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkResult.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\SharePasswordDeliveryMode.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\BrandingAssets.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkHtmlBuilder.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HtmlTemplateSanitizer.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HtmlToPlainTextConverter.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\Strings.cs")
    )
    $references = @(
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Drawing.dll",
        "/reference:System.Web.dll",
        "/reference:System.Web.Extensions.dll"
    )
    Get-ChildItem -Path $vendorDir -Filter "*.dll" | ForEach-Object {
        $references += "/reference:$($_.FullName)"
    }
    $resources = @(
        "/resource:$((Resolve-Path (Join-Path $ProjectRoot 'src\NcTalkOutlookAddIn\Resources\_locales\en\messages.json')).Path),OutlookFileLinkRenderingTests.Resources._locales.en.messages.json",
        "/resource:$((Resolve-Path (Join-Path $ProjectRoot 'src\NcTalkOutlookAddIn\Resources\_locales\de\messages.json')).Path),OutlookFileLinkRenderingTests.Resources._locales.de.messages.json"
    )

    $exe = Join-Path $TempRoot "OutlookFileLinkRenderingTests.exe"
    & $csc /nologo /nowarn:1702 /target:exe "/out:$exe" @references @resources @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Get-ChildItem -Path $vendorDir -Filter "*.dll" | Copy-Item -Force -Destination $TempRoot

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

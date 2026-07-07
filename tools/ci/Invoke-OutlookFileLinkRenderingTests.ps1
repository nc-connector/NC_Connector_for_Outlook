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
        TestAttachmentModeKeepsNextcloudSubpath();
        TestPlainTextKeepsNextcloudSubpath();
        TestSecretLinkLabelHidesLongUrlInHtml();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " FileLink rendering test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook FileLink rendering tests passed.");
        return 0;
    }

    private static void TestAttachmentModeKeepsNextcloudSubpath()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = BuildAttachmentRequest();
        string html = FileLinkHtmlBuilder.Build(result, request, "en");

        Check("Attachment ZIP URL keeps /nc subpath in HTML", html.Contains("https://cloud.example.test/nc/s/AbCd1234/download"), html);
        Check("Attachment ZIP URL does not drop /nc subpath in HTML", !html.Contains("https://cloud.example.test/s/AbCd1234/download"), html);
    }

    private static void TestPlainTextKeepsNextcloudSubpath()
    {
        FileLinkResult result = BuildResult("https://cloud.example.test/nc/s/AbCd1234", "AbCd1234", "Secret!");
        FileLinkRequest request = BuildAttachmentRequest();
        string plainText = FileLinkHtmlBuilder.BuildPlainText(result, request, "en");

        Check("Attachment ZIP URL keeps /nc subpath in plain text", plainText.Contains("https://cloud.example.test/nc/s/AbCd1234/download"), plainText);
        Check("Plain text stays plain", !plainText.Contains("<a "), plainText);
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

    $exe = Join-Path $TempRoot "OutlookFileLinkRenderingTests.exe"
    & $csc /nologo /nowarn:1702 /target:exe "/out:$exe" @references @sources
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

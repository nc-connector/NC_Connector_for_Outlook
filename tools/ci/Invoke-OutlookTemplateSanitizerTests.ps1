Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-sanitizer-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "OutlookTemplateSanitizerTests.cs"
    @'
using System;
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
    }
}

internal static class OutlookTemplateSanitizerTests
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
        TestShareTemplateSanitizer();
        TestEmailSignatureSanitizer();
        TestTalkAppointmentCompatibilityTransform();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " sanitizer test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook template sanitizer tests passed.");
        return 0;
    }

    private static void TestShareTemplateSanitizer()
    {
        string html = "<div onclick=\"alert(1)\"><script>alert(1)</script><a href=\"javascript:alert(1)\">bad</a><a href=\"https://example.test/path\">ok</a>{URL}</div>";
        string sanitized = HtmlTemplateSanitizer.SanitizeShareTemplateHtml(html);
        Check("Share sanitizer removes script tags", !sanitized.Contains("<script"), sanitized);
        Check("Share sanitizer removes event handlers", !sanitized.Contains("onclick"), sanitized);
        Check("Share sanitizer removes javascript URLs", !sanitized.Contains("javascript:"), sanitized);
        Check("Share sanitizer keeps https URLs", sanitized.Contains("https://example.test/path"), sanitized);
        Check("Share sanitizer keeps placeholders", sanitized.Contains("{URL}"), sanitized);
    }

    private static void TestEmailSignatureSanitizer()
    {
        string html = "<table><tr><td style=\"color:#123456\">Name</td></tr></table>";
        string sanitized = HtmlTemplateSanitizer.SanitizeEmailSignatureTemplateHtml(html);
        Check("Signature sanitizer keeps table layout", sanitized.Contains("<table") && sanitized.Contains("<td"), sanitized);
        Check("Signature sanitizer keeps safe inline color", sanitized.Contains("color:"), sanitized);
    }

    private static void TestTalkAppointmentCompatibilityTransform()
    {
        string html = "<div style=\"display:flex;border-radius:8px;color:#0082C9\"><a href=\"https://example.test\" style=\"color:#0082C9\">Talk</a></div>";
        string prepared = HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(html);
        Check("Talk appointment transform strips flex layout", !prepared.Contains("display:flex"), prepared);
        Check("Talk appointment transform strips border radius", !prepared.Contains("border-radius"), prepared);
        Check("Talk appointment transform keeps link", prepared.Contains("https://example.test"), prepared);
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
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HtmlTemplateSanitizer.cs")
    )
    $references = @(
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Web.dll"
    )
    Get-ChildItem -Path $vendorDir -Filter "*.dll" | ForEach-Object {
        $references += "/reference:$($_.FullName)"
    }

    $exe = Join-Path $TempRoot "OutlookTemplateSanitizerTests.exe"
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

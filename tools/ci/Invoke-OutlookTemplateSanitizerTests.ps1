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
using System.IO;
using System.Reflection;
using System.Text;
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
        TestResolverBoundaries();
        TestDependencyScopes();
        TestShareTemplateSanitizer();
        TestDeeplyNestedTemplateSanitizer();
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

    private static void TestDeeplyNestedTemplateSanitizer()
    {
        var payload = new StringBuilder();
        for (int i = 0; i < 513; i++)
        {
            payload.Append("<div>");
        }
        payload.Append("<template><script>alert(1)</script><img src=\"x\" onerror=\"alert(2)\"></template>");
        for (int i = 0; i < 513; i++)
        {
            payload.Append("</div>");
        }

        string sanitized = HtmlTemplateSanitizer.SanitizeShareTemplateHtml(payload.ToString());
        Check("Deep nesting removes template elements", !sanitized.Contains("<template"), sanitized);
        Check("Deep nesting removes script elements", !sanitized.Contains("<script"), sanitized);
        Check("Deep nesting removes event handlers", !sanitized.Contains("onerror"), sanitized);
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

    private static void TestResolverBoundaries()
    {
        MethodInfo resolver = typeof(HtmlTemplateSanitizer).GetMethod(
            "ResolveSanitizerDependency",
            BindingFlags.NonPublic | BindingFlags.Static);
        Check("Sanitizer dependency resolver is available", resolver != null);
        if (resolver == null)
        {
            return;
        }

        Assembly codePages = Assembly.LoadFrom(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System.Text.Encoding.CodePages.dll"));
        const string memoryPrefix = "System.Memory, Version=";
        const string memorySuffix =
            ", Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51";

        Assembly resolved = InvokeResolver(
            resolver,
            memoryPrefix + "4.0.1.1" + memorySuffix,
            codePages);
        Check(
            "Known sanitizer dependency request resolves to packaged System.Memory",
            resolved != null && resolved.GetName().Version == new Version(4, 0, 5, 0),
            resolved == null ? "null" : resolved.FullName);

        Check(
            "Higher dependency versions are not redirected",
            InvokeResolver(resolver, memoryPrefix + "99.0.0.0" + memorySuffix, codePages) == null);
        Check(
            "Wrong dependency public key tokens are not redirected",
            InvokeResolver(
                resolver,
                memoryPrefix + "4.0.1.1, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
                codePages) == null);
        Check(
            "Non-neutral dependency cultures are not redirected",
            InvokeResolver(
                resolver,
                memoryPrefix + "4.0.1.1, Culture=de-DE, PublicKeyToken=cc7b13ffcd2ddd51",
                codePages) == null);
        Check(
            "Unrelated requesting assemblies are not redirected",
            InvokeResolver(resolver, memoryPrefix + "4.0.1.1" + memorySuffix, typeof(string).Assembly) == null);
        Check(
            "Requests without a requesting assembly are not redirected",
            InvokeResolver(resolver, memoryPrefix + "4.0.1.1" + memorySuffix, null) == null);
        Check(
            "Unknown dependency names are not redirected",
            InvokeResolver(
                resolver,
                "Unrelated.Dependency, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null",
                codePages) == null);
    }

    private static void TestDependencyScopes()
    {
        MethodInfo resolver = typeof(HtmlTemplateSanitizer).GetMethod(
            "ResolveSanitizerDependency", BindingFlags.NonPublic | BindingFlags.Static);
        const string request = "AngleSharp, Version=0.17.0.0, Culture=neutral, PublicKeyToken=e83494dcdc6d31ea";
        Check("Scoped requester-less HTML dependency resolves", HtmlTemplateSanitizer.RunWithHtmlDependencies(
            () => InvokeResolver(resolver, request, null)) != null);
        Check("Completed HTML scope does not enable unrelated resolution", InvokeResolver(resolver, request, null) == null);
        HtmlTemplateSanitizer.RunWithHtmlDependencies(() =>
        {
            try
            {
                HtmlTemplateSanitizer.RunWithHtmlDependencies<bool>(() => { throw new InvalidOperationException("test"); });
            }
            catch (InvalidOperationException) { }
            Check("Nested failed HTML operation retains the outer scope", InvokeResolver(resolver, request, null) != null);
            Check("Scoped HTML work cannot redirect a higher version", InvokeResolver(resolver,
                request.Replace("0.17.0.0", "99.0.0.0"), null) == null);
            return true;
        });
        Check("Nested HTML scopes restore requester restrictions", InvokeResolver(resolver, request, null) == null);
        try
        {
            HtmlTemplateSanitizer.RunWithHtmlDependencies<bool>(() => { throw new InvalidOperationException("test"); });
        }
        catch (InvalidOperationException) { }
        Check("Failed HTML scope restores requester restrictions", InvokeResolver(resolver, request, null) == null);
    }

    private static Assembly InvokeResolver(MethodInfo resolver, string requestedName, Assembly requester)
    {
        return resolver.Invoke(
            null,
            new object[] { null, new ResolveEventArgs(requestedName, requester) }) as Assembly;
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
        "/reference:System.Numerics.dll",
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

Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-unit-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "OutlookUtilityTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static void LogException(string category, string message, Exception ex) { }
    }

    internal static class LogCategories
    {
        internal const string Core = "core";
    }
}

internal static class OutlookUtilityTests
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

    private static void Equal(string name, object expected, object actual)
    {
        Check(name, object.Equals(expected, actual), "expected '" + expected + "', got '" + actual + "'");
    }

    public static int Main()
    {
        TestPasswordGenerator();
        TestSizeFormatting();
        TestVersionParsing();
        TestPlainTextUtilities();
        TestBasicAuth();
        TestNcJson();
        TestBackendPolicyStatus();
        TestHtmlToPlainText();
        TestSecretsCrypto();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " unit test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook utility unit tests passed.");
        return 0;
    }

    private static void TestPasswordGenerator()
    {
        string generated = PasswordGenerator.GenerateLocalPassword(4);
        Check("PasswordGenerator enforces minimum length", generated.Length == 8, "length=" + generated.Length);
        Check("PasswordGenerator uses non-empty alphabet", generated.Trim().Length == generated.Length);
    }

    private static void TestSizeFormatting()
    {
        Equal("SizeFormatting 1 MiB", "1.0 MB", SizeFormatting.FormatMegabytes(1024 * 1024, CultureInfo.InvariantCulture));
        Equal("SizeFormatting clamps negative values", "0.0 MB", SizeFormatting.FormatMegabytes(-12, CultureInfo.InvariantCulture));
    }

    private static void TestVersionParsing()
    {
        Version version;
        Check("NextcloudVersionHelper parses version with edition", NextcloudVersionHelper.TryParse("31.0.4 Enterprise", out version));
        Equal("NextcloudVersionHelper parsed edition version", new Version(31, 0, 4), version);
        Check("NextcloudVersionHelper parses pre-release prefix", NextcloudVersionHelper.TryParse("32.0.0-beta1", out version));
        Equal("NextcloudVersionHelper parsed pre-release prefix", new Version(32, 0, 0), version);
        Check("NextcloudVersionHelper rejects empty", !NextcloudVersionHelper.TryParse(" ", out version));
    }

    private static void TestPlainTextUtilities()
    {
        Equal("PlainTextUtilities normalizes CRLF", "a\r\nb\r\nc", PlainTextUtilities.NormalizeCrLf("a\nb\rc"));
        Equal("PlainTextUtilities trims after normalize", "a", PlainTextUtilities.NormalizeCrLfAndTrim("\n a \n"));
    }

    private static void TestBasicAuth()
    {
        string expected = "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes("üser:päss"));
        Equal("HttpAuthUtilities uses UTF-8", expected, HttpAuthUtilities.BuildBasicAuthHeader("üser", "päss"));
    }

    private static void TestNcJson()
    {
        string prepared = NcJson.PrepareJsonPayload(")]}',\n{\"ok\":true}");
        Equal("NcJson removes Angular XSSI prefix", "{\"ok\":true}", prepared);
        prepared = NcJson.PrepareJsonPayload("while(1); {\"ok\":true}");
        Equal("NcJson removes while prefix", "{\"ok\":true}", prepared);
        IDictionary<string, object> payload = NcJson.DeserializeObject("{\"number\":\"7\",\"flag\":true,\"ocs\":{\"data\":{\"id\":\"abc\"},\"meta\":{\"message\":\"OK\"}}}");
        Equal("NcJson GetInt parses string", 7, NcJson.GetInt(payload, "number"));
        Equal("NcJson GetOcsData", "abc", NcJson.GetString(NcJson.GetOcsData(payload), "id"));
    }

    private static void TestBackendPolicyStatus()
    {
        var sharePolicy = new Dictionary<string, object> { { "share_set_password", true }, { "share_expire_days", "14" } };
        var shareEditable = new Dictionary<string, object> { { "share_set_password", false }, { "share_expire_days", true } };
        var status = new BackendPolicyStatus(
            true,
            true,
            true,
            false,
            "",
            "policy",
            "policy_active",
            true,
            true,
            "active",
            sharePolicy,
            new Dictionary<string, object>(),
            new Dictionary<string, object>(),
            shareEditable,
            new Dictionary<string, object>(),
            new Dictionary<string, object>());

        bool value;
        int days;
        Check("BackendPolicyStatus locks non-editable value", status.IsLocked("share", "share_set_password"));
        Check("BackendPolicyStatus does not lock editable value", !status.IsLocked("share", "share_expire_days"));
        Check("BackendPolicyStatus converts bool", status.TryGetPolicyBool("share", "share_set_password", out value) && value);
        Check("BackendPolicyStatus converts int string", status.TryGetPolicyInt("share", "share_expire_days", out days) && days == 14);
        Check("BackendPolicyStatus bool accepts yes", BackendPolicyStatus.TryConvertBool("yes", out value) && value);
    }

    private static void TestHtmlToPlainText()
    {
        string plain = HtmlToPlainTextConverter.Convert("<p>Hello <a href=\"https://example.test\">link</a></p><ul><li>One</li><li>Two</li></ul><script>alert(1)</script>");
        Check("HtmlToPlainText keeps anchor href", plain.Contains("link (https://example.test)"), plain);
        Check("HtmlToPlainText renders list items", plain.Contains("- One") && plain.Contains("- Two"), plain);
        Check("HtmlToPlainText skips script content", !plain.Contains("alert"), plain);
    }

    private static void TestSecretsCrypto()
    {
        SecretsEncryptedPayload payload = SecretsCrypto.EncryptToSecretsPayload("secret");
        byte[] key = Convert.FromBase64String(payload.Key);
        byte[] iv = Convert.FromBase64String(payload.Iv);
        byte[] encrypted = Convert.FromBase64String(payload.Encrypted);
        Equal("SecretsCrypto key length", 32, key.Length);
        Equal("SecretsCrypto iv length", 12, iv.Length);
        Check("SecretsCrypto includes authentication tag", encrypted.Length > "secret".Length);
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\PasswordGenerator.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\SizeFormatting.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudVersionHelper.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\PlainTextUtilities.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HttpAuthUtilities.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\BackendPolicyStatus.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HtmlToPlainTextConverter.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\SecretsCrypto.cs")
    )

    $exe = Join-Path $TempRoot "OutlookUtilityTests.exe"
    $references = @(
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Web.Extensions.dll",
        ("/reference:" + (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\vendor\htmlsanitizer\AngleSharp.dll")),
        ("/reference:" + (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\vendor\htmlsanitizer\AngleSharp.Css.dll"))
    )

    & $csc /nologo /target:exe "/out:$exe" @references @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Get-ChildItem (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\vendor\htmlsanitizer") -Filter "*.dll" |
        Copy-Item -Force -Destination $TempRoot

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

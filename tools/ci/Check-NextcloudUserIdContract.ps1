Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-user-id-contract-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

function Assert-Check {
    Param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
    Write-Host "[OK] $Message"
}

try {
    $fixturePath = Join-Path $ProjectRoot "tests\contracts\current-user.outlook.json"
    Assert-Check (Test-Path -LiteralPath $fixturePath) "Current-user contract fixture exists"
    try {
        $fixture = Get-Content -Raw -LiteralPath $fixturePath | ConvertFrom-Json
    } catch {
        throw "Current-user contract fixture is invalid JSON. $($_.Exception.Message)"
    }
    Assert-Check ($fixture.ocs.data.id -eq "canonical-user") "Current-user fixture exposes the canonical UID"
    Assert-Check ($fixture.ocs.data.email -eq "login@example.test") "Current-user fixture covers an email login alias"

    $identityServicePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\NextcloudUserIdentityService.cs"
    $fileLinkServicePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkService.cs"
    $addressBookPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\IfbAddressBookCache.cs"
    $freeBusyPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FreeBusyServer.cs"
    $appointmentPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Controllers\TalkAppointmentController.cs"
    $httpClientPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\NcHttpClient.cs"
    foreach ($path in @($identityServicePath, $fileLinkServicePath, $addressBookPath, $freeBusyPath, $appointmentPath, $httpClientPath)) {
        Assert-Check (Test-Path -LiteralPath $path) ("Identity contract source exists: " + (Split-Path -Leaf $path))
    }

    $identitySource = Get-Content -Raw -LiteralPath $identityServicePath
    $fileLinkSource = Get-Content -Raw -LiteralPath $fileLinkServicePath
    $addressBookSource = Get-Content -Raw -LiteralPath $addressBookPath
    $freeBusySource = Get-Content -Raw -LiteralPath $freeBusyPath
    $appointmentSource = Get-Content -Raw -LiteralPath $appointmentPath
    $httpClientSource = Get-Content -Raw -LiteralPath $httpClientPath

    Assert-Check ($identitySource.Contains('/ocs/v2.php/cloud/user?format=json')) "Canonical UID resolver uses the documented current-user endpoint"
    Assert-Check ($identitySource.Contains('NcJson.GetOcsData(payload), "id"')) "Canonical UID resolver reads ocs.data.id"
    Assert-Check ([regex]::IsMatch($fileLinkSource, 'NextcloudUserIdentityService\.ResolveCurrentUserId\(\s*_configuration\s*\)')) "FileLink resolves the canonical UID"
    Assert-Check (-not $fileLinkSource.Contains('string username = _configuration.Username')) "FileLink does not use the authentication login as DAV UID"
    Assert-Check ($addressBookSource.Contains('NextcloudUserIdentityService.ResolveCurrentUserId(configuration)')) "CardDAV resolves the canonical UID"
    Assert-Check ($freeBusySource.Contains('NextcloudUserIdentityService.ResolveCurrentUserId(_configuration)')) "CalDAV resolves the canonical UID"
    Assert-Check ($appointmentSource.Contains('NextcloudUserIdentityService.ResolveCurrentUserId(configuration)')) "Appointment address-book lookup resolves the canonical UID"
    Assert-Check ($httpClientSource.Contains('_username = configuration.Username')) "Basic Auth still uses the configured login"

    $testSource = Join-Path $TempRoot "NextcloudUserIdContractTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.Net;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static void Log(string category, string message) { }
        internal static void LogApi(string message) { }
        internal static void LogException(string category, string message, Exception ex) { }
    }

    internal static class LogCategories
    {
        internal const string Api = "api";
    }
}

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class NcHttpRequestOptions
    {
        internal string Method { get; set; }
        internal string Url { get; set; }
        internal string Accept { get; set; }
        internal int TimeoutMs { get; set; }
        internal bool IncludeAuthHeader { get; set; }
        internal bool IncludeOcsApiHeader { get; set; }
        internal bool ParseJson { get; set; }
    }

    internal sealed class NcHttpResponse
    {
        internal bool HasHttpResponse { get; set; }
        internal HttpStatusCode StatusCode { get; set; }
        internal IDictionary<string, object> ParsedJson { get; set; }
        internal string ResponseText { get; set; }
        internal Exception TransportException { get; set; }
    }

    internal sealed class NcHttpClient
    {
        internal static int SendCount;
        internal static string LastAuthenticationLogin;
        internal static NcHttpRequestOptions LastOptions;
        internal static NcHttpResponse NextResponse;

        internal NcHttpClient(TalkServiceConfiguration configuration)
        {
            LastAuthenticationLogin = configuration.Username;
        }

        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            SendCount++;
            LastOptions = options;
            return NextResponse;
        }
    }
}

internal static class NextcloudUserIdContractTests
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
        TalkServiceConfiguration configuration = new TalkServiceConfiguration(
            "https://cloud.example.test/nextcloud/",
            "login@example.test",
            "app-password");
        NcHttpClient.NextResponse = SuccessfulResponse("canonical-user", "login@example.test");

        string userId = NextcloudUserIdentityService.ResolveCurrentUserId(configuration, true);
        Check("Resolver returns ocs.data.id", userId == "canonical-user", userId);
        Check("Resolver preserves the Nextcloud subfolder", NcHttpClient.LastOptions.Url == "https://cloud.example.test/nextcloud/ocs/v2.php/cloud/user?format=json", NcHttpClient.LastOptions.Url);
        Check("Current-user request includes Basic Auth", NcHttpClient.LastOptions.IncludeAuthHeader);
        Check("Current-user request includes the OCS header", NcHttpClient.LastOptions.IncludeOcsApiHeader);
        Check("Authentication keeps the configured email login", NcHttpClient.LastAuthenticationLogin == "login@example.test", NcHttpClient.LastAuthenticationLogin);

        int sendsAfterResolution = NcHttpClient.SendCount;
        string cachedUserId = NextcloudUserIdentityService.ResolveCurrentUserId(configuration);
        Check("Canonical UID is cached for the Outlook session", cachedUserId == "canonical-user" && NcHttpClient.SendCount == sendsAfterResolution);

        NcHttpClient.NextResponse = SuccessfulResponse(null, "login@example.test");
        bool rejectedMissingId = false;
        try
        {
            NextcloudUserIdentityService.ResolveCurrentUserId(configuration, true);
        }
        catch (TalkServiceException)
        {
            rejectedMissingId = true;
        }
        Check("Missing ocs.data.id is rejected without an email fallback", rejectedMissingId);

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " Nextcloud user ID contract test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook Nextcloud user ID contract tests passed.");
        return 0;
    }

    private static NcHttpResponse SuccessfulResponse(string id, string email)
    {
        var data = new Dictionary<string, object>();
        if (id != null)
        {
            data["id"] = id;
        }
        data["email"] = email;
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = HttpStatusCode.OK,
            ParsedJson = new Dictionary<string, object>
            {
                { "ocs", new Dictionary<string, object> { { "data", data } } }
            },
            ResponseText = "{}"
        };
    }
}
'@ | Set-Content -LiteralPath $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    Assert-Check (Test-Path -LiteralPath $csc) "C# compiler is available"
    $exe = Join-Path $TempRoot "NextcloudUserIdContractTests.exe"
    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\NextcloudUserIdentityService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceConfiguration.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceException.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs")
    )
    & $csc /nologo /target:exe "/out:$exe" /reference:System.dll /reference:System.Core.dll /reference:System.Web.Extensions.dll @sources
    if ($LASTEXITCODE -ne 0) {
        throw "Nextcloud user ID contract test compilation failed with exit code $LASTEXITCODE."
    }
    & $exe
    if ($LASTEXITCODE -ne 0) {
        throw "Nextcloud user ID contract tests failed with exit code $LASTEXITCODE."
    }
} finally {
    Remove-Item -LiteralPath $TempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

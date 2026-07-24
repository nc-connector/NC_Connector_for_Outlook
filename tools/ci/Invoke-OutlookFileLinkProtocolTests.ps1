Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path `
    ([System.IO.Path]::GetTempPath()) `
    ("nc4ol-filelink-protocol-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "FileLinkProtocolTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Models
{
    [Flags]
    internal enum FileLinkPermissionFlags
    {
        None = 0,
        Read = 1,
        Write = 2,
        Create = 4,
        Delete = 8
    }

    internal sealed class FileLinkRequest
    {
        internal FileLinkPermissionFlags Permissions { get; set; }
        internal bool PasswordEnabled { get; set; }
        internal string Password { get; set; }
        internal bool ExpireEnabled { get; set; }
        internal DateTime? ExpireDate { get; set; }
        internal bool NoteEnabled { get; set; }
        internal string Note { get; set; }
    }

    internal sealed class FileLinkUploadPhaseProgress
    {
    }
}

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static void Log(string category, string message) { }
        internal static void LogException(
            string category,
            string message,
            Exception ex) { }
    }

    internal static class LogCategories
    {
        internal const string Api = "api";
        internal const string FileLink = "filelink";
    }

    internal static class Strings
    {
        internal static string ErrorServerUnavailable
        {
            get { return "server unavailable"; }
        }

        internal static string FileLinkUploadInsufficientStorage
        {
            get { return "insufficient storage"; }
        }

        internal static string FileLinkWizardFolderCheckFailedFormat
        {
            get { return "folder check failed: {0}"; }
        }

        internal static string FileLinkWizardUploadFailed
        {
            get { return "upload failed"; }
        }
    }

    internal static class ParallelExecution
    {
        internal static void RethrowFirstFailure(
            AggregateException exception,
            CancellationToken cancellationToken)
        {
            throw exception;
        }
    }
}

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class TalkServiceException : Exception
    {
        internal TalkServiceException(
            string message,
            bool isAuthenticationError,
            HttpStatusCode statusCode,
            string responseText)
            : base(message)
        {
            IsAuthenticationError = isAuthenticationError;
            StatusCode = statusCode;
            ResponseText = responseText;
        }

        internal bool IsAuthenticationError { get; private set; }
        internal HttpStatusCode StatusCode { get; private set; }
        internal string ResponseText { get; private set; }
    }

    internal sealed class NcHttpRequestOptions
    {
        internal string Method { get; set; }
        internal string Url { get; set; }
        internal string Payload { get; set; }
        internal string Accept { get; set; }
        internal string ContentType { get; set; }
        internal int TimeoutMs { get; set; }
        internal int ReadWriteTimeoutMs { get; set; }
        internal bool IncludeAuthHeader { get; set; }
        internal bool IncludeOcsApiHeader { get; set; }
        internal bool ParseJson { get; set; }
        internal long ContentLength { get; set; }
        internal bool AllowWriteStreamBuffering { get; set; }
        internal CancellationToken CancellationToken { get; set; }
        internal int ConnectionLimit { get; set; }
        internal IDictionary<string, string> Headers { get; set; }
        internal Action<Stream> BodyWriter { get; set; }
    }

    internal sealed class NcHttpResponse
    {
        internal bool HasHttpResponse { get; set; }
        internal HttpStatusCode StatusCode { get; set; }
        internal string ResponseText { get; set; }
        internal IDictionary<string, object> ParsedJson { get; set; }
        internal WebException TransportException { get; set; }
        internal IDictionary<string, string> Headers { get; set; }
    }

    internal sealed class NcHttpClient
    {
        private readonly Func<NcHttpRequestOptions, NcHttpResponse> _sender;

        internal NcHttpClient(
            Func<NcHttpRequestOptions, NcHttpResponse> sender)
        {
            _sender = sender;
        }

        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            return _sender(options);
        }
    }

    internal sealed class FileLinkUploadContext
    {
        internal string NormalizedBaseUrl { get; set; }
        internal string UserId { get; set; }
        internal string RelativeFolderPath { get; set; }
    }

    internal sealed class FileLinkUploadPlan
    {
        internal FileLinkUploadPlan()
        {
            DirectoriesToCreate = new List<string>();
            Files = new List<object>();
        }

        internal IList<string> DirectoriesToCreate { get; private set; }
        internal IList<object> Files { get; private set; }
        internal long TotalBytes { get; set; }
    }

    internal sealed class FileLinkFolderProgressReporter
    {
        internal FileLinkFolderProgressReporter(
            IProgress<FileLinkUploadPhaseProgress> progress,
            int totalFolders,
            int totalFiles,
            long totalBytes)
        {
        }

        internal void Report(int completedFolders, bool force)
        {
        }
    }
}

internal static class FileLinkProtocolTests
{
    private static int failures;

    private static void Check(
        string name,
        bool condition,
        string detail = "")
    {
        if (condition)
        {
            Console.WriteLine("[OK] " + name);
            return;
        }

        failures++;
        Console.Error.WriteLine(
            "[FAIL] "
            + name
            + (string.IsNullOrEmpty(detail) ? "" : ": " + detail));
    }

    private static void Equal(
        string name,
        object expected,
        object actual)
    {
        Check(
            name,
            object.Equals(expected, actual),
            "expected '" + expected + "', got '" + actual + "'");
    }

    public static int Main()
    {
        TestAutoMkcolHeader();
        TestDavPathNormalization();
        TestMissingResourcePreflight();
        TestExistingResourcePreflight();
        TestUnauthorizedPreflight();
        TestKnownRootCollision();
        TestIndeterminateRootCollision();
        TestOwnedDirectoryRecovery();
        TestInsufficientStorage();
        TestShareCreateSingleRequest();
        TestShareCreateRecovery();
        TestShareCreateTransientRecovery();
        TestShareCreateMalformedRecovery();
        TestShareCreateAbsentRecovery();
        TestShareCreateUnknownRecovery();
        TestShareOcsFailure();

        if (failures > 0)
        {
            Console.Error.WriteLine(
                failures + " FileLink protocol test(s) failed.");
            return 1;
        }

        Console.WriteLine("All FileLink protocol tests passed.");
        return 0;
    }

    private static void TestAutoMkcolHeader()
    {
        Equal(
            "Direct upload header matches the NC32 server",
            "X-NC-WebDAV-Auto-Mkcol",
            FileLinkUploadPolicy.AutoMkcolHeaderName);
    }

    private static void TestDavPathNormalization()
    {
        Equal(
            "DAV URL neutralizes parent-directory segments",
            "https://cloud.example.test/remote.php/dav/files/user/safe/__/file.txt",
            FileLinkDavClient.BuildFileUrl(
                "https://cloud.example.test",
                "user",
                "safe/../file.txt"));
    }

    private static void TestMissingResourcePreflight()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            return Http(HttpStatusCode.NotFound);
        });

        bool exists = client.ResourceExists(
            "https://cloud.example.test/remote.php/dav/files/user/NC%20Connector/20260724_share",
            "folder check failed: {0}",
            CancellationToken.None);
        Check("Missing share target passes the preflight", !exists);
        Equal("Missing target sends one probe", 1, requests.Count);
        Equal("Share target probe uses PROPFIND", "PROPFIND", requests[0].Method);
        Equal("Share target probe uses depth zero", "0", requests[0].Headers["Depth"]);
    }

    private static void TestExistingResourcePreflight()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            return new NcHttpResponse
            {
                HasHttpResponse = true,
                StatusCode = (HttpStatusCode)207,
                ResponseText =
                    "<?xml version=\"1.0\"?>"
                    + "<d:multistatus xmlns:d=\"DAV:\">"
                    + "<d:response><d:propstat><d:prop>"
                    + "<d:getcontentlength>12</d:getcontentlength>"
                    + "</d:prop></d:propstat></d:response>"
                    + "</d:multistatus>"
            };
        });

        bool exists = client.ResourceExists(
            "https://cloud.example.test/remote.php/dav/files/user/NC%20Connector/20260724_share",
            "folder check failed: {0}",
            CancellationToken.None);
        Check(
            "Any existing resource blocks the manual share target",
            exists);
        Equal("Existing target sends one probe", 1, requests.Count);
    }

    private static void TestUnauthorizedPreflight()
    {
        var client = new FileLinkDavClient(options =>
            Http(HttpStatusCode.Unauthorized));
        try
        {
            client.ResourceExists(
                "https://cloud.example.test/remote.php/dav/files/user/NC%20Connector/20260724_share",
                "folder check failed: {0}",
                CancellationToken.None);
            Check(
                "Unauthorized preflight fails closed",
                false,
                "no exception");
        }
        catch (TalkServiceException ex)
        {
            Check(
                "Unauthorized preflight fails closed",
                ex.IsAuthenticationError);
            Equal(
                "Unauthorized preflight keeps the HTTP status",
                HttpStatusCode.Unauthorized,
                ex.StatusCode);
        }
    }

    private static void TestKnownRootCollision()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            return Http(HttpStatusCode.MethodNotAllowed);
        });

        bool created = client.TryCreateShareRoot(
            "https://cloud.example.test",
            "user",
            "NC Connector/share",
            CancellationToken.None);
        Check("Known MKCOL 405 remains a collision", !created);
        Equal("Known collision sends one request", 1, requests.Count);
        Equal("Known collision does not probe", "MKCOL", requests[0].Method);
    }

    private static void TestIndeterminateRootCollision()
    {
        var requests = new List<NcHttpRequestOptions>();
        int requestIndex = 0;
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            requestIndex++;
            if (requestIndex == 1)
            {
                return new NcHttpResponse
                {
                    HasHttpResponse = false,
                    TransportException = new WebException(
                        "timeout",
                        WebExceptionStatus.Timeout)
                };
            }
            return Http(HttpStatusCode.MethodNotAllowed);
        });

        bool created = client.TryCreateShareRoot(
            "https://cloud.example.test",
            "user",
            "NC Connector/share",
            CancellationToken.None);
        Check(
            "Indeterminate MKCOL followed by 405 remains a collision",
            !created);
        Equal("Indeterminate collision request count", 2, requests.Count);
        Equal("Indeterminate collision does not probe", "MKCOL", requests[1].Method);
    }

    private static void TestOwnedDirectoryRecovery()
    {
        var requests = new List<NcHttpRequestOptions>();
        int requestIndex = 0;
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            requestIndex++;
            if (requestIndex == 1)
            {
                return Http(HttpStatusCode.GatewayTimeout);
            }
            if (requestIndex == 2)
            {
                return Http(HttpStatusCode.MethodNotAllowed);
            }
            return new NcHttpResponse
            {
                HasHttpResponse = true,
                StatusCode = (HttpStatusCode)207,
                ResponseText =
                    "<?xml version=\"1.0\"?>"
                    + "<d:multistatus xmlns:d=\"DAV:\">"
                    + "<d:response><d:propstat><d:prop>"
                    + "<d:resourcetype><d:collection/>"
                    + "</d:resourcetype></d:prop>"
                    + "</d:propstat></d:response></d:multistatus>"
            };
        });
        var context = new FileLinkUploadContext
        {
            NormalizedBaseUrl = "https://cloud.example.test",
            UserId = "user",
            RelativeFolderPath = "NC Connector/share"
        };
        var plan = new FileLinkUploadPlan();
        plan.DirectoriesToCreate.Add("folder");

        client.CreatePlannedDirectories(
            context,
            plan,
            null,
            CancellationToken.None);

        Equal("Owned directory recovery request count", 3, requests.Count);
        Equal("Owned directory recovery uses PROPFIND", "PROPFIND", requests[2].Method);
        Equal(
            "Owned directory probe uses Depth 0",
            "0",
            requests[2].Headers["Depth"]);
    }

    private static void TestInsufficientStorage()
    {
        var client = new FileLinkDavClient(
            options => Http((HttpStatusCode)507));
        try
        {
            client.TryCreateShareRoot(
                "https://cloud.example.test",
                "user",
                "NC Connector/share",
                CancellationToken.None);
            Check("HTTP 507 has a dedicated error", false, "no exception");
        }
        catch (TalkServiceException ex)
        {
            Equal(
                "HTTP 507 has a dedicated error",
                Strings.FileLinkUploadInsufficientStorage,
                ex.Message);
        }
    }

    private static void TestShareCreateSingleRequest()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkShareClient(options =>
        {
            requests.Add(options);
            return JsonResponse(
                "{\"ocs\":{\"meta\":{\"status\":\"ok\","
                + "\"statuscode\":200,\"message\":\"OK\"},"
                + "\"data\":{\"id\":\"7\","
                + "\"url\":\"https://cloud.example.test/s/token\","
                + "\"token\":\"token\"}}}");
        });
        var request = new FileLinkRequest
        {
            Permissions =
                FileLinkPermissionFlags.Read
                | FileLinkPermissionFlags.Create,
            PasswordEnabled = true,
            Password = "secret",
            ExpireEnabled = true,
            ExpireDate = new DateTime(2026, 7, 31),
            NoteEnabled = true,
            Note = "hello world"
        };

        FileLinkShareData result = client.Create(
            "https://cloud.example.test",
            "NC Connector/share",
            "Share label",
            request,
            CancellationToken.None);
        Equal("Share create returns ID", "7", result.Id);
        Equal("Share create uses one request", 1, requests.Count);
        Equal("Share create uses POST", "POST", requests[0].Method);
        Check(
            "Share create includes permissions",
            requests[0].Payload.Contains("permissions=5"),
            requests[0].Payload);
        Check(
            "Share create omits legacy publicUpload override",
            !requests[0].Payload.Contains("publicUpload"),
            requests[0].Payload);
        Check(
            "Share create includes password",
            requests[0].Payload.Contains("password=secret"),
            requests[0].Payload);
        Check(
            "Share create includes expiry",
            requests[0].Payload.Contains("expireDate=2026-07-31"),
            requests[0].Payload);
        Check(
            "Share create includes label",
            requests[0].Payload.Contains("label=Share%20label"),
            requests[0].Payload);
        Check(
            "Share create includes note",
            requests[0].Payload.Contains("note=hello%20world"),
            requests[0].Payload);
    }

    private static void TestShareCreateRecovery()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkShareClient(options =>
        {
            requests.Add(options);
            if (options.Method == "POST")
            {
                return new NcHttpResponse
                {
                    HasHttpResponse = false,
                    TransportException = new WebException(
                        "timeout",
                        WebExceptionStatus.Timeout)
                };
            }
            return JsonResponse(
                "{\"ocs\":{\"meta\":{\"status\":\"ok\","
                + "\"statuscode\":200,\"message\":\"OK\"},"
                + "\"data\":[{\"id\":\"8\",\"share_type\":3,"
                + "\"path\":\"/NC Connector/share\","
                + "\"url\":\"https://cloud.example.test/s/recovered\","
                + "\"token\":\"recovered\"}]}}");
        });
        var request = new FileLinkRequest
        {
            Permissions = FileLinkPermissionFlags.Read
        };

        FileLinkShareData result = client.Create(
            "https://cloud.example.test",
            "NC Connector/share",
            "Share",
            request,
            CancellationToken.None);
        Equal("Unknown share POST is recovered", "8", result.Id);
        Equal("Share recovery sends POST and GET", 2, requests.Count);
        Equal("Share recovery starts with POST", "POST", requests[0].Method);
        Equal("Share recovery verifies with GET", "GET", requests[1].Method);
        Check(
            "Share recovery queries the exact path",
            requests[1].Url.Contains(
                "path=%2FNC%20Connector%2Fshare&reshares=false"),
            requests[1].Url);
        Check(
            "Share recovery excludes child shares",
            requests[1].Url.Contains("subfiles=false"),
            requests[1].Url);

        FileLinkShareData cached = client.Create(
            "https://cloud.example.test",
            "NC Connector/share",
            "Share",
            request,
            CancellationToken.None);
        Equal("Recovered share is cached", "8", cached.Id);
        Equal("Recovered share is not created twice", 2, requests.Count);
    }

    private static void TestShareCreateTransientRecovery()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkShareClient(options =>
        {
            requests.Add(options);
            if (options.Method == "POST")
            {
                return Http(HttpStatusCode.GatewayTimeout);
            }
            return JsonResponse(
                "{\"ocs\":{\"meta\":{\"status\":\"ok\","
                + "\"statuscode\":200,\"message\":\"OK\"},"
                + "\"data\":[{\"id\":\"11\",\"share_type\":3,"
                + "\"path\":\"/NC Connector/share\","
                + "\"url\":\"https://cloud.example.test/s/transient\","
                + "\"token\":\"transient\"}]}}");
        });

        FileLinkShareData result = client.Create(
            "https://cloud.example.test",
            "NC Connector/share",
            "Share",
            new FileLinkRequest
            {
                Permissions = FileLinkPermissionFlags.Read
            },
            CancellationToken.None);

        Equal(
            "Transient share response is recovered",
            "11",
            result.Id);
        Equal(
            "Transient share recovery sends POST and GET",
            2,
            requests.Count);
        Equal(
            "Transient share recovery verifies with GET",
            "GET",
            requests[1].Method);
    }

    private static void TestShareCreateAbsentRecovery()
    {
        var requests = new List<NcHttpRequestOptions>();
        int postCount = 0;
        var client = new FileLinkShareClient(options =>
        {
            requests.Add(options);
            if (options.Method == "POST")
            {
                postCount++;
                if (postCount == 1)
                {
                    return new NcHttpResponse
                    {
                        HasHttpResponse = false,
                        TransportException = new WebException(
                            "timeout",
                            WebExceptionStatus.Timeout)
                    };
                }
                return JsonResponse(
                    "{\"ocs\":{\"meta\":{\"status\":\"ok\","
                    + "\"statuscode\":200,\"message\":\"OK\"},"
                    + "\"data\":{\"id\":\"9\","
                    + "\"url\":\"https://cloud.example.test/s/retry\","
                    + "\"token\":\"retry\"}}}");
            }
            return JsonResponse(
                "{\"ocs\":{\"meta\":{\"status\":\"ok\","
                + "\"statuscode\":200,\"message\":\"OK\"},"
                + "\"data\":[]}}");
        });
        var request = new FileLinkRequest
        {
            Permissions = FileLinkPermissionFlags.Read
        };

        try
        {
            client.Create(
                "https://cloud.example.test",
                "NC Connector/share",
                "Share",
                request,
                CancellationToken.None);
            Check(
                "Confirmed absent share keeps the original failure",
                false,
                "no exception");
        }
        catch (TalkServiceException)
        {
            Check(
                "Confirmed absent share keeps the original failure",
                true);
        }

        FileLinkShareData result = client.Create(
            "https://cloud.example.test",
            "NC Connector/share",
            "Share",
            request,
            CancellationToken.None);
        Equal("Confirmed absent share permits a retry", "9", result.Id);
        Equal("Confirmed absent retry sends a second POST", 2, postCount);
        Equal("Confirmed absent retry request count", 3, requests.Count);
    }

    private static void TestShareCreateMalformedRecovery()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkShareClient(options =>
        {
            requests.Add(options);
            if (options.Method == "POST")
            {
                return new NcHttpResponse
                {
                    HasHttpResponse = true,
                    StatusCode = HttpStatusCode.OK,
                    ResponseText = "incomplete"
                };
            }
            return JsonResponse(
                "{\"ocs\":{\"meta\":{\"status\":\"ok\","
                + "\"statuscode\":200,\"message\":\"OK\"},"
                + "\"data\":[{\"id\":\"10\",\"share_type\":3,"
                + "\"path\":\"/NC Connector/share\","
                + "\"url\":\"https://cloud.example.test/s/malformed\","
                + "\"token\":\"malformed\"}]}}");
        });

        FileLinkShareData result = client.Create(
            "https://cloud.example.test",
            "NC Connector/share",
            "Share",
            new FileLinkRequest
            {
                Permissions = FileLinkPermissionFlags.Read
            },
            CancellationToken.None);

        Equal("Malformed share response is recovered", "10", result.Id);
        Equal("Malformed share response uses an exact lookup", 2, requests.Count);
        Equal("Malformed share recovery uses GET", "GET", requests[1].Method);
    }

    private static void TestShareCreateUnknownRecovery()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkShareClient(options =>
        {
            requests.Add(options);
            return new NcHttpResponse
            {
                HasHttpResponse = false,
                TransportException = new WebException(
                    "timeout",
                    WebExceptionStatus.Timeout)
            };
        });
        var request = new FileLinkRequest
        {
            Permissions = FileLinkPermissionFlags.Read
        };

        for (int attempt = 0; attempt < 2; attempt++)
        {
            try
            {
                client.Create(
                    "https://cloud.example.test",
                    "NC Connector/share",
                    "Share",
                    request,
                    CancellationToken.None);
            }
            catch (TalkServiceException)
            {
            }
        }

        Equal(
            "Unknown recovery does not repeat the POST",
            3,
            requests.Count);
        Equal(
            "Unknown recovery rechecks before any retry",
            "GET",
            requests[2].Method);
    }

    private static void TestShareOcsFailure()
    {
        var client = new FileLinkShareClient(options =>
            JsonResponse(
                "{\"ocs\":{\"meta\":{\"status\":\"failure\","
                + "\"statuscode\":997,\"message\":\"Denied\"},"
                + "\"data\":{}}}"));
        try
        {
            client.Create(
                "https://cloud.example.test",
                "NC Connector/share",
                "Share",
                new FileLinkRequest
                {
                    Permissions = FileLinkPermissionFlags.Read
                },
                CancellationToken.None);
            Check(
                "HTTP 200 OCS failure is rejected",
                false,
                "no exception");
        }
        catch (TalkServiceException ex)
        {
            Equal(
                "HTTP 200 OCS failure is rejected",
                "Denied",
                ex.Message);
        }
    }

    private static NcHttpResponse Http(HttpStatusCode statusCode)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = statusCode,
            Headers = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase)
        };
    }

    private static NcHttpResponse JsonResponse(string json)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = HttpStatusCode.OK,
            ResponseText = json,
            ParsedJson = NcJson.DeserializeObject(json),
            Headers = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase)
        };
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $csc = Join-Path `
        $env:WINDIR `
        "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkDavClient.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkDavClient.Probes.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkDavClient.Requests.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkShareClient.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkShareClient.Recovery.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkPath.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkUploadPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs")
    )
    $exe = Join-Path $TempRoot "FileLinkProtocolTests.exe"
    & $csc `
        /nologo `
        /target:exe `
        "/out:$exe" `
        /reference:System.dll `
        /reference:System.Core.dll `
        /reference:System.Web.Extensions.dll `
        @sources
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

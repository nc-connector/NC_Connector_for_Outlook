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
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static bool IsEnabled { get { return false; } }
        internal static void Log(string category, string message) { }
        internal static void LogApi(string message) { }
        internal static void LogException(string category, string message, Exception ex) { }
    }

    internal static class LogCategories
    {
        internal const string Api = "api";
        internal const string Core = "core";
    }

    internal static class Strings
    {
        internal static string FileLinkUploadSourceChanged { get { return "source changed"; } }
        internal static string FileLinkUploadLinkedItemUnsupported { get { return "linked item unsupported"; } }
        internal static string FileLinkWizardUploadCancelledMessage { get { return "upload cancelled"; } }
        internal static string FileLinkWizardUploadFailed { get { return "upload failed"; } }
        internal static string ErrorCredentialsNotVerified { get { return "credentials not verified"; } }
        internal static string ErrorMissingCredentials { get { return "missing credentials"; } }
        internal static string ErrorServerUnavailable { get { return "server unreachable"; } }
        internal static string ErrorConnectionFailed { get { return "connection failed: {0}"; } }
        internal static string TalkVersionUnknown { get { return "unknown"; } }
        internal static string NextcloudMinimumVersionRequiredFormat { get { return "Nextcloud 32 required: {0}"; } }
    }
}

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class NcHttpRequestOptions
    {
        internal string Method, Url, Accept;
        internal int TimeoutMs;
        internal bool IncludeAuthHeader, IncludeOcsApiHeader, ParseJson, ForceFreshConnection, VerifyRejectedCredentials;
    }

    internal sealed class NcHttpResponse
    {
        internal bool HasHttpResponse;
        internal HttpStatusCode StatusCode;
        internal string ResponseText;
        internal IDictionary<string, object> ParsedJson;
        internal WebException TransportException;
        internal long RequestSequence = 1;
    }

    internal sealed class NcHttpClient
    {
        internal static readonly Queue<NcHttpResponse> Responses = new Queue<NcHttpResponse>();
        internal static readonly List<NcHttpRequestOptions> Requests = new List<NcHttpRequestOptions>();
        internal static int ConfirmedVerifications;

        internal NcHttpClient(TalkServiceConfiguration configuration) { }

        internal static void ConfirmVerifiedAuthentication(TalkServiceConfiguration configuration, long requestSequence)
        {
            ConfirmedVerifications++;
        }

        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            Requests.Add(options);
            if (Responses.Count == 0)
            {
                throw new InvalidOperationException("Unexpected connection-verification request.");
            }
            return Responses.Dequeue();
        }

        internal static void Reset(params NcHttpResponse[] responses)
        {
            Responses.Clear();
            Requests.Clear();
            ConfirmedVerifications = 0;
            foreach (NcHttpResponse response in responses) { Responses.Enqueue(response); }
        }
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
        TestCapabilitiesOcsStatus();
        TestConnectionVerification();
        TestComposeLifecycleOriginCompatibility();
        TestComposeShareCleanupTracker();
        TestFileLinkUploadPolicy();
        TestFileLinkPath();
        TestPickerNavigation();
        TestFileLinkSelectionIdentity();
        TestFileLinkQueueSnapshotBuilder();
        TestFileLinkSelectionScanner();
        TestFileLinkUploadPlanner();
        TestPlainTextUtilities();
        TestBasicAuth();
        TestNcJson();
        TestBackendPolicyStatus();
        TestHtmlToPlainText();
        TestEmailSignatureSlotPlacement();
        TestEmailSignatureQuoteSeparator();
        TestSecretsCrypto();
        TestOutlookUiSynchronizationContext();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " unit test(s) failed.");
            return 1;
        }
        Console.WriteLine("All Outlook utility unit tests passed.");
        return 0;
    }

    private static void TestPickerNavigation()
    {
        var navigation = new NcTalkOutlookAddIn.UI.NextcloudPickerNavigation();
        int targetIndex;
        string targetPath;
        Check("Empty picker history cannot go back", !navigation.CanGoBack);
        Check("Empty picker history cannot go forward", !navigation.CanGoForward);
        Check("Empty picker history has no target", !navigation.TryGetTarget(-1, out targetIndex, out targetPath));
        navigation.Record("");
        navigation.Record("/folder");
        navigation.Record("folder");
        Check("Picker history accepts a back target", navigation.TryGetTarget(-1, out targetIndex, out targetPath));
        Equal("Duplicate folder navigation creates no extra step", "", targetPath);
        Check("Failed navigation leaves history position unchanged", navigation.CanGoBack && !navigation.CanGoForward);
        navigation.CompleteHistoryNavigation(targetIndex, targetPath);
        Check("Returning to root enables forward navigation", !navigation.CanGoBack && navigation.CanGoForward);
        Check("Forward navigation retains the folder", navigation.TryGetTarget(1, out targetIndex, out targetPath));
        Equal("Forward target is normalized", "folder", targetPath);
        navigation.CompleteHistoryNavigation(targetIndex, "server-folder");
        navigation.TryGetTarget(-1, out targetIndex, out targetPath);
        navigation.CompleteHistoryNavigation(targetIndex, targetPath);
        navigation.TryGetTarget(1, out targetIndex, out targetPath);
        Equal("History records the server-returned path", "server-folder", targetPath);
        navigation.Record("replacement");
        Check("New navigation discards the old forward branch", !navigation.CanGoForward);
        Check("History refuses an out-of-range target", !navigation.TryGetTarget(2, out targetIndex, out targetPath));
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
        Equal("SizeFormatting bytes", "512 B", SizeFormatting.FormatBytes(512, CultureInfo.InvariantCulture));
        Equal("SizeFormatting scales KiB", "1.5 KB", SizeFormatting.FormatBytes(1536, CultureInfo.InvariantCulture));
        Equal("SizeFormatting bytes per second", "1.5 KB/s", SizeFormatting.FormatBytesPerSecond(1536, CultureInfo.InvariantCulture));
    }

    private static void TestVersionParsing()
    {
        Version version;
        Check("NextcloudVersionHelper parses version with edition", NextcloudVersionHelper.TryParse("31.0.4 Enterprise", out version));
        Equal("NextcloudVersionHelper parsed edition version", new Version(31, 0, 4), version);
        Check("NextcloudVersionHelper parses pre-release prefix", NextcloudVersionHelper.TryParse("32.0.0-beta1", out version));
        Equal("NextcloudVersionHelper parsed pre-release prefix", new Version(32, 0, 0), version);
        Check("NextcloudVersionHelper parses product-prefixed version", NextcloudVersionHelper.TryParse("Nextcloud 34.0.1", out version));
        Equal("NextcloudVersionHelper parsed product-prefixed version", new Version(34, 0, 1), version);
        Check("NextcloudVersionHelper rejects empty", !NextcloudVersionHelper.TryParse(" ", out version));

        IDictionary<string, object> capabilities = NcJson.DeserializeObject(
            "{\"ocs\":{\"data\":{\"version\":{\"major\":32,\"minor\":1,\"micro\":4}}}}");
        string versionText;
        Check(
            "NextcloudVersionHelper extracts structured OCS version",
            NextcloudVersionHelper.TryExtractFromCapabilities(capabilities, out version, out versionText));
        Equal("NextcloudVersionHelper structured OCS version", new Version(32, 1, 4), version);
        Equal("NextcloudVersionHelper structured OCS version text", "32.1.4", versionText);

        Equal("Nextcloud minimum supported major version", 32, NextcloudVersionHelper.MinimumSupportedMajorVersion);
        Check("Nextcloud 31 is rejected", !NextcloudVersionHelper.IsSupported(new Version(31, 0, 9)));
        Check("Nextcloud 32 is supported", NextcloudVersionHelper.IsSupported(new Version(32, 0, 0)));
        Check("Missing Nextcloud version is rejected", !NextcloudVersionHelper.IsSupported(null));
    }

    private static void TestCapabilitiesOcsStatus()
    {
        IDictionary<string, object> success = NcJson.DeserializeObject(
            "{\"ocs\":{\"meta\":{\"status\":\"ok\",\"statuscode\":200,\"message\":\"OK\"},\"data\":{}}}");
        string detail;
        Check(
            "Nextcloud capabilities accepts OCS statuscode 200",
            NcJson.IsOcsSuccess(success, out detail),
            detail);

        IDictionary<string, object> legacySuccess = NcJson.DeserializeObject(
            "{\"ocs\":{\"meta\":{\"status\":\"ok\",\"statuscode\":100},\"data\":{}}}");
        Check(
            "Nextcloud capabilities accepts OCS statuscode 100",
            NcJson.IsOcsSuccess(legacySuccess, out detail),
            detail);

        IDictionary<string, object> zeroSuccess = NcJson.DeserializeObject(
            "{\"ocs\":{\"meta\":{\"status\":\"ok\",\"statuscode\":0},\"data\":{}}}");
        Check(
            "Nextcloud capabilities accepts OCS statuscode 0",
            NcJson.IsOcsSuccess(zeroSuccess, out detail),
            detail);

        IDictionary<string, object> failure = NcJson.DeserializeObject(
            "{\"ocs\":{\"meta\":{\"status\":\"failure\",\"statuscode\":997,\"message\":\"Denied\"},\"data\":{}}}");
        Check(
            "Nextcloud capabilities rejects OCS failure code",
            !NcJson.IsOcsSuccess(failure, out detail));
        Equal(
            "Nextcloud capabilities exposes OCS failure detail",
            "Denied",
            detail);

        IDictionary<string, object> incomplete = NcJson.DeserializeObject(
            "{\"ocs\":{\"meta\":{},\"data\":{}}}");
        Check(
            "Nextcloud capabilities rejects empty OCS metadata",
            !NcJson.IsOcsSuccess(incomplete, out detail));
    }

    private static void TestConnectionVerification()
    {
        var configuration = ConnectionTestConfiguration();
        NcHttpClient.Reset(new NcHttpResponse
        {
            TransportException = new WebException("controlled transport failure")
        });
        TalkServiceException failure = CaptureConnectionFailure(() =>
            new NextcloudCapabilitiesService(configuration).GetRequiredSnapshot(true, true));
        Check("Capabilities without HTTP preserve a typed transport failure",
            failure != null && failure.IsTransportError && !failure.IsAuthenticationError
            && (int)failure.StatusCode == 0);
        Equal("Capabilities preserve transport diagnostic text", "controlled transport failure",
            failure != null ? failure.Message : null);

        NcHttpClient.Reset(new NcHttpResponse());
        failure = CaptureConnectionFailure(() =>
            new NextcloudCapabilitiesService(configuration).GetRequiredSnapshot(true, true));
        Check("Capabilities without HTTP or an exception remain a transport failure",
            failure != null && failure.IsTransportError);
        Equal("Missing transport detail uses the server-unavailable message", Strings.ErrorServerUnavailable,
            failure != null ? failure.Message : null);

        NcHttpClient.Reset(CapabilitiesTestResponse("32.0.0"), UserTestResponse("initial-user"));
        var verifier = new ConnectionVerificationHarness(configuration);
        string message = string.Empty;
        Check("Connection verification accepts supported capabilities and a canonical UID",
            verifier.VerifyConnection(out message));
        Equal("Successful connection verification reports the version", "32.0.0", message);
        Equal("Connection verification checks capabilities and the current UID", 2, NcHttpClient.Requests.Count);
        Check("Both connection-verification requests use fresh connections",
            NcHttpClient.Requests.Count == 2 && NcHttpClient.Requests.All(request => request.ForceFreshConnection));
        Check("Verification uses the authenticated capabilities and user endpoints",
            NcHttpClient.Requests.Count == 2
            && NcHttpClient.Requests[0].Url.EndsWith("/cloud/capabilities?format=json", StringComparison.Ordinal)
            && NcHttpClient.Requests[1].Url.EndsWith("/cloud/user?format=json", StringComparison.Ordinal)
            && NcHttpClient.Requests.All(request => request.IncludeAuthHeader && request.IncludeOcsApiHeader));
        Check("Ordinary action checks cannot retry rejected credentials",
            NcHttpClient.Requests.All(request => !request.VerifyRejectedCredentials)
            && NcHttpClient.ConfirmedVerifications == 0);

        NcHttpClient.Reset(CapabilitiesTestResponse("32.0.0"), UserTestResponse("initial-user"));
        Check("Explicit settings verification may retry rejected credentials", verifier.VerifyConnection(out message, true));
        Check("Explicit verification marks both authenticated checks and confirms only the valid UID",
            NcHttpClient.Requests.Count == 2 && NcHttpClient.Requests.All(request => request.VerifyRejectedCredentials)
            && NcHttpClient.ConfirmedVerifications == 1);

        NcHttpClient.Reset(CapabilitiesTestResponse("32.0.0"), UserTestResponse(string.Empty));
        Check("An invalid UID cannot release an authentication pause", !verifier.VerifyConnection(out message, true)
            && NcHttpClient.ConfirmedVerifications == 0);

        NcHttpClient.Reset(new NcHttpResponse());
        failure = CaptureConnectionFailure(() => verifier.VerifyConnection(out message));
        Check("A warm capability cache cannot hide an unavailable server at action verification",
            failure != null && failure.IsTransportError && NcHttpClient.Requests.Count == 1);

        NcHttpClient.Reset(CapabilitiesTestResponse("32.0.0"), new NcHttpResponse
        {
            TransportException = new WebException("controlled UID transport failure")
        });
        failure = CaptureConnectionFailure(() => verifier.VerifyConnection(out message));
        Check("VerifyConnection propagates UID transport failures despite a cached UID",
            failure != null && failure.IsTransportError && !failure.IsAuthenticationError);
        Check("UID transport failure retains its diagnostic text",
            failure != null && failure.Message.Contains("controlled UID transport failure"));
        Equal("The UID transport failure occurs after a fresh capability request", 2, NcHttpClient.Requests.Count);

        foreach (HttpStatusCode statusCode in new[] { HttpStatusCode.Unauthorized, HttpStatusCode.Forbidden })
        {
            NcHttpClient.Reset(CapabilitiesTestResponse("32.0.0"), new NcHttpResponse
            {
                HasHttpResponse = true,
                StatusCode = statusCode,
                ResponseText = "controlled rejection"
            });
            failure = CaptureConnectionFailure(() => verifier.VerifyConnection(out message));
            Check("VerifyConnection preserves typed UID authentication rejection " + (int)statusCode,
                failure != null && failure.IsAuthenticationError && !failure.IsTransportError
                && failure.StatusCode == statusCode && failure.ResponseBody == "controlled rejection");
        }

        NcHttpClient.Reset(new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = HttpStatusCode.Unauthorized
        });
        failure = CaptureConnectionFailure(() => verifier.VerifyConnection(out message));
        Check("Capability authentication rejection is not a transport outage",
            failure != null && failure.IsAuthenticationError && !failure.IsTransportError);

        foreach (string version in new[] { "31.0.9", "" })
        {
            NcHttpClient.Reset(CapabilitiesTestResponse(version));
            Check("Unsupported or absent server version returns false without a transport exception: " + version,
                !verifier.VerifyConnection(out message));
            Check("Version refusal retains the minimum-version explanation",
                message.StartsWith("Nextcloud 32 required:", StringComparison.Ordinal)
                && message != Strings.ErrorServerUnavailable);
            Equal("Version refusal never probes the UID", 1, NcHttpClient.Requests.Count);
        }

        NcHttpClient.Reset(CapabilitiesTestResponse("32.0.0"), UserTestResponse(""));
        Check("Missing canonical UID returns false instead of a transport exception",
            !verifier.VerifyConnection(out message));
        Check("Missing canonical UID retains the response-validation explanation",
            message.Contains("canonical user ID") && message != Strings.ErrorServerUnavailable);

        foreach (HttpStatusCode statusCode in new[] { (HttpStatusCode)429, HttpStatusCode.InternalServerError })
        {
            NcHttpClient.Reset(CapabilitiesTestResponse("32.0.0"), new NcHttpResponse
            {
                HasHttpResponse = true,
                StatusCode = statusCode
            });
            Check("Ordinary UID HTTP rejection returns false: " + (int)statusCode,
                !verifier.VerifyConnection(out message));
            Check("UID HTTP rejection retains status text rather than an invented transport message",
                message.Contains("HTTP " + (int)statusCode) && message != Strings.ErrorServerUnavailable);
        }

        NcHttpClient.Reset(new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = HttpStatusCode.OK
        });
        failure = CaptureConnectionFailure(() => verifier.VerifyConnection(out message));
        Check("Malformed capabilities keep their non-transport validation failure",
            failure != null && !failure.IsTransportError && !failure.IsAuthenticationError
            && failure.Message == Strings.ErrorCredentialsNotVerified);

        configuration = ConnectionTestConfiguration();
        NcHttpClient.Reset(UserTestResponse("cached-user"));
        Equal("Initial UID lookup caches the canonical identity", "cached-user",
            NextcloudUserIdentityService.ResolveCurrentUserId(configuration));
        Check("Ordinary UID lookup does not force a fresh connection",
            NcHttpClient.Requests.Count == 1 && !NcHttpClient.Requests[0].ForceFreshConnection);
        NcHttpClient.Reset(UserTestResponse("refreshed-user"));
        Equal("Forced UID lookup replaces a cached identity", "refreshed-user",
            NextcloudUserIdentityService.ResolveCurrentUserId(configuration, true));
        Check("Forced UID lookup requests a fresh connection",
            NcHttpClient.Requests.Count == 1 && NcHttpClient.Requests[0].ForceFreshConnection);
        NcHttpClient.Reset();
        Equal("Unforced UID lookup keeps the refreshed cache", "refreshed-user",
            NextcloudUserIdentityService.ResolveCurrentUserId(configuration));
        Equal("Unforced cached UID lookup performs no request", 0, NcHttpClient.Requests.Count);
    }

    private static TalkServiceConfiguration ConnectionTestConfiguration()
    {
        return new TalkServiceConfiguration(
            "https://connection-" + Guid.NewGuid().ToString("N") + ".example.test", "test-user", "test-only");
    }

    private static NcHttpResponse CapabilitiesTestResponse(string version)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = HttpStatusCode.OK,
            ParsedJson = NcJson.DeserializeObject(
                "{\"ocs\":{\"meta\":{\"status\":\"ok\",\"statuscode\":200},\"data\":{"
                + "\"version\":{\"string\":\"" + version + "\"},\"capabilities\":{}}}}")
        };
    }

    private static NcHttpResponse UserTestResponse(string userId)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = HttpStatusCode.OK,
            ParsedJson = NcJson.DeserializeObject("{\"ocs\":{\"data\":{\"id\":\"" + userId + "\"}}}")
        };
    }

    private static TalkServiceException CaptureConnectionFailure(Action action)
    {
        try { action(); }
        catch (TalkServiceException ex) { return ex; }
        return null;
    }

    private static void TestComposeShareCleanupTracker()
    {
        var tracker = new ComposeShareCleanupTracker();
        Check(
            "Compose cleanup tracker rejects null records",
            !tracker.Arm(null));
        Check(
            "Compose cleanup tracker rejects records without a folder",
            !tracker.Arm(new ComposeShareCleanupRecord()));

        ComposeShareCleanupRecord recordA =
            CreateCleanupRecord(
                "account-a",
                "shares/record-a",
                "share-a");
        Check(
            "Compose cleanup tracker arms record A",
            tracker.Arm(recordA));
        Equal(
            "Compose cleanup tracker counts record A",
            1,
            tracker.Count);
        Check(
            "Compose cleanup tracker rejects duplicate record A",
            !tracker.Arm(
                CreateCleanupRecord(
                    "ACCOUNT-A",
                    "SHARES/RECORD-A",
                    "share-a")));

        Equal(
            "Compose cleanup tracker releases record A after write",
            1,
            tracker.ReleaseAll());
        Equal(
            "Compose cleanup tracker is empty after ReleaseAll",
            0,
            tracker.Count);

        ComposeShareCleanupRecord recordB =
            CreateCleanupRecord(
                "account-a",
                "shares/record-b",
                "share-b");
        Check(
            "Compose cleanup tracker arms record B after ReleaseAll",
            tracker.Arm(recordB));
        List<ComposeShareCleanupRecord> drained = tracker.Drain();
        Equal(
            "Compose cleanup tracker drains only record B",
            1,
            drained.Count);
        Check(
            "Compose cleanup tracker preserves drained record B",
            drained.Count == 1
                && object.ReferenceEquals(recordB, drained[0]));
        Equal(
            "Compose cleanup tracker is empty after Drain",
            0,
            tracker.Count);
        Equal(
            "Compose cleanup tracker drains an empty generation once",
            0,
            tracker.Drain().Count);
    }

    private static void TestComposeLifecycleOriginCompatibility()
    {
        var serializer = new JavaScriptSerializer();
        ComposeLifecycleOrigin origin =
            serializer.Deserialize<ComposeLifecycleOrigin>(
                "{\"ServerUrl\":\"https://cloud.example.test\","
                + "\"Username\":\"alice\","
                + "\"AppPassword\":\"secret\","
                + "\"CanonicalUserId\":\"legacy-alice\","
                + "\"AccountFingerprint\":\"legacy-fingerprint\"}");
        Check(
            "Compose origin accepts legacy canonical-user payloads",
            origin != null && origin.IsComplete());
        Equal(
            "Compose origin keeps the legacy account fingerprint",
            "legacy-fingerprint",
            origin != null ? origin.AccountFingerprint : string.Empty);

        ComposeLifecycleOrigin created =
            ComposeLifecycleOrigin.Create(
                new TalkServiceConfiguration(
                    "https://cloud.example.test/",
                    "Alice",
                    "secret"));
        Equal(
            "Compose origin keeps the established account fingerprint",
            "40f4c18aff88b77a30863281fb9f68e8d0f578b93f8ef106518191e52c5aec52",
            created.AccountFingerprint);
    }

    private static ComposeShareCleanupRecord CreateCleanupRecord(
        string accountFingerprint,
        string relativeFolder,
        string shareId)
    {
        return new ComposeShareCleanupRecord
        {
            RelativeFolder = relativeFolder,
            ShareId = shareId,
            ShareLabel = relativeFolder,
            Origin = new ComposeLifecycleOrigin
            {
                AccountFingerprint = accountFingerprint
            }
        };
    }

    private static void TestFileLinkUploadPolicy()
    {
        Equal(
            "FileLink uses the NC32 server AutoMkcol header",
            "X-NC-WebDAV-Auto-Mkcol",
            FileLinkUploadPolicy.AutoMkcolHeaderName);
        Equal(
            "FileLink direct upload limit is 20 MiB",
            20L * 1024L * 1024L,
            FileLinkUploadPolicy.DirectUploadLimitBytes);
        Equal(
            "FileLink chunk minimum is 5 MiB",
            5L * 1024L * 1024L,
            FileLinkUploadPolicy.ChunkUploadMinimumChunkSizeBytes);
        Equal(
            "FileLink standard chunk size is 20 MiB",
            20L * 1024L * 1024L,
            FileLinkUploadPolicy.ChunkUploadChunkSizeBytes);
        Equal(
            "FileLink chunk maximum is 5 GiB",
            5L * 1024L * 1024L * 1024L,
            FileLinkUploadPolicy.ChunkUploadMaximumChunkSizeBytes);
        Equal(
            "FileLink chunk count maximum is 10000",
            10000,
            FileLinkUploadPolicy.ChunkUploadMaxChunks);
        Equal(
            "FileLink maximum file size follows chunk limits",
            FileLinkUploadPolicy.ChunkUploadMaximumChunkSizeBytes
                * FileLinkUploadPolicy.ChunkUploadMaxChunks,
            FileLinkUploadPolicy.ChunkUploadMaximumFileSizeBytes);
        Equal(
            "FileLink bulk candidate limit is 8 MiB",
            8L * 1024L * 1024L,
            FileLinkUploadPolicy.BulkCandidateLimitBytes);
        Equal(
            "FileLink bulk batch byte limit is 20 MiB",
            20L * 1024L * 1024L,
            FileLinkUploadPolicy.BulkBatchLimitBytes);
        Equal("FileLink bulk batch file limit", 100, FileLinkUploadPolicy.BulkBatchFileLimit);
        Equal("FileLink bulk minimum file count", 20, FileLinkUploadPolicy.BulkMinimumFileCount);
        Equal("FileLink maximum parallel requests", 3, FileLinkUploadPolicy.MaxParallelRequests);
        Equal("FileLink maximum request attempts", 3, FileLinkUploadPolicy.MaxRequestAttempts);

        Check(
            "FileLink direct upload includes exact 20 MiB boundary",
            !FileLinkUploadPolicy.ShouldUseChunkedUpload(FileLinkUploadPolicy.DirectUploadLimitBytes));
        Check(
            "FileLink chunked upload starts above 20 MiB",
            FileLinkUploadPolicy.ShouldUseChunkedUpload(FileLinkUploadPolicy.DirectUploadLimitBytes + 1));
        Equal(
            "FileLink direct upload uses one transfer request",
            1,
            FileLinkUploadPolicy.GetTransferRequestCount(
                FileLinkUploadPolicy.DirectUploadLimitBytes));
        Equal(
            "FileLink chunked upload counts folder chunks and move",
            4,
            FileLinkUploadPolicy.GetTransferRequestCount(
                FileLinkUploadPolicy.DirectUploadLimitBytes + 1));
        Equal(
            "FileLink chunk count stays within server limit",
            FileLinkUploadPolicy.ChunkUploadMaxChunks,
            (int)(((FileLinkUploadPolicy.ChunkUploadMaximumFileSizeBytes - 1)
                / FileLinkUploadPolicy.GetChunkUploadChunkSize(
                    FileLinkUploadPolicy.ChunkUploadMaximumFileSizeBytes))
                + 1));
        Equal(
            "FileLink maximum file uses the maximum chunk size",
            FileLinkUploadPolicy.ChunkUploadMaximumChunkSizeBytes,
            FileLinkUploadPolicy.GetChunkUploadChunkSize(
                FileLinkUploadPolicy.ChunkUploadMaximumFileSizeBytes));
        Check(
            "FileLink accepts the exact maximum file size",
            FileLinkUploadPolicy.IsSupportedFileSize(
                FileLinkUploadPolicy.ChunkUploadMaximumFileSizeBytes));
        Check(
            "FileLink rejects a file above the maximum size",
            !FileLinkUploadPolicy.IsSupportedFileSize(
                FileLinkUploadPolicy.ChunkUploadMaximumFileSizeBytes + 1));
        bool oversizedFileRejected = false;
        try
        {
            FileLinkUploadPolicy.GetTransferRequestCount(
                FileLinkUploadPolicy.ChunkUploadMaximumFileSizeBytes + 1);
        }
        catch (ArgumentOutOfRangeException)
        {
            oversizedFileRejected = true;
        }
        Check(
            "FileLink rejects oversized files before request planning",
            oversizedFileRejected);
        Check(
            "FileLink bulk candidate includes exact 8 MiB boundary",
            FileLinkUploadPolicy.IsBulkCandidate(FileLinkUploadPolicy.BulkCandidateLimitBytes));
        Check(
            "FileLink bulk candidate rejects files above 8 MiB",
            !FileLinkUploadPolicy.IsBulkCandidate(FileLinkUploadPolicy.BulkCandidateLimitBytes + 1));
        Check("FileLink bulk candidate rejects negative size", !FileLinkUploadPolicy.IsBulkCandidate(-1));

        Check(
            "FileLink bulk upload requires server capability",
            !FileLinkUploadPolicy.ShouldUseBulkUpload(
                false,
                20,
                20,
                0,
                0,
                1));
        Check(
            "FileLink bulk upload requires minimum file count",
            !FileLinkUploadPolicy.ShouldUseBulkUpload(
                true,
                19,
                19,
                0,
                0,
                1));
        Check(
            "FileLink bulk upload requires a batch request",
            !FileLinkUploadPolicy.ShouldUseBulkUpload(
                true,
                20,
                20,
                0,
                0,
                0));
        Check(
            "FileLink bulk upload accepts exact request-saving threshold",
            FileLinkUploadPolicy.ShouldUseBulkUpload(
                true,
                20,
                20,
                75,
                75,
                1));
        Check(
            "FileLink bulk upload rejects insufficient request savings",
            !FileLinkUploadPolicy.ShouldUseBulkUpload(
                true,
                20,
                20,
                76,
                76,
                1));
        Check(
            "FileLink bulk upload includes non-bulk requests in saving threshold",
            !FileLinkUploadPolicy.ShouldUseBulkUpload(
                true,
                20,
                100,
                0,
                0,
                1));
        Check(
            "FileLink bulk upload counts extra bulk parent requests",
            !FileLinkUploadPolicy.ShouldUseBulkUpload(
                true,
                20,
                20,
                0,
                20,
                1));

        int[] retryable = { 408, 423, 429, 502, 503, 504 };
        foreach (int statusCode in retryable)
        {
            Check(
                "FileLink retryable HTTP " + statusCode,
                FileLinkUploadPolicy.IsRetryableStatusCode(statusCode));
        }
        Check(
            "FileLink HTTP 502 result can be indeterminate",
            FileLinkUploadPolicy.IsIndeterminateStatusCode(502));
        Check(
            "FileLink HTTP 429 result is not indeterminate",
            !FileLinkUploadPolicy.IsIndeterminateStatusCode(429));

        int[] nonRetryable = { 400, 401, 404, 409, 500, 507 };
        foreach (int statusCode in nonRetryable)
        {
            Check(
                "FileLink non-retryable HTTP " + statusCode,
                !FileLinkUploadPolicy.IsRetryableStatusCode(statusCode));
        }
    }

    private static void TestFileLinkPath()
    {
        Equal(
            "FileLink path normalizes separators",
            "one/two/three",
            FileLinkPath.NormalizeRelativePath(
                "one\\two//three"));
        Equal(
            "FileLink path removes empty sanitized segments",
            "one/two",
            FileLinkPath.NormalizeRelativePath(
                "one/   /two"));
        Equal(
            "FileLink path neutralizes current-directory segments",
            "one/_/two",
            FileLinkPath.NormalizeRelativePath(
                "one/./two"));
        Equal(
            "FileLink path neutralizes parent-directory segments",
            "one/__/two",
            FileLinkPath.NormalizeRelativePath(
                "one/../two"));
        Equal(
            "FileLink share folder date is stable",
            "20260723_share",
            FileLinkPath.BuildShareFolderName(
                new DateTime(2026, 7, 23),
                "share"));
        FileLinkShareTarget target = FileLinkPath.ResolveShareTarget(
            "NC Connector\\Team",
            "  quarterly:review  ",
            new DateTime(2026, 7, 23, 23, 59, 59),
            "share");
        Equal(
            "FileLink share target normalizes the configured base path",
            "NC Connector/Team",
            target.BasePath);
        Equal(
            "FileLink share target sanitizes the entered name",
            "quarterly_review",
            target.ShareName);
        Equal(
            "FileLink share target keeps the wizard date",
            "NC Connector/Team/20260723_quarterly_review",
            target.RelativeFolderPath);
        FileLinkShareTarget fallbackTarget =
            FileLinkPath.ResolveShareTarget(
                "NC Connector",
                "   ",
                new DateTime(2026, 7, 24),
                "share");
        Equal(
            "FileLink share target applies the localized fallback",
            "NC Connector/20260724_share",
            fallbackTarget.RelativeFolderPath);
    }

    private static void TestFileLinkUploadPlanner()
    {
        string fixtureRoot = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-plan-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(fixtureRoot);
        try
        {
            string bulkRoot = Path.Combine(fixtureRoot, "bulk");
            Directory.CreateDirectory(bulkRoot);
            Directory.CreateDirectory(Path.Combine(bulkRoot, "empty"));
            for (int index = 0; index < 20; index++)
            {
                File.WriteAllText(
                    Path.Combine(bulkRoot, "file-" + index.ToString("00", CultureInfo.InvariantCulture) + ".txt"),
                    "x");
            }

            var bulkSelection = new FileLinkSelection(
                FileLinkSelectionType.Directory,
                bulkRoot);
            var bulkPlan = BuildFileLinkUploadPlan(
                new List<FileLinkSelection> { bulkSelection },
                true,
                1,
                null,
                CancellationToken.None);
            Equal("FileLink planner scans 20 files", 20, bulkPlan.Files.Count);
            Equal("FileLink planner classifies bulk-capable files", 20, bulkPlan.BulkFiles.Count);
            Equal("FileLink planner creates one bulk batch", 1, bulkPlan.BulkBatches.Count);
            Equal(
                "FileLink planner leaves no direct files with bulk capability",
                0,
                bulkPlan.DirectFileCount);
            Check(
                "FileLink planner preserves an empty directory",
                bulkPlan.DirectoriesToCreate.Any(path => path.EndsWith("/empty", StringComparison.OrdinalIgnoreCase)));

            var directPlan = BuildFileLinkUploadPlan(
                new List<FileLinkSelection> { bulkSelection },
                false,
                1,
                null,
                CancellationToken.None);
            Equal(
                "FileLink planner keeps files direct without bulk capability",
                20,
                directPlan.DirectFileCount);
            Equal("FileLink planner creates no bulk files without capability", 0, directPlan.BulkFiles.Count);
            Equal("FileLink planner creates no bulk batches without capability", 0, directPlan.BulkBatches.Count);

            string directOnlyRoot = Path.Combine(
                fixtureRoot,
                "direct-only");
            Directory.CreateDirectory(directOnlyRoot);
            File.WriteAllText(
                Path.Combine(directOnlyRoot, "file.txt"),
                "direct");
            var directOnlyPlan = BuildFileLinkUploadPlan(
                new List<FileLinkSelection>
                {
                    new FileLinkSelection(
                        FileLinkSelectionType.Directory,
                        directOnlyRoot)
                },
                false,
                1,
                null,
                CancellationToken.None);
            Equal(
                "FileLink planner lets AutoMkcol create direct-only parents",
                0,
                directOnlyPlan.DirectoriesToCreate.Count);

            string sharedDirectRoot = Path.Combine(
                fixtureRoot,
                "shared-direct");
            string sharedDirectLeft = Path.Combine(
                sharedDirectRoot,
                "left");
            string sharedDirectRight = Path.Combine(
                sharedDirectRoot,
                "right");
            Directory.CreateDirectory(sharedDirectLeft);
            Directory.CreateDirectory(sharedDirectRight);
            File.WriteAllText(
                Path.Combine(sharedDirectLeft, "left.txt"),
                "left");
            File.WriteAllText(
                Path.Combine(sharedDirectRight, "right.txt"),
                "right");
            var sharedDirectPlan = BuildFileLinkUploadPlan(
                new List<FileLinkSelection>
                {
                    new FileLinkSelection(
                        FileLinkSelectionType.Directory,
                        sharedDirectRoot)
                },
                false,
                1,
                null,
                CancellationToken.None);
            Equal(
                "FileLink planner creates one shared direct parent",
                1,
                sharedDirectPlan.DirectoriesToCreate.Count);
            Check(
                "FileLink planner leaves single-file child paths to AutoMkcol",
                sharedDirectPlan.DirectoriesToCreate[0].EndsWith(
                    "shared-direct",
                    StringComparison.OrdinalIgnoreCase));

            var nextcloudRoot = new NextcloudStorageEntry(
                "Team: Archive",
                "Team: Archive",
                true,
                0,
                null);
            var nextcloudSelection =
                FileLinkSelection.FromNextcloudFolder(
                    nextcloudRoot,
                    new[]
                    {
                        new NextcloudStorageEntry(
                            "Team: Archive/Empty",
                            "Empty",
                            true,
                            0,
                            null),
                        new NextcloudStorageEntry(
                            "Team: Archive/report.pdf",
                            "report.pdf",
                            false,
                            42,
                            new DateTime(
                                2026,
                                9,
                                11,
                                0,
                                0,
                                0,
                                DateTimeKind.Utc))
                    });
            var nextcloudPlan = BuildFileLinkUploadPlan(
                new List<FileLinkSelection> { nextcloudSelection },
                true,
                1,
                null,
                CancellationToken.None);
            Equal(
                "FileLink planner keeps remote files on the server",
                FileLinkUploadTransport.ServerCopy,
                nextcloudPlan.Files[0].Transport);
            Equal(
                "FileLink planner preserves the Nextcloud source path",
                "Team: Archive/report.pdf",
                nextcloudPlan.Files[0].NextcloudSourcePath);
            Equal(
                "FileLink planner tracks remote file bytes",
                42L,
                nextcloudPlan.TotalBytes);
            Check(
                "FileLink planner keeps an empty Nextcloud folder",
                nextcloudPlan.DirectoriesToCreate.Any(
                    path => path.EndsWith(
                        "/Empty",
                        StringComparison.OrdinalIgnoreCase)));

            string stableRoot = Path.Combine(
                fixtureRoot,
                "stable-queue");
            Directory.CreateDirectory(stableRoot);
            string queuedFile = Path.Combine(
                stableRoot,
                "queued.txt");
            File.WriteAllText(queuedFile, "queued");
            var stableSelection = new FileLinkSelection(
                FileLinkSelectionType.Directory,
                stableRoot);
            var stableSelections = new List<FileLinkSelection>
            {
                stableSelection
            };
            var stableSnapshots =
                new Dictionary<FileLinkSelection, FileLinkQueueNode>
                {
                    {
                        stableSelection,
                        FileLinkQueueSnapshotBuilder.Build(
                            stableSelection,
                            CancellationToken.None)
                    }
                };
            File.WriteAllText(
                Path.Combine(stableRoot, "late.txt"),
                "late");
            FileLinkUploadPlan stablePlan =
                FileLinkUploadPlanBuilder.Build(
                    stableSelections,
                    stableSnapshots,
                    false,
                    1,
                    null,
                    CancellationToken.None);
            Equal(
                "FileLink planner uploads only files captured in the queue",
                1,
                stablePlan.Files.Count);
            Check(
                "FileLink planner keeps the queued local file",
                stablePlan.Files[0].LocalPath.EndsWith(
                    "queued.txt",
                    StringComparison.OrdinalIgnoreCase));

            File.Delete(queuedFile);
            bool removedQueuedFileRejected = false;
            try
            {
                FileLinkUploadPlanBuilder.Build(
                    stableSelections,
                    stableSnapshots,
                    false,
                    1,
                    null,
                    CancellationToken.None);
            }
            catch (TalkServiceException ex)
            {
                removedQueuedFileRejected = string.Equals(
                    ex.Message,
                    "source changed",
                    StringComparison.Ordinal);
            }
            Check(
                "FileLink planner rejects a file removed from the queue snapshot",
                removedQueuedFileRejected);

        }
        finally
        {
            if (Directory.Exists(fixtureRoot))
            {
                Directory.Delete(fixtureRoot, true);
            }
        }
    }

    private static FileLinkUploadPlan BuildFileLinkUploadPlan(
        IList<FileLinkSelection> selections,
        bool bulkUploadSupported,
        int fixedRequestCount,
        Func<FileLinkDuplicateInfo, string> duplicateResolver,
        CancellationToken cancellationToken)
    {
        var snapshots =
            new Dictionary<FileLinkSelection, FileLinkQueueNode>();
        foreach (FileLinkSelection selection in selections)
        {
            snapshots.Add(
                selection,
                FileLinkQueueSnapshotBuilder.Build(
                    selection,
                    cancellationToken));
        }
        return FileLinkUploadPlanBuilder.Build(
            selections,
            snapshots,
            bulkUploadSupported,
            fixedRequestCount,
            duplicateResolver,
            cancellationToken);
    }

    private static void TestFileLinkQueueSnapshotBuilder()
    {
        string fixtureRoot = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-queue-tests-" + Guid.NewGuid().ToString("N"));
        string junctionPath = string.Empty;
        Directory.CreateDirectory(fixtureRoot);
        try
        {
            string selectedFolder = Path.Combine(fixtureRoot, "Selected");
            Directory.CreateDirectory(selectedFolder);
            Directory.CreateDirectory(Path.Combine(selectedFolder, "Empty"));
            Directory.CreateDirectory(Path.Combine(selectedFolder, "Nested"));
            File.WriteAllText(
                Path.Combine(selectedFolder, "root.txt"),
                "root");
            File.WriteAllText(
                Path.Combine(selectedFolder, "Nested", "child.pdf"),
                "child");

            FileLinkQueueNode local = FileLinkQueueSnapshotBuilder.Build(
                new FileLinkSelection(
                    FileLinkSelectionType.Directory,
                    selectedFolder),
                CancellationToken.None);
            Equal("Queue snapshot keeps the selected local folder", "Selected", local.DisplayName);
            Check("Queue snapshot keeps local folders before files", local.Children.Count == 3 && local.Children[0].IsDirectory && local.Children[1].IsDirectory && !local.Children[2].IsDirectory);
            FileLinkQueueNode nested = local.Children.First(node => node.DisplayName == "Nested");
            Check("Queue snapshot includes nested local files", nested.Children.Count == 1 && nested.Children[0].DisplayName == "child.pdf" && nested.Children[0].Length == 5);
            Check("Queue snapshot captures local modification time", nested.Children[0].LastWriteTimeUtc.HasValue);
            FileLinkQueueNode empty = local.Children.First(node => node.DisplayName == "Empty");
            Check("Queue snapshot preserves empty local folders", empty.Children.Count == 0);

            string reparseRoot = Path.Combine(
                fixtureRoot,
                "reparse-selection");
            string reparseTarget = Path.Combine(
                fixtureRoot,
                "reparse-target");
            junctionPath = Path.Combine(
                reparseRoot,
                "linked");
            Directory.CreateDirectory(reparseRoot);
            Directory.CreateDirectory(reparseTarget);
            File.WriteAllText(
                Path.Combine(reparseTarget, "outside.txt"),
                "outside");

            bool junctionCreated = TryCreateDirectoryJunction(
                junctionPath,
                reparseTarget);
            Check(
                "Queue snapshot test creates a directory junction",
                junctionCreated);
            if (junctionCreated)
            {
                bool linkedItemRejected = false;
                try
                {
                    FileLinkQueueSnapshotBuilder.Build(
                        new FileLinkSelection(
                            FileLinkSelectionType.Directory,
                            reparseRoot),
                        CancellationToken.None);
                }
                catch (IOException ex)
                {
                    linkedItemRejected = string.Equals(
                        ex.Message,
                        "linked item unsupported",
                        StringComparison.Ordinal);
                }
                Check(
                    "Queue snapshot rejects a nested directory junction",
                    linkedItemRejected);
            }

            var remoteRoot = new NextcloudStorageEntry(
                "Projects",
                "Projects",
                true,
                0,
                null);
            FileLinkSelection remoteSelection = FileLinkSelection.FromNextcloudFolder(
                remoteRoot,
                new[]
                {
                    new NextcloudStorageEntry(
                        "Projects/Empty",
                        "Empty",
                        true,
                        0,
                        null),
                    new NextcloudStorageEntry(
                        "Projects/report.docx",
                        "report.docx",
                        false,
                        42,
                        null)
                });
            FileLinkQueueNode remote = FileLinkQueueSnapshotBuilder.Build(
                remoteSelection,
                CancellationToken.None);
            Equal("Queue snapshot keeps the selected Nextcloud folder", "Projects", remote.DisplayName);
            Check("Queue snapshot preserves empty Nextcloud folders", remote.Children.Count == 2 && remote.Children[0].IsDirectory && remote.Children[0].Children.Count == 0);
            Check("Queue snapshot exposes the Nextcloud file size", !remote.Children[1].IsDirectory && remote.Children[1].Length == 42);

            bool missingParentRejected = false;
            try
            {
                FileLinkSelection incomplete = FileLinkSelection.FromNextcloudFolder(
                    remoteRoot,
                    new[]
                    {
                        new NextcloudStorageEntry(
                            "Projects/Missing/file.txt",
                            "file.txt",
                            false,
                            1,
                            null)
                    });
                FileLinkQueueSnapshotBuilder.Build(
                    incomplete,
                    CancellationToken.None);
            }
            catch (IOException)
            {
                missingParentRejected = true;
            }
            Check("Queue snapshot rejects an incomplete Nextcloud hierarchy", missingParentRejected);
        }
        finally
        {
            if (!string.IsNullOrEmpty(junctionPath)
                && Directory.Exists(junctionPath))
            {
                Directory.Delete(junctionPath);
            }
            if (Directory.Exists(fixtureRoot))
            {
                Directory.Delete(fixtureRoot, true);
            }
        }
    }

    private static void TestFileLinkSelectionIdentity()
    {
        var localFile = new FileLinkSelection(
            FileLinkSelectionType.File,
            @"C:\Reports\report.pdf");
        var localFileCaseVariant = new FileLinkSelection(
            FileLinkSelectionType.File,
            @"c:\reports\Report.pdf");
        var localSelections = new HashSet<FileLinkSelection>(
            new[] { localFile },
            FileLinkSelection.IdentityComparer);
        Check("Queue identity deduplicates Windows file path case variants", !localSelections.Add(localFileCaseVariant));
        Equal("Per-selection state retains reference identity", 2, new HashSet<FileLinkSelection>(new[] { localFile, localFileCaseVariant }).Count);

        var localFolder = new FileLinkSelection(
            FileLinkSelectionType.Directory,
            @"C:\Reports");
        var localFolderCaseVariant = new FileLinkSelection(
            FileLinkSelectionType.Directory,
            @"c:\reports");
        Check("Queue identity accepts a different Windows selection", localSelections.Add(localFolder));
        Check("Queue identity deduplicates Windows folder path case variants", !localSelections.Add(localFolderCaseVariant));

        var lowerFile = FileLinkSelection.FromNextcloudFile(
            new NextcloudStorageEntry("Reports/report.pdf", "report.pdf", false, 1, null));
        var upperFile = FileLinkSelection.FromNextcloudFile(
            new NextcloudStorageEntry("Reports/Report.pdf", "Report.pdf", false, 2, null));
        var sameFile = FileLinkSelection.FromNextcloudFile(
            new NextcloudStorageEntry("/Reports/report.pdf", "report.pdf", false, 1, null));
        var remoteSelections = new HashSet<FileLinkSelection>(
            new[] { lowerFile },
            FileLinkSelection.IdentityComparer);
        Check("Queue identity retains distinct Nextcloud file path case variants", remoteSelections.Add(upperFile));
        Check("Queue identity deduplicates the exact normalized Nextcloud path", !remoteSelections.Add(sameFile));
        var laterBatch = new HashSet<FileLinkSelection>(
            remoteSelections,
            FileLinkSelection.IdentityComparer);
        Check("Queue identity deduplicates a Nextcloud file from an earlier batch", !laterBatch.Add(upperFile));
        Check("Queue identity separates local paths from identical Nextcloud paths", laterBatch.Add(new FileLinkSelection(FileLinkSelectionType.File, lowerFile.NextcloudPath)));

        var lowerFolder = FileLinkSelection.FromNextcloudFolder(
            new NextcloudStorageEntry("reports", "reports", true, 0, null),
            null);
        var upperFolder = FileLinkSelection.FromNextcloudFolder(
            new NextcloudStorageEntry("Reports", "Reports", true, 0, null),
            null);
        Check("Queue identity accepts a Nextcloud folder", remoteSelections.Add(lowerFolder));
        Check("Queue identity retains distinct Nextcloud folder path case variants", remoteSelections.Add(upperFolder));
        Check("Queue identity deduplicates an exact Nextcloud folder", !remoteSelections.Add(FileLinkSelection.FromNextcloudFolder(new NextcloudStorageEntry("reports", "reports", true, 0, null), null)));

        var selections = new List<FileLinkSelection> { lowerFile, upperFile };
        var snapshots = selections.ToDictionary(
            selection => selection,
            selection => FileLinkQueueSnapshotBuilder.Build(selection, CancellationToken.None));
        int resolverCalls = 0;
        FileLinkSelectionScanResult scan = FileLinkSelectionScanner.Scan(
            selections,
            snapshots,
            info =>
            {
                resolverCalls++;
                return "report-copy.pdf";
            },
            CancellationToken.None);
        Equal("Nextcloud case variants both reach the upload plan", 2, scan.Files.Count);
        Equal("Destination case collisions still invoke the existing rename resolver", 1, resolverCalls);
        Equal("The first Nextcloud source keeps its exact path", lowerFile.NextcloudPath, scan.Files[0].NextcloudSourcePath);
        Equal("The second Nextcloud source keeps its exact path", upperFile.NextcloudPath, scan.Files[1].NextcloudSourcePath);
        Equal("The destination rename remains separate from source identity", "report-copy.pdf", scan.Files[1].RemotePath);
        Equal("Nextcloud case variants retain their combined size", 3L, scan.TotalBytes);
        Equal("The first Nextcloud selection retains its own count", 1, scan.SelectionFileCounts[lowerFile]);
        Equal("The second Nextcloud selection retains its own count", 1, scan.SelectionFileCounts[upperFile]);
    }

    private static void TestFileLinkSelectionScanner()
    {
        string fixtureRoot = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-scan-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(fixtureRoot);
        try
        {
            string selectionRoot = Path.Combine(
                fixtureRoot,
                "selection");
            string nestedRoot = Path.Combine(
                selectionRoot,
                "nested");
            Directory.CreateDirectory(nestedRoot);
            Directory.CreateDirectory(
                Path.Combine(selectionRoot, "empty"));
            File.WriteAllText(
                Path.Combine(nestedRoot, "a.txt"),
                "one");
            File.WriteAllText(
                Path.Combine(nestedRoot, " a.txt"),
                "two");

            int resolverCalls = 0;
            var selection = new FileLinkSelection(
                FileLinkSelectionType.Directory,
                selectionRoot);
            var queueSnapshots =
                new Dictionary<FileLinkSelection, FileLinkQueueNode>
                {
                    {
                        selection,
                        FileLinkQueueSnapshotBuilder.Build(
                            selection,
                            CancellationToken.None)
                    }
                };
            FileLinkSelectionScanResult scan =
                FileLinkSelectionScanner.Scan(
                    new List<FileLinkSelection> { selection },
                    queueSnapshots,
                    info =>
                    {
                        resolverCalls++;
                        return "renamed-"
                            + resolverCalls.ToString(
                                CultureInfo.InvariantCulture)
                            + ".txt";
                    },
                    CancellationToken.None);

            Equal(
                "FileLink scanner invokes resolver for a sanitized collision",
                1,
                resolverCalls);
            Equal(
                "FileLink scanner keeps collision paths unique",
                2,
                scan.Files
                    .Select(file => file.RemotePath)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count());
            Check(
                "FileLink scanner applies the collision rename",
                scan.Files.Any(
                    file => file.RemotePath.EndsWith(
                        "/renamed-1.txt",
                        StringComparison.OrdinalIgnoreCase)));
            Check(
                "FileLink scanner records an empty directory",
                scan.Directories.Any(
                    path => path.EndsWith(
                        "/empty",
                        StringComparison.OrdinalIgnoreCase)));
            Equal(
                "FileLink scanner tracks selection file count",
                2,
                scan.SelectionFileCounts[selection]);
            Equal(
                "FileLink scanner tracks total bytes",
                6L,
                scan.TotalBytes);

            bool missingSnapshotMapRejected = false;
            try
            {
                FileLinkSelectionScanner.Scan(
                    new List<FileLinkSelection> { selection },
                    null,
                    null,
                    CancellationToken.None);
            }
            catch (ArgumentNullException ex)
            {
                missingSnapshotMapRejected = string.Equals(
                    ex.ParamName,
                    "queueSnapshots",
                    StringComparison.Ordinal);
            }
            Check(
                "FileLink scanner rejects a missing snapshot map",
                missingSnapshotMapRejected);

            bool missingSnapshotRejected = false;
            try
            {
                FileLinkSelectionScanner.Scan(
                    new List<FileLinkSelection> { selection },
                    new Dictionary<FileLinkSelection, FileLinkQueueNode>(),
                    null,
                    CancellationToken.None);
            }
            catch (TalkServiceException ex)
            {
                missingSnapshotRejected = string.Equals(
                    ex.Message,
                    "source changed",
                    StringComparison.Ordinal);
            }
            Check(
                "FileLink scanner rejects a missing selection snapshot",
                missingSnapshotRejected);
        }
        finally
        {
            if (Directory.Exists(fixtureRoot))
            {
                Directory.Delete(fixtureRoot, true);
            }
        }
    }

    private static bool TryCreateDirectoryJunction(
        string junctionPath,
        string targetPath)
    {
        string commandProcessor =
            Environment.GetEnvironmentVariable("ComSpec");
        if (string.IsNullOrWhiteSpace(commandProcessor))
        {
            commandProcessor = "cmd.exe";
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = commandProcessor,
            Arguments = "/d /c mklink /J \""
                + junctionPath
                + "\" \""
                + targetPath
                + "\"",
            CreateNoWindow = true,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false
        };
        using (Process process = Process.Start(startInfo))
        {
            if (process == null)
            {
                return false;
            }
            process.StandardOutput.ReadToEnd();
            process.StandardError.ReadToEnd();
            process.WaitForExit();
            return process.ExitCode == 0
                && Directory.Exists(junctionPath)
                && (new DirectoryInfo(junctionPath).Attributes
                    & FileAttributes.ReparsePoint)
                    == FileAttributes.ReparsePoint;
        }
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

    private static void TestEmailSignatureSlotPlacement()
    {
        Equal(
            "Reply signature below _MailOriginal moves above quote",
            EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 140, 180, 100, 102, false));
        Equal(
            "Forward signature below border quote moves above quote",
            EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 240, 280, 200, 200, false));
        Equal(
            "Direct-match native table ending at _MailOriginal stays in place",
            EmailSignatureSlotPlacementDecision.KeepExistingSlot,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 160, 200, 198, 200, false));
        Equal(
            "Managed signature ending at protected insertion point stays in place",
            EmailSignatureSlotPlacementDecision.KeepExistingSlot,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 160, 198, 198, 200, false));
        Equal(
            "Signature crossing actual quote boundary fails closed",
            EmailSignatureSlotPlacementDecision.UnsafeQuoteBoundaryOverlap,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 160, 201, 198, 200, false));
        Equal(
            "Signature starting at actual quote boundary fails closed",
            EmailSignatureSlotPlacementDecision.UnsafeQuoteBoundaryOverlap,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 200, 240, 198, 200, false));
        Equal(
            "Signature entirely below actual quote boundary moves above quote",
            EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 201, 240, 198, 200, false));
        Equal(
            "Inline border target also acts as quote boundary",
            EmailSignatureSlotPlacementDecision.KeepExistingSlot,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 160, 200, 200, 200, false));
        Equal(
            "Authored text after signature moves replacement to safe point",
            EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 40, 80, 160, 162, true));
        Equal(
            "Whitespace after signature keeps existing slot",
            EmailSignatureSlotPlacementDecision.KeepExistingSlot,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 40, 80, 160, 162, false));
        Equal(
            "New-mail authored text still moves replacement to document end",
            EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint,
            EmailSignatureSlotPlacementPolicy.Resolve(false, 40, 80, 160, 160, true));
        Equal(
            "New mail does not apply reply quote correction",
            EmailSignatureSlotPlacementDecision.KeepExistingSlot,
            EmailSignatureSlotPlacementPolicy.Resolve(false, 140, 180, 100, 100, false));
    }

    private static void TestEmailSignatureQuoteSeparator()
    {
        var shareRow = new QuoteTestParagraph(5, 15, true, true);
        var nestedPermissionCell = new QuoteTestParagraph(15, 25, true, true);
        var signature = new QuoteTestParagraph(30, 50, false, true);
        var quote = new QuoteTestParagraph(60, 80, false, true);
        int separator;
        Check(
            "Reply quote scan accepts the divider after a bordered share and signature",
            SignatureQuoteSeparatorHarness.TryFind(
                new QuoteTestDocument(shareRow, nestedPermissionCell, signature, quote),
                true, 30, 50, out separator));
        Equal("Share permission borders cannot become the quote boundary", 60, separator);
        Equal("Outer share table borders are not read", 0, shareRow.BorderReads);
        Equal("Nested permission cell borders are not read", 0, nestedPermissionCell.BorderReads);
        Equal("Existing signature remains excluded from quote-border detection", 0, signature.BorderReads);
        Equal(
            "Send-time placement retains the signature below the share",
            EmailSignatureSlotPlacementDecision.KeepExistingSlot,
            EmailSignatureSlotPlacementPolicy.Resolve(true, 30, 50, separator, separator, false));

        Check(
            "Table-only reply has no safe quote boundary",
            !SignatureQuoteSeparatorHarness.TryFind(
                new QuoteTestDocument(shareRow, nestedPermissionCell),
                false, 0, 0, out separator));

        var paragraphs = new List<QuoteTestParagraph>();
        for (int index = 0; index < 100; index++)
        {
            paragraphs.Add(new QuoteTestParagraph(index * 10, index * 10 + 10, true, true));
        }
        paragraphs.Add(new QuoteTestParagraph(1000, 1020, false, true));
        Check(
            "Multiple shares do not exhaust the quote-divider paragraph budget",
            SignatureQuoteSeparatorHarness.TryFind(
                new QuoteTestDocument(paragraphs.ToArray()), false, 0, 0, out separator));
        Equal("Divider after more than eighty table paragraphs is retained", 1000, separator);

        var unavailable = new QuoteTestParagraph(5, 15, false, true);
        unavailable.Range.ThrowOnInformation = true;
        Check(
            "Failed Word table membership stops quote detection without guessing",
            !SignatureQuoteSeparatorHarness.TryFind(
                new QuoteTestDocument(unavailable, quote), false, 0, 0, out separator));
        Equal("Failed table membership cannot reach border detection", 0, unavailable.BorderReads);
        unavailable.Range.ThrowOnInformation = false;
        unavailable.Range.WithinTable = null;
        Check(
            "Unknown Word table membership also fails closed",
            !SignatureQuoteSeparatorHarness.TryFind(
                new QuoteTestDocument(unavailable, quote), false, 0, 0, out separator));

        paragraphs.Clear();
        for (int index = 0; index < 80; index++)
        {
            paragraphs.Add(new QuoteTestParagraph(index * 10, index * 10 + 10, false, false));
        }
        paragraphs.Add(new QuoteTestParagraph(800, 820, false, true));
        Check(
            "Quote scan still stops after eighty eligible non-table paragraphs",
            !SignatureQuoteSeparatorHarness.TryFind(
                new QuoteTestDocument(paragraphs.ToArray()), false, 0, 0, out separator));
    }

    public sealed class QuoteTestDocument
    {
        public QuoteTestRange Content { get; private set; }

        public QuoteTestDocument(params QuoteTestParagraph[] paragraphs)
        {
            Content = new QuoteTestRange(0, paragraphs.Max(paragraph => paragraph.Range.End), false);
            Content.Paragraphs = new QuoteTestParagraphs(paragraphs);
        }

        public QuoteTestRange Range(int start, int end) { return Content; }
    }

    public sealed class QuoteTestRange
    {
        public int Start { get; private set; }
        public int End { get; private set; }
        public object WithinTable { get; set; }
        public bool ThrowOnInformation { get; set; }
        public QuoteTestParagraphs Paragraphs { get; set; }

        public QuoteTestRange(int start, int end, bool withinTable)
        {
            Start = start;
            End = end;
            WithinTable = withinTable;
        }

        [System.Runtime.CompilerServices.IndexerName("Information")]
        public object this[int information]
        {
            get
            {
                if (ThrowOnInformation || information != 12)
                {
                    throw new InvalidOperationException("Word table membership unavailable.");
                }
                return WithinTable;
            }
        }
    }

    public sealed class QuoteTestParagraphs
    {
        private readonly QuoteTestParagraph[] paragraphs;
        public int Count { get { return paragraphs.Length; } }
        public QuoteTestParagraphs(QuoteTestParagraph[] values) { paragraphs = values; }
        public QuoteTestParagraph Item(int index) { return paragraphs[index - 1]; }
    }

    public sealed class QuoteTestParagraph
    {
        private readonly QuoteTestBorders borders;
        public QuoteTestRange Range { get; private set; }
        public int BorderReads { get; private set; }
        public QuoteTestBorders Borders
        {
            get
            {
                BorderReads++;
                return borders;
            }
        }

        public QuoteTestParagraph(int start, int end, bool withinTable, bool visibleBorder)
        {
            Range = new QuoteTestRange(start, end, withinTable);
            borders = new QuoteTestBorders(visibleBorder);
        }
    }

    public sealed class QuoteTestBorders
    {
        private readonly bool visible;
        public QuoteTestBorders(bool value) { visible = value; }
        public QuoteTestBorder Item(int index)
        {
            return new QuoteTestBorder { LineStyle = visible && index == -1 ? 1 : 0 };
        }
    }

    public sealed class QuoteTestBorder
    {
        public int LineStyle { get; set; }
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

    private static void TestOutlookUiSynchronizationContext()
    {
        OutlookUiSynchronizationContext context = null;
        Exception uiThreadException = null;
        int uiThreadId = 0;
        var ready = new ManualResetEventSlim(false);
        var uiThread = new Thread(() =>
        {
            try
            {
                context = new OutlookUiSynchronizationContext();
                uiThreadId = Thread.CurrentThread.ManagedThreadId;
                ready.Set();
                Application.Run();
            }
            catch (Exception ex)
            {
                uiThreadException = ex;
                ready.Set();
            }
            finally
            {
                if (context != null)
                {
                    context.Dispose();
                }
            }
        });
        uiThread.IsBackground = true;
        uiThread.SetApartmentState(ApartmentState.STA);
        uiThread.Start();

        Check("Outlook UI context becomes ready", ready.Wait(TimeSpan.FromSeconds(5)));
        Check("Outlook UI context initializes on STA", context != null && uiThreadException == null, uiThreadException != null ? uiThreadException.ToString() : "");
        if (context == null)
        {
            return;
        }

        Task<Tuple<int, ApartmentState>> dispatch = context.InvokeAsync(() =>
        {
            var result = Tuple.Create(Thread.CurrentThread.ManagedThreadId, Thread.CurrentThread.GetApartmentState());
            Application.ExitThread();
            return result;
        });

        bool dispatchCompleted = false;
        try
        {
            dispatchCompleted = dispatch.Wait(TimeSpan.FromSeconds(5));
        }
        catch (AggregateException ex)
        {
            Check("Outlook UI dispatch completes", false, ex.ToString());
        }
        if (!dispatch.IsFaulted)
        {
            Check("Outlook UI dispatch completes", dispatchCompleted);
        }
        if (dispatch.Status == TaskStatus.RanToCompletion)
        {
            Equal("Outlook UI dispatch returns to captured thread", uiThreadId, dispatch.Result.Item1);
            Equal("Outlook UI dispatch runs in STA", ApartmentState.STA, dispatch.Result.Item2);
        }
        Check("Outlook UI test thread exits", uiThread.Join(TimeSpan.FromSeconds(5)));
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $signatureInteropSource = Get-Content -Raw -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Controllers\ManagedEmailSignatureController.cs")
    $signatureQuoteParts = @()
    foreach ($methodName in @('TryFindInlineQuoteSeparatorStart', 'ParagraphHasVisibleBorder', 'BorderAtIndexIsVisible')) {
        $method = [regex]::Match($signatureInteropSource, "(?ms)^        private static bool $methodName\(.*?^        }")
        if (-not $method.Success) {
            throw "Could not isolate production signature quote method '$methodName'."
        }
        $signatureQuoteParts += $method.Value
    }
    $tableInformationConstant = [regex]::Match($signatureInteropSource, '(?m)^        private const int WordInformationWithinTable = \d+;')
    if (-not $tableInformationConstant.Success) {
        throw "Could not isolate the Word table information constant."
    }
    $signatureQuoteSource = Join-Path $TempRoot "SignatureQuoteSeparatorHarness.cs"
    @"
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using NcTalkOutlookAddIn.Utilities;

internal static class SignatureQuoteSeparatorHarness
{
$($tableInformationConstant.Value)
    internal static bool TryFind(object document, bool hasExcludedRange, int excludedStart, int excludedEnd, out int position)
    {
        return TryFindInlineQuoteSeparatorStart(document, 0, hasExcludedRange, excludedStart, excludedEnd, out position);
    }
$($signatureQuoteParts -join "`r`n")
}
"@ | Set-Content -LiteralPath $signatureQuoteSource -Encoding UTF8

    $talkServiceSource = Get-Content -Raw -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkService.cs")
    $verificationParts = @()
    foreach ($signature in @('internal bool VerifyConnection', 'private void EnsureConfiguration')) {
        $method = [regex]::Match($talkServiceSource, "(?ms)^        $signature\(.*?^        }")
        if (-not $method.Success) {
            throw "Could not isolate production connection-verification method '$signature'."
        }
        $verificationParts += $method.Value
    }
    $connectionVerificationSource = Join-Path $TempRoot "ConnectionVerificationHarness.cs"
    @"
using System;
using System.Net;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;

internal sealed class ConnectionVerificationHarness
{
    private readonly TalkServiceConfiguration _configuration;
    internal ConnectionVerificationHarness(TalkServiceConfiguration configuration) { _configuration = configuration; }
$($verificationParts -join "`r`n")
}
"@ | Set-Content -LiteralPath $connectionVerificationSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        $signatureQuoteSource,
        $connectionVerificationSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\ComInteropScope.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkSelection.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkQueueNode.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\NextcloudStorageEntry.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\NextcloudCapabilitiesSnapshot.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\ComposeLifecycleOrigin.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\ComposeShareCleanupRecord.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkDuplicateInfo.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkUploadPlan.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkQueueSnapshotBuilder.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkSelectionScanner.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkUploadPlanner.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceConfiguration.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceException.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\NextcloudCapabilitiesService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\NextcloudUserIdentityService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Controllers\ComposeShareCleanupTracker.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\PasswordGenerator.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\SizeFormatting.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudVersionHelper.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudPath.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\NextcloudPickerNavigation.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudUriValidator.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkPath.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkUploadPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\PlainTextUtilities.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HttpAuthUtilities.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\BackendPolicyStatus.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\EmailSignatureSlotPlacementPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HtmlToPlainTextConverter.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\SecretsCrypto.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\OutlookUiSynchronizationContext.cs")
    )

    $exe = Join-Path $TempRoot "OutlookUtilityTests.exe"
    $references = @(
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Windows.Forms.dll",
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

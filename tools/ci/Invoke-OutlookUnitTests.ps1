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
using NcTalkOutlookAddIn.Models;
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
        internal const string Core = "core";
    }

    internal static class Strings
    {
        internal static string FileLinkUploadSourceChanged { get { return "source changed"; } }
        internal static string FileLinkUploadLinkedItemUnsupported { get { return "linked item unsupported"; } }
        internal static string FileLinkWizardUploadCancelledMessage { get { return "upload cancelled"; } }
        internal static string FileLinkWizardUploadFailed { get { return "upload failed"; } }
        internal static string ErrorCredentialsNotVerified { get { return "credentials not verified"; } }
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
        TestFileLinkUploadPolicy();
        TestFileLinkPath();
        TestFileLinkSelectionScanner();
        TestFileLinkUploadPlanner();
        TestPlainTextUtilities();
        TestBasicAuth();
        TestNcJson();
        TestBackendPolicyStatus();
        TestHtmlToPlainText();
        TestEmailSignatureSlotPlacement();
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
            var bulkPlan = FileLinkUploadPlanBuilder.Build(
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

            var directPlan = FileLinkUploadPlanBuilder.Build(
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
            var directOnlyPlan = FileLinkUploadPlanBuilder.Build(
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
            var sharedDirectPlan = FileLinkUploadPlanBuilder.Build(
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

        }
        finally
        {
            if (Directory.Exists(fixtureRoot))
            {
                Directory.Delete(fixtureRoot, true);
            }
        }
    }

    private static void TestFileLinkSelectionScanner()
    {
        string fixtureRoot = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-scan-tests-" + Guid.NewGuid().ToString("N"));
        string junctionPath = string.Empty;
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
            FileLinkSelectionScanResult scan =
                FileLinkSelectionScanner.Scan(
                    new List<FileLinkSelection> { selection },
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
                "FileLink scanner test creates a directory junction",
                junctionCreated);
            if (junctionCreated)
            {
                bool linkedItemRejected = false;
                try
                {
                    FileLinkSelectionScanner.Scan(
                        new List<FileLinkSelection>
                        {
                            new FileLinkSelection(
                                FileLinkSelectionType.Directory,
                                reparseRoot)
                        },
                        null,
                        CancellationToken.None);
                }
                catch (TalkServiceException ex)
                {
                    linkedItemRejected = string.Equals(
                        ex.Message,
                        "linked item unsupported",
                        StringComparison.Ordinal);
                }
                Check(
                    "FileLink scanner rejects a nested directory junction",
                    linkedItemRejected);
            }
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

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\FileLinkSelection.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkDuplicateInfo.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkUploadPlan.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkSelectionScanner.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkUploadPlanner.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\TalkServiceException.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\PasswordGenerator.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\SizeFormatting.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudVersionHelper.cs"),
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

Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path `
    ([System.IO.Path]::GetTempPath()) `
    ("nc4ol-filelink-transfer-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "FileLinkTransferTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class FileLinkUploadPhaseProgress
    {
    }
}

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static void Log(string category, string message)
        {
        }

        internal static void LogException(
            string category,
            string message,
            Exception ex)
        {
        }
    }

    internal static class LogCategories
    {
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

        internal static string FileLinkUploadSourceChanged
        {
            get { return "source changed"; }
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

    internal static class TimeUtilities
    {
        internal static long? ToUnixTimeSeconds(DateTime value)
        {
            DateTime utc = value.Kind == DateTimeKind.Utc
                ? value
                : value.ToUniversalTime();
            return (long)(utc - new DateTime(
                1970,
                1,
                1,
                0,
                0,
                0,
                DateTimeKind.Utc)).TotalSeconds;
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
            StatusCode = statusCode;
        }

        internal HttpStatusCode StatusCode { get; private set; }
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

    internal sealed class FileLinkPlannedFile
    {
        internal FileLinkPlannedFile(
            string localPath,
            string remotePath,
            long length,
            DateTime lastWriteTimeUtc)
        {
            LocalPath = localPath;
            RemotePath = remotePath;
            Length = length;
            LastWriteTimeUtc = lastWriteTimeUtc;
        }

        internal string LocalPath { get; private set; }
        internal string RemotePath { get; private set; }
        internal long Length { get; private set; }
        internal DateTime LastWriteTimeUtc { get; private set; }
        internal string BulkChecksum { get; set; }
    }

    internal sealed class FileLinkBulkUploadBatch
    {
        internal FileLinkBulkUploadBatch(
            IEnumerable<FileLinkPlannedFile> files)
        {
            Files = files.ToList();
        }

        internal IList<FileLinkPlannedFile> Files { get; private set; }
    }

    internal sealed class FileLinkUploadPlan
    {
        internal FileLinkUploadPlan()
        {
            Files = new List<FileLinkPlannedFile>();
            BulkFiles = new List<FileLinkPlannedFile>();
            DirectoriesToCreate = new List<string>();
        }

        internal IList<FileLinkPlannedFile> Files { get; private set; }
        internal IList<FileLinkPlannedFile> BulkFiles { get; private set; }
        internal IList<string> DirectoriesToCreate { get; private set; }
        internal long TotalBytes { get; set; }
    }

    internal sealed class FileLinkUploadContext
    {
        internal string NormalizedBaseUrl { get; set; }
        internal string UserId { get; set; }
        internal string RelativeFolderPath { get; set; }
    }

    internal sealed class FileLinkUploadProgressCoordinator
    {
        internal Action<FileLinkPlannedFile, long, bool> Progress;

        internal void SetFileBytes(
            FileLinkPlannedFile file,
            long uploadedBytes,
            bool force)
        {
            if (Progress != null)
            {
                Progress(file, uploadedBytes, force);
            }
        }
    }

    internal sealed class FileLinkFolderProgressReporter
    {
        internal FileLinkFolderProgressReporter(
            IProgress<NcTalkOutlookAddIn.Models.FileLinkUploadPhaseProgress>
                progress,
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

internal static class FileLinkTransferTests
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
        string tempRoot = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-transfer-fixtures-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        try
        {
            TestDirectRequest(tempRoot);
            TestBulkRetryRewritesSameBody(tempRoot);
            TestBulkPerPathFailure(tempRoot);
            TestChunkRequestSequence(tempRoot);
            TestChunkFailureUsesShortCleanup(tempRoot);
            TestMoveRecoveryWithMatchingLength();
            TestMoveRecoveryRejectsWrongLength();
            TestMoveRecoveryRejectsCollection();
            TestMoveHttpFailureIsNotRetried();
        }
        finally
        {
            Directory.Delete(tempRoot, true);
        }

        if (failures > 0)
        {
            Console.Error.WriteLine(
                failures + " FileLink transfer test(s) failed.");
            return 1;
        }

        Console.WriteLine("All FileLink transfer tests passed.");
        return 0;
    }

    private static void TestDirectRequest(string tempRoot)
    {
        byte[] content = { 10, 20, 30, 40, 50 };
        FileLinkPlannedFile file = CreateFile(
            tempRoot,
            "direct.bin",
            "folder/direct file.bin",
            content);
        var context = new FileLinkUploadContext
        {
            NormalizedBaseUrl = "https://cloud.example.test",
            UserId = "user id",
            RelativeFolderPath = "root"
        };
        NcHttpRequestOptions request = null;
        byte[] sentBody = null;
        var client = new FileLinkDavClient(options =>
        {
            request = options;
            using (var output = new MemoryStream())
            {
                options.BodyWriter(output);
                sentBody = output.ToArray();
            }
            return Http(HttpStatusCode.Created);
        });

        new FileLinkDirectUploader(client).Upload(
            context,
            file,
            new FileLinkUploadProgressCoordinator(),
            CancellationToken.None);

        Equal("Direct upload uses PUT", "PUT", request.Method);
        Equal(
            "Direct upload uses the exact encoded target URL",
            "https://cloud.example.test/remote.php/dav/files/user%20id/root/folder/direct%20file.bin",
            request.Url);
        Equal(
            "Direct upload sends the source length",
            (long)content.Length,
            request.ContentLength);
        Equal(
            "Direct upload sends an octet-stream",
            "application/octet-stream",
            request.ContentType);
        Equal(
            "Direct upload enables server AutoMkcol",
            "1",
            request.Headers[
                FileLinkUploadPolicy.AutoMkcolHeaderName]);
        Check(
            "Direct upload writes the exact source body",
            content.SequenceEqual(sentBody));
    }

    private static void TestBulkRetryRewritesSameBody(string tempRoot)
    {
        FileLinkPlannedFile first = CreateFile(
            tempRoot,
            "first.txt",
            "first.txt",
            new byte[] { 1, 2, 3, 4 });
        FileLinkPlannedFile second = CreateFile(
            tempRoot,
            "second.txt",
            "folder/second.txt",
            new byte[] { 5, 6, 7 });
        first.BulkChecksum = FileLinkSourceFile.ComputeMd5Hex(first);
        second.BulkChecksum = FileLinkSourceFile.ComputeMd5Hex(second);
        var batch = new FileLinkBulkUploadBatch(
            new[] { first, second });
        var context = new FileLinkUploadContext
        {
            NormalizedBaseUrl = "https://cloud.example.test",
            UserId = "user",
            RelativeFolderPath = "root"
        };
        var bodies = new List<byte[]>();
        var requests = new List<NcHttpRequestOptions>();
        var events = new List<string>();
        int attempt = 0;
        var client = new FileLinkDavClient(options =>
        {
            attempt++;
            requests.Add(options);
            events.Add("attempt:" + attempt);
            using (var output = new MemoryStream())
            {
                options.BodyWriter(output);
                bodies.Add(output.ToArray());
            }

            if (attempt == 1)
            {
                return new NcHttpResponse
                {
                    HasHttpResponse = false,
                    TransportException = new WebException(
                        "timeout",
                        WebExceptionStatus.Timeout),
                    Headers = RetryNow()
                };
            }
            if (attempt == 2)
            {
                return Http(HttpStatusCode.ServiceUnavailable);
            }
            return BulkSuccess(first, second);
        });
        var coordinator = new FileLinkUploadProgressCoordinator
        {
            Progress = (file, uploaded, force) =>
                events.Add(
                    (force ? "reset:" : "progress:")
                    + file.RemotePath
                    + ":"
                    + uploaded.ToString(CultureInfo.InvariantCulture))
        };

        new FileLinkBulkUploader(client).Upload(
            context,
            batch,
            coordinator,
            CancellationToken.None);

        Equal("Bulk retries transport and transient HTTP", 3, attempt);
        Equal("Bulk writes a body for every attempt", 3, bodies.Count);
        Check(
            "Bulk retry body is identical after transport failure",
            bodies[0].SequenceEqual(bodies[1]));
        Check(
            "Bulk retry body is identical after HTTP failure",
            bodies[1].SequenceEqual(bodies[2]));
        Check(
            "Bulk uses multipart related content",
            requests.All(
                request => request.ContentType.StartsWith(
                    "multipart/related; boundary=ncconnector-",
                    StringComparison.Ordinal)));
        Check(
            "Bulk sends the exact multipart content length",
            requests
                .Select((request, index) =>
                    request.ContentLength == bodies[index].LongLength)
                .All(matches => matches));
        Check(
            "Bulk keeps content type and length across retries",
            requests.All(
                request =>
                    request.ContentType == requests[0].ContentType
                    && request.ContentLength
                    == requests[0].ContentLength));
        Equal(
            "Bulk resets every file before both retries",
            4,
            events.Count(value => value.StartsWith(
                "reset:",
                StringComparison.Ordinal)));
        int secondAttempt = events.IndexOf("attempt:2");
        int thirdAttempt = events.IndexOf("attempt:3");
        Check(
            "First batch reset happens before attempt two",
            events.Take(secondAttempt).Count(value => value.StartsWith(
                "reset:",
                StringComparison.Ordinal)) == 2);
        Check(
            "Second batch reset happens before attempt three",
            events.Take(thirdAttempt).Count(value => value.StartsWith(
                "reset:",
                StringComparison.Ordinal)) == 4);
    }

    private static void TestBulkPerPathFailure(string tempRoot)
    {
        FileLinkPlannedFile file = CreateFile(
            tempRoot,
            "bulk-error.txt",
            "bulk-error.txt",
            new byte[] { 1 });
        file.BulkChecksum = FileLinkSourceFile.ComputeMd5Hex(file);
        var batch = new FileLinkBulkUploadBatch(new[] { file });
        int requestCount = 0;
        var client = new FileLinkDavClient(options =>
        {
            requestCount++;
            using (var output = new MemoryStream())
            {
                options.BodyWriter(output);
                Equal(
                    "Bulk error request content length matches body",
                    options.ContentLength,
                    output.Length);
            }
            return new NcHttpResponse
            {
                HasHttpResponse = true,
                StatusCode = HttpStatusCode.Created,
                Headers = RetryNow(),
                ParsedJson = new Dictionary<string, object>
                {
                    {
                        "/root/" + file.RemotePath,
                        new Dictionary<string, object>
                        {
                            { "error", true },
                            { "message", "denied" }
                        }
                    }
                }
            };
        });

        try
        {
            new FileLinkBulkUploader(client).Upload(
                new FileLinkUploadContext
                {
                    NormalizedBaseUrl =
                        "https://cloud.example.test",
                    UserId = "user",
                    RelativeFolderPath = "root"
                },
                batch,
                new FileLinkUploadProgressCoordinator(),
                CancellationToken.None);
            Check(
                "Bulk rejects a per-path error",
                false,
                "no exception");
        }
        catch (TalkServiceException)
        {
            Check("Bulk rejects a per-path error", true);
        }
        Equal(
            "Bulk per-path error is not retried",
            1,
            requestCount);
    }

    private static void TestChunkRequestSequence(string tempRoot)
    {
        long fileSize =
            FileLinkUploadPolicy.ChunkUploadChunkSizeBytes + 3;
        FileLinkPlannedFile file = CreateSizedFile(
            tempRoot,
            "chunked.bin",
            "folder/chunked.bin",
            fileSize);
        var context = new FileLinkUploadContext
        {
            NormalizedBaseUrl = "https://cloud.example.test",
            UserId = "user",
            RelativeFolderPath = "root"
        };
        var requests = new List<NcHttpRequestOptions>();
        var bodies = new List<long>();
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            if (options.Method == "PUT")
            {
                using (var output = new MemoryStream())
                {
                    options.BodyWriter(output);
                    bodies.Add(output.Length);
                }
            }
            return Http(HttpStatusCode.Created);
        });

        new FileLinkChunkUploader(client).Upload(
            context,
            file,
            new FileLinkUploadProgressCoordinator(),
            CancellationToken.None);

        Equal("Chunk upload sends four requests", 4, requests.Count);
        Equal("Chunk upload starts with MKCOL", "MKCOL", requests[0].Method);
        Equal("Chunk upload sends first PUT", "PUT", requests[1].Method);
        Equal("Chunk upload sends second PUT", "PUT", requests[2].Method);
        Equal("Chunk upload finishes with MOVE", "MOVE", requests[3].Method);

        string targetUrl =
            "https://cloud.example.test/remote.php/dav/files/user/root/folder/chunked.bin";
        string uploadFolderUrl = requests[0].Url;
        Check(
            "Chunk upload folder uses the user upload namespace",
            uploadFolderUrl.StartsWith(
                "https://cloud.example.test/remote.php/dav/uploads/user/ncconnector-",
                StringComparison.Ordinal));
        Equal(
            "Chunk MKCOL carries the final destination",
            targetUrl,
            requests[0].Headers["Destination"]);
        Equal(
            "First chunk has a numbered DAV path",
            uploadFolderUrl + "/00001",
            requests[1].Url);
        Equal(
            "Second chunk has a numbered DAV path",
            uploadFolderUrl + "/00002",
            requests[2].Url);
        Equal(
            "First chunk sends the standard chunk size",
            FileLinkUploadPolicy.ChunkUploadChunkSizeBytes,
            requests[1].ContentLength);
        Equal(
            "Second chunk sends the remaining bytes",
            3L,
            requests[2].ContentLength);
        Equal(
            "First chunk body matches its request length",
            requests[1].ContentLength,
            bodies[0]);
        Equal(
            "Second chunk body matches its request length",
            requests[2].ContentLength,
            bodies[1]);
        for (int index = 1; index <= 2; index++)
        {
            Equal(
                "Chunk PUT destination " + index,
                targetUrl,
                requests[index].Headers["Destination"]);
            Equal(
                "Chunk PUT total length " + index,
                fileSize.ToString(CultureInfo.InvariantCulture),
                requests[index].Headers["OC-Total-Length"]);
        }
        Equal(
            "Chunk MOVE reads the assembled .file",
            uploadFolderUrl + "/.file",
            requests[3].Url);
        Equal(
            "Chunk MOVE carries the final destination",
            targetUrl,
            requests[3].Headers["Destination"]);
        Equal(
            "Chunk MOVE carries the total length",
            fileSize.ToString(CultureInfo.InvariantCulture),
            requests[3].Headers["OC-Total-Length"]);
    }

    private static void TestChunkFailureUsesShortCleanup(
        string tempRoot)
    {
        FileLinkPlannedFile file = CreateFile(
            tempRoot,
            "chunk-cleanup.bin",
            "chunk-cleanup.bin",
            new byte[] { 1, 2 });
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            if (options.Method == "PUT")
            {
                options.BodyWriter(Stream.Null);
                return Http(HttpStatusCode.BadRequest);
            }
            return options.Method == "DELETE"
                ? Http(HttpStatusCode.NoContent)
                : Http(HttpStatusCode.Created);
        });

        try
        {
            new FileLinkChunkUploader(client).Upload(
                new FileLinkUploadContext
                {
                    NormalizedBaseUrl =
                        "https://cloud.example.test",
                    UserId = "user",
                    RelativeFolderPath = "root"
                },
                file,
                new FileLinkUploadProgressCoordinator(),
                CancellationToken.None);
            Check(
                "Chunk failure raises its upload error",
                false,
                "no exception");
        }
        catch (TalkServiceException)
        {
            Check("Chunk failure raises its upload error", true);
        }

        Equal(
            "Chunk failure cleans its upload folder",
            "DELETE",
            requests[2].Method);
        Equal(
            "Chunk cleanup uses a ten-second timeout",
            10000,
            requests[2].TimeoutMs);
        Equal(
            "Chunk cleanup read timeout is also ten seconds",
            10000,
            requests[2].ReadWriteTimeoutMs);
    }

    private static void TestMoveRecoveryWithMatchingLength()
    {
        var requests = new List<NcHttpRequestOptions>();
        var client = new FileLinkDavClient(options =>
        {
            requests.Add(options);
            if (options.Method == "MOVE")
            {
                return TransportTimeout();
            }
            return ContentLengthResponse(123);
        });

        new FileLinkChunkUploader(client).MoveIntoPlace(
            "https://cloud.example.test/dav/uploads/user/id",
            "https://cloud.example.test/dav/files/user/root/file.bin",
            123,
            DateTime.UtcNow,
            CancellationToken.None);

        Equal("Indeterminate MOVE is sent once", 2, requests.Count);
        Equal("MOVE remains single attempt", "MOVE", requests[0].Method);
        Equal(
            "MOVE recovery probes the target",
            "PROPFIND",
            requests[1].Method);
        Equal(
            "MOVE recovery probe uses Depth 0",
            "0",
            requests[1].Headers["Depth"]);
        Check(
            "MOVE recovery requests content length",
            requests[1].Payload.Contains("getcontentlength"));
        Check(
            "MOVE recovery requests resource type",
            requests[1].Payload.Contains("resourcetype"));
    }

    private static void TestMoveRecoveryRejectsWrongLength()
    {
        int requestCount = 0;
        var client = new FileLinkDavClient(options =>
        {
            requestCount++;
            return options.Method == "MOVE"
                ? TransportTimeout()
                : ContentLengthResponse(122);
        });

        try
        {
            new FileLinkChunkUploader(client).MoveIntoPlace(
                "https://cloud.example.test/dav/uploads/user/id",
                "https://cloud.example.test/dav/files/user/root/file.bin",
                123,
                DateTime.UtcNow,
                CancellationToken.None);
            Check(
                "MOVE recovery rejects a wrong target length",
                false,
                "no exception");
        }
        catch (TalkServiceException)
        {
            Check(
                "MOVE recovery rejects a wrong target length",
                true);
        }
        Equal("Wrong-length recovery sends one probe", 2, requestCount);
    }

    private static void TestMoveRecoveryRejectsCollection()
    {
        int requestCount = 0;
        var client = new FileLinkDavClient(options =>
        {
            requestCount++;
            return options.Method == "MOVE"
                ? TransportTimeout()
                : ContentLengthResponse(123, true);
        });

        try
        {
            new FileLinkChunkUploader(client).MoveIntoPlace(
                "https://cloud.example.test/dav/uploads/user/id",
                "https://cloud.example.test/dav/files/user/root/file.bin",
                123,
                DateTime.UtcNow,
                CancellationToken.None);
            Check(
                "MOVE recovery rejects a DAV collection",
                false,
                "no exception");
        }
        catch (TalkServiceException)
        {
            Check("MOVE recovery rejects a DAV collection", true);
        }
        Equal("Collection recovery sends one probe", 2, requestCount);
    }

    private static void TestMoveHttpFailureIsNotRetried()
    {
        int requestCount = 0;
        var client = new FileLinkDavClient(options =>
        {
            requestCount++;
            return Http(HttpStatusCode.ServiceUnavailable);
        });

        try
        {
            new FileLinkChunkUploader(client).MoveIntoPlace(
                "https://cloud.example.test/dav/uploads/user/id",
                "https://cloud.example.test/dav/files/user/root/file.bin",
                123,
                DateTime.UtcNow,
                CancellationToken.None);
            Check(
                "MOVE HTTP failure stays single attempt",
                false,
                "no exception");
        }
        catch (TalkServiceException)
        {
            Check("MOVE HTTP failure stays single attempt", true);
        }
        Equal("MOVE HTTP failure sends no retry or probe", 1, requestCount);
    }

    private static FileLinkPlannedFile CreateFile(
        string tempRoot,
        string localName,
        string remotePath,
        byte[] content)
    {
        string path = Path.Combine(tempRoot, localName);
        File.WriteAllBytes(path, content);
        var info = new FileInfo(path);
        info.Refresh();
        return new FileLinkPlannedFile(
            path,
            remotePath,
            info.Length,
            info.LastWriteTimeUtc);
    }

    private static FileLinkPlannedFile CreateSizedFile(
        string tempRoot,
        string localName,
        string remotePath,
        long length)
    {
        string path = Path.Combine(tempRoot, localName);
        using (FileStream stream = new FileStream(
            path,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None))
        {
            stream.SetLength(length);
        }
        var info = new FileInfo(path);
        info.Refresh();
        return new FileLinkPlannedFile(
            path,
            remotePath,
            info.Length,
            info.LastWriteTimeUtc);
    }

    private static NcHttpResponse TransportTimeout()
    {
        return new NcHttpResponse
        {
            HasHttpResponse = false,
            TransportException = new WebException(
                "timeout",
                WebExceptionStatus.Timeout),
            Headers = RetryNow()
        };
    }

    private static NcHttpResponse Http(HttpStatusCode statusCode)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = statusCode,
            Headers = RetryNow()
        };
    }

    private static IDictionary<string, string> RetryNow()
    {
        return new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase)
        {
            { "Retry-After", "0" }
        };
    }

    private static NcHttpResponse BulkSuccess(
        FileLinkPlannedFile first,
        FileLinkPlannedFile second)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = HttpStatusCode.Created,
            Headers = RetryNow(),
            ParsedJson = new Dictionary<string, object>
            {
                {
                    "/root/" + first.RemotePath,
                    new Dictionary<string, object>
                    {
                        { "error", false }
                    }
                },
                {
                    "/root/" + second.RemotePath,
                    new Dictionary<string, object>
                    {
                        { "error", false }
                    }
                }
            }
        };
    }

    private static NcHttpResponse ContentLengthResponse(
        long length,
        bool isCollection = false)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = (HttpStatusCode)207,
            Headers = RetryNow(),
            ResponseText =
                "<?xml version=\"1.0\"?>"
                + "<d:multistatus xmlns:d=\"DAV:\">"
                + "<d:response><d:propstat><d:prop>"
                + "<d:resourcetype>"
                + (isCollection ? "<d:collection/>" : "")
                + "</d:resourcetype>"
                + "<d:getcontentlength>"
                + length.ToString(CultureInfo.InvariantCulture)
                + "</d:getcontentlength>"
                + "</d:prop></d:propstat></d:response>"
                + "</d:multistatus>"
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
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkSourceFile.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkBulkUploader.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkDirectUploader.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\FileLinkChunkUploader.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkPath.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\FileLinkUploadPolicy.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs")
    )
    $exe = Join-Path $TempRoot "FileLinkTransferTests.exe"
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

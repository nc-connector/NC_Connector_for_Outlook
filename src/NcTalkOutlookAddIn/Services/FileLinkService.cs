// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
        // Creates Nextcloud shares including uploading local files.
    // Encapsulates the full workflow (create folder, upload, create share, return result).
    internal sealed class FileLinkService
    {
        private const int ShareTypePublicLink = 3;
        private const long DirectUploadLimitBytes = 20L * 1024L * 1024L;
        private const long ChunkUploadChunkSizeBytes = 20L * 1024L * 1024L;
        private const long ChunkUploadMinChunkBytes = 5L * 1024L * 1024L;
        private const int ChunkUploadMaxChunks = 10000;
        private readonly TalkServiceConfiguration _configuration;
        private readonly NcHttpClient _httpClient;
        private static void LogFileLink(string message)
        {
            DiagnosticsLogger.Log(LogCategories.FileLink, message);
        }

        internal FileLinkService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }

            _configuration = configuration;
            _httpClient = new NcHttpClient(configuration);
        }

        internal FileLinkResult CreateFileShare(FileLinkRequest request, IProgress<FileLinkProgress> progress, CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            var uploadContext = PrepareUpload(request, cancellationToken);

            long totalBytes = CalculateTotalSize(request.Items);
            var progressState = new ProgressState(totalBytes, progress);

            UploadSelections(
                uploadContext,
                request.Items,
                new Progress<FileLinkUploadItemProgress>(p =>
                {
                    if (p != null && p.Status == FileLinkUploadStatus.Uploading && p.DeltaBytes > 0)
                    {
                        progressState.AddBytes(p.DeltaBytes, p.Selection != null ? p.Selection.LocalPath : null);
                    }
                }),
                null,
                cancellationToken);

            cancellationToken.ThrowIfCancellationRequested();

            return FinalizeShare(uploadContext, request, cancellationToken);
        }

        internal FileLinkUploadContext PrepareUpload(FileLinkRequest request, CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }

            cancellationToken.ThrowIfCancellationRequested();

            string normalizedBaseUrl = _configuration.GetNormalizedBaseUrl();
            string username = _configuration.Username ?? string.Empty;
            string basePath = NormalizeRelativePath(request.BasePath);
            string sanitizedShareName = SanitizeComponent(request.ShareName);
            if (string.IsNullOrWhiteSpace(sanitizedShareName))
            {
                sanitizedShareName = "Share";
            }

            DateTime shareDate = request.ShareDate.HasValue ? request.ShareDate.Value : DateTime.Now;
            string prefixFormat = NormalizeShareDatePrefixFormat(request.ShareDatePrefixFormat);
            string folderName = shareDate.ToString(prefixFormat, CultureInfo.InvariantCulture) + "_" + sanitizedShareName;
            string relativeFolderPath = CombineRelativePath(basePath, folderName);

            var context = new FileLinkUploadContext(
                normalizedBaseUrl,
                username,
                sanitizedShareName,
                folderName,
                relativeFolderPath);

            EnsureFolderExists(normalizedBaseUrl, username, basePath, cancellationToken, context.KnownFolderPaths);
            EnsureFolderExists(normalizedBaseUrl, username, relativeFolderPath, cancellationToken, context.KnownFolderPaths);
            context.KnownFolderPaths.Add(relativeFolderPath);

            return context;
        }

        internal bool FolderExists(string basePath, string folderName, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            string normalizedBaseUrl = _configuration.GetNormalizedBaseUrl();
            string username = _configuration.Username ?? string.Empty;
            string normalizedBasePath = NormalizeRelativePath(basePath);
            string relativePath = CombineRelativePath(normalizedBasePath, folderName);

            return FolderExistsInternal(normalizedBaseUrl, username, relativePath, cancellationToken);
        }

        internal void DeleteShareFolder(string relativeFolderPath, CancellationToken cancellationToken)
        {
            string normalizedPath = NormalizeRelativePath(relativeFolderPath);
            if (string.IsNullOrWhiteSpace(normalizedPath))
            {
                throw new ArgumentException("relativeFolderPath");
            }

            cancellationToken.ThrowIfCancellationRequested();

            string normalizedBaseUrl = _configuration.GetNormalizedBaseUrl();
            string username = _configuration.Username ?? string.Empty;
            string url = BuildDavUrl(normalizedBaseUrl, username, normalizedPath);
            DiagnosticsLogger.LogApi("DELETE " + url);

            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "DELETE",
                Url = url,
                TimeoutMs = 90000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                DiagnosticsLogger.LogException(LogCategories.Api, "Share folder delete failed (" + url + ").", transport);
                throw new TalkServiceException("Share folder delete failed: " + (transport != null ? transport.Message : "no HTTP response"), false, 0, null);
            }
            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                LogFileLink("Share folder delete skipped (already removed): " + normalizedPath);
                return;
            }
            if (response.StatusCode != HttpStatusCode.NoContent
                && response.StatusCode != HttpStatusCode.OK
                && response.StatusCode != HttpStatusCode.Accepted)
            {
                bool authError = response.StatusCode == HttpStatusCode.Unauthorized
                                 || response.StatusCode == HttpStatusCode.Forbidden;
                throw new TalkServiceException("Share folder could not be deleted: " + normalizedPath, authError, response.StatusCode, response.ResponseText);
            }

            DiagnosticsLogger.LogApi("DELETE " + url + " -> " + response.StatusCode);
        }

        internal void UploadSelections(
            FileLinkUploadContext context,
            IList<FileLinkSelection> selections,
            IProgress<FileLinkUploadItemProgress> progress,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            CancellationToken cancellationToken)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }
            if (selections == null)
            {
                throw new ArgumentNullException("selections");
            }
            foreach (var selection in selections)
            {
                cancellationToken.ThrowIfCancellationRequested();

                long totalBytes = CalculateSelectionSize(selection);
                var tracker = new SelectionUploadTracker(totalBytes);

                ReportProgress(progress, selection, tracker, FileLinkUploadStatus.Uploading, null, 0);

                try
                {
                    if (selection.SelectionType == FileLinkSelectionType.File)
                    {
                        UploadSingleFileSelection(context, selection, tracker, progress, duplicateResolver, cancellationToken);
                    }
                    else
                    {
                        UploadDirectorySelection(context, selection, tracker, progress, duplicateResolver, cancellationToken);
                    }

                    ReportProgress(progress, selection, tracker, FileLinkUploadStatus.Completed, null, 0);
                }
                catch (Exception ex)
                {
                    ReportProgress(progress, selection, tracker, FileLinkUploadStatus.Failed, ex.Message, 0);
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Upload failed for '" + (selection != null ? selection.LocalPath : string.Empty) + "'.", ex);
                    throw;
                }
            }
        }

        internal FileLinkResult FinalizeShare(FileLinkUploadContext context, FileLinkRequest request, CancellationToken cancellationToken)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            var shareData = CreateShare(
                context.NormalizedBaseUrl,
                context.Username,
                context.RelativeFolderPath,
                context.SanitizedShareName,
                request,
                cancellationToken);

            return new FileLinkResult(
                shareData.Url,
                shareData.Id,
                shareData.Token,
                request.PasswordEnabled ? request.Password : null,
                request.ExpireEnabled ? request.ExpireDate : null,
                request.Permissions,
                context.FolderName,
                context.RelativeFolderPath);
        }

        private bool FolderExistsInternal(string baseUrl, string username, string relativePath, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(relativePath))
            {
                return false;
            }
            string url = BuildDavUrl(baseUrl, username, relativePath);
            cancellationToken.ThrowIfCancellationRequested();

            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "PROPFIND",
                Url = url,
                ContentType = "application/xml; charset=utf-8",
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                Headers = new Dictionary<string, string> { { "Depth", "0" } }
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                DiagnosticsLogger.LogException(LogCategories.Api, "DAV folder check failed (" + url + ").", transport);
                throw new TalkServiceException("Folder check failed: " + (transport != null ? transport.Message : "no HTTP response"), false, 0, null);
            }
            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                LogFileLink("DAV folder not found: " + url);
                return false;
            }
            if (response.StatusCode == HttpStatusCode.Unauthorized
                || response.StatusCode == HttpStatusCode.Forbidden)
            {
                throw new TalkServiceException("Folder check failed: HTTP " + (int)response.StatusCode, true, response.StatusCode, response.ResponseText);
            }
            return response.StatusCode == HttpStatusCode.OK
                   || response.StatusCode == HttpStatusCode.NoContent
                   || (int)response.StatusCode == 207;
        }

        private static void ReportProgress(
            IProgress<FileLinkUploadItemProgress> progress,
            FileLinkSelection selection,
            SelectionUploadTracker tracker,
            FileLinkUploadStatus status,
            string message,
            long deltaBytes)
        {
            if (progress == null)
            {
                return;
            }

            progress.Report(new FileLinkUploadItemProgress(
                selection,
                tracker.UploadedBytes,
                tracker.TotalBytes,
                status,
                message,
                deltaBytes));
        }

        private static long CalculateSelectionSize(FileLinkSelection selection)
        {
            if (selection == null)
            {
                return 0;
            }
            if (selection.SelectionType == FileLinkSelectionType.File)
            {
                var info = new FileInfo(selection.LocalPath);
                return info.Exists ? info.Length : 0;
            }

            long total = 0;
            if (Directory.Exists(selection.LocalPath))
            {
                foreach (string file in Directory.EnumerateFiles(selection.LocalPath, "*", SearchOption.AllDirectories))
                {
                    var info = new FileInfo(file);
                    if (info.Exists)
                    {
                        total += info.Length;
                    }
                }
            }
            return total;
        }

        private void UploadSingleFileSelection(
            FileLinkUploadContext context,
            FileLinkSelection selection,
            SelectionUploadTracker tracker,
            IProgress<FileLinkUploadItemProgress> progress,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            CancellationToken cancellationToken)
        {
            string remoteFolder = context.RelativeFolderPath;
            string fileName = SanitizeComponent(Path.GetFileName(selection.LocalPath));
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new TalkServiceException("File has no valid name: " + selection.LocalPath, false, 0, null);
            }
            string uniqueName = EnsureUniqueName(context, remoteFolder, fileName, duplicateResolver, selection, false, cancellationToken);
            string remotePath = CombineRelativePath(remoteFolder, uniqueName);

            UploadFileContent(context, remotePath, selection.LocalPath, tracker, progress, selection, cancellationToken);
        }

        private void UploadDirectorySelection(
            FileLinkUploadContext context,
            FileLinkSelection selection,
            SelectionUploadTracker tracker,
            IProgress<FileLinkUploadItemProgress> progress,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            CancellationToken cancellationToken)
        {
            if (!Directory.Exists(selection.LocalPath))
            {
                return;
            }
            string relativeRoot = SanitizeComponent(Path.GetFileName(selection.LocalPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)));
            if (string.IsNullOrWhiteSpace(relativeRoot))
            {
                relativeRoot = "Folder";
            }
            string uniqueRoot = EnsureUniqueName(context, context.RelativeFolderPath, relativeRoot, duplicateResolver, selection, true, cancellationToken);
            string remoteRoot = CombineRelativePath(context.RelativeFolderPath, uniqueRoot);

            // For directory uploads we enforce MKCOL without relying on cache hints.
            // A stale known-folder cache can otherwise skip folder creation and the first PUT fails with 404.
            EnsureFolderExists(context.NormalizedBaseUrl, context.Username, remoteRoot, cancellationToken);
            context.KnownFolderPaths.Add(remoteRoot);

            foreach (string directory in Directory.EnumerateDirectories(selection.LocalPath, "*", SearchOption.AllDirectories))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string relativeSub = ConvertPath(GetRelativePath(selection.LocalPath, directory));
                string remoteSub = CombineRelativePath(remoteRoot, relativeSub);
                EnsureFolderExists(context.NormalizedBaseUrl, context.Username, remoteSub, cancellationToken);
                context.KnownFolderPaths.Add(remoteSub);
            }
            foreach (string file in Directory.EnumerateFiles(selection.LocalPath, "*", SearchOption.AllDirectories))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string relativeFile = ConvertPath(GetRelativePath(selection.LocalPath, file));
                string remoteFolder = GetRemoteFolder(remoteRoot, relativeFile);
                EnsureFolderExists(context.NormalizedBaseUrl, context.Username, remoteFolder, cancellationToken);
                context.KnownFolderPaths.Add(remoteFolder);

                string remoteFileName = SanitizeComponent(Path.GetFileName(relativeFile));
                if (string.IsNullOrWhiteSpace(remoteFileName))
                {
                    remoteFileName = "File";
                }
                string uniqueName = EnsureUniqueName(context, remoteFolder, remoteFileName, duplicateResolver, selection, false, cancellationToken);
                string remotePath = CombineRelativePath(remoteFolder, uniqueName);

                UploadFileContent(context, remotePath, file, tracker, progress, selection, cancellationToken);
            }
        }

        private void UploadFileContent(
            FileLinkUploadContext context,
            string remotePath,
            string localPath,
            SelectionUploadTracker tracker,
            IProgress<FileLinkUploadItemProgress> progress,
            FileLinkSelection selection,
            CancellationToken cancellationToken)
        {
            var fileInfo = new FileInfo(localPath);
            if (!fileInfo.Exists)
            {
                return;
            }
            string targetUrl = BuildDavUrl(context.NormalizedBaseUrl, context.Username, remotePath);
            if (ShouldUseChunkedUpload(fileInfo))
            {
                UploadFileContentChunked(context, targetUrl, fileInfo, tracker, progress, selection, cancellationToken);
                return;
            }

            UploadFileContentDirect(targetUrl, localPath, fileInfo, tracker, progress, selection, cancellationToken);
        }

        private void UploadFileContentDirect(
            string targetUrl,
            string localPath,
            FileInfo fileInfo,
            SelectionUploadTracker tracker,
            IProgress<FileLinkUploadItemProgress> progress,
            FileLinkSelection selection,
            CancellationToken cancellationToken)
        {
            DiagnosticsLogger.LogApi("PUT " + targetUrl);
            LogFileLink("Upload method selected (method=put, file=\"" + fileInfo.Name + "\", bytes=" + fileInfo.Length.ToString(CultureInfo.InvariantCulture) + ").");

            cancellationToken.ThrowIfCancellationRequested();

            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "PUT",
                Url = targetUrl,
                TimeoutMs = 120000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                ContentLength = fileInfo.Length,
                ContentType = "application/octet-stream",
                BodyWriter = requestStream =>
                {
                    using (FileStream fileStream = fileInfo.OpenRead())
                    {
                        byte[] buffer = new byte[81920];
                        int bytesRead;
                        while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            requestStream.Write(buffer, 0, bytesRead);
                            tracker.AddBytes(bytesRead);
                            ReportProgress(progress, selection, tracker, FileLinkUploadStatus.Uploading, null, bytesRead);
                        }
                    }
                }
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                throw new TalkServiceException("File could not be uploaded: " + (transport != null ? transport.Message : localPath), false, 0, null);
            }
            if (response.StatusCode != HttpStatusCode.Created && response.StatusCode != HttpStatusCode.OK && response.StatusCode != HttpStatusCode.NoContent)
            {
                throw new TalkServiceException("File could not be uploaded: " + localPath, false, response.StatusCode, response.ResponseText);
            }

            DiagnosticsLogger.LogApi("PUT " + targetUrl + " -> " + response.StatusCode);
        }

        private void UploadFileContentChunked(
            FileLinkUploadContext context,
            string targetUrl,
            FileInfo fileInfo,
            SelectionUploadTracker tracker,
            IProgress<FileLinkUploadItemProgress> progress,
            FileLinkSelection selection,
            CancellationToken cancellationToken)
        {
            long totalSize = fileInfo.Length;
            long chunkSize = GetChunkSize(totalSize);
            long chunkCount = (totalSize + chunkSize - 1) / chunkSize;
            if (chunkCount > ChunkUploadMaxChunks)
            {
                throw new TalkServiceException("File could not be uploaded: too many chunks.", false, 0, null);
            }

            string uploadFolderUrl = BuildChunkUploadFolderUrl(context.NormalizedBaseUrl, context.Username);
            LogFileLink(
                "Upload method selected (method=chunked-v2, file=\""
                + fileInfo.Name
                + "\", bytes="
                + totalSize.ToString(CultureInfo.InvariantCulture)
                + ", chunks="
                + chunkCount.ToString(CultureInfo.InvariantCulture)
                + ", chunkSize="
                + chunkSize.ToString(CultureInfo.InvariantCulture)
                + ").");

            bool cleanupRequired = false;
            CreateChunkUploadFolder(uploadFolderUrl, targetUrl, cancellationToken);
            cleanupRequired = true;
            try
            {
                using (FileStream fileStream = fileInfo.OpenRead())
                {
                    byte[] buffer = new byte[81920];
                    for (long index = 0; index < chunkCount; index++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        long offset = index * chunkSize;
                        long length = Math.Min(chunkSize, totalSize - offset);
                        string chunkName = (index + 1).ToString("00000", CultureInfo.InvariantCulture);
                        string chunkUrl = uploadFolderUrl + "/" + chunkName;
                        UploadChunk(chunkUrl, targetUrl, totalSize, fileStream, offset, length, buffer, tracker, progress, selection, cancellationToken);
                    }
                }

                MoveChunkedUpload(uploadFolderUrl, targetUrl, totalSize, cancellationToken);
                cleanupRequired = false;
                LogFileLink("Chunked upload completed (file=\"" + fileInfo.Name + "\", bytes=" + totalSize.ToString(CultureInfo.InvariantCulture) + ").");
            }
            catch
            {
                if (cleanupRequired)
                {
                    CleanupChunkUploadFolder(uploadFolderUrl);
                }
                throw;
            }
        }

        private void CreateChunkUploadFolder(string uploadFolderUrl, string targetUrl, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DiagnosticsLogger.LogApi("MKCOL " + uploadFolderUrl);
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "MKCOL",
                Url = uploadFolderUrl,
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                Headers = new Dictionary<string, string> { { "Destination", targetUrl } }
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                throw new TalkServiceException("File could not be uploaded: " + (transport != null ? transport.Message : "no HTTP response"), false, 0, null);
            }
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new TalkServiceException("File could not be uploaded: chunk folder could not be created.", false, response.StatusCode, response.ResponseText);
            }
            DiagnosticsLogger.LogApi("MKCOL " + uploadFolderUrl + " -> " + response.StatusCode);
        }

        private void UploadChunk(
            string chunkUrl,
            string targetUrl,
            long totalSize,
            FileStream fileStream,
            long offset,
            long length,
            byte[] buffer,
            SelectionUploadTracker tracker,
            IProgress<FileLinkUploadItemProgress> progress,
            FileLinkSelection selection,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DiagnosticsLogger.LogApi("PUT " + chunkUrl);
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "PUT",
                Url = chunkUrl,
                TimeoutMs = 120000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                ContentLength = length,
                ContentType = "application/octet-stream",
                Headers = new Dictionary<string, string>
                {
                    { "Destination", targetUrl },
                    { "OC-Total-Length", totalSize.ToString(CultureInfo.InvariantCulture) }
                },
                BodyWriter = requestStream =>
                {
                    fileStream.Seek(offset, SeekOrigin.Begin);
                    long remaining = length;
                    while (remaining > 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int toRead = (int)Math.Min(buffer.Length, remaining);
                        int bytesRead = fileStream.Read(buffer, 0, toRead);
                        if (bytesRead <= 0)
                        {
                            throw new EndOfStreamException("Unexpected end of file during chunked upload.");
                        }
                        requestStream.Write(buffer, 0, bytesRead);
                        remaining -= bytesRead;
                        tracker.AddBytes(bytesRead);
                        ReportProgress(progress, selection, tracker, FileLinkUploadStatus.Uploading, null, bytesRead);
                    }
                }
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                throw new TalkServiceException("File could not be uploaded: " + (transport != null ? transport.Message : "no HTTP response"), false, 0, null);
            }
            if (!IsSuccessStatus(response.StatusCode))
            {
                throw new TalkServiceException("File could not be uploaded: chunk upload failed.", false, response.StatusCode, response.ResponseText);
            }
            DiagnosticsLogger.LogApi("PUT " + chunkUrl + " -> " + response.StatusCode);
        }

        private void MoveChunkedUpload(string uploadFolderUrl, string targetUrl, long totalSize, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string sourceUrl = uploadFolderUrl + "/.file";
            DiagnosticsLogger.LogApi("MOVE " + sourceUrl);
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "MOVE",
                Url = sourceUrl,
                TimeoutMs = 120000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                Headers = new Dictionary<string, string>
                {
                    { "Destination", targetUrl },
                    { "OC-Total-Length", totalSize.ToString(CultureInfo.InvariantCulture) }
                }
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                throw new TalkServiceException("File could not be uploaded: " + (transport != null ? transport.Message : "no HTTP response"), false, 0, null);
            }
            if (!IsSuccessStatus(response.StatusCode))
            {
                throw new TalkServiceException("File could not be uploaded: chunk assembly failed.", false, response.StatusCode, response.ResponseText);
            }
            DiagnosticsLogger.LogApi("MOVE " + sourceUrl + " -> " + response.StatusCode);
        }

        private void CleanupChunkUploadFolder(string uploadFolderUrl)
        {
            try
            {
                DiagnosticsLogger.LogApi("DELETE " + uploadFolderUrl);
                NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
                {
                    Method = "DELETE",
                    Url = uploadFolderUrl,
                    TimeoutMs = 60000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false
                });
                if (response.HasHttpResponse && (IsSuccessStatus(response.StatusCode) || response.StatusCode == HttpStatusCode.NotFound))
                {
                    DiagnosticsLogger.LogApi("DELETE " + uploadFolderUrl + " -> " + response.StatusCode);
                    return;
                }
                LogFileLink("Chunked upload cleanup failed (status=" + (response.HasHttpResponse ? ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture) : "n/a") + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Chunked upload cleanup failed.", ex);
            }
        }

        private string EnsureUniqueName(
            FileLinkUploadContext context,
            string remoteFolder,
            string sanitizedName,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            FileLinkSelection selection,
            bool isDirectory,
            CancellationToken cancellationToken)
        {
            string folderKey = remoteFolder ?? string.Empty;
            string fullPath = CombineRelativePath(folderKey, sanitizedName);
            HashSet<string> primaryKnownSet = isDirectory ? context.KnownFolderPaths : context.KnownFilePaths;
            HashSet<string> secondaryKnownSet = isDirectory ? context.KnownFilePaths : context.KnownFolderPaths;
            HashSet<string> primaryReservedSet = isDirectory ? context.ReservedFolderPaths : context.ReservedFilePaths;
            HashSet<string> secondaryReservedSet = isDirectory ? context.ReservedFilePaths : context.ReservedFolderPaths;

            while (primaryKnownSet.Contains(fullPath)
                || secondaryKnownSet.Contains(fullPath)
                || primaryReservedSet.Contains(fullPath)
                || secondaryReservedSet.Contains(fullPath))
            {
                if (duplicateResolver == null)
                {
                    throw new TalkServiceException("Duplicate name in target directory: " + sanitizedName, false, 0, null);
                }
                string newName = duplicateResolver(new FileLinkDuplicateInfo(selection, remoteFolder, sanitizedName, isDirectory));
                if (string.IsNullOrWhiteSpace(newName))
                {
                    throw new OperationCanceledException("Upload cancelled: no unique name available.");
                }

            cancellationToken.ThrowIfCancellationRequested();

                sanitizedName = SanitizeComponent(newName);
                if (string.IsNullOrWhiteSpace(sanitizedName))
                {
                    throw new TalkServiceException("Invalid name after rename.", false, 0, null);
                }

                fullPath = CombineRelativePath(folderKey, sanitizedName);
            }

            primaryReservedSet.Add(fullPath);
            return sanitizedName;
        }

        private static string GetRemoteFolder(string root, string relativeFile)
        {
            int index = relativeFile.LastIndexOf('/');
            if (index < 0)
            {
                return root;
            }
            string folder = relativeFile.Substring(0, index);
            return CombineRelativePath(root, folder);
        }

        private void EnsureFolderExists(string baseUrl, string username, string relativePath, CancellationToken cancellationToken)
        {
            EnsureFolderExists(baseUrl, username, relativePath, cancellationToken, null);
        }

        private void EnsureFolderExists(
            string baseUrl,
            string username,
            string relativePath,
            CancellationToken cancellationToken,
            HashSet<string> knownFolderPaths)
        {
            if (string.IsNullOrEmpty(relativePath))
            {
                return;
            }
            string[] segments = relativePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var current = new List<string>();
            foreach (string segment in segments)
            {
                cancellationToken.ThrowIfCancellationRequested();
                current.Add(segment);
                string path = string.Join("/", current.ToArray());
                if (knownFolderPaths != null && knownFolderPaths.Contains(path))
                {
                    continue;
                }
                string url = BuildDavUrl(baseUrl, username, path);
                DiagnosticsLogger.LogApi("MKCOL " + url);
                NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
                {
                    Method = "MKCOL",
                    Url = url,
                    TimeoutMs = 60000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false
                });

                if (!response.HasHttpResponse)
                {
                    Exception transport = response.TransportException;
                    DiagnosticsLogger.LogException(LogCategories.Api, "MKCOL failed (" + url + ").", transport);
                    throw new TalkServiceException("Directory could not be created: " + (transport != null ? transport.Message : path), false, 0, null);
                }
                if (response.StatusCode != HttpStatusCode.Created
                    && response.StatusCode != HttpStatusCode.MethodNotAllowed
                    && response.StatusCode != HttpStatusCode.Conflict)
                {
                    throw new TalkServiceException("Directory could not be created: " + path, false, response.StatusCode, response.ResponseText);
                }
                if (knownFolderPaths != null)
                {
                    knownFolderPaths.Add(path);
                }

                DiagnosticsLogger.LogApi("MKCOL " + url + " -> " + response.StatusCode);
            }
        }

                // Create the public share through the documented OCS create endpoint and then update mutable metadata.
        private ShareData CreateShare(string baseUrl, string username, string relativeFolderPath, string shareName, FileLinkRequest request, CancellationToken cancellationToken)
        {
            string url = baseUrl.TrimEnd('/') + "/ocs/v2.php/apps/files_sharing/api/v1/shares";
            string payload = BuildShareCreatePayload(relativeFolderPath, shareName, request);
            ShareData shareData = ExecuteShareCreateRequest(url, payload, cancellationToken);
            UpdateShareMetadata(baseUrl, shareData, request, cancellationToken);
            return shareData;
        }

                // Build the documented form-encoded create payload for OCS public-link shares.
        private string BuildShareCreatePayload(string relativeFolderPath, string shareName, FileLinkRequest request)
        {
            var builder = new StringBuilder();
            builder.Append("path=").Append(Uri.EscapeDataString("/" + relativeFolderPath));
            builder.Append("&shareType=").Append(ShareTypePublicLink.ToString(CultureInfo.InvariantCulture));
            builder.Append("&permissions=").Append(CalculatePermissionValue(request.Permissions).ToString(CultureInfo.InvariantCulture));

            if (request.PasswordEnabled && !string.IsNullOrEmpty(request.Password))
            {
                builder.Append("&password=").Append(Uri.EscapeDataString(request.Password));
            }
            if (request.ExpireEnabled && request.ExpireDate.HasValue)
            {
                builder.Append("&expireDate=").Append(Uri.EscapeDataString(request.ExpireDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)));
            }
            if (!string.IsNullOrWhiteSpace(shareName))
            {
                builder.Append("&label=").Append(Uri.EscapeDataString(shareName));
            }
            if ((request.Permissions & FileLinkPermissionFlags.Create) == FileLinkPermissionFlags.Create)
            {
                builder.Append("&publicUpload=true");
            }
            return builder.ToString();
        }

                // Update mutable share metadata through the documented OCS update endpoint.
        private void UpdateShareMetadata(string baseUrl, ShareData shareData, FileLinkRequest request, CancellationToken cancellationToken)
        {
            if (shareData == null || string.IsNullOrWhiteSpace(shareData.Id))
            {
                return;
            }
            string url = baseUrl.TrimEnd('/') + "/ocs/v2.php/apps/files_sharing/api/v1/shares/" + Uri.EscapeDataString(shareData.Id);
            var payload = new StringBuilder();
            payload.Append("permissions=").Append(Uri.EscapeDataString(CalculatePermissionValue(request.Permissions).ToString(CultureInfo.InvariantCulture)));
            payload.Append("&publicUpload=").Append((request.Permissions & FileLinkPermissionFlags.Create) == FileLinkPermissionFlags.Create ? "true" : "false");
            payload.Append("&note=").Append(Uri.EscapeDataString(request.NoteEnabled ? (request.Note ?? string.Empty) : string.Empty));
            payload.Append("&attributes=").Append(Uri.EscapeDataString("[]"));
            if (request.ExpireEnabled && request.ExpireDate.HasValue)
            {
                payload.Append("&expireDate=").Append(Uri.EscapeDataString(request.ExpireDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)));
            }
            if (request.PasswordEnabled && !string.IsNullOrEmpty(request.Password))
            {
                payload.Append("&password=").Append(Uri.EscapeDataString(request.Password));
            }

            ExecuteShareMetadataUpdateRequest(url, payload.ToString(), cancellationToken);
        }

        private ShareData ExecuteShareCreateRequest(string url, string formPayload, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            DiagnosticsLogger.LogApi("POST " + url);
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "POST",
                Url = url,
                Payload = formPayload ?? string.Empty,
                Accept = "application/json",
                ContentType = "application/x-www-form-urlencoded",
                TimeoutMs = 90000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                DiagnosticsLogger.LogException(LogCategories.Api, "Share create failed (" + url + ").", transport);
                throw new TalkServiceException("Share creation failed: " + (transport != null ? transport.Message : "no HTTP response"), false, 0, null);
            }

            DiagnosticsLogger.LogApi("POST " + url + " -> " + response.StatusCode);
            if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            {
                string detail = NcJson.ExtractOcsErrorMessage(response.ParsedJson);
                if (string.IsNullOrWhiteSpace(detail))
                {
                    detail = "HTTP " + (int)response.StatusCode;
                }
                bool authError = response.StatusCode == HttpStatusCode.Unauthorized
                                 || response.StatusCode == HttpStatusCode.Forbidden;
                throw new TalkServiceException("Share creation failed: " + detail, authError, response.StatusCode, response.ResponseText);
            }
            return ParseShareData(response.ParsedJson, response.ResponseText);
        }

        private void ExecuteShareMetadataUpdateRequest(string url, string formPayload, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            DiagnosticsLogger.LogApi("PUT " + url);
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "PUT",
                Url = url,
                Payload = formPayload ?? string.Empty,
                Accept = "application/json",
                ContentType = "application/x-www-form-urlencoded",
                TimeoutMs = 90000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                DiagnosticsLogger.LogException(LogCategories.Api, "Share metadata update failed (" + url + ").", transport);
                throw new TalkServiceException("Share metadata update failed: " + (transport != null ? transport.Message : "no HTTP response"), false, 0, null);
            }

            DiagnosticsLogger.LogApi("PUT " + url + " -> " + response.StatusCode);
            if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            {
                string detail = NcJson.ExtractOcsErrorMessage(response.ParsedJson);
                if (string.IsNullOrWhiteSpace(detail))
                {
                    detail = "HTTP " + (int)response.StatusCode;
                }
                bool authError = response.StatusCode == HttpStatusCode.Unauthorized
                                 || response.StatusCode == HttpStatusCode.Forbidden;
                throw new TalkServiceException("Share metadata update failed: " + detail, authError, response.StatusCode, response.ResponseText);
            }
        }

        private static string BuildDavUrl(string baseUrl, string username, string relativePath)
        {
            string[] segments = relativePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string encoded = string.Join("/", segments.Select(Uri.EscapeDataString));
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/files/{1}/{2}",
                baseUrl.TrimEnd('/'),
                Uri.EscapeDataString(username ?? string.Empty),
                encoded);
        }

        private static string BuildChunkUploadFolderUrl(string baseUrl, string username)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/uploads/{1}/{2}",
                baseUrl.TrimEnd('/'),
                Uri.EscapeDataString(username ?? string.Empty),
                Uri.EscapeDataString(BuildChunkUploadId()));
        }

        private static string BuildChunkUploadId()
        {
            return "ncconnector-" + Guid.NewGuid().ToString("N");
        }

        private static bool ShouldUseChunkedUpload(FileInfo fileInfo)
        {
            return fileInfo != null && fileInfo.Length > DirectUploadLimitBytes;
        }

        private static long GetChunkSize(long fileSize)
        {
            long minimumForChunkLimit = (long)Math.Ceiling((double)Math.Max(0, fileSize) / ChunkUploadMaxChunks);
            return Math.Max(ChunkUploadChunkSizeBytes, Math.Max(ChunkUploadMinChunkBytes, minimumForChunkLimit));
        }

        private static bool IsSuccessStatus(HttpStatusCode statusCode)
        {
            int status = (int)statusCode;
            return status >= 200 && status < 300;
        }

        private static ShareData ParseShareData(IDictionary<string, object> parsedJson, string responseText)
        {
            if (parsedJson == null)
            {
                throw new TalkServiceException("Share creation failed: invalid response.", false, 0, responseText);
            }

            IDictionary<string, object> data = NcJson.GetOcsData(parsedJson);
            var result = new ShareData
            {
                Id = NcJson.GetTrimmedString(data, "id"),
                Url = NcJson.GetTrimmedString(data, "url"),
                Token = NcJson.GetTrimmedString(data, "token")
            };

            if (string.IsNullOrWhiteSpace(result.Id) || string.IsNullOrWhiteSpace(result.Url))
            {
                throw new TalkServiceException("Share creation failed: incomplete response.", false, 0, responseText);
            }
            return result;
        }

        internal static string NormalizeRelativePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }
            return string.Join("/", path.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries).Select(SanitizeComponent));
        }

        internal static string CombineRelativePath(string basePath, string component)
        {
            if (string.IsNullOrEmpty(basePath))
            {
                return component ?? string.Empty;
            }
            if (string.IsNullOrEmpty(component))
            {
                return basePath;
            }
            return basePath.TrimEnd('/') + "/" + component.TrimStart('/');
        }

        internal static string SanitizeComponent(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            var invalid = Path.GetInvalidFileNameChars();
            var builder = new StringBuilder(value.Length);
            foreach (char c in value)
            {
                builder.Append(invalid.Contains(c) ? '_' : c);
            }
            return builder.ToString().Trim();
        }

        internal static string NormalizeShareDatePrefixFormat(string value)
        {
            return "yyyyMMdd";
        }

        private static string ConvertPath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return string.Empty;
            }
            return string.Join("/", path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Select(SanitizeComponent));
        }

        private static string GetRelativePath(string root, string fullPath)
        {
            if (string.IsNullOrEmpty(root))
            {
                return fullPath;
            }
            var rootUri = new Uri(AppendDirectorySeparator(root));
            var fullUri = new Uri(fullPath);
            string relative = Uri.UnescapeDataString(rootUri.MakeRelativeUri(fullUri).ToString());
            return relative.Replace('/', Path.DirectorySeparatorChar);
        }

        private static string AppendDirectorySeparator(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return path;
            }
            if (!path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                return path + Path.DirectorySeparatorChar;
            }
            return path;
        }

        private static int CalculatePermissionValue(FileLinkPermissionFlags permissions)
        {
            int value = 0;
            if ((permissions & FileLinkPermissionFlags.Read) == FileLinkPermissionFlags.Read)
            {
                value |= 1;
            }
            if ((permissions & FileLinkPermissionFlags.Write) == FileLinkPermissionFlags.Write)
            {
                value |= 2;
            }
            if ((permissions & FileLinkPermissionFlags.Create) == FileLinkPermissionFlags.Create)
            {
                value |= 4;
            }
            if ((permissions & FileLinkPermissionFlags.Delete) == FileLinkPermissionFlags.Delete)
            {
                value |= 8;
            }
            return value;
        }

        private static long CalculateTotalSize(IEnumerable<FileLinkSelection> items)
        {
            long total = 0;
            foreach (var item in items)
            {
                if (item.SelectionType == FileLinkSelectionType.File)
                {
                    var info = new FileInfo(item.LocalPath);
                    if (info.Exists)
                    {
                        total += info.Length;
                    }
                }
                else if (Directory.Exists(item.LocalPath))
                {
                    foreach (string file in Directory.EnumerateFiles(item.LocalPath, "*", SearchOption.AllDirectories))
                    {
                        var info = new FileInfo(file);
                        if (info.Exists)
                        {
                            total += info.Length;
                        }
                    }
                }
            }
            return total;
        }

        private sealed class SelectionUploadTracker
        {
            internal SelectionUploadTracker(long totalBytes)
            {
                TotalBytes = totalBytes < 0 ? 0 : totalBytes;
            }

            internal long TotalBytes { get; private set; }

            internal long UploadedBytes { get; private set; }

            internal void AddBytes(long bytes)
            {
                if (bytes <= 0)
                {
                    return;
                }

                UploadedBytes += bytes;
                if (TotalBytes > 0 && UploadedBytes > TotalBytes)
                {
                    UploadedBytes = TotalBytes;
                }
            }
        }

        private sealed class ProgressState
        {
            private readonly IProgress<FileLinkProgress> _progress;
            private readonly long _totalBytes;
            private long _uploadedBytes;

            internal ProgressState(long totalBytes, IProgress<FileLinkProgress> progress)
            {
                _totalBytes = totalBytes;
                _progress = progress;
            }

            internal void AddBytes(long bytes, string item)
            {
                if (_totalBytes <= 0 || _progress == null)
                {
                    return;
                }

                _uploadedBytes += bytes;
                _progress.Report(new FileLinkProgress(_totalBytes, _uploadedBytes, item));
            }
        }

        private sealed class ShareData
        {
            internal string Id { get; set; }

            internal string Url { get; set; }

            internal string Token { get; set; }
        }

    }
}


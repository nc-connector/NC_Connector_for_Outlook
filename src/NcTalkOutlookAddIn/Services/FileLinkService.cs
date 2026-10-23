// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Coordinates planning, DAV transfer, and OCS share creation.
    internal sealed class FileLinkService
    {
        private const int AttachmentShareNameCandidateLimit = 1000;

        private readonly TalkServiceConfiguration _configuration;
        private readonly FileLinkDavClient _davClient;
        private readonly FileLinkShareClient _shareClient;
        private readonly FileLinkTransferService _transferService;
        private NextcloudCapabilitiesSnapshot _capabilitiesSnapshot;

        internal FileLinkService(TalkServiceConfiguration configuration)
            : this(configuration, null)
        {
        }

        internal FileLinkService(
            TalkServiceConfiguration configuration,
            NextcloudCapabilitiesSnapshot capabilitiesSnapshot)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }

            _configuration = configuration;
            var httpClient = new NcHttpClient(configuration);
            _davClient = new FileLinkDavClient(httpClient);
            _shareClient = new FileLinkShareClient(httpClient);
            _transferService = new FileLinkTransferService(_davClient);
            _capabilitiesSnapshot = capabilitiesSnapshot;
        }

        internal bool ManualShareFolderExists(
            FileLinkRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (request.AttachmentMode)
            {
                throw new ArgumentException(
                    "Attachment shares use automatic collision resolution.",
                    "request");
            }

            cancellationToken.ThrowIfCancellationRequested();
            using (DiagnosticsLogger.BeginOperation(
                LogCategories.FileLink,
                "FileLink.ManualSharePreflight"))
            {
                ResolveRequiredCapabilities();
                FileLinkShareTarget target = ResolveShareTarget(request);
                string normalizedBaseUrl =
                    _configuration.GetNormalizedBaseUrl();
                string userId =
                    NextcloudUserIdentityService.ResolveCurrentUserId(
                        _configuration);
                bool exists = _davClient.ResourceExists(
                    FileLinkDavClient.BuildFileUrl(
                        normalizedBaseUrl,
                        userId,
                        target.RelativeFolderPath),
                    Strings.FileLinkWizardFolderCheckFailedFormat,
                    cancellationToken);
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Manual share folder preflight completed (exists="
                    + exists.ToString()
                    + ").");
                return exists;
            }
        }

        internal NextcloudStorageListing ListNextcloudDirectory(
            string relativePath,
            CancellationToken cancellationToken)
        {
            string normalizedBaseUrl =
                _configuration.GetNormalizedBaseUrl();
            string userId =
                NextcloudUserIdentityService.ResolveCurrentUserId(
                    _configuration);
            return _davClient.ListDirectory(
                normalizedBaseUrl,
                userId,
                relativePath,
                cancellationToken);
        }

        internal byte[] ReadNextcloudFilePreview(
            NextcloudStorageEntry entry,
            long maximumBytes,
            CancellationToken cancellationToken)
        {
            if (entry == null || entry.IsDirectory)
            {
                throw new ArgumentException(
                    "A Nextcloud file is required.",
                    "entry");
            }

            string normalizedBaseUrl =
                _configuration.GetNormalizedBaseUrl();
            string userId =
                NextcloudUserIdentityService.ResolveCurrentUserId(
                    _configuration);
            return _davClient.ReadFilePreview(
                normalizedBaseUrl,
                userId,
                entry.RelativePath,
                maximumBytes,
                cancellationToken);
        }

        internal byte[] TryReadNextcloudGeneratedPreview(
            NextcloudStorageEntry entry,
            int width,
            int height,
            long maximumBytes,
            CancellationToken cancellationToken)
        {
            if (entry == null || entry.IsDirectory)
            {
                throw new ArgumentException(
                    "A Nextcloud file is required.",
                    "entry");
            }

            return _davClient.TryReadGeneratedPreview(
                _configuration.GetNormalizedBaseUrl(),
                entry.RelativePath,
                width,
                height,
                maximumBytes,
                cancellationToken);
        }

        internal IList<NextcloudStorageEntry> SnapshotNextcloudFolder(
            NextcloudStorageEntry root,
            IProgress<int> progress,
            CancellationToken cancellationToken)
        {
            if (root == null || !root.IsDirectory)
            {
                throw new ArgumentException(
                    "A Nextcloud folder is required.",
                    "root");
            }

            var snapshot = new List<NextcloudStorageEntry>();
            var pending = new Queue<NextcloudStorageEntry>();
            pending.Enqueue(root);
            int completedFolders = 0;
            while (pending.Count > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                NextcloudStorageEntry folder = pending.Dequeue();
                NextcloudStorageListing listing =
                    ListNextcloudDirectory(
                        folder.RelativePath,
                        cancellationToken);
                foreach (NextcloudStorageEntry entry
                    in listing.Entries)
                {
                    snapshot.Add(entry);
                    if (entry.IsDirectory)
                    {
                        pending.Enqueue(entry);
                    }
                }
                completedFolders++;
                if (progress != null)
                {
                    progress.Report(completedFolders);
                }
            }

            return snapshot
                .OrderBy(
                    entry => NextcloudPath.GetDepth(
                        entry.RelativePath))
                .ThenBy(entry => entry.IsDirectory ? 0 : 1)
                .ThenBy(
                    entry => entry.RelativePath,
                    StringComparer.Ordinal)
                .ToList();
        }

        internal FileLinkUploadContext PrepareUpload(
            FileLinkRequest request,
            IList<FileLinkSelection> selections,
            IDictionary<FileLinkSelection, FileLinkQueueNode>
                queueSnapshots,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            IProgress<FileLinkUploadPhaseProgress> phaseProgress,
            CancellationToken cancellationToken,
            FileLinkUploadContext existingContext = null)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (selections == null)
            {
                throw new ArgumentNullException("selections");
            }
            if (queueSnapshots == null)
            {
                throw new ArgumentNullException("queueSnapshots");
            }

            cancellationToken.ThrowIfCancellationRequested();
            FileLinkUploadProgress.ReportPhase(
                phaseProgress,
                FileLinkUploadPhase.Scanning,
                0,
                0,
                0,
                0,
                0,
                0);

            using (DiagnosticsLogger.BeginOperation(
                LogCategories.FileLink,
                "FileLink.UploadPrepare"))
            {
                NextcloudCapabilitiesSnapshot capabilities =
                    ResolveRequiredCapabilities();
                string normalizedBaseUrl =
                    _configuration.GetNormalizedBaseUrl();
                string userId =
                    NextcloudUserIdentityService.ResolveCurrentUserId(
                        _configuration);
                FileLinkShareTarget shareTarget =
                    ResolveShareTarget(request);
                FileLinkUploadPlan plan = FileLinkUploadPlanBuilder.Build(
                    selections,
                    queueSnapshots,
                    capabilities.BulkUploadSupported && existingContext == null,
                    existingContext == null
                        ? FileLinkPath.GetDepth(shareTarget.BasePath) + 1
                        : 0,
                    duplicateResolver,
                    cancellationToken,
                    existingContext);
                _transferService.PrepareBulkChecksums(
                    plan,
                    phaseProgress,
                    cancellationToken);
                LogUploadPlan(selections.Count, plan);

                if (existingContext != null)
                {
                    if (!string.Equals(normalizedBaseUrl, existingContext.NormalizedBaseUrl, StringComparison.Ordinal)
                        || !string.Equals(userId, existingContext.UserId, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException("The upload account changed during the share.");
                    }
                    var appendedContext = new FileLinkUploadContext(
                        existingContext.NormalizedBaseUrl,
                        existingContext.UserId,
                        existingContext.SanitizedShareName,
                        existingContext.FolderName,
                        existingContext.RelativeFolderPath,
                        plan,
                        existingContext);
                    _davClient.CreatePlannedDirectories(
                        appendedContext,
                        plan,
                        phaseProgress,
                        cancellationToken);
                    return appendedContext;
                }

                FileLinkUploadContext context = null;
                bool shareRootCreated = false;
                try
                {
                    _davClient.EnsureFolderPath(
                        normalizedBaseUrl,
                        userId,
                        shareTarget.BasePath,
                        cancellationToken);

                    context = CreateShareRoot(
                        request,
                        plan,
                        normalizedBaseUrl,
                        userId,
                        shareTarget,
                        cancellationToken);
                    shareRootCreated = true;

                    _davClient.CreatePlannedDirectories(
                        context,
                        plan,
                        phaseProgress,
                        cancellationToken);
                    return context;
                }
                catch
                {
                    if (shareRootCreated && context != null)
                    {
                        _davClient.DeleteBestEffort(
                            FileLinkDavClient.BuildFileUrl(
                                context.NormalizedBaseUrl,
                                context.UserId,
                                context.RelativeFolderPath),
                            "Prepared upload root cleanup failed");
                    }
                    throw;
                }
            }
        }

        internal void DeleteShareFolder(
            string relativeFolderPath,
            CancellationToken cancellationToken)
        {
            string normalizedPath = FileLinkPath.NormalizeRelativePath(
                relativeFolderPath);
            if (string.IsNullOrWhiteSpace(normalizedPath))
            {
                throw new ArgumentException("relativeFolderPath");
            }

            cancellationToken.ThrowIfCancellationRequested();
            string normalizedBaseUrl =
                _configuration.GetNormalizedBaseUrl();
            string userId =
                NextcloudUserIdentityService.ResolveCurrentUserId(
                    _configuration);
            _davClient.DeleteShareFolder(
                normalizedBaseUrl,
                userId,
                normalizedPath,
                cancellationToken);
        }

        internal void UploadSelections(
            FileLinkUploadContext context,
            IProgress<FileLinkUploadItemProgress> itemProgress,
            IProgress<FileLinkUploadPhaseProgress> phaseProgress,
            CancellationToken cancellationToken)
        {
            _transferService.Upload(
                context,
                itemProgress,
                phaseProgress,
                cancellationToken);
        }

        internal FileLinkResult FinalizeShare(
            FileLinkUploadContext context,
            FileLinkRequest request,
            CancellationToken cancellationToken)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }

            FileLinkShareData shareData = _shareClient.Create(
                context.NormalizedBaseUrl,
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

        private FileLinkUploadContext CreateShareRoot(
            FileLinkRequest request,
            FileLinkUploadPlan plan,
            string normalizedBaseUrl,
            string userId,
            FileLinkShareTarget initialTarget,
            CancellationToken cancellationToken)
        {
            string resolvedShareName = initialTarget.ShareName;
            FileLinkShareTarget resolvedTarget = initialTarget;
            int candidateLimit = request.AttachmentMode
                ? AttachmentShareNameCandidateLimit
                : 1;

            for (int suffix = 0; suffix < candidateLimit; suffix++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                resolvedShareName = suffix == 0
                    ? initialTarget.ShareName
                    : initialTarget.ShareName
                      + "_"
                      + suffix.ToString(CultureInfo.InvariantCulture);
                resolvedTarget = suffix == 0
                    ? initialTarget
                    : FileLinkPath.ResolveShareTarget(
                        initialTarget.BasePath,
                        resolvedShareName,
                        initialTarget.ShareDate,
                        Strings.FileLinkWizardFallbackShareName);
                if (!_davClient.TryCreateShareRoot(
                    normalizedBaseUrl,
                    userId,
                    resolvedTarget.RelativeFolderPath,
                    cancellationToken))
                {
                    continue;
                }

                if (suffix > 0)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.FileLink,
                        "Attachment share root selected after "
                        + suffix.ToString(CultureInfo.InvariantCulture)
                        + " name collision(s).");
                }
                return new FileLinkUploadContext(
                    normalizedBaseUrl,
                    userId,
                    resolvedShareName,
                    resolvedTarget.FolderName,
                    resolvedTarget.RelativeFolderPath,
                    plan);
            }

            throw new TalkServiceException(
                string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.FileLinkWizardFolderExistsFormat,
                    resolvedTarget.FolderName),
                false,
                HttpStatusCode.MethodNotAllowed,
                null);
        }

        private static FileLinkShareTarget ResolveShareTarget(
            FileLinkRequest request)
        {
            DateTime shareDate = request.ShareDate.HasValue
                ? request.ShareDate.Value
                : DateTime.Now;
            return FileLinkPath.ResolveShareTarget(
                request.BasePath,
                request.ShareName,
                shareDate,
                Strings.FileLinkWizardFallbackShareName);
        }

        private NextcloudCapabilitiesSnapshot ResolveRequiredCapabilities()
        {
            if (_capabilitiesSnapshot != null)
            {
                NextcloudCapabilitiesService.RequireSupportedSnapshot(
                    _capabilitiesSnapshot);
                return _capabilitiesSnapshot;
            }

            _capabilitiesSnapshot = new NextcloudCapabilitiesService(
                _configuration).GetRequiredSnapshot(false, false);
            return _capabilitiesSnapshot;
        }

        private static void LogUploadPlan(
            int selectionCount,
            FileLinkUploadPlan plan)
        {
            DiagnosticsLogger.Log(
                LogCategories.FileLink,
                "Upload plan ready (selections="
                + selectionCount.ToString(CultureInfo.InvariantCulture)
                + ", files="
                + plan.Files.Count.ToString(CultureInfo.InvariantCulture)
                + ", foldersToCreate="
                + plan.DirectoriesToCreate.Count.ToString(
                    CultureInfo.InvariantCulture)
                + ", bytes="
                + plan.TotalBytes.ToString(CultureInfo.InvariantCulture)
                + ", direct="
                + plan.DirectFileCount.ToString(
                    CultureInfo.InvariantCulture)
                + ", chunked="
                + plan.ChunkedFileCount.ToString(
                    CultureInfo.InvariantCulture)
                + ", bulkFiles="
                + plan.BulkFiles.Count.ToString(
                    CultureInfo.InvariantCulture)
                + ", bulkBatches="
                + plan.BulkBatches.Count.ToString(
                    CultureInfo.InvariantCulture)
                + ").");
        }
    }
}

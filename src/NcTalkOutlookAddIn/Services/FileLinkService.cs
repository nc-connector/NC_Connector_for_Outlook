// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
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

        internal FileLinkUploadContext PrepareUpload(
            FileLinkRequest request,
            IList<FileLinkSelection> selections,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            IProgress<FileLinkUploadPhaseProgress> phaseProgress,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (selections == null)
            {
                throw new ArgumentNullException("selections");
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
                string basePath = FileLinkPath.NormalizeRelativePath(
                    request.BasePath);
                string sanitizedShareName =
                    FileLinkPath.SanitizeComponent(request.ShareName);
                if (string.IsNullOrWhiteSpace(sanitizedShareName))
                {
                    sanitizedShareName =
                        Strings.FileLinkWizardFallbackShareName;
                }

                DateTime shareDate = request.ShareDate.HasValue
                    ? request.ShareDate.Value
                    : DateTime.Now;
                FileLinkUploadPlan plan = FileLinkUploadPlanBuilder.Build(
                    selections,
                    capabilities.BulkUploadSupported,
                    FileLinkPath.GetDepth(basePath) + 1,
                    duplicateResolver,
                    cancellationToken);
                _transferService.PrepareBulkChecksums(
                    plan,
                    cancellationToken);
                LogUploadPlan(selections.Count, plan);

                FileLinkUploadContext context = null;
                bool shareRootCreated = false;
                try
                {
                    _davClient.EnsureFolderPath(
                        normalizedBaseUrl,
                        userId,
                        basePath,
                        cancellationToken);

                    context = CreateShareRoot(
                        request,
                        plan,
                        normalizedBaseUrl,
                        userId,
                        basePath,
                        sanitizedShareName,
                        shareDate,
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
            string basePath,
            string sanitizedShareName,
            DateTime shareDate,
            CancellationToken cancellationToken)
        {
            string resolvedShareName = sanitizedShareName;
            string resolvedFolderName = string.Empty;
            int candidateLimit = request.AttachmentMode
                ? AttachmentShareNameCandidateLimit
                : 1;

            for (int suffix = 0; suffix < candidateLimit; suffix++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                resolvedShareName = suffix == 0
                    ? sanitizedShareName
                    : sanitizedShareName
                      + "_"
                      + suffix.ToString(CultureInfo.InvariantCulture);
                resolvedFolderName = FileLinkPath.BuildShareFolderName(
                    shareDate,
                    resolvedShareName);
                string resolvedFolderPath = FileLinkPath.Combine(
                    basePath,
                    resolvedFolderName);
                if (!_davClient.TryCreateShareRoot(
                    normalizedBaseUrl,
                    userId,
                    resolvedFolderPath,
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
                    resolvedFolderName,
                    resolvedFolderPath,
                    plan);
            }

            throw new TalkServiceException(
                string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.FileLinkWizardFolderExistsFormat,
                    resolvedFolderName),
                false,
                HttpStatusCode.MethodNotAllowed,
                null);
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

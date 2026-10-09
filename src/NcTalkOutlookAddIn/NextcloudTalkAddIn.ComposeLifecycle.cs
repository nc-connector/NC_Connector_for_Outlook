// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal void DispatchSeparatePasswordMails(
            string composeKey,
            List<SeparatePasswordDispatchEntry> queue)
        {
            _separatePasswordDeliveryController
                .DispatchSeparatePasswordMailQueue(
                    composeKey,
                    queue);
        }

        internal void QueueCreatedShareCleanup(
            string composeKey,
            FileLinkResult result,
            ComposeLifecycleOrigin origin,
            string reason)
        {
            QueueCreatedShareCleanup(
                composeKey,
                new List<ComposeShareCleanupRecord>
                {
                    new ComposeShareCleanupRecord
                    {
                        RelativeFolder = result != null
                            ? result.RelativePath
                            : string.Empty,
                        ShareId = result != null
                            ? result.ShareId
                            : string.Empty,
                        ShareLabel = result != null
                            ? result.FolderName
                            : string.Empty,
                        Origin = origin
                    }
                },
                reason);
        }

        internal void QueueCreatedShareCleanup(
            string composeKey,
            List<ComposeShareCleanupRecord> records,
            string reason)
        {
            var pending = new List<ComposeShareCleanupRecord>();
            if (records != null)
            {
                for (int i = 0; i < records.Count; i++)
                {
                    ComposeShareCleanupRecord record = records[i];
                    if (record == null
                        || string.IsNullOrWhiteSpace(
                            record.RelativeFolder))
                    {
                        continue;
                    }
                    pending.Add(
                        new ComposeShareCleanupRecord
                        {
                            RelativeFolder =
                                record.RelativeFolder.Trim(),
                            ShareId = record.ShareId
                                ?? string.Empty,
                            ShareLabel = record.ShareLabel
                                ?? string.Empty,
                            Origin = record.Origin != null
                                ? record.Origin.Clone()
                                : null
                        });
                }
            }
            if (pending.Count == 0)
            {
                return;
            }

            LogFileLinkMessage(
                "Compose share cleanup queued (composeKey="
                + (composeKey ?? string.Empty)
                + ", count="
                + pending.Count
                + ", reason="
                + (reason ?? string.Empty)
                + ").");
            if (!_composeShareCleanupService.QueueCleanup(pending, reason))
            {
                LogFileLinkMessage("Compose share cleanup could not be accepted by durable storage; remote files are preserved.");
            }
        }

        private void InitializeComposeShareCleanup(string dataDirectory, string profileScope)
        {
            _composeShareCleanupService.Initialize(dataDirectory, profileScope);
        }

        private void DisposeComposeShareCleanup()
        {
            _composeShareCleanupService.Dispose();
        }

        internal void CaptureSeparatePasswordSignatureSnapshot(
            List<SeparatePasswordDispatchEntry> queue,
            BackendPolicyStatus policyStatus,
            AddinSettings settings,
            string composeKey)
        {
            _separatePasswordDeliveryController
                .CaptureSeparatePasswordSignatureSnapshot(
                    queue,
                    policyStatus,
                    settings,
                    composeKey);
        }

        internal bool TryInsertHtmlIntoMail(
            Outlook.MailItem mail,
            string html)
        {
            return _mailBodyInsertionController.InsertHtmlIntoMail(mail, html);
        }

        internal bool TryInsertPlainTextIntoMail(
            Outlook.MailItem mail,
            string plainText)
        {
            return _mailBodyInsertionController.InsertPlainTextIntoMail(
                mail,
                plainText);
        }

    }
}

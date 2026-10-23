// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription
        {
            private const string SharingBlockProperty = "X-NCC-SHARE-INSERTED";
            private bool _sharingWizardOpen;
            private bool _sharingBlockInserted;
            private bool _sharingNoticeShown;
            private bool _sharingHandoffFailed;
            private FileLinkWizardForm _sharingWizard;
            private readonly List<string> _sharingTemporaryFiles = new List<string>();

            internal bool SharingWizardOpen { get { return _sharingWizardOpen; } }

            internal bool HasInsertedSharingBlock()
            {
                if (_sharingBlockInserted) { return true; }
                Outlook.UserProperties properties = null;
                Outlook.UserProperty property = null;
                try
                {
                    properties = _mail.UserProperties;
                    property = properties[SharingBlockProperty];
                    _sharingBlockInserted = property != null && object.Equals((object)property.Value, true);
                    return _sharingBlockInserted;
                }
                finally
                {
                    ComInteropScope.TryRelease(property, LogCategories.FileLink, "Failed to release share marker.");
                    ComInteropScope.TryRelease(properties, LogCategories.FileLink, "Failed to release share properties.");
                }
            }

            internal void SetSharingBlockInserted(bool inserted)
            {
                Outlook.UserProperties properties = null;
                Outlook.UserProperty property = null;
                try
                {
                    properties = _mail.UserProperties;
                    property = properties[SharingBlockProperty];
                    if (property == null)
                    {
                        if (!inserted) { _sharingBlockInserted = false; return; }
                        property = properties.Add(SharingBlockProperty, Outlook.OlUserPropertyType.olYesNo, false, Type.Missing);
                    }
                    property.Value = inserted;
                    _sharingBlockInserted = inserted;
                }
                finally
                {
                    ComInteropScope.TryRelease(property, LogCategories.FileLink, "Failed to release share marker.");
                    ComInteropScope.TryRelease(properties, LogCategories.FileLink, "Failed to release share properties.");
                }
            }

            private void ShowSharingAlreadyInserted()
            {
                if (_sharingNoticeShown) { return; }
                _sharingNoticeShown = true;
                MessageBox.Show(Strings.SharingAlreadyInserted, Strings.DialogTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            internal void FocusSharingWizard()
            {
                if (_sharingWizard == null || _sharingWizard.IsDisposed) { return; }
                if (_sharingWizard.WindowState == FormWindowState.Minimized)
                {
                    _sharingWizard.WindowState = FormWindowState.Normal;
                }
                _sharingWizard.Activate();
            }

            internal bool TryBeginSharingWizard()
            {
                if (!CanApplyComposeChanges) { return false; }
                if (SharingWizardOpen) { FocusSharingWizard(); return false; }
                if (HasInsertedSharingBlock()) { ShowSharingAlreadyInserted(); return false; }
                // Reserve before connection or policy work can yield to another entry.
                _sharingWizardOpen = true;
                _attachmentEvalTimer.Stop();
                return true;
            }

            internal bool BindSharingWizard(FileLinkWizardForm wizard)
            {
                if (!CanApplyComposeChanges || !SharingWizardOpen) { return false; }
                _sharingWizard = wizard;
                wizard.FlushComposeAttachments = DrainSharingAttachments;
                return DrainSharingAttachments();
            }

            internal bool DrainSharingAttachments()
            {
                if (!CanApplyComposeChanges || _sharingWizard == null || _sharingWizard.IsDisposed)
                {
                    return false;
                }
                if (_attachmentSuppressed || _sharingHandoffFailed) { return false; }
                _attachmentEvaluationPending = false;
                _pendingAddedBatch.Clear();
                var selections = new List<FileLinkSelection>();
                var originals = new List<AttachmentShareOriginal>();
                var temporaryFiles = new List<string>();
                var removed = new List<AttachmentShareOriginal>();
                bool adopted = false;
                _attachmentSuppressed = true;
                try
                {
                    long expectedBytes;
                    int expectedCount = CountPolicyRelevantAttachments(out expectedBytes);
                    CollectAttachmentSelectionsForShare(selections, originals, temporaryFiles);
                    if (selections.Count != expectedCount) { return false; }
                    if (selections.Count > 0)
                    {
                        adopted = _sharingWizard.TryAdoptComposeSelections(selections);
                        if (!adopted) { return false; }
                        // Admission and host removal run without yielding the Outlook STA.
                        foreach (AttachmentShareOriginal original in originals)
                        {
                            if (!IsAutomationFileAttachment(original.Attachment))
                            {
                                throw new InvalidOperationException("Attachment classification changed during queue handoff.");
                            }
                            original.Attachment.Delete();
                            removed.Add(original);
                        }
                        _sharingTemporaryFiles.AddRange(temporaryFiles);
                        temporaryFiles.Clear();
                        LogFileLink("Compose attachments adopted by sharing wizard (composeKey=" + _composeKey
                            + ", count=" + selections.Count.ToString(CultureInfo.InvariantCulture) + ").");
                    }
                    long remainingBytes;
                    return CountPolicyRelevantAttachments(out remainingBytes) == 0;
                }
                catch (Exception ex)
                {
                    _sharingHandoffFailed = true;
                    if (adopted) { _sharingWizard.RollbackComposeSelections(selections); }
                    foreach (AttachmentShareOriginal original in removed)
                    {
                        Outlook.Attachments attachments = null;
                        Outlook.Attachment restored = null;
                        try
                        {
                            attachments = _mail.Attachments;
                            restored = attachments.Add(original.LocalPath, Outlook.OlAttachmentType.olByValue,
                                Type.Missing, Type.Missing);
                        }
                        catch (Exception restoreEx)
                        {
                            // Preserve the materialized source if Outlook cannot restore it.
                            temporaryFiles.Remove(original.LocalPath);
                            DiagnosticsLogger.LogException(LogCategories.FileLink,
                                "Attachment rollback failed; retained local source: " + original.LocalPath, restoreEx);
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(restored, LogCategories.FileLink, "Failed to release restored attachment.");
                            ComInteropScope.TryRelease(attachments, LogCategories.FileLink, "Failed to release restore collection.");
                        }
                    }
                    DiagnosticsLogger.LogException(LogCategories.FileLink,
                        "Compose queue handoff failed; retained original attachments (composeKey=" + _composeKey + ").", ex);
                    MessageBox.Show(Strings.FileLinkWizardUploadFailed, Strings.DialogTitle,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                finally
                {
                    ReleaseAttachmentShareOriginals(originals);
                    CleanupTemporaryFiles(temporaryFiles);
                    EndAttachmentSuppression("queue_handoff");
                }
            }

            internal void EndSharingWizard()
            {
                _sharingWizard = null;
                _sharingWizardOpen = false;
                _sharingHandoffFailed = false;
                CleanupTemporaryFiles(_sharingTemporaryFiles);
                _sharingTemporaryFiles.Clear();
                _pendingAddedBatch.Clear();
                _attachmentEvaluationPending = false;
                ResumeDeferredEmailSignatureApplication("sharing_closed");
            }

            private Task StartComposeAttachmentShareFlowAsync(string trigger, long totalBytes, int thresholdMb, AttachmentBatchInfo lastAdded)
            {
                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState("start_flow", _composeKey, out guardState))
                {
                    return Task.CompletedTask;
                }
                return _owner.RunFileLinkWizardForMailAsync(_mail, new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = string.IsNullOrWhiteSpace(trigger) ? "always" : trigger,
                    AttachmentTotalBytes = Math.Max(0, totalBytes),
                    AttachmentThresholdMb = Math.Max(1, thresholdMb),
                    AttachmentLastName = lastAdded != null ? lastAdded.Name : string.Empty,
                    AttachmentLastSizeBytes = lastAdded != null ? Math.Max(0, lastAdded.SizeBytes) : 0
                });
            }

            private void RunAttachmentFlowTask(Task task, string failureMessage)
            {
                if (task == null)
                {
                    return;
                }
                task.ContinueWith(
                    failedTask => DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        (failureMessage ?? "Compose attachment flow failed")
                        + " (composeKey="
                        + _composeKey
                        + ").",
                        failedTask.Exception),
                    TaskContinuationOptions.OnlyOnFaulted);
            }

            private void EndAttachmentSuppression(string trigger)
            {
                _attachmentSuppressed = false;
                if (_attachmentEvaluationPending) { ScheduleAttachmentEvaluation(); }
                ResumeDeferredEmailSignatureApplication("attachment_" + (trigger ?? string.Empty));
            }
        }
    }
}

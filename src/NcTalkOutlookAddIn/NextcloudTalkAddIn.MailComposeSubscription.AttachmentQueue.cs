// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription
        {
            private sealed class BeforeAddShareEntry
            {
                internal AttachmentBatchEntry Candidate { get; set; }

                internal string LocalPath { get; set; }

                internal int ThresholdMb { get; set; }

                internal bool CleanupLocalPathAfterFlow { get; set; }

                internal string Trigger { get; set; }

            }

            private async Task StartComposeAttachmentShareFlowAsync(string trigger, long totalBytes, int thresholdMb, AttachmentBatchInfo lastAdded)
            {
                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_beforeAddShareFlowRunning
                    || _owner.TryGetAttachmentAutomationGuardState("start_flow", _composeKey, out guardState))
                {
                    return;
                }
                var originals = new List<AttachmentShareOriginal>();
                var tempFiles = new List<string>();
                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = string.IsNullOrWhiteSpace(trigger) ? "always" : trigger,
                    AttachmentTotalBytes = Math.Max(0, totalBytes),
                    AttachmentThresholdMb = Math.Max(1, thresholdMb),
                    AttachmentLastName = lastAdded != null ? (lastAdded.Name ?? string.Empty) : string.Empty,
                    AttachmentLastSizeBytes = lastAdded != null ? Math.Max(0, lastAdded.SizeBytes) : 0
                };
                launchOptions.PrepareInitialSelections = () =>
                    PrepareComposeAttachmentSelections(launchOptions, originals, tempFiles);

                _beforeAddShareFlowRunning = true;
                _attachmentSuppressed = true;
                bool accepted = false;
                ExceptionDispatchInfo failure = null;
                try
                {
                    accepted = await _owner.RunFileLinkWizardForMailAsync(_mail, launchOptions);
                    LogFileLink(
                        "Compose attachment flow completed (composeKey=" + _composeKey
                        + ", trigger=" + launchOptions.AttachmentTrigger
                        + ", wizardAccepted=" + accepted.ToString(CultureInfo.InvariantCulture) + ").");
                }
                catch (Exception ex)
                {
                    failure = ExceptionDispatchInfo.Capture(ex);
                    launchOptions.UnexpectedFailureObserved = !(ex is OperationCanceledException);
                }
                try
                {
                    await FinalizeAttachmentShareFlowAsync(originals, launchOptions, accepted);
                }
                catch (Exception ex)
                {
                    if (failure == null) { failure = ExceptionDispatchInfo.Capture(ex); }
                }
                finally
                {
                    CleanupTemporaryFiles(tempFiles);
                }
                if (failure != null) { failure.Throw(); }
            }

            private bool PrepareComposeAttachmentSelections(
                FileLinkWizardLaunchOptions launchOptions,
                List<AttachmentShareOriginal> originals,
                List<string> tempFiles)
            {
                if (!CanApplyComposeChanges || _mail == null)
                {
                    return false;
                }
                // Keep owned attachment identities through the wizard, never stale collection positions.
                var selections = new List<FileLinkSelection>();
                CollectAttachmentSelectionsForShare(selections, originals, tempFiles);
                if (selections.Count == 0)
                {
                    LogFileLink("Compose attachment flow skipped (composeKey=" + _composeKey + ", reason=no_collectible_files).");
                    return false;
                }
                foreach (FileLinkSelection selection in selections)
                {
                    launchOptions.InitialSelections.Add(selection);
                }
                return true;
            }

            private void QueueBeforeAddAttachmentShareFlow(
                string trigger,
                AttachmentBatchEntry candidate,
                string localPath,
                int thresholdMb,
                bool cleanupLocalPathAfterFlow)
            {
                if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                {
                    LogFileLink("Compose before-attachment-add share queue skipped (composeKey=" + _composeKey + ", reason=missing_local_path).");
                    return;
                }

                _pendingBeforeAddShareEntries.Add(new BeforeAddShareEntry
                {
                    Candidate = new AttachmentBatchEntry
                    {
                        OriginalAttachment = candidate != null ? candidate.OriginalAttachment : null,
                        Name = candidate != null ? (candidate.Name ?? string.Empty) : string.Empty,
                        SizeBytes = Math.Max(0, candidate != null ? candidate.SizeBytes : 0)
                    },
                    LocalPath = localPath,
                    ThresholdMb = Math.Max(1, thresholdMb),
                    CleanupLocalPathAfterFlow = cleanupLocalPathAfterFlow,
                    Trigger = string.IsNullOrWhiteSpace(trigger) ? "always_preadd" : trigger
                });

                LogFileLink(
                    "Compose before-attachment-add candidate queued (composeKey="
                    + _composeKey
                    + ", queued="
                    + _pendingBeforeAddShareEntries.Count.ToString(CultureInfo.InvariantCulture)
                    + ", attachment="
                    + (candidate != null ? (candidate.Name ?? string.Empty) : string.Empty)
                    + ").");

                if (_beforeAddShareFlowRunning)
                {
                    return;
                }
                try
                {
                    _beforeAddShareTimer.Stop();
                    _beforeAddShareTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to schedule queued before-attachment-add share flow (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private async Task RunQueuedBeforeAddAttachmentShareFlowAsync()
            {
                if (_disposed || _beforeAddShareFlowRunning || _pendingBeforeAddShareEntries.Count == 0)
                {
                    return;
                }

                _beforeAddShareFlowRunning = true;
                var batch = new List<BeforeAddShareEntry>(_pendingBeforeAddShareEntries);
                _pendingBeforeAddShareEntries.Clear();
                var originals = new List<AttachmentShareOriginal>();
                var temporaryFiles = new List<string>();
                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = batch.TrueForAll(entry => entry != null
                        && string.Equals(entry.Trigger, "threshold_preadd", StringComparison.OrdinalIgnoreCase))
                        ? "threshold" : "always"
                };
                launchOptions.PrepareInitialSelections = () =>
                {
                    if (!CanApplyComposeChanges) { return false; }
                    var selections = new List<FileLinkSelection>();
                    foreach (BeforeAddShareEntry entry in batch)
                    {
                        CaptureBeforeAddAttachmentOriginal(entry.Candidate, selections, originals, temporaryFiles);
                    }
                    foreach (FileLinkSelection selection in selections) { launchOptions.InitialSelections.Add(selection); }
                    return selections.Count > 0;
                };

                bool accepted = false;
                ExceptionDispatchInfo failure = null;
                try
                {
                    foreach (BeforeAddShareEntry entry in batch)
                    {
                        if (entry == null) { continue; }
                        if (entry.CleanupLocalPathAfterFlow) { temporaryFiles.Add(entry.LocalPath); }
                        launchOptions.AttachmentTotalBytes += Math.Max(0, entry.Candidate != null ? entry.Candidate.SizeBytes : 0);
                        launchOptions.AttachmentThresholdMb = Math.Max(launchOptions.AttachmentThresholdMb, entry.ThresholdMb);
                        launchOptions.AttachmentLastName = entry.Candidate != null ? entry.Candidate.Name : string.Empty;
                        launchOptions.AttachmentLastSizeBytes = Math.Max(0, entry.Candidate != null ? entry.Candidate.SizeBytes : 0);
                    }
                    if (batch.Count > 0)
                    {
                        _attachmentSuppressed = true;
                        _pendingAddedBatch.Clear();
                        accepted = await _owner.RunFileLinkWizardForMailAsync(_mail, launchOptions);
                        LogFileLink(
                            "Compose queued attachment flow completed (composeKey=" + _composeKey
                            + ", queued=" + batch.Count.ToString(CultureInfo.InvariantCulture)
                            + ", wizardAccepted=" + accepted.ToString(CultureInfo.InvariantCulture) + ").");
                    }
                }
                catch (Exception ex)
                {
                    failure = ExceptionDispatchInfo.Capture(ex);
                    launchOptions.UnexpectedFailureObserved = !(ex is OperationCanceledException);
                }
                try
                {
                    await FinalizeAttachmentShareFlowAsync(originals, launchOptions, accepted);
                }
                catch (Exception ex)
                {
                    if (failure == null) { failure = ExceptionDispatchInfo.Capture(ex); }
                }
                finally
                {
                    CleanupTemporaryFiles(temporaryFiles);
                }
                if (failure != null) { failure.Throw(); }
            }

            private async Task FinalizeAttachmentShareFlowAsync(
                List<AttachmentShareOriginal> originals,
                FileLinkWizardLaunchOptions launchOptions,
                bool accepted)
            {
                bool stateReleased = false;
                try
                {
                    await _owner.RunOnOutlookUiThreadAsync(() =>
                    {
                        try
                        {
                            if (accepted) { RemoveSharedAttachmentOriginals(originals, launchOptions.SharedLocalPaths); }
                        }
                        finally
                        {
                            try { ReleaseAttachmentShareOriginals(originals); }
                            finally
                            {
                                _beforeAddShareFlowRunning = false;
                                stateReleased = true;
                                EndAttachmentSuppression("share_flow_complete");
                                if (!accepted && launchOptions.UnexpectedFailureObserved)
                                {
                                    AddinSettings current = _owner._currentSettings;
                                    if (current != null)
                                    {
                                        var configuration = new TalkServiceConfiguration(current.ServerUrl, current.Username, current.AppPassword);
                                        _owner.InvalidateCurrentBackendPolicyCheck(configuration);
                                        BeginAttachmentAutomationSettingsRefresh();
                                        ScheduleEmailSignatureApplication("attachment_failure");
                                    }
                                }
                                RestartBeforeAddShareTimerIfNeeded();
                            }
                        }
                        return true;
                    });
                }
                finally
                {
                    if (!stateReleased)
                    {
                        _attachmentSuppressed = false;
                        _beforeAddShareFlowRunning = false;
                    }
                }
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

            private void RestartBeforeAddShareTimerIfNeeded()
            {
                if (_pendingBeforeAddShareEntries.Count == 0 || _disposed)
                {
                    return;
                }
                try
                {
                    _beforeAddShareTimer.Stop();
                    _beforeAddShareTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to restart queued before-attachment-add share flow (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void EndAttachmentSuppression(string trigger)
            {
                _attachmentSuppressed = false;
                ResumeDeferredEmailSignatureApplication("attachment_" + (trigger ?? string.Empty));
            }

        }
    }
}

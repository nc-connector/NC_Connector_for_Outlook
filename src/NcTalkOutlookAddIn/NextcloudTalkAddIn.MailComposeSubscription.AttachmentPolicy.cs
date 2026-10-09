// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
            internal bool CanApplyComposeChanges
            {
                get { return !_disposed && !_sendAccepted; }
            }

            private AttachmentAutomationSettings _attachmentAutomationSettingsSnapshot;
            private DateTime _attachmentAutomationSettingsSnapshotUtc;
            private Task<AttachmentAutomationSettings> _attachmentAutomationSettingsRefreshTask;
            private int _attachmentAutomationSettingsRefreshGeneration;
            private static readonly TimeSpan AttachmentAutomationSettingsCacheLifetime =
                TimeSpan.FromMinutes(5);

            private sealed class AttachmentAutomationSettings
            {
                internal AddinSettings LocalSettings { get; set; }

                internal bool AlwaysConnector { get; set; }

                internal bool OfferAboveEnabled { get; set; }

                internal bool ThresholdMandatory { get; set; }

                internal int ThresholdMb { get; set; }

                internal long ThresholdBytes { get; set; }

                internal bool EnterpriseRolloutBlocked { get; set; }

                internal bool BackendRequired { get; set; }
            }

            private AttachmentAutomationSettings ReadAttachmentAutomationSettings()
            {
                AttachmentAutomationSettings local = ReadLocalAttachmentAutomationSettings();
                AddinSettings current = _owner._currentSettings;
                BackendPolicyStatus confirmed;
                AttachmentAutomationSettings effective = local;
                if (!HasFreshAttachmentAutomationSettingsSnapshot())
                {
                    BeginAttachmentAutomationSettingsRefresh();
                }
                if (current != null)
                {
                    var configuration = new TalkServiceConfiguration(current.ServerUrl, current.Username, current.AppPassword);
                    if (_owner.TryGetCachedEmailSignaturePolicyStatus(configuration, out confirmed))
                    {
                        effective = ApplyAttachmentAutomationPolicy(local, confirmed);
                    }
                }
                return effective;
            }

            private async Task<AttachmentAutomationSettings> ReadAttachmentAutomationSettingsAsync()
            {
                if (_attachmentAutomationSettingsSnapshot != null)
                {
                    return ReadAttachmentAutomationSettings();
                }

                while (!_disposed)
                {
                    BeginAttachmentAutomationSettingsRefresh();
                    Task<AttachmentAutomationSettings> refreshTask =
                        _attachmentAutomationSettingsRefreshTask;
                    if (refreshTask == null)
                    {
                        return ReadLocalAttachmentAutomationSettings();
                    }

                    AttachmentAutomationSettings refreshed =
                        await refreshTask;
                    if (ReferenceEquals(
                        refreshTask,
                        _attachmentAutomationSettingsRefreshTask))
                    {
                        return _attachmentAutomationSettingsSnapshot
                               ?? refreshed;
                    }
                }
                return ReadLocalAttachmentAutomationSettings();
            }

            private bool HasFreshAttachmentAutomationSettingsSnapshot()
            {
                return _attachmentAutomationSettingsSnapshot != null
                       && DateTime.UtcNow
                       - _attachmentAutomationSettingsSnapshotUtc
                       < AttachmentAutomationSettingsCacheLifetime;
            }

            internal void RefreshAttachmentAutomationSettings()
            {
                if (_disposed)
                {
                    return;
                }

                _attachmentAutomationSettingsRefreshGeneration++;
                _attachmentAutomationSettingsSnapshot = null;
                _attachmentAutomationSettingsSnapshotUtc =
                    DateTime.MinValue;
                _attachmentAutomationSettingsRefreshTask = null;
                BeginAttachmentAutomationSettingsRefresh();
                LogFileLink(
                    "Compose attachment settings refresh requested (composeKey="
                    + _composeKey
                    + ").");
            }

            private void BeginAttachmentAutomationSettingsRefresh()
            {
                if (_disposed
                    || (_attachmentAutomationSettingsRefreshTask != null
                        && !_attachmentAutomationSettingsRefreshTask.IsCompleted))
                {
                    return;
                }

                AttachmentAutomationSettings local =
                    ReadLocalAttachmentAutomationSettings();
                AddinSettings current = _owner._currentSettings;
                if (current == null)
                {
                    _attachmentAutomationSettingsSnapshot = local;
                    _attachmentAutomationSettingsSnapshotUtc = DateTime.UtcNow;
                    return;
                }

                var configuration = new TalkServiceConfiguration(
                    current.ServerUrl,
                    current.Username,
                    current.AppPassword);
                int refreshGeneration =
                    _attachmentAutomationSettingsRefreshGeneration;
                _attachmentAutomationSettingsRefreshTask = RefreshAttachmentAutomationSettingsAsync(
                    local,
                    configuration,
                    refreshGeneration);
                RunAttachmentFlowTask(
                    _attachmentAutomationSettingsRefreshTask,
                    "Compose attachment policy refresh failed");
            }

            private async Task<AttachmentAutomationSettings> RefreshAttachmentAutomationSettingsAsync(
                AttachmentAutomationSettings local,
                TalkServiceConfiguration configuration,
                int refreshGeneration)
            {
                AttachmentAutomationSettings resolved = local;
                bool checkSucceeded = configuration == null || !configuration.IsComplete();
                if (configuration != null && configuration.IsComplete())
                {
                    BackendPolicyStatus policyStatus = await _owner.GetEmailSignaturePolicyStatusAsync(
                        configuration, "compose_attachment_evaluate").ConfigureAwait(false);
                    BackendPolicyStatus currentCheck;
                    checkSucceeded = _owner.TryGetCurrentBackendPolicyCheck(configuration, out currentCheck)
                        && currentCheck.FetchSucceeded;
                    if (!checkSucceeded
                        && _attachmentAutomationSettingsSnapshot != null)
                    {
                        LogFileLink(
                            "Compose attachment policy refresh unavailable; retaining the previous rules (composeKey="
                            + _composeKey
                            + ").");
                        return _attachmentAutomationSettingsSnapshot ?? local;
                    }
                    resolved = ApplyAttachmentAutomationPolicy(
                        local,
                        policyStatus);
                }

                if (!_disposed
                    && refreshGeneration
                    == _attachmentAutomationSettingsRefreshGeneration)
                {
                    _attachmentAutomationSettingsSnapshot = resolved;
                    if (checkSucceeded)
                    {
                        _attachmentAutomationSettingsSnapshotUtc = DateTime.UtcNow;
                    }
                }
                return resolved;
            }

            private AttachmentAutomationSettings ReadLocalAttachmentAutomationSettings()
            {
                _owner.EnsureSettingsLoaded();

                AddinSettings settings = (_owner._currentSettings ?? new AddinSettings()).Clone();
                return ApplyAttachmentAutomationPolicy(
                    new AttachmentAutomationSettings { LocalSettings = settings },
                    null);
            }

            private static AttachmentAutomationSettings ApplyAttachmentAutomationPolicy(
                AttachmentAutomationSettings local,
                BackendPolicyStatus policyStatus)
            {
                if (policyStatus != null && !policyStatus.FetchSucceeded)
                {
                    policyStatus = null;
                }
                AddinSettings settings = local != null && local.LocalSettings != null
                    ? local.LocalSettings : new AddinSettings();
                AttachmentAutomationSettings resolved = BuildAttachmentAutomationSettings(
                    settings.ResolvePolicyDefaults(policyStatus), settings);
                resolved.BackendRequired = settings.IsEnterpriseRollout
                    || (policyStatus != null && policyStatus.IsDomainActive("share"));
                resolved.ThresholdMandatory = resolved.OfferAboveEnabled
                    && policyStatus != null
                    && policyStatus.IsLocked("share", "attachments_min_size_mb")
                    && policyStatus.HasPolicyKey("share", "attachments_min_size_mb");
                resolved.EnterpriseRolloutBlocked = policyStatus != null
                    && policyStatus.FetchSucceeded
                    && !string.IsNullOrEmpty(PolicyUiHelper.GetEnterpriseRolloutNotice(settings, policyStatus));
                if (resolved.EnterpriseRolloutBlocked)
                {
                    resolved.AlwaysConnector = false;
                    resolved.OfferAboveEnabled = false;
                    resolved.ThresholdMandatory = false;
                }
                return resolved;
            }

            private static AttachmentAutomationSettings BuildAttachmentAutomationSettings(
                AddinSettings effective,
                AddinSettings local)
            {
                int thresholdMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(effective.SharingAttachmentsOfferAboveMb);
                bool alwaysConnector = effective.SharingAttachmentsAlwaysConnector;
                return new AttachmentAutomationSettings
                {
                    LocalSettings = local,
                    AlwaysConnector = alwaysConnector,
                    OfferAboveEnabled = effective.SharingAttachmentsOfferAboveEnabled && !alwaysConnector,
                    ThresholdMb = thresholdMb,
                    ThresholdBytes = (long)thresholdMb * 1024L * 1024L
                };
            }

            private bool TryValidateAttachmentPolicyBeforeSend(ref bool cancel)
            {
                if (_disposed || cancel)
                {
                    return !cancel;
                }

                long totalBytes;
                int attachmentCount = CountPolicyRelevantAttachments(out totalBytes);
                if (attachmentCount <= 0)
                {
                    return true;
                }

                AttachmentAutomationSettings settings = ReadAttachmentAutomationSettings();
                if (settings.EnterpriseRolloutBlocked
                    || (!settings.AlwaysConnector
                        && !(settings.ThresholdMandatory && totalBytes > settings.ThresholdBytes)))
                {
                    return true;
                }

                AddinSettings current = _owner._currentSettings;
                var configuration = new TalkServiceConfiguration(
                    current != null ? current.ServerUrl : string.Empty,
                    current != null ? current.Username : string.Empty,
                    current != null ? current.AppPassword : string.Empty);
                BackendPolicyStatus check;
                bool checkCurrent = _owner.TryGetCurrentBackendPolicyCheck(configuration, out check);
                if (!checkCurrent)
                {
                    BeginAttachmentAutomationSettingsRefresh();
                    if (current != null && current.SendPolicyFailClosed)
                    {
                        return BlockSendPolicyFailure(ref cancel, null, true);
                    }
                }
                if (checkCurrent && !check.FetchSucceeded
                    && (!check.IsEndpointMissing || settings.BackendRequired))
                {
                    if (check.IsServiceUnavailable && (current == null || !current.SendPolicyFailClosed))
                    {
                        RecordSendPolicyWarning(false, true, check);
                        return true;
                    }
                    return BlockSendPolicyFailure(ref cancel, check, false);
                }

                cancel = true;
                ShowRequiredAttachmentRoutingNotice();
                LogFileLink(
                    "Compose send blocked by required attachment routing (composeKey="
                    + _composeKey
                    + ", remainingAttachments="
                    + attachmentCount.ToString(CultureInfo.InvariantCulture)
                    + ", thresholdRequired="
                    + settings.ThresholdMandatory.ToString(CultureInfo.InvariantCulture)
                    + ").");
                return false;
            }

            private static void ShowRequiredAttachmentRoutingNotice()
            {
                MessageBox.Show(
                    Strings.AttachmentRoutingRequired,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            private bool PauseUnavailableAttachmentAutomation(AttachmentAutomationSettings settings, long totalBytes)
            {
                AddinSettings current = _owner._currentSettings;
                if (current == null)
                {
                    return false;
                }
                var configuration = new TalkServiceConfiguration(current.ServerUrl, current.Username, current.AppPassword);
                BackendPolicyStatus check;
                if (!_owner.TryGetCurrentBackendPolicyCheck(configuration, out check) || check.FetchSucceeded
                    || (check.IsEndpointMissing && !settings.BackendRequired))
                {
                    return false;
                }
                bool required = settings.AlwaysConnector
                    || (settings.ThresholdMandatory && totalBytes > settings.ThresholdBytes);
                if (required && check.IsServiceUnavailable && !current.SendPolicyFailClosed)
                {
                    RecordSendPolicyWarning(false, true, check);
                    ShowSendPolicyWarning();
                }
                LogFileLink("Compose automatic sharing paused after the current connection check failed (composeKey=" + _composeKey + ").");
                return true;
            }

        }
    }
}

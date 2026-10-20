// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Controllers
{
    // Shares action-entry state and ribbon refreshes without moving feature execution into callbacks.
    internal sealed class RibbonActionController
    {
        private readonly NextcloudTalkAddIn _owner;
        private readonly CalendarResourceController _resources;
        private readonly List<IRibbonUI> _ribbons = new List<IRibbonUI>();
        private Timer _refreshTimer;
        private Task _accessRefresh;
        private bool _started;
        private int _generation;

        internal RibbonActionController(NextcloudTalkAddIn owner, CalendarResourceController resources)
        {
            _owner = owner;
            _resources = resources;
        }

        internal sealed class AccessState
        {
            internal bool CanStart;
            internal string Notice;
        }

        internal static AccessState EvaluateAccess(AddinSettings settings, BackendPolicyStatus status, bool requiresSeat, string pauseReason)
        {
            if (settings == null) { return new AccessState { CanStart = false, Notice = Strings.ErrorMissingCredentials }; }
            if (!settings.IsManagedTransportTlsValid) { return new AccessState { CanStart = false, Notice = Strings.ManagedTlsPolicyInvalid }; }
            if (pauseReason == "rate_limited") { return new AccessState { CanStart = false, Notice = Strings.ConnectionRateLimited }; }
            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            // Enabling sign-in or a fresh check is not permission to open a wizard or mutate an item.
            if (pauseReason == "auth_required" || !configuration.IsComplete())
            {
                return new AccessState { CanStart = true, Notice = pauseReason == "auth_required" ? Strings.ConnectionAuthRequired : Strings.ErrorMissingCredentials };
            }
            string notice = PolicyUiHelper.GetActionUnavailableNotice(settings, status, requiresSeat);
            // A click can explain a refusal; execution still checks the current entitlement.
            return new AccessState { CanStart = true, Notice = notice };
        }

        internal void Start()
        {
            if (_started) { return; }
            _started = true;
            _generation++;
            NextcloudConnectionState.StateChanged += OnConnectionStateChanged;
            StartRefreshTimer();
        }

        internal void Register(IRibbonUI ribbon)
        {
            if (ribbon == null) { return; }
            if (!_ribbons.Contains(ribbon)) { _ribbons.Add(ribbon); }
            StartRefreshTimer();
        }

        private void StartRefreshTimer()
        {
            if (_started && _ribbons.Count != 0 && _refreshTimer == null)
            {
                // Outlook caches disabled callbacks. Re-evaluate them; HTTP still uses the shared TTL.
                _refreshTimer = new Timer { Interval = 15000 };
                _refreshTimer.Tick += OnRefreshTick;
                _refreshTimer.Start();
            }
        }

        private void OnRefreshTick(object sender, EventArgs args) { Invalidate(); }

        internal void Invalidate()
        {
            for (int i = _ribbons.Count - 1; i >= 0; i--)
            {
                try { _ribbons[i].Invalidate(); }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Closed ribbon removed from action refresh.", ex);
                    _ribbons.RemoveAt(i);
                }
            }
        }

        internal void Stop()
        {
            _started = false;
            _generation++;
            NextcloudConnectionState.StateChanged -= OnConnectionStateChanged;
            if (_refreshTimer != null)
            {
                _refreshTimer.Stop();
                _refreshTimer.Tick -= OnRefreshTick;
                _refreshTimer.Dispose();
                _refreshTimer = null;
            }
            _ribbons.Clear();
            _accessRefresh = null;
        }

        private async void OnConnectionStateChanged(ConnectionPauseStatus pause)
        {
            int generation = _generation;
            try
            {
                await _owner.RunOnOutlookUiThreadAsync(() =>
                {
                    if (!_started || generation != _generation) { return; }
                    AddinSettings settings = _owner.CurrentSettings;
                    if (settings != null)
                    {
                        _owner.InvalidateCurrentBackendPolicyCheck(new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword));
                    }
                    Invalidate();
                }).ConfigureAwait(false);
            }
            catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Core, "Ribbon connection-state refresh failed.", ex); }
        }

        private AccessState GetAccess(bool requiresSeat)
        {
            AddinSettings settings = _owner.CurrentSettings;
            if (settings == null) { return EvaluateAccess(null, null, requiresSeat, string.Empty); }
            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            ConnectionPauseStatus pause = NextcloudConnectionState.GetStatus(configuration);
            BackendPolicyStatus status = null;
            if (settings.IsManagedTransportTlsValid && configuration.IsComplete() && !pause.IsPaused
                && (requiresSeat || settings.IsEnterpriseRollout))
            {
                BackendPolicyStatus current;
                if (!_owner.TryGetCurrentBackendPolicyCheck(configuration, out current) && (_accessRefresh == null || _accessRefresh.IsCompleted))
                { _accessRefresh = RefreshAccessAsync(configuration, _generation); }
                if (!_owner.TryGetCachedEmailSignaturePolicyStatus(configuration, out status)) { status = current; }
            }
            return EvaluateAccess(settings, status, requiresSeat, pause.Reason);
        }

        private async Task RefreshAccessAsync(TalkServiceConfiguration configuration, int generation)
        {
            try
            {
                await _owner.GetEmailSignaturePolicyStatusAsync(configuration, "ribbon_access").ConfigureAwait(false);
                await _owner.RunOnOutlookUiThreadAsync(() =>
                {
                    if (_started && generation == _generation) { Invalidate(); }
                }).ConfigureAwait(false);
            }
            catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Core, "Ribbon access refresh failed.", ex); }
        }

        private static bool IsResource(IRibbonControl control)
        {
            return control != null && (control.Id == "NcResourceAppointmentButton" || control.Id == "NcResourceSchedulingButton");
        }

        internal bool IsEnabled(IRibbonControl control)
        {
            if (control == null) { return false; }
            bool resource = IsResource(control);
            return GetAccess(resource).CanStart && (!resource || _resources.IsContextAvailable(control));
        }

        internal string GetTooltip(IRibbonControl control)
        {
            bool resource = IsResource(control);
            string notice = GetAccess(resource).Notice;
            if (notice.Length != 0) { return notice; }
            if (resource) { return _resources.GetContextTooltip(control); }
            return control != null && control.Id == "NcTalkCreateButton" ? Strings.RibbonTalkButtonSuperTip : Strings.RibbonFileLinkButtonSuperTip;
        }
    }
}

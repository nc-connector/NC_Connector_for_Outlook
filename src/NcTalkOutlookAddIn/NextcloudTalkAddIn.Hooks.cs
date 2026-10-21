// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
        // Outlook hook lifecycle for inspector/application events and compose-tracking entry points.
    public sealed partial class NextcloudTalkAddIn
    {
        private Outlook.ApplicationEvents_11_Event _profileLogonEvents;

        private void EnsureInspectorHook()
        {
            if (_profileRuntimeStopped || _outlookApplication == null || _inspectors != null)
            {
                return;
            }
            try
            {
                _inspectors = _outlookApplication.Inspectors;
                if (_inspectors != null)
                {
                    _inspectors.NewInspector += OnNewInspector;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Inspectors.NewInspector.", ex);
                _inspectors = null;
            }
        }

        private void EnsureApplicationHook()
        {
            if (_profileRuntimeStopped || _outlookApplication == null)
            {
                return;
            }
            try
            {
                if (_profileLogonEvents == null)
                {
                    _profileLogonEvents = _outlookApplication as Outlook.ApplicationEvents_11_Event;
                    if (_profileLogonEvents != null)
                    {
                        _profileLogonEvents.MAPILogonComplete += OnMapiLogonComplete;
                    }
                }
                EnsureExplorerHooks();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Explorer lifecycle.", ex);
            }
        }

        private void UnhookApplication()
        {
            if (_profileLogonEvents != null)
            {
                try { _profileLogonEvents.MAPILogonComplete -= OnMapiLogonComplete; }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Outlook profile logon completion.", ex);
                }
                _profileLogonEvents = null;
            }
            UnhookExplorerHooks();
        }

        private void OnMapiLogonComplete()
        {
            try { TryInitializeProfileRuntime(); }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to initialize the selected Outlook profile.", ex);
            }
        }

        private void RefreshProfileRuntimeSurfaces()
        {
            // A Simple MAPI Inspector may already exist when logon finally completes.
            try
            {
                int count = _inspectors != null ? _inspectors.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.Inspector inspector = null;
                    bool retained = false;
                    try
                    {
                        inspector = _inspectors[index];
                        retained = BindInspectorItem(inspector);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to bind an Inspector after profile initialization.", ex);
                    }
                    finally
                    {
                        if (!retained)
                        {
                            ComInteropScope.TryRelease(inspector, LogCategories.Core, "Failed to release an untracked existing Inspector.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to enumerate Inspectors after profile initialization.", ex);
            }
            foreach (string explorerKey in new System.Collections.Generic.List<string>(_hookedExplorers.Keys))
            {
                OnExplorerSelectionChanged(explorerKey);
            }
            _mailComposeSubscriptionRegistry.RefreshAttachmentAutomationSettings();
            InvalidateActionRibbons();
        }

        private void EnsureExplorerHooks()
        {
            if (_outlookApplication == null)
            {
                return;
            }
            try
            {
                if (_explorers == null)
                {
                    _explorers = _outlookApplication.Explorers;
                    _explorersEvents = _explorers as Outlook.ExplorersEvents_Event;
                    if (_explorersEvents != null)
                    {
                        _explorersEvents.NewExplorer += OnNewExplorer;
                    }

                    int count = _explorers != null ? _explorers.Count : 0;
                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.Explorer explorer = null;
                        try
                        {
                            explorer = _explorers[i];
                            if (!HookExplorer(explorer))
                            {
                                ComInteropScope.TryRelease(
                                    explorer,
                                    LogCategories.Core,
                                    "Failed to release an untracked existing Explorer.");
                            }
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook existing Explorer lifecycle.", ex);
                            ComInteropScope.TryRelease(explorer, LogCategories.Core, "Failed to release Explorer after hook failure.");
                        }
                    }
                }

                Outlook.Explorer activeExplorer = null;
                try
                {
                    activeExplorer = _outlookApplication.ActiveExplorer();
                    if (!HookExplorer(activeExplorer))
                    {
                        ComInteropScope.TryRelease(
                            activeExplorer,
                            LogCategories.Core,
                            "Failed to release an untracked ActiveExplorer.");
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook active Explorer lifecycle.", ex);
                    ComInteropScope.TryRelease(activeExplorer, LogCategories.Core, "Failed to release ActiveExplorer after hook failure.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to initialize Explorer hooks.", ex);
            }
        }

        private bool HookExplorer(Outlook.Explorer explorer)
        {
            if (_profileRuntimeStopped || explorer == null)
            {
                return false;
            }
            string explorerKey = ComInteropScope.ResolveIdentityKey(explorer, LogCategories.Core, "Explorer");
            if (string.IsNullOrWhiteSpace(explorerKey) || _hookedExplorerEvents.ContainsKey(explorerKey))
            {
                return false;
            }

            var explorerEvents = explorer as Outlook.ExplorerEvents_10_Event;
            if (explorerEvents == null)
            {
                return false;
            }

            Outlook.ExplorerEvents_10_InlineResponseEventHandler inlineResponseHandler =
                item => OnExplorerInlineResponse(explorerKey, item);
            Outlook.ExplorerEvents_10_InlineResponseCloseEventHandler inlineResponseCloseHandler =
                () => OnExplorerInlineResponseClose(explorerKey);
            Outlook.ExplorerEvents_10_SelectionChangeEventHandler selectionChangeHandler =
                () => OnExplorerSelectionChanged(explorerKey);

            explorerEvents.InlineResponse += inlineResponseHandler;
            try
            {
                explorerEvents.InlineResponseClose += inlineResponseCloseHandler;
            }
            catch (Exception ex)
            {
                try
                {
                    explorerEvents.InlineResponse -= inlineResponseHandler;
                }
                catch (Exception rollbackEx)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to roll back Explorer.InlineResponse hook (explorerKey=" + explorerKey + ").",
                        rollbackEx);
                }
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to hook Explorer.InlineResponseClose (explorerKey=" + explorerKey + ").",
                    ex);
                return false;
            }
            bool selectionChangeHooked = false;
            try
            {
                explorerEvents.SelectionChange += selectionChangeHandler;
                selectionChangeHooked = true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to hook Explorer.SelectionChange (explorerKey=" + explorerKey + ").",
                    ex);
            }

            _hookedExplorerEvents[explorerKey] = explorerEvents;
            _inlineResponseHandlers[explorerKey] = inlineResponseHandler;
            _inlineResponseCloseHandlers[explorerKey] = inlineResponseCloseHandler;
            if (selectionChangeHooked)
            {
                _explorerSelectionChangeHandlers[explorerKey] = selectionChangeHandler;
            }
            _hookedExplorers[explorerKey] = explorer;
            OnExplorerSelectionChanged(explorerKey);
            LogCore(
                "Explorer lifecycle hooked (explorerKey="
                + explorerKey
                + ", calendarSelection="
                + selectionChangeHooked
                + ").");
            return true;
        }

        private void UnhookExplorerHooks()
        {
            foreach (var pair in _hookedExplorerEvents)
            {
                Outlook.ExplorerEvents_10_InlineResponseEventHandler inlineResponseHandler;
                if (_inlineResponseHandlers.TryGetValue(pair.Key, out inlineResponseHandler))
                {
                    try
                    {
                        pair.Value.InlineResponse -= inlineResponseHandler;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to unhook Explorer.InlineResponse (explorerKey=" + pair.Key + ").",
                            ex);
                    }
                }

                Outlook.ExplorerEvents_10_InlineResponseCloseEventHandler inlineResponseCloseHandler;
                if (_inlineResponseCloseHandlers.TryGetValue(pair.Key, out inlineResponseCloseHandler))
                {
                    try
                    {
                        pair.Value.InlineResponseClose -= inlineResponseCloseHandler;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to unhook Explorer.InlineResponseClose (explorerKey=" + pair.Key + ").",
                            ex);
                    }
                }

                Outlook.ExplorerEvents_10_SelectionChangeEventHandler selectionChangeHandler;
                if (_explorerSelectionChangeHandlers.TryGetValue(pair.Key, out selectionChangeHandler))
                {
                    try
                    {
                        pair.Value.SelectionChange -= selectionChangeHandler;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to unhook Explorer.SelectionChange (explorerKey=" + pair.Key + ").",
                            ex);
                    }
                }
            }
            _hookedExplorerEvents.Clear();
            _inlineResponseHandlers.Clear();
            _inlineResponseCloseHandlers.Clear();
            _explorerSelectionChangeHandlers.Clear();
            _inlineResponseSubscriptions.Clear();

            foreach (var pair in _hookedExplorers)
            {
                ComInteropScope.TryRelease(pair.Value, LogCategories.Core, "Failed to release tracked Explorer COM object.");
            }
            _hookedExplorers.Clear();

            if (_explorersEvents != null)
            {
                try
                {
                    _explorersEvents.NewExplorer -= OnNewExplorer;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Explorers.NewExplorer.", ex);
                }
                _explorersEvents = null;
            }

            ComInteropScope.TryFinalRelease(_explorers, LogCategories.Core, "Failed to release Explorers COM object.");
            _explorers = null;
        }

        private void UnhookInspector()
        {
            if (_inspectors == null)
            {
                return;
            }
            try
            {
                _inspectors.NewInspector -= OnNewInspector;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Inspectors.NewInspector.", ex);
            }
            finally
            {
                ComInteropScope.TryFinalRelease(_inspectors, LogCategories.Core, "Failed to release Inspectors COM object.");
                _inspectors = null;
            }
        }

        private void OnNewInspector(Outlook.Inspector inspector)
        {
            if (inspector == null)
            {
                return;
            }
            try
            {
                if (!TryInitializeProfileRuntime()) { return; }
                BindInspectorItem(inspector);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to process NewInspector event.", ex);
            }
        }

        private bool BindInspectorItem(Outlook.Inspector inspector)
        {
            if (inspector == null) { return false; }
            var appointment = inspector.CurrentItem as Outlook.AppointmentItem;
            if (appointment != null)
            {
                EnsureSubscriptionForAppointment(appointment);
            }
            var mail = inspector.CurrentItem as Outlook.MailItem;
            if (mail != null && IsMailComposeCandidate(mail, "new_inspector"))
            {
                string inspectorIdentityKey = ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "Inspector");
                MailComposeSubscription subscription = EnsureMailComposeSubscription(
                    mail,
                    inspectorIdentityKey,
                    false,
                    null,
                    inspector);
                if (subscription != null && DiagnosticsLogger.IsEnabled)
                {
                    LogFileLink(
                        "Compose subscription associated with Inspector surface (inspectorKey="
                        + (inspectorIdentityKey ?? string.Empty)
                        + ").");
                }
                return subscription != null;
            }
            return false;
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            try
            {
                TryInitializeProfileRuntime();
                if (_profileRuntimeStopped) { return; }
                HookExplorer(explorer);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook new Explorer lifecycle.", ex);
            }
        }

        private void OnExplorerInlineResponse(string explorerKey, object item)
        {
            if (!TryInitializeProfileRuntime()) { return; }
            var mail = item as Outlook.MailItem;
            if (mail == null || !IsMailComposeCandidate(mail, "inline_response"))
            {
                return;
            }
            try
            {
                MailComposeSubscription subscription = EnsureMailComposeSubscription(mail, string.Empty, true, explorerKey);
                if (subscription == null)
                {
                    return;
                }

                MailComposeSubscription previousSubscription;
                if (_inlineResponseSubscriptions.TryGetValue(explorerKey, out previousSubscription)
                    && !ReferenceEquals(previousSubscription, subscription))
                {
                    previousSubscription.MarkInlineResponseClosed(explorerKey);
                }
                _inlineResponseSubscriptions[explorerKey] = subscription;
                if (DiagnosticsLogger.IsEnabled)
                {
                    LogFileLink(
                        "Compose subscription bound via Explorer.InlineResponse (explorerKey="
                        + (explorerKey ?? string.Empty)
                        + ").");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to bind compose subscription on Explorer.InlineResponse (explorerKey="
                    + (explorerKey ?? string.Empty)
                    + ").",
                    ex);
            }
        }

        private void OnExplorerInlineResponseClose(string explorerKey)
        {
            try
            {
                MailComposeSubscription subscription;
                if (!_inlineResponseSubscriptions.TryGetValue(explorerKey, out subscription))
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink(
                            "Explorer.InlineResponseClose ignored without tracked compose subscription (explorerKey="
                            + (explorerKey ?? string.Empty)
                            + ").");
                    }
                    return;
                }

                _inlineResponseSubscriptions.Remove(explorerKey);
                subscription.MarkInlineResponseClosed(explorerKey);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to close inline compose surface (explorerKey="
                    + (explorerKey ?? string.Empty)
                    + ").",
                    ex);
            }
        }

        private static bool IsMailComposeCandidate(Outlook.MailItem mail, string reason)
        {
            if (mail == null)
            {
                return false;
            }
            try
            {
                if (mail.Sent)
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink("Mail compose subscription skipped (reason=" + (reason ?? "n/a") + ", sent=True).");
                    }
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to verify mail compose state (reason=" + (reason ?? "n/a") + ").", ex);
                return false;
            }
        }
    }
}


// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
        // Outlook hook lifecycle for inspector/application events and compose-tracking entry points.
    public sealed partial class NextcloudTalkAddIn
    {
        private void EnsureInspectorHook()
        {            if (_outlookApplication == null || _inspectors != null)
            {
                return;
            }
            try
            {
                _inspectors = _outlookApplication.Inspectors;                if (_inspectors != null)
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
        {            if (_outlookApplication == null)
            {
                return;
            }
            try
            {
                EnsureExplorerInlineResponseHooks();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Explorer.InlineResponse.", ex);
            }
        }

        private void UnhookApplication()
        {
            UnhookExplorerInlineResponseHooks();
        }

        private void EnsureExplorerInlineResponseHooks()
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
                            HookInlineResponseExplorer(explorer);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook existing Explorer.InlineResponse.", ex);
                            ComInteropScope.TryRelease(explorer, LogCategories.Core, "Failed to release Explorer after hook failure.");
                        }
                    }
                }

                Outlook.Explorer activeExplorer = null;
                try
                {
                    activeExplorer = _outlookApplication.ActiveExplorer();
                    HookInlineResponseExplorer(activeExplorer);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook active Explorer.InlineResponse.", ex);
                    ComInteropScope.TryRelease(activeExplorer, LogCategories.Core, "Failed to release ActiveExplorer after hook failure.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to ensure Explorer.InlineResponse hooks.", ex);
            }
        }

        private bool HookInlineResponseExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                return false;
            }
            string explorerKey = ComInteropScope.ResolveIdentityKey(explorer, LogCategories.Core, "Explorer");
            if (string.IsNullOrWhiteSpace(explorerKey) || _inlineResponseExplorerEvents.ContainsKey(explorerKey))
            {
                return false;
            }

            var explorerEvents = explorer as Outlook.ExplorerEvents_10_Event;
            if (explorerEvents == null)
            {
                return false;
            }

            explorerEvents.InlineResponse += OnExplorerInlineResponse;
            _inlineResponseExplorerEvents[explorerKey] = explorerEvents;
            _inlineResponseExplorers[explorerKey] = explorer;
            LogCore("Explorer.InlineResponse hooked (explorerKey=" + explorerKey + ").");
            return true;
        }

        private void UnhookExplorerInlineResponseHooks()
        {
            foreach (var pair in _inlineResponseExplorerEvents)
            {
                try
                {
                    pair.Value.InlineResponse -= OnExplorerInlineResponse;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Explorer.InlineResponse.", ex);
                }
            }
            _inlineResponseExplorerEvents.Clear();

            foreach (var pair in _inlineResponseExplorers)
            {
                ComInteropScope.TryRelease(pair.Value, LogCategories.Core, "Failed to release tracked Explorer COM object.");
            }
            _inlineResponseExplorers.Clear();

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
        {            if (_inspectors == null)
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
        {            if (inspector == null)
            {
                return;
            }
            try
            {
                var appointment = inspector.CurrentItem as Outlook.AppointmentItem;                if (appointment != null)
                {
                    EnsureSubscriptionForAppointment(appointment);
                }
                var mail = inspector.CurrentItem as Outlook.MailItem;                if (mail != null && IsMailComposeCandidate(mail, "new_inspector"))
                {
                    string inspectorIdentityKey = ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "Inspector");
                    EnsureMailComposeSubscription(mail, inspectorIdentityKey);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to process NewInspector event.", ex);
            }
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            try
            {
                HookInlineResponseExplorer(explorer);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook new Explorer.InlineResponse.", ex);
            }
        }

        private void OnExplorerInlineResponse(object item)
        {
            var mail = item as Outlook.MailItem;
            if (mail == null || !IsMailComposeCandidate(mail, "inline_response"))
            {
                return;
            }
            try
            {
                EnsureMailComposeSubscription(mail, string.Empty, true);
                if (DiagnosticsLogger.IsEnabled)
                {
                    LogFileLink("Compose subscription ensured via Explorer.InlineResponse.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to ensure compose subscription on Explorer.InlineResponse.", ex);
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


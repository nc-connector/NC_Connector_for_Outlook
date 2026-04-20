/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    /**
     * Outlook hook lifecycle for inspector/application events and compose-tracking entry points.
     */
    public sealed partial class NextcloudTalkAddIn
    {
        private void EnsureInspectorHook()
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (_outlookApplication == null || _inspectors != null)
            {
                return;
            }

            try
            {
                _inspectors = _outlookApplication.Inspectors;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_outlookApplication == null || _applicationEvents != null)
            {
                return;
            }

            try
            {
                _applicationEvents = _outlookApplication as Outlook.ApplicationEvents_11_Event;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_applicationEvents != null)
                {
                    _applicationEvents.ItemLoad += OnApplicationItemLoad;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook ApplicationEvents_11.ItemLoad.", ex);
                _applicationEvents = null;
            }
        }

        private void UnhookApplication()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_applicationEvents == null)
            {
                return;
            }

            try
            {
                _applicationEvents.ItemLoad -= OnApplicationItemLoad;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook ApplicationEvents_11.ItemLoad.", ex);
            }
            finally
            {
                _applicationEvents = null;
            }
        }

        private void UnhookInspector()
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (inspector == null)
            {
                return;
            }

            try
            {
                var appointment = inspector.CurrentItem as Outlook.AppointmentItem;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (appointment != null)
                {
                    EnsureSubscriptionForAppointment(appointment);
                }

                var mail = inspector.CurrentItem as Outlook.MailItem;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (mail != null)
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

        private void OnApplicationItemLoad(object item)
        {
            var mail = item as Outlook.MailItem;
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (mail == null)
            {
                return;
            }

            Outlook.Explorer explorer = null;
            object inlineResponseObject = null;
            try
            {
                string inspectorIdentityKey = ResolveMailInspectorIdentityKey(mail);
                if (!string.IsNullOrWhiteSpace(inspectorIdentityKey))
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink("Application.ItemLoad compose subscription skipped (reason=inspector_context).");
                    }
                    return;
                }

                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_outlookApplication != null)
                {
                    explorer = _outlookApplication.ActiveExplorer();
                }

                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (explorer == null)
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink("Application.ItemLoad compose subscription skipped (reason=no_active_explorer).");
                    }
                    return;
                }

                inlineResponseObject = explorer.ActiveInlineResponse;
                var inlineResponseMail = inlineResponseObject as Outlook.MailItem;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (inlineResponseMail == null)
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink("Application.ItemLoad compose subscription skipped (reason=no_inline_response).");
                    }
                    return;
                }

                if (!ComInteropScope.AreSameObject(
                    mail,
                    inlineResponseMail,
                    LogCategories.FileLink,
                    "ItemLoad.MailItem",
                    "Explorer.ActiveInlineResponse"))
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink("Application.ItemLoad compose subscription skipped (reason=not_active_inline_response).");
                    }
                    return;
                }

                EnsureMailComposeSubscription(mail, string.Empty);
                if (DiagnosticsLogger.IsEnabled)
                {
                    LogFileLink("Compose subscription ensured via Application.ItemLoad (inline).");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to ensure compose subscription on Application.ItemLoad.", ex);
            }
            finally
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (inlineResponseObject != null && !ReferenceEquals(inlineResponseObject, mail))
                {
                    ComInteropScope.TryRelease(
                        inlineResponseObject,
                        LogCategories.FileLink,
                        "Failed to release ActiveInlineResponse COM object in ItemLoad handler.");
                }

                ComInteropScope.TryRelease(
                    explorer,
                    LogCategories.FileLink,
                    "Failed to release ActiveExplorer COM object in ItemLoad handler.");
            }
        }
    }
}

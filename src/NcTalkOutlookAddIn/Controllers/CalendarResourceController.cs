// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // Owns the selection transaction; CalDAV Synchronizer, not this controller, books resources.
    internal sealed class CalendarResourceController
    {
        private readonly NextcloudTalkAddIn _owner;
        private bool _opening;
        private bool _applying;
        private const string AccessProperty = "http://schemas.microsoft.com/mapi/proptag/0x0FF40003";

        internal CalendarResourceController(NextcloudTalkAddIn owner) { _owner = owner; }

        internal bool IsContextAvailable(IRibbonControl control)
        {
            if (_opening) { return false; }
            Outlook.AppointmentItem appointment = GetAppointment(control);
            try { return IsEditable(appointment); }
            finally { Release(appointment); }
        }

        internal string GetContextTooltip(IRibbonControl control)
        {
            Outlook.AppointmentItem appointment = GetAppointment(control);
            try { return IsEditable(appointment) ? Strings.ResourceRibbonTooltip : Strings.ResourceEditableAppointmentRequired; }
            finally { Release(appointment); }
        }

        private Outlook.AppointmentItem GetAppointment(IRibbonControl control)
        {
            Outlook.Inspector inspector = control == null ? null : control.Context as Outlook.Inspector;
            try { return inspector != null ? inspector.CurrentItem as Outlook.AppointmentItem : _owner.GetActiveAppointment(); }
            finally { Release(inspector); }
        }

        internal async Task OpenAsync(IRibbonControl control)
        {
            if (_opening) { return; }
            Outlook.AppointmentItem appointment = GetAppointment(control);
            if (!IsEditable(appointment))
            {
                Release(appointment);
                MessageBox.Show(Strings.ResourceEditableAppointmentRequired, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _opening = true;
            _owner.InvalidateActionRibbons();
            _applying = false;
            Exception failure = null;
            try
            {
                await OpenCoreAsync(appointment).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                failure = ex;
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Resource selection failed.", ex);
            }
            await _owner.RunOnOutlookUiThreadAsync(() =>
            {
                try
                {
                    if (failure != null && _owner.IsItemOpenForRibbonAction(appointment))
                    {
                        MessageBox.Show(PolicyUiHelper.GetConnectionFailureMessage(failure,
                            _applying ? Strings.ResourceApplyFailed : Strings.ResourceLoadFailed), Strings.DialogTitle,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                finally
                {
                    _opening = false;
                    Release(appointment);
                    _owner.InvalidateActionRibbons();
                }
            }).ConfigureAwait(false);
        }

        private async Task OpenCoreAsync(Outlook.AppointmentItem appointment)
        {
            AddinSettings settings = await _owner.EnsureConnectionForActionAsync(appointment, true, "resource_picker").ConfigureAwait(false);
            if (settings == null) { return; }
            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            BackendPolicyStatus status = await Task.Run(() => _owner.FetchEnterpriseRolloutPolicyStatus(configuration, "resource_picker")).ConfigureAwait(false);
            string notice = PolicyUiHelper.GetActionUnavailableNotice(settings, status, true);
            if (notice.Length != 0)
            {
                await ShowNoticeAsync(appointment, notice).ConfigureAwait(false);
                return;
            }
            CalendarResourceDirectory directory = _owner.CalendarResources;
            if (directory == null) { throw new InvalidOperationException("Resource directory is unavailable."); }
            List<CalendarResource> resources = await Task.Run(() => directory.GetResources(configuration, settings.IfbCacheHours, true)).ConfigureAwait(false);
            Selection selected = await _owner.RunOnOutlookUiThreadAsync(() => SelectOnUiThread(appointment, settings, resources, status)).ConfigureAwait(false);
            if (selected == null) { return; }
            // The dialog may remain open across a Seat change. Recheck before mutating Outlook.
            status = await Task.Run(() => _owner.FetchEnterpriseRolloutPolicyStatus(configuration, "resource_apply")).ConfigureAwait(false);
            notice = PolicyUiHelper.GetActionUnavailableNotice(settings, status, true);
            if (notice.Length != 0) { await ShowNoticeAsync(appointment, notice).ConfigureAwait(false); return; }
            NextcloudConnectionState.AssertRequestAllowed(configuration);
            _applying = true;
            await _owner.RunOnOutlookUiThreadAsync(() =>
            {
                if (!CanContinue(appointment, settings)) { return; }
                if (!selected.TimeWindow.Matches(CaptureTimeWindow(appointment)))
                {
                    MessageBox.Show(Strings.ResourceAppointmentChanged, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                NextcloudConnectionState.AssertRequestAllowed(configuration);
                BackendPolicyStatus latest;
                if (!_owner.TryGetCachedEmailSignaturePolicyStatus(configuration, out latest)) { latest = status; }
                string latestNotice = PolicyUiHelper.GetActionUnavailableNotice(settings, latest, true);
                if (latestNotice.Length != 0)
                {
                    PolicyUiHelper.ShowActionUnavailableNotice(latestNotice);
                    return;
                }
                Apply(appointment, resources, selected.Resources, selected.Recipients, selected.Location, selected.MeetingStatus);
            }).ConfigureAwait(false);
        }

        private Task ShowNoticeAsync(Outlook.AppointmentItem appointment, string text)
        {
            return _owner.RunOnOutlookUiThreadAsync(() =>
            {
                if (_owner.IsItemOpenForRibbonAction(appointment))
                { PolicyUiHelper.ShowActionUnavailableNotice(text); }
            });
        }

        private bool CanContinue(Outlook.AppointmentItem appointment, AddinSettings settings)
        {
            return _owner.IsItemOpenForRibbonAction(appointment) && IsEditable(appointment)
                && SettingsWorkflowController.ConnectionSettingsMatch(_owner.CurrentSettings, settings);
        }

        private Selection SelectOnUiThread(Outlook.AppointmentItem appointment, AddinSettings settings, List<CalendarResource> resources, BackendPolicyStatus status)
        {
            if (!CanContinue(appointment, settings)) { return null; }
            List<RecipientState> recipients = CaptureRecipients(appointment);
            string location = appointment.Location ?? string.Empty;
            Outlook.OlMeetingStatus meetingStatus = appointment.MeetingStatus;
            using (var dialog = new CalendarResourceForm(resources, recipients.Select(r => r.Email), settings.IfbEnabled && settings.IsManagedIfbValid, status, settings.ServerUrl))
            {
                var service = new CalDavFreeBusyService(_owner.CalendarResources,
                    new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword), settings.IfbCacheHours);
                AddinSettings preferences = _owner.SettingsStorage.Load();
                dialog.ConfigureAvailability(service, () => CanContinue(appointment, settings) ? CaptureTimeWindow(appointment) : null,
                    preferences.CalendarResourceGrouping, grouping => SaveGrouping(grouping),
                    preferences.CalendarResourceGroupsExpanded, expanded => SaveGroupsExpanded(expanded));
                if (dialog.ShowDialog() != DialogResult.OK) { return null; }
                if (dialog.TimeWindow == null) { return null; }
                return new Selection { Resources = dialog.SelectedResources, Recipients = recipients, Location = location, MeetingStatus = meetingStatus, TimeWindow = dialog.TimeWindow };
            }
        }

        private void SaveGrouping(string grouping)
        {
            try
            {
                if (_owner.SettingsStorage.SaveCalendarResourceGrouping(grouping))
                {
                    _owner.CurrentSettings.CalendarResourceGrouping = grouping;
                    return;
                }
            }
            catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Ifb, "Resource grouping preference could not be saved.", ex); }
            MessageBox.Show(Strings.ResourceGroupingSaveFailed, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void SaveGroupsExpanded(bool expanded)
        {
            try
            {
                if (_owner.SettingsStorage.SaveCalendarResourceGroupsExpanded(expanded))
                {
                    _owner.CurrentSettings.CalendarResourceGroupsExpanded = expanded;
                    return;
                }
            }
            catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Ifb, "Resource group expansion preference could not be saved.", ex); }
            MessageBox.Show(Strings.ResourceGroupingSaveFailed, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        internal static CalendarResourceTimeWindow CaptureTimeWindow(Outlook.AppointmentItem appointment)
        {
            Outlook.TimeZone timeZone = appointment.StartTimeZone;
            try
            {
                return new CalendarResourceTimeWindow(appointment.StartUTC, appointment.EndUTC, TimeZoneInfo.FindSystemTimeZoneById(timeZone.ID));
            }
            finally { Release(timeZone); }
        }

        internal static bool IsEditable(Outlook.AppointmentItem appointment)
        {
            if (appointment == null) { return false; }
            Outlook.PropertyAccessor accessor = null;
            Outlook.MAPIFolder folder = null;
            try
            {
                if (appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting && appointment.MeetingStatus != Outlook.OlMeetingStatus.olMeeting) { return false; }
                if (appointment.ResponseStatus != Outlook.OlResponseStatus.olResponseNone && appointment.ResponseStatus != Outlook.OlResponseStatus.olResponseOrganized) { return false; }
                if (string.IsNullOrEmpty(appointment.EntryID))
                {
                    folder = appointment.Parent as Outlook.MAPIFolder;
                    if (folder == null) { return false; }
                    accessor = folder.PropertyAccessor;
                    return (Convert.ToInt32((object)accessor.GetProperty(AccessProperty)) & 16) != 0;
                }
                accessor = appointment.PropertyAccessor;
                return (Convert.ToInt32((object)accessor.GetProperty(AccessProperty)) & 1) != 0;
            }
            catch (COMException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Appointment write access could not be checked.", ex);
                return false;
            }
            finally { Release(accessor); Release(folder); }
        }

        internal sealed class RecipientState
        {
            internal string Email;
            internal int Type;
            internal string Name;
            internal bool Sendable;
        }

        private sealed class Selection
        {
            internal List<CalendarResource> Resources;
            internal List<RecipientState> Recipients;
            internal string Location;
            internal Outlook.OlMeetingStatus MeetingStatus;
            internal CalendarResourceTimeWindow TimeWindow;
        }

        internal sealed class SelectionPlan
        {
            internal readonly List<CalendarResource> Add = new List<CalendarResource>();
            internal readonly List<int> Remove = new List<int>();
            internal readonly List<int> Retype = new List<int>();
            internal string Location;
        }

        internal static SelectionPlan Plan(List<CalendarResource> directory, List<CalendarResource> selected, List<RecipientState> existing, string currentLocation)
        {
            if (selected.Count(r => r.IsRoom) > 1) { throw new InvalidOperationException("Only one room can be selected."); }
            var available = directory.ToDictionary(r => r.Email, StringComparer.OrdinalIgnoreCase);
            var chosen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var plan = new SelectionPlan { Location = currentLocation };
            foreach (CalendarResource resource in selected)
            {
                CalendarResource known;
                if (!available.TryGetValue(resource.Email, out known) || !chosen.Add(resource.Email))
                { throw new InvalidOperationException("Invalid resource selection."); }
                if (known.IsRoom && known.Location.Length != 0) { plan.Location = known.Location; }
                if (!existing.Any(r => string.Equals(r.Email, known.Email, StringComparison.OrdinalIgnoreCase))) { plan.Add.Add(known); }
            }
            for (int i = 0; i < existing.Count; i++)
            {
                CalendarResource known;
                if (!available.TryGetValue(existing[i].Email, out known)) { continue; }
                if (!chosen.Contains(known.Email))
                {
                    plan.Remove.Add(i);
                    // Clear only the old room's exact location, never a manually edited location.
                    if (known.IsRoom && !selected.Any(r => r.IsRoom) && known.Location.Length != 0 && currentLocation == known.Location) { plan.Location = string.Empty; }
                }
                else if (existing[i].Type != (int)Outlook.OlMeetingRecipientType.olResource) { plan.Retype.Add(i); }
            }
            return plan;
        }

        private static List<RecipientState> CaptureRecipients(Outlook.AppointmentItem appointment)
        {
            var result = new List<RecipientState>();
            Outlook.Recipients recipients = appointment.Recipients;
            try
            {
                for (int i = 1; i <= recipients.Count; i++)
                {
                    Outlook.Recipient recipient = recipients[i];
                    try
                    {
                        result.Add(new RecipientState { Email = OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(recipient) ?? string.Empty,
                            Type = recipient.Type, Name = recipient.Name ?? string.Empty, Sendable = recipient.Sendable });
                    }
                    finally { Release(recipient); }
                }
            }
            finally { Release(recipients); }
            return result;
        }

        internal static string GetSelectionChange(List<RecipientState> current, string location, int meetingStatus,
            List<RecipientState> before, string oldLocation, int oldMeetingStatus)
        {
            // Outlook can expose an unset Location as null; it is the same empty field shown by the picker.
            if (!string.Equals(location ?? string.Empty, oldLocation ?? string.Empty, StringComparison.Ordinal)) { return "location"; }
            if (meetingStatus != oldMeetingStatus) { return "meeting_status"; }
            if (current.Count != before.Count) { return "recipient_count"; }
            if (current.Where((r, i) => r.Email != before[i].Email || r.Type != before[i].Type || r.Name != before[i].Name || r.Sendable != before[i].Sendable).Any()) { return "recipient_details"; }
            return string.Empty;
        }

        private static void Apply(Outlook.AppointmentItem appointment, List<CalendarResource> directory, List<CalendarResource> selected,
            List<RecipientState> before, string oldLocation, Outlook.OlMeetingStatus oldMeetingStatus)
        {
            List<RecipientState> current = CaptureRecipients(appointment);
            string change = GetSelectionChange(current, appointment.Location, (int)appointment.MeetingStatus, before, oldLocation, (int)oldMeetingStatus);
            if (change.Length != 0)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "Resource selection stopped because the appointment changed (field=" + change + ").");
                throw new InvalidOperationException("Appointment changed during resource selection.");
            }
            SelectionPlan plan = Plan(directory, selected, before, oldLocation);
            Outlook.Recipients recipients = appointment.Recipients;
            var added = new List<Outlook.Recipient>();
            var removed = new List<RecipientState>();
            try
            {
                foreach (CalendarResource resource in plan.Add)
                {
                    Outlook.Recipient recipient = recipients.Add(resource.Email);
                    added.Add(recipient);
                    recipient.Type = (int)Outlook.OlMeetingRecipientType.olResource;
                    if (!recipient.Resolve() || !string.Equals(OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(recipient), resource.Email, StringComparison.OrdinalIgnoreCase))
                    { throw new InvalidOperationException("Resource recipient could not be resolved exactly."); }
                }
                foreach (int index in plan.Retype)
                {
                    Outlook.Recipient recipient = recipients[index + 1];
                    try { recipient.Type = (int)Outlook.OlMeetingRecipientType.olResource; }
                    finally { Release(recipient); }
                }
                appointment.Location = plan.Location;
                if (selected.Count != 0 && oldMeetingStatus == Outlook.OlMeetingStatus.olNonMeeting) { appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting; }
                // CalDAV schedules resources; Outlook must not also mail their calendar addresses.
                // Sendable applies only after meeting conversion. Never change existing recipients:
                // switching a previously invited recipient off can make Outlook send a cancellation.
                foreach (Outlook.Recipient recipient in added)
                {
                    recipient.Sendable = false;
                    if (recipient.Sendable) { throw new InvalidOperationException("Resource invitation could not be disabled."); }
                }
                // Delete last: a failed addition must not discard an existing room or resource.
                foreach (int index in plan.Remove.OrderByDescending(i => i))
                {
                    recipients.Remove(index + 1);
                    removed.Add(before[index]);
                }
                DiagnosticsLogger.Log(LogCategories.Ifb, "Resource selection applied (added=" + plan.Add.Count + ", removed=" + plan.Remove.Count
                    + ", smtpInvitationsDisabled=" + added.Count + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Resource update failed; restoring affected appointment fields.", ex);
                foreach (Outlook.Recipient recipient in added) { TryRestore(() => recipient.Delete()); }
                foreach (RecipientState state in removed)
                {
                    TryRestore(() =>
                    {
                        Outlook.Recipient recipient = recipients.Add(state.Email);
                        try
                        {
                            recipient.Type = state.Type;
                            recipient.Resolve();
                            recipient.Sendable = state.Sendable;
                            if (recipient.Sendable != state.Sendable) { throw new InvalidOperationException("Resource invitation state could not be restored."); }
                        }
                        finally { Release(recipient); }
                    });
                }
                foreach (int index in plan.Retype)
                {
                    RecipientState state = before[index];
                    for (int i = 1; i <= recipients.Count; i++)
                    {
                        Outlook.Recipient recipient = recipients[i];
                        try
                        {
                            if (string.Equals(OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(recipient), state.Email, StringComparison.OrdinalIgnoreCase))
                            { TryRestore(() => recipient.Type = state.Type); }
                        }
                        finally { Release(recipient); }
                    }
                }
                TryRestore(() => appointment.Location = oldLocation);
                TryRestore(() => appointment.MeetingStatus = oldMeetingStatus);
                throw;
            }
            finally { foreach (Outlook.Recipient recipient in added) { Release(recipient); } Release(recipients); }
        }

        private static void TryRestore(Action restore)
        {
            try { restore(); }
            catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Ifb, "Appointment resource rollback was incomplete.", ex); }
        }

        private static void Release(object value)
        {
            ComInteropScope.TryRelease(value, LogCategories.Ifb, "Failed to release resource selection COM object.");
        }
    }
}

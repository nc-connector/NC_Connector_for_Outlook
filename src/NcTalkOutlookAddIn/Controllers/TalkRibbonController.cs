/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    /**
     * Handles Talk ribbon interactions including authentication gate, wizard orchestration,
     * room replacement flow, and appointment persistence.
     */
    internal sealed class TalkRibbonController
    {
        private readonly NextcloudTalkAddIn _owner;

        internal TalkRibbonController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal async Task OnTalkButtonPressedAsync(IRibbonControl control)
        {            if (_owner == null)
            {
                return;
            }

            AddinSettings settings = _owner.CurrentSettings;
            if (!_owner.SettingsAreComplete())
            {
                NextcloudTalkAddIn.LogTalkMessage("Talk link cancelled: settings are incomplete.");
                MessageBox.Show(
                    Strings.ErrorMissingCredentials,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                _owner.OnSettingsButtonPressed(control);
                return;
            }
            if (!EnsureAuthenticationValid(control))
            {
                NextcloudTalkAddIn.LogTalkMessage("Talk link cancelled: authentication failed.");
                return;
            }

            Outlook.AppointmentItem appointment = _owner.GetActiveAppointment();            if (appointment == null)
            {
                NextcloudTalkAddIn.LogTalkMessage("Talk link cancelled: no active appointment found.");
                MessageBox.Show(
                    Strings.ErrorNoAppointment,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            var subject = appointment.Subject ?? string.Empty;
            var start = appointment.Start == DateTime.MinValue ? DateTime.Now : appointment.Start;
            var end = appointment.End == DateTime.MinValue ? start.AddHours(1) : appointment.End;
            NextcloudTalkAddIn.LogTalkMessage("Talk link started (subject='" + subject + "', start=" + start.ToString("o") + ", end=" + end.ToString("o") + ").");

            settings = _owner.CurrentSettings ?? new AddinSettings();
            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            Task<BackendPolicyStatus> policyStatusTask = Task.Run(() => _owner.FetchBackendPolicyStatus(configuration, "talk_wizard_open"));
            Task<PasswordPolicyInfo> passwordPolicyTask = Task.Run(() => _owner.FetchPasswordPolicyForTalkWizard(configuration));
            await Task.WhenAll(policyStatusTask, passwordPolicyTask);
            BackendPolicyStatus policyStatus = policyStatusTask.Result;
            PasswordPolicyInfo passwordPolicy = passwordPolicyTask.Result;

            var addressbookCache = new IfbAddressBookCache(_owner.SettingsStorage != null ? _owner.SettingsStorage.DataDirectory : null);
            NextcloudTalkAddIn.LogTalkMessage("System address book status check requested (context=talk_click, forceRefresh=True).");
            var talkClickAddressbookStatus = addressbookCache.GetSystemAddressbookStatus(
                configuration,
                settings.IfbCacheHours,
                true);
            NextcloudTalkAddIn.LogTalkMessage(
                "System address book status result (context=talk_click, available=" + talkClickAddressbookStatus.Available
                + ", count=" + talkClickAddressbookStatus.Count
                + ", hasError=" + (!string.IsNullOrWhiteSpace(talkClickAddressbookStatus.Error)) + ").");
            if (!talkClickAddressbookStatus.Available && !string.IsNullOrWhiteSpace(talkClickAddressbookStatus.Error))
            {
                NextcloudTalkAddIn.LogTalkMessage("System address book unavailable on talk click: " + talkClickAddressbookStatus.Error);
            }

            List<NextcloudUser> userDirectory;
            try
            {
                userDirectory = talkClickAddressbookStatus.Available
                    ? addressbookCache.GetUsers(configuration, settings.IfbCacheHours, false)
                    : new List<NextcloudUser>();
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("System address book could not be loaded: " + ex.Message);
                userDirectory = new List<NextcloudUser>();
            }

            using (var dialog = new TalkLinkForm(
                settings,
                configuration,
                passwordPolicy,
                policyStatus,
                userDirectory,
                talkClickAddressbookStatus,
                subject,
                start,
                end))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Talk link dialog cancelled.");
                    return;
                }
                string descriptionLanguage = NextcloudTalkAddIn.ResolveTalkDescriptionLanguage(
                    policyStatus,
                    settings.EventDescriptionLang);
                string descriptionType = NextcloudTalkAddIn.ResolveTalkEventDescriptionType(policyStatus);
                string invitationTemplate = NextcloudTalkAddIn.ResolveTalkInvitationTemplate(policyStatus);
                string initialDescription;
                try
                {
                    initialDescription = NextcloudTalkAddIn.BuildInitialRoomDescription(
                        dialog.TalkPassword,
                        descriptionLanguage,
                        invitationTemplate);
                }
                catch (Exception ex)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Talk invitation template rendering blocked: " + ex.Message);
                    MessageBox.Show(
                        string.Format(CultureInfo.CurrentCulture, Strings.ErrorCreateRoomUnexpected, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }
                var request = new TalkRoomRequest
                {
                    Title = dialog.TalkTitle,
                    Password = dialog.TalkPassword,
                    LobbyEnabled = dialog.LobbyUntilStart,
                    SearchVisible = dialog.SearchVisible,
                    RoomType = dialog.SelectedRoomType,
                    AppointmentStart = start,
                    AppointmentEnd = end,
                    DescriptionLanguage = descriptionLanguage,
                    DescriptionType = descriptionType,
                    InvitationTemplate = invitationTemplate,
                    Description = initialDescription,
                    AddUsers = dialog.AddUsers,
                    AddGuests = dialog.AddGuests,
                    DelegateModeratorId = dialog.DelegateModeratorId,
                    DelegateModeratorName = dialog.DelegateModeratorName
                };
                NextcloudTalkAddIn.LogTalkMessage("Room request prepared (title='" + request.Title + "', type=" + request.RoomType + ", lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ", passwordSet=" + (!string.IsNullOrEmpty(request.Password)) + ").");

                string existingToken = TalkAppointmentController.GetUserPropertyTextPrefer(appointment, NextcloudTalkAddIn.IcalToken, NextcloudTalkAddIn.PropertyToken);
                if (!string.IsNullOrWhiteSpace(existingToken))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Existing room found (token=" + existingToken + "), replacement requested.");
                    var overwrite = MessageBox.Show(
                        Strings.ConfirmReplaceRoom,
                        Strings.DialogTitle,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (overwrite != DialogResult.Yes)
                    {
                        NextcloudTalkAddIn.LogTalkMessage("Replacement declined, operation ended.");
                        return;
                    }
                    var existingType = TalkAppointmentController.GetRoomType(appointment);
                    bool existingIsEvent = existingType.HasValue && existingType.Value == TalkRoomType.EventConversation;
                    NextcloudTalkAddIn.LogTalkMessage("Attempting to delete existing room (event=" + existingIsEvent + ").");

                    if (!_owner.TryDeleteRoom(existingToken, existingIsEvent))
                    {
                        NextcloudTalkAddIn.LogTalkMessage("Deleting existing room failed.");
                        return;
                    }
                }

                TalkRoomCreationResult result;
                try
                {
                    using (new WaitCursorScope())
                    {
                        NextcloudTalkAddIn.LogTalkMessage("Sending CreateRoom request to Nextcloud.");
                        var service = _owner.CreateTalkService();
                        result = service.CreateRoom(request);
                    }
                    NextcloudTalkAddIn.LogTalkMessage("Room created successfully (token=" + result.RoomToken + ", URL=" + result.RoomUrl + ", event=" + result.CreatedAsEventConversation + ").");
                }
                catch (TalkServiceException ex)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Talk room could not be created: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoom, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        ex.IsAuthenticationError ? MessageBoxIcon.Warning : MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Unexpected error while creating talk room: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoomUnexpected, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                _owner.ApplyRoomToAppointment(appointment, request, result);
                NextcloudTalkAddIn.LogTalkMessage("Room data stored in appointment (EntryID=" + (appointment.EntryID ?? "n/a") + ").");

                MessageBox.Show(
                    string.Format(Strings.InfoRoomCreated, request.Title),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private bool EnsureAuthenticationValid(IRibbonControl control)
        {
            try
            {
                using (new WaitCursorScope())
                {
                    var service = _owner.CreateTalkService();
                    string response;
                    NextcloudTalkAddIn.LogTalkMessage("Starting credential verification request.");
                    if (service.VerifyConnection(out response))
                    {
                        _owner.UpdateStoredServerVersion(response);
                        NextcloudTalkAddIn.LogTalkMessage("Credentials verified (response=" + (string.IsNullOrEmpty(response) ? "OK" : response) + ").");
                        return true;
                    }
                    string message = string.IsNullOrEmpty(response)
                        ? Strings.ErrorCredentialsNotVerified
                        : string.Format(CultureInfo.CurrentCulture, Strings.ErrorCredentialsNotVerifiedFormat, response);
                    NextcloudTalkAddIn.LogTalkMessage("Invalid credentials: " + message);
                    return PromptOpenSettings(message, control);
                }
            }
            catch (TalkServiceException ex)
            {
                string message;
                if ((int)ex.StatusCode == 0)
                {
                    message = Strings.ErrorServerUnavailable;
                }
                else if (ex.IsAuthenticationError)
                {
                    message = string.Format(Strings.ErrorAuthenticationRejected, ex.Message);
                }
                else
                {
                    message = string.Format(Strings.ErrorConnectionFailed, ex.Message);
                }

                NextcloudTalkAddIn.LogTalkMessage("Connection check failed: " + message);
                return PromptOpenSettings(message, control);
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Unexpected error during connection check: " + ex.Message);
                return PromptOpenSettings(string.Format(Strings.ErrorUnknownAuthentication, ex.Message), control);
            }
        }

        private bool PromptOpenSettings(string message, IRibbonControl control)
        {
            var result = MessageBox.Show(
                string.Format(Strings.PromptOpenSettings, message),
                Strings.DialogTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                _owner.OnSettingsButtonPressed(control);
            }
            return false;
        }
    }
}


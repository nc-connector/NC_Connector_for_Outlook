// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;

namespace NcTalkOutlookAddIn.Controllers
{
    // Handles ribbon-driven FileLink launch and wizard execution for mail compose windows.
    internal sealed class FileLinkLaunchController
    {
        private readonly NextcloudTalkAddIn _owner;

        internal FileLinkLaunchController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal async Task OnFileLinkButtonPressedAsync(IRibbonControl control)
        {
            if (_owner == null)
            {
                return;
            }
            if (!_owner.SettingsAreComplete())
            {
                MessageBox.Show(
                    Strings.ErrorMissingCredentials,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                _owner.OnSettingsButtonPressed(control);
                return;
            }

            Outlook.MailItem mail = _owner.GetActiveMailItem();
            if (mail == null)
            {
                MessageBox.Show(
                    Strings.ErrorNoMailItem,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            bool isInlineResponse = _owner.IsActiveInlineResponse(mail);
            _owner.EnsureMailComposeSubscription(mail, isInlineResponse ? string.Empty : _owner.ResolveActiveInspectorIdentityKey(), isInlineResponse);
            await RunFileLinkWizardForMailAsync(mail, null);
        }

        internal async Task<bool> RunFileLinkWizardForMailAsync(Outlook.MailItem mail, FileLinkWizardLaunchOptions launchOptions)
        {
            AddinSettings settings = _owner != null ? _owner.CurrentSettings : null;
            if (_owner == null || mail == null || settings == null)
            {
                return false;
            }
            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl,
                settings.Username,
                settings.AppPassword);
            Task<NextcloudCapabilitiesSnapshot> capabilitiesTask = Task.Run(
                () => new NextcloudCapabilitiesService(configuration)
                    .GetRequiredSnapshot(false, false));
            Task<BackendPolicyStatus> policyStatusTask = Task.Run(() => _owner.FetchBackendPolicyStatus(configuration, "sharing_wizard_open"));
            Task<PasswordPolicyInfo> passwordPolicyTask = Task.Run(() => _owner.FetchPasswordPolicyForFileLinkWizard(configuration));
            string launchFailureMessage = null;
            MessageBoxIcon launchFailureIcon = MessageBoxIcon.Error;
            try
            {
                await Task.WhenAll(
                    capabilitiesTask,
                    policyStatusTask,
                    passwordPolicyTask).ConfigureAwait(false);
            }
            catch (TalkServiceException ex)
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Sharing wizard capability check failed: " + ex.Message);
                launchFailureMessage = ex.Message;
                launchFailureIcon = ex.IsAuthenticationError
                    ? MessageBoxIcon.Warning
                    : MessageBoxIcon.Error;
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Sharing wizard capability check failed unexpectedly: "
                    + ex.Message);
                launchFailureMessage = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.ErrorConnectionFailed,
                    ex.Message);
            }
            if (!string.IsNullOrWhiteSpace(launchFailureMessage))
            {
                return await ShowLaunchFailureAsync(
                    launchFailureMessage,
                    launchFailureIcon).ConfigureAwait(false);
            }

            NextcloudCapabilitiesSnapshot capabilities =
                await capabilitiesTask.ConfigureAwait(false);
            BackendPolicyStatus policyStatus =
                await policyStatusTask.ConfigureAwait(false);
            PasswordPolicyInfo passwordPolicy =
                await passwordPolicyTask.ConfigureAwait(false);

            return await _owner.RunOnOutlookUiThreadAsync(
                () => RunFileLinkWizardOnUiThread(
                    mail,
                    launchOptions,
                    settings,
                    configuration,
                    capabilities,
                    policyStatus,
                    passwordPolicy)).ConfigureAwait(false);
        }

        private Task<bool> ShowLaunchFailureAsync(
            string message,
            MessageBoxIcon icon)
        {
            return _owner.RunOnOutlookUiThreadAsync(
                () =>
                {
                    MessageBox.Show(
                        message,
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        icon);
                    return false;
                });
        }

        private bool RunFileLinkWizardOnUiThread(
            Outlook.MailItem mail,
            FileLinkWizardLaunchOptions launchOptions,
            AddinSettings settings,
            TalkServiceConfiguration configuration,
            NextcloudCapabilitiesSnapshot capabilities,
            BackendPolicyStatus policyStatus,
            PasswordPolicyInfo passwordPolicy)
        {
            string basePath = string.IsNullOrWhiteSpace(settings.FileLinkBasePath)
                ? AddinSettings.DefaultFileLinkBasePath
                : settings.FileLinkBasePath;

            NextcloudTalkAddIn.LogFileLinkMessage(
                "Sharing wizard UI ready (threadId="
                + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(CultureInfo.InvariantCulture)
                + ", apartment="
                + System.Threading.Thread.CurrentThread.GetApartmentState()
                + ").");

            using (var wizard = new FileLinkWizardForm(
                settings,
                configuration,
                capabilities,
                passwordPolicy,
                policyStatus,
                basePath,
                launchOptions))
            {
                if (wizard.ShowDialog() == DialogResult.OK && wizard.Result != null)
                {
                    string languageOverride = settings != null ? settings.ShareBlockLang : "default";
                    bool plainTextCompose = MailInteropController.IsPlainTextMail(mail);
                    NextcloudTalkAddIn.LogFileLinkMessage("Share created (folder=\"" + wizard.Result.FolderName + "\").");

                    bool isInlineResponse = _owner.IsActiveInlineResponse(mail);
                    NextcloudTalkAddIn.MailComposeSubscription composeSubscription = _owner.EnsureMailComposeSubscription(mail, isInlineResponse ? string.Empty : _owner.ResolveActiveInspectorIdentityKey(), isInlineResponse);
                    if (composeSubscription != null)
                    {
                        composeSubscription.ArmShareCleanup(wizard.Result);
                    }
                    string html;
                    string plainText;
                    try
                    {
                        html = plainTextCompose
                            ? string.Empty
                            : FileLinkHtmlBuilder.Build(wizard.Result, wizard.RequestSnapshot, languageOverride, policyStatus);
                        plainText = plainTextCompose
                            ? FileLinkHtmlBuilder.BuildPlainText(wizard.Result, wizard.RequestSnapshot, languageOverride, policyStatus)
                            : string.Empty;
                    }
                    catch (Exception ex)
                    {
                        NextcloudTalkAddIn.LogFileLinkMessage("Share template rendering blocked: " + ex.Message);
                        MessageBox.Show(
                            string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                            Strings.DialogTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return false;
                    }
                    if (composeSubscription != null
                        && wizard.RequestSnapshot != null
                        && wizard.RequestSnapshot.PasswordSeparateEnabled
                        && !string.IsNullOrWhiteSpace(wizard.Result.Password))
                    {
                        string passwordOnlyHtml;
                        string passwordOnlyPlainText;
                        try
                        {
                            passwordOnlyHtml = plainTextCompose
                                ? string.Empty
                                : FileLinkHtmlBuilder.BuildPasswordOnly(wizard.Result, languageOverride, policyStatus);
                            passwordOnlyPlainText = plainTextCompose
                                ? FileLinkHtmlBuilder.BuildPasswordOnlyPlainText(wizard.Result, languageOverride, policyStatus)
                                : string.Empty;
                        }
                        catch (Exception ex)
                        {
                            NextcloudTalkAddIn.LogFileLinkMessage("Password-only template rendering blocked: " + ex.Message);
                            MessageBox.Show(
                                string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                                Strings.DialogTitle,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return false;
                        }

                        composeSubscription.RegisterSeparatePasswordDispatch(
                            wizard.Result,
                            wizard.RequestSnapshot,
                            passwordOnlyHtml,
                            passwordOnlyPlainText,
                            plainTextCompose,
                            languageOverride,
                            policyStatus);
                    }

                    if (plainTextCompose)
                    {
                        _owner.InsertPlainTextIntoMail(mail, plainText);
                    }
                    else
                    {
                        _owner.InsertHtmlIntoMail(mail, html);
                    }
                    return true;
                }
            }
            return false;
        }
    }
}


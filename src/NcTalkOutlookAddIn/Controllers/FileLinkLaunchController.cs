// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Runtime.ExceptionServices;
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

            bool canContinue = await _owner.RunOnOutlookUiThreadAsync(
                () =>
                {
                    if (!_owner.IsItemOpenForRibbonAction(mail))
                    {
                        NextcloudTalkAddIn.LogFileLinkMessage("Sharing launch ended because the original compose item closed.");
                        return false;
                    }
                    bool isInlineResponse = _owner.IsActiveInlineResponse(mail);
                    _owner.EnsureMailComposeSubscription(
                        mail,
                        isInlineResponse ? string.Empty : MailInteropController.ResolveMailInspectorIdentityKey(mail),
                        isInlineResponse);
                    return true;
                }).ConfigureAwait(false);
            if (canContinue)
            {
                await RunFileLinkWizardForMailAsync(mail, null, true).ConfigureAwait(false);
            }
        }

        internal async Task<bool> RunFileLinkWizardForMailAsync(
            Outlook.MailItem mail,
            FileLinkWizardLaunchOptions launchOptions,
            bool allowInteractiveRecovery = false)
        {
            if (_owner == null || mail == null) { return false; }
            NextcloudTalkAddIn.MailComposeSubscription subscription = await _owner.RunOnOutlookUiThreadAsync(() =>
            {
                if (!_owner.IsItemOpenForRibbonAction(mail)) { return null; }
                bool inline = _owner.IsActiveInlineResponse(mail);
                var current = _owner.EnsureMailComposeSubscription(mail,
                    inline ? string.Empty : MailInteropController.ResolveMailInspectorIdentityKey(mail), inline);
                return current != null && current.TryBeginSharingWizard() ? current : null;
            }).ConfigureAwait(false);
            if (subscription == null) { return false; }
            bool accepted = false;
            ExceptionDispatchInfo failure = null;
            try
            {
                accepted = await RunReservedWizardAsync(mail, launchOptions, allowInteractiveRecovery, subscription).ConfigureAwait(false);
            }
            catch (Exception ex) { failure = ExceptionDispatchInfo.Capture(ex); }
            await _owner.RunOnOutlookUiThreadAsync(() => { subscription.EndSharingWizard(); }).ConfigureAwait(false);
            if (failure != null) { failure.Throw(); }
            return accepted;
        }

        private async Task<bool> RunReservedWizardAsync(
            Outlook.MailItem mail,
            FileLinkWizardLaunchOptions launchOptions,
            bool allowInteractiveRecovery,
            NextcloudTalkAddIn.MailComposeSubscription subscription)
        {
            AddinSettings settings = await _owner.EnsureConnectionForActionAsync(
                mail,
                allowInteractiveRecovery,
                "sharing_wizard_open",
                null).ConfigureAwait(false);
            if (settings == null)
            {
                return false;
            }
            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl, settings.Username, settings.AppPassword);
            // The shared entry gate has just refreshed these capabilities over a new connection.
            Task<NextcloudCapabilitiesSnapshot> capabilitiesTask = Task.Run(
                () => new NextcloudCapabilitiesService(configuration).GetRequiredSnapshot(false, false));
            Task<BackendPolicyStatus> policyStatusTask = Task.Run(() => settings.IsEnterpriseRollout
                ? _owner.FetchEnterpriseRolloutPolicyStatus(configuration, "sharing_wizard_open")
                : _owner.FetchBackendPolicyStatus(configuration, "sharing_wizard_open"));
            Task<PasswordPolicyInfo> passwordPolicyTask = Task.Run(() => _owner.FetchPasswordPolicyForFileLinkWizard(configuration));
            Exception launchFailure = null;
            try
            {
                await Task.WhenAll(capabilitiesTask, policyStatusTask, passwordPolicyTask).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { return false; }
            catch (Exception ex) { launchFailure = ex; }
            if (launchFailure != null)
            {

                NextcloudTalkAddIn.LogFileLinkMessage("Sharing prefetch failed: " + launchFailure.Message);
                await _owner.RunOnOutlookUiThreadAsync(() =>
                {
                    if (_owner.IsItemOpenForRibbonAction(mail))
                    {
                        MessageBox.Show(
                            SettingsWorkflowController.GetActionConnectionFailureMessage(launchFailure),
                            Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }).ConfigureAwait(false);
                return false;
            }
            NextcloudCapabilitiesSnapshot capabilities = await capabilitiesTask.ConfigureAwait(false);
            BackendPolicyStatus policyStatus = await policyStatusTask.ConfigureAwait(false);
            PasswordPolicyInfo passwordPolicy = await passwordPolicyTask.ConfigureAwait(false);
            Task<bool> wizardTask = await _owner.RunOnOutlookUiThreadAsync(
                () => RunFileLinkWizardOnUiThread(
                    mail, launchOptions, settings, configuration,
                    capabilities, policyStatus, passwordPolicy, subscription)).ConfigureAwait(false);
            return await wizardTask.ConfigureAwait(false);
        }

        private async Task<bool> RunFileLinkWizardOnUiThread(
            Outlook.MailItem mail,
            FileLinkWizardLaunchOptions launchOptions,
            AddinSettings settings,
            TalkServiceConfiguration configuration,
            NextcloudCapabilitiesSnapshot capabilities,
            BackendPolicyStatus policyStatus,
            PasswordPolicyInfo passwordPolicy,
            NextcloudTalkAddIn.MailComposeSubscription subscription)
        {
            if (!_owner.IsItemOpenForRibbonAction(mail)
                || !SettingsWorkflowController.ConnectionSettingsMatch(_owner.CurrentSettings, settings))
            {
                NextcloudTalkAddIn.LogFileLinkMessage("Sharing wizard ended because the original compose item or credentials changed during prefetch.");
                return false;
            }
            string rolloutNotice = PolicyUiHelper.GetActionUnavailableNotice(settings, policyStatus);
            if (!string.IsNullOrEmpty(rolloutNotice))
            {
                PolicyUiHelper.ShowActionUnavailableNotice(rolloutNotice);
                return false;
            }

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
                launchOptions))
            {
                if (!subscription.BindSharingWizard(wizard)) { return false; }
                var closed = new TaskCompletionSource<DialogResult>();
                wizard.FormClosed += (sender, args) => closed.TrySetResult(wizard.DialogResult);
                wizard.Show();
                DialogResult wizardResult = await closed.Task;
                if (wizardResult == DialogResult.OK && wizard.Result != null)
                {
                    if (!subscription.CanApplyComposeChanges)
                    {
                        _owner.QueueCreatedShareCleanup(string.Empty, wizard.Result,
                            ComposeLifecycleOrigin.Create(configuration), "compose_closed_before_insert");
                        return false;
                    }
                    string languageOverride = settings.ResolvePolicyDefaults(policyStatus).ShareBlockLang;
                    bool plainTextCompose = MailBodyInsertionController.IsPlainTextMail(mail);
                    NextcloudTalkAddIn.LogFileLinkMessage("Share created (folder=\"" + wizard.Result.FolderName + "\").");
                    ComposeLifecycleOrigin origin =
                        ComposeLifecycleOrigin.Create(configuration);
                    string composeKey = ComInteropScope.ResolveIdentityKey(
                        mail,
                        LogCategories.FileLink,
                        "MailItem");
                    string html;
                    string plainText;
                    string passwordOnlyHtml = string.Empty;
                    string passwordOnlyPlainText = string.Empty;
                    string secretsHtmlTemplate = string.Empty;
                    string secretsPlainTextTemplate = string.Empty;
                    bool registerSeparatePassword =
                        wizard.RequestSnapshot != null
                        && wizard.RequestSnapshot.PasswordSeparateEnabled
                        && !string.IsNullOrWhiteSpace(wizard.Result.Password);
                    try
                    {
                        html = plainTextCompose
                            ? string.Empty
                            : FileLinkHtmlBuilder.Build(wizard.Result, wizard.RequestSnapshot, languageOverride, policyStatus);
                        plainText = plainTextCompose
                            ? FileLinkHtmlBuilder.BuildPlainText(wizard.Result, wizard.RequestSnapshot, languageOverride, policyStatus)
                            : string.Empty;
                        if (registerSeparatePassword)
                        {
                            passwordOnlyHtml = plainTextCompose
                                ? string.Empty
                                : FileLinkHtmlBuilder.BuildPasswordOnly(
                                    wizard.Result,
                                    languageOverride,
                                    policyStatus);
                            passwordOnlyPlainText = plainTextCompose
                                ? FileLinkHtmlBuilder.BuildPasswordOnlyPlainText(
                                    wizard.Result,
                                    languageOverride,
                                    policyStatus)
                                : string.Empty;

                            FileLinkResult placeholderResult =
                                BuildSecretPlaceholderResult(wizard.Result);
                            secretsHtmlTemplate = plainTextCompose
                                ? string.Empty
                                : FileLinkHtmlBuilder.BuildPasswordOnly(
                                    placeholderResult,
                                    languageOverride,
                                    policyStatus,
                                    true);
                            secretsPlainTextTemplate = plainTextCompose
                                ? FileLinkHtmlBuilder.BuildPasswordOnlyPlainText(
                                    placeholderResult,
                                    languageOverride,
                                    policyStatus,
                                    true)
                                : string.Empty;
                        }
                    }
                    catch (Exception ex)
                    {

                        NextcloudTalkAddIn.LogFileLinkMessage("Share template rendering blocked: " + ex.Message);
                        MessageBox.Show(
                            string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                            Strings.DialogTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        _owner.QueueCreatedShareCleanup(
                            composeKey,
                            wizard.Result,
                            origin,
                            "render_failed");
                        return false;
                    }

                    if (!subscription.CanApplyComposeChanges)
                    {
                        _owner.QueueCreatedShareCleanup(composeKey, wizard.Result, origin, "compose_subscription_unavailable");
                        return false;
                    }

                    // Write the durable marker before body mutation; a failed marker must not leave an unguarded block.
                    try { subscription.SetSharingBlockInserted(true); }
                    catch (Exception ex)
                    {
                        NextcloudTalkAddIn.LogFileLinkMessage("Share marker could not be written: " + ex.Message);
                        _owner.QueueCreatedShareCleanup(composeKey, wizard.Result, origin, "share_marker_failed");
                        return false;
                    }
                    bool inserted = plainTextCompose
                        ? _owner.TryInsertPlainTextIntoMail(mail, plainText)
                        : _owner.TryInsertHtmlIntoMail(mail, html);
                    if (!inserted)
                    {
                        try { subscription.SetSharingBlockInserted(false); }
                        catch (Exception ex) { NextcloudTalkAddIn.LogFileLinkMessage("Failed to clear share marker: " + ex.Message); }
                        _owner.QueueCreatedShareCleanup(
                            composeKey,
                            wizard.Result,
                            origin,
                            "body_insert_failed");
                        return false;
                    }
                    subscription.ArmShareCleanup(
                        wizard.Result,
                        origin);
                    if (registerSeparatePassword)
                    {
                        subscription.RegisterSeparatePasswordDispatch(
                            wizard.Result,
                            wizard.RequestSnapshot,
                            passwordOnlyHtml,
                            passwordOnlyPlainText,
                            secretsHtmlTemplate,
                            secretsPlainTextTemplate,
                            plainTextCompose,
                            languageOverride,
                            policyStatus,
                            origin);
                    }
                    return true;
                }
            }
            return false;
        }

        private static FileLinkResult BuildSecretPlaceholderResult(
            FileLinkResult result)
        {
            return new FileLinkResult(
                result.ShareUrl,
                result.ShareId,
                result.ShareToken,
                SeparatePasswordDeliveryController.SecretDeliveryPlaceholder,
                result.ExpireDate,
                result.Permissions,
                result.FolderName,
                result.RelativePath);
        }
    }
}


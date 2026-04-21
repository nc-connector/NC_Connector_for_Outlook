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

        internal void OnFileLinkButtonPressed(IRibbonControl control)
        {            if (_owner == null)
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

            Outlook.MailItem mail = _owner.GetActiveMailItem();            if (mail == null)
            {
                MessageBox.Show(
                    Strings.ErrorNoMailItem,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            _owner.EnsureMailComposeSubscription(mail, _owner.ResolveActiveInspectorIdentityKey());
            RunFileLinkWizardForMail(mail, null);
        }

        internal bool RunFileLinkWizardForMail(Outlook.MailItem mail, FileLinkWizardLaunchOptions launchOptions)
        {
            AddinSettings settings = _owner != null ? _owner.CurrentSettings : null;            if (_owner == null || mail == null || settings == null)
            {
                return false;
            }
            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl,
                settings.Username,
                settings.AppPassword);
            // Keep this method synchronous for existing callers and prefetch both policies in parallel.
            Task<BackendPolicyStatus> policyStatusTask = Task.Run(() => _owner.FetchBackendPolicyStatus(configuration, "sharing_wizard_open"));
            Task<PasswordPolicyInfo> passwordPolicyTask = Task.Run(() => _owner.FetchPasswordPolicyForFileLinkWizard(configuration));
            Task.WhenAll(policyStatusTask, passwordPolicyTask).GetAwaiter().GetResult();
            BackendPolicyStatus policyStatus = policyStatusTask.Result;

            string basePath = string.IsNullOrWhiteSpace(settings.FileLinkBasePath)
                ? AddinSettings.DefaultFileLinkBasePath
                : settings.FileLinkBasePath;

            PasswordPolicyInfo passwordPolicy = passwordPolicyTask.Result;

            using (var wizard = new FileLinkWizardForm(settings, configuration, passwordPolicy, policyStatus, basePath, launchOptions))
            {                if (wizard.ShowDialog() == DialogResult.OK && wizard.Result != null)
                {
                    string languageOverride = settings != null ? settings.ShareBlockLang : "default";
                    NextcloudTalkAddIn.LogFileLinkMessage("Share created (folder=\"" + wizard.Result.FolderName + "\").");

                    NextcloudTalkAddIn.MailComposeSubscription composeSubscription = _owner.EnsureMailComposeSubscription(mail, _owner.ResolveActiveInspectorIdentityKey());                    if (composeSubscription != null)
                    {
                        composeSubscription.ArmShareCleanup(wizard.Result);
                    }
                    string html;
                    try
                    {
                        html = FileLinkHtmlBuilder.Build(wizard.Result, wizard.RequestSnapshot, languageOverride, policyStatus);
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
                        try
                        {
                            passwordOnlyHtml = FileLinkHtmlBuilder.BuildPasswordOnly(wizard.Result, languageOverride, policyStatus);
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
                            passwordOnlyHtml);
                    }

                    _owner.InsertHtmlIntoMail(mail, html);
                    return true;
                }
            }
            return false;
        }
    }
}


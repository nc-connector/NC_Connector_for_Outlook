// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Applies backend share defaults, lock state, and administrator guidance to the wizard.
    internal sealed partial class FileLinkWizardForm
    {
        private bool IsPolicyLocked(string key)
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.IsLocked("share", key);
        }

        private void ApplyPolicyDefaultsToSettings()
        {
            if (!PolicyUiHelper.IsPolicyDomainActive(_backendPolicyStatus, "share"))
            {
                return;
            }
            bool policyBool;
            int policyInt;
            string policyString;

            policyString = _backendPolicyStatus.GetPolicyString("share", "share_base_directory");
            if (!string.IsNullOrWhiteSpace(policyString))
            {
                _request.BasePath = policyString;
            }

            policyString = _backendPolicyStatus.GetPolicyString("share", "share_name_template");
            if (!string.IsNullOrWhiteSpace(policyString))
            {
                _defaults.SharingDefaultShareName = policyString;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_upload", out policyBool))
            {
                _defaults.SharingDefaultPermCreate = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_edit", out policyBool))
            {
                _defaults.SharingDefaultPermWrite = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_delete", out policyBool))
            {
                _defaults.SharingDefaultPermDelete = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_set_password", out policyBool))
            {
                _defaults.SharingDefaultPasswordEnabled = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_send_password_separately", out policyBool))
            {
                _defaults.SharingDefaultPasswordSeparateEnabled = policyBool;
            }
            policyString = _backendPolicyStatus.GetPolicyString("share", "share_send_password_mode");
            if (IsPolicyLocked("share_send_password_mode")
                && _backendPolicyStatus.HasPolicyKey("share", "share_send_password_mode"))
            {
                _defaults.SharingDefaultPasswordDeliveryMode = SharePasswordDeliveryPolicy.ParseMode(policyString);
            }
            if (!PolicyUiHelper.HasBackendSeatEntitlement(_backendPolicyStatus))
            {
                _defaults.SharingDefaultPasswordSeparateEnabled = false;
            }
            if (_backendPolicyStatus.TryGetPolicyInt("share", "share_expire_days", out policyInt))
            {
                _defaults.SharingDefaultExpireDays = Math.Max(1, policyInt);
            }
        }

        private void ApplyPolicyWarningUi()
        {
            PolicyUiHelper.ApplyPolicyWarningState(
                _backendPolicyStatus,
                _policyWarningPanel,
                _policyWarningTextLabel);
            LayoutPolicyWarningPanel();
            UpdateStepHostBounds();
            LayoutCurrentStep();
            LayoutProgressPanel();
        }

        private void ApplyPolicyLockState()
        {
            bool lockShareName = IsPolicyLocked("share_name_template");
            bool lockPermCreate = IsPolicyLocked("share_permission_upload");
            bool lockPermWrite = IsPolicyLocked("share_permission_edit");
            bool lockPermDelete = IsPolicyLocked("share_permission_delete");
            bool lockPassword = IsPolicyLocked("share_set_password");
            bool lockPasswordSeparate = IsPolicyLocked("share_send_password_separately");
            bool lockPasswordDeliveryMode = IsPolicyLocked("share_send_password_mode");
            bool lockExpireDays = IsPolicyLocked("share_expire_days");
            bool separatePasswordAvailable = PolicyUiHelper.HasBackendSeatEntitlement(_backendPolicyStatus);
            string separatePasswordUnavailableTooltip = PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(_backendPolicyStatus);
            bool passwordDeliveryModeAvailable = PolicyUiHelper.HasPasswordDeliveryMode(_backendPolicyStatus);
            string passwordDeliveryUnavailableTooltip = PolicyUiHelper.GetPasswordDeliveryModeUnavailableTooltip(_backendPolicyStatus);

            _shareNameTextBox.ReadOnly = _attachmentMode || lockShareName;
            _permissionCreateCheckBox.Enabled = !_attachmentMode && !lockPermCreate;
            _permissionWriteCheckBox.Enabled = !_attachmentMode && !lockPermWrite;
            _permissionDeleteCheckBox.Enabled = !_attachmentMode && !lockPermDelete;
            _passwordToggleCheckBox.Enabled = !lockPassword;
            _expireToggleCheckBox.Enabled = !lockExpireDays;

            _disabledTooltipHints.Apply(_shareNameTextBox, lockShareName ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareName, _shareNameLabel, _titleLabel);
            _disabledTooltipHints.Apply(_permissionCreateCheckBox, lockPermCreate ? Strings.PolicyAdminControlledTooltip : string.Empty, lockPermCreate);
            _disabledTooltipHints.Apply(_permissionWriteCheckBox, lockPermWrite ? Strings.PolicyAdminControlledTooltip : string.Empty, lockPermWrite);
            _disabledTooltipHints.Apply(_permissionDeleteCheckBox, lockPermDelete ? Strings.PolicyAdminControlledTooltip : string.Empty, lockPermDelete);
            _disabledTooltipHints.Apply(
                _passwordToggleCheckBox,
                lockPassword ? Strings.PolicyAdminControlledTooltip : string.Empty,
                lockPassword,
                (Control)null,
                _passwordGenerateButton,
                _passwordTextBox);
            _disabledTooltipHints.Apply(
                _passwordSeparateToggleCheckBox,
                !separatePasswordAvailable
                    ? separatePasswordUnavailableTooltip
                    : (lockPasswordSeparate ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !separatePasswordAvailable || lockPasswordSeparate,
                (Control)null,
                _passwordTextBox);
            _disabledTooltipHints.Apply(
                _passwordDeliveryModeCombo,
                !passwordDeliveryModeAvailable
                    ? passwordDeliveryUnavailableTooltip
                    : (lockPasswordDeliveryMode
                        ? Strings.PolicyAdminControlledTooltip
                        : (!_passwordSeparateToggleCheckBox.Checked ? Strings.SharingPasswordDeliveryEnableSeparateTooltip : string.Empty)),
                !passwordDeliveryModeAvailable
                    || lockPasswordDeliveryMode
                    || !_passwordSeparateToggleCheckBox.Checked,
                _passwordDeliveryModeLabel);
            _disabledTooltipHints.Apply(_expireToggleCheckBox, lockExpireDays ? Strings.PolicyAdminControlledTooltip : string.Empty, lockExpireDays, _expireHintLabel);
            _disabledTooltipHints.Apply(_expireDatePicker, lockExpireDays ? Strings.PolicyAdminControlledTooltip : string.Empty, false, _expireHintLabel);

            UpdatePasswordState();
            UpdateExpireState();
        }

        private static void OpenPolicyAdminGuide()
        {
            BrowserLauncher.OpenUrl(
                Strings.PolicyAdminGuideUrl,
                LogCategories.FileLink,
                "Failed to open policy admin guide URL.");
        }
    }
}

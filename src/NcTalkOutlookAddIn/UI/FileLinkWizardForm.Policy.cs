// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
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
            _request.BasePath = string.IsNullOrWhiteSpace(_defaults.FileLinkBasePath)
                ? AddinSettings.DefaultFileLinkBasePath : _defaults.FileLinkBasePath;
            if (!PolicyUiHelper.HasBackendSeatEntitlement(_backendPolicyStatus))
            {
                _defaults.SharingDefaultPasswordSeparateEnabled = false;
            }
        }

        private void ApplyPolicyWarningUi()
        {
            PolicyUiHelper.ApplyPolicyWarningState(
                _backendPolicyStatus,
                _policyWarningPanel,
                _policyWarningTextLabel,
                _policyWarningTitleLabel,
                _policyWarningLinkLabel,
                _configuration != null ? _configuration.BaseUrl : string.Empty);
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
            bool lockExpireDays = IsPolicyLocked("share_expire_days");

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
            _disabledTooltipHints.Apply(_expireToggleCheckBox, lockExpireDays ? Strings.PolicyAdminControlledTooltip : string.Empty, lockExpireDays, _expireHintLabel);
            _disabledTooltipHints.Apply(_expireDatePicker, lockExpireDays ? Strings.PolicyAdminControlledTooltip : string.Empty, false, _expireHintLabel);

            UpdatePasswordState();
            UpdateExpireState();
        }

    }
}

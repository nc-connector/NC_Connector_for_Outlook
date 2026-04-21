// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
        // Language override and talk room type option handling for Settings UI.
    internal sealed partial class SettingsForm
    {
        private sealed class LanguageOption
        {
            internal LanguageOption(string value, string label, bool enabled = true)
            {
                Value = value ?? string.Empty;
                Label = label ?? value ?? string.Empty;
                Enabled = enabled;
            }

            internal string Value { get; private set; }

            internal string Label { get; private set; }

            internal bool Enabled { get; private set; }

            public override string ToString()
            {
                return Label;
            }
        }

        private sealed class TalkRoomTypeOption
        {
            internal TalkRoomTypeOption(TalkRoomType value, string label)
            {
                Value = value;
                Label = label ?? value.ToString();
            }

            internal TalkRoomType Value { get; private set; }

            internal string Label { get; private set; }

            public override string ToString()
            {
                return Label;
            }
        }

        private static string NormalizeLanguageChoice(string value)
        {
            if (string.Equals((value ?? string.Empty).Trim(), "custom", StringComparison.OrdinalIgnoreCase))
            {
                return "custom";
            }
            return Strings.NormalizeLanguageOverride(value);
        }

        private bool IsCustomLanguageModeAvailable(string domain)
        {
            if (_backendPolicyStatus == null
                || !_backendPolicyStatus.EndpointAvailable
                || !PolicyUiHelper.IsPolicyActive(_backendPolicyStatus))
            {
                return false;
            }
            string normalizedDomain = string.Equals(domain, "talk", StringComparison.OrdinalIgnoreCase)
                ? "talk"
                : "share";
            string languageKey = normalizedDomain == "talk" ? "language_talk_description" : "language_share_html_block";
            string templateKey = normalizedDomain == "talk" ? "talk_invitation_template" : "share_html_block_template";
            string languageValue = NormalizeLanguageChoice(_backendPolicyStatus.GetPolicyString(normalizedDomain, languageKey));
            string templateValue = _backendPolicyStatus.GetPolicyString(normalizedDomain, templateKey);
            return string.Equals(languageValue, "custom", StringComparison.OrdinalIgnoreCase)
                   && !string.IsNullOrWhiteSpace(templateValue);
        }

        private static string GetLanguageLabel(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return string.Empty;
            }
            try
            {
                string cultureCode = code.Replace('_', '-');
                return CultureInfo.GetCultureInfo(cultureCode).NativeName;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve language label for '" + code + "'.", ex);
                return code;
            }
        }

        private void PopulateLanguageOverrideCombo(ComboBox combo, string domain)
        {
            // Dialog lifecycle guard: combo controls can be null during partial init/dispose transitions.
            if (combo == null)
            {
                return;
            }

            combo.Items.Clear();
            foreach (string code in Strings.SupportedLanguageOverrideCodes)
            {
                if (string.Equals(code, "default", StringComparison.OrdinalIgnoreCase))
                {
                    combo.Items.Add(new LanguageOption("default", Strings.LanguageOverrideDefaultOption));
                    continue;
                }

                combo.Items.Add(new LanguageOption(code, GetLanguageLabel(code)));
            }
            if (_backendPolicyStatus != null && _backendPolicyStatus.EndpointAvailable)
            {
                combo.Items.Add(new LanguageOption("custom", Strings.LanguageOverrideCustomOption, IsCustomLanguageModeAvailable(domain)));
            }
            if (combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }
        }

                // Rebuild language override combos so the `custom` option only exists
        // when the backend endpoint is available and only becomes selectable
        // once the corresponding backend policy actually uses a custom template.
        private void RefreshLanguageOverrideCombos(string shareValue, string talkValue)
        {
            PopulateLanguageOverrideCombo(_shareBlockLangCombo, "share");
            PopulateLanguageOverrideCombo(_eventDescriptionLangCombo, "talk");
            SelectLanguageChoice(_shareBlockLangCombo, shareValue);
            SelectLanguageChoice(_eventDescriptionLangCombo, talkValue);
        }

        private static void SelectLanguageChoice(ComboBox combo, string value)
        {            if (combo == null)
            {
                return;
            }
            string normalized = NormalizeLanguageChoice(value);
            foreach (var item in combo.Items)
            {
                var option = item as LanguageOption;
                if (option != null
                    && option.Enabled
                    && string.Equals(option.Value, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedItem = option;
                    combo.Tag = option.Value;
                    return;
                }
            }
            if (combo.Items.Count > 0)
            {
                foreach (var item in combo.Items)
                {
                    var option = item as LanguageOption;                    if (option != null && option.Enabled)
                    {
                        combo.SelectedItem = option;
                        combo.Tag = option.Value;
                        return;
                    }
                }
                combo.SelectedIndex = 0;
            }
        }

        private static string GetSelectedLanguageChoice(ComboBox combo)
        {            if (combo == null)
            {
                return "default";
            }
            var selected = combo.SelectedItem as LanguageOption;
            return selected != null && selected.Enabled ? NormalizeLanguageChoice(selected.Value) : "default";
        }

        private void HandleLanguageComboDrawItem(object sender, DrawItemEventArgs e)
        {
            ComboBox combo = sender as ComboBox;            if (combo == null || e.Index < 0 || e.Index >= combo.Items.Count)
            {
                return;
            }

            e.DrawBackground();
            var option = combo.Items[e.Index] as LanguageOption;
            string label = option != null ? option.Label : Convert.ToString(combo.Items[e.Index], CultureInfo.InvariantCulture);
            bool enabled = option == null || option.Enabled;
            Color foreColor = enabled ? e.ForeColor : SystemColors.GrayText;
            TextRenderer.DrawText(e.Graphics, label ?? string.Empty, e.Font, e.Bounds, foreColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            e.DrawFocusRectangle();
        }

        private void HandleLanguageComboSelectionCommitted(object sender, EventArgs e)
        {
            ComboBox combo = sender as ComboBox;            if (combo == null)
            {
                return;
            }
            var selected = combo.SelectedItem as LanguageOption;            if (selected == null)
            {
                return;
            }
            if (!selected.Enabled)
            {
                SelectLanguageChoice(combo, Convert.ToString(combo.Tag, CultureInfo.InvariantCulture) ?? "default");
                return;
            }

            combo.Tag = selected.Value;
        }

        private void SelectTalkRoomType(TalkRoomType value)
        {
            foreach (var item in _talkDefaultRoomTypeCombo.Items)
            {
                var option = item as TalkRoomTypeOption;                if (option != null && option.Value == value)
                {
                    _talkDefaultRoomTypeCombo.SelectedItem = option;
                    return;
                }
            }
            if (_talkDefaultRoomTypeCombo.Items.Count > 0)
            {
                _talkDefaultRoomTypeCombo.SelectedIndex = 0;
            }
        }

        private TalkRoomType GetSelectedTalkRoomType()
        {
            var selected = _talkDefaultRoomTypeCombo.SelectedItem as TalkRoomTypeOption;
            return selected != null ? selected.Value : TalkRoomType.StandardRoom;
        }

        private void UpdateTalkRoomTypeTooltip()
        {
            if (IsPolicyLocked("talk", "talk_room_type"))
            {
                _disabledTooltipHints.Apply(
                    _talkDefaultRoomTypeCombo,
                    Strings.PolicyAdminControlledTooltip,
                    true,
                    _talkDefaultRoomTypeLabel);
                return;
            }
            var selected = _talkDefaultRoomTypeCombo.SelectedItem as TalkRoomTypeOption;
            TalkRoomType roomType = selected != null ? selected.Value : TalkRoomType.EventConversation;
            _disabledTooltipHints.Apply(
                _talkDefaultRoomTypeCombo,
                roomType == TalkRoomType.EventConversation ? Strings.TooltipRoomTypeEvent : Strings.TooltipRoomTypeStandard,
                false,
                _talkDefaultRoomTypeLabel);
        }
    }
}


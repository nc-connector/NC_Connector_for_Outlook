/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Localization helper for UI-visible strings.
     *
     * Translations are stored as embedded JSON resources under:
     * - Resources/_locales/<language>/messages.json
     *
     * Supported language codes:
     * - de (default), en, fr, cs, es, hu, it, ja, nl, pl, pt_BR, pt_PT, ru, zh_CN, zh_TW
     */
    internal static class Strings
    {
        private const string DefaultLanguageCode = "de";
        private const string EnglishLanguageCode = "en";

        private static readonly string[] SupportedLanguages = new[]
        {
            "de",
            "en",
            "fr",
            "cs",
            "es",
            "hu",
            "it",
            "ja",
            "nl",
            "pl",
            "pt_BR",
            "pt_PT",
            "ru",
            "zh_CN",
            "zh_TW"
        };

        private static readonly string[] SupportedLanguageOverrides = new[]
        {
            "default",
            "en",
            "de",
            "fr",
            "cs",
            "es",
            "hu",
            "it",
            "ja",
            "nl",
            "pl",
            "pt_BR",
            "pt_PT",
            "ru",
            "zh_CN",
            "zh_TW"
        };

        internal static IReadOnlyList<string> SupportedLanguageCodes
        {
            get { return SupportedLanguages; }
        }

        internal static IReadOnlyList<string> SupportedLanguageOverrideCodes
        {
            get { return SupportedLanguageOverrides; }
        }

        private static readonly HashSet<string> SupportedLanguageSet = new HashSet<string>(SupportedLanguages, StringComparer.OrdinalIgnoreCase);

        private static readonly object InitLock = new object();
        private static bool _initialized;
        private static string _preferredUiLanguageCode;
        private static string[] _languageCandidates = new[] { DefaultLanguageCode, EnglishLanguageCode };

        private static readonly object TranslationLock = new object();
        private static readonly Dictionary<string, Dictionary<string, string>> TranslationCache = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        private sealed class LocaleMessage
        {
            public string message { get; set; }
        }

        private static void EnsureInitialized()
        {
            if (_initialized)
            {
                return;
            }

            lock (InitLock)
            {
                if (_initialized)
                {
                    return;
                }

                _languageCandidates = BuildLanguageCandidates();
                if (DiagnosticsLogger.IsEnabled)
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "UI language candidates: " + string.Join(", ", _languageCandidates));
                }
                _initialized = true;
            }
        }

        private static string[] BuildLanguageCandidates()
        {
            var list = new List<string>();

            if (!string.IsNullOrWhiteSpace(_preferredUiLanguageCode))
            {
                AddLanguageCandidate(list, _preferredUiLanguageCode);
            }

            TryAddCulture(list, () => CultureInfo.CurrentUICulture);
            TryAddCulture(list, () => CultureInfo.CurrentCulture);
            TryAddCulture(list, () => CultureInfo.InstalledUICulture);

            AddLanguageCandidate(list, DefaultLanguageCode);
            AddLanguageCandidate(list, EnglishLanguageCode);

            return list.ToArray();
        }

        internal static void SetPreferredUiLanguage(string languageCode)
        {
            try
            {
                string normalized = NormalizeLanguageCode(languageCode);
                if (string.IsNullOrEmpty(normalized) || !SupportedLanguageSet.Contains(normalized))
                {
                    normalized = null;
                }

                lock (InitLock)
                {
                    _preferredUiLanguageCode = normalized;
                    _initialized = false;
                }

                if (DiagnosticsLogger.IsEnabled && !string.IsNullOrEmpty(normalized))
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "Preferred UI language: " + normalized);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Strings.SetPreferredUiLanguage failed.", ex);
            }
        }

        private static void TryAddCulture(List<string> list, Func<CultureInfo> provider)
        {
            try
            {
                CultureInfo culture = provider();
                if (culture == null)
                {
                    return;
                }

                string normalized = NormalizeLanguageCode(culture.Name);
                AddLanguageCandidate(list, normalized);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Strings.TryAddCulture failed.", ex);
            }
        }

        private static void AddLanguageCandidate(List<string> list, string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return;
            }

            if (!SupportedLanguageSet.Contains(code))
            {
                return;
            }

            if (!list.Contains(code))
            {
                list.Add(code);
            }
        }

        private static string NormalizeLanguageCode(string languageCode)
        {
            if (string.IsNullOrWhiteSpace(languageCode))
            {
                return string.Empty;
            }

            string normalized = languageCode.Trim().Replace('-', '_');
            string[] parts = normalized.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                return string.Empty;
            }

            string language = parts[0].ToLowerInvariant();

            if (string.Equals(language, "pt", StringComparison.OrdinalIgnoreCase))
            {
                string region = parts.Length > 1 ? parts[1] : "PT";
                region = region.ToUpperInvariant();
                if (string.Equals(region, "BR", StringComparison.OrdinalIgnoreCase))
                {
                    return "pt_BR";
                }

                return "pt_PT";
            }

            if (string.Equals(language, "zh", StringComparison.OrdinalIgnoreCase))
            {
                string second = parts.Length > 1 ? parts[1] : string.Empty;
                string third = parts.Length > 2 ? parts[2] : string.Empty;

                if (!string.IsNullOrEmpty(third) &&
                    (string.Equals(second, "Hans", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(second, "Hant", StringComparison.OrdinalIgnoreCase)))
                {
                    second = third;
                }

                if (string.Equals(second, "Hans", StringComparison.OrdinalIgnoreCase))
                {
                    second = "CN";
                }
                else if (string.Equals(second, "Hant", StringComparison.OrdinalIgnoreCase))
                {
                    second = "TW";
                }

                if (string.Equals(second, "TW", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(second, "HK", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(second, "MO", StringComparison.OrdinalIgnoreCase))
                {
                    return "zh_TW";
                }

                if (string.Equals(second, "CN", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(second, "SG", StringComparison.OrdinalIgnoreCase))
                {
                    return "zh_CN";
                }

                return "zh_CN";
            }

            return language;
        }

        internal static string NormalizeLanguageOverride(string languageOverride)
        {
            if (string.IsNullOrWhiteSpace(languageOverride))
            {
                return "default";
            }

            string trimmed = languageOverride.Trim();
            if (string.Equals(trimmed, "default", StringComparison.OrdinalIgnoreCase))
            {
                return "default";
            }

            string normalized = NormalizeLanguageCode(trimmed);
            if (string.IsNullOrEmpty(normalized) || !SupportedLanguageSet.Contains(normalized))
            {
                return "default";
            }

            return normalized;
        }

        private static Dictionary<string, string> GetTranslationsForLanguage(string languageCode)
        {
            if (string.IsNullOrEmpty(languageCode))
            {
                return null;
            }

            lock (TranslationLock)
            {
                Dictionary<string, string> cached;
                if (TranslationCache.TryGetValue(languageCode, out cached))
                {
                    return cached;
                }

                Dictionary<string, string> loaded = LoadTranslations(languageCode) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                TranslationCache[languageCode] = loaded;
                return loaded;
            }
        }

        private static Dictionary<string, string> LoadTranslations(string languageCode)
        {
            try
            {
                string resourceName = BuildLocaleResourceName(languageCode);
                Assembly assembly = typeof(Strings).Assembly;

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        DiagnosticsLogger.Log(LogCategories.Core, "Locale resource not found: " + resourceName);
                        return null;
                    }

                    using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                    {
                        string json = reader.ReadToEnd();
                        var serializer = new JavaScriptSerializer();
                        var parsed = serializer.Deserialize<Dictionary<string, LocaleMessage>>(json);

                        var dictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        if (parsed != null)
                        {
                            foreach (var item in parsed)
                            {
                                if (item.Value == null || string.IsNullOrEmpty(item.Value.message))
                                {
                                    continue;
                                }

                                dictionary[item.Key] = ConvertPlaceholders(item.Value.message);
                            }
                        }

                        return dictionary;
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load translations for '" + (languageCode ?? string.Empty) + "'.", ex);
                return null;
            }
        }

        private static string BuildLocaleResourceName(string languageCode)
        {
            string assemblyName = typeof(Strings).Assembly.GetName().Name ?? "NcTalkOutlookAddIn";
            return assemblyName + ".Resources._locales." + languageCode + ".messages.json";
        }

        private static string ConvertPlaceholders(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return message;
            }

            // WebExtension i18n uses "$1", "$2", ... placeholders. Convert to string.Format placeholders.
            for (int i = 9; i >= 1; i--)
            {
                message = message.Replace(
                    "$" + i.ToString(CultureInfo.InvariantCulture),
                    "{" + (i - 1).ToString(CultureInfo.InvariantCulture) + "}");
            }

            return message;
        }

        private static string Get(string key, string defaultValue)
        {
            try
            {
                EnsureInitialized();

                foreach (string code in _languageCandidates)
                {
                    Dictionary<string, string> dictionary = GetTranslationsForLanguage(code);
                    if (dictionary == null)
                    {
                        continue;
                    }

                    string value;
                    if (dictionary.TryGetValue(key, out value) && !string.IsNullOrEmpty(value))
                    {
                        return value;
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Strings.Get failed for key '" + (key ?? string.Empty) + "'.", ex);
            }

            return defaultValue;
        }

        internal static string GetInLanguage(string languageOverride, string key, string defaultValue)
        {
            try
            {
                EnsureInitialized();

                string normalized = NormalizeLanguageOverride(languageOverride);
                if (string.IsNullOrEmpty(normalized) || string.Equals(normalized, "default", StringComparison.OrdinalIgnoreCase))
                {
                    return Get(key, defaultValue);
                }

                foreach (string code in new[] { normalized, DefaultLanguageCode, EnglishLanguageCode })
                {
                    Dictionary<string, string> dictionary = GetTranslationsForLanguage(code);
                    if (dictionary == null)
                    {
                        continue;
                    }

                    string value;
                    if (dictionary.TryGetValue(key, out value) && !string.IsNullOrEmpty(value))
                    {
                        return value;
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Strings.GetInLanguage failed for key '" + (key ?? string.Empty) + "'.", ex);
            }

            return defaultValue;
        }

        private static string BuildTooltip(string title, params string[] lines)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(title))
            {
                parts.Add(title.Trim());
            }

            if (lines != null)
            {
                foreach (string line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        parts.Add(line.Trim());
                    }
                }
            }

            return string.Join("\n\n", parts.ToArray());
        }

        // Settings UI
        internal static string SettingsFormTitle { get { return Get("options_title", "NC Connector for Outlook - Settings"); } }
        internal static string TabGeneral { get { return Get("options_tab_general", "General"); } }
        internal static string TabIfb { get { return Get("outlook_tab_ifb", "IFB"); } }
        internal static string TabTalkLink { get { return Get("options_tab_talklink", "Talk Link"); } }
        internal static string TabAdvanced { get { return Get("options_tab_advanced", "Advanced"); } }
        internal static string TabDebug { get { return Get("options_tab_debug", "Debug"); } }
        internal static string TabAbout { get { return Get("options_tab_about", "About"); } }
        internal static string TabFileLink { get { return Get("options_tab_sharing", "Sharing"); } }

        internal static string SettingsTalkDefaultsGroup { get { return Get("options_talk_defaults_heading", "Default settings"); } }
        internal static string LabelServerUrl { get { return Get("options_base_url_label", "Nextcloud URL"); } }
        internal static string LabelUsername { get { return Get("options_user_label", "Username"); } }
        internal static string LabelAppPassword { get { return Get("options_app_pass_label", "App password"); } }
        internal static string GroupAuthentication { get { return Get("options_auth_heading", "Authentication"); } }
        internal static string RadioManual { get { return Get("options_auth_manual", "Manual (enter username + app password)"); } }
        internal static string RadioLoginFlow { get { return Get("options_auth_login", "Login with Nextcloud (fetch app password automatically)"); } }
        internal static string ButtonLoginFlow { get { return Get("options_loginflow_button", "Login with Nextcloud..."); } }
        internal static string ButtonTestConnection { get { return Get("options_test_button", "Test connection"); } }

        internal static string CheckIfbEnabled { get { return Get("outlook_ifb_enabled", "Enable IFB endpoint"); } }
        internal static string LabelIfbDays { get { return Get("outlook_ifb_days_label", "Range (days):"); } }
        internal static string LabelIfbCacheHours { get { return Get("outlook_ifb_cache_hours_label", "Address book cache (hours):"); } }

        internal static string AdvancedShareBlockLangLabel { get { return Get("options_share_block_lang_label", "Language for sharing HTML block"); } }
        internal static string AdvancedEventDescriptionLangLabel { get { return Get("options_event_description_lang_label", "Language for Talk description text"); } }
        internal static string LanguageOverrideDefaultOption { get { return Get("options_lang_default", "Default (UI)"); } }

        internal static string ButtonSave { get { return Get("options_save_button", "Save"); } }
        internal static string ButtonCancel { get { return Get("ui_button_cancel", "Cancel"); } }

        internal static string DebugCheckbox { get { return Get("outlook_debug_file_label", "Write debug log file"); } }
        internal static string DebugPathPrefix { get { return Get("outlook_debug_location_prefix", "Location: "); } }
        internal static string DebugOpenLog { get { return Get("outlook_debug_open_log", "Open log file"); } }
        internal static string DebugLogMissingMessage { get { return Get("outlook_debug_log_missing", "No log file has been created yet."); } }
        internal static string DebugLogOpenErrorMessage { get { return Get("outlook_debug_log_open_error", "Could not open log file."); } }

        internal static string ButtonNext { get { return Get("sharing_button_next", "Next"); } }
        internal static string ButtonBack { get { return Get("sharing_button_back", "Back"); } }

        // Talk wizard
        internal static string TalkFormTitle { get { return Get("talk_button_create_room", "Create Talk room"); } }
        internal static string TalkTitleLabel { get { return Get("ui_create_title_label", "Title"); } }
        internal static string TalkPasswordSetCheck { get { return Get("talk_password_toggle_label", "Set password"); } }
        internal static string TalkPasswordLabel { get { return Get("sharing_password_label", "Password"); } }
        internal static string TalkPasswordGenerate { get { return Get("ui_password_generate_label", "Generate"); } }
        internal static string TalkLobbyCheck { get { return Get("ui_create_lobby_label", "Lobby until start time"); } }
        internal static string TalkSearchCheck { get { return Get("ui_create_listable_label", "Show in search"); } }
        internal static string TalkSettingsGroup { get { return Get("talk_section_settings", "Talk settings"); } }
        internal static string TalkAddUsersCheck { get { return Get("ui_create_add_users_label", "Add users"); } }
        internal static string TalkAddGuestsCheck { get { return Get("ui_create_add_guests_label", "Add guests"); } }
        internal static string TalkRoomGroup { get { return Get("ui_create_roomtype_label", "Room type"); } }
        internal static string TalkEventRadio { get { return Get("ui_create_mode_event", "Event conversation"); } }
        internal static string TalkStandardRadio { get { return Get("ui_create_mode_standard", "Group conversation"); } }
        internal static string TalkModeratorGroup { get { return Get("ui_create_moderator_label", "Moderator (optional)"); } }
        internal static string TalkModeratorClear { get { return Get("ui_button_clear", "Clear"); } }
        internal static string TalkModeratorHint { get { return Get("ui_create_moderator_hint", "If provided, moderation is transferred to this user after creation and you leave the room."); } }
        internal static string TalkModeratorHintNoDirectory { get { return Get("outlook_moderator_hint_no_directory", "User directory unavailable. Please enter the username manually."); } }
        internal static string TalkModeratorNoMatches { get { return Get("ui_delegate_status_none_with_email", "No matches."); } }
        internal static string TalkVersionUnknown { get { return Get("outlook_version_unknown", "unknown"); } }
        internal static string TalkEventHint { get { return Get("outlook_event_conversation_hint", "Event conversations require Nextcloud 31 or newer.\nDetected version: {0}"); } }

        internal static string DialogOk { get { return Get("ui_button_ok", "OK"); } }
        internal static string DialogCancel { get { return Get("ui_button_cancel", "Cancel"); } }

        internal static string TalkPasswordTooShort { get { return Get("talk_password_policy_error", "Password must be at least {0} characters long."); } }
        internal static string TalkDefaultTitle { get { return Get("ui_default_title", "Meeting"); } }

        internal static string TooltipAddUsers
        {
            get
            {
                return BuildTooltip(
                    Get("talk_option_add_users_title", "👤 Add users"),
                    Get("talk_option_add_users_line1", "Internal Nextcloud users are added directly to the room"),
                    Get("talk_option_add_users_line2", "The room is immediately visible in Nextcloud Talk"));
            }
        }

        internal static string TooltipAddGuests
        {
            get
            {
                return BuildTooltip(
                    Get("talk_option_add_guests_title", "🌍 Add guests"),
                    Get("talk_option_add_guests_line1", "Guests are also added directly to the room"),
                    Get("talk_option_add_guests_line2", "⚠️ They will receive a separate (additional) email with a personal access link"));
            }
        }

        internal static string TooltipSearchVisible
        {
            get
            {
                return BuildTooltip(
                    Get("talk_option_listable_title", "🔍 Show in search"),
                    Get("talk_option_listable_line1", "The room can be found via Talk search"),
                    Get("talk_option_listable_line2", "⚠️ Makes it easier to find again, but is discoverable for all users."));
            }
        }

        internal static string TooltipLobby
        {
            get
            {
                return BuildTooltip(
                    Get("talk_option_lobby_title", "🚪 Lobby until start time"),
                    Get("talk_option_lobby_line1", "Participants are placed in the lobby (waiting room) before the start and cannot join the call yet"),
                    Get("talk_option_lobby_line2", "⚠️ Joining is only possible from the configured start time (or when a moderator ends the lobby)"),
                    Get("talk_option_lobby_line3", "Ideal so nobody joins the meeting/webinar too early."));
            }
        }

        internal static string TooltipRoomTypeEvent
        {
            get
            {
                return BuildTooltip(
                    Get("talk_roomtype_event_title", "📅 Event conversation"),
                    Get("talk_roomtype_event_line1", "Time-limited: for a specific appointment or occasion"),
                    Get("talk_roomtype_event_line2", "Spontaneous & lightweight: suitable for meetings, trainings or webinars"),
                    Get("talk_roomtype_event_line3", "Self-cleaning: removed automatically after some time"));
            }
        }

        internal static string TooltipRoomTypeStandard
        {
            get
            {
                return BuildTooltip(
                    Get("talk_roomtype_group_title", "🗨️ Group conversation"),
                    Get("talk_roomtype_group_line1", "Persistent: remains until it is actively deleted"),
                    Get("talk_roomtype_group_line2", "Everyday use: ideal for teams, departments, ongoing projects and recurring appointments"),
                    Get("talk_roomtype_group_line3", "With history: the entire chat can be reviewed at any time"));
            }
        }

        internal static string StatusLoginFlowStarting { get { return Get("options_loginflow_starting", "Starting login flow..."); } }
        internal static string StatusLoginFlowBrowser { get { return Get("options_loginflow_browser", "Browser opened. Please confirm the login in Nextcloud..."); } }
        internal static string StatusLoginFlowSuccess { get { return Get("options_loginflow_success", "Login successful. App password copied."); } }
        internal static string StatusLoginFlowFailure { get { return Get("outlook_loginflow_failed_with_reason", "Login failed: {0}"); } }
        internal static string StatusMissingFields { get { return Get("options_test_missing", "Please enter URL, username and app password first."); } }
        internal static string StatusServerUrlRequired { get { return Get("options_loginflow_missing", "Please enter the server URL first."); } }
        internal static string StatusInvalidServerUrl { get { return Get("outlook_status_invalid_server_url", "Please enter a valid server URL."); } }
        internal static string StatusTestRunning { get { return Get("options_test_running", "Testing connection..."); } }
        internal static string StatusTestSuccessFormat { get { return Get("outlook_test_success_format", "Connection successful{0}"); } }
        internal static string StatusTestSuccessVersionFormat { get { return Get("outlook_test_version_format", "Nextcloud {0}"); } }
        internal static string StatusTestFailureUnknown { get { return Get("outlook_test_failure_unknown", "Unknown error"); } }
        internal static string StatusTestFailure { get { return Get("outlook_test_failure_format", "Connection failed: {0}"); } }

        internal static string ErrorMissingCredentials { get { return Get("error_credentials_missing", "Please configure server URL, username, and app password in the settings first."); } }
        internal static string ErrorNoAppointment { get { return Get("outlook_error_no_appointment", "Could not determine the current appointment item."); } }
        internal static string ConfirmReplaceRoom { get { return Get("outlook_confirm_replace_room", "A Talk room already exists for this appointment. Replace it?"); } }
        internal static string ErrorCreateRoom { get { return Get("ui_create_failed", "Could not create Nextcloud Talk:\n{0}\nPlease check the options."); } }
        internal static string ErrorCreateRoomUnexpected { get { return Get("outlook_error_create_room_unexpected_format", "Unexpected error while creating the Talk room: {0}"); } }
        internal static string InfoRoomCreated { get { return Get("ui_alert_room_created", "Talk room \"{0}\" has been created."); } }
        internal static string PromptOpenSettings { get { return Get("outlook_prompt_open_settings", "{0}\n\nOpen settings now?"); } }
        internal static string ErrorServerUnavailable { get { return Get("outlook_error_server_unavailable", "The Nextcloud server is currently unreachable. Please check your Internet connection."); } }
        internal static string ErrorAuthenticationRejected { get { return Get("outlook_error_authentication_rejected_format", "Credentials were not accepted: {0}"); } }
        internal static string ErrorConnectionFailed { get { return Get("outlook_error_connection_failed_format", "Connection to the Nextcloud server failed: {0}"); } }
        internal static string ErrorUnknownAuthentication { get { return Get("outlook_error_unknown_authentication_format", "Unknown error during authentication: {0}"); } }
        internal static string ErrorCredentialsNotVerified { get { return Get("outlook_error_credentials_not_verified", "Credentials could not be verified."); } }
        internal static string ErrorCredentialsNotVerifiedFormat { get { return Get("outlook_error_credentials_not_verified_format", "Credentials could not be verified: {0}"); } }

        internal static string WarningIfbStartFailed { get { return Get("outlook_warning_ifb_start_failed_format", "IFB could not be started: {0}"); } }
        internal static string WarningRoomDeleteFailed { get { return Get("error_room_delete_failed", "Room could not be deleted: {0}"); } }
        internal static string WarningLobbyUpdateFailed { get { return Get("error_lobby_set_failed", "Lobby could not be set: {0}"); } }
        internal static string WarningDescriptionUpdateFailed { get { return Get("outlook_warning_description_update_failed_format", "Room description could not be updated: {0}"); } }
        internal static string WarningModeratorTransferFailed { get { return Get("ui_moderator_transfer_failed", "Could not transfer moderator."); } }
        internal static string WarningModeratorTransferFailedWithReasonFormat { get { return Get("ui_moderator_transfer_failed_with_reason", "Could not transfer moderator:\n{0}"); } }

        // Ribbon / Mail compose
        internal static string RibbonAppointmentGroupLabel { get { return Get("outlook_ribbon_group_label", "NC Connector"); } }
        internal static string RibbonTalkButtonLabel { get { return Get("ui_insert_button_label", "Insert Talk link"); } }
        internal static string RibbonTalkButtonScreenTip { get { return Get("outlook_ribbon_talk_screentip", "Create a Nextcloud Talk conversation"); } }
        internal static string RibbonTalkButtonSuperTip { get { return Get("outlook_ribbon_talk_supertip", "Creates a Nextcloud Talk conversation for this appointment."); } }
        internal static string RibbonExplorerTabLabel { get { return Get("outlook_ribbon_tab_label", "NC Connector"); } }
        internal static string RibbonExplorerGroupLabel { get { return Get("outlook_ribbon_config_group", "Configuration"); } }
        internal static string RibbonSettingsButtonLabel { get { return Get("outlook_ribbon_settings_label", "Settings"); } }
        internal static string RibbonSettingsScreenTip { get { return Get("outlook_ribbon_settings_screentip", "NC Connector settings"); } }
        internal static string RibbonSettingsSuperTip { get { return Get("outlook_ribbon_settings_supertip", "Manage server and credentials."); } }
        internal static string RibbonMailGroupLabel { get { return Get("outlook_ribbon_group_label", "NC Connector"); } }
        internal static string RibbonFileLinkButtonLabel { get { return Get("compose_sharing_title", "Add Nextcloud share"); } }
        internal static string RibbonFileLinkButtonScreenTip { get { return Get("outlook_ribbon_sharing_screentip", "Insert Nextcloud share link"); } }
        internal static string RibbonFileLinkButtonSuperTip { get { return Get("outlook_ribbon_sharing_supertip", "Creates a Nextcloud share link for selected files and inserts it into the message."); } }
        internal static string ErrorNoMailItem { get { return Get("outlook_error_no_mail_item", "No active mail item found."); } }
        internal static string ErrorInsertHtmlFailed { get { return Get("outlook_error_insert_html_failed_format", "Could not insert HTML into the message: {0}"); } }

        // About
        internal static string AboutVersionFormat { get { return Get("outlook_about_version_format", "Version: {0}"); } }
        internal static string AboutLicenseLabel { get { return Get("options_about_license_label", "License:"); } }
        internal static string AboutLicenseLink { get { return Get("options_about_license_value", "AGPL v3"); } }
        internal static string AboutSupportNote { get { return Get("options_about_support_note", "So NC Connector for Outlook can continue to be maintained and improved for free."); } }
        internal static string AboutSupportHeading { get { return Get("options_about_support_heading", "Support"); } }
        internal static string AboutSupportLink { get { return Get("options_support_link", "PayPal: Donate"); } }
        internal static string AboutCopyright { get { return Get("outlook_about_copyright", "Copyright 2025 Bastian Kleinschmidt"); } }
        internal static string LicenseFileMissingMessage { get { return Get("outlook_license_file_missing_format", "The license file could not be found:\r\n{0}"); } }
        internal static string LicenseFileOpenErrorMessage { get { return Get("outlook_license_file_open_error_format", "The license file could not be opened: {0}"); } }

        // Sharing defaults (settings)
        internal static string FileLinkBaseLabel { get { return Get("options_sharing_base_label", "Base directory"); } }
        internal static string FileLinkBaseHint { get { return Get("options_sharing_base_hint", "Files are stored beneath this directory (e.g. \"90 Freigaben - extern\")."); } }
        internal static string SharingDefaultsHeading { get { return Get("options_sharing_defaults_heading", "Default settings"); } }
        internal static string SharingDefaultShareNameLabel { get { return Get("options_sharing_default_share_name_label", "Share name"); } }
        internal static string SharingDefaultPermissionsLabel { get { return Get("options_sharing_default_permissions_label", "Default permissions"); } }
        internal static string SharingDefaultPermCreateLabel { get { return Get("options_sharing_default_perm_create", "Upload/Create"); } }
        internal static string SharingDefaultPermWriteLabel { get { return Get("options_sharing_default_perm_write", "Edit"); } }
        internal static string SharingDefaultPermDeleteLabel { get { return Get("options_sharing_default_perm_delete", "Delete"); } }
        internal static string SharingDefaultPasswordLabel { get { return Get("options_sharing_default_password_label", "Set password"); } }
        internal static string SharingDefaultExpireDaysLabel { get { return Get("options_sharing_default_expire_days_label", "Expiration (days)"); } }

        // Sharing wizard
        internal static string FileLinkNoFilesConfirm { get { return Get("sharing_confirm_no_files_message", "No files or folders were added to this share. Recipients can only upload their own files. Do you want to continue?"); } }
        internal static string FileLinkPermissionRead { get { return Get("sharing_permission_read", "Read"); } }
        internal static string FileLinkPermissionCreate { get { return Get("sharing_permission_create", "Upload"); } }
        internal static string FileLinkPermissionWrite { get { return Get("sharing_permission_write", "Modify"); } }
        internal static string FileLinkPermissionDelete { get { return Get("sharing_permission_delete", "Delete"); } }
        internal static string FileLinkWizardTitle { get { return Get("sharing_wizard_title", "Nextcloud Share"); } }
        internal static string FileLinkWizardStepShare { get { return Get("sharing_heading_permissions", "Share settings"); } }
        internal static string FileLinkWizardStepExpire { get { return Get("sharing_step_expire", "Expiration date"); } }
        internal static string FileLinkWizardStepFiles { get { return Get("sharing_step_files_title", "Select files or folders"); } }
        internal static string FileLinkWizardStepNote { get { return Get("sharing_step_note", "Note & finish"); } }
        internal static string FileLinkWizardShareNameLabel { get { return Get("sharing_share_name_label", "Share name"); } }
        internal static string FileLinkWizardPermissionsLabel { get { return Get("sharing_step_security", "Permissions"); } }
        internal static string FileLinkWizardFallbackShareName { get { return Get("outlook_sharing_fallback_share_name", "Share"); } }
        internal static string FileLinkWizardPasswordToggle { get { return Get("sharing_password_toggle", "Set password"); } }
        internal static string FileLinkWizardExpireToggle { get { return Get("sharing_expire_toggle", "Set expiration date"); } }
        internal static string FileLinkWizardExpireHint { get { return Get("sharing_expire_hint", "After this date the link is no longer available."); } }
        internal static string FileLinkWizardBasePathPrefix { get { return Get("outlook_sharing_base_path_prefix", "Base directory: "); } }
        internal static string FileLinkWizardAddFilesButton { get { return Get("sharing_button_add_files", "Add files..."); } }
        internal static string FileLinkWizardAddFolderButton { get { return Get("sharing_button_add_folder", "Add folder..."); } }
        internal static string FileLinkWizardRemoveButton { get { return Get("sharing_button_remove", "Remove"); } }
        internal static string FileLinkWizardColumnPath { get { return Get("sharing_files_table_path", "Path"); } }
        internal static string FileLinkWizardColumnType { get { return Get("sharing_files_table_type", "Type"); } }
        internal static string FileLinkWizardColumnStatus { get { return Get("sharing_files_table_status", "Status"); } }
        internal static string FileLinkWizardTypeFile { get { return Get("sharing_file_type_file", "File"); } }
        internal static string FileLinkWizardTypeFolder { get { return Get("sharing_file_type_folder", "Folder"); } }
        internal static string FileLinkWizardNoteToggle { get { return Get("sharing_note_toggle", "Add note for recipients"); } }
        internal static string FileLinkWizardUploadButton { get { return Get("sharing_button_upload", "Upload"); } }
        internal static string FileLinkWizardFinishButton { get { return Get("sharing_button_finish", "Create share"); } }
        internal static string FileLinkWizardShareNameRequired { get { return Get("sharing_error_share_name", "Please enter a share name."); } }
        internal static string FileLinkWizardExpireMustBeFuture { get { return Get("outlook_sharing_expire_must_be_future", "The expiration date must be in the future."); } }
        internal static string FileLinkWizardSelectFileOrFolder { get { return Get("outlook_sharing_select_file_or_folder", "Please select at least one file or folder."); } }
        internal static string FileLinkWizardUploadFirst { get { return Get("outlook_sharing_upload_first", "Please upload the selected files first."); } }
        internal static string FileLinkWizardCreatingShare { get { return Get("sharing_progress_label", "Creating share..."); } }
        internal static string FileLinkWizardCreateFailedFormat { get { return Get("outlook_sharing_create_failed_format", "Could not create share: {0}"); } }
        internal static string FileLinkWizardFolderExistsFormat { get { return Get("sharing_error_folder_exists", "A share folder with this name already exists. Please choose a different name."); } }
        internal static string FileLinkWizardFolderCheckFailedFormat { get { return Get("outlook_sharing_folder_check_failed_format", "Could not check share folder: {0}"); } }
        internal static string FileLinkWizardStatusSuccess { get { return Get("sharing_status_done_row", "Done"); } }
        internal static string FileLinkWizardStatusError { get { return Get("sharing_status_error_row", "Error"); } }
        internal static string FileLinkWizardStatusCancelled { get { return Get("outlook_sharing_status_cancelled", "Cancelled"); } }
        internal static string FileLinkWizardUploadCancelledMessage { get { return Get("outlook_sharing_upload_cancelled", "Upload was cancelled."); } }
        internal static string FileLinkWizardUploadFailed { get { return Get("outlook_sharing_upload_failed", "Upload failed."); } }
        internal static string FileLinkWizardUploadFailedFormat { get { return Get("outlook_sharing_upload_failed_format", "Upload failed: {0}"); } }
        internal static string FileLinkWizardRenameFolderTitle { get { return Get("outlook_sharing_rename_folder_title", "Rename folder"); } }
        internal static string FileLinkWizardRenameFileTitle { get { return Get("outlook_sharing_rename_file_title", "Rename file"); } }
        internal static string FileLinkWizardRenameFolderPrompt { get { return Get("outlook_sharing_rename_folder_prompt", "Please enter a new folder name:"); } }
        internal static string FileLinkWizardRenameFilePrompt { get { return Get("outlook_sharing_rename_file_prompt", "Please enter a new file name:"); } }

        // Generic dialog title
        internal static string DialogTitle { get { return Get("extName", "NC Connector for Outlook"); } }
    }
}

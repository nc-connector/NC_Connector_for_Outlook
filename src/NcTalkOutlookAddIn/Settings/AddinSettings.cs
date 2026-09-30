// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    // Persistent add-in settings (credentials, sharing/IFB options, etc.).
    internal class AddinSettings
    {
        internal const int DefaultIfbDays = 30;
        internal const int DefaultIfbCacheHours = 24;
        internal const int DefaultIfbPort = 7777;
        internal const int MinIfbPort = 1024;
        internal const int MaxIfbPort = 49151;
        internal const string DefaultFileLinkBasePath = "NC Connector";
        internal const int DefaultSharingAttachmentsOfferAboveMb = 20;

        private Dictionary<string, object> _localPolicyValues = new Dictionary<string, object>(StringComparer.Ordinal);
        private ManagedSetupPolicy _managedSetupPolicy;
        private string _defaultsSource;
        private AuthenticationMode _localAuthMode;
        private bool _localDebugLoggingEnabled;
        private bool _localLogAnonymizationEnabled;
        private bool _localUpdateNotifyEnabled;
        private bool _localTransportTlsUseSystemDefault;
        private bool _localTransportTlsEnable12;
        private bool _localTransportTlsEnable13;
        private bool _localIfbEnabled;
        private int _localIfbDays;
        private int _localIfbCacheHours;
        private int _localIfbPort;

        public AddinSettings()
        {
            ServerUrl = string.Empty;
            Username = string.Empty;
            AppPassword = string.Empty;
            AuthMode = AuthenticationMode.LoginFlow;
            IfbEnabled = false;
            IfbDays = DefaultIfbDays;
            IfbCacheHours = DefaultIfbCacheHours;
            IfbPort = DefaultIfbPort;
            IfbPreviousFreeBusyPath = string.Empty;
            IfbUserDecisionRecorded = false;
            DebugLoggingEnabled = false;
            LogAnonymizationEnabled = true;
            TransportTlsUseSystemDefault = false;
            TransportTlsEnable12 = true;
            TransportTlsEnable13 = false;
            UpdateNotifyEnabled = false;
            UpdateInstallId = string.Empty;
            UpdateLastCheckedAtUtc = string.Empty;
            UpdateLatestVersion = string.Empty;
            UpdateReleaseUrl = string.Empty;
            UpdateDownloadUrl = string.Empty;
            UpdatePublishedAt = string.Empty;
            UpdateChangelogTitle = string.Empty;
            UpdateChangelogText = string.Empty;
            UpdateLastNotifiedVersion = string.Empty;
            UpdateLastNotifiedDateUtc = string.Empty;
            SharingAttachmentLinkTarget = null;
            EmailSignatureOnCompose = null;
            EmailSignatureOnReply = null;
            EmailSignatureOnForward = null;
            ManagedNextcloudUrl = string.Empty;
            ManagedNextcloudUrlSource = string.Empty;
            ManagedNextcloudUrlLocked = false;
            IsEnterpriseRollout = false;
            ShowMainRibbonTab = true;
        }

        public string ServerUrl { get; set; }

        public string Username { get; set; }

        public string AppPassword { get; set; }

        public AuthenticationMode AuthMode
        {
            get { return HasManagedAuthMode ? _managedSetupPolicy.AuthMode : _localAuthMode; }
            set { _localAuthMode = value; }
        }

        internal AuthenticationMode LocalAuthMode { get { return _localAuthMode; } }
        internal bool HasManagedAuthMode { get { return _managedSetupPolicy != null && _managedSetupPolicy.HasAuthModePolicy; } }
        internal bool IsManagedAuthModeValid { get { return !HasManagedAuthMode || _managedSetupPolicy.IsAuthModePolicyValid; } }

        public string DefaultsSource
        {
            get { return _defaultsSource; }
            set { _defaultsSource = BackendPolicyStatus.NormalizeDefaultsSource(value); }
        }

        internal bool HasManagedDefaultsSource { get { return _managedSetupPolicy != null && _managedSetupPolicy.HasDefaultsSourcePolicy; } }
        internal bool IsManagedDefaultsSourceValid { get { return !HasManagedDefaultsSource || _managedSetupPolicy.IsDefaultsSourcePolicyValid; } }

        internal string ResolveDefaultsSource(BackendPolicyStatus status)
        {
            if (status == null || !status.FetchSucceeded || !PolicyUiHelper.HasBackendSeatEntitlement(status))
            {
                return "local";
            }
            if (status.DefaultsSource != null)
            {
                return status.DefaultsSourceEditable && DefaultsSource != null ? DefaultsSource : status.DefaultsSource;
            }
            return HasManagedDefaultsSource ? _managedSetupPolicy.DefaultsSource : DefaultsSource ?? "local";
        }

        internal bool CanEditDefaultsSource(BackendPolicyStatus status)
        {
            if (status == null || !status.FetchSucceeded || !PolicyUiHelper.HasBackendSeatEntitlement(status))
            {
                return false;
            }
            return status.DefaultsSource != null ? status.DefaultsSourceEditable : !HasManagedDefaultsSource;
        }

        public bool IfbEnabled
        {
            get { return HasManagedIfb ? _managedSetupPolicy.IfbEnabled : _localIfbEnabled; }
            set { _localIfbEnabled = value; }
        }

        public int IfbDays
        {
            get { return HasManagedIfb ? _managedSetupPolicy.IfbDays : _localIfbDays; }
            set { _localIfbDays = value; }
        }

        public int IfbCacheHours
        {
            get { return HasManagedIfb ? _managedSetupPolicy.IfbCacheHours : _localIfbCacheHours; }
            set { _localIfbCacheHours = value; }
        }

        public int IfbPort
        {
            get { return HasManagedIfb ? _managedSetupPolicy.IfbPort : _localIfbPort; }
            set { _localIfbPort = value; }
        }

        internal bool LocalIfbEnabled { get { return _localIfbEnabled; } }
        internal int LocalIfbDays { get { return _localIfbDays; } }
        internal int LocalIfbCacheHours { get { return _localIfbCacheHours; } }
        internal int LocalIfbPort { get { return _localIfbPort; } }
        internal bool HasManagedIfb { get { return _managedSetupPolicy != null && _managedSetupPolicy.HasIfbPolicy; } }
        internal bool IsManagedIfbValid { get { return !HasManagedIfb || _managedSetupPolicy.IsIfbPolicyValid; } }

        public string IfbPreviousFreeBusyPath { get; set; }

        public bool IfbUserDecisionRecorded { get; set; }

        public bool DebugLoggingEnabled
        {
            get { return HasManagedLogging ? _managedSetupPolicy.DebugLoggingEnabled : _localDebugLoggingEnabled; }
            set { _localDebugLoggingEnabled = value; }
        }

        public bool LogAnonymizationEnabled
        {
            get { return HasManagedLogging ? _managedSetupPolicy.LogAnonymizationEnabled : _localLogAnonymizationEnabled; }
            set { _localLogAnonymizationEnabled = value; }
        }

        internal bool LocalDebugLoggingEnabled { get { return _localDebugLoggingEnabled; } }
        internal bool LocalLogAnonymizationEnabled { get { return _localLogAnonymizationEnabled; } }
        internal bool HasManagedLogging { get { return _managedSetupPolicy != null && _managedSetupPolicy.HasLoggingPolicy; } }
        internal bool IsManagedLoggingValid { get { return !HasManagedLogging || _managedSetupPolicy.IsLoggingPolicyValid; } }

        public bool TransportTlsUseSystemDefault
        {
            get { return HasManagedTransportTls ? _managedSetupPolicy.TransportTlsUseSystemDefault : _localTransportTlsUseSystemDefault; }
            set { _localTransportTlsUseSystemDefault = value; }
        }

        public bool TransportTlsEnable12
        {
            get { return HasManagedTransportTls ? _managedSetupPolicy.TransportTlsEnable12 : _localTransportTlsEnable12; }
            set { _localTransportTlsEnable12 = value; }
        }

        public bool TransportTlsEnable13
        {
            get { return HasManagedTransportTls ? _managedSetupPolicy.TransportTlsEnable13 : _localTransportTlsEnable13; }
            set { _localTransportTlsEnable13 = value; }
        }

        internal bool LocalTransportTlsUseSystemDefault { get { return _localTransportTlsUseSystemDefault; } }
        internal bool LocalTransportTlsEnable12 { get { return _localTransportTlsEnable12; } }
        internal bool LocalTransportTlsEnable13 { get { return _localTransportTlsEnable13; } }
        internal bool HasManagedTransportTls { get { return _managedSetupPolicy != null && _managedSetupPolicy.HasTransportTlsPolicy; } }
        internal bool IsManagedTransportTlsValid { get { return !HasManagedTransportTls || _managedSetupPolicy.IsTransportTlsPolicyValid; } }

        public bool UpdateNotifyEnabled
        {
            get { return HasManagedUpdateNotify ? _managedSetupPolicy.UpdateNotifyEnabled : _localUpdateNotifyEnabled; }
            set { _localUpdateNotifyEnabled = value; }
        }

        internal bool LocalUpdateNotifyEnabled { get { return _localUpdateNotifyEnabled; } }
        internal bool HasManagedUpdateNotify { get { return _managedSetupPolicy != null && _managedSetupPolicy.HasUpdateNotifyPolicy; } }
        internal bool IsManagedUpdateNotifyValid { get { return !HasManagedUpdateNotify || _managedSetupPolicy.IsUpdateNotifyPolicyValid; } }

        public string UpdateInstallId { get; set; }

        public string UpdateLastCheckedAtUtc { get; set; }

        public string UpdateLatestVersion { get; set; }

        public string UpdateReleaseUrl { get; set; }

        public string UpdateDownloadUrl { get; set; }

        public string UpdatePublishedAt { get; set; }

        public string UpdateChangelogTitle { get; set; }

        public string UpdateChangelogText { get; set; }

        public string UpdateLastNotifiedVersion { get; set; }

        public string UpdateLastNotifiedDateUtc { get; set; }

        public string FileLinkBasePath
        {
            get { return GetLocalValue(DefaultFileLinkBasePath); }
            set { SetLocalValue(value); }
        }


        public string SharingDefaultShareName
        {
            get { return GetLocalValue(Strings.SharingDefaultShareNameLabel); }
            set { SetLocalValue(value); }
        }

        public bool SharingDefaultPermCreate
        {
            get { return GetLocalValue(false); }
            set { SetLocalValue(value); }
        }

        public bool SharingDefaultPermWrite
        {
            get { return GetLocalValue(false); }
            set { SetLocalValue(value); }
        }

        public bool SharingDefaultPermDelete
        {
            get { return GetLocalValue(false); }
            set { SetLocalValue(value); }
        }

        public bool SharingDefaultPasswordEnabled
        {
            get { return GetLocalValue(true); }
            set { SetLocalValue(value); }
        }

        public bool SharingDefaultPasswordSeparateEnabled
        {
            get { return GetLocalValue(false); }
            set { SetLocalValue(value); }
        }

        public SharePasswordDeliveryMode SharingDefaultPasswordDeliveryMode
        {
            get { return GetLocalValue(SharePasswordDeliveryMode.Plain); }
            set { SetLocalValue(value); }
        }

        public int SharingDefaultExpireDays
        {
            get { return GetLocalValue(7); }
            set { SetLocalValue(value); }
        }

        public bool SharingAttachmentsAlwaysConnector
        {
            get { return GetLocalValue(false); }
            set { SetLocalValue(value); }
        }

        public bool SharingAttachmentsOfferAboveEnabled
        {
            get { return GetLocalValue(true); }
            set { SetLocalValue(value); }
        }

        public int SharingAttachmentsOfferAboveMb
        {
            get { return GetLocalValue(DefaultSharingAttachmentsOfferAboveMb); }
            set { SetLocalValue(value); }
        }

        public AttachmentLinkTarget? SharingAttachmentLinkTarget { get; set; }

        public string ShareBlockLang
        {
            get { return GetLocalValue("default"); }
            set { SetLocalValue(value); }
        }

        public string EventDescriptionLang
        {
            get { return GetLocalValue("default"); }
            set { SetLocalValue(value); }
        }

        public bool TalkDefaultLobbyEnabled
        {
            get { return GetLocalValue(true); }
            set { SetLocalValue(value); }
        }

        public bool TalkDefaultSearchVisible
        {
            get { return GetLocalValue(true); }
            set { SetLocalValue(value); }
        }

        public TalkRoomType TalkDefaultRoomType
        {
            get { return GetLocalValue(TalkRoomType.EventConversation); }
            set { SetLocalValue(value); }
        }

        public bool TalkDefaultPasswordEnabled
        {
            get { return GetLocalValue(true); }
            set { SetLocalValue(value); }
        }

        public bool TalkDefaultAddUsers
        {
            get { return GetLocalValue(true); }
            set { SetLocalValue(value); }
        }

        public bool TalkDefaultAddGuests
        {
            get { return GetLocalValue(false); }
            set { SetLocalValue(value); }
        }

        public bool TalkDeleteRoomOnEventDelete
        {
            get { return GetLocalValue(false); }
            set { SetLocalValue(value); }
        }

        public bool? EmailSignatureOnCompose { get; set; }

        public bool? EmailSignatureOnReply { get; set; }

        public bool? EmailSignatureOnForward { get; set; }

        internal string ManagedNextcloudUrl { get; private set; }

        internal string ManagedNextcloudUrlSource { get; private set; }

        internal bool ManagedNextcloudUrlLocked { get; private set; }

        internal bool IsEnterpriseRollout { get; private set; }

        internal bool ShowMainRibbonTab { get; private set; }

        internal bool HasManagedNextcloudUrl
        {
            get { return !string.IsNullOrWhiteSpace(ManagedNextcloudUrl); }
        }

        public AddinSettings Clone()
        {
            var copy = (AddinSettings)MemberwiseClone();
            copy._localPolicyValues = new Dictionary<string, object>(_localPolicyValues, StringComparer.Ordinal);
            return copy;
        }

        internal bool HasLocalValue(string propertyName)
        {
            return _localPolicyValues.ContainsKey(propertyName);
        }

        private T GetLocalValue<T>(T defaultValue, [CallerMemberName] string propertyName = null)
        {
            object value;
            return _localPolicyValues.TryGetValue(propertyName, out value) ? (T)value : defaultValue;
        }

        private void SetLocalValue<T>(T value, [CallerMemberName] string propertyName = null)
        {
            _localPolicyValues[propertyName] = value;
        }

        // Resolve a runtime copy without persisting backend defaults or overwriting local choices.
        internal AddinSettings ResolvePolicyDefaults(BackendPolicyStatus status)
        {
            AddinSettings resolved = Clone();
            bool preferBackendDefaults = ResolveDefaultsSource(status) == "backend";
            ApplyStringPolicy(resolved, status, "share", "share_base_directory", "FileLinkBasePath");
            ApplyStringPolicy(resolved, status, "share", "share_name_template", "SharingDefaultShareName");
            ApplyBoolPolicy(resolved, status, "share", "share_permission_upload", "SharingDefaultPermCreate");
            ApplyBoolPolicy(resolved, status, "share", "share_permission_edit", "SharingDefaultPermWrite");
            ApplyBoolPolicy(resolved, status, "share", "share_permission_delete", "SharingDefaultPermDelete");
            ApplyBoolPolicy(resolved, status, "share", "share_set_password", "SharingDefaultPasswordEnabled");
            ApplyBoolPolicy(resolved, status, "share", "share_send_password_separately", "SharingDefaultPasswordSeparateEnabled");
            ApplyBoolPolicy(resolved, status, "share", "attachments_always_via_ncconnector", "SharingAttachmentsAlwaysConnector");
            ApplyStringPolicy(resolved, status, "share", "language_share_html_block", "ShareBlockLang");
            ApplyBoolPolicy(resolved, status, "talk", "talk_lobby_active", "TalkDefaultLobbyEnabled");
            ApplyBoolPolicy(resolved, status, "talk", "talk_show_in_search", "TalkDefaultSearchVisible");
            ApplyBoolPolicy(resolved, status, "talk", "talk_set_password", "TalkDefaultPasswordEnabled");
            ApplyBoolPolicy(resolved, status, "talk", "talk_add_users", "TalkDefaultAddUsers");
            ApplyBoolPolicy(resolved, status, "talk", "talk_add_guests", "TalkDefaultAddGuests");
            ApplyBoolPolicy(resolved, status, "talk", "talk_delete_room_on_event_delete", "TalkDeleteRoomOnEventDelete");
            ApplyStringPolicy(resolved, status, "talk", "language_talk_description", "EventDescriptionLang");

            int days;
            if (ShouldApplyPolicy(status, "share", "share_expire_days", "SharingDefaultExpireDays")
                && status.TryGetPolicyInt("share", "share_expire_days", out days)
                && days > 0)
            {
                resolved.SharingDefaultExpireDays = Math.Min(3650, days);
            }
            if (ShouldApplyPolicy(status, "share", "share_send_password_mode", "SharingDefaultPasswordDeliveryMode")
                && status.HasPolicyKey("share", "share_send_password_mode"))
            {
                string mode = status.GetPolicyString("share", "share_send_password_mode");
                if (status.IsLocked("share", "share_send_password_mode")
                    || !HasLocalValue("SharingDefaultPasswordDeliveryMode")
                    || string.Equals(mode, "plain", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(mode, "secrets", StringComparison.OrdinalIgnoreCase))
                {
                    resolved.SharingDefaultPasswordDeliveryMode = SharePasswordDeliveryPolicy.ParseMode(mode);
                }
            }
            if (ShouldApplyPolicy(status, "talk", "talk_room_type", "TalkDefaultRoomType"))
            {
                string roomType = status.GetPolicyString("talk", "talk_room_type");
                if (string.Equals(roomType, "event", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(roomType, "group", StringComparison.OrdinalIgnoreCase))
                {
                    resolved.TalkDefaultRoomType = string.Equals(roomType, "event", StringComparison.OrdinalIgnoreCase)
                        ? TalkRoomType.EventConversation : TalkRoomType.StandardRoom;
                }
            }

            bool hasLocalThreshold = HasLocalValue("SharingAttachmentsOfferAboveEnabled")
                                     || HasLocalValue("SharingAttachmentsOfferAboveMb");
            if (status != null && status.IsDomainActive("share")
                && (status.IsLocked("share", "attachments_min_size_mb") || preferBackendDefaults || !hasLocalThreshold)
                && status.HasPolicyKey("share", "attachments_min_size_mb"))
            {
                int threshold;
                if (status.TryGetPolicyInt("share", "attachments_min_size_mb", out threshold))
                {
                    // Older backends accepted zero; retain their established five-MB interpretation.
                    resolved.SharingAttachmentsOfferAboveMb = Math.Min(
                        10240, OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(threshold));
                    resolved.SharingAttachmentsOfferAboveEnabled = true;
                }
                else if (status.GetPolicyValue("share", "attachments_min_size_mb") == null)
                {
                    resolved.SharingAttachmentsOfferAboveEnabled = false;
                }
            }
            resolved.SharingAttachmentLinkTarget = AttachmentLinkTargetPolicy.Resolve(
                SharingAttachmentLinkTarget, status, preferBackendDefaults);
            resolved.EmailSignatureOnCompose = EmailSignaturePolicyService.ResolveFlag(
                status, "email_signature_on_compose", EmailSignatureOnCompose, preferBackendDefaults);
            resolved.EmailSignatureOnReply = EmailSignaturePolicyService.ResolveFlag(
                status, "email_signature_on_reply", EmailSignatureOnReply, preferBackendDefaults);
            resolved.EmailSignatureOnForward = EmailSignaturePolicyService.ResolveFlag(
                status, "email_signature_on_forward", EmailSignatureOnForward, preferBackendDefaults);
            return resolved;
        }

        private bool ShouldApplyPolicy(BackendPolicyStatus status, string domain, string key, string propertyName)
        {
            return status != null && status.IsDomainActive(domain)
                   && (status.IsLocked(domain, key) || ResolveDefaultsSource(status) == "backend" || !HasLocalValue(propertyName));
        }

        private void ApplyBoolPolicy(AddinSettings resolved, BackendPolicyStatus status, string domain, string key, string propertyName)
        {
            bool value;
            if (ShouldApplyPolicy(status, domain, key, propertyName)
                && status.TryGetPolicyBool(domain, key, out value))
            {
                resolved._localPolicyValues[propertyName] = value;
            }
        }

        private void ApplyStringPolicy(AddinSettings resolved, BackendPolicyStatus status, string domain, string key, string propertyName)
        {
            if (ShouldApplyPolicy(status, domain, key, propertyName))
            {
                string value = status.GetPolicyString(domain, key);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    resolved._localPolicyValues[propertyName] = value;
                }
            }
        }

        internal void ApplyManagedSetupPolicy(ManagedSetupPolicy policy)
        {
            _managedSetupPolicy = policy;
            ManagedNextcloudUrl = string.Empty;
            ManagedNextcloudUrlSource = string.Empty;
            ManagedNextcloudUrlLocked = false;
            IsEnterpriseRollout = policy != null && policy.IsEnterpriseRollout;
            ShowMainRibbonTab = policy == null || policy.ShowMainRibbonTab;

            if (policy == null || !policy.HasNextcloudUrl)
            {
                return;
            }

            ManagedNextcloudUrl = policy.NextcloudUrl;
            ManagedNextcloudUrlSource = policy.Source;
            ManagedNextcloudUrlLocked = policy.NextcloudUrlLocked;

            if (ManagedNextcloudUrlLocked || string.IsNullOrWhiteSpace(ServerUrl))
            {
                ServerUrl = ManagedNextcloudUrl;
            }
        }

        internal static int NormalizeIfbPort(int port)
        {
            if (port < MinIfbPort || port > MaxIfbPort)
            {
                return DefaultIfbPort;
            }
            return port;
        }
    }
}

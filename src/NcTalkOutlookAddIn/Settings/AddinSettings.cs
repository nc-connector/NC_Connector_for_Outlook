// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    // Persistent add-in settings (credentials, sharing/IFB options, etc.).
    internal class AddinSettings
    {
        internal const int DefaultIfbPort = 7777;
        internal const int MinIfbPort = 1024;
        internal const int MaxIfbPort = 49151;
        internal const string DefaultFileLinkBasePath = "NC Connector";

        public AddinSettings()
        {
            ServerUrl = string.Empty;
            Username = string.Empty;
            AppPassword = string.Empty;
            AuthMode = AuthenticationMode.LoginFlow;
            IfbEnabled = false;
            IfbDays = 30;
            IfbCacheHours = 24;
            IfbPort = DefaultIfbPort;
            IfbPreviousFreeBusyPath = string.Empty;
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
            FileLinkBasePath = DefaultFileLinkBasePath;
            SharingDefaultShareName = Strings.SharingDefaultShareNameLabel;
            SharingDefaultPermCreate = false;
            SharingDefaultPermWrite = false;
            SharingDefaultPermDelete = false;
            SharingDefaultPasswordEnabled = true;
            SharingDefaultPasswordSeparateEnabled = false;
            SharingDefaultPasswordDeliveryMode = SharePasswordDeliveryMode.Plain;
            SharingDefaultExpireDays = 7;
            SharingAttachmentsAlwaysConnector = false;
            SharingAttachmentsOfferAboveEnabled = true;
            SharingAttachmentsOfferAboveMb = 20;
            SharingAttachmentLinkTarget = null;
            ShareBlockLang = "default";
            EventDescriptionLang = "default";
            TalkDefaultLobbyEnabled = true;
            TalkDefaultSearchVisible = true;
            TalkDefaultRoomType = TalkRoomType.EventConversation;
            TalkDefaultPasswordEnabled = true;
            TalkDefaultAddUsers = true;
            TalkDefaultAddGuests = false;
            TalkDeleteRoomOnEventDelete = false;
            EmailSignatureOnCompose = null;
            EmailSignatureOnReply = null;
            EmailSignatureOnForward = null;
            ManagedNextcloudUrl = string.Empty;
            ManagedNextcloudUrlSource = string.Empty;
            ManagedNextcloudUrlLocked = false;
        }

        public string ServerUrl { get; set; }

        public string Username { get; set; }

        public string AppPassword { get; set; }

        public AuthenticationMode AuthMode { get; set; }

        public bool IfbEnabled { get; set; }

        public int IfbDays { get; set; }

        public int IfbCacheHours { get; set; }

        public int IfbPort { get; set; }

        public string IfbPreviousFreeBusyPath { get; set; }

        public bool DebugLoggingEnabled { get; set; }

        public bool LogAnonymizationEnabled { get; set; }

        public bool TransportTlsUseSystemDefault { get; set; }

        public bool TransportTlsEnable12 { get; set; }

        public bool TransportTlsEnable13 { get; set; }

        public bool UpdateNotifyEnabled { get; set; }

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

        public string FileLinkBasePath { get; set; }


        public string SharingDefaultShareName { get; set; }

        public bool SharingDefaultPermCreate { get; set; }

        public bool SharingDefaultPermWrite { get; set; }

        public bool SharingDefaultPermDelete { get; set; }

        public bool SharingDefaultPasswordEnabled { get; set; }

        public bool SharingDefaultPasswordSeparateEnabled { get; set; }

        public SharePasswordDeliveryMode SharingDefaultPasswordDeliveryMode { get; set; }

        public int SharingDefaultExpireDays { get; set; }

        public bool SharingAttachmentsAlwaysConnector { get; set; }

        public bool SharingAttachmentsOfferAboveEnabled { get; set; }

        public int SharingAttachmentsOfferAboveMb { get; set; }

        public AttachmentLinkTarget? SharingAttachmentLinkTarget { get; set; }

        public string ShareBlockLang { get; set; }

        public string EventDescriptionLang { get; set; }

        public bool TalkDefaultLobbyEnabled { get; set; }

        public bool TalkDefaultSearchVisible { get; set; }

        public TalkRoomType TalkDefaultRoomType { get; set; }

        public bool TalkDefaultPasswordEnabled { get; set; }

        public bool TalkDefaultAddUsers { get; set; }

        public bool TalkDefaultAddGuests { get; set; }

        public bool TalkDeleteRoomOnEventDelete { get; set; }

        public bool? EmailSignatureOnCompose { get; set; }

        public bool? EmailSignatureOnReply { get; set; }

        public bool? EmailSignatureOnForward { get; set; }

        internal string ManagedNextcloudUrl { get; private set; }

        internal string ManagedNextcloudUrlSource { get; private set; }

        internal bool ManagedNextcloudUrlLocked { get; private set; }

        internal bool HasManagedNextcloudUrl
        {
            get { return !string.IsNullOrWhiteSpace(ManagedNextcloudUrl); }
        }

        public AddinSettings Clone()
        {
            var copy = (AddinSettings)MemberwiseClone();
            return copy;
        }

        internal void ApplyManagedSetupPolicy(ManagedSetupPolicy policy)
        {
            ManagedNextcloudUrl = string.Empty;
            ManagedNextcloudUrlSource = string.Empty;
            ManagedNextcloudUrlLocked = false;

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

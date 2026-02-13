/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    /**
     * Persistent add-in settings (credentials, sharing/IFB options, etc.).
     */
    internal class AddinSettings
    {
        public AddinSettings()
        {
            ServerUrl = string.Empty;
            Username = string.Empty;
            AppPassword = string.Empty;
            AuthMode = AuthenticationMode.LoginFlow;
            IfbEnabled = false;
            IfbDays = 30;
            IfbCacheHours = 24;
            IfbPreviousFreeBusyPath = string.Empty;
            DebugLoggingEnabled = false;
            LastKnownServerVersion = string.Empty;
            FileLinkBasePath = "90 Freigaben - extern";
            SharingDefaultShareName = Strings.SharingDefaultShareNameLabel;
            SharingDefaultPermCreate = false;
            SharingDefaultPermWrite = false;
            SharingDefaultPermDelete = false;
            SharingDefaultPasswordEnabled = true;
            SharingDefaultExpireDays = 7;
            ShareBlockLang = "default";
            EventDescriptionLang = "default";
            TalkDefaultLobbyEnabled = true;
            TalkDefaultSearchVisible = true;
            TalkDefaultRoomType = TalkRoomType.EventConversation;
            TalkDefaultPasswordEnabled = true;
            TalkDefaultAddUsers = true;
            TalkDefaultAddGuests = false;
        }

        public string ServerUrl { get; set; }

        public string Username { get; set; }

        public string AppPassword { get; set; }

        public AuthenticationMode AuthMode { get; set; }

        public bool IfbEnabled { get; set; }

        public int IfbDays { get; set; }

        public int IfbCacheHours { get; set; }

        public string IfbPreviousFreeBusyPath { get; set; }

        public bool DebugLoggingEnabled { get; set; }

        public string LastKnownServerVersion { get; set; }

        public string FileLinkBasePath { get; set; }

        public string SharingDefaultShareName { get; set; }

        public bool SharingDefaultPermCreate { get; set; }

        public bool SharingDefaultPermWrite { get; set; }

        public bool SharingDefaultPermDelete { get; set; }

        public bool SharingDefaultPasswordEnabled { get; set; }

        public int SharingDefaultExpireDays { get; set; }

        public string ShareBlockLang { get; set; }

        public string EventDescriptionLang { get; set; }

        public bool TalkDefaultLobbyEnabled { get; set; }

        public bool TalkDefaultSearchVisible { get; set; }

        public TalkRoomType TalkDefaultRoomType { get; set; }

        public bool TalkDefaultPasswordEnabled { get; set; }

        public bool TalkDefaultAddUsers { get; set; }

        public bool TalkDefaultAddGuests { get; set; }

        public AddinSettings Clone()
        {
            var copy = (AddinSettings)MemberwiseClone();
            return copy;
        }
    }
}

// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    // Coordinates the local IFB server and profile-owned Outlook configuration.
    internal sealed class FreeBusyManager : IDisposable
    {
        private readonly FreeBusyServer _server;
        private readonly IfbRegistryOwnershipManager _registryOwnership;
        private readonly string _requestSecret;

        private Outlook.Application _application;

        internal FreeBusyManager(string dataDirectory)
            : this(dataDirectory, "default")
        {
        }

        internal FreeBusyManager(
            string dataDirectory,
            string profileScope,
            Func<TalkServiceConfiguration, BackendPolicyStatus> fetchRolloutPolicy = null)
        {
            var addressBookCache = new IfbAddressBookCache(
                dataDirectory,
                profileScope);
            _server = new FreeBusyServer(addressBookCache, fetchRolloutPolicy);
            _registryOwnership =
                new IfbRegistryOwnershipManager(
                    dataDirectory,
                    profileScope);
            _requestSecret = CreateRequestSecret();
        }

        internal void Initialize(Outlook.Application application)
        {
            _application = application;
        }

        internal void ApplySettings(AddinSettings settings)
        {
            if (_application == null || settings == null)
            {
                return;
            }

            bool credentialsComplete =
                !string.IsNullOrWhiteSpace(settings.ServerUrl)
                && !string.IsNullOrWhiteSpace(settings.Username)
                && !string.IsNullOrEmpty(settings.AppPassword);
            if (!settings.IfbEnabled || !credentialsComplete)
            {
                StopServer();
                _registryOwnership.Restore();
                return;
            }

            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl,
                settings.Username,
                settings.AppPassword);
            if (!configuration.IsComplete())
            {
                StopServer();
                _registryOwnership.Restore();
                return;
            }

            _server.UpdateSettings(
                configuration,
                settings.IfbDays,
                settings.IfbCacheHours,
                _requestSecret,
                settings.IsEnterpriseRollout);
            int ifbPort =
                AddinSettings.NormalizeIfbPort(settings.IfbPort);
            try
            {
                _server.Start(ifbPort);
                _registryOwnership.Apply(
                    GetOutlookVersionSegment(),
                    BuildIfbUrl(settings, _requestSecret),
                    settings);
            }
            catch (HttpListenerException ex)
            {
                StopAndRestoreOwnedSettings();
                throw new InvalidOperationException(
                    BuildIfbStartFailureMessage(ifbPort, ex),
                    ex);
            }
            catch (Exception ex)
            {
                StopAndRestoreOwnedSettings();
                if (ex is InvalidOperationException)
                {
                    throw;
                }
                throw new InvalidOperationException(
                    "IFB server could not be started: "
                    + ex.Message,
                    ex);
            }
        }

        internal static string BuildIfbUrl(
            AddinSettings settings,
            string requestSecret)
        {
            int ifbPort =
                AddinSettings.NormalizeIfbPort(settings.IfbPort);
            return "http://127.0.0.1:"
                   + ifbPort.ToString(CultureInfo.InvariantCulture)
                   + "/nc-ifb/"
                   + (requestSecret ?? string.Empty).Trim()
                   + "/freebusy/%NAME%@%SERVER%.vfb";
        }

        internal static string CreateRequestSecret()
        {
            byte[] bytes = new byte[32];
            using (RandomNumberGenerator random =
                RandomNumberGenerator.Create())
            {
                random.GetBytes(bytes);
            }

            var builder = new StringBuilder(bytes.Length * 2);
            for (int i = 0; i < bytes.Length; i++)
            {
                builder.Append(
                    bytes[i].ToString(
                        "x2",
                        CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }

        private string GetOutlookVersionSegment()
        {
            try
            {
                string versionString =
                    _application != null
                        ? _application.Version
                        : null;
                if (!string.IsNullOrEmpty(versionString))
                {
                    string[] parts = versionString.Split('.');
                    if (parts.Length >= 2)
                    {
                        return parts[0] + "." + parts[1];
                    }
                    if (parts.Length == 1)
                    {
                        return parts[0] + ".0";
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Ifb,
                    "Failed to read Outlook version. Falling back to default version segment.",
                    ex);
            }
            return "16.0";
        }

        private static string BuildIfbStartFailureMessage(
            int ifbPort,
            HttpListenerException ex)
        {
            if (ex != null && ex.ErrorCode == 5)
            {
                return "IFB server could not be started on port "
                       + ifbPort.ToString(CultureInfo.InvariantCulture)
                       + ": access denied (URL reservation missing).";
            }
            if (ex != null && ex.ErrorCode == 32)
            {
                return "IFB server could not be started on port "
                       + ifbPort.ToString(CultureInfo.InvariantCulture)
                       + ": port/prefix is already in use.";
            }
            return "IFB server could not be started: "
                   + (ex != null
                       ? ex.Message
                       : "Unknown listener error.");
        }

        internal void StopServer()
        {
            _server.Stop();
        }

        private void StopAndRestoreOwnedSettings()
        {
            StopServer();
            try
            {
                _registryOwnership.Restore();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Ifb,
                    "Owned Outlook IFB settings could not be restored after server shutdown.",
                    ex);
            }
        }

        public void Dispose()
        {
            StopAndRestoreOwnedSettings();
        }
    }
}

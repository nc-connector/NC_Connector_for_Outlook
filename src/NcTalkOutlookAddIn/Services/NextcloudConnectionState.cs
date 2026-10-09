// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class ConnectionPauseStatus
    {
        internal ConnectionPauseStatus(string reason, DateTime retryAfterUtc)
        {
            Reason = reason ?? string.Empty;
            RetryAfterUtc = retryAfterUtc;
        }

        internal string Reason { get; private set; }
        internal DateTime RetryAfterUtc { get; private set; }
        internal bool IsPaused { get { return Reason.Length != 0; } }
    }

    internal sealed class VerifiedNextcloudIdentity
    {
        internal VerifiedNextcloudIdentity(TalkServiceConfiguration configuration, string userId)
        {
            Configuration = configuration;
            BaseUrl = configuration.GetNormalizedBaseUrl();
            UserId = userId;
        }

        internal string BaseUrl { get; private set; }
        internal string UserId { get; private set; }
        internal TalkServiceConfiguration Configuration { get; private set; }
    }

    internal sealed class PersistedNextcloudConnectionState
    {
        public PersistedNextcloudConnectionState()
        {
            Version = 1;
            RejectedBaseUrl = RateLimitOrigin = KnownBaseUrl = KnownUserId = KnownCredentialFingerprint = string.Empty;
        }

        public int Version { get; set; }
        public bool AuthenticationRejected { get; set; }
        public string RejectedBaseUrl { get; set; }
        public string RateLimitOrigin { get; set; }
        public DateTime RetryAfterUtc { get; set; }
        public string KnownBaseUrl { get; set; }
        public string KnownUserId { get; set; }
        public string KnownCredentialFingerprint { get; set; }
    }

    // Owns the saved connection's request pauses, verification proofs and recovery signals.
    internal static class NextcloudConnectionState
    {
        private sealed class VerificationProof
        {
            internal VerifiedNextcloudIdentity Identity;
            internal DateTime VerifiedAtUtc;
            internal long RequestSequence;
        }

        private static readonly object Sync = new object();
        private static readonly Dictionary<string, VerificationProof> Proofs = new Dictionary<string, VerificationProof>(StringComparer.Ordinal);
        private static readonly TimeSpan ProofLifetime = TimeSpan.FromMinutes(5);
        private static readonly ProtectedJsonStateStoreDefinition<PersistedNextcloudConnectionState> StoreDefinition =
            new ProtectedJsonStateStoreDefinition<PersistedNextcloudConnectionState>(
                "connection-state", "NC4OL::ConnectionState::v1",
                state => state != null && state.Version == 1
                    && state.RejectedBaseUrl != null && state.RateLimitOrigin != null
                    && state.KnownBaseUrl != null && state.KnownUserId != null && state.KnownCredentialFingerprint != null,
                () => new PersistedNextcloudConnectionState(), LogCategories.Api,
                new ProtectedJsonStateStoreMessages(
                    "Recovered Nextcloud connection state from its backup.",
                    "Failed to restore the Nextcloud connection state backup.",
                    "Nextcloud connection state is unreadable; automatic requests remain paused.",
                    "Nextcloud connection state is unreadable; existing files were preserved.",
                    "Nextcloud connection state structure is invalid.",
                    "Failed to load Nextcloud connection state file '"));

        private static ProtectedJsonStateStore<PersistedNextcloudConnectionState> _store;
        private static PersistedNextcloudConnectionState _state = new PersistedNextcloudConnectionState();
        private static TalkServiceConfiguration _savedConfiguration;
        private static VerifiedNextcloudIdentity _verifiedIdentity;
        private static Timer _rateLimitTimer;
        private static long _lastRejectedSequence;
        private static long _lastCommittedSequence;
        private static int _generation;

        internal static event Action<ConnectionPauseStatus> StateChanged;
        internal static event Action<VerifiedNextcloudIdentity> ConnectionRestored;

        internal static void Initialize(string dataDirectory, string profileScope, TalkServiceConfiguration savedConfiguration)
        {
            lock (Sync)
            {
                ShutdownLocked();
                _savedConfiguration = savedConfiguration;
                _store = new ProtectedJsonStateStore<PersistedNextcloudConnectionState>(dataDirectory, profileScope, StoreDefinition);
                _state = _store.Load();
                if (!_store.IsWriteAllowed && savedConfiguration != null)
                {
                    _state.AuthenticationRejected = true;
                    _state.RejectedBaseUrl = savedConfiguration.GetNormalizedBaseUrl();
                }
                ScheduleRateLimitExpiryLocked();
            }
        }

        internal static void Shutdown()
        {
            lock (Sync) { ShutdownLocked(); }
        }

        private static void ShutdownLocked()
        {
            _generation++;
            if (_rateLimitTimer != null) { _rateLimitTimer.Dispose(); _rateLimitTimer = null; }
            _savedConfiguration = null;
            _verifiedIdentity = null;
            _store = null;
            _state = new PersistedNextcloudConnectionState();
            _lastRejectedSequence = _lastCommittedSequence = 0;
            Proofs.Clear();
            ClearIdentityCaches();
        }

        internal static ConnectionPauseStatus GetStatus(TalkServiceConfiguration configuration = null, string requestUrl = null)
        {
            lock (Sync)
            {
                return GetStatusLocked(configuration ?? _savedConfiguration, requestUrl, false, true);
            }
        }

        internal static ConnectionPauseStatus GetRequestStatus(TalkServiceConfiguration configuration, string requestUrl,
            bool verification, bool authenticated)
        {
            lock (Sync) { return GetStatusLocked(configuration, requestUrl, verification, authenticated); }
        }

        internal static void AssertRequestAllowed(TalkServiceConfiguration configuration, bool verification = false)
        {
            ConnectionPauseStatus status = GetRequestStatus(configuration, null, verification, true);
            if (status.IsPaused)
            {
                bool rateLimited = status.Reason == "rate_limited";
                throw new TalkServiceException(rateLimited ? Strings.ConnectionRateLimited : Strings.ConnectionAuthRequired,
                    !rateLimited, rateLimited ? (System.Net.HttpStatusCode)429 : System.Net.HttpStatusCode.Unauthorized, null);
            }
        }

        private static ConnectionPauseStatus GetStatusLocked(TalkServiceConfiguration configuration, string requestUrl,
            bool verification, bool authenticated)
        {
            string origin = GetOrigin(requestUrl ?? (configuration != null ? configuration.GetNormalizedBaseUrl() : string.Empty));
            if (_state.RetryAfterUtc > DateTime.UtcNow && string.Equals(origin, _state.RateLimitOrigin, StringComparison.Ordinal))
            {
                return new ConnectionPauseStatus("rate_limited", _state.RetryAfterUtc);
            }
            if (authenticated && !verification && configuration != null && configuration.IsComplete())
            {
                if (!MatchesSavedConfigurationLocked(configuration)
                    || (_state.AuthenticationRejected && string.Equals(_state.RejectedBaseUrl, configuration.GetNormalizedBaseUrl(), StringComparison.Ordinal)))
                {
                    return new ConnectionPauseStatus("auth_required", DateTime.MinValue);
                }
            }
            return new ConnectionPauseStatus(string.Empty, DateTime.MinValue);
        }

        internal static void BeginVerification(TalkServiceConfiguration configuration)
        {
            lock (Sync) { Proofs.Remove(BuildCredentialFingerprint(configuration)); }
        }

        internal static bool HasFreshVerification(TalkServiceConfiguration configuration)
        {
            lock (Sync)
            {
                VerificationProof proof;
                return Proofs.TryGetValue(BuildCredentialFingerprint(configuration), out proof)
                    && DateTime.UtcNow - proof.VerifiedAtUtc <= ProofLifetime
                    && proof.RequestSequence >= _lastRejectedSequence;
            }
        }

        internal static void RecordVerifiedIdentity(TalkServiceConfiguration configuration, string userId,
            long requestSequence, bool verification)
        {
            if (configuration == null || string.IsNullOrWhiteSpace(userId)) { return; }
            lock (Sync)
            {
                var identity = new VerifiedNextcloudIdentity(configuration, userId);
                if (verification)
                {
                    if (Proofs.Count >= 10) { Proofs.Clear(); }
                    Proofs[BuildCredentialFingerprint(configuration)] = new VerificationProof
                    {
                        Identity = identity, VerifiedAtUtc = DateTime.UtcNow, RequestSequence = requestSequence
                    };
                }
                else if (MatchesSavedConfigurationLocked(configuration)
                    && !GetStatusLocked(configuration, null, false, true).IsPaused)
                {
                    _verifiedIdentity = identity;
                    if (SetKnownIdentityLocked(identity)) { SavePauseLocked(); }
                }
            }
        }

        internal static VerifiedNextcloudIdentity CommitVerifiedSavedConnection(TalkServiceConfiguration configuration)
        {
            VerifiedNextcloudIdentity identity;
            ConnectionPauseStatus status;
            lock (Sync)
            {
                if (!HasFreshVerification(configuration)) { throw new InvalidOperationException(Strings.ErrorCredentialsNotVerified); }
                VerificationProof proof = Proofs[BuildCredentialFingerprint(configuration)];
                identity = proof.Identity;
                PersistedNextcloudConnectionState previous = _state;
                _state = new PersistedNextcloudConnectionState
                {
                    RateLimitOrigin = previous.RateLimitOrigin,
                    RetryAfterUtc = previous.RetryAfterUtc
                };
                SetKnownIdentityLocked(identity);
                try { SaveLocked(); }
                catch { _state = previous; throw; }
                _savedConfiguration = configuration;
                _verifiedIdentity = identity;
                _lastCommittedSequence = proof.RequestSequence;
                _lastRejectedSequence = 0;
                Proofs.Clear();
                ClearIdentityCaches();
                ScheduleRateLimitExpiryLocked();
                status = GetStatusLocked(configuration, null, false, true);
            }
            DiagnosticsLogger.Log(LogCategories.Api, "Verified Nextcloud credentials were saved; matching requests may resume.");
            NotifyStateChanged(status);
            if (!status.IsPaused) { NotifyConnectionRestored(identity); }
            return identity;
        }

        internal static void RemoveSavedCredentials(TalkServiceConfiguration configurationWithEmptyCredentials)
        {
            if (configurationWithEmptyCredentials != null && configurationWithEmptyCredentials.IsComplete())
            {
                throw new ArgumentException("Credential removal requires an empty authentication configuration.");
            }
            ConnectionPauseStatus status;
            lock (Sync)
            {
                _savedConfiguration = configurationWithEmptyCredentials;
                _verifiedIdentity = null;
                Proofs.Clear();
                ClearIdentityCaches();
                status = GetStatusLocked(_savedConfiguration, null, false, true);
            }
            DiagnosticsLogger.Log(LogCategories.Api, "Saved Nextcloud credentials were removed; connection pauses and cleanup ownership are retained.");
            NotifyStateChanged(status);
        }

        internal static VerifiedNextcloudIdentity GetKnownIdentity(TalkServiceConfiguration configuration)
        {
            lock (Sync)
            {
                return configuration != null && !string.IsNullOrWhiteSpace(_state.KnownUserId)
                    && string.Equals(_state.KnownCredentialFingerprint, BuildCredentialFingerprint(configuration), StringComparison.Ordinal)
                    ? new VerifiedNextcloudIdentity(configuration, _state.KnownUserId) : null;
            }
        }

        internal static VerifiedNextcloudIdentity GetVerifiedIdentity()
        {
            lock (Sync)
            {
                return _verifiedIdentity != null && !GetStatusLocked(_savedConfiguration, null, false, true).IsPaused
                    ? _verifiedIdentity : null;
            }
        }

        internal static VerifiedNextcloudIdentity ResolveVerifiedIdentity()
        {
            TalkServiceConfiguration configuration;
            lock (Sync)
            {
                configuration = _savedConfiguration;
                if (configuration == null || !configuration.IsComplete() || GetStatusLocked(configuration, null, false, true).IsPaused) { return null; }
            }
            NextcloudUserIdentityService.ResolveCurrentUserId(configuration, true);
            return GetVerifiedIdentity();
        }

        internal static void RecordResponse(TalkServiceConfiguration configuration, string requestUrl, bool authenticated, NcHttpResponse response)
        {
            ConnectionPauseStatus status = null;
            lock (Sync)
            {
                if (_store == null) { return; }
                if (authenticated && response.StatusCode == System.Net.HttpStatusCode.Unauthorized
                    && MatchesSavedConfigurationLocked(configuration) && response.RequestSequence > _lastCommittedSequence)
                {
                    _state.AuthenticationRejected = true;
                    _state.RejectedBaseUrl = configuration.GetNormalizedBaseUrl();
                    _lastRejectedSequence = Math.Max(_lastRejectedSequence, response.RequestSequence);
                    _verifiedIdentity = null;
                    Proofs.Remove(BuildCredentialFingerprint(configuration));
                    NextcloudCapabilitiesService.ClearCache();
                    SavePauseLocked();
                    status = GetStatusLocked(_savedConfiguration, null, false, true);
                    DiagnosticsLogger.Log(LogCategories.Api, "HTTP 401: saved authentication was rejected; requests remain paused until verified credentials are saved.");
                }
                if ((int)response.StatusCode == 429)
                {
                    string origin = GetOrigin(requestUrl);
                    DateTime retryAt = HttpFailureDiagnostics.ReadRetryAfterUtc(response.Headers);
                    _state.RetryAfterUtc = string.Equals(_state.RateLimitOrigin, origin, StringComparison.Ordinal)
                        && _state.RetryAfterUtc > retryAt ? _state.RetryAfterUtc : retryAt;
                    _state.RateLimitOrigin = origin;
                    SavePauseLocked();
                    ScheduleRateLimitExpiryLocked();
                    status = GetStatusLocked(_savedConfiguration, requestUrl, false, true);
                    DiagnosticsLogger.Log(LogCategories.Api, "HTTP 429: Nextcloud requests are paused until the server retry deadline.");
                }
            }
            if (status != null) { NotifyStateChanged(status); }
        }

        private static bool SetKnownIdentityLocked(VerifiedNextcloudIdentity identity)
        {
            string fingerprint = BuildCredentialFingerprint(identity.Configuration);
            bool changed = _state.KnownCredentialFingerprint != fingerprint || _state.KnownUserId != identity.UserId;
            _state.KnownBaseUrl = identity.BaseUrl;
            _state.KnownUserId = identity.UserId;
            _state.KnownCredentialFingerprint = fingerprint;
            return changed;
        }

        private static bool MatchesSavedConfigurationLocked(TalkServiceConfiguration configuration)
        {
            return configuration != null && _savedConfiguration != null && _savedConfiguration.IsComplete()
                && string.Equals(BuildCredentialFingerprint(configuration), BuildCredentialFingerprint(_savedConfiguration), StringComparison.Ordinal);
        }

        private static string BuildCredentialFingerprint(TalkServiceConfiguration configuration)
        {
            if (configuration == null) { return string.Empty; }
            byte[] input = Encoding.UTF8.GetBytes(configuration.GetNormalizedBaseUrl() + "\n"
                + (configuration.Username ?? string.Empty).Trim() + "\n" + (configuration.AppPassword ?? string.Empty));
            using (SHA256 hash = SHA256.Create()) { return Convert.ToBase64String(hash.ComputeHash(input)); }
        }

        private static string GetOrigin(string url)
        {
            Uri uri;
            return Uri.TryCreate(url, UriKind.Absolute, out uri) ? uri.GetLeftPart(UriPartial.Authority) : string.Empty;
        }

        private static void SaveLocked()
        {
            if (_store == null) { throw new InvalidOperationException("Nextcloud connection state has not been initialized."); }
            _store.Save(_state);
        }

        private static void SavePauseLocked()
        {
            try { SaveLocked(); }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Api,
                    "Nextcloud connection state could not be saved; the current request pause remains active in memory.", ex);
            }
        }

        private static void ClearIdentityCaches()
        {
            NextcloudCapabilitiesService.ClearCache();
            NextcloudUserIdentityService.ClearCache();
        }

        private static void ScheduleRateLimitExpiryLocked()
        {
            if (_rateLimitTimer != null) { _rateLimitTimer.Dispose(); _rateLimitTimer = null; }
            if (_state.RetryAfterUtc <= DateTime.UtcNow) { return; }
            int generation = _generation;
            double delay = Math.Max(1, Math.Min(int.MaxValue, (_state.RetryAfterUtc - DateTime.UtcNow).TotalMilliseconds));
            _rateLimitTimer = new Timer(ignored => ExpireRateLimit(generation), null, (int)delay, Timeout.Infinite);
        }

        private static void ExpireRateLimit(int generation)
        {
            try
            {
                ConnectionPauseStatus status;
                bool restoreCurrentConnection;
                lock (Sync)
                {
                    if (generation != _generation) { return; }
                    if (_state.RetryAfterUtc > DateTime.UtcNow) { ScheduleRateLimitExpiryLocked(); return; }
                    _state.RetryAfterUtc = DateTime.MinValue;
                    SavePauseLocked();
                    status = GetStatusLocked(_savedConfiguration, null, false, true);
                    restoreCurrentConnection = _savedConfiguration != null
                        && string.Equals(GetOrigin(_savedConfiguration.GetNormalizedBaseUrl()), _state.RateLimitOrigin, StringComparison.Ordinal);
                }
                NotifyStateChanged(status);
                if (!status.IsPaused && restoreCurrentConnection)
                {
                    VerifiedNextcloudIdentity identity = ResolveVerifiedIdentity();
                    if (identity != null) { NotifyConnectionRestored(identity); }
                }
            }
            catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Api, "Nextcloud retry-delay recovery failed.", ex); }
        }

        private static void NotifyStateChanged(ConnectionPauseStatus status)
        {
            Action<ConnectionPauseStatus> listeners = StateChanged;
            if (listeners == null) { return; }
            foreach (Action<ConnectionPauseStatus> listener in listeners.GetInvocationList())
            {
                try { listener(status); }
                catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Api, "Nextcloud connection-state listener failed.", ex); }
            }
        }

        private static void NotifyConnectionRestored(VerifiedNextcloudIdentity identity)
        {
            Action<VerifiedNextcloudIdentity> listeners = ConnectionRestored;
            if (listeners == null) { return; }
            foreach (Action<VerifiedNextcloudIdentity> listener in listeners.GetInvocationList())
            {
                try { listener(identity); }
                catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Api, "Nextcloud connection-recovery listener failed.", ex); }
            }
        }
    }
}

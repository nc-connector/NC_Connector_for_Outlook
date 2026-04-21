/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Performs HTTP calls against the Nextcloud Talk REST API.
     */
    internal sealed class TalkService
    {
        private const int RoomTypePublic = 3;
        private const int ListableNone = 0;
        private const int ListableUsers = 1;

        private const string ActorTypeUsers = "users";
        private const string ActorTypeEmails = "emails";

        private readonly TalkServiceConfiguration _configuration;
        private readonly NcHttpClient _httpClient;

        private static void LogTalk(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Talk, message);
        }

        internal TalkService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }

            _configuration = configuration;
            _httpClient = new NcHttpClient(configuration);
        }

        /**
         * Creates a Talk room based on user input and returns token + URL.
         */
        internal TalkRoomCreationResult CreateRoom(TalkRoomRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }

            EnsureConfiguration();

            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string createUrl = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room";

            bool includeEvent = request.RoomType == TalkRoomType.EventConversation;
            if (includeEvent && (!request.AppointmentStart.HasValue || !request.AppointmentEnd.HasValue))
            {
                throw new TalkServiceException("Event conversation requires appointment start/end.", false, HttpStatusCode.BadRequest, null);
            }
            try
            {
                LogTalk("CreateRoom attempt includeEvent=" + includeEvent + " lobby=" + request.LobbyEnabled + " listable=" + request.SearchVisible + " addUsers=" + request.AddUsers + " addGuests=" + request.AddGuests);
                IDictionary<string, object> payload = BuildCreatePayload(request, includeEvent);

                IDictionary<string, object> responseData;
                HttpStatusCode statusCode;
                string responseText = ExecuteJsonRequest("POST", createUrl, payload, out statusCode, out responseData);

                if (!IsSuccessStatus(statusCode))
                {
                    ThrowServiceError(statusCode, responseText, responseData);
                }
                string token = ExtractRoomToken(responseData);
                if (string.IsNullOrEmpty(token))
                {
                    throw new TalkServiceException("Response did not contain a room token.", false, statusCode, responseText);
                }
                string roomUrl = baseUrl + "/call/" + token;

                if (request.LobbyEnabled)
                {
                    // Lobby updates must not abort room creation. If this fails, the room still exists and
                    // Outlook will retry on appointment save (Write event) when the tracking subscription runs.
                    TryUpdateLobbyInternal(token, request.AppointmentStart, baseUrl, true);
                }
                if (!includeEvent)
                {
                    TryUpdateListable(token, request.SearchVisible, baseUrl);
                }
                try
                {
                    TryUpdateDescription(token, request.Description, includeEvent, baseUrl);
                }
                catch (TalkServiceException ex)
                {
                    if (includeEvent && IsEventConversationDescriptionLockError(ex))
                    {
                        // Event conversations may reject direct room description updates with a generic
                        // "event" BadRequest. Keep the event conversation and continue without room-type fallback.
                        LogTalk("CreateRoom description update skipped for event conversation: " + ex.Message);
                    }
                    else
                    {
                        throw;
                    }
                }
                return new TalkRoomCreationResult(token, roomUrl, includeEvent, request.LobbyEnabled, request.SearchVisible);
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "CreateRoom failed (includeEvent=" + includeEvent + ", status=" + (int)ex.StatusCode + ").", ex);
                throw;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "CreateRoom failed unexpectedly (includeEvent=" + includeEvent + ").", ex);
                throw;
            }
        }

        /**
         * Updates the lobby time of an existing Talk room.
         */
        internal void UpdateLobby(string roomToken, DateTime start, DateTime end, bool isEventConversation)
        {
            EnsureConfiguration();

            if (string.IsNullOrWhiteSpace(roomToken))
            {
                return;
            }
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            if (isEventConversation)
            {
                try
                {
                    TryUpdateEventBinding(roomToken, start, end, baseUrl);
                }
                catch (TalkServiceException ex)
                {
                    if (!IsRecoverableEventBindingStatus(ex.StatusCode))
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "UpdateLobby event binding failed (status=" + (int)ex.StatusCode + ").", ex);
                        throw;
                    }

                    DiagnosticsLogger.LogException(LogCategories.Talk, "UpdateLobby event binding not supported/recoverable (status=" + (int)ex.StatusCode + ").", ex);
                }
            }

            TryUpdateLobbyInternal(roomToken, start, baseUrl, false);
        }

        /**
         * Deletes the Talk room (removes the current participant) and ignores 404.
         */
        internal void DeleteRoom(string roomToken, bool isEventConversation)
        {
            EnsureConfiguration();

            if (string.IsNullOrWhiteSpace(roomToken))
            {
                return;
            }
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            TryClearActiveParticipants(roomToken, baseUrl);
            if (isEventConversation)
            {
                TryDetachEventBinding(roomToken, baseUrl);
            }
            string url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(roomToken.Trim());

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            ExecuteJsonRequest("DELETE", url, (string)null, out statusCode, out parsed);

            if (statusCode != HttpStatusCode.OK &&
                statusCode != HttpStatusCode.NoContent &&
                statusCode != HttpStatusCode.NotFound)
            {
                throw new TalkServiceException("Talk room could not be deleted (status " + (int)statusCode + ").", false, statusCode, null);
            }
        }

        internal bool VerifyConnection(out string message)
        {
            EnsureConfiguration();

            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string url = baseUrl + "/ocs/v2.php/cloud/capabilities";

            HttpStatusCode statusCode;
            IDictionary<string, object> data;
            // Force a fresh connection for connectivity diagnostics so TLS mode changes
            // are validated against a new handshake and not masked by pooled keep-alive sockets.
            ExecuteJsonRequest("GET", url, (string)null, out statusCode, out data, true);

            if (!IsSuccessStatus(statusCode))
            {
                message = "HTTP " + (int)statusCode;
                return false;
            }

            IDictionary<string, object> ocs = NcJson.GetDictionary(data, "ocs");
            IDictionary<string, object> meta = NcJson.GetDictionary(ocs, "meta");
            IDictionary<string, object> payload = NcJson.GetDictionary(ocs, "data");

            string version = ExtractVersion(payload);
            string status = NcJson.GetString(meta, "status");
            if (!string.IsNullOrEmpty(version))
            {
                message = version;
            }
            else if (!string.IsNullOrEmpty(status))
            {
                message = status;
            }
            else
            {
                message = "OK";
            }
            return true;
        }

        private IDictionary<string, object> BuildCreatePayload(TalkRoomRequest request, bool includeEvent)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            // Keep both fields for backward compatibility across Talk API variants.
            payload["roomType"] = RoomTypePublic;
            payload["type"] = RoomTypePublic;
            payload["roomName"] = string.IsNullOrWhiteSpace(request.Title) ? "Meeting" : request.Title.Trim();
            payload["listable"] = request.SearchVisible ? ListableUsers : ListableNone;
            payload["participants"] = new Dictionary<string, object>();

            if (!string.IsNullOrWhiteSpace(request.Password))
            {
                payload["password"] = request.Password;
            }
            if (!string.IsNullOrWhiteSpace(request.Description))
            {
                payload["description"] = request.Description.Trim();
            }
            if (includeEvent)
            {
                string objectId = BuildEventObjectId(request);
                if (!string.IsNullOrEmpty(objectId))
                {
                    payload["objectType"] = "event";
                    payload["objectId"] = objectId;
                }
            }
            return payload;
        }

        private void TryUpdateLobbyInternal(string token, DateTime? start, string baseUrl, bool silent)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return;
            }

            // Send a single lobby update request with an optional timer.
            // Some servers react poorly to two-stage calls (first without timer, then with timer) and may keep a
            // wrong start time.
            TrySendLobbyRequest(token, baseUrl, 1, start, silent);
        }

        private bool TrySendLobbyRequest(string token, string baseUrl, int state, DateTime? start, bool silent)
        {
            long? unixStart = TimeUtilities.ToUnixTimeSeconds(start);

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["state"] = state;
            if (unixStart.HasValue)
            {
                payload["timer"] = unixStart.Value;
            }
            string lobbyUrl = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token) + "/webinar/lobby";

            try
            {
                HttpStatusCode statusCode;
                IDictionary<string, object> parsed;
                ExecuteJsonRequest("PUT", lobbyUrl, payload, out statusCode, out parsed);
                if (!IsSuccessStatus(statusCode) && !silent)
                {
                    throw new TalkServiceException("Lobby time could not be set (status " + (int)statusCode + ").", false, statusCode, null);
                }
                return IsSuccessStatus(statusCode);
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Lobby update failed (silent=" + silent + ").", ex);
                if (!silent)
                {
                    throw;
                }
            }
            return false;
        }

        private void TryUpdateEventBinding(string token, DateTime? start, DateTime? end, string baseUrl)
        {
            string objectId = BuildEventObjectId(start, end);
            if (string.IsNullOrEmpty(objectId))
            {
                return;
            }

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["objectType"] = "event";
            payload["objectId"] = objectId;

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            string responseText = ExecuteJsonRequest("PUT",
                                                     baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token) + "/object",
                                                     payload,
                                                     out statusCode,
                                                     out parsed);

            if (!IsSuccessStatus(statusCode))
            {
                ThrowServiceError(statusCode, responseText, parsed);
            }
        }

        private static bool IsRecoverableEventBindingStatus(HttpStatusCode statusCode)
        {
            return statusCode == HttpStatusCode.MethodNotAllowed ||
                   statusCode == HttpStatusCode.NotImplemented ||
                   statusCode == HttpStatusCode.BadRequest;
        }

        private void TryClearActiveParticipants(string token, string baseUrl)
        {
            try
            {
                HttpStatusCode statusCode;
                IDictionary<string, object> parsed;
                ExecuteJsonRequest("DELETE",
                                   baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token) + "/participants/active",
                                   (string)null,
                                   out statusCode,
                                   out parsed);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to clear active participants (best-effort cleanup).", ex);
            }
        }

        private void TryDetachEventBinding(string token, string baseUrl)
        {
            try
            {
                HttpStatusCode statusCode;
                IDictionary<string, object> parsed;
                ExecuteJsonRequest("DELETE",
                                   baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token) + "/object/event",
                                   (string)null,
                                   out statusCode,
                                   out parsed);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to detach event binding (best-effort cleanup).", ex);
            }
        }

        private void TryUpdateListable(string token, bool searchVisible, string baseUrl)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["scope"] = searchVisible ? ListableUsers : ListableNone;

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            ExecuteJsonRequest("PUT",
                               baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token) + "/listable",
                               payload,
                               out statusCode,
                               out parsed);
        }

        internal void UpdateDescription(string token, string description, bool isEventConversation)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return;
            }

            EnsureConfiguration();
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            TryUpdateDescription(token, description, isEventConversation, baseUrl);
        }

        internal void UpdateRoomName(string token, string roomName)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return;
            }

            EnsureConfiguration();
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            TryUpdateRoomName(token, roomName, baseUrl);
        }

        internal bool AddUserParticipant(string token, string userId)
        {
            return AddParticipant(token, ActorTypeUsers, userId);
        }

        internal bool AddGuestParticipant(string token, string email)
        {
            return AddParticipant(token, ActorTypeEmails, email);
        }

        internal bool AddParticipant(string token, string actorType, string actorId)
        {
            if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(actorType) || string.IsNullOrWhiteSpace(actorId))
            {
                return false;
            }

            EnsureConfiguration();
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token.Trim()) + "/participants";

            var payload = new Dictionary<string, object>();
            // The OCS API expects "newParticipant" + "source".
            // "source" matches Talk actor types such as "users" and "emails".
            payload["newParticipant"] = actorId.Trim();
            payload["source"] = actorType.Trim();

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            string responseText = ExecuteJsonRequest("POST", url, payload, out statusCode, out parsed);

            if (IsSuccessStatus(statusCode) || statusCode == HttpStatusCode.Conflict)
            {
                return true;
            }

            ThrowServiceError(statusCode, responseText, parsed);
            return false;
        }

        internal List<TalkParticipant> GetParticipants(string token)
        {
            var participants = new List<TalkParticipant>();
            if (string.IsNullOrWhiteSpace(token))
            {
                return participants;
            }

            EnsureConfiguration();
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token.Trim()) + "/participants?includeStatus=true";

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            string responseText = ExecuteJsonRequest("GET", url, (string)null, out statusCode, out parsed);

            if (!IsSuccessStatus(statusCode))
            {
                ThrowServiceError(statusCode, responseText, parsed);
                return participants;
            }

            object listObj = null;
            IDictionary<string, object> ocs = NcJson.GetDictionary(parsed, "ocs");            if (ocs != null)
            {
                ocs.TryGetValue("data", out listObj);
            }

            object[] list = listObj as object[];            if (list == null)
            {
                // Some Talk versions wrap the list under { data: { participants: [...] } }.
                var dataDict = listObj as IDictionary<string, object>;                if (dataDict != null)
                {
                    object participantsObj;
                    if (dataDict.TryGetValue("participants", out participantsObj))
                    {
                        list = participantsObj as object[];
                    }
                }
            }
            if (list == null)
            {
                return participants;
            }
            foreach (object entry in list)
            {
                var dict = entry as IDictionary<string, object>;                if (dict == null)
                {
                    continue;
                }
                string actorType = NcJson.GetString(dict, "actorType") ?? string.Empty;
                string actorId = NcJson.GetString(dict, "actorId") ?? string.Empty;
                int attendeeId = NcJson.GetInt(dict, "attendeeId");
                participants.Add(new TalkParticipant(actorType, actorId, attendeeId));
            }
            return participants;
        }

        internal bool PromoteModerator(string token, int attendeeId, out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrWhiteSpace(token) || attendeeId <= 0)
            {
                return false;
            }

            EnsureConfiguration();
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token.Trim()) + "/moderators";

            var payload = new Dictionary<string, object>();
            payload["attendeeId"] = attendeeId;

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            string responseText = ExecuteJsonRequest("POST", url, payload, out statusCode, out parsed);

            if (IsSuccessStatus(statusCode) || statusCode == HttpStatusCode.Conflict)
            {
                return true;
            }
            string message = ExtractOcsMessage(parsed);
            if (string.IsNullOrWhiteSpace(message))
            {
                message = "HTTP " + (int)statusCode;
            }

            errorMessage = message;
            return false;
        }

        internal bool LeaveRoom(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            EnsureConfiguration();
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string url = baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token.Trim()) + "/participants/self";

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            ExecuteJsonRequest("DELETE", url, (string)null, out statusCode, out parsed);

            return IsSuccessStatus(statusCode) || statusCode == HttpStatusCode.NotFound;
        }

        /**
         * Reads server-side room traits for cross-client interoperability when local
         * appointment flags are not available yet (for example TB-created events in Outlook).
         */
        internal bool TryReadRoomTraits(string token, out bool? lobbyEnabled, out bool? isEventConversation)
        {
            lobbyEnabled = null;
            isEventConversation = null;

            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            EnsureConfiguration();
            string normalizedToken = token.Trim();
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            bool resolvedAny = false;

            try
            {
                HttpStatusCode statusCode;
                IDictionary<string, object> parsed;
                ExecuteJsonRequest("GET",
                                   baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(normalizedToken) + "/object",
                                   (string)null,
                                   out statusCode,
                                   out parsed);

                if (IsSuccessStatus(statusCode))
                {
                    isEventConversation = true;
                    resolvedAny = true;
                }
                else if (statusCode == HttpStatusCode.NotFound ||
                         statusCode == HttpStatusCode.Conflict ||
                         IsRecoverableEventBindingStatus(statusCode))
                {
                    isEventConversation = false;
                    resolvedAny = true;
                }
            }
            catch (TalkServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.NotFound ||
                    ex.StatusCode == HttpStatusCode.Conflict ||
                    IsRecoverableEventBindingStatus(ex.StatusCode))
                {
                    isEventConversation = false;
                    resolvedAny = true;
                }
                else
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to probe event conversation flag.", ex);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to probe event conversation flag.", ex);
            }
            try
            {
                HttpStatusCode statusCode;
                IDictionary<string, object> parsed;
                ExecuteJsonRequest("GET",
                                   baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(normalizedToken) + "/webinar/lobby",
                                   (string)null,
                                   out statusCode,
                                   out parsed);

                if (IsSuccessStatus(statusCode))
                {
                    IDictionary<string, object> data = NcJson.GetDictionary(NcJson.GetDictionary(parsed, "ocs"), "data");
                    int lobbyState;
                    if (NcJson.TryGetInt(data, "state", out lobbyState))
                    {
                        lobbyEnabled = lobbyState != 0;
                        resolvedAny = true;
                    }
                }
                else if (statusCode == HttpStatusCode.NotFound ||
                         statusCode == HttpStatusCode.MethodNotAllowed ||
                         statusCode == HttpStatusCode.NotImplemented)
                {
                    lobbyEnabled = false;
                    resolvedAny = true;
                }
            }
            catch (TalkServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.NotFound ||
                    ex.StatusCode == HttpStatusCode.MethodNotAllowed ||
                    ex.StatusCode == HttpStatusCode.NotImplemented)
                {
                    lobbyEnabled = false;
                    resolvedAny = true;
                }
                else
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to probe lobby flag.", ex);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to probe lobby flag.", ex);
            }
            return resolvedAny;
        }

        private void TryUpdateDescription(string token, string description, bool isEventConversation, string baseUrl)
        {            if (description == null)
            {
                return;
            }

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["description"] = description.Trim();

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            string responseText = ExecuteJsonRequest("PUT",
                                                     baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token) + "/description",
                                                     payload,
                                                     out statusCode,
                                                     out parsed);

            if (!IsSuccessStatus(statusCode))
            {
                ThrowServiceError(statusCode, responseText, parsed);
            }
        }

        private void TryUpdateRoomName(string token, string roomName, string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(roomName))
            {
                return;
            }

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["roomName"] = roomName.Trim();

            HttpStatusCode statusCode;
            IDictionary<string, object> parsed;
            string responseText = ExecuteJsonRequest("PUT",
                                                     baseUrl + "/ocs/v2.php/apps/spreed/api/v4/room/" + Uri.EscapeDataString(token),
                                                     payload,
                                                     out statusCode,
                                                     out parsed);

            if (!IsSuccessStatus(statusCode))
            {
                ThrowServiceError(statusCode, responseText, parsed);
            }
        }

        private static string ExtractOcsMessage(IDictionary<string, object> responseData)
        {
            return NcJson.ExtractOcsErrorMessage(responseData);
        }

        private void EnsureConfiguration()
        {
            if (!_configuration.IsComplete())
            {
                throw new TalkServiceException("Credentials are incomplete.", true, HttpStatusCode.BadRequest, null);
            }
        }

        private string ExecuteJsonRequest(string method, string url, object payload, out HttpStatusCode statusCode, out IDictionary<string, object> parsedData)
        {
            return ExecuteJsonRequest(method, url, payload, out statusCode, out parsedData, false);
        }

        private string ExecuteJsonRequest(string method, string url, object payload, out HttpStatusCode statusCode, out IDictionary<string, object> parsedData, bool forceFreshConnection)
        {
            string payloadText = null;
            if (payload is string)
            {
                payloadText = (string)payload;
            }
            else if (payload != null)
            {
                payloadText = NcJson.Serialize(payload);
            }

            DiagnosticsLogger.LogApi(method + " " + url);
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = method,
                Url = url,
                Payload = payloadText,
                Accept = "application/json",
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true,
                ForceFreshConnection = forceFreshConnection
            });

            if (!response.HasHttpResponse)
            {                if (response.TransportException != null)
                {
                    HttpFailureInfo failure = response.FailureInfo ?? HttpFailureDiagnostics.Analyze(response.TransportException);
                    DiagnosticsLogger.LogException(LogCategories.Api, "HTTP connection error without HTTP response (" + HttpFailureDiagnostics.BuildLogSummary(response.TransportException, failure) + ").", response.TransportException);
                    throw new TalkServiceException(failure.BuildUserMessage(), false, 0, null, true);
                }

                throw new TalkServiceException("HTTP connection error without HTTP response.", false, 0, null, true);
            }

            statusCode = response.StatusCode;
            DiagnosticsLogger.LogApi(method + " " + url + " -> " + statusCode);            if (response.JsonParseException != null)
            {
                DiagnosticsLogger.LogException(LogCategories.Api, "Failed to parse JSON response.", response.JsonParseException);
            }

            parsedData = response.ParsedJson;
            return response.ResponseText;
        }

        private static bool IsSuccessStatus(HttpStatusCode statusCode)
        {
            return statusCode >= HttpStatusCode.OK && statusCode < HttpStatusCode.MultipleChoices;
        }

        private static string ExtractRoomToken(IDictionary<string, object> responseData)
        {            if (responseData == null)
            {
                return null;
            }

            IDictionary<string, object> ocs = NcJson.GetDictionary(responseData, "ocs");
            IDictionary<string, object> data = NcJson.GetDictionary(ocs, "data");            if (data != null)
            {
                string token = NcJson.GetString(data, "token");
                if (string.IsNullOrEmpty(token))
                {
                    token = NcJson.GetString(data, "roomToken");
                }
                if (!string.IsNullOrEmpty(token))
                {
                    return token;
                }
            }
            return NcJson.GetString(responseData, "token");
        }
        private static string ExtractVersion(IDictionary<string, object> payload)
        {            if (payload == null)
            {
                return null;
            }
            string version = NcJson.GetString(payload, "versionstring") ?? NcJson.GetString(payload, "version");
            string edition = NcJson.GetString(payload, "edition");
            string result = ComposeVersion(version, edition);
            if (!string.IsNullOrEmpty(result))
            {
                return EnsureProductPrefix(result, payload);
            }

            IDictionary<string, object> versionDict = NcJson.GetDictionary(payload, "version");
            result = ComposeVersion(
                NcJson.GetString(versionDict, "string") ?? BuildVersionFromParts(versionDict),
                NcJson.GetString(versionDict, "edition"));
            if (!string.IsNullOrEmpty(result))
            {
                return EnsureProductPrefix(result, payload);
            }

            IDictionary<string, object> nextcloud = NcJson.GetDictionary(payload, "nextcloud");
            IDictionary<string, object> system = NcJson.GetDictionary(nextcloud, "system");            if (system != null)
            {
                result = ComposeVersion(
                    NcJson.GetString(system, "versionstring") ?? NcJson.GetString(system, "version") ?? BuildVersionFromParts(system),
                    NcJson.GetString(system, "edition"));
                if (!string.IsNullOrEmpty(result))
                {
                    string product = NcJson.GetString(system, "productname") ?? NcJson.GetString(nextcloud, "productname");
                    if (!string.IsNullOrEmpty(product))
                    {
                        return product + " " + result;
                    }
                    return EnsureProductPrefix(result, payload);
                }
            }
            return null;
        }

        private static string BuildVersionFromParts(IDictionary<string, object> dictionary)
        {            if (dictionary == null)
            {
                return null;
            }
            string major = NcJson.GetString(dictionary, "major");
            string minor = NcJson.GetString(dictionary, "minor");
            string micro = NcJson.GetString(dictionary, "micro");

            if (string.IsNullOrEmpty(major))
            {
                return null;
            }
            var builder = new StringBuilder();
            builder.Append(major);
            if (!string.IsNullOrEmpty(minor))
            {
                builder.Append(".").Append(minor);
                if (!string.IsNullOrEmpty(micro))
                {
                    builder.Append(".").Append(micro);
                }
            }
            return builder.ToString();
        }

        private static string ComposeVersion(string version, string edition)
        {
            if (string.IsNullOrEmpty(version))
            {
                return null;
            }
            if (!string.IsNullOrEmpty(edition))
            {
                return edition.Equals(version, StringComparison.OrdinalIgnoreCase)
                    ? version
                    : version + " (" + edition + ")";
            }
            return version;
        }

        private static string EnsureProductPrefix(string version, IDictionary<string, object> payload)
        {
            if (string.IsNullOrEmpty(version))
            {
                return null;
            }
            string prefix = "Nextcloud";
            string product = NcJson.GetString(payload, "productname");
            if (!string.IsNullOrEmpty(product))
            {
                prefix = product;
            }
            if (version.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return version;
            }
            return prefix + " " + version;
        }

        private static void ThrowServiceError(HttpStatusCode statusCode, string responseText, IDictionary<string, object> responseData)
        {
            StringBuilder builder = new StringBuilder();
            string error = NcJson.ExtractOcsErrorMessage(responseData);
            if (!string.IsNullOrWhiteSpace(error))
            {
                builder.Append(error);
            }
            if (builder.Length == 0)
            {
                builder.Append("HTTP ").Append((int)statusCode);
            }
            bool authError = statusCode == HttpStatusCode.Unauthorized || statusCode == HttpStatusCode.Forbidden;
            throw new TalkServiceException(builder.ToString(), authError, statusCode, responseText);
        }

        private static bool IsEventConversationDescriptionLockError(TalkServiceException ex)
        {            if (ex == null || ex.StatusCode != HttpStatusCode.BadRequest)
            {
                return false;
            }
            if (string.IsNullOrWhiteSpace(ex.Message))
            {
                return false;
            }
            string normalized = ex.Message.ToLowerInvariant();
            return normalized.IndexOf("event", StringComparison.Ordinal) >= 0;
        }

        private static string BuildEventObjectId(TalkRoomRequest request)
        {
            // Keep null-safe to preserve fail-soft behavior on partial room request objects.
            if (request == null)
            {
                return null;
            }
            return BuildEventObjectId(request.AppointmentStart, request.AppointmentEnd);
        }

        private static string BuildEventObjectId(DateTime? start, DateTime? end)
        {
            long? startEpoch = TimeUtilities.ToUnixTimeSeconds(start);
            long? endEpoch = TimeUtilities.ToUnixTimeSeconds(end);

            if (!startEpoch.HasValue || !endEpoch.HasValue)
            {
                return null;
            }
            return startEpoch.Value + "#" + endEpoch.Value;
        }

        
    }
}





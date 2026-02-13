/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
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
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private static void LogApi(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Api, message);
        }

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

            bool attemptEventConversation = request.RoomType == TalkRoomType.EventConversation
                                            && request.AppointmentStart.HasValue
                                            && request.AppointmentEnd.HasValue;

            List<bool> attempts = new List<bool>();
            if (attemptEventConversation)
            {
                attempts.Add(true);
            }
            attempts.Add(false);

            Exception lastError = null;

            foreach (bool includeEvent in attempts)
            {
                try
                {
                    LogTalk("CreateRoom attempt includeEvent=" + includeEvent + " lobby=" + request.LobbyEnabled + " listable=" + request.SearchVisible + " addUsers=" + request.AddUsers + " addGuests=" + request.AddGuests);
                    IDictionary<string, object> payload = BuildCreatePayload(request, includeEvent);
                    string payloadJson = _serializer.Serialize(payload);

                    IDictionary<string, object> responseData;
                    HttpStatusCode statusCode;
                    string responseText = ExecuteJsonRequest("POST", createUrl, payload, out statusCode, out responseData);

                    if (!IsSuccessStatus(statusCode))
                    {
                        if (includeEvent && ShouldFallbackToStandard(statusCode, responseData))
                        {
                            lastError = null;
                            continue;
                        }

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

                    TryUpdateDescription(token, request.Description, includeEvent, baseUrl);

                    return new TalkRoomCreationResult(token, roomUrl, includeEvent, request.LobbyEnabled, request.SearchVisible);
                }
                catch (TalkServiceException ex)
                {
                    lastError = ex;
                    DiagnosticsLogger.LogException(LogCategories.Talk, "CreateRoom failed (includeEvent=" + includeEvent + ", status=" + (int)ex.StatusCode + ").", ex);
                    if (includeEvent && !ex.IsAuthenticationError && ShouldFallbackToStandard(ex.StatusCode, null))
                    {
                        LogTalk("CreateRoom falling back to standard room due to status=" + (int)ex.StatusCode + ".");
                        continue;
                    }

                    throw;
                }
                catch (Exception ex)
                {
                    lastError = ex;
                    DiagnosticsLogger.LogException(LogCategories.Talk, "CreateRoom failed unexpectedly (includeEvent=" + includeEvent + ").", ex);
                    throw;
                }
            }

            if (lastError != null)
            {
                throw lastError;
            }

            throw new TalkServiceException("Talk room could not be created.", false, HttpStatusCode.InternalServerError, null);
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
            ExecuteJsonRequest("GET", url, (string)null, out statusCode, out data);

            if (!IsSuccessStatus(statusCode))
            {
                message = "HTTP " + (int)statusCode;
                return false;
            }

            IDictionary<string, object> ocs = GetDictionary(data, "ocs");
            IDictionary<string, object> meta = GetDictionary(ocs, "meta");
            IDictionary<string, object> payload = GetDictionary(ocs, "data");

            string version = ExtractVersion(payload);
            string status = GetString(meta, "status");
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
                string payloadJson = _serializer.Serialize(payload);
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
            string payloadJson = _serializer.Serialize(payload);
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
            IDictionary<string, object> ocs = GetDictionary(parsed, "ocs");
            if (ocs != null)
            {
                ocs.TryGetValue("data", out listObj);
            }

            object[] list = listObj as object[];
            if (list == null)
            {
                // Some Talk versions wrap the list under { data: { participants: [...] } }.
                var dataDict = listObj as IDictionary<string, object>;
                if (dataDict != null)
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
                var dict = entry as IDictionary<string, object>;
                if (dict == null)
                {
                    continue;
                }

                string actorType = GetString(dict, "actorType") ?? string.Empty;
                string actorId = GetString(dict, "actorId") ?? string.Empty;
                int attendeeId = GetInt(dict, "attendeeId");
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

        private void TryUpdateDescription(string token, string description, bool isEventConversation, string baseUrl)
        {
            if (isEventConversation)
            {
                return;
            }

            if (description == null)
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

        private static string ExtractOcsMessage(IDictionary<string, object> responseData)
        {
            IDictionary<string, object> meta = GetDictionary(GetDictionary(responseData, "ocs"), "meta");
            IDictionary<string, object> payload = GetDictionary(GetDictionary(responseData, "ocs"), "data");
            string message = GetString(meta, "message");
            if (string.IsNullOrEmpty(message))
            {
                message = GetString(payload, "error");
            }

            return message;
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
            string payloadText = null;
            if (payload is string)
            {
                payloadText = (string)payload;
            }
            else if (payload != null)
            {
                payloadText = _serializer.Serialize(payload);
            }

            return SendJsonRequest(method, url, payloadText, out statusCode, out parsedData);
        }
        private string SendJsonRequest(string method, string url, string payload, out HttpStatusCode statusCode, out IDictionary<string, object> parsedData)
        {
            HttpWebResponse response = null;
            string responseText = null;
            parsedData = null;

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = method;
                request.Accept = "application/json";
                request.Headers["OCS-APIRequest"] = "true";
                request.Headers["Authorization"] = HttpAuthUtilities.BuildBasicAuthHeader(_configuration.Username, _configuration.AppPassword);
                request.Timeout = 60000;

                bool hasBody = !string.Equals(method, "GET", StringComparison.OrdinalIgnoreCase)
                               && !string.Equals(method, "DELETE", StringComparison.OrdinalIgnoreCase);

                if (hasBody)
                {
                    request.ContentType = "application/json";
                    if (!string.IsNullOrEmpty(payload))
                    {
                        using (StreamWriter writer = new StreamWriter(request.GetRequestStream()))
                        {
                            writer.Write(payload);
                        }
                    }
                    else
                    {
                        request.ContentLength = 0;
                    }
                }

                LogApi(method + " " + url);
                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                }
                catch (WebException ex)
                {
                    response = ex.Response as HttpWebResponse;
                    if (response == null)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Api, "HTTP connection error.", ex);
                        throw new TalkServiceException("HTTP connection error: " + ex.Message, false, 0, null);
                    }
                }

                statusCode = response.StatusCode;
                LogApi(method + " " + url + " -> " + statusCode);

                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream ?? Stream.Null))
                {
                    responseText = reader.ReadToEnd();
                }

                if (!string.IsNullOrEmpty(responseText))
                {
                    try
                    {
                        parsedData = _serializer.DeserializeObject(responseText) as IDictionary<string, object>;
                    }
                    catch (Exception ex)
                    {
                        parsedData = null;
                        DiagnosticsLogger.LogException(LogCategories.Api, "Failed to parse JSON response.", ex);
                    }
                }

                return responseText;
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private static bool IsSuccessStatus(HttpStatusCode statusCode)
        {
            return statusCode >= HttpStatusCode.OK && statusCode < HttpStatusCode.MultipleChoices;
        }

        private static string ExtractRoomToken(IDictionary<string, object> responseData)
        {
            if (responseData == null)
            {
                return null;
            }

            IDictionary<string, object> ocs = GetDictionary(responseData, "ocs");
            IDictionary<string, object> data = GetDictionary(ocs, "data");
            if (data != null)
            {
                string token = GetString(data, "token");
                if (string.IsNullOrEmpty(token))
                {
                    token = GetString(data, "roomToken");
                }

                if (!string.IsNullOrEmpty(token))
                {
                    return token;
                }
            }

            return GetString(responseData, "token");
        }

        private static Dictionary<string, object> GetDictionary(IDictionary<string, object> parent, string key)
        {
            if (parent == null)
            {
                return null;
            }

            object value;
            if (parent.TryGetValue(key, out value))
            {
                return value as Dictionary<string, object>;
            }

            return null;
        }

        private static string GetString(IDictionary<string, object> parent, string key)
        {
            if (parent == null)
            {
                return null;
            }

            object value;
            if (parent.TryGetValue(key, out value) && value != null)
            {
                if (value is string)
                {
                    return (string)value;
                }

                if (value is IDictionary<string, object>)
                {
                    return null;
                }

                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return null;
        }

        private static int GetInt(IDictionary<string, object> parent, string key)
        {
            if (parent == null)
            {
                return 0;
            }

            object value;
            if (parent.TryGetValue(key, out value) && value != null)
            {
                if (value is int)
                {
                    return (int)value;
                }

                int parsed;
                if (int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
                {
                    return parsed;
                }
            }

            return 0;
        }

        private static string ExtractVersion(IDictionary<string, object> payload)
        {
            if (payload == null)
            {
                return null;
            }

            string version = GetString(payload, "versionstring") ?? GetString(payload, "version");
            string edition = GetString(payload, "edition");
            string result = ComposeVersion(version, edition);
            if (!string.IsNullOrEmpty(result))
            {
                return EnsureProductPrefix(result, payload);
            }

            IDictionary<string, object> versionDict = GetDictionary(payload, "version");
            result = ComposeVersion(
                GetString(versionDict, "string") ?? BuildVersionFromParts(versionDict),
                GetString(versionDict, "edition"));
            if (!string.IsNullOrEmpty(result))
            {
                return EnsureProductPrefix(result, payload);
            }

            IDictionary<string, object> nextcloud = GetDictionary(payload, "nextcloud");
            IDictionary<string, object> system = GetDictionary(nextcloud, "system");
            if (system != null)
            {
                result = ComposeVersion(
                    GetString(system, "versionstring") ?? GetString(system, "version") ?? BuildVersionFromParts(system),
                    GetString(system, "edition"));
                if (!string.IsNullOrEmpty(result))
                {
                    string product = GetString(system, "productname") ?? GetString(nextcloud, "productname");
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
        {
            if (dictionary == null)
            {
                return null;
            }

            string major = GetString(dictionary, "major");
            string minor = GetString(dictionary, "minor");
            string micro = GetString(dictionary, "micro");

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
            string product = GetString(payload, "productname");
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
            IDictionary<string, object> meta = GetDictionary(GetDictionary(responseData, "ocs"), "meta");
            string message = GetString(meta, "message");
            IDictionary<string, object> payload = GetDictionary(GetDictionary(responseData, "ocs"), "data");
            string errorDetail = GetString(payload, "error");

            StringBuilder builder = new StringBuilder();

            if (!string.IsNullOrEmpty(message))
            {
                builder.Append(message);
            }
            if (!string.IsNullOrEmpty(errorDetail))
            {
                if (builder.Length > 0)
                {
                    builder.Append(" / ");
                }
                builder.Append(errorDetail);
            }
            if (builder.Length == 0)
            {
                builder.Append("HTTP ").Append((int)statusCode);
            }

            bool authError = statusCode == HttpStatusCode.Unauthorized || statusCode == HttpStatusCode.Forbidden;
            throw new TalkServiceException(builder.ToString(), authError, statusCode, responseText);
        }

        private static bool ShouldFallbackToStandard(HttpStatusCode statusCode, IDictionary<string, object> responseData)
        {
            if (statusCode == HttpStatusCode.Conflict ||
                statusCode == HttpStatusCode.NotImplemented ||
                statusCode == HttpStatusCode.BadRequest ||
                statusCode == (HttpStatusCode)422)
            {
                return true;
            }

            IDictionary<string, object> meta = GetDictionary(GetDictionary(responseData, "ocs"), "meta");
            IDictionary<string, object> payload = GetDictionary(GetDictionary(responseData, "ocs"), "data");
            string message = GetString(meta, "message");
            if (string.IsNullOrEmpty(message))
            {
                message = GetString(payload, "error");
            }

            if (!string.IsNullOrEmpty(message))
            {
                string normalized = message.ToLowerInvariant();
                if (normalized.IndexOf("object", StringComparison.Ordinal) >= 0 &&
                    normalized.IndexOf("event", StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string BuildEventObjectId(TalkRoomRequest request)
        {
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




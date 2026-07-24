// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Threading;
using System.Xml;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Handles exact DAV resource probes and their XML responses.
    internal sealed partial class FileLinkDavClient
    {
        private const string DavCollectionProbeXml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<d:propfind xmlns:d=\"DAV:\"><d:prop><d:resourcetype/></d:prop></d:propfind>";
        private const string DavContentLengthProbeXml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<d:propfind xmlns:d=\"DAV:\"><d:prop>"
            + "<d:getcontentlength/><d:resourcetype/>"
            + "</d:prop></d:propfind>";

        internal bool CollectionExists(
            string url,
            string failureMessage,
            CancellationToken cancellationToken)
        {
            NcHttpResponse response = ProbeResource(
                url,
                cancellationToken,
                failureMessage);
            return response.StatusCode != HttpStatusCode.NotFound
                   && ResponseContainsCollection(response.ResponseText);
        }

        internal bool ResourceExists(
            string url,
            string failureMessage,
            CancellationToken cancellationToken)
        {
            NcHttpResponse response = ProbeResource(
                url,
                cancellationToken,
                failureMessage);
            return response.StatusCode != HttpStatusCode.NotFound;
        }

        private NcHttpResponse ProbeResource(
            string url,
            CancellationToken cancellationToken,
            string failureMessage)
        {
            NcHttpResponse response = SendWithRetry(
                () => new NcHttpRequestOptions
                {
                    Method = "PROPFIND",
                    Url = url,
                    Payload = DavCollectionProbeXml,
                    ContentType = "application/xml; charset=utf-8",
                    TimeoutMs = 60000,
                    ReadWriteTimeoutMs = 60000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false,
                    CancellationToken = cancellationToken,
                    ConnectionLimit = FileLinkUploadPolicy.MaxParallelRequests,
                    Headers = new Dictionary<string, string>
                    {
                        { "Depth", "0" }
                    }
                },
                "folder_probe",
                cancellationToken,
                null);

            if (response == null || !response.HasHttpResponse)
            {
                ThrowProbeFailure(
                    response,
                    failureMessage,
                    cancellationToken);
            }
            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                return response;
            }
            bool successful =
                response.StatusCode == HttpStatusCode.OK
                || (int)response.StatusCode == 207;
            if (response.StatusCode == HttpStatusCode.Unauthorized
                || response.StatusCode == HttpStatusCode.Forbidden
                || !successful)
            {
                ThrowProbeFailure(
                    response,
                    failureMessage,
                    cancellationToken);
            }

            return response;
        }

        internal bool ResourceHasContentLength(
            string url,
            long expectedLength,
            string failureMessage,
            CancellationToken cancellationToken)
        {
            if (expectedLength < 0)
            {
                throw new ArgumentOutOfRangeException("expectedLength");
            }

            NcHttpResponse response = SendWithRetry(
                () => new NcHttpRequestOptions
                {
                    Method = "PROPFIND",
                    Url = url,
                    Payload = DavContentLengthProbeXml,
                    ContentType = "application/xml; charset=utf-8",
                    TimeoutMs = 60000,
                    ReadWriteTimeoutMs = 60000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false,
                    CancellationToken = cancellationToken,
                    ConnectionLimit =
                        FileLinkUploadPolicy.MaxParallelRequests,
                    Headers = new Dictionary<string, string>
                    {
                        { "Depth", "0" }
                    }
                },
                "file_length_probe",
                cancellationToken,
                null);

            if (response == null || !response.HasHttpResponse)
            {
                ThrowProbeFailure(
                    response,
                    failureMessage,
                    cancellationToken);
            }
            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                return false;
            }
            if (response.StatusCode == HttpStatusCode.Unauthorized
                || response.StatusCode == HttpStatusCode.Forbidden)
            {
                ThrowProbeFailure(
                    response,
                    failureMessage,
                    cancellationToken);
            }
            if (response.StatusCode != HttpStatusCode.OK
                && (int)response.StatusCode != 207)
            {
                ThrowProbeFailure(
                    response,
                    failureMessage,
                    cancellationToken);
            }

            return ResponseContainsContentLength(
                response.ResponseText,
                expectedLength);
        }

        private static void ThrowProbeFailure(
            NcHttpResponse response,
            string failureMessage,
            CancellationToken cancellationToken)
        {
            bool hasDetailPlaceholder = !string.IsNullOrWhiteSpace(
                                            failureMessage)
                                        && failureMessage.IndexOf(
                                            "{0}",
                                            StringComparison.Ordinal) >= 0;
            if (!hasDetailPlaceholder)
            {
                ThrowFailure(
                    response,
                    failureMessage,
                    cancellationToken);
                return;
            }

            string detail;
            if (response != null
                && !response.HasHttpResponse
                && response.TransportException != null)
            {
                detail = response.TransportException.Message;
            }
            else if (response != null && response.HasHttpResponse)
            {
                detail = "HTTP "
                         + ((int)response.StatusCode).ToString(
                             CultureInfo.InvariantCulture);
            }
            else
            {
                detail = Strings.ErrorServerUnavailable;
            }

            ThrowFailure(
                response,
                string.Format(
                    CultureInfo.CurrentCulture,
                    failureMessage,
                    detail),
                cancellationToken,
                false);
        }

        private static bool ResponseContainsCollection(string responseText)
        {
            if (string.IsNullOrWhiteSpace(responseText))
            {
                return false;
            }

            try
            {
                var document = new XmlDocument();
                document.XmlResolver = null;
                document.LoadXml(responseText);
                XmlNodeList collectionNodes = document.GetElementsByTagName(
                    "collection",
                    "DAV:");
                return collectionNodes != null && collectionNodes.Count > 0;
            }
            catch (XmlException ex)
            {
                throw new TalkServiceException(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.FileLinkWizardFolderCheckFailedFormat,
                        ex.Message),
                    false,
                    0,
                    responseText);
            }
        }

        private static bool ResponseContainsContentLength(
            string responseText,
            long expectedLength)
        {
            if (string.IsNullOrWhiteSpace(responseText))
            {
                return false;
            }

            try
            {
                var document = new XmlDocument();
                document.XmlResolver = null;
                document.LoadXml(responseText);
                XmlNodeList responseNodes = document.GetElementsByTagName(
                    "response",
                    "DAV:");
                if (responseNodes == null || responseNodes.Count == 0)
                {
                    return NodeContainsFileLength(
                        document.DocumentElement,
                        expectedLength);
                }

                foreach (XmlNode responseNode in responseNodes)
                {
                    if (NodeContainsFileLength(
                        responseNode,
                        expectedLength))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (XmlException ex)
            {
                throw new TalkServiceException(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.FileLinkWizardFolderCheckFailedFormat,
                        ex.Message),
                    false,
                    0,
                    responseText);
            }
        }

        private static bool NodeContainsFileLength(
            XmlNode scope,
            long expectedLength)
        {
            XmlElement element = scope as XmlElement;
            if (element == null)
            {
                return false;
            }

            XmlNodeList resourceTypeNodes = element.GetElementsByTagName(
                "resourcetype",
                "DAV:");
            if (resourceTypeNodes == null
                || resourceTypeNodes.Count == 0)
            {
                return false;
            }
            foreach (XmlNode resourceTypeNode in resourceTypeNodes)
            {
                XmlElement resourceTypeElement =
                    resourceTypeNode as XmlElement;
                if (resourceTypeElement != null
                    && resourceTypeElement.GetElementsByTagName(
                        "collection",
                        "DAV:").Count > 0)
                {
                    return false;
                }
            }

            XmlNodeList lengthNodes = element.GetElementsByTagName(
                "getcontentlength",
                "DAV:");
            if (lengthNodes == null)
            {
                return false;
            }
            foreach (XmlNode lengthNode in lengthNodes)
            {
                long actualLength;
                if (lengthNode != null
                    && long.TryParse(
                        lengthNode.InnerText,
                        NumberStyles.Integer,
                        CultureInfo.InvariantCulture,
                        out actualLength)
                    && actualLength == expectedLength)
                {
                    return true;
                }
            }
            return false;
        }
    }
}

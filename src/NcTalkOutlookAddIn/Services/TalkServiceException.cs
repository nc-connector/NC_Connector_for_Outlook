/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Net;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Specialized exception for errors returned by the Nextcloud Talk REST API.
     */
    internal sealed class TalkServiceException : Exception
    {
        internal bool IsAuthenticationError { get; private set; }
        internal HttpStatusCode StatusCode { get; private set; }
        internal string ResponseBody { get; private set; }

        public TalkServiceException(string message, bool authenticationError, HttpStatusCode statusCode, string responseBody)
            : base(message)
        {
            IsAuthenticationError = authenticationError;
            StatusCode = statusCode;
            ResponseBody = responseBody;
        }
    }
}
